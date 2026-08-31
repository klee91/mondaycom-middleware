/**
 * helpers/generate.js — top-level generation + revision orchestration:
 * footer protection, template resolution, clarification questions,
 * generateHTML (routing to design or standard path), reviseHTML.
 */
const { anthropic, GEN_MODEL, CHEAP_MODEL, AI_SYSTEM_PROMPT } = require("./config");
const { fetchTemplateFromSharePoint, fetchTemplateIndex } = require("./sharepoint");
const { fetchWordDocContent, applyHeaderImage, fetchItemColumns } = require("./monday");
const { applyContentImages } = require("./images");
const { parseInstructions, applyVariables, fillTokensFromColumnMap } = require("./instructions");
const { logUsage, extractHtml } = require("./utils");
const {
  isDesignTemplate, designVariant, generateDesignHTML,
  protectFragileRegions, restoreFragileRegions, findCorruptedTokens,
  renumberArticles, fixStripImageDimensions,
} = require("./design");

// ═════════════════════════════════════════════
// Footer protection
// ═════════════════════════════════════════════
function extractFooter(html) {
  const match = html.match(/<!--\s*START FOOTER\s*-->([\s\S]*?)<!--\s*END FOOTER\s*-->/i);
  return match ? match[0] : null;
}
function hasFooter(html) {
  return /<!--\s*START FOOTER\s*-->/i.test(html) && /<!--\s*END FOOTER\s*-->/i.test(html);
}
function reattachFooter(html, footer) {
  const insertPoint = html.lastIndexOf("</tbody></table></div>");
  if (insertPoint !== -1) return html.slice(0, insertPoint) + "\n" + footer + "\n" + html.slice(insertPoint);
  return html.replace("</body>", footer + "\n</body>");
}

async function resolveTemplateName(ticket) {
  const raw   = (ticket.template || "").trim();
  const index = await fetchTemplateIndex();
  const keys  = Object.keys(index);

  // 1. exact, then 2. case-insensitive, then 3. normalized match (handles the
  //    dropdown label not being spelled identically to the SharePoint display
  //    name, e.g. "AI in Focus" vs "AI IN FOCUS NEWSLETTER").
  if (raw && index[raw]) return raw;
  const ci = keys.find(k => k.toLowerCase() === raw.toLowerCase());
  if (ci) return ci;
  const rn = normalizeTemplateName(raw);
  if (rn) {
    const nm = keys.find(k => normalizeTemplateName(k) === rn);
    if (nm) {
      console.log(`[template] "${raw}" matched to "${nm}" by normalized name`);
      return nm;
    }
  }

  // No match — this is almost always a misconfigured dropdown label. Warn
  // LOUDLY with the available options rather than silently defaulting, since a
  // silent default to a general template is exactly what produces a
  // "template returned verbatim" result on a newsletter ticket.
  console.warn(
    `[template] ⚠️ dropdown value "${raw}" did not match any template. ` +
    `Available: [${keys.join(" | ")}]. Defaulting to "CPACOM General" — ` +
    `if this was meant to be a newsletter, the design path will NOT run.`
  );
  return "CPACOM General";
}

// ═════════════════════════════════════════════
// Clarification questions
// ═════════════════════════════════════════════
// Per the brand brief, the agent should CONFIRM missing/ambiguous inputs with
// the requestor rather than silently guessing — especially CTA destination
// URLs, personalization, and (when unclear) the template choice. This runs one
// cheap Haiku pass over the ticket inputs + the generated HTML and returns a
// short list of questions to post back on the ticket. It never blocks the
// draft; the draft is posted alongside the questions so the requestor sees
// both. Returns [] on any failure so generation is never held up.
async function analyzeForQuestions(ticket, vars, generatedHtml) {
  try {
    // Cheap deterministic signals to focus the model.
    const hasAnyUrl = /href="https?:\/\//i.test(generatedHtml);
    const usesFirstName = /\{\{Recipient\.FirstName\}\}/.test(generatedHtml);
    const source = [ticket.description, vars["__freeform__"], vars["BodyContent"]].filter(Boolean).join("\n");

    const response = await anthropic.messages.create({
      model: CHEAP_MODEL,
      max_tokens: 400,
      system:
        "You review a marketing email ticket for a proof generator and list ONLY the " +
        "genuinely blocking clarifications the requestor must answer before the email can " +
        "be finalized. Follow this checklist: (1) every CTA/button must have a real " +
        "destination URL — if any is missing or was left as a placeholder, ask for it; " +
        "(2) confirm whether {{Recipient.FirstName}} personalization is intended if it's " +
        "unclear; (3) any figure, stat, date, discount code, or claim that appears to be " +
        "missing or a placeholder (TBD/TBC/XX) — ask, never invent. " +
        "Respond ONLY with JSON: {\"questions\": [\"...\"]}. Each question one sentence, " +
        "specific and actionable. If nothing is genuinely blocking, return " +
        "{\"questions\": []}. Do NOT ask about styling, tone, or things already provided.",
      messages: [{
        role: "user",
        content:
          `Subject: ${ticket.subjectLine || ticket.name || ""}\n` +
          `Template: ${ticket.template || ""}\n` +
          `Provided content / instructions:\n${source || "(none beyond subject)"}\n\n` +
          `Signals from the generated draft: containsAtLeastOneRealUrl=${hasAnyUrl}, ` +
          `usesFirstNamePersonalization=${usesFirstName}\n\n` +
          `Draft HTML (for detecting empty/placeholder CTAs and TBD copy):\n${(generatedHtml || "").slice(0, 12000)}`,
      }],
    });
    const text = response.content.find(b => b.type === "text")?.text ?? "{}";
    let parsed;
    try { parsed = JSON.parse(text.replace(/```json|```/g, "").trim()); } catch { return []; }
    const qs = Array.isArray(parsed.questions) ? parsed.questions.filter(q => typeof q === "string" && q.trim()) : [];
    return qs.slice(0, 5);
  } catch (err) {
    console.warn(`[agent] clarification analysis skipped: ${err.message}`);
    return [];
  }
}

// Render a questions list into an HTML update snippet, or "" if none.
function renderQuestionsBlock(questions) {
  if (!questions || questions.length === 0) return "";
  const items = questions.map(q => `<li>${q}</li>`).join("");
  return `<p>Before this is final, a few things to confirm — reply with <strong>@MEG</strong> and the answers:</p><ul>${items}</ul>`;
}

// ═════════════════════════════════════════════
// Fallback token fill from Monday columns
// ═════════════════════════════════════════════
// After normal substitution, fill any remaining {{Token}} placeholders from a
// Monday column whose NAME matches the token — leaving anything without a
// confident match (and all Pardot/namespaced/triple-brace tags) untouched.
// Non-blocking: any failure just returns the html unchanged.
async function finalizeTokens(html, ticket) {
  if (!ticket?.id || ticket.id === "manual") return html;
  try {
    const columns = await fetchItemColumns(ticket.id);
    if (!columns.length) return html;
    const { html: out, filled, skipped } = fillTokensFromColumnMap(html, columns);
    if (filled.length) console.log(`[tokens] filled ${filled.length} from columns: ${filled.join(", ")}`);
    if (skipped.length) console.log(`[tokens] left ${skipped.length} unmatched (no column / ambiguous): ${skipped.join(", ")}`);
    return out;
  } catch (err) {
    console.warn(`[tokens] column fill skipped for ${ticket.id}: ${err.message}`);
    return html;
  }
}

// ═════════════════════════════════════════════
// Generate HTML (first proof)
// ═════════════════════════════════════════════
async function generateHTML(ticket, templateName) {
  const templateHtml = await fetchTemplateFromSharePoint(templateName);
  if (!templateHtml) throw new Error(`Could not load template: "${templateName}"`);

  const vars = parseInstructions(ticket.instructions);
  vars["Subject"]   = ticket.subjectLine || ticket.name || "";
  vars["JobNumber"] = ticket.jobNumber || "";

  // Routing diagnostic — makes the "why did it pick this path" question
  // answerable from a single log line.
  console.log(
    `[generate] "${ticket.name}" | dropdown="${ticket.template || ""}" ` +
    `resolved="${templateName}" designPath=${isDesignTemplate(templateName)}`
  );

  // ── Design (repeating-block) templates: model-driven path with protected
  //    regions + deterministic renumbering. Uses the RAW template, so this
  //    runs BEFORE applyVariables to avoid disturbing article markers. ──
  if (isDesignTemplate(templateName)) {
    const docContent = (ticket.id && ticket.id !== "manual")
      ? await fetchWordDocContent(ticket.id) : null;
    const sourceContent = docContent?.html || vars["BodyContent"] || vars["__freeform__"] || "";
    // Only pass __freeform__ as *separate* instructions when it isn't already
    // the source, so it isn't duplicated into the prompt twice.
    ticket.__freeform__ = (sourceContent === vars["__freeform__"]) ? "" : (vars["__freeform__"] || "");
    try {
      let html = await generateDesignHTML(ticket, templateName, templateHtml, sourceContent);
      if (ticket.id && ticket.id !== "manual") {
        html = await applyHeaderImage(html, ticket.id, vars);
        html = await applyContentImages(html, ticket.id, vars);
      }
      return await finalizeTokens(html, ticket);
    } catch (e) {
      if (e.code !== "NO_CANONICAL_BLOCK") throw e;
      console.warn(`[design] falling back to standard path for "${ticket.name}": ${e.message}`);
      // fall through to the standard placeholder path below
    }
  }

  const originalFooter = extractFooter(templateHtml);
  let html = applyVariables(templateHtml, vars, ticket.instructions || "");

  const stillMissing = [];
  if (/BODY_CONTENT_HERE/.test(html) || /\{\{BodyContent\}\}/.test(html) || /\{\{BodyText\}\}/.test(html)) stillMissing.push("BodyContent");
  if (/\{\{PreheaderText\}\}/.test(html)) stillMissing.push("PreheaderText");

  let docContent = null;
  if (stillMissing.includes("BodyContent") && ticket.id && ticket.id !== "manual") {
    docContent = await fetchWordDocContent(ticket.id);
  }

  if (stillMissing.length > 0) {
    const message = await anthropic.messages.create({
      model: GEN_MODEL,
      max_tokens: 16000,
      system: [{ type: "text", text: AI_SYSTEM_PROMPT, cache_control: { type: "ephemeral" } }],
      messages: [{
        role: "user",
        content: `The following placeholders still need content: ${stillMissing.join(", ")}

Ticket data:
Name: ${ticket.name}
Template: ${templateName}
Subject Line: ${vars["Subject"]}
Job Number: ${vars["JobNumber"]}
Description: ${ticket.description || "No description provided."}
Category: ${ticket.category || ""}
Product: ${ticket.product || ""}
${vars["__freeform__"] ? `\nAdditional instructions from the requestor (apply these when generating content):\n${vars["__freeform__"]}` : ""}${docContent ? `\nSOURCE CONTENT (from the attached Word document "${docContent.name}"). Use this as the basis for the email body. Rewrite/reformat it into brand-compliant, table-based email HTML for the BODY_CONTENT_HERE region — preserve its meaning, structure, headings, and links, but apply CPA.com styling. Do not copy any Word styling verbatim:\n${docContent.html}\n` : ""}
HTML TEMPLATE (partially populated — only fill the remaining placeholders listed above):
${html}`,
      }],
    });
    const rawGenerated = message.content.find(b => b.type === "text")?.text ?? "";
    logUsage(`generate "${ticket.name}"`, message);
    html = extractHtml(rawGenerated, html);
  }

  if (originalFooter && !hasFooter(html)) {
    console.warn(`Footer missing for "${ticket.name}" — reattaching.`);
    html = reattachFooter(html, originalFooter);
  }

  if (ticket.id && ticket.id !== "manual") {
    html = await applyHeaderImage(html, ticket.id, vars);
    html = await applyContentImages(html, ticket.id, vars);
  }

  return await finalizeTokens(html, ticket);
}

// ═════════════════════════════════════════════
// Revise HTML (agent feedback loop)
// ═════════════════════════════════════════════
async function reviseHTML(currentHtml, feedback, history, freeform = "", opts = {}) {
  const { variant = null } = opts;

  // ── Input-size guard ──────────────────────────────────────────────────────
  // The context window is ~1M tokens. A single email is ~11K tokens, so a
  // healthy revise prompt is well under 30K. If currentHtml is wildly larger,
  // something upstream duplicated content (e.g. a corrupted/compounded state
  // blob) — sending it would blow the window with a 400. Fail loud and early
  // with a clear message rather than firing a doomed multi-MB request.
  const MAX_HTML_CHARS = 400_000; // ~100K tokens; ~10x a normal email, generous
  if ((currentHtml || "").length > MAX_HTML_CHARS) {
    throw new Error(
      `revise aborted: base HTML is ${currentHtml.length} chars (>${MAX_HTML_CHARS}), ` +
      `far larger than a normal email — likely a corrupted/compounded agent-state blob. ` +
      `Not sending to the model. Reset this item's agent state (delete the ` +
      `agent_state_*.json in Files) and regenerate.`
    );
  }
  // Only the last few feedback rounds matter for applying the next edit; an
  // unbounded history is both useless context and a slow token leak.
  const MAX_HISTORY = 8;
  const recent = history.length > MAX_HISTORY ? history.slice(-MAX_HISTORY) : history;
  const historyText = recent.length
    ? (history.length > MAX_HISTORY ? `(… ${history.length - MAX_HISTORY} earlier round(s) omitted)\n` : "") +
      recent.map((h, i) => `Round ${history.length - recent.length + i + 1}: ${h}`).join("\n")
    : "(none)";

  // Protect fragile regions so a revision can't re-warp strip images, break the
  // bracket buttons, or disturb the footer. The model edits around opaque
  // tokens; we restore them verbatim afterward.
  const { html: protectedHtml, store } = protectFragileRegions(currentHtml);

  const REVISE_RULES = `
REVISION FIDELITY — READ CAREFULLY:
- The CURRENT EMAIL HTML below is the SOURCE OF TRUTH, not a fresh template.
  It already reflects earlier approved corrections by the requestor.
- Apply ONLY the change described in NEW FEEDBACK. Reproduce EVERYTHING else
  byte-for-byte — same text, same values, same attributes, same order.
- Do NOT "restore", "normalize", or "clean up" anything. In particular:
  * If a spot contains a literal value where a {{merge token}} once was (e.g.
    a real date like "August 2026", a real URL, a job number), that is an
    INTENTIONAL prior edit — KEEP IT. Never revert a literal value back to a
    {{...}} token.
  * Do not re-introduce placeholder tokens that are no longer present.
  * Do not change links, dates, numbers, or copy that the feedback didn't
    mention.
- Tokens of the form <!--#PROTECTED:...#--> are opaque; copy them through
  verbatim and never alter their contents.
- Output the COMPLETE HTML, raw only (nothing before the first < or after the
  last >).`;

  const message = await anthropic.messages.create({
    model: GEN_MODEL,
    max_tokens: 16000,
    system: [{ type: "text", text: AI_SYSTEM_PROMPT + "\n" + REVISE_RULES, cache_control: { type: "ephemeral" } }],
    messages: [{
      role: "user",
      content: `You previously generated this email proof. The requestor has reviewed it and given feedback. Apply ONLY the requested change; keep everything else exactly as it is in the CURRENT EMAIL HTML.
${freeform ? `\nSTANDING INSTRUCTIONS FROM REQUESTOR (remain in effect across all revisions):\n${freeform}\n` : ""}
PRIOR FEEDBACK ROUNDS (already applied — do not undo these):
${historyText}

NEW FEEDBACK TO APPLY:
${feedback}

CURRENT EMAIL HTML (source of truth — preserve verbatim except for the change above):
${protectedHtml}`,
    }],
  });

  const rawRevised = message.content.find(b => b.type === "text")?.text ?? "";
  logUsage("revise", message);
  let html = extractHtml(rawRevised, protectedHtml);

  // Integrity + restore protected regions.
  const corrupted = findCorruptedTokens(html, store);
  if (corrupted.length > 0) console.warn(`[revise] ${corrupted.length} corrupted protected token(s): ${corrupted.join(", ")}`);
  html = restoreFragileRegions(html, store).html;

  // Re-run the deterministic design passes so a revision can never regress
  // article numbering or strip-image dimensions.
  if (variant === "aiInFocus") {
    html = renumberArticles(html, variant).html;
    html = (await fixStripImageDimensions(html, variant, "revision")).html;
  }

  return html;
}

module.exports = { extractFooter, hasFooter, reattachFooter, resolveTemplateName, analyzeForQuestions, renderQuestionsBlock, generateHTML, reviseHTML };