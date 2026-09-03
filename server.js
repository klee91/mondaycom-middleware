/**
 * Monday.com Email Template Generator — Full Stack Agent Server (entry point)
 *
 * This file wires up Express, the routes, and the stateful webhook. All domain
 * logic lives in ./helpers/*:
 *   config.js        — shared clients (Anthropic, fetch), env config, constants, system prompt
 *   sharepoint.js    — SharePoint template library (Graph client-credentials)
 *   github-store.js  — GitHub-hosted manifest + brand guide, auto-tagging, RAG retrieval
 *   design.js        — repeating-block (AI-in-Focus / Town Hall) generation pipeline
 *   instructions.js  — prompt-column parsing, brand buttons, body rendering, token substitution
 *   monday.js        — Monday GraphQL, ticket/file reads, Word-doc extraction, header image
 *   monday-state.js  — proof upload, agent-state persistence/recovery, ticket updates
 *   generate.js      — generateHTML / reviseHTML orchestration, clarification questions
 *   utils.js         — tiny shared helpers (logUsage, extractHtml)
 *
 * Required env vars:
 *   MONDAY_API_TOKEN, ANTHROPIC_API_KEY, BOARD_ID
 *   SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_DRIVE_ID
 *   GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, GITHUB_BRANCH,
 *   GITHUB_MANIFEST_PATH, GITHUB_BRAND_GUIDE_PATH
 */

const express = require("express");
const cors    = require("cors");
const path    = require("path");

const { BOARD_ID, INSTRUCTIONS_COLUMN, AGENT_STATE_COLUMN, TEMPLATE_COLUMN, FILES_COLUMN, STATUS_COLUMN } = require("./helpers/config");
const { getSharePointToken, fetchTemplateIndex } = require("./helpers/sharepoint");
const { getManifest } = require("./helpers/github-store");
const { designVariant } = require("./helpers/design");
const { parseInstructions } = require("./helpers/instructions");
const {
  mondayQuery, normalizeTicket, fetchItemById, fetchTicketFiles, fetchWordDocContent, fetchAllItemFiles,
} = require("./helpers/monday");
const {
  uploadToMonday, readAgentMeta, readCurrentHtml, persistAgentState, postUpdate, setStatus, recordProofVersion, findUserIdByName,
} = require("./helpers/monday-state");
const {
  resolveTemplateName, analyzeForQuestions, renderQuestionsBlock, generateHTML, reviseHTML,
} = require("./helpers/generation");

const app = express();
app.use(cors({ origin: "*", methods: ["GET", "POST", "OPTIONS"], allowedHeaders: ["Content-Type", "Authorization"] }));
app.options("*", cors());
app.use(express.json({ limit: "10mb" }));
app.use((req, res, next) => { res.setHeader("Access-Control-Allow-Origin", "*"); next(); });

// ═════════════════════════════════════════════
// Doc gate — the universal precondition for generation
// ═════════════════════════════════════════════
// A readable Word .docx is the SINGLE trigger for building an email. No matter
// the entry point (ticket creation, @MEG mention, manual endpoint), generation
// must not proceed without one. Returns the extracted doc content when present,
// or null when absent. When notify=true and absent, posts a precise message to
// the requestor explaining what to do. retry=true waits for a just-uploaded
// file to finish registering.
async function requireDoc(itemId, { notify = false, retry = true } = {}) {
  const doc = await fetchWordDocContent(itemId, { retry });
  if (doc) return doc;
  if (notify) {
    const found = (await fetchAllItemFiles(itemId).catch(() => []))
      .map(f => f.name)
      .filter(n => !/^agent_state_.*\.json$/i.test(n) && !/_v\d+\.html$/i.test(n)); // hide the agent's own files
    await postUpdate(itemId,
      `<p>I can't start an email draft without a Word document (<strong>.docx</strong>) to build from. ` +
      (found.length
        ? `I can see these attached: ${found.map(n => `<em>${n}</em>`).join(", ")}, but none is a readable .docx in the <strong>Files column</strong>. `
        : `Nothing readable is attached to the <strong>Files column</strong>. `) +
      `If the document is in the <strong>Files gallery</strong>, please move it to the <strong>Files column</strong> ` +
      `(the file cell on the item's row) — automation can't read the gallery. ` +
      `Then reply <strong>@MEG</strong> to generate.</p>`).catch(() => {});
    console.log(`[agent] Item ${itemId} — generation blocked: no readable .docx. Visible non-agent files: [${found.join(", ") || "∅"}].`);
  }
  return null;
}

// ═════════════════════════════════════════════
// ROUTES
// ═════════════════════════════════════════════
app.get("/",       (_, res) => res.json({ status: "ok", service: "Monday Email Agent" }));
app.get("/health", (_, res) => res.json({ status: "ok" }));

app.get("/debug/sharepoint", async (req, res) => {
  try {
    const token = await getSharePointToken();
    const index = await fetchTemplateIndex();
    const firstItemId = Object.values(index)[0];
    const metaRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${firstItemId}?select=id,name,@microsoft.graph.downloadUrl`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const metaData = await metaRes.json();
    let downloadStatus = null, downloadOk = null, contentPreview = null;
    const downloadUrl = metaData["@microsoft.graph.downloadUrl"];
    if (downloadUrl) {
      const fileRes = await fetch(downloadUrl);
      downloadStatus = fileRes.status; downloadOk = fileRes.ok;
      if (fileRes.ok) contentPreview = (await fileRes.text()).substring(0, 200);
    }
    res.json({
      tokenObtained: true,
      driveId: process.env.SHAREPOINT_DRIVE_ID ? "set" : "MISSING",
      metadataStatus: metaRes.status, metadataOk: metaRes.ok,
      metadataResponse: metaRes.ok ? { id: metaData.id, name: metaData.name } : metaData,
      hasDownloadUrl: !!downloadUrl, downloadStatus, downloadOk, contentPreview,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/tickets", async (req, res) => {
  try {
    const data = await mondayQuery(`
      query GetGroups($boardId: ID!) {
        boards(ids: [$boardId]) {
          groups {
            id title
            items_page(limit: 50) {
              items {
                id name
                column_values(ids: [
                  "status","long_text7","text86","formula",
                  "date4","date_mkx4g1zc","dropdown3",
                  "status_1","dropdown2","person","${FILES_COLUMN}","${INSTRUCTIONS_COLUMN}","${AGENT_STATE_COLUMN}","${TEMPLATE_COLUMN}"
                ]) {
                  id text value
                  ... on FormulaValue { display_value }
                }
              }
            }
          }
        }
      }
    `, { boardId: BOARD_ID }, "2025-07");

    const allGroups = data?.boards?.[0]?.groups ?? [];
    const matched   = allGroups.filter(g => g.title.toLowerCase().includes("new request"));
    const groups    = matched.length > 0 ? matched : allGroups;
    const tickets   = groups.flatMap(g => g.items_page?.items ?? []).map(normalizeTicket);
    res.json({ tickets });
  } catch (err) {
    console.error("GET /api/tickets:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/generate", async (req, res) => {
  const { ticket, templateName } = req.body;
  if (!ticket || !templateName) return res.status(400).json({ error: "Missing ticket or templateName" });
  try {
    // Universal gate: a readable .docx is required to build an email. For a real
    // item, verify one is attached before generating. (Skip the check only for
    // synthetic/manual tickets with no id, e.g. local template previews.)
    if (ticket.id && ticket.id !== "manual") {
      const doc = await requireDoc(ticket.id, { notify: false, retry: false });
      if (!doc) return res.status(422).json({ error: "No readable .docx attached to the item's Files column — cannot generate." });
    }
    const html = await generateHTML(ticket, templateName);
    res.json({ html });
  } catch (err) {
    console.error("POST /api/generate:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/upload", async (req, res) => {
  const { itemId, fileName, html } = req.body;
  if (!itemId || !fileName || !html) return res.status(400).json({ error: "Missing itemId, fileName, or html" });
  try {
    const asset = await uploadToMonday(itemId, fileName, html);
    res.json({ success: true, assetId: asset?.id, fileName: asset?.name, url: asset?.url });
  } catch (err) {
    console.error("POST /api/upload:", err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// POST /api/webhook — stateful agent
// In-memory de-duplication of webhook deliveries. Monday can deliver the same
// event more than once (duplicate subscriptions, at-least-once delivery), which
// otherwise produces multiple proofs from a single action. We fingerprint each
// event and skip any repeat seen within the TTL window. In-memory is sufficient
// here (single Render instance); if scaled to multiple instances this would need
// shared storage.
const recentEvents = new Map(); // fingerprint -> timestamp
const DEDUP_TTL_MS = 60 * 1000;
function isDuplicateEvent(event) {
  // Fingerprint: type + item + the salient payload (update text / column).
  const body = (event.body || event.textBody || "").replace(/<[^>]+>/g, "").trim().slice(0, 120);
  const fp = `${event.type}:${event.pulseId}:${event.columnId || ""}:${body}`;
  const now = Date.now();
  // prune old entries
  for (const [k, t] of recentEvents) if (now - t > DEDUP_TTL_MS) recentEvents.delete(k);
  if (recentEvents.has(fp)) return true;
  recentEvents.set(fp, now);
  return false;
}

app.post("/api/webhook", async (req, res) => {
  const { challenge, event } = req.body;
  if (challenge) return res.json({ challenge });
  if (!event) return res.json({ status: "ignored" });

  res.json({ status: "received" });

  if (isDuplicateEvent(event)) {
    console.log(`[agent] Duplicate ${event.type} for ${event.pulseId} within ${DEDUP_TTL_MS / 1000}s — skipping (likely duplicate webhook delivery/subscription).`);
    return;
  }
  console.log(`[agent] event=${event.type} pulseId=${event.pulseId} boardId=${event.boardId}`);

  // Keep the template manifest in sync with the SharePoint folder. getManifest()
  // diffs the live folder against manifest.json and auto-tags + commits any NEW
  // template (see ensureManifestUpToDate). It's cached (10 min) and only WRITES
  // when a new template is found, so this is cheap and usually a no-op — but it's
  // what makes the manifest a living index. Fire-and-forget so it never blocks or
  // breaks proof generation.
  getManifest().catch(err => console.warn(`[manifest] sync skipped: ${err.message}`));

  try {
    if (event.type === "create_pulse") {
      const itemId = String(event.pulseId);

      // Generation requires a readable .docx — the universal trigger. At
      // creation we BLOCK silently if it's missing (placeholder tickets stay
      // quiet — no comment). The requestor kicks off generation later with an
      // @MEG tag, which is where a missing-doc gets an explanatory reply.
      const doc = await requireDoc(itemId, { notify: false, retry: false });
      if (!doc) {
        console.log(`[agent] Item ${itemId} created without a readable .docx — staying quiet, waiting for doc + @MEG.`);
        return;
      }

      console.log(`[agent] New item ${itemId} with doc attached — generating first proof`);
      let ticket = await fetchItemById(itemId);

      if (!ticket.jobNumber) {
        console.log(`[agent] Job Number empty for ${itemId} — retrying once after delay`);
        await new Promise(r => setTimeout(r, 3000));
        ticket = await fetchItemById(itemId);
      }

      const templateName = await resolveTemplateName(ticket);
      const html         = await generateHTML(ticket, templateName);
      const fileName     = `${ticket.name.replace(/\s+/g, "_")}_v1.html`;

      const asset = await uploadToMonday(itemId, fileName, html);
      await persistAgentState(itemId, 1, [], html);
      await setStatus(itemId, "Proofing");
      await recordProofVersion(itemId, { version: 1, fileName, assetUrl: asset?.url });

      const questions = await analyzeForQuestions(ticket, parseInstructions(ticket.instructions), html);
      const requestorId = await findUserIdByName(ticket.requestor);
      const mentionName = ticket.requestor.split(",")[0].trim();
      const mentionTag  = requestorId ? `<strong>@${mentionName}</strong> ` : "";
      await postUpdate(itemId,
        `<p>${mentionTag}Your email proof (v1) is ready and attached to this item's Files. ` +
        `Review it and reply with <strong>@MEG</strong> followed by any changes you'd like. ` +
        `When it's ready, set Status to <strong>Approved</strong>.</p>` +
        renderQuestionsBlock(questions),
        requestorId ? [requestorId] : []
      );
      console.log(`[agent] Item ${itemId} v1 complete${questions.length ? ` (${questions.length} question(s) posted)` : ""}`);
      return;
    }

    if (event.type === "create_update") {
      const itemId = String(event.pulseId);
      const rawBody = event.body || event.textBody || "";
      console.log(`[agent] update raw body on ${itemId}: ${JSON.stringify(rawBody).slice(0, 300)}`);
      const plain = rawBody.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();

      // GUARD 1 — never act on the agent's OWN comments. Its posts contain these
      // signatures (and themselves mention @MEG), which would otherwise
      // re-trigger generation in a loop.
      const SELF_MARKERS = [
        "Your email proof",
        "Before this is final",
        "reply with @MEG and the answers",
        "Updated proof (v",
        "is ready and attached to this item",
      ];
      if (SELF_MARKERS.some(m => plain.includes(m))) {
        console.log(`[agent] Update on ${itemId} is the agent's own comment — ignoring (self-reply guard).`);
        return;
      }

      // GUARD 2 — the trigger must be at the START of the comment. Monday wraps
      // a mention as <span>@MEG</span>, so strip a leading zero-width char /
      // whitespace and require @MEG|@agent as the first token. This deliberately
      // does NOT match @MEG appearing mid-sentence (which is how the agent's own
      // replies reference it), only a genuine leading mention the requestor typed.
      const leading = plain.replace(/^[\uFEFF\u200B\s]+/, "");
      if (!/^@(MEG|agent)\b/i.test(leading)) {
        console.log(`[agent] Update on ${itemId} ignored (no leading @MEG mention). Stripped: "${leading.slice(0, 120)}"`);
        return;
      }
      const feedback = leading.replace(/^@(MEG|agent)\b[:,\s]*/i, "").trim();
      if (!feedback) {
        console.log(`[agent] @MEG mention on ${itemId} had no instruction text after it — treating as a bare trigger.`);
      }

      console.log(`[agent] Feedback on ${itemId}: "${feedback}"`);
      const ticket = await fetchItemById(itemId);

      if ((ticket.status || "").toLowerCase() === "approved") {
        console.log(`[agent] Item ${itemId} already approved — ignoring`);
        return;
      }

      const meta     = await readAgentMeta(itemId);
      const baseHtml = await readCurrentHtml(itemId, meta);

      // Routing is purely: no draft yet → generate a first proof from the doc;
      // a draft already exists → revise it. (There is intentionally no
      // requestor "regenerate" command — requestors only ever kick off the first
      // proof or request revisions.)
      if (!baseHtml) {
        console.log(`[agent] No prior draft on ${itemId} — checking for a document to build from.`);

        // Universal gate: no readable .docx → no draft. Notify the requestor.
        const doc = await requireDoc(itemId, { notify: true, retry: true });
        if (!doc) return;

        const templateName = await resolveTemplateName(ticket);
        const html         = await generateHTML(ticket, templateName);
        const version      = 1;
        const fileName     = `${ticket.name.replace(/\s+/g, "_")}_v${version}.html`;

        const asset = await uploadToMonday(itemId, fileName, html);
        await persistAgentState(itemId, version, [], html);
        await setStatus(itemId, "Proofing");
        await recordProofVersion(itemId, { version, fileName, assetUrl: asset?.url });

        const questions = await analyzeForQuestions(ticket, parseInstructions(ticket.instructions), html);
        const reqId    = await findUserIdByName(ticket.requestor);
        const reqName  = ticket.requestor.split(",")[0].trim();
        const reqTag   = reqId ? `<strong>@${reqName}</strong> ` : "";
        await postUpdate(itemId,
          `<p>${reqTag}Your email proof (v${version}) is ready and attached to this item's Files. ` +
          `Review it and reply with <strong>@MEG</strong> followed by any changes you'd like. ` +
          `When it's ready, set Status to <strong>Approved</strong>.</p>` +
          renderQuestionsBlock(questions),
          reqId ? [reqId] : []
        );
        console.log(`[agent] Item ${itemId} v${version} complete (via @MEG)${questions.length ? ` (${questions.length} question(s) posted)` : ""}`);
        return;
      }

      const ticketVars = parseInstructions(ticket.instructions);
      if (!feedback) {
        console.log(`[agent] Bare @MEG on ${itemId} with an existing draft and no instructions — nothing to revise, skipping.`);
        await postUpdate(itemId,
          `<p>Tag <strong>@MEG</strong> followed by the change you'd like (e.g. "@MEG make the intro shorter"), ` +
          `or set Status to <strong>Approved</strong> when the current proof is ready.</p>`);
        return;
      }
      const reviseTemplate = await resolveTemplateName(ticket);
      const revised  = await reviseHTML(
        baseHtml, feedback, meta.history || [], ticketVars["__freeform__"] || "",
        { variant: designVariant(reviseTemplate) }
      );
      const newRev   = (meta.revision || 1) + 1;
      const fileName = `${ticket.name.replace(/\s+/g, "_")}_v${newRev}.html`;

      const revAsset = await uploadToMonday(itemId, fileName, revised);
      await persistAgentState(itemId, newRev, [...(meta.history || []), feedback], revised);
      await recordProofVersion(itemId, { version: newRev, fileName, assetUrl: revAsset?.url });

      const requestorId = await findUserIdByName(ticket.requestor);
      const mentionName = ticket.requestor.split(",")[0].trim();
      const mentionTag  = requestorId ? `<strong>@${mentionName}</strong> ` : "";
      await postUpdate(itemId,
        `<p>${mentionTag}Updated proof (v${newRev}) is attached with your requested changes. ` +
        `Reply with <strong>@MEG</strong> for more edits, or set Status to <strong>Approved</strong> when ready.</p>`,
        requestorId ? [requestorId] : []
      );
      console.log(`[agent] Item ${itemId} revised to v${newRev}`);
      return;
    }

    if (event.type === "update_column_value" && event.columnId === STATUS_COLUMN) {
      const itemId = String(event.pulseId);
      const label  = event.value?.label?.text || event.value?.label || "";
      if ((label || "").toLowerCase() === "approved") {
        console.log(`[agent] Item ${itemId} approved — finalizing, no further action`);
      }
      return;
    }

    console.log(`[agent] Unhandled event type: ${event.type}`);
  } catch (err) {
    console.error(`[agent] Webhook processing error:`, err.message);
  }
});

app.get("/app", (_, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));
