/**
 * helpers/instructions.js — prompt-column parsing, brand buttons, body rendering, token substitution.
 */
const { BUTTON_COLORS, CONTENT_VARIABLES } = require("./config");

function resolveColor(name) {
  if (!name) return "#72246c";
  const key = name.trim().toLowerCase();
  if (BUTTON_COLORS[key]) return BUTTON_COLORS[key];
  if (/^#?[0-9a-f]{6}$/i.test(key)) return key.startsWith("#") ? key : `#${key}`;
  return "#72246c";
}

function buildButton(text, url, colorName, style = "solid") {
  const color = resolveColor(colorName);
  const isOutline = (style || "").toLowerCase() === "outline";
  const bg  = isOutline ? "#ffffff" : color;
  const txt = isOutline ? color : "#ffffff";
  return `<table align="center" border="0" cellpadding="0" cellspacing="0" style="margin:20px auto;"><tbody><tr><td align="center" bgcolor="${bg}" style="border-radius:3px;background:${bg};"><a href="${url}" style="font-size:16px;font-family:Roboto,Arial,Helvetica,sans-serif;border-radius:3px;padding:12px 18px;border:1px solid ${color};display:inline-block;text-decoration:none;color:${txt};font-weight:300;" target="_blank">${text}</a></td></tr></tbody></table>`;
}

// ═════════════════════════════════════════════
// Prompt column parsing (freeform-first)
// ═════════════════════════════════════════════
const SINGLE_LINE_LABELS = ["PreheaderText", "PreheaderLink", "HeaderImage", "HeaderLink", "Partner"];
const BLOCK_LABELS       = ["BodyContent", "BodyText", "Prompt"];

function parseInstructions(instructions) {
  const vars = {};
  if (!instructions) return vars;

  // 1. Pull out single-line labels (value = remainder of that line), removing
  //    each matched line from the text so it doesn't bleed into the prompt.
  let remaining = instructions;
  for (const label of SINGLE_LINE_LABELS) {
    const re = new RegExp(`^[ \\t]*${label}\\s*:\\s*(.*)$`, "gim");
    remaining = remaining.replace(re, (_m, val) => {
      const v = (val || "").trim().replace(/^"|"$/g, "");
      if (v) vars[label] = v;
      return "";
    });
  }

  // 2. Handle block labels within what's left.
  const blockPattern = new RegExp(
    `(${BLOCK_LABELS.join("|")})\\s*:\\s*([\\s\\S]*?)(?=(?:${BLOCK_LABELS.join("|")})\\s*:|$)`,
    "gi"
  );
  let match;
  let firstBlockIdx = -1;
  while ((match = blockPattern.exec(remaining)) !== null) {
    if (firstBlockIdx === -1) firstBlockIdx = match.index;
    const key = match[1].toLowerCase();
    const val = (match[2] || "").trim().replace(/^"|"$/g, "");
    if (!val) continue;
    if (key === "prompt") {
      vars["__freeform__"] = vars["__freeform__"] ? `${vars["__freeform__"]}\n\n${val}` : val;
    } else {
      vars["BodyContent"] = vars["BodyContent"] ? `${vars["BodyContent"]}\n\n${val}` : val;
    }
  }

  // 3. Everything NOT claimed by a block label is freeform prompt.
  if (firstBlockIdx === -1) {
    const all = remaining.trim();
    if (all) vars["__freeform__"] = vars["__freeform__"] ? `${all}\n\n${vars["__freeform__"]}` : all;
  } else {
    const lead = remaining.slice(0, firstBlockIdx).trim();
    if (lead) vars["__freeform__"] = vars["__freeform__"] ? `${lead}\n\n${vars["__freeform__"]}` : lead;
  }

  return vars;
}

// ─────────────────────────────────────────────
// BodyContent rendering
// ─────────────────────────────────────────────
const BODY_PARAGRAPH_STYLE = "margin:0 0 10px 0;font-family:'Roboto',Arial,Helvetica,sans-serif;font-size:16px;font-weight:300;line-height:22px;color:#000000;";

function renderTextSegment(text) {
  const trimmed = (text || "").trim();
  if (!trimmed) return "";
  return trimmed
    .split(/\n{2,}/)
    .map(p => `<p style="${BODY_PARAGRAPH_STYLE}">${p.trim().replace(/\n/g, "<br>")}</p>`)
    .join("\n");
}

function renderBodyContent(content) {
  if (!content) return "";
  const buttonPattern = /Button:\s*(.+?)[\r\n]+((?:\s*(?:Link|Color|Style):\s*.+[\r\n]*){1,3})/gi;
  let out = "";
  let lastIndex = 0;
  let m;
  while ((m = buttonPattern.exec(content)) !== null) {
    out += renderTextSegment(content.slice(lastIndex, m.index));
    const btnText = m[1].trim();
    const rest    = m[2];
    const link    = (rest.match(/Link:\s*(\S+)/i)  || [])[1];
    const color   = (rest.match(/Color:\s*(.+)/i)  || [])[1];
    const style   = (rest.match(/Style:\s*(\w+)/i) || [])[1];
    if (link) {
      out += "\n" + buildButton(btnText, link.trim(), color && color.trim(), style && style.trim()) + "\n";
    } else {
      out += renderTextSegment(m[0]);
    }
    lastIndex = m.index + m[0].length;
  }
  out += renderTextSegment(content.slice(lastIndex));
  return out;
}

function applyVariables(html, vars, instructions = "") {
  let out = html;

  if (vars["PreheaderText"] !== undefined) {
    let preheader = vars["PreheaderText"];
    if (vars["PreheaderLink"]) {
      preheader = `<a href="${vars["PreheaderLink"]}" style="color:#86387f;text-decoration:underline;" target="_blank">${vars["PreheaderText"]}</a>`;
    }
    out = out.replace(/\{\{PreheaderText\}\}/g, preheader);
  }

  for (const key of CONTENT_VARIABLES) {
    if (vars[key] !== undefined) out = out.replace(new RegExp(`\\{\\{${key}\\}\\}`, "g"), vars[key]);
  }

  if (vars["BodyContent"]) {
    const renderedBody = renderBodyContent(vars["BodyContent"]);
    out = out
      .replace(/BODY_CONTENT_HERE/g, renderedBody)
      .replace(/\{\{BodyContent\}\}/g, renderedBody)
      .replace(/\{\{BodyText\}\}/g, renderedBody);
  }

  if (vars["JobNumber"]) {
    out = out
      .replace(/\{\{JobNumber\}\}/g, vars["JobNumber"])
      .replace(/JOB_NUMBER_HERE/g, vars["JobNumber"])
      .replace(/\bJOB_NUMBER\b/g, vars["JobNumber"]);
  }

  return out;
}

// ─────────────────────────────────────────────
// Fallback token fill from Monday columns
// ─────────────────────────────────────────────
// After the normal substitution, a template may still contain {{Token}}
// placeholders that nothing filled (e.g. {{IssueDate}}, {{HeaderLink}}). As a
// default, try to fill each remaining DOUBLE-brace token from a Monday column
// whose NAME matches the token name; if there's no confident match, leave the
// token exactly as-is.
//
// Safety rules (false fills would put wrong data in a live email):
//   - Only double-brace {{...}}. Triple-brace {{{...}}} are Pardot tags — skip.
//   - Skip namespaced/dotted tokens like {{Recipient.FirstName}} (Pardot).
//   - Match on a normalized name (lowercase, alphanumeric only). Use a column
//     ONLY when the match is unambiguous: exactly one exact-normalized match,
//     or (failing that) exactly one column whose name contains / is contained
//     by the token name. Any ambiguity → leave the token alone.
//   - Never fill from an empty column value.
function normalizeName(s) {
  return (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function fillTokensFromColumnMap(html, columns) {
  // Build a normalized index of non-empty columns, tracking duplicate names.
  const byNorm = new Map();       // normTitle -> { value, title }
  const dupNorm = new Set();
  for (const c of (columns || [])) {
    const value = (c.value ?? "").toString().trim();
    if (!value) continue;
    const n = normalizeName(c.title);
    if (!n) continue;
    if (byNorm.has(n)) dupNorm.add(n);
    else byNorm.set(n, { value, title: c.title });
  }

  const filled = [];
  const skipped = [];

  // Match tokens with optional 3rd brace so we can detect and skip triple.
  const out = html.replace(/(\{\{\{?)\s*([^{}]+?)\s*(\}?\}\})/g, (whole, open, inner, close) => {
    if (open !== "{{" || close !== "}}") return whole; // triple-brace Pardot tag → leave
    const name = inner.trim();
    if (!name || name.includes(".")) return whole;      // namespaced Pardot → leave
    const n = normalizeName(name);
    if (!n) return whole;

    // 1. exact normalized match, unambiguous
    if (byNorm.has(n) && !dupNorm.has(n)) { filled.push(name); return byNorm.get(n).value; }

    // 2. unambiguous contains match (either direction)
    const cands = [...byNorm.entries()]
      .filter(([cn]) => !dupNorm.has(cn) && (cn.includes(n) || n.includes(cn)));
    if (cands.length === 1) { filled.push(name); return cands[0][1].value; }

    skipped.push(name); // no confident match → leave the token exactly as-is
    return whole;
  });

  return { html: out, filled: [...new Set(filled)], skipped: [...new Set(skipped)] };
}

module.exports = { resolveColor, buildButton, parseInstructions, renderTextSegment, renderBodyContent, applyVariables, fillTokensFromColumnMap, SINGLE_LINE_LABELS, BLOCK_LABELS };
