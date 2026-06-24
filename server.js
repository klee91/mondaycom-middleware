/**
 * Monday.com Email Template Generator — Full Stack Agent Server
 *
 * Capabilities:
 *   - Serves frontend UI + REST API (manual generation)
 *   - SharePoint template library (Graph client-credentials)
 *   - Instructions-column variable substitution + brand buttons
 *   - Header image re-hosting from Files column
 *   - Footer + Pardot-tag protection, brand style enforcement
 *   - Stateful agent webhook: item created → proof; @agent feedback → revise;
 *     Status=Approved → silent finalize
 *   - Generation model: claude-opus-4-6 (swap to claude-haiku-4-5-20251001 once fine-tuned)
 *
 * Required env vars:
 *   MONDAY_API_TOKEN, ANTHROPIC_API_KEY, BOARD_ID
 *   SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_DRIVE_ID
 *   HEADER_UPLOAD_FOLDER_URL (or HEADER_UPLOAD_FOLDER_ID)
 */

const express   = require("express");
const cors      = require("cors");
const fetch     = require("node-fetch").default || require("node-fetch");
const FormData  = require("form-data");
const Anthropic = require("@anthropic-ai/sdk").default;
const path      = require("path");

const app = express();
app.use(cors({ origin: "*", methods: ["GET", "POST", "OPTIONS"], allowedHeaders: ["Content-Type", "Authorization"] }));
app.options("*", cors());
app.use(express.json({ limit: "10mb" }));
app.use((req, res, next) => { res.setHeader("Access-Control-Allow-Origin", "*"); next(); });

const MONDAY_API_URL  = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID        = process.env.BOARD_ID;
const anthropic       = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const AGENT_STATE_COLUMN = "long_text_mm4b1t9h"; // Agent State column
const INSTRUCTIONS_COLUMN = "long_text_mm47njms"; // Instructions column
const TEMPLATE_COLUMN = "dropdown_mm4d2e9v"; // Template dropdown column


const CONTENT_VARIABLES = ["PreheaderText", "BodyText", "Subject", "JobNumber"];

const BUTTON_COLORS = {
  "purple": "#72246c", "light purple": "#86387f", "green": "#48a23f", "navy": "#0f206c",
};

const AI_SYSTEM_PROMPT = `You are an email template generator for CPA.com. Populate the HTML template with ticket data, strictly following the CPA.com brand standards below.

OUTPUT RULES:
- Your response must start with < and end with >. Nothing before the opening < and nothing after the closing >.
- Do NOT wrap output in markdown code fences. Do NOT include \`\`\`html or \`\`\` anywhere.
- Keep ALL HTML structure, styles, images, and links completely intact.

PROTECTED — NEVER MODIFY:
- Any existing {{...}} or {{{...}}} tokens are Pardot merge tags. Leave them exactly as-is.
- The section between <!-- START FOOTER --> and <!-- END FOOTER --> must remain completely intact.
- Pre-built buttons already in the template (such as multi-cell "Read more" bracket buttons with side images) must keep their structure. Only change their href or link text if explicitly instructed.

CPA.COM BRAND STANDARDS (STRICT — match exactly, no deviation):
- Font: always 'Roboto', Arial, Helvetica, sans-serif.
- Body copy: font-size 16px, font-weight 300, line-height 22-23px, color #000000.
- Body paragraphs: each <p> uses margin: 0 0 10px 0 (or 0 0 20px 0 between major sections).
- Bold/strong text: font-weight 700.
- Greeting is always: Hi {{Recipient.FirstName}},
- Inline text links: color #86387f, text-decoration underline.
- Section/accent headers: font-weight 700, color #0f206c (navy).
- Bulleted lists: Roboto, 16px, weight 300, line-height 23px, padding-left 20-30px; items margin 0 0 10px 0.
- Do NOT introduce new colors, fonts, font sizes, or spacing values outside these standards.

CONTENT PLACEHOLDERS (only replace these):
- BODY_CONTENT_HERE → email body copy written from the ticket description.
- BUTTON_TEXT → relevant CTA based on ticket context.
- JOB_NUMBER_HERE / JOB_NUMBER → the job number, or remove the line if blank.

TONE: professional, clear, appropriate for accounting and finance professionals. If the description says TBD or TBC, write appropriate placeholder copy based on the subject line and product.`;

// ═════════════════════════════════════════════
// SharePoint: token + template fetch
// ═════════════════════════════════════════════
let sharepointToken = null;
let tokenExpiresAt  = 0;
const templateHtmlCache = {};  // { itemId: { html, fetchedAt } }
let   templateIndexCache = null; // { map: { displayName: itemId }, fetchedAt }
const CACHE_TTL_MS  = 10 * 60 * 1000;

async function getSharePointToken() {
  if (sharepointToken && Date.now() < tokenExpiresAt - 60000) return sharepointToken;

  const url  = `https://login.microsoftonline.com/${process.env.SHAREPOINT_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     process.env.SHAREPOINT_CLIENT_ID,
    client_secret: process.env.SHAREPOINT_CLIENT_SECRET,
    scope:         "https://graph.microsoft.com/.default",
  });
  const res  = await fetch(url, { method: "POST", body });
  const data = await res.json();
  if (!data.access_token) throw new Error(`SharePoint auth failed: ${data.error_description || data.error}`);

  sharepointToken = data.access_token;
  tokenExpiresAt  = Date.now() + data.expires_in * 1000;
  console.log("SharePoint token refreshed.");
  return sharepointToken;
}

// Convert a SharePoint filename to a human-friendly display name.
// e.g. "CPACOM_GENERAL_TEMPLATE.html" → "CPACOM General"
//      "CPA_PARTNER_(DOUBLE).html"    → "CPA Partner (Double)"
function fileNameToDisplayName(filename) {
  return filename
    .replace(/\.html$/i, "")
    .replace(/_TEMPLATE$/i, "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, c => c.toUpperCase());
}

// Fetch (and cache) the folder listing → { displayName: itemId }
async function fetchTemplateIndex() {
  if (templateIndexCache && Date.now() - templateIndexCache.fetchedAt < CACHE_TTL_MS) {
    return templateIndexCache.map;
  }
  const token = await getSharePointToken();
  const url   = `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/root/children?$select=id,name`;
  const res   = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Template index fetch failed: ${res.status} ${res.statusText}`);
  const data = await res.json();
  const map  = {};
  for (const item of (data.value || [])) {
    if (/\.html$/i.test(item.name)) map[fileNameToDisplayName(item.name)] = item.id;
  }
  templateIndexCache = { map, fetchedAt: Date.now() };
  console.log(`Template index refreshed: ${Object.keys(map).join(", ")}`);
  return map;
}

async function fetchTemplateFromSharePoint(templateName) {
  const index = await fetchTemplateIndex();

  // Exact match → case-insensitive match → CPACOM General fallback
  const itemId = index[templateName]
    ?? index[Object.keys(index).find(k => k.toLowerCase() === (templateName || "").toLowerCase())]
    ?? index["CPACOM General"];
  if (!itemId) throw new Error("No templates found in SharePoint folder.");

  const cached = templateHtmlCache[itemId];
  if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) return cached.html;

  const token   = await getSharePointToken();
  const metaRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${itemId}?select=id,name,@microsoft.graph.downloadUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!metaRes.ok) throw new Error(`SharePoint metadata failed: ${metaRes.status} ${metaRes.statusText}`);

  const meta        = await metaRes.json();
  const downloadUrl = meta["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) throw new Error(`No download URL for template "${templateName}"`);

  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) throw new Error(`SharePoint download failed: ${fileRes.status}`);

  const html = await fileRes.text();
  templateHtmlCache[itemId] = { html, fetchedAt: Date.now() };
  console.log(`Fetched template "${templateName}" from SharePoint (${html.length} bytes)`);
  return html;
}

// ═════════════════════════════════════════════
// Monday GraphQL helper
// ═════════════════════════════════════════════
async function mondayQuery(query, variables = {}) {
  const res = await fetch(MONDAY_API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: process.env.MONDAY_API_TOKEN, "API-Version": "2024-01" },
    body: JSON.stringify({ query, variables }),
  });
  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join(", "));
  return data.data;
}

function normalizeTicket(item) {
  const cols = {};
  for (const cv of item.column_values) cols[cv.id] = cv.text ?? cv.value ?? "";
  return {
    id:           item.id,
    name:         item.name,
    status:       cols["status"]             || "",
    description:  cols["long_text7"]         || "",
    subjectLine:  cols["text86"]             || "",
    jobNumber:    cols["formula"]            || "",
    sendDate:     cols["date4"]              || "",
    contentDue:   cols["date_mkx4g1zc"]     || "",
    type:         cols["dropdown3"]          || "",
    category:     cols["status_1"]           || "",
    product:      cols["dropdown2"]          || "",
    requestor:    cols["person"]             || "",
    instructions: cols[INSTRUCTIONS_COLUMN]  || "",
    agentState:   cols[AGENT_STATE_COLUMN]   || "",
    template:     cols[TEMPLATE_COLUMN]       || "",
    hasFiles:     !!cols["files"],
  };
}

async function fetchItemById(itemId) {
  const data = await mondayQuery(`
    query GetItem($itemId: ID!) {
      items(ids: [$itemId]) {
        id name
        column_values(ids: [
          "status","long_text7","text86","formula",
          "date4","date_mkx4g1zc","dropdown3",
          "status_1","dropdown2","person","files","${INSTRUCTIONS_COLUMN}","${AGENT_STATE_COLUMN}","${TEMPLATE_COLUMN}"
        ]) { id text value }
      }
    }
  `, { itemId });
  const item = data?.items?.[0];
  if (!item) throw new Error(`Item ${itemId} not found.`);
  return normalizeTicket(item);
}

// ═════════════════════════════════════════════
// Monday files: list, pick header image, download
// ═════════════════════════════════════════════
async function fetchTicketFiles(itemId) {
  const data = await mondayQuery(`
    query GetFiles($itemId: ID!) {
      items(ids: [$itemId]) { assets { id name url public_url file_extension } }
    }
  `, { itemId });
  return data?.items?.[0]?.assets ?? [];
}

function pickHeaderImage(assets) {
  const IMAGE_EXT = ["png", "jpg", "jpeg", "gif"];
  return assets.find(a => {
    const ext = (a.file_extension || a.name.split(".").pop() || "").toLowerCase();
    return IMAGE_EXT.includes(ext);
  });
}

async function downloadMondayAsset(asset) {
  const src = asset.public_url || asset.url;
  const res = await fetch(src);
  if (!res.ok) throw new Error(`Asset download failed: ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// ═════════════════════════════════════════════
// Header image re-hosting (SharePoint anonymous link)
// ═════════════════════════════════════════════
let cachedFolderId = null;
async function getHeaderFolderId() {
  if (cachedFolderId) return cachedFolderId;
  if (process.env.HEADER_UPLOAD_FOLDER_ID) { cachedFolderId = process.env.HEADER_UPLOAD_FOLDER_ID; return cachedFolderId; }

  const shareUrl = process.env.HEADER_UPLOAD_FOLDER_URL;
  if (!shareUrl) throw new Error("No HEADER_UPLOAD_FOLDER_ID or HEADER_UPLOAD_FOLDER_URL set.");
  const b64 = Buffer.from(shareUrl).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const shareId = "u!" + b64;

  const token = await getSharePointToken();
  const res   = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Folder resolve failed: ${res.status}`);
  const item = await res.json();
  cachedFolderId = item.id;
  console.log(`Resolved header upload folder ID: ${cachedFolderId}`);
  return cachedFolderId;
}

async function rehostToSharePoint(fileName, buffer) {
  const token  = await getSharePointToken();
  const folder = await getHeaderFolderId();

  const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${folder}:/${encodeURIComponent(fileName)}:/content`;
  const upRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
    body: buffer,
  });
  if (!upRes.ok) throw new Error(`SharePoint upload failed: ${upRes.status}`);
  const uploaded = await upRes.json();

  const linkRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${uploaded.id}/createLink`,
    { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify({ type: "view", scope: "anonymous" }) }
  );
  if (!linkRes.ok) throw new Error(`Share link creation failed: ${linkRes.status}`);
  const linkData = await linkRes.json();

  let webUrl = linkData?.link?.webUrl || "";
  if (webUrl && !webUrl.includes("download=1")) webUrl += (webUrl.includes("?") ? "&" : "?") + "download=1";
  return webUrl;
}

function replaceFirstImage(html, newUrl) {
  let replaced = false;
  return html.replace(/(<img\b[^>]*\bsrc=")([^"]*)(")/i, (m, pre, _src, post) => {
    if (replaced) return m;
    replaced = true;
    return pre + newUrl + post;
  });
}

async function applyHeaderImage(html, itemId) {
  try {
    const assets = await fetchTicketFiles(itemId);
    const header = pickHeaderImage(assets);
    if (!header) return html;
    const buf       = await downloadMondayAsset(header);
    const publicUrl = await rehostToSharePoint(`header_${itemId}_${header.name}`, buf);
    if (!publicUrl) return html;
    console.log(`Header image re-hosted: ${publicUrl}`);
    return replaceFirstImage(html, publicUrl);
  } catch (err) {
    console.warn(`Header image swap skipped for ${itemId}: ${err.message}`);
    return html;
  }
}

// ═════════════════════════════════════════════
// Brand buttons
// ═════════════════════════════════════════════
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

function parseButtons(html) {
  const blockPattern = /Button:\s*(.+?)\s*[\r\n]+((?:\s*(?:Link|Color|Style):\s*.+[\r\n]*){1,3})/gi;
  return html.replace(blockPattern, (match, btnText, rest) => {
    const link  = (rest.match(/Link:\s*(\S+)/i)  || [])[1];
    const color = (rest.match(/Color:\s*(.+)/i)  || [])[1];
    const style = (rest.match(/Style:\s*(\w+)/i) || [])[1];
    if (!link) return match;
    return buildButton(btnText.trim(), link.trim(), color && color.trim(), style && style.trim());
  });
}

// ═════════════════════════════════════════════
// Instructions parsing + variable substitution
// ═════════════════════════════════════════════
function parseInstructions(instructions) {
  const vars = {};
  if (!instructions) return vars;
  const VARIABLE_NAMES = ["PreheaderText", "PreheaderLink", "BodyText"];
  const pattern = new RegExp(
    `(${VARIABLE_NAMES.join("|")})\\s*:\\s*"?([^"\\n]*(?:\\n(?!(?:${VARIABLE_NAMES.join("|")})\\s*:)[^\\n]*)*)"?`,
    "gi"
  );
  let match;
  while ((match = pattern.exec(instructions)) !== null) {
    const key = match[1].trim();
    const val = match[2].trim().replace(/^"|"$/g, "");
    if (key && val) vars[key] = val;
  }
  return vars;
}

function applyVariables(html, vars) {
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
  if (vars["BodyText"]) out = out.replace(/BODY_CONTENT_HERE/g, vars["BodyText"]);

  // Job Number: support {{JobNumber}}, JOB_NUMBER_HERE, and plain JOB_NUMBER tokens
  if (vars["JobNumber"]) {
    out = out
      .replace(/\{\{JobNumber\}\}/g, vars["JobNumber"])
      .replace(/JOB_NUMBER_HERE/g, vars["JobNumber"])
      .replace(/\bJOB_NUMBER\b/g, vars["JobNumber"]);
  }

  out = parseButtons(out);
  return out;
}

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

// Resolve which template to use from the ticket's Template dropdown value.
// Falls back to "CPACOM General" if empty or unrecognized.
async function resolveTemplateName(ticket) {
  const raw   = (ticket.template || "").trim();
  const index = await fetchTemplateIndex();
  if (raw && index[raw]) return raw;
  const ci = Object.keys(index).find(k => k.toLowerCase() === raw.toLowerCase());
  if (ci) return ci;
  if (raw) console.log(`[template] "${raw}" not in library — defaulting to CPACOM General`);
  return "CPACOM General";
}

// ═════════════════════════════════════════════
// Generate HTML (first proof)
// ═════════════════════════════════════════════
async function generateHTML(ticket, templateName) {
  const templateHtml = await fetchTemplateFromSharePoint(templateName);
  if (!templateHtml) throw new Error(`Could not load template: "${templateName}"`);

  const originalFooter = extractFooter(templateHtml);

  const vars = parseInstructions(ticket.instructions);
  vars["Subject"]   = ticket.subjectLine || ticket.name || "";
  vars["JobNumber"] = ticket.jobNumber || "";

  let html = applyVariables(templateHtml, vars);

  const stillMissing = [];
  if (/BODY_CONTENT_HERE/.test(html) || /\{\{BodyText\}\}/.test(html)) stillMissing.push("BodyText");
  if (/\{\{PreheaderText\}\}/.test(html)) stillMissing.push("PreheaderText");
  if (/BUTTON_TEXT/.test(html)) stillMissing.push("ButtonText");

  if (stillMissing.length > 0) {
    const message = await anthropic.messages.create({
      model: "claude-opus-4-6",
      max_tokens: 8000,
      system: AI_SYSTEM_PROMPT,
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

HTML TEMPLATE (partially populated — only fill the remaining placeholders listed above):
${html}`,
      }],
    });
    html = message.content.find(b => b.type === "text")?.text ?? html;
    html = html.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
  }

  if (originalFooter && !hasFooter(html)) {
    console.warn(`Footer missing for "${ticket.name}" — reattaching.`);
    html = reattachFooter(html, originalFooter);
  }

  if (ticket.id && ticket.id !== "manual") {
    html = await applyHeaderImage(html, ticket.id);
  }

  return html;
}

// ═════════════════════════════════════════════
// Revise HTML (agent feedback loop)
// ═════════════════════════════════════════════
async function reviseHTML(currentHtml, feedback, history) {
  const originalFooter = extractFooter(currentHtml);
  const historyText = history.length ? history.map((h, i) => `Round ${i + 1}: ${h}`).join("\n") : "(none)";

  const message = await anthropic.messages.create({
    model: "claude-opus-4-6",
    max_tokens: 8000,
    system: AI_SYSTEM_PROMPT,
    messages: [{
      role: "user",
      content: `You previously generated this email proof. The requestor has reviewed it and given feedback. Apply ONLY the requested changes while keeping everything else intact and following all CPA.com brand standards.

PRIOR FEEDBACK ROUNDS:
${historyText}

NEW FEEDBACK TO APPLY:
${feedback}

CURRENT EMAIL HTML:
${currentHtml}`,
    }],
  });

  let html = message.content.find(b => b.type === "text")?.text ?? currentHtml;
  html = html.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
  if (originalFooter && !hasFooter(html)) html = reattachFooter(html, originalFooter);
  return html;
}

// ═════════════════════════════════════════════
// Upload to Monday (HTML or JSON)
// ═════════════════════════════════════════════
async function uploadToMonday(itemId, fileName, content, contentType = "text/html") {
  const query = `
    mutation AddFileToColumn($itemId: ID!, $columnId: String!, $file: File!) {
      add_file_to_column(item_id: $itemId, column_id: $columnId, file: $file) { id name url }
    }
  `;
  const form = new FormData();
  form.append("query", query);
  form.append("variables", JSON.stringify({ itemId, columnId: "files", file: null }));
  form.append("map", JSON.stringify({ "0": ["variables.file"] }));
  const buf = Buffer.from(content, "utf-8");
  form.append("0", buf, { filename: fileName, contentType, knownLength: buf.length });

  const res = await fetch(MONDAY_FILE_URL, {
    method: "POST",
    headers: { Authorization: process.env.MONDAY_API_TOKEN, "API-Version": "2024-01", ...form.getHeaders() },
    body: form,
  });
  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join(", "));
  return data?.data?.add_file_to_column;
}

// ═════════════════════════════════════════════
// Agent state
//   - Heavy data (currentHtml) → agent_state JSON file in Files column
//   - Lightweight metadata (revision, history, stateFile) → Agent State column
// ═════════════════════════════════════════════
async function readAgentMeta(itemId) {
  const data = await mondayQuery(`
    query GetState($itemId: ID!) {
      items(ids: [$itemId]) { column_values(ids: ["${AGENT_STATE_COLUMN}"]) { text } }
    }
  `, { itemId });
  const raw = data?.items?.[0]?.column_values?.[0]?.text || "";
  if (!raw) return { revision: 0, history: [], stateFile: null };
  try { return JSON.parse(raw); } catch { return { revision: 0, history: [], stateFile: null }; }
}

async function readCurrentHtml(itemId, meta) {
  if (!meta?.stateFile) return "";
  try {
    const assets = await fetchTicketFiles(itemId);
    const match  = assets.find(a => a.name === meta.stateFile);
    if (!match) return "";
    const buf  = await downloadMondayAsset(match);
    const blob = JSON.parse(buf.toString("utf-8"));
    return blob.currentHtml || "";
  } catch (err) {
    console.warn(`Could not recover HTML for ${itemId}: ${err.message}`);
    return "";
  }
}

async function persistAgentState(itemId, revision, history, currentHtml) {
  const stateFileName = `agent_state_${itemId}.json`;
  const blob = JSON.stringify({ revision, currentHtml });

  // 1. Heavy state file → Files column
  await uploadToMonday(itemId, stateFileName, blob, "application/json");

  // 2. Lightweight metadata → Agent State column
  const meta  = JSON.stringify({ revision, history, stateFile: stateFileName });
  const value = JSON.stringify({ text: meta });
  await mondayQuery(`
    mutation SetState($itemId: ID!, $boardId: ID!, $val: JSON!) {
      change_column_value(item_id: $itemId, board_id: $boardId, column_id: "${AGENT_STATE_COLUMN}", value: $val) { id }
    }
  `, { itemId, boardId: BOARD_ID, val: value });
}

async function postUpdate(itemId, body) {
  await mondayQuery(`
    mutation PostUpdate($itemId: ID!, $body: String!) {
      create_update(item_id: $itemId, body: $body) { id }
    }
  `, { itemId, body });
}

async function findUserIdByName(name) {
  if (!name) return null;
  // The person column may contain multiple names; use the first
  const first = name.split(",")[0].trim();
  const data = await mondayQuery(`query { users { id name } }`);
  const users = data?.users ?? [];
  const match = users.find(u => u.name.toLowerCase() === first.toLowerCase());
  return match ? match.id : null;
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

// GET /api/tickets
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
                  "status_1","dropdown2","person","files","${INSTRUCTIONS_COLUMN}","${AGENT_STATE_COLUMN}","${TEMPLATE_COLUMN}"
                ]) { id text value }
              }
            }
          }
        }
      }
    `, { boardId: BOARD_ID });

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

// POST /api/generate (manual)
app.post("/api/generate", async (req, res) => {
  const { ticket, templateName } = req.body;
  if (!ticket || !templateName) return res.status(400).json({ error: "Missing ticket or templateName" });
  try {
    const html = await generateHTML(ticket, templateName);
    res.json({ html });
  } catch (err) {
    console.error("POST /api/generate:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/upload (manual)
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
app.post("/api/webhook", async (req, res) => {
  const { challenge, event } = req.body;
  if (challenge) return res.json({ challenge });
  if (!event) return res.json({ status: "ignored" });

  res.json({ status: "received" });

  try {
    // ── Trigger 1: item created → first proof ──
    if (event.type === "create_pulse") {
      const itemId = String(event.pulseId);
      console.log(`[agent] New item ${itemId} — generating first proof`);

      const ticket = await fetchItemById(itemId);

      // Guardrail: only generate when the Instructions column has content
      if (!ticket.instructions || !ticket.instructions.trim()) {
        console.log(`[agent] Item ${itemId} has no Instructions — skipping generation`);
        return;
      }

      const templateName = await resolveTemplateName(ticket);
      const html         = await generateHTML(ticket, templateName);
      const fileName     = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v1.html`;

      // Upload proof + persist state BEFORE notifying — prevents feedback race
      await uploadToMonday(itemId, fileName, html);
      await persistAgentState(itemId, 1, [], html);

      const requestorId = await findUserIdByName(ticket.requestor);
      const mention = requestorId ? `[@${ticket.requestor.split(",")[0].trim()}](${requestorId}) ` : "";
      await postUpdate(itemId,
        `<p>${mention}Your email proof (v1) is ready and attached to this item's Files. ` +
        `Review it and reply with <strong>@agent</strong> followed by any changes you'd like. ` +
        `When it's ready, set Status to <strong>Approved</strong>.</p>`
      );
      console.log(`[agent] Item ${itemId} v1 complete`);
      return;
    }

    // ── Trigger 2: update posted → feedback if prefixed @agent ──
    if (event.type === "create_update") {
      const itemId = String(event.pulseId);
      const body   = (event.body || event.textBody || "").trim();
      const plain  = body.replace(/<[^>]+>/g, "").trim();

      if (!/^@agent\b/i.test(plain)) {
        console.log(`[agent] Update on ${itemId} ignored (no @agent prefix)`);
        return;
      }
      const feedback = plain.replace(/^@agent\b[:,\s]*/i, "").trim();
      if (!feedback) return;

      console.log(`[agent] Feedback on ${itemId}: "${feedback}"`);
      const ticket = await fetchItemById(itemId);

      if ((ticket.status || "").toLowerCase() === "approved") {
        console.log(`[agent] Item ${itemId} already approved — ignoring`);
        return;
      }

      const meta     = await readAgentMeta(itemId);
      const baseHtml = await readCurrentHtml(itemId, meta);
      if (!baseHtml) {
        await postUpdate(itemId, `<p>I don't have a prior draft stored for this item yet. If the first proof was just generated, give it a moment and try again.</p>`);
        return;
      }

      const revised  = await reviseHTML(baseHtml, feedback, meta.history || []);
      const newRev   = (meta.revision || 1) + 1;
      const fileName = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v${newRev}.html`;

      await uploadToMonday(itemId, fileName, revised);
      await persistAgentState(itemId, newRev, [...(meta.history || []), feedback], revised);

      const requestorId = await findUserIdByName(ticket.requestor);
      const mention = requestorId ? `[@${ticket.requestor.split(",")[0].trim()}](${requestorId}) ` : "";
      await postUpdate(itemId,
        `<p>${mention}Updated proof (v${newRev}) is attached with your requested changes. ` +
        `Reply with <strong>@agent</strong> for more edits, or set Status to <strong>Approved</strong> when ready.</p>`
      );
      console.log(`[agent] Item ${itemId} revised to v${newRev}`);
      return;
    }

    // ── Trigger 3: status changed → Approved = silent finalize ──
    if (event.type === "update_column_value" && event.columnId === "status") {
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

// Serve frontend UI
app.get("/app", (_, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));
