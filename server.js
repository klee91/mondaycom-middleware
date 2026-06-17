/**
 * Monday.com Email Template Generator — Full Stack Server
 * Serves frontend UI + API + Monday webhook listener
 * Templates pulled from SharePoint via Microsoft Graph (client credentials)
 *
 * Required env vars:
 *   MONDAY_API_TOKEN
 *   ANTHROPIC_API_KEY
 *   BOARD_ID
 *   SHAREPOINT_TENANT_ID
 *   SHAREPOINT_CLIENT_ID
 *   SHAREPOINT_CLIENT_SECRET
 *   SHAREPOINT_DRIVE_ID
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
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  next();
});

const MONDAY_API_URL  = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID        = process.env.BOARD_ID;
const anthropic       = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Map template names to SharePoint file item IDs
const TEMPLATE_MAP = {
  "AICPA Town Hall Newsletter": "0135ZG5SYRC5GCY6GDPFDJVT3R37W5MFI7",
  "DOTCPA General":             "0135ZG5S4OE2ZR2WA4BZELNKR6CP5ISLZ3",
  "CPACOM General":             "0135ZG5S36KVRZLIB3SFEL3RGEYDHPMOP3",
};

const CONTENT_VARIABLES = [
  "PreheaderText","BodyText","Subject","JobNumber"
];

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
- Body paragraphs: each <p> uses style margin: 0 0 10px 0 (or 0 0 20px 0 between major sections).
- Bold/strong text: font-weight 700.
- Greeting is always: Hi {{Recipient.FirstName}},
- Inline text links: color #86387f, text-decoration underline.
- Section/accent headers: font-weight 700, color #0f206c (navy).
- Bulleted lists: font-family Roboto, font-size 16px, font-weight 300, line-height 23px, padding-left 20-30px. List items use margin 0 0 10px 0.
- Do NOT introduce new colors, fonts, font sizes, or spacing values outside these standards.
- Do NOT add inline styles that conflict with the above.

CONTENT PLACEHOLDERS (only replace these):
- BODY_CONTENT_HERE → email body copy written from the ticket description, following the brand standards above.
- BUTTON_TEXT → relevant CTA based on ticket context.
- JOB_NUMBER_HERE / JOB_NUMBER → the job number, or remove the line if blank.

TONE: professional, clear, appropriate for accounting and finance professionals. If the description says TBD or TBC, write appropriate placeholder copy based on the subject line and product.`;

// ═════════════════════════════════════════════
// SharePoint: client credentials token + template fetch
// ═════════════════════════════════════════════
let sharepointToken = null;
let tokenExpiresAt  = 0;

const templateCache = {};
const CACHE_TTL_MS  = 10 * 60 * 1000; // 10 minutes

async function getSharePointToken() {
  if (sharepointToken && Date.now() < tokenExpiresAt - 60000) {
    return sharepointToken;
  }

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

async function fetchTemplateFromSharePoint(templateName) {
  const itemId = TEMPLATE_MAP[templateName];
  if (!itemId) throw new Error(`Unknown template: "${templateName}"`);

  const cached = templateCache[itemId];
  if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) {
    return cached.html;
  }

  const token = await getSharePointToken();

  // Step 1: fetch item metadata — includes @microsoft.graph.downloadUrl (pre-authenticated)
  const metaRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${itemId}?select=id,name,@microsoft.graph.downloadUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!metaRes.ok) throw new Error(`SharePoint metadata failed: ${metaRes.status} ${metaRes.statusText}`);

  const meta        = await metaRes.json();
  const downloadUrl = meta["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) throw new Error(`No download URL returned for template "${templateName}"`);

  // Step 2: download content from pre-authenticated URL — no auth header
  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) throw new Error(`SharePoint download failed: ${fileRes.status} ${fileRes.statusText}`);

  const html = await fileRes.text();
  templateCache[itemId] = { html, fetchedAt: Date.now() };
  console.log(`Fetched template "${templateName}" from SharePoint (${html.length} bytes)`);
  return html;
}

// ═════════════════════════════════════════════
// Monday GraphQL helper
// ═════════════════════════════════════════════
async function mondayQuery(query, variables = {}) {
  const res = await fetch(MONDAY_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization:  process.env.MONDAY_API_TOKEN,
      "API-Version":  "2024-01",
    },
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
    instructions: cols["long_text_mm47njms"] || "",
    agentState:   cols["long_text_mm4b1t9h"] || "",
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
          "status_1","dropdown2","person","files","long_text_mm47njms","long_text_mm4b1t9h"
        ]) { id text value }
      }
    }
  `, { itemId });
  const item = data?.items?.[0];
  if (!item) throw new Error(`Item ${itemId} not found.`);
  return normalizeTicket(item);
}

// ─────────────────────────────────────────────
// Agent state:
//   - Heavy data (currentHtml) → agent_state JSON file in Files column
//   - Lightweight metadata (revision, history) → Agent State long-text column
// ─────────────────────────────────────────────
const AGENT_STATE_COLUMN = "long_text_mm4b1t9h"; // Agent State column

async function readAgentMeta(itemId) {
  const data = await mondayQuery(`
    query GetState($itemId: ID!) {
      items(ids: [$itemId]) {
        column_values(ids: ["${AGENT_STATE_COLUMN}"]) { text }
      }
    }
  `, { itemId });
  const raw = data?.items?.[0]?.column_values?.[0]?.text || "";
  if (!raw) return { revision: 0, history: [], stateFile: null };
  try { return JSON.parse(raw); }
  catch { return { revision: 0, history: [], stateFile: null }; }
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

  // 1. Upload the heavy state file to Files
  await uploadToMonday(itemId, stateFileName, blob, "application/json");

  // 2. Write lightweight metadata to the column
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
  const data = await mondayQuery(`query { users { id name } }`);
  const users = data?.users ?? [];
  const match = users.find(u => u.name.toLowerCase() === name.toLowerCase());
  return match ? match.id : null;
}

// ─────────────────────────────────────────────
// Agent state: stored as JSON in the Agent State long-text column
// ─────────────────────────────────────────────
const AGENT_STATE_COLUMN = "long_text_mm4b1t9h"; // Agent State column

async function readAgentState(itemId) {
  const data = await mondayQuery(`
    query GetState($itemId: ID!) {
      items(ids: [$itemId]) {
        column_values(ids: ["${AGENT_STATE_COLUMN}"]) { text value }
      }
    }
  `, { itemId });
  const raw = data?.items?.[0]?.column_values?.[0]?.text || "";
  if (!raw) return { revision: 0, history: [], currentHtml: "" };
  try { return JSON.parse(raw); }
  catch { return { revision: 0, history: [], currentHtml: "" }; }
}

async function writeAgentState(itemId, state) {
  const json = JSON.stringify(state);
  // Long-text columns expect a value of the form {"text":"..."}.
  // The GraphQL $val variable must itself be a JSON-encoded string of that object.
  const value = JSON.stringify({ text: json });
  await mondayQuery(`
    mutation SetState($itemId: ID!, $boardId: ID!, $val: JSON!) {
      change_column_value(item_id: $itemId, board_id: $boardId, column_id: "${AGENT_STATE_COLUMN}", value: $val) { id }
    }
  `, { itemId, boardId: BOARD_ID, val: value });
}

// Post an update (comment) on an item, optionally tagging a user by ID
async function postUpdate(itemId, body) {
  await mondayQuery(`
    mutation PostUpdate($itemId: ID!, $body: String!) {
      create_update(item_id: $itemId, body: $body) { id }
    }
  `, { itemId, body });
}

// Look up a user's ID by name (for @-mention); returns null if not found
async function findUserIdByName(name) {
  if (!name) return null;
  const data = await mondayQuery(`query { users { id name } }`);
  const users = data?.users ?? [];
  const match = users.find(u => u.name.toLowerCase() === name.toLowerCase());
  return match ? match.id : null;
}

// ═════════════════════════════════════════════
// Brand-standard button HTML (Outlook-safe, table-based)
// ═════════════════════════════════════════════
function buildButton(text, url) {
  return `<table border="0" cellpadding="0" cellspacing="0" width="100%" role="presentation"><tbody><tr><td align="center" style="padding:20px 0;"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:separate!important;" width="300"><tbody><tr><td align="center" style="background-color:#72246C;border-radius:3px;font-family:Arial,Helvetica,sans-serif;font-size:15px;padding:12px 20px;"><a href="${url}" target="_blank" style="color:#FFFFFF;text-decoration:none;display:inline-block;font-weight:bold;">${text}</a></td></tr></tbody></table></td></tr></tbody></table>`;
}

// Convert marketer-friendly button syntax into brand button HTML
// Recognizes:
//   Button: Register Now
//   Link: https://cpa.com/register
// (the two lines can be in either order, and are matched as a pair)
function parseButtons(html) {
  // Match "Button: <text>" followed within a few lines by "Link: <url>"
  const pattern = /Button:\s*(.+?)\s*[\r\n]+\s*Link:\s*(\S+)/gi;
  let out = html.replace(pattern, (_, text, url) => buildButton(text.trim(), url.trim()));

  // Also handle reverse order: "Link:" first, then "Button:"
  const reversePattern = /Link:\s*(\S+)\s*[\r\n]+\s*Button:\s*(.+?)(?=[\r\n]|$)/gi;
  out = out.replace(reversePattern, (_, url, text) => buildButton(text.trim(), url.trim()));

  return out;
}

// ═════════════════════════════════════════════
// Instructions parsing + variable substitution
// ═════════════════════════════════════════════
function parseInstructions(instructions) {
  const vars = {};
  if (!instructions) return vars;

  const VARIABLE_NAMES = ["PreheaderText","BodyText"];

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
  for (const key of CONTENT_VARIABLES) {
    if (vars[key] !== undefined) {
      out = out.replace(new RegExp(`\\{\\{${key}\\}\\}`, "g"), vars[key]);
    }
  }
  if (vars["BodyText"])  out = out.replace(/BODY_CONTENT_HERE/g, vars["BodyText"]);
  if (vars["JobNumber"]) out = out.replace(/JOB_NUMBER_HERE/g, vars["JobNumber"]).replace(/JOB_NUMBER(?!_)/g, vars["JobNumber"]);

  // Convert any [[BUTTON: text | url]] tokens into brand button HTML
  out = parseButtons(out);
  return out;
}

// ═════════════════════════════════════════════
// Header image: fetch from ticket Files, re-host on SharePoint
// ═════════════════════════════════════════════

// The folder item ID in SharePoint where header images get re-hosted.
// Set HEADER_UPLOAD_FOLDER_ID in env (a folder in the same drive).
async function fetchTicketFiles(itemId) {
  const data = await mondayQuery(`
    query GetFiles($itemId: ID!) {
      items(ids: [$itemId]) {
        assets { id name url public_url file_extension }
      }
    }
  `, { itemId });
  return data?.items?.[0]?.assets ?? [];
}

// Pick the first image asset that is NOT a generated .html proof
function pickHeaderImage(assets) {
  const IMAGE_EXT = ["png", "jpg", "jpeg", "gif"];
  return assets.find(a => {
    const ext = (a.file_extension || a.name.split(".").pop() || "").toLowerCase();
    return IMAGE_EXT.includes(ext);
  });
}

// Download bytes from a Monday asset (public_url works without auth; url needs token)
async function downloadMondayAsset(asset) {
  const src = asset.public_url || asset.url;
  const res = await fetch(src);
  if (!res.ok) throw new Error(`Asset download failed: ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// Resolve the upload folder's item ID. Accepts either a raw item ID
// (HEADER_UPLOAD_FOLDER_ID) or a SharePoint share URL (HEADER_UPLOAD_FOLDER_URL).
let cachedFolderId = null;
async function getHeaderFolderId() {
  if (cachedFolderId) return cachedFolderId;
  if (process.env.HEADER_UPLOAD_FOLDER_ID) {
    cachedFolderId = process.env.HEADER_UPLOAD_FOLDER_ID;
    return cachedFolderId;
  }

  const shareUrl = process.env.HEADER_UPLOAD_FOLDER_URL;
  if (!shareUrl) throw new Error("No HEADER_UPLOAD_FOLDER_ID or HEADER_UPLOAD_FOLDER_URL set.");

  // Encode share URL per Microsoft Graph sharing URL spec
  const b64 = Buffer.from(shareUrl).toString("base64")
    .replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const shareId = "u!" + b64;

  const token = await getSharePointToken();
  const res   = await fetch(
    `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Folder resolve failed: ${res.status}`);
  const item = await res.json();
  cachedFolderId = item.id;
  console.log(`Resolved header upload folder ID: ${cachedFolderId}`);
  return cachedFolderId;
}

// Upload bytes to SharePoint and create an anonymous (public) share link
async function rehostToSharePoint(fileName, buffer) {
  const token  = await getSharePointToken();
  const folder = await getHeaderFolderId();

  // Upload (simple PUT for files under 4MB)
  const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${folder}:/${encodeURIComponent(fileName)}:/content`;
  const upRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
    body: buffer,
  });
  if (!upRes.ok) throw new Error(`SharePoint upload failed: ${upRes.status}`);
  const uploaded = await upRes.json();

  // Create an anonymous view link
  const linkRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${uploaded.id}/createLink`,
    {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ type: "view", scope: "anonymous" }),
    }
  );
  if (!linkRes.ok) throw new Error(`Share link creation failed: ${linkRes.status}`);
  const linkData = await linkRes.json();

  // Convert the share URL to a direct-download form usable in <img src>
  let webUrl = linkData?.link?.webUrl || "";
  if (webUrl && !webUrl.includes("download=1")) {
    webUrl += (webUrl.includes("?") ? "&" : "?") + "download=1";
  }
  return webUrl;
}

// Replace the first <img src="..."> in the HTML with a new URL
function replaceFirstImage(html, newUrl) {
  let replaced = false;
  return html.replace(/(<img\b[^>]*\bsrc=")([^"]*)(")/i, (m, pre, _src, post) => {
    if (replaced) return m;
    replaced = true;
    return pre + newUrl + post;
  });
}

// Full header-swap flow; returns possibly-modified HTML
async function applyHeaderImage(html, itemId) {
  try {
    const assets = await fetchTicketFiles(itemId);
    const header = pickHeaderImage(assets);
    if (!header) return html; // no image attached — keep default

    const buf       = await downloadMondayAsset(header);
    const publicUrl = await rehostToSharePoint(`header_${itemId}_${header.name}`, buf);
    if (!publicUrl) return html;

    console.log(`Header image re-hosted: ${publicUrl}`);
    return replaceFirstImage(html, publicUrl);
  } catch (err) {
    console.warn(`Header image swap skipped for item ${itemId}: ${err.message}`);
    return html; // fail safe — keep default header
  }
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

// ═════════════════════════════════════════════
// Generate HTML
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
  if (/BUTTON_TEXT/.test(html))           stillMissing.push("ButtonText");

  if (stillMissing.length > 0) {
    const message = await anthropic.messages.create({
      model:      "claude-haiku-4-5-20251001",
      max_tokens: 8000,
      system:     AI_SYSTEM_PROMPT,
      messages: [{
        role:    "user",
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

  // Swap header image if one is attached to the ticket's Files column
  if (ticket.id && ticket.id !== "manual") {
    html = await applyHeaderImage(html, ticket.id);
  }

  return html;
}

// ═════════════════════════════════════════════
// Agent: revise an existing draft based on requestor feedback
// ═════════════════════════════════════════════
async function reviseHTML(currentHtml, feedback, history) {
  const originalFooter = extractFooter(currentHtml);

  const historyText = history.length
    ? history.map((h, i) => `Round ${i + 1}: ${h}`).join("\n")
    : "(none)";

  const message = await anthropic.messages.create({
    model:      "claude-haiku-4-5-20251001",
    max_tokens: 8000,
    system:     AI_SYSTEM_PROMPT,
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

  if (originalFooter && !hasFooter(html)) {
    html = reattachFooter(html, originalFooter);
  }
  return html;
}


// ═════════════════════════════════════════════
// Upload to Monday (content can be HTML or JSON)
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
    method:  "POST",
    headers: { Authorization: process.env.MONDAY_API_TOKEN, "API-Version": "2024-01", ...form.getHeaders() },
    body:    form,
  });
  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join(", "));
  return data?.data?.add_file_to_column;
}

// ═════════════════════════════════════════════
// ROUTES
// ═════════════════════════════════════════════
app.get("/",       (_, res) => res.json({ status: "ok", service: "Monday Email Generator" }));
app.get("/health", (_, res) => res.json({ status: "ok" }));

// Debug: tests the exact metadata + download flow generate uses
app.get("/debug/sharepoint", async (req, res) => {
  try {
    const token = await getSharePointToken();
    const firstItemId = Object.values(TEMPLATE_MAP)[0];

    const metaRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${firstItemId}?select=id,name,@microsoft.graph.downloadUrl`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const metaData = await metaRes.json();

    let downloadStatus = null, downloadOk = null, contentPreview = null;
    const downloadUrl = metaData["@microsoft.graph.downloadUrl"];
    if (downloadUrl) {
      const fileRes = await fetch(downloadUrl);
      downloadStatus = fileRes.status;
      downloadOk = fileRes.ok;
      if (fileRes.ok) contentPreview = (await fileRes.text()).substring(0, 200);
    }

    res.json({
      tokenObtained: true,
      driveId: process.env.SHAREPOINT_DRIVE_ID ? "set" : "MISSING",
      firstItemId,
      metadataStatus: metaRes.status,
      metadataOk: metaRes.ok,
      metadataResponse: metaRes.ok ? { id: metaData.id, name: metaData.name } : metaData,
      hasDownloadUrl: !!downloadUrl,
      downloadStatus,
      downloadOk,
      contentPreview,
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
                  "status_1","dropdown2","person","files","long_text_mm47njms"
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

// POST /api/generate
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

// POST /api/upload
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

// POST /api/webhook — Monday agent: handles item creation, feedback, approval
app.post("/api/webhook", async (req, res) => {
  const { challenge, event } = req.body;

  // Monday setup handshake
  if (challenge) return res.json({ challenge });
  if (!event) return res.json({ status: "ignored" });

  // Respond immediately so Monday doesn't retry; process async
  res.json({ status: "received" });

  try {
    // ── Trigger 1: new item created → generate first proof ──
    if (event.type === "create_pulse") {
      const itemId = String(event.pulseId);
      console.log(`[agent] New item ${itemId} — generating first proof`);

      const ticket       = await fetchItemById(itemId);
      const templateName = "CPACOM General"; // default until Template column wired
      const html         = await generateHTML(ticket, templateName);
      const fileName     = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v1.html`;

      // Upload the proof, THEN persist state — and only post "ready" after both succeed.
      await uploadToMonday(itemId, fileName, html);
      await persistAgentState(itemId, 1, [], html);

      // State is guaranteed in place before the requestor is notified — no race.
      const requestorId = await findUserIdByName(ticket.requestor);
      const mention = requestorId ? `[@${ticket.requestor}](${requestorId}) ` : "";
      await postUpdate(itemId,
        `<p>${mention}Your email proof (v1) is ready and attached to this item's Files. ` +
        `Review it and reply with <strong>@agent</strong> followed by any changes you'd like. ` +
        `When it's ready, set Status to <strong>Approved</strong>.</p>`
      );
      console.log(`[agent] Item ${itemId} v1 complete`);
      return;
    }

    // ── Trigger 2: update posted → feedback if it starts with @agent ──
    if (event.type === "create_update") {
      const itemId = String(event.pulseId);
      const body   = (event.body || event.textBody || "").trim();

      // Strip HTML tags for keyword detection
      const plain = body.replace(/<[^>]+>/g, "").trim();
      if (!/^@agent\b/i.test(plain)) {
        console.log(`[agent] Update on ${itemId} ignored (no @agent prefix)`);
        return;
      }

      const feedback = plain.replace(/^@agent\b[:,\s]*/i, "").trim();
      if (!feedback) return;

      console.log(`[agent] Feedback on ${itemId}: "${feedback}"`);

      const ticket = await fetchItemById(itemId);

      // Don't act on approved items
      if ((ticket.status || "").toLowerCase() === "approved") {
        console.log(`[agent] Item ${itemId} already approved — ignoring feedback`);
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

      // Upload proof + persist state BEFORE notifying — no race.
      await uploadToMonday(itemId, fileName, revised);
      await persistAgentState(itemId, newRev, [...(meta.history || []), feedback], revised);

      const requestorId = await findUserIdByName(ticket.requestor);
      const mention = requestorId ? `[@${ticket.requestor}](${requestorId}) ` : "";
      await postUpdate(itemId,
        `<p>${mention}Updated proof (v${newRev}) is attached with your requested changes. ` +
        `Reply with <strong>@agent</strong> for more edits, or set Status to <strong>Approved</strong> when ready.</p>`
      );
      console.log(`[agent] Item ${itemId} revised to v${newRev}`);
      return;
    }

    // ── Trigger 3: status changed → if Approved, agent goes silent ──
    if (event.type === "update_column_value" && event.columnId === "status") {
      const itemId = String(event.pulseId);
      const label  = event.value?.label?.text || event.value?.label || "";
      if ((label || "").toLowerCase() === "approved") {
        console.log(`[agent] Item ${itemId} approved — finalizing, no further action`);
        // Silent per spec — nothing to post or change
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
