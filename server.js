/**
 * Monday.com Email Template Generator — Full Stack Server
 * Serves frontend UI + API + Monday webhook listener
 *
 * Required env vars:
 *   MONDAY_API_TOKEN   — Monday.com Personal API Token
 *   ANTHROPIC_API_KEY  — Anthropic API Key
 *   BOARD_ID           — Monday.com Board ID
 *   MONDAY_SIGNING_SECRET — From Monday app settings (for webhook verification)
 */

const express   = require("express");
const cors      = require("cors");
const fetch     = require("node-fetch").default || require("node-fetch");
const FormData  = require("form-data");
const Anthropic = require("@anthropic-ai/sdk").default;
const crypto    = require("crypto");
const path      = require("path");

const app = express();
app.use(cors({ origin: "*", methods: ["GET", "POST", "OPTIONS"], allowedHeaders: ["Content-Type", "Authorization"] }));
app.options("*", cors());
app.use(express.json({ limit: "10mb" }));

const MONDAY_API_URL  = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID        = process.env.BOARD_ID;
const anthropic       = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ─────────────────────────────────────────────
// Templates
// ─────────────────────────────────────────────
const SHAREPOINT_DRIVE_ID = process.env.SHAREPOINT_DRIVE_ID;

// Map template names to SharePoint file item IDs
const TEMPLATE_MAP = {
  "AICPA Town Hall Newsletter": "0135ZG5SYRC5GCY6GDPFDJVT3R37W5MFI7",
  "DOTCPA General":             "0135ZG5S4OE2ZR2WA4BZELNKR6CP5ISLZ3",
  "CPACOM General":             "0135ZG5S36KVRZLIB3SFEL3RGEYDHPMOP3",
};

// ─────────────────────────────────────────────
// SharePoint: client credentials token fetch
// ─────────────────────────────────────────────
let sharepointToken = null;
let tokenExpiresAt  = 0;

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
const CACHE_TTL_MS  = 10 * 60 * 1000; // 10 minutes

// ─────────────────────────────────────────────
// Helper: fetch template HTML from SharePoint
// ─────────────────────────────────────────────
async function fetchTemplateFromSharePoint(templateName) {
  const itemId = TEMPLATE_MAP[templateName];
  if (!itemId) throw new Error(`Unknown template: "${templateName}"`);

  // Return cached version if still fresh
  const cached = templateCache[itemId];
  if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) {
    return cached.html;
  }

  const url = `https://graph.microsoft.com/v1.0/drives/${SHAREPOINT_DRIVE_ID}/items/${itemId}/content`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${process.env.SHAREPOINT_ACCESS_TOKEN}` }
  });

  if (!res.ok) throw new Error(`SharePoint fetch failed: ${res.status} ${res.statusText}`);

  const html = await res.text();
  templateCache[itemId] = { html, fetchedAt: Date.now() };
  console.log(`Fetched template "${templateName}" from SharePoint (${html.length} bytes)`);
  return html;
}

const AI_SYSTEM_PROMPT = `You are an email template generator. Populate the HTML template with ticket data.
RULES:
- Return ONLY the populated HTML. No explanation, no markdown, no code fences.
- Keep ALL HTML structure, styles, images, and links completely intact.
- CRITICAL: Any existing {{...}} or {{{...}}} tokens in the template are Pardot merge tags. Leave them exactly as-is. Never modify, replace, or remove them.
- CRITICAL: The HTML section between <!-- START FOOTER --> and <!-- END FOOTER --> is the footer. It must ALWAYS be present and completely intact in your output. Never remove, truncate, or modify it under any circumstances.
- Only replace these plain text placeholders:
    BODY_CONTENT_HERE → email body copy written from the ticket description
    BUTTON_TEXT       → a relevant CTA based on ticket context
    JOB_NUMBER_HERE   → the job number, or remove the line if blank
- Write body copy in a professional tone appropriate to the template and product.
- If the description says TBD or TBC, write appropriate placeholder copy based on the subject line and product.`;

// ─────────────────────────────────────────────
// Helper: Monday GraphQL
// ─────────────────────────────────────────────
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

// ─────────────────────────────────────────────
// Helper: normalize Monday column_values
// ─────────────────────────────────────────────
function normalizeTicket(item) {
  const cols = {};
  for (const cv of item.column_values) cols[cv.id] = cv.text ?? cv.value ?? "";
  return {
    id:          item.id,
    name:        item.name,
    status:      cols["status"]        || "",
    description: cols["long_text7"]    || "",
    subjectLine: cols["text86"]        || "",
    jobNumber:   cols["formula"]       || "",
    sendDate:    cols["date4"]         || "",
    contentDue:  cols["date_mkx4g1zc"] || "",
    type:        cols["dropdown3"]     || "",
    category:    cols["status_1"]      || "",
    product:     cols["dropdown2"]     || "",
    requestor:   cols["person"]        || "",
    instructions: cols["long_text_mm47njms"] || "",  // Instructions column
    hasFiles:    !!cols["files"],
  };
}

// ─────────────────────────────────────────────
// Helper: fetch a single item by ID
// ─────────────────────────────────────────────
async function fetchItemById(itemId) {
  const data = await mondayQuery(`
    query GetItem($itemId: ID!) {
      items(ids: [$itemId]) {
        id name
        column_values(ids: [
          "status","long_text7","text86","formula",
          "date4","date_mkx4g1zc","dropdown3",
          "status_1","dropdown2","person","files","long_text_mm47njms"
        ]) { id text value }
      }
    }
  `, { itemId });
  const item = data?.items?.[0];
  if (!item) throw new Error(`Item ${itemId} not found.`);
  return normalizeTicket(item);
}

// ─────────────────────────────────────────────
// Helper: parse Instructions column into a variable map
// Format supported:
//   PreheaderText: "some value"
//   PreheaderText: some value (no quotes)
//   PreheaderText:
//   "some value"   (value on next line)
// ─────────────────────────────────────────────
function parseInstructions(instructions) {
  const vars = {};
  if (!instructions) return vars;

  const VARIABLE_NAMES = [
    "PreheaderText","BodyText","PrimaryLink","PrimaryText",
    "SecondaryLink","SecondText","TertiaryLink","TertiaryText"
  ];

  // Build a regex that matches "VariableName: value" or "VariableName:\n value"
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

// ─────────────────────────────────────────────
// Helper: replace {{Variable}} tokens in HTML
// Only replaces tokens that are NOT Pardot tags
// (Pardot tags are {{Recipient.*}}, {{{...}}}, etc.)
// ─────────────────────────────────────────────
const CONTENT_VARIABLES = [
  "PreheaderText","BodyText","PrimaryLink","PrimaryText",
  "SecondaryLink","SecondText","TertiaryLink","TertiaryText",
  "Subject","JobNumber"
];

function applyVariables(html, vars) {
  let out = html;
  for (const key of CONTENT_VARIABLES) {
    if (vars[key] !== undefined) {
      const regex = new RegExp(`\\{\\{${key}\\}\\}`, "g");
      out = out.replace(regex, vars[key]);
    }
  }
  // Also replace legacy plain-text placeholders
  if (vars["BodyText"])   out = out.replace(/BODY_CONTENT_HERE/g, vars["BodyText"]);
  if (vars["JobNumber"])  out = out.replace(/JOB_NUMBER_HERE/g, vars["JobNumber"])
                                   .replace(/JOB_NUMBER(?!_)/g, vars["JobNumber"]);
  return out;
}
function extractFooter(html) {
  const match = html.match(/<!--\s*START FOOTER\s*-->([\s\S]*?)<!--\s*END FOOTER\s*-->/i);
  return match ? match[0] : null;
}

function hasFooter(html) {
  return /<!--\s*START FOOTER\s*-->/i.test(html) && /<!--\s*END FOOTER\s*-->/i.test(html);
}

function reattachFooter(html, footer) {
  // Insert footer before closing </table> of the main wrapper
  const insertPoint = html.lastIndexOf("</tbody></table></div>");
  if (insertPoint !== -1) {
    return html.slice(0, insertPoint) + "\n" + footer + "\n" + html.slice(insertPoint);
  }
  // Fallback: insert before </body>
  return html.replace("</body>", footer + "\n</body>");
}

// ─────────────────────────────────────────────
// Helper: generate HTML via Claude
// ─────────────────────────────────────────────
async function generateHTML(ticket, templateName) {
  // Fetch template from SharePoint (cached after first fetch)
  const templateHtml = await fetchTemplateFromSharePoint(templateName);
  if (!templateHtml) throw new Error(`Could not load template: "${templateName}"`);

  const originalFooter = extractFooter(templateHtml);

  // Step 1: Parse Instructions column into variable map
  const vars = parseInstructions(ticket.instructions);

  // Step 2: Inject Subject and JobNumber from their own columns
  vars["Subject"]    = ticket.subjectLine || ticket.name || "";
  vars["JobNumber"]  = ticket.jobNumber || "";

  // Step 3: Apply all known variables directly — no AI needed for these
  let html = applyVariables(templateHtml, vars);

  // Step 4: Check which content placeholders still need filling
  const stillMissing = [];
  if (/BODY_CONTENT_HERE/.test(html) || /\{\{BodyText\}\}/.test(html)) stillMissing.push("BodyText");
  if (/\{\{PreheaderText\}\}/.test(html)) stillMissing.push("PreheaderText");
  if (/BUTTON_TEXT/.test(html))           stillMissing.push("ButtonText");

  // Step 5: Only call Claude if there are unfilled placeholders
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

  // Step 6: Footer enforcement
  if (originalFooter && !hasFooter(html)) {
    console.warn(`Footer missing for "${ticket.name}" — reattaching.`);
    html = reattachFooter(html, originalFooter);
  }

  return html;
}

// ─────────────────────────────────────────────
// Helper: upload HTML file to Monday item
// ─────────────────────────────────────────────
async function uploadToMonday(itemId, fileName, html) {
  const query = `
    mutation AddFileToColumn($itemId: ID!, $columnId: String!, $file: File!) {
      add_file_to_column(item_id: $itemId, column_id: $columnId, file: $file) {
        id name url
      }
    }
  `;

  const form = new FormData();
  form.append("query", query);
  form.append("variables", JSON.stringify({ itemId, columnId: "files", file: null }));
  form.append("map", JSON.stringify({ "0": ["variables.file"] }));

  const buf = Buffer.from(html, "utf-8");
  form.append("0", buf, { filename: fileName, contentType: "text/html", knownLength: buf.length });

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

// Root + health
app.get("/",       (_, res) => res.json({ status: "ok", service: "Monday Email Generator" }));
app.get("/health", (_, res) => res.json({ status: "ok" }));

// GET /api/tickets — fetch all groups, filter for New Requests
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

// POST /api/generate — generate HTML from ticket + template
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

// POST /api/upload — upload HTML file to Monday ticket
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

// POST /api/webhook — Monday fires this when a new item is created
app.post("/api/webhook", async (req, res) => {
  const { challenge, event } = req.body;

  // Monday sends a challenge on first setup — must echo it back
  if (challenge) return res.json({ challenge });

  // Only handle item_created events
  if (event?.type !== "create_pulse") {
    return res.json({ status: "ignored" });
  }

  const itemId = String(event.pulseId);
  console.log(`Webhook: new item created — ID ${itemId}`);

  // Respond immediately so Monday doesn't retry
  res.json({ status: "received" });

  // Process asynchronously
  try {
    const ticket = await fetchItemById(itemId);
    console.log(`Processing ticket: ${ticket.name}`);

    // Pick template — from Template column, fallback to CPACOM General
    const templateName = TEMPLATES[ticket.template] ? ticket.template : "CPACOM General";
    console.log(`Using template: ${templateName}`);

    const html     = await generateHTML(ticket, templateName);
    const fileName = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}.html`;
    const asset    = await uploadToMonday(itemId, fileName, html);

    console.log(`✅ Uploaded ${fileName} to item ${itemId} — asset ID: ${asset?.id}`);
  } catch (err) {
    console.error(`Webhook processing error for item ${itemId}:`, err.message);
  }
});

// ─────────────────────────────────────────────
// Serve frontend UI
// ─────────────────────────────────────────────
app.get("/app", (_, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));
