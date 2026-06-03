/**
 * Monday.com Email Template Generator — Backend Server
 * Node.js + Express (CommonJS)
 */

const express    = require("express");
const cors       = require("cors");
const fetch      = require("node-fetch");
const FormData   = require("form-data");
const Anthropic  = require("@anthropic-ai/sdk").default;

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(cors({
  origin: "*",
  methods: ["GET", "POST", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"],
}));
app.options("*", cors());

const MONDAY_API_URL = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID        = process.env.BOARD_ID;

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

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
// Root + health
// ─────────────────────────────────────────────
app.get("/", (_, res) => res.json({ status: "ok", service: "Monday Email Generator API" }));
app.get("/health", (_, res) => res.json({ status: "ok" }));

// ─────────────────────────────────────────────
// GET /api/tickets
// ─────────────────────────────────────────────
app.get("/api/tickets", async (req, res) => {
  try {
    const data = await mondayQuery(`
      query GetGroups($boardId: ID!) {
        boards(ids: [$boardId]) {
          groups {
            id
            title
            items_page(limit: 50) {
              items {
                id
                name
                column_values(ids: [
                  "status", "long_text7", "text86", "formula",
                  "date4", "date_mkx4g1zc", "dropdown3",
                  "status_1", "dropdown2", "person", "files"
                ]) {
                  id text value
                }
              }
            }
          }
        }
      }
    `, { boardId: BOARD_ID });

    const allGroups = data?.boards?.[0]?.groups ?? [];
    const matched   = allGroups.filter(g =>
      g.title.toLowerCase().includes("new request")
    );
    const groups = matched.length > 0 ? matched : allGroups;
    const items  = groups.flatMap(g => g.items_page?.items ?? []);

    const tickets = items.map(item => {
      const cols = {};
      for (const cv of item.column_values) cols[cv.id] = cv.text ?? cv.value;
      return {
        id:          item.id,
        name:        item.name,
        status:      cols["status"]        ?? "",
        description: cols["long_text7"]    ?? "",
        subjectLine: cols["text86"]        ?? "",
        jobNumber:   cols["formula"]       ?? "",
        sendDate:    cols["date4"]         ?? "",
        contentDue:  cols["date_mkx4g1zc"] ?? "",
        type:        cols["dropdown3"]     ?? "",
        category:    cols["status_1"]      ?? "",
        product:     cols["dropdown2"]     ?? "",
        requestor:   cols["person"]        ?? "",
        hasFiles:    !!cols["files"],
      };
    });

    res.json({ tickets });
  } catch (err) {
    console.error("GET /api/tickets error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/generate
// ─────────────────────────────────────────────
app.post("/api/generate", async (req, res) => {
  const { ticket, templateName, templateHtml } = req.body;
  if (!ticket || !templateName || !templateHtml) {
    return res.status(400).json({ error: "Missing ticket, templateName, or templateHtml" });
  }

  try {
    const message = await anthropic.messages.create({
      model:      "claude-haiku-4-5-20251001",
      max_tokens: 8000,
      system: `You are an email template generator. Populate the HTML template with ticket data.
RULES:
- Return ONLY the populated HTML. No explanation, no markdown, no code fences.
- Keep ALL HTML structure, styles, images, and links completely intact.
- CRITICAL: Any existing {{...}} or {{{...}}} tokens already in the template are Pardot merge tags. Leave them exactly as-is. Do NOT modify, replace, or remove them. Examples: {{Recipient.FirstName}}, {{{EmailPreferenceCenter_488}}}, {{Subject}}.
- Only replace these plain text placeholders:
    BODY_CONTENT_HERE → email body copy based on the ticket description
    BUTTON_TEXT       → a relevant CTA based on ticket context
    JOB_NUMBER_HERE   → the job number, or remove the line if blank
- Write body copy in a professional tone appropriate to the template and product.
- If description says TBD or TBC, write placeholder copy based on the subject line and product.`,
      messages: [{
        role:    "user",
        content: `Ticket data:
Name: ${ticket.name}
Template: ${templateName}
Subject Line: ${ticket.subjectLine || ticket.name}
Job Number: ${ticket.jobNumber || ""}
Send Date: ${ticket.sendDate || ""}
Description: ${ticket.description || "No description provided."}
Category: ${ticket.category || ""}
Product: ${ticket.product || ""}

BASE HTML TEMPLATE:
${templateHtml}`,
      }],
    });

    const html = message.content.find(b => b.type === "text")?.text ?? "";
    if (!html) throw new Error("No content returned from Claude.");
    res.json({ html });
  } catch (err) {
    console.error("POST /api/generate error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/upload
// ─────────────────────────────────────────────
app.post("/api/upload", async (req, res) => {
  const { itemId, fileName, html } = req.body;
  if (!itemId || !fileName || !html) {
    return res.status(400).json({ error: "Missing itemId, fileName, or html" });
  }

  try {
    const query = `
      mutation AddFileToColumn($itemId: ID!, $columnId: String!, $file: File!) {
        add_file_to_column(item_id: $itemId, column_id: $columnId, file: $file) {
          id name url
        }
      }
    `;

    const form = new FormData();
    form.append("query", query);
    form.append("variables", JSON.stringify({
      itemId,
      columnId: "files",
      file:     null,
    }));
    form.append("map", JSON.stringify({ "0": ["variables.file"] }));

    const buf = Buffer.from(html, "utf-8");
    form.append("0", buf, {
      filename:    fileName,
      contentType: "text/html",
      knownLength: buf.length,
    });

    const uploadRes = await fetch(MONDAY_FILE_URL, {
      method:  "POST",
      headers: {
        Authorization: process.env.MONDAY_API_TOKEN,
        "API-Version": "2024-01",
        ...form.getHeaders(),
      },
      body: form,
    });

    const uploadData = await uploadRes.json();
    if (uploadData.errors) {
      throw new Error(uploadData.errors.map(e => e.message).join(", "));
    }

    const asset = uploadData?.data?.add_file_to_column;
    res.json({ success: true, assetId: asset?.id, fileName: asset?.name, url: asset?.url });
  } catch (err) {
    console.error("POST /api/upload error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));
