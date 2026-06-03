/**
 * Monday.com Email Template Generator — Backend Server
 * 
 * Stack:  Node.js + Express
 * Deploy: Render / Railway / Vercel / any Node host
 * 
 * Required env vars:
 *   MONDAY_API_TOKEN   — Monday.com API token (Admin → API)
 *   ANTHROPIC_API_KEY  — Anthropic API key
 *   PORT               — (optional) defaults to 3000
 *   ALLOWED_ORIGIN     — (optional) CORS origin, e.g. https://claude.ai
 */

import express from "express";
import cors from "cors";
import fetch from "node-fetch";
import FormData from "form-data";
import Anthropic from "@anthropic-ai/sdk";

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(cors({ origin: process.env.ALLOWED_ORIGIN || "*" }));

const MONDAY_API_URL = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID       = "2120641399";

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ─────────────────────────────────────────────
// Helper: call Monday GraphQL API
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
// GET /api/tickets — fetch New Requests group
// ─────────────────────────────────────────────
app.get("/api/tickets", async (req, res) => {
  try {
    const data = await mondayQuery(`
      query GetNewRequests($boardId: ID!) {
        boards(ids: [$boardId]) {
          groups {
            id title
            items_page(limit: 50) {
              items {
                id name
                column_values(ids: [
                  "status", "long_text7", "text86", "formula",
                  "date4", "date_mkx4g1zc", "dropdown3", "status_1",
                  "dropdown2", "person", "files"
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

    // Match any group whose title contains "new request" (case-insensitive)
    // Falls back to all groups if none found, so the app still works
    const matchedGroups = allGroups.filter(g =>
      g.title.toLowerCase().includes("new request")
    );
    const groups = matchedGroups.length > 0 ? matchedGroups : allGroups;
    const items  = groups.flatMap(g => g.items_page?.items ?? []);

    // Normalize column_values array → flat object
    const tickets = items.map(item => {
      const cols = {};
      for (const cv of item.column_values) cols[cv.id] = cv.text ?? cv.value;
      return {
        id:          item.id,
        name:        item.name,
        status:      cols["status"]           ?? "",
        description: cols["long_text7"]        ?? "",
        subjectLine: cols["text86"]            ?? "",
        jobNumber:   cols["formula"]           ?? "",
        sendDate:    cols["date4"]             ?? "",
        contentDue:  cols["date_mkx4g1zc"]    ?? "",
        type:        cols["dropdown3"]         ?? "",
        category:    cols["status_1"]          ?? "",
        product:     cols["dropdown2"]         ?? "",
        requestor:   cols["person"]            ?? "",
        hasFiles:    cols["files"] !== null && cols["files"] !== "",
      };
    });

    res.json({ tickets });
  } catch (err) {
    console.error("GET /api/tickets error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/generate — generate HTML via Claude
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
- CRITICAL: Any existing {{...}} or {{{...}}} tokens already present in the template are Pardot merge tags and must be left exactly as-is. Do NOT modify, replace, or remove them under any circumstances. Examples: {{Recipient.FirstName}}, {{{EmailPreferenceCenter_488}}}, {{Recipient.Custom_CPA_Domain}}.
- Only replace these specific content placeholders that you insert yourself — do not treat them as Pardot tags:
    BODY_CONTENT_HERE  → replace with email body copy written from the ticket description
    BUTTON_TEXT        → replace with a relevant CTA based on the ticket context
    JOB_NUMBER_HERE    → replace with the job number, or remove the line if blank
- Write body copy in a professional tone appropriate to the template and product.
- If the description says "TBD" or "TBC", write placeholder-appropriate copy based on the subject line and product.`,
      messages: [{
        role: "user",
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
    if (!html) throw new Error("Claude returned no content.");
    res.json({ html });
  } catch (err) {
    console.error("POST /api/generate error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/upload — upload HTML file to Monday
// ─────────────────────────────────────────────
app.post("/api/upload", async (req, res) => {
  const { itemId, fileName, html } = req.body;
  if (!itemId || !fileName || !html) {
    return res.status(400).json({ error: "Missing itemId, fileName, or html" });
  }

  try {
    // Monday's file upload requires a multipart/form-data GraphQL request
    const query = `
      mutation AddFileToColumn($itemId: ID!, $columnId: String!, $file: File!) {
        add_file_to_column(item_id: $itemId, column_id: $columnId, file: $file) {
          id
          name
          url
        }
      }
    `;

    const variables = {
      itemId,
      columnId: "files",
      file: null,  // placeholder — actual file goes in form-data map
    };

    // Build multipart form-data per the Monday file upload spec
    // https://developer.monday.com/api-reference/docs/files
    const form = new FormData();

    form.append("query",     query);
    form.append("variables", JSON.stringify(variables));

    // "map" maps the multipart field "file" to the GraphQL variable path
    form.append("map", JSON.stringify({ "0": ["variables.file"] }));

    // Append the actual HTML file buffer
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
    res.json({
      success:  true,
      assetId:  asset?.id,
      fileName: asset?.name,
      url:      asset?.url,
    });
  } catch (err) {
    console.error("POST /api/upload error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─────────────────────────────────────────────
// Health check
// ─────────────────────────────────────────────
app.get("/health", (_, res) => res.json({ status: "ok" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));
