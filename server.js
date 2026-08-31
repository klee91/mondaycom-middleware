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

const { BOARD_ID, INSTRUCTIONS_COLUMN, AGENT_STATE_COLUMN, TEMPLATE_COLUMN } = require("./helpers/config");
const { getSharePointToken, fetchTemplateIndex } = require("./helpers/sharepoint");
const { getManifest } = require("./helpers/github-store");
const { designVariant } = require("./helpers/design");
const { parseInstructions } = require("./helpers/instructions");
const {
  mondayQuery, normalizeTicket, fetchItemById, fetchTicketFiles, fetchWordDocContent,
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
                  "status_1","dropdown2","person","files","${INSTRUCTIONS_COLUMN}","${AGENT_STATE_COLUMN}","${TEMPLATE_COLUMN}"
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
app.post("/api/webhook", async (req, res) => {
  const { challenge, event } = req.body;
  if (challenge) return res.json({ challenge });
  if (!event) return res.json({ status: "ignored" });

  res.json({ status: "received" });
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

      // At ticket creation, only auto-run if a Word doc is already attached.
      // Requestors often create placeholder tickets before content is ready, so
      // prompt text alone (which may just be a placeholder) is NOT enough to
      // start generation. If there's no doc, do nothing and wait — the
      // requestor kicks it off later by tagging @MEG. Check this FIRST so empty
      // placeholder tickets bail out immediately (no job-number retry delay).
      const hasDoc = await fetchWordDocContent(itemId);
      if (!hasDoc) {
        console.log(`[agent] Item ${itemId} created without a Word doc — waiting for a doc or an @MEG tag before generating.`);
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
      const body   = (event.body || event.textBody || "").trim();
      const plain  = body.replace(/<[^>]+>/g, "").trim();

      if (!/^@(MEG|agent)\b/i.test(plain)) {
        console.log(`[agent] Update on ${itemId} ignored (no @MEG prefix)`);
        return;
      }
      const feedback = plain.replace(/^@(MEG|agent)\b[:,\s]*/i, "").trim();
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
        console.log(`[agent] No prior draft on ${itemId} — generating first proof from @MEG request`);
        const templateName = await resolveTemplateName(ticket);
        const html         = await generateHTML(ticket, templateName);
        const fileName     = `${ticket.name.replace(/\s+/g, "_")}_v1.html`;

        const asset = await uploadToMonday(itemId, fileName, html);
        await persistAgentState(itemId, 1, [], html);
        await setStatus(itemId, "Proofing");
        await recordProofVersion(itemId, { version: 1, fileName, assetUrl: asset?.url });

        const questions = await analyzeForQuestions(ticket, parseInstructions(ticket.instructions), html);
        const reqId    = await findUserIdByName(ticket.requestor);
        const reqName  = ticket.requestor.split(",")[0].trim();
        const reqTag   = reqId ? `<strong>@${reqName}</strong> ` : "";
        await postUpdate(itemId,
          `<p>${reqTag}Your email proof (v1) is ready and attached to this item's Files. ` +
          `Review it and reply with <strong>@MEG</strong> followed by any changes you'd like. ` +
          `When it's ready, set Status to <strong>Approved</strong>.</p>` +
          renderQuestionsBlock(questions),
          reqId ? [reqId] : []
        );
        console.log(`[agent] Item ${itemId} v1 complete (via @MEG)${questions.length ? ` (${questions.length} question(s) posted)` : ""}`);
        return;
      }

      const ticketVars = parseInstructions(ticket.instructions);
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

app.get("/app", (_, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅  Server running on port ${PORT}`));