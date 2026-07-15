/**
 * Monday.com Email Template Generator — Full Stack Agent Server
 *
 * Capabilities:
 *   - Serves frontend UI + REST API (manual generation)
 *   - SharePoint template library (Graph client-credentials)
 *   - Instructions-column variable substitution + brand buttons
 *   - Header image re-hosting from Files column
 *   - Word doc (.docx) in Files column → extracted as body source content
 *   - Footer + Pardot-tag protection, brand style enforcement
 *   - Stateful agent webhook: item created → proof; @agent feedback → revise;
 *     Status=Approved → silent finalize
 *   - Generation model: claude-opus-4-6 (swap to claude-haiku-4-5-20251001 once fine-tuned)
 *
 * Required env vars:
 *   MONDAY_API_TOKEN, ANTHROPIC_API_KEY, BOARD_ID
 *   SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_DRIVE_ID
 */

const express   = require("express");
const cors      = require("cors");
const fetch     = require("node-fetch").default || require("node-fetch");
const FormData  = require("form-data");
const Anthropic = require("@anthropic-ai/sdk").default;
const mammoth   = require("mammoth");
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


const CONTENT_VARIABLES = ["PreheaderText", "Subject", "Partner"];

const BUTTON_COLORS = {
  "purple": "#72246c", "light purple": "#86387f", "green": "#48a23f", "navy": "#0f206c",
};

const AI_SYSTEM_PROMPT = `You are an email template generator for CPA.com. Populate the HTML template with ticket data, strictly following the CPA.com brand standards below.

OUTPUT RULES:
- Output ONLY the raw HTML. Absolutely nothing before the first < and nothing after the last >.
- No preamble, no commentary, no sign-off, no markdown fences. The response is the HTML document and nothing else.
- Keep ALL HTML structure, styles, images, and links completely intact.

PROTECTED — NEVER MODIFY:
- Any existing {{...}} or {{{...}}} tokens are Pardot merge tags. Leave them exactly as-is.
- The section between <!-- START FOOTER --> and <!-- END FOOTER --> must remain completely intact.
- ALL generated and body content (copy, buttons, articles, CTAs, etc.) MUST appear ABOVE the <!-- START FOOTER --> marker — between the header and the footer. Never place any content after <!-- END FOOTER -->. The footer is always the last block in the email body.
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
async function fetchTemplateIndex({ includeNonHtml = false } = {}) {
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
    if (includeNonHtml) {
       map[item.name] = item.id;          // raw filename, not display-cased
    } else if (/\.html$/i.test(item.name)) {
       map[fileNameToDisplayName(item.name)] = item.id;
    }
    // if (/\.html$/i.test(item.name)) map[fileNameToDisplayName(item.name)] = item.id;
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

/**
 * ---------------------------------------------------------------------------
 *
 * Protected-block extract/restore.
 *   Generalizes your existing extractFooter()/reattachFooter() pattern to an
 *   ARRAY of fragile regions. Before the template goes to Claude, fragile
 *   markup (AI-in-Focus bracket buttons, per-article PNG header-strip images,
 *   the footer) is lifted out and replaced with opaque tokens the model can
 *   move but not mangle. After generation, the real markup is restored
 *   verbatim, then per-article numbering is stamped on deterministically.
 *
 * Structured-content-with-model-placement generation.
 *   The model receives (a) the source content loosely (doc/BodyContent), and
 *   (b) ONE canonical article block from the template as the structural
 *   contract. Its job: emit N blocks matching that structure, choosing how
 *   many and what to emphasize. Structure is constrained; editorial judgment
 *   is the model's. That judgment is what improves with testing/training.
 *
 * ---------------------------------------------------------------------------
 */

// Which archetypes/templates use the repeating-block generation path.
// Everything else keeps flowing through the existing placeholder-fill path.
const DESIGN_TEMPLATES = {
  "AI IN FOCUS Newsletter": "aiInFocus",
  "AICPA Town Hall Newsletter": "aicpaTownHall",
};

function isDesignTemplate(templateName) {
  const key = (templateName || "").trim().toLowerCase();
  return Object.keys(DESIGN_TEMPLATES).some(k => k.toLowerCase() === key);
}

function designVariant(templateName) {
  const key = (templateName || "").trim().toLowerCase();
  const match = Object.keys(DESIGN_TEMPLATES).find(k => k.toLowerCase() === key);
  return match ? DESIGN_TEMPLATES[match] : null;
}

// ═════════════════════════════════════════════
// Protected-block extraction / restoration
// ═════════════════════════════════════════════
//
// Each rule matches a fragile region by regex. Order among rules is
// irrelevant to correctness (each scans independently); tokens must be unique,
// and restoration is order-independent because each token is a unique
// sentinel, e.g. <!--#PROTECTED:bracketButton:0#-->.
//
// The regexes below are written against the ACTUAL markup in the two
// templates (verified against the files), not idealized markup:
//
//  - bracketButton: the 3-cell 340px table with ai-nl-btn-left.png / #41b6e6
//    middle cell / ai-nl-btn-right.png. Its href carries the external
//    deep-link anchor (#artN) — protected so the model can't mangle it, and
//    renumbered positionally after restore (see renumberArticles).
//  - articleStripImage: the per-article PNG header-strip <img> tags
//    (article-N-top-sX.png) that frame each AI-in-Focus title. Protected so
//    the model can't alter dimensions or invent filenames; the article NUMBER
//    in the path is stamped positionally after restore. NOTE: only the <img>
//    tags are protected — the title TEXT cell that sits between s2 and s3
//    stays editable, so the model still writes each article's title.
//  - footer: same region your extractFooter() already guards, folded in here
//    so ALL protected regions restore through one path.

const PROTECT_RULES = [
  {
    name: "bracketButton",
    // 340px-wide 3-cell table: left PNG bracket, #41b6e6 middle, right PNG bracket
    re: /<table[^>]*width:\s*340px[^>]*>[\s\S]*?ai-nl-btn-left\.png[\s\S]*?ai-nl-btn-right\.png[\s\S]*?<\/table>/gi,
  },
  {
    name: "articleStripImage",
    // Per-article header-strip PNGs: article-N-top-s1..s4.png. One token per
    // <img>. Does not overlap bracketButton (different filenames).
    re: /<img\b[^>]*article-\d+-top-s\d+\.png[^>]*>/gi,
  },
  {
    name: "footer",
    re: /<!--\s*START FOOTER\s*-->[\s\S]*?<!--\s*END FOOTER\s*-->/gi,
  },
];

/**
 * Extract all protected regions, returning cleaned HTML (tokens in place) plus
 * a restore map. Restoration is exact — the model never influences these bytes.
 */
function protectFragileRegions(html) {
  const store = {}; // token -> original markup
  let out = html;

  for (const rule of PROTECT_RULES) {
    let i = 0;
    out = out.replace(rule.re, (match) => {
      const token = `<!--#PROTECTED:${rule.name}:${i}#-->`;
      store[token] = match;
      i++;
      return token;
    });
  }
  return { html: out, store };
}

/**
 * Restore every protected token. If the model dropped a token entirely (e.g.
 * deleted a whole article that had a bracket button), that protected block is
 * simply not reinserted — which is correct: it went with the block the model
 * chose to remove. Tokens the model kept are restored byte-for-byte. When the
 * model DUPLICATED a token (copied the canonical block N times), every copy is
 * restored to the same original — the per-article numbering is corrected
 * afterward by renumberArticles().
 */
function restoreFragileRegions(html, store) {
  let out = html;
  let restored = 0;
  let missing = 0;
  for (const [token, original] of Object.entries(store)) {
    if (out.includes(token)) {
      out = out.split(token).join(original);
      restored++;
    } else {
      missing++;
    }
  }
  return { html: out, restored, missing, total: Object.keys(store).length };
}

/**
 * Post-generation integrity check for protected tokens: catches the failure
 * mode where the model emitted a MALFORMED token (e.g. altered the sentinel)
 * so it neither matches for restoration nor is cleanly absent. Any leftover
 * "#PROTECTED:" fragment that isn't an exact known token is a red flag.
 */
function findCorruptedTokens(html, store) {
  const known = new Set(Object.keys(store));
  const found = html.match(/<!--#PROTECTED:[^#]*#-->/g) || [];
  return found.filter(t => !known.has(t));
}

// ═════════════════════════════════════════════
// Deterministic per-article renumbering (AI-in-Focus only)
// ═════════════════════════════════════════════
//
// The article NUMBER in three coupled places is purely positional:
//   - image paths        article-N-top-sX.png
//   - title cell id       id="artN"
//   - button deep-link    href="...ai-in-focus/05-26#artN"  (inside the
//                         protected bracketButton, restored just above)
//
// The model is told NOT to manage these numbers. After restore, every article
// carries the canonical block's numbers (all "1"). This pass walks the emitted
// articles in document order and stamps the correct sequential number on each.
// This also cleans up the source template's own numbering drift (the original
// had id run art1,art2,art3,art4,art4 and anchors art1..art5).
//
// Delimiter: the <!-- START ARTICLE --> comment, which begins the canonical
// block and is therefore copied at the head of every emitted article. The head
// (before the first article) and tail (footer etc.) contain none of these
// patterns, so they are safe even though the last segment includes the tail.
function renumberArticles(html, variant) {
  if (variant !== "aiInFocus") return { html, count: 0 };

  const parts = html.split(/<!--\s*START ARTICLE\s*-->/i);
  if (parts.length < 2) {
    return { html, count: 0, warning: "no <!-- START ARTICLE --> markers found; numbering left as-is" };
  }

  let rebuilt = parts[0]; // head, untouched
  let n = 0;
  for (let k = 1; k < parts.length; k++) {
    n += 1;
    const seg = parts[k]
      .replace(/article-\d+-top-s(\d+)\.png/gi, `article-${n}-top-s$1.png`)
      .replace(/id="art\d+"/gi, `id="art${n}"`)
      .replace(/#art\d+\b/gi, `#art${n}`);
    rebuilt += "<!-- START ARTICLE -->" + seg;
  }
  return { html: rebuilt, count: n };
}

// ═════════════════════════════════════════════
// canonical article-block extraction
// ═════════════════════════════════════════════
//
// Rather than hand the model the whole template, we give it exactly
// ONE representative article block as the structural contract, plus the region
// boundaries so it knows where blocks go. The model then emits N blocks that
// match that structure. We keep the FIRST block verbatim as the example and
// remove the rest, leaving a single insertion marker.

/**
 * For AI-in-Focus: articles are delimited by <!-- START ARTICLE --> markers
 * with no end marker (the next START ARTICLE or <!-- END MAIN CONTENT --> is
 * the boundary — this quirk is documented in brand-guide sectionMarkers).
 * For Town Hall: repeating highlight rows keyed by the 320x181 highlight-img
 * cell.
 *
 * Returns { canonicalBlock, shellHtml } where shellHtml has all repeating
 * blocks replaced by a single <!--#ARTICLES#--> insertion marker, and
 * canonicalBlock is one representative block (with its own fragile regions
 * still tokenized, since protectFragileRegions ran first).
 */
function extractArticleShells(html, variant) {
  if (variant === "aiInFocus") {
    // Articles are delimited by <!-- START ARTICLE --> with no end marker; the
    // repeating region ends at <!-- END MAIN CONTENT -->. So:
    //   head              = everything before the first START ARTICLE
    //   repeating region  = first START ARTICLE ... up to END MAIN CONTENT
    //   tail              = END MAIN CONTENT ... EOF (includes footer token)
    const marker = /<!--\s*START ARTICLE\s*-->/i;
    const firstIdx = html.search(marker);
    if (firstIdx === -1) return { canonicalBlock: null, shellHtml: html };

    const endMainRe = /<!--\s*END MAIN CONTENT\s*-->/i;
    const endMainIdx = html.search(endMainRe);
    // If END MAIN CONTENT is absent, fall back to the footer token boundary.
    const footerIdx = html.search(/<!--#PROTECTED:footer/i);
    const regionEnd = endMainIdx !== -1
      ? endMainIdx
      : (footerIdx !== -1 ? footerIdx : html.length);

    const head = html.slice(0, firstIdx);
    const region = html.slice(firstIdx, regionEnd);
    const tail = html.slice(regionEnd); // starts at END MAIN CONTENT, carries footer token through

    // Canonical block = the FIRST article only (first START ARTICLE up to the
    // next START ARTICLE within the region, or the whole region if only one).
    const secondIdx = region.slice(1).search(marker); // search after char 0
    const canonicalBlock = secondIdx === -1
      ? region
      : region.slice(0, secondIdx + 1);

    const shellHtml = head + "<!--#ARTICLES#-->\n" + tail;
    return { canonicalBlock, shellHtml };
  }

  if (variant === "aicpaTownHall") {
    // Repeating unit = the table row containing a highlight-img cell (320x181)
    // plus its adjacent text cell. Capture one full <tr>...</tr> that contains
    // highlight-img as the canonical block.
    const rowRe = /<tr>(?:(?!<\/tr>)[\s\S])*?highlight-img[\s\S]*?<\/tr>/i;
    const firstRow = html.match(rowRe);
    if (!firstRow) return { canonicalBlock: null, shellHtml: html };

    // Replace all such rows with a single insertion marker.
    const allRowsRe = /<tr>(?:(?!<\/tr>)[\s\S])*?highlight-img[\s\S]*?<\/tr>/gi;
    let replaced = false;
    const shellHtml = html.replace(allRowsRe, () => {
      if (replaced) return "";
      replaced = true;
      return "<!--#ARTICLES#-->";
    });
    return { canonicalBlock: firstRow[0], shellHtml };
  }

  return { canonicalBlock: null, shellHtml: html };
}

// ═════════════════════════════════════════════
// model-driven generation for design templates
// ═════════════════════════════════════════════

const DESIGN_SYSTEM_ADDENDUM = `
DESIGN / REPEATING-BLOCK TEMPLATE MODE:
You are populating a newsletter with a repeating article/highlight structure.
You are given:
  1. A SHELL with a single <!--#ARTICLES#--> marker showing where blocks go.
  2. ONE CANONICAL BLOCK showing the exact required structure of each block.
  3. SOURCE CONTENT to draw from.

Your task:
- Decide how many blocks the source content naturally warrants (do not pad).
- For EACH block, reproduce the canonical block's structure EXACTLY, changing
  only the editorial content (title text and body copy). Keep every style
  attribute, table structure, and class identical to the canonical block.
- Begin EACH block with the <!-- START ARTICLE --> comment exactly as it
  appears at the top of the canonical block. This marker delimits block
  boundaries and must be present on every block.
- Tokens of the form <!--#PROTECTED:...#--> are opaque. Copy them through
  verbatim into each block wherever they appear in the canonical block. NEVER
  alter, expand, remove, or reorder the characters inside a protected token.
  They are placeholders for fragile pre-built markup that will be restored
  after you finish.
- Replace the single <!--#ARTICLES#--> marker in the shell with all your
  blocks concatenated in reading order.
- Output the COMPLETE shell with blocks inserted. Raw HTML only, nothing else.

DO NOT MANAGE NUMBERING:
- Do NOT renumber or edit image paths, element ids, or link anchors between
  blocks (e.g. article-N image paths, id="artN", #artN anchors). Copy them
  from the canonical block exactly as-is — including inside protected tokens.
  Sequential per-block numbering is applied automatically after you finish, so
  any numbers you see in the canonical block are placeholders; leave them
  untouched rather than trying to increment them.

EDITORIAL JUDGMENT (this is yours to exercise):
- Order blocks from top-to-bottom in the order they appear in the source
  document or provided body content.
- Write block titles and body copy from the source; keep them tight and
  on-brand. Match the density/tone notes provided.
- If the source clearly maps to fewer blocks than the canonical example
  implies, use fewer. If more, use more. Never invent facts, URLs, or stats
  not present in the source or instructions.`;

/**
 * Generate a design (repeating-block) template, model-driven, with fragile
 * markup protected across the round-trip and per-article numbering stamped
 * deterministically afterward.
 *
 * @param {object}   deps  injected server.js helpers so this module stays
 *                         decoupled: { anthropic, AI_SYSTEM_PROMPT,
 *                         extractHtml, logUsage, getBrandGuide }
 */
async function generateDesignHTML(ticket, templateName, templateHtml, sourceContent, deps) {
  const { anthropic, AI_SYSTEM_PROMPT, extractHtml, logUsage, getBrandGuide } = deps;
  const variant = designVariant(templateName);

  // 1. Protect fragile regions (bracket buttons, strip images, footer) FIRST.
  const { html: protectedHtml, store } = protectFragileRegions(templateHtml);

  // 2. Extract the canonical block + build the shell with an insertion marker.
  const { canonicalBlock, shellHtml } = extractArticleShells(protectedHtml, variant);

  if (!canonicalBlock) {
    // Structure not recognized — signal caller to fall back to the normal path
    // rather than silently producing a broken newsletter.
    const err = new Error(`Could not extract canonical block for "${templateName}" — falling back to standard generation.`);
    err.code = "NO_CANONICAL_BLOCK";
    throw err;
  }

  // 3. Pull variant-specific tone/structure notes from the brand guide so the
  //    prompt carries the documented quirks (font-weight 400 body, navy
  //    section headers, protected bracket buttons, etc.).
  let brandNotes = "";
  try {
    const bg = await getBrandGuide();
    const v = bg?.archetypes?.newsletter?.variants?.[variant];
    if (v) brandNotes = `\nBRAND VARIANT NOTES for this template:\n${v}\n`;
  } catch { /* brand guide optional here; addendum already carries the contract */ }

  const message = await anthropic.messages.create({
    model: "claude-opus-4-6",
    max_tokens: 16000,
    system: [
      { type: "text", text: AI_SYSTEM_PROMPT, cache_control: { type: "ephemeral" } },
      { type: "text", text: DESIGN_SYSTEM_ADDENDUM },
    ],
    messages: [{
      role: "user",
      content:
        `Template: ${templateName}\n` +
        `Subject: ${ticket.subjectLine || ticket.name || ""}\n` +
        brandNotes +
        `\nCANONICAL BLOCK (reproduce this structure exactly per article; ` +
        `protected tokens must be copied through verbatim):\n\n${canonicalBlock}\n\n` +
        `SHELL (insert all blocks where the <!--#ARTICLES#--> marker is):\n\n${shellHtml}\n\n` +
        `SOURCE CONTENT to turn into blocks:\n${sourceContent || "(use the description below)"}\n\n` +
        `Description: ${ticket.description || ""}\n` +
        (ticket.__freeform__ ? `\nAdditional requestor instructions:\n${ticket.__freeform__}\n` : ""),
    }],
  });

  const raw = message.content.find(b => b.type === "text")?.text ?? "";
  logUsage(`generate-design "${ticket.name}"`, message);
  let html = extractHtml(raw, shellHtml);

  // 4. Integrity check BEFORE restoring: did the model corrupt any tokens?
  const corrupted = findCorruptedTokens(html, store);
  if (corrupted.length > 0) {
    console.warn(`[design] ${corrupted.length} corrupted protected token(s) in "${ticket.name}": ${corrupted.join(", ")}`);
  }

  // 5. If the model left the insertion marker unreplaced, it failed the task.
  if (html.includes("<!--#ARTICLES#-->")) {
    console.warn(`[design] Model left #ARTICLES# marker unreplaced for "${ticket.name}" — restoring what we can.`);
  }

  // 6. Restore all protected regions byte-for-byte.
  const { html: restored, restored: n, missing, total } = restoreFragileRegions(html, store);
  console.log(`[design] Protected regions restored: ${n}/${total} (${missing} intentionally dropped with removed blocks)`);

  // 7. Stamp sequential per-article numbering (AI-in-Focus only; no-op elsewhere).
  const { html: finalHtml, count, warning } = renumberArticles(restored, variant);
  if (warning) console.warn(`[design] renumber for "${ticket.name}": ${warning}`);
  else if (count) console.log(`[design] renumbered ${count} article(s) sequentially`);

  return finalHtml;
}

/**
 * ---------------------------------------------------------------------------
 * RAG few-shot template retrieval layer.
 * Drop these functions into server.js alongside fetchTemplateIndex() /
 * fetchTemplateFromSharePoint() — those stay unchanged, templates HTML keeps
 * living in SharePoint, read-only, exactly as today.
 *
 * SOURCE OF TRUTH for *which templates exist* = the SharePoint folder listing
 * itself (fetchTemplateIndex()). manifest.json is an auto-maintained tag
 * cache, not hand-authored — new .html files in the folder get auto-tagged
 * the next time the manifest loads.
 *
 * manifest.json AND brand-guide.json are hosted on GitHub, not SharePoint —
 * SharePoint write scope (Files.ReadWrite.All) wasn't available, and a
 * fine-grained GitHub PAT scoped to just this repo's Contents is much easier
 * to provision than a Graph admin-consent flow. Bonus: every manifest update
 * becomes a git commit, so you get tag-change history for free.
 *
 * New env vars:
 *   GITHUB_TOKEN             - fine-grained PAT, Contents: Read & Write on the repo below
 *   GITHUB_OWNER             - e.g. "cpa-com"
 *   GITHUB_REPO              - e.g. "email-agent-data"
 *   GITHUB_BRANCH            - default "main"
 *   GITHUB_MANIFEST_PATH     - e.g. "manifest.json"
 *   GITHUB_BRAND_GUIDE_PATH  - e.g. "brand-guide.json"
 * ---------------------------------------------------------------------------
 */

let manifestCache   = null; // { data, fetchedAt, sha }
let brandGuideCache = null; // { data, fetchedAt }

const GITHUB_API = "https://api.github.com";
const GH_OWNER  = process.env.GITHUB_OWNER;
const GH_REPO   = process.env.GITHUB_REPO;
const GH_BRANCH = process.env.GITHUB_BRANCH || "main";

// ---------------------------------------------------------------------------
// GitHub Contents API helpers
// ---------------------------------------------------------------------------
async function githubGetFile(path) {
  const res = await fetch(
    `${GITHUB_API}/repos/${GH_OWNER}/${GH_REPO}/contents/${encodeURIComponent(path)}?ref=${GH_BRANCH}`,
    {
      headers: {
        Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
        Accept: "application/vnd.github+json",
      },
    }
  );
  if (res.status === 404) return null; // file doesn't exist yet
  if (!res.ok) throw new Error(`GitHub read failed for ${path}: ${res.status} ${res.statusText}`);
  const json = await res.json();
  const content = Buffer.from(json.content, "base64").toString("utf-8");
  return { content: JSON.parse(content), sha: json.sha };
}

async function githubPutFile(path, data, message, knownSha = null) {
  const sha = knownSha ?? (await githubGetFile(path))?.sha ?? null;

  const res = await fetch(
    `${GITHUB_API}/repos/${GH_OWNER}/${GH_REPO}/contents/${encodeURIComponent(path)}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message,
        content: Buffer.from(JSON.stringify(data, null, 2)).toString("base64"),
        branch: GH_BRANCH,
        ...(sha ? { sha } : {}), // omit sha on first-ever create
      }),
    }
  );

  if (res.status === 409 || res.status === 422) {
    // stale sha (concurrent write) — refetch and retry once
    console.warn(`GitHub write conflict on ${path}, refetching sha and retrying once.`);
    const fresh = await githubGetFile(path);
    return githubPutFileRetry(path, data, message, fresh?.sha ?? null);
  }
  if (!res.ok) throw new Error(`GitHub write failed for ${path}: ${res.status} ${res.statusText}`);
  const json = await res.json();
  return json.content.sha;
}

async function githubPutFileRetry(path, data, message, sha) {
  const res = await fetch(
    `${GITHUB_API}/repos/${GH_OWNER}/${GH_REPO}/contents/${encodeURIComponent(path)}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message,
        content: Buffer.from(JSON.stringify(data, null, 2)).toString("base64"),
        branch: GH_BRANCH,
        ...(sha ? { sha } : {}),
      }),
    }
  );
  if (!res.ok) throw new Error(`GitHub write retry failed for ${path}: ${res.status} ${res.statusText}`);
  const json = await res.json();
  return json.content.sha;
}

async function getManifest() {
  if (manifestCache && Date.now() - manifestCache.fetchedAt < CACHE_TTL_MS) {
    return manifestCache.data;
  }
  const path = process.env.GITHUB_MANIFEST_PATH || "manifest.json";
  const remote = await githubGetFile(path); // null if not created yet
  let data = remote?.content ?? { version: 0, updated: null, templates: [] };
  if (!remote) console.log("No manifest.json in GitHub repo yet — starting from empty manifest.");

  data = await ensureManifestUpToDate(data);
  manifestCache = { data, fetchedAt: Date.now() };
  console.log(`Template manifest refreshed: ${data.templates.length} entries`);
  return data;
}

// ---------------------------------------------------------------------------
// Diff the manifest against the live SharePoint folder listing (still the
// real source of truth for *which templates exist*). Any .html file present
// in the folder but not yet in the manifest gets auto-tagged with one cheap
// Haiku call, then the manifest is committed to GitHub so this only happens
// once per new template.
// ---------------------------------------------------------------------------
async function ensureManifestUpToDate(manifest) {
  const liveIndex = await fetchTemplateIndex({ includeNonHtml: false }); // { displayName: itemId }
  const knownFilenames = new Set(manifest.templates.map(t => t.filename));

  // liveIndex keys are display-cased names; we need the raw filenames too —
  // fetchTemplateIndex({ includeNonHtml: true }) gives raw filenames, reuse that.
  const rawIndex = await fetchTemplateIndex({ includeNonHtml: true });
  const newFilenames = Object.keys(rawIndex).filter(
    name => /\.html$/i.test(name) && !knownFilenames.has(name)
  );

  if (newFilenames.length === 0) return manifest;

  console.log(`Found ${newFilenames.length} untagged template(s): ${newFilenames.join(", ")}`);

  const newEntries = [];
  for (const filename of newFilenames) {
    try {
      const html = await fetchTemplateFromSharePoint(filename.replace(/\.html$/i, ""));
      const entry = await autoTagTemplate(filename, html, manifest.templates);
      newEntries.push(entry);
    } catch (err) {
      console.error(`Auto-tagging failed for ${filename}: ${err.message}`);
    }
  }

  if (newEntries.length === 0) return manifest;

  const updated = {
    version: (manifest.version || 0) + 1,
    updated: new Date().toISOString(),
    templates: [...manifest.templates, ...newEntries],
  };

  await writeManifestToGithub(updated); // commits the update to the GitHub repo
  return updated;
}

// ---------------------------------------------------------------------------
// One cheap Haiku call per new template: infer tags/useCase/archetype/tone
// from the HTML itself. Fed the existing tag vocabulary so it reuses tags
// instead of fragmenting ("CAS workshop" vs "CAS Workshop Promo" etc).
// ---------------------------------------------------------------------------
async function autoTagTemplate(filename, html, existingTemplates) {
  const tagVocab = [...new Set(existingTemplates.flatMap(t => t.tags))];

  const response = await anthropic.messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 400,
    system:
      "You tag marketing email HTML templates for a retrieval library. " +
      "Respond ONLY with JSON matching this shape: " +
      '{"tags": ["..."], "useCase": "...", "archetype": "newsletter|general", ' +
      '"toneNotes": "...", "ctaPattern": "..."}. ' +
      "Prefer reusing tags from the existing vocabulary given below when they fit; " +
      "only add a new tag if nothing existing captures it. Keep useCase, toneNotes, " +
      "and ctaPattern each to one concise sentence.",
    messages: [
      {
        role: "user",
        content:
          `Filename: ${filename}\n\n` +
          `Existing tag vocabulary: ${tagVocab.join(", ") || "(none yet)"}\n\n` +
          `Template HTML:\n${html.slice(0, 12000)}`, // cap to keep the call cheap
      },
    ],
  });

  const text = response.content.find(b => b.type === "text")?.text ?? "{}";
  let parsed;
  try {
    parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
  } catch {
    parsed = {};
  }

  return {
    id: filename.replace(/\.html$/i, "").toLowerCase().replace(/[^a-z0-9]+/g, "-"),
    filename,
    archetype: parsed.archetype || "general",
    tags: parsed.tags || [],
    useCase: parsed.useCase || "",
    toneNotes: parsed.toneNotes || "",
    ctaPattern: parsed.ctaPattern || "",
    addedDate: new Date().toISOString().slice(0, 10),
    autoTagged: true,
  };
}

// ---------------------------------------------------------------------------
// Requires GITHUB_TOKEN with Contents: Read & Write on GITHUB_OWNER/GITHUB_REPO. Degrades
// gracefully — logs once and keeps working from the in-memory cache if the
// token/repo isn't configured yet, just re-tags on every cache expiry
// instead of persisting.
// ---------------------------------------------------------------------------
let githubWriteSupported = true;

async function writeManifestToGithub(manifest) {
  if (!githubWriteSupported) return;
  const path = process.env.GITHUB_MANIFEST_PATH || "manifest.json";
  try {
    await githubPutFile(
      path,
      manifest,
      `auto-tag: ${manifest.templates.length} templates (v${manifest.version})`
    );
    console.log("manifest.json committed to GitHub.");
  } catch (err) {
    githubWriteSupported = false;
    console.warn(
      `manifest.json commit to GitHub failed: ${err.message} — check GITHUB_TOKEN scope ` +
      `and GITHUB_OWNER/GITHUB_REPO. Auto-tagging will still run each cache cycle, it just ` +
      `won't persist between runs until this is fixed.`
    );
  }
}

async function getBrandGuide() {
  if (brandGuideCache && Date.now() - brandGuideCache.fetchedAt < CACHE_TTL_MS) {
    return brandGuideCache.data;
  }
  const path = process.env.GITHUB_BRAND_GUIDE_PATH || "brand-guide.json";
  const remote = await githubGetFile(path);
  if (!remote) throw new Error(`brand-guide.json not found in GitHub repo at path "${path}".`);
  brandGuideCache = { data: remote.content, fetchedAt: Date.now() };
  console.log("Brand guide refreshed from GitHub.");
  return remote.content;
}

// ---------------------------------------------------------------------------
// Tier 1: cheap keyword/tag overlap scoring. No API call.
// ---------------------------------------------------------------------------
function scoreTemplateAgainstText(template, text) {
  const haystack = text.toLowerCase();
  let score = 0;
  for (const tag of template.tags) {
    if (haystack.includes(tag.toLowerCase())) score += 3;
  }
  // loose token overlap as a secondary signal
  const tagTokens = new Set(
    template.tags.join(" ").toLowerCase().split(/\W+/).filter(w => w.length > 3)
  );
  for (const tok of tagTokens) {
    if (haystack.includes(tok)) score += 1;
  }
  return score;
}

function classifyByKeyword(manifest, ticketText) {
  const scored = manifest.templates
    .map(t => ({ template: t, score: scoreTemplateAgainstText(t, ticketText) }))
    .filter(s => s.score > 0)
    .sort((a, b) => b.score - a.score);

  if (scored.length === 0) return { confident: false, matches: [] };

  const topScore = scored[0].score;
  const runnerUpScore = scored[1]?.score ?? 0;
  // "confident" = clear leader, comfortably above zero
  const confident = topScore >= 4 && topScore - runnerUpScore >= 2;

  return { confident, matches: scored.slice(0, 3).map(s => s.template) };
}

// ---------------------------------------------------------------------------
// Tier 2: cheap Haiku call, only invoked when Tier 1 is ambiguous.
// Structured JSON output, minimal tokens (tag list only, not full templates).
// ---------------------------------------------------------------------------
async function classifyWithModel(manifest, ticketText) {
  const tagVocab = [...new Set(manifest.templates.flatMap(t => t.tags))];

  const response = await anthropic.messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 300,
    system:
      "You classify marketing email tickets against a fixed list of template IDs. " +
      "Respond ONLY with JSON: {\"templateIds\": [\"id1\", \"id2\"]} — 1 to 3 ids, " +
      "most relevant first. Prefer reusing the given tag vocabulary in your reasoning; " +
      "do not invent template ids that aren't in the list below.",
    messages: [
      {
        role: "user",
        content:
          `Ticket text:\n${ticketText}\n\n` +
          `Available templates:\n${manifest.templates
            .map(t => `- ${t.id}: ${t.useCase} (tags: ${t.tags.join(", ")})`)
            .join("\n")}\n\n` +
          `Known tag vocabulary: ${tagVocab.join(", ")}`,
      },
    ],
  });

  const text = response.content.find(b => b.type === "text")?.text ?? "{}";
  let parsed;
  try {
    parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
  } catch {
    return [];
  }
  const ids = new Set(parsed.templateIds || []);
  return manifest.templates.filter(t => ids.has(t.id));
}

// ---------------------------------------------------------------------------
// Public entry point: classify a ticket, return 1-3 template manifest entries.
// ---------------------------------------------------------------------------
async function selectRelevantTemplates(ticket) {
  const manifest = await getManifest();
  const ticketText = [
    ticket.description,
    ticket.subjectLine,
    ticket.templateDropdownValue,
    ticket.groupName,
  ]
    .filter(Boolean)
    .join("\n");

  const tier1 = classifyByKeyword(manifest, ticketText);
  if (tier1.confident) {
    console.log(`Template classification (keyword, confident): ${tier1.matches.map(t => t.id).join(", ")}`);
    return tier1.matches;
  }

  console.log("Template classification: keyword match ambiguous, falling back to model.");
  const modelMatches = await classifyWithModel(manifest, ticketText);
  if (modelMatches.length > 0) return modelMatches;

  // last resort: fall back to keyword matches even if not "confident",
  // or empty (caller should fall back to the default General archetype)
  return tier1.matches;
}

// ---------------------------------------------------------------------------
// Strip the footer out of a few-shot HTML example (server reattaches the real
// footer post-generation anyway — see brand-guide.footer.serverSideReattachment).
// Keeps few-shot examples short and avoids ever putting real footer content
// into a "here's a style example" context by mistake.
// ---------------------------------------------------------------------------
function stripFooterForFewShot(html) {
  return html.replace(/<!--\s*START FOOTER\s*-->[\s\S]*?<!--\s*END FOOTER\s*-->/i, "<!-- FOOTER OMITTED: server reattaches actual footer verbatim -->");
}

// ---------------------------------------------------------------------------
// Build the system prompt: static rules + brand guide (cacheable prefix)
// + up to 3 footer-stripped few-shot examples (per-ticket, not cached).
// ---------------------------------------------------------------------------
async function buildSystemPromptWithRAG({ staticBrandRules, ticket }) {
  const brandGuide = await getBrandGuide();
  const selected = await selectRelevantTemplates(ticket);

  const fewShotBlocks = await Promise.all(
    selected.map(async t => {
      const html = await fetchTemplateFromSharePoint(t.filename.replace(/\.html$/i, ""));
      return (
        `### Example: ${t.useCase}\n` +
        `Tags: ${t.tags.join(", ")}\n` +
        `Tone notes: ${t.toneNotes}\n` +
        `CTA pattern: ${t.ctaPattern}\n\n` +
        "```html\n" + stripFooterForFewShot(html) + "\n```"
      );
    })
  );

  return [
    {
      type: "text",
      text: staticBrandRules,
    },
    {
      type: "text",
      text: "BRAND GUIDE (structured):\n" + JSON.stringify(brandGuide, null, 2),
      cache_control: { type: "ephemeral" }, // stable prefix ends here
    },
    {
      type: "text",
      text:
        selected.length > 0
          ? "RELEVANT TEMPLATE EXAMPLES (retrieved for this ticket, footers omitted — do not reproduce footer content from these):\n\n" +
            fewShotBlocks.join("\n\n---\n\n")
          : "No closely-matching prior template found — use the General archetype as the default fallback.",
    },
  ];
}

// ═════════════════════════════════════════════
// Monday GraphQL helper
// ═════════════════════════════════════════════
async function mondayQuery(query, variables = {}, apiVersion = "2024-01") {
  const res = await fetch(MONDAY_API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: process.env.MONDAY_API_TOKEN, "API-Version": apiVersion },
    body: JSON.stringify({ query, variables }),
  });
  const data = await res.json();
  if (data.errors) throw new Error(data.errors.map(e => e.message).join(", "));
  return data.data;
}

function normalizeTicket(item) {
  const cols = {};
  const formulaVals = {};
  for (const cv of item.column_values) {
    cols[cv.id] = cv.text ?? cv.value ?? "";
    // Formula columns return blank text/value — capture display_value instead
    if (cv.display_value !== undefined && cv.display_value !== null && cv.display_value !== "") {
      formulaVals[cv.id] = cv.display_value;
    }
  }
  return {
    id:           item.id,
    name:         item.name,
    status:       cols["status"]             || "",
    description:  cols["long_text7"]         || "",
    subjectLine:  cols["text86"]             || "",
    jobNumber:    formulaVals["formula"]     || cols["formula"] || "",
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
        ]) {
          id text value
          ... on FormulaValue { display_value }
        }
      }
    }
  `, { itemId }, "2025-07");
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


async function downloadMondayAsset(asset) {
  const src = asset.public_url || asset.url;
  const res = await fetch(src);
  if (!res.ok) throw new Error(`Asset download failed: ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// Find a Word document in the ticket's Files column and extract its content as
// clean HTML (headings, bold, lists, links preserved). Returns { name, html }
// or null if none is attached. Legacy binary .doc is not supported — .docx only.
async function fetchWordDocContent(itemId) {
  try {
    const assets = await fetchTicketFiles(itemId);
    const doc = assets.find(a => /\.docx?$/i.test(a.name || ""));
    if (!doc) return null;

    if (/\.doc$/i.test(doc.name)) {
      console.warn(`Legacy .doc not supported ("${doc.name}") — ask the requestor to attach a .docx.`);
      return null;
    }

    const buffer = await downloadMondayAsset(doc);
    const result = await mammoth.convertToHtml({ buffer });
    const html = (result.value || "").trim();
    if (!html) {
      console.warn(`Word doc "${doc.name}" produced no content.`);
      return null;
    }
    console.log(`Extracted Word doc "${doc.name}" (${html.length} chars of HTML)`);
    return { name: doc.name, html };
  } catch (err) {
    console.warn(`Word doc extraction failed for ${itemId}: ${err.message}`);
    return null;
  }
}

// Replace the first <img> src in the HTML, and optionally wrap it in an <a> tag.
// - newUrl: the image src to inject; pass null to leave the existing src unchanged
// - linkUrl: if provided, wraps the <img> in <a href="linkUrl" ...> (optional)
function replaceFirstImage(html, newUrl, linkUrl = "") {
  let replaced = false;
  return html.replace(/(<img\b[^>]*\bsrc=")([^"]*)("[^>]*>)/i, (m, pre, existingSrc, post) => {
    if (replaced) return m;
    replaced = true;
    const resolvedSrc = newUrl !== null ? newUrl : existingSrc;
    const img = pre + resolvedSrc + post;
    if (linkUrl) {
      return `<a href="${linkUrl}" target="_blank" style="display:block;border:0;text-decoration:none;">${img}</a>`;
    }
    return img;
  });
}

// Apply header image and/or header link to the email HTML.
// Behaviour matrix (all conditional):
//   HeaderImage absent,  HeaderLink absent   → no changes; template header left as-is
//   HeaderImage provided, HeaderLink absent  → find asset by filename in Files column,
//                                              use its Monday public_url as the <img> src
//   HeaderImage absent,  HeaderLink provided → wrap existing <img> in <a href="HeaderLink">,
//                                              src left unchanged
//   HeaderImage provided, HeaderLink provided → find asset by filename, swap <img> src
//                                              to its public_url AND wrap in <a>
//
// NOTE: Monday's public_url is a TEMPORARY, expiring link — fine for proofing,
// but the production header must point at a permanent marketing.cpa.com asset
// before the Pardot send.
async function applyHeaderImage(html, itemId, vars = {}) {
  const headerLink     = (vars["HeaderLink"]  || "").trim();
  const headerFilename = (vars["HeaderImage"] || "").trim();

  // Nothing to do — leave the template header completely untouched
  if (!headerFilename && !headerLink) return html;

  try {
    if (headerFilename) {
      // Look up the named file in the Files column
      const assets = await fetchTicketFiles(itemId);
      const match  = assets.find(a => a.name.toLowerCase() === headerFilename.toLowerCase());
      if (!match) {
        console.warn(`Header image "${headerFilename}" not found in Files column for item ${itemId} — header left unchanged.`);
        // Still apply HeaderLink to the existing image if provided
        return headerLink ? replaceFirstImage(html, null, headerLink) : html;
      }
      // Use Monday's temporary public URL directly (no SharePoint re-host)
      const publicUrl = match.public_url || match.url;
      if (!publicUrl) {
        console.warn(`Header image "${headerFilename}" has no public URL — header left unchanged.`);
        return headerLink ? replaceFirstImage(html, null, headerLink) : html;
      }
      console.log(`Header image set from Monday public_url (temporary): ${headerFilename}`);
      return replaceFirstImage(html, publicUrl, headerLink);
    } else {
      // HeaderLink only — wrap the existing <img> without touching its src
      return replaceFirstImage(html, null, headerLink);
    }
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

// ═════════════════════════════════════════════
// Requestor prompt
// ═════════════════════════════════════════════

// ---------------------------------------------------------------------------
// New prompt column is a FREEFORM PROMPT by default.
// Any text not claimed by a recognized label becomes __freeform__ (the prompt
// passed to the model). No "Prompt:" preface required anymore.
//
// Labels are now split into two kinds, which fixes a real fragility: in a
// freeform-first column, a stray label mid-prose used to swallow everything
// after it. So:
//   SINGLE_LINE_LABELS — value is the REST OF THAT LINE only. These are the
//     mechanical directives (asset lookup, token substitution) whose values
//     are always short (a filename, a URL, a name). Safe to drop into prose.
//   BLOCK_LABELS — value runs until the next recognized label or end. These
//     are multi-line content regions (BodyContent) and the legacy Prompt:
//     alias (still accepted, merged into __freeform__).
// ---------------------------------------------------------------------------
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
      return ""; // strip the directive line
    });
  }
   // 2. Handle block labels within what's left. Capture from a block label up
  //    to the next block label (or end).
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
    } else { // bodycontent | bodytext
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
  // Any text before the first recognised label is treated as leading body copy.
  if (firstLabelIndex > 0) {
    const lead = instructions.slice(0, firstLabelIndex).trim();
    if (lead) vars["BodyContent"] = vars["BodyContent"] ? `${lead}\n\n${vars["BodyContent"]}` : lead;
  } else if (firstLabelIndex === -1) {
    // No labels at all — the entire field is body content
    const all = instructions.trim();
    if (all) vars["BodyContent"] = all;
  }

  return vars;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------
if (require.main === module) {
  const cases = [
    {
      name: "Pure prompt, no labels → all __freeform__",
      in: "Write the May AI in Focus newsletter from the attached doc. Keep each article tight.",
      expect: v => v.__freeform__ && !v.BodyContent && !v.HeaderImage,
    },
    {
      name: "Prompt + single-line HeaderImage (label does NOT swallow following prose)",
      in: "Build the general announcement email.\nHeaderImage: spring-promo.png\nMake the tone upbeat and mention the June 1 deadline.",
      expect: v =>
        v.HeaderImage === "spring-promo.png" &&
        /upbeat/.test(v.__freeform__) &&
        /general announcement/.test(v.__freeform__) &&
        !/spring-promo/.test(v.__freeform__),
    },
    {
      name: "Legacy Prompt: label still merges into __freeform__",
      in: "Prompt: Emphasize the early-bird discount and keep it under 150 words.",
      expect: v => /early-bird/.test(v.__freeform__) && !v.BodyContent,
    },
    {
      name: "BodyContent block still captured verbatim (power-user escape hatch)",
      in: "Some framing for the model.\nBodyContent: Hi there,\n\nThis is literal copy.\n\nButton: Register\nLink: https://cpa.com/x",
      expect: v =>
        /literal copy/.test(v.BodyContent) &&
        /Button: Register/.test(v.BodyContent) &&
        /framing for the model/.test(v.__freeform__),
    },
    {
      name: "Multiple single-line labels + prompt",
      in: "PreheaderText: Your monthly roundup\nHeaderImage: hero.png\nPartner: BILL\nWrite a co-branded webinar invite; registration link is in the doc.",
      expect: v =>
        v.PreheaderText === "Your monthly roundup" &&
        v.HeaderImage === "hero.png" &&
        v.Partner === "BILL" &&
        /co-branded webinar/.test(v.__freeform__),
    },
    {
      name: "Empty input → empty vars",
      in: "",
      expect: v => Object.keys(v).length === 0,
    },
  ];
 
  let pass = 0;
  for (const c of cases) {
    const v = parseInstructions(c.in);
    let ok = false;
    try { ok = c.expect(v); } catch { ok = false; }
    console.log(`${ok ? "PASS" : "FAIL"}  ${c.name}`);
    if (!ok) console.log("   got:", JSON.stringify(v));
    if (ok) pass++;
  }
  console.log(`\n${pass}/${cases.length} passed`);
}

// ─────────────────────────────────────────────
// BodyContent rendering: turn a mixed text + Button-block region into HTML,
// preserving the order in which elements were written.
// ─────────────────────────────────────────────
const BODY_PARAGRAPH_STYLE = "margin:0 0 10px 0;font-family:'Roboto',Arial,Helvetica,sans-serif;font-size:16px;font-weight:300;line-height:22px;color:#000000;";

function renderTextSegment(text) {
  const trimmed = (text || "").trim();
  if (!trimmed) return "";
  // Blank line → new paragraph; single newline → <br>
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
    // Text before this button block
    out += renderTextSegment(content.slice(lastIndex, m.index));
    // The button itself
    const btnText = m[1].trim();
    const rest    = m[2];
    const link    = (rest.match(/Link:\s*(\S+)/i)  || [])[1];
    const color   = (rest.match(/Color:\s*(.+)/i)  || [])[1];
    const style   = (rest.match(/Style:\s*(\w+)/i) || [])[1];
    if (link) {
      out += "\n" + buildButton(btnText, link.trim(), color && color.trim(), style && style.trim()) + "\n";
    } else {
      // Malformed button block — keep as text rather than dropping it
      out += renderTextSegment(m[0]);
    }
    lastIndex = m.index + m[0].length;
  }
  // Trailing text after the last button
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

  // Body content: render the mixed text + inline-Button region in order, then
  // drop it into the body placeholder(s). Buttons are NOT injected separately —
  // they appear exactly where the marketer placed them within BodyContent.
  if (vars["BodyContent"]) {
    const renderedBody = renderBodyContent(vars["BodyContent"]);
    out = out
      .replace(/BODY_CONTENT_HERE/g, renderedBody)
      .replace(/\{\{BodyContent\}\}/g, renderedBody)
      .replace(/\{\{BodyText\}\}/g, renderedBody);
  }

  // Job Number: support {{JobNumber}}, JOB_NUMBER_HERE, and plain JOB_NUMBER tokens
  if (vars["JobNumber"]) {
    out = out
      .replace(/\{\{JobNumber\}\}/g, vars["JobNumber"])
      .replace(/JOB_NUMBER_HERE/g, vars["JobNumber"])
      .replace(/\bJOB_NUMBER\b/g, vars["JobNumber"]);
  }

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
// Log token usage + prompt-cache activity for a Claude response.
// cache_creation_input_tokens = tokens written to cache (first call, full price
//   +25%); cache_read_input_tokens = tokens served from cache (~90% cheaper).
// ═════════════════════════════════════════════
function logUsage(label, message) {
  const u = message?.usage || {};
  console.log(
    `[usage] ${label} — in:${u.input_tokens ?? "?"} out:${u.output_tokens ?? "?"} ` +
    `cache_write:${u.cache_creation_input_tokens ?? 0} cache_read:${u.cache_read_input_tokens ?? 0}`
  );
}

// ═════════════════════════════════════════════
// Strip any prose preamble/postamble Claude may have added around the HTML.
// Finds the first < and last > in the response and returns only that slice.
// Falls back to the original if no valid HTML envelope is found.
// ═════════════════════════════════════════════
function extractHtml(raw, fallback) {
  const start = raw.indexOf("<");
  const end   = raw.lastIndexOf(">");
  if (start === -1 || end === -1 || end < start) {
    console.warn("extractHtml: no valid HTML envelope found in Claude response — using fallback");
    return fallback;
  }
  return raw.slice(start, end + 1).trim();
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

  let html = applyVariables(templateHtml, vars, ticket.instructions || "");
  
  if (design.isDesignTemplate(templateName)) {
    const vars = parseInstructions(ticket.instructions);
    const docContent = (ticket.id && ticket.id !== "manual")
      ? await fetchWordDocContent(ticket.id) : null;
    // Source the model lays out: doc first, then literal BodyContent, then the prompt itself.
    const sourceContent = docContent?.html || vars["BodyContent"] || vars["__freeform__"] || "";
    // Only pass __freeform__ as *separate* instructions when it isn't already the source,
    // so we don't duplicate it into the prompt twice.
    ticket.__freeform__ = (sourceContent === vars["__freeform__"]) ? "" : (vars["__freeform__"] || "");
    try {
      let html = await design.generateDesignHTML(ticket, templateName, templateHtml, sourceContent,
        { anthropic, AI_SYSTEM_PROMPT, extractHtml, logUsage, getBrandGuide });
      if (ticket.id && ticket.id !== "manual") html = await applyHeaderImage(html, ticket.id, vars);
      return html;
    } catch (e) {
      if (e.code !== "NO_CANONICAL_BLOCK") throw e;
      // else fall through to the standard placeholder path below
    }
  }
  
  const stillMissing = [];
  if (/BODY_CONTENT_HERE/.test(html) || /\{\{BodyContent\}\}/.test(html) || /\{\{BodyText\}\}/.test(html)) stillMissing.push("BodyContent");
  if (/\{\{PreheaderText\}\}/.test(html)) stillMissing.push("PreheaderText");

  // If the body still needs content, check for an attached Word doc to use as
  // source material. Instructions BodyContent takes precedence (it would have
  // already filled the token above); the doc is the fallback content source.
  let docContent = null;
  if (stillMissing.includes("BodyContent") && ticket.id && ticket.id !== "manual") {
    docContent = await fetchWordDocContent(ticket.id);
  }

  if (stillMissing.length > 0) {
    const message = await anthropic.messages.create({
      model: "claude-opus-4-6",
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
  }

  return html;
}

// ═════════════════════════════════════════════
// Revise HTML (agent feedback loop)
// ═════════════════════════════════════════════
async function reviseHTML(currentHtml, feedback, history, freeform = "") {
  const originalFooter = extractFooter(currentHtml);
  const historyText = history.length ? history.map((h, i) => `Round ${i + 1}: ${h}`).join("\n") : "(none)";

  const message = await anthropic.messages.create({
    model: "claude-opus-4-6",
    max_tokens: 16000,
    system: [{ type: "text", text: AI_SYSTEM_PROMPT, cache_control: { type: "ephemeral" } }],
    messages: [{
      role: "user",
      content: `You previously generated this email proof. The requestor has reviewed it and given feedback. Apply ONLY the requested changes while keeping everything else intact and following all CPA.com brand standards.
${freeform ? `\nSTANDING INSTRUCTIONS FROM REQUESTOR (remain in effect across all revisions):\n${freeform}\n` : ""}
PRIOR FEEDBACK ROUNDS:
${historyText}

NEW FEEDBACK TO APPLY:
${feedback}

CURRENT EMAIL HTML:
${currentHtml}`,
    }],
  });

  const rawRevised = message.content.find(b => b.type === "text")?.text ?? "";
  logUsage("revise", message);
  let html = extractHtml(rawRevised, currentHtml);
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

async function postUpdate(itemId, body, mentionIds = []) {
  const mentionsList = mentionIds.map(id => ({ id: parseInt(id, 10), type: "User" }));
  await mondayQuery(`
    mutation PostUpdate($itemId: ID!, $body: String!, $mentionsList: [UpdateMention]) {
      create_update(item_id: $itemId, body: $body, mentions_list: $mentionsList) { id }
    }
  `, { itemId, body, mentionsList: mentionsList.length ? mentionsList : undefined }, "2025-07");
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
  console.log(`[agent] event=${event.type} pulseId=${event.pulseId} boardId=${event.boardId}`);
  try {
    // ── Trigger 1: item created → first proof ──
    if (event.type === "create_pulse") {
      const itemId = String(event.pulseId);
      console.log(`[agent] New item ${itemId} — generating first proof`);

      let ticket = await fetchItemById(itemId);

      // Formula columns (Job Number) can lag a moment after item creation.
      // If it's empty on first read, wait briefly and re-fetch once.
      if (!ticket.jobNumber) {
        console.log(`[agent] Job Number empty for ${itemId} — retrying once after delay`);
        await new Promise(r => setTimeout(r, 3000));
        ticket = await fetchItemById(itemId);
      }

      // Guardrail: generate when there's Instructions content OR an attached
      // Word doc to use as body source. Skip only when there's neither.
      const hasInstructions = !!(ticket.instructions && ticket.instructions.trim());
      const hasDoc = await fetchWordDocContent(itemId);
      if (!hasInstructions && !hasDoc) {
        console.log(`[agent] Item ${itemId} has no Instructions and no Word doc — skipping generation`);
        return;
      }

      const templateName = await resolveTemplateName(ticket);
      const html         = await generateHTML(ticket, templateName);
      const fileName     = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v1.html`;

      // Upload proof + persist state BEFORE notifying — prevents feedback race
      await uploadToMonday(itemId, fileName, html);
      await persistAgentState(itemId, 1, [], html);

      const requestorId = await findUserIdByName(ticket.requestor);
      const mentionName = ticket.requestor.split(",")[0].trim();
      const mentionTag  = requestorId ? `<strong>@${mentionName}</strong> ` : "";
      await postUpdate(itemId,
        `<p>${mentionTag}Your email proof (v1) is ready and attached to this item's Files. ` +
        `Review it and reply with <strong>@agent</strong> followed by any changes you'd like. ` +
        `When it's ready, set Status to <strong>Approved</strong>.</p>`,
        requestorId ? [requestorId] : []
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

      // No prior draft yet — treat this as a request to generate the FIRST proof
      // (covers the case where the item was created without Instructions/doc, or
      // creation didn't trigger, and the requestor now tags @agent to kick it off).
      if (!baseHtml) {
        console.log(`[agent] No prior draft on ${itemId} — generating first proof from @agent request`);
        const templateName = await resolveTemplateName(ticket);
        const html         = await generateHTML(ticket, templateName);
        const fileName     = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v1.html`;

        await uploadToMonday(itemId, fileName, html);
        await persistAgentState(itemId, 1, [], html);

        const reqId    = await findUserIdByName(ticket.requestor);
        const reqName  = ticket.requestor.split(",")[0].trim();
        const reqTag   = reqId ? `<strong>@${reqName}</strong> ` : "";
        await postUpdate(itemId,
          `<p>${reqTag}Your email proof (v1) is ready and attached to this item's Files. ` +
          `Review it and reply with <strong>@agent</strong> followed by any changes you'd like. ` +
          `When it's ready, set Status to <strong>Approved</strong>.</p>`,
          reqId ? [reqId] : []
        );
        console.log(`[agent] Item ${itemId} v1 complete (via @agent)`);
        return;
      }

      const ticketVars = parseInstructions(ticket.instructions);
      const revised  = await reviseHTML(baseHtml, feedback, meta.history || [], ticketVars["__freeform__"] || "");
      const newRev   = (meta.revision || 1) + 1;
      const fileName = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v${newRev}.html`;

      await uploadToMonday(itemId, fileName, revised);
      await persistAgentState(itemId, newRev, [...(meta.history || []), feedback], revised);

      const requestorId = await findUserIdByName(ticket.requestor);
      const mentionName = ticket.requestor.split(",")[0].trim();
      const mentionTag  = requestorId ? `<strong>@${mentionName}</strong> ` : "";
      await postUpdate(itemId,
        `<p>${mentionTag}Updated proof (v${newRev}) is attached with your requested changes. ` +
        `Reply with <strong>@agent</strong> for more edits, or set Status to <strong>Approved</strong> when ready.</p>`,
        requestorId ? [requestorId] : []
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
