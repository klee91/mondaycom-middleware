/**
 * Monday.com Email Template Generator — Full Stack Agent Server
 *
 * Capabilities:
 *   - Serves frontend UI + REST API (manual generation)
 *   - SharePoint template library (Graph client-credentials)
 *   - Prompt-column freeform input + optional labels + brand buttons
 *   - Header image re-hosting from Files column
 *   - Word doc (.docx) in Files column → extracted as body source content
 *   - Footer + Pardot-tag protection, brand style enforcement
 *   - Design (repeating-block) templates: protected-region round-trip + renumber
 *   - RAG few-shot retrieval (GitHub-hosted manifest + brand guide)
 *   - Stateful agent webhook: item created → proof; @agent feedback → revise;
 *     Status=Approved → silent finalize
 *   - Generation model: claude-opus-4-6 (swap to claude-haiku-4-5-20251001 once fine-tuned)
 *
 * Required env vars:
 *   MONDAY_API_TOKEN, ANTHROPIC_API_KEY, BOARD_ID
 *   SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_DRIVE_ID
 *   GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, GITHUB_BRANCH,
 *   GITHUB_MANIFEST_PATH, GITHUB_BRAND_GUIDE_PATH
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
const INSTRUCTIONS_COLUMN = "long_text_mm47njms"; // Prompt column (display name may be "Prompt"; ID unchanged)
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

function fileNameToDisplayName(filename) {
  return filename
    .replace(/\.html$/i, "")
    .replace(/_TEMPLATE$/i, "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, c => c.toUpperCase());
}

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
  }
  templateIndexCache = { map, fetchedAt: Date.now() };
  console.log(`Template index refreshed: ${Object.keys(map).join(", ")}`);
  return map;
}

async function fetchTemplateFromSharePoint(templateName) {
  const index = await fetchTemplateIndex();

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
// DESIGN (repeating-block) templates: protect / restore / renumber / generate
// ═════════════════════════════════════════════
// Normalize a template name for matching: lowercase, drop the word
// "template"/"newsletter", strip everything non-alphanumeric. This makes
// matching robust across the three naming conventions in play (Monday dropdown
// label, SharePoint filename→display name, and these design keys). E.g.
// "AI in Focus", "AI IN FOCUS NEWSLETTER", "ai-in-focus" all normalize to
// "aiinfocus".
function normalizeTemplateName(name) {
  return (name || "")
    .toLowerCase()
    .replace(/\b(template|newsletter)\b/g, "")
    .replace(/[^a-z0-9]/g, "");
}

// Each design template lists the normalized forms that should map to it.
const DESIGN_TEMPLATES = {
  aiInFocus:     { variant: "aiInFocus",     aliases: ["aiinfocus"] },
  aicpaTownHall: { variant: "aicpaTownHall", aliases: ["aicpatownhall", "townhall"] },
};

function designVariant(templateName) {
  const n = normalizeTemplateName(templateName);
  for (const { variant, aliases } of Object.values(DESIGN_TEMPLATES)) {
    if (aliases.includes(n)) return variant;
  }
  return null;
}

function isDesignTemplate(templateName) {
  return designVariant(templateName) !== null;
}

const PROTECT_RULES = [
  {
    name: "bracketButton",
    re: /<table[^>]*width:\s*340px[^>]*>[\s\S]*?ai-nl-btn-left\.png[\s\S]*?ai-nl-btn-right\.png[\s\S]*?<\/table>/gi,
  },
  {
    name: "articleStripImage",
    re: /<img\b[^>]*article-\d+-top-s\d+\.png[^>]*>/gi,
  },
  {
    name: "footer",
    re: /<!--\s*START FOOTER\s*-->[\s\S]*?<!--\s*END FOOTER\s*-->/gi,
  },
];

function protectFragileRegions(html) {
  const store = {};
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

function findCorruptedTokens(html, store) {
  const known = new Set(Object.keys(store));
  const found = html.match(/<!--#PROTECTED:[^#]*#-->/g) || [];
  return found.filter(t => !known.has(t));
}

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

// ── Per-article strip-image dimension correction (AI-in-Focus) ──
// The article header frame (article-N-top-s1..s4.png) is a DIFFERENT pre-made
// asset per article, each with its own native pixel size (e.g. the top strip
// s1 is 101px tall for article 1 but 41px for article 2). Because every
// article is generated from the canonical (article-1) block and renumber only
// swaps the filename number, each article inherits article-1's declared
// width/height — which stretches the real asset. This pass reads each strip
// image's TRUE dimensions from the PNG header and rewrites the <img> (and its
// enclosing <td> size hint) to match, so nothing is forced out of aspect ratio.
const pngDimCache = new Map(); // url -> { w, h } | "failed"

async function fetchPngDimensions(url) {
  if (pngDimCache.has(url)) {
    const v = pngDimCache.get(url);
    if (v === "failed") throw new Error("cached failure");
    return v;
  }
  // Range request keeps this to a few dozen bytes; falls back fine if the
  // server ignores Range and returns the whole file.
  const res = await fetch(url, { headers: { Range: "bytes=0-33" } });
  if (!res.ok && res.status !== 206) { pngDimCache.set(url, "failed"); throw new Error(`fetch ${res.status}`); }
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.length < 24 || buf.toString("ascii", 1, 4) !== "PNG") { pngDimCache.set(url, "failed"); throw new Error("not a PNG"); }
  const dim = { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  pngDimCache.set(url, dim);
  return dim;
}

async function fixStripImageDimensions(html, variant, label = "") {
  if (variant !== "aiInFocus") return { html, fixed: 0 };

  // Collect the distinct strip-image URLs actually present, and fetch their
  // native dimensions in parallel (cached across generations).
  const urlRe = /<img\b[^>]*\bsrc="([^"]*article-\d+-top-s\d+\.png)"[^>]*>/gi;
  const urls = new Set();
  let m;
  while ((m = urlRe.exec(html)) !== null) urls.add(m[1]);
  if (urls.size === 0) return { html, fixed: 0 };

  const dims = new Map();
  await Promise.all([...urls].map(async (url) => {
    try { dims.set(url, await fetchPngDimensions(url)); }
    catch (e) { console.warn(`[design] strip-image dimensions unavailable for ${url}: ${e.message}`); }
  }));
  if (dims.size === 0) return { html, fixed: 0 };

  // Rewrite each "<td …><img … strip.png …>" pair: patch the img's width/height
  // (attribute + inline style) and the enclosing td's width/height hint.
  let fixed = 0;
  const out = html.replace(
    /(<td\b[^>]*>)(\s*)(<img\b[^>]*\bsrc="([^"]*article-\d+-top-s\d+\.png)"[^>]*>)/gi,
    (whole, tdTag, ws, imgTag, url) => {
      const d = dims.get(url);
      if (!d) return whole;
      const { w, h } = d;
      const newImg = imgTag
        .replace(/\bwidth="\d+"/i, `width="${w}"`)
        .replace(/\bheight="\d+"/i, `height="${h}"`)
        .replace(/(?<!-)\bwidth:\s*\d+px/i, `width: ${w}px`)
        .replace(/(?<!-)\bheight:\s*\d+px/i, `height: ${h}px`);
      const newTd = tdTag
        .replace(/(?<!-)\bwidth:\s*\d+px/i, `width: ${w}px`)
        .replace(/(?<!-)\bheight:\s*\d+px/i, `height: ${h}px`);
      fixed++;
      return newTd + ws + newImg;
    }
  );
  if (fixed) console.log(`[design] corrected strip-image dimensions on ${fixed} image(s)${label ? ` for "${label}"` : ""}`);
  return { html: out, fixed };
}

// Detect the BODY region (between header-end and footer-start), returning
// { head, body, tail }. Body is what the model may re-map; head and tail are
// preserved verbatim. Uses a prioritized set of markers because the templates
// don't share one convention (per the brand-guide sectionMarkers gaps):
//   1. START MAIN CONTENT ... END MAIN CONTENT   (AI in Focus)
//   2. END HEADER ... START FOOTER               (stated future convention)
//   3. END PREHEADER/END HEADER ... START FOOTER
//   4. fallback: first hero/banner block end ... footer boundary
// The footer boundary itself falls back through START FOOTER →
// "Black Footer Bar : BEGIN" (present in all six templates). Run this AFTER
// protectFragileRegions, so the footer is already a token we can locate.
function detectBodyRegion(html) {
  const find = (re) => { const i = html.search(re); return i === -1 ? null : i; };

  // Body START. Every template now carries <!-- START HEADER -->/<!-- END
  // HEADER -->, so END HEADER is the canonical, clean cut point. The rest are
  // defensive fallbacks for any template that somehow lacks it.
  const endHeader = find(/<!--\s*END\s+HEADER\s*-->/i);
  const startMain = find(/<!--\s*START MAIN CONTENT\s*-->/i);
  const endPre    = find(/<!--\s*END\s+PREHEADER\s*-->/i);
  const a2Begin   = find(/<!--\s*A2 \(callout block area\)\s*:\s*BEGIN\s*-->/i);
  const greeting  = find(/(?:Hi|Hello)\s*\{\{Recipient\.FirstName\}\}/i);

  let bodyStart, startMarkerLen = 0;
  const setStart = (idx, re) => { bodyStart = idx; startMarkerLen = re ? html.slice(idx).match(re)[0].length : 0; };
  if      (endHeader !== null) setStart(endHeader, /<!--\s*END\s+HEADER\s*-->/i);
  else if (startMain !== null) setStart(startMain, /<!--\s*START MAIN CONTENT\s*-->/i);
  else if (endPre    !== null) setStart(endPre,    /<!--\s*END\s+PREHEADER\s*-->/i);
  else if (a2Begin   !== null) setStart(a2Begin,   /<!--\s*A2 \(callout block area\)\s*:\s*BEGIN\s*-->/i);
  else if (greeting  !== null) {
    const before = html.slice(0, greeting);
    const cellStart = before.lastIndexOf("<td");
    bodyStart = cellStart !== -1 ? cellStart : greeting; startMarkerLen = 0;
  }
  else return null;

  // Body END = START FOOTER (now standard). protectFragileRegions runs first,
  // so the footer is a token sitting exactly where START FOOTER was. Fallbacks:
  // the literal START FOOTER marker, or the Black Footer Bar begin comment.
  const footerTok  = find(/<!--#PROTECTED:footer/i);
  const startFoot  = find(/<!--\s*START\s+FOOTER\s*-->/i);
  const blackBar   = find(/<!--\s*Black Footer Bar\s*:\s*BEGIN\s*-->/i);
  const bodyEnd = footerTok ?? startFoot ?? blackBar;
  if (bodyEnd === null || bodyEnd <= bodyStart) return null;

  return {
    head: html.slice(0, bodyStart + startMarkerLen),
    body: html.slice(bodyStart + startMarkerLen, bodyEnd),
    tail: html.slice(bodyEnd),
  };
}

// Within a body region, isolate the repeating ARTICLE structure and replace it
// with a single <!--#ARTICLES#--> marker, returning the canonical block for
// structure-locking. The non-article parts of the body remain in the returned
// bodyWithMarker for the model to re-map freely (hybrid mode).
function splitBodyArticles(body, variant) {
  if (variant === "aiInFocus") {
    const marker = /<!--\s*START ARTICLE\s*-->/i;
    const firstIdx = body.search(marker);
    if (firstIdx === -1) return { canonicalBlock: null, bodyWithMarker: body };

    // IMPORTANT structure note: all AI-in-Focus articles live inside ONE shared
    // row+cell that opens at "ARTICLE MODULE START" (<tr><td border="0"
    // valign="top">) and whose matching </td></tr> sits after END MAIN CONTENT.
    // That shared cell (together with the START MAIN CONTENT column cell) is
    // what constrains the articles to the 680px centered column — each article's
    // own inner <table width:100%> is correct BECAUSE it's nested in that cell.
    //
    // So the shared wrapper must be PRESERVED, not regenerated: keep its opener
    // on the intro side and its closer on the trailing side, and let the marker
    // sit between them. The repeating unit the model reproduces is therefore
    // only the inner <!-- START ARTICLE --><table>…</table> block — NOT a new
    // row/cell.
    const regionStart = firstIdx;                    // keep <tr><td> opener in intro
    const endMainIdx  = body.search(/<!--\s*END MAIN CONTENT\s*-->/i);
    const regionEnd   = endMainIdx !== -1 ? endMainIdx : body.length; // </td></tr> stays in trailing

    const intro    = body.slice(0, regionStart);     // includes ARTICLE MODULE START + <tr><td>
    const region   = body.slice(firstIdx, regionEnd); // the stacked inner article tables
    const trailing = body.slice(regionEnd);           // END MAIN CONTENT + </td></tr> + spacer

    const secondIdx = region.slice(1).search(marker);
    const canonicalBlock = secondIdx === -1 ? region : region.slice(0, secondIdx + 1);

    const bodyWithMarker = intro + "<!--#ARTICLES#-->\n" + trailing;
    return { canonicalBlock, bodyWithMarker };
  }

  if (variant === "aicpaTownHall") {
    const rowRe = /<tr>(?:(?!<\/tr>)[\s\S])*?highlight-img[\s\S]*?<\/tr>/i;
    const firstRow = body.match(rowRe);
    if (!firstRow) return { canonicalBlock: null, bodyWithMarker: body };

    const allRowsRe = /<tr>(?:(?!<\/tr>)[\s\S])*?highlight-img[\s\S]*?<\/tr>/gi;
    let replaced = false;
    const bodyWithMarker = body.replace(allRowsRe, () => {
      if (replaced) return "";
      replaced = true;
      return "<!--#ARTICLES#-->";
    });
    return { canonicalBlock: firstRow[0], bodyWithMarker };
  }

  return { canonicalBlock: null, bodyWithMarker: body };
}

const DESIGN_SYSTEM_ADDENDUM = `
DESIGN / REPEATING-BLOCK TEMPLATE MODE (HYBRID BODY MAPPING):
You are re-populating the BODY of a newsletter — the region between the header
and the footer. The header and footer are handled outside your output; do NOT
reproduce them.

You are given:
  1. A BODY SKELETON: the current body, containing SAMPLE/placeholder content
     and exactly one <!--#ARTICLES#--> marker where the repeating article
     blocks belong.
  2. ONE CANONICAL ARTICLE BLOCK: the exact, structure-locked layout each
     article must use.
  3. SOURCE CONTENT: the real content, extracted from a document, in its
     natural top-to-bottom order.

HOW TO MAP (this is the core task):
- Read the SOURCE CONTENT top-to-bottom and read the BODY SKELETON top-to-bottom.
  Compare their layouts and decide, by best judgment, which piece of source
  content corresponds to which region of the skeleton.
- Do NOT expect the source to contain placement tokens or codes. It usually
  will not. Infer placement from meaning and order: an opening greeting/intro
  paragraph maps to the skeleton's intro region; a list of links/resources maps
  to a list region; the repeating articles map to the article region.

TWO KINDS OF REGION, TWO DIFFERENT RULES:

A) NON-ARTICLE REGIONS (intro, resource lists, any prose/list in the skeleton
   that is NOT the <!--#ARTICLES#--> marker):
   - You MAY re-map these with judgment: replace the sample text with the
     corresponding source content, and adjust list items to match the source
     (add/remove <li> items as the source requires).
   - PRESERVE the surrounding HTML table structure, tags, classes, and inline
     styles. Change text and list items — not the scaffolding.
   - Leave any {{...}} / {{{...}}} merge tokens exactly as they are.

B) THE ARTICLE REGION (the <!--#ARTICLES#--> marker):
   - Replace the single marker with ONE canonical block per source article.
   - Reproduce the CANONICAL ARTICLE BLOCK structure EXACTLY per article,
     changing ONLY the title text and body copy. Keep every style attribute,
     table, class, and comment identical to the canonical block.
   - Begin EACH article block with the <!-- START ARTICLE --> comment exactly
     as it appears at the top of the canonical block.
   - Use as many blocks as the source has articles — no more, no fewer.

PRESERVE TABLE ATTRIBUTES EXACTLY:
- When you reproduce the canonical article block, copy EVERY opening tag
  verbatim — including width, align, valign, height, bgcolor, and the full
  style attribute. Do NOT normalize, simplify, or "tidy" them. In particular:
  - Never change a table's width (e.g. do not rewrite width="680" or
    width:680px to width:100%, and vice-versa). Keep whatever the canonical
    block has, character-for-character.
  - Never drop align="center" (or align="right"/valign) from a tag that has it.
  - Keep margin:auto and any other alignment styles intact.
  These attributes are what keep each block aligned with the surrounding
  680px centered column; altering them makes blocks misalign.
- The <!--#ARTICLES#--> marker is already positioned INSIDE the correct
  wrapping row/cell. Do NOT add a new <tr>, <td>, or wrapper table around your
  article blocks — emit only the inner block structure exactly as the canonical
  block shows, starting at its <!-- START ARTICLE --> comment.

PROTECTED TOKENS:
- Tokens of the form <!--#PROTECTED:...#--> are opaque placeholders for fragile
  pre-built markup. Copy them through verbatim wherever they appear. NEVER
  alter, expand, remove, or reorder the characters inside a protected token.

DO NOT MANAGE NUMBERING:
- Do NOT renumber image paths, element ids, or link anchors between blocks
  (article-N paths, id="artN", #artN). Copy them from the canonical block as-is.
  Sequential numbering is applied automatically after you finish.

OUTPUT:
- Output ONLY the re-mapped BODY (starting at the first element of the skeleton,
  ending at the last). Do NOT include the header or footer. Raw HTML only,
  nothing before the first < or after the last >.

NEVER invent facts, URLs, or statistics not present in the source or instructions.`;

async function generateDesignHTML(ticket, templateName, templateHtml, sourceContent, deps) {
  const { anthropic, AI_SYSTEM_PROMPT, extractHtml, logUsage, getBrandGuide } = deps;
  const variant = designVariant(templateName);

  // 1. Protect fragile regions FIRST (footer becomes a token we can locate).
  const { html: protectedHtml, store } = protectFragileRegions(templateHtml);

  // 2. Isolate the BODY region (header-end → footer-start). Head + tail are
  //    preserved verbatim; only the body is re-mapped.
  const region = detectBodyRegion(protectedHtml);
  if (!region) {
    const err = new Error(`Could not detect body region for "${templateName}" — falling back to standard generation.`);
    err.code = "NO_CANONICAL_BLOCK";
    throw err;
  }

  // 3. Within the body, lock the article structure behind an #ARTICLES# marker
  //    and keep the non-article body editable (hybrid mode).
  const { canonicalBlock, bodyWithMarker } = splitBodyArticles(region.body, variant);
  if (!canonicalBlock) {
    const err = new Error(`Could not extract canonical article block for "${templateName}" — falling back to standard generation.`);
    err.code = "NO_CANONICAL_BLOCK";
    throw err;
  }

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
        `\nCANONICAL ARTICLE BLOCK (structure-locked; reproduce exactly per ` +
        `article, protected tokens verbatim):\n\n${canonicalBlock}\n\n` +
        `BODY SKELETON (re-map source content onto this; replace the ` +
        `<!--#ARTICLES#--> marker with one canonical block per source article):\n\n${bodyWithMarker}\n\n` +
        `SOURCE CONTENT (top-to-bottom; map by meaning and order):\n${sourceContent || "(use the description below)"}\n\n` +
        `Description: ${ticket.description || ""}\n` +
        (ticket.__freeform__ ? `\nAdditional requestor instructions:\n${ticket.__freeform__}\n` : ""),
    }],
  });

  const raw = message.content.find(b => b.type === "text")?.text ?? "";
  logUsage(`generate-design "${ticket.name}"`, message);
  const mappedBody = extractHtml(raw, bodyWithMarker);

  // 4. Reassemble: head (verbatim) + mapped body + tail (verbatim, footer token).
  let html = region.head + "\n" + mappedBody + "\n" + region.tail;

  const corrupted = findCorruptedTokens(html, store);
  if (corrupted.length > 0) {
    console.warn(`[design] ${corrupted.length} corrupted protected token(s) in "${ticket.name}": ${corrupted.join(", ")}`);
  }
  if (html.includes("<!--#ARTICLES#-->")) {
    console.warn(`[design] Model left #ARTICLES# marker unreplaced for "${ticket.name}" — restoring what we can.`);
  }

  // 5. Restore protected regions byte-for-byte.
  const { html: restored, restored: n, missing, total } = restoreFragileRegions(html, store);
  console.log(`[design] Protected regions restored: ${n}/${total} (${missing} intentionally dropped with removed blocks)`);

  // 6. Stamp sequential per-article numbering (AI-in-Focus only; no-op elsewhere).
  const { html: renumbered, count, warning } = renumberArticles(restored, variant);
  if (warning) console.warn(`[design] renumber for "${ticket.name}": ${warning}`);
  else if (count) console.log(`[design] renumbered ${count} article(s) sequentially`);

  // 7. Correct each article's header-strip image dimensions to the real asset
  //    sizes (they differ per article; renumber only swaps the filename).
  const { html: finalHtml } = await fixStripImageDimensions(renumbered, variant, ticket.name);

  return finalHtml;
}

// ═════════════════════════════════════════════
// RAG few-shot retrieval layer (GitHub-hosted manifest + brand guide)
// ═════════════════════════════════════════════
let manifestCache   = null;
let brandGuideCache = null;

const GITHUB_API = "https://api.github.com";
const GH_OWNER  = process.env.GITHUB_OWNER;
const GH_REPO   = process.env.GITHUB_REPO;
const GH_BRANCH = process.env.GITHUB_BRANCH || "main";

async function githubGetFile(p) {
  const res = await fetch(
    `${GITHUB_API}/repos/${GH_OWNER}/${GH_REPO}/contents/${encodeURIComponent(p)}?ref=${GH_BRANCH}`,
    { headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}`, Accept: "application/vnd.github+json" } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub read failed for ${p}: ${res.status} ${res.statusText}`);
  const json = await res.json();
  const content = Buffer.from(json.content, "base64").toString("utf-8");
  return { content: JSON.parse(content), sha: json.sha };
}

async function githubPutFile(p, data, message, knownSha = null) {
  const sha = knownSha ?? (await githubGetFile(p))?.sha ?? null;
  const res = await fetch(
    `${GITHUB_API}/repos/${GH_OWNER}/${GH_REPO}/contents/${encodeURIComponent(p)}`,
    {
      method: "PUT",
      headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}`, Accept: "application/vnd.github+json", "Content-Type": "application/json" },
      body: JSON.stringify({
        message,
        content: Buffer.from(JSON.stringify(data, null, 2)).toString("base64"),
        branch: GH_BRANCH,
        ...(sha ? { sha } : {}),
      }),
    }
  );
  if (res.status === 409 || res.status === 422) {
    console.warn(`GitHub write conflict on ${p}, refetching sha and retrying once.`);
    const fresh = await githubGetFile(p);
    return githubPutFileRetry(p, data, message, fresh?.sha ?? null);
  }
  if (!res.ok) throw new Error(`GitHub write failed for ${p}: ${res.status} ${res.statusText}`);
  const json = await res.json();
  return json.content.sha;
}

async function githubPutFileRetry(p, data, message, sha) {
  const res = await fetch(
    `${GITHUB_API}/repos/${GH_OWNER}/${GH_REPO}/contents/${encodeURIComponent(p)}`,
    {
      method: "PUT",
      headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}`, Accept: "application/vnd.github+json", "Content-Type": "application/json" },
      body: JSON.stringify({
        message,
        content: Buffer.from(JSON.stringify(data, null, 2)).toString("base64"),
        branch: GH_BRANCH,
        ...(sha ? { sha } : {}),
      }),
    }
  );
  if (!res.ok) throw new Error(`GitHub write retry failed for ${p}: ${res.status} ${res.statusText}`);
  const json = await res.json();
  return json.content.sha;
}

async function getManifest() {
  if (manifestCache && Date.now() - manifestCache.fetchedAt < CACHE_TTL_MS) {
    return manifestCache.data;
  }
  const p = process.env.GITHUB_MANIFEST_PATH || "manifest.json";
  const remote = await githubGetFile(p);
  let data = remote?.content ?? { version: 0, updated: null, templates: [] };
  if (!remote) console.log("No manifest.json in GitHub repo yet — starting from empty manifest.");
  data = await ensureManifestUpToDate(data);
  manifestCache = { data, fetchedAt: Date.now() };
  console.log(`Template manifest refreshed: ${data.templates.length} entries`);
  return data;
}

async function ensureManifestUpToDate(manifest) {
  await fetchTemplateIndex({ includeNonHtml: false });
  const knownFilenames = new Set(manifest.templates.map(t => t.filename));
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
  await writeManifestToGithub(updated);
  return updated;
}

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
          `Template HTML:\n${html.slice(0, 12000)}`,
      },
    ],
  });
  const text = response.content.find(b => b.type === "text")?.text ?? "{}";
  let parsed;
  try { parsed = JSON.parse(text.replace(/```json|```/g, "").trim()); } catch { parsed = {}; }
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

let githubWriteSupported = true;

async function writeManifestToGithub(manifest) {
  if (!githubWriteSupported) return;
  const p = process.env.GITHUB_MANIFEST_PATH || "manifest.json";
  try {
    await githubPutFile(p, manifest, `auto-tag: ${manifest.templates.length} templates (v${manifest.version})`);
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
  const p = process.env.GITHUB_BRAND_GUIDE_PATH || "brand-guide.json";
  const remote = await githubGetFile(p);
  if (!remote) throw new Error(`brand-guide.json not found in GitHub repo at path "${p}".`);
  brandGuideCache = { data: remote.content, fetchedAt: Date.now() };
  console.log("Brand guide refreshed from GitHub.");
  return remote.content;
}

function scoreTemplateAgainstText(template, text) {
  const haystack = text.toLowerCase();
  let score = 0;
  for (const tag of template.tags) {
    if (haystack.includes(tag.toLowerCase())) score += 3;
  }
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
  const confident = topScore >= 4 && topScore - runnerUpScore >= 2;
  return { confident, matches: scored.slice(0, 3).map(s => s.template) };
}

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
  try { parsed = JSON.parse(text.replace(/```json|```/g, "").trim()); } catch { return []; }
  const ids = new Set(parsed.templateIds || []);
  return manifest.templates.filter(t => ids.has(t.id));
}

async function selectRelevantTemplates(ticket) {
  const manifest = await getManifest();
  const ticketText = [ticket.description, ticket.subjectLine, ticket.templateDropdownValue, ticket.groupName]
    .filter(Boolean).join("\n");
  const tier1 = classifyByKeyword(manifest, ticketText);
  if (tier1.confident) {
    console.log(`Template classification (keyword, confident): ${tier1.matches.map(t => t.id).join(", ")}`);
    return tier1.matches;
  }
  console.log("Template classification: keyword match ambiguous, falling back to model.");
  const modelMatches = await classifyWithModel(manifest, ticketText);
  if (modelMatches.length > 0) return modelMatches;
  return tier1.matches;
}

function stripFooterForFewShot(html) {
  return html.replace(/<!--\s*START FOOTER\s*-->[\s\S]*?<!--\s*END FOOTER\s*-->/i, "<!-- FOOTER OMITTED: server reattaches actual footer verbatim -->");
}

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
    { type: "text", text: staticBrandRules },
    { type: "text", text: "BRAND GUIDE (structured):\n" + JSON.stringify(brandGuide, null, 2), cache_control: { type: "ephemeral" } },
    {
      type: "text",
      text: selected.length > 0
        ? "RELEVANT TEMPLATE EXAMPLES (retrieved for this ticket, footers omitted — do not reproduce footer content from these):\n\n" + fewShotBlocks.join("\n\n---\n\n")
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
// Monday files
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

async function applyHeaderImage(html, itemId, vars = {}) {
  const headerLink     = (vars["HeaderLink"]  || "").trim();
  const headerFilename = (vars["HeaderImage"] || "").trim();
  if (!headerFilename && !headerLink) return html;
  try {
    if (headerFilename) {
      const assets = await fetchTicketFiles(itemId);
      const match  = assets.find(a => a.name.toLowerCase() === headerFilename.toLowerCase());
      if (!match) {
        console.warn(`Header image "${headerFilename}" not found in Files column for item ${itemId} — header left unchanged.`);
        return headerLink ? replaceFirstImage(html, null, headerLink) : html;
      }
      const publicUrl = match.public_url || match.url;
      if (!publicUrl) {
        console.warn(`Header image "${headerFilename}" has no public URL — header left unchanged.`);
        return headerLink ? replaceFirstImage(html, null, headerLink) : html;
      }
      console.log(`Header image set from Monday public_url (temporary): ${headerFilename}`);
      return replaceFirstImage(html, publicUrl, headerLink);
    } else {
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

function logUsage(label, message) {
  const u = message?.usage || {};
  console.log(
    `[usage] ${label} — in:${u.input_tokens ?? "?"} out:${u.output_tokens ?? "?"} ` +
    `cache_write:${u.cache_creation_input_tokens ?? 0} cache_read:${u.cache_read_input_tokens ?? 0}`
  );
}

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
      let html = await generateDesignHTML(ticket, templateName, templateHtml, sourceContent,
        { anthropic, AI_SYSTEM_PROMPT, extractHtml, logUsage, getBrandGuide });
      if (ticket.id && ticket.id !== "manual") html = await applyHeaderImage(html, ticket.id, vars);
      return html;
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
// Upload to Monday
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
  await uploadToMonday(itemId, stateFileName, blob, "application/json");
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
  try {
    if (event.type === "create_pulse") {
      const itemId = String(event.pulseId);
      console.log(`[agent] New item ${itemId} — generating first proof`);

      let ticket = await fetchItemById(itemId);

      if (!ticket.jobNumber) {
        console.log(`[agent] Job Number empty for ${itemId} — retrying once after delay`);
        await new Promise(r => setTimeout(r, 3000));
        ticket = await fetchItemById(itemId);
      }

      const hasInstructions = !!(ticket.instructions && ticket.instructions.trim());
      const hasDoc = await fetchWordDocContent(itemId);
      if (!hasInstructions && !hasDoc) {
        console.log(`[agent] Item ${itemId} has no Prompt and no Word doc — skipping generation`);
        return;
      }

      const templateName = await resolveTemplateName(ticket);
      const html         = await generateHTML(ticket, templateName);
      const fileName     = `${ticket.name.replace(/\s+/g, "_")}_${ticket.jobNumber || itemId}_v1.html`;

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
