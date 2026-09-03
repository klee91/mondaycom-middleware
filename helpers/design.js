/**
 * helpers/design.js — repeating-block (AI-in-Focus / Town Hall) generation:
 * template detection, protected-region round-trip, article renumber,
 * strip-image dimension correction, CTA-link injection, body-region mapping.
 */
const { fetch, anthropic, GEN_MODEL, CHEAP_MODEL, AI_SYSTEM_PROMPT } = require("./config");
const { getBrandGuide } = require("./github-store");
const { logUsage, extractHtml } = require("./utils");

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

// ── Per-article CTA link injection (AI-in-Focus) ──
// Each article's CTA is a PROTECTED bracket button, so the model can't write the
// destination into it — after restore+renumber every button points at the
// template placeholder anchor (…/ai-in-focus/05-26#artN). The real destinations
// live in the source doc as one link per article.
//
// We must NOT assume the CTA label — it may be "Deep dive", "Read more",
// "Learn more", "Full story", or unlabeled, and it varies by request. So rather
// than string-match a label, we let the model read the doc top-to-bottom and
// report each article's CTA URL in order (its judgment about which link is the
// story link), then deterministically stamp those onto the buttons by position.
//
// Safety: only URLs that actually appear in the source doc are injected, so the
// model can never introduce an invented destination (per the brand rule).

// Map every href in the source to its exact authored form, keyed by an
// entity-normalized version, so a model-returned URL can be matched back to the
// precise string to inject (preserving &amp; etc.).
function docHrefIndex(sourceContent) {
  const map = new Map();
  for (const m of (sourceContent || "").matchAll(/href="([^"]*)"/gi)) {
    const norm = m[1].replace(/&amp;/gi, "&").trim();
    if (norm && !map.has(norm)) map.set(norm, m[1]);
  }
  return map;
}

// Ask the model which link is each article's CTA, in article order. Flexible by
// design: no reliance on the link's label text. Returns URLs in the doc's exact
// authored form, validated to exist in the doc; unknown/empty slots become "".
async function extractArticleCtaUrls(sourceContent, articleCount, label = "") {
  if (!sourceContent || !articleCount) return [];
  if (!/<a\b[^>]*href=/i.test(sourceContent)) return []; // no links → nothing to map
  const hrefIndex = docHrefIndex(sourceContent);
  if (hrefIndex.size === 0) return [];

  try {
    const response = await anthropic.messages.create({
      model: CHEAP_MODEL,
      max_tokens: 600,
      system:
        "You identify the primary call-to-action (CTA) link for each article in a newsletter " +
        "source document — the link a reader follows to read that article's full story. Do NOT " +
        "rely on the link's wording: the CTA may be labeled 'Deep dive', 'Read more', 'Learn more', " +
        "'Full story', an arrow, or nothing consistent. Judge by role and position, not text. " +
        "Links in an introduction or a resources/round-up list that appear BEFORE the articles are " +
        "NOT article CTAs — exclude them. Return ONLY JSON: {\"ctaUrls\": [\"url\", ...]} containing " +
        "exactly one entry per article, in the same top-to-bottom order the articles appear. Copy " +
        "each URL verbatim from the document. If a given article has no CTA link, use \"\" for that " +
        "slot to preserve order. Never invent a URL.",
      messages: [{
        role: "user",
        content:
          `This newsletter has ${articleCount} article(s), in order.\n\n` +
          `Source document:\n${sourceContent.slice(0, 16000)}`,
      }],
    });
    const text = response.content.find(b => b.type === "text")?.text ?? "{}";
    let parsed;
    try { parsed = JSON.parse(text.replace(/```json|```/g, "").trim()); } catch { return []; }
    const raw = Array.isArray(parsed.ctaUrls) ? parsed.ctaUrls : [];
    // Validate each against the doc; map to the exact authored href or drop.
    return raw.map(u => {
      if (typeof u !== "string" || !u.trim()) return "";
      const norm = u.replace(/&amp;/gi, "&").trim();
      return hrefIndex.get(norm) || "";
    });
  } catch (err) {
    console.warn(`[design] CTA URL identification skipped${label ? ` for "${label}"` : ""}: ${err.message}`);
    return [];
  }
}

// Stamp the per-article CTA URLs (already in article order) onto the protected
// bracket buttons, replacing the #artN placeholder anchor.
function applyArticleCtaLinks(html, ctaUrls, variant, label = "") {
  if (variant !== "aiInFocus") return { html, applied: 0 };
  if (!ctaUrls || ctaUrls.length === 0) return { html, applied: 0 };

  const parts = html.split(/<!--\s*START ARTICLE\s*-->/i);
  if (parts.length < 2) return { html, applied: 0 };

  let rebuilt = parts[0]; // head, untouched
  let applied = 0;
  for (let k = 1; k < parts.length; k++) {
    let seg = parts[k];
    const url = ctaUrls[k - 1];
    if (url) {
      // Replace ONLY the bracket-button anchor hrefs (…#artN). All three <a>
      // tags in the button share that href, so this catches the full button.
      const before = seg;
      seg = seg.replace(/href="[^"]*#art\d+"/gi, `href="${url}"`);
      if (seg !== before) applied++;
    }
    rebuilt += "<!-- START ARTICLE -->" + seg;
  }

  const articleCount = parts.length - 1;
  const provided = ctaUrls.filter(Boolean).length;
  if (applied < articleCount) {
    console.warn(`[design] applied ${applied}/${articleCount} article CTA link(s)${label ? ` for "${label}"` : ""} (${provided} URL(s) identified) — remaining buttons left on placeholder anchor`);
  } else if (applied) {
    console.log(`[design] applied ${applied} article CTA link(s)${label ? ` for "${label}"` : ""}`);
  }
  return { html: rebuilt, applied };
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

async function generateDesignHTML(ticket, templateName, templateHtml, sourceContent) {
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
    model: GEN_MODEL,
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
  const { html: sized } = await fixStripImageDimensions(renumbered, variant, ticket.name);

  // 8. Inject each article's CTA destination from the source doc into the
  //    protected bracket button (replacing the #artN placeholder). The CTA link
  //    is identified by role/position, not by any assumed label text.
  const articleCount = (sized.match(/<!--\s*START ARTICLE\s*-->/gi) || []).length;
  const ctaUrls = await extractArticleCtaUrls(sourceContent, articleCount, ticket.name);
  const { html: finalHtml } = applyArticleCtaLinks(sized, ctaUrls, variant, ticket.name);

  return finalHtml;
}

module.exports = { DESIGN_TEMPLATES, normalizeTemplateName, designVariant, isDesignTemplate, protectFragileRegions, restoreFragileRegions, findCorruptedTokens, renumberArticles, fetchPngDimensions, fixStripImageDimensions, docHrefIndex, extractArticleCtaUrls, applyArticleCtaLinks, detectBodyRegion, splitBodyArticles, generateDesignHTML };
