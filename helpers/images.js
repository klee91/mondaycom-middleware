/**
 * helpers/images.js — content (body) image placement from Files-column attachments.
 *
 * Requestors attach images to the ticket's Files column and describe, in the
 * prompt, where they want them ("add chart.png to the second article",
 * "replace the product shot with newphoto.png", etc). This module:
 *   1. lists the attached IMAGE assets (each carries a temporary Monday public_url),
 *   2. lists the REPLACEABLE <img> slots in the generated HTML (everything that
 *      isn't a protected structural image — article header strips, bracket
 *      buttons),
 *   3. asks the model to map attached images → slots by the requestor's intent
 *      (flexible: no assumed phrasing or label),
 *   4. deterministically swaps each targeted <img>'s src to the asset's
 *      public_url — the model never writes the URL, only chooses the mapping.
 *
 * Nothing is forced: images the requestor didn't ask to place are ignored, and
 * any slot without a confident mapping is left exactly as-is. Like the header
 * image, the public_url is TEMPORARY (proof only) — swap to a permanent asset
 * in Salesforce before send.
 */
const { anthropic, CHEAP_MODEL } = require("./config");
const { fetchAllItemFiles } = require("./monday");

const IMAGE_EXT_RE = /\.(png|jpe?g|gif|webp|svg)$/i;

// <img> whose src matches these are STRUCTURAL and must never be replaced by a
// requestor image (they're the template's own framing assets).
const PROTECTED_IMG_SRC = [
  /article-\d+-top-s\d+\.png/i, // AI-in-Focus article header strips
  /ai-nl-btn-(left|right)\.png/i, // bracket-button end caps
];

function isProtectedSrc(src) {
  return PROTECTED_IMG_SRC.some(re => re.test(src || ""));
}

// List replaceable content <img> slots with light context (alt text + a little
// surrounding text) so the model can tell them apart. Returns [{ src, alt,
// context }] in document order; protected/structural images are excluded.
function findReplaceableImages(html) {
  const slots = [];
  const imgRe = /<img\b[^>]*>/gi;
  let m;
  while ((m = imgRe.exec(html)) !== null) {
    const tag = m[0];
    const src = (tag.match(/\bsrc="([^"]*)"/i) || [])[1] || "";
    if (!src || isProtectedSrc(src)) continue;
    const alt = (tag.match(/\balt="([^"]*)"/i) || [])[1] || "";
    // grab ~120 chars of visible text just before the image for positional context
    const before = html.slice(Math.max(0, m.index - 400), m.index)
      .replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    const context = before.slice(-120);
    slots.push({ src, alt, context });
  }
  return slots;
}

// Apply a validated plan: [{ src, publicUrl }] — swap each matching <img>'s src.
// Exact src match; protected images are refused even if a plan names them.
function executeImagePlan(html, plan) {
  let applied = 0;
  let out = html;
  for (const { src, publicUrl } of plan) {
    if (!src || !publicUrl || isProtectedSrc(src)) continue;
    // Replace the src of the FIRST <img> whose src exactly equals `src`.
    const re = new RegExp(`(<img\\b[^>]*\\bsrc=")${src.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}("[^>]*>)`, "i");
    const next = out.replace(re, `$1${publicUrl}$2`);
    if (next !== out) { out = next; applied++; }
  }
  return { html: out, applied };
}

// Does the prompt reference any of the attached images / ask for image work?
// Cheap gate so we only spend a model call when there's real image intent.
function hasImageIntent(images, promptText) {
  const t = (promptText || "").toLowerCase();
  if (!t) return false;
  for (const a of images) {
    const name = (a.name || "").toLowerCase();
    const base = name.replace(IMAGE_EXT_RE, "");
    if (name && t.includes(name)) return true;
    if (base.length >= 3 && t.includes(base)) return true;
  }
  return false;
}

// Ask the model to map attached images → replaceable slots by the requestor's
// intent. Returns [{ image: <filename>, slotSrc: <exact current src> }].
// Validated by the caller against real assets + slots.
async function planImagePlacements(promptText, images, slots) {
  if (!images.length || !slots.length) return [];
  try {
    const response = await anthropic.messages.create({
      model: CHEAP_MODEL,
      max_tokens: 500,
      system:
        "You place a requestor's uploaded images into the image slots of a marketing email. " +
        "You are given the requestor's instructions, the list of ATTACHED IMAGE filenames, and the " +
        "list of replaceable image SLOTS (each with its current src and nearby text for context). " +
        "Map each image the requestor wants placed to the single best-matching slot, using their " +
        "instructions and the slot context (e.g. 'the second article', 'the product shot'). Do NOT " +
        "place an image the requestor didn't ask about, and do NOT map two images to the same slot. " +
        "Respond ONLY with JSON: {\"placements\":[{\"image\":\"<exact attached filename>\",\"slotSrc\":\"<exact current src of the chosen slot>\"}]}. " +
        "Copy filenames and srcs verbatim from the lists. If you can't confidently place an image, omit it.",
      messages: [{
        role: "user",
        content:
          `Requestor instructions:\n${promptText || "(none)"}\n\n` +
          `ATTACHED IMAGES:\n${images.map(a => `- ${a.name}`).join("\n")}\n\n` +
          `REPLACEABLE SLOTS (in document order):\n${slots.map((s, i) => `- [slot ${i + 1}] src="${s.src}" alt="${s.alt}" context="…${s.context}"`).join("\n")}`,
      }],
    });
    const text = response.content.find(b => b.type === "text")?.text ?? "{}";
    let parsed;
    try { parsed = JSON.parse(text.replace(/```json|```/g, "").trim()); } catch { return []; }
    return Array.isArray(parsed.placements) ? parsed.placements : [];
  } catch (err) {
    console.warn(`[images] placement planning skipped: ${err.message}`);
    return [];
  }
}

// Entry point: place requestor-attached content images into the HTML body.
// Non-blocking; returns the html unchanged on any miss or failure.
async function applyContentImages(html, itemId, vars = {}) {
  if (!itemId || itemId === "manual") return html;
  try {
    const assets = await fetchAllItemFiles(itemId);
    const images = assets.filter(a => IMAGE_EXT_RE.test(a.name || "") && (a.public_url || a.url));
    if (images.length === 0) return html;

    const promptText = [vars["__freeform__"], vars["BodyContent"], vars["HeaderImage"]].filter(Boolean).join("\n");
    if (!hasImageIntent(images, promptText)) return html; // no image request → don't force anything

    const slots = findReplaceableImages(html);
    if (slots.length === 0) return html;

    // Deterministic fast path: exactly one referenced image + one slot → no model call.
    const referenced = images.filter(a => {
      const n = (a.name || "").toLowerCase(); const b = n.replace(IMAGE_EXT_RE, "");
      const t = promptText.toLowerCase();
      return t.includes(n) || (b.length >= 3 && t.includes(b));
    });
    let plan = [];
    if (referenced.length === 1 && slots.length === 1) {
      plan = [{ image: referenced[0].name, slotSrc: slots[0].src }];
    } else {
      plan = await planImagePlacements(promptText, images, slots);
    }
    if (!plan.length) return html;

    // Validate the plan against real assets + real slots, resolve public_urls.
    const assetByName = new Map(images.map(a => [(a.name || "").toLowerCase(), a]));
    const slotSrcs = new Set(slots.map(s => s.src));
    const usedSlots = new Set();
    const validated = [];
    for (const p of plan) {
      const asset = assetByName.get((p.image || "").toLowerCase());
      if (!asset) continue;                              // unknown filename → drop (no invented images)
      if (!slotSrcs.has(p.slotSrc) || usedSlots.has(p.slotSrc)) continue; // unknown/duplicate slot → drop
      const publicUrl = asset.public_url || asset.url;
      if (!publicUrl) continue;
      usedSlots.add(p.slotSrc);
      validated.push({ src: p.slotSrc, publicUrl });
    }
    if (!validated.length) return html;

    const { html: out, applied } = executeImagePlan(html, validated);
    if (applied) console.log(`[images] placed ${applied} content image(s) from Files column (temporary public_url — swap before send)`);
    return out;
  } catch (err) {
    console.warn(`[images] content image placement skipped for ${itemId}: ${err.message}`);
    return html;
  }
}

module.exports = { applyContentImages, findReplaceableImages, executeImagePlan, hasImageIntent, isProtectedSrc, planImagePlacements };
