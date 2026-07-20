/**
 * helpers/monday.js — Monday GraphQL, ticket/file reads, Word-doc extraction, header image.
 */
const mammoth = require("mammoth");
const { fetch, MONDAY_API_URL, BOARD_ID, AGENT_STATE_COLUMN, INSTRUCTIONS_COLUMN, TEMPLATE_COLUMN } = require("./config");

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
      items(ids: [$itemId]) { assets { id name url public_url file_extension created_at } }
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

    // CRITICAL: by default mammoth inlines embedded images as base64 data URIs.
    // A newsletter doc with a few pasted images becomes MILLIONS of tokens
    // (one real doc measured ~749K tokens, ~746K of it base64) and blows the
    // model's context window. The doc is only a source of TEXT and LINKS — the
    // email uses hosted assets, not whatever was pasted into Word — so drop the
    // image data entirely instead of embedding it.
    const result = await mammoth.convertToHtml({ buffer }, {
      convertImage: mammoth.images.imgElement(() => ({ src: "" })), // no base64
    });
    let html = (result.value || "").trim();
    // Belt-and-suspenders: strip any base64 data URI that slipped through, and
    // drop now-empty <img> tags so they don't clutter the source content.
    html = html
      .replace(/\s*src="data:[^"]*"/gi, ' src=""')
      .replace(/<img\b[^>]*\bsrc=""[^>]*>/gi, "");

    if (!html) {
      console.warn(`Word doc "${doc.name}" produced no content.`);
      return null;
    }
    const approxTokens = Math.round(html.length / 4);
    console.log(`Extracted Word doc "${doc.name}" (${html.length} chars, ~${approxTokens} tokens after image strip)`);
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

module.exports = { mondayQuery, normalizeTicket, fetchItemById, fetchTicketFiles, downloadMondayAsset, fetchWordDocContent, replaceFirstImage, applyHeaderImage };