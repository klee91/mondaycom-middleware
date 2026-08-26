/**
 * helpers/sharepoint.js — SharePoint (Graph client-credentials) auth + template library.
 */
const { fetch, CACHE_TTL_MS } = require("./config");

// module-local caches
let sharepointToken = null;
let tokenExpiresAt  = 0;
const templateHtmlCache = {};  // { itemId: { html, fetchedAt } }
let   templateIndexCache = null; // { map: { displayName: itemId }, fetchedAt }

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

module.exports = { getSharePointToken, fileNameToDisplayName, fetchTemplateIndex, fetchTemplateFromSharePoint };
