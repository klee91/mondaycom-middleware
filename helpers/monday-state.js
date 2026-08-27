/**
 * helpers/monday-state.js — proof upload, agent-state persistence/recovery, ticket updates.
 */
const FormData = require("form-data");
const { fetch, MONDAY_FILE_URL, BOARD_ID, AGENT_STATE_COLUMN } = require("./config");
const { mondayQuery, fetchTicketFiles, downloadMondayAsset } = require("./monday");

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
    // add_file_to_column APPENDS, so multiple same-named state files can
    // accumulate. Always pick the MOST RECENT one, not the first .find() hit.
    const matches = assets
      .filter(a => a.name === meta.stateFile)
      .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));
    const match = matches[0];
    if (!match) return "";
    if (matches.length > 1) {
      console.warn(`[agent] ${matches.length} copies of ${meta.stateFile} in Files for ${itemId} — using newest (${match.created_at}). Older copies should be purged.`);
    }
    const buf  = await downloadMondayAsset(match);
    const blob = JSON.parse(buf.toString("utf-8"));
    const html = blob.currentHtml || "";
    // Guard: a corrupted/compounded blob can be enormous. Refuse to feed it back.
    if (html.length > 400_000) {
      console.warn(`[agent] recovered HTML for ${itemId} is ${html.length} chars — abnormally large, treating as unrecoverable.`);
      return "";
    }
    return html;
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

// Set the item's Status column to a label (e.g. "Proofing"). Uses
// change_simple_column_value, which for a status column matches the label text.
// The label must exist on the board's Status column exactly as given. Any
// failure is logged and swallowed so it never blocks proof delivery.
async function setStatus(itemId, label) {
  try {
    await mondayQuery(`
      mutation SetStatus($itemId: ID!, $boardId: ID!, $label: String!) {
        change_simple_column_value(item_id: $itemId, board_id: $boardId, column_id: "status", value: $label) { id }
      }
    `, { itemId, boardId: BOARD_ID, label });
    console.log(`[agent] Item ${itemId} status set to "${label}"`);
  } catch (err) {
    console.warn(`[agent] Could not set status to "${label}" for ${itemId}: ${err.message}`);
  }
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

module.exports = { uploadToMonday, readAgentMeta, readCurrentHtml, persistAgentState, postUpdate, setStatus, findUserIdByName };
