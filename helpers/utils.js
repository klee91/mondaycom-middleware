/**
 * helpers/utils.js — tiny dependency-free helpers shared across modules.
 */
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
  
  module.exports = { logUsage, extractHtml };