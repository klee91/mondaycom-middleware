/**
 * helpers/config.js — shared clients, env-derived config, and constants.
 * Every other helper imports from here so there's a single source of truth
 * for the Anthropic client, fetch, Monday column IDs, and the system prompt.
 */
const fetch     = require("node-fetch").default || require("node-fetch");
const Anthropic = require("@anthropic-ai/sdk").default;

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Models (kept here so a swap is one edit)
const GEN_MODEL   = "claude-opus-4-6";
const CHEAP_MODEL = "claude-haiku-4-5-20251001";

// Monday
const MONDAY_API_URL  = "https://api.monday.com/v2";
const MONDAY_FILE_URL = "https://api.monday.com/v2/file";
const BOARD_ID        = process.env.BOARD_ID;

// Monday column IDs
const AGENT_STATE_COLUMN  = "long_text_mm4b1t9h"; // Agent State column
const INSTRUCTIONS_COLUMN = "long_text_mm47njms"; // Prompt column (display name may be "Prompt"; ID unchanged)
const TEMPLATE_COLUMN     = "dropdown_mm4d2e9v";  // Template dropdown column
// Proof History: a long-text column that accumulates one line per uploaded
// version (timestamp + filename + asset link). Set PROOF_HISTORY_COLUMN in the
// environment to the real column ID; if unset, proof-history recording is a
// no-op (never blocks generation).
const PROOF_HISTORY_COLUMN = process.env.PROOF_HISTORY_COLUMN || "";

const CONTENT_VARIABLES = ["PreheaderText", "Subject", "Partner"];

const BUTTON_COLORS = {
  "purple": "#72246c", "light purple": "#86387f", "green": "#48a23f", "navy": "#0f206c",
};

// Shared cache TTL used by SharePoint + GitHub-store caches
const CACHE_TTL_MS = 10 * 60 * 1000;

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

module.exports = {
  fetch, anthropic, GEN_MODEL, CHEAP_MODEL,
  MONDAY_API_URL, MONDAY_FILE_URL, BOARD_ID,
  AGENT_STATE_COLUMN, INSTRUCTIONS_COLUMN, TEMPLATE_COLUMN, PROOF_HISTORY_COLUMN,
  CONTENT_VARIABLES, BUTTON_COLORS, CACHE_TTL_MS, AI_SYSTEM_PROMPT,
};