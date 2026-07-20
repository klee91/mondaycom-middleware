/**
 * helpers/github-store.js — GitHub-hosted manifest + brand guide, auto-tagging, and RAG few-shot retrieval.
 */
const { fetch, anthropic, CHEAP_MODEL, CACHE_TTL_MS } = require("./config");
const { fetchTemplateIndex, fetchTemplateFromSharePoint } = require("./sharepoint");

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
    model: CHEAP_MODEL,
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
    model: CHEAP_MODEL,
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

module.exports = { getManifest, ensureManifestUpToDate, autoTagTemplate, writeManifestToGithub, getBrandGuide, selectRelevantTemplates, buildSystemPromptWithRAG, classifyByKeyword, classifyWithModel };