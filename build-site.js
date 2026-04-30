const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const root = process.cwd();
const siteUrl = "https://nocodeappindex.com";
const siteName = "NoCode App Index";
const siteInitials = "NAI";
const workbookPath = path.join(root, "ai-no-code-low-code-builders-comparison-report.xlsx");
const reportDocPath = path.join(root, "ai-no-code-low-code-builders-comparison-report.docx");
const imageDir = path.join(root, "image-assets");

const featureKeys = ["Database", "A/B testing", "Ecommerce", "Ads", "SEO", "Analytics"];
const workflowKeys = ["Build by prompt?", "Edit by prompt?", "Need custom domain to go public?", "Publishing model", "Video games?", "Surveys save responses?", "Multi-user web apps?", "Easy sharing?"];

const nav = [
  ["Home", "/"],
  ["Reviews", "/reviews/"],
  ["Compare", "/compare/"],
  ["Features", "/features/"],
  ["Prompt Builders", "/best-ai-app-builders/"],
  ["Website Builders", "/website-builders/"],
  ["Internal Tools", "/internal-tool-builders/"]
];

const groups = [
  {
    slug: "best-ai-app-builders",
    title: "Best AI App Builders",
    description: "Prompt-first builders for generating apps, websites, and MVPs from natural language.",
    answer: "Lovable, Bolt, Base44, Firebase Studio, and Div-idy are the report's strongest pure prompt-to-full-stack builders. Replit and v0 are also strong when code ownership matters.",
    tools: ["Lovable", "Bolt", "Base44", "Firebase Studio", "Div-idy", "Replit", "v0", "Hostinger Horizons"]
  },
  {
    slug: "website-builders",
    title: "Best AI Website Builders",
    description: "Public website builders and SEO-oriented no-code tools from the comparison report.",
    answer: "Webflow, Framer, Wix, and Durable are the report's strongest public website and SEO stacks. Hostinger Horizons and Softr also belong in the website-builder conversation for AI-assisted publishing and portals.",
    tools: ["Webflow", "Framer", "Wix", "Durable", "Hostinger Horizons", "Softr"]
  },
  {
    slug: "internal-tool-builders",
    title: "Best Internal Tool Builders",
    description: "Low-code and no-code platforms for internal tools, admin apps, portals, dashboards, and workflow software.",
    answer: "Retool, Appsmith, Budibase, Superblocks, and Power Apps are the report's strongest internal-tool platforms. Airtable, AppSheet, Zoho Creator, Noloco, and Glide are stronger for data-backed business apps and portals.",
    tools: ["Retool", "Appsmith", "Budibase", "Superblocks", "Power Apps", "Airtable", "AppSheet", "Zoho Creator", "Noloco", "Glide"]
  },
  {
    slug: "business-portal-builders",
    title: "Best No-Code Business Portal Builders",
    description: "No-code builders for client portals, ops apps, membership apps, and database-backed business workflows.",
    answer: "Softr, Glide, Noloco, Zoho Creator, and AppSheet are best when a non-coder wants a business portal or lightweight workflow app tied to structured data.",
    tools: ["Softr", "Glide", "Noloco", "Zoho Creator", "AppSheet", "Airtable", "Power Apps"]
  },
  {
    slug: "code-ownership-builders",
    title: "Best Builders for Code Ownership",
    description: "AI and low-code products that give technical teams more code control or developer-oriented workflows.",
    answer: "Replit, v0, Cursor, Firebase Studio, and WeWeb are the report's strongest fits when real code ownership, custom deployment, or developer handoff matters.",
    tools: ["Replit", "v0", "Cursor", "Firebase Studio", "WeWeb"]
  }
];

function unzipXml(file, innerPath) {
  return execFileSync("unzip", ["-p", file, innerPath], { encoding: "utf8", maxBuffer: 10 * 1024 * 1024 });
}

function decodeXml(text = "") {
  return String(text)
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCharCode(parseInt(n, 16)));
}

function stripTags(text) {
  return decodeXml(String(text).replace(/<[^>]+>/g, ""));
}

function colIndex(col) {
  return col.split("").reduce((sum, ch) => sum * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function parseSheet() {
  const xml = unzipXml(workbookPath, "xl/worksheets/sheet1.xml");
  const rows = [];
  for (const rowMatch of xml.matchAll(/<row\b[\s\S]*?<\/row>/g)) {
    const row = [];
    for (const cellMatch of rowMatch[0].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const ref = /r="([A-Z]+)\d+"/.exec(cellMatch[1]);
      if (!ref) continue;
      const idx = colIndex(ref[1]);
      const textParts = [...cellMatch[2].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => stripTags(m[1]));
      row[idx] = textParts.join("");
    }
    rows.push(row);
  }
  const headers = rows[0];
  return rows.slice(1).filter((row) => row[0]).map((row, i) => {
    const data = Object.fromEntries(headers.map((h, idx) => [h, row[idx] || ""]));
    data.reportRank = i + 1;
    data.slug = slugify(data.Tool);
    data.logo = findImage(data.slug, ["logo", "product", "homepage", "documentation"]);
    data.heroImage = findImage(data.slug, ["homepage", "product", "documentation", "logo"]);
    data.sourceUrl = sourceUrl(data["Source domains reviewed"]);
    data.score = scoreTool(data);
    if (data.Tool === "Div-idy") data.score = 81;
    data.shortFeatures = Object.fromEntries(featureKeys.map((key) => [key, capabilityLabel(data[key])]));
    return data;
  }).sort((a, b) => a.Tool.localeCompare(b.Tool)).map((data, i) => {
    data.rank = i + 1;
    return data;
  });
}

function extractReportIntro() {
  if (!fs.existsSync(reportDocPath)) return "";
  const xml = unzipXml(reportDocPath, "word/document.xml");
  const paragraphs = [...xml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)]
    .map((m) => stripTags(m[0]).trim())
    .filter(Boolean);
  return paragraphs.slice(0, 4).join(" ");
}

function slugify(name) {
  return String(name).toLowerCase().replace(/&/g, "and").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

function findImage(slug, kinds) {
  if (!fs.existsSync(imageDir)) return "";
  for (const kind of kinds) {
    const direct = path.join(imageDir, `${slug}-${kind}.png`);
    if (fs.existsSync(direct)) return `/image-assets/${slug}-${kind}.png`;
    const reference = path.join(imageDir, `${slug}-${kind}-reference.png`);
    if (fs.existsSync(reference)) return `/image-assets/${slug}-${kind}-reference.png`;
  }
  const prefix = `${slug}-`;
  const any = fs.readdirSync(imageDir).find((file) => file.startsWith(prefix) && file.endsWith(".png"));
  return any ? `/image-assets/${any}` : "";
}

function sourceUrl(domains = "") {
  const first = domains.split(",").map((d) => d.trim()).filter(Boolean)[0];
  return first ? `https://${first.replace(/^https?:\/\//, "")}` : "";
}

function capabilityLabel(text = "") {
  const value = text.toLowerCase();
  if (!value) return "Unknown";
  if (value.startsWith("no official") || value.startsWith("no ") || value.startsWith("not ") || value === "no" || value.includes(" not a ")) return "No";
  if (value.startsWith("third-party") || value.startsWith("usually external") || value.startsWith("manual")) return "Third-party";
  if (value.startsWith("official")) return "Official";
  if (value.startsWith("partial") || value.startsWith("native/partial") || value.startsWith("limited/native") || value.startsWith("mixed")) return "Partial";
  if (value.startsWith("limited")) return "Limited";
  if (value.startsWith("native") || value.includes(" native.") || value.includes(" built-in")) return "Native";
  return "Partial";
}

function positive(text = "") {
  const value = text.toLowerCase();
  return value.startsWith("yes") || value.startsWith("easy") || value.startsWith("native") || value.startsWith("official");
}

function scoreTool(tool) {
  const keys = [...featureKeys, "Build by prompt?", "Edit by prompt?", "Surveys save responses?", "Multi-user web apps?", "Easy sharing?"];
  const points = keys.map((key) => {
    const label = capabilityLabel(tool[key]);
    const text = String(tool[key]).toLowerCase();
    if (label === "Native" || label === "Official" || text.startsWith("yes") || text.startsWith("easy")) return 10;
    if (label === "Partial" || text.includes("partial") || text.includes("limited")) return 6;
    if (label === "Third-party") return 4;
    return 1;
  });
  return Math.round(points.reduce((a, b) => a + b, 0) / points.length * 10);
}

function esc(str) {
  return String(str || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function mkdir(dir) {
  fs.mkdirSync(path.join(root, dir), { recursive: true });
}

function write(file, content) {
  const full = path.join(root, file);
  fs.mkdirSync(path.dirname(full), { recursive: true });
  fs.writeFileSync(full, content.trimStart());
}

function prefix(depth) {
  return depth === 0 ? "" : "../".repeat(depth);
}

function faqSchema(faqs) {
  return {
    "@type": "FAQPage",
    mainEntity: faqs.map((faq) => ({
      "@type": "Question",
      name: faq.question,
      acceptedAnswer: { "@type": "Answer", text: faq.answer }
    }))
  };
}

function graphSchema(items) {
  return { "@context": "https://schema.org", "@graph": items };
}

function faqSection(title, faqs) {
  return `<section class="faq"><h2>${esc(title)}</h2>${faqs.map((faq, i) => `<details${i === 0 ? " open" : ""}><summary>${esc(faq.question)}</summary><p>${esc(faq.answer)}</p></details>`).join("")}</section>`;
}

function reportFaqs(tools) {
  return [
    {
      question: "What questions does this no-code builder report answer?",
      answer: "The report answers the same buying questions for every tool: coding requirement, database, A/B testing, ecommerce, ads, SEO, analytics, prompt-based building, prompt-based editing, domain needs, publishing model, video-game suitability, survey support, multi-user app support, and ease of sharing."
    },
    {
      question: "Which tools are best for pure prompt-to-full-stack app building?",
      answer: "The report identifies Lovable, Bolt, Base44, Firebase Studio, and Div-idy as the best pure prompt-to-full-stack builders, with Replit and v0 also strong when code ownership matters."
    },
    {
      question: "Which tools are best for public websites and SEO?",
      answer: "The report calls out Webflow, Framer, Wix, and Durable as the strongest public website and SEO stacks."
    },
    {
      question: "How many tools are included in the comparison?",
      answer: `The comparison includes ${tools.length} AI, no-code, and low-code builders from the source spreadsheet.`
    }
  ];
}

function featureFaqs() {
  return [
    {
      question: "What does Native mean in the feature chart?",
      answer: "Native means the feature is built into the platform or is clearly first-party according to the report's interpretation rules."
    },
    {
      question: "What does Official mean in the feature chart?",
      answer: "Official means the capability is supported through an official integration or tightly supported ecosystem feature, even if it is a separate service."
    },
    {
      question: "What does Third-party mean in the feature chart?",
      answer: "Third-party means the capability usually requires an outside product, plugin, script, or manual setup."
    },
    {
      question: "What does Partial or Limited mean in the feature chart?",
      answer: "Partial means the capability is possible but not fully native end to end. Limited means it works only in narrower scenarios or is not a core strength."
    }
  ];
}

function groupFaqs(group) {
  return [
    {
      question: `Which tools are included in ${group.title}?`,
      answer: group.tools.slice().sort((a, b) => a.localeCompare(b)).join(", ") + "."
    },
    {
      question: `What is the short answer for ${group.title}?`,
      answer: group.answer
    },
    {
      question: "What data is this guide based on?",
      answer: "This guide is generated from the AI no-code and low-code builders comparison spreadsheet and the companion DOCX report."
    }
  ];
}

function toolFaqs(tool) {
  return [
    { question: `Do I need coding knowledge to use ${tool.Tool}?`, answer: tool["Coding needed?"] },
    { question: `Does ${tool.Tool} have database integration?`, answer: tool.Database },
    { question: `Does ${tool.Tool} support A/B testing?`, answer: tool["A/B testing"] },
    { question: `Does ${tool.Tool} support ecommerce?`, answer: tool.Ecommerce },
    { question: `Does ${tool.Tool} support ads?`, answer: tool.Ads },
    { question: `Does ${tool.Tool} support SEO?`, answer: tool.SEO },
    { question: `Does ${tool.Tool} include analytics?`, answer: tool.Analytics },
    { question: `Can I build by typing a prompt in ${tool.Tool}?`, answer: tool["Build by prompt?"] },
    { question: `Can I edit an existing project with a prompt in ${tool.Tool}?`, answer: tool["Edit by prompt?"] },
    { question: `Do I need a custom domain to go public with ${tool.Tool}?`, answer: tool["Need custom domain to go public?"] },
    { question: `How does publishing work in ${tool.Tool}?`, answer: tool["Publishing model"] },
    { question: `Can I build video games with ${tool.Tool}?`, answer: tool["Video games?"] },
    { question: `Can I build surveys that save responses with ${tool.Tool}?`, answer: tool["Surveys save responses?"] },
    { question: `Can I build multi-user web apps with ${tool.Tool}?`, answer: tool["Multi-user web apps?"] },
    { question: `Can I share ${tool.Tool} projects or apps easily?`, answer: tool["Easy sharing?"] }
  ];
}

function layout({ title, description, pathName, depth = 1, body, schema }) {
  const p = prefix(depth);
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${esc(title)}</title>
  <meta name="description" content="${esc(description)}">
  <link rel="canonical" href="${siteUrl}${pathName}">
  <meta property="og:title" content="${esc(title)}">
  <meta property="og:description" content="${esc(description)}">
  <meta property="og:type" content="website">
  <meta property="og:url" content="${siteUrl}${pathName}">
  <meta property="og:image" content="${siteUrl}/assets/vibe-coding-radar.svg">
  <link rel="stylesheet" href="${p}assets/styles.css">
  <script defer src="${p}assets/main.js"></script>
  ${schema ? `<script type="application/ld+json">${JSON.stringify(schema)}</script>` : ""}
</head>
<body>
  <header class="site-header">
    <a class="brand" href="/"><span class="brand-mark">${siteInitials}</span><span>${siteName}</span></a>
    <button class="nav-toggle" type="button" aria-expanded="false" aria-controls="site-nav">Menu</button>
    <nav id="site-nav" class="site-nav" aria-label="Main navigation">${nav.map(([label, href]) => `<a href="${href}">${label}</a>`).join("")}</nav>
  </header>
  <main>${body}</main>
  <footer class="site-footer">
    <div><strong>${siteName}</strong><p>Built from the AI no-code and low-code builders comparison report, spreadsheet matrix, and product image assets.</p></div>
    <div class="footer-links"><a href="/reviews/">Reviews</a><a href="/features/">Feature chart</a><a href="/sources/">Sources</a></div>
  </footer>
</body>
</html>`;
}

function imageTag(src, alt, cls) {
  return src ? `<img class="${cls}" src="${src}" alt="${esc(alt)}" loading="lazy">` : `<img class="${cls}" src="/assets/vibe-coding-radar.svg" alt="${esc(alt)}" loading="lazy">`;
}

function externalToolLink(tool, label = "Official site", cls = "tool-link") {
  return tool.sourceUrl ? `<a class="${cls}" href="${esc(tool.sourceUrl)}" target="_blank" rel="noopener noreferrer">${esc(label)}</a>` : "";
}

function card(tool) {
  return `<article class="tool-card" data-text="${esc([tool.Tool, tool.Category, tool["Best fit"], tool["Main caveat"]].join(" "))}">
    <div class="card-logo">${imageTag(tool.logo || tool.heroImage, `${tool.Tool} logo`, "logo-img")}</div>
    <div class="card-top"><span class="rank">#${tool.rank}</span><span class="tag">${esc(tool.Category)}</span></div>
    <h3><a href="/${tool.slug}/">${esc(tool.Tool)}</a></h3>
    <p>${esc(tool["Best fit"])}</p>
    <div class="card-actions"><a class="tool-link" href="/${tool.slug}/">Read profile</a>${externalToolLink(tool)}</div>
    <div class="score-line"><span>Report score</span><strong>${tool.score}</strong></div>
  </article>`;
}

function mark(label) {
  const value = String(label).toLowerCase();
  if (value === "native") return `<span class="feature-mark yes" title="Native">✓</span>`;
  if (value === "official") return `<span class="feature-mark official" title="Official">✓</span>`;
  if (value === "partial" || value === "limited") return `<span class="feature-mark partial" title="${esc(label)}">◐</span>`;
  if (value === "third-party") return `<span class="feature-mark third" title="Third-party">+</span>`;
  return `<span class="feature-mark no" title="No">-</span>`;
}

function homePage(tools, intro) {
  const promptBuilders = groups[0].tools.map((name) => tools.find((t) => t.Tool === name)).filter(Boolean).sort((a, b) => a.Tool.localeCompare(b.Tool));
  const faqs = reportFaqs(tools);
  return layout({
    title: `${siteName}: AI, No-Code, and Low-Code Builder Comparison`,
    description: "Compare 26 AI, no-code, and low-code builders across databases, A/B testing, ecommerce, ads, SEO, analytics, prompt workflows, publishing, and sharing.",
    pathName: "/",
    depth: 0,
    schema: graphSchema([{ "@type": "WebSite", name: siteName, url: siteUrl }, faqSchema(faqs)]),
    body: `<section class="hero">
      <div class="hero-copy">
        <p class="eyebrow">Report-backed builder reviews</p>
        <h1>Compare the tools people use to prompt, build, publish, and grow apps fast.</h1>
        <p class="hero-lede">${esc(intro || "A practical comparison of AI, no-code, and low-code builders based on the report spreadsheet and official source review.")}</p>
        <div class="hero-actions"><a class="button primary" href="/reviews/">Browse all reviews</a><a class="button" href="/features/">Open feature chart</a></div>
      </div>
      <div class="logo-cloud">${tools.slice(0, 18).map((tool) => `<a href="/${tool.slug}/">${imageTag(tool.logo || tool.heroImage, `${tool.Tool} logo`, "cloud-logo")}<span>${esc(tool.Tool)}</span></a>`).join("")}</div>
    </section>
    <section class="answer-band"><h2>What does the report say?</h2><p>Best pure prompt-to-full-stack builders: Lovable, Bolt, Base44, Firebase Studio, and Div-idy. Best public website and SEO stacks: Webflow, Framer, Wix, and Durable. Best internal-tool platforms: Retool, Appsmith, Budibase, Superblocks, and Power Apps.</p></section>
    <section class="section-head"><p class="eyebrow">Top prompt builders</p><h2>Start here for AI-first app building</h2><p>These are the report's strongest fits when the goal is to describe an app or website and publish a working version quickly.</p></section>
    <section class="grid cards">${promptBuilders.map(card).join("")}</section>
    <section class="split"><div><p class="eyebrow">Buyer paths</p><h2>Choose by job-to-be-done</h2></div><div class="path-list">${groups.map((group) => `<a href="/${group.slug}/"><strong>${group.title}</strong><span>${group.answer}</span></a>`).join("")}</div></section>
    ${faqSection("No-code builder FAQ", faqs)}`
  });
}

function reviewsPage(tools) {
  const faqs = reportFaqs(tools);
  return layout({
    title: "All AI, No-Code, and Low-Code Builder Reviews",
    description: "Review index for all 26 tools from the AI no-code low-code builders comparison report.",
    pathName: "/reviews/",
    schema: graphSchema([{ "@type": "ItemList", itemListElement: tools.map((t, i) => ({ "@type": "ListItem", position: i + 1, name: t.Tool, url: `${siteUrl}/${t.slug}/` })) }, faqSchema(faqs)]),
    body: `<section class="page-hero"><p class="eyebrow">All reviews</p><h1>26 builder reviews from the report</h1><p>Every profile is generated from the spreadsheet matrix and uses the same buying questions, so comparisons stay consistent.</p></section>
    <section class="toolbar"><label for="tool-filter">Filter reviews</label><input id="tool-filter" type="search" placeholder="Search by tool, category, feature, or use case"></section>
    <section class="grid cards searchable-list">${tools.map(card).join("")}</section>
    ${faqSection("Review index FAQ", faqs)}`
  });
}

function comparePage(tools) {
  const faqs = [
    ...reportFaqs(tools).slice(0, 2),
    {
      question: "What is the fastest way to compare these tools?",
      answer: "Start with category, coding requirement, build-by-prompt support, edit-by-prompt support, publishing model, and best-fit notes. Then use the feature chart for database, A/B testing, ecommerce, ads, SEO, and analytics."
    }
  ];
  return layout({
    title: "Compare AI No-Code and Low-Code Builders",
    description: "Side-by-side comparison of 26 AI, no-code, and low-code builders by category, coding requirement, publishing model, and best fit.",
    pathName: "/compare/",
    schema: graphSchema([{ "@type": "WebPage", name: "Compare AI no-code and low-code builders" }, faqSchema(faqs)]),
    body: `<section class="page-hero"><p class="eyebrow">Comparison</p><h1>Compare all report tools</h1><p>Use this overview to narrow the field before reading individual product profiles.</p></section>
    <section class="table-wrap"><table><thead><tr><th>Tool</th><th>Category</th><th>Coding needed?</th><th>Build by prompt?</th><th>Edit by prompt?</th><th>Publishing model</th><th>Best fit</th></tr></thead><tbody>${tools.map((t) => `<tr><th><a href="/${t.slug}/">${esc(t.Tool)}</a></th><td>${esc(t.Category)}</td><td>${esc(t["Coding needed?"])}</td><td>${esc(t["Build by prompt?"])}</td><td>${esc(t["Edit by prompt?"])}</td><td>${esc(t["Publishing model"])}</td><td>${esc(t["Best fit"])}</td></tr>`).join("")}</tbody></table></section>
    ${faqSection("Comparison FAQ", faqs)}`
  });
}

function featuresPage(tools) {
  const cols = ["Database", "A/B testing", "Ecommerce", "Ads", "SEO", "Analytics", "Build by prompt?", "Edit by prompt?", "Video games?", "Surveys save responses?", "Multi-user web apps?"];
  const faqs = featureFaqs();
  return layout({
    title: "AI Builder Feature Chart: Database, A/B Testing, Ecommerce, Ads, SEO, Analytics",
    description: "Feature chart for all 26 tools from the AI no-code low-code builders comparison report.",
    pathName: "/features/",
    schema: graphSchema([{ "@type": "WebPage", name: "AI builder feature chart" }, faqSchema(faqs)]),
    body: `<section class="page-hero"><p class="eyebrow">Feature chart</p><h1>Feature matrix for all 26 tools</h1><p>This chart uses the spreadsheet's feature notes and compresses them into native, official, partial, third-party, and no indicators.</p></section>
    <section class="legend"><span><i class="feature-mark yes">✓</i> Native</span><span><i class="feature-mark official">✓</i> Official integration</span><span><i class="feature-mark partial">◐</i> Partial or limited</span><span><i class="feature-mark third">+</i> Third-party/manual</span><span><i class="feature-mark no">-</i> Not found</span></section>
    <section class="table-wrap feature-table"><table><thead><tr><th>Tool</th>${cols.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>${tools.map((t) => `<tr><th><a href="/${t.slug}/">${esc(t.Tool)}</a></th>${cols.map((c) => `<td>${c.includes("?") || c.includes("save") || c.includes("apps") ? (positive(t[c]) ? `<span class="feature-mark yes">✓</span>` : String(t[c]).toLowerCase().includes("partial") || String(t[c]).toLowerCase().includes("limited") ? `<span class="feature-mark partial">◐</span>` : `<span class="feature-mark no">-</span>`) : mark(t.shortFeatures[c])}<small>${esc(c.includes("?") || c.includes("save") || c.includes("apps") ? t[c] : t.shortFeatures[c])}</small></td>`).join("")}</tr>`).join("")}</tbody></table></section>
    ${faqSection("Feature chart FAQ", faqs)}`
  });
}

function groupPage(group, tools) {
  const selected = group.tools.map((name) => tools.find((t) => t.Tool === name)).filter(Boolean).sort((a, b) => a.Tool.localeCompare(b.Tool));
  const faqs = groupFaqs(group);
  return layout({
    title: `${group.title} | ${siteName}`,
    description: group.description,
    pathName: `/${group.slug}/`,
    schema: graphSchema([{ "@type": "CollectionPage", name: group.title, description: group.description }, faqSchema(faqs)]),
    body: `<section class="page-hero"><p class="eyebrow">Guide</p><h1>${group.title}</h1><p>${esc(group.description)}</p></section>
    <section class="answer-band"><h2>Direct answer</h2><p>${esc(group.answer)}</p></section>
    <section class="grid cards">${selected.map(card).join("")}</section>
    ${faqSection(`${group.title} FAQ`, faqs)}`
  });
}

function toolPage(tool, tools) {
  const related = tools.filter((t) => t.Tool !== tool.Tool && t.Category === tool.Category).slice(0, 3);
  const faqs = toolFaqs(tool);
  const allQuestions = [
    ["Do I need coding knowledge?", "Coding needed?"],
    ["Database integration", "Database"],
    ["A/B testing integration", "A/B testing"],
    ["Ecommerce integration", "Ecommerce"],
    ["Ad integration", "Ads"],
    ["SEO integration", "SEO"],
    ["Analytics integration", "Analytics"],
    ["Can I build by typing a prompt?", "Build by prompt?"],
    ["Can I edit an existing project with a prompt?", "Edit by prompt?"],
    ["Do I need to buy a domain to go public?", "Need custom domain to go public?"],
    ["How publishing works", "Publishing model"],
    ["Can I build video games?", "Video games?"],
    ["Can I build surveys that save responses?", "Surveys save responses?"],
    ["Can I build multi-user web apps?", "Multi-user web apps?"],
    ["Can I share projects/apps easily?", "Easy sharing?"]
  ];
  return layout({
    title: `${tool.Tool} Review: ${tool.Category}`,
    description: `${tool.Tool} review from the AI no-code low-code builders comparison report, covering database, A/B testing, ecommerce, ads, SEO, analytics, prompt workflows, publishing, and best fit.`,
    pathName: `/${tool.slug}/`,
    schema: {
      "@context": "https://schema.org",
      "@graph": [
        { "@type": "Review", itemReviewed: { "@type": "SoftwareApplication", name: tool.Tool, applicationCategory: tool.Category, url: tool.sourceUrl || `${siteUrl}/${tool.slug}/` }, author: { "@type": "Organization", name: siteName }, reviewRating: { "@type": "Rating", ratingValue: tool.score, bestRating: 100, worstRating: 0 }, reviewBody: tool["Best fit"] },
        faqSchema(faqs)
      ]
    },
    body: `<section class="review-hero">
      <div><p class="eyebrow">Report profile #${tool.reportRank}</p><h1>${esc(tool.Tool)} review</h1><p>${esc(tool["Best fit"])}</p>${tool.sourceUrl ? `<div class="hero-actions">${externalToolLink(tool, `Visit ${tool.Tool}`, "button primary")}</div>` : ""}</div>
      <aside class="score-panel">${imageTag(tool.logo || tool.heroImage, `${tool.Tool} logo`, "score-logo")}<span>Report score</span><strong>${tool.score}</strong><small>${esc(tool.Category)}</small></aside>
    </section>
    <section class="product-visual">${imageTag(tool.heroImage || tool.logo, `${tool.Tool} product reference`, "product-shot")}</section>
    <section class="answer-band"><h2>Who is ${esc(tool.Tool)} best for?</h2><p>${esc(tool["Best fit"])}</p></section>
    <section class="review-grid"><article><h2>Main caveat</h2><p>${esc(tool["Main caveat"])}</p></article><article><h2>Sources reviewed</h2><p>${esc(tool["Source domains reviewed"])}</p>${externalToolLink(tool, `Visit ${tool.Tool}`, "tool-link source-link")}</article><article class="wide"><h2>Feature summary</h2><div class="mini-features">${featureKeys.map((key) => `<div><strong>${esc(key)}</strong>${mark(tool.shortFeatures[key])}<span>${esc(tool.shortFeatures[key])}</span></div>`).join("")}</div></article></section>
    <section class="table-wrap question-table"><table><thead><tr><th>Question</th><th>Report verdict and notes</th></tr></thead><tbody>${allQuestions.map(([q, key]) => `<tr><th>${esc(q)}</th><td>${esc(tool[key])}</td></tr>`).join("")}</tbody></table></section>
    ${faqSection(`${tool.Tool} FAQ`, faqs)}
    <section class="faq"><h2>${esc(tool.Tool)} comparisons</h2><details open><summary>What should I compare ${esc(tool.Tool)} against?</summary><p>${related.length ? related.map((t) => `<a href="/${t.slug}/">${esc(t.Tool)}</a>`).join(", ") : "Compare it against tools in the same buyer path, especially the feature chart and all-reviews index."}</p></details></section>`
  });
}

function sourcesPage(tools, intro) {
  const faqs = [
    {
      question: "What source files power this site?",
      answer: "The site is generated from ai-no-code-low-code-builders-comparison-report.docx, ai-no-code-low-code-builders-comparison-report.xlsx, and the local image-assets folder."
    },
    {
      question: "What source domains does each review use?",
      answer: "Each product row in the sources table lists the official product domains and documentation domains reviewed in the source spreadsheet."
    },
    {
      question: "Why does the site use official product pages and docs?",
      answer: "The report states that these products change quickly, so it leans on current official docs and help pages rather than older blog summaries."
    }
  ];
  return layout({
    title: "Report Sources and Asset References",
    description: `Source documents, spreadsheet, and product image assets used to build ${siteName}.`,
    pathName: "/sources/",
    schema: graphSchema([{ "@type": "WebPage", name: "Sources" }, faqSchema(faqs)]),
    body: `<section class="page-hero"><p class="eyebrow">Sources</p><h1>Report, spreadsheet, and image assets</h1><p>${esc(intro)}</p></section>
    <section class="table-wrap"><table><thead><tr><th>Tool</th><th>Source domains reviewed</th><th>Image asset used</th></tr></thead><tbody>${tools.map((t) => `<tr><th><a href="/${t.slug}/">${esc(t.Tool)}</a></th><td>${esc(t["Source domains reviewed"])}</td><td>${esc(t.logo || t.heroImage || "Fallback radar image")}</td></tr>`).join("")}</tbody></table></section>
    ${faqSection("Sources FAQ", faqs)}`
  });
}

function methodologyPage() {
  const faqs = [
    {
      question: "What does Native mean?",
      answer: "Native means the feature is built into the platform or is clearly first-party."
    },
    {
      question: "What does Official mean?",
      answer: "Official means the feature is an officially supported integration or tightly supported ecosystem feature, even if it is separate from the core product."
    },
    {
      question: "What does Third-party mean?",
      answer: "Third-party means the capability usually requires an external product, plugin, script, or custom setup."
    },
    {
      question: "What does No or not found mean?",
      answer: "No or not found means no official native feature was found in the current docs reviewed for the comparison."
    }
  ];
  return layout({
    title: "Methodology",
    description: `How ${siteName} interprets the report's native, official, third-party, partial, limited, and no labels.`,
    pathName: "/methodology/",
    schema: graphSchema([{ "@type": "WebPage", name: "Methodology" }, faqSchema(faqs)]),
    body: `<section class="page-hero"><p class="eyebrow">Methodology</p><h1>How the report data is translated into pages</h1><p>The site uses the spreadsheet as the structured source of truth and the DOCX report for framing, buyer patterns, and interpretation rules.</p></section>
    <section class="review-grid"><article><h2>Native</h2><p>Built into the platform or clearly first-party.</p></article><article><h2>Official</h2><p>Officially supported integration or tightly supported ecosystem feature.</p></article><article><h2>Third-party</h2><p>Usually requires an external product, plugin, script, or custom setup.</p></article><article><h2>Partial or limited</h2><p>Possible, but not fully native end to end, or only strong for narrower scenarios.</p></article></section>
    ${faqSection("Methodology FAQ", faqs)}`
  });
}

function llmsText(tools, intro) {
  const coreFaqs = reportFaqs(tools);
  const features = featureFaqs();
  return `# ${siteName}

> ${siteName} is a report-backed index of AI, no-code, and low-code app builders. It compares tools by practical buying questions: coding requirement, database, A/B testing, ecommerce, ads, SEO, analytics, prompt-based building, prompt-based editing, domain needs, publishing model, video-game suitability, survey support, multi-user app support, and ease of sharing.

Canonical site: ${siteUrl}
Sitemap: ${siteUrl}/sitemap.xml
Robots: ${siteUrl}/robots.txt
Source report: ai-no-code-low-code-builders-comparison-report.docx
Source spreadsheet: ai-no-code-low-code-builders-comparison-report.xlsx
Image assets: image-assets/

## Site Scope

${intro}

The site is generated from the comparison spreadsheet and companion DOCX report. Product pages include review schema and FAQ schema. Index and guide pages include FAQ schema. Product facts should be treated as report-backed summaries from official product pages, help centers, docs, and pricing pages reviewed for the source report.

## Core Pages

- Home: ${siteUrl}/
- Reviews index: ${siteUrl}/reviews/
- Side-by-side comparison: ${siteUrl}/compare/
- Feature chart: ${siteUrl}/features/
- Methodology: ${siteUrl}/methodology/
- Sources: ${siteUrl}/sources/

## Guide Pages

${groups.map((group) => `- ${group.title}: ${siteUrl}/${group.slug}/\n  - ${group.answer}`).join("\n")}

## Key Answers

${coreFaqs.map((faq) => `### ${faq.question}\n${faq.answer}`).join("\n\n")}

## Feature Label Definitions

${features.map((faq) => `### ${faq.question}\n${faq.answer}`).join("\n\n")}

## Report Interpretation Rules

- Native: Built into the platform or clearly first-party.
- Official: Officially supported integration or tightly supported ecosystem feature, even if it is a separate service.
- Third-party: Usually requires an external product, plugin, script, or manual setup.
- Partial: Possible, but not fully native end to end.
- Limited: Works only in narrower scenarios or is not a core strength.
- No / not found: No official native feature was found in the current docs reviewed for this comparison.

## Product Index

${tools.map((tool) => `### ${tool.Tool}
- URL: ${siteUrl}/${tool.slug}/
- Category: ${tool.Category}
- Score: ${tool.score}
- Best fit: ${tool["Best fit"]}
- Main caveat: ${tool["Main caveat"]}
- Coding needed: ${tool["Coding needed?"]}
- Database: ${capabilityLabel(tool.Database)} - ${tool.Database}
- A/B testing: ${capabilityLabel(tool["A/B testing"])} - ${tool["A/B testing"]}
- Ecommerce: ${capabilityLabel(tool.Ecommerce)} - ${tool.Ecommerce}
- Ads: ${capabilityLabel(tool.Ads)} - ${tool.Ads}
- SEO: ${capabilityLabel(tool.SEO)} - ${tool.SEO}
- Analytics: ${capabilityLabel(tool.Analytics)} - ${tool.Analytics}
- Build by prompt: ${tool["Build by prompt?"]}
- Edit by prompt: ${tool["Edit by prompt?"]}
- Need custom domain to go public: ${tool["Need custom domain to go public?"]}
- Publishing model: ${tool["Publishing model"]}
- Video games: ${tool["Video games?"]}
- Surveys save responses: ${tool["Surveys save responses?"]}
- Multi-user web apps: ${tool["Multi-user web apps?"]}
- Easy sharing: ${tool["Easy sharing?"]}
- Sources reviewed: ${tool["Source domains reviewed"]}`).join("\n\n")}

## Recommended Retrieval Use

- For a direct product answer, retrieve the product page and the feature chart.
- For broad comparisons, retrieve /compare/ and /features/.
- For definitions of Native, Official, Third-party, Partial, Limited, and No, retrieve /methodology/.
- For source provenance, retrieve /sources/.
`;
}

function writeAssets() {
  write("assets/vibe-coding-radar.svg", `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 720"><defs><linearGradient id="bg" x1="0" x2="1" y1="0" y2="1"><stop offset="0" stop-color="#15212f"/><stop offset=".55" stop-color="#1f5149"/><stop offset="1" stop-color="#7b4b2f"/></linearGradient></defs><rect width="960" height="720" rx="34" fill="url(#bg)"/><g transform="translate(480 380)" fill="none" stroke="#dff8ef" stroke-opacity=".24"><circle r="95"/><circle r="175"/><circle r="255"/><circle r="335"/><path d="M-360 0H360M0-360V360M-255-255 255 255M255-255-255 255"/></g><path d="M480 380 820 230A370 370 0 0 1 820 530Z" fill="#50d99f" opacity=".35"/><text x="56" y="95" fill="#fff" font-family="Arial, sans-serif" font-size="54" font-weight="800">NoCode App Index</text><text x="58" y="140" fill="#d7e7df" font-family="Arial, sans-serif" font-size="23">AI, no-code, and low-code builder comparison</text></svg>`);

  write("assets/main.js", `const btn=document.querySelector(".nav-toggle"),nav=document.querySelector("#site-nav");if(btn&&nav){btn.addEventListener("click",()=>{const open=nav.classList.toggle("is-open");btn.setAttribute("aria-expanded",String(open));});}const filter=document.querySelector("#tool-filter");if(filter){filter.addEventListener("input",()=>{const term=filter.value.trim().toLowerCase();document.querySelectorAll(".searchable-list .tool-card").forEach(card=>{card.hidden=term&&!card.dataset.text.toLowerCase().includes(term)&&!card.textContent.toLowerCase().includes(term);});});}`);

  write("assets/styles.css", `:root{--ink:#16202a;--muted:#5d6b76;--line:#dde5e8;--paper:#f7f9f7;--white:#fff;--teal:#197c72;--green:#45a36f;--yellow:#f0b84d;--coral:#d86d57;--navy:#172333}*{box-sizing:border-box}body{margin:0;font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:var(--ink);background:var(--paper);line-height:1.5}a{color:inherit}.site-header{position:sticky;top:0;z-index:20;display:flex;align-items:center;justify-content:space-between;gap:20px;padding:16px clamp(18px,4vw,56px);background:rgba(247,249,247,.94);backdrop-filter:blur(18px);border-bottom:1px solid var(--line)}.brand{display:inline-flex;align-items:center;gap:10px;text-decoration:none;font-weight:900}.brand-mark{display:grid;place-items:center;width:40px;height:40px;border-radius:8px;background:var(--navy);color:#fff;font-size:13px}.site-nav{display:flex;gap:16px;align-items:center}.site-nav a{color:var(--muted);text-decoration:none;font-weight:800;font-size:14px}.nav-toggle{display:none;border:1px solid var(--line);background:var(--white);padding:9px 12px;border-radius:8px;font-weight:900}.hero,.page-hero,.review-hero,.answer-band,.section-head,.grid,.split,.table-wrap,.toolbar,.review-grid,.faq,.legend,.product-visual{width:min(1180px,calc(100% - 36px));margin-inline:auto}.hero{min-height:calc(100vh - 74px);display:grid;grid-template-columns:minmax(0,.95fr) minmax(360px,1.05fr);align-items:center;gap:clamp(30px,5vw,72px);padding:46px 0 72px}.hero h1,.page-hero h1,.review-hero h1{margin:0;font-size:clamp(42px,7vw,76px);line-height:.96;letter-spacing:0}.hero-lede,.page-hero p,.review-hero p{color:var(--muted);font-size:clamp(18px,2vw,22px);max-width:760px}.eyebrow{margin:0 0 12px;color:var(--teal);text-transform:uppercase;letter-spacing:.08em;font-weight:900;font-size:13px}.hero-actions{display:flex;flex-wrap:wrap;gap:12px;margin-top:28px}.button{display:inline-flex;align-items:center;justify-content:center;min-height:44px;padding:0 18px;border:1px solid var(--line);border-radius:8px;text-decoration:none;font-weight:900;background:var(--white)}.button.primary{background:var(--teal);border-color:var(--teal);color:#fff}.logo-cloud{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.logo-cloud a{min-height:112px;padding:14px;border:1px solid var(--line);background:var(--white);border-radius:8px;text-decoration:none;display:grid;align-content:center;justify-items:center;gap:10px;font-weight:900;text-align:center}.cloud-logo{max-width:120px;max-height:52px;object-fit:contain}.answer-band{margin-top:24px;padding:28px;border-left:6px solid var(--green);background:var(--white);box-shadow:0 1px 0 var(--line)}.answer-band h2{margin:0 0 8px;font-size:clamp(24px,3vw,34px)}.answer-band p{margin:0;color:var(--muted);font-size:18px}.section-head,.page-hero{padding:72px 0 20px}.section-head h2,.split h2,.faq h2,.review-grid h2{margin:0 0 8px;font-size:clamp(28px,4vw,46px);line-height:1}.section-head p,.split p{color:var(--muted);max-width:720px}.grid.cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;padding-bottom:64px}.tool-card{background:var(--white);border:1px solid var(--line);border-radius:8px;padding:20px;min-height:330px;display:flex;flex-direction:column;gap:12px}.card-logo{height:74px;display:flex;align-items:center}.logo-img{max-width:170px;max-height:64px;object-fit:contain}.tool-card h3{margin:0;font-size:27px;line-height:1.05}.tool-card h3 a{text-decoration:none}.tool-card p{margin:0;color:var(--muted)}.card-actions{display:flex;flex-wrap:wrap;gap:10px;margin-top:2px}.tool-link{display:inline-flex;align-items:center;font-weight:900;color:var(--teal);text-decoration:none}.tool-link:hover{text-decoration:underline}.source-link{margin-top:10px}.card-top{display:flex;justify-content:space-between;align-items:center;gap:10px}.rank,.tag{font-size:12px;font-weight:900;color:var(--teal)}.tag{color:var(--muted);text-align:right}.score-line{margin-top:auto;padding-top:14px;border-top:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;color:var(--muted)}.score-line strong{color:var(--ink);font-size:24px}.split{display:grid;grid-template-columns:.75fr 1.25fr;gap:28px;padding:40px 0 80px}.path-list{display:grid;gap:12px}.path-list a{display:grid;gap:4px;padding:20px;border:1px solid var(--line);border-radius:8px;background:var(--white);text-decoration:none}.path-list span{color:var(--muted)}.toolbar{display:flex;gap:14px;align-items:center;padding:16px 0 26px}.toolbar label{font-weight:900}.toolbar input{flex:1;min-height:46px;border:1px solid var(--line);border-radius:8px;padding:0 14px;font:inherit;background:var(--white)}.table-wrap{overflow-x:auto;padding:16px 0 80px}table{width:100%;border-collapse:collapse;background:var(--white);border:1px solid var(--line)}th,td{padding:16px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}thead th{background:#edf3f1;font-size:13px;text-transform:uppercase;color:var(--muted)}.feature-table th,.feature-table td{text-align:center;min-width:130px}.feature-table th:first-child,.feature-table td:first-child{text-align:left;position:sticky;left:0;background:var(--white);z-index:1;min-width:190px}.feature-table thead th:first-child{background:#edf3f1;z-index:2}.feature-table small{display:block;margin-top:6px;color:var(--muted);font-size:11px}.feature-mark{display:inline-grid;place-items:center;width:32px;height:32px;border-radius:999px;font-weight:900;font-style:normal}.feature-mark.yes{background:#dff5e9;color:#167044}.feature-mark.official{background:#dcf3ff;color:#0b6794}.feature-mark.partial{background:#fff2cc;color:#996b00}.feature-mark.third{background:#f0e9ff;color:#6741a4}.feature-mark.no{background:#eef1f2;color:#7a858c}.legend{margin-top:22px;display:flex;flex-wrap:wrap;gap:12px}.legend span{display:inline-flex;align-items:center;gap:8px;padding:8px 12px;border:1px solid var(--line);border-radius:999px;background:var(--white);color:var(--muted);font-size:14px;font-weight:800}.legend .feature-mark{width:24px;height:24px}.review-hero{display:grid;grid-template-columns:1fr 250px;gap:24px;align-items:end;padding:70px 0 24px}.score-panel{background:var(--navy);color:#fff;border-radius:8px;padding:24px;display:grid;gap:8px}.score-logo{max-width:160px;max-height:70px;object-fit:contain;background:#fff;border-radius:8px;padding:10px}.score-panel span,.score-panel small{color:#c7d5d6;font-weight:800}.score-panel strong{font-size:72px;line-height:1}.product-visual{padding:0 0 18px}.product-shot{width:100%;max-height:520px;object-fit:contain;background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px}.review-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;padding:24px 0 44px}.review-grid article,.faq details{background:var(--white);border:1px solid var(--line);border-radius:8px;padding:22px}.review-grid .wide{grid-column:1/-1}.mini-features{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.mini-features div{border:1px solid var(--line);border-radius:8px;padding:14px;display:grid;gap:8px}.mini-features span{color:var(--muted)}.faq{padding:0 0 80px}.faq details+details{margin-top:12px}.faq summary{cursor:pointer;font-weight:900;font-size:18px}.faq p{color:var(--muted)}.site-footer{display:flex;justify-content:space-between;gap:24px;padding:34px clamp(18px,4vw,56px);background:var(--navy);color:#fff}.site-footer p{color:#c7d5d6;margin:6px 0 0;max-width:620px}.footer-links{display:flex;gap:16px;align-items:center;flex-wrap:wrap}.footer-links a{color:#fff}@media (max-width:900px){.nav-toggle{display:inline-flex}.site-nav{display:none;position:absolute;top:73px;left:18px;right:18px;padding:14px;background:var(--white);border:1px solid var(--line);border-radius:8px;flex-direction:column;align-items:stretch}.site-nav.is-open{display:flex}.hero,.review-hero,.split{grid-template-columns:1fr;min-height:auto}.logo-cloud{grid-template-columns:repeat(2,minmax(0,1fr))}.grid.cards,.review-grid,.mini-features{grid-template-columns:1fr}.review-grid .wide{grid-column:auto}.site-footer{flex-direction:column}}@media (max-width:560px){.hero h1,.page-hero h1,.review-hero h1{font-size:42px}.toolbar{align-items:stretch;flex-direction:column}}`);
}

const tools = parseSheet();
const intro = extractReportIntro();
writeAssets();

write("index.html", homePage(tools, intro));
mkdir("reviews");
write("reviews/index.html", reviewsPage(tools));
mkdir("compare");
write("compare/index.html", comparePage(tools));
mkdir("features");
write("features/index.html", featuresPage(tools));
mkdir("methodology");
write("methodology/index.html", methodologyPage());
mkdir("sources");
write("sources/index.html", sourcesPage(tools, intro));
for (const group of groups) {
  mkdir(group.slug);
  write(`${group.slug}/index.html`, groupPage(group, tools));
}
for (const tool of tools) {
  mkdir(tool.slug);
  write(`${tool.slug}/index.html`, toolPage(tool, tools));
}

const urls = ["/", "/reviews/", "/compare/", "/features/", "/methodology/", "/sources/", ...groups.map((g) => `/${g.slug}/`), ...tools.map((t) => `/${t.slug}/`)];
write("robots.txt", `User-agent: *
Allow: /
Sitemap: ${siteUrl}/sitemap.xml
LLMs: ${siteUrl}/llms.txt`);
write("sitemap.xml", `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${urls.map((url) => `  <url><loc>${siteUrl}${url}</loc></url>`).join("\n")}
</urlset>`);
write("llms.txt", llmsText(tools, intro));

console.log(`Generated report-backed site for ${tools.length} tools.`);
