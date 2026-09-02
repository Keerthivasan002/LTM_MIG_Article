"use strict";

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DOWNLOAD_DIR = path.join(__dirname, "downloads");
const RESULTS_PATH = path.join(__dirname, "migration-results.csv");
const TEMPLATE_TABLES = [
  "kb_knowledge",
  "kb_template_how_to",
  "kb_template_faq",
  "kb_template_what_is",
  "kb_template_kcs_article",
  "kb_template_known_error_article",
];

loadEnv(path.join(__dirname, ".env"));

const SNOW_INSTANCE = normalizeSnowInstance(required("SNOW_INSTANCE"));
const SNOW_USER = required("SNOW_USER");
const SNOW_PASSWORD = required("SNOW_PASSWORD");
const FS_DOMAIN = normalizeFsDomain(required("FS_DOMAIN"));
const FS_API_KEY = required("FS_API_KEY");
const FS_FOLDER_ID = required("FS_FOLDER_ID");
const EXCEL_PATH = path.resolve(
  __dirname,
  process.env.EXCEL_PATH || "articles.xlsx"
);

const snowAuth = basicAuth(SNOW_USER, SNOW_PASSWORD);
const fsAuth = basicAuth(FS_API_KEY, "X");

async function main() {
  fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });

  const jobs = readJobsFromExcel(EXCEL_PATH);
  if (!jobs.length) {
    throw new Error(
      `No SNOW_ARTICLE_SYS_ID values found in ${EXCEL_PATH}. Put sys_ids in that column.`
    );
  }

  const results = loadPreviousResults(RESULTS_PATH);
  const processed = new Set(
    results
      .filter((row) => row.ok)
      .map((row) => String(row.sysId).toLowerCase())
  );
  const pending = jobs.filter((job) => !processed.has(job.sysId));
  const skipped = jobs.length - pending.length;

  console.log("Testing ServiceNow and Freshservice access...");
  await snowGet("/api/now/table/kb_knowledge", { sysparm_limit: 1 });
  const folder = await fsGet(`/api/v2/solutions/folders/${FS_FOLDER_ID}`);
  const folderName = (folder.folder && folder.folder.name) || FS_FOLDER_ID;
  console.log(`Using Freshservice folder ${FS_FOLDER_ID} (${folderName})`);
  console.log(
    `Excel rows: ${jobs.length}, already done: ${skipped}, to process: ${pending.length}\n`
  );

  if (!pending.length) {
    console.log("Nothing new to migrate. Already processed rows were skipped.");
    return;
  }

  for (let i = 0; i < pending.length; i++) {
    const job = pending[i];
    const label = `${i + 1}/${pending.length} ${job.sysId}`;
    try {
      const result = await migrateOne(job, folder.folder || { id: FS_FOLDER_ID });
      upsertResult(results, result);
      console.log(
        `OK   ${label} -> ${result.number} -> FS ${result.freshserviceId}`
      );
    } catch (err) {
      upsertResult(results, {
        sysId: job.sysId,
        table: job.table || "",
        ok: false,
        error: err.message || String(err),
      });
      console.error(`FAIL ${label}: ${err.message || err}`);
    } finally {
      removeDownloadDir(job.sysId);
      writeResultsCsv(RESULTS_PATH, results);
    }
  }

  const ok = results.filter((r) => r.ok).length;
  console.log(`\nDone. ${ok}/${results.length} succeeded. Log: ${RESULTS_PATH}`);
}

async function migrateOne(job, folder) {
  const article = await getKnowledgeArticle(job.sysId, job.table);
  article.html = composeArticleHtml(article);
  const files = await collectArticleFiles(article);
  const saved = await downloadFiles(article.sys_id, files);
  const folderId = job.folderId || folder.id;
  const created = await createArticleWithAttachments(folderId, article, saved);
  const updated = await rewriteInlineImages(created, article, saved);
  return {
    sysId: article.sys_id,
    table: article._table,
    number: article.number,
    title: snowValue(article.short_description),
    freshserviceId: updated.id,
    attachments: (updated.attachments || []).length,
    ok: true,
    url: `https://${FS_DOMAIN}/a/solutions/articles/${updated.id}`,
  };
}

function readJobsFromExcel(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(
      `Excel file not found: ${filePath}. Create it with a SNOW_ARTICLE_SYS_ID column.`
    );
  }

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const jobs = [];
  const seen = new Set();

  for (const row of rows) {
    const mapped = mapRow(row);
    const sysId = String(mapped.snow_article_sys_id || mapped.sys_id || "")
      .trim()
      .toLowerCase();
    if (!sysId) continue;
    if (seen.has(sysId)) continue;
    seen.add(sysId);
    jobs.push({
      sysId,
      table: String(
        mapped.snow_article_table || mapped.table || mapped.sys_class_name || ""
      ).trim(),
      folderId: String(mapped.fs_folder_id || mapped.folder_id || "").trim(),
    });
  }
  return jobs;
}

function mapRow(row) {
  const mapped = {};
  for (const [key, value] of Object.entries(row)) {
    mapped[normalizeHeader(key)] = value;
  }
  return mapped;
}

function normalizeHeader(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "");
}

async function getKnowledgeArticle(sysId, hintedTable) {
  const tryTables = [
    hintedTable,
    "kb_knowledge",
    ...TEMPLATE_TABLES,
  ].filter((name, i, all) => name && all.indexOf(name) === i);

  let base = null;
  let foundTable = "";
  for (const table of tryTables) {
    const row = await snowGetOrNull(`/api/now/table/${table}/${sysId}`);
    if (row && row.sys_id) {
      base = row;
      foundTable = table;
      break;
    }
  }
  if (!base) {
    throw new Error(
      `Article ${sysId} was not found on kb_knowledge or template tables (ACL or wrong id).`
    );
  }

  const className = snowValue(base.sys_class_name) || foundTable;
  if (className && className !== foundTable) {
    const extra = await snowGetOrNull(`/api/now/table/${className}/${sysId}`);
    if (extra && extra.sys_id) {
      return { ...base, ...extra, _table: className };
    }
  }
  return { ...base, _table: foundTable || className };
}

function composeArticleHtml(article) {
  const sections = [
    ["Introduction", "kb_introduction"],
    ["Instructions", "kb_instructions"],
    ["Question", "kb_question"],
    ["Answer", "kb_answer"],
    ["Explanation", "kb_explanation"],
    ["Issue", "kb_issue"],
    ["Environment", "kb_environment"],
    ["Cause", "kb_cause"],
    ["Resolution", "kb_resolution"],
    ["Description", "kb_description"],
    ["Workaround", "kb_workaround"],
  ];
  const parts = [];
  for (const [label, field] of sections) {
    const value = snowValue(article[field]);
    if (value) parts.push(`<h2>${label}</h2>${toHtml(value)}`);
  }
  const text = snowValue(article.text);
  if (!parts.length && text) parts.push(toHtml(text));
  if (!parts.length) {
    parts.push(toHtml(snowValue(article.short_description) || "Untitled"));
  }
  return parts.join("\n");
}

async function collectArticleFiles(article) {
  const tableNames = [
    article._table,
    "kb_knowledge",
    snowValue(article.sys_class_name),
  ].filter(Boolean);
  const listed = [];
  for (const tableName of [...new Set(tableNames)]) {
    const rows = await snowGet("/api/now/attachment", {
      sysparm_query: `table_name=${tableName}^table_sys_id=${article.sys_id}`,
      sysparm_limit: 100,
    });
    listed.push(...(Array.isArray(rows) ? rows : []));
  }

  const bySysId = new Map();
  for (const att of listed) {
    const sysId = snowValue(att.sys_id);
    if (!sysId) continue;
    bySysId.set(sysId, {
      sys_id: sysId,
      file_name: att.file_name || `${sysId}.bin`,
      content_type: att.content_type || "application/octet-stream",
      inline: false,
    });
  }

  const html = [
    article.html,
    article.text,
    article.kb_introduction,
    article.kb_instructions,
    article.kb_question,
    article.kb_answer,
    article.kb_explanation,
    article.kb_issue,
    article.kb_resolution,
  ]
    .map((v) => snowValue(v))
    .join("\n");
  for (const sysId of extractAttachmentSysIds(html)) {
    if (bySysId.has(sysId)) {
      bySysId.get(sysId).inline = true;
      continue;
    }
    bySysId.set(sysId, {
      sys_id: sysId,
      file_name: `inline-${sysId}.bin`,
      content_type: "application/octet-stream",
      inline: true,
    });
  }

  return [...bySysId.values()];
}

async function downloadFiles(articleSysId, files) {
  const dir = path.join(DOWNLOAD_DIR, String(articleSysId).toLowerCase());
  fs.mkdirSync(dir, { recursive: true });
  const saved = [];
  for (const file of files) {
    const safeName = sanitizeFileName(file.file_name);
    const dest = path.join(dir, `${file.sys_id.slice(0, 8)}-${safeName}`);
    const buffer = await snowDownload(
      `/api/now/attachment/${file.sys_id}/file`
    );
    fs.writeFileSync(dest, buffer);
    saved.push({ ...file, dest, size: buffer.length });
  }
  return saved;
}

async function createArticleWithAttachments(folderId, article, files) {
  const title = `${snowValue(article.short_description) || "Untitled"}`;
  const description = article.html || composeArticleHtml(article);

  const form = new FormData();
  form.set("title", title);
  form.set("description", description);
  form.set("folder_id", String(folderId));
  form.set("status", "1");
  form.set("article_type", "1");

  for (const file of files) {
    const bytes = fs.readFileSync(file.dest);
    form.append(
      "attachments[]",
      new Blob([bytes], { type: file.content_type }),
      path.basename(file.dest)
    );
  }

  try {
    const created = await fsForm("POST", "/api/v2/solutions/articles", form);
    return created.article;
  } catch (err) {
    console.log(
      `     multipart create failed (${err.message}); creating JSON then attaching files.`
    );
    const created = await fsJson("POST", "/api/v2/solutions/articles", {
      title,
      description,
      folder_id: folderId,
      status: 1,
      article_type: 1,
    });
    if (!files.length) return created.article;

    const update = new FormData();
    for (const file of files) {
      const bytes = fs.readFileSync(file.dest);
      update.append(
        "attachments[]",
        new Blob([bytes], { type: file.content_type }),
        path.basename(file.dest)
      );
    }
    const patched = await fsForm(
      "PUT",
      `/api/v2/solutions/articles/${created.article.id}`,
      update
    );
    return patched.article;
  }
}

function decodeEntities(value) {
  return String(value || "")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) =>
      String.fromCharCode(parseInt(hex, 16))
    )
    .replace(/&#(\d+);/g, (_, num) => String.fromCharCode(Number(num)))
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");
}

function findFsAttachment(file, attachments) {
  const destName = path.basename(file.dest || "");
  const names = [destName, file.file_name, destName.replace(/\s+/g, "_")]
    .filter(Boolean)
    .map((n) => n.toLowerCase().replace(/\s+/g, "_"));
  const prefix = String(file.sys_id || "").slice(0, 8).toLowerCase();
  return (attachments || []).find((a) => {
    const name = String(a.name || a.content_file_name || "")
      .toLowerCase()
      .replace(/\s+/g, "_");
    return names.includes(name) || (prefix && name.startsWith(prefix));
  });
}

async function rewriteInlineImages(fsArticle, snowArticle, files) {
  let html = decodeEntities(
    fsArticle.description || snowArticle.html || composeArticleHtml(snowArticle)
  );
  const original = html;
  const attachments = fsArticle.attachments || [];

  for (const file of files) {
    const match = findFsAttachment(file, attachments);
    const url =
      match &&
      (match.attachment_url ||
        match.attachment_url_for_export ||
        match.canonical_url);
    if (!url) continue;

    html = html.replace(/<img\b[^>]*>/gi, (tag) => {
      const decodedTag = decodeEntities(tag);
      if (!decodedTag.toLowerCase().includes(file.sys_id.toLowerCase())) {
        return tag;
      }
      if (/\bsrc=/i.test(decodedTag)) {
        return decodedTag.replace(
          /\bsrc=(["'])[\s\S]*?\1/i,
          `src=$1${url}$1`
        );
      }
      return decodedTag.replace(/<img/i, `<img src="${url}"`);
    });

    html = html.replace(
      new RegExp(
        `(?:https?:\\/\\/[^"'\\s]+)?\\/?sys_attachment\\.do\\?sys_id=${file.sys_id}[^"'\\s]*`,
        "gi"
      ),
      url
    );
  }

  if (html === original) return fsArticle;

  const updated = await fsJson(
    "PUT",
    `/api/v2/solutions/articles/${fsArticle.id}`,
    { description: html }
  );
  return updated.article;
}

function extractAttachmentSysIds(html) {
  const decoded = decodeEntities(html);
  const ids = new Set();
  const patterns = [
    /sys_attachment\.do\?sys_id=([a-f0-9]{32})/gi,
    /\/api\/now\/attachment\/([a-f0-9]{32})/gi,
  ];
  for (const pattern of patterns) {
    let m;
    while ((m = pattern.exec(decoded))) ids.add(m[1].toLowerCase());
  }
  return ids;
}

function toHtml(text) {
  const value = String(text || "").trim();
  if (!value) return "<p>(empty ServiceNow article body)</p>";
  if (/<[a-z][\s\S]*>/i.test(value)) return value;
  return `<p>${escapeHtml(value)}</p>`;
}

function escapeHtml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function sanitizeFileName(name) {
  return String(name || "file.bin").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_");
}

function snowValue(field) {
  if (field == null) return "";
  if (typeof field === "object") return field.value || field.display_value || "";
  return String(field);
}

function removeDownloadDir(articleSysId) {
  if (!articleSysId) return;
  const dir = path.join(DOWNLOAD_DIR, String(articleSysId).toLowerCase());
  fs.rmSync(dir, { recursive: true, force: true });
  const alt = path.join(DOWNLOAD_DIR, String(articleSysId));
  if (alt !== dir) fs.rmSync(alt, { recursive: true, force: true });
}

function loadPreviousResults(filePath) {
  if (!fs.existsSync(filePath)) return [];
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return rows
    .map((row) => {
      const mapped = mapRow(row);
      return {
        ok: String(mapped.ok || "").toLowerCase() === "yes",
        sysId: String(mapped.sys_id || "").trim().toLowerCase(),
        table: mapped.table || "",
        number: mapped.number || "",
        title: mapped.title || "",
        freshserviceId: mapped.freshservice_id || "",
        attachments: mapped.attachments || "",
        url: mapped.url || "",
        error: mapped.error || "",
      };
    })
    .filter((row) => row.sysId);
}

function upsertResult(results, result) {
  const id = String(result.sysId || "").toLowerCase();
  result.sysId = id;
  const idx = results.findIndex(
    (row) => String(row.sysId).toLowerCase() === id
  );
  if (idx >= 0) results[idx] = { ...results[idx], ...result };
  else results.push(result);
}

function writeResultsCsv(filePath, results) {
  const header = [
    "ok",
    "sys_id",
    "table",
    "number",
    "title",
    "freshservice_id",
    "attachments",
    "url",
    "error",
  ];
  const lines = [header.join(",")];
  for (const row of results) {
    lines.push(
      [
        row.ok ? "yes" : "no",
        csv(row.sysId),
        csv(row.table),
        csv(row.number),
        csv(row.title),
        csv(row.freshserviceId),
        csv(row.attachments),
        csv(row.url),
        csv(row.error),
      ].join(",")
    );
  }
  fs.writeFileSync(filePath, lines.join("\n"));
}

function csv(value) {
  const text = value == null ? "" : String(value);
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

async function snowGet(pathname, query = {}) {
  const url = new URL(pathname, SNOW_INSTANCE);
  for (const [key, value] of Object.entries(query)) {
    url.searchParams.set(key, String(value));
  }
  const res = await request(url, {
    headers: {
      Authorization: snowAuth,
      Accept: "application/json",
    },
  });
  const body = await readJson(res);
  if (!res.ok) {
    throw new Error(
      `ServiceNow GET ${pathname} failed ${res.status}: ${snippet(body)}`
    );
  }
  return body.result;
}

async function snowGetOrNull(pathname, query = {}) {
  try {
    return await snowGet(pathname, query);
  } catch (err) {
    if (/failed 404/.test(err.message)) return null;
    throw err;
  }
}

async function snowDownload(pathname) {
  const url = new URL(pathname, SNOW_INSTANCE);
  const res = await request(url, {
    headers: {
      Authorization: snowAuth,
      Accept: "*/*",
    },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(
      `ServiceNow download ${pathname} failed ${res.status}: ${snippet(body)}`
    );
  }
  return Buffer.from(await res.arrayBuffer());
}

async function fsGet(pathname) {
  const res = await request(`https://${FS_DOMAIN}${pathname}`, {
    headers: {
      Authorization: fsAuth,
      Accept: "application/json",
    },
  });
  const body = await readJson(res);
  if (!res.ok) {
    throw new Error(
      `Freshservice GET ${pathname} failed ${res.status}: ${snippet(body)}`
    );
  }
  return body;
}

async function fsJson(method, pathname, payload) {
  const res = await request(`https://${FS_DOMAIN}${pathname}`, {
    method,
    headers: {
      Authorization: fsAuth,
      Accept: "application/json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  const body = await readJson(res);
  if (!res.ok) {
    throw new Error(
      `Freshservice ${method} ${pathname} failed ${res.status}: ${snippet(body)}`
    );
  }
  return body;
}

async function fsForm(method, pathname, form) {
  const res = await request(`https://${FS_DOMAIN}${pathname}`, {
    method,
    headers: {
      Authorization: fsAuth,
      Accept: "application/json",
    },
    body: form,
  });
  const body = await readJson(res);
  if (!res.ok) {
    throw new Error(
      `Freshservice ${method} ${pathname} failed ${res.status}: ${snippet(body)}`
    );
  }
  return body;
}

async function request(url, options, attempt = 1) {
  const res = await fetch(url, options);
  if (res.status === 429 && attempt < 5) {
    const wait = Number(res.headers.get("retry-after") || 2) * 1000;
    await delay(wait);
    return request(url, options, attempt + 1);
  }
  return res;
}

async function readJson(res) {
  const text = await res.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return { raw: text };
  }
}

function snippet(body) {
  const text = typeof body === "string" ? body : JSON.stringify(body);
  return text.slice(0, 500);
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function basicAuth(user, password) {
  return `Basic ${Buffer.from(`${user}:${password}`).toString("base64")}`;
}

function required(name) {
  const value = (process.env[name] || "").trim();
  if (!value) {
    throw new Error(`Missing ${name}. Fill it in .env before running.`);
  }
  return value;
}

function normalizeSnowInstance(value) {
  let host = value.replace(/\/+$/, "");
  if (!/^https?:\/\//i.test(host)) host = `https://${host}`;
  return host;
}

function normalizeFsDomain(value) {
  return value
    .replace(/^https?:\/\//i, "")
    .replace(/\/+$/, "")
    .replace(/\/api\/v2.*$/, "");
}

function loadEnv(filePath) {
  if (!fs.existsSync(filePath)) return;
  for (const line of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq < 1) continue;
    const key = trimmed.slice(0, eq).trim();
    let val = trimmed.slice(eq + 1).trim();
    if (
      (val.startsWith('"') && val.endsWith('"')) ||
      (val.startsWith("'") && val.endsWith("'"))
    ) {
      val = val.slice(1, -1);
    }
    if (process.env[key] == null) process.env[key] = val;
  }
}

main().catch((err) => {
  console.error("\nMigration failed:");
  console.error(err.message || err);
  process.exit(1);
});
