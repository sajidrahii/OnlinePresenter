// Optional: set these if auto-detection fails (e.g. custom domain or local testing).
const CONFIG = { owner: "", repo: "", branch: "", folder: "presentations" };

const $ = (id) => document.getElementById(id);
const esc = (s) => s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const nice = (f) => f.replace(/\.pptx$/i, "").replace(/[-_]+/g, " ");
const size = (b) => (b > 1e6 ? (b / 1e6).toFixed(1) + " MB" : Math.max(1, Math.round(b / 1e3)) + " KB");
let items = [];

function repoInfo() {
  if (CONFIG.owner && CONFIG.repo) return CONFIG;
  const m = location.hostname.match(/^(.+)\.github\.io$/);
  const repo = location.pathname.split("/")[1];
  return m ? { owner: m[1], repo: repo || m[1] + ".github.io", branch: CONFIG.branch, folder: CONFIG.folder } : null;
}

async function fromManifest() {
  try { // presentations.json: [{ "file": "a.pptx", "title": "...", "description": "...", "size": 123 }]
    const r = await fetch("presentations.json", { cache: "no-cache" });
    return r.ok ? await r.json() : [];
  } catch (e) { return []; }
}
async function fromApi() { // works on github.io, or when owner/repo are set in CONFIG
  const info = repoInfo();
  if (!info) return [];
  try {
    const ref = info.branch ? "?ref=" + info.branch : "";
    const r = await fetch(`https://api.github.com/repos/${info.owner}/${info.repo}/contents/${info.folder}${ref}`);
    return r.ok ? (await r.json()).map((f) => ({ file: f.name, size: f.size })) : [];
  } catch (e) { return []; }
}
async function fromDirListing() { // works on local servers that show folder indexes
  try {
    const r = await fetch(CONFIG.folder + "/");
    if (!r.ok) return [];
    const doc = new DOMParser().parseFromString(await r.text(), "text/html");
    return [...doc.querySelectorAll("a")].map((a) => decodeURIComponent((a.getAttribute("href") || "").split("/").pop())).map((file) => ({ file }));
  } catch (e) { return []; }
}

async function load() {
  $("grid").innerHTML = '<div class="skel"></div>'.repeat(3);
  const [manifest, api, dir] = await Promise.all([fromManifest(), fromApi(), fromDirListing()]);
  const map = new Map();
  [...dir, ...api, ...manifest].forEach((m) => { // later sources add titles/descriptions
    if (m && /\.pptx$/i.test(m.file || "")) map.set(m.file, { ...map.get(m.file), ...m });
  });
  items = [...map.values()].map((m) => ({ file: m.file, size: m.size || 0, title: m.title || nice(m.file), desc: m.description || "" }));
  items.sort((a, b) => a.title.localeCompare(b.title));
  render();
}

function render() {
  const q = $("q").value.trim().toLowerCase();
  const list = items.filter((i) => (i.title + " " + i.desc).toLowerCase().includes(q));
  $("count").textContent = items.length ? `(${list.length})` : "";
  $("msg").hidden = list.length > 0;
  $("msg").innerHTML = items.length
    ? "No presentations match your search."
    : `<b>No presentations yet.</b> Add <code>.pptx</code> files to the <code>${CONFIG.folder}</code> folder in your repository and refresh this page.`;
  $("grid").innerHTML = list.map((i) =>
    `<a class="card" href="viewer.html?file=${encodeURIComponent(i.file)}">
      <div class="tile">${esc(i.title)}</div>
      <div class="info"><b>${esc(i.title)}</b><span>${esc(i.desc || (i.size ? size(i.size) : "PowerPoint"))}</span></div></a>`).join("");
}

$("q").addEventListener("input", render);
load();
