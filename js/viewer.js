const $ = (id) => document.getElementById(id);
const FOLDER = "presentations";
const file = new URLSearchParams(location.search).get("file") || "";
let buf, prev, index = 0, ratio = 16 / 9;

function fail(msg) { $("status").hidden = false; $("status").textContent = msg; }

function build() {
  const st = $("stage");
  const w = Math.floor(Math.min(st.clientWidth, st.clientHeight * ratio));
  $("slide").innerHTML = "";
  prev?.destroy?.();
  prev = pptxPreview.init($("slide"), { width: w, height: Math.floor(w / ratio), mode: "slide" });
  return prev.preview(buf.slice(0));
}

function update() {
  $("pos").textContent = `${index + 1} / ${prev.slideCount}`;
  $("prev").disabled = index === 0;
  $("next").disabled = index >= prev.slideCount - 1;
}
function go(d) {
  const n = index + d;
  if (!prev || n < 0 || n >= prev.slideCount) return;
  index = n;
  prev.renderSingleSlide(index);
  update();
}

async function init() {
  if (!/^[\w\-. ()]+\.pptx$/i.test(file)) return fail("No valid presentation selected. Go back to the library and pick one.");
  const title = file.replace(/\.pptx$/i, "").replace(/[-_]+/g, " ");
  $("title").textContent = title;
  document.title = title + " · Slide Library";
  const url = `${FOLDER}/${encodeURIComponent(file)}`;
  $("dl").href = url;
  try {
    const r = await fetch(url);
    if (!r.ok) throw new Error("missing");
    buf = await r.arrayBuffer();
  } catch (e) { return fail(`Couldn't load "${file}". Check that it exists in the ${FOLDER} folder.`); }
  try {
    await build();
    const p = prev.pptx;
    if (p?.width && p?.height && Math.abs(p.width / p.height - ratio) > 0.01) { ratio = p.width / p.height; await build(); }
    $("status").hidden = true;
    update();
  } catch (e) { fail("This file couldn't be displayed. It may be damaged or use features the viewer doesn't support. You can still download it."); }
}

let t;
addEventListener("resize", () => { clearTimeout(t); t = setTimeout(async () => { if (!buf) return; await build(); prev.renderSingleSlide(index); update(); }, 200); });
addEventListener("keydown", (e) => {
  if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") { e.preventDefault(); go(1); }
  if (e.key === "ArrowLeft" || e.key === "PageUp") { e.preventDefault(); go(-1); }
  if (e.key.toLowerCase() === "f") toggleFs();
});
function toggleFs() { document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen?.(); }
$("prev").onclick = () => go(-1);
$("next").onclick = () => go(1);
$("fs").onclick = toggleFs;
$("copy").onclick = async () => { try { await navigator.clipboard.writeText(location.href); $("copy").textContent = "Copied"; } catch (e) { $("copy").textContent = "Copy failed"; } setTimeout(() => ($("copy").textContent = "Copy link"), 1500); };
// Touch swipe
let x0; $("stage").addEventListener("touchstart", (e) => (x0 = e.touches[0].clientX), { passive: true });
$("stage").addEventListener("touchend", (e) => { const d = e.changedTouches[0].clientX - x0; if (Math.abs(d) > 50) go(d < 0 ? 1 : -1); });
init();
