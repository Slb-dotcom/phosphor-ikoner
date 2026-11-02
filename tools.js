// Fanerne Værktøjer, Bibliotek og Tjek i panelet, plus genvejs-baren (AltGr+K).
// Kører i samme runtime som index.html og commands.js og bruger derfra: state, toast, insertSvg, escapeHtml, hasPptApi.
(() => {
const $ = (id) => document.getElementById(id);
const run = (fn) => PowerPoint.run(fn);
const TEXTY = ["GeometricShape", "TextBox", "Placeholder"];
const CM = 28.3465;
const themeHex = (name, fallback) => state.theme.find((c) => c.name === name)?.hex || fallback;
const num = (v, d) => { const n = parseFloat(String(v).replace(",", ".")); return isNaN(n) ? d : n; };
const fail = (e) => { console.error(e); toast("Det lykkedes ikke: " + (e?.message || "ukendt fejl")); };

// ---------- Faner ----------
function showTab(tab) {
  document.body.dataset.tab = tab;
  document.querySelectorAll("#tabs button").forEach((b) => b.setAttribute("aria-selected", b.dataset.tab === tab));
  document.querySelectorAll(".view").forEach((v) => { v.hidden = v.id !== "view-" + tab; });
  try { localStorage.setItem("kvadrant.tab", tab); } catch {}
  if (tab === "library") renderLibrary();
}
async function openTab(tab) {
  try { await Office.addin.showAsTaskpane(); } catch (e) { console.error(e); }
  showTab(tab);
}

// Tekst i figurerne på en slide: [{ shape, text }]
async function slideTexts(ctx, slide) {
  slide.shapes.load("items/id,items/name,items/type");
  await ctx.sync();
  const out = [];
  for (const s of slide.shapes.items.filter((x) => TEXTY.includes(x.type))) {
    const r = s.textFrame.textRange;
    r.load("text");
    try { await ctx.sync(); out.push({ shape: s, text: r.text.trim() }); } catch { /* figuren har ingen tekst */ }
  }
  return out;
}
const titleOf = (texts) => {
  const real = texts.filter((t) => t.text && !/^\d{1,2}$/.test(t.text));
  return (real.find((t) => /title|titel/i.test(t.shape.name)) || real[0])?.text.replace(/\s+/g, " ") || "";
};

// ---------- Agenda ----------
let agendaItems = [];
async function agendaScan() {
  $("agendaList").innerHTML = '<div class="hint">Læser slides …</div>';
  try {
    agendaItems = await run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const tags = slides.items.map((s) => { s.layout.load("name"); const t = s.tags.getItemOrNullObject("KVADRANT"); t.load("value"); return t; });
      await ctx.sync();
      const found = [];
      for (const [i, s] of slides.items.entries()) {
        if (!tags[i].isNullObject && tags[i].value === "agenda") continue;
        const texts = (await slideTexts(ctx, s)).filter((t) => t.text);
        const title = titleOf(texts);
        if (!title) continue;
        // Afsnits-slide: layout hedder fx "Section Header", eller en kort slide med et nummer som "01".
        const divider = /section|afsnit|divider|sektion|kapitel|chapter/i.test(s.layout.name) ||
          (texts.length <= 3 && texts.some((t) => /^\d{1,2}$/.test(t.text)));
        found.push({ id: s.id, index: i, title, on: divider });
      }
      // Ingen afsnits-slides: vis alle titler, så man selv kan vælge.
      return found.some((f) => f.on) ? found.filter((f) => f.on) : found;
    });
  } catch (e) { fail(e); agendaItems = []; }
  renderAgenda();
}
function renderAgenda() {
  const list = $("agendaList");
  if (!agendaItems.length) { list.innerHTML = '<div class="hint">Fandt ingen slides med titler.</div>'; $("agendaInsert").disabled = true; return; }
  list.innerHTML = agendaItems.map((it, i) => `<div class="agenda-item">
    <span class="no">${it.index + 1}</span><input type="checkbox" data-i="${i}" ${it.on ? "checked" : ""} aria-label="Med i agendaen">
    <input type="text" data-i="${i}" value="${escapeHtml(it.title)}" aria-label="Agendapunkt"></div>`).join("");
  $("agendaInsert").disabled = false;
}

async function agendaInsert() {
  const chosen = agendaItems.filter((it) => it.on);
  if (!chosen.length) { toast("Vælg mindst ét punkt"); return; }
  const each = $("agendaEach").checked;
  $("agendaInsert").disabled = true;
  try {
    await run(async (ctx) => {
      // Fjern agendaer, Kvadrant har lavet før.
      let slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const tags = slides.items.map((s) => { const t = s.tags.getItemOrNullObject("KVADRANT"); t.load("value"); return t; });
      await ctx.sync();
      slides.items.forEach((s, i) => { if (!tags[i].isNullObject && tags[i].value === "agenda") s.delete(); });
      await ctx.sync();

      slides = ctx.presentation.slides;
      slides.load("items/id");
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items/id");
      await ctx.sync();
      const ids = slides.items.map((s) => s.id);
      const ref = slides.items.find((s) => s.id === chosen[0].id) || slides.items[0];
      ref.slideMaster.load("id");
      ref.slideMaster.layouts.load("items/id,items/name");
      await ctx.sync();
      const layouts = ref.slideMaster.layouts.items;
      const layout = layouts.find((l) => /agenda/i.test(l.name)) || layouts.find((l) => /title only|kun titel|titel alene/i.test(l.name))
        || layouts.find((l) => /blank|tom/i.test(l.name)) || layouts[0];
      const after = sel.items[0] ? ids.indexOf(sel.items[0].id) + 1 : 1;
      const jobs = each ? chosen.map((c, k) => ({ at: ids.indexOf(c.id), active: k })).filter((j) => j.at >= 0)
        : [{ at: after, active: -1 }];
      jobs.sort((a, b) => b.at - a.at);   // bagfra, så de øvrige pladser ikke flytter sig
      for (const j of jobs) await addAgendaSlide(ctx, layout.id, ref.slideMaster.id, chosen, j.active, j.at);
    });
    toast(each ? `Indsat ${chosen.length} agendaer` : "Agenda indsat");
    agendaScan();
  } catch (e) { fail(e); } finally { $("agendaInsert").disabled = false; }
}

async function addAgendaSlide(ctx, layoutId, slideMasterId, items, active, at) {
  ctx.presentation.slides.add({ layoutId, slideMasterId });
  await ctx.sync();
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const slide = slides.items[slides.items.length - 1];
  slide.moveTo(at);
  slide.tags.add("KVADRANT", "agenda");
  slide.shapes.load("items/id,items/name,items/type");
  await ctx.sync();
  // Titlen i layoutets titelfelt; tomme indholdsfelter fjernes, sidefod og sidetal bliver.
  let titled = false;
  for (const s of slide.shapes.items.filter((x) => x.type === "Placeholder")) {
    if (!titled && /title|titel/i.test(s.name)) { s.textFrame.textRange.text = "Agenda"; titled = true; }
    else if (!/footer|sidefod|slide number|slidenummer|date|dato/i.test(s.name)) s.delete();
  }
  const { w, h } = state.slideSize;
  const left = w * 0.08;
  if (!titled) {
    const t = slide.shapes.addTextBox("Agenda", { left, top: h * 0.08, width: w * 0.84, height: 48 });
    t.textFrame.textRange.font.size = 28;
    t.textFrame.textRange.font.bold = true;
  }
  const ink = themeHex("Tekst 1", "27272C"), accent = themeHex("Tekst 2", "262139"), muted = "A3A1B3";
  const top = h * 0.26, rowH = Math.min(44, (h * 0.62) / items.length), size = Math.round(Math.max(11, Math.min(20, rowH * 0.45)));
  items.forEach((it, i) => {
    const on = active < 0 || i === active;
    const rows = [[String(i + 1).padStart(2, "0"), left, 44, on ? accent : muted], [it.title, left + 54, w * 0.7, on ? ink : muted]];
    for (const [text, x, width, color] of rows) {
      const box = slide.shapes.addTextBox(text, { left: x, top: top + i * rowH, width, height: rowH });
      box.name = `Agenda ${i + 1}`;
      box.textFrame.autoSizeSetting = "AutoSizeNone";
      box.textFrame.verticalAlignment = "Middle";
      const f = box.textFrame.textRange.font;
      f.size = size;
      f.color = "#" + color;
      f.bold = i === active;
    }
  });
  await ctx.sync();
}

// ---------- Stempler ----------
const STAMPS = [
  { id: "stampDraft", label: "Udkast", text: "UDKAST" },
  { id: "stampConfidential", label: "Fortroligt", text: "FORTROLIGT" },
  { id: "stampDiscussion", label: "Til diskussion", text: "TIL DISKUSSION" },
  { id: "stampSource", label: "Kilde", text: "Kilde: ", source: true },
  { id: "stampSourceEn", label: "Source", text: "Source: ", source: true },
];
async function addStamp({ text, source }) {
  await run(async (ctx) => {
    const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
    const { w, h } = state.slideSize;
    const box = source
      ? slide.shapes.addTextBox(text, { left: w * 0.045, top: h - 34, width: w * 0.6, height: 18 })
      : slide.shapes.addTextBox(text, { left: w - 36 - 150, top: 18, width: 150, height: 24 });
    box.name = source ? "Kilde" : "Stempel " + text;
    box.textFrame.autoSizeSetting = "AutoSizeNone";
    const f = box.textFrame.textRange.font;
    if (source) {
      f.size = 8;
      f.color = "#7A7A7A";
    } else {
      const color = "#" + themeHex("Accent 6", "E2562C");
      f.size = 10; f.bold = true; f.color = color;
      box.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
      box.textFrame.verticalAlignment = "Middle";
      box.lineFormat.visible = true;
      box.lineFormat.color = color;
      box.lineFormat.weight = 1.25;
    }
    box.load("id");
    await ctx.sync();
    slide.setSelectedShapes([box.id]);
    await ctx.sync();
  });
}

// ---------- Harvey balls ----------
const HARVEY = [0, 25, 50, 75, 100];
function harveySvg(p, hex, pt) {
  const r = 116, c = 128, a = (p / 100) * 2 * Math.PI;
  const x = (c + r * Math.sin(a)).toFixed(2), y = (c - r * Math.cos(a)).toFixed(2);
  const fill = p <= 0 ? "" : p >= 100 ? `<circle cx="${c}" cy="${c}" r="${r}"/>`
    : `<path d="M${c} ${c} L${c} ${c - r} A${r} ${r} 0 ${p > 50 ? 1 : 0} 1 ${x} ${y} Z"/>`;
  const size = pt ? `width="${pt}pt" height="${pt}pt"` : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256" ${size}><circle cx="${c}" cy="${c}" r="${r}" fill="#FFFFFF" stroke="#${hex}" stroke-width="12"/><g fill="#${hex}">${fill}</g></svg>`;
}
async function addHarvey(p) {
  const pt = 0.7 * CM, { w, h } = state.slideSize;
  await insertSvg(harveySvg(p, state.hex, pt), { imageLeft: (w - pt) / 2, imageTop: (h - pt) / 2, imageWidth: pt, imageHeight: pt });
  if (hasPptApi()) {
    await run(async (ctx) => {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load("items/id");
      await ctx.sync();
      const s = sel.items[sel.items.length - 1];
      if (s) s.name = `Harvey ${p}%`;
      await ctx.sync();
    }).catch(() => {});
  }
}

// ---------- Kvadrant-tjek ----------
const rgb = (hex) => [0, 2, 4].map((i) => parseInt(hex.substr(i, 2), 16));
// Lysere eller mørkere udgave af en temafarve (som i PowerPoints farvevælger) tæller som temafarve.
function onPalette(hex, palette) {
  const c = rgb(hex);
  if (Math.max(...c) - Math.min(...c) < 10) return true;   // sort, hvid og gråtoner
  return palette.some((p) => {
    const b = rgb(p);
    if (c.every((v, i) => Math.abs(v - b[i]) < 6)) return true;
    const tint = c.map((v, i) => (b[i] >= 250 ? null : (v - b[i]) / (255 - b[i]))).filter((t) => t !== null);
    const shade = c.map((v, i) => (b[i] <= 5 ? null : v / b[i])).filter((t) => t !== null);
    const same = (ts) => ts.length && Math.max(...ts) - Math.min(...ts) < 0.06 && ts[0] >= -0.02 && ts[0] <= 1.02;
    return same(tint) || same(shade);
  });
}
function nearest(hex, palette) {
  const c = rgb(hex);
  return palette.reduce((best, p) => {
    const d = rgb(p).reduce((n, v, i) => n + (v - c[i]) ** 2, 0);
    return d < best.d ? { p, d } : best;
  }, { p: palette[0], d: Infinity }).p;
}

const CHECKS = {
  font: { title: "Skrifttyper uden for temaet", fix: "Brug temaets skrift" },
  color: { title: "Farver uden for paletten", fix: "Brug nærmeste temafarve" },
  off: { title: "Uden for sliden", fix: "Flyt ind på sliden" },
  empty: { title: "Tomme pladsholdere", fix: "Slet dem" },
};
let issues = null;

async function runCheck() {
  $("checkRun").disabled = true;
  $("checkStatus").textContent = "Læser temaet …";
  try {
    if (!state.fonts.length || !state.theme.length) {
      const p = await readPresentation(0).catch(() => ({}));
      if (p.fonts?.length) state.fonts = p.fonts;
      if (p.theme) applyTheme(p.theme, p.themeName);
    }
    const families = state.fonts.map((f) => f.split(" ")[0].toLowerCase());
    const palette = state.theme.map((c) => c.hex);
    const { w, h } = state.slideSize;
    issues = { font: [], color: [], off: [], empty: [] };
    await run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      for (const [i, slide] of slides.items.entries()) {
        $("checkStatus").textContent = `Slide ${i + 1} af ${slides.items.length} …`;
        slide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
        await ctx.sync();
        const base = (s) => ({ slide: i + 1, slideId: slide.id, shapeId: s.id, name: s.name });
        for (const s of slide.shapes.items) {
          if (s.left < -2 || s.top < -2 || s.left + s.width > w + 2 || s.top + s.height > h + 2) issues.off.push({ ...base(s), detail: "" });
        }
        for (const s of slide.shapes.items.filter((x) => TEXTY.includes(x.type))) {
          const r = s.textFrame.textRange;
          r.load("text,font/name,font/color");
          const fill = s.type === "GeometricShape" ? s.fill : null;
          fill?.load("type,foregroundColor");
          try { await ctx.sync(); } catch { continue; }
          const text = r.text.trim();
          if (!text) { if (s.type === "Placeholder") issues.empty.push({ ...base(s), detail: "" }); }
          else {
            const font = r.font.name;
            if (font && families.length && !families.includes(font.split(" ")[0].toLowerCase())) issues.font.push({ ...base(s), detail: font, title: /title|titel/i.test(s.name) });
            const color = r.font.color?.replace("#", "");
            if (color && palette.length && !onPalette(color, palette)) issues.color.push({ ...base(s), detail: "#" + color, kind: "text" });
          }
          const fc = fill?.type === "Solid" && fill.foregroundColor?.replace("#", "");
          if (fc && palette.length && !onPalette(fc, palette)) issues.color.push({ ...base(s), detail: "#" + fc + " (fyld)", kind: "fill", hex: fc });
        }
      }
    });
    $("checkStatus").textContent = "";
    renderIssues();
  } catch (e) { fail(e); $("checkStatus").textContent = ""; } finally { $("checkRun").disabled = false; }
}

function renderIssues() {
  const total = Object.values(issues).reduce((n, l) => n + l.length, 0);
  const fonts = state.fonts.join(" / ");
  $("checkResult").innerHTML = (total ? "" : `<div class="card"><b>Alt ser fint ud.</b><div class="hint">Temaets skrifter: ${escapeHtml(fonts)}</div></div>`) +
    Object.entries(CHECKS).filter(([k]) => issues[k].length).map(([k, c]) => `<div class="card issue-group">
      <div class="issue-head"><b>${c.title}<span class="badge">${issues[k].length}</span></b><button class="ghost" data-fix="${k}">${c.fix}</button></div>
      ${issues[k].slice(0, 40).map((it, i) => `<div class="issue"><span>Slide ${it.slide} · ${escapeHtml(it.name)}${it.detail ? " · " + escapeHtml(it.detail) : ""}</span><button data-go="${k}:${i}">Gå til</button></div>`).join("")}
      ${issues[k].length > 40 ? `<div class="hint">… og ${issues[k].length - 40} mere</div>` : ""}
      ${k === "font" ? `<div class="hint">Temaets skrifter: ${escapeHtml(fonts)}</div>` : ""}
    </div>`).join("");
}

async function goTo(it) {
  await run(async (ctx) => {
    ctx.presentation.setSelectedSlides([it.slideId]);
    await ctx.sync();
    ctx.presentation.slides.getItem(it.slideId).setSelectedShapes([it.shapeId]);
    await ctx.sync();
  }).catch(fail);
}

async function fixAll(kind) {
  const list = issues[kind];
  const [major, minor] = [state.fonts[0], state.fonts[1] || state.fonts[0]];
  const palette = state.theme.map((c) => c.hex);
  const { w, h } = state.slideSize;
  try {
    await run(async (ctx) => {
      for (const it of list) {
        const s = ctx.presentation.slides.getItem(it.slideId).shapes.getItem(it.shapeId);
        if (kind === "font") s.textFrame.textRange.font.name = it.title ? major : minor;
        if (kind === "color" && it.kind === "text") s.textFrame.textRange.font.color = "#" + nearest(it.detail.slice(1), palette);
        if (kind === "color" && it.kind === "fill") s.fill.setSolidColor("#" + nearest(it.hex, palette));
        if (kind === "empty") s.delete();
        if (kind === "off") {
          s.load("left,top,width,height");
          await ctx.sync();
          if (s.width <= w) s.left = Math.min(Math.max(0, s.left), w - s.width);
          if (s.height <= h) s.top = Math.min(Math.max(0, s.top), h - s.height);
        }
      }
      await ctx.sync();
    });
    toast(`Rettet ${list.length}`);
    issues[kind] = [];
    renderIssues();
  } catch (e) { fail(e); }
}

// ---------- Bibliotek (byggeklodser) ----------
// Gemmes lokalt i IndexedDB: "meta" (titel, miniature) og "data" (slidens .pptx som base64) hver for sig,
// så listen ikke skal læse de store filer.
const db = new Promise((res, rej) => {
  const r = indexedDB.open("kvadrant", 1);
  r.onupgradeneeded = () => { r.result.createObjectStore("meta", { keyPath: "id" }); r.result.createObjectStore("data", { keyPath: "id" }); };
  r.onsuccess = () => res(r.result);
  r.onerror = () => rej(r.error);
});
async function dbDo(store, mode, fn) {
  const d = await db;
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode);
    const req = fn(t.objectStore(store));
    t.oncomplete = () => res(req?.result);
    t.onerror = () => rej(t.error);
  });
}

async function saveSlides(all) {
  const source = decodeURIComponent((Office.context.document.url || "Præsentation").split(/[\\/]/).pop()).replace(/\.pptx?$/i, "");
  let ids;
  try {
    ids = await run(async (ctx) => {
      const coll = all ? ctx.presentation.slides : ctx.presentation.getSelectedSlides();
      coll.load("items/id");
      await ctx.sync();
      return coll.items.map((s) => s.id);
    });
  } catch (e) { fail(e); return; }
  if (!ids.length) { toast("Markér de slides, der skal gemmes"); return; }
  const buttons = [$("libSaveSel"), $("libSaveAll")];
  buttons.forEach((b) => { b.disabled = true; });
  try {
    for (const [k, id] of ids.entries()) {
      toast(`Gemmer ${k + 1} af ${ids.length} …`);
      const rec = await run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(id);
        const data = slide.exportAsBase64();
        const img = slide.getImageAsBase64({ width: 400 });
        await ctx.sync();
        const title = titleOf(await slideTexts(ctx, slide)) || `Slide ${k + 1}`;
        return { data: data.value, thumb: img.value, title };
      });
      const key = crypto.randomUUID();
      await dbDo("data", "readwrite", (st) => st.put({ id: key, data: rec.data }));
      await dbDo("meta", "readwrite", (st) => st.put({ id: key, title: rec.title, thumb: rec.thumb, source, added: Date.now() + k }));
    }
    toast(`Gemt ${ids.length} slides i biblioteket`);
  } catch (e) { fail(e); } finally { buttons.forEach((b) => { b.disabled = false; }); }
  renderLibrary();
}

async function renderLibrary() {
  let items = [];
  try { items = await dbDo("meta", "readonly", (st) => st.getAll()) || []; } catch (e) { console.error(e); }
  const q = $("libSearch").value.trim().toLowerCase();
  items = items.filter((it) => !q || (it.title + " " + it.source).toLowerCase().includes(q)).sort((a, b) => a.added - b.added);
  if (!items.length) {
    $("libList").innerHTML = `<div class="hint" style="padding:4px 2px">${q ? "Intet matcher søgningen." : "Biblioteket er tomt. Åbn en præsentation med jeres standard-slides, markér dem i oversigten til venstre, og tryk “Gem markerede slides”."}</div>`;
    return;
  }
  const groups = {};
  items.forEach((it) => (groups[it.source] ||= []).push(it));
  $("libList").innerHTML = Object.entries(groups).map(([src, list]) => `<div class="lib-source">${escapeHtml(src)} · ${list.length}</div>
    <div class="lib-grid" style="margin-top:8px">${list.map((it) => `<button class="block" data-id="${it.id}" title="Indsæt: ${escapeHtml(it.title)}">
      <img src="data:image/png;base64,${it.thumb}" alt=""><div class="t">${escapeHtml(it.title)}</div>
      <span class="x" data-del="${it.id}" title="Fjern fra biblioteket">×</span></button>`).join("")}</div>`).join("");
}

async function insertBlock(id) {
  try {
    const rec = await dbDo("data", "readonly", (st) => st.get(id));
    if (!rec) return;
    await run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items/id");
      await ctx.sync();
      const opts = { formatting: $("libTheme").checked ? "UseDestinationTheme" : "KeepSourceFormatting" };
      if (sel.items[0]) opts.targetSlideId = sel.items[0].id;
      ctx.presentation.insertSlidesFromBase64(rec.data, opts);
      await ctx.sync();
    });
    toast("Slide indsat");
  } catch (e) { fail(e); }
}
async function deleteBlock(id) {
  await dbDo("meta", "readwrite", (st) => st.delete(id));
  await dbDo("data", "readwrite", (st) => st.delete(id));
  renderLibrary();
}

// ---------- Genvejs-baren (AltGr+K) ----------
let bar = null;
function showBar() {
  return new Promise((done) => {
    if (bar) { bar.close(); bar = null; done(); return; }
    const url = new URL("keys.html", location.href).href;
    Office.context.ui.displayDialogAsync(url, { width: 55, height: 26, displayInIframe: false }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { console.error(r.error); done(); return; }
      bar = r.value;
      bar.addEventHandler(Office.EventType.DialogMessageReceived, (m) => {
        let msg = {};
        try { msg = JSON.parse(m.message); } catch {}
        if (!msg.keep && bar) { bar.close(); bar = null; }
        if (msg.cmd && typeof window[msg.cmd] === "function") window[msg.cmd]();
      });
      bar.addEventHandler(Office.EventType.DialogEventReceived, () => { bar = null; });
      done();
    });
  });
}

// ---------- Kommandoer til bånd, genveje og baren ----------
const cmd = (fn) => async (event) => { try { await fn(); } catch (e) { fail(e); } finally { event?.completed?.(); } };
const TOOL_COMMANDS = {
  showBar: cmd(showBar),
  openAgenda: cmd(async () => { await openTab("tools"); agendaScan(); }),
  openTools: cmd(() => openTab("tools")),
  openLibrary: cmd(() => openTab("library")),
  openCheck: cmd(async () => { await openTab("check"); runCheck(); }),
  ...Object.fromEntries(STAMPS.map((s) => [s.id, cmd(() => addStamp(s))])),
  ...Object.fromEntries(HARVEY.map((p) => ["harvey" + p, cmd(() => addHarvey(p))])),
};

function wireTools() {
  $("tabs").addEventListener("click", (e) => { const b = e.target.closest("button[data-tab]"); if (b) showTab(b.dataset.tab); });
  $("agendaScan").addEventListener("click", agendaScan);
  $("agendaInsert").addEventListener("click", agendaInsert);
  $("agendaList").addEventListener("input", (e) => {
    const it = agendaItems[e.target.dataset.i];
    if (!it) return;
    if (e.target.type === "checkbox") it.on = e.target.checked; else it.title = e.target.value;
  });
  $("gridRun").addEventListener("click", () => window.kvadrantGrid({
    cols: Math.round(num($("gridCols").value, 0)), gap: num($("gridGap").value, 0.3) * CM, same: $("gridSame").checked,
  }));
  $("stamps").innerHTML = STAMPS.map((s) => `<button class="chip" data-stamp="${s.id}">${s.label}</button>`).join("");
  $("stamps").addEventListener("click", (e) => { const b = e.target.closest("[data-stamp]"); if (b) addStamp(STAMPS.find((s) => s.id === b.dataset.stamp)).catch(fail); });
  const renderHarvey = () => { $("harvey").innerHTML = HARVEY.map((p) => `<button class="chip" data-p="${p}" title="${p} %">${harveySvg(p, state.hex)}${p}%</button>`).join(""); };
  renderHarvey();
  document.querySelector(".top").addEventListener("click", () => setTimeout(renderHarvey, 0));
  $("harvey").addEventListener("click", (e) => { const b = e.target.closest("[data-p]"); if (b) addHarvey(+b.dataset.p).catch(fail); });

  $("libSaveSel").addEventListener("click", () => saveSlides(false));
  $("libSaveAll").addEventListener("click", () => saveSlides(true));
  $("libSearch").addEventListener("input", renderLibrary);
  $("libList").addEventListener("click", (e) => {
    const del = e.target.closest("[data-del]");
    if (del) { e.stopPropagation(); deleteBlock(del.dataset.del); return; }
    const b = e.target.closest(".block");
    if (b) insertBlock(b.dataset.id);
  });

  $("checkRun").addEventListener("click", runCheck);
  $("checkResult").addEventListener("click", (e) => {
    const go = e.target.closest("[data-go]"), fix = e.target.closest("[data-fix]");
    if (go) { const [k, i] = go.dataset.go.split(":"); goTo(issues[k][+i]); }
    if (fix) fixAll(fix.dataset.fix);
  });

  let tab = "icons";
  try { tab = localStorage.getItem("kvadrant.tab") || "icons"; } catch {}
  showTab(["icons", "tools", "library", "check"].includes(tab) ? tab : "icons");
}

wireTools();
window.kvadrantPane = { showTab };
const register = () => {
  for (const [name, fn] of Object.entries(TOOL_COMMANDS)) {
    window[name] = fn;
    Office.actions?.associate?.(name, fn);
  }
};
if (window.Office) Office.onReady(register); else register();
})();
