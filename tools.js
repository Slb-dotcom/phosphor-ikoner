// Motoren bag pop-up-vinduet (tools.html) og genvejs-baren (keys.html).
// Kører i panelets runtime (index.html) og bruger derfra: state, toast, insertSvg, readPresentation, applyTheme, hasPptApi.
// Vinduerne har ingen PowerPoint-API, så de sender { id, op, args } hertil og får { id, ok, result | error } tilbage.
(() => {
const run = (fn) => PowerPoint.run(fn);
const TEXTY = ["GeometricShape", "TextBox", "Placeholder"];
const CM = 28.3465;
const themeHex = (name, fallback) => state.theme.find((c) => c.name === name)?.hex || fallback;

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
function agendaScan() {
  return run(async (ctx) => {
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
}

async function agendaInsert(items, each) {
  const chosen = items.filter((it) => it.on && it.title.trim());
  if (!chosen.length) throw new Error("Vælg mindst ét punkt");
  const tpl = await dbDo("data", "readonly", (st) => st.get("_agendaTemplate")).catch(() => null);
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
    for (const j of jobs) {
      if (tpl) await addFromTemplate(ctx, tpl.data, chosen, j.active, j.at);
      else await addAgendaSlide(ctx, layout.id, ref.slideMaster.id, chosen, j.active, j.at);
    }
  });
  return each ? chosen.length : 1;
}

// Gemmer den markerede slide som agenda-design (samme IndexedDB som slidebanken, id "_agendaTemplate").
async function agendaTemplateSave() {
  const rec = await run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    await ctx.sync();
    const slide = sel.items[0];
    if (!slide) throw new Error("Markér agenda-sliden i oversigten til venstre");
    const data = slide.exportAsBase64();
    const img = slide.getImageAsBase64({ width: 400 });
    await ctx.sync();
    return { data: data.value, thumb: img.value, title: titleOf(await slideTexts(ctx, slide)) || "Agenda" };
  });
  await dbDo("data", "readwrite", (st) => st.put({ id: "_agendaTemplate", data: rec.data }));
  await dbDo("meta", "readwrite", (st) => st.put({ id: "_agendaTemplate", title: rec.title, thumb: rec.thumb, source: "_", added: Date.now() }));
  return rec.title;
}

// Agenda i skabelonens design: sliden sættes ind, og listen ("01 …", "02 …") skrives om med de valgte punkter.
// Formateringen fra skabelonens første linje bevares; ved et aktivt afsnit gråtones de andre linjer.
async function addFromTemplate(ctx, data, items, active, at) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const before = new Set(slides.items.map((s) => s.id));
  const opts = { formatting: "KeepSourceFormatting" };
  if (at > 0 && slides.items[at - 1]) opts.targetSlideId = slides.items[at - 1].id;
  ctx.presentation.insertSlidesFromBase64(data, opts);
  await ctx.sync();
  const after = ctx.presentation.slides;
  after.load("items/id");
  await ctx.sync();
  const slide = after.items.find((s) => !before.has(s.id));
  if (!slide) return;
  if (at === 0) slide.moveTo(0);
  slide.tags.add("KVADRANT", "agenda");
  await ctx.sync();
  const texts = (await slideTexts(ctx, slide)).filter((t) => t.text);
  const lines = (t) => t.text.split(/\r\n|\r|\n|\v/).filter((l) => l.trim());
  const list = texts.find((t) => lines(t).length > 1 && /^\s*0?1\b/.test(t.text))
    || texts.filter((t) => !/^agenda$/i.test(t.text)).sort((a, b) => lines(b).length - lines(a).length || b.text.length - a.text.length)[0];
  if (!list) return;
  const out = items.map((it, i) => `${String(i + 1).padStart(2, "0")} ${it.title.trim()}`);
  const range = list.shape.textFrame.textRange;
  range.text = out.join("\n");
  await ctx.sync();
  if (active < 0) return;
  let pos = 0;
  out.forEach((l, i) => {
    if (i !== active) range.getSubstring(pos, l.length).font.color = "#A3A1B3";
    pos += l.length + 1;
  });
  await ctx.sync();
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
    const rows = [[String(i + 1).padStart(2, "0"), left, 44, on ? accent : muted], [it.title.trim(), left + 54, w * 0.7, on ? ink : muted]];
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
const STAMPS = {
  stampDraft: { text: "UDKAST" },
  stampConfidential: { text: "FORTROLIGT" },
  stampDiscussion: { text: "TIL DISKUSSION" },
  stampSource: { text: "Kilde: ", source: true },
  stampSourceEn: { text: "Source: ", source: true },
};
async function addStamp(id) {
  const { text, source } = STAMPS[id];
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

// ---------- Harvey ball ----------
// p = 0-100 (alle værdier, fx 33); fyldet starter kl. 12 og går med uret.
function harveySvg(p, hex, pt) {
  const r = 116, c = 128, a = (p / 100) * 2 * Math.PI;
  const x = (c + r * Math.sin(a)).toFixed(2), y = (c - r * Math.cos(a)).toFixed(2);
  const fill = p <= 0 ? "" : p >= 100 ? `<circle cx="${c}" cy="${c}" r="${r}"/>`
    : `<path d="M${c} ${c} L${c} ${c - r} A${r} ${r} 0 ${p > 50 ? 1 : 0} 1 ${x} ${y} Z"/>`;
  const size = pt ? `width="${pt}pt" height="${pt}pt"` : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256" ${size}><circle cx="${c}" cy="${c}" r="${r}" fill="#FFFFFF" stroke="#${hex}" stroke-width="12"/><g fill="#${hex}">${fill}</g></svg>`;
}
async function addHarvey(p, hex = state.hex, cm = 0.7) {
  p = Math.max(0, Math.min(100, Number(p) || 0));
  const pt = cm * CM, { w, h } = state.slideSize;
  await insertSvg(harveySvg(p, hex, pt), { imageLeft: (w - pt) / 2, imageTop: (h - pt) / 2, imageWidth: pt, imageHeight: pt });
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

async function ensureTheme() {
  if (state.fonts.length && state.theme.length) return;
  const p = await readPresentation(0).catch(() => ({}));
  if (p.slideSize) state.slideSize = p.slideSize;
  if (p.fonts?.length) state.fonts = p.fonts;
  if (p.theme) applyTheme(p.theme, p.themeName);
}

async function runCheck(progress) {
  await ensureTheme();
  const families = state.fonts.map((f) => f.split(" ")[0].toLowerCase());
  const palette = state.theme.map((c) => c.hex);
  const { w, h } = state.slideSize;
  const issues = { font: [], color: [], off: [], empty: [] };
  await run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    for (const [i, slide] of slides.items.entries()) {
      progress(`Slide ${i + 1} af ${slides.items.length} …`);
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
          if (color && palette.length && !onPalette(color, palette)) issues.color.push({ ...base(s), detail: "#" + color, kind: "text", hex: color });
        }
        const fc = fill?.type === "Solid" && fill.foregroundColor?.replace("#", "");
        if (fc && palette.length && !onPalette(fc, palette)) issues.color.push({ ...base(s), detail: "#" + fc + " (fyld)", kind: "fill", hex: fc });
      }
    }
  });
  return { issues, fonts: state.fonts };
}

function goTo(it) {
  return run(async (ctx) => {
    ctx.presentation.setSelectedSlides([it.slideId]);
    await ctx.sync();
    ctx.presentation.slides.getItem(it.slideId).setSelectedShapes([it.shapeId]);
    await ctx.sync();
  });
}

async function fixAll(kind, list) {
  const [major, minor] = [state.fonts[0], state.fonts[1] || state.fonts[0]];
  const palette = state.theme.map((c) => c.hex);
  const { w, h } = state.slideSize;
  await run(async (ctx) => {
    for (const it of list) {
      const s = ctx.presentation.slides.getItem(it.slideId).shapes.getItem(it.shapeId);
      if (kind === "font") s.textFrame.textRange.font.name = it.title ? major : minor;
      if (kind === "color" && it.kind === "text") s.textFrame.textRange.font.color = "#" + nearest(it.hex, palette);
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
  return list.length;
}

// ---------- Bibliotek (byggeklodser) ----------
// IndexedDB "kvadrant": "meta" (titel, miniature) og "data" (slidens .pptx som base64) hver for sig, så listen er hurtig.
// tools.html læser listen direkte (samme domæne); her gemmes og indsættes.
let dbPromise;
function db() {
  return (dbPromise ||= new Promise((res, rej) => {
    const r = indexedDB.open("kvadrant", 1);
    r.onupgradeneeded = () => { r.result.createObjectStore("meta", { keyPath: "id" }); r.result.createObjectStore("data", { keyPath: "id" }); };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  }));
}
async function dbDo(store, mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode);
    const req = fn(t.objectStore(store));
    t.oncomplete = () => res(req?.result);
    t.onerror = () => rej(t.error);
  });
}

async function saveSlides(all, progress) {
  const source = decodeURIComponent((Office.context.document.url || "Præsentation").split(/[\\/]/).pop()).replace(/\.pptx?$/i, "");
  const ids = await run(async (ctx) => {
    const coll = all ? ctx.presentation.slides : ctx.presentation.getSelectedSlides();
    coll.load("items/id");
    await ctx.sync();
    return coll.items.map((s) => s.id);
  });
  if (!ids.length) throw new Error("Markér de slides, der skal gemmes");
  for (const [k, id] of ids.entries()) {
    progress(`Gemmer ${k + 1} af ${ids.length} …`);
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
  return ids.length;
}

async function insertBlock(id, useTheme) {
  const rec = await dbDo("data", "readonly", (st) => st.get(id));
  if (!rec) throw new Error("Sliden findes ikke længere i biblioteket");
  await run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    await ctx.sync();
    const opts = { formatting: useTheme ? "UseDestinationTheme" : "KeepSourceFormatting" };
    if (sel.items[0]) opts.targetSlideId = sel.items[0].id;
    ctx.presentation.insertSlidesFromBase64(rec.data, opts);
    await ctx.sync();
  });
}
async function deleteBlock(id) {
  await dbDo("meta", "readwrite", (st) => st.delete(id));
  await dbDo("data", "readwrite", (st) => st.delete(id));
}

// ---------- Det, vinduerne kan bede om ----------
const OPS = {
  context: async () => { await ensureTheme().catch(() => {}); return { theme: state.theme, hex: state.hex, fonts: state.fonts }; },
  agendaScan, agendaInsert, agendaTemplateSave,
  grid: (opts) => window.kvadrantGrid({ cols: opts.cols, gap: opts.gapCm * CM, same: opts.same }),
  stamp: addStamp,
  harvey: addHarvey,
  saveSlides: (all, progress) => saveSlides(all, progress),
  insertBlock, deleteBlock,
  check: (progress) => runCheck(progress),
  goTo, fixAll,
};
// Ops, der sender fremskridt undervejs, får en progress-funktion som sidste argument.
const WITH_PROGRESS = new Set(["saveSlides", "check"]);

// ---------- Vinduer ----------
let dialog = null;
const closeDialog = () => { try { dialog?.close(); } catch {} dialog = null; };

function openDialog(page, view, size) {
  return new Promise((done) => {
    closeDialog();
    const url = new URL(page, location.href);
    if (view) url.hash = view;
    const open = (retry) => Office.context.ui.displayDialogAsync(url.href, { ...size, displayInIframe: false }, (r) => {
      // 12007: der er allerede et vindue åbent; det forrige lukker lige nu.
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        if (retry && r.error?.code === 12007) { setTimeout(() => open(false), 400); return; }
        console.error(r.error); toast("Vinduet kunne ikke åbnes"); done(); return;
      }
      const d = (dialog = r.value);
      const reply = (msg) => { try { d.messageChild(JSON.stringify(msg)); } catch (e) { console.error(e); } };
      d.addEventHandler(Office.EventType.DialogMessageReceived, async (m) => {
        let msg = {};
        try { msg = JSON.parse(m.message); } catch {}
        if (msg.close) { closeDialog(); return; }
        if (msg.cmd) {   // genvejs-baren
          if (!msg.keep) closeDialog();
          if (typeof window[msg.cmd] === "function") window[msg.cmd]();
          return;
        }
        if (!msg.op || !OPS[msg.op]) return;
        const args = [...(msg.args || [])];
        if (WITH_PROGRESS.has(msg.op)) args.push((text) => reply({ id: msg.id, progress: text }));
        try { reply({ id: msg.id, ok: true, result: await OPS[msg.op](...args) }); }
        catch (e) { console.error(e); reply({ id: msg.id, ok: false, error: e?.message || String(e) }); }
      });
      d.addEventHandler(Office.EventType.DialogEventReceived, () => { if (dialog === d) dialog = null; });
      done();
    });
    open(true);
  });
}

const TOOLS_SIZE = { width: 46, height: 72 };
const SMALL_SIZE = { width: 22, height: 40 };
const toolsView = (view) => openDialog("tools.html", view, view === "harvey" ? SMALL_SIZE : TOOLS_SIZE);
function showBar() {
  if (dialog) { closeDialog(); return Promise.resolve(); }   // Ctrl+Shift+K igen lukker
  return openDialog("keys.html", "", { width: 50, height: 22 });
}

// ---------- Kommandoer til bånd, genveje og baren ----------
const cmd = (fn) => async (event) => {
  try { await fn(); } catch (e) { console.error(e); toast("Det lykkedes ikke: " + (e?.message || "ukendt fejl")); }
  finally { event?.completed?.(); }
};
const TOOL_COMMANDS = {
  showBar: cmd(showBar),
  openTools: cmd(() => toolsView("agenda")),
  openAgenda: cmd(() => toolsView("agenda")),
  openGrid: cmd(() => toolsView("grid")),
  openStamps: cmd(() => toolsView("stamps")),
  openHarvey: cmd(() => toolsView("harvey")),
  openLibrary: cmd(() => toolsView("library")),
  openCheck: cmd(() => toolsView("check")),
  ...Object.fromEntries(Object.keys(STAMPS).map((id) => [id, cmd(() => addStamp(id))])),
  ...Object.fromEntries([0, 25, 50, 75, 100].map((p) => ["harvey" + p, cmd(() => addHarvey(p))])),
};

const register = () => {
  for (const [name, fn] of Object.entries(TOOL_COMMANDS)) {
    window[name] = fn;
    Office.actions?.associate?.(name, fn);
  }
};
if (window.Office) Office.onReady(register); else register();
})();
