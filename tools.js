// Motoren bag pop-up-vinduet (tools.html) og genvejs-baren (keys.html).
// Kører i panelets runtime (index.html) og bruger derfra: state, toast, insertSvg, readPresentation, applyTheme, hasPptApi.
// Vinduerne har ingen PowerPoint-API, så de sender { id, op, args } hertil og får { id, ok, result | error } tilbage.
(() => {
const run = (fn) => PowerPoint.run(fn);
const TOOLS_BUILD = "2026-11-03b";   // skift ved hver upload af tools.html
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
      const divider = /section|afsnit|divider|sektion|kapitel|chapter|breaker/i.test(s.layout.name) ||
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
    const layout = layouts.find((l) => /agenda|table of contents|indhold/i.test(l.name)) || layouts.find((l) => /title only|kun titel|titel alene/i.test(l.name))
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

// ---------- Kvadrant-bokse ----------
// Faste farveregler: underoverskrift = Tekst 2 ved 90 % gennemsigtighed, baggrundsboks = Baggrund 2 ved 75 %,
// COMMENTARY-overskrift = Tekst 2 (0 %) med hvide versaler, COMMENTARY-tekst = Baggrund 2 (0 %), topjusteret.
// Alle bokse har runde hjørner og 10 pt tekst i temaets skrift.
const BOXES = {
  boxSubheader: { name: "Underoverskrift", fill: "Dark2", transparency: 0.9, text: "Sub-header", align: "Center", vAlign: "Middle", w: 0.55, h: 20 },
  boxBackground: { name: "Baggrundsboks", fill: "Light2", transparency: 0.75, text: "", w: 0.55, h: 0.5 },
  boxCommentaryHead: { name: "COMMENTARY", fill: "Dark2", transparency: 0, text: "COMMENTARY", white: true, align: "Center", vAlign: "Middle", w: 0.25, h: 20, right: true },
  boxCommentaryBody: { name: "Commentary", fill: "Light2", transparency: 0, text: "Kommentar", vAlign: "Top", bullets: true, w: 0.25, h: 0.55, right: true, below: true },
};
// Temafarven fra sliden (PowerPointApi 1.10), ellers fra det læste tema, ellers Kvadrants standard.
const THEME_FALLBACK = { Dark1: ["Tekst 1", "27272C"], Dark2: ["Tekst 2", "262139"], Light2: ["Baggrund 2", "E0E3EC"] };
async function themeColor(ctx, slide, name) {
  try { const r = slide.themeColorScheme.getThemeColor(name); await ctx.sync(); if (r.value) return r.value.replace("#", ""); } catch {}
  return themeHex(...THEME_FALLBACK[name]);
}
async function addBox(id) {
  const b = BOXES[id];
  await run(async (ctx) => {
    const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
    const { w, h } = state.slideSize;
    const bw = b.w * w, bh = b.h <= 1 ? b.h * h : b.h;
    const left = b.right ? w * 0.72 : w * 0.06, top = b.below ? h * 0.27 + 22 : h * 0.27;
    const shape = slide.shapes.addGeometricShape("RoundRectangle", { left, top, width: bw, height: bh });
    shape.name = "Kvadrant " + b.name;
    shape.fill.setSolidColor("#" + (await themeColor(ctx, slide, b.fill)));
    shape.fill.transparency = b.transparency;
    shape.lineFormat.visible = false;
    const tf = shape.textFrame;
    tf.textRange.text = b.text;
    tf.verticalAlignment = b.vAlign || "Top";
    tf.textRange.font.size = 10;
    tf.textRange.font.color = b.white ? "#FFFFFF" : "#" + (await themeColor(ctx, slide, "Dark1"));
    tf.textRange.paragraphFormat.horizontalAlignment = b.align || "Left";
    if (b.bullets) tf.textRange.paragraphFormat.bulletFormat.visible = true;
    shape.load("id");
    await ctx.sync();
    try { shape.adjustments.set(0, Math.min(0.5, 4 / Math.min(bw, bh))); await ctx.sync(); } catch { /* ældre PowerPoint: standardrunding */ }
    slide.setSelectedShapes([shape.id]);
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
  // Kvadrant-regler: overskrift 25 pt og højst 2 linjer, brødtekst i bokse 10 pt, "COMMENTARY" med versaler,
  // tekst i bokse topjusteret, kilde på indholdsslides og ikke for tekst-tunge slides.
  const issues = { font: [], color: [], off: [], empty: [], title: [], titleLong: [], body: [], commentary: [], align: [], source: [], heavy: [] };
  await run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    slides.items.forEach((s) => s.layout.load("name"));
    await ctx.sync();
    for (const [i, slide] of slides.items.entries()) {
      progress(`Slide ${i + 1} af ${slides.items.length} …`);
      slide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      await ctx.sync();
      const base = (s) => ({ slide: i + 1, slideId: slide.id, shapeId: s.id, name: s.name });
      for (const s of slide.shapes.items) {
        if (s.left < -2 || s.top < -2 || s.left + s.width > w + 2 || s.top + s.height > h + 2) issues.off.push({ ...base(s), detail: "" });
      }
      let words = 0, hasSource = false, hasTitle = false;
      for (const s of slide.shapes.items.filter((x) => TEXTY.includes(x.type))) {
        const r = s.textFrame.textRange;
        r.load("text,font/name,font/color,font/size");
        s.textFrame.load("verticalAlignment");
        const fill = s.type === "GeometricShape" ? s.fill : null;
        fill?.load("type,foregroundColor");
        try { await ctx.sync(); } catch { continue; }
        const text = r.text.trim(), size = r.font.size, isTitle = /title|titel/i.test(s.name);
        if (!text) { if (s.type === "Placeholder") issues.empty.push({ ...base(s), detail: "" }); }
        else {
          const font = r.font.name;
          if (font && families.length && !families.includes(font.split(" ")[0].toLowerCase())) issues.font.push({ ...base(s), detail: font, title: isTitle });
          const color = r.font.color?.replace("#", "");
          if (color && palette.length && !onPalette(color, palette)) issues.color.push({ ...base(s), detail: "#" + color, kind: "text", hex: color });
          words += (text.match(/\S+/g) || []).length;
          if (/^\s*(source|sources|kilde|kilder)\s*:/im.test(text)) hasSource = true;
          if (isTitle) {
            hasTitle = true;
            if (size && size !== 25) issues.title.push({ ...base(s), detail: `${size} pt` });
            const perLine = Math.max(10, s.width / ((size || 25) * 0.52));   // ca. tegn pr. linje
            const lines = text.split(/\r|\n|\v/).reduce((n, l) => n + Math.max(1, Math.ceil(l.length / perLine)), 0);
            if (lines > 2) issues.titleLong.push({ ...base(s), detail: `ca. ${lines} linjer` });
          }
          if (/^commentary$/i.test(text) && text !== "COMMENTARY") issues.commentary.push({ ...base(s), detail: text });
          const boxText = !isTitle && s.type !== "Placeholder" && text.length >= 40 && !/^\s*(source|kilde)/i.test(text);
          if (boxText && size && size >= 6 && size < 18 && size !== 10) issues.body.push({ ...base(s), detail: `${size} pt` });
          if (boxText && fill?.type === "Solid" && text.length >= 120 && !/^Top/.test(s.textFrame.verticalAlignment || "")) issues.align.push({ ...base(s), detail: "" });
        }
        const fc = fill?.type === "Solid" && fill.foregroundColor?.replace("#", "");
        if (fc && palette.length && !onPalette(fc, palette)) issues.color.push({ ...base(s), detail: "#" + fc + " (fyld)", kind: "fill", hex: fc });
      }
      const content = hasTitle && slide.shapes.items.length >= 4 && !/breaker|cover|end|agenda|table of contents|section|afsnit/i.test(slide.layout.name);
      if (content && !hasSource) issues.source.push({ slide: i + 1, slideId: slide.id, name: "Ingen kilde", detail: "" });
      if (words > 220) issues.heavy.push({ slide: i + 1, slideId: slide.id, name: `${words} ord`, detail: "" });
    }
  });
  return { issues, fonts: state.fonts };
}

function goTo(it) {
  return run(async (ctx) => {
    ctx.presentation.setSelectedSlides([it.slideId]);
    await ctx.sync();
    if (!it.shapeId) return;   // fund for hele sliden (fx manglende kilde)
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
      if (kind === "title") s.textFrame.textRange.font.size = 25;
      if (kind === "body") s.textFrame.textRange.font.size = 10;
      if (kind === "commentary") s.textFrame.textRange.text = "COMMENTARY";
      if (kind === "align") s.textFrame.verticalAlignment = "Top";
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

// ---------- Fælles slidebank ----------
// Slidebanken er filen "Kvadrant Slidebank.pptx" i den fælles mappe. Hver slide har skjulte mærker (tags):
// KV_ID "GTM-03-004" (praksis-trin-løbenummer), KV_PRACTICE, KV_STEP og KV_DESC (beskrivelsen, personen skrev).
// Når filen er åben, læses den ind i IndexedDB ("bank:<id>"), så alle præsentationer kan indsætte derfra.
const PRACTICES = [["CT", "Commercial Transformation"], ["GTM", "Go-To-Market"], ["TS", "Transaction Services"]];
const STEPS = [["01", "Introduktion"], ["02", "Breakers"], ["03", "Metoder & tilgang"], ["04", "Proces & tidsplan"],
  ["05", "Analyser & frameworks"], ["06", "Cases"], ["07", "Team & investering"], ["08", "Afslutning & appendix"]];
const BANK_TAGS = ["KV_ID", "KV_PRACTICE", "KV_STEP", "KV_DESC"];
const isBankDoc = () => /slidebank/i.test(decodeURIComponent(Office.context.document.url || ""));

async function slideTags(ctx, slides) {
  const t = slides.map((s) => BANK_TAGS.map((k) => { const x = s.tags.getItemOrNullObject(k); x.load("value"); return x; }));
  await ctx.sync();
  return t.map((row) => Object.fromEntries(BANK_TAGS.map((k, i) => [k, row[i].isNullObject ? "" : row[i].value])));
}

// Mærker slides med næste ledige nummer i praksis+trin, og sorterer banken efter nummer (umærkede slides, fx forsiden, først).
async function bankAssign(ctx, slideIds, practice, step, desc) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const tags = await slideTags(ctx, slides.items);
  const prefix = `${practice}-${step}-`;
  let seq = Math.max(0, ...tags.map((t) => t.KV_ID).filter((id) => id.startsWith(prefix)).map((id) => +id.slice(prefix.length) || 0));
  const ids = [];
  for (const sid of slideIds) {
    const s = ctx.presentation.slides.getItem(sid);
    const id = prefix + String(++seq).padStart(3, "0");
    s.tags.add("KV_ID", id); s.tags.add("KV_PRACTICE", practice); s.tags.add("KV_STEP", step); s.tags.add("KV_DESC", desc || "");
    ids.push(id);
  }
  await ctx.sync();
  await bankSort(ctx);
  return ids;
}
async function bankSort(ctx) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const tags = await slideTags(ctx, slides.items);
  const order = slides.items.map((s, i) => ({ s, id: tags[i].KV_ID, i }))
    .sort((a, b) => (!!a.id - !!b.id) || (a.id ? a.id.localeCompare(b.id) : a.i - b.i));
  order.forEach((o, i) => o.s.moveTo(i));
  await ctx.sync();
}

// Gemmer én bank-slide (eksport + miniature) i IndexedDB.
async function storeBankSlide(slideId, info) {
  const rec = await run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(slideId);
    const data = slide.exportAsBase64();
    const img = slide.getImageAsBase64({ width: 400 });
    await ctx.sync();
    return { data: data.value, thumb: img.value, title: titleOf(await slideTexts(ctx, slide)) };
  });
  const key = "bank:" + info.KV_ID;
  await dbDo("data", "readwrite", (st) => st.put({ id: key, data: rec.data }));
  await dbDo("meta", "readwrite", (st) => st.put({ id: key, source: "bank", code: info.KV_ID, practice: info.KV_PRACTICE,
    step: info.KV_STEP, desc: info.KV_DESC, title: rec.title || info.KV_ID, thumb: rec.thumb, added: Date.now() }));
}

async function bankStatus() {
  const metas = await dbDo("meta", "readonly", (st) => st.getAll());
  let synced = null;
  try { synced = localStorage.getItem("kvadrant.bankSynced"); } catch {}
  let bankUrl = null;
  try { bankUrl = localStorage.getItem("kvadrant.bankUrl"); } catch {}
  return { isBank: isBankDoc(), practices: PRACTICES, steps: STEPS, synced, canOpen: !!bankUrl,
    count: metas.filter((m) => m.source === "bank").length, pending: metas.filter((m) => m.source === "pending").length };
}

async function bankSync(progress) {
  if (!isBankDoc()) throw new Error("Åbn Kvadrant Slidebank.pptx fra den fælles mappe først");
  const recs = await run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const tags = await slideTags(ctx, slides.items);
    return slides.items.map((s, i) => ({ slideId: s.id, ...tags[i] })).filter((r) => r.KV_ID);
  });
  const old = (await dbDo("meta", "readonly", (st) => st.getAll())).filter((m) => m.source === "bank");
  for (const m of old) { await dbDo("meta", "readwrite", (st) => st.delete(m.id)); await dbDo("data", "readwrite", (st) => st.delete(m.id)); }
  for (const [k, r] of recs.entries()) { progress(`Læser ${k + 1} af ${recs.length} …`); await storeBankSlide(r.slideId, r); }
  try { localStorage.setItem("kvadrant.bankSynced", new Date().toISOString()); } catch {}
  return recs.length;
}

// I slidebank-filen: mærk de markerede slides. Andre steder: læg dem i kø, til nogen åbner slidebanken.
async function bankAdd(practice, step, desc) {
  if (isBankDoc()) {
    const done = await run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items/id");
      await ctx.sync();
      if (!sel.items.length) throw new Error("Markér de slides, der skal i slidebanken");
      const sids = sel.items.map((s) => s.id);
      const ids = await bankAssign(ctx, sids, practice, step, desc);
      return sids.map((sid, i) => ({ slideId: sid, KV_ID: ids[i], KV_PRACTICE: practice, KV_STEP: step, KV_DESC: desc }));
    });
    for (const d of done) await storeBankSlide(d.slideId, d);
    return { added: done.map((d) => d.KV_ID), queued: 0 };
  }
  const ids = await run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    await ctx.sync();
    return sel.items.map((s) => s.id);
  });
  if (!ids.length) throw new Error("Markér de slides, der skal i slidebanken");
  for (const sid of ids) {
    const rec = await run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(sid);
      const data = slide.exportAsBase64();
      const img = slide.getImageAsBase64({ width: 400 });
      await ctx.sync();
      return { data: data.value, thumb: img.value, title: titleOf(await slideTexts(ctx, slide)) };
    });
    const key = "pending:" + crypto.randomUUID();
    await dbDo("data", "readwrite", (st) => st.put({ id: key, data: rec.data }));
    await dbDo("meta", "readwrite", (st) => st.put({ id: key, source: "pending", practice, step, desc, title: rec.title, thumb: rec.thumb, added: Date.now() }));
  }
  return { added: [], queued: ids.length };
}

// I slidebank-filen: sæt slides fra køen ind, mærk dem og gem dem.
async function bankFlush(progress) {
  if (!isBankDoc()) throw new Error("Åbn Kvadrant Slidebank.pptx fra den fælles mappe først");
  const pending = (await dbDo("meta", "readonly", (st) => st.getAll())).filter((m) => m.source === "pending");
  const added = [];
  for (const [k, p] of pending.entries()) {
    progress(`Sætter ${k + 1} af ${pending.length} ind …`);
    const rec = await dbDo("data", "readonly", (st) => st.get(p.id));
    const info = await run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const before = new Set(slides.items.map((s) => s.id));
      ctx.presentation.insertSlidesFromBase64(rec.data, { formatting: "KeepSourceFormatting", targetSlideId: slides.items[slides.items.length - 1].id });
      await ctx.sync();
      const after = ctx.presentation.slides;
      after.load("items/id");
      await ctx.sync();
      const sid = after.items.find((s) => !before.has(s.id))?.id;
      const [id] = await bankAssign(ctx, [sid], p.practice, p.step, p.desc);
      return { slideId: sid, KV_ID: id, KV_PRACTICE: p.practice, KV_STEP: p.step, KV_DESC: p.desc };
    });
    await storeBankSlide(info.slideId, info);
    await dbDo("meta", "readwrite", (st) => st.delete(p.id));
    await dbDo("data", "readwrite", (st) => st.delete(p.id));
    added.push(info.KV_ID);
  }
  return added;
}

// ---------- Talformat ----------
// "1234567" / "1.234.567" / "1,2345e6" -> "1,2 mio. kr." (dansk) eller "DKK 1.2m" (engelsk).
function parseNumber(s) {
  s = s.replace(/[\s ]/g, "");
  const dot = s.lastIndexOf("."), comma = s.lastIndexOf(",");
  let dec = null;
  if (dot >= 0 && comma >= 0) dec = dot > comma ? "." : ",";
  else if (dot >= 0 || comma >= 0) {
    const sep = dot >= 0 ? "." : ",", parts = s.split(sep);
    dec = parts.length === 2 && parts[1].length !== 3 ? sep : null;   // "1.234" og "1.234.567" er tusinder, "12,5" er decimaler
  }
  const thou = dec === "." ? "," : dec === "," ? "." : /[.,]/.exec(s)?.[0];
  if (thou) s = s.split(thou).join("");
  if (dec) s = s.replace(dec, ".");
  return parseFloat(s);
}
const UNITS = { da: [[1e9, "mia."], [1e6, "mio."], [1e3, "t."]], en: [[1e9, "bn"], [1e6, "m"], [1e3, "k"]] };
function formatNumber(v, { lang = "da", unit = "auto", decimals = 1, currency = "" }) {
  const units = UNITS[lang];
  const [div, word] = unit === "auto" ? (units.find(([d]) => Math.abs(v) >= d) || [1, ""]) : unit === "none" ? [1, ""] : units.find(([d]) => d === +unit) || [1, ""];
  const n = (v / div).toLocaleString(lang === "da" ? "da-DK" : "en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  if (lang === "da") return [n, word, currency].filter(Boolean).join(" ");
  return `${currency ? currency + " " : ""}${n}${word}`;
}
// Tal + evt. enhed og valuta, der allerede står efter tallet ("12,5 mio kr"), som så indgår i værdien og erstattes.
const NUM_RX = /(?<![\w.,])(-?\d{1,3}(?:[.,  ]\d{3})+(?:[.,]\d+)?|-?\d+(?:[.,]\d+)?)(?![.,]?\d)(\s*(?:mia\.?|mio\.?|millioner|million|mn|bn|t\.|k|m)(?![a-zæøå]))?(\s*(?:kr\.?|DKK|EUR|USD|€)(?![a-zæøå]))?/gi;
const SCALE = (w) => (/^(mia|bn)/i.test(w) ? 1e9 : /^(mio|mill|mn|m)/i.test(w) ? 1e6 : /^(t|k)/i.test(w) ? 1e3 : 1);
function reformatText(text, opts) {
  let count = 0;
  const out = text.replace(NUM_RX, (m, num, scale = "", cur = "") => {
    let v = parseNumber(num);
    if (isNaN(v)) return m;
    scale = scale.trim(); cur = cur.trim();
    if (!scale && !cur && /^\d{4}$/.test(num) && v >= 1900 && v <= 2100) return m;   // årstal
    if (!scale && !cur && Math.abs(v) < 1000) return m;                             // små tal (antal, procenter) lades være
    v *= SCALE(scale);
    count++;
    const currency = opts.currency === "keep" ? (cur ? cur.replace(/^kr$/i, "kr.") : "") : opts.currency;
    return formatNumber(v, { ...opts, currency });
  });
  return { out, count };
}
async function formatNumbers(opts) {
  return run(async (ctx) => {
    let ranges = [];
    const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
    tr.load("length,text");
    await ctx.sync();
    if (!tr.isNullObject && tr.length > 0) ranges = [tr];
    else {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load("items/type");
      await ctx.sync();
      ranges = sel.items.filter((s) => TEXTY.includes(s.type)).map((s) => s.textFrame.textRange);
      ranges.forEach((r) => r.load("text"));
      await ctx.sync();
    }
    if (!ranges.length) throw new Error("Markér teksten eller tekstboksene med tallene");
    let total = 0;
    for (const r of ranges) {
      const { out, count } = reformatText(r.text, opts);
      if (count) { r.text = out; total += count; }
    }
    await ctx.sync();
    return total;
  });
}

// ---------- Rens før afsendelse ----------
// Office.js kan ikke slette noter og kommentarer, så filen læses, renses og åbnes som en ny præsentation (originalen røres ikke).
async function readFileBytes() {
  const ok = (r) => r.status === Office.AsyncResultStatus.Succeeded;
  const file = await new Promise((res, rej) => Office.context.document.getFileAsync(Office.FileType.Compressed,
    { sliceSize: 4194304 }, (r) => (ok(r) ? res(r.value) : rej(r.error))));
  const slices = [];
  try {
    for (let i = 0; i < file.sliceCount; i++) slices.push(await new Promise((res, rej) => file.getSliceAsync(i, (r) => (ok(r) ? res(r.value.data) : rej(r.error)))));
  } finally { file.closeAsync(); }
  const bytes = new Uint8Array(slices.reduce((n, s) => n + s.length, 0));
  let o = 0;
  for (const s of slices) { bytes.set(s, o); o += s.length; }
  return bytes;
}
async function cleanCopy(progress) {
  progress("Læser præsentationen …");
  const files = fflate.unzipSync(await readFileBytes());
  const str = (k) => fflate.strFromU8(files[k]), put = (k, s) => { files[k] = fflate.strToU8(s); };
  const gone = (k) => /^ppt\/(notesSlides\/|comments\/|commentAuthors\.xml$|authors\.xml$)/.test(k);
  const notes = Object.keys(files).filter((k) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k)).length;
  const comments = Object.keys(files).filter((k) => /^ppt\/comments\/.*\.xml$/.test(k)).length;
  progress("Fjerner noter og kommentarer …");
  for (const k of Object.keys(files)) if (gone(k)) delete files[k];
  const dropRel = /Type="[^"]*\/(notesSlide|comments|commentAuthors|authors)"/;
  for (const k of Object.keys(files).filter((k) => /^ppt\/(slides\/_rels\/slide\d+\.xml|_rels\/presentation\.xml)\.rels$/.test(k))) {
    const removed = [];
    put(k, str(k).replace(/<Relationship\b[^>]*\/>/g, (m) => { if (!dropRel.test(m)) return m; removed.push(/Id="([^"]+)"/.exec(m)[1]); return ""; }));
    const part = k.replace("_rels/", "").replace(/\.rels$/, "");
    if (removed.length && files[part]) {
      let xml = str(part);
      for (const id of removed) xml = xml.replace(new RegExp(`<p:ext\\b(?:(?!</p:ext>)[\\s\\S])*r:id="${id}"(?:(?!</p:ext>)[\\s\\S])*</p:ext>`, "g"), "");
      put(part, xml.replace(/<p:extLst>\s*<\/p:extLst>/g, ""));
    }
  }
  put("[Content_Types].xml", str("[Content_Types].xml").replace(/<Override PartName="\/ppt\/(notesSlides|comments)\/[^"]*"[^>]*\/>|<Override PartName="\/ppt\/(commentAuthors|authors)\.xml"[^>]*\/>/g, ""));
  progress("Åbner den rensede kopi …");
  const zipped = fflate.zipSync(files, { level: 6 });
  let b64 = "";
  for (let i = 0; i < zipped.length; i += 0x8000) b64 += String.fromCharCode.apply(null, zipped.subarray(i, i + 0x8000));
  await PowerPoint.createPresentation(btoa(b64));
  return { notes, comments };
}

// ---------- Grafer efter Kvadrant-regler ----------
// Office.js kan ikke formatere diagrammer, men en slide kan eksporteres som .pptx, hvor diagrammet er XML.
// XML'en rettes til Kvadrant-stil, og sliden sættes ind igen som en "spejling" lige efter originalen, som så slettes.
// Gælder PowerPoints egne diagrammer (ikke think-cell).
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart", DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_PALETTE = ["tx2", "accent5", "accent1", "accent3", "accent2", "accent6", "accent4"];

function formatChartXml(xml, { size = null, gridlines = false, legendBottom = true } = {}) {   // size = null: behold tekststørrelsen
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const frag = (s) => doc.importNode(new DOMParser().parseFromString(
    `<r xmlns:c="${CHART_NS}" xmlns:a="${DML_NS}">${s}</r>`, "application/xml").documentElement.firstElementChild, true);
  const kids = (el, name) => [...el.children].filter((c) => c.namespaceURI === CHART_NS && c.localName === name);
  const all = (name, ns = CHART_NS) => [...doc.getElementsByTagNameNS(ns, name)];
  // Erstat <c:spPr> eller indsæt den efter det sidste af de elementer, der skal stå før den.
  const setSpPr = (parent, inner, after) => {
    const sp = frag(`<c:spPr>${inner}</c:spPr>`);
    const old = kids(parent, "spPr")[0];
    if (old) { parent.replaceChild(sp, old); return; }
    const anchor = [...parent.children].filter((c) => after.includes(c.localName)).pop();
    anchor ? anchor.after(sp) : parent.prepend(sp);
  };
  const solid = (clr) => `<a:solidFill><a:schemeClr val="${clr}"/></a:solidFill>`;
  const NONE = "<a:noFill/><a:ln><a:noFill/></a:ln>";

  // Ingen baggrund eller ramme om diagram og plotområde.
  const space = doc.documentElement, chart = kids(space, "chart")[0], plot = chart && kids(chart, "plotArea")[0];
  if (chart) setSpPr(space, `${solid("bg1")}<a:ln><a:noFill/></a:ln>`, ["chart"]);   // Kvadrant: hvid baggrund i alle grafer
  if (plot) setSpPr(plot, NONE, ["layout", "barChart", "bar3DChart", "lineChart", "line3DChart", "pieChart", "pie3DChart", "doughnutChart",
    "areaChart", "area3DChart", "scatterChart", "radarChart", "bubbleChart", "catAx", "valAx", "dateAx", "serAx", "dTable"]);

  // Serier i temaets farver: søjler/flader fyldt, linjer som streg.
  for (const type of ["barChart", "bar3DChart", "areaChart", "area3DChart", "lineChart", "line3DChart", "scatterChart", "radarChart"]) {
    for (const group of all(type)) {
      const line = /line|scatter|radar/i.test(type);
      kids(group, "ser").forEach((ser, i) => {
        if (kids(ser, "dPt").length) return;   // bevidst fremhævede punkter: seriens farver bevares
        const clr = CHART_PALETTE[i % CHART_PALETTE.length];
        setSpPr(ser, line ? `<a:ln w="28575" cap="rnd">${solid(clr)}<a:round/></a:ln>` : `${solid(clr)}<a:ln><a:noFill/></a:ln>`, ["idx", "order", "tx"]);
      });
      const gap = kids(group, "gapWidth")[0];
      if (gap) gap.setAttribute("val", "75");
    }
  }
  // Lagkage/donut: hvert stykke sin farve.
  for (const group of [...all("pieChart"), ...all("pie3DChart"), ...all("doughnutChart")]) {
    for (const ser of kids(group, "ser")) {
      const n = +(ser.getElementsByTagNameNS(CHART_NS, "ptCount")[0]?.getAttribute("val") || 0);
      kids(ser, "dPt").forEach((d) => d.remove());
      const anchor = [...ser.children].filter((c) => ["idx", "order", "tx", "spPr", "explosion"].includes(c.localName)).pop();
      let at = anchor;
      for (let p = 0; p < n; p++) {
        const d = frag(`<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr>${solid(CHART_PALETTE[p % CHART_PALETTE.length])}<a:ln w="12700"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></c:spPr></c:dPt>`);
        at ? at.after(d) : ser.prepend(d);
        at = d;
      }
    }
  }

  // Akser: ingen tick-mærker, ingen hjælpelinjer, kun en tynd lys linje på kategoriaksen.
  for (const name of ["catAx", "valAx", "dateAx", "serAx"]) {
    for (const ax of all(name)) {
      if (!gridlines) [...kids(ax, "majorGridlines"), ...kids(ax, "minorGridlines")].forEach((g) => g.remove());
      kids(ax, "majorTickMark").forEach((t) => t.setAttribute("val", "none"));
      kids(ax, "minorTickMark").forEach((t) => t.setAttribute("val", "none"));
      const lineXml = name === "valAx" ? NONE
        : `<a:noFill/><a:ln w="9525">${`<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill>`}</a:ln>`;
      setSpPr(ax, lineXml, ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos"]);
    }
  }

  // Forklaring nederst.
  if (legendBottom) for (const lg of all("legend")) kids(lg, "legendPos").forEach((p) => p.setAttribute("val", "b"));

  // Al tekst: temaets brødtekstskrift, grå, samme størrelse.
  for (const r of [...all("defRPr", DML_NS), ...all("rPr", DML_NS)]) {
    if (size) r.setAttribute("sz", String(Math.round(size * 100)));
    [...r.children].filter((c) => /Fill$|^latin$|^ea$|^cs$/.test(c.localName)).forEach((c) => c.remove());
    const fill = frag(`<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>`);
    const ln = [...r.children].find((c) => c.localName === "ln");
    ln ? ln.after(fill) : r.prepend(fill);
    const effects = [...r.children].filter((c) => ["effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill"].includes(c.localName)).pop() || fill;
    effects.after(frag(`<a:latin typeface="+mn-lt"/>`), frag(`<a:ea typeface="+mn-ea"/>`), frag(`<a:cs typeface="+mn-cs"/>`));
  }
  return new XMLSerializer().serializeToString(doc);
}

const b64ToBytes = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
const bytesToB64 = (bytes) => { let s = ""; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); };

async function formatCharts(opts, progress) {
  const ids = await run(async (ctx) => {
    const coll = opts.all ? ctx.presentation.slides : ctx.presentation.getSelectedSlides();
    coll.load("items/id");
    await ctx.sync();
    return coll.items.map((s) => s.id);
  });
  if (!ids.length) throw new Error("Markér de slides med grafer");
  let slidesDone = 0, charts = 0;
  const newIds = [];
  for (const [k, sid] of ids.entries()) {
    progress(`Slide ${k + 1} af ${ids.length} …`);
    const b64 = await run(async (ctx) => { const r = ctx.presentation.slides.getItem(sid).exportAsBase64(); await ctx.sync(); return r.value; });
    const files = fflate.unzipSync(b64ToBytes(b64));
    const keys = Object.keys(files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f));
    if (!keys.length) continue;
    for (const f of keys) files[f] = fflate.strToU8(formatChartXml(fflate.strFromU8(files[f]), opts));
    const fixed = bytesToB64(fflate.zipSync(files, { level: 6 }));
    // Spejlingen sættes ind lige efter originalen, som derefter slettes.
    const id = await run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const before = new Set(slides.items.map((s) => s.id));
      ctx.presentation.insertSlidesFromBase64(fixed, { formatting: "UseDestinationTheme", targetSlideId: sid });
      await ctx.sync();
      const after = ctx.presentation.slides;
      after.load("items/id");
      await ctx.sync();
      const nid = after.items.find((s) => !before.has(s.id))?.id;
      if (!nid) throw new Error("Den rettede slide kunne ikke sættes ind");
      ctx.presentation.slides.getItem(sid).delete();
      await ctx.sync();
      return nid;
    });
    newIds.push(id);
    slidesDone++; charts += keys.length;
  }
  if (newIds.length) await run(async (ctx) => { ctx.presentation.setSelectedSlides(newIds); await ctx.sync(); }).catch(() => {});
  return { slides: slidesDone, charts };
}
window.kvadrantFormatChartXml = formatChartXml;   // til test

// ---------- Det, vinduerne kan bede om ----------
const OPS = {
  context: () => ({ theme: state.theme, hex: state.hex, fonts: state.fonts }),   // svarer straks; temaet læses kun, når det skal bruges (Tjek)
  agendaScan, agendaInsert, agendaTemplateSave,
  grid: (opts) => window.kvadrantGrid({ cols: opts.cols, gap: opts.gapCm * CM, same: opts.same }),
  stamp: addStamp,
  box: addBox,
  harvey: addHarvey,
  saveSlides: (all, progress) => saveSlides(all, progress),
  insertBlock, deleteBlock,
  check: (progress) => runCheck(progress),
  goTo, fixAll,
  bankStatus, bankSync: (progress) => bankSync(progress), bankAdd, bankFlush: (progress) => bankFlush(progress),
  formatNumbers, cleanCopy: (progress) => cleanCopy(progress),
  formatCharts: (opts, progress) => formatCharts(opts, progress),
};
// Ops, der sender fremskridt undervejs, får en progress-funktion som sidste argument.
const WITH_PROGRESS = new Set(["saveSlides", "check", "bankSync", "bankFlush", "cleanCopy", "formatCharts"]);

// ---------- Vinduer ----------
let dialog = null;
const closeDialog = () => { try { dialog?.close(); } catch {} dialog = null; };

function openDialog(page, view, size) {
  return new Promise((done) => {
    closeDialog();
    const url = new URL(page, location.href);
    url.searchParams.set("v", TOOLS_BUILD);   // ny version = ny adresse, så Office ikke viser en gemt, gammel side
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

const TOOLS_SIZE = { width: 96, height: 94 };   // næsten fuld skærm; Esc lukker
const SMALL_SIZE = { width: 22, height: 40 };
const toolsView = (view) => openDialog("tools.html", view, view === "harvey" ? SMALL_SIZE : TOOLS_SIZE);
function showBar() {
  if (dialog) { closeDialog(); return Promise.resolve(); }   // Ctrl+Shift+K igen lukker
  return toolsView("keys");
}

// ---------- Kommandoer til bånd, genveje og baren ----------
const cmd = (fn) => async (event) => {
  try { await fn(); } catch (e) { console.error(e); toast("Det lykkedes ikke: " + (e?.message || "ukendt fejl")); }
  finally { event?.completed?.(); }
};
const TOOL_COMMANDS = {
  showBar: cmd(showBar),
  openTools: cmd(() => toolsView("keys")),
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
  bankOnOpen();
};

// Når slidebank-filen er åben: husk dens adresse (til "Åbn slidebanken"), lad Kvadrant starte af sig selv næste gang filen
// åbnes (gælder kun denne fil), og sæt slides fra køen ind med det samme.
async function bankOnOpen() {
  if (!window.Office?.context?.document || !isBankDoc()) return;
  try { localStorage.setItem("kvadrant.bankUrl", Office.context.document.url); } catch {}
  Office.addin?.setStartupBehavior?.(Office.StartupBehavior.load).catch(() => {});
  try {
    const { pending, synced } = await bankStatus();
    if (pending) { const ids = await bankFlush(() => {}); toast(`Slidebanken: ${ids.length} nye slides sat ind (${ids.join(", ")}). Husk at gemme.`); }
    else if (!synced) await bankSync(() => {});
  } catch (e) { console.error(e); }
}

// Åbner slidebank-filen i PowerPoint (adressen huskes, første gang filen har været åben med Kvadrant).
async function openBank() {
  let url = null;
  try { url = localStorage.getItem("kvadrant.bankUrl"); } catch {}
  if (!url) throw new Error("Åbn Kvadrant Slidebank.pptx fra den fælles mappe én gang, så husker Kvadrant den");
  const target = /^https?:/i.test(url) ? "ms-powerpoint:ofe|u|" + url : "file:///" + url.replace(/\\/g, "/");
  Office.context.ui.openBrowserWindow(target);
}
OPS.openBank = openBank;

// Slidebank som mappe: vinduet (tools.html) skriver filerne; her eksporteres de markerede slides til IndexedDB ("outbox:…"),
// fordi slides er for store til at sende som beskeder mellem vindue og runtime.
OPS.exportSelected = async () => {
  const ids = await run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    await ctx.sync();
    return sel.items.map((s) => s.id);
  });
  if (!ids.length) throw new Error("Markér de slides, der skal gemmes, i oversigten til venstre");
  const keys = [];
  for (const sid of ids) {
    const rec = await run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(sid);
      const data = slide.exportAsBase64();
      const img = slide.getImageAsBase64({ width: 400 });
      await ctx.sync();
      return { data: data.value, thumb: img.value, title: titleOf(await slideTexts(ctx, slide)) };
    });
    const key = "outbox:" + crypto.randomUUID();
    await dbDo("data", "readwrite", (st) => st.put({ id: key, ...rec }));
    keys.push(key);
  }
  return keys;
};
if (window.Office) Office.onReady(register); else register();
})();
