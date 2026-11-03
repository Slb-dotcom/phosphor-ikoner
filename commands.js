// Kommandoerne bag knapperne på fanen "Kvadrant" (se build-ribbon.mjs for navnene).
// Alle arbejder på de markerede figurer via PowerPoint JavaScript API (PowerPointApi 1.5+).
// Filen indlæses af index.html (delt runtime), så alt ligger i en funktion for ikke at støde sammen med panelets navne.
(() => {
// fflate (zip) hentes først, når en funktion skal bruge den, så panel og vinduer starter hurtigere.
let fflatePromise;
window.kvLoadFflate = () => (fflatePromise ||= window.fflate ? Promise.resolve() : new Promise((res, rej) => {
  const s = document.createElement("script");
  s.src = "https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js";
  s.onload = res; s.onerror = () => rej(new Error("Kunne ikke hente zip-biblioteket"));
  document.head.appendChild(s);
}));

const STEP = 0.1 * 28.3465;   // 0,1 cm i punkter, bruges af "Øg/Mindsk afstand"
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder"];

// Kører fn(shapes, ctx) på de markerede figurer og afslutter altid kommandoen.
function onShapes(min, fn) {
  return async (event) => {
    try {
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
        await ctx.sync();
        if (sel.items.length >= min) {
          await fn(sel.items, ctx);
          await ctx.sync();
        }
      });
    } catch (e) {
      console.error(e);
    } finally {
      event?.completed?.();   // genvejstaster kalder uden event
    }
  };
}

// Tekst: den markerede tekst, hvis man står i en tekstboks; ellers al tekst i de markerede figurer.
function onText(fn) {
  return async (event) => {
    try {
      await PowerPoint.run(async (ctx) => {
        let ranges = [];
        const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
        tr.load("length");
        await ctx.sync();
        if (!tr.isNullObject && tr.length > 0) ranges = [tr];
        else {
          const sel = ctx.presentation.getSelectedShapes();
          sel.load("items/type");
          await ctx.sync();
          ranges = sel.items.filter((s) => TEXT_TYPES.includes(s.type)).map((s) => s.textFrame.textRange);
        }
        if (!ranges.length) return;
        ranges.forEach((r) => r.load("font/size,font/bold,font/italic,font/underline,paragraphFormat/bulletFormat/visible"));
        await ctx.sync();
        fn(ranges);
        await ctx.sync();
      });
    } catch (e) {
      console.error(e);
    } finally {
      event?.completed?.();
    }
  };
}

// ---------- Slidestørrelse (Office.js har ingen API til den) ----------
let slideSize = null;
async function getSlideSize() {
  if (slideSize) return slideSize;
  const key = "kvadrant.slideSize." + (Office.context.document.url || "");
  try { const s = JSON.parse(localStorage.getItem(key)); if (s?.w) return (slideSize = s); } catch {}
  slideSize = { w: 960, h: 540 };
  try {
    const ok = (r) => r.status === Office.AsyncResultStatus.Succeeded;
    const file = await new Promise((res, rej) => Office.context.document.getFileAsync(Office.FileType.Compressed,
      { sliceSize: 4194304 }, (r) => (ok(r) ? res(r.value) : rej(r.error))));
    const slices = [];
    try {
      for (let i = 0; i < file.sliceCount; i++) {
        slices.push(await new Promise((res, rej) => file.getSliceAsync(i, (r) => (ok(r) ? res(r.value.data) : rej(r.error)))));
      }
    } finally { file.closeAsync(); }
    const bytes = new Uint8Array(slices.reduce((n, s) => n + s.length, 0));
    let o = 0;
    for (const s of slices) { bytes.set(s, o); o += s.length; }
    await window.kvLoadFflate();
    const files = fflate.unzipSync(bytes, { filter: (f) => f.name === "ppt/presentation.xml" });
    const sz = new DOMParser().parseFromString(fflate.strFromU8(files["ppt/presentation.xml"]), "application/xml")
      .getElementsByTagName("p:sldSz")[0];
    slideSize = { w: sz.getAttribute("cx") / 12700, h: sz.getAttribute("cy") / 12700 };
    localStorage.setItem(key, JSON.stringify(slideSize));
  } catch (e) { console.error(e); }
  return slideSize;
}

// Ramme om figurerne; ved én figur er rammen sliden.
async function bounds(shapes) {
  if (shapes.length === 1) { const { w, h } = await getSlideSize(); return { l: 0, t: 0, r: w, b: h }; }
  return {
    l: Math.min(...shapes.map((s) => s.left)), t: Math.min(...shapes.map((s) => s.top)),
    r: Math.max(...shapes.map((s) => s.left + s.width)), b: Math.max(...shapes.map((s) => s.top + s.height)),
  };
}

// ---------- Juster ----------
// Den først markerede figur står stille, og de øvrige justeres efter den. Ved én figur justeres der efter sliden.
const align = (fn) => onShapes(1, async (shapes) => {
  if (shapes.length === 1) { fn(shapes[0], await bounds(shapes)); return; }
  const [ref, ...rest] = shapes;
  const b = { l: ref.left, t: ref.top, r: ref.left + ref.width, b: ref.top + ref.height };
  rest.forEach((s) => fn(s, b));
});
const alignLeft = align((s, b) => { s.left = b.l; });
const alignRight = align((s, b) => { s.left = b.r - s.width; });
const alignCenter = align((s, b) => { s.left = (b.l + b.r) / 2 - s.width / 2; });
const alignTop = align((s, b) => { s.top = b.t; });
const alignBottom = align((s, b) => { s.top = b.b - s.height; });
const alignMiddle = align((s, b) => { s.top = (b.t + b.b) / 2 - s.height / 2; });

// Flytter markeringen samlet, så den står midt på sliden.
const centerOnSlide = onShapes(1, async (shapes) => {
  const { w, h } = await getSlideSize();
  const l = Math.min(...shapes.map((s) => s.left)), t = Math.min(...shapes.map((s) => s.top));
  const r = Math.max(...shapes.map((s) => s.left + s.width)), b = Math.max(...shapes.map((s) => s.top + s.height));
  const dx = (w - (r - l)) / 2 - l, dy = (h - (b - t)) / 2 - t;
  shapes.forEach((s) => { s.left += dx; s.top += dy; });
});

// ---------- Afstand ----------
// Første figur (længst til venstre/øverst) bliver stående; de øvrige flyttes step pr. mellemrum.
function nudge(axis, step) {
  const pos = axis === "h" ? "left" : "top";
  return onShapes(2, (shapes) => {
    [...shapes].sort((a, b) => a[pos] - b[pos]).forEach((s, i) => { s[pos] += i * step; });
  });
}
const spaceHMore = nudge("h", STEP), spaceHLess = nudge("h", -STEP);
const spaceVMore = nudge("v", STEP), spaceVLess = nudge("v", -STEP);

// Lige store mellemrum mellem første og sidste figur.
function distribute(axis) {
  const pos = axis === "h" ? "left" : "top", size = axis === "h" ? "width" : "height";
  return onShapes(3, (shapes) => {
    const s = [...shapes].sort((a, b) => a[pos] - b[pos]);
    const span = s[s.length - 1][pos] + s[s.length - 1][size] - s[0][pos];
    const gap = (span - s.reduce((n, x) => n + x[size], 0)) / (s.length - 1);
    let at = s[0][pos];
    s.forEach((x) => { x[pos] = at; at += x[size] + gap; });
  });
}
const distributeH = distribute("h"), distributeV = distribute("v");

// Gitter: objekterne tages i læserækkefølge (rækker oppefra, så venstre mod højre) og stilles fra øverste venstre hjørne.
// cols = 0: lige så mange kolonner som den længste række, de står i nu.
function gridCmd({ cols = 0, gap = 0.3 * 28.3465, same = false } = {}) {
  return onShapes(2, (shapes) => {
    const avgH = shapes.reduce((n, s) => n + s.height, 0) / shapes.length;
    const rows = [];
    for (const s of [...shapes].sort((a, b) => a.top - b.top)) {
      const r = rows[rows.length - 1];
      if (r && s.top - r[0].top < avgH / 2) r.push(s); else rows.push([s]);
    }
    const order = rows.flatMap((r) => r.sort((a, b) => a.left - b.left));
    const c = cols > 0 ? cols : Math.max(...rows.map((r) => r.length));
    const cw = Math.max(...order.map((s) => s.width)), ch = Math.max(...order.map((s) => s.height));
    const x0 = Math.min(...order.map((s) => s.left)), y0 = Math.min(...order.map((s) => s.top));
    order.forEach((s, i) => {
      const w = same ? cw : s.width, h = same ? ch : s.height;
      if (same) { s.width = cw; s.height = ch; }
      s.left = x0 + (i % c) * (cw + gap) + (cw - w) / 2;
      s.top = y0 + Math.floor(i / c) * (ch + gap) + (ch - h) / 2;
    });
  });
}
const gridAuto = gridCmd();
window.kvadrantGrid = (opts) => gridCmd(opts)();

// ---------- Form ----------
// Den først markerede figur er reference.
const sameSize = onShapes(2, ([ref, ...rest]) => rest.forEach((s) => { s.width = ref.width; s.height = ref.height; }));
const sameWidth = onShapes(2, ([ref, ...rest]) => rest.forEach((s) => { s.width = ref.width; }));
const sameHeight = onShapes(2, ([ref, ...rest]) => rest.forEach((s) => { s.height = ref.height; }));

// Bytter plads på to figurer (deres midtpunkter), uden at ændre størrelsen.
const swapShapes = onShapes(2, ([a, b]) => {
  const ca = [a.left + a.width / 2, a.top + a.height / 2], cb = [b.left + b.width / 2, b.top + b.height / 2];
  a.left = cb[0] - a.width / 2; a.top = cb[1] - a.height / 2;
  b.left = ca[0] - b.width / 2; b.top = ca[1] - b.height / 2;
});

const autosizeText = onShapes(1, (shapes) => shapes.filter((s) => TEXT_TYPES.includes(s.type))
  .forEach((s) => { s.textFrame.autoSizeSetting = "AutoSizeShapeToFitText"; }));

// Skrift (navn, størrelse, farve, fed, kursiv, understreget) fra den først markerede til resten.
const copyType = onShapes(2, async ([ref, ...rest], ctx) => {
  const f = ref.textFrame.textRange.font;
  f.load("name,size,color,bold,italic,underline");
  await ctx.sync();
  rest.filter((s) => TEXT_TYPES.includes(s.type)).forEach((s) => {
    const t = s.textFrame.textRange.font;
    for (const k of ["name", "size", "color", "bold", "italic", "underline"]) if (f[k] !== null && f[k] !== undefined) t[k] = f[k];
  });
});

// Form (de gule håndtag) fra den først markerede til resten, så rundingen har samme faktiske radius uanset størrelse.
// Håndtagene måles i forhold til figurens korteste side, så værdien skaleres med forholdet mellem de korteste sider.
// Office.js kan ikke skifte figurtype, så figurerne skal være af samme type (samme antal håndtag).
const copyShape = onShapes(2, async ([ref, ...rest], ctx) => {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) { window.toast?.("Kræver en nyere PowerPoint"); return; }
  const shapes = [ref, ...rest];
  shapes.forEach((s) => s.adjustments.load("count"));
  await ctx.sync();
  const n = ref.adjustments.count;
  if (!n) { window.toast?.("Den første figur har ingen håndtag at kopiere"); return; }
  const values = Array.from({ length: n }, (_, i) => ref.adjustments.get(i));
  await ctx.sync();
  const side = (s) => Math.min(s.width, s.height) || 1;
  let done = 0;
  for (const s of rest) {
    if (s.adjustments.count !== n) continue;
    const k = side(ref) / side(s);
    values.forEach((v, i) => {
      let x = v.value * k;
      if (v.value >= 0 && v.value <= 0.5) x = Math.min(0.5, Math.max(0, x));   // hjørneradius kan højst være den halve side
      s.adjustments.set(i, x);
    });
    done++;
  }
  window.toast?.(done === rest.length ? "Formen er kopieret" : `Kopieret til ${done} af ${rest.length} (de andre er en anden figurtype)`);
});

// Én tekstboks pr. afsnit, stablet i den oprindelige boks' plads.
const splitTextBox = onShapes(1, async (shapes, ctx) => {
  for (const s of shapes.filter((x) => TEXT_TYPES.includes(x.type))) {
    const range = s.textFrame.textRange;
    range.load("text,font/name,font/size,font/color");
    await ctx.sync();
    const paras = range.text.split(/\r\n|\r|\n|\v/).map((p) => p.trim()).filter(Boolean);
    if (paras.length < 2) continue;
    const h = s.height / paras.length;
    const slide = s.getParentSlide ? s.getParentSlide() : ctx.presentation.getSelectedSlides().getItemAt(0);
    paras.forEach((p, i) => {
      const box = slide.shapes.addTextBox(p, { left: s.left, top: s.top + i * h, width: s.width, height: h });
      box.name = `${s.name} ${i + 1}`;
      if (range.font.name) box.textFrame.textRange.font.name = range.font.name;
      if (range.font.size) box.textFrame.textRange.font.size = range.font.size;
      if (range.font.color) box.textFrame.textRange.font.color = range.font.color;
    });
    s.delete();
  }
});

const zOrder = (where) => onShapes(1, (shapes) => shapes.forEach((s) => s.setZOrder(where)));
const bringToFront = zOrder("BringToFront"), sendToBack = zOrder("SendToBack");

// ---------- Tekst ----------
const step = (size, dir) => {
  if (!size) return dir > 0 ? 12 : 10;
  return dir > 0 ? (FONT_SIZES.find((s) => s > size) || size + 8) : ([...FONT_SIZES].reverse().find((s) => s < size) || Math.max(1, size - 1));
};
const fontBigger = onText((rs) => rs.forEach((r) => { r.font.size = step(r.font.size, 1); }));
const fontSmaller = onText((rs) => rs.forEach((r) => { r.font.size = step(r.font.size, -1); }));
// Slår til, medmindre alt allerede er slået til (som i PowerPoint).
const toggle = (prop, on, off) => onText((rs) => {
  const allOn = rs.every((r) => r.font[prop] === on);
  rs.forEach((r) => { r.font[prop] = allOn ? off : on; });
});
const toggleBold = toggle("bold", true, false);
const toggleItalic = toggle("italic", true, false);
const toggleUnderline = toggle("underline", "Single", "None");
const paragraph = (value) => onText((rs) => rs.forEach((r) => { r.paragraphFormat.horizontalAlignment = value; }));
const textLeft = paragraph("Left"), textCenter = paragraph("Center"), textRight = paragraph("Right"), textJustify = paragraph("Justify");
const toggleBullets = onText((rs) => {
  const allOn = rs.every((r) => r.paragraphFormat.bulletFormat.visible === true);
  rs.forEach((r) => { r.paragraphFormat.bulletFormat.visible = !allOn; });
});
const fontColor = (hex) => onText((rs) => rs.forEach((r) => { r.font.color = "#" + hex; }));

// ---------- Registrering ----------
const COMMANDS = {
  alignLeft, alignRight, alignCenter, alignTop, alignBottom, alignMiddle, centerOnSlide,
  spaceHMore, spaceHLess, spaceVMore, spaceVLess, distributeH, distributeV, gridAuto,
  sameSize, sameWidth, sameHeight, swapShapes, autosizeText, copyType, copyShape, splitTextBox, bringToFront, sendToBack,
  fontBigger, fontSmaller, toggleBold, toggleItalic, toggleUnderline, textLeft, textCenter, textRight, textJustify, toggleBullets,
  fontColor3C2C1C: fontColor("3C2C1C"), fontColorB8862E: fontColor("B8862E"), fontColor000000: fontColor("000000"),
  fontColor7A7A7A: fontColor("7A7A7A"), fontColorFFFFFF: fontColor("FFFFFF"),
  showPane: async (event) => { try { await Office.addin.showAsTaskpane(); } catch (e) { console.error(e); } finally { event?.completed?.(); } },
};

Office.onReady(() => {
// fflate (zip) hentes først, når en funktion skal bruge den, så panel og vinduer starter hurtigere.
let fflatePromise;
window.kvLoadFflate = () => (fflatePromise ||= window.fflate ? Promise.resolve() : new Promise((res, rej) => {
  const s = document.createElement("script");
  s.src = "https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js";
  s.onload = res; s.onerror = () => rej(new Error("Kunne ikke hente zip-biblioteket"));
  document.head.appendChild(s);
}));
  for (const [name, fn] of Object.entries(COMMANDS)) {
    window[name] = fn;
    Office.actions?.associate?.(name, fn);
  }
});
})();
