// Fælles værktøj til .pptx-filer som zip (bruges af både panelets runtime og tools.html).
// files = { "ppt/slides/slide1.xml": Uint8Array, ... } fra fflate.unzipSync.
// Kræver at fflate er hentet. Alt lægges på window.KvPptx.
(() => {
const dec = (u8) => fflate.strFromU8(u8), enc = (s) => fflate.strToU8(s);
const attr = (xml, name) => new RegExp(`\\s${name}="([^"]*)"`).exec(xml)?.[1];
const relsPath = (part) => part === "" ? "_rels/.rels" : part.replace(/([^/]+)$/, "_rels/$1.rels");
function resolve(base, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const seg of target.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
  return parts.join("/");
}
// Relationerne fra en del: [{ xml, id, type, path }]; path er null for eksterne links.
function rels(files, part) {
  const k = relsPath(part);
  if (!files[k]) return [];
  return [...dec(files[k]).matchAll(/<Relationship\b[^>]*?\/>/g)].map(([xml]) => ({
    xml, id: attr(xml, "Id"), type: attr(xml, "Type") || "",
    path: /TargetMode="External"/.test(xml) ? null : resolve(part, attr(xml, "Target") || ""),
  }));
}
function dropRels(files, part, ids) {
  const k = relsPath(part);
  if (!files[k] || !ids.length) return;
  files[k] = enc(dec(files[k]).replace(/<Relationship\b[^>]*?\/>/g, (m) => (ids.includes(attr(m, "Id")) ? "" : m)));
}
const str = (files, k) => (files[k] ? dec(files[k]) : "");
const put = (files, k, s) => { files[k] = enc(s); };

// Sletter alle dele, der ikke længere kan nås fra roden, og rydder [Content_Types].xml.
function gc(files) {
  const keep = new Set(["[Content_Types].xml", "_rels/.rels"]);
  const queue = [""];
  while (queue.length) {
    const part = queue.pop();
    if (part) { keep.add(part); keep.add(relsPath(part)); }
    for (const r of rels(files, part)) if (r.path && files[r.path] && !keep.has(r.path)) queue.push(r.path);
  }
  for (const k of Object.keys(files)) if (!keep.has(k) && !k.endsWith("/")) delete files[k];
  put(files, "[Content_Types].xml", str(files, "[Content_Types].xml")
    .replace(/<Override PartName="\/([^"]+)"[^>]*\/>/g, (m, p) => (files[p] ? m : "")));
  return files;
}

// Slides i rækkefølge: [{ index, id (sldId), rid, path, layout, master }]
function slideList(files) {
  const pres = str(files, "ppt/presentation.xml");
  const byId = Object.fromEntries(rels(files, "ppt/presentation.xml").map((r) => [r.id, r]));
  return [...pres.matchAll(/<p:sldId\b[^>]*\/>/g)].map(([xml], index) => {
    const rid = attr(xml, "r:id"), path = byId[rid]?.path;
    const layout = path && rels(files, path).find((r) => r.type.endsWith("/slideLayout"))?.path;
    const master = layout && rels(files, layout).find((r) => r.type.endsWith("/slideMaster"))?.path;
    return { index, id: attr(xml, "id"), rid, path, layout, master };
  });
}

// Behold kun slides med de givne numre (0-baseret); resten fjernes helt, inkl. noter, billeder og sektionernes henvisninger.
function keepSlides(files, keep) {
  const keepSet = new Set(keep);
  const gone = slideList(files).filter((s) => !keepSet.has(s.index));
  if (!gone.length) return files;
  const goneRid = new Set(gone.map((s) => s.rid)), goneId = new Set(gone.map((s) => s.id));
  let pres = str(files, "ppt/presentation.xml")
    .replace(/<p:sldId\b[^>]*\/>/g, (m) => (goneRid.has(attr(m, "r:id")) ? "" : m))
    .replace(/<p14:sldId\b[^>]*\/>/g, (m) => (goneId.has(attr(m, "id")) ? "" : m))
    .replace(/<p:custShowLst>[\s\S]*?<\/p:custShowLst>/, "");
  put(files, "ppt/presentation.xml", pres);
  dropRels(files, "ppt/presentation.xml", [...goneRid]);
  for (const s of gone) delete files[s.path];
  return gc(files);
}

// Fjerner masters, som ingen slide bruger (og med layouts: true også ubrugte layouts i de masters, der bliver).
function dropUnused(files, { layouts = false } = {}) {
  const slides = slideList(files);
  const usedLayouts = new Set(slides.map((s) => s.layout).filter(Boolean));
  const usedMasters = new Set(slides.map((s) => s.master).filter(Boolean));
  const presRels = rels(files, "ppt/presentation.xml");
  const masters = presRels.filter((r) => r.type.endsWith("/slideMaster"));
  if (!usedMasters.size && masters[0]) usedMasters.add(masters[0].path);   // en præsentation skal have mindst én master
  const goneMasters = masters.filter((r) => !usedMasters.has(r.path));
  const goneRids = new Set(goneMasters.map((r) => r.id));
  put(files, "ppt/presentation.xml", str(files, "ppt/presentation.xml").replace(/<p:sldMasterId\b[^>]*\/>/g, (m) => (goneRids.has(attr(m, "r:id")) ? "" : m)));
  dropRels(files, "ppt/presentation.xml", [...goneRids]);
  let goneLayouts = 0;
  if (layouts) {
    for (const m of masters.filter((r) => usedMasters.has(r.path))) {
      const unused = rels(files, m.path).filter((r) => r.type.endsWith("/slideLayout") && !usedLayouts.has(r.path)).map((r) => r.id);
      if (!unused.length) continue;
      put(files, m.path, str(files, m.path).replace(/<p:sldLayoutId\b[^>]*\/>/g, (x) => (unused.includes(attr(x, "r:id")) ? "" : x)));
      dropRels(files, m.path, unused);
      goneLayouts += unused.length;
    }
  }
  gc(files);
  return { masters: goneMasters.length, layouts: goneLayouts };
}

// ---------- Masters og temaer ----------
const ELIXIRR_THEMES = /^(cover slides|content & sections|template \/ text|template \/ content)$/i;
function family({ name, minor, dk2 }) {
  if (ELIXIRR_THEMES.test(name) || dk2 === "300165") return "elixirr";
  if (/^inter/i.test(minor) || /kvadrant consulting/i.test(name)) return "kvNew";
  if (/open sans/i.test(minor) || /^kvadrant theme/i.test(name)) return "kvOld";
  return "other";
}
const FAMILY_NAMES = { elixirr: "Elixirr", kvNew: "Kvadrant · ny CVI (Inter)", kvOld: "Kvadrant · gammel CVI (Open Sans)", other: "Andet" };

function masterInfo(files) {
  const slides = slideList(files);
  const masters = rels(files, "ppt/presentation.xml").filter((r) => r.type.endsWith("/slideMaster")).map((r) => {
    const theme = rels(files, r.path).find((x) => x.type.endsWith("/theme"))?.path;
    const t = str(files, theme);
    const font = (tag) => new RegExp(`<a:${tag}>\\s*<a:latin typeface="([^"]*)"`).exec(t)?.[1] || "";
    const info = { path: r.path, name: unesc(attr(t, "name") || ""), major: font("majorFont"), minor: font("minorFont"),
      dk2: /<a:dk2>\s*<a:(?:srgbClr val|sysClr[^>]*lastClr)="([0-9A-Fa-f]{6})"/.exec(t)?.[1]?.toUpperCase() || "",
      layouts: rels(files, r.path).filter((x) => x.type.endsWith("/slideLayout")).length,
      slides: slides.filter((s) => s.master === r.path).map((s) => s.index) };
    return { ...info, family: family(info), familyName: FAMILY_NAMES[family(info)] };
  });
  return { masters, slides: slides.map((s) => ({ index: s.index, master: s.master, family: masters.find((m) => m.path === s.master)?.family || "other" })) };
}

// Sektioner: [{ name, slides: [index, ...] }]
function sections(files) {
  const pres = str(files, "ppt/presentation.xml");
  const indexById = Object.fromEntries(slideList(files).map((s) => [s.id, s.index]));
  return [...pres.matchAll(/<p14:section\b([^>]*)>([\s\S]*?)<\/p14:section>/g)].map(([, a, body]) => ({
    name: (attr(a, "name") || "").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&apos;/g, "'"),
    slides: [...body.matchAll(/<p14:sldId id="(\d+)"/g)].map(([, id]) => indexById[id]).filter((i) => i != null),
  }));
}

// Tekst pr. slide (afsnit adskilt af linjeskift).
function slideText(files, path) {
  return [...str(files, path).matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
    .map(([, p]) => [...p.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(([, t]) => t).join(""))
    .map((t) => t.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/\u200b/g, "").replace(/\s+/g, " ").trim())
    .filter(Boolean).join("\n");
}

// ---------- Billeder ----------
// Pixelstørrelse fra filens hoved (PNG og JPEG); null for andre formater.
function imageSize(u8) {
  if (u8[0] === 0x89 && u8[1] === 0x50) return { w: (u8[16] << 24 | u8[17] << 16 | u8[18] << 8 | u8[19]) >>> 0, h: (u8[20] << 24 | u8[21] << 16 | u8[22] << 8 | u8[23]) >>> 0, type: "png" };
  if (u8[0] === 0xff && u8[1] === 0xd8) {
    let i = 2;
    while (i + 9 < u8.length) {
      if (u8[i] !== 0xff) { i++; continue; }
      const m = u8[i + 1], len = u8[i + 2] << 8 | u8[i + 3];
      if (m >= 0xc0 && m <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(m)) return { h: u8[i + 5] << 8 | u8[i + 6], w: u8[i + 7] << 8 | u8[i + 8], type: "jpeg" };
      i += 2 + len;
    }
  }
  return null;
}
// Hvert billede og hvor stort det højst vises (i tommer, korrigeret for beskæring), så det kan skaleres til fx 150 dpi.
function imageUsage(files) {
  const use = {};
  const parts = Object.keys(files).filter((k) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(k));
  for (const part of parts) {
    const byId = Object.fromEntries(rels(files, part).map((r) => [r.id, r.path]));
    for (const [block] of str(files, part).matchAll(/<p:(pic|sp)\b[\s\S]*?<\/p:\1>/g)) {
      const embed = /<a:blip\b[^>]*r:embed="([^"]+)"/.exec(block)?.[1];
      const media = embed && byId[embed];
      if (!media || !files[media]) continue;
      const ext = /<a:ext cx="(\d+)" cy="(\d+)"/.exec(block);
      if (!ext) continue;
      const src = /<a:srcRect\b([^>]*)\/>/.exec(block)?.[1] || "";
      const pct = (n) => (+attr(src, n) || 0) / 100000;
      const visW = Math.max(0.05, 1 - pct("l") - pct("r")), visH = Math.max(0.05, 1 - pct("t") - pct("b"));
      const wIn = +ext[1] / 914400 / visW, hIn = +ext[2] / 914400 / visH;
      const u = (use[media] ||= { media, wIn: 0, hIn: 0, uses: 0 });
      u.wIn = Math.max(u.wIn, wIn); u.hIn = Math.max(u.hIn, hIn); u.uses++;
    }
  }
  return Object.values(use).map((u) => ({ ...u, bytes: files[u.media].length, size: imageSize(files[u.media]) }));
}

// Store billeder, der kan skaleres ned til dpi uden at blive vist mindre: [{ media, bytes, w, h, dpi, tw, th }]
function oversized(files, dpi = 150) {
  return imageUsage(files).filter((u) => u.size && u.bytes > 60000).map((u) => {
    const cur = Math.min(u.size.w / u.wIn, u.size.h / u.hIn);
    const f = dpi / cur;
    return { ...u, w: u.size.w, h: u.size.h, dpi: Math.round(cur), tw: Math.round(u.size.w * f), th: Math.round(u.size.h * f) };
  }).filter((u) => u.dpi > dpi * 1.15).sort((a, b) => b.bytes - a.bytes);
}

// ---------- Layouts og masters ----------
// Relativ sti fra én del til en anden, fx ppt/slides/slide1.xml → ../slideLayouts/slideLayout3.xml
function relTo(from, to) {
  const a = from.split("/").slice(0, -1), b = to.split("/");
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...Array(a.length - i).fill(".."), ...b.slice(i)].join("/");
}
function retarget(files, part, rel, toPath) {
  const k = relsPath(part);
  put(files, k, str(files, k).replace(rel.xml, rel.xml.replace(/Target="[^"]*"/, `Target="${relTo(part, toPath)}"`)));
}
const unesc = (s) => s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
function layoutsOf(files, master) {
  return rels(files, master).filter((r) => r.type.endsWith("/slideLayout") && files[r.path]).map((r) => {
    const x = str(files, r.path);
    return { path: r.path, master, name: unesc(/<p:cSld\b[^>]*\bname="([^"]*)"/.exec(x)?.[1] || ""), type: /<p:sldLayout\b[^>]*\btype="(\w+)"/.exec(x)?.[1] || "" };
  });
}
const layoutRel = (files, slide) => rels(files, slide).find((r) => r.type.endsWith("/slideLayout"));

// En slide uden noter og kommentarer på masterens tomme layout (bruges til figurer).
function blankSlide(files, slide) {
  dropRels(files, slide, rels(files, slide).filter((r) => /\/(notesSlide|comments|comment)$/.test(r.type) || /\/comments$/.test(r.type)).map((r) => r.id));
  const rel = layoutRel(files, slide);
  const master = rel && rels(files, rel.path).find((r) => r.type.endsWith("/slideMaster"))?.path;
  if (master) {
    const all = layoutsOf(files, master);
    const blank = all.filter((l) => l.type === "blank" || /blank|empty|\btom\b/i.test(l.name));
    const pick = blank.find((l) => /white|hvid/i.test(l.name)) || blank[0];
    if (pick && pick.path !== rel.path) retarget(files, slide, rel, pick.path);
  }
  dropUnused(files, { layouts: true });
  return files;
}

// Kopierer dele (fx en master med layouts, tema og billeder) fra en anden pakke ind med nye, ledige navne.
function copyParts(dst, src, roots) {
  const map = {}, taken = new Set(Object.keys(dst));
  const srcCt = str(src, "[Content_Types].xml");
  let ct = str(dst, "[Content_Types].xml");
  const fresh = (p) => {
    const m = /^(.*?)(\d*)(\.[^./]+)$/.exec(p);
    let n = 1, out;
    do out = `${m[1]}${n++}${m[3]}`; while (taken.has(out));
    taken.add(out);
    return out;
  };
  const visit = (p) => {
    if (map[p]) return map[p];
    const np = (map[p] = fresh(p));
    dst[np] = src[p];
    const rs = rels(src, p);
    if (rs.length) {
      let x = str(src, relsPath(p));
      for (const r of rs) {
        if (!r.path || !src[r.path] || /\/(slide|notesSlide)$/.test(r.type)) continue;
        const t = visit(r.path);
        x = x.replace(r.xml, r.xml.replace(/Target="[^"]*"/, `Target="${relTo(np, t)}"`));
      }
      put(dst, relsPath(np), x);
    }
    const ov = new RegExp(`<Override PartName="/${p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}" ContentType="([^"]+)"`).exec(srcCt);
    if (ov) ct = ct.replace("</Types>", `<Override PartName="/${np}" ContentType="${ov[1]}"/></Types>`);
    else {
      const ext = np.split(".").pop().toLowerCase();
      const def = new RegExp(`<Default Extension="${ext}" ContentType="([^"]+)"`, "i").exec(srcCt);
      if (def && !new RegExp(`<Default Extension="${ext}"`, "i").test(ct)) ct = ct.replace("<Default ", `<Default Extension="${ext}" ContentType="${def[1]}"/><Default `);
    }
    return np;
  };
  roots.forEach(visit);
  put(dst, "[Content_Types].xml", ct);
  return map;
}
// Henter masters fra et template (src) ind i præsentationen og giver dem og deres layouts nye, unikke id'er. Returnerer de nye stier.
function importMasters(files, src, masterPaths) {
  const map = copyParts(files, src, masterPaths);
  const PRES = "ppt/presentation.xml";
  let pres = str(files, PRES);
  const masters = rels(files, PRES).filter((r) => r.type.endsWith("/slideMaster"));
  const ids = [...pres.matchAll(/<p:sldMasterId\b[^>]*\bid="(\d+)"/g)].map((m) => +m[1]);
  for (const m of masters) if (files[m.path]) ids.push(...[...str(files, m.path).matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g)].map((x) => +x[1]));
  let next = Math.max(2147483647, ...ids) + 1;
  let rid = Math.max(0, ...rels(files, PRES).map((r) => +(/\d+$/.exec(r.id)?.[0] || 0))) + 1;
  const added = [];
  let relXml = str(files, relsPath(PRES)), list = "";
  for (const p of masterPaths) {
    const np = map[p];
    added.push(np);
    const mid = next++;
    put(files, np, str(files, np).replace(/(<p:sldLayoutId\b[^>]*\bid=")\d+/g, (_, a) => a + next++));
    relXml = relXml.replace("</Relationships>", `<Relationship Id="rId${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="${relTo(PRES, np)}"/></Relationships>`);
    list += `<p:sldMasterId id="${mid}" r:id="rId${rid++}"/>`;
  }
  put(files, relsPath(PRES), relXml);
  put(files, PRES, pres.replace("</p:sldMasterIdLst>", list + "</p:sldMasterIdLst>"));
  return added;
}
// Fjerner de givne masters (som ingen slide længere bruger) helt fra præsentationen.
function removeMasters(files, paths) {
  const PRES = "ppt/presentation.xml";
  const gone = rels(files, PRES).filter((r) => r.type.endsWith("/slideMaster") && paths.includes(r.path)).map((r) => r.id);
  put(files, PRES, str(files, PRES).replace(/<p:sldMasterId\b[^>]*\/>/g, (m) => (gone.includes(attr(m, "r:id")) ? "" : m)));
  dropRels(files, PRES, gone);
  gc(files);
  return gone.length;
}

// Hvilken slags layout det er, ud fra navn og type (Kvadrant og Elixirr bruger forskellige navne).
const KINDS = [["cover", /cover|front ?page|title slide|forside/i], ["divider", /breaker|divider|section header|afsnit/i],
  ["agenda", /agenda|table of contents|indhold/i], ["end", /(^|[\s/_])end\b|slut/i], ["blank", /empty|blank|\btom\b/i]];
const layoutKind = (l) => (l.type === "blank" ? "blank" : l.type === "title" ? "cover" : l.type === "secHead" ? "divider" : KINDS.find(([, rx]) => rx.test(l.name))?.[0] || "content");
const ELIXIRR_LAYOUT = { cover: /^core \/ cover \/ 1$/i, divider: /^core \/ divider \/ 1$/i, agenda: /^core \/ agenda \/ white$/i,
  end: /^core \/ end \/ light$/i, blank: /^block \/ blank \/ white$/i, content: /^text \/ 1$/i };
function pickLayout(targets, kind) {
  return targets.find((l) => ELIXIRR_LAYOUT[kind].test(l.name)) || targets.find((l) => layoutKind(l) === kind)
    || targets.find((l) => ELIXIRR_LAYOUT.content.test(l.name)) || targets[0];
}

// Pladsholdere arver placering og tekstformat fra layoutet. Før en slide skifter layout, skrives det, den arver, ind i sliden,
// så indholdet bliver, hvor det er. Tomme pladsholdere, som det nye layout ikke har, fjernes.
const phOf = (block) => { const a = /<p:ph\b([^>]*)\/?>/.exec(block)?.[1]; return a == null ? null : { type: attr(a, "type") || "body", idx: attr(a, "idx") }; };
const TITLE_PH = /^(title|ctrTitle)$/;
// Pladsholderen i et layout, en slide-pladsholder arver fra: samme idx, ellers samme type (titler matcher titler).
function phMatch(xml, ph) {
  const blocks = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((m) => m[0]).filter((b) => phOf(b));
  if (ph.idx != null) { const hit = blocks.find((b) => phOf(b).idx === ph.idx); if (hit) return hit; }
  return blocks.find((b) => { const p = phOf(b); return p.type === ph.type || (TITLE_PH.test(p.type) && TITLE_PH.test(ph.type)); });
}
// Masteren har kun titel, brødtekst, dato, sidefod og sidenummer.
const masterPh = (ph) => ({ type: TITLE_PH.test(ph.type) ? "title" : ["dt", "ftr", "sldNum"].includes(ph.type) ? ph.type : "body" });
const xfrmIn = (x) => /<a:xfrm\b[\s\S]*?<\/a:xfrm>/.exec(x || "")?.[0];
function freezeAndMove(files, slide, newLayout) {
  const rel = layoutRel(files, slide);
  const oldLayout = rel.path, master = rels(files, oldLayout).find((r) => r.type.endsWith("/slideMaster"))?.path;
  const lx = str(files, oldLayout), mx = str(files, master), nx = str(files, newLayout);
  const xml = str(files, slide).replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (block) => {
    const ph = phOf(block);
    if (!ph) return block;
    const inLayout = phMatch(lx, ph), inMaster = phMatch(mx, masterPh(ph));
    const target = phMatch(nx, ph);
    const fits = !!target && (ph.idx == null ? true : phOf(target).idx === ph.idx || TITLE_PH.test(ph.type));
    const text = [...block.matchAll(/<a:t>([^<]*)<\/a:t>/g)].some((m) => m[1].trim());
    if (!fits && !text && !/<a:blip\b/.test(block)) return "";
    let b = block;
    const spPr = /<p:spPr\b[^>]*\/>|<p:spPr\b[^>]*>[\s\S]*?<\/p:spPr>/.exec(b)?.[0] || "";
    const xf = !/<a:xfrm\b/.test(spPr) && (xfrmIn(inLayout) || xfrmIn(inMaster));
    if (xf) b = /<p:spPr\/>/.test(b) ? b.replace("<p:spPr/>", `<p:spPr>${xf}</p:spPr>`) : b.replace(/<p:spPr\b([^>]*)>/, `<p:spPr$1>${xf}`);
    const from = inLayout;
    if (from) {
      const ls = /<a:lstStyle>[\s\S]*?<\/a:lstStyle>/.exec(from)?.[0];
      if (ls && /<a:lstStyle\/>/.test(b)) b = b.replace("<a:lstStyle/>", ls);
      const bp = /<a:bodyPr\b[^>]*?(\/>|>[\s\S]*?<\/a:bodyPr>)/.exec(from)?.[0];
      if (bp && /<a:bodyPr\/>/.test(b)) b = b.replace("<a:bodyPr/>", bp);
    }
    return b;
  });
  put(files, slide, xml);
  retarget(files, slide, rel, newLayout);
}

// Flytter slides på de masters, der skal fjernes, til Elixirr, og fjerner masterne. template: pakken fra Elixirr-templatet,
// bruges kun, hvis præsentationen ikke selv har en Elixirr-master.
function moveToElixirr(files, { remove = [], template = null } = {}) {
  const info = masterInfo(files);
  const gone = new Set(remove);
  const moving = slideList(files).filter((s) => gone.has(s.master));
  let targets = info.masters.filter((m) => m.family === "elixirr" && !gone.has(m.path)).map((m) => m.path);
  let imported = 0;
  if (moving.length && !targets.length) {
    if (!template) throw new Error("Præsentationen har ingen Elixirr-master, og Elixirr-templatet blev ikke fundet");
    const tm = masterInfo(template).masters.filter((m) => m.family === "elixirr").map((m) => m.path);
    if (!tm.length) throw new Error("Templatet har ingen Elixirr-master");
    targets = importMasters(files, template, tm);
    imported = targets.length;
  }
  const layouts = targets.flatMap((m) => layoutsOf(files, m));
  const moved = [];
  for (const s of moving) {
    const from = layoutsOf(files, s.master).find((l) => l.path === s.layout) || { name: "", type: "" };
    const to = pickLayout(layouts, layoutKind(from));
    freezeAndMove(files, s.path, to.path);
    moved.push({ slide: s.index + 1, from: from.name, to: to.name });
  }
  const removed = removeMasters(files, [...gone]);
  return { moved, removed, imported };
}

const zip = (files) => fflate.zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) =>
  [k, /\.(jpe?g|png|gif|mp4|m4a|wdp)$/i.test(k) ? [v, { level: 0 }] : [v, { level: 6 }]])));

window.KvPptx = { rels, gc, slideList, keepSlides, dropUnused, masterInfo, sections, slideText, imageSize, imageUsage, oversized, zip, FAMILY_NAMES, family,
  layoutsOf, blankSlide, importMasters, removeMasters, moveToElixirr, layoutKind };
})();
