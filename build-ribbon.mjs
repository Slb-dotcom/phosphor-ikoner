// Generates ../manifest.template.xml (the "Kvadrant" ribbon tab) and assets/ribbon-icons.json
// (the icons render-ribbon-icons.ps1 turns into PNGs). Edit RIBBON below, then run:
//   node build-ribbon.mjs  &&  powershell -File ..\render-ribbon-icons.ps1
import { writeFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const INK = "404040";

// Button: [id, label, icon, function, tip]. Menu: { menu: [id, label, icon, tip], items: [buttons] }.
const RIBBON = [
  { id: "Icons", label: "Ikoner", controls: [{ pane: true }] },
  // PowerPoints egne grupper Skrifttype og Afsnit (samme som på fanen Hjem).
  { office: "GroupFont" },
  { office: "GroupParagraph" },
  { id: "Align", label: "Juster", controls: [
    ["ALeft", "Venstre", "align-left", "alignLeft", "Flugt til venstre efter det først markerede objekt (ét objekt: til slidens kant)"],
    ["ACenter", "Center", "align-center-horizontal", "alignCenter", "Centrér vandret efter det først markerede objekt"],
    ["ARight", "Højre", "align-right", "alignRight", "Flugt til højre efter det først markerede objekt"],
    ["ATop", "Top", "align-top", "alignTop", "Flugt i toppen efter det først markerede objekt"],
    ["AMiddle", "Midte", "align-center-vertical", "alignMiddle", "Centrér lodret efter det først markerede objekt"],
    ["ABottom", "Bund", "align-bottom", "alignBottom", "Flugt i bunden efter det først markerede objekt"],
  ] },
  { id: "Space", label: "Afstand", controls: [
    ["HMore", "Øg vandret", "arrows-out-line-horizontal", "spaceHMore", "Øg den vandrette afstand med 0,1 cm"],
    ["HLess", "Mindsk vandret", "arrows-in-line-horizontal", "spaceHLess", "Mindsk den vandrette afstand med 0,1 cm"],
    ["HDist", "Fordel vandret", "columns", "distributeH", "Fordel objekterne jævnt vandret"],
    ["VMore", "Øg lodret", "arrows-out-line-vertical", "spaceVMore", "Øg den lodrette afstand med 0,1 cm"],
    ["VLess", "Mindsk lodret", "arrows-in-line-vertical", "spaceVLess", "Mindsk den lodrette afstand med 0,1 cm"],
    ["VDist", "Fordel lodret", "rows", "distributeV", "Fordel objekterne jævnt lodret"],
  ] },
  { id: "Shape", label: "Form", controls: [
    { menu: ["Size", "Samme størrelse", "resize", "Gør objekterne lige store som det først markerede"], items: [
      ["SameSize", "Samme størrelse", "resize", "sameSize", "Samme bredde og højde som det først markerede"],
      ["SameW", "Samme bredde", "arrows-horizontal", "sameWidth", "Samme bredde som det først markerede"],
      ["SameH", "Samme højde", "arrows-vertical", "sameHeight", "Samme højde som det først markerede"],
    ] },
    ["Swap", "Byt plads", "swap", "swapShapes", "Byt plads på to markerede objekter"],
    ["Split", "Del tekstboks", "split-vertical", "splitTextBox", "Del en tekstboks i én boks pr. afsnit"],
    ["Autosize", "Tilpas tekst", "arrows-in-simple", "autosizeText", "Tilpas formen til teksten"],
    { menu: ["Copy", "Kopiér", "paint-brush", "Kopiér skrift eller form fra det først markerede objekt"], items: [
      ["CopyType", "Kopiér skrift", "paint-brush", "copyType", "Giv de øvrige markerede objekter skriften fra det først markerede"],
      ["CopyShape", "Kopiér form", "frame-corners", "copyShape", "Giv de øvrige markerede figurer samme runding og håndtag som den først markerede, med samme faktiske radius uanset størrelse"],
    ] },
    { menu: ["Order", "Rækkefølge", "stack", "Placér forrest eller bagest"], items: [
      ["Front", "Forrest", "selection-foreground", "bringToFront", "Placér forrest"],
      ["Back", "Bagest", "selection-background", "sendToBack", "Placér bagest"],
    ] },
  ] },
  { id: "Insert", label: "Indsæt", controls: [
    ["Agenda", "Agenda", "list-numbers", "openAgenda", "Lav agenda-slides ud fra afsnits-slides, med det aktuelle afsnit fremhævet"],
    ["Grid", "Gitter", "grid-four", "openGrid", "Stil de markerede objekter i rækker og kolonner med samme afstand"],
    { menu: ["Stamp", "Stempel", "seal", "Udkast, fortroligt og kildehenvisning"], items: [
      ["Draft", "Udkast", "seal-warning", "stampDraft", "Sæt et UDKAST-stempel på sliden"],
      ["Conf", "Fortroligt", "lock-simple", "stampConfidential", "Sæt et FORTROLIGT-stempel på sliden"],
      ["Disc", "Til diskussion", "seal-question", "stampDiscussion", "Sæt et TIL DISKUSSION-stempel på sliden"],
      ["Source", "Kilde", "quotes", "stampSource", "Indsæt en kildehenvisning nederst"],
      ["SourceEn", "Source", "quotes", "stampSourceEn", "Insert a source line at the bottom"],
    ] },
    { menu: ["Boxes", "Kvadrant-bokse", "frame-corners", "Underoverskrift, baggrundsboks og COMMENTARY i Kvadrants farver"], items: [
      ["BoxSub", "Underoverskrift", "rows", "boxSubheader", "Bjælke i Tekst 2 ved 90 % gennemsigtighed"],
      ["BoxBg", "Baggrundsboks", "frame-corners", "boxBackground", "Boks i Baggrund 2 ved 75 % gennemsigtighed"],
      ["BoxComHead", "COMMENTARY-overskrift", "quotes", "boxCommentaryHead", "COMMENTARY i Tekst 2 med hvide versaler"],
      ["BoxComBody", "COMMENTARY-tekst", "quotes", "boxCommentaryBody", "Kommentarboks i Baggrund 2, topjusteret med punkter"],
    ] },
    { menu: ["Harvey", "Harvey ball", "circle-half", "Indsæt en Harvey ball"], items: [
      ["Hv25", "25 %", "circle-half", "harvey25", "Harvey ball fyldt 25 %"],
      ["Hv50", "50 %", "circle-half", "harvey50", "Harvey ball fyldt 50 %"],
      ["Hv75", "75 %", "circle-half", "harvey75", "Harvey ball fyldt 75 %"],
      ["Hv100", "100 %", "circle-half", "harvey100", "Harvey ball fyldt helt"],
      ["HvOther", "Andet …", "circle-half", "openHarvey", "Skriv selv procenten i et lille felt"],
    ] },
    ["Library", "Slidebank", "books", "openLibrary", "Indsæt gemte Kvadrant-slides fra slidebanken"],
    ["Check", "Tjek", "seal-check", "openCheck", "Tjek skrifttyper, farver, tekst uden for sliden og tomme pladsholdere"],
  ] },
  { id: "Slide", label: "Slide", controls: [
    ["CenterSlide", "Centrér på slide", "crosshair-simple", "centerOnSlide", "Centrér de markerede objekter på sliden"],
    ["Keys", "Genveje", "keyboard", "showBar", "Åbn Kvadrant-vinduet med alle genveje og værktøjer"],
  ] },
];

// Genvejstaster (shortcuts.json): Ctrl+Skift+bogstav, samme bogstav som i genvejs-baren (keys.html).
// Bogstaver, PowerPoint selv bruger med Ctrl+Skift (C, V, G, H, F, P, Z og pilene), bruger Ctrl+Alt+Skift i stedet.
const MOD = "Ctrl+Shift+", ALT = "Ctrl+Alt+Shift+";
const SHORTCUTS = {
  showPane: ["I", "Åbn ikonpanelet"], showBar: ["K", "Åbn Kvadrant-vinduet med genveje"],
  alignLeft: ["L"], alignRight: ["R"], alignCenter: ["E"], alignTop: ["T"], alignBottom: ["B"], alignMiddle: ["M"],
  centerOnSlide: ["O"], distributeH: ["J"], distributeV: ["U"], gridAuto: ["N", "Stil de markerede objekter i et gitter"],
  sameSize: ["S"], sameWidth: ["W"], sameHeight: ["Y"], swapShapes: ["X"], copyType: ["Q"], copyShape: ["D"],
  spaceHMore: ["Right", , ALT], spaceHLess: ["Left", , ALT], spaceVMore: ["Down", , ALT], spaceVLess: ["Up", , ALT],
  bringToFront: ["F", , ALT], sendToBack: ["G", , ALT], openCheck: ["K", , ALT],
};
const KEY_NAMES = { Right: "→", Left: "←", Down: "↓", Up: "↑" };
const keyLabel = (fn) => SHORTCUTS[fn] && `${SHORTCUTS[fn][2] === ALT ? "Ctrl+Alt+Skift" : "Ctrl+Skift"}+${KEY_NAMES[SHORTCUTS[fn][0]] || SHORTCUTS[fn][0]}`;
const actionNames = {};

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
const icons = new Map(), shorts = [], longs = [];
// Office afviser manifestet, hvis et resid er over 32 tegn, så billederne får korte id'er (I0, I1, ...).
const iconRef = (spec) => {
  const s = typeof spec === "string" ? { name: spec, weight: "regular", color: INK } : spec;
  const key = `${s.name}-${s.weight}-${s.color}`.toLowerCase();
  if (!icons.has(key)) icons.set(key, { file: `rb-${key}`, id: `I${icons.size}`, ...s });
  return key;
};
const iconXml = (key, pad) => ["16", "32", "80"].map((sz) => `${pad}<bt:Image size="${sz}" resid="${icons.get(key).id}.${sz}"/>`).join("\n");
const strings = (id, label, tip) => { shorts.push([`L.${id}`, label]); longs.push([`T.${id}`, tip]); };

function button([id, label, icon, fn, tip], pad, tag = "Control") {
  actionNames[fn] = tip;
  strings(id, label, keyLabel(fn) ? `${tip} (${keyLabel(fn)})` : tip);
  const key = iconRef(icon);
  const open = tag === "Control" ? `<Control xsi:type="Button" id="K.${id}">` : `<Item id="K.${id}">`;
  return `${pad}${open}
${pad}  <Label resid="L.${id}"/>
${pad}  <Supertip><Title resid="L.${id}"/><Description resid="T.${id}"/></Supertip>
${pad}  <Icon>
${iconXml(key, pad + "    ")}
${pad}  </Icon>
${pad}  <Action xsi:type="ExecuteFunction"><FunctionName>${fn}</FunctionName></Action>
${pad}</${tag}>`;
}

function control(c, pad) {
  if (c.pane) {
    strings("Open", "Ikonpanel", `Åbn panelet med alle Phosphor-ikoner: søg, vælg tykkelse og farve, og indsæt. (${keyLabel("showPane")})`);
    return `${pad}<Control xsi:type="Button" id="Phosphor.Open">
${pad}  <Label resid="L.Open"/>
${pad}  <Supertip><Title resid="L.Open"/><Description resid="T.Open"/></Supertip>
${pad}  <Icon>
${pad}    <bt:Image size="16" resid="Icon.16"/>
${pad}    <bt:Image size="32" resid="Icon.32"/>
${pad}    <bt:Image size="80" resid="Icon.80"/>
${pad}  </Icon>
${pad}  <Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/></Action>
${pad}</Control>`;
  }
  if (c.menu) {
    const [id, label, icon, tip] = c.menu;
    strings(id, label, tip);
    const key = iconRef(icon);
    return `${pad}<Control xsi:type="Menu" id="K.${id}">
${pad}  <Label resid="L.${id}"/>
${pad}  <Supertip><Title resid="L.${id}"/><Description resid="T.${id}"/></Supertip>
${pad}  <Icon>
${iconXml(key, pad + "    ")}
${pad}  </Icon>
${pad}  <Items>
${c.items.map((b) => button(b, pad + "    ", "Item")).join("\n")}
${pad}  </Items>
${pad}</Control>`;
  }
  return button(c, pad);
}

const groups = RIBBON.map((g) => {
  if (g.office) return `              <OfficeGroup id="${g.office}"/>`;
  shorts.push([`G.${g.id}`, g.label]);
  const first = g.controls[0];
  const gIcon = first.pane ? null : iconRef(first.menu ? first.menu[2] : first[2]);
  const gIconXml = gIcon ? iconXml(gIcon, "                  ")
    : ["16", "32", "80"].map((s) => `                  <bt:Image size="${s}" resid="Icon.${s}"/>`).join("\n");
  return `              <Group id="K.G.${g.id}">
                <Label resid="G.${g.id}"/>
                <Icon>
${gIconXml}
                </Icon>
${g.controls.map((c) => control(c, "                ")).join("\n")}
              </Group>`;
}).join("\n");

const images = [...icons.values()].flatMap((ic) => ["16", "32", "80"].map((s) =>
  `        <bt:Image id="${ic.id}.${s}" DefaultValue="{{BASE}}/assets/${ic.file}-${s}.png"/>`)).join("\n");

const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<!-- Generated by web-panel/build-ribbon.mjs; install-panel.ps1 replaces {{BASE}} and writes manifest.xml. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>6f1c7a52-3e0b-4d8e-9a41-2b7c5d9e8f30</Id>
  <Version>1.7.0.0</Version>
  <ProviderName>Kvadrant</ProviderName>
  <DefaultLocale>da-DK</DefaultLocale>
  <DisplayName DefaultValue="Kvadrant"/>
  <Description DefaultValue="Phosphor-ikoner og værktøjer til justering, afstand, form og tekst."/>
  <IconUrl DefaultValue="{{BASE}}/assets/kv-32.png"/>
  <HighResolutionIconUrl DefaultValue="{{BASE}}/assets/kv-64.png"/>
  <SupportUrl DefaultValue="https://phosphoricons.com"/>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{{BASE}}/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <!-- Delt runtime: panel og knapper kører i samme side, der bliver indlæst én gang og så bliver liggende. -->
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long"/>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Phosphor.Tab">
${groups}
              <Label resid="Tab.Label"/>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="{{BASE}}/assets/kv-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="{{BASE}}/assets/kv-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="{{BASE}}/assets/kv-80.png"/>
${images}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="{{BASE}}/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Tab.Label" DefaultValue="Kvadrant"/>
${shorts.map(([id, v]) => `        <bt:String id="${id}" DefaultValue="${esc(v)}"/>`).join("\n")}
      </bt:ShortStrings>
      <bt:LongStrings>
${longs.map(([id, v]) => `        <bt:String id="${id}" DefaultValue="${esc(v)}"/>`).join("\n")}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
  <ExtendedOverrides Url="{{BASE}}/shortcuts.json"/>
</OfficeApp>
`;

// Navne i shortcuts.json må kun bestå af bogstaver, tal, - _ og + (ellers afviser Office hele tilføjelsesprogrammet),
// så de rigtige danske tekster ligger under "resources" og refereres som ${resource.<id>}.
const shortcuts = {
  actions: Object.keys(SHORTCUTS).map((id) => ({ id, type: "ExecuteFunction", name: `\${resource.${id}}` })),
  shortcuts: Object.entries(SHORTCUTS).map(([action, [key, , mod]]) => ({ action, key: { default: (mod || MOD) + key } })),
  resources: { default: Object.fromEntries(Object.keys(SHORTCUTS).map((id) =>
    [id, { value: (SHORTCUTS[id][1] || actionNames[id] || id).slice(0, 250) }])) },
};
writeFileSync(join(here, "shortcuts.json"), JSON.stringify(shortcuts, null, 1));
writeFileSync(join(here, "..", "manifest.template.xml"), manifest);
writeFileSync(join(here, "assets", "ribbon-icons.json"), JSON.stringify([...icons.values()], null, 1));
console.log(`groups ${RIBBON.length}, icons ${icons.size}, strings ${shorts.length}`);
