// Generates ../manifest.template.xml (the "Kvadrant" ribbon tab) and assets/ribbon-icons.json
// (the icons render-ribbon-icons.ps1 turns into PNGs). Edit RIBBON below, then run:
//   node build-ribbon.mjs  &&  powershell -File ..\render-ribbon-icons.ps1
import { writeFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const INK = "404040";
const COLORS = [["3C2C1C", "Mørkebrun"], ["B8862E", "Guld"], ["000000", "Sort"], ["7A7A7A", "Grå"], ["FFFFFF", "Hvid"]];

// Button: [id, label, icon, function, tip]. Menu: { menu: [id, label, icon, tip], items: [buttons] }.
const RIBBON = [
  { id: "Icons", label: "Ikoner", controls: [{ pane: true }] },
  { id: "Text", label: "Tekst", controls: [
    ["FontUp", "Større", "text-aa", "fontBigger", "Gør teksten større"],
    ["FontDown", "Mindre", "text-t", "fontSmaller", "Gør teksten mindre"],
    { menu: ["Style", "Stil", "text-b", "Fed, kursiv og understreget"], items: [
      ["Bold", "Fed", "text-b", "toggleBold", "Fed tekst"],
      ["Italic", "Kursiv", "text-italic", "toggleItalic", "Kursiv tekst"],
      ["Underline", "Understreget", "text-underline", "toggleUnderline", "Understreget tekst"],
    ] },
    { menu: ["Para", "Justering", "text-align-left", "Justér teksten vandret"], items: [
      ["TLeft", "Venstre", "text-align-left", "textLeft", "Venstrestillet tekst"],
      ["TCenter", "Centreret", "text-align-center", "textCenter", "Centreret tekst"],
      ["TRight", "Højre", "text-align-right", "textRight", "Højrestillet tekst"],
      ["TJustify", "Lige margener", "text-align-justify", "textJustify", "Lige margener"],
    ] },
    ["Bullets", "Punkter", "list-bullets", "toggleBullets", "Slå punkttegn til eller fra"],
    { menu: ["Color", "Farve", "palette", "Skift tekstfarve"], items: COLORS.map(([hex, name]) =>
      [`C${hex}`, name, { name: "circle", weight: "fill", color: hex }, `fontColor${hex}`, `Tekstfarve ${name} (#${hex})`]) },
  ] },
  { id: "Align", label: "Juster", controls: [
    ["ALeft", "Venstre", "align-left", "alignLeft", "Flugt de markerede objekter til venstre (ét objekt: til slidens kant)"],
    ["ACenter", "Center", "align-center-horizontal", "alignCenter", "Centrér objekterne vandret"],
    ["ARight", "Højre", "align-right", "alignRight", "Flugt objekterne til højre"],
    ["ATop", "Top", "align-top", "alignTop", "Flugt objekterne i toppen"],
    ["AMiddle", "Midte", "align-center-vertical", "alignMiddle", "Centrér objekterne lodret"],
    ["ABottom", "Bund", "align-bottom", "alignBottom", "Flugt objekterne i bunden"],
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
    ["CopyType", "Kopiér format", "paint-brush", "copyType", "Giv de øvrige markerede objekter skriften fra det først markerede"],
    { menu: ["Order", "Rækkefølge", "stack", "Placér forrest eller bagest"], items: [
      ["Front", "Forrest", "selection-foreground", "bringToFront", "Placér forrest"],
      ["Back", "Bagest", "selection-background", "sendToBack", "Placér bagest"],
    ] },
  ] },
  { id: "Slide", label: "Slide", controls: [
    ["CenterSlide", "Centrér på slide", "crosshair-simple", "centerOnSlide", "Centrér de markerede objekter på sliden"],
  ] },
];

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
const icons = new Map(), shorts = [], longs = [];
const iconRef = (spec) => {
  const s = typeof spec === "string" ? { name: spec, weight: "regular", color: INK } : spec;
  const key = `${s.name}-${s.weight}-${s.color}`.toLowerCase();
  icons.set(key, { file: `rb-${key}`, ...s });
  return key;
};
const iconXml = (key, pad) => ["16", "32", "80"].map((sz) => `${pad}<bt:Image size="${sz}" resid="I.${key}.${sz}"/>`).join("\n");
const strings = (id, label, tip) => { shorts.push([`L.${id}`, label]); longs.push([`T.${id}`, tip]); };

function button([id, label, icon, fn, tip], pad, tag = "Control") {
  strings(id, label, tip);
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
    strings("Open", "Ikonpanel", "Åbn panelet med alle Phosphor-ikoner: søg, vælg tykkelse og farve, og indsæt.");
    return `${pad}<Control xsi:type="Button" id="Phosphor.Open">
${pad}  <Label resid="L.Open"/>
${pad}  <Supertip><Title resid="L.Open"/><Description resid="T.Open"/></Supertip>
${pad}  <Icon>
${pad}    <bt:Image size="16" resid="Icon.16"/>
${pad}    <bt:Image size="32" resid="Icon.32"/>
${pad}    <bt:Image size="80" resid="Icon.80"/>
${pad}  </Icon>
${pad}  <Action xsi:type="ShowTaskpane"><TaskpaneId>PhosphorPane</TaskpaneId><SourceLocation resid="Taskpane.Url"/></Action>
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

const images = [...icons.keys()].flatMap((k) => ["16", "32", "80"].map((s) =>
  `        <bt:Image id="I.${k}.${s}" DefaultValue="{{BASE}}/assets/rb-${k}-${s}.png"/>`)).join("\n");

const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<!-- Generated by web-panel/build-ribbon.mjs; install-panel.ps1 replaces {{BASE}} and writes manifest.xml. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>6f1c7a52-3e0b-4d8e-9a41-2b7c5d9e8f30</Id>
  <Version>1.1.0.0</Version>
  <ProviderName>Kvadrant</ProviderName>
  <DefaultLocale>da-DK</DefaultLocale>
  <DisplayName DefaultValue="Kvadrant"/>
  <Description DefaultValue="Phosphor-ikoner og værktøjer til justering, afstand, form og tekst."/>
  <IconUrl DefaultValue="{{BASE}}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{{BASE}}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://phosphoricons.com"/>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{{BASE}}/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
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
        <bt:Image id="Icon.16" DefaultValue="{{BASE}}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="{{BASE}}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="{{BASE}}/assets/icon-80.png"/>
${images}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="{{BASE}}/index.html"/>
        <bt:Url id="Commands.Url" DefaultValue="{{BASE}}/commands.html"/>
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
</OfficeApp>
`;

writeFileSync(join(here, "..", "manifest.template.xml"), manifest);
writeFileSync(join(here, "assets", "ribbon-icons.json"), JSON.stringify([...icons.values()], null, 1));
console.log(`groups ${RIBBON.length}, icons ${icons.size}, strings ${shorts.length}`);
