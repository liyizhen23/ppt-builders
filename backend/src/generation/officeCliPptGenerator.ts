import { mkdtemp, readFile, rm } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import type { EvidenceBlock, EvidenceIndex } from "../reports/reportParser.js";
import type { TemplateProfile, TemplateSlideRole } from "../templates/templateProfile.js";
import { runOfficeCli } from "../rendering/officeCliRuntime.js";

export const officeCliGenerationRules = [
  "Only use report evidence and explicit user instructions; do not invent claims.",
  "Use officeCLI as the native deck builder for overall PPT generation.",
  "Build slides from blank pages with explicit officeCLI shapes, text, notes, positions, fonts, and colors.",
  "Use the uploaded template only as visual inspiration for colors and fonts; do not clone template slides.",
  "Keep local partial-edit features on their existing editing path."
] as const;

export interface OfficeCliPptGenerationInput {
  instruction: string;
  reportFileName: string;
  templateFileName: string;
  evidenceIndex: EvidenceIndex;
  templateProfile: TemplateProfile;
  templateSourcePath?: string;
  templateBuffer?: Buffer;
}

export interface OfficeCliPptGenerationResult {
  deckId: string;
  pptxBase64: string;
  summary: string;
  qa: string;
  deckSpec: OfficeCliDeckSpec;
  templateReplacement: {
    selectedSlideIndex: number;
    selectedRole: string;
    replacedSlots: Array<{
      shapeId: string;
      slotType: string;
    }>;
    slides: Array<{
      slideId: string;
      selectedSlideIndex: number;
      selectedRole: string;
      replacedSlots: Array<{
        shapeId: string;
        slotType: string;
      }>;
    }>;
  };
}

export interface OfficeCliDeckSpec {
  deckId: string;
  title: string;
  generationEngine: "officecli";
  generationRules: readonly string[];
  slides: OfficeCliSlideSpec[];
  sourceEvidenceIds: string[];
}

export interface OfficeCliSlideSpec {
  slideId: string;
  kind: "cover" | "agenda" | "section" | "content" | "summary";
  title: string;
  bullets: string[];
  sourceEvidenceIds: string[];
  preferredTemplateRole: TemplateSlideRole;
}

interface NativeTheme {
  primary: string;
  secondary: string;
  accent: string;
  background: string;
  panel: string;
  text: string;
  muted: string;
  titleFont: string;
  bodyFont: string;
}

interface NativeShape {
  text?: string;
  x: string;
  y: string;
  width: string;
  height: string;
  preset?: "rect" | "roundRect" | "ellipse" | "triangle" | "diamond" | "rightArrow";
  fill?: string;
  line?: string;
  lineWidth?: string;
  lineDash?: string;
  color?: string;
  size?: string;
  bold?: boolean;
  font?: string;
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  margin?: string;
  opacity?: number;
  shadow?: string;
  name?: string;
}

export async function generatePptWithOfficeCli(input: OfficeCliPptGenerationInput): Promise<OfficeCliPptGenerationResult> {
  const deckSpec = buildOfficeCliDeckSpec(input);
  const theme = buildNativeTheme(input.templateProfile);
  const workDir = await mkdtemp(join(tmpdir(), "ppt-builders-officecli-native-"));
  const outputPath = join(workDir, "generated.pptx");

  try {
    await runOfficeCli(["create", outputPath], { cwd: workDir });

    for (let index = 0; index < deckSpec.slides.length; index += 1) {
      await renderNativeSlide({
        outputPath,
        slideNumber: index + 1,
        slideSpec: deckSpec.slides[index],
        slideCount: deckSpec.slides.length,
        theme
      });
    }

    const qa = await collectOfficeCliQa(outputPath, workDir);
    const pptxBase64 = (await readFile(outputPath)).toString("base64");
    await runOfficeCli(["close", outputPath], { cwd: workDir, timeoutMs: 30_000 }).catch(() => undefined);

    return {
      deckId: deckSpec.deckId,
      pptxBase64,
      summary: `Generated a ${deckSpec.slides.length}-slide deck with the officeCLI-native deck builder.`,
      qa,
      deckSpec,
      templateReplacement: {
        selectedSlideIndex: 0,
        selectedRole: "officecli-native",
        replacedSlots: [],
        slides: deckSpec.slides.map((slide) => ({
          slideId: slide.slideId,
          selectedSlideIndex: 0,
          selectedRole: "officecli-native",
          replacedSlots: []
        }))
      }
    };
  } finally {
    await rm(workDir, { recursive: true, force: true });
  }
}

export function buildOfficeCliDeckSpec(input: {
  instruction: string;
  reportFileName: string;
  evidenceIndex: EvidenceIndex;
}): OfficeCliDeckSpec {
  const evidence = input.evidenceIndex.blocks.filter((block) => block.text.trim());
  const selectedEvidence = selectEvidence(evidence, input.instruction, 16);
  const title = makeDeckTitle(input.instruction, selectedEvidence, input.reportFileName);
  const sections = buildContentSections(selectedEvidence).slice(0, 5);
  const slides: OfficeCliSlideSpec[] = [
    {
      slideId: "slide_cover",
      kind: "cover",
      title,
      bullets: [input.reportFileName, "基于报告证据生成"],
      sourceEvidenceIds: selectedEvidence.slice(0, 1).map((block) => block.evidenceId),
      preferredTemplateRole: "cover"
    },
    {
      slideId: "slide_agenda",
      kind: "agenda",
      title: "目录",
      bullets: sections.map((section) => section.title),
      sourceEvidenceIds: selectedEvidence.slice(0, 4).map((block) => block.evidenceId),
      preferredTemplateRole: "agenda"
    },
    ...sections.map((section, index) => ({
      slideId: `slide_content_${index + 1}`,
      kind: "content" as const,
      title: section.title,
      bullets: section.bullets,
      sourceEvidenceIds: section.evidenceIds,
      preferredTemplateRole: section.hasVisualHint ? ("content_image" as const) : ("content_text" as const)
    })),
    {
      slideId: "slide_summary",
      kind: "summary",
      title: "结论与启示",
      bullets: buildSummaryBullets(selectedEvidence),
      sourceEvidenceIds: selectedEvidence.slice(0, 4).map((block) => block.evidenceId),
      preferredTemplateRole: "summary"
    }
  ];

  return {
    deckId: `deck_${Date.now()}`,
    title,
    generationEngine: "officecli",
    generationRules: officeCliGenerationRules,
    slides,
    sourceEvidenceIds: Array.from(new Set(slides.flatMap((slide) => slide.sourceEvidenceIds)))
  };
}

async function renderNativeSlide(input: {
  outputPath: string;
  slideNumber: number;
  slideSpec: OfficeCliSlideSpec;
  slideCount: number;
  theme: NativeTheme;
}) {
  const background = input.slideSpec.kind === "cover" ? input.theme.primary : input.theme.background;
  await addSlide(input.outputPath, background);

  if (input.slideSpec.kind === "cover") {
    await renderCoverSlide(input);
  } else if (input.slideSpec.kind === "agenda") {
    await renderAgendaSlide(input);
  } else if (input.slideSpec.kind === "summary") {
    await renderSummarySlide(input);
  } else {
    await renderContentSlide(input);
  }

  await renderFooter(input.outputPath, input.slideNumber, input.slideSpec, input.slideCount, input.theme);
  await addNotes(input.outputPath, input.slideNumber, buildSpeakerNotes(input.slideSpec));
}

async function renderCoverSlide(input: {
  outputPath: string;
  slideNumber: number;
  slideSpec: OfficeCliSlideSpec;
  slideCount: number;
  theme: NativeTheme;
}) {
  const { outputPath, slideNumber, slideSpec, theme } = input;
  await addShape(outputPath, slideNumber, {
    x: "0in",
    y: "0in",
    width: "13.333in",
    height: "7.5in",
    preset: "rect",
    fill: theme.primary,
    line: "none",
    name: "cover_background"
  });
  await addShape(outputPath, slideNumber, {
    x: "8.7in",
    y: "-0.15in",
    width: "4.9in",
    height: "7.8in",
    preset: "rect",
    fill: theme.secondary,
    line: "none",
    opacity: 0.35,
    name: "cover_accent_panel"
  });
  await addShape(outputPath, slideNumber, {
    text: slideSpec.title,
    x: "0.85in",
    y: "1.85in",
    width: "8.4in",
    height: "1.3in",
    fill: "none",
    line: "none",
    color: "FFFFFF",
    size: "38pt",
    bold: true,
    font: theme.titleFont,
    margin: "0.05in",
    name: "cover_title"
  });
  await addShape(outputPath, slideNumber, {
    text: slideSpec.bullets.join("\n"),
    x: "0.9in",
    y: "3.35in",
    width: "6.6in",
    height: "0.8in",
    fill: "none",
    line: "none",
    color: "F4ECF7",
    size: "17pt",
    font: theme.bodyFont,
    margin: "0.05in",
    name: "cover_subtitle"
  });
  await addEvidenceTower(outputPath, slideNumber, theme, "9.05in", "1.45in", "3.2in", "4.65in");
}

async function renderAgendaSlide(input: {
  outputPath: string;
  slideNumber: number;
  slideSpec: OfficeCliSlideSpec;
  slideCount: number;
  theme: NativeTheme;
}) {
  const { outputPath, slideNumber, slideSpec, theme } = input;
  await renderTitle(outputPath, slideNumber, slideSpec.title, theme);
  const bullets = slideSpec.bullets.length > 0 ? slideSpec.bullets : ["核心内容", "方法与证据", "结论启示"];
  for (let index = 0; index < bullets.slice(0, 5).length; index += 1) {
    const y = 1.55 + index * 0.95;
    await addShape(outputPath, slideNumber, {
      x: "0.9in",
      y: `${y}in`,
      width: "11.3in",
      height: "0.68in",
      preset: "roundRect",
      fill: index % 2 === 0 ? "F8F5FA" : "F5F8FA",
      line: "E8E1EA:0.8:solid",
      shadow: "000000",
      name: `agenda_card_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: String(index + 1).padStart(2, "0"),
      x: "1.15in",
      y: `${y + 0.12}in`,
      width: "0.52in",
      height: "0.42in",
      preset: "ellipse",
      fill: theme.primary,
      line: "none",
      color: "FFFFFF",
      size: "13pt",
      bold: true,
      font: theme.bodyFont,
      align: "center",
      valign: "middle",
      margin: "0.02in",
      name: `agenda_number_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: bullets[index],
      x: "1.9in",
      y: `${y + 0.1}in`,
      width: "9.7in",
      height: "0.45in",
      fill: "none",
      line: "none",
      color: theme.text,
      size: "20pt",
      bold: true,
      font: theme.titleFont,
      margin: "0.02in",
      name: `agenda_text_${index + 1}`
    });
  }
}

async function renderContentSlide(input: {
  outputPath: string;
  slideNumber: number;
  slideSpec: OfficeCliSlideSpec;
  slideCount: number;
  theme: NativeTheme;
}) {
  const { outputPath, slideNumber, slideSpec, theme } = input;
  await renderTitle(outputPath, slideNumber, slideSpec.title, theme);
  const bullets = normalizeBullets(slideSpec.bullets, 4);

  for (let index = 0; index < bullets.length; index += 1) {
    const y = 1.45 + index * 1.05;
    await addShape(outputPath, slideNumber, {
      x: "0.75in",
      y: `${y}in`,
      width: "5.45in",
      height: "0.82in",
      preset: "roundRect",
      fill: "FFFFFF",
      line: "E1E6EA:0.8:solid",
      shadow: "000000",
      name: `content_card_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: `${index + 1}`,
      x: "0.98in",
      y: `${y + 0.24}in`,
      width: "0.34in",
      height: "0.34in",
      preset: "ellipse",
      fill: index % 2 === 0 ? theme.secondary : theme.accent,
      line: "none",
      color: "FFFFFF",
      size: "11pt",
      bold: true,
      font: theme.bodyFont,
      align: "center",
      valign: "middle",
      margin: "0.01in",
      name: `content_card_number_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: bullets[index],
      x: "1.45in",
      y: `${y + 0.16}in`,
      width: "4.45in",
      height: "0.5in",
      fill: "none",
      line: "none",
      color: theme.text,
      size: "14.5pt",
      font: theme.bodyFont,
      margin: "0.02in",
      name: `content_card_text_${index + 1}`
    });
  }

  await renderEvidenceVisual(outputPath, slideNumber, slideSpec, theme);
}

async function renderSummarySlide(input: {
  outputPath: string;
  slideNumber: number;
  slideSpec: OfficeCliSlideSpec;
  slideCount: number;
  theme: NativeTheme;
}) {
  const { outputPath, slideNumber, slideSpec, theme } = input;
  await renderTitle(outputPath, slideNumber, slideSpec.title, theme);
  const bullets = normalizeBullets(slideSpec.bullets, 3);
  const fills = ["F6EFF8", "F0F7FA", "F7F5F0"];
  for (let index = 0; index < bullets.length; index += 1) {
    const x = 0.8 + index * 4.12;
    await addShape(outputPath, slideNumber, {
      x: `${x}in`,
      y: "1.75in",
      width: "3.55in",
      height: "3.6in",
      preset: "roundRect",
      fill: fills[index % fills.length],
      line: "DFD7E2:0.8:solid",
      shadow: "000000",
      name: `summary_card_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: `0${index + 1}`,
      x: `${x + 0.3}in`,
      y: "2.1in",
      width: "0.7in",
      height: "0.45in",
      fill: "none",
      line: "none",
      color: theme.primary,
      size: "24pt",
      bold: true,
      font: theme.titleFont,
      name: `summary_index_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: bullets[index],
      x: `${x + 0.35}in`,
      y: "2.82in",
      width: "2.85in",
      height: "1.6in",
      fill: "none",
      line: "none",
      color: theme.text,
      size: "18pt",
      bold: true,
      font: theme.titleFont,
      margin: "0.04in",
      name: `summary_text_${index + 1}`
    });
  }
}

async function renderTitle(outputPath: string, slideNumber: number, title: string, theme: NativeTheme) {
  await addShape(outputPath, slideNumber, {
    text: title,
    x: "0.65in",
    y: "0.35in",
    width: "10.4in",
    height: "0.55in",
    fill: "none",
    line: "none",
    color: theme.primary,
    size: "28pt",
    bold: true,
    font: theme.titleFont,
    margin: "0.02in",
    name: "slide_title"
  });
  await addShape(outputPath, slideNumber, {
    x: "0.65in",
    y: "0.98in",
    width: "11.9in",
    height: "0.03in",
    preset: "rect",
    fill: theme.primary,
    line: "none",
    name: "title_rule"
  });
}

async function renderFooter(
  outputPath: string,
  slideNumber: number,
  slideSpec: OfficeCliSlideSpec,
  slideCount: number,
  theme: NativeTheme
) {
  const color = slideSpec.kind === "cover" ? "F6EEF7" : theme.muted;
  await addShape(outputPath, slideNumber, {
    text: `officeCLI-native | ${slideNumber}/${slideCount}`,
    x: "10.45in",
    y: "7.05in",
    width: "2.15in",
    height: "0.22in",
    fill: "none",
    line: "none",
    color,
    size: "8.5pt",
    font: theme.bodyFont,
    align: "right",
    margin: "0.01in",
    name: "footer_page"
  });
}

async function renderEvidenceVisual(outputPath: string, slideNumber: number, slideSpec: OfficeCliSlideSpec, theme: NativeTheme) {
  await addShape(outputPath, slideNumber, {
    x: "6.75in",
    y: "1.35in",
    width: "5.75in",
    height: "4.85in",
    preset: "roundRect",
    fill: theme.panel,
    line: "E0D9E5:0.8:solid",
    name: "visual_panel"
  });
  await addShape(outputPath, slideNumber, {
    text: "数据证据层级",
    x: "7.25in",
    y: "1.7in",
    width: "4.7in",
    height: "0.42in",
    fill: "none",
    line: "none",
    color: theme.primary,
    size: "19pt",
    bold: true,
    font: theme.titleFont,
    align: "center",
    name: "visual_title"
  });

  const labels = buildEvidenceLabels(slideSpec);
  const colors = [theme.secondary, theme.primary, theme.accent];
  for (let index = 0; index < labels.length; index += 1) {
    const y = 2.45 + index * 0.9;
    await addShape(outputPath, slideNumber, {
      x: "7.55in",
      y: `${y}in`,
      width: "3.95in",
      height: "0.58in",
      preset: "roundRect",
      fill: "FFFFFF",
      line: `${colors[index % colors.length]}:1.2:solid`,
      name: `visual_level_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      x: "7.75in",
      y: `${y + 0.16}in`,
      width: "0.25in",
      height: "0.25in",
      preset: "ellipse",
      fill: colors[index % colors.length],
      line: "none",
      name: `visual_dot_${index + 1}`
    });
    await addShape(outputPath, slideNumber, {
      text: labels[index],
      x: "8.15in",
      y: `${y + 0.1}in`,
      width: "3.05in",
      height: "0.34in",
      fill: "none",
      line: "none",
      color: theme.text,
      size: "13pt",
      bold: true,
      font: theme.bodyFont,
      margin: "0.01in",
      name: `visual_label_${index + 1}`
    });
  }

  await addShape(outputPath, slideNumber, {
    text: `${slideSpec.sourceEvidenceIds.length || 1}`,
    x: "9.05in",
    y: "5.1in",
    width: "0.75in",
    height: "0.52in",
    preset: "ellipse",
    fill: theme.primary,
    line: "none",
    color: "FFFFFF",
    size: "21pt",
    bold: true,
    font: theme.titleFont,
    align: "center",
    valign: "middle",
    margin: "0.01in",
    name: "visual_evidence_count"
  });
  await addShape(outputPath, slideNumber, {
    text: "条报告证据",
    x: "9.85in",
    y: "5.22in",
    width: "1.4in",
    height: "0.28in",
    fill: "none",
    line: "none",
    color: theme.muted,
    size: "11pt",
    font: theme.bodyFont,
    margin: "0.01in",
    name: "visual_evidence_caption"
  });
}

async function addEvidenceTower(
  outputPath: string,
  slideNumber: number,
  theme: NativeTheme,
  x: string,
  y: string,
  width: string,
  height: string
) {
  await addShape(outputPath, slideNumber, {
    x,
    y,
    width,
    height,
    preset: "roundRect",
    fill: "FFFFFF",
    line: "FFFFFF:0.5:solid",
    opacity: 0.16,
    name: "cover_visual_shell"
  });
  const rows = [
    { label: "静态供给", fill: theme.secondary },
    { label: "体验表达", fill: theme.accent },
    { label: "动态行为", fill: "FFFFFF" }
  ];
  for (let index = 0; index < rows.length; index += 1) {
    await addShape(outputPath, slideNumber, {
      text: rows[index].label,
      x: "9.55in",
      y: `${2.05 + index * 1.05}in`,
      width: "2.25in",
      height: "0.62in",
      preset: "roundRect",
      fill: rows[index].fill,
      line: "FFFFFF:0.8:solid",
      color: index === 2 ? theme.primary : "FFFFFF",
      size: "16pt",
      bold: true,
      font: theme.titleFont,
      align: "center",
      valign: "middle",
      margin: "0.03in",
      name: `cover_visual_level_${index + 1}`
    });
  }
}

async function addSlide(outputPath: string, background: string) {
  await runOfficeCli(["add", outputPath, "/", "--type", "slide", "--prop", "layout=blank", "--prop", `background=#${cleanHex(background)}`]);
}

async function addNotes(outputPath: string, slideNumber: number, text: string) {
  await runOfficeCli(["add", outputPath, `/slide[${slideNumber}]`, "--type", "notes", "--prop", `text=${text}`]).catch(async () => {
    await runOfficeCli(["set", outputPath, `/slide[${slideNumber}]`, "--prop", `notes=${text}`]);
  });
}

async function addShape(outputPath: string, slideNumber: number, shape: NativeShape) {
  const props = shapeToProps(shape);
  const args = ["add", outputPath, `/slide[${slideNumber}]`, "--type", "shape"];
  for (const prop of props) {
    args.push("--prop", prop);
  }
  await runOfficeCli(args);
}

function shapeToProps(shape: NativeShape) {
  const props = [
    `x=${shape.x}`,
    `y=${shape.y}`,
    `width=${shape.width}`,
    `height=${shape.height}`,
    `preset=${shape.preset ?? "rect"}`,
    `autoFit=normal`
  ];

  if (shape.text !== undefined) props.push(`text=${shape.text}`);
  if (shape.fill) props.push(`fill=${shape.fill === "none" ? "none" : `#${cleanHex(shape.fill)}`}`);
  if (shape.line) props.push(`line=${shape.line === "none" ? "none" : normalizeLine(shape.line)}`);
  if (shape.lineWidth) props.push(`lineWidth=${shape.lineWidth}`);
  if (shape.lineDash) props.push(`lineDash=${shape.lineDash}`);
  if (shape.color) props.push(`color=${shape.color === "none" ? "none" : `#${cleanHex(shape.color)}`}`);
  if (shape.size) props.push(`size=${shape.size}`);
  if (shape.bold !== undefined) props.push(`bold=${shape.bold}`);
  if (shape.font) {
    props.push(`font=${shape.font}`, `font.ea=${shape.font}`, `font.latin=${shape.font}`);
  }
  if (shape.align) props.push(`align=${shape.align}`);
  if (shape.valign) props.push(`valign=${shape.valign}`);
  if (shape.margin) props.push(`margin=${shape.margin}`);
  if (shape.opacity !== undefined) props.push(`opacity=${shape.opacity}`);
  if (shape.shadow) props.push(`shadow=#${cleanHex(shape.shadow)}`);
  if (shape.name) props.push(`name=${shape.name}`);
  return props;
}

function buildNativeTheme(profile: TemplateProfile): NativeTheme {
  const colors = profile.theme.colors.map(cleanHex).filter(isHexColor);
  const primary = chooseReadableAccent(colors, "6F1D7A");
  const secondary = chooseDifferentColor(colors, primary, "62A8C8");
  const accent = chooseDifferentColor(colors, primary, "9B2A92", [secondary]);
  const fonts = profile.theme.fonts.filter(Boolean);
  const cjkFont = fonts.find((font) => /yahei|hei|song|noto|source han|等线|黑体|宋体/i.test(font));

  return {
    primary,
    secondary,
    accent,
    background: "FBFAFC",
    panel: "F6F1F7",
    text: "202124",
    muted: "6B6470",
    titleFont: cjkFont ?? fonts[0] ?? "Microsoft YaHei",
    bodyFont: cjkFont ?? fonts[1] ?? fonts[0] ?? "Microsoft YaHei"
  };
}

function chooseReadableAccent(colors: string[], fallback: string) {
  return colors.find((color) => isReadableAccent(color)) ?? fallback;
}

function chooseDifferentColor(colors: string[], primary: string, fallback: string, excluded: string[] = []) {
  return (
    colors.find(
      (color) => isReadableAccent(color) && colorDistance(color, primary) > 90 && excluded.every((item) => colorDistance(color, item) > 65)
    ) ?? fallback
  );
}

function isReadableAccent(color: string) {
  const [r, g, b] = hexToRgb(color);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const saturation = max - min;
  const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return saturation > 35 && luminance > 40 && luminance < 210;
}

function colorDistance(a: string, b: string) {
  const [ar, ag, ab] = hexToRgb(a);
  const [br, bg, bb] = hexToRgb(b);
  return Math.sqrt((ar - br) ** 2 + (ag - bg) ** 2 + (ab - bb) ** 2);
}

function hexToRgb(color: string) {
  const hex = cleanHex(color);
  return [Number.parseInt(hex.slice(0, 2), 16), Number.parseInt(hex.slice(2, 4), 16), Number.parseInt(hex.slice(4, 6), 16)];
}

function normalizeLine(value: string) {
  const parts = value.split(":");
  if (parts.length === 0) return value;
  parts[0] = `#${cleanHex(parts[0])}`;
  return parts.join(":");
}

function cleanHex(value: string) {
  return value.replace(/^#/, "").toUpperCase();
}

function isHexColor(value: string) {
  return /^[0-9A-F]{6}$/.test(value);
}

function buildContentSections(evidence: EvidenceBlock[]) {
  const groups = groupEvidence(evidence).slice(0, 5);
  if (groups.length === 0) {
    return [
      {
        title: "核心内容",
        bullets: ["报告中未解析到足够文本，请补充报告内容或生成要求。"],
        evidenceIds: [],
        hasVisualHint: false
      }
    ];
  }

  return groups.map((group, index) => ({
    title: inferSectionTitle(group, index + 1),
    bullets: group
      .filter((block) => !isLikelyHeading(block.text))
      .slice(0, 4)
      .map((block) => compactBullet(block.text)),
    evidenceIds: group.map((block) => block.evidenceId),
    hasVisualHint: group.some((block) => /图|表|数据|figure|chart|image|poi|轨迹|位置|空间/i.test(block.text))
  }));
}

function groupEvidence(evidence: EvidenceBlock[]) {
  const headings = evidence.filter((block) => isLikelyHeading(block.text));
  if (headings.length >= 2) {
    const groups: EvidenceBlock[][] = [];
    let current: EvidenceBlock[] = [];
    for (const block of evidence.slice(0, 28)) {
      if (isLikelyHeading(block.text) && current.length > 0) {
        groups.push(current);
        current = [block];
      } else {
        current.push(block);
      }
    }
    if (current.length > 0) {
      groups.push(current);
    }
    return groups.filter((group) => group.some((block) => !isLikelyHeading(block.text)));
  }

  const groups: EvidenceBlock[][] = [];
  for (let index = 0; index < evidence.length; index += 3) {
    groups.push(evidence.slice(index, index + 3));
  }
  return groups;
}

function selectEvidence(evidence: EvidenceBlock[], instruction: string, limit: number) {
  const terms = tokenize(`${instruction} ${evidence.slice(0, 3).map((block) => block.text).join(" ")}`);
  const scored = evidence.map((block) => {
    const text = block.text.toLowerCase();
    const hits = terms.filter((term) => text.includes(term)).length;
    const structureBoost = isLikelyHeading(block.text) ? 0.5 : 0;
    return {
      block,
      score: hits + structureBoost
    };
  });
  const selected = scored
    .sort((a, b) => b.score - a.score || a.block.order - b.block.order)
    .map((item) => item.block)
    .slice(0, limit)
    .sort((a, b) => a.order - b.order);
  return selected.length > 0 ? selected : evidence.slice(0, limit);
}

function inferSectionTitle(group: EvidenceBlock[], index: number) {
  const heading = group.find((block) => isLikelyHeading(block.text));
  if (heading) {
    return cleanTitle(heading.text);
  }
  const first = group[0]?.text ?? "";
  const keywordTitle = titleFromKeywords(first);
  return keywordTitle || `核心发现 ${index}`;
}

function titleFromKeywords(text: string) {
  if (/数据|证据|指标|样本|统计/.test(text)) return "数据证据与指标";
  if (/方法|模型|框架|流程|识别/.test(text)) return "方法框架";
  if (/结果|发现|分析|影响/.test(text)) return "结果分析";
  if (/空间|城市|区域|地理|位置/.test(text)) return "空间特征";
  if (/结论|建议|启示/.test(text)) return "结论启示";
  return "";
}

function buildSummaryBullets(evidence: EvidenceBlock[]) {
  const bullets = evidence
    .filter((block) => /结论|建议|启示|结果|发现|表明|说明/.test(block.text))
    .slice(0, 3)
    .map((block) => compactBullet(block.text));
  return bullets.length > 0 ? bullets : evidence.slice(0, 3).map((block) => compactBullet(block.text));
}

function makeDeckTitle(instruction: string, evidence: EvidenceBlock[], reportFileName: string) {
  const cleaned = cleanTitle(instruction);
  if (cleaned && cleaned !== "生成 PPT") {
    return cleaned;
  }
  const heading = evidence.find((block) => isLikelyHeading(block.text));
  return heading ? cleanTitle(heading.text) : reportFileName.replace(/\.[^.]+$/, "");
}

function buildEvidenceLabels(slideSpec: OfficeCliSlideSpec) {
  const text = `${slideSpec.title} ${slideSpec.bullets.join(" ")}`;
  const labels = [];
  if (/poi|路网|设施|站点|供给|空间/i.test(text)) labels.push("静态供给");
  if (/手机|信令|wifi|gps|lbs|轨迹|行为|动态/i.test(text)) labels.push("动态行为");
  if (/问卷|访谈|观察|体验|感知|遥感|校验/i.test(text)) labels.push("体验校验");
  return [...labels, "指标提取", "交叉验证", "结论归纳"].slice(0, 3);
}

function normalizeBullets(bullets: string[], count: number) {
  const normalized = bullets.map(compactBullet).filter(Boolean);
  while (normalized.length < count) {
    normalized.push(["提取报告证据", "归纳关键发现", "形成页面结论", "保留验证依据"][normalized.length] ?? "补充证据");
  }
  return normalized.slice(0, count);
}

function buildSpeakerNotes(slideSpec: OfficeCliSlideSpec) {
  const evidence = slideSpec.sourceEvidenceIds.length > 0 ? slideSpec.sourceEvidenceIds.join(", ") : "无明确证据编号";
  return [`${slideSpec.title}`, ...slideSpec.bullets.slice(0, 4).map((bullet) => `- ${bullet}`), `证据来源: ${evidence}`].join("\n");
}

function cleanTitle(text: string) {
  const cleaned = normalizeWhitespace(text)
    .replace(/^#+\s*/, "")
    .replace(/^[一二三四五六七八九十\d]+[、.)）\-\s]*/, "");
  return cleaned.length > 28 ? `${cleaned.slice(0, 28)}...` : cleaned;
}

function compactBullet(text: string) {
  const cleaned = normalizeWhitespace(text).replace(/^[-•·\s]*/, "");
  return cleaned.length > 58 ? `${cleaned.slice(0, 58)}...` : cleaned;
}

function isLikelyHeading(text: string) {
  const cleaned = normalizeWhitespace(text);
  return cleaned.length > 0 && cleaned.length <= 32 && !/[。；;,.，]/.test(cleaned);
}

function normalizeWhitespace(text: string) {
  return text.replace(/\s+/g, " ").trim();
}

function tokenize(text: string) {
  return Array.from(new Set(text.toLowerCase().match(/[a-z0-9]{2,}|[\u4e00-\u9fa5]{2,}/g) ?? []));
}

async function collectOfficeCliQa(outputPath: string, workDir: string) {
  const messages: string[] = [];

  try {
    const issues = await runOfficeCli(["view", outputPath, "issues", "--limit", "12"], { cwd: workDir, timeoutMs: 60_000 });
    const issueText = `${issues.stdout}${issues.stderr}`.trim();
    if (issueText) {
      messages.push(issueText);
    }
  } catch (error) {
    messages.push(error instanceof Error ? error.message : "officeCLI issues check failed.");
  }

  try {
    await runOfficeCli(["validate", outputPath], { cwd: workDir, timeoutMs: 60_000 });
    messages.push("officeCLI validate passed.");
  } catch (error) {
    messages.push(error instanceof Error ? `officeCLI validate reported issues: ${error.message}` : "officeCLI validate reported issues.");
  }

  return messages.filter(Boolean).join("\n").trim() || "officeCLI native generation completed.";
}
