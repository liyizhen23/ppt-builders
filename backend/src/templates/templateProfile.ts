import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { buildTemplateCapabilities, TemplateCapabilities } from "./templateCapabilities.js";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  parseTagValue: false,
  parseAttributeValue: false
});

export interface TemplateProfile {
  templateId: string;
  sourceFileName: string;
  generatedAt: string;
  counts: {
    slides: number;
    layouts: number;
    masters: number;
    media: number;
  };
  pageSize: {
    cx: number | null;
    cy: number | null;
  };
  slides: TemplateSlideProfile[];
  layouts: TemplatePartProfile[];
  masters: TemplatePartProfile[];
  media: TemplateMediaProfile[];
  theme: {
    fonts: string[];
    colors: string[];
  };
  capabilities: TemplateCapabilities;
}

export interface TemplateSlideProfile {
  index: number;
  partName: string;
  title: string | null;
  role: TemplateSlideRole;
  textBoxes: TemplateTextBoxProfile[];
  pictureBoxes: TemplatePictureBoxProfile[];
  placeholders: TemplatePlaceholderProfile[];
  shapeCount: number;
}

export interface TemplatePartProfile {
  partName: string;
  name: string | null;
  placeholders: TemplatePlaceholderProfile[];
  shapeCount: number;
}

export interface TemplateTextBoxProfile {
  id: string;
  name: string | null;
  text: string;
  bbox: BBox | null;
  fontSize: number | null;
  fontFace: string | null;
  color: string | null;
  placeholderType: string | null;
}

export interface TemplatePictureBoxProfile {
  id: string;
  name: string | null;
  bbox: BBox | null;
  placeholderType: string | null;
}

export interface TemplatePlaceholderProfile {
  id: string;
  name: string | null;
  type: string | null;
  idx: string | null;
  bbox: BBox | null;
}

export interface TemplateMediaProfile {
  partName: string;
  extension: string;
  size: number;
}

export interface BBox {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

export type TemplateSlideRole =
  | "cover"
  | "agenda"
  | "section_divider"
  | "content_text"
  | "content_image"
  | "content_chart"
  | "comparison"
  | "summary"
  | "closing"
  | "unknown";

export async function analyzePptxTemplate(input: {
  templateId: string;
  sourceFileName: string;
  buffer: Buffer;
}): Promise<TemplateProfile> {
  const zip = await JSZip.loadAsync(input.buffer);
  const names = Object.keys(zip.files);
  const slideNames = sortedPartNames(names, /^ppt\/slides\/slide\d+\.xml$/);
  const layoutNames = sortedPartNames(names, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/);
  const masterNames = sortedPartNames(names, /^ppt\/slideMasters\/slideMaster\d+\.xml$/);
  const mediaNames = sortedPartNames(names, /^ppt\/media\/.+$/);

  const presentationXml = await readXml(zip, "ppt/presentation.xml");
  const pageSize = extractPageSize(presentationXml);

  const slides: TemplateSlideProfile[] = [];
  for (let index = 0; index < slideNames.length; index += 1) {
    const partName = slideNames[index];
    const xml = await readXml(zip, partName);
    const slide = parseSlide(partName, index + 1, xml);
    slides.push({
      ...slide,
      role: inferSlideRole(slide)
    });
  }

  const layouts = await parseParts(zip, layoutNames);
  const masters = await parseParts(zip, masterNames);
  const media = await parseMedia(zip, mediaNames);
  const theme = await extractTheme(zip, names);

  const rawProfile = {
    templateId: input.templateId,
    sourceFileName: input.sourceFileName,
    generatedAt: new Date().toISOString(),
    counts: {
      slides: slideNames.length,
      layouts: layoutNames.length,
      masters: masterNames.length,
      media: mediaNames.length
    },
    pageSize,
    slides,
    layouts,
    masters,
    media,
    theme
  };

  return {
    ...rawProfile,
    capabilities: buildTemplateCapabilities(rawProfile)
  };
}

async function parseParts(zip: JSZip, partNames: string[]): Promise<TemplatePartProfile[]> {
  const parts: TemplatePartProfile[] = [];
  for (const partName of partNames) {
    const xml = await readXml(zip, partName);
    const parsed = parseSlideLikeXml(xml);
    parts.push({
      partName,
      name: extractCommonSlideDataName(xml),
      placeholders: parsed.placeholders,
      shapeCount: parsed.shapeCount
    });
  }

  return parts;
}

function parseSlide(partName: string, index: number, xml: unknown): Omit<TemplateSlideProfile, "role"> {
  const parsed = parseSlideLikeXml(xml);
  const title = parsed.textBoxes.find((box) => box.placeholderType === "title")?.text || parsed.textBoxes[0]?.text || null;

  return {
    index,
    partName,
    title: normalizeText(title),
    textBoxes: parsed.textBoxes,
    pictureBoxes: parsed.pictureBoxes,
    placeholders: parsed.placeholders,
    shapeCount: parsed.shapeCount
  };
}

function parseSlideLikeXml(xml: unknown) {
  const root = getSlideRoot(xml);
  const tree = root?.["p:cSld"]?.["p:spTree"];
  const shapes = asArray(tree?.["p:sp"]);
  const pictures = asArray(tree?.["p:pic"]);

  const textBoxes: TemplateTextBoxProfile[] = [];
  const placeholders: TemplatePlaceholderProfile[] = [];

  for (const shape of shapes) {
    const cNvPr = shape?.["p:nvSpPr"]?.["p:cNvPr"];
    const placeholder = shape?.["p:nvSpPr"]?.["p:nvPr"]?.["p:ph"];
    const bbox = extractBBox(shape?.["p:spPr"]?.["a:xfrm"]);
    const id = String(cNvPr?.id ?? "");
    const name = cNvPr?.name ? String(cNvPr.name) : null;
    const placeholderType = placeholder?.type ? String(placeholder.type) : null;

    if (placeholder || bbox) {
      placeholders.push({
        id,
        name,
        type: placeholderType,
        idx: placeholder?.idx ? String(placeholder.idx) : null,
        bbox
      });
    }

    const text = extractText(shape?.["p:txBody"]);
    if (text) {
      textBoxes.push({
        id,
        name,
        text,
        bbox,
        fontSize: extractFontSize(shape?.["p:txBody"]),
        fontFace: extractFontFace(shape?.["p:txBody"]),
        color: extractColor(shape?.["p:txBody"]),
        placeholderType
      });
    }
  }

  const pictureBoxes: TemplatePictureBoxProfile[] = pictures.map((picture) => {
    const cNvPr = picture?.["p:nvPicPr"]?.["p:cNvPr"];
    const placeholder = picture?.["p:nvPicPr"]?.["p:nvPr"]?.["p:ph"];
    return {
      id: String(cNvPr?.id ?? ""),
      name: cNvPr?.name ? String(cNvPr.name) : null,
      bbox: extractBBox(picture?.["p:spPr"]?.["a:xfrm"]),
      placeholderType: placeholder?.type ? String(placeholder.type) : null
    };
  });

  return {
    textBoxes,
    pictureBoxes,
    placeholders,
    shapeCount: shapes.length + pictures.length
  };
}

function inferSlideRole(slide: Omit<TemplateSlideProfile, "role">): TemplateSlideRole {
  const text = `${slide.title ?? ""} ${slide.textBoxes.map((box) => box.text).join(" ")}`.toLowerCase();
  const title = slide.title ?? "";
  const textCount = slide.textBoxes.length;
  const pictureCount = slide.pictureBoxes.length;

  if (slide.index === 1 || /汇报|报告|presentation|title/i.test(title)) {
    return "cover";
  }
  if (/目录|agenda|contents/.test(text)) {
    return "agenda";
  }
  if (/谢谢|thank|致谢/.test(text)) {
    return "closing";
  }
  if (/总结|结论|summary|conclusion/.test(text)) {
    return "summary";
  }
  if (/第[一二三四五六七八九十\d]+章|chapter|section/.test(text) && textCount <= 4) {
    return "section_divider";
  }
  if (/对比|比较|comparison|vs\.?/.test(text)) {
    return "comparison";
  }
  if (/图|表|chart|数据/.test(text) || pictureCount >= 1) {
    return pictureCount >= 1 ? "content_image" : "content_chart";
  }
  if (textCount > 0) {
    return "content_text";
  }

  return "unknown";
}

async function extractTheme(zip: JSZip, names: string[]) {
  const themeNames = sortedPartNames(names, /^ppt\/theme\/theme\d+\.xml$/);
  const fonts = new Set<string>();
  const colors = new Set<string>();

  for (const themeName of themeNames) {
    const theme = await readXml(zip, themeName);
    walk(theme, (node) => {
      if (!isRecord(node)) {
        return;
      }
      if (typeof node.typeface === "string" && node.typeface) {
        fonts.add(node.typeface);
      }
      if (typeof node.val === "string" && /^[0-9A-Fa-f]{6}$/.test(node.val)) {
        colors.add(node.val.toUpperCase());
      }
    });
  }

  return {
    fonts: [...fonts].slice(0, 40),
    colors: [...colors].slice(0, 40)
  };
}

async function parseMedia(zip: JSZip, mediaNames: string[]): Promise<TemplateMediaProfile[]> {
  const media: TemplateMediaProfile[] = [];
  for (const partName of mediaNames) {
    const file = zip.file(partName);
    if (!file) {
      continue;
    }
    const data = await file.async("nodebuffer");
    media.push({
      partName,
      extension: partName.split(".").pop()?.toLowerCase() ?? "",
      size: data.length
    });
  }

  return media;
}

async function readXml(zip: JSZip, partName: string) {
  const file = zip.file(partName);
  if (!file) {
    return null;
  }

  return parser.parse(await file.async("text"));
}

function sortedPartNames(names: string[], pattern: RegExp) {
  return names
    .filter((name) => pattern.test(name))
    .sort((a, b) => naturalPartNumber(a) - naturalPartNumber(b));
}

function naturalPartNumber(name: string) {
  const match = name.match(/(\d+)\.xml$/);
  return match ? Number(match[1]) : 0;
}

function getSlideRoot(xml: unknown): any {
  if (!isRecord(xml)) {
    return null;
  }

  return xml["p:sld"] ?? xml["p:sldLayout"] ?? xml["p:sldMaster"] ?? null;
}

function extractCommonSlideDataName(xml: unknown) {
  const root = getSlideRoot(xml);
  const name = root?.["p:cSld"]?.name;
  return name ? String(name) : null;
}

function extractPageSize(xml: unknown) {
  const size = isRecord(xml) ? (xml["p:presentation"] as any)?.["p:sldSz"] : null;
  return {
    cx: toNumberOrNull(size?.cx),
    cy: toNumberOrNull(size?.cy)
  };
}

function extractBBox(xfrm: any): BBox | null {
  const off = xfrm?.["a:off"];
  const ext = xfrm?.["a:ext"];
  const x = toNumberOrNull(off?.x);
  const y = toNumberOrNull(off?.y);
  const cx = toNumberOrNull(ext?.cx);
  const cy = toNumberOrNull(ext?.cy);

  if (x === null || y === null || cx === null || cy === null) {
    return null;
  }

  return { x, y, cx, cy };
}

function extractText(txBody: any) {
  const fragments: string[] = [];
  walk(txBody, (node) => {
    if (isRecord(node) && typeof node["a:t"] === "string") {
      fragments.push(node["a:t"]);
    }
  });
  return normalizeText(fragments.join(" "));
}

function extractFontSize(txBody: any) {
  let size: number | null = null;
  walk(txBody, (node) => {
    if (size !== null || !isRecord(node)) {
      return;
    }
    const sz = node?.["a:rPr"]?.sz ?? node?.["a:defRPr"]?.sz;
    const numeric = toNumberOrNull(sz);
    if (numeric !== null) {
      size = numeric / 100;
    }
  });
  return size;
}

function extractFontFace(txBody: any) {
  let face: string | null = null;
  walk(txBody, (node) => {
    if (face !== null || !isRecord(node)) {
      return;
    }
    const typeface = node?.["a:latin"]?.typeface ?? node?.["a:ea"]?.typeface;
    if (typeof typeface === "string" && typeface) {
      face = typeface;
    }
  });
  return face;
}

function extractColor(txBody: any) {
  let color: string | null = null;
  walk(txBody, (node) => {
    if (color !== null || !isRecord(node)) {
      return;
    }
    const val = node?.["a:solidFill"]?.["a:srgbClr"]?.val;
    if (typeof val === "string" && /^[0-9A-Fa-f]{6}$/.test(val)) {
      color = val.toUpperCase();
    }
  });
  return color;
}

function walk(value: unknown, visit: (node: unknown) => void) {
  visit(value);
  if (Array.isArray(value)) {
    for (const child of value) {
      walk(child, visit);
    }
  } else if (isRecord(value)) {
    for (const child of Object.values(value)) {
      walk(child, visit);
    }
  }
}

function asArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function toNumberOrNull(value: unknown) {
  if (value === undefined || value === null || value === "") {
    return null;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function normalizeText(value: string | null | undefined) {
  const normalized = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  return normalized || null;
}

function isRecord(value: unknown): value is Record<string, any> {
  return typeof value === "object" && value !== null;
}
