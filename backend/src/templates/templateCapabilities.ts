import type {
  BBox,
  TemplatePictureBoxProfile,
  TemplateProfile,
  TemplateSlideProfile,
  TemplateSlideRole,
  TemplateTextBoxProfile
} from "./templateProfile.js";

export interface TemplateCapabilities {
  roleIndex: Record<TemplateSlideRole, number[]>;
  recommendedSlides: Record<TemplateSlideRole, number[]>;
  replaceableSlots: TemplateReplaceableSlot[];
  styleTokens: TemplateStyleTokens;
  slideSummaries: TemplateSlideSummary[];
}

export interface TemplateReplaceableSlot {
  slideIndex: number;
  shapeId: string;
  shapeName: string | null;
  slotType: "title" | "subtitle" | "body" | "caption" | "image";
  bbox: BBox | null;
  placeholderType: string | null;
  sampleText?: string;
  confidence: number;
}

export interface TemplateStyleTokens {
  title: TemplateTextStyle | null;
  subtitle: TemplateTextStyle | null;
  body: TemplateTextStyle | null;
  caption: TemplateTextStyle | null;
  palette: string[];
  fonts: string[];
}

export interface TemplateTextStyle {
  fontFace: string | null;
  fontSize: number | null;
  color: string | null;
  sampleSlideIndex: number;
  sampleShapeId: string;
}

export interface TemplateSlideSummary {
  slideIndex: number;
  role: TemplateSlideRole;
  title: string | null;
  textSlotCount: number;
  imageSlotCount: number;
  replaceableSlotCount: number;
  score: number;
}

const roles: TemplateSlideRole[] = [
  "cover",
  "agenda",
  "section_divider",
  "content_text",
  "content_image",
  "content_chart",
  "comparison",
  "summary",
  "closing",
  "unknown"
];

export function buildTemplateCapabilities(profile: Omit<TemplateProfile, "capabilities">): TemplateCapabilities {
  const roleIndex = createRoleIndex(profile.slides);
  const replaceableSlots = profile.slides.flatMap((slide) => inferReplaceableSlots(profile, slide));
  const slideSummaries = profile.slides.map((slide) => summarizeSlide(slide, replaceableSlots));
  const recommendedSlides = createRecommendedSlides(roleIndex, slideSummaries);

  return {
    roleIndex,
    recommendedSlides,
    replaceableSlots,
    styleTokens: extractStyleTokens(profile, replaceableSlots),
    slideSummaries
  };
}

function createRoleIndex(slides: TemplateSlideProfile[]) {
  const index = emptyRoleMap();
  for (const slide of slides) {
    index[slide.role].push(slide.index);
  }
  return index;
}

function inferReplaceableSlots(profile: Omit<TemplateProfile, "capabilities">, slide: TemplateSlideProfile) {
  const slots: TemplateReplaceableSlot[] = [];
  const sortedTextBoxes = slide.textBoxes
    .filter((box) => isVisibleBox(profile, box.bbox) && isTemplateText(box.text))
    .sort((a, b) => top(a.bbox) - top(b.bbox));

  for (let index = 0; index < sortedTextBoxes.length; index += 1) {
    const box = sortedTextBoxes[index];
    const slotType = inferTextSlotType(profile, slide, box, index);
    slots.push({
      slideIndex: slide.index,
      shapeId: box.id,
      shapeName: box.name,
      slotType,
      bbox: box.bbox,
      placeholderType: box.placeholderType,
      sampleText: box.text,
      confidence: scoreTextSlotConfidence(slotType, box)
    });
  }

  for (const picture of slide.pictureBoxes.filter((box) => isReplaceablePicture(profile, box))) {
    slots.push({
      slideIndex: slide.index,
      shapeId: picture.id,
      shapeName: picture.name,
      slotType: "image",
      bbox: picture.bbox,
      placeholderType: picture.placeholderType,
      confidence: scoreImageSlotConfidence(profile, picture)
    });
  }

  return slots;
}

function inferTextSlotType(
  profile: Omit<TemplateProfile, "capabilities">,
  slide: TemplateSlideProfile,
  box: TemplateTextBoxProfile,
  orderIndex: number
): TemplateReplaceableSlot["slotType"] {
  const text = box.text.toLowerCase();
  const pageHeight = profile.pageSize.cy ?? 6858000;
  const y = top(box.bbox);
  const fontSize = box.fontSize ?? 0;

  if (box.placeholderType === "title" || fontSize >= 28 || orderIndex === 0 || /标题|title|chapter|section/.test(text)) {
    return slide.role === "cover" && orderIndex > 0 ? "subtitle" : "title";
  }
  if (y > pageHeight * 0.78 || fontSize <= 12 || /注释|说明|caption|source|来源/.test(text)) {
    return "caption";
  }
  if (/副标题|subtitle|汇报人|日期|author/.test(text)) {
    return "subtitle";
  }
  return "body";
}

function scoreTextSlotConfidence(slotType: TemplateReplaceableSlot["slotType"], box: TemplateTextBoxProfile) {
  let score = 0.55;
  if (box.placeholderType) {
    score += 0.15;
  }
  if (/标题|正文|文本|内容|添加|placeholder|title/i.test(box.text)) {
    score += 0.2;
  }
  if (slotType === "title" && (box.fontSize ?? 0) >= 24) {
    score += 0.15;
  }
  return clamp(score);
}

function isReplaceablePicture(profile: Omit<TemplateProfile, "capabilities">, picture: TemplatePictureBoxProfile) {
  if (!picture.bbox) {
    return false;
  }
  const pageWidth = profile.pageSize.cx ?? 12192000;
  const pageHeight = profile.pageSize.cy ?? 6858000;
  const area = picture.bbox.cx * picture.bbox.cy;
  const pageArea = pageWidth * pageHeight;

  if (area >= pageArea * 0.75) {
    return false;
  }
  if (picture.bbox.cx < pageWidth * 0.08 || picture.bbox.cy < pageHeight * 0.08) {
    return false;
  }
  return true;
}

function scoreImageSlotConfidence(profile: Omit<TemplateProfile, "capabilities">, picture: TemplatePictureBoxProfile) {
  if (!picture.bbox) {
    return 0.2;
  }
  const pageWidth = profile.pageSize.cx ?? 12192000;
  const pageHeight = profile.pageSize.cy ?? 6858000;
  const areaRatio = (picture.bbox.cx * picture.bbox.cy) / (pageWidth * pageHeight);
  return clamp(0.45 + Math.min(areaRatio, 0.35));
}

function summarizeSlide(slide: TemplateSlideProfile, slots: TemplateReplaceableSlot[]): TemplateSlideSummary {
  const slideSlots = slots.filter((slot) => slot.slideIndex === slide.index);
  const textSlotCount = slideSlots.filter((slot) => slot.slotType !== "image").length;
  const imageSlotCount = slideSlots.filter((slot) => slot.slotType === "image").length;
  const score = clamp(0.25 + textSlotCount * 0.08 + imageSlotCount * 0.12);

  return {
    slideIndex: slide.index,
    role: slide.role,
    title: slide.title,
    textSlotCount,
    imageSlotCount,
    replaceableSlotCount: slideSlots.length,
    score
  };
}

function createRecommendedSlides(roleIndex: Record<TemplateSlideRole, number[]>, summaries: TemplateSlideSummary[]) {
  const bySlide = new Map(summaries.map((summary) => [summary.slideIndex, summary]));
  const recommended = emptyRoleMap();

  for (const role of roles) {
    recommended[role] = roleIndex[role]
      .map((slideIndex) => bySlide.get(slideIndex))
      .filter((summary): summary is TemplateSlideSummary => Boolean(summary))
      .sort((a, b) => b.score - a.score || a.slideIndex - b.slideIndex)
      .slice(0, 8)
      .map((summary) => summary.slideIndex);
  }

  return recommended;
}

function emptyRoleMap(): Record<TemplateSlideRole, number[]> {
  return {
    cover: [],
    agenda: [],
    section_divider: [],
    content_text: [],
    content_image: [],
    content_chart: [],
    comparison: [],
    summary: [],
    closing: [],
    unknown: []
  };
}

function extractStyleTokens(
  profile: Omit<TemplateProfile, "capabilities">,
  slots: TemplateReplaceableSlot[]
): TemplateStyleTokens {
  return {
    title: findStyleToken(profile, slots, "title"),
    subtitle: findStyleToken(profile, slots, "subtitle"),
    body: findStyleToken(profile, slots, "body"),
    caption: findStyleToken(profile, slots, "caption"),
    palette: profile.theme.colors,
    fonts: profile.theme.fonts
  };
}

function findStyleToken(
  profile: Omit<TemplateProfile, "capabilities">,
  slots: TemplateReplaceableSlot[],
  slotType: TemplateReplaceableSlot["slotType"]
): TemplateTextStyle | null {
  const candidates = slots
    .filter((candidate) => candidate.slotType === slotType)
    .map((slot) => {
      const slide = profile.slides.find((candidate) => candidate.index === slot.slideIndex);
      const box = slide?.textBoxes.find((candidate) => candidate.id === slot.shapeId);
      return { slot, box };
    })
    .filter((candidate): candidate is { slot: TemplateReplaceableSlot; box: TemplateTextBoxProfile } =>
      Boolean(candidate.box)
    )
    .filter((candidate) => isReasonableStyleSample(slotType, candidate.box));

  const selected = candidates.sort((a, b) => {
    if (slotType === "title" || slotType === "subtitle") {
      return (b.box.fontSize ?? 0) - (a.box.fontSize ?? 0) || b.slot.confidence - a.slot.confidence;
    }
    return b.slot.confidence - a.slot.confidence || (b.box.fontSize ?? 0) - (a.box.fontSize ?? 0);
  })[0];

  if (!selected) {
    return null;
  }

  const { slot, box } = selected;
  if (!box) {
    return null;
  }
  return {
    fontFace: box.fontFace,
    fontSize: box.fontSize,
    color: box.color,
    sampleSlideIndex: slot.slideIndex,
    sampleShapeId: slot.shapeId
  };
}

function isReasonableStyleSample(slotType: TemplateReplaceableSlot["slotType"], box: TemplateTextBoxProfile) {
  const size = box.fontSize ?? 0;
  if (size <= 0) {
    return true;
  }
  if (slotType === "title") {
    return size >= 16 && size <= 72;
  }
  if (slotType === "subtitle") {
    return size >= 10 && size <= 40;
  }
  if (slotType === "body") {
    return size >= 8 && size <= 32;
  }
  if (slotType === "caption") {
    return size >= 6 && size <= 18;
  }
  return true;
}

function isTemplateText(text: string) {
  return text.trim().length > 0;
}

function isVisibleBox(profile: Omit<TemplateProfile, "capabilities">, bbox: BBox | null) {
  if (!bbox) {
    return false;
  }
  const pageWidth = profile.pageSize.cx ?? 12192000;
  const pageHeight = profile.pageSize.cy ?? 6858000;
  return bbox.cx > pageWidth * 0.02 && bbox.cy > pageHeight * 0.015;
}

function top(bbox: BBox | null) {
  return bbox?.y ?? 0;
}

function clamp(value: number) {
  return Math.max(0, Math.min(1, Number(value.toFixed(3))));
}
