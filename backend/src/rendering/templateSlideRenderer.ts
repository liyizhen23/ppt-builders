import { createRequire } from "node:module";
import type { TemplateProfile } from "../templates/templateProfile.js";
import type { TemplateReplaceableSlot } from "../templates/templateCapabilities.js";

const require = createRequire(import.meta.url);
const PptxGenerator = require("pptxgenjs") as { new (): any };

const wideLayout = {
  width: 13.333,
  height: 7.5
};

export interface TemplateReplacementInput {
  profile: TemplateProfile;
  reportFileName: string;
  templateFileName: string;
  instruction: string;
}

export interface TemplateReplacementResult {
  pptxBase64: string;
  selectedSlideIndex: number;
  selectedRole: string;
  replacedSlots: Array<{
    shapeId: string;
    slotType: string;
  }>;
}

export async function renderTemplateReplacementDeck(input: TemplateReplacementInput): Promise<TemplateReplacementResult> {
  const selectedSlideIndex = chooseTemplateSlide(input.profile);
  const selectedSlide = input.profile.slides.find((slide) => slide.index === selectedSlideIndex) ?? input.profile.slides[0];
  const selectedSlots = input.profile.capabilities.replaceableSlots.filter((slot) => slot.slideIndex === selectedSlide.index);
  const titleSlot = selectSlot(selectedSlots, "title");
  const bodySlots = selectedSlots.filter((slot) => slot.slotType === "body").slice(0, 3);
  const imageSlot = selectSlot(selectedSlots, "image");

  const pptx = new PptxGenerator();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "AI PPT Builder";
  pptx.subject = "Template replacement smoke test";
  pptx.title = "Template Replacement";
  pptx.company = "ppt-builders";
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: input.profile.capabilities.styleTokens.title?.fontFace ?? "Microsoft YaHei",
    bodyFontFace: input.profile.capabilities.styleTokens.body?.fontFace ?? "Microsoft YaHei",
    lang: "zh-CN"
  };

  const slide = pptx.addSlide();
  applyTemplateInspiredBackground(slide, input.profile);

  const replacedSlots: TemplateReplacementResult["replacedSlots"] = [];
  const title = normalizeInstructionTitle(input.instruction);
  const bullets = buildBodyBullets(input);

  if (titleSlot?.bbox) {
    addTextInSlot(slide, input.profile, titleSlot, title, "title");
    replacedSlots.push({ shapeId: titleSlot.shapeId, slotType: titleSlot.slotType });
  } else {
    slide.addText(title, {
      x: 0.7,
      y: 0.55,
      w: 11.8,
      h: 0.65,
      fontFace: input.profile.capabilities.styleTokens.title?.fontFace ?? "Microsoft YaHei",
      fontSize: saneFontSize(input.profile.capabilities.styleTokens.title?.fontSize, 30),
      bold: true,
      color: input.profile.capabilities.styleTokens.title?.color ?? "1F2933",
      margin: 0
    });
  }

  if (bodySlots.length > 0) {
    bodySlots.forEach((slot, index) => {
      if (!slot.bbox) {
        return;
      }
      addTextInSlot(slide, input.profile, slot, bullets[index] ?? bullets[bullets.length - 1], "body");
      replacedSlots.push({ shapeId: slot.shapeId, slotType: slot.slotType });
    });
  } else {
    slide.addText(bullets.map((bullet) => `• ${bullet}`).join("\n"), {
      x: 0.9,
      y: 1.65,
      w: imageSlot?.bbox ? 6.2 : 11.2,
      h: 3.2,
      fontFace: input.profile.capabilities.styleTokens.body?.fontFace ?? "Microsoft YaHei",
      fontSize: saneFontSize(input.profile.capabilities.styleTokens.body?.fontSize, 18),
      color: input.profile.capabilities.styleTokens.body?.color ?? "25313D",
      breakLine: false,
      fit: "shrink",
      margin: 0.08
    });
  }

  if (imageSlot?.bbox) {
    addImagePlaceholder(slide, input.profile, imageSlot);
    replacedSlots.push({ shapeId: imageSlot.shapeId, slotType: imageSlot.slotType });
  }

  addFooter(slide, input.profile, selectedSlide.index, input.templateFileName);

  const output = await pptx.write({ outputType: "base64" });
  return {
    pptxBase64: String(output),
    selectedSlideIndex: selectedSlide.index,
    selectedRole: selectedSlide.role,
    replacedSlots
  };
}

function chooseTemplateSlide(profile: TemplateProfile) {
  const preferred =
    profile.capabilities.recommendedSlides.content_text[0] ??
    profile.capabilities.recommendedSlides.content_image[0] ??
    profile.capabilities.slideSummaries
      .slice()
      .sort((a, b) => b.replaceableSlotCount - a.replaceableSlotCount || a.slideIndex - b.slideIndex)[0]?.slideIndex;

  return preferred ?? profile.slides[0]?.index ?? 1;
}

function selectSlot(slots: TemplateReplaceableSlot[], slotType: TemplateReplaceableSlot["slotType"]) {
  return slots
    .filter((slot) => slot.slotType === slotType)
    .sort((a, b) => b.confidence - a.confidence)[0];
}

function addTextInSlot(
  slide: any,
  profile: TemplateProfile,
  slot: TemplateReplaceableSlot,
  text: string,
  styleType: "title" | "body"
) {
  const box = emuToInches(profile, slot.bbox);
  const style = styleType === "title" ? profile.capabilities.styleTokens.title : profile.capabilities.styleTokens.body;
  slide.addText(text, {
    ...box,
    fontFace: style?.fontFace ?? "Microsoft YaHei",
    fontSize: saneFontSize(style?.fontSize, styleType === "title" ? 30 : 18),
    bold: styleType === "title",
    color: style?.color ?? (styleType === "title" ? "1F2933" : "25313D"),
    fit: "shrink",
    valign: "mid",
    margin: 0.06,
    breakLine: false
  });
}

function addImagePlaceholder(slide: any, profile: TemplateProfile, slot: TemplateReplaceableSlot) {
  const box = emuToInches(profile, slot.bbox);
  slide.addShape("rect", {
    ...box,
    fill: { color: "F3F6F8", transparency: 6 },
    line: { color: primaryColor(profile), width: 1.2, dash: "dash" }
  });
  slide.addText("Image / figure slot", {
    x: box.x + 0.12,
    y: box.y + box.h / 2 - 0.15,
    w: Math.max(box.w - 0.24, 0.5),
    h: 0.3,
    fontFace: profile.capabilities.styleTokens.caption?.fontFace ?? "Microsoft YaHei",
    fontSize: 11,
    color: "667085",
    align: "center",
    margin: 0
  });
}

function applyTemplateInspiredBackground(slide: any, profile: TemplateProfile) {
  slide.background = { color: "FFFFFF" };
  const color = primaryColor(profile);
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: wideLayout.width,
    h: 0.16,
    fill: { color },
    line: { color }
  });
  slide.addShape("rect", {
    x: 0,
    y: wideLayout.height - 0.1,
    w: wideLayout.width,
    h: 0.1,
    fill: { color, transparency: 18 },
    line: { color, transparency: 100 }
  });
}

function addFooter(slide: any, profile: TemplateProfile, selectedSlideIndex: number, templateFileName: string) {
  slide.addText(`Template slide ${selectedSlideIndex} • ${templateFileName}`, {
    x: 0.65,
    y: 7.08,
    w: 11.9,
    h: 0.25,
    fontFace: profile.capabilities.styleTokens.caption?.fontFace ?? "Microsoft YaHei",
    fontSize: saneFontSize(profile.capabilities.styleTokens.caption?.fontSize, 10),
    color: "667085",
    margin: 0
  });
}

function buildBodyBullets(input: TemplateReplacementInput) {
  return [
    `报告文件：${input.reportFileName}`,
    `生成要求：${input.instruction.trim() || "未提供具体生成要求"}`,
    `已根据 Template Profile 选择模板页并填充可替换槽位。`
  ];
}

function normalizeInstructionTitle(instruction: string) {
  const cleaned = instruction.trim().replace(/\s+/g, " ");
  if (!cleaned) {
    return "模板页替换生成测试";
  }
  return cleaned.length > 34 ? `${cleaned.slice(0, 34)}...` : cleaned;
}

function emuToInches(profile: TemplateProfile, bbox: TemplateReplaceableSlot["bbox"]) {
  if (!bbox) {
    return { x: 0.75, y: 1.2, w: 11, h: 0.6 };
  }

  const pageWidth = profile.pageSize.cx ?? 12192000;
  const pageHeight = profile.pageSize.cy ?? 6858000;
  return {
    x: (bbox.x / pageWidth) * wideLayout.width,
    y: (bbox.y / pageHeight) * wideLayout.height,
    w: Math.max((bbox.cx / pageWidth) * wideLayout.width, 0.3),
    h: Math.max((bbox.cy / pageHeight) * wideLayout.height, 0.2)
  };
}

function saneFontSize(value: number | null | undefined, fallback: number) {
  if (!value || value < 6 || value > 72) {
    return fallback;
  }
  return value;
}

function primaryColor(profile: TemplateProfile) {
  return profile.capabilities.styleTokens.palette.find((color) => color !== "FFFFFF") ?? "2F5D8C";
}
