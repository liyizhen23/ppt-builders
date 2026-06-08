import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const PptxGenerator = require("pptxgenjs") as { new (): any };

export interface ReflowSlideInput {
  instruction: string;
  pageText: string;
  selectedImageFileName?: string | null;
}

export async function renderReflowSlide(input: ReflowSlideInput) {
  const pptx = new PptxGenerator();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "AI PPT Builder";
  pptx.subject = "Current slide reflow replacement";
  pptx.title = normalizeTitle(input.instruction, input.pageText);
  pptx.company = "ppt-builders";
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
    lang: "zh-CN"
  };

  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.16,
    fill: { color: "0066CC" },
    line: { color: "0066CC" }
  });

  const title = normalizeTitle(input.instruction, input.pageText);
  const bullets = buildBullets(input.pageText, input.instruction);
  const hasImage = Boolean(input.selectedImageFileName);

  slide.addText(title, {
    x: 0.72,
    y: 0.55,
    w: 11.8,
    h: 0.55,
    fontFace: "Microsoft YaHei",
    fontSize: 28,
    bold: true,
    color: "1D1D1F",
    margin: 0
  });

  slide.addText(bullets.map((bullet) => `• ${bullet}`).join("\n"), {
    x: 0.78,
    y: 1.45,
    w: hasImage ? 6.2 : 11.6,
    h: 4.7,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    color: "333333",
    fit: "shrink",
    breakLine: false,
    margin: 0.08,
    paraSpaceAfterPt: 8
  });

  if (hasImage) {
    slide.addShape("rect", {
      x: 7.35,
      y: 1.45,
      w: 4.95,
      h: 4.25,
      fill: { color: "F5F5F7" },
      line: { color: "D2D2D7", width: 1.2, dash: "dash" }
    });
    slide.addText(input.selectedImageFileName ?? "Selected image", {
      x: 7.55,
      y: 3.35,
      w: 4.55,
      h: 0.45,
      fontFace: "Microsoft YaHei",
      fontSize: 13,
      color: "666666",
      align: "center",
      margin: 0
    });
  }

  slide.addText("Reflow replacement slide. Insert after the current slide, then compare and keep the better version.", {
    x: 0.72,
    y: 7.05,
    w: 11.6,
    h: 0.25,
    fontFace: "Microsoft YaHei",
    fontSize: 9,
    color: "777777",
    margin: 0
  });

  const output = await pptx.write({ outputType: "base64" });
  return {
    pptxBase64: String(output),
    slideSpec: {
      title,
      bullets,
      selectedImageFileName: input.selectedImageFileName ?? null
    },
    qa: "Reflow MVP generated a replacement slide from page text and instruction. Exact original shape positions are not read yet."
  };
}

function normalizeTitle(instruction: string, pageText: string) {
  const source = instruction.trim() || pageText.trim() || "当前页重排";
  const firstLine = source.split(/\r?\n/).map((line) => line.trim()).find(Boolean) ?? "当前页重排";
  return firstLine.length > 34 ? `${firstLine.slice(0, 34)}...` : firstLine;
}

function buildBullets(pageText: string, instruction: string) {
  const source = pageText.trim() || instruction.trim();
  const candidates = source
    .split(/[\r\n。；;]/)
    .map((item) => item.trim())
    .filter(Boolean)
    .filter((item) => item.length > 4);

  if (candidates.length > 0) {
    return candidates.slice(0, 5).map((item) => (item.length > 56 ? `${item.slice(0, 56)}...` : item));
  }

  return ["保留当前页核心信息", "压缩文字层级", "为图片或图表预留清晰区域"];
}
