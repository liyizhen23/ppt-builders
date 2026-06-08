import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { tmpdir } from "node:os";
import { listAssets } from "../assets/assetLibraryStore.js";
import type { EvidenceBlock, EvidenceIndex } from "../reports/reportParser.js";
import type { TemplateProfile, TemplateSlideRole } from "../templates/templateProfile.js";
import type { TemplateReplaceableSlot } from "../templates/templateCapabilities.js";
import { runOfficeCli } from "../rendering/officeCliRuntime.js";

export const officeCliGenerationRules = [
  "Only use report evidence and explicit user instructions; do not invent claims.",
  "Strictly preserve the uploaded template page layouts by cloning template slides with officeCLI.",
  "Replace only content slots inferred from the template; protect logos, school emblems, page numbers, footers, and decorative shapes.",
  "Prefer replacing template content pictures with uploaded image assets when a safe image slot is available.",
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

type ImageAsset = Awaited<ReturnType<typeof listAssets>>[number];

export async function generatePptWithOfficeCli(input: OfficeCliPptGenerationInput): Promise<OfficeCliPptGenerationResult> {
  const deckSpec = buildOfficeCliDeckSpec(input);
  const workDir = await mkdtemp(join(tmpdir(), "ppt-builders-officecli-template-"));
  const outputPath = join(workDir, "generated.pptx");

  try {
    if (input.templateSourcePath) {
      await copyFile(input.templateSourcePath, outputPath);
    } else if (input.templateBuffer) {
      await writeFile(outputPath, input.templateBuffer);
    } else {
      throw new Error("officeCLI 生成缺少模板源文件。");
    }

    const imageAssets = await listAssets("image");
    const renderedSlides = [];
    for (const slideSpec of deckSpec.slides) {
      const selectedSlideIndex = chooseTemplateSlide(input.templateProfile, slideSpec);
      const selectedSlide = input.templateProfile.slides.find((slide) => slide.index === selectedSlideIndex);
      await runOfficeCli(["add", outputPath, "/", "--from", `/slide[${selectedSlideIndex}]`], { cwd: workDir });
      renderedSlides.push({
        slideSpec,
        selectedSlideIndex,
        selectedRole: selectedSlide?.role ?? "unknown"
      });
    }

    for (let slideIndex = input.templateProfile.counts.slides; slideIndex >= 1; slideIndex -= 1) {
      await runOfficeCli(["remove", outputPath, `/slide[${slideIndex}]`], { cwd: workDir });
    }

    const slideResults = [];
    for (let index = 0; index < renderedSlides.length; index += 1) {
      const generatedSlideIndex = index + 1;
      const rendered = renderedSlides[index];
      const replacedSlots = await replaceTemplateSlots({
        outputPath,
        generatedSlideIndex,
        profile: input.templateProfile,
        selectedSlideIndex: rendered.selectedSlideIndex,
        slideSpec: rendered.slideSpec,
        imageAssets
      });

      await addNotes(outputPath, generatedSlideIndex, buildSpeakerNotes(rendered.slideSpec));
      slideResults.push({
        slideId: rendered.slideSpec.slideId,
        selectedSlideIndex: rendered.selectedSlideIndex,
        selectedRole: rendered.selectedRole,
        replacedSlots
      });
    }

    const qa = await collectOfficeCliQa(outputPath, workDir);
    const pptxBase64 = (await readFile(outputPath)).toString("base64");
    await runOfficeCli(["close", outputPath], { cwd: workDir, timeoutMs: 30_000 }).catch(() => undefined);

    return {
      deckId: deckSpec.deckId,
      pptxBase64,
      summary: `Generated a ${deckSpec.slides.length}-slide deck by cloning template layouts with officeCLI and replacing safe content slots.`,
      qa,
      deckSpec,
      templateReplacement: {
        selectedSlideIndex: slideResults[0]?.selectedSlideIndex ?? 0,
        selectedRole: slideResults[0]?.selectedRole ?? "unknown",
        replacedSlots: slideResults.flatMap((slide) => slide.replacedSlots),
        slides: slideResults
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

async function replaceTemplateSlots(input: {
  outputPath: string;
  generatedSlideIndex: number;
  profile: TemplateProfile;
  selectedSlideIndex: number;
  slideSpec: OfficeCliSlideSpec;
  imageAssets: ImageAsset[];
}) {
  const allSlots = input.profile.capabilities.replaceableSlots.filter((slot) => slot.slideIndex === input.selectedSlideIndex);
  const safeSlots = allSlots.filter((slot) => !isProtectedSlot(input.profile, slot));
  const replacedSlots: Array<{ shapeId: string; slotType: string }> = [];

  const titleSlot = selectSlot(safeSlots, "title") ?? selectSlot(safeSlots, "subtitle");
  if (titleSlot && (await setShapeText(input.outputPath, input.generatedSlideIndex, titleSlot.shapeId, input.slideSpec.title))) {
    replacedSlots.push({ shapeId: titleSlot.shapeId, slotType: titleSlot.slotType });
  }

  const bodySlots = safeSlots
    .filter((slot) => slot.slotType === "body" || slot.slotType === "caption" || slot.slotType === "subtitle")
    .filter((slot) => slot.shapeId !== titleSlot?.shapeId)
    .sort((a, b) => b.confidence - a.confidence)
    .slice(0, Math.max(1, Math.min(input.slideSpec.bullets.length, 4)));
  const bodyTexts = input.slideSpec.kind === "agenda" ? input.slideSpec.bullets : normalizeBullets(input.slideSpec.bullets, bodySlots.length || 1);

  for (let index = 0; index < bodySlots.length; index += 1) {
    const slot = bodySlots[index];
    const text = bodyTexts[index] ?? bodyTexts[bodyTexts.length - 1];
    if (text && (await setShapeText(input.outputPath, input.generatedSlideIndex, slot.shapeId, text))) {
      replacedSlots.push({ shapeId: slot.shapeId, slotType: slot.slotType });
    }
  }

  const imageSlot = safeSlots
    .filter((slot) => slot.slotType === "image")
    .sort((a, b) => b.confidence - a.confidence)[0];
  const imageAsset = chooseImageAsset(input.imageAssets, input.slideSpec, input.generatedSlideIndex);
  if (imageSlot && imageAsset) {
    const imagePath = resolve(process.cwd(), "..", "asset-library", "files", imageAsset.storedFileName);
    if (await setPictureSource(input.outputPath, input.generatedSlideIndex, imageSlot.shapeId, imagePath, imageAsset.sourceFileName)) {
      replacedSlots.push({ shapeId: imageSlot.shapeId, slotType: "image" });
    }
  }

  return replacedSlots;
}

async function setShapeText(outputPath: string, slideIndex: number, shapeId: string, text: string) {
  if (!shapeId || !text.trim()) {
    return false;
  }

  try {
    await runOfficeCli([
      "set",
      outputPath,
      `/slide[${slideIndex}]/shape[@id=${shapeId}]`,
      "--prop",
      `text=${text}`,
      "--prop",
      "autoFit=normal"
    ]);
    return true;
  } catch {
    return false;
  }
}

async function setPictureSource(outputPath: string, slideIndex: number, pictureId: string, imagePath: string, alt: string) {
  if (!pictureId || !imagePath) {
    return false;
  }

  try {
    await runOfficeCli([
      "set",
      outputPath,
      `/slide[${slideIndex}]/picture[@id=${pictureId}]`,
      "--prop",
      `src=${imagePath}`,
      "--prop",
      `alt=${alt}`
    ]);
    return true;
  } catch {
    return false;
  }
}

function chooseTemplateSlide(profile: TemplateProfile, slideSpec: OfficeCliSlideSpec) {
  const preferred = profile.capabilities.recommendedSlides[slideSpec.preferredTemplateRole][0];
  if (preferred) {
    return preferred;
  }

  if (slideSpec.kind === "content") {
    return (
      profile.capabilities.recommendedSlides.content_image[0] ??
      profile.capabilities.recommendedSlides.content_text[0] ??
      profile.capabilities.recommendedSlides.unknown[0] ??
      profile.slides[0]?.index ??
      1
    );
  }

  return profile.capabilities.recommendedSlides.content_text[0] ?? profile.slides[0]?.index ?? 1;
}

function selectSlot(slots: TemplateReplaceableSlot[], slotType: TemplateReplaceableSlot["slotType"]) {
  return slots
    .filter((slot) => slot.slotType === slotType)
    .sort((a, b) => b.confidence - a.confidence)[0];
}

function isProtectedSlot(profile: TemplateProfile, slot: TemplateReplaceableSlot) {
  const text = `${slot.shapeName ?? ""} ${slot.sampleText ?? ""}`.toLowerCase();
  if (/logo|徽|校徽|清华|tsinghua|页码|page|footer|页脚|日期|date/.test(text)) {
    return true;
  }
  if (slot.slotType === "image") {
    return isProtectedPictureSlot(profile, slot);
  }
  if (!slot.bbox) {
    return false;
  }
  const pageHeight = profile.pageSize.cy ?? 6858000;
  return slot.bbox.y > pageHeight * 0.88 && slot.slotType !== "body";
}

function isProtectedPictureSlot(profile: TemplateProfile, slot: TemplateReplaceableSlot) {
  if (!slot.bbox) {
    return true;
  }
  const pageWidth = profile.pageSize.cx ?? 12192000;
  const pageHeight = profile.pageSize.cy ?? 6858000;
  const areaRatio = (slot.bbox.cx * slot.bbox.cy) / (pageWidth * pageHeight);
  const rightEdge = slot.bbox.x + slot.bbox.cx;
  const inTopRight = slot.bbox.x > pageWidth * 0.72 && slot.bbox.y < pageHeight * 0.18;
  const nearTopRight = rightEdge > pageWidth * 0.82 && slot.bbox.y < pageHeight * 0.22;
  const likelyLogoScale = areaRatio < 0.05;
  return (inTopRight || nearTopRight) && likelyLogoScale;
}

function chooseImageAsset(assets: ImageAsset[], slideSpec: OfficeCliSlideSpec, slideIndex: number) {
  if (assets.length === 0) {
    return null;
  }
  const terms = tokenize(`${slideSpec.title} ${slideSpec.bullets.join(" ")}`);
  const scored = assets.map((asset, index) => {
    const text = `${asset.sourceFileName} ${asset.notes}`.toLowerCase();
    return {
      asset,
      score: terms.filter((term) => text.includes(term)).length,
      index
    };
  });
  return scored.sort((a, b) => b.score - a.score || a.index - b.index)[(slideIndex - 1) % scored.length]?.asset ?? assets[0];
}

async function addNotes(outputPath: string, slideNumber: number, text: string) {
  await runOfficeCli(["add", outputPath, `/slide[${slideNumber}]`, "--type", "notes", "--prop", `text=${text}`]).catch(async () => {
    await runOfficeCli(["set", outputPath, `/slide[${slideNumber}]`, "--prop", `notes=${text}`]);
  });
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
  const suppressedMessages: string[] = [];

  try {
    const issues = await runOfficeCli(["view", outputPath, "issues", "--limit", "12"], { cwd: workDir, timeoutMs: 60_000 });
    const filtered = filterOfficeCliOutput(`${issues.stdout}${issues.stderr}`, [
      /Picture ".+" is missing alt text \(accessibility issue\)/i,
      /^Found \d+ issue\(s\):$/i,
      /^Format Issues \(\d+\):$/i
    ]);
    if (filtered.visible) {
      messages.push(filtered.visible);
    }
    if (filtered.suppressedCount > 0) {
      suppressedMessages.push(`已隐藏 ${filtered.suppressedCount} 条模板图片 alt text 可访问性提示。`);
    }
  } catch (error) {
    messages.push(error instanceof Error ? error.message : "officeCLI issues check failed.");
  }

  try {
    await runOfficeCli(["validate", outputPath], { cwd: workDir, timeoutMs: 60_000 });
    messages.push("officeCLI validate passed.");
  } catch (error) {
    const rawMessage = error instanceof Error ? error.message : "officeCLI validate reported issues.";
    const filtered = filterOfficeCliOutput(rawMessage, [
      /The 'mod' attribute is not declared\./i,
      /Path: \/p:sldLayout\[1\]\/p:extLst\[1\]/i,
      /Part: \/ppt\/slideLayouts\/slideLayout\d+\.xml/i,
      /^Found \d+ validation error\(s\):$/i
    ]);
    if (filtered.visible) {
      messages.push(`officeCLI validate reported issues: ${filtered.visible}`);
    }
    if (filtered.suppressedCount > 0) {
      suppressedMessages.push(`已隐藏 ${filtered.suppressedCount} 条模板 slideLayout 扩展属性 schema 提示。`);
    }
  }

  const visible = messages.filter(Boolean).join("\n").trim();
  const suppressed = suppressedMessages.join(" ");
  if (visible) {
    return [visible, suppressed].filter(Boolean).join("\n");
  }
  return suppressed ? `officeCLI 生成完成。${suppressed}` : "officeCLI 生成完成，未发现阻断性 QA 问题。";
}

function filterOfficeCliOutput(output: string, suppressPatterns: RegExp[]) {
  let suppressedCount = 0;
  const visibleLines = output
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => {
      if (!line) {
        return false;
      }
      if (/^Command failed:/i.test(line)) {
        suppressedCount += 1;
        return false;
      }
      const suppressed = suppressPatterns.some((pattern) => pattern.test(line));
      if (suppressed) {
        suppressedCount += 1;
        return false;
      }
      return true;
    });

  return {
    visible: visibleLines.join("\n"),
    suppressedCount
  };
}
