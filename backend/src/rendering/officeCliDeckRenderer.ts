import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import type { DeckPlan, SlideSpec } from "../deckPlan/deckPlanSchema.js";
import type { TemplateProfile } from "../templates/templateProfile.js";
import type { TemplateReplaceableSlot } from "../templates/templateCapabilities.js";
import { runOfficeCli } from "./officeCliRuntime.js";

export interface OfficeCliDeckInput {
  deckPlan: DeckPlan;
  profile: TemplateProfile;
  reportFileName: string;
  templateFileName: string;
  templateSourcePath?: string;
  templateBuffer?: Buffer;
}

export interface OfficeCliDeckResult {
  pptxBase64: string;
  slides: Array<{
    slideId: string;
    selectedSlideIndex: number;
    selectedRole: string;
    replacedSlots: Array<{
      shapeId: string;
      slotType: string;
    }>;
  }>;
  qa: string;
}

export async function renderOfficeCliDeck(input: OfficeCliDeckInput): Promise<OfficeCliDeckResult> {
  const workDir = await mkdtemp(join(tmpdir(), "ppt-builders-officecli-"));
  const outputPath = join(workDir, "generated.pptx");

  try {
    if (input.templateSourcePath) {
      await copyFile(input.templateSourcePath, outputPath);
    } else if (input.templateBuffer) {
      await writeFile(outputPath, input.templateBuffer);
    } else {
      throw new Error("officeCLI 生成缺少模板源文件。");
    }

    const renderedSlides = [];
    for (const slideSpec of input.deckPlan.slides) {
      const selectedSlideIndex = chooseTemplateSlide(input.profile, slideSpec);
      await runOfficeCli(["add", outputPath, "/", "--from", `/slide[${selectedSlideIndex}]`], { cwd: workDir });
      const selectedSlide = input.profile.slides.find((slide) => slide.index === selectedSlideIndex) ?? input.profile.slides[0];
      renderedSlides.push({
        slideSpec,
        selectedSlideIndex,
        selectedRole: selectedSlide?.role ?? "unknown"
      });
    }

    for (let slideIndex = input.profile.counts.slides; slideIndex >= 1; slideIndex -= 1) {
      await runOfficeCli(["remove", outputPath, `/slide[${slideIndex}]`], { cwd: workDir });
    }

    const slideResults: OfficeCliDeckResult["slides"] = [];
    for (let index = 0; index < renderedSlides.length; index += 1) {
      const generatedSlideIndex = index + 1;
      const rendered = renderedSlides[index];
      const replacedSlots = await replaceSlideTextSlots({
        outputPath,
        generatedSlideIndex,
        profile: input.profile,
        slideSpec: rendered.slideSpec
      });

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
      pptxBase64,
      slides: slideResults,
      qa
    };
  } finally {
    await rm(workDir, { recursive: true, force: true });
  }
}

async function replaceSlideTextSlots(input: {
  outputPath: string;
  generatedSlideIndex: number;
  profile: TemplateProfile;
  slideSpec: SlideSpec;
}) {
  const slots = input.profile.capabilities.replaceableSlots.filter(
    (slot) => slot.slideIndex === chooseTemplateSlide(input.profile, input.slideSpec)
  );
  const titleSlot = selectSlot(slots, "title");
  const bodySlots = slots.filter((slot) => slot.slotType === "body").sort((a, b) => b.confidence - a.confidence).slice(0, 3);
  const replacedSlots: Array<{ shapeId: string; slotType: string }> = [];

  if (titleSlot) {
    const replaced = await setShapeText(input.outputPath, input.generatedSlideIndex, titleSlot.shapeId, input.slideSpec.title);
    if (replaced) {
      replacedSlots.push({ shapeId: titleSlot.shapeId, slotType: titleSlot.slotType });
    }
  }

  const bodyTexts = input.slideSpec.blocks
    .filter((block) => block.type === "body" || block.type === "callout")
    .map((block) => block.text);

  if (bodySlots.length > 0) {
    for (let index = 0; index < bodySlots.length; index += 1) {
      const slot = bodySlots[index];
      const text = bodyTexts[index] ?? bodyTexts[bodyTexts.length - 1];
      if (!text) {
        continue;
      }
      const replaced = await setShapeText(input.outputPath, input.generatedSlideIndex, slot.shapeId, text);
      if (replaced) {
        replacedSlots.push({ shapeId: slot.shapeId, slotType: slot.slotType });
      }
    }
  } else if (bodyTexts.length > 0) {
    const fallbackSlot = firstBodyLikeTextBox(input.profile, input.slideSpec);
    if (fallbackSlot) {
      const replaced = await setShapeText(input.outputPath, input.generatedSlideIndex, fallbackSlot.id, bodyTexts.join("\n"));
      if (replaced) {
        replacedSlots.push({ shapeId: fallbackSlot.id, slotType: "body" });
      }
    }
  }

  return replacedSlots;
}

async function setShapeText(outputPath: string, slideIndex: number, shapeId: string, text: string) {
  if (!shapeId || !text.trim()) {
    return false;
  }

  try {
    await runOfficeCli(["set", outputPath, `/slide[${slideIndex}]/shape[@id=${shapeId}]`, "--prop", `text=${text}`]);
    return true;
  } catch {
    return false;
  }
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

function chooseTemplateSlide(profile: TemplateProfile, slideSpec: SlideSpec) {
  if (slideSpec.templateIntent.preferredSlideIndex) {
    return slideSpec.templateIntent.preferredSlideIndex;
  }

  const preferredForRole = profile.capabilities.recommendedSlides[slideSpec.templateIntent.preferredRole][0];
  if (preferredForRole) {
    return preferredForRole;
  }

  return profile.capabilities.recommendedSlides.content_text[0] ?? profile.slides[0]?.index ?? 1;
}

function selectSlot(slots: TemplateReplaceableSlot[], slotType: TemplateReplaceableSlot["slotType"]) {
  return slots
    .filter((slot) => slot.slotType === slotType)
    .sort((a, b) => b.confidence - a.confidence)[0];
}

function firstBodyLikeTextBox(profile: TemplateProfile, slideSpec: SlideSpec) {
  const selectedSlideIndex = chooseTemplateSlide(profile, slideSpec);
  const slide = profile.slides.find((candidate) => candidate.index === selectedSlideIndex);
  return slide?.textBoxes
    .filter((box) => box.text && !box.placeholderType?.includes("title"))
    .sort((a, b) => (b.bbox?.cy ?? 0) - (a.bbox?.cy ?? 0))[0];
}
