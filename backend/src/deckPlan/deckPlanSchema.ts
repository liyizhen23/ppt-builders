import { z } from "zod";
import { EvidenceIndex, findRelevantEvidence } from "../reports/reportParser.js";
import type { TemplateProfile, TemplateSlideRole } from "../templates/templateProfile.js";

export const visualSpecSchema = z.object({
  type: z.enum(["none", "image", "chart", "table", "diagram"]).default("none"),
  intent: z.string().optional(),
  sourceEvidenceId: z.string().optional()
});

export const slideBlockSchema = z.object({
  type: z.enum(["title", "subtitle", "body", "caption", "callout"]),
  text: z.string().min(1),
  sourceEvidenceId: z.string().optional()
});

export const slideSpecSchema = z.object({
  slideId: z.string().min(1),
  role: z.enum([
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
  ]),
  title: z.string().min(1),
  blocks: z.array(slideBlockSchema).min(1),
  visual: visualSpecSchema.default({ type: "none" }),
  templateIntent: z.object({
    preferredRole: z.enum([
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
    ]),
    preferredSlideIndex: z.number().int().positive().optional(),
    requiredSlots: z.array(z.enum(["title", "subtitle", "body", "caption", "image"])).default(["title", "body"])
  }),
  sourceEvidenceIds: z.array(z.string()).default([])
});

export const deckPlanSchema = z.object({
  deckId: z.string().min(1),
  version: z.literal("0.1"),
  title: z.string().min(1),
  source: z.object({
    reportFileName: z.string().min(1),
    templateFileName: z.string().min(1),
    instruction: z.string().default("")
  }),
  slides: z.array(slideSpecSchema).min(1),
  constraints: z.object({
    maxSlides: z.number().int().positive().default(1),
    language: z.string().default("zh-CN"),
    requireEvidence: z.boolean().default(false)
  }),
  validation: z.object({
    schemaValid: z.boolean(),
    warnings: z.array(z.string())
  })
});

export type DeckPlan = z.infer<typeof deckPlanSchema>;
export type SlideSpec = z.infer<typeof slideSpecSchema>;
export type SlideBlock = z.infer<typeof slideBlockSchema>;

export interface BuildDeckPlanInput {
  reportFileName: string;
  templateFileName: string;
  instruction: string;
  templateProfile: TemplateProfile;
  evidenceIndex?: EvidenceIndex;
}

export function buildSingleSlideDeckPlan(input: BuildDeckPlanInput): DeckPlan {
  const preferredRole = choosePreferredRole(input.instruction, input.templateProfile);
  const preferredSlideIndex = input.templateProfile.capabilities.recommendedSlides[preferredRole][0];
  const title = normalizeTitle(input.instruction);
  const relevantEvidence = input.evidenceIndex ? findRelevantEvidence(input.evidenceIndex, input.instruction, 3) : [];
  const fallbackEvidenceId = `report:${input.reportFileName}:placeholder`;
  const evidenceIds = relevantEvidence.length > 0 ? relevantEvidence.map((block) => block.evidenceId) : [fallbackEvidenceId];
  const bodyBlocks = relevantEvidence.length > 0 ? evidenceToBlocks(relevantEvidence) : fallbackBlocks(input, fallbackEvidenceId);

  const candidate = {
    deckId: `deck_${Date.now()}`,
    version: "0.1",
    title,
    source: {
      reportFileName: input.reportFileName,
      templateFileName: input.templateFileName,
      instruction: input.instruction
    },
    slides: [
      {
        slideId: "slide_1",
        role: preferredRole,
        title,
        blocks: [
          {
            type: "title",
            text: title,
            sourceEvidenceId: evidenceIds[0]
          },
          ...bodyBlocks
        ],
        visual: inferVisual(input.instruction),
        templateIntent: {
          preferredRole,
          preferredSlideIndex,
          requiredSlots: inferRequiredSlots(input.instruction)
        },
        sourceEvidenceIds: evidenceIds
      }
    ],
    constraints: {
      maxSlides: 1,
      language: "zh-CN",
      requireEvidence: false
    },
    validation: {
      schemaValid: true,
      warnings: []
    }
  };

  const parsed = deckPlanSchema.safeParse(candidate);
  if (!parsed.success) {
    return deckPlanSchema.parse({
      ...candidate,
      validation: {
        schemaValid: false,
        warnings: parsed.error.issues.map((issue) => `${issue.path.join(".")}: ${issue.message}`)
      }
    });
  }

  return parsed.data;
}

export function buildReportDeckPlan(input: BuildDeckPlanInput): DeckPlan {
  const contentRole = choosePreferredRole(input.instruction, input.templateProfile);
  const evidence = input.evidenceIndex ? findRelevantEvidence(input.evidenceIndex, input.instruction, 18) : [];
  const title = normalizeTitle(input.instruction);
  const warnings: string[] = [];

  if (!input.evidenceIndex || input.evidenceIndex.blocks.length === 0) {
    warnings.push("No report evidence was available; generated a fallback single-slide deck.");
    return buildSingleSlideDeckPlan(input);
  }

  const contentSlides = chunkEvidence(evidence.length > 0 ? evidence : input.evidenceIndex.blocks.slice(0, 18), 3)
    .slice(0, 5)
    .map((chunk, index) => {
      const slideTitle = index === 0 ? title : deriveSlideTitle(chunk[0]?.text, index + 1);
      const evidenceIds = chunk.map((block) => block.evidenceId);
      return {
        slideId: `slide_${index + 1}`,
        role: contentRole,
        title: slideTitle,
        blocks: [
          {
            type: "title" as const,
            text: slideTitle,
            sourceEvidenceId: evidenceIds[0]
          },
          ...evidenceToBlocks(chunk)
        ],
        visual: inferVisual(input.instruction),
        templateIntent: {
          preferredRole: contentRole,
          preferredSlideIndex: pickRecommendedSlide(input.templateProfile, contentRole, index),
          requiredSlots: inferRequiredSlots(input.instruction)
        },
        sourceEvidenceIds: evidenceIds
      };
    });

  const candidate = {
    deckId: `deck_${Date.now()}`,
    version: "0.1",
    title,
    source: {
      reportFileName: input.reportFileName,
      templateFileName: input.templateFileName,
      instruction: input.instruction
    },
    slides: contentSlides,
    constraints: {
      maxSlides: contentSlides.length,
      language: "zh-CN",
      requireEvidence: true
    },
    validation: {
      schemaValid: true,
      warnings
    }
  };

  const parsed = deckPlanSchema.safeParse(candidate);
  if (!parsed.success) {
    return deckPlanSchema.parse({
      ...candidate,
      validation: {
        schemaValid: false,
        warnings: parsed.error.issues.map((issue) => `${issue.path.join(".")}: ${issue.message}`)
      }
    });
  }

  return parsed.data;
}

function evidenceToBlocks(evidence: ReturnType<typeof findRelevantEvidence>) {
  return evidence.map((block) => ({
    type: "body" as const,
    text: compactEvidenceText(block.text),
    sourceEvidenceId: block.evidenceId
  }));
}

function fallbackBlocks(input: BuildDeckPlanInput, evidenceId: string) {
  return [
    {
      type: "body" as const,
      text: `报告文件：${input.reportFileName}`,
      sourceEvidenceId: evidenceId
    },
    {
      type: "body" as const,
      text: `生成要求：${input.instruction.trim() || "未提供具体生成要求"}`,
      sourceEvidenceId: evidenceId
    },
    {
      type: "body" as const,
      text: "已根据 DeckPlan schema 和 Template Profile 选择模板页并填充槽位。",
      sourceEvidenceId: evidenceId
    }
  ];
}

function compactEvidenceText(text: string) {
  const cleaned = text.replace(/\s+/g, " ").trim();
  return cleaned.length > 90 ? `${cleaned.slice(0, 90)}...` : cleaned;
}

function chunkEvidence<T>(items: T[], size: number) {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
}

function deriveSlideTitle(text: string | undefined, slideNumber: number) {
  const cleaned = (text ?? "").replace(/\s+/g, " ").trim();
  if (!cleaned) {
    return `Report insight ${slideNumber}`;
  }
  return cleaned.length > 28 ? `${cleaned.slice(0, 28)}...` : cleaned;
}

function pickRecommendedSlide(profile: TemplateProfile, role: TemplateSlideRole, offset: number) {
  const slides = profile.capabilities.recommendedSlides[role];
  if (slides.length === 0) {
    return undefined;
  }
  return slides[offset % slides.length];
}

export function validateDeckPlan(input: unknown) {
  return deckPlanSchema.safeParse(input);
}

function choosePreferredRole(instruction: string, profile: TemplateProfile): TemplateSlideRole {
  const text = instruction.toLowerCase();
  const role: TemplateSlideRole = /图|图片|figure|image|流程|方法|diagram/.test(text) ? "content_image" : "content_text";
  if (profile.capabilities.recommendedSlides[role].length > 0) {
    return role;
  }
  return profile.capabilities.recommendedSlides.content_text.length > 0 ? "content_text" : "content_image";
}

function inferVisual(instruction: string) {
  if (/图|图片|figure|image/.test(instruction)) {
    return { type: "image" as const, intent: "Use image slot if available." };
  }
  if (/流程|方法|框架|diagram/.test(instruction)) {
    return { type: "diagram" as const, intent: "Use image/diagram slot if available." };
  }
  return { type: "none" as const };
}

function inferRequiredSlots(instruction: string): Array<"title" | "body" | "image"> {
  if (/图|图片|figure|image|流程|方法|diagram/.test(instruction)) {
    return ["title", "body", "image"];
  }
  return ["title", "body"];
}

function normalizeTitle(instruction: string) {
  const cleaned = instruction.trim().replace(/\s+/g, " ");
  if (!cleaned) {
    return "基于报告生成的一页 PPT";
  }
  return cleaned.length > 34 ? `${cleaned.slice(0, 34)}...` : cleaned;
}
