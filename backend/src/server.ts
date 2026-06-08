import Fastify from "fastify";
import type { FastifyRequest } from "fastify";
import cors from "@fastify/cors";
import multipart from "@fastify/multipart";
import { z } from "zod";
import {
  deleteAsset,
  listAssets,
  readAssetBase64,
  saveAsset,
  type AssetKind
} from "./assets/assetLibraryStore.js";
import { getPublicAiSettings } from "./config/aiConfig.js";
import { buildReportDeckPlan } from "./deckPlan/deckPlanSchema.js";
import { buildImageSelectionPlan, buildSelectedImageEditPlan } from "./edits/imageEditPlanner.js";
import { renderReflowSlide } from "./rendering/reflowSlideRenderer.js";
import { buildSelectedTextEditPlan } from "./edits/textEditPlanner.js";
import { autofixPageQa, checkPageQa } from "./qa/qaValidator.js";
import { renderTemplateReplacementDeckPlan } from "./rendering/templateSlideRenderer.js";
import {
  getCurrentReport,
  saveCurrentReport,
  summarizeCurrentReport
} from "./reports/currentReportStore.js";
import {
  analyzeTemplateBuffer,
  getDefaultTemplate,
  saveDefaultTemplate
} from "./templates/defaultTemplateStore.js";

const envSchema = z.object({
  PORT: z.coerce.number().int().positive().default(3000),
  HOST: z.string().default("127.0.0.1"),
  FRONTEND_ORIGIN: z.string().default("https://localhost:5173")
});

const env = envSchema.parse(process.env);

const app = Fastify({
  logger: true
});

await app.register(cors, {
  origin: [env.FRONTEND_ORIGIN, "http://localhost:5173"],
  methods: ["GET", "POST", "DELETE"]
});

await app.register(multipart, {
  limits: {
    fileSize: 80 * 1024 * 1024,
    files: 50,
    fields: 8
  }
});

app.get("/health", async () => {
  return {
    ok: true,
    service: "ai-ppt-plugin-backend",
    ai: getPublicAiSettings()
  };
});

app.get("/api/settings/ai", async () => {
  return getPublicAiSettings();
});

app.get("/api/templates/default", async () => {
  const record = await getDefaultTemplate();
  return summarizeDefaultTemplate(record);
});

app.get("/api/reports/current", async () => {
  return summarizeCurrentReport(await getCurrentReport());
});

app.post("/api/reports/current", async (request, reply) => {
  const received = await readGenerationRequest(request);
  if (!received.report) {
    return reply.code(400).send({
      error: "report file is required"
    });
  }

  const currentReport = await saveCurrentReport({
    fileName: received.report.fileName,
    buffer: received.report.buffer
  });

  return summarizeCurrentReport(currentReport);
});

app.post("/api/templates/default", async (request, reply) => {
  const parts = request.parts();
  let templateFile: UploadedFile | null = null;

  for await (const part of parts) {
    if (part.type === "file" && part.fieldname === "template") {
      templateFile = {
        fieldName: part.fieldname,
        fileName: part.filename,
        mimeType: part.mimetype,
        size: 0,
        buffer: await part.toBuffer()
      };
      templateFile.size = templateFile.buffer.length;
    }
  }

  if (!templateFile) {
    return reply.code(400).send({
      error: "template file is required"
    });
  }

  const record = await saveDefaultTemplate({
    fileName: templateFile.fileName,
    buffer: templateFile.buffer
  });

  return summarizeDefaultTemplate(record);
});

app.post("/api/templates/analyze", async (request, reply) => {
  const parts = request.parts();
  let templateFile: UploadedFile | null = null;

  for await (const part of parts) {
    if (part.type === "file" && part.fieldname === "template") {
      const buffer = await part.toBuffer();
      templateFile = {
        fieldName: part.fieldname,
        fileName: part.filename,
        mimeType: part.mimetype,
        size: buffer.length,
        buffer
      };
    }
  }

  if (!templateFile) {
    return reply.code(400).send({
      error: "template file is required"
    });
  }

  const profile = await analyzeTemplateBuffer({
    fileName: templateFile.fileName,
    buffer: templateFile.buffer
  });

  return {
    profile
  };
});

app.get("/api/assets", async (request) => {
  const query = z
    .object({
      kind: z.enum(["image", "table"]).optional()
    })
    .parse(request.query);

  return {
    assets: await listAssets(query.kind)
  };
});

app.post("/api/assets", async (request, reply) => {
  const parts = request.parts();
  let kind: AssetKind = "image";
  let notes = "";
  const saved = [];

  for await (const part of parts) {
    if (part.type === "file") {
      const buffer = await part.toBuffer();
      saved.push(
        await saveAsset({
          kind,
          sourceFileName: part.filename,
          mimeType: part.mimetype,
          buffer,
          notes
        })
      );
    } else if (part.fieldname === "kind") {
      const value = String(part.value ?? "image");
      kind = value === "table" ? "table" : "image";
    } else if (part.fieldname === "notes") {
      notes = String(part.value ?? "");
    }
  }

  if (saved.length === 0) {
    return reply.code(400).send({
      error: "at least one asset file is required"
    });
  }

  return {
    assets: saved
  };
});

app.get("/api/assets/:assetId/base64", async (request, reply) => {
  const params = z.object({ assetId: z.string() }).parse(request.params);
  const result = await readAssetBase64(params.assetId);
  if (!result) {
    return reply.code(404).send({
      error: "asset not found"
    });
  }

  return result;
});

app.delete("/api/assets/:assetId", async (request, reply) => {
  const params = z.object({ assetId: z.string() }).parse(request.params);
  const deleted = await deleteAsset(params.assetId);
  if (!deleted) {
    return reply.code(404).send({
      error: "asset not found"
    });
  }

  return {
    ok: true
  };
});

app.post("/api/edits/selection/text", async (request) => {
  const body = z
    .object({
      instruction: z.string().default(""),
      selectedText: z.string().default(""),
      context: z.string().optional()
    })
    .parse(request.body);

  return {
    editPlan: await buildSelectedTextEditPlan(body)
  };
});

app.post("/api/edits/selection/image", async (request) => {
  const body = z
    .object({
      instruction: z.string().default(""),
      imageFileName: z.string().optional(),
      imageMimeType: z.string().optional()
    })
    .parse(request.body);

  return {
    editPlan: buildSelectedImageEditPlan(body)
  };
});

app.post("/api/edits/selection/image/select", async (request) => {
  const body = z
    .object({
      instruction: z.string().default(""),
      pageText: z.string().default(""),
      candidates: z
        .array(
          z.object({
            id: z.string(),
            fileName: z.string(),
            mimeType: z.string().optional(),
            notes: z.string().optional()
          })
        )
        .default([])
    })
    .parse(request.body);

  return {
    editPlan: buildImageSelectionPlan(body)
  };
});

app.post("/api/edits/slide/reflow", async (request) => {
  const body = z
    .object({
      instruction: z.string().default(""),
      pageText: z.string().default(""),
      selectedImageFileName: z.string().nullable().optional()
    })
    .parse(request.body);

  return renderReflowSlide(body);
});

app.post("/api/qa/check", async (request) => {
  const body = z
    .object({
      pageText: z.string().default(""),
      instruction: z.string().optional()
    })
    .parse(request.body);

  return checkPageQa(body);
});

app.post("/api/qa/autofix", async (request) => {
  const body = z
    .object({
      pageText: z.string().default(""),
      instruction: z.string().optional()
    })
    .parse(request.body);

  return autofixPageQa(body);
});

app.post("/api/decks/plan", async (request, reply) => {
  const received = await readGenerationRequest(request);
  const currentReport = await resolveReport(received.report);

  if (!currentReport) {
    return reply.code(400).send({
      error: "report file is required when no current report is saved"
    });
  }

  const { template, templateProfile, defaultTemplateUsed } = await resolveTemplate(received.template);
  const deckPlan = buildReportDeckPlan({
    reportFileName: currentReport.sourceFileName,
    templateFileName: template.fileName,
    instruction: received.instruction,
    templateProfile,
    evidenceIndex: currentReport.evidenceIndex
  });

  return {
    deckPlan,
    received: {
      report: summarizeCurrentReport(currentReport),
      template: toUploadedFileSummary(template),
      currentReportUsed: !received.report,
      defaultTemplateUsed
    }
  };
});

app.post("/api/decks/generate", async (request, reply) => {
  const received = await readGenerationRequest(request);
  const currentReport = await resolveReport(received.report);

  if (!currentReport) {
    return reply.code(400).send({
      error: "report file is required when no current report is saved"
    });
  }

  const { template, templateProfile, defaultTemplateUsed } = await resolveTemplate(received.template);
  const deckPlan = buildReportDeckPlan({
    reportFileName: currentReport.sourceFileName,
    templateFileName: template.fileName,
    instruction: received.instruction,
    templateProfile,
    evidenceIndex: currentReport.evidenceIndex
  });

  const rendered = await renderTemplateReplacementDeckPlan({
    profile: templateProfile,
    reportFileName: currentReport.sourceFileName,
    templateFileName: template.fileName,
    instruction: received.instruction,
    slideSpec: deckPlan.slides[0],
    slideSpecs: deckPlan.slides
  });

  return reply.send({
    deckId: `deck_${Date.now()}`,
    pptxBase64: rendered.pptxBase64,
    summary: `Generated a ${deckPlan.slides.length}-slide deck from report evidence using template replacement.`,
    qa: deckPlan.validation.warnings.length > 0
      ? deckPlan.validation.warnings.join(" ")
      : "Generated from parsed report evidence. Template styling is approximated; original PPTX slide backgrounds are not copied yet.",
    received: {
      report: summarizeCurrentReport(currentReport),
      template: toUploadedFileSummary(template),
      currentReportUsed: !received.report,
      defaultTemplateUsed
    },
    deckPlan,
    templateReplacement: {
      selectedSlideIndex: rendered.slides[0]?.selectedSlideIndex ?? 0,
      selectedRole: rendered.slides[0]?.selectedRole ?? "unknown",
      replacedSlots: rendered.slides.flatMap((slide) => slide.replacedSlots),
      slides: rendered.slides
    }
  });
});

try {
  await app.listen({
    host: env.HOST,
    port: env.PORT
  });
} catch (error) {
  app.log.error(error);
  process.exit(1);
}

interface UploadedFileSummary {
  fieldName: string;
  fileName: string;
  mimeType: string;
  size: number;
}

interface UploadedFile extends UploadedFileSummary {
  buffer: Buffer;
}

async function readGenerationRequest(request: FastifyRequest) {
  const parts = request.parts();
  const received = {
    report: null as UploadedFile | null,
    template: null as UploadedFile | null,
    instruction: ""
  };

  for await (const part of parts) {
    if (part.type === "file") {
      const bytes = await part.toBuffer();
      const summary = {
        fieldName: part.fieldname,
        fileName: part.filename,
        mimeType: part.mimetype,
        size: bytes.length,
        buffer: bytes
      };

      if (part.fieldname === "report") {
        received.report = summary;
      }

      if (part.fieldname === "template") {
        received.template = summary;
      }
    } else if (part.fieldname === "instruction") {
      received.instruction = String(part.value ?? "");
    }
  }

  return received;
}

async function resolveTemplate(uploadedTemplate: UploadedFile | null) {
  if (uploadedTemplate) {
    return {
      template: uploadedTemplate,
      templateProfile: await analyzeTemplateBuffer({
        fileName: uploadedTemplate.fileName,
        buffer: uploadedTemplate.buffer
      }),
      defaultTemplateUsed: false
    };
  }

  const defaultTemplate = await getDefaultTemplate();
  return {
    template: {
      fieldName: "template",
      fileName: defaultTemplate.sourceFileName,
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      size: defaultTemplate.profile.media.reduce((sum, media) => sum + media.size, 0),
      buffer: Buffer.alloc(0)
    },
    templateProfile: defaultTemplate.profile,
    defaultTemplateUsed: true
  };
}

async function resolveReport(uploadedReport: UploadedFile | null) {
  if (uploadedReport) {
    return saveCurrentReport({
      fileName: uploadedReport.fileName,
      buffer: uploadedReport.buffer
    });
  }

  return getCurrentReport();
}

function toUploadedFileSummary(file: UploadedFile): UploadedFileSummary {
  return {
    fieldName: file.fieldName,
    fileName: file.fileName,
    mimeType: file.mimeType,
    size: file.size
  };
}

function summarizeDefaultTemplate(record: Awaited<ReturnType<typeof getDefaultTemplate>>) {
  return {
    templateId: record.templateId,
    sourceFileName: record.sourceFileName,
    profilePath: record.profilePath,
    counts: record.profile.counts,
    roles: record.profile.slides.reduce<Record<string, number>>((acc, slide) => {
      acc[slide.role] = (acc[slide.role] ?? 0) + 1;
      return acc;
    }, {}),
    capabilities: {
      replaceableSlots: record.profile.capabilities.replaceableSlots.length,
      recommendedSlides: record.profile.capabilities.recommendedSlides,
      styleTokens: record.profile.capabilities.styleTokens
    },
    profile: record.profile
  };
}
