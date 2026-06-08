import Fastify from "fastify";
import cors from "@fastify/cors";
import multipart from "@fastify/multipart";
import { z } from "zod";
import { getPublicAiSettings } from "./config/aiConfig.js";
import { renderTemplateReplacementDeck } from "./rendering/templateSlideRenderer.js";
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
  methods: ["GET", "POST"]
});

await app.register(multipart, {
  limits: {
    fileSize: 80 * 1024 * 1024,
    files: 2,
    fields: 4
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

app.post("/api/decks/generate", async (request, reply) => {
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

  if (!received.report) {
    return reply.code(400).send({
      error: "report file is required"
    });
  }

  const defaultTemplate = received.template ? null : await getDefaultTemplate();
  const templateProfile = received.template
    ? await analyzeTemplateBuffer({
        fileName: received.template.fileName,
        buffer: received.template.buffer
      })
    : defaultTemplate!.profile;
  const template = received.template ?? {
    fieldName: "template",
    fileName: defaultTemplate!.sourceFileName,
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    size: defaultTemplate!.profile.media.reduce((sum, media) => sum + media.size, 0),
    buffer: Buffer.alloc(0)
  };

  const rendered = await renderTemplateReplacementDeck({
    profile: templateProfile,
    reportFileName: received.report.fileName,
    templateFileName: template.fileName,
    instruction: received.instruction
  });

  return reply.send({
    deckId: `deck_${Date.now()}`,
    pptxBase64: rendered.pptxBase64,
    summary: `Generated a one-slide template replacement deck using template slide ${rendered.selectedSlideIndex}.`,
    qa: "Template replacement smoke test only: report parsing, copied template backgrounds, and content QA are not implemented yet.",
    received: {
      report: toUploadedFileSummary(received.report),
      template: toUploadedFileSummary(template),
      defaultTemplateUsed: !received.template
    },
    templateReplacement: {
      selectedSlideIndex: rendered.selectedSlideIndex,
      selectedRole: rendered.selectedRole,
      replacedSlots: rendered.replacedSlots
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
