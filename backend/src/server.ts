import Fastify from "fastify";
import cors from "@fastify/cors";
import multipart from "@fastify/multipart";
import { createRequire } from "node:module";
import { z } from "zod";
import {
  analyzeTemplateBuffer,
  getDefaultTemplate,
  saveDefaultTemplate
} from "./templates/defaultTemplateStore.js";

const require = createRequire(import.meta.url);
const PptxGenerator = require("pptxgenjs") as { new (): any };

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
    service: "ai-ppt-plugin-backend"
  };
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
  const template = received.template ?? {
    fieldName: "template",
    fileName: defaultTemplate!.sourceFileName,
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    size: defaultTemplate!.profile.media.reduce((sum, media) => sum + media.size, 0),
    buffer: Buffer.alloc(0)
  };

  const generationInput = {
    report: received.report,
    template,
    instruction: received.instruction
  };
  const pptxBase64 = await createSmokeTestDeck(generationInput);

  return reply.send({
    deckId: `deck_${Date.now()}`,
    pptxBase64,
    summary: "Generated a one-slide PPTX smoke test deck.",
    qa: "Smoke test only: report parsing, template page replacement, and content QA are not implemented yet.",
    received: {
      report: toUploadedFileSummary(received.report),
      template: toUploadedFileSummary(template),
      defaultTemplateUsed: !received.template
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

async function createSmokeTestDeck(input: {
  report: UploadedFileSummary;
  template: UploadedFileSummary;
  instruction: string;
}) {
  const pptx = new PptxGenerator();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "AI PPT Builder";
  pptx.subject = "Smoke test deck";
  pptx.title = "AI PPT Builder Smoke Test";
  pptx.company = "ppt-builders";
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
    lang: "zh-CN"
  };

  const slide = pptx.addSlide();
  slide.background = { color: "F7F8FA" };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.18,
    fill: { color: "2F5D8C" },
    line: { color: "2F5D8C" }
  });
  slide.addText("AI PPT Builder", {
    x: 0.6,
    y: 0.55,
    w: 5.5,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: "1F2933",
    margin: 0
  });
  slide.addText("PPTX generation smoke test", {
    x: 0.62,
    y: 1.08,
    w: 7.5,
    h: 0.3,
    fontFace: "Microsoft YaHei",
    fontSize: 12,
    color: "5F6B7A",
    margin: 0
  });
  slide.addText(
    [
      { text: "Report: ", options: { bold: true } },
      { text: `${input.report.fileName} (${formatBytes(input.report.size)})\n` },
      { text: "Template: ", options: { bold: true } },
      { text: `${input.template.fileName} (${formatBytes(input.template.size)})\n` },
      { text: "Instruction: ", options: { bold: true } },
      { text: input.instruction.trim() || "No instruction provided." }
    ],
    {
      x: 0.75,
      y: 1.8,
      w: 11.8,
      h: 2.1,
      fontFace: "Microsoft YaHei",
      fontSize: 15,
      breakLine: false,
      color: "25313D",
      fit: "shrink",
      valign: "mid",
      fill: { color: "FFFFFF" },
      line: { color: "D8DEE6", width: 1 },
      margin: 0.18
    }
  );
  slide.addText("Next step: replace this smoke-test slide with template-based rendering.", {
    x: 0.75,
    y: 5.95,
    w: 11.8,
    h: 0.35,
    fontFace: "Microsoft YaHei",
    fontSize: 11,
    color: "667085",
    margin: 0
  });

  const output = await pptx.write({ outputType: "base64" });
  return String(output);
}

function formatBytes(bytes: number) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  const kb = bytes / 1024;
  if (kb < 1024) {
    return `${kb.toFixed(1)} KB`;
  }

  return `${(kb / 1024).toFixed(1)} MB`;
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
    profile: record.profile
  };
}
