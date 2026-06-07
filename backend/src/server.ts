import Fastify from "fastify";
import cors from "@fastify/cors";
import multipart from "@fastify/multipart";
import { z } from "zod";

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

app.post("/api/decks/generate", async (request, reply) => {
  const parts = request.parts();
  const received = {
    report: null as UploadedFileSummary | null,
    template: null as UploadedFileSummary | null,
    instruction: ""
  };

  for await (const part of parts) {
    if (part.type === "file") {
      const bytes = await part.toBuffer();
      const summary = {
        fieldName: part.fieldname,
        fileName: part.filename,
        mimeType: part.mimetype,
        size: bytes.length
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

  if (!received.report || !received.template) {
    return reply.code(400).send({
      error: "report and template files are required"
    });
  }

  return reply.code(501).send({
    error: "PPTX generation is not implemented yet",
    nextStep: "Implement the PPT read/write minimal loop and return pptxBase64.",
    received
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
