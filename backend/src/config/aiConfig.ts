import "dotenv/config";
import { z } from "zod";

const aiEnvSchema = z.object({
  AI_PROVIDER: z.enum(["openai", "custom"]).default("openai"),
  AI_API_KEY: z.string().optional(),
  AI_BASE_URL: z.string().url().default("https://api.openai.com/v1"),
  AI_MODEL: z.string().default("gpt-4.1-mini")
});

const aiEnv = aiEnvSchema.parse(process.env);

export interface AiRuntimeConfig {
  provider: "openai" | "custom";
  apiKey: string | null;
  baseUrl: string;
  model: string;
}

export function getAiRuntimeConfig(): AiRuntimeConfig {
  return {
    provider: aiEnv.AI_PROVIDER,
    apiKey: aiEnv.AI_API_KEY ?? null,
    baseUrl: aiEnv.AI_BASE_URL,
    model: aiEnv.AI_MODEL
  };
}

export function getPublicAiSettings() {
  const config = getAiRuntimeConfig();
  return {
    provider: config.provider,
    configured: Boolean(config.apiKey),
    baseUrlHost: safeHost(config.baseUrl),
    model: config.model
  };
}

export function requireAiRuntimeConfig() {
  const config = getAiRuntimeConfig();

  if (!config.apiKey) {
    throw new Error("AI_API_KEY is not configured. Create backend/.env from backend/.env.example.");
  }

  return config as AiRuntimeConfig & { apiKey: string };
}

function safeHost(url: string) {
  try {
    return new URL(url).host;
  } catch {
    return "invalid-url";
  }
}
