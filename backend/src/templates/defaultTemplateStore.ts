import { mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { basename, join, resolve } from "node:path";
import { analyzePptxTemplate, TemplateProfile } from "./templateProfile.js";

const rootDir = resolve(process.cwd(), "..");
const storageDir = resolve(process.env.TEMPLATE_STORAGE_DIR ?? join(rootDir, "templates", "default"));
const defaultTsinghuaTemplatePath = resolve(
  process.env.DEFAULT_TEMPLATE_PATH ?? join(rootDir, "清华大学2025年度演示文稿系列模板2-通用主题.pptx")
);

export interface DefaultTemplateRecord {
  templateId: string;
  sourceFileName: string;
  sourcePath: string;
  profilePath: string;
  profile: TemplateProfile;
}

export async function getDefaultTemplate(): Promise<DefaultTemplateRecord> {
  const storedSourcePath = join(storageDir, "source.pptx");
  const sourcePath = existsSync(storedSourcePath) ? storedSourcePath : defaultTsinghuaTemplatePath;
  return analyzeAndPersistDefault(sourcePath, basename(sourcePath), existsSync(storedSourcePath));
}

export async function saveDefaultTemplate(input: {
  fileName: string;
  buffer: Buffer;
}): Promise<DefaultTemplateRecord> {
  await mkdir(storageDir, { recursive: true });
  const sourcePath = join(storageDir, "source.pptx");
  await writeFile(sourcePath, input.buffer);
  return analyzeAndPersistDefault(sourcePath, input.fileName, true);
}

export async function analyzeTemplateBuffer(input: {
  fileName: string;
  buffer: Buffer;
  templateId?: string;
}) {
  return analyzePptxTemplate({
    templateId: input.templateId ?? slugify(input.fileName.replace(/\.pptx$/i, "")),
    sourceFileName: input.fileName,
    buffer: input.buffer
  });
}

async function analyzeAndPersistDefault(sourcePath: string, sourceFileName: string, isStored: boolean) {
  await mkdir(storageDir, { recursive: true });
  const profilePath = join(storageDir, "template-profile.json");
  const buffer = await readFile(sourcePath);
  const profile = await analyzePptxTemplate({
    templateId: isStored ? "user-default-template" : "tsinghua-2025-general-2",
    sourceFileName,
    buffer
  });
  await writeFile(profilePath, JSON.stringify(profile, null, 2), "utf8");

  return {
    templateId: profile.templateId,
    sourceFileName,
    sourcePath,
    profilePath,
    profile
  };
}

function slugify(value: string) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}
