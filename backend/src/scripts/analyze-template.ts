import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, join, resolve } from "node:path";
import { analyzePptxTemplate } from "../templates/templateProfile.js";

const [inputPath, outputDirArg] = process.argv.slice(2);

if (!inputPath) {
  console.error("Usage: tsx src/scripts/analyze-template.ts <template.pptx> [output-dir]");
  process.exit(1);
}

const resolvedInput = resolve(inputPath);
const outputDir = resolve(outputDirArg ?? "templates/analyzed-template");
const buffer = await readFile(resolvedInput);
const templateId = slugify(basename(resolvedInput, ".pptx"));
const profile = await analyzePptxTemplate({
  templateId,
  sourceFileName: basename(resolvedInput),
  buffer
});

await mkdir(outputDir, { recursive: true });
await writeFile(join(outputDir, "template-profile.json"), JSON.stringify(profile, null, 2), "utf8");

console.log(
  JSON.stringify(
    {
      output: join(outputDir, "template-profile.json"),
      templateId: profile.templateId,
      counts: profile.counts,
      roles: profile.slides.reduce<Record<string, number>>((acc, slide) => {
        acc[slide.role] = (acc[slide.role] ?? 0) + 1;
        return acc;
      }, {})
    },
    null,
    2
  )
);

function slugify(value: string) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}
