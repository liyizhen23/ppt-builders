import { mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { basename, join, resolve } from "node:path";
import { EvidenceIndex, parseReportToEvidence } from "./reportParser.js";

const rootDir = resolve(process.cwd(), "..");
const storageDir = resolve(process.env.REPORT_STORAGE_DIR ?? join(rootDir, "reports", "current"));
const sourcePath = join(storageDir, "source");
const metaPath = join(storageDir, "report-meta.json");
const evidencePath = join(storageDir, "evidence-index.json");

export interface CurrentReportRecord {
  reportId: string;
  sourceFileName: string;
  sourcePath: string;
  evidencePath: string;
  evidenceIndex: EvidenceIndex;
}

interface ReportMeta {
  reportId: string;
  sourceFileName: string;
  sourcePath: string;
}

export async function saveCurrentReport(input: {
  fileName: string;
  buffer: Buffer;
}): Promise<CurrentReportRecord> {
  await mkdir(storageDir, { recursive: true });
  const reportId = `report_${Date.now()}`;
  const storedSourcePath = `${sourcePath}${extension(input.fileName)}`;
  await writeFile(storedSourcePath, input.buffer);
  const evidenceIndex = await parseReportToEvidence({
    reportId,
    fileName: input.fileName,
    buffer: input.buffer
  });
  const meta: ReportMeta = {
    reportId,
    sourceFileName: input.fileName,
    sourcePath: storedSourcePath
  };
  await writeFile(metaPath, JSON.stringify(meta, null, 2), "utf8");
  await writeFile(evidencePath, JSON.stringify(evidenceIndex, null, 2), "utf8");

  return {
    reportId,
    sourceFileName: input.fileName,
    sourcePath: storedSourcePath,
    evidencePath,
    evidenceIndex
  };
}

export async function getCurrentReport(): Promise<CurrentReportRecord | null> {
  if (!existsSync(metaPath) || !existsSync(evidencePath)) {
    return null;
  }

  const meta = JSON.parse(await readFile(metaPath, "utf8")) as ReportMeta;
  const evidenceIndex = JSON.parse(await readFile(evidencePath, "utf8")) as EvidenceIndex;

  return {
    reportId: meta.reportId,
    sourceFileName: meta.sourceFileName,
    sourcePath: meta.sourcePath,
    evidencePath,
    evidenceIndex
  };
}

export function summarizeCurrentReport(record: CurrentReportRecord | null) {
  if (!record) {
    return {
      configured: false
    };
  }

  return {
    configured: true,
    reportId: record.reportId,
    sourceFileName: basename(record.sourceFileName),
    evidenceBlocks: record.evidenceIndex.blocks.length,
    stats: record.evidenceIndex.stats
  };
}

function extension(fileName: string) {
  const match = fileName.match(/\.[^.]+$/);
  return match?.[0] ?? ".bin";
}
