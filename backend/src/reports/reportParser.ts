import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  parseTagValue: false,
  parseAttributeValue: false
});

export interface EvidenceIndex {
  reportId: string;
  sourceFileName: string;
  generatedAt: string;
  blocks: EvidenceBlock[];
  stats: {
    paragraphs: number;
    tables: number;
    characters: number;
  };
}

export interface EvidenceBlock {
  evidenceId: string;
  type: "paragraph" | "table" | "text";
  order: number;
  text: string;
  keywords: string[];
}

export async function parseReportToEvidence(input: {
  reportId: string;
  fileName: string;
  buffer: Buffer;
}): Promise<EvidenceIndex> {
  const lower = input.fileName.toLowerCase();
  if (lower.endsWith(".docx")) {
    return parseDocx(input);
  }

  const text = input.buffer.toString("utf8");
  return buildEvidenceFromText({
    reportId: input.reportId,
    sourceFileName: input.fileName,
    paragraphs: splitTextBlocks(text),
    tables: []
  });
}

export function findRelevantEvidence(index: EvidenceIndex, instruction: string, limit = 4) {
  const terms = tokenize(instruction);
  const scored = index.blocks.map((block) => {
    const text = block.text.toLowerCase();
    const keywordHits = block.keywords.filter((keyword) => terms.includes(keyword)).length;
    const textHits = terms.filter((term) => text.includes(term)).length;
    const poiBoost = /poi|兴趣点|感知|方法|流程|轨迹|地理|位置/i.test(block.text) ? 1.5 : 0;
    return {
      block,
      score: keywordHits * 2 + textHits + poiBoost
    };
  });

  const matches = scored
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.block.order - b.block.order)
    .map((item) => item.block)
    .slice(0, limit);

  return matches.length > 0 ? matches : index.blocks.slice(0, limit);
}

async function parseDocx(input: { reportId: string; fileName: string; buffer: Buffer }) {
  const zip = await JSZip.loadAsync(input.buffer);
  const document = await readXml(zip, "word/document.xml");
  const body = document?.["w:document"]?.["w:body"];
  const paragraphs: string[] = [];
  const tables: string[] = [];

  for (const child of children(body)) {
    if (child.key === "w:p") {
      const text = normalizeText(extractText(child.value));
      if (text) {
        paragraphs.push(text);
      }
    } else if (child.key === "w:tbl") {
      const text = normalizeText(extractText(child.value));
      if (text) {
        tables.push(text);
      }
    }
  }

  return buildEvidenceFromText({
    reportId: input.reportId,
    sourceFileName: input.fileName,
    paragraphs,
    tables
  });
}

function buildEvidenceFromText(input: {
  reportId: string;
  sourceFileName: string;
  paragraphs: string[];
  tables: string[];
}): EvidenceIndex {
  const blocks: EvidenceBlock[] = [];
  let order = 1;

  for (const paragraph of input.paragraphs) {
    blocks.push({
      evidenceId: `${input.reportId}:p${order}`,
      type: "paragraph",
      order,
      text: paragraph,
      keywords: tokenize(paragraph).slice(0, 20)
    });
    order += 1;
  }

  for (const table of input.tables) {
    blocks.push({
      evidenceId: `${input.reportId}:table${order}`,
      type: "table",
      order,
      text: table,
      keywords: tokenize(table).slice(0, 20)
    });
    order += 1;
  }

  return {
    reportId: input.reportId,
    sourceFileName: input.sourceFileName,
    generatedAt: new Date().toISOString(),
    blocks,
    stats: {
      paragraphs: input.paragraphs.length,
      tables: input.tables.length,
      characters: blocks.reduce((sum, block) => sum + block.text.length, 0)
    }
  };
}

async function readXml(zip: JSZip, partName: string) {
  const file = zip.file(partName);
  if (!file) {
    return null;
  }

  return parser.parse(await file.async("text"));
}

function extractText(value: unknown): string {
  const fragments: string[] = [];
  walk(value, (node) => {
    if (isRecord(node) && typeof node["w:t"] === "string") {
      fragments.push(node["w:t"]);
    }
  });
  return fragments.join("");
}

function splitTextBlocks(text: string) {
  return text
    .split(/\r?\n+/)
    .map((line) => normalizeText(line))
    .filter((line): line is string => Boolean(line));
}

function tokenize(text: string) {
  const terms = new Set<string>();
  const normalized = text.toLowerCase();
  for (const match of normalized.matchAll(/[a-z0-9]{2,}|[\u4e00-\u9fa5]{2,}/g)) {
    terms.add(match[0]);
  }
  return [...terms].filter((term) => !stopwords.has(term));
}

function children(value: unknown): Array<{ key: string; value: unknown }> {
  if (!isRecord(value)) {
    return [];
  }

  const result: Array<{ key: string; value: unknown }> = [];
  for (const [key, child] of Object.entries(value)) {
    if (!key.startsWith("w:")) {
      continue;
    }
    if (Array.isArray(child)) {
      result.push(...child.map((item) => ({ key, value: item })));
    } else {
      result.push({ key, value: child });
    }
  }
  return result;
}

function walk(value: unknown, visit: (node: unknown) => void) {
  visit(value);
  if (Array.isArray(value)) {
    for (const child of value) {
      walk(child, visit);
    }
  } else if (isRecord(value)) {
    for (const child of Object.values(value)) {
      walk(child, visit);
    }
  }
}

function normalizeText(value: string | null | undefined) {
  const normalized = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  return normalized || null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

const stopwords = new Set([
  "the",
  "and",
  "for",
  "with",
  "this",
  "that",
  "其中",
  "一个",
  "进行",
  "生成",
  "报告",
  "方法",
  "可以",
  "基于"
]);
