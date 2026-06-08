import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const libraryRoot = path.resolve(process.cwd(), "..", "asset-library");
const fileRoot = path.join(libraryRoot, "files");
const indexPath = path.join(libraryRoot, "index.json");

export type AssetKind = "image" | "table";

export interface AssetRecord {
  assetId: string;
  kind: AssetKind;
  sourceFileName: string;
  storedFileName: string;
  mimeType: string;
  size: number;
  notes: string;
  createdAt: string;
}

export async function listAssets(kind?: AssetKind) {
  const assets = await readIndex();
  return kind ? assets.filter((asset) => asset.kind === kind) : assets;
}

export async function saveAsset(input: {
  kind: AssetKind;
  sourceFileName: string;
  mimeType: string;
  buffer: Buffer;
  notes?: string;
}) {
  await ensureLibrary();
  const assetId = `asset_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const extension = path.extname(input.sourceFileName) || extensionFromMime(input.mimeType);
  const storedFileName = `${assetId}${extension}`;
  await writeFile(path.join(fileRoot, storedFileName), input.buffer);

  const record: AssetRecord = {
    assetId,
    kind: input.kind,
    sourceFileName: input.sourceFileName,
    storedFileName,
    mimeType: input.mimeType,
    size: input.buffer.length,
    notes: input.notes ?? "",
    createdAt: new Date().toISOString()
  };

  const assets = await readIndex();
  assets.push(record);
  await writeIndex(assets);
  return record;
}

export async function getAsset(assetId: string) {
  const assets = await readIndex();
  return assets.find((asset) => asset.assetId === assetId) ?? null;
}

export async function readAssetBase64(assetId: string) {
  const asset = await getAsset(assetId);
  if (!asset) {
    return null;
  }

  const buffer = await readFile(path.join(fileRoot, asset.storedFileName));
  return {
    asset,
    base64: buffer.toString("base64")
  };
}

export async function deleteAsset(assetId: string) {
  const assets = await readIndex();
  const asset = assets.find((item) => item.assetId === assetId);
  if (!asset) {
    return false;
  }

  await rm(path.join(fileRoot, asset.storedFileName), { force: true });
  await writeIndex(assets.filter((item) => item.assetId !== assetId));
  return true;
}

async function readIndex(): Promise<AssetRecord[]> {
  await ensureLibrary();
  try {
    const raw = await readFile(indexPath, "utf8");
    return JSON.parse(raw) as AssetRecord[];
  } catch {
    return [];
  }
}

async function writeIndex(records: AssetRecord[]) {
  await ensureLibrary();
  await writeFile(indexPath, JSON.stringify(records, null, 2), "utf8");
}

async function ensureLibrary() {
  await mkdir(fileRoot, { recursive: true });
}

function extensionFromMime(mimeType: string) {
  if (mimeType === "image/png") {
    return ".png";
  }
  if (mimeType === "image/jpeg") {
    return ".jpg";
  }
  if (mimeType === "text/csv") {
    return ".csv";
  }
  return ".bin";
}
