export interface AssetRecord {
  assetId: string;
  kind: "image" | "table";
  sourceFileName: string;
  storedFileName: string;
  mimeType: string;
  size: number;
  notes: string;
  createdAt: string;
}

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export async function listAssets(kind?: "image" | "table") {
  const suffix = kind ? `?kind=${encodeURIComponent(kind)}` : "";
  const response = await fetch(`${apiBaseUrl}/api/assets${suffix}`);
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `读取素材库失败：${response.status}`);
  }

  return (await response.json()) as { assets: AssetRecord[] };
}

export async function uploadAssets(input: {
  kind: "image" | "table";
  files: File[];
  notes?: string;
}) {
  const formData = new FormData();
  formData.append("kind", input.kind);
  formData.append("notes", input.notes ?? "");
  input.files.forEach((file) => formData.append("asset", file));

  const response = await fetch(`${apiBaseUrl}/api/assets`, {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `保存素材失败：${response.status}`);
  }

  return (await response.json()) as { assets: AssetRecord[] };
}

export async function getAssetBase64(assetId: string) {
  const response = await fetch(`${apiBaseUrl}/api/assets/${encodeURIComponent(assetId)}/base64`);
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `读取素材内容失败：${response.status}`);
  }

  return (await response.json()) as { asset: AssetRecord; base64: string };
}

export async function deleteAsset(assetId: string) {
  const response = await fetch(`${apiBaseUrl}/api/assets/${encodeURIComponent(assetId)}`, {
    method: "DELETE"
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `删除素材失败：${response.status}`);
  }

  return (await response.json()) as { ok: boolean };
}
