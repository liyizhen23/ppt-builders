export interface AiSettingsResult {
  provider: "openai" | "custom";
  configured: boolean;
  baseUrlHost: string;
  model: string;
}

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export async function getAiSettings(): Promise<AiSettingsResult> {
  const response = await fetch(`${apiBaseUrl}/api/settings/ai`);

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `读取 AI API 设置失败：${response.status}`);
  }

  return (await response.json()) as AiSettingsResult;
}
