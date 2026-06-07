export interface GenerateDeckInput {
  reportFile: File;
  templateFile: File;
  instruction: string;
}

export interface GenerateDeckResult {
  deckId?: string;
  pptxBase64: string;
  summary?: string;
  qa?: string;
}

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export async function generateDeck(input: GenerateDeckInput): Promise<GenerateDeckResult> {
  const formData = new FormData();
  formData.append("report", input.reportFile);
  formData.append("template", input.templateFile);
  formData.append("instruction", input.instruction);

  const response = await fetch(`${apiBaseUrl}/api/decks/generate`, {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `生成请求失败：${response.status}`);
  }

  const result = (await response.json()) as Partial<GenerateDeckResult>;

  if (!result.pptxBase64) {
    throw new Error("后端响应缺少 pptxBase64。");
  }

  return {
    deckId: result.deckId,
    pptxBase64: result.pptxBase64,
    summary: result.summary,
    qa: result.qa
  };
}
