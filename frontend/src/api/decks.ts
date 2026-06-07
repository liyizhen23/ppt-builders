export interface GenerateDeckInput {
  reportFile: File;
  templateFile: File | null;
  instruction: string;
}

export interface GenerateDeckResult {
  deckId?: string;
  pptxBase64: string;
  summary?: string;
  qa?: string;
}

export interface DefaultTemplateResult {
  templateId: string;
  sourceFileName: string;
  counts: {
    slides: number;
    layouts: number;
    masters: number;
    media: number;
  };
  roles: Record<string, number>;
  capabilities?: {
    replaceableSlots: number;
    recommendedSlides: Record<string, number[]>;
    styleTokens: {
      title: unknown;
      subtitle: unknown;
      body: unknown;
      caption: unknown;
      palette: string[];
      fonts: string[];
    };
  };
}

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export async function generateDeck(input: GenerateDeckInput): Promise<GenerateDeckResult> {
  const formData = new FormData();
  formData.append("report", input.reportFile);
  if (input.templateFile) {
    formData.append("template", input.templateFile);
  }
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

export async function getDefaultTemplate(): Promise<DefaultTemplateResult> {
  const response = await fetch(`${apiBaseUrl}/api/templates/default`);

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `读取默认模板失败：${response.status}`);
  }

  return (await response.json()) as DefaultTemplateResult;
}

export async function saveDefaultTemplate(templateFile: File): Promise<DefaultTemplateResult> {
  const formData = new FormData();
  formData.append("template", templateFile);

  const response = await fetch(`${apiBaseUrl}/api/templates/default`, {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `保存默认模板失败：${response.status}`);
  }

  return (await response.json()) as DefaultTemplateResult;
}
