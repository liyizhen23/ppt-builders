const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export interface ReflowSlideResult {
  pptxBase64: string;
  slideSpec: {
    title: string;
    bullets: string[];
    selectedImageFileName: string | null;
  };
  qa: string;
}

export async function reflowCurrentSlide(input: {
  instruction: string;
  pageText: string;
  selectedImageFileName?: string | null;
}) {
  const response = await fetch(`${apiBaseUrl}/api/edits/slide/reflow`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(input)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `当前页重排失败：${response.status}`);
  }

  return (await response.json()) as ReflowSlideResult;
}
