const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export interface QaIssue {
  id: string;
  severity: "low" | "medium" | "high";
  category: "content" | "layout" | "source" | "style";
  message: string;
  suggestion: string;
}

export interface QaCheckResult {
  passed: boolean;
  issues: QaIssue[];
  summary: string;
}

export interface QaAutofixResult extends QaCheckResult {
  fixedPageText: string;
}

export async function checkPageQa(input: { pageText: string; instruction?: string }) {
  const response = await fetch(`${apiBaseUrl}/api/qa/check`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(input)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `QA 检查失败：${response.status}`);
  }

  return (await response.json()) as QaCheckResult;
}

export async function autofixPageQa(input: { pageText: string; instruction?: string }) {
  const response = await fetch(`${apiBaseUrl}/api/qa/autofix`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(input)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `QA 自动修复失败：${response.status}`);
  }

  return (await response.json()) as QaAutofixResult;
}
