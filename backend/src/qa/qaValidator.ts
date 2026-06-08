export interface QaCheckInput {
  pageText: string;
  instruction?: string;
}

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

export function checkPageQa(input: QaCheckInput): QaCheckResult {
  const text = input.pageText.trim();
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const sentences = text.split(/[。；;.!?！？]/).map((item) => item.trim()).filter(Boolean);
  const issues: QaIssue[] = [];

  if (!text) {
    issues.push({
      id: "empty-page-text",
      severity: "high",
      category: "content",
      message: "当前页内容为空，无法进行可靠重排或检查。",
      suggestion: "先读取选区或粘贴当前页标题、正文和图片说明。"
    });
  }

  if (text.length > 520) {
    issues.push({
      id: "too-much-text",
      severity: "high",
      category: "layout",
      message: "当前页文字量偏大，直接放入一页容易拥挤。",
      suggestion: "压缩为 3-5 条要点，或拆分为两页。"
    });
  }

  if (sentences.length > 7) {
    issues.push({
      id: "too-many-points",
      severity: "medium",
      category: "layout",
      message: "当前页信息点过多，读者难以快速扫描。",
      suggestion: "保留最关键的 3-5 个信息点，其余内容移入备注或下一页。"
    });
  }

  if (lines.length <= 1 && text.length > 120) {
    issues.push({
      id: "missing-hierarchy",
      severity: "medium",
      category: "style",
      message: "文本缺少标题和要点层级。",
      suggestion: "将第一句作为标题，其余内容拆成要点。"
    });
  }

  if (!/[图表]|figure|chart|image|图片|表格/i.test(`${text} ${input.instruction ?? ""}`)) {
    issues.push({
      id: "visual-intent-missing",
      severity: "low",
      category: "layout",
      message: "未检测到明确的图表或图片意图。",
      suggestion: "如果该页需要配图，请在重排要求中说明图片或表格用途。"
    });
  }

  return {
    passed: issues.filter((issue) => issue.severity !== "low").length === 0,
    issues,
    summary: issues.length === 0 ? "QA passed. No obvious page-level issues found." : `QA found ${issues.length} issue(s).`
  };
}

export function autofixPageQa(input: QaCheckInput): QaAutofixResult {
  const checked = checkPageQa(input);
  const text = input.pageText.trim();
  const units = text
    .split(/[\r\n。；;.!?！？]/)
    .map((item) => item.trim())
    .filter(Boolean);

  const titleSource = input.instruction?.trim() || units[0] || "当前页重点";
  const title = titleSource.length > 32 ? `${titleSource.slice(0, 32)}...` : titleSource;
  const bullets = units
    .filter((item) => item !== units[0])
    .slice(0, 5)
    .map((item) => (item.length > 58 ? `${item.slice(0, 58)}...` : item));

  const fixedPageText = [title, ...bullets.map((bullet) => `- ${bullet}`)].join("\n");

  return {
    ...checked,
    fixedPageText
  };
}
