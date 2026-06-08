export interface CurrentReportResult {
  configured: boolean;
  reportId?: string;
  sourceFileName?: string;
  evidenceBlocks?: number;
  stats?: {
    paragraphs: number;
    tables: number;
    characters: number;
  };
}

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export async function getCurrentReport(): Promise<CurrentReportResult> {
  const response = await fetch(`${apiBaseUrl}/api/reports/current`);

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `读取当前报告失败：${response.status}`);
  }

  return (await response.json()) as CurrentReportResult;
}
