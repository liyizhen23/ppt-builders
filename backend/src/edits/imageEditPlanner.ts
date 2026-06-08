export interface ImageEditRequest {
  instruction: string;
  imageFileName?: string;
  imageMimeType?: string;
}

export interface ImageEditPlan {
  editType: "selected_image_replace";
  target: {
    selectionType: "image";
  };
  instruction: string;
  imageFileName: string | null;
  operation: "replace_image" | "format_guidance";
  needsConfirmation: boolean;
  clarificationQuestion: string | null;
  apply: {
    method: "office_set_selected_image";
    preservesAspectRatio: boolean;
    keepsOriginalFrameWhenPowerPointAllows: boolean;
  };
  qa: string;
}

export interface ImageLibraryCandidate {
  id: string;
  fileName: string;
  mimeType?: string;
  notes?: string;
}

export interface ImageSelectionRequest {
  instruction: string;
  pageText: string;
  candidates: ImageLibraryCandidate[];
}

export interface ImageSelectionPlan {
  editType: "image_library_select";
  target: {
    selectionType: "image";
  };
  instruction: string;
  pageText: string;
  selectedImageId: string | null;
  selectedImageFileName: string | null;
  confidence: "low" | "medium" | "high";
  reason: string;
  candidatesReviewed: number;
  needsConfirmation: boolean;
  qa: string;
}

export function buildSelectedImageEditPlan(input: ImageEditRequest): ImageEditPlan {
  const instruction = input.instruction.trim();
  const imageFileName = input.imageFileName?.trim() || null;

  if (!instruction) {
    return {
      editType: "selected_image_replace",
      target: {
        selectionType: "image"
      },
      instruction,
      imageFileName,
      operation: "format_guidance",
      needsConfirmation: true,
      clarificationQuestion: "请说明希望如何处理图片，例如替换当前图片、保持比例、居中或裁剪为模板图片框比例。",
      apply: {
        method: "office_set_selected_image",
        preservesAspectRatio: true,
        keepsOriginalFrameWhenPowerPointAllows: true
      },
      qa: "No image edit was generated because the instruction is empty."
    };
  }

  return {
    editType: "selected_image_replace",
    target: {
      selectionType: "image"
    },
    instruction,
    imageFileName,
    operation: imageFileName ? "replace_image" : "format_guidance",
    needsConfirmation: true,
    clarificationQuestion: imageFileName
      ? null
      : "当前没有上传替换图片。第一版可以生成图片处理建议；真正替换图片需要先选择一张本地图片。",
    apply: {
      method: "office_set_selected_image",
      preservesAspectRatio: true,
      keepsOriginalFrameWhenPowerPointAllows: true
    },
    qa: imageFileName
      ? "Image replacement plan generated. Office.js will insert the uploaded image into the current selection after confirmation."
      : "Format-only image edits are planned but not directly applied in this MVP because PowerPoint Office.js exposes limited shape formatting controls."
  };
}

export function buildImageSelectionPlan(input: ImageSelectionRequest): ImageSelectionPlan {
  const candidates = input.candidates.filter((candidate) => candidate.fileName.trim());
  const query = `${input.instruction} ${input.pageText}`.trim();

  if (candidates.length === 0) {
    return {
      editType: "image_library_select",
      target: {
        selectionType: "image"
      },
      instruction: input.instruction,
      pageText: input.pageText,
      selectedImageId: null,
      selectedImageFileName: null,
      confidence: "low",
      reason: "没有可供选择的候选图片。",
      candidatesReviewed: 0,
      needsConfirmation: true,
      qa: "No image was selected because the candidate list is empty."
    };
  }

  const queryTokens = tokenize(query);
  const scored = candidates
    .map((candidate) => {
      const candidateText = `${candidate.fileName} ${candidate.notes ?? ""}`;
      const candidateTokens = tokenize(candidateText);
      const overlap = candidateTokens.filter((token) => queryTokens.includes(token));
      const semanticBoost = semanticHints(query).filter((hint) => candidateText.toLowerCase().includes(hint));
      const chartBoost = /图表|图|表|chart|figure|diagram/i.test(query) && /图|表|chart|figure|diagram/i.test(candidateText);
      return {
        candidate,
        score: overlap.length * 2 + semanticBoost.length * 3 + (chartBoost ? 2 : 0),
        overlap,
        semanticBoost
      };
    })
    .sort((a, b) => b.score - a.score);

  const best = scored[0];
  const confidence = best.score >= 6 ? "high" : best.score >= 2 ? "medium" : "low";
  const reasonParts = [
    best.overlap.length ? `匹配关键词：${best.overlap.slice(0, 6).join("、")}` : "",
    best.semanticBoost.length ? `语义提示：${best.semanticBoost.slice(0, 4).join("、")}` : ""
  ].filter(Boolean);

  return {
    editType: "image_library_select",
    target: {
      selectionType: "image"
    },
    instruction: input.instruction,
    pageText: input.pageText,
    selectedImageId: best.candidate.id,
    selectedImageFileName: best.candidate.fileName,
    confidence,
    reason: reasonParts.length
      ? reasonParts.join("；")
      : "未找到明显关键词匹配，先选择候选图片列表中的第一张。建议给图片文件名加入主题关键词。",
    candidatesReviewed: candidates.length,
    needsConfirmation: true,
    qa: "Image library selection uses page text, user instruction, file names, and optional notes. Image visual understanding is not enabled in this MVP."
  };
}

function tokenize(text: string) {
  const normalized = text.toLowerCase();
  const latin = normalized.match(/[a-z0-9_+-]{2,}/g) ?? [];
  const chinese = normalized.match(/[\u4e00-\u9fa5]{2,}/g) ?? [];
  const chineseFragments = chinese.flatMap((chunk) => {
    const fragments: string[] = [];
    for (let size = 2; size <= Math.min(4, chunk.length); size += 1) {
      for (let index = 0; index <= chunk.length - size; index += 1) {
        fragments.push(chunk.slice(index, index + size));
      }
    }
    return fragments;
  });
  return Array.from(new Set([...latin, ...chinese, ...chineseFragments]));
}

function semanticHints(text: string) {
  const lower = text.toLowerCase();
  const hints: string[] = [];

  if (/poi|兴趣点|类别|分类|类型|业态/.test(lower)) {
    hints.push("poi", "分类", "类别", "类型", "兴趣点", "业态", "category", "type");
  }

  if (/方法|流程|框架|路径|技术路线/.test(lower)) {
    hints.push("方法", "流程", "框架", "路线", "method", "workflow", "framework");
  }

  if (/结果|对比|分析|统计|分布/.test(lower)) {
    hints.push("结果", "对比", "分析", "统计", "分布", "result", "compare", "analysis");
  }

  return hints;
}
