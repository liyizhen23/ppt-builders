export interface TextEditPlan {
  editType: "selected_text_rewrite";
  target: {
    selectionType: "text";
  };
  instruction: string;
  originalText: string;
  replacementText: string;
  needsConfirmation: boolean;
  clarificationQuestion: string | null;
  qa: string;
  model: string | null;
  layoutSuggestion: TextLayoutSuggestion;
}

export interface TextLayoutSuggestion {
  strategy: "keep" | "expand_height" | "shift_down" | "shrink_font" | "reflow_slide";
  reason: string;
  estimatedOriginalChars: number;
  estimatedReplacementChars: number;
  relativeLengthChange: number;
  suggestedDeltaY: number;
  suggestedHeightScale: number;
  suggestedFontScale: number;
  applyMode: "advisory" | "requires_shape_api" | "use_reflow";
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

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || "";

export async function planSelectedTextEdit(input: {
  instruction: string;
  selectedText: string;
  context?: string;
}): Promise<TextEditPlan> {
  const response = await fetch(`${apiBaseUrl}/api/edits/selection/text`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(input)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `局部文本编辑请求失败：${response.status}`);
  }

  const result = (await response.json()) as { editPlan?: TextEditPlan };
  if (!result.editPlan) {
    throw new Error("后端响应缺少 editPlan。");
  }

  return result.editPlan;
}

export async function planSelectedImageEdit(input: {
  instruction: string;
  imageFileName?: string;
  imageMimeType?: string;
}): Promise<ImageEditPlan> {
  const response = await fetch(`${apiBaseUrl}/api/edits/selection/image`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(input)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `局部图片编辑请求失败：${response.status}`);
  }

  const result = (await response.json()) as { editPlan?: ImageEditPlan };
  if (!result.editPlan) {
    throw new Error("后端响应缺少 image editPlan。");
  }

  return result.editPlan;
}

export async function planImageLibrarySelection(input: {
  instruction: string;
  pageText: string;
  candidates: Array<{
    id: string;
    fileName: string;
    mimeType?: string;
    notes?: string;
  }>;
}): Promise<ImageSelectionPlan> {
  const response = await fetch(`${apiBaseUrl}/api/edits/selection/image/select`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(input)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `图片库选图请求失败：${response.status}`);
  }

  const result = (await response.json()) as { editPlan?: ImageSelectionPlan };
  if (!result.editPlan) {
    throw new Error("后端响应缺少 image selection editPlan。");
  }

  return result.editPlan;
}
