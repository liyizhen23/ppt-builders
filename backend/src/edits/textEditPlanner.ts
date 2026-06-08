import { getAiRuntimeConfig } from "../config/aiConfig.js";

export interface TextEditRequest {
  instruction: string;
  selectedText: string;
  context?: string;
}

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

export async function buildSelectedTextEditPlan(input: TextEditRequest): Promise<TextEditPlan> {
  const instruction = input.instruction.trim();
  const selectedText = input.selectedText.trim();

  if (!instruction) {
    return buildPlan(input, {
      replacementText: selectedText,
      clarificationQuestion: "请先告诉我你希望如何修改选中的文本。",
      qa: "No edit was generated because the instruction is empty.",
      model: null
    });
  }

  if (!selectedText) {
    return buildPlan(input, {
      replacementText: "",
      clarificationQuestion: "我没有读取到选中文本。请先在 PowerPoint 中选中文本框里的文字，再点击读取选区。",
      qa: "No edit was generated because no selected text was provided.",
      model: null
    });
  }

  const aiResult = await tryGenerateWithAi({
    instruction,
    selectedText,
    context: input.context
  });

  if (aiResult) {
    return buildPlan(input, {
      replacementText: aiResult.replacementText,
      clarificationQuestion: null,
      qa: "AI rewrite generated. Please review before applying to the selected text box.",
      model: aiResult.model
    });
  }

  return buildPlan(input, {
    replacementText: fallbackRewrite(selectedText, instruction),
    clarificationQuestion: null,
    qa: "Fallback rewrite generated because AI runtime was unavailable.",
    model: null
  });
}

async function tryGenerateWithAi(input: TextEditRequest) {
  const config = getAiRuntimeConfig();
  if (!config.apiKey) {
    return null;
  }

  const response = await fetch(joinApiUrl(config.baseUrl, "chat/completions"), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: config.model,
      temperature: 0.2,
      messages: [
        {
          role: "system",
          content:
            "You rewrite selected PowerPoint text. Return only the replacement text. Preserve facts, do not invent information, and keep it suitable for a slide text box."
        },
        {
          role: "user",
          content: [
            `Instruction: ${input.instruction}`,
            input.context ? `Context: ${input.context}` : "",
            "Selected text:",
            input.selectedText
          ]
            .filter(Boolean)
            .join("\n")
        }
      ]
    })
  });

  if (!response.ok) {
    return null;
  }

  const payload = (await response.json()) as {
    choices?: Array<{
      message?: {
        content?: string;
      };
    }>;
  };
  const replacementText = payload.choices?.[0]?.message?.content?.trim();

  if (!replacementText) {
    return null;
  }

  return {
    replacementText,
    model: config.model
  };
}

function buildPlan(
  input: TextEditRequest,
  result: {
    replacementText: string;
    clarificationQuestion: string | null;
    qa: string;
    model: string | null;
  }
): TextEditPlan {
  return {
    editType: "selected_text_rewrite",
    target: {
      selectionType: "text"
    },
    instruction: input.instruction,
    originalText: input.selectedText,
    replacementText: result.replacementText,
    needsConfirmation: true,
    clarificationQuestion: result.clarificationQuestion,
    qa: result.qa,
    model: result.model,
    layoutSuggestion: buildTextLayoutSuggestion(input.selectedText, result.replacementText)
  };
}

function buildTextLayoutSuggestion(originalText: string, replacementText: string): TextLayoutSuggestion {
  const originalChars = measureSlideText(originalText);
  const replacementChars = measureSlideText(replacementText);
  const relativeLengthChange = round(replacementChars / Math.max(originalChars, 1), 2);
  const originalLines = countLines(originalText);
  const replacementLines = countLines(replacementText);
  const addedLines = replacementLines - originalLines;

  if (relativeLengthChange <= 1.15 && addedLines <= 1) {
    return {
      strategy: "keep",
      reason: "替换文字与原文字长度接近，通常可以保持当前文本框位置和尺寸。",
      estimatedOriginalChars: originalChars,
      estimatedReplacementChars: replacementChars,
      relativeLengthChange,
      suggestedDeltaY: 0,
      suggestedHeightScale: 1,
      suggestedFontScale: 1,
      applyMode: "advisory"
    };
  }

  if (relativeLengthChange <= 1.6 && addedLines <= 2) {
    return {
      strategy: "expand_height",
      reason: "替换文字略长，建议保留当前位置，优先增高文本框以避免溢出。",
      estimatedOriginalChars: originalChars,
      estimatedReplacementChars: replacementChars,
      relativeLengthChange,
      suggestedDeltaY: 0.05,
      suggestedHeightScale: clamp(round(relativeLengthChange, 2), 1.15, 1.45),
      suggestedFontScale: 1,
      applyMode: "requires_shape_api"
    };
  }

  if (relativeLengthChange <= 2.2 && addedLines <= 4) {
    return {
      strategy: "shift_down",
      reason: "替换文字明显变长，建议将文本框略微下移并增高，避免挤压页面上方内容。",
      estimatedOriginalChars: originalChars,
      estimatedReplacementChars: replacementChars,
      relativeLengthChange,
      suggestedDeltaY: 0.18,
      suggestedHeightScale: clamp(round(relativeLengthChange * 0.82, 2), 1.25, 1.75),
      suggestedFontScale: 0.96,
      applyMode: "requires_shape_api"
    };
  }

  return {
    strategy: "reflow_slide",
    reason: "替换文字过长，不适合只改当前文本框，建议使用当前页重排生成替代页。",
    estimatedOriginalChars: originalChars,
    estimatedReplacementChars: replacementChars,
    relativeLengthChange,
    suggestedDeltaY: 0.28,
    suggestedHeightScale: 1.8,
    suggestedFontScale: 0.9,
    applyMode: "use_reflow"
  };
}

function measureSlideText(text: string) {
  return text.replace(/\s+/g, "").length;
}

function countLines(text: string) {
  return text.split(/\r?\n/).filter((line) => line.trim().length > 0).length;
}

function round(value: number, digits: number) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function fallbackRewrite(text: string, instruction: string) {
  if (/缩短|精简|简洁|三条|要点/.test(instruction)) {
    return text
      .split(/[。；;\n]/)
      .map((item) => item.trim())
      .filter(Boolean)
      .slice(0, 3)
      .map((item) => `- ${item}`)
      .join("\n");
  }

  return text;
}

function joinApiUrl(baseUrl: string, path: string) {
  return `${baseUrl.replace(/\/+$/, "")}/${path.replace(/^\/+/, "")}`;
}
