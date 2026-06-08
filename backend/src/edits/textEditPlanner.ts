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
    model: result.model
  };
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
