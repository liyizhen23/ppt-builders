import type { TextLayoutSuggestion } from "../api/edits";

export function isPowerPointHost() {
  return Boolean(window.Office && Office.context?.host === Office.HostType.PowerPoint);
}

export async function insertSlidesFromBase64(pptxBase64: string) {
  await PowerPoint.run(async (context) => {
    context.presentation.insertSlidesFromBase64(pptxBase64, {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting
    });

    await context.sync();
  });
}

export async function readSelectedText() {
  return new Promise<string>((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(String(result.value ?? ""));
        return;
      }

      reject(new Error(result.error.message));
    });
  });
}

export async function replaceSelectedText(text: string) {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      text,
      {
        coercionType: Office.CoercionType.Text
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(new Error(result.error.message));
      }
    );
  });
}

export async function adjustSelectedTextBoxLayout(suggestion: TextLayoutSuggestion) {
  if (suggestion.strategy === "keep") {
    return false;
  }

  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();

    const shape = shapes.items[0];
    if (!shape) {
      return false;
    }

    shape.load("top,height");
    const textFrame = shape.getTextFrameOrNullObject();
    textFrame.load("isNullObject,wordWrap");
    const font = textFrame.textRange.font;
    font.load("size");
    await context.sync();

    if (textFrame.isNullObject) {
      return false;
    }

    const deltaPoints = suggestion.suggestedDeltaY * 72;
    const heightScale = Math.max(suggestion.suggestedHeightScale, 1);

    if (deltaPoints > 0) {
      shape.top += deltaPoints;
    }

    if (heightScale > 1) {
      shape.height = Math.max(shape.height * heightScale, shape.height + 8);
    }

    if (typeof font.size === "number" && suggestion.suggestedFontScale < 1) {
      font.size = Math.max(8, Math.round(font.size * suggestion.suggestedFontScale * 10) / 10);
    }

    textFrame.wordWrap = true;
    await context.sync();
    return true;
  });
}

export async function replaceSelectedImage(imageBase64: string) {
  await assertSelectedImageTargetAllowed();

  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      imageBase64,
      {
        coercionType: Office.CoercionType.Image
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(new Error(result.error.message));
      }
    );
  });
}

export async function replaceSelectedShapeTexts(replacements: string[]) {
  const nextTexts = replacements.map((text) => text.trim()).filter(Boolean);
  if (nextTexts.length === 0) {
    return 0;
  }

  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();

    let replacementIndex = 0;

    for (const shape of shapes.items) {
      if (replacementIndex >= nextTexts.length) {
        break;
      }

      const textFrame = shape.getTextFrameOrNullObject();
      textFrame.load("isNullObject");
      await context.sync();

      if (textFrame.isNullObject) {
        continue;
      }

      textFrame.textRange.text = nextTexts[replacementIndex];
      replacementIndex += 1;
    }

    await context.sync();
    return replacementIndex;
  });
}

async function assertSelectedImageTargetAllowed() {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();

    if (shapes.items.length === 0) {
      throw new Error("请先选中正文区域内要替换的图片或图片占位框。");
    }

    for (const shape of shapes.items) {
      shape.load("left,top,width,height,name");
    }
    await context.sync();

    const protectedShape = shapes.items.find((shape) => isLikelyHeaderLogo(shape));
    if (protectedShape) {
      throw new Error(
        `当前选区位于右上角页眉/校徽保护区域（${protectedShape.name || "未命名形状"}），已阻止图片替换。请选中正文中的图片占位框后再应用。`
      );
    }
  });
}

function isLikelyHeaderLogo(shape: PowerPoint.Shape) {
  const left = numberOrZero(shape.left);
  const top = numberOrZero(shape.top);
  const width = numberOrZero(shape.width);
  const height = numberOrZero(shape.height);
  const name = String(shape.name ?? "").toLowerCase();

  const nameLooksProtected = /logo|校徽|tsinghua|清华/.test(name);
  const inTopRightHeader = top <= 70 && left >= 560 && width <= 180 && height <= 90;

  return nameLooksProtected || inTopRightHeader;
}

function numberOrZero(value: number | undefined) {
  return typeof value === "number" && Number.isFinite(value) ? value : 0;
}
