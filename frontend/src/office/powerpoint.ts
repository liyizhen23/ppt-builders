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
