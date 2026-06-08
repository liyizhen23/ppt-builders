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
