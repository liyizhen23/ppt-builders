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
