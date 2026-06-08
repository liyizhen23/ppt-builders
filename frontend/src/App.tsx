import { FormEvent, useEffect, useMemo, useState } from "react";
import { AssetRecord, getAssetBase64, listAssets, uploadAssets } from "./api/assets";
import {
  DefaultTemplateResult,
  generateDeck,
  GenerateDeckResult,
  getDefaultTemplate,
  saveDefaultTemplate
} from "./api/decks";
import {
  ImageEditPlan,
  ImageSelectionPlan,
  planImageLibrarySelection,
  planSelectedImageEdit,
  planSelectedTextEdit,
  TextEditPlan
} from "./api/edits";
import { CurrentReportResult, getCurrentReport } from "./api/reports";
import { autofixPageQa, checkPageQa, QaCheckResult } from "./api/qa";
import { reflowCurrentSlide, ReflowSlideResult } from "./api/reflow";
import { AiSettingsResult, getAiSettings } from "./api/settings";
import {
  adjustSelectedTextBoxLayout,
  insertSlidesFromBase64,
  isPowerPointHost,
  readSelectedText,
  replaceSelectedImage,
  replaceSelectedShapeTexts,
  replaceSelectedText
} from "./office/powerpoint";

type Status = "idle" | "generating" | "ready" | "inserting" | "inserted" | "saving" | "editing" | "error";
type WorkspaceMode = "generate" | "edit";
type EditMode = "text" | "image" | "slide";

export function App() {
  const [mode, setMode] = useState<WorkspaceMode>("generate");
  const [editMode, setEditMode] = useState<EditMode>("text");
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [currentReport, setCurrentReport] = useState<CurrentReportResult | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [defaultTemplate, setDefaultTemplate] = useState<DefaultTemplateResult | null>(null);
  const [aiSettings, setAiSettings] = useState<AiSettingsResult | null>(null);
  const [instruction, setInstruction] = useState("");
  const [result, setResult] = useState<GenerateDeckResult | null>(null);

  const [selectedText, setSelectedText] = useState("");
  const [editInstruction, setEditInstruction] = useState("");
  const [proposalText, setProposalText] = useState("");
  const [editPlan, setEditPlan] = useState<TextEditPlan | null>(null);

  const [pageText, setPageText] = useState("");
  const [imageInstruction, setImageInstruction] = useState("");
  const [shapeTextReplacements, setShapeTextReplacements] = useState("");
  const [imageFile, setImageFile] = useState<File | null>(null);
  const [imageBase64, setImageBase64] = useState("");
  const [imagePlan, setImagePlan] = useState<ImageEditPlan | null>(null);
  const [imageSelectionPlan, setImageSelectionPlan] = useState<ImageSelectionPlan | null>(null);
  const [assets, setAssets] = useState<AssetRecord[]>([]);
  const [assetUploadFiles, setAssetUploadFiles] = useState<File[]>([]);
  const [assetNotes, setAssetNotes] = useState("");

  const [slideInstruction, setSlideInstruction] = useState("");
  const [reflowResult, setReflowResult] = useState<ReflowSlideResult | null>(null);
  const [qaResult, setQaResult] = useState<QaCheckResult | null>(null);

  const [status, setStatus] = useState<Status>("idle");
  const [message, setMessage] = useState("选择工作模式。完整生成需要报告；局部编辑不需要上传报告。");

  useEffect(() => {
    refreshInitialData();
  }, []);

  const imageAssets = assets.filter((asset) => asset.kind === "image");
  const canGenerate = useMemo(
    () => (reportFile !== null || currentReport?.configured === true) && status !== "generating",
    [currentReport, reportFile, status]
  );
  const canSaveDefault = templateFile !== null && status !== "saving";
  const canInsert = Boolean(result?.pptxBase64) && status !== "inserting";
  const canPlanTextEdit = editInstruction.trim().length > 0 && selectedText.trim().length > 0 && status !== "editing";
  const canApplyTextEdit = proposalText.trim().length > 0 && status !== "editing";
  const canPlanImageEdit = imageInstruction.trim().length > 0 && status !== "editing";
  const canApplyImageEdit = Boolean(imagePlan && imageBase64 && imageFile) && status !== "editing";
  const canApplyShapeTexts = shapeTextReplacements.trim().length > 0 && status !== "editing";
  const canPlanImageSelection =
    imageAssets.length > 0 && (pageText.trim().length > 0 || imageInstruction.trim().length > 0) && status !== "editing";
  const canApplyImageSelection = Boolean(imageSelectionPlan?.selectedImageId) && status !== "editing";
  const canReflow = (pageText.trim().length > 0 || slideInstruction.trim().length > 0) && status !== "editing";
  const canInsertReflow = Boolean(reflowResult?.pptxBase64) && status !== "inserting";
  const canCheckQa = pageText.trim().length > 0 && status !== "editing";

  async function refreshInitialData() {
    getDefaultTemplate()
      .then((template) => setDefaultTemplate(template))
      .catch((error) => {
        setStatus("error");
        setMessage(error instanceof Error ? error.message : "读取默认模板失败。");
      });
    getAiSettings()
      .then((settings) => setAiSettings(settings))
      .catch(() => setAiSettings(null));
    getCurrentReport()
      .then((report) => setCurrentReport(report))
      .catch(() => setCurrentReport(null));
    refreshAssets();
  }

  async function refreshAssets() {
    try {
      const response = await listAssets();
      setAssets(response.assets);
    } catch {
      setAssets([]);
    }
  }

  const handleGenerate = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!reportFile && !currentReport?.configured) {
      setStatus("error");
      setMessage("完整生成需要报告文件。首次上传后，后续生成可以复用当前报告。");
      return;
    }

    setStatus("generating");
    setMessage(templateFile ? "正在使用上传模板生成 PPT。" : "正在使用默认模板生成 PPT。");
    setResult(null);

    try {
      const generated = await generateDeck({ reportFile, templateFile, instruction });
      setResult(generated);
      setCurrentReport(await getCurrentReport());
      setStatus("ready");
      setMessage(generated.summary || "后端已返回 PPTX Base64，可以插入 PowerPoint。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "生成失败。");
    }
  };

  const handleSaveDefaultTemplate = async () => {
    if (!templateFile) {
      setStatus("error");
      setMessage("请先选择一个 PPT 模板。");
      return;
    }

    setStatus("saving");
    setMessage("正在解析并保存默认模板。");

    try {
      const saved = await saveDefaultTemplate(templateFile);
      setDefaultTemplate(saved);
      setStatus("idle");
      setMessage(`默认模板已保存：${saved.sourceFileName}`);
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "保存默认模板失败。");
    }
  };

  const handleInsert = async () => {
    if (!result?.pptxBase64) {
      setStatus("error");
      setMessage("没有可插入的 PPTX Base64。");
      return;
    }
    await insertPptx(result.pptxBase64, "已插入到当前演示文稿。");
  };

  const handleReadSelection = async () => {
    if (!isPowerPointHost()) {
      setStatus("error");
      setMessage("请在 PowerPoint 插件任务窗格中读取选区。");
      return;
    }

    setStatus("editing");
    setMessage("正在读取当前选中文本。");
    setEditPlan(null);
    setProposalText("");

    try {
      const text = await readSelectedText();
      setSelectedText(text);
      setPageText(text);
      setStatus("idle");
      setMessage(text.trim() ? "已读取选中文本。" : "未读取到文本，请先选中文本框内的文字。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "读取选区失败。");
    }
  };

  const handlePlanTextEdit = async () => {
    setStatus("editing");
    setMessage("正在生成局部文本修改方案。");
    setEditPlan(null);
    setProposalText("");

    try {
      const plan = await planSelectedTextEdit({ instruction: editInstruction, selectedText });
      setEditPlan(plan);
      setProposalText(plan.replacementText);
      setStatus("ready");
      setMessage(plan.clarificationQuestion || "已生成文本修改方案，请确认后应用。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "生成局部文本修改方案失败。");
    }
  };

  const handleApplyTextEdit = async () => {
    if (!proposalText.trim()) {
      setStatus("error");
      setMessage("没有可应用的修改文本。");
      return;
    }

    if (!isPowerPointHost()) {
      setStatus("error");
      setMessage("当前不在 PowerPoint Add-in 环境中，无法替换选中文本。");
      return;
    }

    setStatus("editing");
    setMessage("正在替换 PowerPoint 当前选中文本。");

    try {
      await replaceSelectedText(proposalText);
      const layoutAdjusted =
        editPlan?.layoutSuggestion.strategy && editPlan.layoutSuggestion.strategy !== "keep"
          ? await adjustSelectedTextBoxLayout(editPlan.layoutSuggestion).catch(() => false)
          : false;
      setSelectedText(proposalText);
      setStatus("inserted");
      setMessage(
        layoutAdjusted
          ? "已替换文字，并按排版建议调整了当前文本框。"
          : editPlan?.layoutSuggestion.strategy && editPlan.layoutSuggestion.strategy !== "keep"
            ? "已替换文字。该修改可能影响文本框排版，请按排版建议手动调整，或使用当前页重排生成替代页。"
          : "已应用到当前选中文本。"
      );
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "应用文本修改失败。");
    }
  };

  const handleImageFileChange = async (file: File | null) => {
    setImageFile(file);
    setImagePlan(null);
    setImageSelectionPlan(null);
    setImageBase64("");

    if (!file) {
      return;
    }

    try {
      setImageBase64(await fileToBase64(file));
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "读取图片失败。");
    }
  };

  const handleUploadAssets = async () => {
    if (assetUploadFiles.length === 0) {
      setStatus("error");
      setMessage("请先选择要保存到素材库的图片或表格文件。");
      return;
    }

    setStatus("saving");
    setMessage("正在保存素材到本地素材库。");

    try {
      await uploadAssets({
        kind: inferAssetKind(assetUploadFiles),
        files: assetUploadFiles,
        notes: assetNotes
      });
      setAssetUploadFiles([]);
      setAssetNotes("");
      await refreshAssets();
      setStatus("idle");
      setMessage("素材已保存到本地素材库。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "保存素材失败。");
    }
  };

  const handlePlanImageEdit = async () => {
    setStatus("editing");
    setMessage("正在生成图片替换和格式调整方案。");
    setImagePlan(null);
    setImageSelectionPlan(null);

    try {
      const plan = await planSelectedImageEdit({
        instruction: imageInstruction,
        imageFileName: imageFile?.name,
        imageMimeType: imageFile?.type
      });
      setImagePlan(plan);
      setStatus("ready");
      setMessage(plan.clarificationQuestion || "已生成图片修改方案，请确认后应用。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "生成图片修改方案失败。");
    }
  };

  const handlePlanImageSelection = async () => {
    setStatus("editing");
    setMessage("正在从素材库中选择最合适的一张图片。");
    setImagePlan(null);
    setImageSelectionPlan(null);

    try {
      const plan = await planImageLibrarySelection({
        instruction: imageInstruction,
        pageText,
        candidates: imageAssets.map((asset) => ({
          id: asset.assetId,
          fileName: asset.sourceFileName,
          mimeType: asset.mimeType,
          notes: asset.notes
        }))
      });
      setImageSelectionPlan(plan);
      setStatus("ready");
      setMessage(plan.selectedImageFileName ? `已选择：${plan.selectedImageFileName}` : "没有选出可用图片。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "素材库选图失败。");
    }
  };

  const handleApplyImageEdit = async () => {
    if (!imageFile || !imageBase64) {
      setStatus("error");
      setMessage("请先选择一张要替换的图片。");
      return;
    }
    await applyImageBase64(imageBase64);
  };

  const handleApplyImageSelection = async () => {
    if (!imageSelectionPlan?.selectedImageId) {
      setStatus("error");
      setMessage("没有可应用的素材库图片。");
      return;
    }

    try {
      const result = await getAssetBase64(imageSelectionPlan.selectedImageId);
      await applyImageBase64(result.base64);
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "应用素材库图片失败。");
    }
  };

  const handleApplyShapeTexts = async () => {
    if (!shapeTextReplacements.trim()) {
      setStatus("error");
      setMessage("请先填写要写入形状的小标题，每行一个。");
      return;
    }

    if (!isPowerPointHost()) {
      setStatus("error");
      setMessage("当前不在 PowerPoint Add-in 环境中，无法修改形状文字。");
      return;
    }

    setStatus("editing");
    setMessage("正在修改当前选中形状中的文字。");

    try {
      const appliedCount = await replaceSelectedShapeTexts(shapeTextReplacements.split(/\r?\n/));
      setStatus(appliedCount > 0 ? "inserted" : "error");
      setMessage(
        appliedCount > 0
          ? `已修改 ${appliedCount} 个选中形状的文字。`
          : "未找到可写入文字的选中形状。请选中圆中的文字形状或扇区形状后再试。"
      );
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "修改形状文字失败。");
    }
  };

  const handleReflow = async () => {
    setStatus("editing");
    setMessage("正在生成当前页替代页。");
    setReflowResult(null);

    try {
      const selectedImageFileName = imageSelectionPlan?.selectedImageFileName ?? imageFile?.name ?? null;
      const reflowed = await reflowCurrentSlide({
        instruction: slideInstruction,
        pageText,
        selectedImageFileName
      });
      setReflowResult(reflowed);
      setStatus("ready");
      setMessage("已生成当前页替代页，可以插入到当前文稿中对比。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "当前页重排失败。");
    }
  };

  const handleInsertReflow = async () => {
    if (!reflowResult?.pptxBase64) {
      setStatus("error");
      setMessage("没有可插入的替代页。");
      return;
    }
    await insertPptx(reflowResult.pptxBase64, "替代页已插入到当前演示文稿。");
  };

  const handleQaCheck = async () => {
    setStatus("editing");
    setMessage("正在检查当前页质量。");
    setQaResult(null);

    try {
      const checked = await checkPageQa({
        pageText,
        instruction: slideInstruction
      });
      setQaResult(checked);
      setStatus("ready");
      setMessage(checked.passed ? "QA 通过，未发现明显问题。" : checked.summary);
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "QA 检查失败。");
    }
  };

  const handleQaAutofix = async () => {
    setStatus("editing");
    setMessage("正在自动修复当前页文本结构。");

    try {
      const fixed = await autofixPageQa({
        pageText,
        instruction: slideInstruction
      });
      setQaResult(fixed);
      setPageText(fixed.fixedPageText);
      setStatus("ready");
      setMessage("已生成自动修复后的页面文本，可继续生成替代页。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "QA 自动修复失败。");
    }
  };

  async function applyImageBase64(base64: string) {
    if (!isPowerPointHost()) {
      setStatus("error");
      setMessage("当前不在 PowerPoint Add-in 环境中，无法替换图片。");
      return;
    }

    setStatus("editing");
    setMessage("正在把图片应用到当前 PowerPoint 选区。");

    try {
      await replaceSelectedImage(base64);
      setStatus("inserted");
      setMessage("图片已应用到当前选区。若 PowerPoint 未保持原图片框，请撤销后改用选中图片占位框再试。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "应用图片修改失败。");
    }
  }

  async function insertPptx(pptxBase64: string, successMessage: string) {
    if (!isPowerPointHost()) {
      setStatus("error");
      setMessage("当前不在 PowerPoint Add-in 环境中，无法执行插入。");
      return;
    }

    setStatus("inserting");
    setMessage("正在插入到当前 PowerPoint。");

    try {
      await insertSlidesFromBase64(pptxBase64);
      setStatus("inserted");
      setMessage(successMessage);
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "插入失败。");
    }
  }

  return (
    <main className="taskpane">
      <header className="appHeader">
        <div>
          <p className="eyebrow">PowerPoint Add-in</p>
          <h1>AI PPT Builder</h1>
        </div>
        <span className={`statusPill status-${status}`}>{statusLabel(status)}</span>
      </header>

      <nav className="modeSwitch" aria-label="工作模式">
        <button className={mode === "generate" ? "active" : ""} type="button" onClick={() => setMode("generate")}>
          生成
        </button>
        <button className={mode === "edit" ? "active" : ""} type="button" onClick={() => setMode("edit")}>
          编辑
        </button>
      </nav>

      <section className="panel compactPanel">
        <h2>AI API</h2>
        <p className="message">
          {aiSettings
            ? `${aiSettings.configured ? "已配置" : "未配置"}：${aiSettings.provider} / ${aiSettings.model} / ${
                aiSettings.baseUrlHost
              }`
            : "未读取到 AI API 设置。"}
        </p>
      </section>

      {mode === "generate" ? (
        <GeneratePanel
          reportFile={reportFile}
          currentReport={currentReport}
          templateFile={templateFile}
          defaultTemplate={defaultTemplate}
          instruction={instruction}
          result={result}
          message={message}
          canGenerate={canGenerate}
          canSaveDefault={canSaveDefault}
          canInsert={canInsert}
          onReportChange={setReportFile}
          onTemplateChange={setTemplateFile}
          onInstructionChange={setInstruction}
          onGenerate={handleGenerate}
          onSaveDefaultTemplate={handleSaveDefaultTemplate}
          onInsert={handleInsert}
        />
      ) : (
        <EditPanel
          editMode={editMode}
          selectedText={selectedText}
          editInstruction={editInstruction}
          editPlan={editPlan}
          proposalText={proposalText}
          imageInstruction={imageInstruction}
          shapeTextReplacements={shapeTextReplacements}
          pageText={pageText}
          imageFile={imageFile}
          imageAssets={imageAssets}
          imagePlan={imagePlan}
          imageSelectionPlan={imageSelectionPlan}
          assetUploadFiles={assetUploadFiles}
          assetNotes={assetNotes}
          slideInstruction={slideInstruction}
          reflowResult={reflowResult}
          qaResult={qaResult}
          message={message}
          canPlanTextEdit={canPlanTextEdit}
          canApplyTextEdit={canApplyTextEdit}
          canPlanImageEdit={canPlanImageEdit}
          canApplyImageEdit={canApplyImageEdit}
          canApplyShapeTexts={canApplyShapeTexts}
          canPlanImageSelection={canPlanImageSelection}
          canApplyImageSelection={canApplyImageSelection}
          canReflow={canReflow}
          canInsertReflow={canInsertReflow}
          canCheckQa={canCheckQa}
          onEditModeChange={setEditMode}
          onSelectedTextChange={setSelectedText}
          onEditInstructionChange={setEditInstruction}
          onProposalTextChange={setProposalText}
          onImageInstructionChange={setImageInstruction}
          onShapeTextReplacementsChange={setShapeTextReplacements}
          onPageTextChange={setPageText}
          onImageFileChange={handleImageFileChange}
          onAssetUploadFilesChange={setAssetUploadFiles}
          onAssetNotesChange={setAssetNotes}
          onUploadAssets={handleUploadAssets}
          onReadSelection={handleReadSelection}
          onPlanTextEdit={handlePlanTextEdit}
          onApplyTextEdit={handleApplyTextEdit}
          onPlanImageEdit={handlePlanImageEdit}
          onApplyImageEdit={handleApplyImageEdit}
          onApplyShapeTexts={handleApplyShapeTexts}
          onPlanImageSelection={handlePlanImageSelection}
          onApplyImageSelection={handleApplyImageSelection}
          onSlideInstructionChange={setSlideInstruction}
          onReflow={handleReflow}
          onInsertReflow={handleInsertReflow}
          onQaCheck={handleQaCheck}
          onQaAutofix={handleQaAutofix}
        />
      )}
    </main>
  );
}

function GeneratePanel(props: {
  reportFile: File | null;
  currentReport: CurrentReportResult | null;
  templateFile: File | null;
  defaultTemplate: DefaultTemplateResult | null;
  instruction: string;
  result: GenerateDeckResult | null;
  message: string;
  canGenerate: boolean;
  canSaveDefault: boolean;
  canInsert: boolean;
  onReportChange: (file: File | null) => void;
  onTemplateChange: (file: File | null) => void;
  onInstructionChange: (instruction: string) => void;
  onGenerate: (event: FormEvent<HTMLFormElement>) => void;
  onSaveDefaultTemplate: () => void;
  onInsert: () => void;
}) {
  return (
    <>
      <form className="panel" onSubmit={props.onGenerate}>
        <FileField
          id="report"
          label="报告文件（完整生成需要）"
          hint={
            props.currentReport?.configured
              ? `不选择则复用当前报告：${props.currentReport.sourceFileName}`
              : "DOCX、Markdown 或文本材料"
          }
          accept=".pdf,.docx,.md,.markdown,.txt"
          file={props.reportFile}
          onChange={props.onReportChange}
        />

        {props.currentReport?.configured ? (
          <div className="templateSummary">
            <strong>当前报告</strong>
            <span>{props.currentReport.sourceFileName}</span>
            <small>
              {props.currentReport.evidenceBlocks} evidence blocks / {props.currentReport.stats?.paragraphs ?? 0}{" "}
              paragraphs / {props.currentReport.stats?.tables ?? 0} tables
            </small>
          </div>
        ) : null}

        <FileField
          id="template"
          label="PPT 模板（可选）"
          hint={
            props.defaultTemplate
              ? `不选择则使用默认模板：${props.defaultTemplate.sourceFileName}`
              : "不选择则使用清华默认模板"
          }
          accept=".pptx"
          file={props.templateFile}
          onChange={props.onTemplateChange}
        />

        <div className="templateActions">
          <button
            className="secondaryButton"
            type="button"
            onClick={props.onSaveDefaultTemplate}
            disabled={!props.canSaveDefault}
          >
            设为默认模板
          </button>
        </div>

        {props.defaultTemplate ? (
          <div className="templateSummary">
            <strong>当前默认模板</strong>
            <span>{props.defaultTemplate.sourceFileName}</span>
            <small>
              {props.defaultTemplate.counts.slides} slides / {props.defaultTemplate.counts.layouts} layouts /{" "}
              {props.defaultTemplate.counts.masters} masters / {props.defaultTemplate.counts.media} media
            </small>
          </div>
        ) : null}

        <label className="field">
          <span>生成要求</span>
          <textarea
            value={props.instruction}
            onChange={(event) => props.onInstructionChange(event.target.value)}
            placeholder="例如：按照报告中的 POI 感知方法生成一页 PPT。"
            rows={4}
          />
        </label>

        <button className="primaryButton" type="submit" disabled={!props.canGenerate}>
          生成 PPT
        </button>
      </form>

      <ResultPanel result={props.result} message={props.message} canInsert={props.canInsert} onInsert={props.onInsert} />
    </>
  );
}

function EditPanel(props: {
  editMode: EditMode;
  selectedText: string;
  editInstruction: string;
  editPlan: TextEditPlan | null;
  proposalText: string;
  imageInstruction: string;
  shapeTextReplacements: string;
  pageText: string;
  imageFile: File | null;
  imageAssets: AssetRecord[];
  imagePlan: ImageEditPlan | null;
  imageSelectionPlan: ImageSelectionPlan | null;
  assetUploadFiles: File[];
  assetNotes: string;
  slideInstruction: string;
  reflowResult: ReflowSlideResult | null;
  qaResult: QaCheckResult | null;
  message: string;
  canPlanTextEdit: boolean;
  canApplyTextEdit: boolean;
  canPlanImageEdit: boolean;
  canApplyImageEdit: boolean;
  canApplyShapeTexts: boolean;
  canPlanImageSelection: boolean;
  canApplyImageSelection: boolean;
  canReflow: boolean;
  canInsertReflow: boolean;
  canCheckQa: boolean;
  onEditModeChange: (mode: EditMode) => void;
  onSelectedTextChange: (text: string) => void;
  onEditInstructionChange: (instruction: string) => void;
  onProposalTextChange: (text: string) => void;
  onImageInstructionChange: (instruction: string) => void;
  onShapeTextReplacementsChange: (text: string) => void;
  onPageTextChange: (text: string) => void;
  onImageFileChange: (file: File | null) => void;
  onAssetUploadFilesChange: (files: File[]) => void;
  onAssetNotesChange: (notes: string) => void;
  onUploadAssets: () => void;
  onReadSelection: () => void;
  onPlanTextEdit: () => void;
  onApplyTextEdit: () => void;
  onPlanImageEdit: () => void;
  onApplyImageEdit: () => void;
  onApplyShapeTexts: () => void;
  onPlanImageSelection: () => void;
  onApplyImageSelection: () => void;
  onSlideInstructionChange: (instruction: string) => void;
  onReflow: () => void;
  onInsertReflow: () => void;
  onQaCheck: () => void;
  onQaAutofix: () => void;
}) {
  return (
    <>
      <section className="panel">
        <div className="chatHeader">
          <h2>局部编辑</h2>
        </div>

        <nav className="triSwitch" aria-label="局部编辑类型">
          <button className={props.editMode === "text" ? "active" : ""} type="button" onClick={() => props.onEditModeChange("text")}>
            文本
          </button>
          <button className={props.editMode === "image" ? "active" : ""} type="button" onClick={() => props.onEditModeChange("image")}>
            图片
          </button>
          <button className={props.editMode === "slide" ? "active" : ""} type="button" onClick={() => props.onEditModeChange("slide")}>
            当前页
          </button>
        </nav>

        {props.editMode === "text" ? (
          <TextEditControls
            selectedText={props.selectedText}
            editInstruction={props.editInstruction}
            canPlanTextEdit={props.canPlanTextEdit}
            onSelectedTextChange={props.onSelectedTextChange}
            onEditInstructionChange={props.onEditInstructionChange}
            onReadSelection={props.onReadSelection}
            onPlanTextEdit={props.onPlanTextEdit}
          />
        ) : null}

        {props.editMode === "image" ? (
          <ImageEditControls
            imageInstruction={props.imageInstruction}
            shapeTextReplacements={props.shapeTextReplacements}
            pageText={props.pageText}
            imageFile={props.imageFile}
            imageAssets={props.imageAssets}
            assetUploadFiles={props.assetUploadFiles}
            assetNotes={props.assetNotes}
            canPlanImageEdit={props.canPlanImageEdit}
            canApplyShapeTexts={props.canApplyShapeTexts}
            canPlanImageSelection={props.canPlanImageSelection}
            onImageInstructionChange={props.onImageInstructionChange}
            onShapeTextReplacementsChange={props.onShapeTextReplacementsChange}
            onPageTextChange={props.onPageTextChange}
            onImageFileChange={props.onImageFileChange}
            onAssetUploadFilesChange={props.onAssetUploadFilesChange}
            onAssetNotesChange={props.onAssetNotesChange}
            onUploadAssets={props.onUploadAssets}
            onPlanImageEdit={props.onPlanImageEdit}
            onApplyShapeTexts={props.onApplyShapeTexts}
            onPlanImageSelection={props.onPlanImageSelection}
          />
        ) : null}

        {props.editMode === "slide" ? (
          <SlideReflowControls
            pageText={props.pageText}
            slideInstruction={props.slideInstruction}
            canReflow={props.canReflow}
            canCheckQa={props.canCheckQa}
            onPageTextChange={props.onPageTextChange}
            onSlideInstructionChange={props.onSlideInstructionChange}
            onReadSelection={props.onReadSelection}
            onReflow={props.onReflow}
            onQaCheck={props.onQaCheck}
            onQaAutofix={props.onQaAutofix}
          />
        ) : null}
      </section>

      <section className="panel">
        <h2>对话确认</h2>
        <p className="message">{props.message}</p>

        {props.editMode === "text" && props.editPlan ? (
          <div className="proposal">
            <strong>建议替换为</strong>
            <textarea value={props.proposalText} onChange={(event) => props.onProposalTextChange(event.target.value)} rows={6} />
            <small>
              {props.editPlan.model ? `模型：${props.editPlan.model}。` : ""}
              {props.editPlan.qa}
            </small>
            <TextLayoutSuggestionCard plan={props.editPlan} />
          </div>
        ) : null}

        {props.editMode === "image" && props.imagePlan ? (
          <div className="proposal">
            <strong>图片方案</strong>
            <p className="message">
              {props.imagePlan.operation === "replace_image"
                ? `替换为：${props.imagePlan.imageFileName}`
                : "当前为格式调整建议，尚未选择替换图片。"}
            </p>
            <small>{props.imagePlan.qa}</small>
          </div>
        ) : null}

        {props.editMode === "image" && props.imageSelectionPlan ? (
          <div className="proposal">
            <strong>素材库选图方案</strong>
            <p className="message">
              {props.imageSelectionPlan.selectedImageFileName
                ? `选择：${props.imageSelectionPlan.selectedImageFileName}（置信度：${props.imageSelectionPlan.confidence}）`
                : "没有选出可用图片。"}
            </p>
            <small>
              {props.imageSelectionPlan.reason}
              <br />
              {props.imageSelectionPlan.qa}
            </small>
          </div>
        ) : null}

        {props.editMode === "slide" && props.reflowResult ? (
          <div className="proposal">
            <strong>替代页方案</strong>
            <p className="message">{props.reflowResult.slideSpec.title}</p>
            <small>{props.reflowResult.qa}</small>
          </div>
        ) : null}

        {props.editMode === "slide" && props.qaResult ? (
          <div className="proposal">
            <strong>QA 检查</strong>
            <p className="message">{props.qaResult.summary}</p>
            <small>
              {props.qaResult.issues
                .map((issue) => `${issue.severity}: ${issue.message} ${issue.suggestion}`)
                .join("\n")}
            </small>
          </div>
        ) : null}

        {props.editMode === "text" ? (
          <button className="secondaryButton" type="button" onClick={props.onApplyTextEdit} disabled={!props.canApplyTextEdit}>
            确认并应用到选区
          </button>
        ) : null}

        {props.editMode === "image" ? (
          <div className="buttonStack">
            <button className="secondaryButton" type="button" onClick={props.onApplyImageEdit} disabled={!props.canApplyImageEdit}>
              确认并替换为指定图片
            </button>
            <button
              className="secondaryButton"
              type="button"
              onClick={props.onApplyImageSelection}
              disabled={!props.canApplyImageSelection}
            >
              确认并插入素材库图片
            </button>
          </div>
        ) : null}

        {props.editMode === "slide" ? (
          <button className="secondaryButton" type="button" onClick={props.onInsertReflow} disabled={!props.canInsertReflow}>
            插入替代页
          </button>
        ) : null}
      </section>
    </>
  );
}

function TextLayoutSuggestionCard(props: { plan: TextEditPlan }) {
  const suggestion = props.plan.layoutSuggestion;
  const strategyLabel: Record<TextEditPlan["layoutSuggestion"]["strategy"], string> = {
    keep: "保持当前文本框",
    expand_height: "增高文本框",
    shift_down: "下移并增高",
    shrink_font: "缩小字号",
    reflow_slide: "生成替代页"
  };
  const applyModeLabel: Record<TextEditPlan["layoutSuggestion"]["applyMode"], string> = {
    advisory: "可直接应用文字",
    requires_shape_api: "需要文本框位置/尺寸调整",
    use_reflow: "建议使用当前页重排"
  };

  return (
    <div className="layoutSuggestion">
      <strong>排版建议：{strategyLabel[suggestion.strategy]}</strong>
      <p>{suggestion.reason}</p>
      <dl>
        <div>
          <dt>长度变化</dt>
          <dd>
            {suggestion.estimatedOriginalChars} → {suggestion.estimatedReplacementChars} 字，约 {suggestion.relativeLengthChange}x
          </dd>
        </div>
        <div>
          <dt>建议调整</dt>
          <dd>
            下移 {suggestion.suggestedDeltaY.toFixed(2)}，高度 {suggestion.suggestedHeightScale.toFixed(2)}x，字号{" "}
            {suggestion.suggestedFontScale.toFixed(2)}x
          </dd>
        </div>
        <div>
          <dt>应用方式</dt>
          <dd>{applyModeLabel[suggestion.applyMode]}</dd>
        </div>
      </dl>
      {suggestion.applyMode !== "advisory" ? (
        <small>应用时会尝试自动移动/缩放当前选中文本框；如果 PowerPoint 选区不支持 shape 调整，请按建议手动调整，或使用当前页重排生成替代页。</small>
      ) : null}
    </div>
  );
}

function TextEditControls(props: {
  selectedText: string;
  editInstruction: string;
  canPlanTextEdit: boolean;
  onSelectedTextChange: (text: string) => void;
  onEditInstructionChange: (instruction: string) => void;
  onReadSelection: () => void;
  onPlanTextEdit: () => void;
}) {
  return (
    <>
      <button className="secondaryButton smallButton readButton" type="button" onClick={props.onReadSelection}>
        读取选区
      </button>
      <label className="field">
        <span>当前选中文本</span>
        <textarea
          value={props.selectedText}
          onChange={(event) => props.onSelectedTextChange(event.target.value)}
          placeholder="在 PowerPoint 中选中文字后点击“读取选区”，也可以手动粘贴文本。"
          rows={5}
        />
      </label>
      <label className="field">
        <span>告诉 agent 如何修改</span>
        <textarea
          value={props.editInstruction}
          onChange={(event) => props.onEditInstructionChange(event.target.value)}
          placeholder="例如：压缩成三条要点，保持原意，适合放在答辩 PPT 中。"
          rows={4}
        />
      </label>
      <button className="primaryButton" type="button" onClick={props.onPlanTextEdit} disabled={!props.canPlanTextEdit}>
        生成文本方案
      </button>
    </>
  );
}

function ImageEditControls(props: {
  imageInstruction: string;
  shapeTextReplacements: string;
  pageText: string;
  imageFile: File | null;
  imageAssets: AssetRecord[];
  assetUploadFiles: File[];
  assetNotes: string;
  canPlanImageEdit: boolean;
  canApplyShapeTexts: boolean;
  canPlanImageSelection: boolean;
  onImageInstructionChange: (instruction: string) => void;
  onShapeTextReplacementsChange: (text: string) => void;
  onPageTextChange: (text: string) => void;
  onImageFileChange: (file: File | null) => void;
  onAssetUploadFilesChange: (files: File[]) => void;
  onAssetNotesChange: (notes: string) => void;
  onUploadAssets: () => void;
  onPlanImageEdit: () => void;
  onApplyShapeTexts: () => void;
  onPlanImageSelection: () => void;
}) {
  return (
    <>
      <label className="field">
        <span>当前页背景信息</span>
        <textarea
          value={props.pageText}
          onChange={(event) => props.onPageTextChange(event.target.value)}
          placeholder="粘贴当前页标题、要点和说明。AI 会据此从素材库中选图。"
          rows={4}
        />
      </label>

      <section className="templateSummary">
        <strong>本地素材库</strong>
        <span>{props.imageAssets.length} 张图片已保存</span>
        <small>{props.imageAssets.slice(0, 4).map((asset) => asset.sourceFileName).join("；") || "暂无图片素材"}</small>
      </section>

      <MultiFileField
        id="asset-library-upload"
        label="保存图片/表格到素材库"
        hint="素材会保存在本地 asset-library，不会提交到 GitHub。"
        accept="image/png,image/jpeg,image/gif,image/bmp,.csv,.xlsx,.xls"
        files={props.assetUploadFiles}
        onChange={(files) => props.onAssetUploadFilesChange(Array.from(files ?? []))}
      />
      <label className="field">
        <span>素材备注（可选）</span>
        <input
          value={props.assetNotes}
          onChange={(event) => props.onAssetNotesChange(event.target.value)}
          placeholder="例如：POI 分类图、实验结果表、研究区地图"
        />
      </label>
      <button className="secondaryButton readButton" type="button" onClick={props.onUploadAssets} disabled={!props.assetUploadFiles.length}>
        保存到素材库
      </button>

      <FileField
        id="replacement-image"
        label="指定替换图片（可选）"
        hint="选择本地图片后，可确认替换当前 PowerPoint 选区。"
        accept="image/png,image/jpeg,image/gif,image/bmp"
        file={props.imageFile}
        onChange={props.onImageFileChange}
      />

      <label className="field">
        <span>图片处理要求</span>
        <textarea
          value={props.imageInstruction}
          onChange={(event) => props.onImageInstructionChange(event.target.value)}
          placeholder="例如：这一页讲 POI 的类别，从素材库中选择最相关的分类图表并插入。"
          rows={4}
        />
      </label>

      <label className="field">
        <span>形状中文字（每行一个）</span>
        <textarea
          value={props.shapeTextReplacements}
          onChange={(event) => props.onShapeTextReplacementsChange(event.target.value)}
          placeholder={"例如：\n静态供给\n体验表达\n动态行为"}
          rows={3}
        />
        <small>用于圆形、扇区、流程框等 PowerPoint 形状。请先选中这些形状，再按行写入对应小标题。</small>
      </label>
      <div className="buttonStack">
        <button className="primaryButton" type="button" onClick={props.onPlanImageEdit} disabled={!props.canPlanImageEdit}>
          生成指定图片方案
        </button>
        <button className="secondaryButton" type="button" onClick={props.onApplyShapeTexts} disabled={!props.canApplyShapeTexts}>
          应用到选中形状文字
        </button>
        <button
          className="primaryButton"
          type="button"
          onClick={props.onPlanImageSelection}
          disabled={!props.canPlanImageSelection}
        >
          AI 从素材库中选图
        </button>
      </div>
    </>
  );
}

function SlideReflowControls(props: {
  pageText: string;
  slideInstruction: string;
  canReflow: boolean;
  canCheckQa: boolean;
  onPageTextChange: (text: string) => void;
  onSlideInstructionChange: (instruction: string) => void;
  onReadSelection: () => void;
  onReflow: () => void;
  onQaCheck: () => void;
  onQaAutofix: () => void;
}) {
  return (
    <>
      <label className="field">
        <span>当前页内容</span>
        <textarea
          value={props.pageText}
          onChange={(event) => props.onPageTextChange(event.target.value)}
          placeholder="粘贴当前页标题、正文、图片说明，或点击读取选区。"
          rows={5}
        />
      </label>
      <button className="secondaryButton smallButton readButton" type="button" onClick={props.onReadSelection}>
        读取选区作为当前页内容
      </button>
      <label className="field">
        <span>重排要求</span>
        <textarea
          value={props.slideInstruction}
          onChange={(event) => props.onSlideInstructionChange(event.target.value)}
          placeholder="例如：按模板重排，左侧放三条要点，右侧预留 POI 分类图。"
          rows={4}
        />
      </label>
      <div className="buttonStack">
        <button className="secondaryButton" type="button" onClick={props.onQaCheck} disabled={!props.canCheckQa}>
          QA 检查
        </button>
        <button className="secondaryButton" type="button" onClick={props.onQaAutofix} disabled={!props.canCheckQa}>
          自动修复文本结构
        </button>
        <button className="primaryButton" type="button" onClick={props.onReflow} disabled={!props.canReflow}>
          生成当前页替代页
        </button>
      </div>
    </>
  );
}

function ResultPanel(props: {
  result: GenerateDeckResult | null;
  message: string;
  canInsert: boolean;
  onInsert: () => void;
}) {
  return (
    <section className="panel">
      <h2>任务状态</h2>
      <p className="message">{props.message}</p>
      {props.result?.qa ? (
        <div className="qaBox">
          <strong>QA</strong>
          <p>{props.result.qa}</p>
        </div>
      ) : null}
      {props.result?.templateReplacement ? (
        <div className="templateSummary">
          <strong>模板替换</strong>
          <span>
            使用第 {props.result.templateReplacement.selectedSlideIndex} 页（
            {props.result.templateReplacement.selectedRole}）
          </span>
          <small>已替换 {props.result.templateReplacement.replacedSlots.length} 个槽位</small>
        </div>
      ) : null}
      {props.result?.deckPlan ? (
        <div className="templateSummary">
          <strong>DeckPlan</strong>
          <span>{props.result.deckPlan.title}</span>
          <small>
            {props.result.deckPlan.slides.length} slide / schema{" "}
            {props.result.deckPlan.validation.schemaValid ? "valid" : "invalid"}
          </small>
        </div>
      ) : null}
      <button className="secondaryButton" type="button" onClick={props.onInsert} disabled={!props.canInsert}>
        插入到 PowerPoint
      </button>
    </section>
  );
}

function FileField(props: {
  id: string;
  label: string;
  hint: string;
  accept: string;
  file: File | null;
  onChange: (file: File | null) => void;
}) {
  return (
    <label className="field" htmlFor={props.id}>
      <span>{props.label}</span>
      <input id={props.id} type="file" accept={props.accept} onChange={(event) => props.onChange(event.target.files?.[0] ?? null)} />
      <small>{props.file ? props.file.name : props.hint}</small>
    </label>
  );
}

function MultiFileField(props: {
  id: string;
  label: string;
  hint: string;
  accept: string;
  files: File[];
  onChange: (files: FileList | null) => void;
}) {
  return (
    <label className="field" htmlFor={props.id}>
      <span>{props.label}</span>
      <input id={props.id} type="file" accept={props.accept} multiple onChange={(event) => props.onChange(event.target.files)} />
      <small>{props.files.length ? props.files.map((file) => file.name).join("；") : props.hint}</small>
    </label>
  );
}

function fileToBase64(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const value = String(reader.result ?? "");
      resolve(value.includes(",") ? value.split(",")[1] : value);
    };
    reader.onerror = () => reject(new Error("无法读取图片文件。"));
    reader.readAsDataURL(file);
  });
}

function inferAssetKind(files: File[]) {
  return files.every((file) => file.type.startsWith("image/")) ? "image" : "table";
}

function statusLabel(status: Status) {
  switch (status) {
    case "generating":
      return "生成中";
    case "ready":
      return "待确认";
    case "inserting":
      return "插入中";
    case "inserted":
      return "已完成";
    case "saving":
      return "保存中";
    case "editing":
      return "编辑中";
    case "error":
      return "错误";
    default:
      return "待开始";
  }
}
