import { FormEvent, useEffect, useMemo, useState } from "react";
import {
  DefaultTemplateResult,
  generateDeck,
  GenerateDeckResult,
  getDefaultTemplate,
  saveDefaultTemplate
} from "./api/decks";
import { insertSlidesFromBase64, isPowerPointHost } from "./office/powerpoint";

type Status = "idle" | "generating" | "ready" | "inserting" | "inserted" | "saving" | "error";

export function App() {
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [defaultTemplate, setDefaultTemplate] = useState<DefaultTemplateResult | null>(null);
  const [instruction, setInstruction] = useState("");
  const [result, setResult] = useState<GenerateDeckResult | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [message, setMessage] = useState("等待上传报告。未选择模板时会使用默认模板。");

  useEffect(() => {
    getDefaultTemplate()
      .then((template) => setDefaultTemplate(template))
      .catch((error) => {
        setStatus("error");
        setMessage(error instanceof Error ? error.message : "读取默认模板失败。");
      });
  }, []);

  const canGenerate = useMemo(() => reportFile !== null && status !== "generating", [reportFile, status]);
  const canSaveDefault = templateFile !== null && status !== "saving";
  const canInsert = Boolean(result?.pptxBase64) && status !== "inserting";

  const handleGenerate = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!reportFile) {
      setStatus("error");
      setMessage("请先选择报告文件。");
      return;
    }

    setStatus("generating");
    setMessage(templateFile ? "正在使用上传模板生成测试 PPT。" : "正在使用默认模板生成测试 PPT。");
    setResult(null);

    try {
      const generated = await generateDeck({
        reportFile,
        templateFile,
        instruction
      });
      setResult(generated);
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

    if (!isPowerPointHost()) {
      setStatus("error");
      setMessage("当前不在 PowerPoint Add-in 环境中，无法执行插入。");
      return;
    }

    setStatus("inserting");
    setMessage("正在插入到当前 PowerPoint。");

    try {
      await insertSlidesFromBase64(result.pptxBase64);
      setStatus("inserted");
      setMessage("已插入到当前演示文稿。");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "插入失败。");
    }
  };

  return (
    <main className="taskpane">
      <header className="appHeader">
        <div>
          <p className="eyebrow">PowerPoint Add-in</p>
          <h1>AI PPT Builder</h1>
        </div>
        <span className={`statusPill status-${status}`}>{statusLabel(status)}</span>
      </header>

      <form className="panel" onSubmit={handleGenerate}>
        <FileField
          id="report"
          label="报告文件"
          hint="PDF、DOCX、Markdown 或文本材料"
          accept=".pdf,.docx,.md,.markdown,.txt"
          file={reportFile}
          onChange={setReportFile}
        />

        <FileField
          id="template"
          label="PPT 模板（可选）"
          hint={defaultTemplate ? `不选择则使用默认模板：${defaultTemplate.sourceFileName}` : "不选择则使用清华默认模板"}
          accept=".pptx"
          file={templateFile}
          onChange={setTemplateFile}
        />

        <div className="templateActions">
          <button className="secondaryButton" type="button" onClick={handleSaveDefaultTemplate} disabled={!canSaveDefault}>
            设为默认模板
          </button>
        </div>

        {defaultTemplate ? (
          <div className="templateSummary">
            <strong>当前默认模板</strong>
            <span>{defaultTemplate.sourceFileName}</span>
            <small>
              {defaultTemplate.counts.slides} slides / {defaultTemplate.counts.layouts} layouts /{" "}
              {defaultTemplate.counts.masters} masters / {defaultTemplate.counts.media} media
            </small>
          </div>
        ) : null}

        <label className="field">
          <span>生成要求</span>
          <textarea
            value={instruction}
            onChange={(event) => setInstruction(event.target.value)}
            placeholder="例如：按照报告中的 POI 感知方法生成一页 PPT。"
            rows={4}
          />
        </label>

        <button className="primaryButton" type="submit" disabled={!canGenerate}>
          生成测试 PPT
        </button>
      </form>

      <section className="panel">
        <h2>任务状态</h2>
        <p className="message">{message}</p>

        {result?.qa ? (
          <div className="qaBox">
            <strong>QA</strong>
            <p>{result.qa}</p>
          </div>
        ) : null}

        <button className="secondaryButton" type="button" onClick={handleInsert} disabled={!canInsert}>
          插入到 PowerPoint
        </button>
      </section>
    </main>
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
      <input
        id={props.id}
        type="file"
        accept={props.accept}
        onChange={(event) => props.onChange(event.target.files?.[0] ?? null)}
      />
      <small>{props.file ? props.file.name : props.hint}</small>
    </label>
  );
}

function statusLabel(status: Status) {
  switch (status) {
    case "generating":
      return "生成中";
    case "ready":
      return "可插入";
    case "inserting":
      return "插入中";
    case "inserted":
      return "已完成";
    case "saving":
      return "保存中";
    case "error":
      return "错误";
    default:
      return "待开始";
  }
}
