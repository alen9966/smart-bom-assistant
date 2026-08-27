(function () {
  "use strict";

  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => [...document.querySelectorAll(selector)];

  const state = {
    mode: "bom",
    documents: [],
    activeDocumentId: null,
    templates: {},
    exportMode: "separate",
    encodingPreference: "auto",
    detectedEncoding: "",
    encodingLabel: "",
    fileKind: "",
    mojibakeScore: 0,
    file: null,
    workbook: null,
    analysis: null,
    sheet: null,
    mapping: [],
    rows: [],
    mappingCheck: null,
    filter: "all",
    selectedOutputs: new Set(),
    truncated: false,
    toastTimer: null,
    outputDirectory: "",
    outputDirectoryHandle: null
  };

  const editableColumns = [
    { key: "partNumber", label: "物料编码" },
    { key: "description", label: "物料名称" },
    { key: "model", label: "型号 / 规格" },
    { key: "designator", label: "位号" },
    { key: "quantityRaw", label: "原数量" },
    { key: "unit", label: "单位" },
    { key: "vendor", label: "厂家" },
    { key: "source", label: "来源" }
  ];

  const metadataIds = ["projectCode", "productModel", "projectName", "batch", "multiplier", "variant", "department", "pcbRevision", "pcbVendor"];
  const downloadUrls = new Set();

  function base64ToArrayBuffer(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes.buffer;
  }

  function loadBuiltInTemplates() {
    Object.entries(window.DEFAULT_TEMPLATE_FILES || {}).forEach(([key, definition]) => {
      try {
        const workbook = XLSX.read(base64ToArrayBuffer(definition.base64), { type: "array", cellStyles: true, cellFormula: true, cellDates: true, cellNF: true });
        state.templates[key] = {
          file: { name: definition.filename },
          workbook,
          inspection: definition.inspection,
          builtIn: true,
          exactDefinition: definition
        };
      } catch (error) {
        console.error(`默认模板加载失败：${key}`, error);
      }
    });
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]);
  }

  function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function toast(message, type) {
    const element = $("#toast");
    element.textContent = message;
    element.className = `toast show${type === "error" ? " error" : ""}`;
    clearTimeout(state.toastTimer);
    state.toastTimer = setTimeout(() => { element.className = "toast"; }, 3200);
  }

  function verifyRuntimeDependencies() {
    const dependencies = [
      ["中文代码页", window.cptable && window.cptable[936]],
      ["XLSX", window.XLSX],
      ["JSZip", window.JSZip],
      ["BOMCore", window.BOMCore],
      ["ExcelReader", window.ExcelReader],
      ["TemplateExporter", window.TemplateExporter]
    ];
    const missing = dependencies.filter(([, value]) => !value).map(([name]) => name);
    if (!missing.length) return true;

    const message = `程序安装包不完整，缺少运行组件：${missing.join("、")}。请重新解压完整 ZIP，不要只复制 index.html。`;
    console.error(message);
    $("#fileInput").disabled = true;
    $("#generateButton").disabled = true;
    setNotice("warning", message);
    toast(message, "error");
    return false;
  }

  function setStep(step) {
    $$(".progress-steps li").forEach((item) => {
      const itemStep = Number(item.dataset.step);
      item.classList.toggle("active", itemStep === step);
      item.classList.toggle("done", itemStep < step);
      if (itemStep < step) item.querySelector(":scope > span").textContent = "✓";
      else item.querySelector(":scope > span").textContent = String(itemStep);
    });
  }

  function setNotice(kind, message) {
    const notice = $("#recognitionNotice");
    notice.className = `notice ${kind}`;
    notice.textContent = message;
  }

  function saveMetadata() {
    const data = {};
    metadataIds.forEach((id) => { data[id] = $(`#${id}`).value; });
    try { localStorage.setItem("smart-bom-metadata", JSON.stringify(data)); } catch (_) { /* 浏览器禁用存储时忽略 */ }
  }

  function loadMetadata() {
    try {
      const data = JSON.parse(localStorage.getItem("smart-bom-metadata") || "{}");
      metadataIds.forEach((id) => { if (data[id] !== undefined) $(`#${id}`).value = data[id]; });
    } catch (_) { /* 损坏的本地设置不影响使用 */ }
  }

  function metadata() {
    const result = {};
    metadataIds.forEach((id) => { result[id] = $(`#${id}`).value.trim(); });
    result.multiplier = Number($("#multiplier").value || 1);
    result.mode = state.mode;
    result.sourceFile = state.file ? state.file.name : "";
    const all = rowsForAllDocuments();
    const manualVariant = String($("#variant").value || "").trim();
    if (manualVariant) {
      result.variant = manualVariant;
    } else {
      const variantRow = all.find((row) => row && String(row.variant || "").trim());
      result.variant = variantRow ? String(variantRow.variant).trim() : "整机采购";
    }
    return result;
  }

  function effectiveFallbackVariant() {
    return String($("#variant").value || "").trim();
  }

  function effectiveImportMultiplier() {
    return state.mode === "assembly" ? 1 : Number($("#multiplier").value || 1);
  }

  function recalculateDocumentQuantities() {
    captureActiveDocument();
    const multiplier = effectiveImportMultiplier();
    const fallback = effectiveFallbackVariant();
    state.documents.forEach((documentItem) => {
      documentItem.rows = (documentItem.rows || []).map((row) => BOMCore.revalidateRow(row, multiplier, fallback));
    });
    const active = state.documents.find((documentItem) => documentItem.id === state.activeDocumentId);
    if (active) state.rows = active.rows;
  }

  function availableOutputKeys() {
    return Object.entries(BOMCore.OUTPUT_DEFS)
      .filter(([, definition]) => definition.modes.includes(state.mode))
      .map(([key]) => key);
  }

  function renderOutputOptions(resetSelection) {
    const available = availableOutputKeys();
    const primary = available.filter((key) => BOMCore.OUTPUT_DEFS[key].primary);
    const secondary = available.filter((key) => !BOMCore.OUTPUT_DEFS[key].primary);
    if (resetSelection) state.selectedOutputs = new Set(primary.length ? primary : available);
    state.selectedOutputs = new Set([...state.selectedOutputs].filter((key) => available.includes(key)));
    const container = $("#outputList");
    const renderOption = (key) => {
      const definition = BOMCore.OUTPUT_DEFS[key];
      const template = state.templates[key];
      return `<label class="output-option ${definition.primary ? "primary-output" : ""}">
        <input type="checkbox" value="${key}" ${state.selectedOutputs.has(key) ? "checked" : ""}>
        <span><strong>${definition.label}${definition.primary ? '<i class="common-badge">常用</i>' : ""}</strong><small>${definition.description}</small></span>
        <span class="output-actions">
          <em data-output-count="${key}">—</em>
          <button class="template-button ${template ? "has-template" : ""}" data-template-key="${key}" type="button" title="${template ? escapeHtml(template.file.name) : "上传 Excel 模板"}">${template ? (template.builtIn ? "公司默认模板" : escapeHtml(template.file.name)) : "指定模板"}</button>
          ${template && !template.builtIn ? `<button class="template-button" data-template-clear="${key}" type="button">恢复公司默认</button>` : ""}
          <input class="template-input" data-template-input="${key}" type="file" accept=".xlsx,.xls,.xlsm" hidden>
        </span>
      </label>`;
    };
    container.innerHTML = `${primary.length ? `<div class="output-group-title">常用清单 · 默认勾选</div>${primary.map(renderOption).join("")}` : ""}
      ${secondary.length ? `<details class="other-outputs"><summary>其他报表（${secondary.length} 类）</summary><div class="other-output-list">${secondary.map(renderOption).join("")}</div></details>` : ""}`;
    $("#modeExportHint").textContent = state.mode === "bom" ? "默认生成装配、采购和外协阻容三类常用清单" : "默认生成外协阻容备料清单；仅外协阻容数量放大";
    $("#toggleAllOutputs").textContent = available.length && available.every((key) => state.selectedOutputs.has(key)) ? "取消全选" : "全选";
    container.querySelectorAll("input").forEach((input) => input.addEventListener("change", () => {
      if (input.classList.contains("template-input")) return;
      if (input.checked) state.selectedOutputs.add(input.value);
      else state.selectedOutputs.delete(input.value);
      renderHeaderChecklist();
      updateGenerationState();
    }));
    container.querySelectorAll("[data-template-key]").forEach((button) => button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      container.querySelector(`[data-template-input="${button.dataset.templateKey}"]`).click();
    }));
    container.querySelectorAll("[data-template-input]").forEach((input) => input.addEventListener("change", (event) => {
      event.stopPropagation();
      handleTemplateFile(input.dataset.templateInput, input.files[0]);
    }));
    container.querySelectorAll("[data-template-clear]").forEach((button) => button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const definition = (window.DEFAULT_TEMPLATE_FILES || {})[button.dataset.templateClear];
      if (definition) {
        const workbook = XLSX.read(base64ToArrayBuffer(definition.base64), { type: "array", cellStyles: true, cellFormula: true, cellDates: true, cellNF: true });
        state.templates[button.dataset.templateClear] = { file: { name: definition.filename }, workbook, inspection: definition.inspection, builtIn: true, exactDefinition: definition };
      } else delete state.templates[button.dataset.templateClear];
      renderOutputOptions(false);
      toast("已恢复该清单的公司默认模板。");
    }));
    updateOutputCounts();
    renderHeaderChecklist();
  }

  function setMode(mode) {
    state.mode = mode;
    if (state.documents.length) recalculateDocumentQuantities();
    $$(".mode-card").forEach((card) => card.classList.toggle("active", card.dataset.mode === mode));
    renderOutputOptions(true);
    renderHeaderChecklist();
    updateGenerationState();
  }

  function captureActiveDocument() {
    const documentItem = state.documents.find((item) => item.id === state.activeDocumentId);
    if (!documentItem) return;
    Object.assign(documentItem, {
      file: state.file,
      workbook: state.workbook,
      analysis: state.analysis,
      sheet: state.sheet,
      mapping: state.mapping,
      rows: state.rows,
      mappingCheck: state.mappingCheck,
      truncated: state.truncated,
      encodingPreference: state.encodingPreference,
      detectedEncoding: state.detectedEncoding,
      encodingLabel: state.encodingLabel,
      fileKind: state.fileKind,
      mojibakeScore: state.mojibakeScore
    });
  }

  function resetWorkspace() {
    state.documents = [];
    state.file = null;
    state.workbook = null;
    state.analysis = null;
    state.sheet = null;
    state.mapping = [];
    state.rows = [];
    state.mappingCheck = null;
    state.truncated = false;
    state.encodingPreference = "auto";
    state.detectedEncoding = "";
    state.encodingLabel = "";
    state.fileKind = "";
    state.mojibakeScore = 0;
    state.activeDocumentId = null;
    $("#fileInput").value = "";
    $("#dropZone").classList.remove("hidden");
    $("#fileSummary").classList.add("hidden");
    $("#recognitionNotice").classList.add("hidden");
    $("#replaceFile").classList.add("hidden");
    $("#bomFileList").classList.add("hidden");
    $("#bomFileList").innerHTML = "";
    $("#mappingPanel").classList.add("hidden");
    $("#previewPanel").classList.add("hidden");
    renderPcbRevisionHint();
    setStep(1);
    updateOutputCounts();
    updateGenerationState();
  }

  function documentPcbRevision(documentItem) {
    return BOMCore.pcbRevisionFromFilename(documentItem && documentItem.file && documentItem.file.name);
  }

  function renderPcbRevisionHint() {
    const hint = $("#pcbRevisionHint");
    if (!hint) return;
    const revisions = [...new Set(state.documents.map(documentPcbRevision).filter(Boolean))];
    const unnamed = state.documents.filter((item) => !documentPcbRevision(item)).length;
    if (!state.documents.length) {
      hint.textContent = "上传 BOM 后自动从文件名识别版号，一般只需填写板厂。";
      return;
    }
    if (revisions.length) {
      hint.textContent = `已从文件名识别 ${revisions.length} 个版号：${revisions.join("、")}${unnamed ? `；另有 ${unnamed} 个文件未能识别，可在上方补填` : "。分别生成时各用本文件版号，合并时每个版号一行。"}`;
      return;
    }
    hint.textContent = "当前文件名未能识别版号，可手动填写；填了板厂后才会追加印制板。";
  }

  function renderBomFileList() {
    const container = $("#bomFileList");
    renderPcbRevisionHint();
    if (!state.documents.length) {
      container.classList.add("hidden");
      return;
    }
    container.classList.remove("hidden");
    container.innerHTML = state.documents.map((documentItem, index) => {
      const errorCount = (documentItem.rows || []).filter((row) => row.errors.length).length;
      const revision = documentPcbRevision(documentItem);
      return `<button class="bom-file-chip ${documentItem.id === state.activeDocumentId ? "active" : ""}" data-document-id="${documentItem.id}" type="button"><strong>${index + 1}. ${escapeHtml(documentItem.file.name)}</strong><span class="chip-meta">${revision ? `<em class="pcb-rev" title="PCB版号">${escapeHtml(revision)}</em>` : ""}<em>${errorCount ? `${errorCount} 错误` : `${documentItem.rows.length} 行`}</em></span></button>`;
    }).join("");
    container.querySelectorAll("[data-document-id]").forEach((button) => button.addEventListener("click", () => activateDocument(button.dataset.documentId)));
  }

  function activateDocument(id) {
    captureActiveDocument();
    const documentItem = state.documents.find((item) => item.id === id);
    if (!documentItem) return;
    state.activeDocumentId = documentItem.id;
    state.file = documentItem.file;
    state.workbook = documentItem.workbook;
    state.analysis = documentItem.analysis;
    state.sheet = documentItem.sheet;
    state.mapping = documentItem.mapping;
    state.rows = documentItem.rows;
    state.mappingCheck = documentItem.mappingCheck;
    state.truncated = documentItem.truncated;
    state.encodingPreference = documentItem.encodingPreference || "auto";
    state.detectedEncoding = documentItem.detectedEncoding || "";
    state.encodingLabel = documentItem.encodingLabel || "";
    state.fileKind = documentItem.fileKind || "";
    state.mojibakeScore = documentItem.mojibakeScore || 0;
    $("#dropZone").classList.add("hidden");
    $("#fileSummary").classList.remove("hidden");
    $("#replaceFile").classList.remove("hidden");
    $("#mappingPanel").classList.remove("hidden");
    $("#previewPanel").classList.remove("hidden");
    $("#fileName").textContent = documentItem.file.name;
    $("#fileMeta").textContent = `${formatSize(documentItem.file.size)} · ${documentItem.sheet.estimatedDataRows.toLocaleString("zh-CN")} 个候选数据行 · ${documentItem.encodingLabel || "自动编码"} · 批量中共 ${state.documents.length} 个 BOM`;
    populateSheetSelect();
    $("#sheetSelect").value = String(documentItem.sheet.index);
    $("#manualHeaderRow").value = documentItem.sheet.headerRowIndex + 1;
    $("#encodingSelect").value = documentItem.encodingPreference || "auto";
    renderMapping();
    renderRecognitionStats();
    renderPreview();
    renderBomFileList();
    updateOutputCounts();
    updateGenerationState();
    setStep(2);
  }

  function clearFile() {
    captureActiveDocument();
    const removeIndex = state.documents.findIndex((item) => item.id === state.activeDocumentId);
    if (removeIndex >= 0) state.documents.splice(removeIndex, 1);
    if (!state.documents.length) {
      resetWorkspace();
      return;
    }
    const next = state.documents[Math.min(removeIndex, state.documents.length - 1)];
    activateDocument(next.id);
    toast("已从批量任务中移除该 BOM。");
  }

  function populateSheetSelect() {
    const select = $("#sheetSelect");
    select.innerHTML = state.analysis.sheets.map((sheet) => {
      const suffix = sheet.recognized ? ` · ${sheet.matchedCount} 个字段` : " · 未自动识别";
      return `<option value="${sheet.index}">${escapeHtml(sheet.name)}${suffix}</option>`;
    }).join("");
    select.value = String(state.analysis.bestSheetIndex);
  }

  function selectSheet(index, preserveManual) {
    const baseSheet = state.analysis.sheets.find((sheet) => sheet.index === Number(index)) || state.analysis.sheets[0];
    state.sheet = baseSheet.recognized ? baseSheet : BOMCore.reanalyzeWithHeader(baseSheet, 0, 1);
    $("#manualHeaderRow").value = state.sheet.headerRowIndex + 1;
    state.mapping = state.sheet.suggestions.map((suggestion) => suggestion.field || "");
    renderMapping();
    processRows();
    if (!preserveManual) {
      const message = state.sheet.recognized
        ? `已在“${state.sheet.name}”第 ${state.sheet.headerRowIndex + 1}${state.sheet.headerSpan === 2 ? `–${state.sheet.headerRowIndex + 2}` : ""} 行找到表头，共匹配 ${state.sheet.matchedCount} 个字段。`
        : `没有可靠识别该工作表的表头。请在右上角指定表头行，再确认字段对应关系。`;
      setNotice(state.sheet.recognized ? "success" : "warning", message);
    }
    renderRecognitionStats();
  }

  function renderRecognitionStats() {
    if (!state.sheet || !state.analysis) return;
    $("#sheetCount").textContent = state.analysis.sheets.length;
    $("#headerRow").textContent = state.sheet.headerSpan === 2
      ? `${state.sheet.headerRowIndex + 1}-${state.sheet.headerRowIndex + 2}`
      : String(state.sheet.headerRowIndex + 1);
    $("#matchedCount").textContent = state.sheet.matchedCount || 0;
    $("#recognitionScore").textContent = `${state.sheet.recognitionScore || 0}%`;
  }

  function renderMapping() {
    const container = $("#mappingGrid");
    container.innerHTML = state.sheet.headers.map((header, index) => {
      const suggestion = state.sheet.suggestions[index] || { confidence: 0, field: "" };
      const level = suggestion.confidence >= 0.84 ? "high" : suggestion.confidence >= 0.6 ? "medium" : "low";
      const confidenceLabel = suggestion.confidence >= 0.84 ? `高 ${Math.round(suggestion.confidence * 100)}%` : suggestion.confidence >= 0.6 ? `待确认 ${Math.round(suggestion.confidence * 100)}%` : "未识别";
      const options = [`<option value="">忽略此列</option>`, ...BOMCore.FIELD_DEFS.map((field) => `<option value="${field.key}" ${state.mapping[index] === field.key ? "selected" : ""}>${field.label}</option>`)].join("");
      return `<div class="mapping-item ${level}">
        <div class="mapping-source"><strong title="${escapeHtml(header.label)}">${header.column} · ${escapeHtml(header.label)}</strong><span class="confidence">${confidenceLabel}</span></div>
        <select data-column-index="${index}" aria-label="${escapeHtml(header.label)} 对应字段">${options}</select>
      </div>`;
    }).join("");
    container.querySelectorAll("select").forEach((select) => select.addEventListener("change", () => {
      state.mapping[Number(select.dataset.columnIndex)] = select.value;
      renderHeaderChecklist();
      processRows();
    }));
    renderHeaderChecklist();
  }

  function renderHeaderChecklist() {
    const grid = $("#headerCheckGrid");
    const summary = $("#headerCheckSummary");
    if (!grid || !summary || !state.sheet) return;
    const templateMatches = {};
    [...state.selectedOutputs].forEach((key) => {
      const template = state.templates[key];
      if (!template || !template.inspection || !Array.isArray(template.inspection.matches)) return;
      templateMatches[key] = template.inspection.matches.map((match) => ({
        key: match.key,
        label: (() => {
          const sheet = template.workbook && template.workbook.Sheets[template.inspection.sheetName];
          const address = XLSX.utils.encode_cell({ r: template.inspection.headerRowIndex, c: match.colIndex });
          const headerValue = sheet && sheet[address] && sheet[address].v;
          return headerValue || (BOMCore.OUTPUT_COLUMNS[key] && (BOMCore.OUTPUT_COLUMNS[key].find((column) => column[1] === match.key) || [])[0]);
        })()
      }));
    });
    const checks = BOMCore.templateColumnStatus(state.mapping, state.rows, [...state.selectedOutputs], templateMatches);
    const missingCount = checks.filter((item) => item.status === "missing").length;
    summary.textContent = !state.selectedOutputs.size ? "未勾选清单" : `${checks.length} 列 · ${missingCount ? `缺少 ${missingCount}` : "均可生成"}`;
    summary.className = missingCount ? "has-missing" : "all-present";
    if (!checks.length) {
      grid.innerHTML = `<div class="header-check-empty">请先勾选需要生成的清单。</div>`;
      return;
    }
    const statusMeta = {
      present: { icon: "✓", label: "已识别" },
      derived: { icon: "↗", label: "可推导" },
      automatic: { icon: "A", label: "自动填写" },
      optional: { icon: "○", label: "可留空" },
      missing: { icon: "!", label: "缺少" }
    };
    grid.innerHTML = checks.map((item) => `<article class="header-check-item ${item.status}">
      <div class="header-check-status"><span>${statusMeta[item.status].icon}</span><strong>${escapeHtml(item.label)}</strong><em>${statusMeta[item.status].label}</em></div>
      <p><b>${escapeHtml(item.outputLabel)}</b><br>${escapeHtml(item.detail)}</p>
    </article>`).join("");
  }

  function processRows() {
    if (!state.sheet) return;
    const result = BOMCore.extractRows(state.sheet, state.mapping, effectiveImportMultiplier(), { fallbackVariant: effectiveFallbackVariant() });
    state.rows = result.rows.map((row) => ({ ...row, sourceFile: state.file ? state.file.name : "" }));
    state.mappingCheck = result.mappingCheck;
    state.truncated = result.truncated;
    const issue = result.mappingCheck.issues[0];
    const mappingHint = $("#mappingHint");
    if (issue) {
      mappingHint.textContent = issue.message;
      mappingHint.style.color = issue.level === "error" ? "var(--red)" : "var(--amber)";
    } else {
      mappingHint.textContent = `字段对应有效，已读取 ${result.rows.length.toLocaleString("zh-CN")} 行数据。`;
      mappingHint.style.color = "var(--green)";
    }
    renderPreview();
    updateOutputCounts();
    updateGenerationState();
    captureActiveDocument();
  }

  function issueText(row) {
    const issues = [...row.errors, ...row.warnings];
    return issues.map((issue) => issue.message).join("；");
  }

  function filteredRows() {
    return state.rows.map((row, index) => ({ row, index })).filter(({ row }) => {
      if (state.filter === "error") return row.errors.length > 0;
      if (state.filter === "warning") return row.warnings.length > 0 && !row.errors.length;
      return true;
    });
  }

  function renderPreview() {
    const head = $("#previewHead");
    const body = $("#previewBody");
    const errors = state.rows.filter((row) => row.errors.length).length;
    const warnings = state.rows.filter((row) => row.warnings.length && !row.errors.length).length;
    $("#allCount").textContent = state.rows.length;
    $("#errorCount").textContent = errors;
    $("#warningCount").textContent = warnings;
    const columns = state.mode === "assembly"
      ? [{ key: "drawingNo", label: "图号" }, { key: "assemblyName", label: "装配项目" }, ...editableColumns]
      : editableColumns;
    head.innerHTML = `<tr><th>状态</th><th>原行</th>${columns.map((column) => `<th>${column.label}</th>`).join("")}<th>具体问题</th></tr>`;

    if (!state.mappingCheck || !state.mappingCheck.valid) {
      body.innerHTML = `<tr><td colspan="${columns.length + 3}" style="padding:28px;text-align:center;color:var(--muted)">请先完成上方必需字段的对应设置。</td></tr>`;
      $("#rowLimitNote").textContent = "";
      return;
    }
    const visible = filteredRows();
    const limited = visible.slice(0, 200);
    if (!limited.length) {
      body.innerHTML = `<tr><td colspan="${columns.length + 3}" style="padding:28px;text-align:center;color:var(--muted)">${state.rows.length ? "当前筛选下没有数据问题。" : "表头下方没有找到有效数据行。"}</td></tr>`;
    } else {
      body.innerHTML = limited.map(({ row, index }) => {
        const statusText = row.status === "error" ? "!" : row.status === "warning" ? "?" : "✓";
        const issue = issueText(row);
        return `<tr class="${row.status === "error" ? "has-error" : row.status === "warning" ? "has-warning" : ""}">
          <td><span class="row-status ${row.status}" title="${escapeHtml(issue || "校验通过")}">${statusText}</span></td>
          <td>${row.sourceRow}</td>
          ${columns.map((column) => {
            const invalid = row.errors.some((item) => item.field === column.key || (item.field === "identity" && ["partNumber", "description", "model", "drawingNo", "assemblyName"].includes(column.key)));
            return `<td><input class="cell-input ${invalid ? "invalid" : ""}" data-row-index="${index}" data-field="${column.key}" value="${escapeHtml(row[column.key])}" aria-label="第 ${row.sourceRow} 行 ${column.label}"></td>`;
          }).join("")}
          <td><div class="issue-message ${row.status}">${escapeHtml(issue || "—")}</div></td>
        </tr>`;
      }).join("");
      body.querySelectorAll(".cell-input").forEach((input) => input.addEventListener("change", () => {
        const rowIndex = Number(input.dataset.rowIndex);
        const updated = { ...state.rows[rowIndex], [input.dataset.field]: input.value.trim() };
        state.rows[rowIndex] = BOMCore.revalidateRow(updated, effectiveImportMultiplier(), effectiveFallbackVariant());
        captureActiveDocument();
        renderPreview();
        updateOutputCounts();
        updateGenerationState();
      }));
    }
    const notes = [];
    if (visible.length > 200) notes.push(`当前显示前 200 行，共 ${visible.length} 行`);
    else notes.push(`当前显示 ${visible.length} 行`);
    if (state.truncated) notes.push(`原表超过 ${BOMCore.MAX_ROWS.toLocaleString("zh-CN")} 行，超出部分未读取`);
    $("#rowLimitNote").textContent = notes.join("；");
  }

  function rowsForAllDocuments() {
    return state.documents.flatMap((documentItem) => documentItem.id === state.activeDocumentId ? state.rows : (documentItem.rows || []));
  }

  function documentsReadyForExport() {
    return state.documents.every((documentItem) => {
      const rows = documentItem.id === state.activeDocumentId ? state.rows : (documentItem.rows || []);
      const mappingCheck = documentItem.id === state.activeDocumentId ? state.mappingCheck : documentItem.mappingCheck;
      const validRows = rows.filter((row) => !row.errors.length).length;
      return mappingCheck && mappingCheck.valid && rows.length && ($("#skipInvalid").checked ? validRows > 0 : true);
    });
  }

  function updateOutputCounts() {
    const rows = rowsForAllDocuments();
    if (!rows.length) {
      $$('[data-output-count]').forEach((item) => { item.textContent = "—"; });
      return;
    }
    const outputs = BOMCore.buildOutputs(rows, {
      skipInvalid: $("#skipInvalid").checked,
      purchaseMode: $("#purchaseMode").value,
      multiplier: Number($("#multiplier").value || 1)
    });
    availableOutputKeys().forEach((key) => {
      const element = $(`[data-output-count="${key}"]`);
      if (element) element.textContent = `${outputs[key].length} 行${state.documents.length > 1 ? "（合计）" : ""}`;
    });
  }

  function updateGenerationState() {
    const button = $("#generateButton");
    const note = $("#generationNote");
    const allRows = rowsForAllDocuments();
    const validMapping = Boolean(state.documents.length && documentsReadyForExport());
    const errorCount = allRows.filter((row) => row.errors.length).length;
    const canGenerate = Boolean(state.documents.length && validMapping && allRows.length && state.selectedOutputs.size);
    button.disabled = !canGenerate;
    if (!state.file) note.textContent = "请先导入一份料单";
    else if (!validMapping) note.textContent = "请完成必需字段的对应设置";
    else if (!allRows.length) note.textContent = "表头下方没有找到可处理的数据";
    else if (!state.selectedOutputs.size) note.textContent = "请至少勾选一种清单";
    else if (errorCount) note.textContent = `有 ${errorCount} 行错误；生成时将${$("#skipInvalid").checked ? "跳过错误行，并附问题清单" : "保留问题清单"}`;
    else {
      const vendor = $("#pcbVendor").value.trim();
      const revisions = [...new Set(state.documents.map(documentPcbRevision).filter(Boolean))];
      const manualRevision = $("#pcbRevision").value.trim();
      if (vendor && (revisions.length || manualRevision)) {
        const count = revisions.length || 1;
        note.textContent = `${state.documents.length} 个 BOM、${allRows.length} 行数据已就绪；采购清单将追加 ${count} 个印制板版号`;
      } else if (!vendor && (revisions.length || manualRevision)) {
        note.textContent = `${state.documents.length} 个 BOM、${allRows.length} 行数据已就绪；填写 PCB 板厂后才会追加印制板`;
      } else {
        note.textContent = `${state.documents.length} 个 BOM、${allRows.length} 行数据已就绪`;
      }
    }
  }

  async function handleFile(file) {
    if (!file) return;
    const extension = (file.name.split(".").pop() || "").toLowerCase();
    if (!["xlsx", "xls", "xlsm", "csv"].includes(extension)) {
      toast("文件类型不支持，请选择 Excel 或 CSV 文件。", "error");
      return;
    }
    if (file.size > 60 * 1024 * 1024) {
      toast("文件超过 60 MB。建议先删除图片等非料单内容后重试。", "error");
      return;
    }
    state.file = file;
    $("#dropZone").classList.add("hidden");
    $("#fileSummary").classList.remove("hidden");
    $("#replaceFile").classList.remove("hidden");
    $("#fileName").textContent = file.name;
    $("#fileMeta").textContent = `${formatSize(file.size)} · 正在读取并扫描表头…`;
    setNotice("success", "正在扫描所有工作表，请稍候…");
    $("#generateButton").disabled = true;

    try {
      await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 20)));
      const arrayBuffer = await file.arrayBuffer();
      if (extension === "csv") {
        let csvText;
        try {
          csvText = new TextDecoder("utf-8", { fatal: true }).decode(arrayBuffer);
        } catch (_) {
          csvText = new TextDecoder("gb18030").decode(arrayBuffer);
        }
        state.workbook = XLSX.read(csvText.replace(/^\uFEFF/, ""), { type: "string", cellDates: false, cellFormula: false, cellNF: false });
      } else {
        state.workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false, cellFormula: false, cellNF: false });
      }
      state.analysis = BOMCore.analyzeWorkbook(state.workbook);
      populateSheetSelect();
      selectSheet(state.analysis.bestSheetIndex);
      $("#fileMeta").textContent = `${formatSize(file.size)} · ${state.sheet.estimatedDataRows.toLocaleString("zh-CN")} 个候选数据行`;
      $("#mappingPanel").classList.remove("hidden");
      $("#previewPanel").classList.remove("hidden");
      setStep(2);
      if (!state.analysis.hasRecognizedSheet) {
        setNotice("warning", "未能可靠定位表头。请在“表头行”中填写实际行号，再用下拉框确认字段；程序会说明还缺少什么。 ");
      }
      toast("文件读取完成，已给出字段匹配建议。 ");
    } catch (error) {
      console.error(error);
      clearFile();
      const raw = String(error && error.message ? error.message : error);
      const message = /password|encrypted|密码|加密/i.test(raw)
        ? "该 Excel 文件可能受密码保护，请取消保护后再导入。"
        : `无法读取该文件：${raw || "文件内容损坏或格式不受支持"}`;
      toast(message, "error");
    }
  }

  function validateImportFile(file) {
    const extension = (file.name.split(".").pop() || "").toLowerCase();
    if (!["xlsx", "xls", "xlsm", "csv"].includes(extension)) throw new Error("文件类型不支持，请选择 Excel 或 CSV 文件");
    if (file.size > 60 * 1024 * 1024) throw new Error("文件超过 60 MB，请先删除图片等非料单内容");
    return extension;
  }

  async function readWorkbookFile(file, encodingPreference) {
    validateImportFile(file);
    const arrayBuffer = await file.arrayBuffer();
    return ExcelReader.readWorkbook(arrayBuffer, { filename: file.name, encoding: encodingPreference || "auto" });
  }

  async function parseDocument(file, encodingPreference) {
    const readResult = await readWorkbookFile(file, encodingPreference || "auto");
    const workbook = readResult.workbook;
    const analysis = BOMCore.analyzeWorkbook(workbook);
    const baseSheet = analysis.sheets.find((sheet) => sheet.index === analysis.bestSheetIndex) || analysis.sheets[0];
    const sheet = baseSheet.recognized ? baseSheet : BOMCore.reanalyzeWithHeader(baseSheet, 0, 1);
    const mapping = sheet.suggestions.map((suggestion) => suggestion.field || "");
    const extracted = BOMCore.extractRows(sheet, mapping, effectiveImportMultiplier(), { fallbackVariant: effectiveFallbackVariant() });
    return {
      id: `bom-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
      file,
      workbook,
      analysis,
      sheet,
      mapping,
      rows: extracted.rows.map((row) => ({ ...row, sourceFile: file.name })),
      mappingCheck: extracted.mappingCheck,
      truncated: extracted.truncated,
      encodingPreference: encodingPreference || "auto",
      detectedEncoding: readResult.encoding,
      encodingLabel: readResult.encodingLabel,
      fileKind: readResult.kind,
      mojibakeScore: readResult.mojibakeScore,
      encodingSuspicious: readResult.suspicious
    };
  }

  async function handleFiles(fileList) {
    const files = [...(fileList || [])];
    if (!files.length) return;
    $("#dropZone").classList.add("hidden");
    $("#fileSummary").classList.remove("hidden");
    $("#replaceFile").classList.remove("hidden");
    $("#fileName").textContent = `正在加入 ${files.length} 个 BOM…`;
    $("#fileMeta").textContent = "正在逐个扫描工作表和表头";
    setNotice("success", "正在读取批量文件，请稍候…");
    $("#generateButton").disabled = true;
    await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 20)));

    const added = [];
    const failed = [];
    for (const file of files) {
      try {
        const documentItem = await parseDocument(file);
        state.documents.push(documentItem);
        added.push(documentItem);
      } catch (error) {
        console.error(error);
        const raw = String(error && error.message ? error.message : error);
        failed.push(`${file.name}：${/password|encrypted|密码|加密/i.test(raw) ? "文件受密码保护" : raw}`);
      }
    }
    $("#fileInput").value = "";
    if (added.length) {
      activateDocument(added[added.length - 1].id);
      const active = added[added.length - 1];
      if (active.encodingSuspicious) {
        setNotice("warning", `文件已按“${active.encodingLabel}”读取，但内容仍疑似乱码。请在“读取编码”中选择 GBK/GB18030 或其他编码后重读。`);
      } else {
        setNotice(active.analysis.hasRecognizedSheet ? "success" : "warning", active.analysis.hasRecognizedSheet
          ? `批量任务已有 ${state.documents.length} 个 BOM；“${active.file.name}”已按 ${active.encodingLabel} 读取。`
          : `“${active.file.name}”未能可靠定位表头，请尝试切换读取编码，或手动指定表头行和字段。`);
      }
      toast(`已加入 ${added.length} 个 BOM${failed.length ? `，${failed.length} 个失败` : ""}。`);
    } else if (state.documents.length) {
      activateDocument(state.activeDocumentId || state.documents[0].id);
      toast(`没有加入新文件：${failed.join("；")}`, "error");
    } else {
      resetWorkspace();
      toast(`文件读取失败：${failed.join("；")}`, "error");
    }
    renderBomFileList();
  }

  async function handleTemplateFile(outputKey, file) {
    if (!file) return;
    try {
      const readResult = await readWorkbookFile(file, "auto");
      const workbook = readResult.workbook;
      const inspection = BOMCore.inspectTemplate(workbook, outputKey);
      if (!inspection.valid) throw new Error(inspection.message);
      state.templates[outputKey] = { file, workbook, inspection, builtIn: false };
      renderOutputOptions(false);
      toast(`${BOMCore.OUTPUT_DEFS[outputKey].label}模板已识别：${inspection.matchedCount} 个字段。`);
    } catch (error) {
      console.error(error);
      toast(`模板无法使用：${error.message || error}`, "error");
    }
  }

  async function reloadActiveDocumentWithEncoding() {
    captureActiveDocument();
    const index = state.documents.findIndex((item) => item.id === state.activeDocumentId);
    if (index < 0) return;
    const current = state.documents[index];
    const preference = $("#encodingSelect").value;
    $("#reloadEncoding").disabled = true;
    $("#reloadEncoding").textContent = "重读中…";
    setNotice("success", `正在按“${ExcelReader.ENCODINGS[preference].label}”重新读取文件…`);
    try {
      const replacement = await parseDocument(current.file, preference);
      replacement.id = current.id;
      state.documents[index] = replacement;
      activateDocument(replacement.id);
      setNotice(replacement.encodingSuspicious ? "warning" : "success", replacement.encodingSuspicious
        ? `已按 ${replacement.encodingLabel} 重读，但仍检测到异常字符。可以尝试另一种编码。`
        : `已按 ${replacement.encodingLabel} 重读，当前识别 ${replacement.sheet.matchedCount || 0} 个字段。`);
      toast("文件已重新解码并完成表头识别。");
    } catch (error) {
      console.error(error);
      activateDocument(current.id);
      toast(`重新读取失败：${error.message || error}`, "error");
    } finally {
      $("#reloadEncoding").disabled = false;
      $("#reloadEncoding").textContent = "按此编码重读";
    }
  }

  function changeManualHeaderRow() {
    if (!state.sheet || !state.analysis) return;
    const rowNumber = Number($("#manualHeaderRow").value);
    if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > 200) {
      toast("表头行请输入 1–200 之间的整数。", "error");
      return;
    }
    const baseSheet = state.analysis.sheets.find((sheet) => sheet.index === state.sheet.index);
    state.sheet = BOMCore.reanalyzeWithHeader(baseSheet, rowNumber - 1, 1);
    state.mapping = state.sheet.suggestions.map((suggestion) => suggestion.field || "");
    renderMapping();
    processRows();
    renderRecognitionStats();
    setNotice(state.sheet.recognized ? "success" : "warning", state.sheet.recognized
      ? `已按第 ${rowNumber} 行重新识别，共匹配 ${state.sheet.matchedCount} 个字段。请确认黄色项目。`
      : `已切换到第 ${rowNumber} 行，但自动匹配字段较少。请在下方手动指定“数量”和物料标识字段。`);
  }

  function generateWorkbook() {
    try {
      setStep(3);
      const options = {
        selectedOutputs: [...state.selectedOutputs],
        skipInvalid: $("#skipInvalid").checked,
        purchaseMode: $("#purchaseMode").value
      };
      const result = BOMCore.buildWorkbook(state.rows, metadata(), options);
      XLSX.writeFile(result.workbook, result.filename, { compression: true });
      toast(`已生成 ${result.filename}`);
      $("#generationNote").textContent = `已生成 ${options.selectedOutputs.length} 类清单${result.outputs.issues.length ? `，并附 ${result.outputs.issues.length} 条问题说明` : ""}`;
    } catch (error) {
      console.error(error);
      toast(`生成失败：${error.message || error}`, "error");
      setStep(2);
    }
  }

  function safeDownloadPart(value) {
    return String(value || "BOM").replace(/[<>:\"/\\|?*\x00-\x1F]/g, "_").slice(0, 70);
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    downloadUrls.add(url);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    // 不按固定时间释放：部分电脑写盘或杀毒扫描较慢。页面关闭时由
    // 浏览器统一释放，保证每个 Excel 下载全程可读。
  }

  function downloadBytes(bytes, filename, mimeType) {
    const normalized = ZipUtils.bytesOf(bytes);
    const exactBuffer = normalized.buffer.slice(normalized.byteOffset, normalized.byteOffset + normalized.byteLength);
    downloadBlob(new Blob([exactBuffer], { type: mimeType || "application/octet-stream" }), filename);
  }

  function isDesktopApp() {
    return Boolean(window.desktopAPI && window.desktopAPI.isDesktop);
  }

  async function initializeOutputDirectory() {
    const pathInput = $("#outputPath");
    const chooseButton = $("#chooseOutputPath");
    if (!isDesktopApp()) {
      pathInput.value = "系统下载目录";
      chooseButton.disabled = false;
      if (typeof window.showDirectoryPicker === "function") {
        chooseButton.title = "选择生成文件的保存文件夹";
        $("#outputPathHint").textContent = "可选择文件夹；未选择时使用系统下载目录";
      } else {
        chooseButton.title = "当前打开方式不支持网页直接选择保存文件夹";
        $("#outputPathHint").textContent = "当前打开方式只能使用系统下载目录";
      }
      return;
    }
    state.outputDirectory = await window.desktopAPI.getOutputDirectory();
    pathInput.value = state.outputDirectory;
    $("#outputPathHint").textContent = "生成的 Excel 将直接保存到这里";
  }

  async function chooseOutputDirectory() {
    try {
      if (isDesktopApp()) {
        const selected = await window.desktopAPI.chooseOutputDirectory();
        if (!selected) return;
        state.outputDirectory = selected;
        $("#outputPath").value = selected;
        toast(`保存位置已改为：${selected}`);
        return;
      }

      if (typeof window.showDirectoryPicker !== "function") {
        toast("当前浏览器或打开方式不支持选择保存文件夹，请用最新版 Edge/Chrome 打开；文件仍会保存到系统下载目录。", "error");
        return;
      }

      const handle = await window.showDirectoryPicker({ mode: "readwrite" });
      state.outputDirectoryHandle = handle;
      state.outputDirectory = handle.name;
      $("#outputPath").value = `已选择：${handle.name}`;
      $("#outputPathHint").textContent = "生成的 Excel 将直接保存到所选文件夹";
      toast(`保存位置已改为：${handle.name}`);
    } catch (error) {
      if (error && error.name === "AbortError") {
        toast("未更改保存位置");
        return;
      }
      console.error(error);
      toast(`无法选择保存位置：${error.message || error}`, "error");
    }
  }

  async function buildOutputFiles(rows, metadataValue, options) {
    return TemplateExporter.buildFiles(rows, metadataValue, options, BOMCore, XLSX);
  }

  async function downloadOutputFiles(files) {
    const list = Array.isArray(files) ? files : files.files;
    if (isDesktopApp()) {
      const result = await window.desktopAPI.saveOutputFiles({
        files: list.map((file) => ({ filename: file.filename, bytes: ZipUtils.bytesOf(file.bytes) }))
      });
      state.outputDirectory = result.outputDirectory;
      $("#outputPath").value = result.outputDirectory;
      return result;
    }
    if (state.outputDirectoryHandle) {
      const saved = [];
      for (const file of list) {
        const filename = file.filename;
        const normalized = ZipUtils.bytesOf(file.bytes);
        const exactBuffer = normalized.buffer.slice(normalized.byteOffset, normalized.byteOffset + normalized.byteLength);
        const fileHandle = await state.outputDirectoryHandle.getFileHandle(filename, { create: true });
        const writable = await fileHandle.createWritable();
        try {
          await writable.write(exactBuffer);
          await writable.close();
        } catch (error) {
          await writable.abort().catch(() => {});
          throw error;
        }
        saved.push(filename);
      }
      return { outputDirectory: state.outputDirectoryHandle.name, saved };
    }
    for (const file of list) {
      downloadBytes(file.bytes, file.filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      // 给浏览器足够时间创建每个下载任务，避免连续点击被合并或漏掉。
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
    return { outputDirectory: "系统下载目录", saved: list.map((file) => file.filename) };
  }

  function completionMessage(count, result) {
    return isDesktopApp() || state.outputDirectoryHandle
      ? `已保存 ${count} 个 Excel 到：${result.outputDirectory}`
      : `已下载 ${count} 个 Excel`;
  }

  async function generateBatchWorkbook() {
    try {
      captureActiveDocument();
      setStep(3);
      $("#generateButton").disabled = true;
      $("#generationNote").textContent = state.exportMode === "separate" ? "正在逐个生成 Excel…" : "正在合并所有 BOM…";
      const commonOptions = {
        selectedOutputs: [...state.selectedOutputs],
        skipInvalid: $("#skipInvalid").checked,
        purchaseMode: $("#purchaseMode").value,
        templates: state.templates
      };
      const baseMetadata = metadata();

      if (state.exportMode === "combined") {
        const combinedRows = state.documents.flatMap((documentItem) => documentItem.rows || []);
        const combinedMetadata = {
          ...baseMetadata,
          sourceFile: `${state.documents.length} 个 BOM 合并`,
          pcbRevisions: [...new Set(state.documents.map(documentPcbRevision).filter(Boolean))]
        };
        const result = await buildOutputFiles(combinedRows, combinedMetadata, commonOptions);
        const saved = await downloadOutputFiles(result.files);
        toast(completionMessage(result.files.length, saved));
        $("#generationNote").textContent = `${completionMessage(result.files.length, saved)}，共 ${combinedRows.length} 行源数据`;
      } else if (state.documents.length === 1) {
        const documentItem = state.documents[0];
        const result = await buildOutputFiles(documentItem.rows, { ...baseMetadata, sourceFile: documentItem.file.name }, commonOptions);
        const prefix = safeDownloadPart(documentItem.file.name.replace(/\.[^.]+$/, ""));
        result.files.forEach((file) => { file.filename = `${prefix}_${file.filename}`; });
        const saved = await downloadOutputFiles(result.files);
        toast(completionMessage(result.files.length, saved));
        $("#generationNote").textContent = completionMessage(result.files.length, saved);
      } else {
        const outputFiles = [];
        for (const documentItem of state.documents) {
          const result = await buildOutputFiles(documentItem.rows, { ...baseMetadata, sourceFile: documentItem.file.name }, commonOptions);
          const prefix = safeDownloadPart(documentItem.file.name.replace(/\.[^.]+$/, ""));
          result.files.forEach((file) => outputFiles.push({ filename: `${prefix}_${file.filename}`, bytes: file.bytes }));
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
        const saved = await downloadOutputFiles(outputFiles);
        toast(completionMessage(outputFiles.length, saved));
        $("#generationNote").textContent = completionMessage(outputFiles.length, saved);
      }
    } catch (error) {
      console.error(error);
      toast(`生成失败：${error.message || error}`, "error");
      setStep(2);
    } finally {
      updateGenerationState();
    }
  }

  function bindEvents() {
    $$(".mode-card").forEach((card) => card.addEventListener("click", () => setMode(card.dataset.mode)));
    $("#fileInput").addEventListener("change", (event) => handleFiles(event.target.files));
    $("#replaceFile").addEventListener("click", () => $("#fileInput").click());
    $("#clearFile").addEventListener("click", clearFile);

    const dropZone = $("#dropZone");
    ["dragenter", "dragover"].forEach((name) => dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      dropZone.classList.add("dragover");
    }));
    ["dragleave", "drop"].forEach((name) => dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      dropZone.classList.remove("dragover");
    }));
    dropZone.addEventListener("drop", (event) => handleFiles(event.dataTransfer.files));

    $("#sheetSelect").addEventListener("change", (event) => selectSheet(event.target.value));
    $("#manualHeaderRow").addEventListener("change", changeManualHeaderRow);
    $("#reloadEncoding").addEventListener("click", reloadActiveDocumentWithEncoding);
    $("#resetMapping").addEventListener("click", () => {
      state.mapping = state.sheet.suggestions.map((suggestion) => suggestion.field || "");
      renderMapping();
      processRows();
      toast("已恢复本次智能匹配结果。 ");
    });

    $$(".filter-chip").forEach((button) => button.addEventListener("click", () => {
      state.filter = button.dataset.filter;
      $$(".filter-chip").forEach((item) => item.classList.toggle("active", item === button));
      renderPreview();
    }));

    metadataIds.forEach((id) => $(`#${id}`).addEventListener("change", () => {
      saveMetadata();
      const needRecalc = (id === "multiplier" || id === "variant") && state.documents.length;
      if (id === "multiplier") {
        const multiplier = Number($("#multiplier").value);
        if (!Number.isFinite(multiplier) || multiplier <= 0) {
          toast("生产数量必须大于 0。", "error");
          updateGenerationState();
          return;
        }
      }
      if (needRecalc) {
        const effectiveMult = state.mode === "assembly" ? 1 : Number($("#multiplier").value || 1);
        const fallback = effectiveFallbackVariant();
        captureActiveDocument();
        state.documents.forEach((documentItem) => {
          documentItem.rows = documentItem.rows.map((row) => BOMCore.revalidateRow(row, effectiveMult, fallback));
        });
        const active = state.documents.find((documentItem) => documentItem.id === state.activeDocumentId);
        if (active) state.rows = active.rows;
        renderPreview();
      }
      renderHeaderChecklist();
      updateOutputCounts();
      updateGenerationState();
    }));

    ["purchaseMode", "skipInvalid"].forEach((id) => $(`#${id}`).addEventListener("change", () => {
      updateOutputCounts();
      updateGenerationState();
    }));

    $("#toggleAllOutputs").addEventListener("click", () => {
      const available = availableOutputKeys();
      const allSelected = available.every((key) => state.selectedOutputs.has(key));
      state.selectedOutputs = new Set(allSelected ? [] : available);
      renderOutputOptions(false);
      renderHeaderChecklist();
      $("#toggleAllOutputs").textContent = allSelected ? "全选" : "取消全选";
      updateGenerationState();
    });
    $$('input[name="exportMode"]').forEach((input) => input.addEventListener("change", () => {
      state.exportMode = input.value;
      $$(".export-mode-option").forEach((option) => option.classList.toggle("active", option.contains(input) && input.checked));
      updateGenerationState();
    }));
    $("#generateButton").addEventListener("click", generateBatchWorkbook);
    $("#chooseOutputPath").addEventListener("click", chooseOutputDirectory);

    const helpDialog = $("#helpDialog");
    $("#openHelp").addEventListener("click", () => helpDialog.showModal());
    $("#closeHelp").addEventListener("click", () => helpDialog.close());
    helpDialog.addEventListener("click", (event) => { if (event.target === helpDialog) helpDialog.close(); });
    const changelogDialog = $("#changelogDialog");
    $("#openChangelog").addEventListener("click", () => changelogDialog.showModal());
    $("#closeChangelog").addEventListener("click", () => changelogDialog.close());
    changelogDialog.addEventListener("click", (event) => { if (event.target === changelogDialog) changelogDialog.close(); });
  }

  bindEvents();
  if (verifyRuntimeDependencies()) {
    initializeOutputDirectory().catch((error) => {
      console.error(error);
      toast(`无法读取保存位置：${error.message || error}`, "error");
    });
    loadMetadata();
    loadBuiltInTemplates();
    renderOutputOptions(true);
    updateGenerationState();
  }
})();
