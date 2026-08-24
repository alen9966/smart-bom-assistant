(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("xlsx-js-style"));
  } else {
    root.BOMCore = factory(root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX) {
  "use strict";

  const MAX_ROWS = 50000;
  const MAX_SCAN_ROWS = 40;
  const MAX_SCAN_COLS = 80;

  const FIELD_DEFS = [
    { key: "category", label: "类别", aliases: ["类别", "分类", "物料类别", "物料分类", "class", "category", "part type", "type"] },
    { key: "designator", label: "位号", aliases: ["位号", "元件位号", "参考位号", "器件位号", "reference", "designator", "refdes", "ref des", "location"] },
    { key: "partNumber", label: "物料编码", aliases: ["物料编码", "物料编号", "物料号", "物资编码", "物资编号", "编码", "存货编码", "零件编号", "part number", "part no", "partnumber", "pn", "item code", "material code", "sku"] },
    { key: "description", label: "物料名称", aliases: ["物料名称", "名称", "元件名称", "零件名称", "品名", "描述", "器件描述", "description", "comment", "part name", "item name", "component"] },
    { key: "model", label: "型号 / 规格", aliases: ["型号", "规格", "型号规格", "型号/规格", "规格型号", "参数", "数值", "型号value", "value", "model", "spec", "specification", "device"] },
    { key: "footprint", label: "封装", aliases: ["封装", "封装形式", "PCB封装", "器件封装", "footprint", "package", "pcb footprint"] },
    { key: "quality", label: "质量等级", aliases: ["质量等级", "质量", "质量级别", "等级", "质量等级quality", "quality", "quality level", "grade"] },
    { key: "quantity", label: "数量", aliases: ["数量", "用量", "单机用量", "需求数量", "总数量", "装配数量", "数量quantity", "qty", "quantity", "count", "amount"] },
    { key: "unit", label: "单位", aliases: ["单位", "计量单位", "基本单位", "unit", "uom"] },
    { key: "vendor", label: "厂家 / 供应商", aliases: ["厂家", "厂商", "制造商", "供应商", "品牌", "厂家vendor", "vendor", "supplier", "manufacturer", "mfr", "brand"] },
    { key: "manufacturerPartNumber", label: "厂家料号", aliases: ["厂家料号", "制造商料号", "厂商型号", "厂家型号", "mpn", "manufacturer part number", "mfr part"] },
    { key: "source", label: "来源 / 自制外购", aliases: ["来源", "物料来源", "采购属性", "自制外购", "自制/外购", "外购", "物料属性", "供货方式", "source", "make or buy", "make/buy", "procurement type"] },
    { key: "drawingNo", label: "图号", aliases: ["图号", "图纸号", "零件图号", "drawing", "drawing no", "drawing number"] },
    { key: "assemblyName", label: "装配项目", aliases: ["装配清单", "装配项目", "装配名称", "组件名称", "组件", "总成", "assembly", "assembly name", "subassembly"] },
    { key: "variant", label: "装配变量", aliases: ["装配变量", "变量", "变量名", "variant", "variation", "bom variant"] },
    { key: "remark", label: "备注", aliases: ["备注", "说明", "注释", "特殊要求", "remark", "remarks", "note", "notes"] }
  ];

  const FIELD_MAP = Object.fromEntries(FIELD_DEFS.map((field) => [field.key, field]));

  const OUTPUT_DEFS = {
    assembly: { label: "装配清单", description: "逐项保留位号、图号与装配来源", modes: ["bom"], primary: true },
    picking: { label: "领料单", description: "按物料合并数量并汇总位号", modes: ["bom", "assembly"] },
    procurement: { label: "采购清单", description: "仅汇总判定为外购的物料", modes: ["bom"], primary: true },
    detail: { label: "明细表", description: "标准化后的完整可追溯明细", modes: ["bom", "assembly"] },
    purchased: { label: "外购件汇总表", description: "按厂家、型号归集外购需求", modes: ["bom", "assembly"] },
    subcontract: { label: "外协阻容备料清单", description: "仅筛选电阻、电容，并按生产数量计算", modes: ["bom", "assembly"], primary: true }
  };

  const HEADER_REQUIREMENTS = [
    { key: "category", outputs: ["picking", "procurement", "detail", "purchased", "subcontract"], impact: "影响领料单和采购申请表的加粗分类行；缺少时只能按位号、名称和型号推断，分类可能不准确。" },
    { key: "designator", outputs: ["assembly", "picking", "detail", "purchased", "subcontract"], impact: "装配清单无法显示位号，领料汇总也不能追溯到具体元件位置。" },
    { key: "partNumber", outputs: ["assembly", "picking", "procurement", "detail", "purchased", "subcontract"], impact: "物料编码/代号列会为空；合并物料时只能改用型号或名称，存在错并风险。" },
    { key: "description", outputs: ["assembly", "picking", "procurement", "detail", "purchased", "subcontract"], impact: "装配清单的名称可能为空；阻容会尝试从型号推导，其他物料无法保证补齐。" },
    { key: "model", outputs: ["assembly", "picking", "procurement", "detail", "purchased", "subcontract"], impact: "型号/规格列会为空，并影响物料识别、阻容名称封装推导和汇总准确性。" },
    { key: "footprint", outputs: ["assembly"], impact: "装配清单的封装列可能为空；常用阻容会尝试从型号推导，其他器件无法保证补齐。" },
    { key: "quality", outputs: ["assembly", "picking", "procurement", "detail", "purchased", "subcontract"], impact: "领料单和采购申请表的质量等级列会为空。" },
    { key: "quantity", outputs: ["assembly", "picking", "procurement", "detail", "purchased", "subcontract"], blocking: true, impact: "无法计算单机数量、领料数量、装机总数和采购总数，会阻止生成清单。" },
    { key: "unit", outputs: ["picking", "procurement", "detail", "purchased", "subcontract"], impact: "单位缺失时统一按“件”处理，可能与原始计量单位不一致。" },
    { key: "vendor", outputs: ["assembly", "picking", "procurement", "detail", "purchased", "subcontract"], impact: "厂家/供应商列会为空，并影响采购清单及外购件按厂家汇总。" },
    { key: "manufacturerPartNumber", outputs: ["procurement", "detail", "purchased", "subcontract"], impact: "厂家料号无法输出，采购核对和外购件追溯信息不完整。" },
    { key: "source", outputs: ["procurement", "detail", "purchased"], impact: "无法直接判断自制或外购；程序只能结合类别、厂家等信息推断，采购清单可能多项或漏项。" },
    { key: "drawingNo", outputs: ["assembly", "picking", "detail"], impact: "装配清单和领料单的图号列会为空。" },
    { key: "assemblyName", outputs: ["assembly", "detail"], impact: "无法保留原 BOM 中的装配项目/组件名称，装配层级追溯会不完整。" },
    { key: "variant", outputs: ["assembly"], impact: "装配清单的装配变量列会为空；建议在原始 BOM 中添加 Variant 或装配变量列，或在页面“项目信息”的“装配变量”框中手动填写。" },
    { key: "remark", outputs: ["assembly", "picking", "procurement", "detail", "subcontract"], impact: "备注、特殊要求和缺件说明不会带入导出清单。" }
  ];

  function headerRequirementStatus(mapping, mode, selectedOutputs) {
    const mapped = new Set((mapping || []).filter(Boolean));
    const selected = new Set((selectedOutputs || []).filter((key) => OUTPUT_DEFS[key] && OUTPUT_DEFS[key].modes.includes(mode || "bom")));
    return HEADER_REQUIREMENTS.filter((requirement) => requirement.outputs.some((key) => selected.has(key))).map((requirement) => ({
      ...requirement,
      label: FIELD_MAP[requirement.key].label,
      present: mapped.has(requirement.key),
      affectedOutputs: requirement.outputs.filter((key) => selected.has(key)).map((key) => OUTPUT_DEFS[key].label)
    }));
  }

  const TEMPLATE_FIELD_RULES = {
    _index: { label: "序号", automatic: "程序按输出顺序自动填写" },
    _batch: { label: "批次", automatic: "由右侧项目信息中的批次自动填写" },
    _zero: { label: "固定数值 0", automatic: "由模板规则自动填写" },
    _yes: { label: "是否采购", automatic: "由模板规则自动填写“是”" },
    _electricalAssembly: { label: "装配方式", automatic: "由模板规则自动填写“电装”" },
    componentType: { label: "元件类型", sources: ["category", "description", "designator", "model"], derived: "由类别、名称、位号或型号判断电阻/电容", optional: false },
    modelOrDescription: { label: "型号 / 名称", sources: ["model", "description", "partNumber"], derived: "按型号、名称、物料编码依次取值", optional: false },
    baseQuantity: { label: "单机数量", sources: ["quantity"], derived: "由原数量解析得到", blocking: true },
    quantity: { label: "总数量", sources: ["quantity"], derived: "由原数量和生产数量计算", blocking: true },
    designatorSummary: { label: "位号汇总", sources: ["designator"], derived: "由位号汇总得到", optional: true },
    traceSummary: { label: "来源追溯", automatic: "由源文件、工作表和原行号自动生成" },
    purchaseBasis: { label: "外购判断依据", sources: ["source", "category", "vendor"], derived: "由来源、类别和厂家综合判断", optional: true },
    sourceFile: { label: "源文件", automatic: "由导入文件名自动生成" },
    sourceSheet: { label: "原工作表", automatic: "由导入工作表自动生成" },
    sourceRow: { label: "原行号", automatic: "由原始行位置自动生成" }
  };

  function ruleForTemplateKey(key) {
    if (TEMPLATE_FIELD_RULES[key]) return TEMPLATE_FIELD_RULES[key];
    const field = FIELD_MAP[key];
    const optionalKeys = new Set(["category", "designator", "partNumber", "description", "model", "footprint", "quality", "unit", "vendor", "manufacturerPartNumber", "source", "drawingNo", "assemblyName", "variant", "remark"]);
    const derived = key === "description" ? "可由型号或类别补齐"
      : key === "footprint" ? "常用阻容可由型号、类别或位号推导"
        : key === "unit" ? "缺少时程序按“件”处理"
          : key === "variant" ? "缺少时程序按“整机采购”处理"
            : "";
    return { label: field ? field.label : key, sources: [key], derived, optional: optionalKeys.has(key) };
  }

  function valueExists(rows, key) {
    return (rows || []).some((row) => {
      if (key === "baseQuantity") return Number.isFinite(Number(row.baseQuantity)) && Number(row.baseQuantity) > 0;
      if (key === "quantity") return Number.isFinite(Number(row.quantity)) && Number(row.quantity) >= 0;
      return text(row && row[key]) !== "";
    });
  }

  function templateColumnStatus(mapping, rows, selectedOutputs, templateMatches) {
    const mapped = new Set((mapping || []).filter(Boolean));
    const results = [];
    for (const outputKey of selectedOutputs || []) {
      if (!OUTPUT_DEFS[outputKey]) continue;
      const matches = templateMatches && Array.isArray(templateMatches[outputKey])
        ? templateMatches[outputKey]
        : (OUTPUT_COLUMNS[outputKey] || []).map(([label, key]) => ({ label, key }));
      for (const match of matches) {
        const key = match.key;
        if (!key) continue;
        const rule = ruleForTemplateKey(key);
        const sources = rule.sources || [];
        const directSources = sources.filter((source) => mapped.has(source));
        const hasOutputValue = valueExists(rows, key)
          || (key === "modelOrDescription" && sources.some((source) => valueExists(rows, source)))
          || (key === "componentType" && (rows || []).some((row) => resistorCapacitorType(row)));
        let status;
        let detail;
        if (rule.automatic) {
          status = "automatic";
          detail = rule.automatic;
        } else if (mapped.has(key)) {
          status = "present";
          detail = `原表已识别：${FIELD_MAP[key] ? FIELD_MAP[key].label : rule.label}`;
        } else if (hasOutputValue) {
          status = "derived";
          const sourceLabels = directSources.map((source) => FIELD_MAP[source] ? FIELD_MAP[source].label : source);
          detail = `${rule.derived || "程序已根据现有数据生成"}${sourceLabels.length ? `；来源：${sourceLabels.join("、")}` : ""}`;
        } else if (rule.blocking) {
          status = "missing";
          detail = `模板此列无法生成；需要原表字段：${sources.map((source) => FIELD_MAP[source] ? FIELD_MAP[source].label : source).join("、")}`;
        } else {
          status = "optional";
          detail = `${rule.derived ? `${rule.derived}；当前没有可用值。` : "原表没有对应值。"}模板允许此列留空，不影响生成。`;
        }
        results.push({ outputKey, outputLabel: OUTPUT_DEFS[outputKey].label, key, label: match.label || rule.label, status, detail });
      }
    }
    return results;
  }

  function text(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  function normalizeText(value) {
    return text(value)
      .normalize("NFKC")
      .toLowerCase()
      .replace(/[\s\r\n\t_\-—–/\\|:：·.,，。()（）\[\]【】<>《》]/g, "");
  }

  function bigrams(value) {
    const normalized = normalizeText(value);
    if (normalized.length < 2) return normalized ? [normalized] : [];
    const result = [];
    for (let i = 0; i < normalized.length - 1; i += 1) result.push(normalized.slice(i, i + 2));
    return result;
  }

  function diceSimilarity(a, b) {
    const aa = bigrams(a);
    const bb = bigrams(b);
    if (!aa.length || !bb.length) return 0;
    const counts = new Map();
    aa.forEach((item) => counts.set(item, (counts.get(item) || 0) + 1));
    let overlap = 0;
    bb.forEach((item) => {
      if ((counts.get(item) || 0) > 0) {
        overlap += 1;
        counts.set(item, counts.get(item) - 1);
      }
    });
    return (2 * overlap) / (aa.length + bb.length);
  }

  function aliasScore(header, alias) {
    const h = normalizeText(header);
    const a = normalizeText(alias);
    if (!h || !a) return 0;
    if (h === a) return 1;
    if (h.length >= 2 && a.length >= 2 && (h.includes(a) || a.includes(h))) {
      const ratio = Math.min(h.length, a.length) / Math.max(h.length, a.length);
      return 0.78 + ratio * 0.17;
    }
    return diceSimilarity(h, a) * 0.82;
  }

  function rankHeader(header) {
    const ranked = FIELD_DEFS.map((field) => {
      const confidence = Math.max(...field.aliases.map((alias) => aliasScore(header, alias)));
      return { field: field.key, label: field.label, confidence };
    }).sort((a, b) => b.confidence - a.confidence);
    return ranked;
  }

  function columnName(index) {
    let result = "";
    let value = index + 1;
    while (value > 0) {
      value -= 1;
      result = String.fromCharCode(65 + (value % 26)) + result;
      value = Math.floor(value / 26);
    }
    return result;
  }

  function sheetRows(sheet) {
    if (!sheet || !sheet["!ref"]) return [];
    return XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: true,
      range: 0
    }).slice(0, MAX_ROWS + MAX_SCAN_ROWS);
  }

  function combinedHeader(rows, rowIndex, span, colIndex) {
    const pieces = [];
    for (let offset = 0; offset < span; offset += 1) {
      const value = text((rows[rowIndex + offset] || [])[colIndex]);
      if (value && !pieces.includes(value)) pieces.push(value);
    }
    return pieces.join(" / ");
  }

  function evaluateHeaderCandidate(rows, rowIndex, span) {
    const rowA = rows[rowIndex] || [];
    const rowB = span === 2 ? (rows[rowIndex + 1] || []) : [];
    const colCount = Math.min(MAX_SCAN_COLS, Math.max(rowA.length, rowB.length));
    const candidates = [];
    let numericCells = 0;
    let filledCells = 0;

    for (let col = 0; col < colCount; col += 1) {
      const header = combinedHeader(rows, rowIndex, span, col);
      if (!header) continue;
      filledCells += 1;
      if (/^[\d.,%+\-\s]+$/.test(header)) numericCells += 1;
      const ranked = rankHeader(header);
      if (ranked[0].confidence >= 0.53) {
        candidates.push({ colIndex: col, header, ...ranked[0], alternatives: ranked.slice(0, 3) });
      }
    }

    const bestByField = new Map();
    candidates.forEach((candidate) => {
      const current = bestByField.get(candidate.field);
      if (!current || candidate.confidence > current.confidence) bestByField.set(candidate.field, candidate);
    });
    const unique = [...bestByField.values()];
    const hasQuantity = unique.some((item) => item.field === "quantity");
    const hasIdentity = unique.some((item) => ["partNumber", "model", "description", "drawingNo", "assemblyName"].includes(item.field));
    const numericRatio = filledCells ? numericCells / filledCells : 1;
    const confidenceSum = unique.reduce((sum, item) => sum + item.confidence, 0);
    const score = confidenceSum * 10 + unique.length * 2.8 + (hasQuantity ? 4 : 0) + (hasIdentity ? 3 : 0) - numericRatio * 8 - (span === 2 ? 0.4 : 0);
    return { rowIndex, span, score, unique, candidates, filledCells, numericRatio };
  }

  function analyzeSheet(sheet, name, index) {
    const rows = sheetRows(sheet);
    let best = null;
    const scanCount = Math.min(MAX_SCAN_ROWS, rows.length);
    for (let rowIndex = 0; rowIndex < scanCount; rowIndex += 1) {
      [1, 2].forEach((span) => {
        if (rowIndex + span > rows.length) return;
        const candidate = evaluateHeaderCandidate(rows, rowIndex, span);
        if (!best || candidate.score > best.score) best = candidate;
      });
    }

    if (!best || best.unique.length < 2 || best.score < 11) {
      return {
        name,
        index,
        rows,
        recognized: false,
        score: best ? best.score : 0,
        headerRowIndex: 0,
        headerSpan: 1,
        dataStartIndex: 1,
        headers: [],
        suggestions: []
      };
    }

    const headerRowIndex = best.rowIndex;
    const headerSpan = best.span;
    const nextRows = rows.slice(headerRowIndex, Math.min(rows.length, headerRowIndex + headerSpan + 20));
    const colCount = Math.min(MAX_SCAN_COLS, Math.max(0, ...nextRows.map((row) => row.length)));
    const headers = [];
    const suggestions = [];

    for (let col = 0; col < colCount; col += 1) {
      const rawHeader = combinedHeader(rows, headerRowIndex, headerSpan, col) || `未命名列 ${columnName(col)}`;
      const ranked = rankHeader(rawHeader);
      headers.push({ colIndex: col, label: rawHeader, column: columnName(col) });
      suggestions.push({ colIndex: col, header: rawHeader, field: ranked[0].confidence >= 0.56 ? ranked[0].field : "", confidence: ranked[0].confidence, alternatives: ranked.slice(0, 3) });
    }

    const winners = new Map();
    suggestions.forEach((suggestion) => {
      if (!suggestion.field) return;
      const existing = winners.get(suggestion.field);
      if (!existing || suggestion.confidence > existing.confidence) winners.set(suggestion.field, suggestion);
    });
    suggestions.forEach((suggestion) => {
      if (suggestion.field && winners.get(suggestion.field) !== suggestion) suggestion.field = "";
    });

    const matched = suggestions.filter((item) => item.field);
    const averageConfidence = matched.length ? matched.reduce((sum, item) => sum + item.confidence, 0) / matched.length : 0;
    const coverage = Math.min(1, matched.length / 7);
    const recognitionScore = Math.round((averageConfidence * 0.75 + coverage * 0.25) * 100);

    return {
      name,
      index,
      rows,
      recognized: true,
      score: best.score,
      recognitionScore,
      headerRowIndex,
      headerSpan,
      dataStartIndex: headerRowIndex + headerSpan,
      headers,
      suggestions,
      matchedCount: matched.length,
      estimatedDataRows: rows.slice(headerRowIndex + headerSpan).filter((row) => row.some((value) => text(value))).length
    };
  }

  function reanalyzeWithHeader(sheetAnalysis, headerRowIndex, headerSpan) {
    const rows = sheetAnalysis.rows || [];
    const rowIndex = Math.max(0, Math.min(rows.length - 1, Number(headerRowIndex) || 0));
    const span = headerSpan === 2 ? 2 : 1;
    const candidate = evaluateHeaderCandidate(rows, rowIndex, span);
    const nextRows = rows.slice(rowIndex, Math.min(rows.length, rowIndex + span + 20));
    const colCount = Math.min(MAX_SCAN_COLS, Math.max(0, ...nextRows.map((row) => row.length)));
    const headers = [];
    const suggestions = [];
    for (let col = 0; col < colCount; col += 1) {
      const rawHeader = combinedHeader(rows, rowIndex, span, col) || `未命名列 ${columnName(col)}`;
      const ranked = rankHeader(rawHeader);
      headers.push({ colIndex: col, label: rawHeader, column: columnName(col) });
      suggestions.push({ colIndex: col, header: rawHeader, field: ranked[0].confidence >= 0.56 ? ranked[0].field : "", confidence: ranked[0].confidence, alternatives: ranked.slice(0, 3) });
    }
    const winners = new Map();
    suggestions.forEach((suggestion) => {
      if (!suggestion.field) return;
      const existing = winners.get(suggestion.field);
      if (!existing || suggestion.confidence > existing.confidence) winners.set(suggestion.field, suggestion);
    });
    suggestions.forEach((suggestion) => {
      if (suggestion.field && winners.get(suggestion.field) !== suggestion) suggestion.field = "";
    });
    const matched = suggestions.filter((item) => item.field);
    const averageConfidence = matched.length ? matched.reduce((sum, item) => sum + item.confidence, 0) / matched.length : 0;
    return {
      ...sheetAnalysis,
      recognized: matched.length >= 2,
      score: candidate.score,
      recognitionScore: Math.round((averageConfidence * 0.75 + Math.min(1, matched.length / 7) * 0.25) * 100),
      headerRowIndex: rowIndex,
      headerSpan: span,
      dataStartIndex: rowIndex + span,
      headers,
      suggestions,
      matchedCount: matched.length,
      estimatedDataRows: rows.slice(rowIndex + span).filter((row) => row.some((value) => text(value))).length
    };
  }

  function analyzeWorkbook(workbook) {
    if (!workbook || !Array.isArray(workbook.SheetNames) || !workbook.SheetNames.length) {
      throw new Error("没有找到可读取的工作表。请确认文件不是空文件或受密码保护。");
    }
    const sheets = workbook.SheetNames.map((name, index) => analyzeSheet(workbook.Sheets[name], name, index));
    const recognizedSheets = sheets.filter((sheet) => sheet.recognized).sort((a, b) => b.score - a.score);
    return {
      sheets,
      recognizedSheets,
      bestSheetIndex: recognizedSheets.length ? recognizedSheets[0].index : 0,
      hasRecognizedSheet: recognizedSheets.length > 0
    };
  }

  function parseQuantity(value) {
    if (typeof value === "number" && Number.isFinite(value)) return { value, valid: true };
    const raw = text(value).replace(/，/g, ",").replace(/,/g, "");
    if (!raw) return { value: null, valid: false, reason: "数量为空" };
    const cleaned = raw.replace(/\s*(pcs?|pieces?|个|件|套|只|台|米|m|kg|克|g)\s*$/i, "").trim();
    if (!/^[+\-]?(?:\d+\.?\d*|\.\d+)$/.test(cleaned)) {
      return { value: null, valid: false, reason: `无法把“${text(value)}”识别为数量` };
    }
    const numeric = Number(cleaned);
    if (!Number.isFinite(numeric)) return { value: null, valid: false, reason: "数量不是有效数字" };
    if (numeric <= 0) return { value: numeric, valid: false, reason: numeric === 0 ? "数量不能为 0" : "数量不能为负数" };
    return { value: numeric, valid: true };
  }

  function mappedFields(mapping) {
    return new Set(mapping.filter(Boolean));
  }

  function validateMapping(mapping) {
    const used = mapping.filter(Boolean);
    const duplicates = used.filter((field, index) => used.indexOf(field) !== index);
    const fields = mappedFields(mapping);
    const issues = [];
    if (duplicates.length) issues.push({ level: "error", message: `字段重复对应：${[...new Set(duplicates)].map((key) => FIELD_MAP[key].label).join("、")}` });
    if (!fields.has("quantity")) issues.push({ level: "error", message: "尚未指定“数量”列" });
    if (!fields.has("variant")) issues.push({ level: "warning", message: "原始 BOM 未包含“装配变量（Variant）”列，可在页面“项目信息”的“装配变量”框中手动填写；未填写时按“整机采购”处理" });
    if (!["partNumber", "model", "description", "drawingNo", "assemblyName"].some((field) => fields.has(field))) {
      issues.push({ level: "error", message: "至少指定物料编码、型号/规格、物料名称、图号或装配项目中的一项" });
    }
    if (!fields.has("unit")) issues.push({ level: "warning", message: "未找到单位列，导出时将使用“件”" });
    if (!fields.has("source") && !fields.has("category") && !fields.has("vendor")) {
      issues.push({ level: "warning", message: "缺少来源、类别和厂家，外购件只能按默认规则判断" });
    }
    return { valid: !issues.some((item) => item.level === "error"), issues };
  }

  function engineeringValue(value, unit) {
    if (!Number.isFinite(value)) return "";
    const scales = unit === "Ω"
      ? [[1e6, "M"], [1e3, "K"], [1, ""]]
      : [[1e6, "u"], [1e3, "n"], [1, "p"]];
    for (const [scale, suffix] of scales) {
      if (value >= scale || scale === 1) {
        const number = value / scale;
        const shown = Number.isInteger(number) ? String(number) : String(Number(number.toFixed(3)));
        return `${shown}${suffix}${unit}`;
      }
    }
    return "";
  }

  function codedNumber(code) {
    const normalized = text(code).toUpperCase();
    if (/^\d{1,3}R\d$/.test(normalized)) return Number(normalized.replace("R", "."));
    if (/^R\d{2}$/.test(normalized)) return Number(`0.${normalized.slice(1)}`);
    if (/^\d{3}$/.test(normalized)) return Number(normalized.slice(0, 2)) * (10 ** Number(normalized[2]));
    return null;
  }

  function inferPassiveFields(row) {
    const model = text(row.model).split(/[，,]/)[0].trim();
    const category = normalizeText(row.category);
    const designator = text(row.designator).split(/[，,;；\s/]+/)[0] || "";
    const result = { description: text(row.description), footprint: text(row.footprint) };

    const capacitor = model.toUpperCase().match(/^(0201|0402|0603|0805|1206|1210|1812)[A-Z]{1,3}(\dR\d|R\d{2}|\d{3})[A-Z](\dR\d|\d{3})[A-Z]*$/);
    if ((category.includes("电容") || /^C\d/i.test(designator)) && capacitor) {
      if (!result.footprint) result.footprint = `${capacitor[1]}C`;
      if (!result.description) {
        const capacitance = codedNumber(capacitor[2]);
        const voltage = codedNumber(capacitor[3]);
        const voltageText = voltage >= 1000 ? `${Number((voltage / 1000).toFixed(3))}kV` : `${voltage}V`;
        result.description = `${engineeringValue(capacitance, "F")}/${voltageText}`;
      }
      return result;
    }

    const resistor = model.toUpperCase().match(/^(RC|RS)-0(1|2|3|5)[A-Z]?(\d{1,3}R\d|R\d{2}|\d{3,4})[A-Z]*$/);
    if ((category.includes("电阻") || /^R\d/i.test(designator)) && resistor) {
      const packageMap = { "1": "0201R", "2": "0402R", "3": "0603R", "5": "0805R" };
      if (!result.footprint) result.footprint = packageMap[resistor[2]] || "";
      if (!result.description) {
        const code = resistor[3];
        let resistance = null;
        if (/^\d{4}$/.test(code)) resistance = Number(code.slice(0, 3)) * (10 ** Number(code[3]));
        else resistance = codedNumber(code);
        result.description = engineeringValue(resistance, "Ω");
      }
      return result;
    }

    const size = model.match(/(?:^|[_-])(0201|0402|0603|0805|1206|1210|2012|3216|3225)(?:$|[_-])/i);
    if (!result.footprint && size) result.footprint = size[1];
    if (!result.description) result.description = model || text(row.category);
    return result;
  }

  function buildNormalizedRow(rawRow, sourceRow, sheetName, mapping, multiplier, fallbackVariant) {
    const row = {
      category: "", designator: "", partNumber: "", description: "", model: "", footprint: "", quality: "",
      quantityRaw: "", baseQuantity: null, quantity: null, unit: "", vendor: "", manufacturerPartNumber: "", source: "",
      drawingNo: "", assemblyName: "", variant: "", remark: "", sourceSheet: sheetName, sourceRow, errors: [], warnings: []
    };
    mapping.forEach((field, colIndex) => {
      if (!field) return;
      const value = rawRow[colIndex];
      if (field === "quantity") row.quantityRaw = text(value);
      else row[field] = text(value);
    });
    const inferred = inferPassiveFields(row);
    if (!row.description) row.description = inferred.description;
    if (!row.footprint) row.footprint = inferred.footprint;
    const parsed = parseQuantity(row.quantityRaw);
    if (!parsed.valid) row.errors.push({ field: "quantity", message: parsed.reason });
    else {
      row.baseQuantity = parsed.value;
      row.quantity = parsed.value * multiplier;
    }

    if (![row.partNumber, row.model, row.description, row.drawingNo, row.assemblyName].some(Boolean)) {
      row.errors.push({ field: "identity", message: "缺少物料标识：编码、型号、名称、图号和装配项目均为空" });
    }
    if (!row.unit) {
      row.unit = "件";
      row.warnings.push({ field: "unit", message: "单位为空，已按“件”处理" });
    }
    if (!row.variant) {
      const fb = text(fallbackVariant);
      if (fb) {
        row.variant = fb;
        row.warnings.push({ field: "variant", message: `装配变量为空，已采用项目信息中填写的“${fb}”` });
      } else {
        row.variant = "整机采购";
        row.warnings.push({ field: "variant", message: "装配变量为空，已按“整机采购”处理；可在“项目信息”中手动指定" });
      }
    }
    if (!row.partNumber && row.model) row.warnings.push({ field: "partNumber", message: "没有物料编码，汇总时将使用型号/名称作为识别依据" });
    row.status = row.errors.length ? "error" : row.warnings.length ? "warning" : "ok";
    return row;
  }

  function extractRows(sheetAnalysis, mapping, multiplier, options) {
    const mappingCheck = validateMapping(mapping);
    if (!mappingCheck.valid) return { rows: [], mappingCheck, truncated: false, skippedBlankRows: 0 };
    const numericMultiplier = Number(multiplier);
    if (!Number.isFinite(numericMultiplier) || numericMultiplier <= 0) {
      return { rows: [], mappingCheck: { valid: false, issues: [{ level: "error", message: "生产数量必须是大于 0 的数字" }] }, truncated: false, skippedBlankRows: 0 };
    }
    const fallbackVariant = options && options.fallbackVariant ? String(options.fallbackVariant) : "";
    const rows = [];
    let skippedBlankRows = 0;
    const sourceRows = sheetAnalysis.rows.slice(sheetAnalysis.dataStartIndex, sheetAnalysis.dataStartIndex + MAX_ROWS);
    sourceRows.forEach((rawRow, offset) => {
      const hasData = mapping.some((field, colIndex) => field && text(rawRow[colIndex]));
      if (!hasData) {
        skippedBlankRows += 1;
        return;
      }
      rows.push(buildNormalizedRow(rawRow, sheetAnalysis.dataStartIndex + offset + 1, sheetAnalysis.name, mapping, numericMultiplier, fallbackVariant));
    });
    return {
      rows,
      mappingCheck,
      truncated: sheetAnalysis.rows.length - sheetAnalysis.dataStartIndex > MAX_ROWS,
      skippedBlankRows
    };
  }

  function revalidateRow(row, multiplier, fallbackVariant) {
    const clean = { ...row, errors: [], warnings: [] };
    const parsed = parseQuantity(clean.quantityRaw);
    if (!parsed.valid) {
      clean.baseQuantity = null;
      clean.quantity = null;
      clean.errors.push({ field: "quantity", message: parsed.reason });
    } else {
      clean.baseQuantity = parsed.value;
      clean.quantity = parsed.value * Number(multiplier || 1);
    }
    if (![clean.partNumber, clean.model, clean.description, clean.drawingNo, clean.assemblyName].some((value) => text(value))) {
      clean.errors.push({ field: "identity", message: "缺少物料标识：编码、型号、名称、图号和装配项目均为空" });
    }
    if (!clean.unit) {
      clean.unit = "件";
      clean.warnings.push({ field: "unit", message: "单位为空，已按“件”处理" });
    }
    if (!clean.variant) {
      const fb = text(fallbackVariant);
      if (fb) {
        clean.variant = fb;
        clean.warnings.push({ field: "variant", message: `装配变量为空，已采用项目信息中填写的“${fb}”` });
      } else {
        clean.variant = "整机采购";
        clean.warnings.push({ field: "variant", message: "装配变量为空，已按“整机采购”处理；可在“项目信息”中手动指定" });
      }
    }
    if (!clean.partNumber && clean.model) clean.warnings.push({ field: "partNumber", message: "没有物料编码，汇总时将使用型号/名称作为识别依据" });
    clean.status = clean.errors.length ? "error" : clean.warnings.length ? "warning" : "ok";
    return clean;
  }

  function purchaseDecision(row, mode) {
    if (mode === "all") return { purchased: true, basis: "设置为全部外购" };
    const combined = normalizeText([row.source, row.category].filter(Boolean).join(" "));
    const vendor = text(row.vendor);
    const selfKeywords = ["自制", "自产", "自研", "内部加工", "inhouse", "make", "自产件"];
    const buyKeywords = ["外购", "采购", "标准件", "成品", "外协", "电子元器件", "buy", "purchase", "purchased"];
    if (selfKeywords.some((keyword) => combined.includes(normalizeText(keyword)))) return { purchased: false, basis: "来源/类别标记为自制" };
    if (buyKeywords.some((keyword) => combined.includes(normalizeText(keyword)))) return { purchased: true, basis: "来源/类别含外购关键词" };
    if (mode === "strict") return { purchased: false, basis: "无明确外购关键词" };
    if (vendor) return { purchased: true, basis: "存在厂家/供应商" };
    return { purchased: true, basis: "无明确属性，按外购处理" };
  }

  function identityKey(row, includeVendor) {
    const primary = row.partNumber || row.manufacturerPartNumber || row.model || row.description || row.drawingNo || row.assemblyName;
    return [primary, row.model, row.quality, row.unit, includeVendor ? row.vendor : ""].map(normalizeText).join("|");
  }

  function uniqueJoined(values, limit) {
    const unique = [...new Set(values.flatMap((value) => text(value).split(/[，,;；、\s]+/)).filter(Boolean))];
    const joined = unique.join("、");
    if (!limit || joined.length <= limit) return joined;
    return `${joined.slice(0, limit - 1)}…`;
  }

  function aggregateRows(rows, includeVendor) {
    const groups = new Map();
    rows.forEach((row) => {
      const key = identityKey(row, includeVendor);
      if (!groups.has(key)) {
        groups.set(key, {
          ...row,
          quantity: 0,
          designators: [],
          sourceRows: [],
          purchaseBases: []
        });
      }
      const group = groups.get(key);
      group.quantity += Number(row.quantity || 0);
      group.designators.push(row.designator);
      group.sourceRows.push(`${row.sourceFile ? `${row.sourceFile} / ` : ""}${row.sourceSheet}!${row.sourceRow}`);
      if (row.purchaseBasis) group.purchaseBases.push(row.purchaseBasis);
    });
    return [...groups.values()].map((group) => ({
      ...group,
      // 位号完整保留，避免 Excel 单元格显示不下时被程序侧先截断丢失
      designatorSummary: uniqueJoined(group.designators, 30000),
      traceSummary: uniqueJoined(group.sourceRows, 220),
      purchaseBasis: uniqueJoined(group.purchaseBases, 100)
    }));
  }

  function summarizeIssues(rows) {
    const issues = [];
    rows.forEach((row) => {
      row.errors.forEach((issue) => issues.push({ level: "错误", row, issue }));
      row.warnings.forEach((issue) => issues.push({ level: "提醒", row, issue }));
    });
    return issues;
  }

  function isPrintedBoard(row) {
    const category = normalizeText(row && row.category);
    if (category.includes("印制板") || category === "pcb") return true;
    const descriptive = normalizeText([row && row.description, row && row.model, row && row.partNumber, row && row.drawingNo].filter(Boolean).join(" "));
    if (/印制板|电路板|(?:^|[^a-z])pcb(?:板|板子)?(?:[^a-z]|$)/.test(descriptive)) return true;
    const designator = text(row && row.designator).split(/[，,;；\s/]+/)[0];
    return /^PCB\d*$/i.test(designator);
  }

  function makePrintedBoardRow(options) {
    const revision = text(options && options.pcbRevision);
    const vendor = text(options && options.pcbVendor);
    const multiplier = Number(options && options.multiplier || 1) || 1;
    return {
      category: "印制板",
      partNumber: "",
      description: revision,
      model: revision,
      quality: "I",
      designator: "",
      baseQuantity: 1,
      quantityRaw: "1",
      quantity: multiplier,
      unit: "块",
      vendor,
      manufacturerPartNumber: "",
      source: "外购",
      drawingNo: revision,
      assemblyName: "",
      remark: "",
      purchased: true,
      purchaseBasis: "页面填写的印制板",
      sourceFile: "项目信息",
      sourceSheet: "项目信息",
      sourceRow: "",
      errors: [],
      warnings: [],
      status: "ok"
    };
  }

  function withPrintedBoard(rows, options) {
    const vendor = text(options && options.pcbVendor);
    const revision = text(options && options.pcbRevision);
    const next = (rows || []).map((row) => {
      if (!isPrintedBoard(row)) return row;
      return {
        ...row,
        category: "印制板",
        vendor: vendor || row.vendor,
        model: text(row.model) || revision || row.model,
        description: text(row.description) || revision || row.description
      };
    });
    if (!revision) return next;
    const already = next.some((row) => isPrintedBoard(row) && [row.model, row.description, row.drawingNo, row.partNumber].some((value) => text(value) === revision));
    if (already) return next;
    return [...next, makePrintedBoardRow(options || {})];
  }

  function resistorCapacitorType(row) {
    const descriptive = normalizeText([row.category, row.description, row.assemblyName].filter(Boolean).join(" "));
    const model = text(row.model).toLowerCase();
    const designators = text(row.designator).split(/[，,;；\s/]+/).filter(Boolean);
    if (/电阻|排阻|可调电阻|resistor|res\b/i.test(descriptive) || model.includes("ω") || /(?:ohm|Ω)/i.test(model)) return "电阻";
    if (/电容|钽电容|capacitor|cap\b/i.test(descriptive) || /(?:^|[^a-z])\d+(?:\.\d+)?\s*(?:pf|nf|uf|μf|µf|mf)(?:[^a-z]|$)/i.test(model)) return "电容";
    if (designators.some((item) => /^R\d/i.test(item))) return "电阻";
    if (designators.some((item) => /^C\d/i.test(item))) return "电容";
    if (descriptive.includes("阻容")) return "阻容";
    return "";
  }

  function baseQuantityOf(row) {
    if (Number.isFinite(Number(row.baseQuantity)) && Number(row.baseQuantity) > 0) return Number(row.baseQuantity);
    const parsed = parseQuantity(row.quantityRaw);
    if (parsed.valid) return parsed.value;
    return Number(row.quantity || 0);
  }

  function buildOutputs(rows, options) {
    const validRows = options.skipInvalid === false ? rows : rows.filter((row) => !row.errors.length);
    const withPurchase = validRows.map((row) => {
      const decision = purchaseDecision(row, options.purchaseMode || "auto");
      return { ...row, purchased: decision.purchased, purchaseBasis: decision.basis };
    });
    const purchasedRows = withPurchase.filter((row) => row.purchased);
    const subcontractMultiplier = Number(options.multiplier || 1);
    const subcontractSource = withPurchase
      .map((row) => ({ ...row, componentType: resistorCapacitorType(row) }))
      .filter((row) => row.componentType)
      .map((row) => ({ ...row, quantity: baseQuantityOf(row) * subcontractMultiplier }));
    const picking = aggregateRows(withPurchase, false);
    const procurement = withPrintedBoard(aggregateRows(purchasedRows, true), options);
    const purchased = aggregateRows(purchasedRows, true).sort((a, b) => text(a.vendor).localeCompare(text(b.vendor), "zh-CN"));
    const subcontract = aggregateRows(subcontractSource, true).map((row) => ({
      ...row,
      baseQuantity: subcontractMultiplier ? Number(row.quantity || 0) / subcontractMultiplier : Number(row.quantity || 0)
    }));
    return { source: rows, validRows: withPurchase, assembly: withPurchase, picking, procurement, detail: withPurchase, purchased, subcontract, issues: summarizeIssues(rows) };
  }

  const OUTPUT_COLUMNS = {
    assembly: [
      ["序号", "_index"], ["图号", "drawingNo"], ["装配项目", "assemblyName"], ["位号", "designator"],
      ["物料编码", "partNumber"], ["物料名称", "description"], ["型号 / 规格", "model"], ["质量等级", "quality"],
      ["数量", "quantity"], ["单位", "unit"], ["厂家", "vendor"], ["厂家料号", "manufacturerPartNumber"],
      ["来源", "source"], ["封装", "footprint"], ["备注", "remark"], ["源文件", "sourceFile"], ["原工作表", "sourceSheet"], ["原行号", "sourceRow"]
    ],
    picking: [
      ["序号", "_index"], ["物料编码", "partNumber"], ["物料名称", "description"], ["型号 / 规格", "model"],
      ["质量等级", "quality"], ["领料数量", "quantity"], ["单位", "unit"], ["位号汇总", "designatorSummary"],
      ["类别", "category"], ["厂家", "vendor"], ["备注", "remark"], ["来源追溯", "traceSummary"]
    ],
    procurement: [
      ["序号", "_index"], ["物料编码", "partNumber"], ["物料名称", "description"], ["型号 / 规格", "model"],
      ["厂家料号", "manufacturerPartNumber"], ["厂家 / 供应商", "vendor"], ["采购数量", "quantity"], ["单位", "unit"],
      ["质量等级", "quality"], ["外购判断依据", "purchaseBasis"], ["位号汇总", "designatorSummary"], ["备注", "remark"]
    ],
    detail: [
      ["序号", "_index"], ["类别", "category"], ["物料编码", "partNumber"], ["位号", "designator"],
      ["物料名称", "description"], ["型号 / 规格", "model"], ["质量等级", "quality"], ["数量", "quantity"],
      ["单位", "unit"], ["厂家", "vendor"], ["厂家料号", "manufacturerPartNumber"], ["来源", "source"],
      ["图号", "drawingNo"], ["装配项目", "assemblyName"], ["备注", "remark"], ["源文件", "sourceFile"], ["原工作表", "sourceSheet"], ["原行号", "sourceRow"]
    ],
    purchased: [
      ["序号", "_index"], ["厂家 / 供应商", "vendor"], ["物料编码", "partNumber"], ["物料名称", "description"],
      ["型号 / 规格", "model"], ["厂家料号", "manufacturerPartNumber"], ["总数量", "quantity"], ["单位", "unit"],
      ["质量等级", "quality"], ["位号汇总", "designatorSummary"], ["外购判断依据", "purchaseBasis"], ["来源追溯", "traceSummary"]
    ],
    subcontract: [
      ["序号", "_index"], ["元件类型", "componentType"], ["物料编码", "partNumber"], ["物料名称", "description"],
      ["型号 / 规格", "model"], ["厂家料号", "manufacturerPartNumber"], ["厂家 / 供应商", "vendor"], ["备料数量", "quantity"],
      ["单位", "unit"], ["质量等级", "quality"], ["位号汇总", "designatorSummary"], ["来源追溯", "traceSummary"], ["备注", "remark"]
    ]
  };

  const COLORS = {
    green: "176C52", green2: "248366", greenSoft: "E4F2EB", ink: "19362D", muted: "65756E",
    line: "DCE4DF", paper: "FFFEFA", amber: "FFF3D7", red: "FCE9E6", white: "FFFFFF"
  };

  function borderStyle(color) {
    return {
      top: { style: "thin", color: { rgb: color } }, bottom: { style: "thin", color: { rgb: color } },
      left: { style: "thin", color: { rgb: color } }, right: { style: "thin", color: { rgb: color } }
    };
  }

  function applySheetStyle(sheet, columnCount, dataCount, headerRowIndex, numericColumnIndexes) {
    const endCol = XLSX.utils.encode_col(Math.max(0, columnCount - 1));
    sheet["!merges"] = sheet["!merges"] || [];
    sheet["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(0, columnCount - 1) } });
    sheet["!autofilter"] = { ref: `A${headerRowIndex + 1}:${endCol}${headerRowIndex + dataCount + 1}` };
    sheet["!freeze"] = { xSplit: 0, ySplit: headerRowIndex + 1, topLeftCell: `A${headerRowIndex + 2}`, activePane: "bottomLeft", state: "frozen" };
    sheet["!rows"] = [{ hpt: 28 }, { hpt: 22 }, { hpt: 22 }, { hpt: 22 }, { hpt: 8 }, { hpt: 24 }];

    for (let col = 0; col < columnCount; col += 1) {
      const titleCell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
      if (titleCell) titleCell.s = { fill: { fgColor: { rgb: COLORS.green } }, font: { color: { rgb: COLORS.white }, bold: true, sz: 16 }, alignment: { horizontal: "left", vertical: "center" } };
      const headerCell = sheet[XLSX.utils.encode_cell({ r: headerRowIndex, c: col })];
      if (headerCell) headerCell.s = { fill: { fgColor: { rgb: COLORS.greenSoft } }, font: { color: { rgb: COLORS.ink }, bold: true }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: borderStyle("C6D6CE") };
    }
    for (let row = 1; row <= 3; row += 1) {
      for (let col = 0; col < columnCount; col += 1) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        if (!cell) continue;
        cell.s = { fill: { fgColor: { rgb: row % 2 ? "F4F8F5" : COLORS.paper } }, font: { color: { rgb: col % 2 === 0 ? COLORS.muted : COLORS.ink }, bold: col % 2 === 0 }, alignment: { vertical: "center" } };
      }
    }
    for (let row = headerRowIndex + 1; row <= headerRowIndex + dataCount; row += 1) {
      for (let col = 0; col < columnCount; col += 1) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        if (!cell) continue;
        cell.s = { fill: { fgColor: { rgb: row % 2 ? COLORS.white : "F8FAF8" } }, font: { color: { rgb: COLORS.ink }, sz: 10 }, border: { bottom: { style: "hair", color: { rgb: "E7ECE9" } } }, alignment: { vertical: "center", wrapText: col > 1 } };
        if (numericColumnIndexes.includes(col)) cell.z = "#,##0.####";
      }
    }
  }

  function makeDataSheet(title, rows, columns, metadata) {
    const meta = metadata || {};
    const headerRowIndex = 5;
    const data = [
      [title],
      ["项目工号", meta.projectCode || "", "产品型号", meta.productModel || "", "项目名称", meta.projectName || ""],
      ["批次", meta.batch || "", "生产数量", Number(meta.multiplier || 1), "使用部门", meta.department || ""],
      ["生成时间", meta.generatedAt || "", "数据来源", meta.sourceFile || "", "处理模式", meta.mode === "assembly" ? "装配清单导入" : "AD / BOM 导入"],
      [],
      columns.map(([label]) => label),
      ...rows.map((row, index) => columns.map(([, key]) => key === "_index" ? index + 1 : (row[key] ?? "")))
    ];
    const sheet = XLSX.utils.aoa_to_sheet(data);
    const widths = columns.map(([label, key]) => {
      const base = ["quantity", "sourceRow", "_index"].includes(key) ? 10 : ["designatorSummary", "traceSummary", "remark"].includes(key) ? 28 : 17;
      return { wch: Math.max(base, Math.min(32, text(label).length * 2 + 3)) };
    });
    sheet["!cols"] = widths;
    const numeric = columns.map(([, key], index) => ["quantity", "sourceRow", "_index"].includes(key) ? index : -1).filter((index) => index >= 0);
    applySheetStyle(sheet, columns.length, rows.length, headerRowIndex, numeric);
    return sheet;
  }

  function makeIssueSheet(issues, metadata) {
    const rows = issues.map((item) => ({
      level: item.level,
      sourceSheet: item.row.sourceSheet,
      sourceRow: item.row.sourceRow,
      field: FIELD_MAP[item.issue.field] ? FIELD_MAP[item.issue.field].label : item.issue.field,
      message: item.issue.message,
      identity: item.row.partNumber || item.row.model || item.row.description || item.row.drawingNo || item.row.assemblyName || "",
      quantityRaw: item.row.quantityRaw || ""
    }));
    const columns = [["级别", "level"], ["原工作表", "sourceSheet"], ["原行号", "sourceRow"], ["涉及字段", "field"], ["具体说明", "message"], ["物料标识", "identity"], ["原数量", "quantityRaw"]];
    const sheet = makeDataSheet("数据问题清单", rows, columns, metadata);
    const headerRowIndex = 5;
    rows.forEach((row, index) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: headerRowIndex + 1 + index, c: 0 })];
      if (!cell) return;
      cell.s = { ...(cell.s || {}), fill: { fgColor: { rgb: row.level === "错误" ? COLORS.red : COLORS.amber } }, font: { color: { rgb: row.level === "错误" ? "9E3B34" : "8B5B0B" }, bold: true } };
    });
    return sheet;
  }

  function makeReadmeSheet(outputs, selectedOutputs, metadata) {
    const rows = [
      ["智料单 · 导出说明"],
      ["项目工号", metadata.projectCode || "", "项目名称", metadata.projectName || ""],
      ["源文件", metadata.sourceFile || "", "生成时间", metadata.generatedAt || ""],
      [],
      ["清单", "行数", "说明"],
      ...selectedOutputs.map((key) => [OUTPUT_DEFS[key].label, outputs[key].length, OUTPUT_DEFS[key].description]),
      ["数据问题清单", outputs.issues.length, "错误与提醒均保留原工作表和行号，便于追溯"],
      [],
      ["外购判断规则"],
      ["自动判断", "来源/类别明确写有自制时排除；写有外购、采购、标准件等关键词或存在厂家时视为外购；无属性时为避免漏采，默认按外购处理。"],
      ["数量计算", metadata.mode === "assembly"
        ? `装配文件的领料、明细和外购汇总保持原数量；仅外协备料清单 = 原数量 × 生产数量（${metadata.multiplier || 1}）`
        : `AD / BOM 清单数量 = 原表数量 × 生产数量（${metadata.multiplier || 1}）；外协备料清单不会重复相乘`],
      ["隐私", "本文件由本地浏览器生成，源料单未上传。"]
    ];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }, { s: { r: 8, c: 0 }, e: { r: 8, c: 3 } }];
    sheet["!cols"] = [{ wch: 20 }, { wch: 28 }, { wch: 18 }, { wch: 70 }];
    sheet["!rows"] = [{ hpt: 31 }];
    [0, 8].forEach((row) => {
      for (let col = 0; col < 4; col += 1) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        if (cell) cell.s = { fill: { fgColor: { rgb: COLORS.green } }, font: { color: { rgb: COLORS.white }, bold: true, sz: row === 0 ? 16 : 12 }, alignment: { vertical: "center" } };
      }
    });
    for (let col = 0; col < 3; col += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: 4, c: col })];
      if (cell) cell.s = { fill: { fgColor: { rgb: COLORS.greenSoft } }, font: { bold: true, color: { rgb: COLORS.ink } }, border: borderStyle("C6D6CE") };
    }
    return sheet;
  }

  function deepClone(value) {
    return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
  }

  function outputColumnAliases(column) {
    const [label, key] = column;
    const aliases = [label];
    if (key === "_index") aliases.push("序号", "项次", "行号", "no", "number");
    if (key === "quantity") aliases.push("数量", "总数量", "需求数量", "领料数量", "采购数量", "用量", "qty", "quantity");
    if (key === "designatorSummary") aliases.push("位号", "位号汇总", "参考位号", "designator");
    if (key === "traceSummary") aliases.push("来源追溯", "数据来源", "原始行", "trace");
    if (key === "purchaseBasis") aliases.push("外购判断依据", "判断依据", "采购属性", "来源判断");
    if (FIELD_MAP[key]) aliases.push(...FIELD_MAP[key].aliases, FIELD_MAP[key].label);
    return [...new Set(aliases)];
  }

  function rankOutputHeader(value, columns) {
    let best = { key: "", label: "", confidence: 0 };
    columns.forEach((column) => {
      const confidence = Math.max(...outputColumnAliases(column).map((alias) => aliasScore(value, alias)));
      if (confidence > best.confidence) best = { key: column[1], label: column[0], confidence };
    });
    return best;
  }

  function inspectTemplate(workbook, outputKey) {
    if (!workbook || !workbook.SheetNames || !OUTPUT_COLUMNS[outputKey]) {
      return { valid: false, message: "模板文件无效或清单类型不受支持。" };
    }
    const columns = OUTPUT_COLUMNS[outputKey];
    let best = null;
    workbook.SheetNames.forEach((sheetName) => {
      const rows = sheetRows(workbook.Sheets[sheetName]);
      const scanRows = Math.min(60, rows.length);
      for (let rowIndex = 0; rowIndex < scanRows; rowIndex += 1) {
        const row = rows[rowIndex] || [];
        const matches = [];
        const usedKeys = new Set();
        for (let colIndex = 0; colIndex < Math.min(MAX_SCAN_COLS, row.length); colIndex += 1) {
          const value = text(row[colIndex]);
          if (!value) continue;
          const ranked = rankOutputHeader(value, columns);
          if (ranked.confidence >= 0.62 && !usedKeys.has(ranked.key)) {
            matches.push({ colIndex, header: value, ...ranked });
            usedKeys.add(ranked.key);
          }
        }
        const hasQuantity = matches.some((item) => item.key === "quantity");
        const hasIdentity = matches.some((item) => ["partNumber", "model", "description", "drawingNo", "assemblyName"].includes(item.key));
        const score = matches.reduce((sum, item) => sum + item.confidence, 0) * 10 + matches.length * 2 + (hasQuantity ? 3 : 0) + (hasIdentity ? 2 : 0);
        if (!best || score > best.score) best = { sheetName, rowIndex, matches, score, hasQuantity, hasIdentity };
      }
    });
    if (!best || best.matches.length < 2 || !best.hasQuantity) {
      return {
        valid: false,
        message: `没有在模板中找到“${OUTPUT_DEFS[outputKey].label}”所需表头。模板至少应包含数量列和一个物料字段。`,
        matchedCount: best ? best.matches.length : 0
      };
    }
    return {
      valid: true,
      outputKey,
      sheetName: best.sheetName,
      headerRowIndex: best.rowIndex,
      dataStartIndex: best.rowIndex + 1,
      matches: best.matches,
      matchedCount: best.matches.length,
      confidence: Math.round(best.matches.reduce((sum, item) => sum + item.confidence, 0) / best.matches.length * 100),
      message: `已在“${best.sheetName}”第 ${best.rowIndex + 1} 行识别 ${best.matches.length} 个模板字段。`
    };
  }

  function setCellValue(cell, value) {
    const result = cell ? deepClone(cell) : {};
    delete result.f;
    delete result.F;
    delete result.w;
    result.v = value === null || value === undefined ? "" : value;
    result.t = typeof result.v === "number" ? "n" : typeof result.v === "boolean" ? "b" : "s";
    return result;
  }

  function fillTemplateMetadata(sheet, metadata, beforeRowIndex) {
    if (!sheet["!ref"]) return;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const fields = {
      "项目工号": metadata.projectCode || "",
      "产品型号": metadata.productModel || "",
      "项目名称": metadata.projectName || "",
      "批次": metadata.batch || "",
      "生产数量": Number(metadata.multiplier || 1),
      "使用部门": metadata.department || "",
      "装配变量": metadata.variant || "整机采购",
      "源文件": metadata.sourceFile || "",
      "生成时间": metadata.generatedAt || ""
    };
    const placeholders = {
      "项目工号": "projectCode", "产品型号": "productModel", "项目名称": "projectName", "批次": "batch",
      "生产数量": "multiplier", "使用部门": "department", "装配变量": "variant", "源文件": "sourceFile", "生成时间": "generatedAt"
    };
    for (let row = range.s.r; row <= Math.min(range.e.r, 120); row += 1) {
      for (let col = range.s.c; col <= Math.min(range.e.c, 50); col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = sheet[address];
        if (!cell || typeof cell.v !== "string") continue;
        let value = cell.v;
        Object.entries(placeholders).forEach(([label, key]) => {
          value = value.replace(new RegExp(`\\{\\{\\s*${label}\\s*\\}\\}`, "g"), text(metadata[key]));
          value = value.replace(new RegExp(`\\{\\{\\s*${key}\\s*\\}\\}`, "gi"), text(metadata[key]));
        });
        if (value !== cell.v) sheet[address] = setCellValue(cell, value);
        const exactLabel = Object.keys(fields).find((label) => normalizeText(cell.v) === normalizeText(label));
        if (exactLabel && row < (beforeRowIndex ?? Number.POSITIVE_INFINITY) && col < range.e.c) {
          const nextAddress = XLSX.utils.encode_cell({ r: row, c: col + 1 });
          const nextCell = sheet[nextAddress];
          sheet[nextAddress] = setCellValue(nextCell, fields[exactLabel]);
        }
      }
    }
  }

  function shiftTemplateRows(sheet, startRowIndex, rowCount) {
    if (!rowCount || rowCount < 1 || !sheet["!ref"]) return;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let row = range.e.r; row >= startRowIndex; row -= 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const source = XLSX.utils.encode_cell({ r: row, c: col });
        const target = XLSX.utils.encode_cell({ r: row + rowCount, c: col });
        if (sheet[source]) sheet[target] = sheet[source];
        else delete sheet[target];
        delete sheet[source];
      }
    }
    if (sheet["!rows"]) {
      for (let row = sheet["!rows"].length - 1; row >= startRowIndex; row -= 1) {
        sheet["!rows"][row + rowCount] = sheet["!rows"][row];
        delete sheet["!rows"][row];
      }
    }
    if (sheet["!merges"]) {
      sheet["!merges"].forEach((merge) => {
        if (merge.s.r >= startRowIndex) {
          merge.s.r += rowCount;
          merge.e.r += rowCount;
        }
      });
    }
    range.e.r += rowCount;
    sheet["!ref"] = XLSX.utils.encode_range(range);
  }

  function templateValue(row, key, index, metadata) {
    if (key === "_index") return index + 1;
    if (key === "_batch") return metadata.batch || "";
    if (key === "_zero") return 0;
    if (key === "modelOrDescription") return row.model || row.description || row.partNumber || "";
    if (key === "baseQuantity") return baseQuantityOf(row);
    return row[key] ?? "";
  }

  function makeTemplateSheet(templateWorkbook, outputKey, rows, metadata, suppliedInspection) {
    const inspection = suppliedInspection && suppliedInspection.valid ? suppliedInspection : inspectTemplate(templateWorkbook, outputKey);
    if (!inspection.valid) throw new Error(`${OUTPUT_DEFS[outputKey].label}模板无法使用：${inspection.message}`);
    const original = templateWorkbook.Sheets[inspection.sheetName];
    const sheet = deepClone(original);
    let originalRange = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const dataStart = inspection.dataStartIndex;
    const configuredDataEnd = Number.isInteger(inspection.dataEndIndex) ? inspection.dataEndIndex : null;
    const capacity = configuredDataEnd === null ? Number.POSITIVE_INFINITY : configuredDataEnd - dataStart + 1;
    if (Number.isInteger(inspection.footerStartIndex) && rows.length > capacity) {
      shiftTemplateRows(sheet, inspection.footerStartIndex, rows.length - capacity);
      originalRange = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    }
    const lastWriteRow = dataStart + Math.max(0, rows.length - 1);
    const clearEndRow = configuredDataEnd === null ? Math.max(originalRange.e.r, lastWriteRow) : Math.max(configuredDataEnd, lastWriteRow);

    inspection.matches.forEach((match) => {
      const styleSourceAddress = XLSX.utils.encode_cell({ r: dataStart, c: match.colIndex });
      const styleSource = sheet[styleSourceAddress] || sheet[XLSX.utils.encode_cell({ r: inspection.headerRowIndex, c: match.colIndex })];
      for (let rowIndex = dataStart; rowIndex <= clearEndRow; rowIndex += 1) {
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: match.colIndex });
        const existing = sheet[address] || styleSource;
        sheet[address] = setCellValue(existing, "");
      }
      rows.forEach((row, index) => {
        const address = XLSX.utils.encode_cell({ r: dataStart + index, c: match.colIndex });
        const value = templateValue(row, match.key, index, metadata);
        sheet[address] = setCellValue(sheet[address] || styleSource, value);
        if (match.key === "quantity" && !sheet[address].z) sheet[address].z = "#,##0.####";
      });
    });

    if (sheet["!rows"] && sheet["!rows"][dataStart]) {
      for (let rowIndex = dataStart + 1; rowIndex <= lastWriteRow; rowIndex += 1) {
        if (!sheet["!rows"][rowIndex]) sheet["!rows"][rowIndex] = deepClone(sheet["!rows"][dataStart]);
      }
    }
    originalRange.e.r = Math.max(originalRange.e.r, lastWriteRow);
    originalRange.e.c = Math.max(originalRange.e.c, ...inspection.matches.map((item) => item.colIndex));
    sheet["!ref"] = XLSX.utils.encode_range(originalRange);
    if (sheet["!autofilter"] && sheet["!autofilter"].ref) {
      const filterRange = XLSX.utils.decode_range(sheet["!autofilter"].ref);
      filterRange.e.r = Math.max(filterRange.e.r, lastWriteRow);
      sheet["!autofilter"].ref = XLSX.utils.encode_range(filterRange);
    }
    fillTemplateMetadata(sheet, metadata, inspection.headerRowIndex);
    return sheet;
  }

  function safeFilePart(value) {
    return text(value).replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").slice(0, 60);
  }

  function buildWorkbook(rows, metadata, options) {
    const outputs = buildOutputs(rows, {
      ...options,
      multiplier: Number(metadata.multiplier || 1),
      pcbRevision: metadata.pcbRevision,
      pcbVendor: metadata.pcbVendor
    });
    const selectedOutputs = (options.selectedOutputs || []).filter((key) => OUTPUT_DEFS[key]);
    if (!selectedOutputs.length) throw new Error("请至少选择一种需要生成的清单。");
    const now = new Date();
    const generatedAt = now.toLocaleString("zh-CN", { hour12: false });
    const fullMetadata = { ...metadata, generatedAt, multiplier: Number(metadata.multiplier || 1) };
    const workbook = XLSX.utils.book_new();
    workbook.Props = {
      Title: `${fullMetadata.projectName || fullMetadata.projectCode || "项目"}料单`,
      Subject: "由智料单自动识别并生成",
      Author: "智料单",
      CreatedDate: now
    };
    XLSX.utils.book_append_sheet(workbook, makeReadmeSheet(outputs, selectedOutputs, fullMetadata), "导出说明");
    selectedOutputs.forEach((key) => {
      const template = options.templates && options.templates[key];
      const sheet = template
        ? makeTemplateSheet(template.workbook, key, outputs[key], fullMetadata, template.inspection)
        : makeDataSheet(OUTPUT_DEFS[key].label, outputs[key], OUTPUT_COLUMNS[key], fullMetadata);
      XLSX.utils.book_append_sheet(workbook, sheet, OUTPUT_DEFS[key].label);
    });
    if (outputs.issues.length) XLSX.utils.book_append_sheet(workbook, makeIssueSheet(outputs.issues, fullMetadata), "数据问题清单");
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
    const prefix = safeFilePart(fullMetadata.projectCode || fullMetadata.projectName || "项目");
    return { workbook, filename: `${prefix}_料单_${stamp}.xlsx`, outputs };
  }

  return {
    MAX_ROWS,
    FIELD_DEFS,
    FIELD_MAP,
    OUTPUT_DEFS,
    OUTPUT_COLUMNS,
    HEADER_REQUIREMENTS,
    headerRequirementStatus,
    templateColumnStatus,
    normalizeText,
    rankHeader,
    parseQuantity,
    analyzeSheet,
    reanalyzeWithHeader,
    analyzeWorkbook,
    validateMapping,
    extractRows,
    revalidateRow,
    purchaseDecision,
    isPrintedBoard,
    buildOutputs,
    inspectTemplate,
    makeTemplateSheet,
    buildWorkbook
  };
});
