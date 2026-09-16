(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("jszip"));
  else root.TemplateExporter = factory(root.JSZip);
})(typeof self !== "undefined" ? self : this, function (JSZip) {
  "use strict";

  function xmlEscape(value) {
    return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  }

  function base64Bytes(base64) {
    if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(base64, "base64"));
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  function rowNumber(address) {
    return Number(String(address).match(/\d+$/)[0]);
  }

  function columnName(index) {
    let value = index + 1;
    let result = "";
    while (value) {
      value -= 1;
      result = String.fromCharCode(65 + value % 26) + result;
      value = Math.floor(value / 26);
    }
    return result;
  }

  function columnIndex(address) {
    const letters = String(address).match(/^[A-Z]+/i)[0].toUpperCase();
    let value = 0;
    for (const letter of letters) value = value * 26 + letter.charCodeAt(0) - 64;
    return value - 1;
  }

  function cellStyle(xml, address) {
    const rowInfo = rowBounds(xml, rowNumber(address));
    const regex = new RegExp(`<(?:x:)?c(?=\\s)[^>]*\\br=["']${xmlEscape(address)}["'][^>]*?(?:\\s*\\/>|>[\\s\\S]*?<\\/(?:x:)?c>)`, "i");
    const match = rowInfo.value.match(regex);
    if (!match) return "";
    const opening = match[0].match(/^<(?:x:)?c([^>]*?)(?:>|\/>)/i);
    const style = opening && opening[1].match(/\bs=["'](\d+)["']/i);
    return style ? style[1] : "";
  }

  function isFormulaValue(value) {
    return value && typeof value === "object" && typeof value.formula === "string";
  }

  function cellXml(address, attributes, value) {
    if (isFormulaValue(value)) {
      const cached = Number(value.value);
      const body = Number.isFinite(cached)
        ? `<f>${xmlEscape(value.formula)}</f><v>${cached}</v>`
        : `<f>${xmlEscape(value.formula)}</f>`;
      return `<c${attributes} t="n">${body}</c>`;
    }
    if (value === "" || value === null || value === undefined) return `<c${attributes}></c>`;
    if (typeof value === "number" && Number.isFinite(value)) return `<c${attributes} t="n"><v>${value}</v></c>`;
    return `<c${attributes} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  function setCell(xml, address, value, styleId) {
    const row = rowNumber(address);
    const rowRegex = new RegExp(`<(?:x:)?row\\b[^>]*\\br=["']${row}["'][^>]*>[\\s\\S]*?<\\/(?:x:)?row>`, "i");
    const rowMatch = rowRegex.exec(xml);
    if (!rowMatch) {
      if ((value === "" || value === null || value === undefined) && styleId === undefined) return xml;
      throw new Error(`默认模板缺少第 ${row} 行`);
    }
    const rowInfo = { value: rowMatch[0], index: rowMatch.index };
    const shortAddress = xmlEscape(address);
    const regex = new RegExp(`<(?:x:)?c(?=\\s)[^>]*\\br=["']${shortAddress}["'][^>]*?(?:\\s*\\/>|>[\\s\\S]*?<\\/(?:x:)?c>)`, "i");
    const match = rowInfo.value.match(regex);
    if (!match && (value === "" || value === null || value === undefined) && styleId === undefined) return xml;
    let cell;
    let updatedRow;
    if (!match) {
      const style = styleId === undefined || styleId === "" ? "" : ` s="${styleId}"`;
      cell = cellXml(address, ` r="${address}"${style}`, value);
      const targetColumn = columnIndex(address);
      const existingCells = [...rowInfo.value.matchAll(/<(?:x:)?c(?=\s)[^>]*\br=["']([A-Z]+)\d+["'][^>]*?(?:\s*\/>|>[\s\S]*?<\/(?:x:)?c>)/gi)];
      const nextCell = existingCells.find((item) => columnIndex(item[1]) > targetColumn);
      updatedRow = nextCell
        ? `${rowInfo.value.slice(0, nextCell.index)}${cell}${rowInfo.value.slice(nextCell.index)}`
        : rowInfo.value.replace(/<\/(?:x:)?row>$/i, `${cell}</row>`);
    } else {
      const opening = match[0].match(/^<(?:x:)?c([^>]*?)(?:>|\/>)/i);
      let attributes = (opening ? opening[1] : ` r="${address}"`).replace(/\s+t=["'][^"']*["']/gi, "").replace(/\s*\/\s*$/, "");
      if (styleId !== undefined) {
        attributes = attributes.replace(/\s+s=["'][^"']*["']/gi, "");
        if (styleId !== "") attributes += ` s="${styleId}"`;
      }
      cell = cellXml(address, attributes, value);
      updatedRow = `${rowInfo.value.slice(0, match.index)}${cell}${rowInfo.value.slice(match.index + match[0].length)}`;
    }
    return `${xml.slice(0, rowInfo.index)}${updatedRow}${xml.slice(rowInfo.index + rowInfo.value.length)}`;
  }

  function metadataText(template, metadata) {
    const productModel = String(metadata.productModel || "").trim();
    const boardNo = String(metadata.boardNo || metadata.pcbRevision || "").trim();
    const productModelBoard = productModel && boardNo ? `${productModel}/${boardNo}` : (productModel || boardNo || "");
    return String(template || "")
      .replace(/\{\{使用部门\}\}/g, metadata.department || "")
      .replace(/\{\{项目工号\}\}/g, metadata.projectCode || "")
      .replace(/\{\{产品型号板号\}\}/g, productModelBoard)
      .replace(/\{\{产品型号\}\}/g, productModel)
      .replace(/\{\{板号\}\}/g, boardNo)
      .replace(/\{\{PCB版号\}\}/g, boardNo)
      .replace(/\{\{项目名称\}\}/g, metadata.projectName || "")
      .replace(/\{\{生产数量\}\}/g, String(metadata.multiplier || 1))
      .replace(/\{\{批次\}\}/g, metadata.batch || "")
      .replace(/\{\{装配变量\}\}/g, metadata.variant || "整机采购")
      .replace(/\{\{源文件\}\}/g, metadata.sourceFile || "")
      .replace(/\{\{生成时间\}\}/g, metadata.generatedAt || new Date().toLocaleString("zh-CN", { hour12: false }));
  }

  function baseQuantity(row) {
    const value = Number(row.baseQuantity);
    if (Number.isFinite(value) && value > 0) return value;
    const parsed = Number(row.quantityRaw);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
    return Number(row.quantity || 0);
  }

  function templateValue(row, key, index, metadata, options) {
    if (key === "_index") return index + 1;
    if (row.__category) return key === "modelOrDescription" ? row.category : "";
    if (!key) return "";
    if (key === "_batch") return metadata.batch || "";
    if (key === "_zero") return 0;
    if (key === "_electricalAssembly") return "\u7535\u88c5";
    if (key === "_yes") return "是";
    if (key === "modelOrDescription") return row.model || row.description || row.partNumber || "";
    if (key === "baseQuantity") return baseQuantity(row);
    if (key === "quantity" && options && options.quantityFormula && !row.__category) {
      const formula = options.quantityFormula;
      const rowNumber = options.dataStartRow + index;
      const baseCol = columnName(Number(formula.baseColIndex));
      const multiplierCell = String(formula.multiplierCell || "G3").toUpperCase();
      const cached = Number(row.quantity);
      return {
        formula: `${baseCol}${rowNumber}*$${multiplierCell.replace(/^([A-Z]+)(\d+)$/, "$1$$$2")}`,
        value: Number.isFinite(cached) ? cached : baseQuantity(row) * Number(metadata.multiplier || 1)
      };
    }
    return row[key] ?? "";
  }

  function rowBounds(xml, oneBasedRow) {
    const regex = new RegExp(`<(?:x:)?row\\b[^>]*\\br=["']${oneBasedRow}["'][^>]*>[\\s\\S]*?<\\/(?:x:)?row>`, "i");
    const match = regex.exec(xml);
    if (!match) throw new Error(`默认模板缺少第 ${oneBasedRow} 行`);
    return { regex, value: match[0] };
  }

  function rewriteRowNumber(rowXml, fromRow, toRow) {
    return rowXml
      .replace(new RegExp(`(<(?:x:)?row\\b[^>]*\\br=["'])${fromRow}(["'])`, "i"), `$1${toRow}$2`)
      .replace(new RegExp(`(<(?:x:)?c\\b[^>]*\\br=["'][A-Z]+)${fromRow}(["'])`, "gi"), `$1${toRow}$2`);
  }

  function shiftFooter(xml, startOneBasedRow, shift) {
    if (shift <= 0) return xml;
    const rows = [];
    const regex = /<(?:x:)?row\b[^>]*\br=["'](\d+)["'][^>]*>[\s\S]*?<\/(?:x:)?row>/gi;
    let match;
    while ((match = regex.exec(xml))) if (Number(match[1]) >= startOneBasedRow) rows.push({ row: Number(match[1]), xml: match[0] });
    rows.sort((a, b) => b.row - a.row).forEach((item) => {
      xml = xml.replace(item.xml, rewriteRowNumber(item.xml, item.row, item.row + shift));
    });
    xml = xml.replace(/(<(?:x:)?mergeCell\b[^>]*\bref=["'])([A-Z]+)(\d+):([A-Z]+)(\d+)(["'])/gi, (whole, lead, col1, row1, col2, row2, tail) => {
      const r1 = Number(row1), r2 = Number(row2);
      return `${lead}${col1}${r1 >= startOneBasedRow ? r1 + shift : r1}:${col2}${r2 >= startOneBasedRow ? r2 + shift : r2}${tail}`;
    });
    return xml;
  }

  function extendRows(xml, inspection, rowCount) {
    const start = inspection.dataStartIndex + 1;
    const end = inspection.dataEndIndex + 1;
    const capacity = end - start + 1;
    if (rowCount <= capacity) return xml;
    const extra = rowCount - capacity;
    const footer = inspection.footerStartIndex + 1;
    xml = shiftFooter(xml, footer, extra);
    const templateRow = rowBounds(xml, end).value;
    const inserted = [];
    for (let offset = 1; offset <= extra; offset += 1) inserted.push(rewriteRowNumber(templateRow, end, end + offset));
    return xml.replace(templateRow, `${templateRow}${inserted.join("")}`);
  }

  function hideRowsInRange(xml, startRow, endRow) {
    if (!Number.isFinite(startRow) || !Number.isFinite(endRow) || endRow < startRow) return xml;
    let next = xml;
    for (let row = startRow; row <= endRow; row += 1) {
      const regex = new RegExp(`<(?:x:)?row\\b([^>]*\\br=["']${row}["'][^>]*)(/?)>`, "i");
      next = next.replace(regex, (whole, attrs, selfClosing) => {
        if (/\bhidden=["']1["']/i.test(attrs)) return whole;
        const cleaned = attrs.replace(/\s+hidden=["'][^"']*["']/gi, "");
        return selfClosing ? `<row${cleaned} hidden="1"/>` : `<row${cleaned} hidden="1">`;
      });
    }
    return next;
  }

  function removeSheetBreaks(xml) {
    return xml
      .replace(/<(?:x:)?rowBreaks\b[^>]*>[\s\S]*?<\/(?:x:)?rowBreaks>/gi, "")
      .replace(/<(?:x:)?colBreaks\b[^>]*>[\s\S]*?<\/(?:x:)?colBreaks>/gi, "")
      .replace(/<(?:x:)?rowBreaks\b[^>]*\/>/gi, "")
      .replace(/<(?:x:)?colBreaks\b[^>]*\/>/gi, "");
  }

  function contentPrintLastRow(xml, inspection, dataStartRow, dataRowCount) {
    const capacity = inspection.dataEndIndex - inspection.dataStartIndex + 1;
    const extra = Math.max(0, dataRowCount - capacity);
    const lastDataRow = dataRowCount > 0 ? dataStartRow + dataRowCount - 1 : dataStartRow - 1;
    if (inspection.footerStartIndex == null) {
      return Math.max(lastDataRow, inspection.headerRowIndex + 1, 1);
    }
    const footerStart = inspection.footerStartIndex + 1 + extra;
    const footerRows = [...xml.matchAll(/<(?:x:)?row\b[^>]*\br=["'](\d+)["']/gi)]
      .map((match) => Number(match[1]))
      .filter((row) => row >= footerStart);
    const footerEnd = footerRows.length ? Math.max(...footerRows) : footerStart;
    return Math.max(lastDataRow, footerEnd, 1);
  }

  function collapseUnusedPrintRows(xml, inspection, dataStartRow, dataRowCount) {
    const capacity = inspection.dataEndIndex - inspection.dataStartIndex + 1;
    const extra = Math.max(0, dataRowCount - capacity);
    const lastDataRow = dataRowCount > 0 ? dataStartRow + dataRowCount - 1 : dataStartRow - 1;
    const templateDataEnd = inspection.dataEndIndex + 1 + extra;
    let next = removeSheetBreaks(xml);
    if (inspection.footerStartIndex != null) {
      const footerStart = inspection.footerStartIndex + 1 + extra;
      const hideFrom = Math.max(lastDataRow + 1, dataStartRow);
      if (hideFrom <= footerStart - 1) next = hideRowsInRange(next, hideFrom, footerStart - 1);
    } else if (lastDataRow < templateDataEnd) {
      const hideFrom = Math.max(lastDataRow + 1, dataStartRow);
      next = hideRowsInRange(next, hideFrom, templateDataEnd);
    }
    return next;
  }

  function updateDimension(xml, lastRow) {
    return xml.replace(/(<(?:x:)?dimension\b[^>]*\bref=["'][A-Z]+\d+:[A-Z]+)(\d+)(["'])/i, (whole, lead, current, tail) => `${lead}${lastRow}${tail}`);
  }

  function setOpenView(xml, viewMode) {
    const pageBreakPreview = viewMode === "pageBreakPreview";
    xml = xml.replace(/<(?:x:)?sheetView\b([^>]*)>/i, (whole, attributes) => {
      const clean = attributes
        .replace(/\s+(?:view|topLeftCell|zoomScale|zoomScaleNormal|zoomScaleSheetLayoutView)=["'][^"']*["']/gi, "");
      return `<sheetView${clean} view="${pageBreakPreview ? "pageBreakPreview" : "normal"}" topLeftCell="A1" zoomScale="100" zoomScaleNormal="100">`;
    });
    return xml.replace(/<(?:x:)?selection\b[^>]*\/>/i, '<selection activeCell="A1" sqref="A1"/>');
  }

  function setPageFooter(xml, footerText) {
    if (!footerText) return xml;
    const footer = `<headerFooter><oddFooter>${xmlEscape(footerText)}</oddFooter></headerFooter>`;
    if (/<(?:x:)?headerFooter\b/i.test(xml)) {
      return xml.replace(/<(?:x:)?headerFooter\b[^>]*>[\s\S]*?<\/(?:x:)?headerFooter>/i, (existing) => {
        if (/<(?:x:)?oddFooter\b/i.test(existing)) {
          return existing.replace(/<(?:x:)?oddFooter\b[^>]*>[\s\S]*?<\/(?:x:)?oddFooter>/i, `<oddFooter>${xmlEscape(footerText)}</oddFooter>`);
        }
        return existing.replace(/<\/(?:x:)?headerFooter>/i, `<oddFooter>${xmlEscape(footerText)}</oddFooter></headerFooter>`);
      });
    }
    if (/<(?:x:)?pageSetup\b[^>]*\/>/i.test(xml)) return xml.replace(/(<(?:x:)?pageSetup\b[^>]*\/>)/i, `$1${footer}`);
    if (/<(?:x:)?pageMargins\b[^>]*\/>/i.test(xml)) return xml.replace(/(<(?:x:)?pageMargins\b[^>]*\/>)/i, `$1${footer}`);
    return xml.replace(/<\/(?:x:)?worksheet>$/i, `${footer}</worksheet>`);
  }

  function ensureFitToPageWidth(xml) {
    let next = xml;
    if (/<sheetPr\b/i.test(next)) {
      if (/<pageSetUpPr\b/i.test(next)) {
        next = next.replace(/<pageSetUpPr\b[^>]*\/>/i, '<pageSetUpPr fitToPage="1"/>');
      } else if (/<sheetPr\b[^>]*\/>/i.test(next)) {
        next = next.replace(/<(?:x:)?sheetPr\b([^>]*)\/>/i, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
      } else {
        next = next.replace(/<(?:x:)?sheetPr\b([^>]*)>/i, (open, attrs) => {
          if (/\/\s*$/.test(attrs)) return open;
          return `<sheetPr${attrs}><pageSetUpPr fitToPage="1"/>`;
        });
      }
    } else {
      next = next.replace(/<(?:x:)?worksheet\b[^>]*>/i, (open) => `${open}<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`);
    }
    if (/<pageSetup\b/i.test(next)) {
      next = next.replace(/<(?:x:)?pageSetup\b([^>]*)\/>/i, (whole, attrs) => {
        let cleaned = attrs
          .replace(/\s+scale=["'][^"']*["']/gi, "")
          .replace(/\s+fitToWidth=["'][^"']*["']/gi, "")
          .replace(/\s+fitToHeight=["'][^"']*["']/gi, "");
        return `<pageSetup${cleaned} fitToWidth="1" fitToHeight="0"/>`;
      });
    }
    return next;
  }

  async function setPrintArea(zip, inspection, lastRow) {
    if (!inspection.printAreaMaxCol) return;
    const path = "xl/workbook.xml";
    const entry = zip.file(path);
    if (!entry) return;
    let xml = await entry.async("string");
    const sheet = String(inspection.sheetName || "Sheet1").replace(/'/g, "''");
    const area = `'${sheet}'!$A$1:$${inspection.printAreaMaxCol}$${lastRow}`;
    const title = inspection.printTitleRows ? `'${sheet}'!${inspection.printTitleRows}` : "";
    const replaceName = (name, value) => {
      const regex = new RegExp(`(<(?:x:)?definedName\\b[^>]*\\bname=["']${name.replace(/[.*+?^${}()|[\\]\\]/g, "\\$&")}["'][^>]*>)[\\s\\S]*?(<\\/(?:x:)?definedName>)`, "i");
      if (regex.test(xml)) xml = xml.replace(regex, (whole, lead, tail) => `${lead}${xmlEscape(value)}${tail}`);
      else {
        const node = `<definedName name="${name}" localSheetId="0">${xmlEscape(value)}</definedName>`;
        if (/<(?:x:)?definedNames\b/i.test(xml)) xml = xml.replace(/<\/(?:x:)?definedNames>/i, `${node}</definedNames>`);
        else xml = xml.replace(/<\/(?:x:)?sheets>/i, `</sheets><definedNames>${node}</definedNames>`);
      }
    };
    replaceName("_xlnm.Print_Area", area);
    if (title) replaceName("_xlnm.Print_Titles", title);
    zip.file(path, xml);
  }

  function removeWorksheetFormulas(xml) {
    return xml.replace(/<f\b[^>]*>[\s\S]*?<\/f>/gi, "").replace(/<f\b[^>]*\/>/gi, "");
  }

  async function sanitizeWorkbook(zip) {
    Object.keys(zip.files).filter((name) => /^xl\/externalLinks\//i.test(name)).forEach((name) => zip.remove(name));
    zip.remove("xl/calcChain.xml");

    const typesEntry = zip.file("[Content_Types].xml");
    if (typesEntry) {
      let xml = await typesEntry.async("string");
      xml = xml.replace(/<Override\b[^>]*PartName=["'][^"']*(?:calcChain|externalLinks)[^"']*["'][^>]*\/>/gi, "");
      zip.file("[Content_Types].xml", xml);
    }

    const relsPath = "xl/_rels/workbook.xml.rels";
    const relsEntry = zip.file(relsPath);
    if (relsEntry) {
      let xml = await relsEntry.async("string");
      xml = xml.replace(/<Relationship\b[^>]*Type=["'][^"']*(?:calcChain|externalLink)[^"']*["'][^>]*\/>/gi, "");
      zip.file(relsPath, xml);
    }

    const workbookPath = "xl/workbook.xml";
    const workbookEntry = zip.file(workbookPath);
    if (workbookEntry) {
      let xml = await workbookEntry.async("string");
      xml = xml.replace(/<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>/gi, "");
      if (/<calcPr\b/i.test(xml)) {
        xml = xml.replace(/<calcPr\b[^>]*\/?>(?:<\/calcPr>)?/i, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>');
      } else {
        xml = xml.replace(/<\/workbook>$/i, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
      }
      zip.file(workbookPath, xml);
    }
  }

  function materialCategory(row) {
    const explicit = String(row.category || "").trim();
    if (explicit && !/^(物料|器件|元件|其他|其它)$/i.test(explicit)) return explicit;
    const description = `${row.description || ""} ${row.model || ""}`;
    const designator = String(row.designator || row.designatorSummary || "").split(/[，,;；\s/]+/)[0];
    if (/印制板|电路板|(?:^|[^a-zA-Z])pcb(?:板)?(?:[^a-zA-Z]|$)/i.test(description) || /^PCB\d*$/i.test(designator)) return "印制板";
    if (/按键|开关|button|switch/i.test(description) || /^(SW|K)\d/i.test(designator)) return "按键";
    if (/磁珠|ferrite/i.test(description) || /^FB\d/i.test(designator)) return "磁珠";
    if (/电感|inductor/i.test(description) || /^L\d/i.test(designator)) return "电感";
    if (/电容|capacitor/i.test(description) || /^C\d/i.test(designator)) return "电容";
    if (/电阻|resistor/i.test(description) || /^R\d/i.test(designator)) return "电阻";
    if (/晶振|oscillator|crystal/i.test(description) || /^(Y|X)\d/i.test(designator)) return "晶振";
    if (/连接器|接插件|插座|插头|connector/i.test(description) || /^(J|P|CON)\d/i.test(designator)) return "接插件";
    if (/^(U|IC)\d/i.test(designator)) return "集成电路";
    if (/^(D|Q|VR)\d/i.test(designator)) return "半导体";
    return "其他";
  }

  const WRAP_TEXT_KEYS = new Set([
    "designator",
    "designatorSummary",
    "description",
    "model",
    "modelOrDescription",
    "remark",
    "drawingNo",
    "assemblyName",
    "vendor"
  ]);

  const LEFT_ALIGN_KEYS = new Set(["designator", "description", "model", "modelOrDescription"]);

  function visualWidth(text) {
    let width = 0;
    for (const char of String(text || "")) {
      width += /[\u1100-\u115F\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]/.test(char) ? 2 : 1;
    }
    return width;
  }

  function wrapMaxWidth(colWidth) {
    // Slightly tighter than column width so Excel wrap + row height stay ahead of clipping.
    return Math.max(8, Math.floor(Number(colWidth || 10) * 0.92));
  }

  function hardWrapByWidth(text, maxWidth) {
    if (!text || visualWidth(text) <= maxWidth) return [text || ""];
    const lines = [];
    let current = "";
    for (const char of text) {
      const candidate = current + char;
      if (current && visualWidth(candidate) > maxWidth) {
        lines.push(current);
        current = char;
      } else {
        current = candidate;
      }
    }
    if (current) lines.push(current);
    return lines;
  }

  function softWrapBySeparators(value, maxWidth) {
    const source = String(value ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
    if (!source || !maxWidth || maxWidth < 4 || visualWidth(source) <= maxWidth) return source;
    const lines = [];
    for (const paragraph of source.split("\n")) {
      if (!paragraph) {
        lines.push("");
        continue;
      }
      const tokens = paragraph.split(/([，,;；、])/).filter((part) => part !== "");
      let current = "";
      tokens.forEach((token) => {
        const candidate = current + token;
        if (current && visualWidth(candidate) > maxWidth) {
          lines.push(current.replace(/[，,;；、\s]+$/g, ""));
          current = /^[，,;；、]$/.test(token) ? "" : token;
        } else {
          current = candidate;
        }
      });
      if (current) lines.push(current.replace(/[，,;；、\s]+$/g, ""));
    }
    const separatorWrapped = lines.filter((line, index, arr) => line || (index > 0 && index < arr.length - 1));
    return separatorWrapped.flatMap((line) => hardWrapByWidth(line, maxWidth)).join("\n");
  }

  function parseColumnWidths(xml, maxCol) {
    const widths = Array.from({ length: maxCol + 1 }, () => 10);
    const cols = xml.match(/<(?:x:)?cols\b[^>]*>[\s\S]*?<\/(?:x:)?cols>/i);
    if (!cols) return widths;
    [...cols[0].matchAll(/<(?:x:)?col\b([^>]*?)\/>/gi)].forEach((match) => {
      const attrs = match[1];
      const min = Number((attrs.match(/\bmin=["'](\d+)["']/i) || [])[1] || 0);
      const max = Number((attrs.match(/\bmax=["'](\d+)["']/i) || [])[1] || min);
      const width = Number((attrs.match(/\bwidth=["']([\d.]+)["']/i) || [])[1] || 10);
      for (let col = min - 1; col <= max - 1 && col <= maxCol; col += 1) {
        if (col >= 0) widths[col] = width;
      }
    });
    return widths;
  }

  function presentTemplateValue(value, key, colWidth) {
    if (!WRAP_TEXT_KEYS.has(key) || value == null || typeof value === "number" || isFormulaValue(value)) return value;
    return softWrapBySeparators(value, wrapMaxWidth(colWidth));
  }

  function countDisplayLines(value, colWidth) {
    const text = String(value ?? "");
    if (!text) return 1;
    const maxWidth = wrapMaxWidth(colWidth);
    return text.split(/\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(visualWidth(line) / maxWidth)), 0);
  }

  function rowHeightBaseline(xml, sampleRow) {
    const open = xml.match(new RegExp(`<(?:x:)?row\\b[^>]*\\br=["']${sampleRow}["'][^>]*>`, "i"));
    const ht = open && open[0].match(/\bht=["']([\d.]+)["']/i);
    const value = ht ? Number(ht[1]) : 15.75;
    if (value >= 24) return { line: value / 2, min: value };
    return { line: Math.max(14.25, value), min: value };
  }

  function setRowHeight(xml, row, height) {
    const regex = new RegExp(`<(?:x:)?row\\b([^>]*\\br=["']${row}["'][^>]*)(/?)>`, "i");
    return xml.replace(regex, (whole, attrs, selfClosing) => {
      let next = attrs
        .replace(/\s+ht=["'][^"']*["']/gi, "")
        .replace(/\s+customHeight=["'][^"']*["']/gi, "");
      next += ` ht="${Number(height.toFixed(2))}" customHeight="1"`;
      return selfClosing ? `<row${next}/>` : `<row${next}>`;
    });
  }

  function ensureWrapTextStyles(stylesXml, styleIds) {
    const wanted = new Set([...styleIds].map(String).filter(Boolean));
    if (!wanted.size) return stylesXml;
    return stylesXml.replace(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/i, (whole, attrs, body) => {
      let index = 0;
      const nextBody = body.replace(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/gi, (xf) => {
        const current = String(index);
        index += 1;
        if (!wanted.has(current) || /wrapText=["']1["']/i.test(xf)) return xf;
        if (/<alignment\b/i.test(xf)) {
          return xf.replace(/<alignment\b([^>]*?)\/>/i, (align, alignAttrs) => {
            if (/wrapText=/i.test(alignAttrs)) {
              return `<alignment${alignAttrs.replace(/\swrapText=["'][^"']*["']/i, "")} wrapText="1"/>`;
            }
            return `<alignment${alignAttrs} wrapText="1"/>`;
          });
        }
        if (/\/>\s*$/.test(xf)) {
          const open = xf.replace(/\/>\s*$/, "");
          const withApply = /applyAlignment=/i.test(open) ? open : `${open} applyAlignment="1"`;
          return `${withApply}><alignment vertical="center" wrapText="1"/></xf>`;
        }
        return xf.replace(/^(<xf\b[^>]*>)/i, (open) => {
          const withApply = /applyAlignment=/i.test(open) ? open : open.replace(/>$/, ` applyAlignment="1">`);
          return `${withApply}<alignment vertical="center" wrapText="1"/>`;
        });
      });
      return `<cellXfs${attrs}>${nextBody}</cellXfs>`;
    });
  }

  function listCellXfs(stylesXml) {
    const match = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i);
    if (!match) return [];
    return [...match[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/gi)].map((item) => item[0]);
  }

  function xfWithLeftAlign(xf) {
    let result = xf;
    if (/<alignment\b/i.test(result)) {
      result = result.replace(/<alignment\b([^>]*?)\/>/i, (whole, attrs) => {
        const cleaned = attrs
          .replace(/\shorizontal=["'][^"']*["']/gi, "")
          .replace(/\svertical=["'][^"']*["']/gi, "")
          .replace(/\swrapText=["'][^"']*["']/gi, "");
        return `<alignment${cleaned} horizontal="left" vertical="center" wrapText="1"/>`;
      });
    } else if (/\/>\s*$/.test(result)) {
      const open = result.replace(/\/>\s*$/, "");
      const withApply = /applyAlignment=/i.test(open) ? open : `${open} applyAlignment="1"`;
      result = `${withApply}><alignment horizontal="left" vertical="center" wrapText="1"/></xf>`;
    } else {
      result = result.replace(/^(<xf\b[^>]*>)/i, (open) => {
        const withApply = /applyAlignment=/i.test(open) ? open : open.replace(/>$/, ` applyAlignment="1">`);
        return `${withApply}<alignment horizontal="left" vertical="center" wrapText="1"/>`;
      });
    }
    if (!/applyAlignment=/i.test(result)) {
      result = result.replace(/^(<xf\b)/i, "$1 applyAlignment=\"1\"");
    }
    return result;
  }

  function appendLeftAlignStyles(stylesXml, styleIds) {
    const wanted = [...new Set([...styleIds].map(String).filter(Boolean))];
    if (!wanted.length) return { xml: stylesXml, map: {} };
    const xfs = listCellXfs(stylesXml);
    const map = {};
    const additions = [];
    let nextIndex = xfs.length;
    wanted.forEach((id) => {
      const xf = xfs[Number(id)];
      if (!xf) return;
      if (/horizontal=["']left["']/i.test(xf)) {
        map[id] = id;
        return;
      }
      map[id] = String(nextIndex);
      additions.push(xfWithLeftAlign(xf));
      nextIndex += 1;
    });
    if (!additions.length) return { xml: stylesXml, map };
    return {
      xml: stylesXml.replace(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/i, (whole, attrs, body) => {
        const nextAttrs = /\bcount=/i.test(attrs)
          ? attrs.replace(/\bcount=["']\d+["']/i, `count="${nextIndex}"`)
          : `${attrs} count="${nextIndex}"`;
        return `<cellXfs${nextAttrs}>${body}${additions.join("")}</cellXfs>`;
      }),
      map
    };
  }

  function remapCellStyle(xml, address, styleId) {
    if (styleId === undefined || styleId === null || styleId === "") return xml;
    const row = rowNumber(address);
    const rowRegex = new RegExp(`<(?:x:)?row\\b[^>]*\\br=["']${row}["'][^>]*>[\\s\\S]*?<\\/(?:x:)?row>`, "i");
    const rowMatch = rowRegex.exec(xml);
    if (!rowMatch) return xml;
    const shortAddress = xmlEscape(address);
    const updatedRow = rowMatch[0].replace(
      new RegExp(`<(?:x:)?c(?=\\s)([^>]*\\br=["']${shortAddress}["'][^>]*?)(/?)>`, "i"),
      (whole, attrs, selfClosing) => {
        let next = attrs.replace(/\s+s=["'][^"']*["']/gi, "");
        next += ` s="${styleId}"`;
        return selfClosing ? `<c${next}/>` : `<c${next}>`;
      }
    );
    return `${xml.slice(0, rowMatch.index)}${updatedRow}${xml.slice(rowMatch.index + rowMatch[0].length)}`;
  }

  function categorizedRows(rows, categoryKey) {
    const groups = new Map();
    rows.forEach((row) => {
      const category = String(categoryKey && row[categoryKey] || "").trim() || materialCategory(row);
      if (!groups.has(category)) groups.set(category, []);
      groups.get(category).push(row);
    });
    return [...groups.entries()].flatMap(([category, items]) => [{ __category: true, category }, ...items]);
  }

  async function fillTemplate(definition, outputRows, metadata) {
    if (!JSZip) throw new Error("ZIP 组件未加载");
    const inspection = definition.inspection;
    const zip = await JSZip.loadAsync(base64Bytes(definition.base64));
    const sheetPath = inspection.sheetPath || "xl/worksheets/sheet1.xml";
    let xml = await zip.file(sheetPath).async("string");
    const rowsForTemplate = inspection.categorize ? categorizedRows(outputRows, inspection.categoryKey) : outputRows;
    const styleRows = inspection.rowStyleRows || {};
    const dataMaxCol = Number.isInteger(inspection.dataMaxCol) ? inspection.dataMaxCol : Math.max(...inspection.matches.map((item) => item.colIndex));
    const styles = {};
    Object.entries(styleRows).forEach(([kind, oneBasedRow]) => {
      styles[kind] = [];
      for (let colIndex = 0; colIndex <= dataMaxCol; colIndex += 1) styles[kind][colIndex] = cellStyle(xml, `${columnName(colIndex)}${oneBasedRow}`);
    });
    Object.entries(inspection.styleOverrides || {}).forEach(([kind, overrides]) => {
      styles[kind] = styles[kind] || [];
      Object.entries(overrides).forEach(([colIndex, styleId]) => { styles[kind][Number(colIndex)] = String(styleId); });
    });
    xml = extendRows(xml, inspection, rowsForTemplate.length);
    const columnWidths = parseColumnWidths(xml, dataMaxCol);
    const heightBase = rowHeightBaseline(xml, inspection.dataStartIndex + 1);

    Object.entries(inspection.metadataCells || {}).forEach(([address, template]) => {
      xml = setCell(xml, address, metadataText(template, metadata));
    });

    const dataStartRow = inspection.dataStartIndex + 1;
    const clearEndRow = inspection.dataEndIndex + 1 + Math.max(0, rowsForTemplate.length - (inspection.dataEndIndex - inspection.dataStartIndex + 1));
    for (let row = dataStartRow; row <= clearEndRow; row += 1) {
      for (let colIndex = 0; colIndex <= dataMaxCol; colIndex += 1) xml = setCell(xml, `${columnName(colIndex)}${row}`, "");
    }

    const valueOptions = { dataStartRow, quantityFormula: null };
    const usedStyleIds = new Set();
    const leftAlignCells = [];
    Object.values(styles).forEach((rowStyles) => {
      (rowStyles || []).forEach((styleId) => { if (styleId) usedStyleIds.add(String(styleId)); });
    });

    function writeCell(data, index, match, styleId) {
      if (styleId) usedStyleIds.add(String(styleId));
      const raw = templateValue(data, match.key, index, metadata, valueOptions);
      const presented = presentTemplateValue(raw, match.key, columnWidths[match.colIndex]);
      const address = `${columnName(match.colIndex)}${dataStartRow + index}`;
      xml = setCell(xml, address, presented, styleId);
      if (!data.__category && LEFT_ALIGN_KEYS.has(match.key)) {
        leftAlignCells.push({ address, key: match.key });
      }
      return presented;
    }

    const rowContents = rowsForTemplate.map(() => []);
    inspection.matches.forEach((match) => {
      rowsForTemplate.forEach((data, index) => {
        const kind = data.__category ? "category" : "item";
        const styleId = styles[kind] ? styles[kind][match.colIndex] : undefined;
        const presented = writeCell(data, index, match, styleId);
        rowContents[index].push({ colIndex: match.colIndex, value: presented });
      });
    });
    if (styles.category || styles.item) {
      rowsForTemplate.forEach((data, index) => {
        const kind = data.__category ? "category" : "item";
        rowContents[index] = [];
        for (let colIndex = 0; colIndex <= dataMaxCol; colIndex += 1) {
          const match = inspection.matches.find((item) => item.colIndex === colIndex) || { colIndex, key: "" };
          const presented = writeCell(data, index, match, styles[kind][colIndex]);
          rowContents[index].push({ colIndex, value: presented });
        }
      });
    }
    xml = removeWorksheetFormulas(xml);
    if (inspection.quantityFormula) {
      const quantityMatch = inspection.matches.find((match) => match.key === "quantity");
      if (quantityMatch) {
        const formulaOptions = { dataStartRow, quantityFormula: inspection.quantityFormula };
        const col = columnName(quantityMatch.colIndex);
        rowsForTemplate.forEach((data, index) => {
          if (data.__category) return;
          const styleId = styles.item ? styles.item[quantityMatch.colIndex] : undefined;
          if (styleId) usedStyleIds.add(String(styleId));
          xml = setCell(xml, `${col}${dataStartRow + index}`, templateValue(data, "quantity", index, metadata, formulaOptions), styleId);
        });
      }
    }

    rowsForTemplate.forEach((data, index) => {
      inspection.matches.forEach((match) => {
        const styleId = cellStyle(xml, `${columnName(match.colIndex)}${dataStartRow + index}`);
        if (styleId) usedStyleIds.add(String(styleId));
      });
      if (data.__category) return;
      const lines = Math.max(
        1,
        ...rowContents[index].map((cell) => countDisplayLines(cell.value, columnWidths[cell.colIndex]))
      );
      const height = Math.max(heightBase.min, Number((lines * heightBase.line * 1.08).toFixed(2)));
      xml = setRowHeight(xml, dataStartRow + index, height);
    });

    const capacity = inspection.dataEndIndex - inspection.dataStartIndex + 1;
    const extraRows = Math.max(0, rowsForTemplate.length - capacity);
    xml = collapseUnusedPrintRows(xml, inspection, dataStartRow, rowsForTemplate.length);
    const printLastRow = contentPrintLastRow(xml, inspection, dataStartRow, rowsForTemplate.length);
    xml = setOpenView(xml, inspection.openView || "normal");
    xml = setPageFooter(xml, inspection.pageFooter);
    if (inspection.fitToPageWidth) xml = ensureFitToPageWidth(xml);
    xml = updateDimension(xml, printLastRow);

    const stylesEntry = zip.file("xl/styles.xml");
    if (stylesEntry && (usedStyleIds.size || leftAlignCells.length)) {
      let stylesXml = ensureWrapTextStyles(await stylesEntry.async("string"), usedStyleIds);
      const leftBases = new Set();
      leftAlignCells.forEach((cell) => {
        const styleId = cellStyle(xml, cell.address);
        if (styleId) leftBases.add(String(styleId));
      });
      const left = appendLeftAlignStyles(stylesXml, leftBases);
      stylesXml = left.xml;
      leftAlignCells.forEach((cell) => {
        const baseStyle = cellStyle(xml, cell.address);
        const mapped = left.map[baseStyle];
        if (mapped != null) xml = remapCellStyle(xml, cell.address, mapped);
      });
      zip.file("xl/styles.xml", stylesXml);
    }

    zip.file(sheetPath, xml);
    await setPrintArea(zip, inspection, printLastRow);
    await sanitizeWorkbook(zip);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  function safePart(value) {
    return String(value || "项目").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").slice(0, 70);
  }

  async function buildFiles(rows, metadata, options, core, XLSX) {
    const outputs = core.buildOutputs(rows, {
      ...options,
      mode: metadata.mode || options.mode || "bom",
      multiplier: Number(metadata.multiplier || 1),
      pcbRevision: metadata.pcbRevision,
      pcbVendor: metadata.pcbVendor,
      pcbRevisions: metadata.pcbRevisions,
      sourceFile: metadata.sourceFile
    });
    const selected = (options.selectedOutputs || []).filter((key) => core.OUTPUT_DEFS[key]);
    if (!selected.length) throw new Error("请至少选择一种需要生成的清单");
    const files = [];
    const generatedAt = new Date().toLocaleString("zh-CN", { hour12: false });
    const fullMetadata = { ...metadata, generatedAt };
    for (const key of selected) {
      const active = options.templates && options.templates[key];
      const exactDefinition = active && active.exactDefinition;
      if (exactDefinition) {
        const bytes = await fillTemplate(exactDefinition, outputs[key], fullMetadata);
        files.push({ key, filename: `${safePart(metadata.projectCode || metadata.projectName)}_${core.OUTPUT_DEFS[key].label}.xlsx`, bytes });
      } else {
        const built = core.buildWorkbook(rows, metadata, { ...options, selectedOutputs: [key], templates: active ? { [key]: active } : {} });
        files.push({ key, filename: built.filename.replace("_料单_", `_${core.OUTPUT_DEFS[key].label}_`), bytes: new Uint8Array(XLSX.write(built.workbook, { type: "array", bookType: "xlsx", compression: true })) });
      }
    }
    return { files, outputs };
  }

  return { fillTemplate, buildFiles, setCell, metadataText, materialCategory, categorizedRows, softWrapBySeparators, visualWidth, ensureWrapTextStyles };
});
