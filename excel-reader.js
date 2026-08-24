(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    if (!global.cptable) global.cptable = require("codepage");
    module.exports = factory(require("xlsx-js-style"), require("jszip"), require("./bom-core.js"));
  } else {
    root.ExcelReader = factory(root.XLSX, root.JSZip, root.BOMCore);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX, JSZip, BOMCore) {
  "use strict";

  const ENCODINGS = {
    auto: { label: "自动识别", decoder: null, codepage: null },
    utf8: { label: "UTF-8", decoder: "utf-8", codepage: 65001 },
    gb18030: { label: "简体中文 GB18030 / GBK", decoder: "gb18030", codepage: 936 },
    big5: { label: "繁体中文 Big5", decoder: "big5", codepage: 950 },
    western: { label: "西文 Windows-1252", decoder: "windows-1252", codepage: 1252 }
  };

  function bytesOf(arrayBuffer) {
    return arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  }

  function startsWith(bytes, signature) {
    return signature.every((value, index) => bytes[index] === value);
  }

  function decodePreview(bytes) {
    return new TextDecoder("windows-1252").decode(bytes.slice(0, 2048)).replace(/^\uFEFF/, "").trimStart().toLowerCase();
  }

  function detectFileKind(arrayBuffer) {
    const bytes = bytesOf(arrayBuffer);
    if (startsWith(bytes, [0x50, 0x4B, 0x03, 0x04]) || startsWith(bytes, [0x50, 0x4B, 0x05, 0x06]) || startsWith(bytes, [0x50, 0x4B, 0x07, 0x08])) return "xlsx";
    if (startsWith(bytes, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])) return "xls";
    const preview = decodePreview(bytes);
    if (preview.startsWith("<?xml") || preview.includes("<workbook") || preview.includes("<ss:workbook")) return "xml";
    if (preview.startsWith("<html") || preview.startsWith("<!doctype html") || preview.includes("<table")) return "html";
    return "text";
  }

  function sampleWorkbookStrings(workbook, maxStrings) {
    const result = [];
    if (!workbook || !workbook.SheetNames) return result;
    for (const sheetName of workbook.SheetNames.slice(0, 5)) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || !sheet["!ref"]) continue;
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      for (let row = range.s.r; row <= Math.min(range.e.r, range.s.r + 80); row += 1) {
        for (let col = range.s.c; col <= Math.min(range.e.c, range.s.c + 40); col += 1) {
          const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
          if (cell && typeof cell.v === "string" && cell.v.trim()) result.push(cell.v.trim());
          if (result.length >= maxStrings) return result;
        }
      }
    }
    return result;
  }

  function mojibakeScore(workbook) {
    const strings = sampleWorkbookStrings(workbook, 800);
    if (!strings.length) return 80;
    let score = 0;
    let readableChinese = 0;
    let suspiciousStrings = 0;
    const commonMojibake = /锟斤拷|烫烫|屯屯|鏂|鐗|缂|浣|鏁|鍨|绫|娴|璇|瀵|绛|鍚|浜|绠|銆|鈥|闁|闂|闄|閿|瀹|纭|鎬|鍙|鎴/g;
    strings.forEach((value) => {
      const replacement = (value.match(/\uFFFD/g) || []).length;
      const control = (value.match(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g) || []).length;
      const privateUse = (value.match(/[\uE000-\uF8FF]/g) || []).length;
      const extendedLatin = (value.match(/[ÃÂÐÑÒÓÔÕÖØÙÚÛÜÝÞßæåçèéïðñòóùúûü]/g) || []).length;
      const mojibake = (value.match(commonMojibake) || []).length;
      const boxes = (value.match(/[□�]/g) || []).length;
      const chinese = (value.match(/[\u3400-\u9FFF]/g) || []).length;
      readableChinese += chinese;
      const local = replacement * 24 + control * 20 + privateUse * 16 + boxes * 18 + extendedLatin * 5 + mojibake * 3;
      score += local;
      if (local >= 8) suspiciousStrings += 1;
    });
    score += suspiciousStrings * 2;
    if (!readableChinese && strings.some((value) => /[\x80-\xFF]/.test(value))) score += 18;
    try {
      const analysis = BOMCore.analyzeWorkbook(workbook);
      if (analysis.hasRecognizedSheet) {
        const best = analysis.sheets.find((sheet) => sheet.index === analysis.bestSheetIndex);
        score -= Math.min(30, (best && best.recognitionScore ? best.recognitionScore : 0) * 0.3);
      } else {
        score += 12;
      }
    } catch (_) {
      score += 20;
    }
    return Math.max(0, Math.round(score));
  }

  function decodeText(arrayBuffer, encodingKey) {
    const encoding = ENCODINGS[encodingKey] || ENCODINGS.utf8;
    return new TextDecoder(encoding.decoder || "utf-8").decode(bytesOf(arrayBuffer)).replace(/^\uFEFF/, "");
  }

  function readTextCandidate(arrayBuffer, encodingKey) {
    const config = ENCODINGS[encodingKey];
    const decoded = decodeText(arrayBuffer, encodingKey);
    return XLSX.read(decoded, { type: "string", cellDates: false, cellFormula: false, cellNF: false, codepage: config.codepage });
  }

  function readBinaryCandidate(arrayBuffer, encodingKey) {
    const config = ENCODINGS[encodingKey] || ENCODINGS.gb18030;
    return XLSX.read(bytesOf(arrayBuffer), { type: "array", cellDates: false, cellFormula: false, cellNF: false, codepage: config.codepage || 936 });
  }

  async function repairZipXml(arrayBuffer, encodingKey) {
    if (!JSZip) throw new Error("ZIP 解码组件未加载");
    const zip = await JSZip.loadAsync(arrayBuffer);
    const xmlNames = Object.keys(zip.files).filter((name) => /^(?:xl\/.*|docProps\/.*)\.xml$/i.test(name));
    for (const name of xmlNames) {
      const entry = zip.file(name);
      if (!entry) continue;
      const raw = await entry.async("uint8array");
      let decoded = new TextDecoder(ENCODINGS[encodingKey].decoder).decode(raw);
      decoded = decoded.replace(/(<\?xml[^>]*encoding=["'])[^"']+(["'])/i, "$1UTF-8$2");
      zip.file(name, decoded);
    }
    const repaired = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 3 } });
    return XLSX.read(repaired, { type: "array", cellDates: false, cellFormula: false, cellNF: false });
  }

  async function candidate(label, encodingKey, loader) {
    try {
      const workbook = await loader();
      return { workbook, encoding: encodingKey, label, score: mojibakeScore(workbook) };
    } catch (error) {
      return { error, encoding: encodingKey, label, score: Number.POSITIVE_INFINITY };
    }
  }

  async function readWorkbook(arrayBuffer, options) {
    const preference = options && options.encoding && ENCODINGS[options.encoding] ? options.encoding : "auto";
    const kind = detectFileKind(arrayBuffer);
    const candidates = [];

    if (kind === "xlsx") {
      if (preference === "auto" || preference === "utf8") {
        candidates.push(await candidate("标准 XLSX / UTF-8", "utf8", () => readBinaryCandidate(arrayBuffer, "utf8")));
      }
      // 某些 ERP / AD 软件生成的文件扩展名和 ZIP 结构都是 XLSX，但内部
      // 个别 XML 实际使用 GBK，且错误字符很少，单靠乱码阈值会漏检。
      // 自动模式始终比较标准、GBK、Big5 三种结果，再按表头识别和乱码分数选择。
      const repairKeys = preference === "auto" ? ["gb18030", "big5"] : [preference];
      for (const key of repairKeys) {
        if (key === "utf8" || key === "western") continue;
        candidates.push(await candidate(`修复型 XLSX / ${ENCODINGS[key].label}`, key, () => repairZipXml(arrayBuffer, key)));
      }
    } else if (kind === "xls") {
      const keys = preference === "auto" ? ["gb18030", "utf8", "big5", "western"] : [preference];
      for (const key of keys) candidates.push(await candidate(`旧版 XLS / ${ENCODINGS[key].label}`, key, () => readBinaryCandidate(arrayBuffer, key)));
    } else {
      const keys = preference === "auto" ? ["utf8", "gb18030", "big5", "western"] : [preference];
      for (const key of keys) candidates.push(await candidate(`${kind.toUpperCase()} / ${ENCODINGS[key].label}`, key, () => readTextCandidate(arrayBuffer, key)));
    }

    const successful = candidates.filter((item) => item.workbook).sort((a, b) => a.score - b.score);
    if (!successful.length) {
      const detail = candidates.find((item) => item.error);
      throw new Error(detail && detail.error ? detail.error.message : "没有可用的解码方式");
    }
    const best = successful[0];
    return {
      workbook: best.workbook,
      kind,
      encoding: best.encoding,
      encodingLabel: best.label,
      mojibakeScore: best.score,
      suspicious: best.score >= 25,
      candidates: successful.map((item) => ({ encoding: item.encoding, label: item.label, score: item.score }))
    };
  }

  return { ENCODINGS, detectFileKind, mojibakeScore, readWorkbook, decodeText };
});
