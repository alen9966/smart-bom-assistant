(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("jszip"));
  else root.ZipUtils = factory(root.JSZip);
})(typeof self !== "undefined" ? self : this, function (JSZip) {
  "use strict";

  function bytesOf(value) {
    if (value instanceof Uint8Array) return value;
    if (value instanceof ArrayBuffer) return new Uint8Array(value);
    if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
    if (typeof Buffer !== "undefined" && Buffer.isBuffer(value)) return new Uint8Array(value);
    throw new Error("ZIP 中存在无法识别的文件数据");
  }

  function safeZipPath(value) {
    return String(value || "file")
      .replace(/\\/g, "/")
      .split("/")
      .filter(Boolean)
      .map((part) => part.replace(/[<>:\"|?*\x00-\x1F]/g, "_").replace(/[. ]+$/g, "_") || "file")
      .join("/");
  }

  function uniqueZipPath(path, used) {
    const safe = safeZipPath(path);
    if (!used.has(safe.toLowerCase())) {
      used.add(safe.toLowerCase());
      return safe;
    }
    const slash = safe.lastIndexOf("/");
    const folder = slash >= 0 ? safe.slice(0, slash + 1) : "";
    const filename = slash >= 0 ? safe.slice(slash + 1) : safe;
    const dot = filename.lastIndexOf(".");
    const stem = dot > 0 ? filename.slice(0, dot) : filename;
    const extension = dot > 0 ? filename.slice(dot) : "";
    let index = 2;
    let candidate;
    do candidate = `${folder}${stem} (${index++})${extension}`;
    while (used.has(candidate.toLowerCase()));
    used.add(candidate.toLowerCase());
    return candidate;
  }

  function findEocd(bytes) {
    const minimum = Math.max(0, bytes.length - 65557);
    for (let index = bytes.length - 22; index >= minimum; index -= 1) {
      if (bytes[index] === 0x50 && bytes[index + 1] === 0x4b && bytes[index + 2] === 0x05 && bytes[index + 3] === 0x06) return index;
    }
    return -1;
  }

  function validateZipBytes(input, expectedFiles) {
    const bytes = bytesOf(input);
    if (bytes.length < 22 || bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04) {
      throw new Error("压缩包文件头不完整，已停止下载");
    }
    const eocd = findEocd(bytes);
    if (eocd < 0) throw new Error("压缩包中央目录不完整，已停止下载");
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const entryCount = view.getUint16(eocd + 10, true);
    const directorySize = view.getUint32(eocd + 12, true);
    const directoryOffset = view.getUint32(eocd + 16, true);
    const commentLength = view.getUint16(eocd + 20, true);
    if (directoryOffset + directorySize > eocd || eocd + 22 + commentLength !== bytes.length) {
      throw new Error("压缩包长度校验失败，已停止下载");
    }
    if (Number.isInteger(expectedFiles) && entryCount !== expectedFiles) {
      throw new Error(`压缩包应包含 ${expectedFiles} 个文件，实际为 ${entryCount} 个`);
    }
    return { bytes, entryCount, directorySize, directoryOffset };
  }

  async function createZipBytes(entries) {
    if (!JSZip) throw new Error("批量打包组件未加载，请重新打开程序");
    if (!Array.isArray(entries) || !entries.length) throw new Error("没有可加入压缩包的文件");
    const zip = new JSZip();
    const used = new Set();
    const paths = [];
    entries.forEach((entry) => {
      const path = uniqueZipPath(entry.path || entry.filename, used);
      const bytes = bytesOf(entry.bytes);
      if (!bytes.length) throw new Error(`${path} 的文件内容为空`);
      zip.file(path, bytes, { binary: true, createFolders: false, date: new Date(2000, 0, 1) });
      paths.push(path);
    });
    const bytes = await zip.generateAsync({
      type: "uint8array",
      platform: "DOS",
      // XLSX 自身已经是 ZIP；再次压缩几乎不省空间，STORE 对 Windows
      // 资源管理器以及企业杀毒软件的兼容性更稳定。
      compression: "STORE",
      streamFiles: false
    });
    validateZipBytes(bytes, paths.length);
    const reopened = await JSZip.loadAsync(bytes, { checkCRC32: true });
    const actualFiles = Object.keys(reopened.files).filter((name) => !reopened.files[name].dir);
    if (actualFiles.length !== paths.length) throw new Error("压缩包复检时文件数量不一致");
    for (let fileIndex = 0; fileIndex < paths.length; fileIndex += 1) {
      const original = bytesOf(entries[fileIndex].bytes);
      const extracted = await reopened.file(paths[fileIndex]).async("uint8array");
      if (extracted.length !== original.length) throw new Error(`${paths[fileIndex]} 在 ZIP 复检时长度不一致`);
      for (let byteIndex = 0; byteIndex < original.length; byteIndex += 1) {
        if (extracted[byteIndex] !== original[byteIndex]) throw new Error(`${paths[fileIndex]} 在 ZIP 复检时内容不一致`);
      }
    }
    return { bytes, paths };
  }

  return { bytesOf, safeZipPath, uniqueZipPath, validateZipBytes, createZipBytes };
});
