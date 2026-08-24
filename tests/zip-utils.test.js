const assert = require("node:assert/strict");
const JSZip = require("jszip");
const ZipUtils = require("../zip-utils.js");

async function run() {
  const one = new Uint8Array([0x50, 0x4b, 1, 2, 3, 4]);
  const two = new Uint8Array([9, 8, 7, 6, 5, 4, 3, 2, 1]);
  const packed = await ZipUtils.createZipBytes([
    { path: "BOM-A/装配清单.xlsx", bytes: one },
    { path: "BOM-A/装配清单.xlsx", bytes: two },
    { path: "BOM-B/采购清单.xlsx", bytes: one },
  ]);
  const validation = ZipUtils.validateZipBytes(packed.bytes, 3);
  assert.equal(validation.entryCount, 3, "ZIP 中央目录应包含三个文件");
  assert.deepEqual(packed.paths, ["BOM-A/装配清单.xlsx", "BOM-A/装配清单 (2).xlsx", "BOM-B/采购清单.xlsx"], "同名文件应自动编号");
  const reopened = await JSZip.loadAsync(packed.bytes, { checkCRC32: true });
  const names = Object.keys(reopened.files).filter((name) => !reopened.files[name].dir);
  assert.deepEqual(names, packed.paths, "ZIP 复检后路径应完整且顺序一致");
  assert.deepEqual(await reopened.file(packed.paths[1]).async("uint8array"), two, "ZIP 内容应逐字节一致");
  const localHeader = packed.bytes;
  assert.equal(localHeader[8], 0, "多清单 ZIP 应使用 STORE 模式的压缩方法低字节");
  assert.equal(localHeader[9], 0, "多清单 ZIP 应使用 STORE 模式的压缩方法高字节");
  const packedAgain = await ZipUtils.createZipBytes([
    { path: "BOM-A/装配清单.xlsx", bytes: one },
    { path: "BOM-A/装配清单.xlsx", bytes: two },
    { path: "BOM-B/采购清单.xlsx", bytes: one },
  ]);
  assert.deepEqual(packedAgain.bytes, packed.bytes, "固定时间戳和 STORE 模式应生成完全确定的 ZIP 字节");

  const truncated = packed.bytes.slice(0, packed.bytes.length - 8);
  assert.throws(() => ZipUtils.validateZipBytes(truncated, 3), /中央目录|长度/, "截断的 ZIP 必须被拦截");
  console.log("zip creation, duplicate naming and integrity tests passed");
}

run().catch((error) => { console.error(error); process.exitCode = 1; });
