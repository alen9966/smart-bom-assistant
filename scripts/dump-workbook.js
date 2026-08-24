const XLSX = require("xlsx-js-style");

const source = process.argv[2];
if (!source) throw new Error("Please provide an XLSX path.");
const workbook = XLSX.readFile(source, { raw: false });
for (const name of workbook.SheetNames) {
  console.log(`### ${name}`);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: false, defval: "" });
  rows.forEach((row, index) => console.log(JSON.stringify([index + 1, ...row])));
}
