const XLSX = require("xlsx-js-style");
for (const filename of ["templates/default-assembly.xlsx", ".qa-default-templates/generated-defaults.xlsx", "tests/tmp/built-in-default-output.xlsx"]) {
  const wb = XLSX.readFile(filename, { cellStyles: true });
  console.log(`\n${filename}`, wb.SheetNames);
  const ws = wb.Sheets[filename.includes("default-assembly") ? wb.SheetNames[0] : wb.SheetNames[1]];
  for (const address of ["A1", "A5", "A7", "B7"]) console.log(address, JSON.stringify(ws[address]));
}
