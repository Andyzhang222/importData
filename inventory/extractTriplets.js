const xlsx = require("xlsx");
const workbook = xlsx.readFile("1125.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

const colGroups = [
  0, 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76,
  80, 84,
];

const extracted = [];
const binCodeSet = new Set();

for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
  const row = data[rowIndex];
  if (!row) continue;

  for (let col of colGroups) {
    let binCode = row[col];
    const productCode = row[col + 1];
    const quantity = row[col + 2];

    // ✅ 如果 binCode 为空，继承上一行对应列的 binCode
    if ((binCode === undefined || binCode === "") && rowIndex > 0) {
      const prevRow = data[rowIndex - 1];
      binCode = prevRow?.[col] || "";
    }

    // ✅ 提取有效数据（输出到 filtered_output.xlsx）
    if (
      typeof binCode === "string" &&
      binCode.trim() !== "" &&
      !binCode.includes("区") &&
      productCode &&
      quantity
    ) {
      extracted.push([binCode, productCode, quantity]);
    }

    // ✅ 收集所有 binCode（无论是否有产品）
    if (
      typeof row[col] === "string" &&
      row[col].trim() !== "" &&
      !row[col].includes("区")
    ) {
      binCodeSet.add(row[col]);
    }
  }
}

// ✅ 写入 filtered_output.xlsx
const filteredSheet = xlsx.utils.aoa_to_sheet([
  ["binCode", "productCode", "quantity"],
  ...extracted,
]);
const filteredBook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(filteredBook, filteredSheet, "Filtered");
xlsx.writeFile(filteredBook, "filtered_output.xlsx");

// ✅ 写入 1125binCode.xlsx（每个 binCode 一行，默认 quantity 为 0）
const binCodeList = Array.from(binCodeSet).map((code) => [code, 0]);
const binSheet = xlsx.utils.aoa_to_sheet([
  ["binCode", "quantity"],
  ...binCodeList,
]);
const binBook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(binBook, binSheet, "AllBinCodes");
xlsx.writeFile(binBook, "1125binCode.xlsx");

console.log(`✅ 提取完成：
- 有效记录行数: ${extracted.length}
- 唯一 binCode 数量: ${binCodeSet.size}`);
