const xlsx = require("xlsx");

const workbook = xlsx.readFile("1125.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// 这些是 productCode 所在列的索引（B=1, F=5, J=9, ...）
const productCodeColumns = [
  1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77,
  81, 85, 89, 93, 97, 101, 105, 109,
];

const productCodeSet = new Set();

productCodeColumns.forEach((colIndex) => {
  for (let row = 0; row < data.length; row++) {
    const productCode = data[row][colIndex];
    if (
      productCode &&
      typeof productCode === "string" &&
      !productCode.includes("区")
    ) {
      productCodeSet.add(productCode.trim());
    }
  }
});

// 构建数据行
const outputData = [["productCode", "createdAt", "updatedAt"]];
const fixedDate = "2025-04-07 15:32";

productCodeSet.forEach((code) => {
  outputData.push([code, fixedDate, fixedDate]);
});

// 写入新 Excel 文件
const newSheet = xlsx.utils.aoa_to_sheet(outputData);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Products");
xlsx.writeFile(newWorkbook, "products_output.xlsx");

console.log("✅ 产品提取完成！文件名：products_output.xlsx");
