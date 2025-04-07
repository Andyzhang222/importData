const xlsx = require("xlsx");

const workbook = xlsx.readFile("1125.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// 这些是目标列的索引（A=0, E=4, I=8, ...）
const columnsToExtract = [
  0, 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76,
  80, 84, 88, 92, 96, 100, 104, 108,
];

const singleColumnData = [];

columnsToExtract.forEach((colIndex) => {
  for (let row = 0; row < data.length; row++) {
    const value = data[row][colIndex];
    if (value !== undefined && value !== "" && !String(value).includes("区")) {
      singleColumnData.push([value]); // 每一行是一个数组，形成单列
    }
  }
});

const newSheet = xlsx.utils.aoa_to_sheet(singleColumnData);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "SingleColumn");
xlsx.writeFile(newWorkbook, "output_single_column.xlsx");

console.log("✅ 单列提取完成！文件名：output_single_column.xlsx");
