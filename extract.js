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
    const rowData = data[row];

    const value = rowData[colIndex];

    if (value === undefined || value === null) {
      continue; // 空值，跳过
    }

    const valueStr = String(value).trim();

    if (
      valueStr === "" || // 空字符串
      /[\u4e00-\u9fa5]/.test(valueStr) || // 有中文
      valueStr.includes("区") // 包含"区"
    ) {
      continue; // 只跳过这一个 cell，不跳整行
    }

    singleColumnData.push([valueStr]);
  }
});

const newSheet = xlsx.utils.aoa_to_sheet(singleColumnData);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "SingleColumn");
xlsx.writeFile(newWorkbook, "1125c.xlsx");

console.log("✅ 指定列处理完成，导出成功！");
