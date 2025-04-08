const xlsx = require("xlsx");

// 读取 Excel 文件
const workbook = xlsx.readFile("PICK.xls"); // 改成你的文件名
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

// 用 Map 去重 binCode
const uniqueBins = new Map();

data.forEach((row) => {
  const binCode = row.binCode;
  if (!uniqueBins.has(binCode)) {
    uniqueBins.set(binCode, {
      binCode,
      type: row.type,
      warehouseID: row.warehouseID,
      createdAt: row.createdAt,
      updatedAt: row.updatedAt,
    });
  }
});

// 转成数组写入新 Excel
const cleanedData = Array.from(uniqueBins.values());

const newSheet = xlsx.utils.json_to_sheet(cleanedData);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "UniqueBins");

xlsx.writeFile(newWorkbook, "deduplicated_bins.xlsx");

console.log("✅ Done! Saved to deduplicated_bins.xlsx");
