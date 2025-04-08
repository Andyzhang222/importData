const xlsx = require("xlsx");
const fs = require("fs");
const csv = require("csv-parser");

const workbook = xlsx.readFile("1125.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

const binStartColumns = [
  0, 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76,
  80, 84, 88, 92, 96, 100, 104, 108,
];

const binMap = {};

// Step 1: 读取 bin.csv 建立 binCode => binID 的映射
fs.createReadStream("bin.csv")
  .pipe(csv())
  .on("data", (row) => {
    binMap[row.binCode] = row.binID;
  })
  .on("end", () => {
    
    const extractedData = [["binCode", "productCode", "quantity", "binID"]];

    binStartColumns.forEach((colIndex) => {
      let currentBinCode = null;

      for (let row = 0; row < data.length; row++) {
        const maybeBinCode = data[row][colIndex];
        const productCode = data[row][colIndex + 1];
        const quantity = data[row][colIndex + 2];

        // 如果 binCode 有值且不是带“区”的，就更新 currentBinCode
        if (
          maybeBinCode &&
          !String(maybeBinCode).includes("区") &&
          String(maybeBinCode).trim() !== ""
        ) {
          currentBinCode = maybeBinCode;
        }

        // 如果没有 binCode 或者 productCode 或者数量，跳过
        if (
          !currentBinCode ||
          !productCode ||
          !quantity ||
          String(productCode).trim() === "" ||
          String(quantity).trim() === ""
        ) {
          continue;
        }

        const binID = binMap[currentBinCode] || "";
        extractedData.push([currentBinCode, productCode, quantity, binID]);
      }
    });

    // 输出 Excel 文件
    const newSheet = xlsx.utils.aoa_to_sheet(extractedData);
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Result");
    xlsx.writeFile(newWorkbook, "final_output_with_binID.xlsx");

    console.log("✅ 导出完成！文件名：final_output_with_binID.xlsx");
  });
