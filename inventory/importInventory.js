const fs = require("fs");
const xlsx = require("xlsx");
const path = require("path");

// 读取 Excel 文件
const importDataFromExcel = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet);
  return data;
};

// 生成 SQL 插入语句
const generateSeedData = (data) => {
  const seedData = [];

  data.forEach((item) => {
    seedData.push({
      productCode: item.productCode,
      binID: item.binID,
      quantity: item.quantity,
      createdAt: new Date(), // 使用当前时间
      updatedAt: new Date(), // 使用当前时间
    });
  });

  // 将数据转换为 Sequelize 插入格式
  return `module.exports = {
  up: async (queryInterface, Sequelize) => {
    await queryInterface.bulkInsert('Inventory', ${JSON.stringify(
      seedData,
      null,
      2
    )});
  },

  down: async (queryInterface, Sequelize) => {
    await queryInterface.bulkDelete('Inventory', null, {});
  }
};`;
};

// 执行生成操作
const generateSeedFile = async (filePath) => {
  const data = await importDataFromExcel(filePath);

  const seedData = generateSeedData(data);
  const outputPath = path.resolve(__dirname, "seed_inventory.js");

  fs.writeFileSync(outputPath, seedData, "utf8");

  console.log("Seed file generated at:", outputPath);
};

// 调用生成 seed 文件
const filePath = path.resolve(__dirname, "inventory.xlsx"); // 修改为你的 Excel 文件路径
generateSeedFile(filePath);
