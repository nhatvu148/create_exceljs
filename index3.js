const Excel = require("exceljs");
const workbook = new Excel.Workbook();

const url = "./template.xlsx";

workbook.xlsx.readFile(url).then(function() {
  const worksheet = workbook.getWorksheet(2);

  const row = worksheet.getRow(8);
  row.getCell(2).value = "John Doe";
  row.getCell(3).value = new Date(1970, 1, 1);
  row.getCell(4).value = 1234;
  row.commit();
  return workbook.xlsx.writeFile(url);
});
