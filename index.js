const Excel = require("exceljs");

const workbook = new Excel.Workbook();

// const readmyFile = async () => {
//   const myWb = await workbook.xlsx.readFile("myfile.xlsx");
//   console.log(myWb._worksheets[30]);
// };

// readmyFile();

const sheet = workbook.addWorksheet("Hello", {
  views: [{ showGridLines: true }]
});
sheet.columns = [
  { width: 8.29 },
  { width: 11.57 },
  { width: 22.71 },
  { width: 8.29 },
  { width: 8.29 },
  { width: 14 },
  { width: 9.43 },
  { width: 9.29 },
  { width: 8.29 },
  { width: 8.29 },
  { width: 8.29 },
  { width: 8.29 }
];

for (let i = 6; i < 14; i++) {
  sheet.getRow(i).height = 39;
}

sheet.getColumn("B").values = [
  "Time Sheet",
  null,
  "DIV / DEP",
  "Development",
  null,
  "Project ID",
  20012,
  20012,
  20012,
  20012,
  null,
  20022,
  20022,
  null,
  "Total Work Load"
];

sheet.getColumn("C").values = [
  null,
  null,
  "Name",
  "Kitami",
  null,
  "Project Name",
  "Jupiter V4.1",
  "Jupiter V4.1",
  "Jupiter V4.1",
  "Jupiter V4.1",
  null,
  "Oasis",
  "Oasis"
];

sheet.getColumn("D").values = [
  null,
  null,
  null,
  null,
  null,
  "Deadline",
  20191201,
  20191201,
  20191201,
  20191201,
  null,
  20200301,
  20200301
];

sheet.getColumn("F").values = [
  null,
  null,
  null,
  null,
  null,
  "Expected Completed Date",
  20191201,
  20191201,
  20191201,
  20191201,
  null,
  20200301,
  20200301
];

sheet.getColumn("G").values = [
  null,
  null,
  null,
  null,
  null,
  "Percent Completed",
  0,
  0,
  0,
  0,
  null,
  0,
  0
];

sheet.getColumn("H").values = [
  null,
  null,
  "YEAR",
  2020,
  null,
  "Working time (h)",
  16,
  8,
  3,
  5,
  null,
  3,
  4.5,
  null,
  39.5
];

sheet.getColumn("I").values = [
  null,
  null,
  "MONTH",
  2,
  null,
  "Comments",
  "[MeshCleanup - Bug #10675] Crash happens Mesh Statistic Check.",
  "CM2 Meshers integration.",
  "[Home - Feature #10652] Preferences > Default Setting: Key Bind for Collapse",
  "[Home - Feature #10652] Preferences > Default Setting: Key Bind for Collapse",
  null,
  "[Oasis-A - Feature #10640] (Feedback) Base Model V2 | Mount",
  "[Oasis-A - Feature #10641] (Feedback) Base Model V2 | Load Case"
];

sheet.getColumn("J").values = [null, null, "DAY", 17];
sheet.getColumn("K").values = [null, null, "TO", 21];

sheet.getCell("B4").border = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" }
};

const sheet2 = workbook.addWorksheet("My 2nd Sheet");
console.log(workbook);

workbook.xlsx.writeFile("My_WB.xlsx").then(function() {
  console.log("Success");
});
