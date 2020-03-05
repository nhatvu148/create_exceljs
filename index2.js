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

sheet.addTable({
  name: "MyTable",
  ref: "B7",
  headerRow: true,
  totalsRow: true,
  style: {
    // theme: "TableStyleDark3",
    showRowStripes: true
  },
  columns: [
    {
      name: "Project ID",
      totalsRowLabel: "Total Work Load:",
      filterButton: false
    },
    { name: "Project Name", filterButton: false },
    {
      name: "Deadline",
      //   totalsRowLabel: "Total Work Load:",
      filterButton: false
    },
    {
      name: "Expected Completed Date",
      //   totalsRowLabel: "Total Work Load:",
      filterButton: false
    },
    {
      name: "Percent  Completed",
      //   totalsRowLabel: "Total Work Load:",
      filterButton: false
    },
    { name: "Working time (h)", totalsRowFunction: "sum", filterButton: false },
    { name: "Comment", totalsRowLabel: "Total Work Load:", filterButton: false }
  ],
  rows: [
    [200012, "Jupiter V4.1", 20191201, 20191201, 0, 16, "Some comments"],
    [200012, "Jupiter V4.1", 20191201, 20191201, 0, 16, "Some comments"],
    [200012, "Jupiter V4.1", 20191201, 20191201, 0, 16, "Some comments"]
  ]
});

sheet.getCell("B4").border = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" }
};

const sheet2 = workbook.addWorksheet("My 2nd Sheet");
console.log(workbook);

workbook.xlsx.writeFile("My_WB2.xlsx").then(function() {
  console.log("Success");
});
