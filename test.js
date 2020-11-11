"user strict";
const Excel = require("exceljs");

const workbook = new Excel.Workbook();
workbook.xlsx.readFile("limit.xlsx").then(()=>{
  let worksheet = workbook.getWorksheet("data");


  let column = worksheet.actualColumnCount;
  let row = worksheet.actualRowCount;
  
  console.log("actual");
  console.log(column);
  console.log(row);
  column = worksheet.columnCount;
  row = worksheet.rowCount;
  
  console.log("count");
  console.log(column);
  console.log(row);

  
  // worksheet.spliceRows(11,2);
  // worksheet.spliceColumns(4,2);
  worksheet.unMergeCells("A1:Z1");
  workbook.xlsx.writeFile("limit2.xlsx");
  
});
