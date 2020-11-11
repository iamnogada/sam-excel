"user strict";
const Excel = require("exceljs");
const meta = require("./input");

/*
Excel Create
*/
let workbook = new Excel.Workbook();
workbook.xlsx.readFile("sample.xlsx").then(()=>{
  let ws = workbook.getWorksheet(1);
  if(undefined === ws) {
    process.abort();
  }

  ws.unMergeCells("A1:A3");
  ws.insertRow(2,1);
  cleanHeader(ws);
  console.log(`col:${ws.columnCount}`);
  cleanNo(ws);
  console.log(`col:${ws.columnCount}`);
  cleanRow(ws);
  cleanColumn(ws);
  
  workbook.xlsx.writeFile("result.xlsx");
});

const cleanHeader = (ws)=>{
  for(let i=1;i<=ws.columnCount;i++){
    ws.getRow(2).getCell(i).value = ws.getRow(1).getCell(i).value;
  }
  ws.spliceRows(1,1);
}
const cleanNo = (ws)=>{
  let tag = `${ws.getRow(1).getCell(1).value}`
  if("no" == tag.toLowerCase()){
    ws.spliceColumns(1,1);
  }
}

let msg =[];
const cleanRow = (ws) => {
  let count = ws.rowCount;
  let deleteCount =0;
  for(let i=count;i>3;i--){
    if(0 == ws.getRow(i).actualCellCount){
      ws.spliceRows(i,1);
      deleteCount++;
    }
  }
  if(0 !== deleteCount){
    msg.push(`Exclude empty row: ${deleteCount}`);
  }
  
}
const cleanColumn = (ws) => {
  let count = ws.columnCount;
  let deleteCount =0;
  for(let i=count;i>0;i--){
    if(4 ==ws.getColumn(i).values.length){
      ws.spliceColumns(i,1);
      deleteCount++;
    }
    
  }
  if(0 !== deleteCount){
    msg.push(`Exclude empty column: ${deleteCount}`);
  }
  
}