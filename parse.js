"user strict";
const Excel = require("exceljs");
const Master = require("./master.js");
/*
Excel Create
*/
let workbook = new Excel.Workbook();
let ws={};
let msg =[];
const headerHeight = 4;
const property = "Property";

workbook.xlsx.readFile("uploaded.xlsx").then(()=>{
  ws = workbook.getWorksheet(1);
  if(undefined === ws) {
    process.abort();
  }

  // Excel Tool
  ws.cell = function (col, row) {
    return this.getRow(row).getCell(col);
  }.bind(ws);
  ws.address = function (col, row) {
    return this.getRow(row).getCell(col).address;
  }.bind(ws);
  ws.row = function (row) {
    return this.getRow(row);
  }.bind(ws);
  ws.col = function (col) {
    return this.getColumn(col);
  }.bind(ws);

  //1. unmerge head
  cleanHeader();
  //2. clean row
  cleanRow();
  //3. clean column
  cleanColumn();
  //4. remove no column
  cleanNo();
  //5. extract meta
  let manifest = getManifest();
  if(!manifest) process.abort();
  //6. verify meta
  let msg = verify(Master,manifest);
  console.log(msg);
  //7. export csv
  
  workbook.xlsx.writeFile("result.xlsx");
  ws.spliceRows(1,3);
  workbook.csv.writeFile("result.csv");
});

const cleanHeader = ()=>{
  let header = ws.row(1).values;
  ws.unMergeCells(`${ws.address(1,1)}:${ws.address(ws.columnCount,1)}`);
  ws.row(1).values = header;
}

const cleanNo = ()=>{
  let tag = `${ws.getRow(1).getCell(1).value}`
  if("no" == tag.toLowerCase()){
    ws.spliceColumns(1,1);
  }
}


const cleanRow = () => {
  let count = ws.rowCount;
  let deleteCount =0;
  for(let i=count;i>headerHeight;i--){
    if(1 == ws.getRow(i).actualCellCount){
      ws.spliceRows(i,1);
      deleteCount++;
    }
  }
  if(0 !== deleteCount){
    msg.push(`Exclude empty row: ${deleteCount}`);
  }
  
}
const cleanColumn = () => {
  let count = ws.columnCount;
  let deleteCount =0;
  for(let i=count;i>0;i--){
    
    if(5 == ws.getColumn(i).values.length){
      ws.spliceColumns(i,1);
      deleteCount++;
    }
    
  }
  if(0 !== deleteCount){
    msg.push(`Exclude empty column: ${deleteCount}`);
  }
}

const getManifest = ()=>{
  let manifest = {};
  for(let i=1;i<=ws.columnCount;i++){
    let group = ws.cell(i,1).value;
    if(!group) continue;
    if(!manifest[group]) manifest[group]= [];
    manifest[group].push({
      displayname:ws.cell(i,2).value,
      tag:ws.cell(i,3).value,
      key:ws.cell(i,4).value
    });
  }
  return manifest;
}

const verify = (master, manifest)=>{
  let msg=[];
  for (const [key, value] of Object.entries(manifest)) {
    let data = master[key];
    if(!data){
      msg.push(`Group ${key} is not allowed`);
      continue;
    }
    value.forEach(element => {
      if(!findKey(data,element.key)) msg.push(`Key ${element.key} is not allowed`)
    });
  }
  return msg;
}
const findKey=(master, key)=>{
  for(let i=0;i<master.length;i++){
    if(key == master[i].key) {
      return true;}
  }
  return false;
}