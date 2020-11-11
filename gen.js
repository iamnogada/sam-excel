"user strict";
const Excel = require("exceljs");
const meta = require("./input");
/*
Excel Create
*/
let workbook = new Excel.Workbook();
let ws = workbook.addWorksheet("dataset");

const templateRow = 100;
const headerHeight = 4;
const property = "Property";
const border = {
  right: { style: "thin" },
  bottom: { style: "thin" },
};

// Excel 작업을 위한 Tool object
let CURRENT = 1;
ws.cell = function (col, row) {
  return this.getRow(row).getCell(col);
}.bind(ws);
ws.address = function (col, row) {
  return this.getRow(row).getCell(col).address;
}.bind(ws);
ws.row = function (row) {
  return this.getRow(row);
}.bind(ws);

const display = (group,header) => {
  let length = header.length;
  //group name
  ws.cell(CURRENT, 1).value = group;
  // merge
  ws.mergeCells(
    `${ws.address(CURRENT, 1)}:${ws.address(CURRENT + length - 1, 1)}`
  );
  ws.cell(CURRENT, 1).border=border;
  for (let i = 0; i < length; i++) {
    ws.cell(CURRENT + i, 2).value = header[i].displayname;
    ws.cell(CURRENT + i, 2).border = border;
    ws.cell(CURRENT + i, 3).value = header[i].tag;
    ws.cell(CURRENT + i, 3).border = border;
    ws.cell(CURRENT + i, 4).value = header[i].key;
    ws.cell(CURRENT + i, 4).border = border;
  }
  CURRENT += length;
};

//1. Display No Column and merge cells
ws.cell(CURRENT, 1).value = "No";
ws.mergeCells(`${ws.address(CURRENT, 1)}:${ws.address(CURRENT, 4)}`);
ws.cell(CURRENT, 1).border=border;
CURRENT = 2;

//2. Display Property Column
if (undefined === meta[property]) console.log("Error");
display(property, meta[property]);


//3. Display Feature Column
Object.keys(meta).forEach((key)=>{
  if(property != key){
    display(key, meta[key]);
  }
});

//4. Set Style: Group Column Right Broder, Header height
for(let i=1;i<=4;i++){
  ws.row(i).alignment = {
    wrapText: true,
    vertical: "middle",
    horizontal: "center",
  };
}
ws.row(1).font={bold:true};
ws.row(1).height=30;


//5. Numbering in No Column

// ws.getRow(1).alignment = {
//   wrapText: true,
//   vertical: "middle",
//   horizontal: "center",
// };
// ws.getRow(2).alignment = {
//   wrapText: true,
//   vertical: "middle",
//   horizontal: "center",
// };
// ws.getRow(3).alignment = {
//   wrapText: true,
//   vertical: "middle",
//   horizontal: "center",
// };
// ws.getRow(1).height =30;
// ws.getRow(2).height =30;
// ws.getRow(3).height =30;
// ws.getRow(1).font={bold:true};

// header.forEach((element) => {
//   element(ws);
// });
// for(let i=1;i<=templateRow;i++){
//   let cell = ws.getRow(i+3).getCell(1);
//   cell.value =i;
//   cell.alignment = {horizontal: "center"}
// }

workbook.xlsx.writeFile("sample.xlsx");

// const displayHeader = (index, label, data) => {
//   return (ws) => {
//     ws.mergeCells(`${pos(index, 1)}:${pos(index + data.length - 1, 1)}`);
//     let row = ws.getRow(1).getCell(index);
//     row.value = label;
//     row.border = border;
//     for (let i = 0; i < data.length; i++) {
//       let cell = ws.getRow(2).getCell(i + index);
//       cell.value = `${data[i].company}`;
//       cell.border = border;
//       cell = ws.getRow(3).getCell(i + index);
//       cell.value = `${data[i].data}`;
//       cell.border = border;
//     }
//   };
// };

// const displayNo = (index, label) => {
//   return (ws) => {
//     ws.getRow(1).getCell(index).value = label;
//     ws.mergeCells(`${pos(index, 1)}:${pos(index, 3)}`);
//     ws.getRow(1).getCell(1).border = border;
//   };
// };
