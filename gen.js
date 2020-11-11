"user strict";
const Excel = require("exceljs");
const meta = require("./input");

/*
Excel Create
*/
let workbook = new Excel.Workbook();
let ws = workbook.addWorksheet("dataset");
// ws.properties.defaultRowHeight = 30;

const templateRow = 100;
const border = {
  right: { style: "thin" },
  bottom: { style: "thin" }
};
// const pos = (col, row) => {
//   return `${String.fromCharCode(64 + col)}${row}`;
// };

const _pos = (ws)=>{
  return (col,row)=>{
    return ws.getRow(row).getCell(col).address;
  }
}
const pos = _pos(ws);

const displayHeader = (index, label, data) => {
  return (ws) => {
    ws.mergeCells(`${pos(index, 1)}:${pos(index + data.length - 1, 1)}`);
    let row = ws.getRow(1).getCell(index);
    row.value = label;
    row.border = border;
    for (let i = 0; i < data.length; i++) {
      let cell = ws.getRow(2).getCell(i + index);
      cell.value = `${data[i].company}`;
      cell.border = border;
      cell = ws.getRow(3).getCell(i + index);
      cell.value = `${data[i].data}`;
      cell.border = border;
    }
  };
};

const displayProperty = (index, label, data) => {
  return (ws) => {
    ws.mergeCells(`${pos(index, 1)}:${pos(index + data.length - 1, 1)}`);
    let row = ws.getRow(1);
    row.getCell(index).value = label;
    row.getCell(index).border = border;

    for (let i = 0; i < data.length; i++) {
      let cell = ws.getRow(2).getCell(i + index);
      cell.value = data[i].unit;//`${data[i].abbr}(${data[i].unit})`;
      cell.border = border;
      cell = ws.getRow(3).getCell(i + index);
      cell.value = `${data[i].name}`;
      cell.border = border;
    }
  };
};
const displayNo = (index, label) => {
  return (ws) => {
    ws.getRow(1).getCell(index).value = label;
    ws.mergeCells(`${pos(index, 1)}:${pos(index, 3)}`);
    ws.getRow(1).getCell(1).border = border;
  };
};
// No Columns
let header = [];
header.push(displayNo(1, "No"));
ws.getColumn(1).border = {
  right: { style: "thin" },
};
let index = 2;
// Property Columns
header.push(displayProperty(2, "Property", meta.label));
index += meta.label.length;
ws.getColumn(meta.label.length + 1).border = {
  right: { style: "thin" },
};
// Feature Columns
meta.feature.forEach((element) => {
  header.push(displayHeader(index, element.group, element.grade));
  index += element.grade.length;
  ws.getColumn(index - 1).border = {
    right: { style: "thin" },
  };
});

// console.log(header);

ws.getRow(1).alignment = {
  wrapText: true,
  vertical: "middle",
  horizontal: "center",
};
ws.getRow(2).alignment = {
  wrapText: true,
  vertical: "middle",
  horizontal: "center",
};
ws.getRow(3).alignment = {
  wrapText: true,
  vertical: "middle",
  horizontal: "center",
};
ws.getRow(1).height =30;
ws.getRow(2).height =30;
ws.getRow(3).height =30;
ws.getRow(1).font={bold:true};

header.forEach((element) => {
  element(ws);
});
for(let i=1;i<=templateRow;i++){
  let cell = ws.getRow(i+3).getCell(1);
  cell.value =i;
  cell.alignment = {horizontal: "center"}
}


workbook.xlsx.writeFile("sample.xlsx");
