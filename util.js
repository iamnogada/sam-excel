_pos = (ws) =>{
    return (col,row) =>{
        ws.getRow(row).getCell(col).address;
    }
}
_row = (ws) =>{
    return (row) =>{
        console.log(`ROW:${row}`)
        ws.getRow(row);
    }
}

module.exports = {_pos,_row}