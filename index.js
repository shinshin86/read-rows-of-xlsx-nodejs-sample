const XLSX = require("xlsx");
const Utils = XLSX.utils;
const fileName = process.argv[2];
const sheetName = process.argv[3];
const targetRowIndex = process.argv[4];
const targetItem = process.argv[5];

const book = XLSX.readFile(fileName);
const sheet = book.Sheets[sheetName];
const range = sheet["!ref"];
const decodeRange = Utils.decode_range(range);

let count = 1;
let resultObj = {};

for (let rowIndex = decodeRange.s.r; rowIndex <= decodeRange.e.r; rowIndex++) {
  const address = Utils.encode_cell({ r: rowIndex, c: targetRowIndex });
  if (typeof sheet[address] !== "undefined") {
    if (sheet[address].h.match(targetItem)) {
      console.log(sheet[address].h);
      resultObj["name"] = sheet[address].h;
      for (
        let colIndex = decodeRange.s.c;
        colIndex <= decodeRange.e.c;
        colIndex++
      ) {
        const address = Utils.encode_cell({ r: rowIndex, c: colIndex });
        const cell = sheet[address];
        if (typeof cell !== "undefined") {
          resultObj[count++] = cell.h;
        }
      }
    }
  }
}

console.log(resultObj);
