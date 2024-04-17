import xlsx from "xlsx";
import fs from "fs";

const workbook = xlsx.readFile("test.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const excelData = xlsx.utils.sheet_to_json(worksheet, { raw: true });

const getFechaDesde = () => {
  const date = new Date();
  const day = date.getDate() + 1;
  let month = date.getMonth() + 1;
  if (month < 10) {
    month = `0${month}`;
  }
  const year = date.getFullYear();
  const fechaDesde = `${day}${month}${year}`;
  return fechaDesde;
};

const fechaDesde = getFechaDesde();
let ahkScript = "";
let currentSuc = 0;
const keysArray = [
  "a",
  "b",
  "c",
  "d",
  "e",
  "f",
  "g",
  "h",
  "i",
  "j",
  "k",
  "l",
  "m",
  "n",
  "o",
  "p",
  "q",
  "r",
  "s",
  "t",
  "u",
  "v",
  "w",
  "x",
  "y",
  "z",
];
let keysCounter = 0;
let sucCounter = [];

excelData.forEach((row, i) => {
  const desc = row.desc * 100;
  if (currentSuc === 0) {
    ahkScript += `^${keysArray[keysCounter]}::{\n`;
    currentSuc = row.Suc;
    sucCounter.push({"suc":row.Suc,"keys": `ctrl + ${keysArray[keysCounter]}`});
  } else if (currentSuc !== row.Suc) {
    keysCounter++;
    currentSuc = row.Suc;
    sucCounter.push({"suc":row.Suc,"keys": `ctrl + ${keysArray[keysCounter]}`});
    ahkScript += `Return\n
                  }\n
                  ^${keysArray[keysCounter]}::{\n`;
  }
  const fechaHasta = xlsx.SSF.format("ddmmyyyy", row.Hasta);
  //console.log(row.Suc, row.cod, desc, fechaHasta);
  ahkScript += `SendText "${row.cod}"\n
                Send "{Tab Down}"\n
                Send "{Tab Up}"\n
                SendText "${fechaDesde}"\n
                Send "{Tab Down}"\n
                Send "{Tab Up}"\n
                SendText "${fechaHasta}"\n
                Send "{Enter Down}"\n
                Send "{Enter Up}"\n               
                Sleep 500\n
                SendText "${desc}"\n
                Send "{Enter Down}"\n
                Send "{Enter Up}"\n
                Sleep 500\n           
  `;
});

ahkScript += `Return\n
}`;
fs.writeFileSync("datos.ahk", ahkScript);
fs.writeFileSync("datos.json", JSON.stringify(excelData));
console.log("Archivo generado en: datos.ahk");
console.table(sucCounter);
