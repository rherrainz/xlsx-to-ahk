import xlsx from "xlsx";
import fs from "fs";
import cp from "child_process";

//se busca el archivo y la hoja del excel que trabajamos
const workbook = xlsx.readFile("test.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

//Asigna la información del excel a un objeto
const excelData = xlsx.utils.sheet_to_json(worksheet, { raw: true });

//Función para obtener la fecha actual y mostrarla en el formato correcto
const getFechaDesde = () => {
  const date = new Date();
  let day = date.getDate() + 1;
  if (day < 10) {
    day = `0${day}`;
  }
  let month = date.getMonth() + 1;
  if (month < 10) {
    month = `0${month}`;
  }
  const year = date.getFullYear();
  const fechaDesde = `${day}${month}${year}`;
  return fechaDesde;
};

const fechaDesde = getFechaDesde();

// variable donde guarda el texto
let ahkScript = "";
let currentSuc = 0;

// arreglo donde van las teclas que vamos a usar en el script
const keysArray = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"];
let keysCounter = 0;
let sucCounter = [];

//recorremos el objeto y vamos armando el script
excelData.forEach((row) => {
  const desc = row.desc * 100;
  const cod =  row.cod;
  let strCod = cod.toString();
  let paddedCod = strCod.padStart(7, "0");
  if (currentSuc === 0) {
    ahkScript += `^${keysArray[keysCounter]}::{\n`;
    currentSuc = row.Suc;
    sucCounter.push({"suc":row.Suc,"keys": `ctrl + ${keysArray[keysCounter]}`});
  } else if (currentSuc !== row.Suc) {
    keysCounter++;
    currentSuc = row.Suc;
    sucCounter.push({"suc":row.Suc,"keys": `ctrl + ${keysArray[keysCounter]}`});
    ahkScript += `Return\n}\n^${keysArray[keysCounter]}::{\n`;
  }
  const fechaHasta = xlsx.SSF.format("ddmmyyyy", row.Hasta);
  //console.log(row.Suc, row.cod, desc, fechaHasta);
  ahkScript += `SendText "${paddedCod}"\nSend "{Tab Down}"\nSend "{Tab Up}"\nSendText "${fechaDesde}"\nSend "{Tab Down}"\nSend "{Tab Up}"\nSend "{Tab Down}"\nSend "{Tab Up}"\nSendText "${fechaHasta}"\nSend "{Enter Down}"\nSend "{Enter Up}"\nSleep 500\nSendText "${desc}"\nSend "{Enter Down}"\nSend "{Enter Up}"\nSleep 500\n`;
});

//final del scropt
ahkScript += `Return\n}`;
//se escribe el string en el archivo del script
fs.writeFileSync("datos.ahk", ahkScript);
//se escribe el string en un json (para verificar que se haya generado correctamente)
fs.writeFileSync("datos.json", JSON.stringify(excelData));

//avisamos que se generó bien el archivo
console.log("Archivo generado en: datos.ahk");
//mostramos una tabla con los comandos del ahk
console.table(sucCounter);

//ejecutamos el script
cp.exec("datos.ahk", (err, stdout, stderr) => {
  if (err) {
    console.error(err);
    return;
  }
  stdout ? console.log(`stdout: ${stdout}`) : console.log("Archivo ejecutado correctamente");
  stderr ? console.error(`stderr: ${stderr}`) : null;
});