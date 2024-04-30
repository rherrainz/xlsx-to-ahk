import xlsx from "xlsx";
import fs from "fs";
import cp from "child_process";

//se busca el archivo y la hoja del excel que trabajamos
const workbook = xlsx.readFile("test.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

//Asigna la información del excel a un objeto
const excelData = xlsx.utils.sheet_to_json(worksheet, { raw: true });

//Función para obtener la fecha cuando comienza la acción (mañana) y mostrarla en el formato correcto
const getFechaDesde = () => {
  const date = new Date();
  console.log(date);
  let day = date.getDate() + 1;

  let month = date.getMonth() + 1;
  let year = date.getFullYear();
  
  if (day > 31) {
    day = 1;
    month = month + 1;
  }
  if ((day>30)&&(month==4||month==6||month==9||month==11)){
    day = 1;
    month = month + 1;    
  }
  if ((day>28)&&(month==2)){
    day = 1;
    month = month + 1;    
  }
  if (month > 12) {
    month = 1;
    year++;
  }
  if (day < 10) {
      day = `0${day}`;
    }
  if (month < 10) {
      month = `0${month}`;
    }
  const fechaDesde = `${day}${month}${year}`;
  return fechaDesde;
};

const fechaDesde = getFechaDesde();

// variable donde guarda el texto
let ahkScript = "";
let currentSuc = 0;

// arreglo donde van las teclas que vamos a usar en el script
const keysArray = ["1","2","3","4","5","6","7","8","9","0","a","b","d","e","f","g","h","i","j","k","l","m","n","o","r","s","t","u","w","y","z"];
let keysCounter = 0;
let sucCounter = [];

//recorremos el objeto y vamos armando el script
excelData.forEach((row) => {
  //se pasa el descuento de decimal a porcentaje y se escribe como string
  const desc = row.desc * 100;
  const cod =  row.cod;
  let strCod = cod.toString();
  //se pasa el código a 7 cifras y se rellena con ceros a la izquierda
  let paddedCod = strCod.padStart(7, "0");
  //detecta si es la primera sucursal o si hay cambios
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
  //calcula la fecha hasta pasando el formato de excel a ddmmyyyy
  const fechaHasta = xlsx.SSF.format("ddmmyyyy", row.Hasta);   
  //escribimos el script
  ahkScript += `SendText "${paddedCod}"\nSend "{Tab Down}"\nSend "{Tab Up}"\nSendText "${fechaDesde}"\nSend "{Tab Down}"\nSend "{Tab Up}"\nSend "{Tab Down}"\nSend "{Tab Up}"\nSendText "${fechaHasta}"\nSend "{Enter Down}"\nSend "{Enter Up}"\nSleep 500\nSendText "${desc}"\nSend "{Enter Down}"\nSend "{Enter Up}"\nSleep 500\n`;
});

//final del script
ahkScript += `Return\n}\n^q::ExitApp`;

//se escribe el string en el archivo del script
fs.writeFileSync("datos.ahk", ahkScript);
//se escribe el string en un json (para verificar que se haya generado correctamente)
fs.writeFileSync("datos.json", JSON.stringify(excelData));

//avisamos que se generó bien el archivo
console.log("Archivo generado en: datos.ahk");
//mostramos una tabla con los comandos del ahk
console.table(sucCounter);
console.log("Para ejecutar el script presiona Ctrl + 1, Ctrl + 2, Ctrl + 3, etc.");
console.log('Para cerrar el script presiona Ctrl + q');

//ejecutamos el script
cp.exec("datos.ahk", (err, stdout, stderr) => {
  if (err) {
    console.error(err);
    return;
  }
  stdout ? console.log(`stdout: ${stdout}`) : console.log("Archivo ejecutado correctamente");
  stderr ? console.error(`stderr: ${stderr}`) : null;
});