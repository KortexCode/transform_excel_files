/*Este módulo filtrará las competencias y las horas asociadas a ella por cada ficha que se
se le envíe*/
/* const fs = require("fs-extra");
const XLSX = require("xlsx");
const path = require("path") ; */

import fs from "fs-extra";
import path from "path" ;
import XLSX from "xlsx";

export default function leerArchivoAFiltrar(numFicha){
  const rutaOrigenAfiltrar = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas/hojaParaFiltrar";
  // 1. Leer todos los archivos en la carpeta origen
  const archivos = fs.readdirSync(rutaOrigenAfiltrar);
  const rutaArchivo = path.join(rutaOrigenAfiltrar, archivos[0]);
  const libro = XLSX.readFile(rutaArchivo);
  const hoja = libro.Sheets[libro.SheetNames[0]];
  
  // 2. Convertir la hoja a JSON para manipular los datos
  const datos = XLSX.utils.sheet_to_json(hoja);
  const ficha = datos.filter(datos => datos['Codigo Ficha'] == numFicha);
  const programListDuplicate = ficha.map(datos => datos['Competencia']);
  const programList = [...new Set(programListDuplicate)]
  const horasProgramadas = programList.map(competencia => {
    let horas = 0;
    ficha.forEach(ficha => {
      if(ficha['Competencia'] === competencia) {
        horas += ficha['Horas Programadas']
      }
    });
    return {
      competencia,
      horas
    }
  })
  return horasProgramadas;
}

