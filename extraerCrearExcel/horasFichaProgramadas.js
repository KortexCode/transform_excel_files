import fs from "fs-extra";
import path from "path";
import Exceljs from "exceljs";
import { programsMilson } from "./tokenListMilson.js";
import { programsJuan } from "./tokenListJuan.js";

const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas";
const rutaNueva = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas/Resultados";


function leerArchivoDeExcel(rutaArchivo, rutaNuevo) {
  const {tokenList, programName} = programsJuan.ejecucionDeProgramDeportivo2;
  console.log('\n')
  console.log("Comenzando.....")
  console.log('\n')

  tokenList.forEach(async (token) => {
    const workBook = new Exceljs.Workbook();
    await workBook.xlsx.readFile(rutaArchivo);
    const sheet = workBook.getWorksheet(workBook.id);
    const rutaNuevoArchivo = path.join(rutaNuevo, `${token}.xlsx`);
    console.log('Creando..', rutaNuevoArchivo)
    //Verificar si el archivo ya existe en la ruta de destino
    if (fs.existsSync(rutaNuevoArchivo)) {
      console.log(`⚠️ El archivo ${token}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`);
      console.log("-----------------")
      return; // Saltar a la siguiente iteración del bucle
    }
    //Agregar una hoja (copia de la original) con el nombre de la ficha del ciclo actual.
    const newSheet = workBook.addWorksheet('ficha');
    const columnsList = ['C', 'D', 'E', 'F'];//Array de columnas donde están los datos a extraer.
    console.log("Creando Nuevo libro..📖")
    columnsList.forEach(col => {
      let filaInicio = 2 //Fila donde se comienza a ingresar las competencias.
      let iniciarDatos = false; //Fila donde se establece condición para ingresar datos.
      if(col === 'D' || col === 'C') {
        sheet.getColumn(col).eachCell((cell) => {
            if(cell.value == null) return;
            const cellValueToText = cell.text.trim().replace(/\s+/g, ' ');
            /* console.log(cellValueToText, cellValueToText == 'COMPETENCIA', 'COMPETENCIA') */
            if(cellValueToText == 'COMPETENCIA') {
                newSheet.getCell('A1').value = "NOMBRE PROGRAMA";
                newSheet.getCell('B1').value = 'FICHA';
                newSheet.getCell('C1').value = cell.value;
                iniciarDatos = true;
                return;
            }
            if(iniciarDatos){
                newSheet.getCell(`A${filaInicio}`).value = programName;
                newSheet.getCell(`B${filaInicio}`).value = token;
                newSheet.getCell(`C${filaInicio}`).value = cell.value;
                filaInicio++;
            }
        });
      }   
        
      if(col === 'F' || col === 'E') {
        sheet.getColumn(col).eachCell((cell, rowNumber) => {
            const competenciaD = sheet.getCell(`D${rowNumber}`).value; //Habilitar línea si son de Nestor
            if(!competenciaD) return; //Habilitar línea si son de Nestor
            if(cell.value == null) { 
              if(iniciarDatos) filaInicio++;
              return;
            };
            const cellValueToText = cell.text.trim().replace(/\s+/g, ' ');
            if(cellValueToText == 'HORAS DE DISEÑO') {
              newSheet.getCell('D1').value = 'HORAS DISEÑO';
              iniciarDatos = true;
              return;
            }
            if(iniciarDatos){
              newSheet.getCell(`D${filaInicio}`).value = cell.value;
              filaInicio++;
            }
        });
      }
    });
    workBook.removeWorksheet(sheet.name);
    await workBook.xlsx.writeFile(rutaNuevoArchivo);
    console.log(`Libro de Excel ${token}.xlsx creado 😁`)
  });
  
/*   newHoja.eachRow((row, index) => {
    console.log(row.getCell('F').value, index)
  });
     */
}


try {
  // 1. Leer todos los archivos a manipular en la carpeta origen 
  const archivos = fs.readdirSync(rutaOrigen);
  
  archivos.forEach((archivo) => {
    //2. Se extrae el nombre base y extensión de archivo
    const extension = path.extname(archivo).toLowerCase();
    const nombreBase = path.basename(archivo, extension);
    
    console.log("nombre base: ", nombreBase + " " + extension);

    // 5. Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls"].includes(extension) || [".xlsx"].includes(extension)) {
      //.6 Se genera la ruta donde está el archivo que se va a manipular
      const rutaArchivo = path.join(rutaOrigen, archivo);
      console.log("RUTA ARCHIVO", rutaArchivo) 
      //7. Leer archivo, copiar valores y crear excel resultante
      leerArchivoDeExcel(rutaArchivo, rutaNueva);
      console.log("🎉 Creación de archivo Completado.");
    } else {
      console.log(`⬅️ Omitido: ${archivo} (no es un archivo Excel con extensión requerida)`);
      console.log("-----------------")
    }   
  });
} catch (error) {
  console.log("Error inesperado!!")
  console.log(error)
}
console.log("🏅 Proceso Termiando.");