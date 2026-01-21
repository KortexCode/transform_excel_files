const fs = require("fs-extra");
const path = require("path") ;
const Exceljs = require("exceljs");
const leerArchivoAFiltrar = require("./filtrarTablaPowerBi");

const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas";
const rutaNueva = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas/Resultados";

const tokenList = [
  '3293258',
 ];

 function copiarFormato(origen, destino) {
  // 1. Valor
  /* destino.value = origen.result ?? origen.value; */
  destino.value = origen.value;

  // 2. Estilo
  /* destino.style = JSON.parse(JSON.stringify(origen.style)); */

  // 3. Bordes
  if (origen.border) {
    destino.border = JSON.parse(JSON.stringify(origen.border));
  }

  // 4. Relleno
  if (origen.fill) {
    destino.fill = JSON.parse(JSON.stringify(origen.fill));
  }

  // 5. Fuente
  if (origen.font) {
    destino.font = JSON.parse(JSON.stringify(origen.font));
  }

  // 6. Alineación
  /* if (origen.alignment) {
    destino.alignment = JSON.parse(JSON.stringify(origen.alignment));
  } */

  // 7. Formato numérico
  if (origen.numFmt) {
    destino.numFmt = origen.numFmt;
  }
}

async function leerArchivoDeExcel(rutaArchivo, rutaNuevo) {
  const workBook = new Exceljs.Workbook();
  await workBook.xlsx.readFile(rutaArchivo);
  const hoja = workBook.getWorksheet(workBook.id);
  
  tokenList.forEach((token) => {
    //Extraer las competencias y sus horas en ejecutadas por cada token
    const competenciaHorasList = leerArchivoAFiltrar(token);
    //Agregar una hoja (copia de la original) con el nombre de la ficha del ciclo actual
    const newHoja = workBook.addWorksheet(token);
    //Recorrer la hoja original
    hoja.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const filaNueva = newHoja.getRow(rowNumber); //Obtener fila actual de la hoja nueva
      //Obtener celda actual en la fila de la hoja nueva
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        copiarFormato(cell, filaNueva.getCell(colNumber)); 
        /* filaNueva.getCell(colNumber).value = cell.value; */
      });
      //En contrar coincidencia entre las competencias y el valor de la celda actual
      //de la fila de la hoja origen
      const matchData = competenciaHorasList.find((compHora) => {
        const valorCompetencia = compHora.competencia.toLowerCase().replace(/\s+/g,' ').trim();
        const valor = row.getCell('D').value ?? "Soy Nulo XD";
        if(typeof valor == "number") {
          return valor == valorCompetencia 
        }
        const valorCelda = valor.toLowerCase().replace(/\s+/g,' ').trim();
        return valorCelda == valorCompetencia  
      });
      if(matchData){
        filaNueva.getCell('G').value = Math.ceil(matchData.horas);
        console.log("Valor", filaNueva.getCell('G').value) 
      }
      filaNueva.commit();
    });
  });
  
/*   newHoja.eachRow((row, index) => {
    console.log(row.getCell('F').value, index)
  });
     */
  await workBook.xlsx.writeFile(rutaNuevo)  
}


try {
  // 1. Leer todos los archivos a manipular en la carpeta origen 
  const archivos = fs.readdirSync(rutaOrigen);
  
  archivos.forEach((archivo) => {
    //2. Se extrae el nombre base y extensión de archivo
    const extension = path.extname(archivo).toLowerCase();
    const nombreBase = path.basename(archivo, extension);
    const rutaNuevoArchivo = path.join(rutaNueva, `${nombreBase}.xlsx`);
    console.log("nombre base: ", nombreBase + " " + extension);

    // 5. Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls"].includes(extension) || [".xlsx"].includes(extension)) {
      //.6 Se genera la ruta donde está el archivo que se va a manipular
      const rutaArchivo = path.join(rutaOrigen, archivo);
      console.log("RUTA ARCHIVO", rutaArchivo)
      // 7. Verificar si el archivo ya existe en la ruta de destino
      if (fs.existsSync(rutaNuevoArchivo)) {
        console.log(`⚠️ El archivo ${nombreBase}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`);
        console.log("-----------------")
        return; // Saltar a la siguiente iteración del bucle
      }
      //8. Leer archivo, copiar valores y crear excel resultante
      leerArchivoDeExcel(rutaArchivo, rutaNuevoArchivo)
      .catch(e => console.log("Error durant lectura de archivo: ", e));
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