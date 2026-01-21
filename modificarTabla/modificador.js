const fs = require('fs-extra');
const path = require( 'path');
const excelJs = require('exceljs');
//Ruta origen de los archivos
const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas/archivos_terminados2";
const rutaArchivoGuardado = "/mnt/c/Users/KortexCode/Downloads/Documentos de trabajo/programacion_horas/archivos_terminados2/resultado";
async function modificarTablaExcel(rutaArchivo, rutaArchivoFinal) {
    //Crear un nuevo libro limpio
    const workBook = new excelJs.Workbook();
    //Leer el libro nuevo y copiar los datos del libro origen(objetivo)
    await workBook.xlsx.readFile(rutaArchivo);
    //Extraer la hoja
    const sheet = workBook.getWorksheet(workBook.id);
    const sheetName = sheet.name;
    console.log("Nombre Hoja: ", sheetName)
    const newSheet = workBook.addWorksheet('copia');
    const columnsList = ['C', 'D', 'E', 'F'];
    columnsList.forEach(col => {
        let filaInicio = 2 //Fila donde se comienza a ingresar datos
        let iniciarDatos = false; //Fila donde se establece condición para ingresar datos
        if(col === 'D' || col === 'C') {
            sheet.getColumn(col).eachCell((cell) => {
                if(cell.value == null) return;
                const cellValueToText = cell.text.trim().replace(/\s+/g, ' ');
                /* console.log(cellValueToText, cellValueToText == 'COMPETENCIA', 'COMPETENCIA') */
                if(cellValueToText == 'COMPETENCIA') {
                    newSheet.getCell('A1').value = cell.value;
                    iniciarDatos = true;
                    return;
                }
                if(iniciarDatos){
                    if(cell.value == null) return;
                    newSheet.getCell(`A${filaInicio}`).value = cell.value;
                    filaInicio++;
                }
            });
        }   
        
        if(col === 'F' || col === 'E') {
            sheet.getColumn(col).eachCell((cell) => {
                if(cell.value == null) return;
                const cellValueToText = cell.text.trim().replace(/\s+/g, ' ');
                if(cellValueToText == 'HORAS DE DISEÑO') {
                    newSheet.getCell('B1').value = cell.value;
                    iniciarDatos = true;
                    return;
                }
                if(iniciarDatos){
                    if(cell.value == null) return;
                    newSheet.getCell(`B${filaInicio}`).value = cell.value;
                    filaInicio++;
                }
            });
        }
    });
    /* workBook.removeWorksheet(sheet.id); */
    console.log("Se crea libro :", sheetName);
    await workBook.xlsx.writeFile(rutaArchivoFinal) 
}    

try {
  // 1. Leer todos los archivos a manipular en la carpeta origen 
  const archivos = fs.readdirSync(rutaOrigen);
  //2. Por cada archvio se realizará una modificación
  archivos.forEach((archivo) => {
    //3.. Se extrae el nombre base y extensión del archivo actual
    const extension = path.extname(archivo).toLowerCase();
    const nombreBase = path.basename(archivo, extension);
 
    // 4. Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls"].includes(extension) || [".xlsx"].includes(extension)) {
        //.5 Se genera la ruta donde está el archivo que se va a manipular
        const rutaArchivo = rutaOrigen + '/' + archivo;
        console.log("Ruta archivo: ", rutaArchivo)
        const rutaArchivoFinal = rutaArchivoGuardado + '/' + archivo;
        // 7. Verificar si el archivo ya existe en la ruta de destino
        if (fs.existsSync(rutaArchivoFinal)) {
            console.log(`⚠️ El archivo ${nombreBase}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`);
            console.log("-----------------")
            return; 
        }
        //modificar tabla
        modificarTablaExcel(rutaArchivo, rutaArchivoFinal);
    } else {
      console.log(`⬅️ Omitido: ${archivo} (no es un archivo Excel con extensión requerida)`);
      console.log("-----------------")
    }   
  });
} catch (error) {
  console.log("Error inesperado!!");
  console.log(error);
}
console.log("🏅 Proceso Termiando.");