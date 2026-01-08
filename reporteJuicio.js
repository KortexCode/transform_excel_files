const fs = require("fs-extra");
const path = require("path") ;
const XLSX = require("xlsx");

/* const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/temp_tarea/Reportes de juicio"; */

// Ruta de destino (donde guardarás los nuevos .xlsx)
//const rutaAlternativa = "/mnt/c/Users/KortexCode/Servicio Nacional de Aprendizaje/Analítica 9521 - Documentos/9521";
const rutaOrigen = "/mnt/c/Users/KortexCode/Servicio Nacional de Aprendizaje/Analítica 9521 - Documentos/General/JUICIOS EVALUATIVOS";

try {
  // Crear la carpeta de destino si no existe
  /* fs.ensureDirSync(rutaOrigen); */
  
  // Leer todos los archivos en la carpeta origen
  const archivos = fs.readdirSync(rutaOrigen);
  /*console.log("archivos leidos", archivos)*/

  // Procesar cada archivo
  archivos.forEach((archivo) => {
    const extension = path.extname(archivo).toLowerCase();
    const nombreBase = path.basename(archivo, extension);
    console.log("nombre base: ", nombreBase + " " + extension)

    // Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls"].includes(extension)) {
      const rutaArchivo = path.join(rutaOrigen, archivo);
      const libro = XLSX.readFile(rutaArchivo);

      const hoja = libro.Sheets[libro.SheetNames[0]];
      const fichaCaracterizacion = hoja['C3'].v;
      console.log("Ficha de caracterización: ", fichaCaracterizacion)
    
      // Nueva ruta con extensión .xlsx
      const rutaNuevoArchivo = path.join(rutaOrigen, `${fichaCaracterizacion}.xlsx`);
       console.log("Ruta archivo nuevo: ", rutaNuevoArchivo);

      //Verificar si el archivo ya existe en la ruta de destino
      if (fs.existsSync(rutaNuevoArchivo)) {
        console.log(`⚠️ El archivo ${fichaCaracterizacion}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`);
        console.log("-----------------")
        return; // Saltar a la siguiente iteración del bucle
      }

      // Guardar el archivo convertido a XLSX
      XLSX.writeFile(libro, rutaNuevoArchivo);
      console.log(`✅ Convertido: ${archivo} → ${fichaCaracterizacion}.xlsx`);

      //Elimanar el archivo anterior con extesión xls
      if (archivo.endsWith(".xls")) {
        fs.unlinkSync(rutaArchivo); // elimina el archivo original .xls
        console.log(`✖️ Eliminado: ${archivo} → ${nombreBase}.xls`);
      }
      console.log("-----------------")
    } else {
      console.log(`⬅️ Omitido: ${archivo} (no es un archivo Excel con extensión xls)`);
      console.log("-----------------")
    }
  });

  console.log("🎉 Conversión Completada.");
  
} catch (error) {
  console.log("Error inesperado!!")
  console.log(error)
}
console.log("🏅 Proceso Termiando.");


   
