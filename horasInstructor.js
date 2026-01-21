const fs = require("fs-extra");
const path = require("path");
const XLSX = require("xlsx");

// Ruta de origen (donde están tus archivos actuales)
/* const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/fichas_evaluacion_automate/Horas descargadas"; */
const rutaOrigen =
  "/mnt/c/Users/KortexCode/Documents/servicio_nacional_sena/9521_2026/";

try {
  // Crear la carpeta de destino si no existe
  /* fs.ensureDirSync(rutaOrigen); */

  // Leer todos los archivos en la carpeta origen
  const archivos = fs.readdirSync(rutaOrigen);
  console.log("archivos leidos", archivos);

  // Procesar cada archivo
  archivos.forEach((archivo) => {
    const extension = path.extname(archivo).toLowerCase();
    console.log("extensión", extension);
    const nombreBase = path.basename(archivo, extension);
    console.log("nombre base", nombreBase);
    let nombreArchivo = "";

    const rutaArchivo = path.join(rutaOrigen, archivo);
    //Verificar si el archivo ya existe en la ruta de destino
    
    // Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls"].includes(extension)) {
      nombreArchivo = archivo.includes("r.xls") ? "ENERO" : archivo.includes("(1).xls") ? "FEBRERO" :
      archivo.includes("(2).xls") ? "MARZO" : archivo.includes("(3).xls") ? "ABRIL" :
      archivo.includes("(4).xls") ? "MAYO" : archivo.includes("(5).xls") ? "JUNIO" :
      archivo.includes("(6).xls") ? "JULIO" : archivo.includes("(7).xls") ? "AGOSTO" :
      archivo.includes("(8).xls") ? "SEPTIEMBRE" : archivo.includes("(9).xls") ? "OCTUBRE" :
      archivo.includes("(10).xls") ? "NOVIEMBRE" : archivo.includes("(11).xls") ? "DICIEMBRE" : false;
      if(!nombreArchivo){
        console.log("No se encontraron nombres de archivos válidos")
        return;
      }

      //Se define la nueva ruta para guardar el archivo convertido
      const rutaNuevaArchivo = path.join(rutaOrigen, `${nombreArchivo}.xlsx`);
      if (fs.existsSync(rutaNuevaArchivo)) {
        console.log(
          `⚠️ El archivo ${nombreArchivo}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`
        );
        console.log("-----------------");
        return; 
      }

      //Se lee el archivo xls
      const libro = XLSX.readFile(rutaArchivo);

      // Guardar el archivo convertido a XLSX
      XLSX.writeFile(libro, rutaNuevaArchivo);
      console.log(`✅ Convertido: ${archivo} → ${nombreArchivo}.xlsx`);

      //Elimanar el archivo anterior con extesión xls
      if (archivo.endsWith(".xls")) {
        fs.unlinkSync(rutaArchivo); 
        console.log(`✖️ Eliminado: ${archivo}`);
      }
      console.log("-----------------");

    } else {
      console.log(
        `⬅️ Omitido: ${archivo} (no es un archivo Excel con extensión xls)`
      );
      console.log("-----------------");
    }
  });
} catch (error) {
  console.log("Error inesperado!!");
  console.log(error);
}

console.log("🎉 Conversión completada.");
