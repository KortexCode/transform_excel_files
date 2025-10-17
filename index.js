const fs = require("fs-extra");
const path = require("path") ;
const XLSX = require("xlsx");

// 🔧 Ruta de origen (donde están tus archivos actuales)
const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/fichas_evaluacion_automate/Fichas_Descargadasv2";

// 📁 Ruta de destino (donde guardarás los nuevos .xlsx)
const rutaDestino = "/mnt/c/Users/KortexCode/Servicio Nacional de Aprendizaje/Analítica 9521 - Documentos/General/JUICIOS EVALUATIVOS";

try {
  // Crear la carpeta de destino si no existe
  fs.ensureDirSync(rutaDestino);
  
  // 🔍 Leer todos los archivos en la carpeta origen
  const archivos = fs.readdirSync(rutaOrigen);
  console.log("archivos leidos", archivos)
  // 📦 Procesar cada archivo
  archivos.forEach((archivo) => {
    const extension = path.extname(archivo).toLowerCase();
    console.log("extensión", extension)
    const nombreBase = path.basename(archivo, extension);
    console.log("nombre base", nombreBase)
    // 📄 Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls", ".xlsx", ".csv"].includes(extension)) {
      const rutaArchivo = path.join(rutaOrigen, archivo);
      const libro = XLSX.readFile(rutaArchivo);
      console.log("XLS lee", rutaArchivo)
  
      // 📤 Nueva ruta con extensión .xlsx
      const nuevoArchivo = path.join(rutaDestino, `${nombreBase}.xlsx`);
  
      // 💾 Guardar el archivo convertido
      XLSX.writeFile(libro, nuevoArchivo);
  
      /* console.log(`✅ Convertido: ${archivo} → ${nombreBase}.xlsx`); */
    } else {
      /* console.log(`⏭️ Omitido: ${archivo} (no es un archivo Excel)`); */
    }});
  
} catch (error) {
  console.log("Error inesperado!!")
  console.log(error)
}


console.log("🎉 Conversión completada.");


 /*
      Aquí se construye la ruta del nuevo archivo convertido, esta vez con la extensión `.xlsx`.

      `path.join(rutaDestino, ...)` asegura que el nuevo archivo se guarde dentro de la carpeta de destino.

      `${nombreBase}.xlsx` genera el nuevo nombre manteniendo la parte base original,
      evitando sobreescribir archivos y manteniendo coherencia entre nombres de entrada y salida.
    */
    

    /*
      `XLSX.writeFile(libro, nuevoArchivo)` escribe el objeto del libro (`libro`) en un archivo físico en disco.
      
      - Convierte las estructuras en memoria a formato Excel OpenXML (.xlsx).
      - Crea el archivo si no existe o lo reemplaza si ya existía.
      - Usa las APIs nativas de Node.js para escribir los bytes de salida en el sistema de archivos.

      Este proceso es sincrónico y bloquea la ejecución hasta que la escritura finaliza correctamente.
    */

   
