const fs = require("fs-extra");
const path = require("path") ;
const XLSX = require("xlsx");

const cedulas = require("./datosFiltroTabla.js")

const rutaOrigen = "/mnt/c/Users/KortexCode/Downloads/temp_tarea";
const rutaNueva = "/mnt/c/Users/KortexCode/Downloads/temp_tarea/pago_instructores"

// Ruta de destino (donde guardarás los nuevos .xlsx)
//const rutaAlternativa = "/mnt/c/Users/KortexCode/Servicio Nacional de Aprendizaje/Analítica 9521 - Documentos/9521";
/* const rutaOrigen = "/mnt/c/Users/KortexCode/Servicio Nacional de Aprendizaje/Analítica 9521 - Documentos/General/JUICIOS EVALUATIVOS";
 */

try {
  // Crear la carpeta de destino si no existe
  /* fs.ensureDirSync(rutaOrigen); */
  
  // 1. Leer todos los archivos en la carpeta origen
  const archivos = fs.readdirSync(rutaOrigen);
  /*console.log("archivos leidos", archivos)*/

  // 1.1 Procesar cada archivo
  archivos.forEach((archivo) => {
    const extension = path.extname(archivo).toLowerCase();
    const nombreBase = path.basename(archivo, extension);
    console.log("nombre base: ", nombreBase + " " + extension)

    // 1.2 Leer solo archivos de Excel (.xls, .csv, .xlsx)
    if ([".xls"].includes(extension) || [".xlsx"].includes(extension)) {
      const rutaArchivo = path.join(rutaOrigen, archivo);
      const libro = XLSX.readFile(rutaArchivo);

      const hoja = libro.Sheets[libro.SheetNames[0]];

      // 2. Convertir la hoja a JSON para manipular los datos
      const datos = XLSX.utils.sheet_to_json(hoja);
      // 3. Filtrar los datos por las columnas 
      console.log(cedulas)
      cedulas.forEach((ced) => {
        archivoFiltrado = datos.filter((row) => row["Identificacion"] == ced &&
          row["Tipo Doc Soporte Compromiso"] == "CONTRATO DE PRESTACION DE SERVICIOS - PROFESIONALES")

        if(archivoFiltrado.length > 0) {
          console.log("Se obtuvieron coincidencias ✔️") 
        }
        else {
          console.log(`No hay concidencias con la Cédula: ${ced}`) 
          return
        }
      
        // 4. Crear una nueva hoja con los datos filtrados
        const nombreInstructor = archivoFiltrado[0]["Nombre Razon Social"]
        const nuevaHoja = XLSX.utils.json_to_sheet(archivoFiltrado);
        
        console.log("nombre: ", nombreInstructor)
        
        // 5. Crear un nuevo libro
        const nuevoWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(nuevoWorkbook, nuevaHoja, nombreInstructor.slice(0, 15));
        
        // 6. Nueva ruta con extensión .xlsx
        const rutaNuevoArchivo = path.join(rutaNueva,`${nombreInstructor}.xlsx`);
        console.log("Ruta archivo nuevo: ", rutaNuevoArchivo);
        // 7. Verificar si el archivo ya existe en la ruta de destino
        if (fs.existsSync(rutaNuevoArchivo)) {
          console.log(`⚠️ El archivo ${nombreInstructor}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`);
          console.log("-----------------")
          return; // Saltar a la siguiente iteración del bucle
        }

        // Guardar el archivo convertido a XLSX
        XLSX.writeFile(nuevoWorkbook, rutaNuevoArchivo);
        console.log(`✅ Convertido: ${archivo} → ${nombreInstructor}.xlsx`);

        console.log("-----------------")
      })
           
    } else {
      console.log(`⬅️ Omitido: ${archivo} (no es un archivo Excel con extensión requerida)`);
      console.log("-----------------")
    }
  });

  console.log("🎉 Conversión Completada.");
  
} catch (error) {
  console.log("Error inesperado!!")
  console.log(error)
}
console.log("🏅 Proceso Termiando.");