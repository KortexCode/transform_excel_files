import fs from "fs-extra";
import path from "path";
import XLSX from "xlsx";

export default function transformReports(rutaOrigen, reporte) {
  const tipoReporte = {
    juicio: "Juicios",
    instructorFicha: "instructor_Ficha"
  }
  let fichaCaracterizacion = "";
  const existedFile = [];

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
      if ([".xls", ".XLS"].includes(extension)) {
        const rutaArchivo = path.join(rutaOrigen, archivo);
        const libro = XLSX.readFile(rutaArchivo);
      
        const hoja = libro.Sheets[libro.SheetNames[0]];
        if(reporte == tipoReporte.juicio) {
          fichaCaracterizacion = hoja['C3'].v;
        }
        else if(reporte == tipoReporte.instructorFicha) {
          fichaCaracterizacion = hoja['B2'].v;
        }
        else {
          throw Error("No se definió un tipo de reporte Válido😪");
        }
        console.log("Ficha de caracterización: ", fichaCaracterizacion)
      
        // Nueva ruta con extensión .xlsx
        const rutaNuevoArchivo = path.join(rutaOrigen, `${fichaCaracterizacion}.xlsx`);
         console.log("Ruta archivo nuevo: ", rutaNuevoArchivo);
      
        //Verificar si el archivo ya existe en la ruta de destino
        if (fs.existsSync(rutaNuevoArchivo)) {
          console.log(`⚠️ El archivo ${fichaCaracterizacion}.xlsx ya existe en la ruta de destino. Se omitirá la conversión.`);
          console.log("-----------------")
          existedFile.push(`El archivo ${fichaCaracterizacion}.xlsx ya existe en la ruta de destino.`)
          return; // Saltar a la siguiente iteración del bucle
        }
      
        // Guardar el archivo convertido a XLSX
        XLSX.writeFile(libro, rutaNuevoArchivo);
        console.log(`✅ Convertido: ${archivo} → ${fichaCaracterizacion}.xlsx`);
      
        //Elimanar el archivo anterior con extesión xls
        if (archivo.endsWith(".xls") || archivo.endsWith(".XLS")) {
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

  console.log("NOVEDADES:")
  if(existedFile.length){
    existedFile.forEach(msg => console.log(msg))
  }else {
    console.log("Sin novedades...🙂")
  }

  console.log("🏅 Proceso Termiando.");
}



   
