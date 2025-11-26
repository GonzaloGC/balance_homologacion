// Importar los módulos necesarios
const fs = require('fs/promises'); // Usamos la versión de promesas para async/await
const path = require('path');
const xlsx = require('xlsx');

// --- Configuración ---
// Define la ruta a la carpeta principal que quieres escanear.
// path.join asegura que la ruta funcione en diferentes sistemas operativos.
// __dirname es una variable global de Node.js que contiene la ruta del directorio donde se ejecuta el script actual.
const parentFolderPath = '/mnt/c/Users/gg/Documents/Ingeaudit Info/HOMOLOGACION ADMINISTRATIVA/2-Balances Automaticos/archivo balance automatico/10 al 16 de noviembre del 2025';
// const parentFolderPath = 'C:\Users\gg\Documents\Ingeaudit Info\HOMOLOGACION DOCUMENTAL\test'; /* Se ocupa esta ruta en windows */
// const parentFolderPath = path.join(__dirname, '7_al_13_de_Abril_de_2025'); /* Se ocupa esta ruta en windows */
const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
const filename = `reportes_balance_${timestamp}.xlsx`
// Define el nombre y la ruta del archivo Excel de salida.
const outputExcelPath = path.join('/mnt/c/Users/gg/Documents/Ingeaudit Info/HOMOLOGACION ADMINISTRATIVA/2-Balances Automaticos/reportes_balance/', filename);
// const outputExcelPath = '/mnt/c/Users/gg/Documents/Ingeaudit Info/HOMOLOGACION ADMINISTRATIVA/test/reportes_balance/reportes_balance.xlsx';
// const outputExcelPath = path.join(__dirname, 'reportes_balance.xlsx');
// --- Fin Configuración ---

// Función principal asíncrona para poder usar await
async function generarReporteExcel() {
    console.log(`Iniciando escaneo de la carpeta: ${parentFolderPath}`);

    const dataForExcel = []; // Array para almacenar los datos de cada fila del Excel
    const headers = ["Nombre", "Fecha de Registro", "Cantidad", "Monto"]; // Definir encabezados explícitamente

    try {
        // 1. Leer el contenido de la carpeta principal
        const topLevelEntries = await fs.readdir(parentFolderPath, { withFileTypes: true });

        // Filtrar para obtener solo las carpetas (los "Nombres")
        const topLevelFolders = topLevelEntries.filter(dirent => dirent.isDirectory());
        console.log(`Se encontraron ${topLevelFolders.length} carpetas principales.`);

        // 2. Recorrer cada carpeta principal ("Nombre")
        for (const topLevelDirent of topLevelFolders) {
            const topLevelFolderName = topLevelDirent.name;
            const topLevelFolderPath = path.join(parentFolderPath, topLevelFolderName);
            console.log(`  Procesando carpeta principal: ${topLevelFolderName}`);

            try {
                // 3. Leer el contenido de la carpeta principal actual
                const subEntries = await fs.readdir(topLevelFolderPath, { withFileTypes: true });

                // Filtrar para obtener solo las subcarpetas (las "Fechas de Registro")
                const subFolders = subEntries.filter(dirent => dirent.isDirectory());

                if (subFolders.length === 0) {
                    console.log(`    -> No se encontraron subcarpetas (Fecha de Registro) en ${topLevelFolderName}.`);
                    // Opcional: podrías agregar una fila con Cantidad 0 si es necesario,
                    // pero el requerimiento pide contar archivos *dentro* de las subcarpetas.
                    // Si no hay subcarpetas, no hay nada que contar según la lógica pedida.
                }

                // 4. Recorrer cada subcarpeta ("Fecha de Registro")
                for (const subFolderDirent of subFolders) {
                    const subFolderName = subFolderDirent.name; // Este es el valor para "Fecha de Registro"
                    const subFolderPath = path.join(topLevelFolderPath, subFolderName);
                    let fileCount = 0; // Inicializar contador de archivos

                    try {
                        // 5. Leer el contenido de la subcarpeta actual
                        const fileEntries = await fs.readdir(subFolderPath, { withFileTypes: true });

                        // 6. Filtrar y contar solo los archivos (ignorar otras subcarpetas dentro)
                        const files = fileEntries.filter(dirent => dirent.isFile());
                        fileCount = files.length; // La cantidad de archivos encontrados

                        console.log(`      -> Subcarpeta '${subFolderName}' encontrada con ${fileCount} archivos.`);
                        // Asignar una fórmula a la celda C2
                        // const formula = worksheet['D2'] = { f: 'C2*750' };

                        // 7. Agregar la información al array de datos para el Excel
                        dataForExcel.push({
                            // "Nombre": topLevelFolderName,
                            // "Fecha de Registro": subFolderName, // Usando el nombre de columna solicitado
                            "Nombre": subFolderName,
                            "Fecha de Registro": topLevelFolderName, // Usando el nombre de columna solicitado
                            "Cantidad": fileCount,
                            // "monto": formula
                        });

                    } catch (fileReadError) {
                        // Manejar errores si no se puede leer el contenido de una subcarpeta (ej. permisos)
                        console.error(`      -> Error al leer contenido de la subcarpeta ${subFolderPath}:`, fileReadError.message);
                        // Opcionalmente, agregar una fila indicando el error
                        dataForExcel.push({
                            "Nombre": topLevelFolderName,
                            "Fecha de Registro": subFolderName,
                            "Cantidad": 'Error al leer' // O podrías poner 0 o -1
                        });
                    }
                } // Fin del bucle de subcarpetas

            } catch (subReadError) {
                // Manejar errores si no se puede leer el contenido de una carpeta principal (ej. permisos)
                console.error(`  Error al leer contenido de la carpeta principal ${topLevelFolderPath}:`, subReadError.message);
                // Podrías decidir saltar esta carpeta principal o agregar una fila de error
            }
        } // Fin del bucle de carpetas principales

        // 8. Verificar si se recopilaron datos
        if (dataForExcel.length === 0) {
            console.log("\nNo se encontraron datos válidos para generar el Excel. Asegúrate de que la estructura de carpetas exista y sea correcta.");
            return; // Salir si no hay nada que escribir
        }

        // 9. Crear el archivo Excel
        console.log("\nGenerando archivo Excel...");

        // Crear un nuevo libro de trabajo (Workbook)
        const workbook = xlsx.utils.book_new();

        // Convertir el array de objetos JSON a una hoja de cálculo (Worksheet)
        // Usamos la opción 'header' para asegurarnos de que las columnas estén en el orden deseado
        const worksheet = xlsx.utils.json_to_sheet(dataForExcel, { header: headers });

        // Añadir la hoja de cálculo al libro de trabajo con un nombre (ej. 'Reporte')
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Balance');

        // Escribir el libro de trabajo a un archivo .xlsx
        xlsx.writeFile(workbook, outputExcelPath);

        console.log(`¡Archivo Excel "${outputExcelPath}" creado exitosamente!`);

    } catch (error) {
        // Manejar errores generales (ej. la carpeta principal no existe, no hay permisos)
        if (error.code === 'ENOENT') {
            console.error(`\nError Crítico: La carpeta principal especificada no existe: ${parentFolderPath}`);
            console.error("Por favor, verifica la ruta en la sección de Configuración del script.");
        } else if (error.code === 'EACCES') {
            console.error(`\nError Crítico: Permiso denegado para leer la carpeta: ${parentFolderPath}`);
            console.error("Asegúrate de que el script tenga permisos de lectura para esta carpeta y sus subcarpetas.");
        } else {
            console.error("\nOcurrió un error inesperado durante el proceso:", error);
        }
    }
}

// --- Ejecutar la función principal ---
generarReporteExcel();