const xlsx = require("xlsx");

/**
 * Procesa un archivo Excel y convierte sus datos al formato JSON.
 * @param {string} filePath - Ruta al archivo Excel.
 * @param {Object} options - Opciones adicionales (como seleccionar una hoja específica).
 * @returns {Array<Object>} - Datos extraídos en formato JSON.
 */
function procesarExcel(filePath, options = {}) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = options.sheetName || workbook.SheetNames[0]; // Por defecto, la primera hoja
    const worksheet = workbook.Sheets[sheetName];

    // Convertir los datos a JSON sin transformar las columnas
    return xlsx.utils.sheet_to_json(worksheet, { defval: null });
  } catch (error) {
    throw new Error(`Error al procesar el archivo Excel: ${error.message}`);
  }
}

module.exports = { procesarExcel };
