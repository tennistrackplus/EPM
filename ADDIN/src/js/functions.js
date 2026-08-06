/**
 * Devuelve el texto indicado en el parámetro Display.
 * @customfunction EPM_VALUE
 * @param {string} dimension Nombre de la dimensión.
 * @param {string} atributo Nombre del atributo.
 * @param {any} valor Valor asignado.
 * @param {string} display Texto a mostrar en la celda.
 * @returns {string} El texto proporcionado en display.
 */
function EPM_VALUE(dimension, atributo, valor, display) {
    return display;
}

// Registro explícito de la función en Office.js
CustomFunctions.associate("EPM_VALUE", EPM_VALUE);