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

/**
 * Línea calculada de un eje Estático (o del Estático de un Híbrido): la
 * celda muestra el texto; al refrescar, Draco copia la fórmula del 2º
 * parámetro a cada celda de valores de la línea, igual que copiar y pegar
 * esta celda (referencias relativas), con el formato y las barras opcionales.
 * @customfunction EPM_FORMULA
 * @param {string} texto Texto a mostrar en la cabecera.
 * @param {any} [formula] Fórmula a copiar (p. ej. A6-A7).
 * @param {string} [formato] Formato de número opcional (p. ej. 0.0%).
 * @param {string} [barras] Barras opcionales (p. ej. barra|invertir|izq=40|dcha=25).
 * @returns {string} El texto.
 */
function EPM_FORMULA(texto, formula, formato, barras) {
    return texto;
}

// Registro explícito de la función en Office.js
CustomFunctions.associate("EPM_VALUE", EPM_VALUE);
CustomFunctions.associate("EPM_FORMULA", EPM_FORMULA);