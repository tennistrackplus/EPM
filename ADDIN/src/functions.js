/**
 * Funciones personalizadas de Draco: DRACO.EPM_VALUE y DRACO.EPM_FORMULA.
 * ---------------------------------------------------------------------
 * Las mismas que el XLAM (EPM_VALUE / EPM_FORMULA en VBA) para donde no
 * hay XLAM: Excel online y escritorio sin él. Office exige un espacio de
 * nombres (DRACO.), así que el add-in escribe un nombre u otro según el
 * equipo y adapta las fórmulas del libro al abrirlo (ver ensureDracoFnMode
 * en commands.js). Se cargan en el runtime compartido (taskpane.html).
 *
 * Igual que en VBA, solo devuelven el texto a mostrar: el add-in lee sus
 * argumentos de la fórmula, no del resultado.
 */
(function () {
    "use strict";

    function isErr(v) { return v instanceof Error || (v && typeof v === "object" && v.type === "Error"); }
    function txt(v) {
        if (v === null || v === undefined || isErr(v)) return "";
        if (Array.isArray(v)) return txt(v[0] && v[0][0]);
        return String(v);
    }

    /** EPM_VALUE(Dimension, Atributo, Valor, Display) -> Display */
    function epmValue(dimension, atributo, valor, display) {
        return txt(display);
    }

    /** EPM_FORMULA(Texto, [Formula], [Formato], [Barras]) -> Texto */
    function epmFormula(texto, formula, formato, barras) {
        return txt(texto);
    }

    if (typeof CustomFunctions !== "undefined" && CustomFunctions.associate) {
        CustomFunctions.associate("EPM_VALUE", epmValue);
        CustomFunctions.associate("EPM_FORMULA", epmFormula);
    }
})();
