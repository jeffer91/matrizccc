// Pruebas automáticas sin dependencias externas ni archivos de estudiantes.
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

let writtenBook = null;
const xlsx = {
  utils: {
    book_new: () => ({ sheets: [] }),
    aoa_to_sheet: rows => ({ rows: structuredClone(rows) }),
    book_append_sheet: (book, sheet, name) => book.sheets.push({ sheet, name }),
    sheet_to_json: (sheet, options) => options.range
      ? [sheet.matrix[0]]
      : sheet.matrix,
    decode_range: () => ({ s: { r: 0, c: 0 }, e: { r: 99, c: 6 } })
  },
  read: () => testWorkbook,
  write: (book, options) => {
    writtenBook = { book, options };
    return new Uint8Array([80, 75, 3, 4]);
  }
};
const context = {
  window: { XLSX: xlsx, JSZip: {} },
  XLSX: xlsx,
  setTimeout,
  console,
  document: { addEventListener: () => {} }
};
vm.createContext(context);
const source = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8");
const code = source + "\n globalThis.testApi = { normalHeader, findType, cleanMatrix, cleanedName, outputBytes, readExcel };";
vm.runInContext(code, context, { filename: "app.js" });
const { normalHeader, findType, cleanMatrix, cleanedName, outputBytes, readExcel } = context.testApi;

function equalRows(actual, expected) {
  assert.deepEqual(JSON.parse(JSON.stringify(actual)), JSON.parse(JSON.stringify(expected)));
}
assert.equal(normalHeader("  descripcionComponente\u00a0\u200b"), "descripcionComponente");
assert.equal(findType(["codigoComponente","ordenComponente","descripcionComponente\u00a0","descripcionComponente2","descripcionComponente3"]), "base");
assert.equal(findType(["codigoComponente","ordenComponente","descripcionComponente"]), "unidades");
assert.equal(findType(["nivel","Mecanismo","tema","descripcion"]), "actividades");
assert.equal(findType(["columna desconocida"]), null);

const sourceRows = [
  ["nivel", "Mecanismo", "tema", "descripcion"],
  [2, "Taller Practico 1", "Álgebra básica", "Acentos, saltos\ny símbolos: % $"],
  [null, null, null, null],
  [2, "Actividad Contacto Docente 1", "Álgebra 2", "Segundo párrafo"],
  [null, null, null, null],
  [null, null, null, null]
];
const clean = cleanMatrix(sourceRows, "actividades");
assert.equal(clean.rows.length, 4, "Se quitan filas vacías finales");
equalRows(clean.rows[2], [null, null, null, null]);
assert.equal(clean.rows[1][3], "Acentos, saltos\ny símbolos: % $");
assert.equal(clean.width, 4);
assert.equal(clean.corrected.length, 0);

const dirtyHeaders = cleanMatrix([
  ["codigoComponente", "ordenComponente", "descripcionComponente\u00a0"],
  [3, 1, "Texto de Matemática de Redes"]
], "unidades");
assert.equal(dirtyHeaders.rows[0][2], "descripcionComponente");
assert.equal(dirtyHeaders.corrected.length, 1);

for (const type of ["base", "unidades", "actividades"]) {
  assert.match(cleanedName({type}), /^.+_SIN_FORMATO\.xlsx$/);
}
assert.equal(new Set(["base", "unidades", "actividades"].map(type => cleanedName({type}))).size, 3);

outputBytes({rows: clean.rows});
assert.equal(writtenBook.book.sheets.length, 1);
assert.equal(writtenBook.book.sheets[0].name, "Datos");
equalRows(writtenBook.book.sheets[0].sheet.rows, clean.rows);
assert.equal(writtenBook.options.bookType, "xlsx");

let testWorkbook = null;
async function testRead(type, headers) {
  const matrix = [headers, [...headers].map((_,i)=>i), [null,null,null,null,null]];
  testWorkbook = { SheetNames: ["Hoja1"], Sheets: { Hoja1: { "!ref": "A1:F3", matrix } } };
  const file = { name: "entrada.xlsx", size: 1000, arrayBuffer: async () => new ArrayBuffer(4) };
  const result = await readExcel(file, type);
  assert.equal(result.type, type);
  assert.equal(result.dataCount, 1);
  assert.ok(result.bytes.length > 0);
}
Promise.resolve()
  .then(() => testRead("base", ["codigoComponente", "ordenComponente", "descripcionComponente\u00a0", "descripcionComponente2", "descripcionComponente3"]))
  .then(() => testRead("unidades", ["codigoComponente", "ordenComponente", "descripcionComponente"]))
  .then(() => testRead("actividades", ["nivel", "Mecanismo", "tema", "descripcion"]))
  .then(() => console.log("OK: plantillas, limpieza de formato, datos, nombres y lectura de los 3 tipos"))
  .catch(error => { console.error(error); process.exitCode = 1; });
