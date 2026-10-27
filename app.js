"use strict";

const SCHEMAS = {
  base: {
    title: "Redes base",
    headers: ["codigoComponente", "ordenComponente", "descripcionComponente", "descripcionComponente2", "descripcionComponente3"]
  },
  unidades: {
    title: "PEA · Unidades",
    headers: ["codigoComponente", "ordenComponente", "descripcionComponente"]
  },
  actividades: {
    title: "PEA · Actividades",
    headers: ["nivel", "Mecanismo", "tema", "descripcion"]
  }
};
const TYPES = ["base", "unidades", "actividades"];
const MAX_FILE_SIZE = 15 * 1024 * 1024;
const MAX_CELLS = 250000;
const loaded = Object.create(null);
let reading = false;
const $ = selector => document.querySelector(selector);

function normalHeader(value) {
  return String(value == null ? "" : value)
    .replace(/[\u00A0\u200B-\u200D\uFEFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function headerKey(value) {
  return normalHeader(value).toLowerCase();
}

function findType(headerRow) {
  const keys = headerRow.map(headerKey);
  const isMatch = type => SCHEMAS[type].headers.every((h, index) => keys[index] === headerKey(h));
  if (isMatch("base")) return "base";
  if (isMatch("actividades")) return "actividades";
  if (isMatch("unidades")) return "unidades";
  return null;
}

function showNotice(message, error = false) {
  const el = $("#notice");
  el.hidden = !message;
  el.classList.toggle("error", Boolean(error));
  el.textContent = message || "";
}

function shortText(value, max = 150) {
  const text = String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  return text.length > max ? text.slice(0, max - 1) + "…" : text;
}

function hasValue(cell) {
  return cell !== undefined && cell !== null && cell !== "";
}

function cleanMatrix(matrix, type) {
  const headers = SCHEMAS[type].headers;
  let lastRow = matrix.length - 1;
  while (lastRow > 0 && !(matrix[lastRow] || []).some(hasValue)) lastRow--;
  const cut = matrix.slice(0, lastRow + 1);
  const lastColumnWithContent = Math.max(0, ...cut.map(row => {
    for (let i = (row || []).length - 1; i >= 0; i--) {
      if (hasValue(row[i])) return i + 1;
    }
    return 0;
  }));
  const width = Math.max(headers.length, lastColumnWithContent);

  const rows = cut.map((row, rowIndex) => Array.from({ length: width }, (_, column) => {
    const cell = row && row[column];
    if (rowIndex === 0 && column < headers.length) return headers[column];
    return cell === "" || cell == null ? null : cell;
  }));
  const corrected = headers.reduce((changes, header, index) => {
    const original = (matrix[0] || [])[index];
    if (String(original) !== header) changes.push(header);
    return changes;
  }, []);
  const warnings = [];
  if (width > headers.length) {
    warnings.push("Existen " + (width - headers.length) + " columnas adicionales con contenido: se conservarán.");
  }
  return { rows, corrected, warnings, width };
}

function inspectWorkbook(book) {
  const candidates = [];
  for (const sheetName of book.SheetNames) {
    const ws = book.Sheets[sheetName];
    if (!ws || !ws["!ref"]) continue;
    // Solo la fila 1: no recorrer cientos o miles de filas para identificar la plantilla.
    const firstRow = { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } };
    const rawHeader = XLSX.utils.sheet_to_json(ws, {
      header: 1, raw: true, defval: null, blankrows: true, range: firstRow
    })[0] || [];
    const type = findType(rawHeader);
    if (type) candidates.push({ ws, sheetName, type });
  }
  return candidates[0] || null;
}

async function readExcel(file, intendedType) {
  if (!window.XLSX) throw new Error("No se pudo cargar el lector de Excel. Comprueba tu conexión a internet y actualiza la página.");
  if (!/\.xlsx?$/i.test(file.name)) throw new Error("El archivo «" + file.name + "» debe ser Excel (.xlsx o .xls).");
  if (file.size > MAX_FILE_SIZE) throw new Error("«" + file.name + "» supera el límite de 15 MB.");
  let book;
  try {
    // Evitar cargar estilos, textos de presentación y datos de formatos que se desecharán.
    book = XLSX.read(await file.arrayBuffer(), {
      type: "array", cellStyles: false, cellNF: false,
      cellText: false, cellDates: false, cellFormula: true
    });
  } catch (_) {
    throw new Error("No se pudo abrir «" + file.name + "». Comprueba que sea un archivo Excel válido.");
  }
  const candidate = inspectWorkbook(book);
  if (!candidate) throw new Error(
    "No se reconoció el formato de «" + file.name + "». Revisa que la primera fila tenga exactamente los encabezados institucionales."
  );
  if (intendedType && intendedType !== candidate.type) {
    throw new Error("«" + file.name + "» corresponde a " + SCHEMAS[candidate.type].title + ", no a " + SCHEMAS[intendedType].title + ". Usa su tarjeta correcta o la carga múltiple.");
  }
  const { ws, sheetName, type } = candidate;
  const bounds = XLSX.utils.decode_range(ws["!ref"]);
  const cells = (bounds.e.r - bounds.s.r + 1) * (bounds.e.c - bounds.s.c + 1);
  if (cells > MAX_CELLS) throw new Error("«" + file.name + "» tiene demasiadas celdas para procesarse de forma segura.");

  const matrix = XLSX.utils.sheet_to_json(ws, {
    header: 1, raw: true, defval: null, blankrows: true
  });
  const cleaned = cleanMatrix(matrix, type);
  const formulaCells = Object.values(ws).filter(cell => cell && typeof cell === "object" && cell.f).length;
  if (formulaCells) cleaned.warnings.push("Se detectaron " + formulaCells + " fórmulas; se exportarán sus valores guardados, no las fórmulas.");
  if (Array.isArray(ws["!merges"]) && ws["!merges"].length) {
    cleaned.warnings.push("Las celdas combinadas se dejarán sin combinar.");
  }

  // El XLSX limpio queda preparado al cargar; la descarga no vuelve a convertirlo.
  const bytes = outputBytes({ rows: cleaned.rows });
  return {
    type, sheetName, name: file.name, rows: cleaned.rows, bytes,
    corrected: cleaned.corrected, warnings: cleaned.warnings,
    dataCount: Math.max(0, cleaned.rows.length - 1), width: cleaned.width
  };
}

async function acceptFiles(fileList, intendedType = null) {
  const files = Array.from(fileList || []);
  if (!files.length || reading) return;
  reading = true;
  showNotice("Leyendo " + files.length + (files.length === 1 ? " archivo…" : " archivos…"));
  render();
  try {
    // Lectura concurrente: no obliga a cargar primero uno, luego otro, luego otro.
    const results = await Promise.allSettled(files.map(file => readExcel(file, intendedType)));
    const notes = [];
    let success = 0;
    const assigned = new Set();
    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      if (result.status === "fulfilled") {
        const item = result.value;
        if (assigned.has(item.type)) {
          notes.push("ERROR: Se repitió " + SCHEMAS[item.type].title +
            ". Selecciona solamente un archivo por tipo.");
          continue;
        }
        assigned.add(item.type);
        loaded[item.type] = item;
        success++;
        if (item.corrected.length) {
          notes.push(item.name + ": encabezados normalizados (" +
            item.corrected.join(", ") + ").");
        }
        item.warnings.forEach(warning => notes.push(item.name + ": " + warning));
      } else {
        const error = result.reason;
        notes.push("ERROR: " + (error && error.message ? error.message : "No se pudo procesar un archivo."));
      }
    }
    const failed = files.length - success;
    const ready = TYPES.filter(type => loaded[type]).length;
    const summary = ready === 3
      ? "¡Los tres Excel están limpios! Pulsa «Descargar los 3 Excel»."
      : success
        ? success + (success === 1 ? " archivo preparado." : " archivos preparados.")
        : "";
    showNotice([summary, ...notes].filter(Boolean).join("\n"), failed > 0);
  } finally {
    reading = false;
    render();
  }
}

function render() {
  const count = TYPES.filter(type => loaded[type]).length;
  $("#readyCount").innerHTML = count + ' <small>/ 3</small>';
  $("#fileProgress").textContent = count + " de 3 listos";
  $("#downloadAll").disabled = reading || count !== 3 || !window.XLSX || !window.JSZip;
  for (const type of TYPES) {
    const item = loaded[type];
    const card = document.querySelector('[data-role="' + type + '"]');
    card.classList.toggle("ready", Boolean(item));
    card.querySelector("[data-state]").textContent = item ? "Listo" : "Pendiente";
    card.querySelector("[data-details]").textContent = item
      ? item.name + " · " + item.dataCount + " registros · " + item.width + " columnas"
      : "Ningún archivo cargado.";
    card.querySelector("[data-download]").disabled = reading || !item || !window.XLSX;
  }
  renderPreviews();
}

function renderPreviews() {
  const preview = $("#preview");
  const wrapper = $("#previewTables");
  wrapper.replaceChildren();
  preview.hidden = !TYPES.some(type => loaded[type]);
  for (const type of TYPES) {
    const item = loaded[type];
    if (!item) continue;
    const section = document.createElement("div");
    section.className = "preview-card";
    const title = document.createElement("div");
    title.className = "preview-title";
    title.textContent = SCHEMAS[type].title;
    section.append(title);
    const scroller = document.createElement("div");
    scroller.className = "table-scroller";
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const thRow = document.createElement("tr");
    (item.rows[0] || []).forEach(cell => {
      const th = document.createElement("th");
      th.textContent = shortText(cell, 90);
      thRow.append(th);
    });
    thead.append(thRow);
    table.append(thead);
    const tbody = document.createElement("tbody");
    item.rows.slice(1, 4).forEach(row => {
      const tr = document.createElement("tr");
      row.forEach(cell => {
        const td = document.createElement("td");
        td.textContent = shortText(cell);
        tr.append(td);
      });
      tbody.append(tr);
    });
    table.append(tbody);
    scroller.append(table);
    section.append(scroller);
    wrapper.append(section);
  }
}

function cleanedName(item) {
  const names = {
    base: "Redes_base_SIN_FORMATO.xlsx",
    unidades: "PEA_unidades_SIN_FORMATO.xlsx",
    actividades: "PEA_actividades_SIN_FORMATO.xlsx"
  };
  return names[item.type];
}

function outputBytes(item) {
  const book = XLSX.utils.book_new();
  // Creating a new sheet from raw values intentionally discards ALL input styles.
  const sheet = XLSX.utils.aoa_to_sheet(item.rows, { cellDates: false });
  XLSX.utils.book_append_sheet(book, sheet, "Datos");
  return XLSX.write(book, { type: "array", bookType: "xlsx", compression: true });
}

function downloadBlob(blob, filename) {
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function downloadOne(type) {
  const item = loaded[type];
  if (!item) return;
  try {
    downloadBlob(new Blob([item.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), cleanedName(item));
    showNotice("Descargado: " + cleanedName(item) + ". Sin estilos y conservando sus datos.", false);
  } catch (e) {
    showNotice("No se pudo generar el archivo: " + (e.message || e), true);
  }
}

async function downloadAll() {
  if (!TYPES.every(type => loaded[type])) {
    showNotice("Selecciona los tres archivos antes de descargar el paquete.", true);
    return;
  }
  if (!window.JSZip) {
    showNotice("No se cargó el compresor ZIP. Comprueba tu conexión o descarga cada Excel por separado.", true);
    return;
  }
  const button = $("#downloadAll");
  button.disabled = true;
  const oldHtml = button.innerHTML;
  button.textContent = "Preparando descarga…";
  try {
    const zip = new JSZip();
    for (const type of TYPES) {
      const item = loaded[type];
      zip.file(cleanedName(item), item.bytes);
    }
    // Los XLSX ya son ZIP comprimidos: volver a comprimirlos solo consume tiempo.
    const blob = await zip.generateAsync({ type: "blob", compression: "STORE" });
    downloadBlob(blob, "Matriz_CCC_3_Excel_sin_formato.zip");
    showNotice("Descarga preparada: ZIP con Redes base, PEA unidades y PEA actividades, sin formato. Los documentos originales no se modificaron.");
  } catch (e) {
    showNotice("No se pudo generar el ZIP: " + (e.message || e), true);
  } finally {
    button.innerHTML = oldHtml;
    render();
  }
}

function resetFiles() {
  TYPES.forEach(type => delete loaded[type]);
  document.querySelectorAll('input[type="file"]').forEach(input => { input.value = ""; });
  showNotice("");
  render();
}

document.addEventListener("DOMContentLoaded", () => {
  $("#chooseAll").addEventListener("click", event => { event.stopPropagation(); $("#multiInput").click(); });
  $("#multiInput").addEventListener("change", event => acceptFiles(event.target.files));

  for (const type of TYPES) {
    document.querySelector('[data-pick="' + type + '"]').addEventListener("click", () => {
      document.querySelector('[data-file="' + type + '"]').click();
    });
    document.querySelector('[data-file="' + type + '"]').addEventListener("change", event => acceptFiles(event.target.files, type));
    document.querySelector('[data-download="' + type + '"]').addEventListener("click", () => downloadOne(type));
  }
  $("#downloadAll").addEventListener("click", downloadAll);
  $("#reset").addEventListener("click", resetFiles);

  const drop = $("#dropzone");
  drop.addEventListener("click", event => {
    if (event.target.closest("button")) return;
    $("#multiInput").click();
  });
  drop.addEventListener("keydown", event => {
    if (event.target === drop && (event.key === "Enter" || event.key === " ")) {
      event.preventDefault();
      $("#multiInput").click();
    }
  });
  ["dragenter", "dragover"].forEach(name => drop.addEventListener(name, event => {
    event.preventDefault();
    drop.classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach(name => drop.addEventListener(name, event => {
    event.preventDefault();
    drop.classList.remove("dragging");
  }));
  drop.addEventListener("drop", event => acceptFiles(event.dataTransfer.files));
  render();
});
