/**
 * ELOFEST 2026 — receptor para Google Sheets
 *
 * Opción recomendada: abre la hoja donde guardarás las respuestas,
 * ve a Extensiones > Apps Script y pega este archivo completo.
 */

const NOMBRE_HOJA = "Invitados";
const MAX_PERSONAS = 20;

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const datos = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const nombre = limpiarTexto_(datos.nombre, 100);
    const asistire = limpiarTexto_(datos.asistire, 2);
    const personas = Number(datos.personas);

    if (nombre.length < 3) {
      return responder_({ ok: false, message: "El nombre no es válido." });
    }

    if (!["Sí", "No"].includes(asistire)) {
      return responder_({ ok: false, message: "La respuesta de asistencia no es válida." });
    }

    const total = asistire === "No" ? 0 : personas;
    if (!Number.isInteger(total) || total < 0 || total > MAX_PERSONAS) {
      return responder_({ ok: false, message: "El número de personas no es válido." });
    }

    const libro = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = libro.getSheetByName(NOMBRE_HOJA);

    if (!hoja) {
      hoja = libro.insertSheet(NOMBRE_HOJA);
    }

    if (hoja.getLastRow() === 0) {
      hoja.appendRow(["Fecha y hora", "Nombre completo", "¿Asistirá?", "Número de personas"]);
      hoja.getRange(1, 1, 1, 4)
        .setFontWeight("bold")
        .setBackground("#7d2018")
        .setFontColor("#ffffff");
      hoja.setFrozenRows(1);
    }

    hoja.appendRow([new Date(), nombre, asistire, total]);

    return responder_({ ok: true, message: "Respuesta registrada." });
  } catch (error) {
    console.error(error);
    return responder_({ ok: false, message: "Error interno al guardar la respuesta." });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return responder_({ ok: true, message: "El formulario ELOFEST está conectado." });
}

function limpiarTexto_(valor, maximo) {
  return String(valor || "")
    .trim()
    .replace(/[<>]/g, "")
    .slice(0, maximo);
}

function responder_(contenido) {
  return ContentService
    .createTextOutput(JSON.stringify(contenido))
    .setMimeType(ContentService.MimeType.JSON);
}
