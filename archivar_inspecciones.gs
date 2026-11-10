// ===== Archivado mensual: Inspecciones -> Inspecciones Histórico =====
// Pegar en el Apps Script del Sheet InspeccionTiempoReal como archivo NUEVO (no reemplaza nada de lo existente).
// Mantiene en "Inspecciones" el mes actual + el mes anterior; lo más viejo pasa a "Inspecciones Histórico"
// (mismas columnas). Solo mueve el bloque de filas viejas del inicio de la hoja (las filas están en orden de llegada).
// Copia primero al Histórico y después borra del origen: si algo falla a medias, la fila queda duplicada, nunca perdida.
// Uso: 1) Ejecutar archivarInspecciones una vez (si el registro dice "faltan filas", ejecutarla de nuevo).
//      2) Ejecutar instalarArchivadoMensual una vez -> corre solo el día 1 de cada mes a las 3 a.m.

var ARCH_ORIGEN = 'Inspecciones';
var ARCH_DESTINO = 'Inspecciones Histórico';
var ARCH_LOTE = 2000;            // filas por vuelta (cada vuelta bloquea solo unos segundos)
var ARCH_MAX_MS = 4.5 * 60000;   // Apps Script corta a los 6 min

function archivarCorte_() {
  // primer día del mes anterior, 00:00 hora Costa Rica (= 06:00 UTC)
  var ahora = new Date(Date.now() - 6 * 3600000);
  return new Date(Date.UTC(ahora.getUTCFullYear(), ahora.getUTCMonth() - 1, 1, 6, 0, 0));
}

function archivarInspecciones() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(ARCH_ORIGEN), dst = ss.getSheetByName(ARCH_DESTINO);
  if (!src || !dst) throw new Error('No existe la hoja ' + (!src ? ARCH_ORIGEN : ARCH_DESTINO));
  var cols = src.getLastColumn();
  var hdrS = src.getRange(1, 1, 1, cols).getValues()[0].map(function (h) { return String(h).trim(); });
  var hdrD = dst.getRange(1, 1, 1, dst.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  if (hdrS.join('|') !== hdrD.slice(0, cols).join('|')) throw new Error('Las columnas de las dos hojas no coinciden; no se movió nada.');
  var colF = hdrS.indexOf('fecha');
  if (colF < 0) throw new Error('No encuentro la columna fecha');
  var corte = archivarCorte_(), total = 0, faltan = false;
  var lock = LockService.getScriptLock();

  while (true) {
    var seguir = false;
    lock.waitLock(20000); // doPost (guardado de tablets) usa el mismo lock: espera solo lo que dura esta vuelta
    try {
      var last = src.getLastRow();
      if (last >= 2) {
        var vals = src.getRange(2, 1, Math.min(ARCH_LOTE, last - 1), cols).getValues();
        var k = 0; // filas viejas seguidas desde arriba
        while (k < vals.length) {
          var f = vals[k][colF], d = (f instanceof Date) ? f : new Date(f);
          if (isNaN(d.getTime()) || d >= corte) break;
          k++;
        }
        if (k > 0) {
          dst.getRange(dst.getLastRow() + 1, 1, k, cols).setValues(vals.slice(0, k));
          SpreadsheetApp.flush();
          src.deleteRows(2, k);
          SpreadsheetApp.flush();
          total += k;
          seguir = (k === vals.length); // si k < vuelta, ya llegó a filas recientes
        }
      }
    } finally {
      lock.releaseLock();
    }
    if (!seguir) break;
    if (Date.now() - t0 > ARCH_MAX_MS) { faltan = true; break; }
    Utilities.sleep(500); // deja pasar guardados pendientes
  }
  var msg = 'Archivadas ' + total + ' filas anteriores a ' + Utilities.formatDate(corte, 'America/Costa_Rica', 'yyyy-MM-dd') +
            (faltan ? ' · faltan filas: ejecutar de nuevo' : ' · listo');
  Logger.log(msg);
  return msg;
}

function instalarArchivadoMensual() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'archivarInspecciones') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archivarInspecciones').timeBased().onMonthDay(1).atHour(3).inTimezone('America/Costa_Rica').create();
  Logger.log('Listo: archivado automático el día 1 de cada mes, 3 a.m.');
}
