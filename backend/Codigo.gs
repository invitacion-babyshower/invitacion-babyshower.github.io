/**
 * API de la invitación: busca invitados y guarda confirmaciones en la hoja "Invitados".
 * Se pega en la planilla (Extensiones → Apps Script) y se implementa como "Aplicación web".
 *
 * Avisos opcionales por WhatsApp (CallMeBot): en Configuración del proyecto → Propiedades del script,
 * agregá WSP_TEL (ej. 5491126799971) y WSP_APIKEY (la clave que te manda CallMeBot).
 * El mail de aviso va siempre a la cuenta dueña del script.
 */
const HOJA = "Invitados";
const CORTE = new Date("2027-01-16T23:59:59-03:00"); // último momento para confirmar o cambiar
const MAX_RESULTADOS = 8;

function doGet(e) {
  const p = e.parameter;
  if (p.accion === "buscar") return json(buscar(p.q || ""));
  if (p.accion === "grupo") return json(grupo(p.g || ""));
  return json({error: "accion"});
}

function doPost(e) {
  let datos;
  try { datos = JSON.parse(e.postData.contents); } catch (err) { return json({error: "formato"}); }
  return json(guardar(datos));
}

// ---------- Lectura ----------
const norm = t => String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();

function filas() {
  const v = SpreadsheetApp.getActive().getSheetByName(HOJA).getDataRange().getValues();
  return v.slice(1).map((r, i) => ({fila: i + 2, g: String(r[0]).trim(), n: String(r[1]).trim(), a: String(r[2]).trim(), asiste: String(r[3]).trim()}))
    .filter(p => p.g && p.n);
}

// Solo devuelve coincidencias (nunca la lista completa): mínimo 3 letras y todas las palabras tienen que aparecer
function buscar(q) {
  if (norm(q).replace(/\s/g, "").length < 3) return [];
  const toks = norm(q).split(/\s+/);
  return filas().filter(p => toks.every(t => norm(p.n + " " + p.a).includes(t)))
    .slice(0, MAX_RESULTADOS).map(p => ({n: p.n, a: p.a, g: p.g}));
}

function grupo(g) {
  return {abierto: new Date() <= CORTE, miembros: filas().filter(p => p.g === g).map(p => ({fila: p.fila, n: p.n, a: p.a, asiste: p.asiste}))};
}

// ---------- Escritura ----------
function guardar(d) {
  if (new Date() > CORTE) return {error: "cerrado"};
  if (!d || !d.g || !Array.isArray(d.respuestas) || !d.respuestas.length) return {error: "datos"};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const miembros = filas().filter(p => p.g === d.g);
    const porFila = new Map(miembros.map(p => [p.fila, p]));
    const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA), ahora = new Date(), cambios = [];
    d.respuestas.forEach(r => {
      const p = porFila.get(Number(r.fila));
      if (!p || (r.asiste !== "sí" && r.asiste !== "no")) return; // solo filas de ese grupo y valores válidos
      hoja.getRange(p.fila, 4, 1, 2).setValues([[r.asiste, ahora]]);
      cambios.push(`${r.asiste === "sí" ? "✅" : "❌"} ${p.n} ${p.a}`);
    });
    if (!cambios.length) return {error: "datos"};
    avisar(d.g, cambios);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function avisar(g, cambios) {
  const texto = `Confirmación de ${g}:\n${cambios.join("\n")}`;
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), `🧸 Confirmación: ${g}`, texto); } catch (err) { console.error(err); }
  const prop = PropertiesService.getScriptProperties(), tel = prop.getProperty("WSP_TEL"), key = prop.getProperty("WSP_APIKEY");
  if (tel && key) {
    try { UrlFetchApp.fetch(`https://api.callmebot.com/whatsapp.php?phone=${tel}&apikey=${key}&text=${encodeURIComponent(texto)}`); } catch (err) { console.error(err); }
  }
}

const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// Prueba rápida desde el editor: Ejecutar → probar
function probar() {
  if (buscar("ab").length !== 0) throw new Error("mínimo 3 letras"); // Apps Script no tiene console.assert
  console.log(JSON.stringify(buscar("jose")), JSON.stringify(buscar("PAUCAR").length));
}
