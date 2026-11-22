/**
 * API de la invitación: cada familia confirma desde su link personal (?f=código) y se guarda en la hoja "Invitados".
 * Se pega en la planilla (Extensiones → Apps Script) y se implementa como "Aplicación web".
 * Al actualizar el código: Implementar → Gestionar implementaciones → editar → Versión nueva (así el link no cambia).
 *
 * Hojas:
 * - Invitados: grupo | nombre | apellido | asiste | actualizado
 * - Links: una fila por grupo con su código secreto, el link y un botón para mandarlo por WhatsApp (menú 🧸 Invitación → Generar links)
 * - Historial: cada cambio de respuesta, con fecha
 *
 * Avisos (Configuración del proyecto → Propiedades del script; los datos de contacto no van en el código, que es público):
 * - El mail va siempre a la cuenta dueña del script. AVISO_MAILS: otros mails separados por coma.
 * - WhatsApp (CallMeBot): WSP con "telefono:apikey" separados por coma (ej. 5491126799971:123456).
 *   Cada teléfono activa CallMeBot por su cuenta y recibe su propia apikey.
 */
const HOJA = "Invitados", HOJA_LINKS = "Links", HOJA_HISTORIAL = "Historial";
const SITIO = "https://invitacion-babyshower.github.io/";
const CORTE = new Date("2027-01-16T23:59:59-03:00"); // último momento para confirmar o cambiar
const MENSAJE = "¡Hola! 🧸🎈 Te compartimos la invitación al Baby Shower y Revelación de José y Johanna. Desde este link podés ver todos los detalles y confirmar tu asistencia:";

function doGet(e) {
  const g = grupoDe(e.parameter.f);
  if (!g) return json({error: "codigo"});
  return json({g, abierto: new Date() <= CORTE, miembros: filas().filter(p => p.g === g).map(p => ({fila: p.fila, n: p.n, a: p.a, asiste: p.asiste}))});
}

function doPost(e) {
  let datos;
  try { datos = JSON.parse(e.postData.contents); } catch (err) { return json({error: "formato"}); }
  return json(guardar(datos));
}

// ---------- Lectura ----------
const libro = () => SpreadsheetApp.getActive();

function filas() {
  const v = libro().getSheetByName(HOJA).getDataRange().getValues();
  return v.slice(1).map((r, i) => ({fila: i + 2, g: String(r[0]).trim(), n: String(r[1]).trim(), a: String(r[2]).trim(), asiste: String(r[3]).trim()}))
    .filter(p => p.g && p.n);
}

// El código del link es la única forma de llegar a un grupo: sin buscador, nadie ve ni toca a otra familia
function grupoDe(codigo) {
  codigo = String(codigo || "").trim();
  if (codigo.length < 8) return null;
  const hoja = libro().getSheetByName(HOJA_LINKS);
  if (!hoja) return null;
  const fila = hoja.getDataRange().getValues().slice(1).find(r => String(r[1]).trim() === codigo);
  return fila ? String(fila[0]).trim() : null;
}

// ---------- Escritura ----------
function guardar(d) {
  if (new Date() > CORTE) return {error: "cerrado"};
  const g = d && grupoDe(d.f);
  if (!g) return {error: "codigo"};
  if (!Array.isArray(d.respuestas) || !d.respuestas.length) return {error: "datos"};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const porFila = new Map(filas().filter(p => p.g === g).map(p => [p.fila, p]));
    const hoja = libro().getSheetByName(HOJA), ahora = new Date(), cambios = [], historial = [];
    d.respuestas.forEach(r => {
      const p = porFila.get(Number(r.fila));
      if (!p || (r.asiste !== "sí" && r.asiste !== "no")) return; // solo personas de ese grupo y valores válidos
      hoja.getRange(p.fila, 4, 1, 2).setValues([[r.asiste, ahora]]);
      cambios.push(`${r.asiste === "sí" ? "✅" : "❌"} ${p.n} ${p.a}`);
      if (p.asiste !== r.asiste) historial.push([ahora, g, `${p.n} ${p.a}`, p.asiste || "(sin responder)", r.asiste]);
    });
    if (!cambios.length) return {error: "datos"};
    if (historial.length) {
      const h = hojaHistorial();
      h.getRange(h.getLastRow() + 1, 1, historial.length, 5).setValues(historial);
    }
    avisar(g, cambios);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function hojaHistorial() {
  let h = libro().getSheetByName(HOJA_HISTORIAL);
  if (!h) {
    h = libro().insertSheet(HOJA_HISTORIAL);
    h.appendRow(["fecha", "grupo", "persona", "antes", "ahora"]);
    h.getRange("A1:E1").setFontWeight("bold");
    h.setFrozenRows(1);
  }
  return h;
}

function avisar(g, cambios) {
  const texto = `Confirmación de ${g}:\n${cambios.join("\n")}`;
  const prop = PropertiesService.getScriptProperties();
  const lista = k => (prop.getProperty(k) || "").split(",").map(s => s.trim()).filter(Boolean);
  const mails = [Session.getEffectiveUser().getEmail(), ...lista("AVISO_MAILS")];
  try { MailApp.sendEmail(mails.join(","), `🧸 Confirmación: ${g}`, texto); } catch (err) { console.error(err); }
  lista("WSP").forEach(par => {
    const [tel, key] = par.split(":");
    try { UrlFetchApp.fetch(`https://api.callmebot.com/whatsapp.php?phone=${tel}&apikey=${key}&text=${encodeURIComponent(texto)}`); } catch (err) { console.error(err); }
  });
}

// ---------- Links por familia ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🧸 Invitación").addItem("Generar links", "generarLinks").addToUi();
}

// Agrega una fila por cada grupo nuevo; los códigos existentes no cambian (los links ya enviados siguen andando)
function generarLinks() {
  let hoja = libro().getSheetByName(HOJA_LINKS);
  if (!hoja) {
    hoja = libro().insertSheet(HOJA_LINKS);
    hoja.appendRow(["grupo", "código", "link", "enviar"]);
    hoja.getRange("A1:D1").setFontWeight("bold");
    hoja.setFrozenRows(1);
  }
  const existentes = new Set(hoja.getDataRange().getValues().slice(1).map(r => String(r[0]).trim()));
  const nuevos = [...new Set(filas().map(p => p.g))].filter(g => !existentes.has(g));
  nuevos.forEach(g => {
    const codigo = Utilities.getUuid().replace(/-/g, "").slice(0, 12);
    const link = `${SITIO}?f=${codigo}`;
    const wa = `https://wa.me/?text=${encodeURIComponent(`${MENSAJE}\n${link}`)}`;
    hoja.appendRow([g, codigo, link, ""]);
    // Link en la celda (no fórmula HYPERLINK: su separador cambia según el idioma de la planilla)
    hoja.getRange(hoja.getLastRow(), 4).setRichTextValue(SpreadsheetApp.newRichTextValue().setText("📲 Enviar por WhatsApp").setLinkUrl(wa).build());
  });
  hoja.autoResizeColumns(1, 4);
  return nuevos.length;
}

const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// Prueba rápida desde el editor: Ejecutar → probar
function probar() {
  generarLinks();
  const [g, codigo] = libro().getSheetByName(HOJA_LINKS).getRange(2, 1, 1, 2).getValues()[0];
  if (grupoDe(codigo) !== g) throw new Error("el código no lleva a su grupo");
  if (grupoDe("cualquiera") !== null || grupoDe("") !== null) throw new Error("un código inventado no debería andar");
  console.log(`OK: ${g} → ${codigo}`);
}
