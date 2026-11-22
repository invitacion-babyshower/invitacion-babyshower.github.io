/**
 * API de la invitación: cada familia confirma desde su link personal (?f=código) y se guarda en la hoja "Invitados".
 * Se pega en la planilla (Extensiones → Apps Script) y se implementa como "Aplicación web".
 * Al actualizar el código: Implementar → Gestionar implementaciones → editar → Versión nueva (así el link no cambia).
 *
 * Hojas:
 * - Invitados: grupo | nombre | apellido | asiste | actualizado
 * - Links: una fila por grupo con su código secreto, el link, un botón para mandarlo por WhatsApp (menú 🧸 Invitación → Generar links)
 *   y la fecha en que se mandó (la anota el panel: ver crearPanel)
 * - Historial: cada cambio de respuesta, con fecha
 *
 * Avisos (Configuración del proyecto → Propiedades del script; los datos de contacto no van en el código, que es público):
 * - El mail va siempre a la cuenta dueña del script. AVISO_MAILS: otros mails separados por coma.
 * - WhatsApp (CallMeBot): WSP con "telefono:apikey" separados por coma (ej. 5491126799971:123456).
 *   Cada teléfono activa CallMeBot por su cuenta y recibe su propia apikey.
 */
const HOJA = "Invitados", HOJA_LINKS = "Links", HOJA_HISTORIAL = "Historial";
const SITIO = "https://invitacion-babyshower.github.io/";
const CORTE = new Date("2027-01-03T23:59:59-03:00"); // último momento para confirmar o cambiar
const MENSAJE = "¡Hola! 🧸🎈 Te compartimos la invitación al Baby Shower y Revelación de José y Johanna. Desde este link podés ver todos los detalles y confirmar tu asistencia:";

function doGet(e) {
  if (e.parameter.panel !== undefined) return panel(e.parameter.panel);
  if (e.parameter.asistencia !== undefined) return asistencia(e.parameter.asistencia);
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
    hoja.appendRow([g, codigo, `${SITIO}?f=${codigo}`, ""]);
  });
  // Rehace el botón de WhatsApp de todas las filas (si cambia MENSAJE, se actualiza sin tocar los códigos).
  // api.whatsapp.com y no wa.me: la redirección de wa.me rompe los emojis (llegan como ��).
  // Link en la celda (no fórmula HYPERLINK: su separador cambia según el idioma de la planilla)
  const n = hoja.getLastRow() - 1;
  if (n > 0) hoja.getRange(2, 4, n, 1).setRichTextValues(hoja.getRange(2, 3, n, 1).getValues().map(([link]) =>
    [SpreadsheetApp.newRichTextValue().setText("📲 Enviar por WhatsApp").setLinkUrl(waLink(link)).build()]));
  hoja.autoResizeColumns(1, 4);
  return nuevos.length;
}

const waLink = link => `https://api.whatsapp.com/send?text=${encodeURIComponent(`${MENSAJE}\n${link}`)}`;

// ---------- Panel para mandar las invitaciones (celular) ----------
// Link: <URL de la web app>?panel=<clave>. La clave vive en la propiedad PANEL (no en el código, que es público);
// se crea con crearPanel. Quien tenga el link ve los links de todas las familias: compartirlo solo con quien manda.
const panelOk = clave => !!clave && clave === PropertiesService.getScriptProperties().getProperty("PANEL");

function crearPanel() {
  const prop = PropertiesService.getScriptProperties();
  if (!prop.getProperty("PANEL")) prop.setProperty("PANEL", Utilities.getUuid().replace(/-/g, ""));
  console.log(`${ScriptApp.getService().getUrl()}?panel=${prop.getProperty("PANEL")}`);
}

const esc = t => String(t).replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));

function panel(clave) {
  if (!panelOk(clave)) return HtmlService.createHtmlOutput("Link no válido.");
  const tz = Session.getScriptTimeZone();
  const familias = libro().getSheetByName(HOJA_LINKS).getDataRange().getValues().slice(1).filter(r => r[0] && r[1]);
  const items = familias.map(([g, codigo, link, , enviado]) => `
    <li class="${enviado ? "ok" : ""}"><div><b>${esc(g)}</b><small>${enviado ? "✓ Enviada " + Utilities.formatDate(new Date(enviado), tz, "dd/MM HH:mm") : "Sin enviar"}</small></div>
    <a href="${esc(waLink(link))}" target="_blank" onclick="marcar(this, '${esc(codigo)}')">📲 Enviar</a></li>`).join("");
  const html = `<style>
    body{font-family:system-ui,sans-serif;margin:0;padding:16px;background:#FBF6EF;color:#6B4E39}
    h1{font-size:1.3rem;margin:0 0 4px} p{margin:0 0 16px;color:#86624A}
    ul{list-style:none;padding:0;margin:0;display:grid;gap:10px}
    li{display:flex;align-items:center;justify-content:space-between;gap:12px;background:#fff;border-radius:14px;padding:12px 14px;box-shadow:0 4px 14px rgba(107,78,57,.1)}
    li div{display:grid;gap:2px} small{color:#86624A} li.ok{opacity:.6} li.ok small{color:#117A4B}
    a{flex:none;background:#117A4B;color:#fff;text-decoration:none;font-weight:600;padding:12px 16px;border-radius:999px}
    a.ver{display:block;text-align:center;margin:0 0 16px;background:#fff;color:#6B4E39;border:1.5px solid #E8D9C5}
  </style>
  <h1>🧸 Invitaciones para mandar</h1>
  <p>Tocá <b>Enviar</b>, elegí el contacto en WhatsApp y mandalo. Queda marcada como enviada para todos.</p>
  <a class="ver" href="${SITIO}asistencia#${esc(clave)}" target="_top">👥 Ver quién viene</a>
  <ul>${items}</ul>
  <script>
    function marcar(a, codigo){
      const li = a.closest("li");
      google.script.run.withSuccessHandler(t => { li.classList.add("ok"); li.querySelector("small").textContent = "✓ Enviada " + t; })
        .marcarEnviado(${JSON.stringify(clave)}, codigo);
    }
  </script>`;
  return HtmlService.createHtmlOutput(html).setTitle("Invitaciones").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ---------- Quién viene (celular) ----------
// Link: <URL de la web app>?asistencia=<clave>, misma clave PANEL. Solo lectura: la hoja Invitados agrupada por familia.
function asistencia(clave) {
  if (!panelOk(clave)) return HtmlService.createHtmlOutput("Link no válido.");
  const personas = filas(), grupos = new Map();
  personas.forEach(p => (grupos.get(p.g) || grupos.set(p.g, []).get(p.g)).push(p));
  const cuenta = l => {
    const si = l.filter(p => p.asiste === "sí").length, no = l.filter(p => p.asiste === "no").length;
    return {si, no, pend: l.length - si - no};
  };
  const estado = a => a === "sí" ? ["si", "✅ Asiste"] : a === "no" ? ["no", "❌ No asiste"] : ["pend", "⏳ Sin responder"];
  const t = cuenta(personas);
  const familias = [...grupos].map(([g, l]) => {
    const c = cuenta(l);
    return `<section><h2>${esc(g)}</h2><small>${c.si} asisten · ${c.no} no · ${c.pend} sin responder</small><ul>${l.map(p => {
      const [k, txt] = estado(p.asiste);
      return `<li class="${k}"><span>${esc(p.n)} ${esc(p.a)}</span><b>${txt}</b></li>`;
    }).join("")}</ul></section>`;
  }).join("");
  const html = `<style>
    body{font-family:system-ui,sans-serif;margin:0;padding:16px;background:#FBF6EF;color:#6B4E39}
    h1{font-size:1.3rem;margin:0 0 12px}
    .totales{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:16px}
    .totales div{background:#fff;border-radius:14px;padding:10px;text-align:center;box-shadow:0 4px 14px rgba(107,78,57,.1)}
    .totales b{display:block;font-size:1.6rem}
    section{background:#fff;border-radius:14px;padding:12px 14px;margin-bottom:12px;box-shadow:0 4px 14px rgba(107,78,57,.1)}
    h2{font-size:1.05rem;margin:0} small{color:#86624A}
    ul{list-style:none;padding:0;margin:8px 0 0}
    li{display:flex;justify-content:space-between;gap:12px;padding:8px 0;border-top:1px solid #F1E7DB}
    li b{flex:none;font-weight:600} .si b{color:#117A4B} .no b{color:#B3261E} .pend b{color:#86624A;font-weight:400}
  </style>
  <h1>🧸 ¿Quién viene?</h1>
  <div class="totales"><div><b>${t.si}</b>✅ Asisten</div><div><b>${t.no}</b>❌ No asisten</div><div><b>${t.pend}</b>⏳ Sin responder</div></div>
  ${familias}`;
  return HtmlService.createHtmlOutput(html).setTitle("Asistencia").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Lo llama el panel al tocar Enviar: anota la fecha en la columna E de Links
function marcarEnviado(clave, codigo) {
  if (!panelOk(clave)) throw new Error("clave");
  const hoja = libro().getSheetByName(HOJA_LINKS);
  if (!hoja.getRange("E1").getValue()) hoja.getRange("E1").setValue("enviado").setFontWeight("bold");
  const i = hoja.getDataRange().getValues().findIndex(r => String(r[1]).trim() === String(codigo));
  if (i < 1) throw new Error("codigo");
  const ahora = new Date();
  hoja.getRange(i + 1, 5).setValue(ahora);
  return Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM HH:mm");
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
