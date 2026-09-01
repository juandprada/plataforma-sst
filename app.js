"use strict";

// Estado en memoria.
let EMPRESAS = [];
let FORMATOS = [];
let ENCABEZADO_TPL = "";
let PRESUPUESTO = null; // data/presupuesto.json (parámetros por empresa + catálogo)
let CAPACITACIONES = null; // data/capacitaciones.json (contenido de la página 1 del acta)
let GTC45_DATA = null;    // data/gtc45.json (mapeo de peligros + resultados por empresa)
let INSPECCIONES = null;  // data/inspecciones.json
let CUENTAS_COBRO_DATA = null; // data/cuentas_cobro.json

// Formato cuya página 1 depende de la capacitación elegida (ver CAPACITACIONES).
const FORMATO_ASISTENCIA = "asistencia-a-capacitacion";
// Formato calculado: matriz de riesgos IPVER (GTC45). Depende del cargo elegido.
const FORMATO_IPVER = "matriz-ipver";
// Formato cuya página 1 depende de la inspección elegida (ver INSPECCIONES).
const FORMATO_INSPECCION = "formato-inspeccion";
// Formato de Cuenta de Cobro.
const FORMATO_CUENTA_COBRO = "cuenta-de-cobro";

const $ = (sel) => document.querySelector(sel);

function setEstado(msg, esError = false) {
  const el = $("#estado");
  el.textContent = msg || "";
  el.classList.toggle("error", !!esError);
}

// Escapa texto para insertarlo con seguridad como HTML.
function escapeHTML(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Reemplaza {{TOKEN}}. `raw` = claves cuyo valor ya es HTML y no se escapa.
function fillTokens(tpl, ctx, raw = []) {
  return tpl.replace(/\{\{\s*([A-Z0-9_]+)\s*\}\}/g, (match, key) => {
    if (!(key in ctx)) {
      console.warn("Token sin dato:", key);
      return "";
    }
    return raw.includes(key) ? ctx[key] : escapeHTML(ctx[key]);
  });
}

// Nombre propio a Título ("LILIANA MARÍA OROZCO" -> "Liliana María Orozco").
function tituloCase(str) {
  return String(str)
    .toLowerCase()
    .replace(/(^|[\s/.-])(\p{L})/gu, (_, sep, letra) => sep + letra.toUpperCase());
}

function logoHTML(empresa) {
  if (empresa.LOGO) {
    return `<img class="dh-logo-img" src="${empresa.LOGO}" alt="Logo ${escapeHTML(
      empresa.EMPRESA
    )}">`;
  }
  return `<span class="dh-logo-fallback">${escapeHTML(empresa.EMPRESA)}</span>`;
}

// ---- Presupuesto (formato calculado) -------------------------------------
// El .xlsx original es una calculadora: la tabla sale de 5 parámetros por empresa
// (trabajadores, extintores, botiquines, aires, pago de asesoría) y de un catálogo
// de ítems con su precio base. Ambos viven en data/presupuesto.json.
//
// Encima del cálculo van los AJUSTES: valor unitario y/o cantidad puestos a mano para
// una empresa y un ítem concretos (panel "¿Deseas modificar…?"). Precedencia:
//   catálogo + parámetros  <  presupuesto.json:ajustes_por_empresa  <  localStorage
// Pegar los ajustes locales en `ajustes_por_empresa` los vuelve permanentes para todos.

const pesos = (n) => "$ " + Math.round(n).toLocaleString("es-CO");

// Resuelve un valor unitario o una cantidad según su regla del catálogo:
//   {fijo: n} | {fijo: n, por: "<param>"} (n × parámetro) | {param: "<param>"}
function valorRegla(regla, params) {
  if (!regla) return 0;
  if ("param" in regla) return Number(params[regla.param]) || 0;
  const base = Number(regla.fijo) || 0;
  return "por" in regla ? base * (Number(params[regla.por]) || 0) : base;
}

const AJUSTES_KEY = "sst.presupuesto.ajustes.v1";
let AJUSTES = {}; // { <empresa._id>: { "<grupo>|<item>": {vu?, cant?} } }

function cargarAjustes() {
  try {
    AJUSTES = JSON.parse(localStorage.getItem(AJUSTES_KEY) || "{}") || {};
  } catch (err) {
    console.warn("Ajustes de presupuesto ilegibles, se ignoran.", err);
    AJUSTES = {};
  }
}

function guardarAjustes() {
  try {
    localStorage.setItem(AJUSTES_KEY, JSON.stringify(AJUSTES));
  } catch (err) {
    console.warn("No se pudieron guardar los ajustes.", err);
  }
}

// Ajustes efectivos de una empresa: los del JSON (permanentes) + los locales.
function ajustesDe(empresaId) {
  const delJSON = (PRESUPUESTO && PRESUPUESTO.ajustes_por_empresa) || {};
  const merge = {};
  for (const [clave, val] of Object.entries(delJSON[empresaId] || {})) {
    merge[clave] = { ...val };
  }
  for (const [clave, val] of Object.entries(AJUSTES[empresaId] || {})) {
    merge[clave] = { ...merge[clave], ...val };
  }
  return merge;
}

// Guarda (o borra, con valor null/NaN) un ajuste local. Si el valor vuelve a coincidir
// con el calculado, se borra: así el ítem queda "sin ajuste" y sigue las fórmulas.
function setAjuste(empresaId, clave, campo, valor, base) {
  const porEmpresa = AJUSTES[empresaId] || (AJUSTES[empresaId] = {});
  const item = porEmpresa[clave] || (porEmpresa[clave] = {});
  if (valor == null || !Number.isFinite(valor) || valor === base) delete item[campo];
  else item[campo] = valor;
  if (!Object.keys(item).length) delete porEmpresa[clave];
  if (!Object.keys(porEmpresa).length) delete AJUSTES[empresaId];
  guardarAjustes();
}

// Filas calculadas del presupuesto (con ajustes aplicados). Única fuente de verdad:
// la usan tanto la tabla del documento como el panel de edición.
function filasPresupuesto(empresa) {
  if (!PRESUPUESTO) return null;
  const params = {
    ...PRESUPUESTO.parametros_default,
    ...(PRESUPUESTO.parametros_por_empresa[empresa._id] || {}),
  };
  const ajustes = ajustesDe(empresa._id);

  return PRESUPUESTO.catalogo.map((grupo) => ({
    grupo: grupo.grupo,
    items: grupo.items.map((item) => {
      const clave = grupo.grupo + "|" + item.nombre;
      const aj = ajustes[clave] || {};
      const vuBase = valorRegla(item.vu, params);
      const cantBase = valorRegla(item.cant, params);
      const vu = Number.isFinite(aj.vu) ? aj.vu : vuBase;
      const cant = Number.isFinite(aj.cant) ? aj.cant : cantBase;
      return { nombre: item.nombre, clave, vu, cant, vuBase, cantBase };
    }),
  }));
}

// Arma la tabla del presupuesto (grupos + subtotales + totales) para una empresa.
function tablaPresupuestoHTML(empresa) {
  const grupos = filasPresupuesto(empresa);
  if (!grupos) return "<p>No se pudo cargar el presupuesto.</p>";

  const filas = [];
  let total = 0;
  for (const grupo of grupos) {
    filas.push(
      `<tr class="pr-grupo"><td colspan="4">${escapeHTML(grupo.grupo)}</td></tr>`
    );
    let subtotal = 0;
    for (const item of grupo.items) {
      const t = item.vu * item.cant;
      subtotal += t;
      filas.push(
        `<tr><td>${escapeHTML(item.nombre)}</td>` +
          `<td class="pr-num">${pesos(item.vu)}</td>` +
          `<td class="pr-num">${item.cant}</td>` +
          `<td class="pr-num">${pesos(t)}</td></tr>`
      );
    }
    total += subtotal;
    filas.push(
      `<tr class="pr-subtotal"><td colspan="3">SUBTOTAL</td>` +
        `<td class="pr-num">${pesos(subtotal)}</td></tr>`
    );
  }
  filas.push(
    `<tr class="pr-total"><td colspan="3">PRESUPUESTO TOTAL ASIGNADO</td>` +
      `<td class="pr-num">${pesos(total)}</td></tr>`,
    `<tr class="pr-total"><td colspan="3">PRESUPUESTO MENSUAL</td>` +
      `<td class="pr-num">${pesos(total / 12)}</td></tr>`
  );

  return (
    '<table class="doc-tabla tabla-presupuesto">' +
    "<colgroup><col style=\"width:46%\"><col style=\"width:19%\">" +
    "<col style=\"width:14%\"><col style=\"width:21%\"></colgroup>" +
    "<tr><th>ITEM</th><th>VALOR UNITARIO</th><th>CANTIDADES</th><th>VALOR TOTAL</th></tr>" +
    filas.join("") +
    "</table>"
  );
}

// ---- Panel de ajustes (solo en el formato presupuesto) --------------------
// No se ve por defecto: con el formato presupuesto aparece una sola línea preguntando
// si se desea modificar algo; el editor se despliega solo si el usuario dice que sí.

let editorAbierto = false;
let empresaPanel = null; // empresa que está mostrando el editor
let empresaPintada = null; // _id de la empresa cuyos inputs están en el DOM

function renderPanelPresupuesto(formato, empresa) {
  const panel = $("#panel-presupuesto");
  if (!panel) return;
  if (formato.id !== "presupuesto" || !PRESUPUESTO) {
    panel.hidden = true;
    editorAbierto = false;
    $("#pp-editor").hidden = true;
    return;
  }
  panel.hidden = false;
  empresaPanel = empresa;

  const n = Object.keys(ajustesDe(empresa._id)).length;
  $("#pp-aviso").textContent = n
    ? ` (hay ${n} ${n === 1 ? "ítem ajustado" : "ítems ajustados"} para esta empresa)`
    : "";
  $("#pp-editor").hidden = !editorAbierto;
  // Repintar solo si cambió la empresa: rehacer los inputs en cada regeneración le
  // quitaría el foco (y el valor a medio escribir) a quien está editando.
  if (editorAbierto && empresaPintada !== empresa._id) pintarItems(empresa);
}

// Una fila por ítem: nombre + valor unitario + cantidad, precargados con lo vigente.
function pintarItems(empresa) {
  empresaPintada = empresa._id;
  const grupos = filasPresupuesto(empresa);
  const partes = [];
  for (const grupo of grupos) {
    partes.push(`<h3 class="pp-grupo">${escapeHTML(grupo.grupo)}</h3>`);
    for (const item of grupo.items) {
      const editadoVu = item.vu !== item.vuBase;
      const editadoCant = item.cant !== item.cantBase;
      partes.push(
        `<div class="pp-item">` +
          `<span class="pp-nombre">${escapeHTML(item.nombre)}</span>` +
          `<label class="pp-campo">Valor unitario` +
          `<input type="number" min="0" step="1000" value="${item.vu}"` +
          ` data-clave="${escapeHTML(item.clave)}" data-campo="vu"` +
          ` data-base="${item.vuBase}" class="${editadoVu ? "pp-editado" : ""}"></label>` +
          `<label class="pp-campo">Cantidad` +
          `<input type="number" min="0" step="1" value="${item.cant}"` +
          ` data-clave="${escapeHTML(item.clave)}" data-campo="cant"` +
          ` data-base="${item.cantBase}" class="${editadoCant ? "pp-editado" : ""}"></label>` +
          `</div>`
      );
    }
  }
  $("#pp-items").innerHTML = partes.join("");
}

function conectarPanelPresupuesto() {
  $("#pp-abrir").addEventListener("click", () => {
    editorAbierto = true;
    $("#pp-editor").hidden = false;
    if (empresaPanel) pintarItems(empresaPanel);
  });

  $("#pp-cerrar").addEventListener("click", () => {
    editorAbierto = false;
    $("#pp-editor").hidden = true;
  });

  // `change` (no `input`): así el PDF se regenera al salir del campo, no por tecla.
  $("#pp-items").addEventListener("change", (ev) => {
    const input = ev.target.closest("input[data-clave]");
    if (!input || !empresaPanel) return;
    const base = Number(input.dataset.base);
    const valor = input.value === "" ? null : Number(input.value);
    setAjuste(empresaPanel._id, input.dataset.clave, input.dataset.campo, valor, base);
    if (valor == null || !Number.isFinite(valor)) input.value = base; // vacío = calculado
    input.classList.toggle("pp-editado", Number(input.value) !== base);
    solicitarGeneracion();
  });

}

// ---- Matriz IPVER (formato calculado — riesgos GTC45) ---------------------
// La tabla se calcula por empresa y cargo desde data/gtc45.json (storage).
// El storage combina el mapeo de peligros (catálogo estático del CSV) con los
// resultados de la encuesta SurveyJS por empresa. El script tools/gtc45_to_json.py
// actualiza el storage con merge/upsert cuando llegan encuestas nuevas.

// Tablas GTC45 (constantes del estándar, no van en el JSON).
const TABLA_ND = {
  "Muy Alto": 10, "Alto": 6, "Medio": 2, "Bajo": null, "No Aplica": 0,
};
const TABLA_NE = {
  "Continua": 4, "Frecuente": 3, "Ocasional": 2, "Esporadica": 1, "No Aplica": 0,
};

// Normaliza nombre de empresa para buscar match en el storage (misma lógica que el
// script Python: minúsculas, sin sufijos societarios, sin espacios extra).
function normalizarNombre(nombre) {
  return nombre.trim().toLowerCase()
    .replace(/\b(s\.?a\.?s\.?|s\.?a\.?|ltda\.?|e\.?u\.?)\b/g, "")
    .replace(/\s+/g, " ").trim();
}

// Busca la empresa del selector en los resultados GTC45 (matching flexible).
function buscarEmpresaGTC45(empresa) {
  if (!GTC45_DATA || !GTC45_DATA.resultados_por_empresa) return null;
  const clave = normalizarNombre(empresa.EMPRESA || "");
  return GTC45_DATA.resultados_por_empresa[clave] || null;
}

// Cargo seleccionado en el selector de la IPVER.
function cargoIPVERElegido() {
  const sel = $("#sel-cargo-ipver");
  return sel ? sel.value : "Administrativo";
}

// ¿El peligro aplica al cargo seleccionado?
function peligroAplicaACargo(peligro, cargo) {
  if (!peligro || !peligro.cargos || !peligro.cargos.length) return false;
  const cargoLow = cargo.toLowerCase();
  return peligro.cargos.some((c) => {
    const cl = c.toLowerCase();
    return cl === "todos" || cl === cargoLow;
  });
}

// Calcula los valores GTC45 para un peligro evaluado.
function calcularRiesgoGTC45(deficiencia, exposicion, nc) {
  const nd = TABLA_ND[deficiencia];
  const ne = TABLA_NE[exposicion] || 0;
  const esBajo = deficiencia === "Bajo";

  let np, nr;
  if (esBajo) {
    np = null;  // Bajo → directo a Nivel IV
    nr = 20;    // Valor fijo según la metodología
  } else {
    np = (nd || 0) * ne;
    nr = np * nc;
  }

  let inp;
  if (esBajo) inp = "N/A";
  else if (np >= 24) inp = "Muy Alto";
  else if (np >= 10) inp = "Alto";
  else if (np >= 6) inp = "Medio";
  else inp = "Bajo";

  let inr;
  if (nr >= 600) inr = "I";
  else if (nr >= 150) inr = "II";
  else if (nr >= 40) inr = "III";
  else inr = "IV";

  const aceptabilidad = {
    "I": "NO ACEPTABLE",
    "II": "NO ACEPTABLE O ACEPTABLE CON CONTROL ESPECÍFICO",
    "III": "MEJORABLE",
    "IV": "ACEPTABLE",
  }[inr];

  return { nd: esBajo ? "—" : nd, ne, np: esBajo ? "—" : np, nc, nr, inp, inr, aceptabilidad };
}

// Clase CSS para colorear la celda de nivel de riesgo.
function claseNivelRiesgo(inr) {
  return { "I": "nr-i", "II": "nr-ii", "III": "nr-iii", "IV": "nr-iv" }[inr] || "";
}

// Arma la tabla HTML de la matriz IPVER para una empresa y cargo.
function tablaIPVERHTML(empresa) {
  const datos = buscarEmpresaGTC45(empresa);
  if (!datos) {
    return '<p class="ipver-sin-datos">Esta empresa no tiene datos de inspección GTC45. ' +
      'Llene la encuesta de inspección y ejecute <code>tools/gtc45_to_json.py</code> ' +
      'para generar los datos.</p>';
  }

  const cargo = cargoIPVERElegido();
  const mapeo = GTC45_DATA.mapeo_peligros || [];
  const peligros = datos.peligros || {};

  // Filtrar peligros que aplican al cargo seleccionado.
  const filas = [];
  for (const cat of mapeo) {
    const p = peligros[cat.panel_id];
    if (!p) continue;  // No evaluado o "No Aplica"
    if (!peligroAplicaACargo(p, cargo)) continue;

    const r = calcularRiesgoGTC45(p.deficiencia, p.exposicion, cat.nc);
    filas.push({ cat, p, r });
  }

  if (!filas.length) {
    return `<p class="ipver-sin-datos">No hay peligros identificados para el cargo ` +
      `<strong>${escapeHTML(cargo)}</strong> en esta empresa.</p>`;
  }

  // Encabezado de la tabla (columnas GTC45).
  const th = `<tr class="ipver-header">
    <th rowspan="2">Proceso</th>
    <th rowspan="2">Peligro</th>
    <th rowspan="2">Efectos Posibles</th>
    <th colspan="3">Controles Existentes</th>
    <th colspan="7">Evaluación del Riesgo</th>
    <th rowspan="2">Aceptabilidad</th>
    <th rowspan="2">Requisito Legal</th>
  </tr>
  <tr class="ipver-header">
    <th>Fuente</th><th>Medio</th><th>Individuo / EPP</th>
    <th>ND</th><th>NE</th><th>NP</th><th>NC</th><th>NR</th>
    <th>Int. NR</th><th>Int. NP</th>
  </tr>`;

  const rows = filas.map(({ cat, p, r }) => {
    const controlInd = [cat.control_individuo, cat.control_epp]
      .filter((x) => x && x !== "NA" && x !== "No aplica").join("; ");
    const controlFuente = (!cat.control_fuente || cat.control_fuente === "NA")
      ? "—" : cat.control_fuente;
    return `<tr>
      <td>${escapeHTML(cat.proceso)}</td>
      <td>${escapeHTML(cat.actividades)}</td>
      <td>${escapeHTML(cat.efectos_posibles)}</td>
      <td>${escapeHTML(controlFuente)}</td>
      <td>${escapeHTML(cat.control_medio)}</td>
      <td>${escapeHTML(controlInd || "—")}</td>
      <td class="ipver-num">${r.nd}</td>
      <td class="ipver-num">${r.ne}</td>
      <td class="ipver-num">${r.np}</td>
      <td class="ipver-num">${r.nc}</td>
      <td class="ipver-num ${claseNivelRiesgo(r.inr)}">${r.nr}</td>
      <td class="ipver-num ${claseNivelRiesgo(r.inr)}">${r.inr}</td>
      <td class="ipver-num">${escapeHTML(r.inp)}</td>
      <td class="${claseNivelRiesgo(r.inr)}">${escapeHTML(r.aceptabilidad)}</td>
      <td class="ipver-legal">${escapeHTML(cat.requisito_legal)}</td>
    </tr>`;
  });

  // Encabezado informativo: empresa, cargo, # trabajadores.
  const trab = datos.trabajadores || {};
  const totalTrab = (trab.admin || 0) + (trab.operarios || 0) +
    (trab.servicios || 0) + (trab.conductores || 0);
  const infoHeader = `<div class="ipver-info">
    <p><strong>Cargo / Área evaluada:</strong> ${escapeHTML(cargo)} &nbsp;|&nbsp;
    <strong>Total trabajadores:</strong> ${totalTrab}
    (Admin: ${trab.admin || 0}, Operarios: ${trab.operarios || 0}, ` +
    `Serv. Generales: ${trab.servicios || 0}, Conductores: ${trab.conductores || 0})</p>
  </div>`;

  return infoHeader +
    '<table class="doc-tabla tabla-ipver">' + th + rows.join("") + '</table>';
}

// Muestra el selector de cargo solo en el formato IPVER (en el resto no pinta nada).
function renderSelectorCargoIPVER(formato) {
  const campo = $("#campo-cargo-ipver");
  if (campo) campo.hidden = formato.id !== FORMATO_IPVER;
}

// ---- Asistencia a capacitación (formato calculado) -----------------------
// La página 1 del acta cambia según la capacitación: su frase de objetivo y el temario
// salen de data/capacitaciones.json (lo edita el desarrollador, no el técnico; la app
// solo deja elegir). Lo que NO cambia entre capacitaciones —las viñetas de cultura
// preventiva y el marco normativo— se queda escrito en la plantilla.

// Todas las capacitaciones en una lista plana (los grupos son solo para el desplegable).
function capacitaciones() {
  return CAPACITACIONES ? CAPACITACIONES.grupos.flatMap((g) => g.items) : [];
}

function capacitacionElegida() {
  const items = capacitaciones();
  const sel = $("#sel-capacitacion");
  return items.find((c) => c.id === (sel && sel.value)) || items[0] || null;
}

function puntosCapacitacionHTML(cap) {
  if (!cap) return "";
  return (
    "<ul>" + cap.puntos.map((p) => `<li>${escapeHTML(p)}</li>`).join("") + "</ul>"
  );
}

// Muestra el selector solo en el formato de asistencia (en el resto no pinta nada).
function renderSelectorCapacitacion(formato) {
  const campo = $("#campo-capacitacion");
  if (campo) campo.hidden = formato.id !== FORMATO_ASISTENCIA;
}

// ---- Inspecciones (formato calculado) ------------------------------------
function inspecciones() {
  return INSPECCIONES ? INSPECCIONES.items : [];
}

function inspeccionElegida() {
  const items = inspecciones();
  const sel = $("#sel-inspeccion");
  return items.find((i) => i.id === (sel && sel.value)) || items[0] || null;
}

function puntosInspeccionHTML(insp) {
  if (!insp || !insp.puntos) return "";
  return "<ul>" + insp.puntos.map((p) => `<li>${escapeHTML(p)}</li>`).join("") + "</ul>";
}

function renderSelectorInspeccion(formato) {
  const campo = $("#campo-inspeccion");
  if (campo) campo.hidden = formato.id !== FORMATO_INSPECCION;
}

// ---- Cuenta de Cobro (formato mensual de honorarios SG-SST) -------------
const MESES_ES = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
];

// Convierte un número entero a texto en español (soporta millones y formato bancario/factura).
function numeroALetras(n) {
  n = Math.floor(Number(n) || 0);
  if (n === 0) return "CERO PESOS";
  const u = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
  const d1 = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
  const d2 = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
  const c3 = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];

  function w_1_99(x) {
    if (x < 10) return u[x];
    if (x < 20) return d1[x - 10];
    const d = Math.floor(x / 10);
    const r = x % 10;
    if (x === 20) return "VEINTE";
    if (r === 0) return d2[d];
    if (d === 2) return "VEINTI" + u[r];
    return d2[d] + " Y " + u[r];
  }

  function w_1_999(x) {
    if (x === 0) return "";
    if (x === 100) return "CIEN";
    const c = Math.floor(x / 100);
    const r = x % 100;
    const cent = c3[c];
    const rest = w_1_99(r);
    return [cent, rest].filter(Boolean).join(" ");
  }

  const mm = Math.floor(n / 1e6);
  const rmm = n % 1e6;
  const th = Math.floor(rmm / 1000);
  const rth = rmm % 1000;

  const parts = [];
  if (mm > 0) {
    if (mm === 1) parts.push("UN MILLÓN");
    else parts.push(w_1_999(mm) + " MILLONES");
    if (th === 0 && rth === 0) parts.push("DE");
  }
  if (th > 0) {
    if (th === 1) parts.push("MIL");
    else parts.push(w_1_999(th) + " MIL");
  }
  if (rth > 0) {
    parts.push(w_1_999(rth));
  }
  parts.push("PESOS");
  return parts.join(" ").replace(/\s+/g, " ").trim();
}

// Obtiene el último día hábil (lunes a viernes) del mes y año indicados.
function ultimoDiaHabil(anio, mes) {
  // mes: 1-12. El día 0 del mes siguiente es el último día del mes deseado.
  const fecha = new Date(anio, mes, 0);
  const diaSemana = fecha.getDay(); // 0 = Domingo, 6 = Sábado
  if (diaSemana === 0) fecha.setDate(fecha.getDate() - 2);
  else if (diaSemana === 6) fecha.setDate(fecha.getDate() - 1);
  return fecha.getDate();
}

function diaATexto(dia) {
  const u = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
  const d1 = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
  if (dia < 10) return u[dia];
  if (dia < 20) return d1[dia - 10];
  if (dia === 20) return "VEINTE";
  if (dia < 30) return "VEINTI" + u[dia % 10];
  if (dia === 30) return "TREINTA";
  if (dia === 31) return "TREINTA Y UNO";
  return String(dia);
}

function valoresComunesCuentaCobro() {
  return (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.valores_comunes) || [
    150000, 162500, 200000, 250000, 450000, 500000
  ];
}

function valorCuentaCobro(empresa) {
  if (!empresa) return 250000;
  const selValor = $("#sel-valor-cc");
  if (selValor && selValor.value !== "" && Number(selValor.value) > 0) {
    return Number(selValor.value);
  }
  const tarifas = (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifas_por_empresa) || {};
  return tarifas[empresa._id] || (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifa_default) || 250000;
}

function renderSelectorCuentaCobro(formato, empresa) {
  const campoMes = $("#campo-mes-cc");
  const campoValor = $("#campo-valor-cc");
  const esCC = formato.id === FORMATO_CUENTA_COBRO;
  if (campoMes) campoMes.hidden = !esCC;
  if (campoValor) campoValor.hidden = !esCC;

  if (esCC && empresa) {
    const selValor = $("#sel-valor-cc");
    if (selValor && (!selValor.dataset.empresa || selValor.dataset.empresa !== empresa._id)) {
      const tarifas = (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifas_por_empresa) || {};
      const val = tarifas[empresa._id] || (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifa_default) || 250000;
      selValor.value = String(val);
      selValor.dataset.empresa = empresa._id;
    }
  }
}

async function fetchText(url) {
  const r = await fetch(url, { cache: "no-cache" });
  if (!r.ok) throw new Error(`No se pudo cargar ${url} (${r.status})`);
  return r.text();
}

// ---- Plan de Respuestas (seguimiento de encuestas SG-SST) -------------------
// Fuente de catalogo de encuestas: workers/config.json del repo juandprada/respuestasencuestas.
// Fuente de respuestas: GitHub Contents API (repo publico; sin token requerido).
const PLAN_BASE = "https://juandprada.github.io/respuestasencuestas";
const PLAN_REPO_RAW = "https://raw.githubusercontent.com/juandprada/respuestasencuestas/main/workers/config.json";
const PLAN_REPO_API = "https://api.github.com/repos/juandprada/respuestasencuestas/contents";

function planSlugify(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

// nombre del archivo del worker: <slug_empresa>_<ISO8601>_<uuid8>.json
function planParseFilename(name) {
  const m = String(name).match(/^(.*?)_(\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}-\d{3}Z)_[a-f0-9]{8}\.json$/);
  if (!m) return null;
  return { empresa: m[1], ts: m[2], fecha: m[2].replace(/T(\d{2})-(\d{2})-(\d{2})-/, "T$1:$2:$3.") };
}

async function planFetchJSON(url) {
  const r = await fetch(url, { cache: "no-cache" });
  if (!r.ok) {
    const err = new Error(`${r.status}`);
    err.status = r.status;
    throw err;
  }
  return r.json();
}

async function planRespuestasPorEncuesta(slug) {
  const url = `${PLAN_REPO_API}/respuestas/${slug}`;
  try {
    const items = await planFetchJSON(url);
    return items.filter((it) => it.type === "file" && it.name.endsWith(".json"));
  } catch (e) {
    if (e.status === 404) return [];
    throw e;
  }
}

function planFechaBonita(iso) {
  if (!iso) return "";
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})-(\d{3})Z$/);
  if (!m) return iso;
  return `${m[3]}/${m[2]}/${m[1]} ${m[4]}:${m[5]}:${m[6]}Z`;
}

async function renderPlanRespuestas(empresa) {
  setEstado("Buscando encuestas realizadas…");
  const visor = $("#visor");
  const slugEmpresa = planSlugify(empresa.EMPRESA);

  // 1) catalogo = encuestas publicadas (config.json del repo)
  let cfg;
  try {
    cfg = await planFetchJSON(PLAN_REPO_RAW);
  } catch (e) {
    setEstado("No se pudo leer workers/config.json del repo: " + e.message, true);
    return;
  }
  const encuestas = Array.isArray(cfg.encuestas) ? cfg.encuestas : [];

  // 2) respuestas por encuesta
  let info = [];
  let tasaError = null;
  for (const enc of encuestas) {
    try {
      const archivos = await planRespuestasPorEncuesta(enc.slug);
      const mios = archivos
        .map((it) => ({ it, meta: planParseFilename(it.name) }))
        .filter((x) => x.meta && x.meta.empresa === slugEmpresa)
        .sort((a, b) => (a.meta.ts < b.meta.ts ? 1 : -1));
      info.push({ enc, total: archivos.length, ultimo: mios[0] || null, mios });
    } catch (e) {
      tasaError = e;
      info.push({ enc, total: 0, ultimo: null, mios: [] });
    }
  }

  // 3) pintar panel tipo "estatico"
  const hechas = info.filter((x) => x.ultimo).length;
  const filas = info.map(({ enc, ultimo, mios }) => {
    const urlEncuesta = `${PLAN_BASE}/encuestas/${enc.slug}.html?empresa=${encodeURIComponent(slugEmpresa)}`;
    const completada = !!ultimo;
    const it = completada ? ultimo.it : null;
    const fecha = completada ? planFechaBonita(ultimo.meta.ts) : "";
    const jsonUrl = completada
      ? `https://github.com/juandprada/respuestasencuestas/blob/main/${it.path.replace(/ /g, "%20")}`
      : "";
    const verLink = completada && mios.length > 0
      ? `<a href="https://github.com/juandprada/respuestasencuestas/tree/main/${enc.slug}" target="_blank" rel="noopener">todas (${info.length ? mios.length : 0})</a>`
      : "";
    return `
      <tr style="background:${completada ? "#f2fbf4" : "#fdf4f4"}">
        <td style="text-align:center; width:56px; font-size:1.3rem">${completada ? "✅" : "⬜"}</td>
        <td style="font-weight:600">${escapeHTML(enc.title)}</td>
        <td><code>${escapeHTML(enc.codigo || "—")}</code></td>
        <td style="text-align:center">${completada ? "Completada" : "Pendiente"}</td>
        <td>${fecha}</td>
        <td>
          <a href="${urlEncuesta}" target="_blank" rel="noopener" style="font-weight:600">${completada ? "Volver a responder" : "Completar ahora →"}</a>
          ${completada ? ` · <a href="${jsonUrl}" target="_blank" rel="noopener">ver JSON</a>` : ""}
        </td>
      </tr>`;
  }).join("");

  const errorNota = tasaError ? `<p style="color:#b21d21; font-size:.9rem; margin:6px 0">
    ⚠ GitHub API devolvió un error (${tasaError.status || tasaError.message}). Puede ser límite de
    tasa (60/h sin token) o un repo movido a privado.</p>` : "";

  visor.hidden = false;
  visor.innerHTML = `
    <div style="padding: 26px; background: #fff; max-width: 1060px; margin: 0 auto; box-shadow: 0 2px 8px rgba(0,0,0,.07); font-family: Arial, sans-serif;">
      <h2 style="margin:0 0 2px; color:#1f4e79; font-size: 1.1rem;">PLAN DE RESPUESTAS — ENCUESTAS SG-SST</h2>
      <div style="font-size:.85rem; color:#555;">Empresa: <strong>${escapeHTML(empresa.EMPRESA)}</strong> &nbsp;·&nbsp; ${hechas} de ${encuestas.length} encuestas completadas</div>
      <hr style="border:none; border-top:2px solid #1f4e79; margin:10px 0">
      ${errorNota}
      <table style="width:100%; border-collapse:collapse; font-size:.92rem">
        <thead>
          <tr style="background:#f2f2f2">
            <th style="border:1px solid #000; padding:6px">Estado</th>
            <th style="border:1px solid #000; padding:6px">Encuesta</th>
            <th style="border:1px solid #000; padding:6px">Código</th>
            <th style="border:1px solid #000; padding:6px">Situación</th>
            <th style="border:1px solid #000; padding:6px">Última realización</th>
            <th style="border:1px solid #000; padding:6px">Acciones</th>
          </tr>
        </thead>
        <tbody>${filas}</tbody>
      </table>
      <p style="font-size:.8rem; color:#666; margin-top:10px">
        Los enlaces "Completar ahora" abren la encuesta pública ya con el nombre de la empresa prellenado.
        Las respuestas quedan en el repo GitHub <code>juandprada/respuestasencuestas</code> vía el Worker.
      </p>
    </div>`;
  $("#salida").innerHTML = "";
  document.title = `Plan de Respuestas - ${empresa.EMPRESA}`;
  setEstado("");
}

async function fetchJSON(url) {
  return JSON.parse(await fetchText(url));
}

function poblarSelects() {
  const selF = $("#sel-formato");
  const selE = $("#sel-empresa");

  // Formatos agrupados por categoria.
  const grupos = {};
  for (const f of FORMATOS) (grupos[f.categoria] ||= []).push(f);
  selF.innerHTML = "";
  for (const cat of Object.keys(grupos).sort()) {
    const og = document.createElement("optgroup");
    og.label = cat;
    for (const f of grupos[cat]) {
      const o = document.createElement("option");
      o.value = f.id;
      o.textContent = f.nombre;
      og.appendChild(o);
    }
    selF.appendChild(og);
  }

  selE.innerHTML = "";
  for (const e of EMPRESAS) {
    const o = document.createElement("option");
    o.value = e._id;
    o.textContent = e.EMPRESA;
    selE.appendChild(o);
  }

  // Capacitaciones agrupadas por área, igual que en la matriz de capacitación.
  const selC = $("#sel-capacitacion");
  if (selC && CAPACITACIONES) {
    selC.innerHTML = "";
    for (const grupo of CAPACITACIONES.grupos) {
      const og = document.createElement("optgroup");
      og.label = grupo.area;
      for (const c of grupo.items) {
        const o = document.createElement("option");
        o.value = c.id;
        o.textContent = c.nombre;
        og.appendChild(o);
      }
      selC.appendChild(og);
    }
  }

  // Inspecciones disponibles.
  const selI = $("#sel-inspeccion");
  if (selI && INSPECCIONES) {
    selI.innerHTML = "";
    for (const insp of INSPECCIONES.items) {
      const o = document.createElement("option");
      o.value = insp.id;
      o.textContent = insp.nombre;
      selI.appendChild(o);
    }
  }

  // Meses disponibles para Cuenta de Cobro.
  const selM = $("#sel-mes-cc");
  if (selM) {
    selM.innerHTML = "";
    MESES_ES.forEach((nombre, idx) => {
      const o = document.createElement("option");
      o.value = String(idx + 1);
      o.textContent = tituloCase(nombre);
      selM.appendChild(o);
    });
    const mesActual = new Date().getMonth() + 1;
    selM.value = String(mesActual);
  }

  // Valores mensuales para Cuenta de Cobro.
  const selV = $("#sel-valor-cc");
  if (selV) {
    selV.innerHTML = "";
    const comunes = valoresComunesCuentaCobro();
    for (const val of comunes) {
      const o = document.createElement("option");
      o.value = String(val);
      o.textContent = "$" + val.toLocaleString("es-CO");
      selV.appendChild(o);
    }
  }

  // Años disponibles: el actual y el anterior (hoy: 2026 y 2025). Muchos documentos
  // se firman para el año en curso, pero a veces hay que reponer los del año pasado.
  const selA = $("#sel-anio");
  const actual = new Date().getFullYear();
  selA.innerHTML = "";
  for (const a of [actual, actual - 1, actual - 2]) {
    const o = document.createElement("option");
    o.value = String(a);
    o.textContent = String(a);
    selA.appendChild(o);
  }
  selA.value = String(actual); // por defecto, el año en curso
}

async function generar() {
  try {
    setEstado("Generando…");

    const formato = FORMATOS.find((f) => f.id === $("#sel-formato").value);
    const empresa = EMPRESAS.find((e) => e._id === $("#sel-empresa").value);
    if (!formato || !empresa) {
      setEstado("Selecciona formato y empresa.", true);
      return;
    }
    const anio = Number($("#sel-anio").value) || new Date().getFullYear();

    if (formato.estatico) {
      setEstado(`Listo para descargar: ${formato.nombre}.`);
      $("#salida").innerHTML = "";
      const visor = $("#visor");
      visor.hidden = false;
      
      let botones = "";
      const estaticos = Array.isArray(formato.estatico) ? formato.estatico : [{ nombre: formato.nombre, url: formato.estatico }];
      
      for (const est of estaticos) {
        botones += `
          <a href="${est.url}" download="${est.nombre}.pdf" style="text-decoration:none;text-align:center;color:#333;padding:40px;border-radius:10px;background:#fff;box-shadow:0 4px 6px rgba(0,0,0,0.1);transition:transform 0.2s;margin:15px;min-width:200px;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
            <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom:15px;color:#d9534f;">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            <br>
            <strong style="font-size:18px;">Descargar<br>${est.nombre}</strong>
          </a>
        `;
      }

      visor.innerHTML = `
        <div style="display:flex;flex-direction:row;flex-wrap:wrap;align-items:center;justify-content:center;height:100%;background:#f9f9f9;border:1px dashed #ccc;">
          ${botones}
        </div>`;
      return;
    }

    // FORMATO ESPECIAL: Plan de Respuestas (encuestas internas SG-SST).
    // No genera PDF; pinta el seguimiento de encuestas de la empresa: cuáles están
    // realizadas (fecha y archivo) y links para completar las faltantes.
    if (formato.id === "plan-respuestas") {
      await renderPlanRespuestas(empresa);
      return;
    }

    // Panel de ajustes: solo se muestra (y solo pregunta) en el formato presupuesto.
    renderPanelPresupuesto(formato, empresa);
    // Selector de capacitación: solo en el acta de asistencia.
    renderSelectorCapacitacion(formato);
    // Selector de cargo: solo en la matriz IPVER.
    renderSelectorCargoIPVER(formato);
    // Selector de inspección: solo en el formato de inspección.
    renderSelectorInspeccion(formato);
    // Selector y valor de Cuenta de Cobro: solo en cuenta-de-cobro.
    renderSelectorCuentaCobro(formato, empresa);
    // Panel de extras Matriz Legal: solo en matriz-legal.
    const panelMatrizLegal = $("#panel-matriz-legal");
    if (panelMatrizLegal) panelMatrizLegal.hidden = formato.id !== "matriz-legal";
    
    const cap = capacitacionElegida();
    const insp = inspeccionElegida();

    const mesCC = Number(($("#sel-mes-cc") && $("#sel-mes-cc").value) || (new Date().getMonth() + 1)) || 1;
    const valorCC = valorCuentaCobro(empresa);
    const diaNumCC = ultimoDiaHabil(anio, mesCC);
    const diaTxtCC = diaATexto(diaNumCC);
    const mesNombreCC = MESES_ES[mesCC - 1] || "ENERO";
    const numeroCC = `${String(mesCC).padStart(2, "0")}-${String(anio).slice(-2)}`;
    const consultora = (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.consultora) || {
      nombre: "KAREN LIZETH BENSUR MURIEL",
      nit: "1.143.973.774-7",
      concepto: "Asesoría en el desarrollo del SG-SST",
      ciudad: "CALI",
    };

    let archivoTpl = formato.archivo;
    let tituloFmt = formato.titulo || formato.nombre;
    if (formato.id === FORMATO_INSPECCION && insp) {
      archivoTpl = insp.archivo;
      tituloFmt = insp.nombre;
    }
    const cuerpoTpl = await fetchText(`plantillas/${archivoTpl}`);

    // Contexto = datos de empresa + metadatos del formato + logo (HTML).
    const ctx = {
      ...empresa,
      // Nombre del representante en Título (viene en MAYÚSCULAS desde el Excel),
      // para unificar con la firma de la consultora que va en Título.
      REPRESENTANTE_LEGAL: tituloCase(empresa.REPRESENTANTE_LEGAL || ""),
      TITULO: tituloFmt,
      CODIGO: formato.codigo || "",
      VERSION: formato.version || "",
      // Año del documento: lo elige el usuario en el selector (encabezado y textos
      // tipo "para los años {{ANIO}} y {{ANIO_SIGUIENTE}}").
      ANIO: String(anio),
      ANIO_SIGUIENTE: String(anio + 1),
      LOGO: logoHTML(empresa),
      // Firma de la consultora (Karen Lizeth Bensur): se inserta automáticamente.
      FIRMA_CONSULTORA:
        '<img class="firma-img" src="assets/firma-karen.png" alt="Firma consultora">',
      // Tabla del presupuesto: se calcula por empresa (ver data/presupuesto.json).
      TABLA_PRESUPUESTO: tablaPresupuestoHTML(empresa),
      // Tabla de la matriz IPVER: se calcula por empresa y cargo (ver data/gtc45.json).
      TABLA_IPVER: tablaIPVERHTML(empresa),
      // Página 1 del acta de asistencia (ver data/capacitaciones.json).
      CAPACITACION: cap ? cap.nombre : "",
      CAPACITACION_OBJETIVO: cap ? cap.objetivo : "",
      CAPACITACION_PUNTOS: puntosCapacitacionHTML(cap),
      // Página 1 del formato de inspección.
      INSPECCION_NOMBRE: insp ? insp.nombre : "",
      INSPECCION_OBJETIVO: insp ? insp.objetivo : "",
      INSPECCION_PUNTOS: puntosInspeccionHTML(insp),
      // Tokens de Cuenta de Cobro.
      NUMERO_CUENTA: numeroCC,
      CONSULTORA_NOMBRE: consultora.nombre,
      CONSULTORA_NIT: consultora.nit,
      VALOR_LETRAS: numeroALetras(valorCC),
      VALOR_PESOS: "$" + Math.round(valorCC).toLocaleString("es-CO"),
      CONCEPTO: consultora.concepto,
      MES_NOMBRE: mesNombreCC,
      DIA_TEXTO: diaTxtCC,
      DIA_NUMERO: String(diaNumCC),
      CIUDAD: consultora.ciudad,
    };

    const raw = [
      "LOGO", "FIRMA_CONSULTORA", "TABLA_PRESUPUESTO", "TABLA_IPVER",
      "CAPACITACION_PUNTOS", "INSPECCION_PUNTOS", "TITULO"
    ];
    const encabezado = fillTokens(ENCABEZADO_TPL, ctx, raw);
    const cuerpo = fillTokens(cuerpoTpl, ctx, raw);

    // Orientación de página (vertical por defecto / horizontal si aplica).
    let orientacion = formato.orientacion || "vertical";
    if (formato.id === FORMATO_INSPECCION && insp && insp.orientacion) {
      orientacion = insp.orientacion;
    }
    const horizontal = orientacion === "horizontal";
    // Los formatos de inspección arrancan más arriba: margen superior 0.4in
    // (el resto de la plataforma conserva 0.6in uniforme). Aplica a la impresión
    // del navegador (@page aquí) y al PDF (generarPDF lee #salida[data-mtop]).
    const mtop = formato.id === FORMATO_INSPECCION ? "0.4in" : "0.6in";
    document.getElementById("page-orient").textContent =
      `@page { size: Letter ${horizontal ? "landscape" : "portrait"}; margin: ${mtop} 0.6in 0.6in; }`;
    $("#salida").classList.toggle("horizontal", horizontal);
    $("#salida").dataset.mtop = mtop;

    // Ámbito por formato: la clase doc--<id> permite ajustes de CSS específicos de
    // un formato sin afectar a los demás (ver styles.css). Las variantes de tabla
    // (.tabla-firmas, .tabla-form…) cubren lo común; esto es el escape para excepciones.
    const scope = "doc--" + String(formato.id).replace(/[^a-z0-9-]/gi, "");
    $("#salida").innerHTML =
      `<article class="doc ${scope}">${encabezado}` +
      `<div class="doc-body">${cuerpo}</div></article>`;

    if (formato.id === "matriz-legal") {
      const docOut = $("#salida");
      if (!$("#ml-chk-alturas").checked) {
        docOut.querySelectorAll('tr[data-extra="alturas"]').forEach(el => el.remove());
      }
      if (!$("#ml-chk-confinados").checked) {
        docOut.querySelectorAll('tr[data-extra="confinados"]').forEach(el => el.remove());
      }
      if (!$("#ml-chk-pesv").checked) {
        docOut.querySelectorAll('tr[data-extra="pesv"]').forEach(el => el.remove());
      }
    }

    // Página por defecto en cada encabezado; al descargar se corrige al total real.
    const celdasPag = $("#salida").querySelectorAll(".dh-pagina");
    celdasPag.forEach((c, i) => (c.textContent = i + 1 + " de " + celdasPag.length));

    document.title = `${formato.nombre} - ${empresa.EMPRESA}`;
    // Genera el PDF y lo muestra en el visor.
    await generarPDF();
  } catch (err) {
    console.error(err);
    setEstado(err.message || "Error al generar.", true);
  }
}

let ultimoBlobUrl = null;

// Genera el PDF (sin diálogo de impresión, para evitar el encabezado nativo del
// navegador en móvil) y lo muestra en un visor embebido con controles nativos.
// Espera a que la fuente del documento esté cargada. html2canvas CAPTURA LA PANTALLA:
// si la woff2 aún no llegó, el PDF sale dibujado con la fuente de reemplazo (más
// angosta) y con otros cortes de línea que el Word. Solo la primera vez tarda algo;
// después está en la caché del navegador.
async function esperarFuente() {
  if (!document.fonts) return;
  try {
    await Promise.all([
      document.fonts.load('400 16px "Doc Sans"'),
      document.fonts.load('700 16px "Doc Sans"'),
    ]);
    await document.fonts.ready;
  } catch (err) {
    console.warn("No se pudo confirmar la carga de la fuente del documento.", err);
  }
}

async function generarPDF() {
  const doc = $("#salida .doc");
  if (!doc) {
    setEstado("Primero genera la vista previa.", true);
    return;
  }
  await esperarFuente();
  const horizontal = $("#salida").classList.contains("horizontal");
  // Ancho = área de contenido (página menos 0.6in de margen a cada lado):
  // horizontal 11in-1.2in=9.8in, vertical 8.5in-1.2in=7.3in (a 96dpi).
  const anchoPx = horizontal ? 940 : 700;
  const nombre =
    (document.title || "documento").replace(/[^\wáéíóúñ .-]+/gi, "").trim() ||
    "documento";

  // Margen superior reducido (0.4in) SOLO en los formatos de inspección; lo fija
  // renderFormato en #salida[data-mtop]. El array de html2pdf es [arriba, izq,
  // abajo, der] — verificado en el vendor: addImage usa margin[1]=X, margin[0]=Y.
  const mtop = parseFloat($("#salida").dataset.mtop || "0.6") || 0.6;
  const opt = {
    margin: [mtop, 0.6, 0.6, 0.6], // (in): arriba 0.4 solo inspecciones; resto 0.6
    filename: nombre + ".pdf",
    image: { type: "jpeg", quality: 0.98 },
    // Sin windowWidth/width: dependían del devicePixelRatio (escalado de Windows)
    // y desplazaban el PDF. El ancho lo fija doc.style.width (ver abajo).
    html2canvas: {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
    },
    jsPDF: {
      unit: "in",
      format: "letter",
      orientation: horizontal ? "landscape" : "portrait",
    },
    pagebreak: {
      mode: ["css", "legacy"],
      before: [".salto-pagina", ".pb-antes"], // corte forzado en bordes limpios
      avoid: ["tr", ".doc-header"], // no partir filas ni el encabezado repetido
    },
  };

  setEstado("Generando PDF…");
  window.scrollTo(0, 0); // evita que html2canvas capture con desplazamiento vertical
  // Oculta la barra de scroll durante la captura. Con página alta (el visor de 85vh),
  // Windows muestra una barra clásica de ~17px; html2canvas usa clientWidth (que la
  // resta) y desplaza el contenido a la derecha, cortando el borde derecho. En Android
  // (scrollbars superpuestas) no pasa. overflow:hidden la quita solo durante el render.
  const overflowPrevio = document.documentElement.style.overflow;
  document.documentElement.style.overflow = "hidden";
  const anchoPrevio = doc.style.width;
  doc.style.width = anchoPx + "px"; // fija el ancho durante la captura
  const celdasPagina = doc.querySelectorAll(".dh-pagina");
  try {
    // Paso 1: contar las páginas reales y numerar cada encabezado ("i de N").
    if (celdasPagina.length) {
      const pdf = await html2pdf().set(opt).from(doc).toPdf().get("pdf");
      const N = pdf.internal.getNumberOfPages();
      celdasPagina.forEach((c, i) => (c.textContent = i + 1 + " de " + N));
    }
    // Paso 2: generar el PDF como blob y mostrarlo en el visor.
    const url = await html2pdf().set(opt).from(doc).outputPdf("bloburl");
    if (ultimoBlobUrl) URL.revokeObjectURL(ultimoBlobUrl);
    ultimoBlobUrl = url;

    const visor = $("#visor");
    visor.hidden = false;
    visor.innerHTML =
      `<iframe class="visor-frame" title="Vista del PDF" src="${url}"></iframe>`;
    // Desplazamiento INSTANTÁNEO, no "smooth": la animación seguía corriendo cuando
    // empezaba la siguiente generación y movía la página DESPUÉS del scrollTo(0,0) de
    // abajo, así que html2canvas capturaba desplazado ~37px. Eso metía una banda blanca
    // arriba de cada página y recortaba el final de la última (se veía en la página de
    // firmas del acta de asistencia, que es la más alta).
    visor.scrollIntoView({ behavior: "auto", block: "start" });
    setEstado("PDF listo: usa el visor para hacer zoom, imprimir o descargar.");
  } catch (err) {
    console.error(err);
    setEstado("No se pudo generar el PDF: " + (err.message || err), true);
  } finally {
    doc.style.width = anchoPrevio;
    document.documentElement.style.overflow = overflowPrevio;
  }
}

// Serializa las generaciones: una sola a la vez (dos html2canvas en paralelo sobre
// el mismo #salida se corrompen y gana la que termina última, no la seleccionada).
// Si el usuario cambia de selección durante un render, al terminar se repite con lo
// último elegido.
let genSeq = 0;
let genRunning = false;
async function solicitarGeneracion() {
  genSeq++;
  if (genRunning) return; // la que corre ya tomará el genSeq más reciente al terminar
  genRunning = true;
  try {
    let last;
    do {
      last = genSeq;
      await generar(); // usa los valores ACTUALES de los selects
    } while (genSeq !== last); // cambió durante el render -> repetir con lo último
  } finally {
    genRunning = false;
  }
}

async function init() {
  try {
    setEstado("Cargando datos…");
    [EMPRESAS, FORMATOS, ENCABEZADO_TPL, PRESUPUESTO, CAPACITACIONES, GTC45_DATA, INSPECCIONES, CUENTAS_COBRO_DATA] =
      await Promise.all([
        fetchJSON("data/empresas.json"),
        fetchJSON("plantillas/manifest.json"),
        fetchText("partials/encabezado.html"),
        fetchJSON("data/presupuesto.json"),
        fetchJSON("data/capacitaciones.json"),
        fetchJSON("data/gtc45.json"),
        fetchJSON("data/inspecciones.json"),
        fetchJSON("data/cuentas_cobro.json"),
      ]);
    cargarAjustes(); // ajustes de presupuesto guardados en este navegador
    poblarSelects();
    conectarPanelPresupuesto();
    setEstado("");
    // Auto-genera el PDF al elegir formato o empresa (sin botones), serializado.
    $("#sel-formato").addEventListener("change", () => {
      const formato = FORMATOS.find((f) => f.id === $("#sel-formato").value);
      const empresa = EMPRESAS.find((e) => e._id === $("#sel-empresa").value);
      if (formato && formato.id === FORMATO_CUENTA_COBRO && empresa) {
        const selValor = $("#sel-valor-cc");
        if (selValor) {
          const tarifas = (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifas_por_empresa) || {};
          const val = tarifas[empresa._id] || (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifa_default) || 250000;
          selValor.value = String(val);
          selValor.dataset.empresa = empresa._id;
        }
      }
      solicitarGeneracion();
    });
    $("#sel-empresa").addEventListener("change", () => {
      const formato = FORMATOS.find((f) => f.id === $("#sel-formato").value);
      const empresa = EMPRESAS.find((e) => e._id === $("#sel-empresa").value);
      if (formato && formato.id === FORMATO_CUENTA_COBRO && empresa) {
        const selValor = $("#sel-valor-cc");
        if (selValor) {
          const tarifas = (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifas_por_empresa) || {};
          const val = tarifas[empresa._id] || (CUENTAS_COBRO_DATA && CUENTAS_COBRO_DATA.tarifa_default) || 250000;
          selValor.value = String(val);
          selValor.dataset.empresa = empresa._id;
        }
      }
      solicitarGeneracion();
    });
    $("#sel-anio").addEventListener("change", solicitarGeneracion);
    $("#sel-capacitacion").addEventListener("change", solicitarGeneracion);
    $("#sel-cargo-ipver").addEventListener("change", solicitarGeneracion);
    $("#sel-inspeccion").addEventListener("change", solicitarGeneracion);
    $("#sel-mes-cc").addEventListener("change", solicitarGeneracion);
    $("#sel-valor-cc").addEventListener("change", solicitarGeneracion);
    $("#ml-chk-alturas").addEventListener("change", solicitarGeneracion);
    $("#ml-chk-confinados").addEventListener("change", solicitarGeneracion);
    $("#ml-chk-pesv").addEventListener("change", solicitarGeneracion);
    solicitarGeneracion(); // genera el primer documento con la selección por defecto
  } catch (err) {
    console.error(err);
    setEstado(
      "No se pudieron cargar los datos. Si abriste el archivo directamente, " +
        "usa un servidor local (ver README).",
      true
    );
  }
}

init();
