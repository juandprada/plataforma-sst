"use strict";

// Estado en memoria.
let EMPRESAS = [];
let FORMATOS = [];
let ENCABEZADO_TPL = "";
let PRESUPUESTO = null; // data/presupuesto.json (parámetros por empresa + catálogo)
let CAPACITACIONES = null; // data/capacitaciones.json (contenido de la página 1 del acta)

// Formato cuya página 1 depende de la capacitación elegida (ver CAPACITACIONES).
const FORMATO_ASISTENCIA = "asistencia-a-capacitacion";

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

async function fetchText(url) {
  const r = await fetch(url, { cache: "no-cache" });
  if (!r.ok) throw new Error(`No se pudo cargar ${url} (${r.status})`);
  return r.text();
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

  // Años disponibles: el actual y el anterior (hoy: 2026 y 2025). Muchos documentos
  // se firman para el año en curso, pero a veces hay que reponer los del año pasado.
  const selA = $("#sel-anio");
  const actual = new Date().getFullYear();
  selA.innerHTML = "";
  for (const a of [actual, actual - 1]) {
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

    // Panel de ajustes: solo se muestra (y solo pregunta) en el formato presupuesto.
    renderPanelPresupuesto(formato, empresa);
    // Selector de capacitación: solo en el acta de asistencia.
    renderSelectorCapacitacion(formato);
    const cap = capacitacionElegida();

    const cuerpoTpl = await fetchText(`plantillas/${formato.archivo}`);

    // Contexto = datos de empresa + metadatos del formato + logo (HTML).
    const ctx = {
      ...empresa,
      // Nombre del representante en Título (viene en MAYÚSCULAS desde el Excel),
      // para unificar con la firma de la consultora que va en Título.
      REPRESENTANTE_LEGAL: tituloCase(empresa.REPRESENTANTE_LEGAL || ""),
      TITULO: formato.titulo || formato.nombre,
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
      // Página 1 del acta de asistencia (ver data/capacitaciones.json).
      CAPACITACION: cap ? cap.nombre : "",
      CAPACITACION_OBJETIVO: cap ? cap.objetivo : "",
      CAPACITACION_PUNTOS: puntosCapacitacionHTML(cap),
    };

    const raw = [
      "LOGO", "FIRMA_CONSULTORA", "TABLA_PRESUPUESTO", "CAPACITACION_PUNTOS",
    ];
    const encabezado = fillTokens(ENCABEZADO_TPL, ctx, raw);
    const cuerpo = fillTokens(cuerpoTpl, ctx, raw);

    // Orientación de página (vertical por defecto / horizontal si aplica).
    const horizontal = (formato.orientacion || "vertical") === "horizontal";
    document.getElementById("page-orient").textContent =
      `@page { size: Letter ${horizontal ? "landscape" : "portrait"}; margin: 0.6in; }`;
    $("#salida").classList.toggle("horizontal", horizontal);

    // Ámbito por formato: la clase doc--<id> permite ajustes de CSS específicos de
    // un formato sin afectar a los demás (ver styles.css). Las variantes de tabla
    // (.tabla-firmas, .tabla-form…) cubren lo común; esto es el escape para excepciones.
    const scope = "doc--" + String(formato.id).replace(/[^a-z0-9-]/gi, "");
    $("#salida").innerHTML =
      `<article class="doc ${scope}">${encabezado}` +
      `<div class="doc-body">${cuerpo}</div></article>`;

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

  const opt = {
    margin: 0.6, // margen uniforme (in) en TODAS las páginas y lados
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
    [EMPRESAS, FORMATOS, ENCABEZADO_TPL, PRESUPUESTO, CAPACITACIONES] =
      await Promise.all([
        fetchJSON("data/empresas.json"),
        fetchJSON("plantillas/manifest.json"),
        fetchText("partials/encabezado.html"),
        fetchJSON("data/presupuesto.json"),
        fetchJSON("data/capacitaciones.json"),
      ]);
    cargarAjustes(); // ajustes de presupuesto guardados en este navegador
    poblarSelects();
    conectarPanelPresupuesto();
    setEstado("");
    // Auto-genera el PDF al elegir formato o empresa (sin botones), serializado.
    $("#sel-formato").addEventListener("change", solicitarGeneracion);
    $("#sel-empresa").addEventListener("change", solicitarGeneracion);
    $("#sel-anio").addEventListener("change", solicitarGeneracion);
    $("#sel-capacitacion").addEventListener("change", solicitarGeneracion);
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
