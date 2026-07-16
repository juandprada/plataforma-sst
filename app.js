"use strict";

// Estado en memoria.
let EMPRESAS = [];
let FORMATOS = [];
let ENCABEZADO_TPL = "";

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
      ANIO: formato.anio || String(new Date().getFullYear()),
      ANIO_SIGUIENTE: String(
        Number(formato.anio || new Date().getFullYear()) + 1
      ),
      LOGO: logoHTML(empresa),
      // Firma de la consultora (Karen Lizeth Bensur): se inserta automáticamente.
      FIRMA_CONSULTORA:
        '<img class="firma-img" src="assets/firma-karen.png" alt="Firma consultora">',
    };

    const raw = ["LOGO", "FIRMA_CONSULTORA"];
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
async function generarPDF() {
  const doc = $("#salida .doc");
  if (!doc) {
    setEstado("Primero genera la vista previa.", true);
    return;
  }
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
    visor.scrollIntoView({ behavior: "smooth", block: "start" });
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
    [EMPRESAS, FORMATOS, ENCABEZADO_TPL] = await Promise.all([
      fetchJSON("data/empresas.json"),
      fetchJSON("plantillas/manifest.json"),
      fetchText("partials/encabezado.html"),
    ]);
    poblarSelects();
    setEstado("");
    // Auto-genera el PDF al elegir formato o empresa (sin botones), serializado.
    $("#sel-formato").addEventListener("change", solicitarGeneracion);
    $("#sel-empresa").addEventListener("change", solicitarGeneracion);
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
