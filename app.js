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
    $("#btn-pdf").disabled = true;

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

    $("#salida").innerHTML =
      `<article class="doc">${encabezado}` +
      `<div class="doc-body">${cuerpo}</div></article>`;

    document.title = `${formato.nombre} - ${empresa.EMPRESA}`;
    $("#btn-pdf").disabled = false;
    setEstado(`Vista previa lista: ${formato.nombre} — ${empresa.EMPRESA}.`);
    $("#salida").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) {
    console.error(err);
    setEstado(err.message || "Error al generar.", true);
  }
}

// Descarga el PDF sin diálogo de impresión (evita el encabezado nativo del
// navegador en móvil). Rasteriza con html2pdf: ancho fijo para que no se
// deforme en pantallas pequeñas.
async function descargarPDF() {
  const doc = $("#salida .doc");
  if (!doc) {
    setEstado("Primero genera la vista previa.", true);
    return;
  }
  const horizontal = $("#salida").classList.contains("horizontal");
  const anchoPx = horizontal ? 1056 : 816; // 11in / 8.5in a 96dpi
  const nombre =
    (document.title || "documento").replace(/[^\wáéíóúñ .-]+/gi, "").trim() ||
    "documento";

  setEstado("Generando PDF…");
  $("#btn-pdf").disabled = true;
  const anchoPrevio = doc.style.width;
  doc.style.width = anchoPx + "px"; // fija el ancho durante la captura
  try {
    await html2pdf()
      .set({
        margin: 0,
        filename: nombre + ".pdf",
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: {
          scale: 2,
          useCORS: true,
          backgroundColor: "#ffffff",
          windowWidth: anchoPx,
          width: anchoPx,
        },
        jsPDF: {
          unit: "in",
          format: "letter",
          orientation: horizontal ? "landscape" : "portrait",
        },
        pagebreak: { mode: ["css", "legacy"], before: ".salto-pagina", avoid: "tr" },
      })
      .from(doc)
      .save();
    setEstado("PDF descargado.");
  } catch (err) {
    console.error(err);
    setEstado("No se pudo generar el PDF: " + (err.message || err), true);
  } finally {
    doc.style.width = anchoPrevio;
    $("#btn-pdf").disabled = false;
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
    $("#btn-generar").addEventListener("click", generar);
    $("#btn-pdf").addEventListener("click", descargarPDF);
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
