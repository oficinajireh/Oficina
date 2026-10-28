/* =====================================================================
   deposito.js — funciones del DEPÓSITO (ventas mayoristas)
   ---------------------------------------------------------------------
   El depósito usa su propio backend (mismo code.gs, MODO_NEGOCIO =
   DEPOSITO) y el mismo panel. Este archivo agrega lo que es propio:
   - Dólar del depósito fijado a mano (widget del dólar + Configuración)
   - Precios en dólares de los productos
   - Cotizaciones (presupuestos) A4, boletas de entrega A4
   - Ficha del cliente completa
   - Traspasos de mercadería al local (y su recepción en el local)
   Se apoya en lo que ya define admin.js (API_URL, fetchAPI, toast,
   confirmarAccion, escapeHtml, configNegocioCache, esModoDeposito...).
===================================================================== */

/* ---------- utilidades ---------- */

/** POST al backend con la sesión del panel (el token lo agrega veek-auth.js) */
async function depPost(accion, datos, config) {
  const r = await fetchAPI(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: accion, ...(datos || {}) })
  }, config || {});
  return r.json();
}

/** GET al backend (lecturas) */
async function depGet(accion, params, config) {
  const qs = new URLSearchParams({ action: accion, ...(params || {}) });
  const r = await fetchAPI(API_URL + "?" + qs.toString(), {}, config || {});
  return r.json();
}

const depPesos = (n) => "$" + Math.round(Number(n) || 0).toLocaleString("es-AR");
const depUsd = (n) => "US$ " + (Number(n) || 0).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const depEsAdmin = () => (typeof obtenerRolActual === "function" ? obtenerRolActual() : "admin") === "admin";

function depFechaHora(iso) {
  const f = new Date(iso);
  if (!iso || isNaN(f)) return "";
  return f.toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "2-digit", hour: "2-digit", minute: "2-digit" });
}

/** Crea (una sola vez) un modal con la misma estética que los del panel */
function depModal(id, titulo, cuerpoHtml, accionesHtml, ancho) {
  let bd = document.getElementById(id);
  if (!bd) {
    bd = document.createElement("div");
    bd.className = "product-modal-backdrop";
    bd.id = id;
    bd.innerHTML = `<div class="product-modal" style="max-width:${ancho || 460}px;width:96vw;">
      <div class="product-modal-header"><h6></h6><button type="button" data-cerrar>✕</button></div>
      <div class="product-modal-body"></div>
      <div class="product-modal-actions"></div></div>`;
    document.body.appendChild(bd);
    bd.addEventListener("click", (e) => {
      if (e.target.closest("[data-cerrar]") || (e.target === bd && !bd.dataset.fijo)) bd.classList.remove("show");
    });
  }
  bd.querySelector("h6").textContent = titulo;
  bd.querySelector(".product-modal-body").innerHTML = cuerpoHtml;
  bd.querySelector(".product-modal-actions").innerHTML = accionesHtml || "";
  bd.classList.add("show");
  return bd;
}

/* =====================================================================
   DÓLAR DEL DEPÓSITO
   Se fija a mano. Se ve en el widget del dólar (POS e Inicio) junto a
   las cotizaciones del mercado, y desde ahí se puede cambiar o tomar
   el valor de una de ellas con un toque.
===================================================================== */
let depDolar = null; // { valor, fecha, usuario }

/** Dólar a usar en el depósito (0 si no está fijado) */
function dolarDelDeposito() {
  return depDolar && depDolar.valor > 0 ? Number(depDolar.valor) : 0;
}

/** Lo toma de la configuración ya cargada (sin pedir nada al servidor) */
function depDolarDesdeConfig(cfg) {
  if (cfg && cfg.dolarDeposito && cfg.dolarDeposito.valor > 0) depDolar = cfg.dolarDeposito;
  depDolarAplicar();
}

async function depDolarRefrescar() {
  if (!esModoDeposito()) return;
  try {
    const d = await depGet("dolarDeposito", {}, { timeoutMs: 15000, reintentos: 1 });
    if (d && d.success) { depDolar = { valor: d.valor, fecha: d.fecha, usuario: d.usuario, historial: d.historial }; depDolarAplicar(); }
    return d;
  } catch (e) { return null; }
}

/** Refleja el dólar en el widget y en los campos "Tipo de cambio" */
function depDolarAplicar() {
  if (typeof dolarBurbujaRender === "function") dolarBurbujaRender();
  if (typeof depRepintarPrecios === "function" && esModoDeposito()) depRepintarPrecios();
  depDolarTarjetaConfig();
  if (!esModoDeposito() || !(dolarDelDeposito() > 0)) return;
  // El tipo de cambio del ingreso de mercadería sale del dólar del depósito
  if (typeof ingresoTC !== "undefined") {
    ingresoTC = { ...ingresoTC, valor: dolarDelDeposito(), tipo: "deposito", manual: false, fallo: false, fecha: depDolar.fecha || "" };
    if (typeof tcRender === "function") tcRender();
    if (typeof tcAplicarACampos === "function") tcAplicarACampos(false);
  }
}

/** Filas del widget en el depósito: primero el del depósito, después las del mercado con "usar" */
function depDolarFilasBurbuja(filasMercado) {
  const admin = depEsAdmin();
  const v = dolarDelDeposito();
  const fila = `<div class="db-fila db-dep"><span>Depósito</span><b>${v ? "$" + v.toLocaleString("es-AR", { maximumFractionDigits: 2 }) : "sin fijar"}</b>` +
    (admin ? `<button type="button" class="db-btn" title="Cambiar el dólar del depósito" onclick="event.stopPropagation(); abrirDolarDeposito()">✏️</button>` : "") + `</div>` +
    (depDolar && depDolar.fecha ? `<div class="db-sub">fijado ${escapeHtml(_dbFechaCorta(new Date(depDolar.fecha)))}</div>` : "");
  const mercado = (filasMercado || []).map(f => f.html.replace("</div>", (admin ? `<button type="button" class="db-btn db-usar" title="Usar este valor como dólar del depósito" onclick="event.stopPropagation(); usarDolarMercado('${f.tipo}', ${f.valor})">usar</button>` : "") + "</div>"));
  return fila + (mercado.length ? '<div class="db-sep"></div>' + mercado.join("") : "");
}

function usarDolarMercado(tipo, valor) {
  const nombre = (typeof TC_TIPOS !== "undefined" && TC_TIPOS[tipo]) || tipo;
  confirmarAccion(`¿Fijar el dólar del depósito en $${Number(valor).toLocaleString("es-AR", { maximumFractionDigits: 2 })} (${nombre})? Los productos con precio en dólares se recalculan con este valor.`,
    () => guardarDolarDeposito(valor), "Dólar del depósito");
}

async function guardarDolarDeposito(valor) {
  try {
    const r = await depPost("guardarDolarDeposito", { valor: String(valor) });
    if (!r.success) { toast(r.message || "No se pudo guardar el dólar", "error"); return false; }
    depDolar = { valor: r.valor, fecha: r.fecha, usuario: sessionStorage.getItem("usuarioLogueado") || "" };
    if (typeof configNegocioCache !== "undefined") configNegocioCache.dolarDeposito = depDolar;
    depDolarAplicar();
    toast(`Dólar del depósito: $${Number(r.valor).toLocaleString("es-AR", { maximumFractionDigits: 2 })}` +
      (r.productosActualizados ? ` · ${r.productosActualizados} producto(s) en dólares actualizados` : ""), "success");
    // los precios en pesos de los productos en dólares cambiaron: que el POS los tome
    if (r.productosActualizados && typeof window.refrescarCatalogoPOS === "function") window.refrescarCatalogoPOS();
    return true;
  } catch (e) {
    toast("No se pudo guardar el dólar: " + e.message, "error");
    return false;
  }
}

/** Modal para cambiar el dólar del depósito (con su historial) */
async function abrirDolarDeposito() {
  if (!depEsAdmin()) { toast("Solo el administrador cambia el dólar del depósito.", "error"); return; }
  const v = dolarDelDeposito();
  const mercado = (typeof dbDatos !== "undefined" && dbDatos && dbDatos.valores) ? dbDatos.valores : {};
  const botones = Object.keys(mercado).filter(t => mercado[t] > 0).map(t =>
    `<button type="button" class="btn btn-outline-secondary btn-sm" onclick="document.getElementById('depDolarInput').value='${mercado[t]}'">${escapeHtml((TC_TIPOS && TC_TIPOS[t]) || t)} $${Number(mercado[t]).toLocaleString("es-AR", { maximumFractionDigits: 2 })}</button>`).join(" ");
  const bd = depModal("modalDolarDeposito", "Dólar del depósito", `
    <p class="text-muted" style="font-size:13px;">Se usa en las ventas, cotizaciones y pagos del depósito, y para pasar a pesos los productos con precio en dólares. Solo cambia cuando lo cambiás vos.</p>
    <label class="form-label">Valor del dólar</label>
    <input type="number" class="form-control" id="depDolarInput" min="0" step="0.01" value="${v || ""}" placeholder="Ej: 1450">
    ${botones ? `<div class="form-text" style="margin:8px 0 4px;">Tomar una cotización del mercado (venta):</div><div class="d-flex flex-wrap gap-1">${botones}</div>` : ""}
    <div style="margin-top:14px;font-size:12px;font-weight:700;color:var(--slate-500);text-transform:uppercase;">Últimos cambios</div>
    <div id="depDolarHistorial" style="font-size:12.5px;max-height:180px;overflow:auto;">Cargando…</div>`,
    `<button class="btn btn-outline-secondary" data-cerrar>Cancelar</button><button class="btn btn-primary" id="depDolarGuardar">Guardar</button>`);
  bd.querySelector("#depDolarGuardar").onclick = async () => {
    const valor = Number(String(document.getElementById("depDolarInput").value).replace(",", "."));
    if (!(valor > 0)) { toast("Ingresá un valor válido", "error"); return; }
    const btn = bd.querySelector("#depDolarGuardar"); btn.disabled = true; btn.textContent = "Guardando…";
    const ok = await guardarDolarDeposito(valor);
    btn.disabled = false; btn.textContent = "Guardar";
    if (ok) bd.classList.remove("show");
  };
  setTimeout(() => { const i = document.getElementById("depDolarInput"); if (i) { i.focus(); i.select(); } }, 50);
  const d = await depDolarRefrescar();
  const cont = document.getElementById("depDolarHistorial");
  if (!cont) return;
  const hist = (d && d.historial) || [];
  cont.innerHTML = hist.length ? hist.slice(0, 15).map(h =>
    `<div style="display:flex;justify-content:space-between;gap:10px;padding:3px 0;border-bottom:1px solid var(--slate-100);">
      <span>${escapeHtml(depFechaHora(h.fecha))}${h.usuario ? " · " + escapeHtml(h.usuario) : ""}</span>
      <b class="mono">$${Number(h.valor).toLocaleString("es-AR", { maximumFractionDigits: 2 })}</b></div>`).join("")
    : '<span class="text-muted">Todavía no se fijó nunca.</span>';
}

/** Tarjeta en Configuración → Negocio (solo depósito) */
function depDolarTarjetaConfig() {
  const card = document.getElementById("cardDolarDeposito");
  if (!card) return;
  card.style.display = esModoDeposito() ? "" : "none";
  if (!esModoDeposito()) return;
  const v = dolarDelDeposito();
  document.getElementById("cfgDolarDepositoValor").textContent = v ? "$" + v.toLocaleString("es-AR", { maximumFractionDigits: 2 }) : "Sin fijar";
  document.getElementById("cfgDolarDepositoInfo").textContent = depDolar && depDolar.fecha
    ? "Fijado el " + depFechaHora(depDolar.fecha) + (depDolar.usuario ? " por " + depDolar.usuario : "") : "";
}

/* =====================================================================
   PRECIOS EN DÓLARES (Productos, solo depósito)
   El precio en pesos se calcula en el backend con el dólar del día y se
   guarda en el PRECIO de siempre (el POS, los pedidos y la app sin
   conexión siguen trabajando en pesos).
===================================================================== */
async function guardarPrecioUsdProducto(codigo, precioUsd, precioCajaUsd) {
  const r = await depPost("guardarPreciosUsd", { precios: [{ codigo, precioUsd: Number(precioUsd) || 0, precioCajaUsd: Number(precioCajaUsd) || 0 }] });
  if (!r.success) throw new Error(r.message || "No se pudo guardar el precio en dólares");
  return r;
}

/** Producto (en memoria del panel) por código */
function depProducto(codigo) {
  const lista = (typeof productosAdminGlobal !== "undefined" && productosAdminGlobal) || [];
  return lista.find(x => String(x.CODIGO) === String(codigo)) || null;
}

/** Al abrir el formulario de producto (nuevo o editar): campos en dólares, solo en el depósito */
function depPmPrepararUsd(codigo) {
  const wrap = document.getElementById("pmUsdWrap");
  if (!wrap) return;
  const dep = esModoDeposito();
  wrap.style.display = dep ? "" : "none";
  const p = codigo ? depProducto(codigo) : null;
  const usd = p && Number(p.PRECIO_USD) > 0 ? Number(p.PRECIO_USD) : "";
  const cajaUsd = p && Number(p.PRECIO_CAJA_USD) > 0 ? Number(p.PRECIO_CAJA_USD) : "";
  document.getElementById("pmPrecioUsd").value = usd;
  document.getElementById("pmPrecioCajaUsd").value = cajaUsd;
  window._depPmUsdAntes = { usd: Number(usd) || 0, caja: Number(cajaUsd) || 0 };
  const cl = document.getElementById("pmCodigoLocal");
  if (cl) { cl.value = p && p.CODIGO_LOCAL ? String(p.CODIGO_LOCAL) : ""; window._depPmCodigoLocalAntes = cl.value; }
  depPmUsdCambio();
}

/** Con precio en dólares, el precio en pesos se calcula solo (no se edita a mano) */
function depPmUsdCambio() {
  const precio = document.getElementById("pmPrecio");
  const precioCaja = document.getElementById("pmPrecioCaja");
  const info = document.getElementById("pmUsdInfo");
  if (!precio || !info) return;
  const usd = Number(document.getElementById("pmPrecioUsd").value) || 0;
  const cajaUsd = Number(document.getElementById("pmPrecioCajaUsd").value) || 0;
  const dolar = dolarDelDeposito();
  const dep = esModoDeposito();
  precio.readOnly = dep && usd > 0;
  if (precioCaja) precioCaja.readOnly = dep && cajaUsd > 0;
  if (!dep || (!usd && !cajaUsd)) { info.textContent = dep ? "Si cargás el precio en dólares, el precio en pesos se calcula con el dólar del día y se actualiza solo cuando cambia." : ""; return; }
  if (!(dolar > 0)) { info.innerHTML = '<span style="color:var(--red-500,#ef4444)">Primero fijá el dólar del depósito (Configuración → Negocio).</span>'; return; }
  if (usd > 0) precio.value = Math.round(usd * dolar);
  if (cajaUsd > 0 && precioCaja) precioCaja.value = Math.round(cajaUsd * dolar);
  info.textContent = `Con el dólar del depósito ($${dolar.toLocaleString("es-AR", { maximumFractionDigits: 2 })}): ` +
    [usd > 0 ? `unidad ${depPesos(usd * dolar)}` : "", cajaUsd > 0 ? `caja ${depPesos(cajaUsd * dolar)}` : ""].filter(Boolean).join(" · ") +
    ". Se actualiza solo cuando cambia el dólar.";
}

/**
 * Se envuelven las funciones del formulario de producto (las de admin.js
 * o las de la app de escritorio, que las reemplaza): al abrir se cargan
 * los campos en dólares y, si el producto se guardó, se guarda su
 * precio en dólares en el backend del depósito.
 */
function depInstalarEnFormularioProducto() {
  if (window._depFormProductoListo) return;
  window._depFormProductoListo = true;
  const envolverApertura = (nombre) => {
    const orig = window[nombre];
    if (typeof orig !== "function") return;
    window[nombre] = function (codigo) {
      const r = orig.apply(this, arguments);
      Promise.resolve(r).then(() => depPmPrepararUsd(nombre === "editarProducto" ? codigo : null));
      return r;
    };
  };
  envolverApertura("nuevoProducto");
  envolverApertura("editarProducto");

  const guardarOrig = window.guardarProductoForm;
  if (typeof guardarOrig === "function") {
    window.guardarProductoForm = async function () {
      if (!esModoDeposito()) return guardarOrig.apply(this, arguments);
      depPmUsdCambio();
      const usd = Number(document.getElementById("pmPrecioUsd").value) || 0;
      const cajaUsd = document.getElementById("pmVentaPorCaja").checked ? (Number(document.getElementById("pmPrecioCajaUsd").value) || 0) : 0;
      const codigo = document.getElementById("pmCodigo").value.trim();
      const antes = window._depPmUsdAntes || { usd: 0, caja: 0 };
      if ((usd > 0 || cajaUsd > 0) && !(dolarDelDeposito() > 0)) { toast("Primero fijá el dólar del depósito (Configuración → Negocio).", "error"); return; }
      const codigoLocal = String((document.getElementById("pmCodigoLocal") || {}).value || "").trim();
      const r = await guardarOrig.apply(this, arguments);
      const cerrado = !document.getElementById("productModalBackdrop").classList.contains("show");
      if (!cerrado) return r;
      if (codigoLocal !== (window._depPmCodigoLocalAntes || "")) depGuardarCodigoLocalConReintento(codigo, codigoLocal, 0);
      if (usd === antes.usd && cajaUsd === antes.caja) return r;
      depGuardarUsdConReintento(codigo, usd, cajaUsd, 0);
      return r;
    };
  }
}

/** Un producto nuevo creado sin conexión llega al backend recién al sincronizar: se reintenta un rato */
async function depGuardarUsdConReintento(codigo, usd, cajaUsd, intento) {
  try {
    const r = await guardarPrecioUsdProducto(codigo, usd, cajaUsd);
    if (r.noEncontrados && r.noEncontrados.length) throw Object.assign(new Error("todavía no está en el servidor"), { reintentar: true });
    toast(usd || cajaUsd ? `Precio en dólares guardado (${depUsd(usd || cajaUsd)})` : "El producto quedó con precio en pesos", "success");
    if (typeof cargarProductos === "function") setTimeout(cargarProductos, 500);
  } catch (e) {
    if ((e.reintentar || /no se encontró/i.test(e.message)) && intento < 6) { setTimeout(() => depGuardarUsdConReintento(codigo, usd, cajaUsd, intento + 1), 15000); return; }
    toast("El producto se guardó, pero no su precio en dólares: " + e.message + ". Volvé a abrirlo y guardalo de nuevo.", "error");
  }
}

/** Código del producto en el local (traspasos); como el precio en dólares, se reintenta si el producto recién se creó sin conexión */
async function depGuardarCodigoLocalConReintento(codigo, codigoLocal, intento) {
  try {
    const r = await depPost("guardarCodigosLocal", { codigos: [{ codigo, codigoLocal }] });
    if (!r.success || (r.noEncontrados && r.noEncontrados.length)) throw Object.assign(new Error(r.message || "todavía no está en el servidor"), { reintentar: true });
    const p = depProducto(codigo); if (p) p.CODIGO_LOCAL = codigoLocal;
  } catch (e) {
    if (e.reintentar && intento < 6) { setTimeout(() => depGuardarCodigoLocalConReintento(codigo, codigoLocal, intento + 1), 15000); return; }
    toast("El producto se guardó, pero no su código en el local: " + e.message, "error");
  }
}

/* ---------- arranque ---------- */
/** Lo que envuelve funciones del panel espera a que la app de escritorio haya puesto las suyas */
function depCuandoPanelListo(fn) {
  const inicio = Date.now();
  (function esperar() {
    const escritorio = !!window.veekpos;
    if (!escritorio || window.__posOfflineListo || Date.now() - inicio > 20000) { fn(); return; }
    setTimeout(esperar, 300);
  })();
}
/** Modo y dólar a partir de la configuración del negocio (en la app de escritorio, también la guardada) */
function depAplicarConfig(cfg) {
  if (!cfg) return;
  if (cfg.modoNegocio) {
    try { sessionStorage.setItem("modoNegocio", cfg.modoNegocio); } catch (e) {}
    if (typeof aplicarModoNegocio === "function") aplicarModoNegocio();
  }
  depDolarDesdeConfig(cfg);
}

function depEnvolverConfigNegocio() {
  const orig = window.cargarConfigNegocioDesdeBackend;
  if (typeof orig !== "function" || orig._dep) return;
  const envuelta = async function () {
    const cfg = await orig.apply(this, arguments);
    try { depAplicarConfig(cfg || (typeof configNegocioCache !== "undefined" ? configNegocioCache : null)); } catch (e) {}
    return cfg;
  };
  envuelta._dep = true;
  window.cargarConfigNegocioDesdeBackend = envuelta;
}

window.addEventListener("load", () => depCuandoPanelListo(() => {
  depInstalarEnFormularioProducto();
  depEnvolverConfigNegocio();
  if (typeof configNegocioCache !== "undefined") depAplicarConfig(configNegocioCache);
  if (esModoDeposito()) depDolarRefrescar();
}));

/* =====================================================================
   PEDIDOS EN DÓLARES (depósito)
   El ticket del POS está en pesos. Un pedido en dólares guarda cada
   precio y el total en dólares: los productos con precio en dólares
   usan ese precio exacto; el resto se divide por el tipo de cambio.
   Descuentos y recargos se aplican en la misma proporción.
===================================================================== */
function depConvertirPedidoAUsd(carrito, subtotalArs, totalArs, tipoCambio) {
  const tc = Number(tipoCambio) || 0;
  if (!(tc > 0)) return { carrito, total: totalArs };
  const r2 = (n) => Math.round(Number(n) * 100) / 100;
  const items = (carrito || []).map(it => {
    const p = depProducto(it.CODIGO);
    const usdExacto = p && Number(p.PRECIO_USD) > 0 && !it.esCaja && Math.round(Number(p.PRECIO_USD) * tc) === Math.round(Number(it.PRECIO))
      ? Number(p.PRECIO_USD) : null;
    return { ...it, PRECIO: usdExacto !== null ? usdExacto : r2(Number(it.PRECIO) / tc), PRECIO_ARS: Number(it.PRECIO) };
  });
  const sumaUsd = items.reduce((a, it) => a + it.PRECIO * Number(it.cantidad || 0), 0);
  const factor = subtotalArs > 0 ? Number(totalArs) / Number(subtotalArs) : 1;
  const total = r2(sumaUsd * factor);
  return { carrito: items, total, subtotal: r2(sumaUsd), descuento: r2(Math.max(0, sumaUsd - total)) };
}

/* =====================================================================
   COBRO EN DÓLARES (modal "Finalizar venta")
   Se cobra con el dólar del depósito (se puede cambiar para esa venta).
   Si los dólares no alcanzan, el resto se cobra en pesos; si sobran, el
   vuelto se da en pesos. La caja de dólares y la de pesos quedan bien
   en el cierre (ver getCierreCajaResumen en el backend).
===================================================================== */

function depMfvTotalArs() {
  const sub = (typeof ticketPOS !== "undefined" ? ticketPOS : []).reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
  return typeof calcularDescuentoPOS === "function" ? calcularDescuentoPOS(sub).total : sub;
}

/** El total del modal ya tiene el descuento/recargo que se está escribiendo */
function depMfvTotalModal() {
  const el = document.getElementById("mfvTotal");
  if (!el) return depMfvTotalArs();
  const n = Number(el.textContent.replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", "."));
  return isFinite(n) && n > 0 ? n : depMfvTotalArs();
}

const depRedondearUsd = (n) => Math.round((Number(n) || 0) * 100) / 100;

function depMfvPreparar() {
  const tcEl = document.getElementById("mfvUsdTc");
  if (tcEl && !(Number(tcEl.value) > 0)) tcEl.value = dolarDelDeposito() || "";
  depMfvActualizar();
}

/** Recalcula total en dólares y el resto/vuelto en pesos. tocado=true cuando el cajero escribió los dólares */
function depMfvActualizar(tocado) {
  const tc = Number((document.getElementById("mfvUsdTc") || {}).value) || 0;
  const recEl = document.getElementById("mfvUsdRecibido");
  const totalEl = document.getElementById("mfvUsdTotal");
  const restoEl = document.getElementById("mfvUsdResto");
  const labelEl = document.getElementById("mfvUsdRestoLabel");
  const infoEl = document.getElementById("mfvUsdInfo");
  if (!recEl || !totalEl || !restoEl) return;
  if (tocado) recEl.dataset.tocado = "1";
  const total = depMfvTotalModal();
  if (!(tc > 0)) {
    totalEl.textContent = "—"; restoEl.textContent = "—";
    if (infoEl) infoEl.textContent = "Fijá el dólar del día (o escribilo acá para esta venta).";
    return;
  }
  const totalUsd = depRedondearUsd(total / tc);
  totalEl.textContent = depUsd(totalUsd);
  // Si el cajero no escribió nada, se propone el total justo
  if (!recEl.dataset.tocado) recEl.value = totalUsd;
  const p = depCalcularPagoUsd(total, Number(recEl.value) || 0, tc);
  restoEl.classList.remove("negativo");
  if (p.restoArs > 0) { labelEl.textContent = "Resto en pesos"; restoEl.textContent = depPesos(p.restoArs); }
  else if (p.vueltoArs > 0) { labelEl.textContent = "Vuelto en pesos"; restoEl.textContent = depPesos(p.vueltoArs); }
  else { labelEl.textContent = "Resto en pesos"; restoEl.textContent = "$0"; }
  if (infoEl) {
    const tcDep = dolarDelDeposito();
    infoEl.textContent = (tcDep && Math.abs(tcDep - tc) > 0.001 ? "Ojo: usás un dólar distinto al del día (" + depPesos(tcDep) + "). " : "") +
      (p.restoArs > 0 ? "El resto se cobra en efectivo (pesos)." : p.vueltoArs > 0 ? "El vuelto se da en efectivo (pesos)." : "");
  }
}

/** { montoUsd, tipoCambio, restoArs, vueltoArs } — el redondeo es a pesos enteros */
function depCalcularPagoUsd(totalArs, montoUsd, tc) {
  const enPesos = Math.round(montoUsd * tc);
  const dif = Math.round(totalArs) - enPesos;
  // diferencias de menos de un centavo de dólar se toman como pago justo
  const tolerancia = Math.max(1, Math.ceil(tc * 0.01));
  return {
    montoUsd: depRedondearUsd(montoUsd), tipoCambio: Number(tc),
    restoArs: dif >= tolerancia ? dif : 0,
    vueltoArs: -dif >= tolerancia ? -dif : 0
  };
}

/** Lee el cobro del modal para guardar la venta, o { error } */
function depMfvLeerPago(totalArs) {
  const tc = Number((document.getElementById("mfvUsdTc") || {}).value) || 0;
  const usd = Number((document.getElementById("mfvUsdRecibido") || {}).value) || 0;
  if (!(tc > 0)) return { error: "Ingresá el valor del dólar" };
  if (!(usd > 0)) return { error: "Ingresá cuántos dólares recibiste" };
  return depCalcularPagoUsd(totalArs, usd, tc);
}

function depTextoPagoUsd(p) {
  let t = "Pagó " + depUsd(p.montoUsd) + " a " + depPesos(p.tipoCambio);
  if (p.restoArs > 0) t += " + " + depPesos(p.restoArs) + " en pesos";
  if (p.vueltoArs > 0) t += ", vuelto " + depPesos(p.vueltoArs) + " en pesos";
  return t;
}

/** Vuelve a pedir el catálogo del POS (tras recalcular precios, mandar o recibir un traspaso) */
window.refrescarCatalogoPOS = function () {
  setTimeout(() => {
    try { if (typeof invalidarCache === "function") invalidarCache("productos"); } catch (e) {}
    try { if (typeof actualizarCatalogoPOSManual === "function") actualizarCatalogoPOSManual(); } catch (e) {}
  }, 300);
};

/* =====================================================================
   PRECIOS A LA VISTA EN DÓLARES (depósito)
   En el depósito los precios se ven en dólares (preferido) y, chico,
   en pesos. Con el botón "US$ / $" se pasa a ver en pesos (queda
   guardado en esta PC). El ticket y el cobro siguen calculándose en
   pesos por dentro: lo que cambia es cómo se muestran.
===================================================================== */
const DEP_VISTA_KEY = "veekpos_deposito_vista_moneda";

function depMonedaVista() {
  try { return localStorage.getItem(DEP_VISTA_KEY) === "ARS" ? "ARS" : "USD"; } catch (e) { return "USD"; }
}

/** true si en esta PC del depósito los precios se muestran en dólares (y hay dólar fijado) */
function depVistaUsd() {
  return esModoDeposito() && depMonedaVista() === "USD" && dolarDelDeposito() > 0;
}

/** Dólares de un precio en pesos: el precio en dólares exacto del producto si lo tiene y coincide */
function depArsAUsd(ars, producto) {
  const tc = dolarDelDeposito();
  if (!(tc > 0)) return 0;
  const usd = producto ? Number(producto.PRECIO_USD) || 0 : 0;
  if (usd > 0 && Math.abs(Math.round(usd * tc) - Number(ars)) <= 1) return usd;
  return Math.round(Number(ars) / tc * 100) / 100;
}

/** HTML del precio para la grilla, la tabla y el ticket. conPesos=true agrega los pesos abajo, en chico */
function depPrecioVista(ars, producto, conPesos) {
  const pesos = "$" + Number(ars || 0).toLocaleString("es-AR");
  if (!depVistaUsd()) return pesos;
  const usd = depUsd(depArsAUsd(ars, producto));
  return conPesos ? `${usd}<small style="display:block;font-size:.72em;opacity:.65;font-weight:600;">${pesos}</small>` : `${usd} <span style="opacity:.6;">(${pesos})</span>`;
}

/** Total del ticket en dólares, debajo del total en pesos */
function depTotalesTicketUsd(items, totalArs) {
  const totalEl = document.getElementById("totalPOS");
  if (!totalEl) return;
  const fila = totalEl.closest(".trow");
  let extra = document.getElementById("totalUsdPOS");
  if (!depVistaUsd() || !items || !items.length) { if (extra) extra.remove(); return; }
  if (!extra && fila) {
    extra = document.createElement("div");
    extra.id = "totalUsdPOS";
    extra.className = "trow";
    extra.style.cssText = "font-weight:800;font-size:1.15em;color:var(--green-700,#15803d);";
    fila.parentNode.insertBefore(extra, fila);
  }
  // suma renglón por renglón (respeta los precios exactos en dólares) y aplica el mismo descuento/recargo
  const subArs = items.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
  const subUsd = items.reduce((a, i) => a + depArsAUsd(i.PRECIO, i) * i.cantidad, 0);
  const usd = subArs > 0 ? Math.round(subUsd * (totalArs / subArs) * 100) / 100 : 0;
  extra.innerHTML = `<span>Total en dólares</span><span class="money">${depUsd(usd)}</span>`;
}

/** Al lado del total del cobro: lo mismo en dólares */
function depMfvTituloUsd(totalArs) {
  const t = document.getElementById("mfvTotal");
  if (!t) return;
  let el = document.getElementById("mfvTotalUsd");
  if (!depVistaUsd()) { if (el) el.remove(); return; }
  if (!el) {
    el = document.createElement("span");
    el.id = "mfvTotalUsd";
    el.style.cssText = "font-family:monospace;color:var(--slate-500);font-size:.85em;margin-left:6px;";
    t.after(el);
  }
  const items = typeof ticketPOS !== "undefined" ? ticketPOS : [];
  const subArs = items.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
  const subUsd = items.reduce((a, i) => a + depArsAUsd(i.PRECIO, i) * i.cantidad, 0);
  el.textContent = "· " + depUsd(subArs > 0 ? Math.round(subUsd * (totalArs / subArs) * 100) / 100 : totalArs / dolarDelDeposito());
}

function depAlternarMonedaVista() {
  const nueva = depMonedaVista() === "USD" ? "ARS" : "USD";
  try { localStorage.setItem(DEP_VISTA_KEY, nueva); } catch (e) {}
  if (nueva === "USD" && !(dolarDelDeposito() > 0)) toast("Fijá el dólar del depósito para ver los precios en dólares", "error");
  else toast(nueva === "USD" ? "Precios en dólares (y en pesos, en chico)" : "Precios en pesos", "success");
  depRepintarPrecios();
}

function depRepintarPrecios() {
  ["btnMonedaVistaPOS", "btnMonedaVistaProductos"].forEach(id => {
    const b = document.getElementById(id);
    if (b) b.textContent = depMonedaVista() === "USD" ? "Ver en $" : "Ver en US$";
  });
  try { if (typeof renderPosGrid === "function") renderPosGrid((document.getElementById("posBusqueda") || {}).value?.trim() || ""); } catch (e) {}
  try { if (typeof renderTicketPOS === "function") renderTicketPOS(); } catch (e) {}
  try { const sec = document.getElementById("productos"); if (sec && sec.style.display === "block" && typeof filtrarProductos === "function") filtrarProductos(); } catch (e) {}
}
