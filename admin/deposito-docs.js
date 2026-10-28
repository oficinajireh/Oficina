/* =====================================================================
   deposito-docs.js — documentos A4 del DEPÓSITO
   ---------------------------------------------------------------------
   - Cotizaciones: se arman desde el ticket del POS (o desde cero),
     se guardan en el backend (hoja COTIZACIONES, número COT-0001…),
     se imprimen en A4, se buscan, se editan y se pasan al POS para
     venderlas o hacer el pedido (quedan "convertidas").
   - Boleta de entrega A4: original y duplicado en la misma hoja, con
     línea de firma, desde un pedido o una venta.
   - Datos de los documentos (CUIT, condición de IVA, condiciones...).
   Usa lo de admin.js y deposito.js (depPost, depModal, depPesos, depUsd,
   dolarDelDeposito, ticketPOS, productosPOS, confirmarAccion, toast...).
===================================================================== */

/* ---------- base de las hojas A4 ---------- */

const DEP_CSS_A4 = `
.a4doc{font-family:Arial,Helvetica,sans-serif;color:#111;font-size:10pt;line-height:1.35}
.a4doc *{box-sizing:border-box}
.a4doc .cab{display:flex;justify-content:space-between;align-items:flex-start;gap:6mm;border-bottom:2px solid #111;padding-bottom:3mm;margin-bottom:4mm}
.a4doc .neg-nombre{font-size:17pt;font-weight:800;letter-spacing:-.3px;line-height:1.1}
.a4doc .neg-sub{font-size:9pt;color:#444;margin-top:1mm}
.a4doc .doc-caja{border:1.5px solid #111;border-radius:2mm;padding:2mm 4mm;text-align:right;min-width:58mm}
.a4doc .doc-tipo{font-size:13pt;font-weight:800;letter-spacing:.5px}
.a4doc .doc-num{font-size:12pt;font-weight:700;font-family:"Courier New",monospace}
.a4doc .doc-dato{font-size:9pt;color:#333}
.a4doc .cliente{border:1px solid #bbb;border-radius:2mm;padding:2.5mm 4mm;margin-bottom:4mm;display:grid;grid-template-columns:1fr 1fr;gap:.8mm 6mm;font-size:9.5pt}
.a4doc .cliente b{font-weight:700}
.a4doc .cliente .ancho{grid-column:1 / -1}
.a4doc table.items{width:100%;border-collapse:collapse;font-size:9.5pt}
.a4doc table.items th{background:#eee;border:1px solid #999;padding:1.5mm 2mm;text-align:left;font-size:8.5pt;text-transform:uppercase;letter-spacing:.3px}
.a4doc table.items td{border:1px solid #bbb;padding:1.3mm 2mm;vertical-align:top}
.a4doc .der{text-align:right;white-space:nowrap}
.a4doc .cen{text-align:center}
.a4doc .mono{font-family:"Courier New",monospace;font-size:8.5pt}
.a4doc .totales{margin-left:auto;margin-top:3mm;width:72mm;border-collapse:collapse;font-size:10pt}
.a4doc .totales td{padding:1mm 2mm}
.a4doc .totales tr.total td{border-top:2px solid #111;font-size:13pt;font-weight:800;padding-top:1.5mm}
.a4doc .nota{margin-top:4mm;font-size:9pt;border-top:1px dashed #999;padding-top:2mm;white-space:pre-wrap}
.a4doc .pie{margin-top:5mm;font-size:8pt;color:#555;text-align:center}
.a4doc .firmas{display:flex;gap:10mm;margin-top:12mm}
.a4doc .firmas div{flex:1;border-top:1px solid #111;padding-top:1mm;font-size:8.5pt;text-align:center}
.a4doc .copia{position:relative}
.a4doc .etiqueta-copia{position:absolute;top:-3.6mm;right:0;font-size:7.5pt;font-weight:700;letter-spacing:1px;color:#555}
.a4doc .media-hoja{height:136mm;overflow:hidden}
.a4doc .corte{border-top:1px dashed #888;margin:3mm 0;position:relative}
.a4doc .corte span{position:absolute;left:50%;top:-2.2mm;transform:translateX(-50%);background:#fff;padding:0 2mm;font-size:7pt;color:#888}
.a4doc .salto{page-break-before:always}
.a4doc.compacta{font-size:9pt}
.a4doc.compacta .cab{padding-bottom:2mm;margin-bottom:2.5mm}
.a4doc.compacta .neg-nombre{font-size:14pt}
.a4doc.compacta .cliente{padding:1.5mm 3mm;margin-bottom:2.5mm;font-size:9pt}
.a4doc.compacta table.items td{padding:.8mm 1.5mm}
.a4doc.compacta .firmas{margin-top:8mm}
`;

function depCfgDocs() {
  const c = (typeof obtenerConfigNegocio === "function" ? obtenerConfigNegocio() : null) || {};
  return { ...c, doc: c.documentos || {} };
}

const depEsc = (t) => (typeof escapeHtml === "function" ? escapeHtml(String(t == null ? "" : t)) : String(t == null ? "" : t));

function depFechaCorta(iso) {
  const f = iso instanceof Date ? iso : new Date(iso);
  return isNaN(f) ? "" : f.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" });
}

/** Importe según la moneda del documento */
function depMonto(n, moneda) {
  return moneda === "USD" ? depUsd(n) : depPesos(n);
}

/** Cantidad sin decimales de más (2 / 2,5) */
function depCant(n) {
  return (Number(n) || 0).toLocaleString("es-AR", { maximumFractionDigits: 3 });
}

/** Encabezado del negocio + recuadro con el tipo y número del documento */
function depCabeceraA4(tipo, numero, lineasDerecha) {
  const c = depCfgDocs();
  const tels = [c.telefono1, c.telefono2].filter(Boolean).join(" · ");
  const datos = [c.direccion, tels ? "Tel: " + tels : "", c.doc.email, c.doc.cuit ? "CUIT " + c.doc.cuit : "", c.doc.condicionIva]
    .filter(Boolean).map(depEsc).join(" · ");
  return `<div class="cab">
    <div>
      <div class="neg-nombre">${depEsc(c.nombre || "")}</div>
      ${c.subtitulo && c.subtitulo !== "Punto de Venta" ? `<div class="neg-sub">${depEsc(c.subtitulo)}</div>` : ""}
      <div class="neg-sub">${datos}</div>
    </div>
    <div class="doc-caja">
      <div class="doc-tipo">${depEsc(tipo)}</div>
      ${numero ? `<div class="doc-num">${depEsc(numero)}</div>` : ""}
      ${(lineasDerecha || []).map(l => `<div class="doc-dato">${l}</div>`).join("")}
    </div>
  </div>`;
}

function depBloqueClienteA4(cl) {
  cl = cl || {};
  const linea = (k, v, ancho) => v ? `<div class="${ancho ? "ancho" : ""}">${k}: <b>${depEsc(v)}</b></div>` : "";
  const loc = [cl.localidad, cl.provincia].filter(Boolean).join(", ");
  return `<div class="cliente">
    <div class="ancho">Cliente: <b>${depEsc(cl.nombre || "—")}</b>${cl.empresa ? " · " + depEsc(cl.empresa) : ""}</div>
    ${linea("Dirección", [cl.direccion, loc].filter(Boolean).join(" — "), true)}
    ${linea("Teléfono", cl.telefono)}${linea("CUIT/DNI", cl.cuit)}${linea("Email", cl.email)}
  </div>`;
}

/** Imprime HTML en hoja A4 (mismo mecanismo que el resto de las impresiones A4 del panel) */
function depImprimirA4(html) {
  const termica = document.getElementById("thermalPrintFrame");
  if (termica) termica.innerHTML = "";
  const area = document.getElementById("etiquetasPrintArea");
  area.innerHTML = `<style>${DEP_CSS_A4}</style>` + html;
  if (typeof _setPrintPageSize === "function") _setPrintPageSize("CUSTOM", "@media print { @page { size:A4; margin:10mm; } }");
  setTimeout(() => window.print(), 150);
}

/* =====================================================================
   COTIZACIONES
===================================================================== */

let depCotEditor = null;          // cotización que se está armando en el editor
let depCotizacionEnCurso = null;  // cotización pasada al POS (para marcarla convertida al vender / hacer el pedido)
let _cotListaActual = [];
let _cotPausa = null;

const DEP_ESTADOS_COT = {
  PENDIENTE:  { txt: "Pendiente",  color: "#b45309", fondo: "#fef3c7" },
  ACEPTADA:   { txt: "Aceptada",   color: "#1d4ed8", fondo: "#dbeafe" },
  CONVERTIDA: { txt: "Convertida", color: "#15803d", fondo: "#dcfce7" },
  VENCIDA:    { txt: "Vencida",    color: "#6b7280", fondo: "#f3f4f6" },
  RECHAZADA:  { txt: "Rechazada",  color: "#b91c1c", fondo: "#fee2e2" },
  ANULADA:    { txt: "Anulada",    color: "#6b7280", fondo: "#f3f4f6" }
};

function depBadgeEstadoCot(estado) {
  const e = DEP_ESTADOS_COT[estado] || DEP_ESTADOS_COT.PENDIENTE;
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;font-size:11.5px;font-weight:700;color:${e.color};background:${e.fondo};">${e.txt}</span>`;
}

function cotBuscarConPausa() {
  clearTimeout(_cotPausa);
  _cotPausa = setTimeout(cargarCotizaciones, 350);
}

async function cargarCotizaciones() {
  const tbody = document.getElementById("tablaCotizaciones");
  if (!tbody) return;
  const q = (document.getElementById("cotBuscar") || {}).value || "";
  const estado = (document.getElementById("cotEstado") || {}).value || "";
  if (!_cotListaActual.length) tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">Cargando...</td></tr>`;
  try {
    const r = await depPost("cotizaciones", { q, estado, limite: 300 }, { timeoutMs: 20000 });
    if (!r.success) throw new Error(r.message || "error");
    _cotListaActual = r.cotizaciones || [];
    renderCotizaciones();
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-danger py-4">No se pudieron cargar las cotizaciones (${depEsc(e.message)}). Las cotizaciones necesitan conexión.</td></tr>`;
  }
}

function renderCotizaciones() {
  const tbody = document.getElementById("tablaCotizaciones");
  if (!tbody) return;
  if (!_cotListaActual.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">No hay cotizaciones${(document.getElementById("cotBuscar") || {}).value ? " que coincidan con la búsqueda" : " todavía. Armá una desde el Punto de Venta (botón 📝) o con “Nueva cotización”"}.</td></tr>`;
    return;
  }
  tbody.innerHTML = _cotListaActual.map(c => {
    const editable = c.estado !== "CONVERTIDA" && c.estado !== "ANULADA";
    return `<tr>
      <td class="mono" style="font-weight:700;">${depEsc(c.cotizacionId)}</td>
      <td>${depFechaCorta(c.fecha)}</td>
      <td>${depEsc(c.cliente)}<div class="text-muted" style="font-size:11.5px;">${c.cantidadItems} producto(s)${c.vendedor ? " · " + depEsc(c.vendedor) : ""}</div></td>
      <td class="text-end money" style="font-weight:700;">${depMonto(c.total, c.moneda)}</td>
      <td>${c.vence ? depFechaCorta(c.vence) : "—"}</td>
      <td>${depBadgeEstadoCot(c.estado)}${c.convertidaA ? `<div class="text-muted" style="font-size:11px;">${depEsc(c.convertidaA)}</div>` : ""}</td>
      <td class="text-end" style="white-space:nowrap;">
        <button class="btn btn-sm btn-outline-secondary" onclick="imprimirCotizacionId('${depEsc(c.cotizacionId)}')" title="Imprimir A4">🖨️</button>
        ${editable ? `<button class="btn btn-sm btn-outline-primary ms-1" onclick="pasarCotizacionAlPOS('${depEsc(c.cotizacionId)}')" title="Cargar en el Punto de Venta para venderla o hacer el pedido">Al POS</button>` : ""}
        <button class="btn btn-sm btn-outline-secondary ms-1" onclick="menuCotizacion('${depEsc(c.cotizacionId)}')" title="Más acciones">⋯</button>
      </td>
    </tr>`;
  }).join("");
}

async function depTraerCotizacion(id) {
  const r = await depPost("cotizacion", { cotizacionId: id }, { timeoutMs: 20000 });
  if (!r.success) throw new Error(r.message || "No se encontró la cotización");
  return r.cotizacion;
}

async function imprimirCotizacionId(id) {
  try { imprimirCotizacion(await depTraerCotizacion(id)); }
  catch (e) { toast("No se pudo abrir la cotización: " + e.message, "error"); }
}

function menuCotizacion(id) {
  const c = _cotListaActual.find(x => x.cotizacionId === id);
  if (!c) return;
  const editable = c.estado !== "CONVERTIDA" && c.estado !== "ANULADA";
  const boton = (txt, fn, clase) => `<button type="button" class="btn ${clase || "btn-outline-secondary"} w-100 mb-2" onclick="${fn}">${txt}</button>`;
  depModal("modalMenuCotizacion", c.cotizacionId + " · " + c.cliente, `
    ${editable ? boton("✏️ Editar", `cerrarMenuCot(); abrirEditorCotizacionId('${id}')`) : ""}
    ${boton("📄 Duplicar (nueva cotización con los mismos productos)", `cerrarMenuCot(); duplicarCotizacion('${id}')`)}
    ${editable && c.estado !== "ACEPTADA" ? boton("👍 Marcar como aceptada", `cerrarMenuCot(); estadoCotizacion('${id}', 'ACEPTADA')`) : ""}
    ${editable && c.estado !== "RECHAZADA" ? boton("👎 Marcar como rechazada", `cerrarMenuCot(); estadoCotizacion('${id}', 'RECHAZADA')`) : ""}
    ${editable && c.estado !== "PENDIENTE" && c.estado !== "VENCIDA" ? boton("↩️ Volver a pendiente", `cerrarMenuCot(); estadoCotizacion('${id}', 'PENDIENTE')`) : ""}
    ${editable ? boton("🚫 Anular", `cerrarMenuCot(); estadoCotizacion('${id}', 'ANULADA')`, "btn-outline-danger") : ""}
  `, `<button type="button" class="btn btn-outline-secondary" data-cerrar>Cerrar</button>`, 420);
}
function cerrarMenuCot() { const m = document.getElementById("modalMenuCotizacion"); if (m) m.classList.remove("show"); }

function estadoCotizacion(id, estado) {
  const hacer = async () => {
    try {
      const r = await depPost("estadoCotizacion", { cotizacionId: id, estado });
      if (!r.success) { toast(r.message || "No se pudo cambiar el estado", "error"); return; }
      toast(`${id}: ${DEP_ESTADOS_COT[estado].txt.toLowerCase()}`, "success");
      cargarCotizaciones();
    } catch (e) { toast("Sin conexión: " + e.message, "error"); }
  };
  if (estado === "ANULADA") confirmarAccion(`¿Anular la cotización ${id}? Se conserva en la lista, pero ya no se puede editar ni pasar al POS.`, hacer, "Anular cotización");
  else hacer();
}

/* ---------- editor ---------- */

function depProductoPorCodigo(codigo) {
  const lista = (typeof productosPOS !== "undefined" && productosPOS.length ? productosPOS : (typeof productosAdminGlobal !== "undefined" ? productosAdminGlobal : [])) || [];
  return lista.find(p => String(p.CODIGO) === String(codigo)) || null;
}

/** Precio unitario de un renglón del ticket en la moneda pedida */
function depPrecioEnMoneda(item, moneda, tc) {
  const ars = Number(item.PRECIO) || 0;
  if (moneda !== "USD") return Math.round(ars);
  const p = depProductoPorCodigo(item.CODIGO);
  const usd = p ? Number(p.PRECIO_USD) || 0 : 0;
  // si el producto tiene precio en dólares y el ticket usa su precio de lista, va el exacto
  if (usd > 0 && Math.abs(Math.round(usd * tc) - ars) <= 1) return usd;
  return tc > 0 ? Math.round(ars / tc * 100) / 100 : 0;
}

function abrirCotizacionDesdePOS() {
  if (typeof ticketPOS === "undefined" || !ticketPOS.length) { toast("Agregá productos al ticket para cotizarlos", "error"); return; }
  const subtotal = ticketPOS.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
  const { montoDescuento } = calcularDescuentoPOS(subtotal);
  const tc = dolarDelDeposito();
  // en dólares si el depósito está mostrando los precios en dólares
  const enUsd = typeof depVistaUsd === "function" && depVistaUsd();
  abrirEditorCotizacion({
    desdePOS: true,
    moneda: enUsd ? "USD" : "ARS",
    tipoCambio: tc,
    cliente: depCotizacionEnCurso ? depCotizacionEnCurso.cliente : {},
    clienteId: depCotizacionEnCurso ? depCotizacionEnCurso.clienteId : "",
    items: ticketPOS.map(i => ({ codigo: String(i.CODIGO || ""), producto: i.PRODUCTO, cantidad: i.cantidad, precio: depPrecioEnMoneda(i, enUsd ? "USD" : "ARS", tc) })),
    descuento: montoDescuento > 0 ? (enUsd ? Math.round(montoDescuento / tc * 100) / 100 : Math.round(montoDescuento)) : 0,
    observaciones: ""
  });
}

async function abrirEditorCotizacionId(id) {
  try {
    const c = await depTraerCotizacion(id);
    abrirEditorCotizacion({ ...c, cliente: { nombre: c.cliente, ...(c.clienteDatos || {}) } });
  } catch (e) { toast("No se pudo abrir la cotización: " + e.message, "error"); }
}

async function duplicarCotizacion(id) {
  try {
    const c = await depTraerCotizacion(id);
    abrirEditorCotizacion({ ...c, cotizacionId: "", estado: "", cliente: { nombre: c.cliente, ...(c.clienteDatos || {}) }, tipoCambio: c.moneda === "USD" ? (dolarDelDeposito() || c.tipoCambio) : c.tipoCambio });
  } catch (e) { toast("No se pudo duplicar: " + e.message, "error"); }
}

function abrirEditorCotizacion(datos) {
  const cfg = depCfgDocs();
  datos = datos || {};
  depCotEditor = {
    cotizacionId: datos.cotizacionId || "",
    desdePOS: !!datos.desdePOS,
    moneda: (datos.moneda || (typeof depVistaUsd === "function" && depVistaUsd() ? "USD" : "ARS")) === "USD" ? "USD" : "ARS",
    tipoCambio: Number(datos.tipoCambio) || dolarDelDeposito() || 0,
    clienteId: datos.clienteId || "",
    cliente: { ...(datos.cliente || {}) },
    items: (datos.items || []).map(i => ({ codigo: i.codigo || "", producto: i.producto || "", cantidad: Number(i.cantidad) || 1, precio: Number(i.precio) || 0 })),
    descuento: Number(datos.descuento) || 0,
    validezDias: datos.validezDias !== undefined && datos.validezDias !== "" ? Number(datos.validezDias) : (Number(cfg.doc.validezCotizacion) || 7),
    observaciones: datos.observaciones || ""
  };
  const cl = depCotEditor.cliente;
  const v = (x) => depEsc(x || "");
  const titulo = depCotEditor.cotizacionId ? "Editar cotización " + depCotEditor.cotizacionId : "Nueva cotización";
  const bd = depModal("modalCotizacion", titulo, `
    <div class="mb-2">
      <label class="form-label" style="font-size:12px;font-weight:700;">Cliente</label>
      <select class="form-select form-select-sm mb-2" id="cotClienteSel" onchange="cotElegirCliente(this.value)">
        <option value="">— Cliente nuevo / escribir los datos —</option>
      </select>
      <div class="row g-2">
        <div class="col-md-6"><input class="form-control form-control-sm" id="cotCliNombre" placeholder="Nombre o razón social *" value="${v(cl.nombre)}"></div>
        <div class="col-md-6"><input class="form-control form-control-sm" id="cotCliCuit" placeholder="CUIT / DNI" value="${v(cl.cuit)}"></div>
        <div class="col-md-6"><input class="form-control form-control-sm" id="cotCliDireccion" placeholder="Dirección" value="${v(cl.direccion)}"></div>
        <div class="col-md-3"><input class="form-control form-control-sm" id="cotCliLocalidad" placeholder="Localidad" value="${v(cl.localidad)}"></div>
        <div class="col-md-3"><input class="form-control form-control-sm" id="cotCliTelefono" placeholder="Teléfono" value="${v(cl.telefono)}"></div>
      </div>
    </div>
    <div class="row g-2 mb-2 align-items-end">
      <div class="col-4">
        <label class="form-label" style="font-size:12px;font-weight:700;">Moneda</label>
        <select class="form-select form-select-sm" id="cotMoneda" onchange="cotCambiarMoneda(this.value)">
          <option value="ARS" ${depCotEditor.moneda === "ARS" ? "selected" : ""}>Pesos</option>
          <option value="USD" ${depCotEditor.moneda === "USD" ? "selected" : ""}>Dólares</option>
        </select>
      </div>
      <div class="col-4">
        <label class="form-label" style="font-size:12px;font-weight:700;">Dólar</label>
        <input type="number" step="0.01" min="0" class="form-control form-control-sm" id="cotTc" value="${depCotEditor.tipoCambio || ""}" oninput="cotCambiarTc(this.value)">
      </div>
      <div class="col-4">
        <label class="form-label" style="font-size:12px;font-weight:700;">Validez (días)</label>
        <input type="number" min="0" class="form-control form-control-sm" id="cotValidez" value="${depCotEditor.validezDias}">
      </div>
    </div>
    <div class="table-responsive" style="max-height:38vh;overflow:auto;border:1px solid var(--slate-200);border-radius:8px;">
      <table class="table table-sm mb-0 align-middle" style="font-size:13px;">
        <thead style="position:sticky;top:0;background:var(--slate-50,#f8fafc);z-index:1;"><tr><th>Producto</th><th style="width:92px;">Cant.</th><th style="width:110px;">Precio</th><th class="text-end" style="width:100px;">Subtotal</th><th style="width:32px;"></th></tr></thead>
        <tbody id="cotItems"></tbody>
      </table>
    </div>
    <div class="d-flex gap-2 mt-2">
      <input class="form-control form-control-sm" id="cotAgregar" list="cotProductosLista" placeholder="Agregar producto: escribí código o nombre y elegilo de la lista" onchange="cotAgregarDesdeBuscador(this)">
      <datalist id="cotProductosLista"></datalist>
    </div>
    <div class="row g-2 mt-2">
      <div class="col-md-7"><textarea class="form-control form-control-sm" id="cotObs" rows="2" placeholder="Observaciones (salen impresas)">${v(depCotEditor.observaciones)}</textarea></div>
      <div class="col-md-5">
        <div class="d-flex justify-content-between align-items-center" style="font-size:13px;"><span>Subtotal</span><b id="cotSubtotal">—</b></div>
        <div class="d-flex justify-content-between align-items-center mt-1" style="font-size:13px;gap:8px;"><span>Descuento</span><input type="number" step="0.01" min="0" class="form-control form-control-sm text-end" style="max-width:110px;" id="cotDescuento" value="${depCotEditor.descuento || ""}" oninput="cotRender(true)"></div>
        <div class="d-flex justify-content-between align-items-center mt-1" style="font-size:16px;"><span style="font-weight:700;">Total</span><b id="cotTotal" style="font-size:18px;">—</b></div>
      </div>
    </div>
  `, `
    <button type="button" class="btn btn-outline-secondary" data-cerrar>Cancelar</button>
    <button type="button" class="btn btn-outline-primary" onclick="guardarCotizacionEditor(false)">Guardar</button>
    <button type="button" class="btn btn-primary" onclick="guardarCotizacionEditor(true)">Guardar e imprimir</button>
  `, 760);
  bd.dataset.fijo = "1"; // con tantos datos, un clic afuera no la cierra
  cotLlenarClientes();
  cotLlenarBuscador();
  cotRender();
  setTimeout(() => { const el = document.getElementById(cl.nombre ? "cotAgregar" : "cotClienteSel"); if (el) el.focus(); }, 80);
}

async function cotLlenarClientes() {
  const sel = document.getElementById("cotClienteSel");
  if (!sel) return;
  let lista = (typeof _clientesPedidoCache !== "undefined" && _clientesPedidoCache.length) ? _clientesPedidoCache : (typeof clientesGlobal !== "undefined" ? clientesGlobal : []);
  if (!lista.length && typeof cargarClientesParaPedido === "function") { try { lista = await cargarClientesParaPedido(); } catch (e) {} }
  [...(lista || [])].sort((a, b) => String(a.NOMBRE || "").localeCompare(String(b.NOMBRE || ""))).forEach(c => {
    const o = document.createElement("option");
    o.value = c.CLIENTE_ID || c.DNI || "";
    o.textContent = (c.NOMBRE || c.CLIENTE || "") + (c.DNI ? " · " + c.DNI : "");
    sel.appendChild(o);
  });
  if (depCotEditor && depCotEditor.clienteId) sel.value = depCotEditor.clienteId;
}

function cotElegirCliente(id) {
  if (!depCotEditor) return;
  depCotEditor.clienteId = id || "";
  if (!id) return;
  const lista = (typeof _clientesPedidoCache !== "undefined" && _clientesPedidoCache.length) ? _clientesPedidoCache : (typeof clientesGlobal !== "undefined" ? clientesGlobal : []);
  const c = (lista || []).find(x => String(x.CLIENTE_ID || x.DNI || "") === String(id));
  if (!c) return;
  const set = (i, val) => { const el = document.getElementById(i); if (el) el.value = val || ""; };
  set("cotCliNombre", c.NOMBRE || c.CLIENTE); set("cotCliCuit", c.DNI); set("cotCliDireccion", c.DIRECCION);
  set("cotCliLocalidad", c.LOCALIDAD); set("cotCliTelefono", c.TELEFONO);
}

function cotLlenarBuscador() {
  const dl = document.getElementById("cotProductosLista");
  if (!dl) return;
  const lista = (typeof productosPOS !== "undefined" && productosPOS.length ? productosPOS : (typeof productosAdminGlobal !== "undefined" ? productosAdminGlobal : [])) || [];
  dl.innerHTML = lista.slice(0, 5000).map(p => `<option value="${depEsc(p.CODIGO)} — ${depEsc(p.PRODUCTO)}"></option>`).join("");
}

function cotAgregarDesdeBuscador(input) {
  const texto = String(input.value || "");
  const codigo = texto.split(" — ")[0].trim();
  const p = depProductoPorCodigo(codigo);
  if (!p) { if (texto) toast("Elegí un producto de la lista", "error"); return; }
  const ya = depCotEditor.items.find(i => i.codigo === String(p.CODIGO));
  if (ya) ya.cantidad += 1;
  else depCotEditor.items.push({ codigo: String(p.CODIGO), producto: p.PRODUCTO, cantidad: 1, precio: depPrecioEnMoneda(p, depCotEditor.moneda, depCotEditor.tipoCambio) });
  input.value = "";
  cotRender();
  input.focus();
}

function cotCambiarMoneda(moneda) {
  const ed = depCotEditor;
  if (!ed || moneda === ed.moneda) return;
  const tc = Number(document.getElementById("cotTc").value) || 0;
  if (!(tc > 0)) { toast("Ingresá el dólar para pasar la cotización a otra moneda", "error"); document.getElementById("cotMoneda").value = ed.moneda; return; }
  ed.tipoCambio = tc;
  ed.items.forEach(i => {
    if (moneda === "USD") i.precio = depPrecioEnMoneda({ CODIGO: i.codigo, PRECIO: i.precio }, "USD", tc);
    else i.precio = Math.round(i.precio * tc);
  });
  const desc = Number(document.getElementById("cotDescuento").value) || 0;
  document.getElementById("cotDescuento").value = desc ? (moneda === "USD" ? Math.round(desc / tc * 100) / 100 : Math.round(desc * tc)) : "";
  ed.moneda = moneda;
  cotRender();
}

function cotCambiarTc(v) {
  if (depCotEditor) depCotEditor.tipoCambio = Number(v) || 0;
  cotRender(true);
}

function cotCambiarItem(idx, campo, valor) {
  const it = depCotEditor && depCotEditor.items[idx];
  if (!it) return;
  it[campo] = Math.max(0, Number(String(valor).replace(",", ".")) || 0);
  cotRender(true);
}

function cotQuitarItem(idx) {
  depCotEditor.items.splice(idx, 1);
  cotRender();
}

function cotTotales() {
  const ed = depCotEditor;
  const r = (n) => ed.moneda === "USD" ? Math.round(n * 100) / 100 : Math.round(n);
  const subtotal = r(ed.items.reduce((a, i) => a + r(i.precio) * i.cantidad, 0));
  const descuento = Math.min(subtotal, Math.max(0, r(Number((document.getElementById("cotDescuento") || {}).value) || 0)));
  return { subtotal, descuento, total: r(subtotal - descuento) };
}

/** soloTotales=true mientras se escribe en un campo de la tabla (no se redibuja, para no perder el foco) */
function cotRender(soloTotales) {
  const ed = depCotEditor;
  if (!ed) return;
  const tb = document.getElementById("cotItems");
  const paso = ed.moneda === "USD" ? "0.01" : "1";
  if (!soloTotales && tb) {
    tb.innerHTML = ed.items.length ? ed.items.map((i, n) => `<tr>
      <td>${depEsc(i.producto)}<div class="text-muted mono" style="font-size:11px;">${depEsc(i.codigo)}</div></td>
      <td><input type="number" min="0" step="any" class="form-control form-control-sm" value="${i.cantidad}" oninput="cotCambiarItem(${n}, 'cantidad', this.value)"></td>
      <td><input type="number" min="0" step="${paso}" class="form-control form-control-sm" value="${i.precio}" oninput="cotCambiarItem(${n}, 'precio', this.value)"></td>
      <td class="text-end money" id="cotSub${n}"></td>
      <td><button type="button" class="btn btn-sm btn-link text-danger p-0" onclick="cotQuitarItem(${n})" title="Quitar">✕</button></td>
    </tr>`).join("") : `<tr><td colspan="5" class="text-center text-muted py-3">Sin productos — agregalos con el buscador de abajo</td></tr>`;
  }
  ed.items.forEach((i, n) => { const el = document.getElementById("cotSub" + n); if (el) el.textContent = depMonto(i.precio * i.cantidad, ed.moneda); });
  const t = cotTotales();
  document.getElementById("cotSubtotal").textContent = depMonto(t.subtotal, ed.moneda);
  document.getElementById("cotTotal").textContent = depMonto(t.total, ed.moneda);
}

async function guardarCotizacionEditor(imprimir) {
  const ed = depCotEditor;
  if (!ed) return;
  const val = (id) => String((document.getElementById(id) || {}).value || "").trim();
  const cliente = { id: ed.clienteId, nombre: val("cotCliNombre"), cuit: val("cotCliCuit"), direccion: val("cotCliDireccion"), localidad: val("cotCliLocalidad"), telefono: val("cotCliTelefono") };
  if (!cliente.nombre) { toast("Escribí el nombre del cliente", "error"); document.getElementById("cotCliNombre").focus(); return; }
  const items = ed.items.filter(i => i.cantidad > 0);
  if (!items.length) { toast("La cotización no tiene productos", "error"); return; }
  const tc = Number(val("cotTc")) || 0;
  if (ed.moneda === "USD" && !(tc > 0)) { toast("Ingresá el dólar de la cotización", "error"); return; }
  const t = cotTotales();
  const botones = document.querySelectorAll("#modalCotizacion .product-modal-actions button");
  botones.forEach(b => b.disabled = true);
  try {
    const r = await depPost("guardarCotizacion", {
      cotizacionId: ed.cotizacionId, cliente, moneda: ed.moneda, tipoCambio: tc,
      items, descuento: t.descuento, validezDias: Number(val("cotValidez")) || 0, observaciones: val("cotObs"),
      vendedor: sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("usuarioLogueado") || ""
    }, { timeoutMs: 25000 });
    if (!r.success) { toast(r.message || "No se pudo guardar la cotización", "error"); return; }
    document.getElementById("modalCotizacion").classList.remove("show");
    toast(`Cotización ${r.cotizacion.cotizacionId} guardada`, "success");
    if (ed.desdePOS) {
      // el ticket ya quedó guardado en la cotización: se libera para el próximo cliente
      ticketPOS = [];
      if (typeof resetearDescuentoPOS === "function") resetearDescuentoPOS();
      if (typeof renderTicketPOS === "function") renderTicketPOS();
      depQuitarCotizacionEnCurso();
    }
    if (imprimir) imprimirCotizacion(r.cotizacion);
    const sec = document.getElementById("cotizaciones");
    if (sec && sec.style.display === "block") cargarCotizaciones();
  } catch (e) {
    toast("No se pudo guardar (las cotizaciones necesitan conexión): " + e.message, "error");
  } finally {
    botones.forEach(b => b.disabled = false);
  }
}

/* ---------- hoja A4 de la cotización ---------- */

function htmlCotizacion(c) {
  const cfg = depCfgDocs();
  const m = c.moneda;
  const filas = (c.items || []).map(i => `<tr>
    <td class="mono">${depEsc(i.codigo)}</td><td>${depEsc(i.producto)}</td>
    <td class="cen">${depCant(i.cantidad)}</td><td class="der">${depMonto(i.precio, m)}</td><td class="der">${depMonto(i.precio * i.cantidad, m)}</td>
  </tr>`).join("");
  const lineas = [
    "Fecha: <b>" + depFechaCorta(c.fecha) + "</b>",
    c.vence ? "Válida hasta: <b>" + depFechaCorta(c.vence) + "</b>" : "",
    m === "USD" ? "Precios en <b>dólares</b>" : "",
    c.vendedor ? "Atendió: " + depEsc(c.vendedor) : ""
  ].filter(Boolean);
  const cliente = { nombre: c.cliente, ...(c.clienteDatos || {}) };
  const condiciones = [cfg.doc.condicionesCotizacion, m === "USD" && c.tipoCambio ? `Dólar de referencia: ${depPesos(c.tipoCambio)} (los importes en dólares se pueden abonar en pesos al dólar del día de pago).` : ""].filter(Boolean).join("\n");
  return `<div class="a4doc">
    ${depCabeceraA4("COTIZACIÓN", c.cotizacionId, lineas)}
    ${depBloqueClienteA4(cliente)}
    <table class="items"><thead><tr><th style="width:24mm;">Código</th><th>Descripción</th><th class="cen" style="width:16mm;">Cant.</th><th class="der" style="width:28mm;">Precio unit.</th><th class="der" style="width:30mm;">Importe</th></tr></thead>
    <tbody>${filas}</tbody></table>
    <table class="totales">
      ${c.descuento ? `<tr><td>Subtotal</td><td class="der">${depMonto(c.subtotal, m)}</td></tr><tr><td>Descuento</td><td class="der">−${depMonto(c.descuento, m)}</td></tr>` : ""}
      <tr class="total"><td>TOTAL</td><td class="der">${depMonto(c.total, m)}</td></tr>
    </table>
    ${c.observaciones ? `<div class="nota"><b>Observaciones:</b> ${depEsc(c.observaciones)}</div>` : ""}
    ${condiciones ? `<div class="nota">${depEsc(condiciones)}</div>` : ""}
    <div class="pie">Documento no válido como factura · ${depEsc(cfg.nombre || "")}</div>
  </div>`;
}

function imprimirCotizacion(c) { depImprimirA4(htmlCotizacion(c)); }

/* ---------- pasar al POS (vender o hacer el pedido) ---------- */

async function pasarCotizacionAlPOS(id) {
  let c;
  try { c = await depTraerCotizacion(id); }
  catch (e) { toast("No se pudo abrir la cotización: " + e.message, "error"); return; }
  if (c.estado === "CONVERTIDA" || c.estado === "ANULADA") { toast("Esta cotización ya está " + c.estado.toLowerCase(), "error"); return; }
  const cargar = () => {
    const tc = dolarDelDeposito() || c.tipoCambio;
    const faltan = [];
    ticketPOS = (c.items || []).map(i => {
      const p = depProductoPorCodigo(i.codigo);
      if (!p) faltan.push(i.producto);
      let precio = c.moneda === "USD" ? Math.round(i.precio * tc) : Math.round(i.precio);
      // precio en dólares redondeado a centavos: si coincide con el de lista, va el de lista (sin diferencias de $1–$7)
      if (c.moneda === "USD" && p && Math.abs(Number(p.PRECIO) - i.precio * tc) <= tc * 0.005 + 0.5) precio = Number(p.PRECIO);
      return { ...(p || {}), CODIGO: i.codigo || (p && p.CODIGO) || "", PRODUCTO: i.producto, PRECIO: precio, cantidad: Number(i.cantidad) || 1 };
    });
    if (typeof resetearDescuentoPOS === "function") resetearDescuentoPOS();
    // el descuento de la cotización pasa como descuento en pesos
    if (c.descuento > 0) {
      descuentoTipoPOS = "MONTO"; ajusteModoPOS = "DESCUENTO";
      const val = c.moneda === "USD" ? Math.round(c.descuento * tc) : Math.round(c.descuento);
      const inp = document.getElementById("descuentoValorInput"); if (inp) inp.value = val;
      const mot = document.getElementById("descuentoMotivoInput"); if (mot) mot.value = "Cotización " + c.cotizacionId;
      if (typeof aplicarDescuentoPOS === "function") aplicarDescuentoPOS();
    }
    depCotizacionEnCurso = { id: c.cotizacionId, clienteId: c.clienteId, cliente: { nombre: c.cliente, ...(c.clienteDatos || {}) } };
    mostrarSeccion("pos");
    if (typeof renderTicketPOS === "function") renderTicketPOS();
    depPintarCotizacionEnCurso();
    toast(`${c.cotizacionId} cargada en el ticket` + (c.moneda === "USD" ? ` (en pesos al dólar ${depPesos(tc)})` : "") +
      (faltan.length ? `. Ojo: ${faltan.length} producto(s) ya no están en el catálogo` : ""), faltan.length ? "error" : "success");
  };
  if (typeof ticketPOS !== "undefined" && ticketPOS.length) confirmarAccion("El ticket actual tiene productos. ¿Reemplazarlos por los de la cotización " + c.cotizacionId + "?", cargar, "Pasar cotización al POS");
  else cargar();
}

function depPintarCotizacionEnCurso() {
  let el = document.getElementById("depCotEnCurso");
  if (!depCotizacionEnCurso) { if (el) el.remove(); return; }
  if (!el) {
    el = document.createElement("div");
    el.id = "depCotEnCurso";
    el.style.cssText = "display:flex;align-items:center;gap:8px;justify-content:space-between;background:#ede9fe;color:#5b21b6;border-radius:8px;padding:6px 10px;margin:6px 0;font-size:12.5px;font-weight:600;";
    const barra = document.querySelector("#pos .pos-action-bar");
    if (barra) barra.parentNode.insertBefore(el, barra); else return;
  }
  el.innerHTML = `<span>📝 ${depEsc(depCotizacionEnCurso.id)} · ${depEsc(depCotizacionEnCurso.cliente.nombre || "")}</span>
    <button type="button" class="btn btn-sm btn-link p-0" style="color:#5b21b6;" title="Desvincular (el ticket queda igual)" onclick="depQuitarCotizacionEnCurso()">✕</button>`;
}

function depQuitarCotizacionEnCurso() {
  depCotizacionEnCurso = null;
  depPintarCotizacionEnCurso();
}

async function depMarcarCotizacionConvertida(texto) {
  const c = depCotizacionEnCurso;
  if (!c) return;
  depQuitarCotizacionEnCurso();
  try {
    const r = await depPost("estadoCotizacion", { cotizacionId: c.id, estado: "CONVERTIDA", convertidaA: texto });
    if (!r.success) throw new Error(r.message);
    toast(`${c.id} quedó como convertida`, "success");
  } catch (e) {
    toast(`No se pudo marcar ${c.id} como convertida (${e.message}). Marcala desde Cotizaciones.`, "error");
  }
}

/** Ventas y pedidos hechos con una cotización cargada la marcan como convertida */
function depInstalarConversionCotizaciones() {
  if (window._depConvListo) return;
  window._depConvListo = true;
  const ventaOrig = window.registrarVentaCobrada;
  if (typeof ventaOrig === "function") {
    window.registrarVentaCobrada = async function (venta) {
      const enCurso = depCotizacionEnCurso;
      const r = await ventaOrig.apply(this, arguments);
      if (enCurso && depCotizacionEnCurso === enCurso) depMarcarCotizacionConvertida("Venta " + ((r && r.ventaId) || (venta && venta.ventaIdTemp) || ""));
      return r;
    };
  }
  const pedidoOrig = window.confirmarPedidoAdmin;
  if (typeof pedidoOrig === "function") {
    window.confirmarPedidoAdmin = async function () {
      const enCurso = depCotizacionEnCurso;
      const nombre = String((document.getElementById("paNombre") || {}).value || "").trim();
      const r = await pedidoOrig.apply(this, arguments);
      if (enCurso && depCotizacionEnCurso === enCurso && (typeof ticketPOS === "undefined" || ticketPOS.length === 0)) {
        depMarcarCotizacionConvertida("Pedido " + new Date().toLocaleDateString("es-AR") + (nombre ? " · " + nombre : ""));
      }
      return r;
    };
  }
  const abrirPedidoOrig = window.abrirModalPedidoAdmin;
  if (typeof abrirPedidoOrig === "function") {
    window.abrirModalPedidoAdmin = function () {
      const r = abrirPedidoOrig.apply(this, arguments);
      const c = depCotizacionEnCurso;
      if (c) {
        const set = (id, v) => { const el = document.getElementById(id); if (el && v) el.value = v; };
        if (c.clienteId) { const sel = document.getElementById("paClienteSelector"); if (sel) { sel.value = c.clienteId; if (sel.value === c.clienteId && typeof paSeleccionarCliente === "function") { paSeleccionarCliente(c.clienteId); return r; } } }
        set("paNombre", c.cliente.nombre); set("paDireccion", c.cliente.direccion); set("paLocalidad", c.cliente.localidad);
        set("paProvincia", c.cliente.provincia); set("paTelefono", c.cliente.telefono); set("paDni", c.cliente.cuit);
      }
      return r;
    };
  }
  const vaciarOrig = window.vaciarTicketPOS;
  if (typeof vaciarOrig === "function") {
    window.vaciarTicketPOS = function () {
      const r = vaciarOrig.apply(this, arguments);
      if (typeof ticketPOS === "undefined" || ticketPOS.length === 0) depQuitarCotizacionEnCurso();
      return r;
    };
  }
}

/* =====================================================================
   BOLETA DE ENTREGA A4
   Original y duplicado en la misma hoja (si entran; si no, una hoja
   cada una), con línea de firma "Recibí conforme".
===================================================================== */

let depBoletaDatos = null;
const DEP_BOLETA_PRECIOS_KEY = "veekpos_boleta_con_precios";
const DEP_BOLETA_MAX_MEDIA_HOJA = 11; // renglones que entran en media hoja

function depAbrirBoleta(datos) {
  depBoletaDatos = datos;
  let conPrecios = true;
  try { conPrecios = localStorage.getItem(DEP_BOLETA_PRECIOS_KEY) !== "0"; } catch (e) {}
  const cl = datos.cliente || {};
  const v = (x) => depEsc(x || "");
  depModal("modalBoletaEntrega", "Boleta de entrega · " + (datos.numero || ""), `
    <div class="row g-2">
      <div class="col-md-6"><label class="form-label" style="font-size:12px;">Cliente</label><input class="form-control form-control-sm" id="bolNombre" value="${v(cl.nombre)}"></div>
      <div class="col-md-6"><label class="form-label" style="font-size:12px;">CUIT / DNI</label><input class="form-control form-control-sm" id="bolCuit" value="${v(cl.cuit)}"></div>
      <div class="col-md-8"><label class="form-label" style="font-size:12px;">Dirección de entrega</label><input class="form-control form-control-sm" id="bolDireccion" value="${v(cl.direccion)}"></div>
      <div class="col-md-4"><label class="form-label" style="font-size:12px;">Localidad</label><input class="form-control form-control-sm" id="bolLocalidad" value="${v(cl.localidad)}"></div>
      <div class="col-md-6"><label class="form-label" style="font-size:12px;">Teléfono</label><input class="form-control form-control-sm" id="bolTelefono" value="${v(cl.telefono)}"></div>
      <div class="col-md-6"><label class="form-label" style="font-size:12px;">Transporte / entregó</label><input class="form-control form-control-sm" id="bolTransporte" value="${v(datos.transporte)}"></div>
      <div class="col-12"><label class="form-label" style="font-size:12px;">Observaciones</label><input class="form-control form-control-sm" id="bolObs" value="${v(datos.observaciones)}"></div>
    </div>
    <div class="form-check mt-2">
      <input class="form-check-input" type="checkbox" id="bolPrecios" ${conPrecios ? "checked" : ""}>
      <label class="form-check-label" for="bolPrecios">Mostrar precios e importes</label>
    </div>
    <div class="text-muted mt-1" style="font-size:12px;">${(datos.items || []).length} renglón(es) · ${(datos.items || []).length <= DEP_BOLETA_MAX_MEDIA_HOJA ? "original y duplicado en una sola hoja A4" : "original y duplicado en hojas separadas"}.</div>
  `, `<button type="button" class="btn btn-outline-secondary" data-cerrar>Cancelar</button>
      <button type="button" class="btn btn-primary" onclick="imprimirBoletaEntrega()">Imprimir</button>`, 620);
}

function imprimirBoletaEntrega() {
  const d = depBoletaDatos;
  if (!d) return;
  const val = (id) => String((document.getElementById(id) || {}).value || "").trim();
  const conPrecios = !!(document.getElementById("bolPrecios") || {}).checked;
  try { localStorage.setItem(DEP_BOLETA_PRECIOS_KEY, conPrecios ? "1" : "0"); } catch (e) {}
  const datos = { ...d, cliente: { ...(d.cliente || {}), nombre: val("bolNombre"), cuit: val("bolCuit"), direccion: val("bolDireccion"), localidad: val("bolLocalidad"), telefono: val("bolTelefono") },
    transporte: val("bolTransporte"), observaciones: val("bolObs") };
  const m = document.getElementById("modalBoletaEntrega"); if (m) m.classList.remove("show");
  depImprimirA4(htmlBoletaEntrega(datos, conPrecios));
}

function htmlBoletaEntrega(d, conPrecios) {
  const cfg = depCfgDocs();
  const m = d.moneda === "USD" ? "USD" : "ARS";
  const items = d.items || [];
  const mediaHoja = items.length <= DEP_BOLETA_MAX_MEDIA_HOJA;
  const totalUnidades = items.reduce((a, i) => a + (Number(i.cantidad) || 0), 0);
  const copia = (etiqueta) => `<div class="copia ${mediaHoja ? "media-hoja" : ""}">
    <div class="etiqueta-copia">${etiqueta}</div>
    ${depCabeceraA4("BOLETA DE ENTREGA", d.numero, [
      "Fecha: <b>" + depFechaCorta(d.fecha || new Date()) + "</b>",
      d.referencia ? depEsc(d.referencia) : "",
      d.transporte ? "Transporte: " + depEsc(d.transporte) : ""
    ].filter(Boolean))}
    ${depBloqueClienteA4(d.cliente)}
    <table class="items"><thead><tr><th style="width:22mm;">Código</th><th>Descripción</th><th class="cen" style="width:16mm;">Cant.</th>${conPrecios ? `<th class="der" style="width:26mm;">Precio</th><th class="der" style="width:28mm;">Importe</th>` : ""}</tr></thead>
    <tbody>${items.map(i => `<tr><td class="mono">${depEsc(i.codigo)}</td><td>${depEsc(i.producto)}</td><td class="cen">${depCant(i.cantidad)}</td>${conPrecios ? `<td class="der">${depMonto(i.precio, m)}</td><td class="der">${depMonto(i.precio * i.cantidad, m)}</td>` : ""}</tr>`).join("")}</tbody></table>
    <table class="totales">
      <tr><td>Unidades</td><td class="der">${depCant(totalUnidades)}</td></tr>
      ${conPrecios ? `<tr class="total"><td>TOTAL</td><td class="der">${depMonto(d.total, m)}</td></tr>` : ""}
    </table>
    ${d.observaciones ? `<div class="nota"><b>Observaciones:</b> ${depEsc(d.observaciones)}</div>` : ""}
    <div class="nota" style="border:none;padding-top:0;">${depEsc(cfg.doc.leyendaBoleta || "Recibí la mercadería detallada de conformidad.")}</div>
    <div class="firmas"><div>Firma</div><div>Aclaración</div><div>DNI</div></div>
  </div>`;
  return mediaHoja
    ? `<div class="a4doc compacta">${copia("ORIGINAL")}<div class="corte"><span>✂ cortar</span></div>${copia("DUPLICADO")}</div>`
    : `<div class="a4doc">${copia("ORIGINAL")}<div class="salto"></div>${copia("DUPLICADO")}</div>`;
}

/** Desde el detalle de un pedido */
function abrirBoletaDesdePedido() {
  if (typeof pedidoDetalleActual === "undefined" || !pedidoDetalleActual) return;
  const { pedido: p, detalle } = pedidoDetalleActual;
  depAbrirBoleta({
    numero: "PED " + p.PEDIDO_ID, fecha: new Date(), referencia: p.FECHA ? "Pedido del " + depFechaCorta(p.FECHA) : "",
    moneda: String(p.MONEDA || "ARS").toUpperCase(), total: Number(p.TOTAL) || 0,
    cliente: { nombre: p.CLIENTE, direccion: p.DIRECCION, localidad: p.LOCALIDAD, provincia: p.PROVINCIA, telefono: p.TELEFONO, cuit: p.DNI },
    transporte: p.EMPRESA || "", observaciones: p.OBSERVACIONES || "",
    items: (detalle || []).map(i => ({ codigo: i.CODIGO || "", producto: i.PRODUCTO, cantidad: Number(i.cantidad ?? i.CANTIDAD) || 0, precio: Number(i.PRECIO) || 0 }))
  });
}

function depItemsDeVenta(carrito) {
  let items = carrito;
  if (typeof items === "string") { try { items = JSON.parse(items); } catch (e) { items = []; } }
  return (Array.isArray(items) ? items : []).map(i => ({ codigo: i.CODIGO || "", producto: i.PRODUCTO, cantidad: Number(i.cantidad ?? i.CANTIDAD) || 0, precio: Number(i.PRECIO) || 0 }));
}

/** Desde el recibo de la venta recién hecha */
function abrirBoletaDesdeUltimaVenta() {
  const u = typeof ultimaVentaImprimible !== "undefined" ? ultimaVentaImprimible : null;
  if (!u) { toast("No hay venta para la boleta", "error"); return; }
  depAbrirBoleta({ numero: u.ventaId, fecha: u.fecha, moneda: "ARS", total: u.total, cliente: {}, items: depItemsDeVenta(u.items) });
}

/** Desde el Historial de ventas */
function abrirBoletaDesdeVenta(v) {
  if (!v) return;
  const items = depItemsDeVenta(v.CARRITO);
  if (!items.length) { toast("Esta venta no tiene el detalle de productos guardado", "error"); return; }
  depAbrirBoleta({ numero: String(v.VENTA_ID || v.ID || ""), fecha: new Date(), referencia: v.FECHA ? "Venta del " + depFechaCorta(v.FECHA) : "",
    moneda: "ARS", total: Number(v.TOTAL) || 0, cliente: {}, items });
}

/* =====================================================================
   DATOS DE LOS DOCUMENTOS (Configuración → Negocio)
===================================================================== */

function depDocumentosTarjetaConfig() {
  const d = depCfgDocs().doc;
  const set = (id, v) => { const el = document.getElementById(id); if (el && document.activeElement !== el) el.value = v == null ? "" : v; };
  set("docCuit", d.cuit); set("docCondicionIva", d.condicionIva); set("docEmail", d.email);
  set("docValidez", d.validezCotizacion || 7); set("docCondiciones", d.condicionesCotizacion); set("docLeyendaBoleta", d.leyendaBoleta);
}

async function guardarConfigDocumentos() {
  const val = (id) => String((document.getElementById(id) || {}).value || "").trim();
  const datos = { cuit: val("docCuit"), condicionIva: val("docCondicionIva"), email: val("docEmail"), validezCotizacion: val("docValidez") || "7",
    condicionesCotizacion: val("docCondiciones"), leyendaBoleta: val("docLeyendaBoleta") };
  try {
    const r = await depPost("guardarConfigDocumentos", datos);
    if (!r.success) { toast(r.message || "No se pudo guardar", "error"); return; }
    if (typeof configNegocioCache !== "undefined") configNegocioCache.documentos = { ...datos, validezCotizacion: Number(datos.validezCotizacion) || 7 };
    toast("Datos de los documentos guardados", "success");
  } catch (e) { toast("Sin conexión: " + e.message, "error"); }
}

/* ---------- arranque ---------- */
window.addEventListener("load", () => depCuandoPanelListo(() => {
  depInstalarConversionCotizaciones();
  const cfgOrig = window.mostrarConfigTab;
  if (typeof cfgOrig === "function" && !cfgOrig._depDocs) {
    window.mostrarConfigTab = function (tab) {
      const r = cfgOrig.apply(this, arguments);
      if (tab === "negocio" && esModoDeposito()) depDocumentosTarjetaConfig();
      return r;
    };
    window.mostrarConfigTab._depDocs = true;
  }
}));
