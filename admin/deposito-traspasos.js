/* =====================================================================
   deposito-traspasos.js — mercadería del depósito al local
   ---------------------------------------------------------------------
   DEPÓSITO: arma el traspaso (a mano, escaneando o desde el ticket del
   POS con 🚚), lo manda: su stock baja y queda "en camino". Imprime el
   remito de traspaso A4 (original y duplicado, con firmas).
   LOCAL: ve lo que viene en camino y confirma lo que llegó (con
   diferencias si las hay): recién ahí suma el stock, como ingreso de
   mercadería del depósito.
   Vinculación (una sola vez): el depósito muestra su clave de
   traspasos; el local la pega junto con la dirección del depósito.
===================================================================== */

let depTraspasoArmado = []; // [{codigo, producto, cantidad, stock}]
let _trasListaDeposito = [];
let _trasDelDeposito = null; // respuesta de traspasosDelDeposito (local)

const DEP_ESTADOS_TRA = {
  EN_CAMINO:                { txt: "En camino",               color: "#b45309", fondo: "#fef3c7" },
  RECIBIDO:                 { txt: "Recibido",                color: "#15803d", fondo: "#dcfce7" },
  RECIBIDO_CON_DIFERENCIAS: { txt: "Recibido con diferencias", color: "#b91c1c", fondo: "#fee2e2" },
  ANULADO:                  { txt: "Anulado",                 color: "#6b7280", fondo: "#f3f4f6" }
};
function depBadgeTra(e) {
  const x = DEP_ESTADOS_TRA[e] || { txt: e, color: "#333", fondo: "#eee" };
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;font-size:11.5px;font-weight:700;color:${x.color};background:${x.fondo};">${x.txt}</span>`;
}

/** ¿Se muestra "Traspasos" en el menú? En el depósito siempre; en el local, si está vinculado */
function depTraspasosVisible() {
  if (esModoDeposito()) return true;
  const cfg = typeof configNegocioCache !== "undefined" ? configNegocioCache : {};
  return !!(cfg && cfg.depositoVinculado);
}
function depAplicarMenuTraspasos() {
  const ver = depTraspasosVisible();
  document.querySelectorAll("[data-traspasos]").forEach(a => {
    const destino = a.getAttribute("data-target");
    a.style.display = ver && (!destino || seccionPermitidaParaRol(destino)) ? "" : "none";
  });
}

function cargarTraspasos() {
  const sec = document.getElementById("traspasos");
  if (!sec) return;
  if (esModoDeposito()) depRenderSeccionDeposito(); else depRenderSeccionLocal();
}

/* =================== DEPÓSITO =================== */

function depRenderSeccionDeposito() {
  const sec = document.getElementById("traspasos");
  if (!sec.dataset.vista || sec.dataset.vista !== "deposito") {
    sec.dataset.vista = "deposito";
    sec.innerHTML = `
      <div class="page-header">
        <div><h1>Traspasos al local</h1><div class="sub">Lo que mandás baja del stock del depósito y queda en camino hasta que el local confirma qué llegó</div></div>
        <div class="d-flex gap-2"><button class="btn btn-outline-secondary" onclick="depCargarTraspasosDeposito()">↻ Actualizar</button></div>
      </div>
      <div class="card p-3 mb-3" id="trasArmado">
        <h6 style="margin-bottom:8px;">Nuevo traspaso</h6>
        <div class="d-flex gap-2 flex-wrap">
          <input class="form-control" id="trasBuscar" list="trasProductosLista" style="max-width:520px;" placeholder="Escaneá el código o escribí el producto y elegilo de la lista" onkeydown="if(event.key==='Enter'){event.preventDefault(); depTrasAgregar(this);}" onchange="depTrasAgregar(this)">
          <datalist id="trasProductosLista"></datalist>
        </div>
        <div class="table-responsive mt-2">
          <table class="table table-sm align-middle mb-0" style="font-size:13.5px;">
            <thead><tr><th>Producto</th><th style="width:150px;">Código en el local</th><th class="text-end" style="width:110px;">Stock depósito</th><th style="width:110px;">Cantidad</th><th style="width:36px;"></th></tr></thead>
            <tbody id="trasItems"></tbody>
          </table>
        </div>
        <div class="d-flex gap-2 mt-2 flex-wrap align-items-center">
          <input class="form-control" id="trasObs" style="max-width:520px;" placeholder="Observaciones (transporte, chofer, bultos…)">
          <div class="ms-auto d-flex gap-2 align-items-center">
            <span class="text-muted" id="trasResumen" style="font-size:13px;"></span>
            <button class="btn btn-outline-secondary" onclick="depTrasVaciar()">Vaciar</button>
            <button class="btn btn-success fw-bold" id="btnEnviarTraspaso" onclick="depTrasEnviar()">🚚 Mandar al local</button>
          </div>
        </div>
      </div>
      <div class="card p-3">
        <div class="table-responsive">
          <table class="table table-hover align-middle mb-0" style="font-size:13.5px;">
            <thead><tr><th>Número</th><th>Fecha</th><th class="text-end">Unidades</th><th>Estado</th><th>Recepción</th><th></th></tr></thead>
            <tbody id="trasLista"><tr><td colspan="6" class="text-center text-muted py-4">Cargando...</td></tr></tbody>
          </table>
        </div>
      </div>`;
  }
  depTrasLlenarBuscador();
  depTrasRenderArmado();
  depCargarTraspasosDeposito();
}

function depTrasProductos() {
  return (typeof productosPOS !== "undefined" && productosPOS.length ? productosPOS : (typeof productosAdminGlobal !== "undefined" ? productosAdminGlobal : [])) || [];
}

function depTrasLlenarBuscador() {
  const dl = document.getElementById("trasProductosLista");
  if (!dl) return;
  dl.innerHTML = depTrasProductos().slice(0, 5000).map(p => `<option value="${depEsc(p.CODIGO)} — ${depEsc(p.PRODUCTO)}${p.CODIGO_LOCAL ? " (local: " + depEsc(p.CODIGO_LOCAL) + ")" : ""}"></option>`).join("");
}

function depTrasAgregar(input, cantidad) {
  const texto = String(input.value || "").trim();
  if (!texto) return;
  const codigo = texto.split(" — ")[0].trim();
  const lista = depTrasProductos();
  const p = lista.find(x => String(x.CODIGO) === codigo) || lista.find(x => String(x.CODIGO_CAJA || "") === codigo)
    || lista.find(x => String(x.CODIGO_LOCAL || "").trim() !== "" && String(x.CODIGO_LOCAL).trim() === codigo); // también con el código del local
  if (!p) { toast("No se encontró el producto " + codigo, "error"); input.select(); return; }
  const cant = cantidad || (String(p.CODIGO_CAJA || "") === codigo && Number(p.UNIDADES_POR_CAJA) > 0 ? Number(p.UNIDADES_POR_CAJA) : 1);
  const ya = depTraspasoArmado.find(i => i.codigo === String(p.CODIGO));
  if (ya) ya.cantidad += cant;
  else depTraspasoArmado.push({ codigo: String(p.CODIGO), producto: p.PRODUCTO, cantidad: cant, stock: Number(p.STOCK) || 0, codigoLocal: String(p.CODIGO_LOCAL || "") });
  input.value = "";
  depTrasRenderArmado();
  input.focus();
}

function depTrasRenderArmado() {
  const tb = document.getElementById("trasItems");
  if (!tb) return;
  tb.innerHTML = depTraspasoArmado.length ? depTraspasoArmado.map((i, n) => `<tr>
    <td>${depEsc(i.producto)}<div class="text-muted mono" style="font-size:11px;">${depEsc(i.codigo)}</div></td>
    <td><input class="form-control form-control-sm mono" value="${depEsc(i.codigoLocal || "")}" placeholder="igual" title="Código de este producto en el local (vacío = el mismo). Se recuerda." oninput="depTrasCodigoLocal(${n}, this.value)"></td>
    <td class="text-end ${i.cantidad > i.stock ? "text-danger fw-bold" : "text-muted"}">${depCant(i.stock)}</td>
    <td><input type="number" min="0" step="any" class="form-control form-control-sm" value="${i.cantidad}" oninput="depTrasCantidad(${n}, this.value)"></td>
    <td><button class="btn btn-sm btn-link text-danger p-0" onclick="depTrasQuitar(${n})" title="Quitar">✕</button></td>
  </tr>`).join("") : `<tr><td colspan="5" class="text-center text-muted py-3">Escaneá o buscá los productos que van al local (o armalos en el POS y tocá 🚚)</td></tr>`;
  depTrasResumen();
}
function depTrasResumen() {
  const el = document.getElementById("trasResumen");
  if (el) {
    const u = depTraspasoArmado.reduce((a, i) => a + (Number(i.cantidad) || 0), 0);
    el.textContent = depTraspasoArmado.length ? `${depTraspasoArmado.length} producto(s) · ${depCant(u)} unidades` : "";
  }
}
function depTrasCantidad(n, v) { if (depTraspasoArmado[n]) depTraspasoArmado[n].cantidad = Math.max(0, Number(v) || 0); depTrasResumen(); }
function depTrasCodigoLocal(n, v) { if (depTraspasoArmado[n]) depTraspasoArmado[n].codigoLocal = String(v || "").trim(); }
function depTrasQuitar(n) { depTraspasoArmado.splice(n, 1); depTrasRenderArmado(); }
function depTrasVaciar() { depTraspasoArmado = []; depTrasRenderArmado(); }

/** 🚚 en el POS: el ticket pasa a ser el traspaso */
function traspasoDesdePOS() {
  if (typeof ticketPOS === "undefined" || !ticketPOS.length) { toast("Agregá al ticket los productos que van al local", "error"); return; }
  ticketPOS.forEach(i => {
    const ya = depTraspasoArmado.find(x => x.codigo === String(i.CODIGO));
    if (ya) ya.cantidad += Number(i.cantidad) || 0;
    else depTraspasoArmado.push({ codigo: String(i.CODIGO), producto: i.PRODUCTO, cantidad: Number(i.cantidad) || 0, stock: Number(i.STOCK) || 0, codigoLocal: String(i.CODIGO_LOCAL || "") });
  });
  // el stock que muestra el ticket ya está descontado en pantalla: se toma el del catálogo
  depTraspasoArmado.forEach(i => { const p = depTrasProductos().find(x => String(x.CODIGO) === i.codigo); if (p) i.stock = Number(p.STOCK) || 0; });
  ticketPOS = [];
  if (typeof resetearDescuentoPOS === "function") resetearDescuentoPOS();
  if (typeof renderTicketPOS === "function") renderTicketPOS();
  mostrarSeccion("traspasos");
  toast("Revisá el traspaso y tocá “Mandar al local”", "success");
}

function depTrasEnviar() {
  const items = depTraspasoArmado.filter(i => i.cantidad > 0).map(i => ({ codigo: i.codigo, producto: i.producto, cantidad: i.cantidad, codigoLocal: i.codigoLocal || "" }));
  if (!items.length) { toast("El traspaso no tiene productos", "error"); return; }
  const unidades = items.reduce((a, i) => a + i.cantidad, 0);
  const faltan = depTraspasoArmado.filter(i => i.cantidad > i.stock);
  const enviar = async (permitirSinStock) => {
    const btn = document.getElementById("btnEnviarTraspaso");
    if (btn) btn.disabled = true;
    try {
      const r = await depPost("crearTraspaso", { items, observaciones: (document.getElementById("trasObs") || {}).value || "", permitirSinStock: !!permitirSinStock,
        usuarioNombre: sessionStorage.getItem("nombreUsuario") || "" }, { timeoutMs: 30000 });
      if (!r.success) {
        if (r.sinStock) { confirmarAccion(r.message + ". ¿Mandar igual? (el stock del depósito queda en negativo)", () => enviar(true), "Stock insuficiente"); return; }
        toast(r.message || "No se pudo armar el traspaso", "error"); return;
      }
      toast(`Traspaso ${r.traspaso.traspasoId} en camino (${depCant(r.traspaso.unidades)} unidades)`, "success");
      depTraspasoArmado = [];
      const obs = document.getElementById("trasObs"); if (obs) obs.value = "";
      depTrasRenderArmado();
      // stock del depósito en pantalla
      Object.keys(r.stock || {}).forEach(c => { depTrasProductos().forEach(p => { if (String(p.CODIGO) === c) p.STOCK = r.stock[c]; }); });
      if (typeof window.refrescarCatalogoPOS === "function") window.refrescarCatalogoPOS();
      depCargarTraspasosDeposito();
      confirmarAccion(`¿Imprimir el remito del traspaso ${r.traspaso.traspasoId} (original y duplicado)?`, () => imprimirRemitoTraspaso(r.traspaso), "Remito de traspaso");
    } catch (e) {
      toast("No se pudo mandar (los traspasos necesitan conexión): " + e.message, "error");
    } finally {
      if (btn) btn.disabled = false;
    }
  };
  confirmarAccion(`¿Mandar al local ${items.length} producto(s), ${depCant(unidades)} unidades?${faltan.length ? ` Ojo: ${faltan.length} producto(s) superan el stock del depósito.` : ""} El stock del depósito baja ahora.`, () => enviar(false), "Traspaso al local");
}

async function depCargarTraspasosDeposito() {
  const tb = document.getElementById("trasLista");
  if (!tb) return;
  try {
    const r = await depPost("traspasos", { limite: 200 }, { timeoutMs: 20000 });
    if (!r.success) throw new Error(r.message);
    _trasListaDeposito = r.traspasos || [];
    tb.innerHTML = _trasListaDeposito.length ? _trasListaDeposito.map(t => {
      const dif = (t.diferencias || []).map(d => `${depEsc(d.producto)}: mandado ${depCant(d.enviado)}, llegó ${depCant(d.recibido)}`).join("<br>");
      return `<tr>
        <td class="mono" style="font-weight:700;">${depEsc(t.traspasoId)}</td>
        <td>${depFechaHora(t.fecha)}${t.usuario ? `<div class="text-muted" style="font-size:11px;">${depEsc(t.usuario)}</div>` : ""}</td>
        <td class="text-end">${depCant(t.unidades)}${t.unidadesRecibidas !== null && t.unidadesRecibidas !== t.unidades ? `<div class="text-danger" style="font-size:11px;">llegaron ${depCant(t.unidadesRecibidas)}</div>` : ""}</td>
        <td>${depBadgeTra(t.estado)}</td>
        <td style="font-size:12px;">${t.fechaRecepcion ? depFechaHora(t.fechaRecepcion) + (t.recibio ? " · " + depEsc(t.recibio) : "") : ""}${dif ? `<div class="text-danger mt-1">${dif}</div>` : ""}</td>
        <td class="text-end" style="white-space:nowrap;">
          <button class="btn btn-sm btn-outline-secondary" onclick="imprimirRemitoTraspasoId('${depEsc(t.traspasoId)}')" title="Remito A4">🖨️</button>
          ${t.estado === "EN_CAMINO" ? `<button class="btn btn-sm btn-outline-danger ms-1" onclick="depAnularTraspaso('${depEsc(t.traspasoId)}')">Anular</button>` : ""}
        </td></tr>`;
    }).join("") : `<tr><td colspan="6" class="text-center text-muted py-4">Todavía no mandaste traspasos.</td></tr>`;
  } catch (e) {
    tb.innerHTML = `<tr><td colspan="6" class="text-center text-danger py-4">No se pudieron cargar los traspasos (${depEsc(e.message)}).</td></tr>`;
  }
}

function imprimirRemitoTraspasoId(id) {
  const t = _trasListaDeposito.find(x => x.traspasoId === id) || ((_trasDelDeposito && [...(_trasDelDeposito.enCamino || []), ...(_trasDelDeposito.recibidos || [])]) || []).find(x => x.traspasoId === id);
  if (t) imprimirRemitoTraspaso(t);
}

function depAnularTraspaso(id) {
  confirmarAccion(`¿Anular el traspaso ${id}? La mercadería vuelve al stock del depósito y el local ya no lo va a poder recibir.`, async () => {
    try {
      const r = await depPost("anularTraspaso", { traspasoId: id });
      if (!r.success) { toast(r.message || "No se pudo anular", "error"); return; }
      toast(`${id} anulado: el stock volvió al depósito`, "success");
      if (typeof window.refrescarCatalogoPOS === "function") window.refrescarCatalogoPOS();
      depCargarTraspasosDeposito();
    } catch (e) { toast("Sin conexión: " + e.message, "error"); }
  }, "Anular traspaso");
}

/* ---------- remito de traspaso A4 ---------- */
function htmlRemitoTraspaso(t) {
  const items = t.items || [];
  const mediaHoja = items.length <= DEP_BOLETA_MAX_MEDIA_HOJA + 2;
  const copia = (etq) => `<div class="copia ${mediaHoja ? "media-hoja" : ""}">
    <div class="etiqueta-copia">${etq}</div>
    ${depCabeceraA4("REMITO DE TRASPASO", t.traspasoId, ["Fecha: <b>" + depFechaCorta(t.fecha) + "</b>", "Del <b>depósito</b> al <b>local</b>", t.usuario ? "Armó: " + depEsc(t.usuario) : ""].filter(Boolean))}
    <table class="items"><thead><tr><th style="width:24mm;">Cód. depósito</th><th style="width:22mm;">Cód. local</th><th>Producto</th><th class="cen" style="width:20mm;">Cantidad</th><th class="cen" style="width:20mm;">Llegó</th></tr></thead>
    <tbody>${items.map(i => `<tr><td class="mono">${depEsc(i.codigo)}</td><td class="mono">${depEsc(i.codigoLocalSugerido || i.codigoLocal || "—")}</td><td>${depEsc(i.producto)}</td><td class="cen"><b>${depCant(i.cantidad)}</b></td><td></td></tr>`).join("")}</tbody></table>
    <table class="totales"><tr class="total"><td>UNIDADES</td><td class="der">${depCant(t.unidades)}</td></tr></table>
    ${t.observaciones ? `<div class="nota"><b>Observaciones:</b> ${depEsc(t.observaciones)}</div>` : ""}
    <div class="firmas"><div>Entregó (depósito)</div><div>Transportó</div><div>Recibió (local)</div></div>
  </div>`;
  return mediaHoja
    ? `<div class="a4doc compacta">${copia("ORIGINAL")}<div class="corte"><span>✂ cortar</span></div>${copia("DUPLICADO")}</div>`
    : `<div class="a4doc">${copia("ORIGINAL")}<div class="salto"></div>${copia("DUPLICADO")}</div>`;
}
function imprimirRemitoTraspaso(t) { depImprimirA4(htmlRemitoTraspaso(t)); }

/* =================== LOCAL =================== */

function depRenderSeccionLocal() {
  const sec = document.getElementById("traspasos");
  if (sec.dataset.vista !== "local") {
    sec.dataset.vista = "local";
    sec.innerHTML = `
      <div class="page-header">
        <div><h1>Traspasos desde el depósito</h1><div class="sub">Confirmá lo que llegó: el stock del local se suma recién al recibir</div></div>
        <div class="d-flex gap-2"><button class="btn btn-outline-secondary" onclick="depCargarTraspasosLocal()">↻ Actualizar</button></div>
      </div>
      <h6 style="margin:4px 0 8px;">En camino</h6>
      <div id="trasEnCamino" class="mb-3"><div class="text-muted">Consultando al depósito...</div></div>
      <h6 style="margin:4px 0 8px;">Últimos recibidos</h6>
      <div class="card p-3"><div class="table-responsive"><table class="table table-sm align-middle mb-0" style="font-size:13px;">
        <thead><tr><th>Número</th><th>Enviado</th><th>Recibido</th><th class="text-end">Unidades</th><th>Estado</th><th></th></tr></thead>
        <tbody id="trasRecibidos"></tbody></table></div></div>`;
  }
  depCargarTraspasosLocal();
}

async function depCargarTraspasosLocal() {
  const caja = document.getElementById("trasEnCamino");
  const tb = document.getElementById("trasRecibidos");
  if (!caja) return;
  // para elegir el producto del local al recibir hace falta su lista (con los ocultos)
  if (!depProductosDelLocal().length || (typeof productosAdminGlobal !== "undefined" && !productosAdminGlobal.length)) {
    try { if (typeof cargarProductos === "function") await cargarProductos(); } catch (e) {}
  }
  try {
    const r = await depPost("traspasosDelDeposito", {}, { timeoutMs: 30000 });
    if (!r.success) throw new Error(r.message || "error");
    _trasDelDeposito = r;
    caja.innerHTML = r.enCamino.length ? r.enCamino.map(t => {
      const sinProducto = t.items.filter(i => !i.existeEnLocal);
      return `<div class="card p-3 mb-2" style="border-left:4px solid #f59e0b;">
        <div class="d-flex justify-content-between align-items-center flex-wrap gap-2">
          <div><b class="mono">${depEsc(t.traspasoId)}</b> · ${depFechaHora(t.fecha)} · <b>${depCant(t.unidades)}</b> unidades en ${t.items.length} producto(s)${t.observaciones ? ` · <span class="text-muted">${depEsc(t.observaciones)}</span>` : ""}</div>
          <div class="d-flex gap-2"><button class="btn btn-sm btn-outline-secondary" onclick="imprimirRemitoTraspasoId('${depEsc(t.traspasoId)}')">🖨️ Remito</button>
          <button class="btn btn-sm btn-success fw-bold" onclick="depAbrirRecepcion('${depEsc(t.traspasoId)}')">Recibir</button></div>
        </div>
        <div class="text-muted mt-1" style="font-size:12.5px;">${t.items.slice(0, 6).map(i => `${depEsc(i.producto)} ×${depCant(i.cantidad)}`).join(" · ")}${t.items.length > 6 ? " · …" : ""}</div>
        ${sinProducto.length ? `<div class="text-danger mt-1" style="font-size:12px;">⚠️ ${sinProducto.length} producto(s) sin su código del local (${sinProducto.map(i => depEsc(i.producto)).join(", ")}): al recibir, elegí qué producto del local es cada uno (queda recordado para la próxima).</div>` : ""}
      </div>`;
    }).join("") : `<div class="card p-3 text-muted">No hay traspasos en camino.</div>`;
    tb.innerHTML = (r.recibidos || []).length ? r.recibidos.map(t => `<tr>
      <td class="mono">${depEsc(t.traspasoId)}</td><td>${depFechaHora(t.fecha)}</td><td>${depFechaHora(t.fechaRecepcion)}${t.recibio ? " · " + depEsc(t.recibio) : ""}</td>
      <td class="text-end">${depCant(t.unidadesRecibidas ?? t.unidades)}${t.unidadesRecibidas !== null && t.unidadesRecibidas !== t.unidades ? ` <span class="text-muted">de ${depCant(t.unidades)}</span>` : ""}</td>
      <td>${depBadgeTra(t.estado)}</td>
      <td class="text-end"><button class="btn btn-sm btn-link p-0" onclick="imprimirRemitoTraspasoId('${depEsc(t.traspasoId)}')">🖨️</button></td></tr>`).join("")
      : `<tr><td colspan="6" class="text-center text-muted py-3">Todavía no recibiste traspasos.</td></tr>`;
  } catch (e) {
    caja.innerHTML = `<div class="card p-3 text-danger">No se pudo consultar al depósito: ${depEsc(e.message)}</div>`;
  }
}

function depAbrirRecepcion(id) {
  const t = _trasDelDeposito && _trasDelDeposito.enCamino.find(x => x.traspasoId === id);
  if (!t) return;
  const bd = depModal("modalRecepcionTraspaso", "Recibir " + t.traspasoId, `
    <p class="text-muted" style="font-size:13px;">Contá lo que llegó. Si algo no coincide, corregí la cantidad: la diferencia queda anotada acá y en el depósito.</p>
    <div class="table-responsive" style="max-height:50vh;overflow:auto;">
      <table class="table table-sm align-middle mb-0" style="font-size:13px;">
        <thead><tr><th>Producto (depósito)</th><th style="width:230px;">Producto en el local</th><th class="text-end" style="width:80px;">Mandaron</th><th style="width:96px;">Llegó</th></tr></thead>
        <tbody>${t.items.map((i, n) => `<tr>
          <td>${depEsc(i.producto)}<div class="text-muted mono" style="font-size:11px;">${depEsc(i.codigo)}</div></td>
          <td><input class="form-control form-control-sm" list="recProductosLocal" data-cl="${n}" value="${depEsc(i.codigoLocalSugerido || "")}" placeholder="Código o nombre del local" oninput="depRecepcionResumen()" onchange="depRecepcionNormalizar(this)">
            <div class="mono" style="font-size:11px;" id="recCl${n}"></div></td>
          <td class="text-end">${depCant(i.cantidad)}</td>
          <td><input type="number" min="0" step="any" class="form-control form-control-sm" data-rec="${n}" value="${i.cantidad}" oninput="depRecepcionResumen()"></td></tr>`).join("")}</tbody>
      </table>
    </div>
    <datalist id="recProductosLocal">${depProductosDelLocal().slice(0, 5000).map(p => `<option value="${depEsc(p.CODIGO)} — ${depEsc(p.PRODUCTO)}"></option>`).join("")}</datalist>
    <div class="mt-2" id="recResumen" style="font-size:13px;"></div>
  `, `<button type="button" class="btn btn-outline-secondary" data-cerrar>Cancelar</button>
      <button type="button" class="btn btn-success fw-bold" id="btnConfirmarRecepcion" onclick="depConfirmarRecepcion('${depEsc(id)}')">Confirmar recepción</button>`, 640);
  bd.dataset.fijo = "1";
  bd.dataset.traspaso = id;
  depRecepcionResumen();
}

function depRecepcionLeer() {
  const id = (document.getElementById("modalRecepcionTraspaso") || {}).dataset?.traspaso;
  const t = _trasDelDeposito && _trasDelDeposito.enCamino.find(x => x.traspasoId === id);
  if (!t) return null;
  const recibidos = t.items.map((i, n) => ({ codigo: i.codigo, cantidad: Math.max(0, Number((document.querySelector(`#modalRecepcionTraspaso [data-rec="${n}"]`) || {}).value) || 0),
    codigoLocal: String((document.querySelector(`#modalRecepcionTraspaso [data-cl="${n}"]`) || {}).value || "").split(" — ")[0].trim() }));
  return { t, recibidos };
}

function depRecepcionResumen() {
  const d = depRecepcionLeer();
  const el = document.getElementById("recResumen");
  if (!d || !el) return;
  const dif = d.t.items.filter((i, n) => d.recibidos[n].cantidad !== Number(i.cantidad));
  const total = d.recibidos.reduce((a, r) => a + r.cantidad, 0);
  // cada renglón: ¿el código elegido existe en el local?
  let sinProducto = 0;
  d.recibidos.forEach((r, n) => {
    const p = depProductoLocal(r.codigoLocal);
    const el2 = document.getElementById("recCl" + n);
    if (!p && r.cantidad > 0) sinProducto++;
    if (el2) el2.innerHTML = p ? `<span class="text-success">✓ ${depEsc(p.PRODUCTO)}</span>` : (r.codigoLocal ? `<span class="text-danger">no existe en el local</span>` : `<span class="text-danger">elegí el producto del local</span>`);
  });
  const btn = document.getElementById("btnConfirmarRecepcion");
  if (btn) btn.textContent = sinProducto ? `Confirmar (${sinProducto} sin producto quedan afuera)` : "Confirmar recepción";
  el.innerHTML = dif.length
    ? `<span class="text-danger fw-bold">${dif.length} diferencia(s)</span> · llegan ${depCant(total)} de ${depCant(d.t.unidades)} unidades`
    : `<span class="text-success fw-bold">Llegó todo</span> · ${depCant(total)} unidades`;
}

/** Productos del local para elegir al recibir (incluye los ocultos del catálogo) */
function depProductosDelLocal() {
  const admin = typeof productosAdminGlobal !== "undefined" ? productosAdminGlobal : [];
  return (admin && admin.length ? admin : (typeof productosPOS !== "undefined" ? productosPOS : [])) || [];
}
function depProductoLocal(codigo) {
  const c = String(codigo || "").trim().toUpperCase();
  if (!c) return null;
  return depProductosDelLocal().find(p => String(p.CODIGO).trim().toUpperCase() === c) || null;
}
/** Si eligió de la lista ("CÓDIGO — Nombre"), queda solo el código */
function depRecepcionNormalizar(input) {
  input.value = String(input.value || "").split(" — ")[0].trim();
  depRecepcionResumen();
}

async function depConfirmarRecepcion(id) {
  const d = depRecepcionLeer();
  if (!d) return;
  const btn = document.getElementById("btnConfirmarRecepcion");
  if (btn) { btn.disabled = true; btn.textContent = "Confirmando..."; }
  try {
    const r = await depPost("recibirTraspaso", { traspasoId: id, recibidos: d.recibidos, usuarioNombre: sessionStorage.getItem("nombreUsuario") || "" }, { timeoutMs: 45000 });
    if (!r.success) { toast(r.message || "No se pudo confirmar", "error"); if (r.yaProcesado) { document.getElementById("modalRecepcionTraspaso").classList.remove("show"); depCargarTraspasosLocal(); } return; }
    document.getElementById("modalRecepcionTraspaso").classList.remove("show");
    toast(`${id} recibido: se sumó el stock al local` + (r.noExisten && r.noExisten.length ? ` (quedaron afuera ${r.noExisten.length} producto(s) que no existen acá)` : ""), r.noExisten && r.noExisten.length ? "error" : "success");
    if (typeof window.refrescarCatalogoPOS === "function") window.refrescarCatalogoPOS();
    if (typeof invalidarCache === "function") { try { invalidarCache("productos"); } catch (e) {} }
    depCargarTraspasosLocal();
  } catch (e) {
    toast("No se pudo confirmar (sin conexión): " + e.message + ". Probá de nuevo: nunca se suma dos veces.", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = "Confirmar recepción"; }
  }
}

/* =================== CONFIGURACIÓN → SISTEMA =================== */

async function depTarjetaTraspasosConfig() {
  const card = document.getElementById("cardTraspasosConfig");
  if (!card) return;
  const admin = depEsAdmin();
  if (!admin) { card.style.display = "none"; return; }
  card.style.display = "";
  if (esModoDeposito()) {
    card.innerHTML = `<h6 style="margin-bottom:4px;">Traspasos al local</h6>
      <p class="text-muted" style="font-size:13px;margin-bottom:8px;">En el panel del <b>local</b> (Configuración → Sistema → Depósito vinculado) pegá esta dirección y esta clave. Con eso el local ve lo que le mandás y confirma lo que llega.</p>
      <label class="form-label" style="font-size:12px;">Dirección de este backend (depósito)</label>
      <div class="input-group input-group-sm mb-2"><input class="form-control mono" id="trasCfgUrl" readonly value="${depEsc(typeof API_URL !== "undefined" ? API_URL : "")}"><button class="btn btn-outline-secondary" onclick="depCopiar('trasCfgUrl')">Copiar</button></div>
      <label class="form-label" style="font-size:12px;">Clave de traspasos</label>
      <div class="input-group input-group-sm"><input class="form-control mono" id="trasCfgClave" readonly value="Cargando..."><button class="btn btn-outline-secondary" onclick="depCopiar('trasCfgClave')">Copiar</button>
      <button class="btn btn-outline-danger" onclick="depRegenerarClaveTraspasos()" title="La clave anterior deja de funcionar">Cambiar clave</button></div>`;
    try {
      const r = await depPost("claveTraspasos", {});
      document.getElementById("trasCfgClave").value = r.success ? r.clave : (r.message || "No se pudo leer");
    } catch (e) { document.getElementById("trasCfgClave").value = "Sin conexión"; }
    return;
  }
  const cfg = typeof configNegocioCache !== "undefined" ? configNegocioCache : {};
  let sugerida = "";
  try { if (window.veekpos && window.veekpos.perfilInfo) sugerida = ((await window.veekpos.perfilInfo()).urls || {}).deposito || ""; } catch (e) {}
  card.innerHTML = `<h6 style="margin-bottom:4px;">Depósito vinculado</h6>
    <p class="text-muted" style="font-size:13px;margin-bottom:8px;">${cfg.depositoVinculado ? "✅ Este local está vinculado con el depósito: los traspasos aparecen en <b>Mercadería → Traspasos</b>." : "Para recibir mercadería del depósito: pegá la dirección y la clave que muestra el depósito en Configuración → Sistema → Traspasos al local."}</p>
    <label class="form-label" style="font-size:12px;">Dirección del backend del depósito</label>
    <input class="form-control form-control-sm mono mb-2" id="trasVincUrl" placeholder="https://script.google.com/macros/s/…/exec" value="${depEsc(sugerida)}">
    <label class="form-label" style="font-size:12px;">Clave de traspasos</label>
    <input class="form-control form-control-sm mono mb-2" id="trasVincClave" placeholder="trp_…" autocomplete="off">
    <div class="d-flex gap-2 align-items-center flex-wrap">
      <button class="btn btn-primary btn-sm" onclick="depVincularDeposito()">${cfg.depositoVinculado ? "Volver a vincular" : "Vincular"}</button>
      ${cfg.depositoVinculado ? `<button class="btn btn-outline-danger btn-sm" onclick="depDesvincularDeposito()">Desvincular</button>` : ""}
      <span id="trasVincEstado" style="font-size:12.5px;"></span>
    </div>`;
}

function depCopiar(id) {
  const el = document.getElementById(id);
  if (!el) return;
  const hecho = () => toast("Copiado", "success");
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(el.value).then(hecho, () => { el.select(); document.execCommand("copy"); hecho(); });
  else { el.select(); document.execCommand("copy"); hecho(); }
}

function depRegenerarClaveTraspasos() {
  confirmarAccion("¿Cambiar la clave de traspasos? El local deja de ver los traspasos hasta que pegues la clave nueva en su panel.", async () => {
    try {
      const r = await depPost("claveTraspasos", { nueva: true });
      if (r.success) { document.getElementById("trasCfgClave").value = r.clave; toast("Clave nueva generada: pegala en el panel del local", "success"); }
      else toast(r.message || "No se pudo cambiar", "error");
    } catch (e) { toast("Sin conexión: " + e.message, "error"); }
  }, "Cambiar clave");
}

async function depVincularDeposito() {
  const url = String(document.getElementById("trasVincUrl").value || "").trim();
  const clave = String(document.getElementById("trasVincClave").value || "").trim();
  const est = document.getElementById("trasVincEstado");
  if (!url || !clave) { toast("Completá la dirección y la clave", "error"); return; }
  est.textContent = "Probando la conexión con el depósito...";
  try {
    const r = await depPost("vincularDeposito", { url, clave }, { timeoutMs: 40000 });
    if (!r.success) { est.innerHTML = `<span class="text-danger">${depEsc(r.message || "No se pudo vincular")}</span>`; return; }
    if (typeof configNegocioCache !== "undefined") configNegocioCache.depositoVinculado = true;
    est.innerHTML = `<span class="text-success">Vinculado con ${depEsc(r.deposito || "el depósito")}${r.enCamino ? ` · ${r.enCamino} traspaso(s) en camino` : ""}</span>`;
    depAplicarMenuTraspasos();
    toast("Local vinculado con el depósito", "success");
    setTimeout(depTarjetaTraspasosConfig, 1500);
  } catch (e) { est.innerHTML = `<span class="text-danger">Sin conexión: ${depEsc(e.message)}</span>`; }
}

function depDesvincularDeposito() {
  confirmarAccion("¿Desvincular este local del depósito? Dejás de ver los traspasos (no se borra nada).", async () => {
    try {
      const r = await depPost("vincularDeposito", { url: "", clave: "" });
      if (!r.success) { toast(r.message || "No se pudo", "error"); return; }
      if (typeof configNegocioCache !== "undefined") configNegocioCache.depositoVinculado = false;
      depAplicarMenuTraspasos();
      depTarjetaTraspasosConfig();
    } catch (e) { toast("Sin conexión: " + e.message, "error"); }
  }, "Desvincular depósito");
}

/* ---------- arranque ---------- */
window.addEventListener("load", () => depCuandoPanelListo(() => {
  depAplicarMenuTraspasos();
  // la configuración llega después: se vuelve a mirar cuando se aplica
  const orig = window.depAplicarConfig;
  if (typeof orig === "function" && !orig._tras) {
    window.depAplicarConfig = function () { const r = orig.apply(this, arguments); try { depAplicarMenuTraspasos(); } catch (e) {} return r; };
    window.depAplicarConfig._tras = true;
  }
  setTimeout(depAplicarMenuTraspasos, 3000);
}));
