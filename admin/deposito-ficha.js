/* =====================================================================
   deposito-ficha.js — ficha completa del cliente en el DEPÓSITO
   ---------------------------------------------------------------------
   Se suma arriba del detalle de cuenta del cliente (el mismo modal de
   siempre, con saldos ARS/USD, pagos, pedidos y deudas extra):
   - Datos: dirección y localidad, teléfono (WhatsApp), CUIT, email, notas
   - Acciones: nueva cotización, estado de cuenta A4, editar datos
   - Sus cotizaciones (reimprimir / pasar al POS)
   El estado de cuenta A4 arma los movimientos en orden (pedidos y
   cargos = debe; pagos = haber) con saldo corrido en pesos y en dólares.
===================================================================== */

function depFichaCliente() {
  return (typeof detalleClienteDatosActuales !== "undefined" && detalleClienteDatosActuales) ? detalleClienteDatosActuales : null;
}

function depRenderFicha() {
  const data = depFichaCliente();
  const modal = document.getElementById("detalleClienteModalBackdrop");
  if (!modal) return;
  let caja = document.getElementById("depFicha");
  if (!esModoDeposito() || !data || !data.cliente) { if (caja) caja.remove(); return; }
  const cuerpo = modal.querySelector(".product-modal-body");
  if (!caja) {
    caja = document.createElement("div");
    caja.id = "depFicha";
    cuerpo.insertBefore(caja, cuerpo.firstChild);
  }
  const c = data.cliente;
  const esCredito = String(c.A_CREDITO || "").toUpperCase() === "SI";
  const tel = String(c.TELEFONO || "").trim();
  const telWa = tel.replace(/\D/g, "");
  const wa = telWa.length >= 8 ? `https://wa.me/${telWa.startsWith("54") ? telWa : "549" + telWa.replace(/^0/, "")}` : "";
  const loc = [c.LOCALIDAD, c.PROVINCIA].filter(Boolean).join(", ");
  const fila = (k, v) => v ? `<div><span class="text-muted" style="font-size:11.5px;">${k}</span><div style="font-weight:600;">${v}</div></div>` : "";
  const alta = c.FECHA_ALTA ? depFechaCorta(c.FECHA_ALTA) : "";
  caja.innerHTML = `
    <div class="card p-3 mb-2" style="border-left:4px solid #f59e0b;">
      <div class="d-flex justify-content-between align-items-start gap-2 flex-wrap">
        <div>
          <span style="display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;${esCredito ? "background:#fee2e2;color:#b91c1c;" : "background:#e0f2fe;color:#0369a1;"}">${esCredito ? "CUENTA CORRIENTE" : "CONTADO"}</span>
          ${c.EMPRESA ? `<span class="text-muted ms-1" style="font-size:12.5px;">${depEsc(c.EMPRESA)}</span>` : ""}
          ${alta ? `<span class="text-muted ms-1" style="font-size:11.5px;">· cliente desde ${alta}</span>` : ""}
        </div>
        <div class="d-flex gap-1 flex-wrap">
          <button type="button" class="btn btn-sm btn-primary" onclick="depFichaNuevaCotizacion()">📝 Nueva cotización</button>
          <button type="button" class="btn btn-sm btn-outline-secondary" onclick="depAbrirEstadoCuenta()" ${esCredito ? "" : "disabled title='Solo para clientes a crédito'"}>📄 Estado de cuenta A4</button>
          <button type="button" class="btn btn-sm btn-outline-secondary" onclick="depFichaEditar()">✏️ Editar datos</button>
        </div>
      </div>
      <div class="d-grid mt-2" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px 14px;font-size:13px;">
        ${fila("Dirección", depEsc([c.DIRECCION, loc].filter(Boolean).join(" — ")))}
        ${fila("Teléfono", tel ? depEsc(tel) + (wa ? ` <a href="${wa}" target="_blank" rel="noopener" style="font-size:12px;">WhatsApp ↗</a>` : "") : "")}
        ${fila("CUIT / DNI", depEsc(c.DNI || ""))}
        ${fila("Email", c.EMAIL ? `<a href="mailto:${depEsc(c.EMAIL)}">${depEsc(c.EMAIL)}</a>` : "")}
      </div>
      ${c.NOTAS ? `<div class="mt-2" style="font-size:12.5px;background:#fffbeb;border-radius:6px;padding:6px 8px;white-space:pre-wrap;">📌 ${depEsc(c.NOTAS)}</div>` : ""}
      ${!c.DIRECCION && !tel ? `<div class="mt-2 text-muted" style="font-size:12px;">Faltan la dirección y el teléfono: completalos con “Editar datos” para que salgan en las cotizaciones y boletas.</div>` : ""}
    </div>
    <div class="card p-2 mb-3" id="depFichaCotizaciones">
      <div class="d-flex justify-content-between align-items-center px-1"><b style="font-size:13px;">Cotizaciones</b><span class="text-muted" style="font-size:12px;" id="depFichaCotEstado">Cargando...</span></div>
      <div id="depFichaCotLista"></div>
    </div>`;
  depFichaCargarCotizaciones(c);
}

async function depFichaCargarCotizaciones(c) {
  const lista = document.getElementById("depFichaCotLista");
  const estado = document.getElementById("depFichaCotEstado");
  if (!lista) return;
  try {
    let r = await depPost("cotizaciones", { clienteId: c.CLIENTE_ID, limite: 15 }, { timeoutMs: 15000 });
    // las armadas escribiendo el nombre a mano no tienen el ID del cliente
    if (r.success && !r.cotizaciones.length && c.NOMBRE) {
      const porNombre = await depPost("cotizaciones", { q: c.NOMBRE, limite: 15 }, { timeoutMs: 15000 });
      if (porNombre.success) r = { success: true, cotizaciones: porNombre.cotizaciones.filter(x => String(x.cliente).trim().toLowerCase() === String(c.NOMBRE).trim().toLowerCase()) };
    }
    if (!r.success) throw new Error(r.message);
    const cot = r.cotizaciones || [];
    estado.textContent = cot.length ? cot.length + (cot.length === 15 ? "+ " : " ") + "en total" : "";
    lista.innerHTML = cot.length ? `<table class="table table-sm mb-0" style="font-size:12.5px;"><tbody>${cot.map(x => `<tr>
      <td class="mono">${depEsc(x.cotizacionId)}</td><td>${depFechaCorta(x.fecha)}</td><td class="text-end money">${depMonto(x.total, x.moneda)}</td>
      <td>${depBadgeEstadoCot(x.estado)}</td>
      <td class="text-end" style="white-space:nowrap;"><button class="btn btn-sm btn-link p-0" onclick="imprimirCotizacionId('${depEsc(x.cotizacionId)}')" title="Imprimir">🖨️</button>
      ${x.estado !== "CONVERTIDA" && x.estado !== "ANULADA" ? `<button class="btn btn-sm btn-link p-0 ms-2" onclick="cerrarModalDetalleCliente(); pasarCotizacionAlPOS('${depEsc(x.cotizacionId)}')">Al POS</button>` : ""}</td>
    </tr>`).join("")}</tbody></table>` : `<div class="text-muted px-1 py-1" style="font-size:12.5px;">Todavía no tiene cotizaciones.</div>`;
  } catch (e) {
    estado.textContent = "";
    lista.innerHTML = `<div class="text-muted px-1" style="font-size:12px;">No se pudieron cargar (necesita conexión).</div>`;
  }
}

function depFichaNuevaCotizacion() {
  const data = depFichaCliente();
  if (!data) return;
  const c = data.cliente;
  abrirEditorCotizacion({
    clienteId: c.CLIENTE_ID,
    cliente: { nombre: c.NOMBRE, direccion: c.DIRECCION, localidad: c.LOCALIDAD, provincia: c.PROVINCIA, telefono: c.TELEFONO, cuit: c.DNI, email: c.EMAIL, empresa: c.EMPRESA },
    items: []
  });
}

function depFichaEditar() {
  const data = depFichaCliente();
  if (!data) return;
  const id = data.cliente.CLIENTE_ID;
  // la lista guardada puede no tener las columnas nuevas (localidad, email, notas): se completa con la ficha
  if (typeof clientesGlobal !== "undefined") {
    const enLista = clientesGlobal.find(x => String(x.CLIENTE_ID) === String(id));
    if (enLista) Object.assign(enLista, data.cliente); else clientesGlobal.push({ ...data.cliente });
  }
  cerrarModalDetalleCliente();
  abrirModalEditarCliente(id);
}

/* ---------- estado de cuenta A4 ---------- */

function depMovimientosCuenta(data) {
  const movs = [];
  const fecha = (f) => { const d = new Date(f); return isNaN(d) ? new Date(0) : d; };
  (data.pedidos || []).forEach(p => {
    const usd = String(p.MONEDA || "ARS").toUpperCase() === "USD";
    const t = Number(p.TOTAL) || 0;
    movs.push({ fecha: fecha(p.FECHA), concepto: "Pedido " + (p.PEDIDO_ID || "") + (p.ESTADO ? " (" + String(p.ESTADO).toLowerCase() + ")" : "") + (usd && p.TIPO_CAMBIO ? " · dólar " + depPesos(p.TIPO_CAMBIO) : ""),
      debeArs: usd ? 0 : t, debeUsd: usd ? t : 0, haberArs: 0, haberUsd: 0 });
  });
  (data.deudasExtra || []).forEach(d => {
    const usd = String(d.MONEDA || "ARS").toUpperCase() === "USD";
    const m = Number(d.MONTO) || 0;
    movs.push({ fecha: fecha(d.FECHA), concepto: d.CONCEPTO || "Cargo", debeArs: usd ? 0 : m, debeUsd: usd ? m : 0, haberArs: 0, haberUsd: 0 });
  });
  (data.pagos || []).forEach(p => {
    const mon = String(p.MONEDA_PAGO || "ARS").toUpperCase();
    const forma = String(p.FORMA_PAGO || "").toLowerCase();
    let concepto = "Pago " + (mon === "USD" ? depUsd(p.MONTO_PAGO) : depPesos(p.MONTO_PAGO)) + (forma ? " · " + forma : "");
    if (Number(p.TIPO_CAMBIO) > 0 && Number(p.MONTO_APLICADO_ARS) > 0 && Number(p.MONTO_APLICADO_USD) > 0) concepto += " · dólar " + depPesos(p.TIPO_CAMBIO);
    else if (Number(p.TIPO_CAMBIO) > 0 && ((mon === "USD" && Number(p.MONTO_APLICADO_ARS) > 0) || (mon === "ARS" && Number(p.MONTO_APLICADO_USD) > 0))) concepto += " · dólar " + depPesos(p.TIPO_CAMBIO);
    if (p.OBSERVACIONES) concepto += " · " + p.OBSERVACIONES;
    movs.push({ fecha: fecha(p.FECHA), concepto, debeArs: 0, debeUsd: 0, haberArs: Number(p.MONTO_APLICADO_ARS) || 0, haberUsd: Number(p.MONTO_APLICADO_USD) || 0 });
  });
  movs.sort((a, b) => a.fecha - b.fecha);
  let sa = 0, su = 0;
  movs.forEach(m => { sa += m.debeArs - m.haberArs; su += m.debeUsd - m.haberUsd; m.saldoArs = Math.round(sa); m.saldoUsd = Math.round(su * 100) / 100; });
  return movs;
}

function depAbrirEstadoCuenta() {
  const data = depFichaCliente();
  if (!data) return;
  const movs = depMovimientosCuenta(data);
  const primero = movs.length ? movs[0].fecha : new Date();
  const iso = (d) => new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  depModal("modalEstadoCuenta", "Estado de cuenta · " + data.cliente.NOMBRE, `
    <div class="row g-2">
      <div class="col-6"><label class="form-label" style="font-size:12px;">Desde</label><input type="date" class="form-control form-control-sm" id="ecDesde" value="${iso(primero)}"></div>
      <div class="col-6"><label class="form-label" style="font-size:12px;">Hasta</label><input type="date" class="form-control form-control-sm" id="ecHasta" value="${iso(new Date())}"></div>
    </div>
    <div class="text-muted mt-2" style="font-size:12px;">Los movimientos anteriores a “Desde” se resumen en el saldo anterior. ${movs.length} movimiento(s) en total.</div>
  `, `<button type="button" class="btn btn-outline-secondary" data-cerrar>Cancelar</button>
      <button type="button" class="btn btn-primary" onclick="imprimirEstadoCuenta()">Imprimir</button>`, 420);
}

function imprimirEstadoCuenta() {
  const data = depFichaCliente();
  if (!data) return;
  const desdeTxt = (document.getElementById("ecDesde") || {}).value;
  const hastaTxt = (document.getElementById("ecHasta") || {}).value;
  const desde = desdeTxt ? new Date(desdeTxt + "T00:00:00") : new Date(0);
  const hasta = hastaTxt ? new Date(hastaTxt + "T23:59:59") : new Date();
  const m = document.getElementById("modalEstadoCuenta"); if (m) m.classList.remove("show");
  depImprimirA4(htmlEstadoCuenta(data, desde, hasta));
}

function htmlEstadoCuenta(data, desde, hasta) {
  const todos = depMovimientosCuenta(data);
  const anteriores = todos.filter(x => x.fecha < desde);
  const periodo = todos.filter(x => x.fecha >= desde && x.fecha <= hasta);
  const ant = anteriores.length ? anteriores[anteriores.length - 1] : { saldoArs: 0, saldoUsd: 0 };
  const hayUsd = todos.some(x => x.debeUsd || x.haberUsd);
  const fin = periodo.length ? periodo[periodo.length - 1] : ant;
  const c = data.cliente;
  const celda = (v, usd) => v ? (usd ? depUsd(v) : depPesos(v)) : "";
  const filas = periodo.map(x => `<tr>
    <td>${depFechaCorta(x.fecha)}</td><td>${depEsc(x.concepto)}</td>
    <td class="der">${celda(x.debeArs)}</td><td class="der">${celda(x.haberArs)}</td><td class="der"><b>${depPesos(x.saldoArs)}</b></td>
    ${hayUsd ? `<td class="der">${celda(x.debeUsd, true)}</td><td class="der">${celda(x.haberUsd, true)}</td><td class="der"><b>${depUsd(x.saldoUsd)}</b></td>` : ""}
  </tr>`).join("");
  return `<div class="a4doc">
    ${depCabeceraA4("ESTADO DE CUENTA", "", ["Del <b>" + depFechaCorta(desde.getTime() ? desde : (todos[0] || {}).fecha || new Date()) + "</b> al <b>" + depFechaCorta(hasta) + "</b>", "Emitido: " + depFechaCorta(new Date())])}
    ${depBloqueClienteA4({ nombre: c.NOMBRE, empresa: c.EMPRESA, direccion: c.DIRECCION, localidad: c.LOCALIDAD, provincia: c.PROVINCIA, telefono: c.TELEFONO, cuit: c.DNI, email: c.EMAIL })}
    <table class="items" style="font-size:8.8pt;">
      <thead>
        ${hayUsd ? `<tr><th rowspan="2" style="width:20mm;">Fecha</th><th rowspan="2">Concepto</th><th colspan="3" class="cen">Pesos</th><th colspan="3" class="cen">Dólares</th></tr>` : ""}
        <tr>${hayUsd ? "" : `<th style="width:20mm;">Fecha</th><th>Concepto</th>`}<th class="der">Debe</th><th class="der">Haber</th><th class="der">Saldo</th>${hayUsd ? `<th class="der">Debe</th><th class="der">Haber</th><th class="der">Saldo</th>` : ""}</tr>
      </thead>
      <tbody>
        <tr style="background:#f6f6f6;"><td></td><td><i>Saldo anterior</i></td><td></td><td></td><td class="der"><b>${depPesos(ant.saldoArs)}</b></td>${hayUsd ? `<td></td><td></td><td class="der"><b>${depUsd(ant.saldoUsd)}</b></td>` : ""}</tr>
        ${filas || `<tr><td colspan="${hayUsd ? 8 : 5}" class="cen" style="color:#777;">Sin movimientos en el período</td></tr>`}
      </tbody>
    </table>
    <table class="totales" style="width:90mm;">
      <tr class="total"><td>SALDO EN PESOS</td><td class="der">${depPesos(Math.max(0, fin.saldoArs))}</td></tr>
      ${hayUsd ? `<tr class="total"><td>SALDO EN DÓLARES</td><td class="der">${depUsd(Math.max(0, fin.saldoUsd))}</td></tr>` : ""}
      ${fin.saldoArs < 0 || fin.saldoUsd < 0 ? `<tr><td colspan="2" style="font-size:8.5pt;color:#555;">Saldo a favor del cliente: ${fin.saldoArs < 0 ? depPesos(-fin.saldoArs) : ""} ${fin.saldoUsd < 0 ? depUsd(-fin.saldoUsd) : ""}</td></tr>` : ""}
    </table>
    <div class="pie">Ante cualquier diferencia, comunicate con nosotros · ${depEsc((depCfgDocs().nombre) || "")}</div>
  </div>`;
}

/* ---------- arranque ---------- */
window.addEventListener("load", () => depCuandoPanelListo(() => {
  const orig = window.abrirModalDetalleCliente;
  if (typeof orig !== "function" || orig._depFicha) return;
  const envuelta = async function () {
    const viejo = document.getElementById("depFicha"); if (viejo) viejo.remove();
    const r = await orig.apply(this, arguments);
    try { depRenderFicha(); } catch (e) { console.error("ficha del cliente:", e); }
    return r;
  };
  envuelta._depFicha = true;
  window.abrirModalDetalleCliente = envuelta;
}));
