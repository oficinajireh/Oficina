/* =====================================================================
   facturacion-arca.js — factura electrónica con ARCA (opcional).

   Solo funciona en la app de escritorio (window.veekpos.arca): ahí vive
   el certificado y es la que habla con ARCA (ver desktop/src/arca/).
   En la versión web este archivo no hace nada.

   Qué agrega al panel:
   - Configuración → Facturación: datos fiscales, certificado, prueba.
   - POS → Finalizar venta: "Facturar" + datos del cliente (opcional).
   - Recibo: estado de la factura (emitida / en cola) e impresión fiscal.
   - Historial de ventas y detalle de pedidos: "Facturar" / "Imprimir factura".
   - Anular una venta facturada emite la nota de crédito sola.
===================================================================== */
(function () {
  "use strict";

  const disponible = () => !!(window.veekpos && typeof window.veekpos.arca === "function");
  const esc = (t) => (typeof escapeHtml === "function" ? escapeHtml(String(t === undefined || t === null ? "" : t)) : String(t || ""));
  const plata = (n) => "$" + Number(n || 0).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const aviso = (m, t) => (typeof toast === "function" ? toast(m, t) : console.log(m));

  const CONDICIONES = [
    [5, "Consumidor Final"], [1, "IVA Responsable Inscripto"], [6, "Responsable Monotributo"],
    [4, "IVA Sujeto Exento"], [7, "Sujeto No Categorizado"], [13, "Monotributista Social"],
    [16, "Monotributo Trabajador Independiente Promovido"], [15, "IVA No Alcanzado"],
    [10, "IVA Liberado – Ley 19.640"], [9, "Cliente del Exterior"], [8, "Proveedor del Exterior"]
  ];
  const NOMBRE_CONDICION = Object.fromEntries(CONDICIONES);
  const CLASE_A = [1, 6, 13, 16];
  const LIMITE_CF = 10000000;
  const DOC_NOMBRE = { 80: "CUIT", 86: "CUIL", 96: "DNI", 99: "" };
  const NOMBRE_ESTADO = { EMITIDA: "Emitida", PENDIENTE: "En cola", EMITIENDO: "Confirmando con ARCA", RECHAZADA: "Rechazada", ANULADA: "Cancelada" };

  let estado = null;          // estado de la facturación (de esta caja o de la caja servidor)
  let rolRed = null;
  let captura = null;         // lo elegido en "Finalizar venta" para la venta en curso
  const cache = {};           // ref → factura pública (con .notaCredito si hay)

  async function llamar(accion, datos) {
    try { return await window.veekpos.arca(accion, datos || {}); }
    catch (e) { return { success: false, message: String(e.message || e) }; }
  }

  function leerLS(clave, defecto) { try { const v = localStorage.getItem(clave); return v ? JSON.parse(v) : defecto; } catch (e) { return defecto; } }
  function escribirLS(clave, v) { try { localStorage.setItem(clave, JSON.stringify(v)); } catch (e) {} }

  /* ------------------------------------------------------------------
     Estado
  ------------------------------------------------------------------ */
  async function refrescarEstado() {
    if (!disponible()) return null;
    try { const r = await window.veekpos.obtenerRolRed(); rolRed = r && r.rol; } catch (e) {}
    const r = await llamar("estado");
    if (r && r.success) { estado = r.estado; escribirLS("veekpos_arca_estado", estado); }
    else if (!estado) estado = leerLS("veekpos_arca_estado", null); // caja cliente sin el servidor: lo último conocido
    aplicarEstadoUI();
    return estado;
  }

  const activa = () => !!(estado && estado.activa && estado.lista);

  function claseFactura(condicionReceptor, doc) {
    if (!estado || estado.condicion !== "RI") return "C";
    return CLASE_A.includes(Number(condicionReceptor)) && String(doc || "").replace(/\D/g, "").length === 11 ? "A" : "B";
  }

  function aplicarEstadoUI() {
    const alic = document.getElementById("pmAlicuotaWrap");
    if (alic) alic.style.display = activa() && estado.condicion === "RI" ? "" : "none";
    const wrap = document.getElementById("mfvArcaWrap");
    if (wrap) wrap.style.display = activa() ? "" : "none";
  }

  /* ------------------------------------------------------------------
     Configuración
  ------------------------------------------------------------------ */
  async function cargarConfig() {
    const card = document.getElementById("cardFacturacionArca");
    if (!card) return;
    card.style.display = "block";
    const box = document.getElementById("arcaEstadoBox");
    const form = document.getElementById("arcaFormularioWrap");
    await refrescarEstado();

    if (rolRed === "cliente") {
      form.style.display = "none";
      box.innerHTML = estado && estado.activa
        ? `<div style="width:100%;"><div style="color:var(--green-600);font-weight:600;">Configurada en la caja servidor</div>
           <div class="text-muted" style="font-size:12.5px;margin-top:2px;">${esc(estado.razonSocial)} · CUIT ${esc(estado.cuit)} · Punto de venta ${esc(estado.ptoVta)} · ${estado.condicion === "RI" ? "Facturas A y B" : "Factura C"}${estado.ambiente !== "produccion" ? " · <b>PRUEBAS</b>" : ""}. Esta caja le pide las facturas a la caja servidor.</div></div>`
        : `<span class="text-muted">La facturación se configura una sola vez, en la caja servidor (ahí queda el certificado y la numeración). ${estado ? "Todavía no está activada ahí." : "No hay conexión con la caja servidor."}</span>`;
      return;
    }

    form.style.display = "block";
    if (!estado) { box.innerHTML = `<span class="text-muted">No se pudo leer la configuración.</span>`; return; }
    const set = (id, v) => { const el = document.getElementById(id); if (el) { if (el.type === "checkbox") el.checked = !!v; else el.value = v === undefined || v === null ? "" : v; } };
    set("arcaActiva", estado.activa);
    set("arcaCuit", estado.cuit);
    set("arcaCondicion", estado.condicion);
    set("arcaRazonSocial", estado.razonSocial);
    set("arcaPtoVta", estado.ptoVta);
    set("arcaDomicilio", estado.domicilio);
    set("arcaIibb", estado.iibb);
    set("arcaInicioActividades", estado.inicioActividades);
    set("arcaConcepto", String(estado.concepto || 1));
    set("arcaAlicuotaDefault", String(estado.alicuotaDefault || 21));
    set("arcaAmbiente", estado.ambiente);
    set("arcaFacturarPorDefecto", estado.facturarPorDefecto);
    actualizarVisibilidadAlicuota();
    pintarEstadoConfig();
  }

  function actualizarVisibilidadAlicuota() {
    const w = document.getElementById("arcaAlicuotaWrap");
    const c = document.getElementById("arcaCondicion");
    if (w && c) w.style.display = c.value === "RI" ? "" : "none";
  }

  function pintarEstadoConfig() {
    const box = document.getElementById("arcaEstadoBox");
    const certBox = document.getElementById("arcaCertificadoBox");
    if (!box || !estado) return;
    const cert = estado.certificado;
    if (certBox) {
      if (cert && !cert.error) {
        const vence = new Date(cert.hasta);
        const dias = Math.floor((vence - Date.now()) / 86400000);
        certBox.innerHTML = `<span style="color:var(--green-600);font-weight:600;">Certificado cargado</span> — ${esc(cert.titular)}${cert.cuit ? " · CUIT " + esc(cert.cuit) : ""} · emitido por ${esc(cert.emisor)} · vence el ${vence.toLocaleDateString("es-AR")}${dias < 30 ? ` <b style="color:var(--red-500);">(${dias < 0 ? "vencido" : "faltan " + dias + " días"})</b>` : ""}`;
      } else if (cert && cert.error) {
        certBox.innerHTML = `<span style="color:var(--red-500);">${esc(cert.error)}</span>`;
      } else if (estado.hayPedidoPendiente) {
        certBox.innerHTML = `Pedido de certificado generado. Subilo a ARCA y cargá acá el certificado (.crt) que te devuelve.`;
      } else {
        certBox.textContent = "Sin certificado. Generá el pedido (paso 1), subilo a ARCA y cargá el certificado que te da (paso 2).";
      }
    }
    const partes = [];
    if (estado.activa && estado.lista) {
      partes.push(`<div style="color:var(--green-600);font-weight:600;">Activada${estado.ambiente !== "produccion" ? " — en PRUEBAS (homologación), sin validez fiscal" : ""}</div>`);
      partes.push(`<div class="text-muted" style="font-size:12.5px;margin-top:2px;">${estado.condicion === "RI" ? "Responsable Inscripto: Factura A a inscriptos y monotributistas con CUIT, Factura B al resto." : "Monotributo: Factura C."} Punto de venta ${esc(estado.ptoVta)}.</div>`);
    } else {
      partes.push(`<span class="text-muted">Desactivada — el POS no ofrece facturar.</span>`);
    }
    if (estado.pendientes) partes.push(`<div style="margin-top:4px;">${estado.pendientes} comprobante(s) en cola, se emiten solos cuando hay conexión con ARCA.</div>`);
    if (estado.conError) partes.push(`<div style="margin-top:4px;color:var(--red-500);">${estado.conError} comprobante(s) rechazados por ARCA — revisalos en "Ver comprobantes".</div>`);
    box.innerHTML = `<div style="width:100%;">${partes.join("")}</div>`;
  }

  function leerFormulario() {
    const v = (id) => { const el = document.getElementById(id); return el ? (el.type === "checkbox" ? el.checked : el.value.trim()) : undefined; };
    return {
      activa: v("arcaActiva"), cuit: v("arcaCuit"), condicion: v("arcaCondicion"), razonSocial: v("arcaRazonSocial"),
      ptoVta: v("arcaPtoVta"), domicilio: v("arcaDomicilio"), iibb: v("arcaIibb"), inicioActividades: v("arcaInicioActividades"),
      concepto: Number(v("arcaConcepto")) || 1, alicuotaDefault: v("arcaAlicuotaDefault"), ambiente: v("arcaAmbiente"),
      facturarPorDefecto: v("arcaFacturarPorDefecto")
    };
  }

  async function guardarConfig() {
    const datos = leerFormulario();
    const anterior = estado ? estado.ambiente : "homologacion";
    const guardar = async () => {
      const r = await llamar("guardarConfig", datos);
      if (r.config) { estado = r.config; escribirLS("veekpos_arca_estado", estado); }
      aplicarEstadoUI();
      await cargarConfig();
      aviso(r.success ? "Configuración de facturación guardada" : r.message, r.success ? "success" : "error");
    };
    if (datos.ambiente === "produccion" && anterior !== "produccion" && datos.activa) {
      confirmarAccion("A partir de ahora las facturas son REALES y quedan informadas en ARCA. Usá el certificado de producción (el de pruebas no sirve). ¿Pasar a producción?", guardar, "Facturación en producción");
      return;
    }
    await guardar();
  }

  function descargarTexto(nombre, texto, tipo) {
    const blob = new Blob([texto], { type: tipo || "text/plain" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nombre;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  async function generarPedidoCertificado() {
    const datos = leerFormulario();
    const hacer = async () => {
      aviso("Generando la clave… puede tardar unos segundos", "success");
      const r = await llamar("generarPedidoCertificado", { cuit: datos.cuit, razonSocial: datos.razonSocial });
      if (!r.success) { aviso(r.message, "error"); return; }
      descargarTexto("pedido-certificado-arca-" + String(datos.cuit).replace(/\D/g, "") + ".csr", r.csrPem, "application/pkcs10");
      aviso("Pedido de certificado descargado: subilo a ARCA", "success");
      cargarConfig();
    };
    if (estado && estado.certificado && !estado.certificado.error) {
      confirmarAccion("Ya hay un certificado cargado. Generar un pedido nuevo no lo borra: el actual sigue funcionando hasta que cargues el nuevo. ¿Continuar?", hacer, "Nuevo pedido de certificado");
      return;
    }
    await hacer();
  }

  function cargarCertificado(input) {
    const archivo = input.files && input.files[0];
    input.value = "";
    if (!archivo) return;
    const lector = new FileReader();
    lector.onload = async () => {
      let texto = String(lector.result || "");
      // ARCA entrega PEM; por las dudas se acepta también DER (binario)
      if (!/BEGIN CERTIFICATE/.test(texto)) {
        try {
          const bin = new Uint8Array(await archivo.arrayBuffer());
          let s = ""; bin.forEach(b => { s += String.fromCharCode(b); });
          texto = "-----BEGIN CERTIFICATE-----\n" + btoa(s).replace(/(.{64})/g, "$1\n") + "\n-----END CERTIFICATE-----\n";
        } catch (e) {}
      }
      const r = await llamar("cargarCertificado", { certPem: texto });
      aviso(r.success ? "Certificado cargado" : r.message, r.success ? "success" : "error");
      cargarConfig();
    };
    lector.readAsText(archivo);
  }

  async function probarConexion() {
    const out = document.getElementById("arcaPruebaResultado");
    if (out) out.innerHTML = `<span class="text-muted">Probando con ARCA…</span>`;
    // Lo que está escrito se guarda antes de probar (si no, se probaría lo viejo)
    const datos = leerFormulario();
    await llamar("guardarConfig", { ...datos, activa: estado ? estado.activa && datos.activa : false });
    const r = await llamar("probarConexion");
    if (!out) return;
    out.innerHTML = (r.pasos || []).map(p => `<div style="color:${p.ok ? "var(--green-600)" : "var(--red-500)"};">${p.ok ? "Bien" : "Falla"} — <b>${esc(p.paso)}</b>: ${esc(p.detalle)}</div>`).join("") || esc(r.message || "Sin respuesta");
    refrescarEstado().then(pintarEstadoConfig);
  }

  function mostrarAyudaCertificado() {
    abrirModal("Cómo sacar el certificado de ARCA", `
      <ol style="font-size:13.5px; padding-left:18px; line-height:1.55;">
        <li>Completá CUIT y razón social acá y tocá <b>1. Generar pedido de certificado</b>. Se descarga un archivo <code>.csr</code> (la clave privada queda guardada y cifrada en esta computadora; no sale de acá).</li>
        <li><b>Para pruebas (homologación):</b> entrá a ARCA con tu clave fiscal → servicio <b>"WSASS – Autogestión Certificados Homologación"</b> (si no lo tenés, agregalo desde el Administrador de Relaciones). Creá un certificado pegando el contenido del <code>.csr</code>, y después en "Crear autorización a servicio" autorizá el servicio <b>wsfe</b> para tu CUIT.</li>
        <li><b>Para producción:</b> servicio <b>"Administración de Certificados Digitales"</b> → agregá un alias (ej. <i>veekpos</i>) → subí el <code>.csr</code> → descargá el certificado. Después, en el <b>Administrador de Relaciones de Clave Fiscal</b> → Nueva relación → ARCA → WebServices → <b>Facturación Electrónica</b>, y elegí como representante el certificado (alias) que creaste.</li>
        <li>Creá un <b>punto de venta</b> del tipo <b>"Factura Electrónica – Monotributo/Responsable Inscripto – Web Services"</b> en "Administración de puntos de venta y domicilios". Ese número va en "Punto de venta".</li>
        <li>Tocá <b>2. Cargar certificado</b> y elegí el archivo que te dio ARCA. Después <b>Probar conexión</b>, activá y guardá.</li>
      </ol>
      <p class="text-muted" style="font-size:12.5px;margin:0;">El certificado de pruebas y el de producción son distintos: al pasar a producción hay que generar un pedido nuevo y repetir el paso 3. Los certificados duran 2 años.</p>`);
  }

  /* ------------------------------------------------------------------
     Modal genérico (se agrega al final del body → queda arriba de todo)
  ------------------------------------------------------------------ */
  function abrirModal(titulo, cuerpoHtml, acciones, ancho) {
    cerrarModal();
    const fondo = document.createElement("div");
    fondo.className = "product-modal-backdrop show";
    fondo.id = "arcaModalBackdrop";
    fondo.style.zIndex = "4000";
    fondo.innerHTML = `
      <div class="product-modal" style="width:min(${ancho || 560}px,94vw);">
        <div class="product-modal-header"><h6>${esc(titulo)}</h6><button type="button" data-cerrar>✕</button></div>
        <div class="product-modal-body" id="arcaModalCuerpo">${cuerpoHtml}</div>
        <div class="product-modal-actions" id="arcaModalAcciones"></div>
      </div>`;
    document.body.appendChild(fondo);
    fondo.querySelector("[data-cerrar]").onclick = cerrarModal;
    fondo.addEventListener("mousedown", (e) => { if (e.target === fondo) cerrarModal(); });
    const pie = fondo.querySelector("#arcaModalAcciones");
    (acciones || [{ texto: "Cerrar", clase: "btn-outline-secondary", accion: cerrarModal }]).forEach(a => {
      const b = document.createElement("button");
      b.className = "btn " + (a.clase || "btn-outline-secondary");
      b.textContent = a.texto;
      if (a.id) b.id = a.id;
      b.onclick = a.accion;
      pie.appendChild(b);
    });
    return fondo;
  }
  function cerrarModal() { const m = document.getElementById("arcaModalBackdrop"); if (m) m.remove(); }

  /* ------------------------------------------------------------------
     Datos del cliente (formulario reutilizable)
  ------------------------------------------------------------------ */
  function htmlReceptor(prefijo, inicial) {
    const r = inicial || {};
    const opciones = CONDICIONES.map(([v, n]) => `<option value="${v}" ${Number(r.condicionIva || 5) === v ? "selected" : ""}>${esc(n)}</option>`).join("");
    return `
      <div class="row g-2">
        <div class="col-6"><input type="text" class="form-control form-control-sm" id="${prefijo}Doc" inputmode="numeric" placeholder="DNI o CUIT (opcional)" value="${esc(r.doc || "")}"></div>
        <div class="col-6"><select class="form-select form-select-sm" id="${prefijo}Cond">${opciones}</select></div>
        <div class="col-12"><input type="text" class="form-control form-control-sm" id="${prefijo}Nombre" placeholder="Nombre o razón social" value="${esc(r.nombre || "")}"></div>
        <div class="col-12"><input type="text" class="form-control form-control-sm" id="${prefijo}Domicilio" placeholder="Domicilio (para Factura A)" value="${esc(r.domicilio || "")}"></div>
      </div>
      <div id="${prefijo}Tipo" style="font-size:12.5px;margin-top:6px;"></div>`;
  }

  function leerReceptor(prefijo) {
    const v = (s) => { const el = document.getElementById(prefijo + s); return el ? el.value.trim() : ""; };
    return { doc: v("Doc"), condicionIva: Number(v("Cond")) || 5, nombre: v("Nombre"), domicilio: v("Domicilio") };
  }

  /** Valida lo cargado; devuelve el mensaje de error o "" */
  function validarReceptor(r, total) {
    const doc = String(r.doc || "").replace(/\D/g, "");
    if (doc && !(doc.length === 11 || (doc.length >= 7 && doc.length <= 8))) return "El documento tiene que ser un DNI (7 u 8 números) o un CUIT (11 números)";
    if (estado && estado.condicion === "RI" && CLASE_A.includes(Number(r.condicionIva)) && doc.length !== 11) return "Para " + NOMBRE_CONDICION[r.condicionIva] + " hace falta el CUIT (Factura A)";
    if (!doc && Number(total) >= LIMITE_CF) return "Desde $" + LIMITE_CF.toLocaleString("es-AR") + " hay que identificar al cliente (DNI o CUIT)";
    if (claseFactura(r.condicionIva, doc) === "A" && !r.nombre) return "Para la Factura A completá la razón social del cliente";
    return "";
  }

  function conectarReceptor(prefijo, totalFn) {
    const doc = document.getElementById(prefijo + "Doc");
    const cond = document.getElementById(prefijo + "Cond");
    const nombre = document.getElementById(prefijo + "Nombre");
    const dom = document.getElementById(prefijo + "Domicilio");
    const tipo = document.getElementById(prefijo + "Tipo");
    const pintar = () => {
      const r = leerReceptor(prefijo);
      const clase = claseFactura(r.condicionIva, r.doc);
      const err = validarReceptor(r, totalFn ? totalFn() : 0);
      tipo.innerHTML = err
        ? `<span style="color:var(--red-500);">${esc(err)}</span>`
        : `<span class="text-muted">Se emite <b>Factura ${clase}</b>${r.doc ? "" : " a Consumidor Final"}${estado && estado.ambiente !== "produccion" ? " (de prueba)" : ""}.</span>`;
    };
    // Recordar clientes por documento (solo en esta computadora)
    doc.addEventListener("change", () => {
      const d = doc.value.replace(/\D/g, "");
      const guardado = (leerLS("veekpos_arca_clientes", {}) || {})[d];
      if (guardado) {
        if (!nombre.value) nombre.value = guardado.nombre || "";
        if (!dom.value) dom.value = guardado.domicilio || "";
        if (guardado.condicionIva) cond.value = String(guardado.condicionIva);
      } else if (d.length === 11 && Number(cond.value) === 5) {
        // Un CUIT de empresa (30/33/34) casi seguro es Responsable Inscripto
        if (/^3[034]/.test(d)) cond.value = "1";
      }
      pintar();
    });
    [doc, cond, nombre].forEach(el => el.addEventListener("input", pintar));
    cond.addEventListener("change", pintar);
    pintar();
    return pintar;
  }

  function recordarReceptor(r) {
    const d = String(r.doc || "").replace(/\D/g, "");
    if (!d) return;
    const todos = leerLS("veekpos_arca_clientes", {}) || {};
    todos[d] = { nombre: r.nombre, domicilio: r.domicilio, condicionIva: r.condicionIva };
    const claves = Object.keys(todos);
    if (claves.length > 500) delete todos[claves[0]];
    escribirLS("veekpos_arca_clientes", todos);
  }

  /* ------------------------------------------------------------------
     POS: bloque "Facturar" en Finalizar venta
  ------------------------------------------------------------------ */
  let mfvFacturar = false;

  function asegurarBloqueModal() {
    if (document.getElementById("mfvArcaWrap")) return;
    const ref = document.getElementById("mfvRecibidoWrap");
    if (!ref) return;
    const div = document.createElement("div");
    div.id = "mfvArcaWrap";
    div.style.display = "none";
    div.innerHTML = `
      <div class="mfv-label">Factura electrónica</div>
      <div class="discount-mode-row">
        <button type="button" class="discount-mode-btn" data-modo="SIN" id="mfvArcaNo" onclick="ArcaUI.elegirFacturar(false)">Sin factura</button>
        <button type="button" class="discount-mode-btn" data-modo="SI" id="mfvArcaSi" onclick="ArcaUI.elegirFacturar(true)">Facturar</button>
      </div>
      <div id="mfvArcaDatos" style="display:none;">${htmlReceptor("mfvArca")}</div>`;
    ref.parentNode.insertBefore(div, ref);
    conectarReceptor("mfvArca", totalTicket);
  }

  function totalTicket() {
    try {
      const subtotal = ticketPOS.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
      return calcularDescuentoPOS(subtotal).total;
    } catch (e) { return 0; }
  }

  function elegirFacturar(si) {
    mfvFacturar = !!si;
    const no = document.getElementById("mfvArcaNo"), b = document.getElementById("mfvArcaSi"), datos = document.getElementById("mfvArcaDatos");
    if (no) no.classList.toggle("active", !mfvFacturar);
    if (b) b.classList.toggle("active", mfvFacturar);
    if (datos) datos.style.display = mfvFacturar ? "" : "none";
  }

  /** Llamado al abrir "Finalizar venta" */
  function alAbrirModalVenta() {
    if (!disponible()) return;
    asegurarBloqueModal();
    aplicarEstadoUI();
    // Cada venta arranca limpia, con el "facturar" por defecto de Configuración
    ["Doc", "Nombre", "Domicilio"].forEach(s => { const el = document.getElementById("mfvArca" + s); if (el) el.value = ""; });
    const cond = document.getElementById("mfvArcaCond"); if (cond) cond.value = "5";
    elegirFacturar(activa() && !!estado.facturarPorDefecto);
    const doc = document.getElementById("mfvArcaDoc"); if (doc) doc.dispatchEvent(new Event("input"));
    refrescarEstado().then(() => { if (!activa()) elegirFacturar(false); });
  }

  /**
   * Llamado al confirmar la venta (antes de cerrar el modal). Devuelve false
   * si los datos del cliente no alcanzan para la factura (el modal sigue abierto).
   */
  function capturarModal(total) {
    captura = null;
    if (!disponible() || !activa() || !mfvFacturar) return true;
    const r = leerReceptor("mfvArca");
    const err = validarReceptor(r, total);
    if (err) {
      aviso(err, "error");
      const d = document.getElementById("mfvArcaDoc"); if (d) d.focus();
      return false;
    }
    captura = { facturar: true, receptor: r };
    return true;
  }

  function itemsParaFactura(items) {
    return (items || []).map(it => {
      let alicuota = it.ALICUOTA_IVA;
      if (alicuota === undefined || alicuota === "") {
        try {
          const p = (typeof productosPOS !== "undefined" ? productosPOS : []).find(x => String(x.CODIGO) === String(it.CODIGO));
          if (p) alicuota = p.ALICUOTA_IVA;
        } catch (e) {}
      }
      return { nombre: it.PRODUCTO || it.NOMBRE || "Producto", cantidad: Number(it.cantidad) || 0, precio: Number(it.PRECIO) || 0, alicuota: alicuota === undefined ? "" : alicuota, codigo: it.CODIGO || "" };
    });
  }

  /**
   * Llamado apenas se registró la venta en el POS (camino normal y Mercado Pago).
   * venta: { ref, items, total, fecha }
   */
  function trasVenta(venta) {
    const bloque = asegurarBloqueRecibo();
    const elegido = captura; captura = null;
    if (!disponible() || !elegido || !elegido.facturar || !venta || !venta.ref) { if (bloque) bloque.style.display = "none"; return; }
    const destino = typeof ultimaVentaImprimible !== "undefined" ? ultimaVentaImprimible : null;
    recordarReceptor(elegido.receptor);
    pintarRecibo({ estado: "EMITIENDO", ref: venta.ref }, true);
    const promesa = llamar("solicitarFactura", {
      ref: venta.ref,
      venta: { items: itemsParaFactura(venta.items), total: venta.total, fecha: (venta.fecha || new Date()).toISOString ? (venta.fecha || new Date()).toISOString() : venta.fecha },
      receptor: elegido.receptor
    }).then(r => {
      const f = r && r.success ? r.factura : { ref: venta.ref, estado: "RECHAZADA", error: (r && r.message) || "No se pudo pedir la factura" };
      cache[venta.ref] = f;
      if (destino) destino.factura = f;
      pintarRecibo(f, destino === (typeof ultimaVentaImprimible !== "undefined" ? ultimaVentaImprimible : null));
      if (f.estado === "EMITIDA") aviso(f.tipoNombre + " " + f.numero + " emitida", "success");
      else if (f.estado === "RECHAZADA") aviso("ARCA rechazó la factura: " + (f.error || ""), "error");
      else seguirFactura(venta.ref, destino);
      return f;
    });
    if (destino) { destino.facturaRef = venta.ref; destino.facturaPromesa = promesa; }
  }

  /**
   * La factura quedó pendiente (ARCA tardó, o se emite en la caja servidor):
   * se consulta cada tanto un par de minutos para actualizar el recibo. Si
   * sigue sin salir, queda en cola igual y se ve en el historial.
   */
  function seguirFactura(ref, destino) {
    let vueltas = 0;
    const tic = async () => {
      if (++vueltas > 40) return;
      const actual = cache[ref];
      if (actual && (actual.estado === "EMITIDA" || actual.estado === "RECHAZADA")) return; // ya llegó (ej. por el aviso del proceso principal)
      const r = await llamar("obtenerPorRefs", { refs: [ref] });
      const f = r && r.success && r.facturas[ref];
      if (f && (f.estado === "EMITIDA" || f.estado === "RECHAZADA")) {
        cache[ref] = f;
        if (destino) destino.factura = f;
        pintarRecibo(f, typeof ultimaVentaImprimible !== "undefined" && destino === ultimaVentaImprimible);
        document.querySelectorAll(".arca-celda").forEach(c => { if (c.dataset.arcaRef === ref) pintarCelda(c); });
        return;
      }
      setTimeout(tic, vueltas < 10 ? 3000 : 6000);
    };
    setTimeout(tic, 3000);
  }

  function asegurarBloqueRecibo() {
    let b = document.getElementById("receiptArca");
    if (b) return b;
    const ref = document.getElementById("receiptCambioWrap");
    if (!ref) return null;
    b = document.createElement("div");
    b.id = "receiptArca";
    b.style.cssText = "display:none;margin-top:10px;padding:8px 10px;border-radius:8px;font-size:13px;text-align:left;";
    ref.parentNode.insertBefore(b, ref.nextSibling);
    return b;
  }

  function pintarRecibo(f, esLaActual) {
    if (!esLaActual) return;
    const b = asegurarBloqueRecibo();
    if (!b) return;
    b.style.display = "block";
    if (f.estado === "EMITIDA") {
      b.style.background = "var(--green-50, #ecfdf5)";
      b.innerHTML = `<b>${esc(f.tipoNombre)} ${esc(f.numero)}</b><br><span class="text-muted">CAE ${esc(f.cae)} · vence ${esc(formatearFecha(f.caeVto))}</span>${f.ambiente !== "produccion" ? `<br><span style="color:var(--red-500);">Comprobante de prueba, sin validez fiscal</span>` : ""}`;
    } else if (f.estado === "EMITIENDO" && !f.error) {
      b.style.background = "var(--slate-100, #f1f5f9)";
      b.innerHTML = `<span class="text-muted">Emitiendo la factura con ARCA…</span>`;
    } else if (f.estado === "PENDIENTE" || f.estado === "EMITIENDO") {
      b.style.background = "var(--amber-50, #fffbeb)";
      b.innerHTML = `<b>Factura en cola</b><br><span class="text-muted">Se emite sola apenas haya conexión con ARCA. ${esc(f.error || "")}</span>`;
    } else {
      b.style.background = "var(--red-50, #fef2f2)";
      b.innerHTML = `<b style="color:var(--red-500);">Factura rechazada</b><br><span class="text-muted">${esc(f.error || "")}</span><br>Podés corregir los datos y reintentar desde el historial de ventas.`;
    }
  }

  /** Para imprimir el recibo: espera la factura un rato si todavía se está emitiendo */
  async function facturaParaImprimir(u) {
    if (!u) return null;
    if (u.factura && u.factura.estado === "EMITIDA") return u.factura;
    if (u.facturaPromesa) {
      aviso("Esperando la factura de ARCA para imprimirla…", "success");
      const f = await Promise.race([u.facturaPromesa, new Promise(r => setTimeout(() => r(null), 16000))]);
      if (f && f.estado === "EMITIDA") return f;
      if (u.factura && u.factura.estado === "EMITIDA") return u.factura; // llegó consultando
      aviso("La factura todavía no salió: se imprime el ticket común (la factura se puede imprimir después desde el historial)", "error");
    }
    return null;
  }

  /* ------------------------------------------------------------------
     Historial de ventas y pedidos
  ------------------------------------------------------------------ */
  const refDeVenta = (v) => String(v.CLIENTE_VENTA_ID || "").trim() || ("VEN:" + String(v.VENTA_ID || v.ID || ""));
  let relojCeldas = null;

  /** HTML de la celda de factura de una fila del historial (se completa después) */
  function celdaVenta(v, anulada) {
    if (!disponible()) return "";
    const vid = String(v.VENTA_ID || v.ID || "");
    if (/^RES-/i.test(vid)) return "";
    clearTimeout(relojCeldas);
    relojCeldas = setTimeout(actualizarCeldas, 120);
    return `<span class="arca-celda ms-1" data-arca-ref="${esc(refDeVenta(v))}" data-arca-vid="${esc(vid)}" data-arca-anulada="${anulada ? 1 : 0}"></span>`;
  }

  async function actualizarCeldas() {
    const celdas = Array.from(document.querySelectorAll(".arca-celda[data-arca-ref]"));
    if (!celdas.length) return;
    if (!estado) await refrescarEstado();
    const refs = [...new Set(celdas.map(c => c.dataset.arcaRef))];
    const faltan = refs.filter(r => !(r in cache));
    for (let i = 0; i < faltan.length; i += 200) {
      const r = await llamar("obtenerPorRefs", { refs: faltan.slice(i, i + 200) });
      if (r && r.success) faltan.slice(i, i + 200).forEach(ref => { cache[ref] = r.facturas[ref] || null; });
    }
    celdas.forEach(pintarCelda);
  }

  function pintarCelda(c) {
    const ref = c.dataset.arcaRef, anulada = c.dataset.arcaAnulada === "1";
    const f = cache[ref];
    const refJs = esc(JSON.stringify(ref));
    if (!f || f.estado === "ANULADA") {
      c.innerHTML = activa() && !anulada
        ? `<button class="btn btn-sm btn-outline-primary" onclick='ArcaUI.facturarVentaHistorial(${refJs}, ${esc(JSON.stringify(c.dataset.arcaVid))})' title="Emitir la factura electrónica de esta venta">Facturar</button>`
        : "";
      return;
    }
    if (f.estado === "EMITIDA") {
      const nc = f.notaCredito && f.notaCredito.estado === "EMITIDA" ? ` <span class="badge bg-secondary" title="${esc(f.notaCredito.tipoNombre + " " + f.notaCredito.numero)}">NC ${esc(f.notaCredito.numero)}</span>` : "";
      c.innerHTML = `<button class="btn btn-sm btn-outline-success" onclick='ArcaUI.imprimirPorRef(${refJs})' title="${esc(f.tipoNombre + " " + f.numero + " · CAE " + f.cae)}">${esc(f.clase || "")} ${esc(f.numero)}</button>${nc}`;
      return;
    }
    if (f.estado === "RECHAZADA") {
      c.innerHTML = `<button class="btn btn-sm btn-outline-danger" onclick='ArcaUI.facturarVentaHistorial(${refJs}, ${esc(JSON.stringify(c.dataset.arcaVid))})' title="${esc(f.error)}">Factura rechazada — corregir</button>`;
      return;
    }
    c.innerHTML = `<span class="badge bg-warning text-dark" title="${esc(f.error || "Se emite sola cuando haya conexión con ARCA")}">Factura en cola</span>`;
  }

  function ventaDelHistorial(vid) {
    try { if (typeof _ventasMapPOS !== "undefined" && _ventasMapPOS[vid]) return _ventasMapPOS[vid]; } catch (e) {}
    return null;
  }

  function facturarVentaHistorial(ref, vid) {
    const v = ventaDelHistorial(vid);
    if (!v) { aviso("No se encontró la venta", "error"); return; }
    let items = [];
    try { items = typeof v.CARRITO === "string" ? JSON.parse(v.CARRITO) : (v.CARRITO || []); } catch (e) {}
    const previa = cache[ref] && cache[ref].receptor;
    abrirFacturar({
      ref, titulo: "Facturar venta " + vid, items, total: Number(v.TOTAL || 0), fecha: v.FECHA,
      receptor: previa ? { doc: previa.docNro && previa.docNro !== "0" ? previa.docNro : "", condicionIva: previa.condicionIva, nombre: previa.nombre, domicilio: previa.domicilio } : null,
      formaPago: v.FORMA_PAGO || v.PAGO || ""
    });
  }

  /** Detalle de un pedido web: se agrega el bloque de la factura */
  async function enDetallePedido(pedido, detalle) {
    if (!disponible() || !pedido) return;
    if (!estado) await refrescarEstado();
    const cuerpo = document.getElementById("pedidoDetalleBody");
    if (!cuerpo) return;
    const ref = "PED-" + String(pedido.PEDIDO_ID);
    let f = cache[ref];
    if (f === undefined) {
      const r = await llamar("obtenerPorRefs", { refs: [ref] });
      f = cache[ref] = r && r.success ? (r.facturas[ref] || null) : null;
    }
    if (!pedidoDetalleActual || String(pedidoDetalleActual.pedido.PEDIDO_ID) !== String(pedido.PEDIDO_ID)) return;
    const viejo = document.getElementById("arcaPedidoBloque"); if (viejo) viejo.remove();
    const div = document.createElement("div");
    div.id = "arcaPedidoBloque";
    div.className = "config-preview-hint mt-3";
    const esUSD = String(pedido.MONEDA || "ARS").toUpperCase() === "USD";
    if (f && f.estado === "EMITIDA") {
      div.innerHTML = `<b style="white-space:nowrap;">${esc(f.tipoNombre)} ${esc(f.numero)}</b> · CAE ${esc(f.cae)}${f.notaCredito && f.notaCredito.estado === "EMITIDA" ? ` · anulada con ${esc(f.notaCredito.tipoNombre)} ${esc(f.notaCredito.numero)}` : ""}
        <button class="btn btn-sm btn-outline-success ms-2" onclick='ArcaUI.imprimirPorRef(${esc(JSON.stringify(ref))})'>Imprimir factura</button>
        ${!f.notaCredito ? `<button class="btn btn-sm btn-outline-danger ms-1" onclick='ArcaUI.notaCreditoPedido(${esc(JSON.stringify(ref))})'>Nota de crédito</button>` : ""}`;
    } else if (f && (f.estado === "PENDIENTE" || f.estado === "EMITIENDO")) {
      div.innerHTML = `<b>Factura en cola</b> — se emite sola cuando haya conexión con ARCA.`;
    } else if (!activa()) {
      return;
    } else if (esUSD) {
      div.innerHTML = `<span class="text-muted">Los pedidos en dólares no se facturan desde acá.</span>`;
    } else {
      div.innerHTML = `${f && f.estado === "RECHAZADA" ? `<span style="color:var(--red-500);">ARCA rechazó la factura: ${esc(f.error)}</span><br>` : ""}
        <button class="btn btn-sm btn-primary" id="btnArcaFacturarPedido">Facturar pedido</button>`;
      setTimeout(() => {
        const b = document.getElementById("btnArcaFacturarPedido");
        if (b) b.onclick = () => abrirFacturar({
          ref, titulo: "Facturar pedido " + pedido.PEDIDO_ID,
          items: (detalle || []).map(d => ({ CODIGO: d.CODIGO, PRODUCTO: d.PRODUCTO, PRECIO: Number(d.PRECIO) || 0, cantidad: Number(d.cantidad) || 0 })),
          total: Number(pedido.TOTAL || 0), fecha: new Date(),
          receptor: { doc: String(pedido.DNI || "").replace(/\D/g, ""), nombre: pedido.CLIENTE || "", domicilio: [pedido.DIRECCION, pedido.LOCALIDAD, pedido.PROVINCIA].filter(Boolean).join(", "), condicionIva: String(pedido.DNI || "").replace(/\D/g, "").length === 11 && /^3[034]/.test(String(pedido.DNI).replace(/\D/g, "")) ? 1 : 5 },
          formaPago: pedido.FORMA_PAGO || "",
          alTerminar: () => enDetallePedido(pedido, detalle)
        });
      }, 0);
    }
    div.innerHTML = `<div style="width:100%;">${div.innerHTML}</div>`;
    cuerpo.appendChild(div);
  }

  function notaCreditoPedido(ref) {
    confirmarAccion("Se emite una nota de crédito por el total de la factura de este pedido (la anula ante ARCA). ¿Continuar?", async () => {
      const r = await llamar("solicitarNotaCredito", { ref, motivo: "Anulación del pedido" });
      resultadoNC(r, ref);
      if (pedidoDetalleActual) enDetallePedido(pedidoDetalleActual.pedido, pedidoDetalleActual.detalle);
    }, "Nota de crédito");
  }

  /** Modal "Facturar" (historial y pedidos) */
  function abrirFacturar(o) {
    const lineas = (o.items || []).map(i => `<tr><td>${Number(i.cantidad)}x ${esc(i.PRODUCTO)}</td><td style="text-align:right;">${plata(Number(i.PRECIO) * Number(i.cantidad))}</td></tr>`).join("");
    abrirModal(o.titulo || "Facturar", `
      <table style="width:100%;font-size:13px;margin-bottom:6px;"><tbody>${lineas}</tbody>
        <tfoot><tr style="border-top:1px solid var(--slate-200);"><td><b>Total</b></td><td style="text-align:right;"><b>${plata(o.total)}</b></td></tr></tfoot></table>
      <div class="mfv-label">Cliente</div>
      ${htmlReceptor("arcaFac", o.receptor)}
      <div id="arcaFacResultado" style="margin-top:8px;"></div>`, [
      { texto: "Cancelar", clase: "btn-outline-secondary", accion: cerrarModal },
      { texto: "Emitir factura", clase: "btn-success fw-bold", id: "btnArcaEmitir", accion: () => emitirDesdeModal(o) }
    ], 480);
    conectarReceptor("arcaFac", () => o.total);
  }

  async function emitirDesdeModal(o) {
    const r = leerReceptor("arcaFac");
    const err = validarReceptor(r, o.total);
    const out = document.getElementById("arcaFacResultado");
    if (err) { out.innerHTML = `<span style="color:var(--red-500);">${esc(err)}</span>`; return; }
    const btn = document.getElementById("btnArcaEmitir");
    btn.disabled = true; btn.textContent = "Emitiendo…";
    recordarReceptor(r);
    const fecha = o.fecha ? new Date(o.fecha) : new Date();
    const res = await llamar("solicitarFactura", { ref: o.ref, venta: { items: itemsParaFactura(o.items), total: o.total, fecha: isNaN(fecha) ? new Date().toISOString() : fecha.toISOString() }, receptor: r });
    const f = res && res.success ? res.factura : null;
    if (f) cache[o.ref] = f;
    if (!f) { btn.disabled = false; btn.textContent = "Emitir factura"; out.innerHTML = `<span style="color:var(--red-500);">${esc((res && res.message) || "No se pudo emitir")}</span>`; return; }
    if (f.estado === "EMITIDA") {
      cerrarModal();
      aviso(f.tipoNombre + " " + f.numero + " emitida", "success");
      imprimirComprobante(f, o.formaPago);
    } else if (f.estado === "RECHAZADA") {
      btn.disabled = false; btn.textContent = "Reintentar";
      out.innerHTML = `<span style="color:var(--red-500);">ARCA la rechazó: ${esc(f.error)}</span>`;
      return;
    } else {
      cerrarModal();
      aviso("La factura quedó en cola: se emite sola cuando haya conexión con ARCA", "success");
    }
    document.querySelectorAll(`.arca-celda`).forEach(c => { if (c.dataset.arcaRef === o.ref) pintarCelda(c); });
    if (o.alTerminar) o.alTerminar();
  }

  /* ------------------------------------------------------------------
     Anulación → nota de crédito
  ------------------------------------------------------------------ */
  async function trasAnulacion(venta, motivo) {
    if (!disponible() || !venta) return;
    const refs = [refDeVenta(venta), "VEN:" + String(venta.VENTA_ID || venta.ID || "")];
    let ref = refs.find(r => cache[r] && cache[r].estado);
    if (!ref) {
      const r = await llamar("obtenerPorRefs", { refs });
      if (r && r.success) { refs.forEach(x => { cache[x] = r.facturas[x] || null; }); ref = refs.find(x => r.facturas[x]); }
    }
    if (!ref) return; // la venta no tenía factura
    const res = await llamar("solicitarNotaCredito", { ref, motivo });
    resultadoNC(res, ref);
  }

  function resultadoNC(res, ref) {
    delete cache[ref];
    if (!res || !res.success) { aviso("No se pudo emitir la nota de crédito: " + ((res && res.message) || ""), "error"); return; }
    if (res.cancelada) { aviso("La factura de esta venta todavía no se había emitido: se canceló", "success"); return; }
    const f = res.factura;
    if (!f) return;
    if (f.estado === "EMITIDA") { aviso(f.tipoNombre + " " + f.numero + " emitida", "success"); imprimirComprobante(f, ""); }
    else if (f.estado === "RECHAZADA") aviso("ARCA rechazó la nota de crédito: " + f.error, "error");
    else aviso("La nota de crédito quedó en cola: se emite sola al volver la conexión con ARCA", "success");
  }

  /* ------------------------------------------------------------------
     Impresión
  ------------------------------------------------------------------ */
  async function imprimirPorRef(ref) {
    let f = cache[ref];
    if (!f) { const r = await llamar("obtenerPorRefs", { refs: [ref] }); f = cache[ref] = r && r.success ? r.facturas[ref] : null; }
    if (!f || f.estado !== "EMITIDA") { aviso("No hay factura emitida para imprimir", "error"); return; }
    imprimirComprobante(f, "");
  }

  function imprimirComprobante(f, formaPago) {
    const items = (f.lineas || []).map(l => ({ PRODUCTO: l.nombre, PRECIO: Number(l.precio), cantidad: Number(l.cantidad) }));
    const bruto = items.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
    const dif = Math.round((bruto - Number(f.total)) * 100) / 100;
    const ajuste = Math.abs(dif) >= 0.01 ? { monto: dif, etiqueta: "" } : null;
    _ejecutarImpresion(f.numero, items, Number(f.total), formaPago || "", fechaDeComprobante(f), ajuste, f, null);
  }

  function fechaDeComprobante(f) {
    const s = String(f.fecha || "");
    return s.length === 8 ? new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8)), 12) : new Date();
  }
  function formatearFecha(aaaammdd) {
    const s = String(aaaammdd || "");
    return s.length === 8 ? s.slice(6, 8) + "/" + s.slice(4, 6) + "/" + s.slice(0, 4) : s;
  }

  function codigoTipo(f) { return String(f.cbteTipo || "").padStart(3, "0"); }

  function lineasEmisor(f) {
    const e = f.emisor || {};
    const inicio = e.inicioActividades ? e.inicioActividades.split("-").reverse().join("/") : "";
    return [
      e.razonSocial || "",
      e.domicilio || "",
      "CUIT: " + formatoCuit(e.cuit),
      e.iibb ? "Ing. Brutos: " + e.iibb : "",
      inicio ? "Inicio de actividades: " + inicio : "",
      e.condicion === "RI" ? "IVA Responsable Inscripto" : "Responsable Monotributo"
    ].filter(Boolean);
  }

  function formatoCuit(c) { const s = String(c || "").replace(/\D/g, ""); return s.length === 11 ? s.slice(0, 2) + "-" + s.slice(2, 10) + "-" + s.slice(10) : s; }

  function lineasReceptor(f) {
    const r = f.receptor || {};
    const out = [];
    if (r.docTipo && Number(r.docTipo) !== 99) out.push((DOC_NOMBRE[r.docTipo] || "Doc") + ": " + (Number(r.docTipo) === 80 || Number(r.docTipo) === 86 ? formatoCuit(r.docNro) : r.docNro));
    if (r.nombre) out.push(r.nombre);
    out.push("Condición IVA: " + (NOMBRE_CONDICION[r.condicionIva] || "Consumidor Final"));
    if (r.domicilio) out.push("Domicilio: " + r.domicilio);
    return out;
  }

  const textoPrueba = "COMPROBANTE DE PRUEBA - SIN VALIDEZ FISCAL";

  /** Imagen del QR fiscal (data URL) — con la librería del panel (vendor/qrcode.min.js) */
  function qrDataUrl(texto) {
    try {
      if (!window.QRCode) return "";
      const tmp = document.createElement("div");
      new QRCode(tmp, { text: texto, width: 300, height: 300, correctLevel: QRCode.CorrectLevel.M });
      const cv = tmp.querySelector("canvas");
      return cv ? cv.toDataURL("image/png") : "";
    } catch (e) { return ""; }
  }

  // ---- Ticket HTML (diálogo de impresión) ----
  function htmlEncabezado(f) {
    const prueba = f.ambiente !== "produccion" ? `<div class="th-center" style="font-size:9pt;margin:1mm 0;">*** ${textoPrueba} ***</div>` : "";
    return `
      <div class="th-center" style="font-size:8.5pt;font-weight:bold;">${lineasEmisor(f).map(esc).join("<br>")}</div>
      <hr class="th-sep-solid">
      ${prueba}
      <div class="th-center th-big">${esc(f.tipoNombre)}</div>
      <div class="th-center" style="font-size:8.5pt;">Cód. ${codigoTipo(f)} — ORIGINAL</div>
      <div class="th-center">Nº ${esc(f.numero)}</div>
      <div class="th-center" style="font-size:8.5pt;">Fecha: ${esc(formatearFecha(f.fecha))}</div>
      <hr class="th-sep">
      <div style="font-size:8.5pt;">${lineasReceptor(f).map(esc).join("<br>")}</div>
      ${f.asociado ? `<div style="font-size:8.5pt;margin-top:1mm;">Comprobante asociado: Cód. ${String(f.asociado.tipo).padStart(3, "0")} Nº ${String(f.asociado.ptoVta).padStart(5, "0")}-${String(f.asociado.nro).padStart(8, "0")}</div>` : ""}`;
  }

  /** Factura A: precios sin IVA por línea y el IVA discriminado */
  function htmlFilasA(f) {
    return (f.lineas || []).map(l => {
      const neto = Number(l.importe) / (1 + Number(l.alicuota) / 100);
      const unit = neto / (Number(l.cantidad) || 1);
      return `<tr><td colspan="2" style="padding-bottom:0;">${esc(l.nombre)} (IVA ${String(l.alicuota).replace(".", ",")}%)</td></tr>
        <tr><td>${Number(l.cantidad)} x ${plata(unit)}</td><td style="text-align:right;">${plata(neto)}</td></tr>`;
    }).join("");
  }
  function htmlTotalesA(f) {
    let h = `<tr><td>Neto gravado</td><td style="text-align:right;">${plata(f.impNeto)}</td></tr>`;
    (f.ivas || []).forEach(i => { h += `<tr><td>IVA ${String(i.alicuota).replace(".", ",")}%</td><td style="text-align:right;">${plata(i.importe)}</td></tr>`; });
    h += `<tr class="th-total-row"><td><strong>TOTAL</strong></td><td style="text-align:right;"><strong>${plata(f.total)}</strong></td></tr>`;
    return h;
  }
  function htmlPie(f) {
    const qr = f.qrUrl ? qrDataUrl(f.qrUrl) : "";
    return `
      <hr class="th-sep">
      ${f.clase === "B" ? `<div style="font-size:8.5pt;">Régimen de Transparencia Fiscal al Consumidor (Ley 27.743)<br>IVA contenido: ${plata(f.ivaContenido)}</div><hr class="th-sep">` : ""}
      <div style="font-size:9pt;">CAE: ${esc(f.cae)}<br>Vto. CAE: ${esc(formatearFecha(f.caeVto))}</div>
      ${qr ? `<div class="th-center" style="margin-top:2mm;"><img src="${qr}" style="width:32mm;height:32mm;" alt="QR ARCA"></div>` : ""}
      <div class="th-center" style="font-size:8pt;">Comprobante autorizado por ARCA</div>
      ${f.ambiente !== "produccion" ? `<div class="th-center" style="font-size:9pt;margin-top:1mm;">*** ${textoPrueba} ***</div>` : ""}`;
  }

  // ---- Ticket ESC/POS (impresora térmica directa) ----
  function escposEncabezado(b, f) {
    b.center(); b.bold();
    lineasEmisor(f).forEach(t => { b.text(t); b.feed(1); });
    b.boldOff(); b.left(); b.sepSolid(); b.center();
    if (f.ambiente !== "produccion") { b.text(textoPrueba); b.feed(1); }
    b.doubleH(); b.bold(); b.text(f.tipoNombre); b.feed(1); b.normal(); b.boldOff();
    b.text("Cod. " + codigoTipo(f) + " - ORIGINAL"); b.feed(1);
    b.bold(); b.text("N " + f.numero); b.feed(1); b.boldOff();
    b.text("Fecha: " + formatearFecha(f.fecha)); b.feed(1);
    b.left(); b.sep();
    lineasReceptor(f).forEach(t => { b.text(t); b.feed(1); });
    if (f.asociado) { b.text("Comprobante asociado: Cod. " + String(f.asociado.tipo).padStart(3, "0") + " N " + String(f.asociado.ptoVta).padStart(5, "0") + "-" + String(f.asociado.nro).padStart(8, "0")); b.feed(1); }
  }
  function escposFilasA(b, f) {
    (f.lineas || []).forEach(l => {
      const neto = Number(l.importe) / (1 + Number(l.alicuota) / 100);
      const unit = neto / (Number(l.cantidad) || 1);
      b.bold(); b.text(l.nombre + " (IVA " + String(l.alicuota).replace(".", ",") + "%)"); b.feed(1);
      b.row(Number(l.cantidad) + " x " + plata(unit), plata(neto));
    });
  }
  function escposTotalesA(b, f) {
    b.sep();
    b.row("Neto gravado", plata(f.impNeto));
    (f.ivas || []).forEach(i => b.row("IVA " + String(i.alicuota).replace(".", ",") + "%", plata(i.importe)));
  }
  function escposPie(b, f) {
    b.sep();
    if (f.clase === "B") {
      b.text("Regimen de Transparencia Fiscal al"); b.feed(1);
      b.text("Consumidor (Ley 27.743)"); b.feed(1);
      b.row("IVA contenido", plata(f.ivaContenido));
      b.sep();
    }
    b.bold(); b.text("CAE: " + f.cae); b.feed(1); b.text("Vto. CAE: " + formatearFecha(f.caeVto)); b.feed(1); b.boldOff();
    if (f.qrUrl) {
      b.center();
      b.push(qrEscPos(f.qrUrl));
      b.feed(1);
      b.text("Comprobante autorizado por ARCA"); b.feed(1);
      if (f.ambiente !== "produccion") { b.bold(); b.text(textoPrueba); b.feed(1); b.boldOff(); }
      b.left();
    }
  }
  /** QR nativo de la impresora (GS ( k, modelo 2) */
  function qrEscPos(texto) {
    const datos = Array.from(new TextEncoder().encode(texto));
    const largo = datos.length + 3;
    return [
      0x1D, 0x28, 0x6B, 0x04, 0x00, 0x31, 0x41, 0x32, 0x00,          // modelo 2
      0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x43, 0x05,                // tamaño del módulo
      0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x45, 0x31,                // corrección M
      0x1D, 0x28, 0x6B, largo & 0xFF, (largo >> 8) & 0xFF, 0x31, 0x50, 0x30, ...datos,
      0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x51, 0x30                 // imprimir
    ];
  }

  /* ------------------------------------------------------------------
     Comprobantes (listado)
  ------------------------------------------------------------------ */
  async function abrirComprobantes() {
    abrirModal("Comprobantes emitidos", `<div class="text-muted">Cargando…</div>`, null, 860);
    const r = await llamar("listar", { limite: 300 });
    const cuerpo = document.getElementById("arcaModalCuerpo");
    if (!cuerpo) return;
    const lista = (r && r.facturas) || [];
    if (!lista.length) { cuerpo.innerHTML = `<div class="text-muted">Todavía no hay comprobantes.</div>`; return; }
    lista.forEach(f => { if (f.ref) cache[f.ref] = cache[f.ref] || null; });
    cuerpo.innerHTML = `<div style="max-height:60vh;overflow:auto;"><table class="table table-sm" style="font-size:12.5px;">
      <thead><tr><th>Fecha</th><th>Comprobante</th><th>Cliente</th><th style="text-align:right;">Total</th><th>Estado</th><th></th></tr></thead>
      <tbody>${lista.map(f => {
        const r = f.receptor || {};
        const cliente = r.nombre || (Number(r.docTipo) === 99 || !r.docTipo ? "Consumidor Final" : (DOC_NOMBRE[r.docTipo] || "") + " " + r.docNro);
        const acciones = f.estado === "EMITIDA"
          ? `<button class="btn btn-sm btn-outline-secondary" data-arca-imp="${esc(f.id)}">Imprimir</button>`
          : f.estado === "RECHAZADA" || f.estado === "PENDIENTE"
            ? `<button class="btn btn-sm btn-outline-primary" data-arca-reint="${esc(f.id)}">Reintentar</button> <button class="btn btn-sm btn-outline-danger" data-arca-desc="${esc(f.id)}">Descartar</button>`
            : "";
        return `<tr>
          <td>${esc(formatearFecha(f.fecha))}</td>
          <td>${esc(f.tipoNombre)} <span style="white-space:nowrap;">${esc(f.numero || "")}</span>${f.ambiente !== "produccion" ? ' <span class="badge bg-secondary">prueba</span>' : ""}</td>
          <td>${esc(cliente)}</td>
          <td style="text-align:right;">${plata(f.total)}</td>
          <td title="${esc(f.error || "")}">${esc(NOMBRE_ESTADO[f.estado] || f.estado)}${f.error && f.estado !== "EMITIDA" ? `<div class="text-muted" style="font-size:11px;max-width:220px;">${esc(f.error)}</div>` : ""}</td>
          <td style="white-space:nowrap;">${acciones}</td></tr>`;
      }).join("")}</tbody></table></div>`;
    const porId = Object.fromEntries(lista.map(f => [f.id, f]));
    cuerpo.querySelectorAll("[data-arca-imp]").forEach(b => b.onclick = () => imprimirComprobante(porId[b.dataset.arcaImp], ""));
    cuerpo.querySelectorAll("[data-arca-reint]").forEach(b => b.onclick = async () => {
      b.disabled = true; b.textContent = "…";
      const res = await llamar("reintentar", { id: b.dataset.arcaReint });
      const f = res && res.factura;
      aviso(f && f.estado === "EMITIDA" ? f.tipoNombre + " " + f.numero + " emitida" : (f && f.error) || (res && res.message) || "Sigue pendiente", f && f.estado === "EMITIDA" ? "success" : "error");
      abrirComprobantes();
    });
    cuerpo.querySelectorAll("[data-arca-desc]").forEach(b => b.onclick = () => confirmarAccion("El comprobante deja de intentarse y la venta queda sin factura. ¿Descartarlo?", async () => {
      const res = await llamar("descartar", { id: b.dataset.arcaDesc });
      aviso(res.success ? "Descartado" : res.message, res.success ? "success" : "error");
      abrirComprobantes();
    }, "Descartar comprobante"));
  }

  /* ------------------------------------------------------------------
     Arranque
  ------------------------------------------------------------------ */
  function iniciar() {
    if (!disponible()) return;
    const card = document.getElementById("cardFacturacionArca");
    if (card) card.style.display = "block";
    const cond = document.getElementById("arcaCondicion");
    if (cond) cond.addEventListener("change", actualizarVisibilidadAlicuota);
    if (typeof window.mostrarConfigTab === "function") {
      const original = window.mostrarConfigTab;
      window.mostrarConfigTab = function (tabId, btn) { const r = original.apply(this, arguments); if (tabId === "facturacion") cargarConfig(); return r; };
    }
    if (typeof window.veekpos.alEmitirFactura === "function") {
      window.veekpos.alEmitirFactura((f) => {
        if (!f || !f.ref) return;
        if (f.tipoDoc === "NOTA_CREDITO") { delete cache[f.refOriginal]; return; }
        cache[f.ref] = f;
        try {
          if (typeof ultimaVentaImprimible !== "undefined" && ultimaVentaImprimible && ultimaVentaImprimible.facturaRef === f.ref) {
            ultimaVentaImprimible.factura = f;
            pintarRecibo(f, true);
          }
        } catch (e) {}
        document.querySelectorAll(".arca-celda").forEach(c => { if (c.dataset.arcaRef === f.ref) pintarCelda(c); });
      });
    }
    refrescarEstado();
    // Si la pestaña Facturación ya estaba elegida de antes
    setTimeout(() => { const p = document.getElementById("cfgTab-facturacion"); if (p && p.classList.contains("active")) cargarConfig(); }, 800);
  }

  window.ArcaUI = {
    disponible, activa, refrescarEstado, cargarConfig, guardarConfig, generarPedidoCertificado, cargarCertificado,
    probarConexion, mostrarAyudaCertificado, abrirComprobantes,
    alAbrirModalVenta, elegirFacturar, capturarModal, trasVenta, facturaParaImprimir,
    celdaVenta, facturarVentaHistorial, enDetallePedido, notaCreditoPedido, trasAnulacion, imprimirPorRef, imprimirComprobante,
    // impresión (los usan buildThermalHTML / buildThermalESCPOS)
    htmlEncabezado, htmlFilasA, htmlTotalesA, htmlPie, escposEncabezado, escposFilasA, escposTotalesA, escposPie,
    _cache: cache
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar);
  else iniciar();
})();
