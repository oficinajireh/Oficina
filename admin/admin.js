/* ===================================================================
   VeekPOS — panel admin (app logic v2)
   • All original Apps Script API calls preserved
   • Thermal print (POS80 80mm) added
   • Dashboard POS summary added
   • Responsive mobile nav sync added
=================================================================== */

// API_URL dinámica — se carga desde config.json (o SQLite en Electron)
// al iniciar. Permite instalar el mismo código para distintos clientes
// sin modificar nada manualmente. Sin valor por defecto a propósito:
// si no hay config.json ni api_url guardada, esta instalación todavía
// no está configurada y no debe poder hablar con el backend de otro
// cliente por error (ver comentario más abajo, cerca de DOMContentLoaded).
let API_URL = "";

/* ===================================================================
   ZONA HORARIA DEL NEGOCIO (Configuración → Negocio)
   Todas las fechas y horas que muestra el panel, y el "hoy" de los
   cálculos, usan la zona del negocio y no la de la computadora (si la
   PC tiene mal la zona, antes las horas salían corridas). Además, varios
   "hoy" se calculaban en UTC: después de las 21 h ya era "mañana".
=================================================================== */
const ZONA_HORARIA_DEFECTO = "America/Argentina/Buenos_Aires";
let _zonaHorariaPanel = (() => { try { return localStorage.getItem("vpos_zona_horaria") || ZONA_HORARIA_DEFECTO; } catch (e) { return ZONA_HORARIA_DEFECTO; } })();
function _zonaNegocio() { return _zonaHorariaPanel; }
function _fijarZonaNegocio(z) {
  if (!z) return;
  try { new Intl.DateTimeFormat("es-AR", { timeZone: z }); } catch (e) { return; }
  _zonaHorariaPanel = z;
  try { localStorage.setItem("vpos_zona_horaria", z); } catch (e) {}
}
/** "yyyy-mm-dd" de una fecha, en la zona del negocio */
function _fechaISOEnZona(fecha) {
  const d = fecha instanceof Date ? fecha : new Date(fecha);
  if (isNaN(d)) return "";
  try { return new Intl.DateTimeFormat("en-CA", { timeZone: _zonaNegocio(), year: "numeric", month: "2-digit", day: "2-digit" }).format(d); }
  catch (e) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
}
(function () {
  // toLocaleString / toLocaleDateString / toLocaleTimeString de fechas: si no
  // piden una zona, usan la del negocio.
  ["toLocaleString", "toLocaleDateString", "toLocaleTimeString"].forEach(m => {
    const original = Date.prototype[m];
    Date.prototype[m] = function (locales, opciones) {
      if (opciones && opciones.timeZone) return original.call(this, locales, opciones);
      try { return original.call(this, locales, Object.assign({}, opciones || {}, { timeZone: _zonaHorariaPanel })); }
      catch (e) { return original.call(this, locales, opciones); }
    };
  });
})();

/**
 * Reemplazo de fetch() para las llamadas al backend, con timeout
 * automático y reintentos seguros — pensado para conexiones
 * inestables (wifi que se corta un momento, 4G que titila, etc.).
 *
 * Antes, un corte momentáneo de conexión dejaba el fetch() colgado
 * esperando una respuesta que podía no llegar nunca — sin timeout, el
 * navegador puede tardar más de un minuto en darse por vencido solo.
 * Eso se sentía como que "el sistema se congeló", cuando en realidad
 * solo hacía falta reintentar.
 *
 * - LECTURAS (sin body / GET, la gran mayoría de las llamadas: cargar
 *   productos, pedidos, clientes, reportes, etc.): si fallan por
 *   timeout o corte de red, se reintentan solas 1 vez más — no tienen
 *   ningún efecto secundario, así que reintentar es 100% seguro.
 * - MUTACIONES (POST: crear/editar/eliminar algo, registrar un pago,
 *   etc.): acá NO se reintenta solo. Si la operación ya había llegado
 *   al servidor y solo se cortó la respuesta de vuelta, reintentar a
 *   ciegas podría duplicarla (dos clientes, dos pagos, dos deudas).
 *   Solo se les pone un límite de tiempo para que al menos fallen
 *   rápido y avisen con un error claro, en vez de dejar el botón
 *   "Guardando..." colgado para siempre — así el usuario puede
 *   revisar si se guardó o no antes de reintentar a mano.
 */
/* Pedidos al servidor en curso — los usan los botones "↻ Actualizar"
   para mostrar que están trabajando y cuándo terminaron (ver
   iniciarBotonesActualizar). */
const _redEnCurso = { n: 0, fallos: 0, cambios: 0 };
function _avisarCambioRed() { _redEnCurso.cambios++; }

async function fetchAPI(url, opciones = {}, config = {}) {
  _redEnCurso.n++; _avisarCambioRed();
  try {
    const r = await _fetchAPIBase(url, opciones, config);
    if (!r || !r.ok) _redEnCurso.fallos++;
    return r;
  } catch (e) {
    _redEnCurso.fallos++;
    throw e;
  } finally {
    _redEnCurso.n = Math.max(0, _redEnCurso.n - 1); _avisarCambioRed();
  }
}

async function _fetchAPIBase(url, opciones = {}, config = {}) {
  // El backend (code.gs) está armado sobre doGet con ?action=..., así
  // que la GRAN MAYORÍA de las mutaciones (actualizarProducto,
  // guardarProducto, sumarStockProducto, guardarPedidoAdmin, etc.)
  // viajan como GET, sin `method: "POST"` — el método HTTP solo no
  // alcanza para distinguir lectura de escritura acá. Antes, cualquier
  // mutación por GET se trataba como "lectura": timeout corto (10s,
  // insuficiente si el backend tarda por contención de lock — hasta
  // 20-30s, ver code.gs) y, peor, REINTENTO AUTOMÁTICO al vencer ese
  // timeout — exactamente lo que el comentario de más arriba dice que
  // nunca hay que hacer con una mutación (puede duplicarla), y lo que
  // generaba el caso real: el cambio sí llegaba a guardarse en Sheets,
  // pero el cliente ya había abortado pensando que era una lectura y
  // mostraba "error de conexión".
  //
  // Por eso, además del método HTTP, se mira el nombre de la acción en
  // la URL: en este backend las mutaciones siempre arrancan con un
  // verbo (guardar/actualizar/eliminar/crear/editar/registrar/...),
  // mientras que las lecturas son sustantivos (productos, pedidos,
  // clientes...). Si matchea un verbo de mutación, se trata como
  // mutación SIN IMPORTAR el método HTTP.
  const accionUrl = (() => {
    try { return new URL(url, window.location.href).searchParams.get("action") || ""; }
    catch (e) { return ""; }
  })();
  const esMutacionPorNombre = /^(guardar|actualizar|eliminar|crear|editar|registrar|cambiar|anular|borrar|marcar|activar|fijar|probar|cancelar|confirmar|aplicar|reordenar|subir|migrar|inicializar|sumar)/i.test(accionUrl);

  const esLectura = (!opciones.method || opciones.method === "GET") && !esMutacionPorNombre;
  const timeoutMs = config.timeoutMs || (esLectura ? 25000 : 30000);
  const maxIntentos = esLectura ? (config.reintentos ?? 2) : 1;

  let ultimoError;
  for (let intento = 1; intento <= maxIntentos; intento++) {
    const controlador = new AbortController();
    const timer = setTimeout(() => controlador.abort(), timeoutMs);
    try {
      const respuesta = await fetch(url, { ...opciones, signal: controlador.signal });
      clearTimeout(timer);
      return respuesta;
    } catch (error) {
      clearTimeout(timer);
      ultimoError = error;
      if (window.anotarErrorParaInforme) window.anotarErrorParaInforme("servidor", "Falló la llamada " + (accionUrl || (opciones.body ? (() => { try { return JSON.parse(opciones.body).action; } catch (e) { return ""; } })() : "")) + " (intento " + intento + "): " + (error && error.name === "AbortError" ? "sin respuesta en " + Math.round(timeoutMs / 1000) + " s" : (error && error.message)), "", "");
      if (!esLectura || intento === maxIntentos) throw ultimoError;
      // Corte momentáneo — pequeña espera antes de reintentar
      await new Promise(r => setTimeout(r, 400 * intento));
    }
  }
  throw ultimoError;
}

async function cargarConfigCliente() {
  // App de escritorio (Electron): la API URL vive en SQLite, guardada
  // por setup.html vía fijarConfigLocal("api_url", ...) — nunca hay un
  // config.json físico en este modo. Antes esta función solo miraba
  // config.json, así que en Electron el fetch siempre fallaba y
  // admin.js se quedaba con el default hardcodeado de arriba (la URL
  // de OTRA instalación) sin ningún aviso — cualquier venta o cambio
  // de stock se habría guardado en la planilla equivocada.
  const bridge = window.posOffline || window.veekpos;
  console.log("[cargarConfigCliente] bridge detectado:", bridge ? (window.posOffline ? "posOffline" : "veekpos") : "ninguno (modo navegador)");
  if (bridge && typeof bridge.obtenerConfigLocal === "function") {
    try {
      const guardada = await bridge.obtenerConfigLocal("api_url", "");
      console.log("[cargarConfigCliente] api_url leída de SQLite:", JSON.stringify(guardada));
      if (guardada) {
        API_URL = guardada;
        return true; // ya configurada vía Instalador (setup.html)
      }
    } catch (e) {
      console.error("[cargarConfigCliente] Error al leer api_url desde el almacenamiento local:", e);
    }
    console.log("[cargarConfigCliente] Sin api_url guardada — se considera instalación nueva.");
    // Estamos en Electron pero todavía no hay api_url guardada: no
    // caer al config.json de navegador (no aplica en este modo) ni al
    // default de otra instalación — se trata como no configurada.
    return false;
  }

  // Navegador normal (sin Electron): mismo mecanismo de siempre.
  try {
    const res = await fetch("../config.json?_=" + Date.now(), { cache: "no-store" });
    if (res.ok) {
      const cfg = await res.json();
      if (cfg.apiUrl) API_URL = cfg.apiUrl;
      return true; // config.json existe: esta instalación ya fue configurada alguna vez
    }
  } catch(e) {
    console.log("config.json no encontrado, usando URL por defecto");
  }
  return false;
}

let pedidosGlobal = [];
let _avisoPedidosCaido = false; // evita repetir el toast de error de conexión en cada ciclo del polling

// Store last completed sale for "print from receipt modal"
let ultimaVentaImprimible = null;

// Store POS sales loaded for dashboard
let ventasPOSGlobal = [];

// El chequeo de "¿está logueado?" se hizo ANTES una función síncrona
// que corría acá mismo, redirigiendo a login.html apenas se cargaba
// el script — sin siquiera esperar a ver si config.json existía. El
// problema: en una instalación recién copiada para un cliente nuevo
// (sin config.json todavía, porque no pasó por el Instalador), esto
// mandaba a login.html, que terminaba usando la URL de Apps Script
// de Jireh que viene por defecto en el código — pidiendo la
// contraseña de OTRA cuenta que el cliente nuevo ni siquiera tiene.
// Ahora esa decisión se toma recién dentro de DOMContentLoaded, una
// vez que ya se supo si esta instalación tiene config.json propio o no.

/* =========================================================
   AVISO DE PEDIDO WEB NUEVO
   Cada ~30 s (o ~60 s con la ventana minimizada) se pregunta al
   backend por action=avisoPedidos — no lee hojas, es liviano. Si
   entró un pedido del catálogo: sonido, notificación de Windows,
   aviso en pantalla y un número en "Pedidos web" del menú, que se
   borra al entrar a esa sección. Se puede apagar por computadora en
   Configuración → Esta computadora.
========================================================= */
const AVISO_PED_LS = "vpos_aviso_pedidos";      // { sonido, notificacion }
const AVISO_PED_VISTO_LS = "vpos_aviso_pedidos_visto"; // n ya visto en esta computadora
let _avisoPedN = null;          // último n del servidor
let _avisoPedUltimoChequeo = 0;
let _avisoPedDeshabilitado = false;

function prefsAvisoPedidos() {
  try { return Object.assign({ sonido: true, notificacion: true }, JSON.parse(localStorage.getItem(AVISO_PED_LS) || "{}")); }
  catch (e) { return { sonido: true, notificacion: true }; }
}

function guardarPrefAvisoPedidos(clave, valor) {
  const p = prefsAvisoPedidos(); p[clave] = !!valor;
  try { localStorage.setItem(AVISO_PED_LS, JSON.stringify(p)); } catch (e) {}
  if (clave === "notificacion" && valor) pedirPermisoNotificaciones(true);
  if (clave === "sonido" && valor) sonarAvisoPedido();
  pintarPrefsAvisoPedidos();
}

function pintarPrefsAvisoPedidos() {
  const p = prefsAvisoPedidos();
  const s = document.getElementById("cfgAvisoPedSonido"), n = document.getElementById("cfgAvisoPedNotif");
  if (s) s.checked = p.sonido;
  if (n) n.checked = p.notificacion;
  const est = document.getElementById("cfgAvisoPedEstado");
  if (est) {
    if (!("Notification" in window)) est.textContent = "Este navegador no permite notificaciones: se avisa con sonido y en pantalla.";
    else if (Notification.permission === "denied") est.textContent = "Las notificaciones están bloqueadas para este sitio (se pueden habilitar desde el candado de la barra de direcciones).";
    else est.textContent = "";
  }
}

function pedirPermisoNotificaciones(forzar) {
  if (!("Notification" in window) || Notification.permission !== "default") return;
  if (!forzar && !prefsAvisoPedidos().notificacion) return;
  try { Notification.requestPermission().then(pintarPrefsAvisoPedidos).catch(() => {}); } catch (e) {}
}

function _pedidosVistos() {
  let crudo = null;
  try { crudo = localStorage.getItem(AVISO_PED_VISTO_LS); } catch (e) {}
  if (crudo === null || crudo === "") return null;
  const v = Number(crudo);
  return isFinite(v) ? v : null;
}

function pintarBadgePedidos() {
  const visto = _pedidosVistos();
  const pendientes = _avisoPedN !== null && visto !== null ? Math.max(0, _avisoPedN - visto) : 0;
  document.querySelectorAll('a[data-target="pedidos"]').forEach(a => {
    let b = a.querySelector(".nav-aviso");
    if (!pendientes) { if (b) b.remove(); return; }
    if (!b) { b = document.createElement("span"); b.className = "nav-aviso"; a.appendChild(b); }
    b.textContent = pendientes > 99 ? "99+" : String(pendientes);
    b.title = pendientes === 1 ? "1 pedido web nuevo" : pendientes + " pedidos web nuevos";
  });
}

function marcarPedidosVistos() {
  if (_avisoPedN === null) return;
  try { localStorage.setItem(AVISO_PED_VISTO_LS, String(_avisoPedN)); } catch (e) {}
  pintarBadgePedidos();
}

/** Tres tonos tipo "timbre", distinto del beep de transferencias. */
function sonarAvisoPedido() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    [659, 880, 1319].forEach((freq, i) => {
      const t = ctx.currentTime + i * 0.16;
      const osc = ctx.createOscillator(), gain = ctx.createGain();
      osc.type = "triangle"; osc.frequency.value = freq;
      gain.gain.setValueAtTime(0.0001, t);
      gain.gain.exponentialRampToValueAtTime(0.22, t + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.001, t + 0.3);
      osc.connect(gain).connect(ctx.destination);
      osc.start(t); osc.stop(t + 0.32);
    });
    setTimeout(() => { try { ctx.close(); } catch (e) {} }, 1200);
  } catch (e) { /* sin audio: el aviso visual alcanza */ }
}

function _abrirPedidosDesdeAviso() {
  try { if (window.veekpos && window.veekpos.enfocarVentana) window.veekpos.enfocarVentana(); else window.focus(); } catch (e) {}
  mostrarSeccion("pedidos");
}

function avisarPedidosNuevos(nuevos, cantidad) {
  const ultimo = nuevos[0] || {};
  const monto = ultimo.total ? " · $" + Number(ultimo.total).toLocaleString("es-AR") : "";
  const titulo = cantidad === 1 ? "Pedido web nuevo" : `🛒 ${cantidad} pedidos web nuevos`;
  const cuerpo = cantidad === 1
    ? `${ultimo.cliente || "Cliente"}${monto} (${ultimo.id || ""})`
    : `El último: ${ultimo.cliente || "cliente"}${monto}`;
  const prefs = prefsAvisoPedidos();
  if (prefs.sonido) sonarAvisoPedido();
  if (prefs.notificacion && "Notification" in window && Notification.permission === "granted") {
    try {
      const n = new Notification(titulo, { body: cuerpo, tag: "veekpos-pedido-web", renotify: true, icon: "../icon-192.png" });
      n.onclick = () => { _abrirPedidosDesdeAviso(); n.close(); };
    } catch (e) {}
  }
  if (typeof toast === "function") toast(`${titulo}: ${cuerpo}`, "success");
  // Si la caja ya está mirando Pedidos, se refresca la lista.
  const secPed = document.getElementById("pedidos");
  if (secPed && secPed.style.display === "block") { cargarPedidos(); marcarPedidosVistos(); }
}

async function chequearPedidosNuevos(forzar) {
  if (_avisoPedDeshabilitado || !navigator.onLine) return;
  const rol = obtenerRolActual();
  if (rol !== "admin" && rol !== "vendedor") return;
  const ahora = Date.now();
  if (!forzar && document.hidden && ahora - _avisoPedUltimoChequeo < 55000) return;
  _avisoPedUltimoChequeo = ahora;
  try {
    const res = await fetchAPI(API_URL + "?action=avisoPedidos", {}, { reintentos: 1, timeoutMs: 15000 });
    const data = await res.json();
    if (!data || !data.success) {
      if (data && data.sinPermiso) _avisoPedDeshabilitado = true;
      return;
    }
    const n = Number(data.n) || 0;
    const visto = _pedidosVistos();
    if (visto === null || visto > n) {
      // Primera vez en esta computadora (o se reinició el contador):
      // lo que ya había no se anuncia.
      try { localStorage.setItem(AVISO_PED_VISTO_LS, String(n)); } catch (e) {}
    } else if (_avisoPedN !== null && n > _avisoPedN) {
      const cantidad = n - _avisoPedN;
      avisarPedidosNuevos((data.recientes || []).slice(0, cantidad), cantidad);
    } else if (_avisoPedN === null && n > visto) {
      // Entraron pedidos mientras el panel estaba cerrado: solo el número en el menú.
    }
    _avisoPedN = n;
    pintarBadgePedidos();
  } catch (e) { /* sin red o backend viejo: se reintenta en el próximo ciclo */ }
}

function iniciarAvisoPedidos() {
  pintarPrefsAvisoPedidos();
  // El permiso de notificaciones se pide con el primer clic (los
  // navegadores no lo muestran sin una acción de la persona).
  document.addEventListener("pointerdown", () => pedirPermisoNotificaciones(false), { once: true, capture: true });
  setTimeout(() => chequearPedidosNuevos(true), 2500);
  setInterval(() => chequearPedidosNuevos(false), 30000);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) chequearPedidosNuevos(true); });
}

function iniciarPollingSecciones() {
  const offsetInicial = Math.floor(Math.random() * 4000);

  setTimeout(() => {
    ejecutarPollingSecciones();
    setInterval(ejecutarPollingSecciones, 30000); // 15 s — near real-time without hammering the API
  }, offsetInicial);
}

// Polling más liviano: sin interacción del usuario por más de 5 minutos
// (caja abierta pero nadie mirando la pantalla) se consulta 1 de cada 4
// ciclos (~2 min) en vez de cada 30s; y sin red no se consulta.
let _ultimaInteraccionUsuario = Date.now();
let _cicloPolling = 0;
["pointerdown", "keydown", "wheel", "touchstart"].forEach(ev =>
  document.addEventListener(ev, () => { _ultimaInteraccionUsuario = Date.now(); }, { passive: true, capture: true })
);

function ejecutarPollingSecciones() {
  if (document.hidden) return; // pestaña en segundo plano: no consultar
  if (!navigator.onLine) return;
  _cicloPolling++;
  const inactivo = Date.now() - _ultimaInteraccionUsuario > 5 * 60 * 1000;
  if (inactivo && _cicloPolling % 4 !== 0) return;

  const dashboardVisible = document.getElementById("dashboard").style.display === "block";
  const pedidosVisible   = document.getElementById("pedidos").style.display === "block";
  const clientesVisible  = document.getElementById("clientes").style.display === "block";

  // Antes esto llamaba a cargarMetricas() siempre, aunque el cajero
  // estuviera en Productos o en el POS vendiendo — son llamadas al
  // backend (Apps Script + Sheets) que no hacían falta y competían
  // con lo que el cajero estaba haciendo en ese momento. Ahora solo
  // se actualiza la pantalla que efectivamente está en uso.
  if (dashboardVisible) { cargarMetricas(); cargarVentasPOS(); }
  if (pedidosVisible)   cargarPedidos();
  if (clientesVisible)  cargarClientes();
}

document.addEventListener("DOMContentLoaded", async () => {
  const yaConfigurado = await cargarConfigCliente();

  if (!yaConfigurado) {
    const bridge = window.posOffline || window.veekpos;

    if (bridge) {
      // App de escritorio (Electron) sin api_url guardada todavía: el
      // flujo correcto es setup.html (pide API URL y activa la
      // la licencia), no la sección "Instalador" de abajo — esa es
      // para el modo navegador/multi-cliente y genera un config.json
      // que en Electron no se usa para nada. Mandar ahí por error
      // dejaría a la persona completando un formulario que no guarda
      // los datos donde admin.js realmente los busca.
      window.location.href = "setup.html";
      return;
    }

    // Instalación nueva en modo navegador, sin config.json propio
    // todavía — no tiene sentido pedir login (no hay ninguna cuenta
    // real detrás todavía, y menos la de otro cliente). Se muestra
    // directo el Instalador, sin arrancar nada que dependa de un
    // backend que no existe.
    document.querySelectorAll(".sidebar a[data-target], #bottomNavLinks a[data-target]").forEach(a => {
      if (a.getAttribute("data-target") !== "instalador") a.style.display = "none";
    });
    mostrarSeccion("instalador");
    toast("Primera vez acá — completá el Instalador para configurar este negocio", "success");

    const cat = document.getElementById("loadingCat");
    if (cat) { cat.style.opacity = "0"; setTimeout(() => cat.remove(), 500); }
    return;
  }

  if (sessionStorage.getItem("admin") !== "true") {
    window.location.href = "login.html";
    return;
  }

  // En el navegador hace falta un token del backend (sesiones de antes de
  // esta versión no lo tienen: se pide login de nuevo). En la app de
  // escritorio se puede entrar sin token (login offline): el POS funciona
  // y la cola espera a que haya sesión para subir.
  if (window.VeekAuth && !VeekAuth.tieneSesion() && !(window.posOffline || window.veekpos)) {
    sessionStorage.removeItem("admin");
    window.location.href = "login.html";
    return;
  }

  aplicarPermisosPorRol();
  aplicarModoNegocio();
  mostrarAvisoPasswordDeFabrica();

  // Verificar licencia (solo en Electron con window.veekpos disponible)
  aplicarEstadoLicencia();
  // Caja cliente: cuando aparece la caja servidor, se revisa al instante
  if (window.veekpos && typeof window.veekpos.alEncontrarServidor === "function") window.veekpos.alEncontrarServidor(() => aplicarEstadoLicencia());
  // Actualización automática: avisa cuando hay una versión lista o el backend quedó atrás
  iniciarAvisosActualizacion();

  // Pantalla inicial: el POS (si el rol lo permite; si no, el dashboard).
  // Nada de lo que sigue bloquea el arranque: antes se esperaba a las
  // metricas/ventas del dashboard (Apps Script, ~20s) antes de quitar la
  // pantalla de carga, aunque el usuario ni estuviera en el dashboard.
  // En Electron, pos-offline.js se inyecta desde preload.js y envuelve
  // asegurarProductosPOS (catalogo desde SQLite). Se espera (max 3s) a que
  // este listo, para no abrir el POS con la version online lenta.
  if (window.veekpos || window.posOffline) {
    for (let i = 0; i < 60 && !window.__posOfflineListo; i++) await new Promise(r => setTimeout(r, 50));
  }
  mostrarSeccion(seccionPermitidaParaRol("pos") ? "pos" : "dashboard");

  // Quitar el loading ya mismo
  const cat = document.getElementById("loadingCat");
  if (cat) {
    cat.style.opacity = "0";
    setTimeout(() => cat.remove(), 500);
  }

  cargarConfigNegocioDesdeBackend();
  reconectarImpresoraUSBSiPosible();
  iniciarPollingSecciones();
  iniciarAvisoPedidos();

  // Ventas pendientes de sesiones anteriores: aviso + reintento en segundo plano
  actualizarBadgeVentasPendientes(_leerColaVentasPendientes().length);
  sincronizarVentasPendientes();

  setupScannerListener();
  suscribirseATransferenciasMP();
});


/* ===================== UTILS ===================== */

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text || "";
  return div.innerHTML;
}

/** Igual que escapeHtml, pero además neutraliza el apóstrofe (como entidad
 *  HTML &#39;, que el navegador decodifica de nuevo a ' al leer el atributo).
 *  Usar SIEMPRE que un valor se inserte dentro de un string JS delimitado
 *  por comillas simples dentro de un onclick/onchange generado con innerHTML
 *  — si no, un apóstrofe en el dato (código, nombre, motivo, etc.) corta el
 *  string antes de tiempo y rompe el atributo con "missing ) after argument list". */
function escapeJsAttr(text) {
  // Orden correcto: primero se escapa para el string JS (\ y '), y RECIÉN
  // DESPUÉS para el atributo HTML. Antes solo se convertía ' en &#39;, que el
  // navegador decodifica de nuevo a ' ANTES de ejecutar el JS, así que el
  // apóstrofe seguía cortando el string. Sirve tanto para onclick="..." como
  // para onclick='...'.
  const js = String(text ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/'/g, "\\'")
    .replace(/\r?\n/g, " ");
  return js
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Devuelve una versión "demorada" de una función: si se la llama
 * varias veces seguidas (por ejemplo, una vez por cada tecla mientras
 * alguien escribe rápido en un buscador), solo se ejecuta la ÚLTIMA
 * llamada, esperando `demoraMs` desde que paró de llamarse — evita
 * redibujar una tabla o grilla completa en cada tecla, cuando en la
 * práctica el usuario ya tipeó la letra siguiente antes de llegar a
 * ver el resultado intermedio.
 *
 * Uso: const buscarConDemora = debounce(buscarFunction, 150);
 *      buscarConDemora(); // se llama igual que la función original
 */
function debounce(fn, demoraMs) {
  let temporizador = null;
  return function (...args) {
    clearTimeout(temporizador);
    temporizador = setTimeout(() => fn.apply(this, args), demoraMs);
  };
}

function actualizarElemento(id, valor) {
  const el = document.getElementById(id);
  if (el) el.textContent = valor;
}

/* ===================== CONFIGURACIÓN DEL NEGOCIO (encabezado del ticket) ===================== */

// Valores por defecto — genéricos a propósito, sin ningún nombre de
// cliente hardcodeado. Solo se usan como placeholder mientras todavía
// no cargó la config real desde el backend (hoja CONFIGURACION); una
// vez que carga, siempre gana lo que esté guardado ahí.
const CONFIG_NEGOCIO_DEFAULT = {
  nombre:     "",
  subtitulo:  "Punto de Venta",
  direccion:  "",
  telefono1:  "",
  telefono2:  "",
  pie:        "¡Gracias por su compra!",

  bannerTitulo:    "",
  bannerSubtitulo: "",
  bannerImagen:    "",
  tema:            "navy",

  cbuTransferencia:  "",
  nombreTitularCbu:  ""
};

// Caché en memoria de la config, para que imprimir un ticket no tenga
// que esperar una llamada al servidor cada vez. Se carga al iniciar la
// app y se refresca cada vez que se guarda desde el formulario.
let configNegocioCache = { ...CONFIG_NEGOCIO_DEFAULT };

/** Synchronous read used by the print functions — always returns instantly from the in-memory cache */
function obtenerConfigNegocio() {
  return configNegocioCache;
}

/** Fetches the saved config from the backend (hoja CONFIGURACION) and refreshes the in-memory cache */
async function cargarConfigNegocioDesdeBackend() {
  try {
    const response = await fetchAPI(API_URL + "?action=configuracionNegocio");
    const data = await response.json();
    if (data.success && data.config) {
      configNegocioCache = { ...CONFIG_NEGOCIO_DEFAULT, ...data.config };
      if (data.config.modoNegocio) {
        try { sessionStorage.setItem("modoNegocio", data.config.modoNegocio); } catch (e) {}
        aplicarModoNegocio();
      }
      if (typeof depDolarDesdeConfig === "function") depDolarDesdeConfig(configNegocioCache);
      if (configNegocioCache.zonaHoraria) _fijarZonaNegocio(configNegocioCache.zonaHoraria);
      aplicarSidebarBrand(configNegocioCache);
      aplicarTemaAdmin(configNegocioCache);
    }
  } catch (error) {
    console.warn("No se pudo cargar la configuración del negocio:", error);
  }
  return configNegocioCache;
}

/** Loads the saved config into the form fields (called when the Configuración section opens) */
async function cargarConfigNegocioForm() {
  const form = document.getElementById("cfgNombreLocal");
  if (form) form.placeholder = "Cargando...";

  const cfg = await cargarConfigNegocioDesdeBackend();

  // Setter seguro: si el input no existe en este HTML, no rompe el resto de la carga.
  const setVal = (id, valor) => {
    const el = document.getElementById(id);
    if (el) el.value = valor ?? "";
    else console.warn(`cargarConfigNegocioForm: no existe #${id} en este HTML, se omite.`);
  };

  setVal("cfgNombreLocal", cfg.nombre);
  setVal("cfgSubtitulo", cfg.subtitulo);
  setVal("cfgDireccion", cfg.direccion);
  setVal("cfgTelefono1", cfg.telefono1);
  setVal("cfgTelefono2", cfg.telefono2);
  setVal("cfgPie", cfg.pie);
  const selZona = document.getElementById("cfgZonaHoraria");
  if (selZona) {
    const z = cfg.zonaHoraria || _zonaNegocio();
    if (![...selZona.options].some(o => o.value === z)) selZona.add(new Option(z, z));
    selZona.value = z;
    _mostrarHoraZona();
  }

  cargarAparienciaForm(cfg);
  cargarBeneficiosForm(cfg);
  setVal("cfgCbuTransferencia", cfg.cbuTransferencia ?? CONFIG_NEGOCIO_DEFAULT.cbuTransferencia);
  setVal("cfgNombreTitularCbu", cfg.nombreTitularCbu ?? CONFIG_NEGOCIO_DEFAULT.nombreTitularCbu);
  setVal("cfgBannerTopMensajes", cfg.bannerTopMensajes ?? "");
  setVal("cfgTransportesNoDisponibles", cfg.transportesNoDisponibles ?? "");
  cargarSidebarForm(cfg);
  cargarDriveProductosForm(cfg);
  cargarDrivePedidosForm(cfg);
  cargarUrlCatalogoForm(cfg);

  if (form) form.placeholder = "Ej: Mi Negocio";
}

/** Muestra la hora actual en la zona elegida, para comprobar que es la correcta */
function _mostrarHoraZona() {
  const sel = document.getElementById("cfgZonaHoraria"), out = document.getElementById("cfgZonaHorariaHora");
  if (!sel || !out) return;
  try { out.textContent = "Ahora en esa zona: " + new Date().toLocaleString("es-AR", { timeZone: sel.value, weekday: "long", hour: "2-digit", minute: "2-digit" }); }
  catch (e) { out.textContent = ""; }
}

/** Reads the form fields and saves them to the backend (hoja CONFIGURACION) */
async function guardarConfigNegocioForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("El nombre del local es obligatorio", "error");
    return;
  }

  const cfg = {
    nombre,
    subtitulo: document.getElementById("cfgSubtitulo").value.trim(),
    direccion: document.getElementById("cfgDireccion").value.trim(),
    telefono1: document.getElementById("cfgTelefono1").value.trim(),
    telefono2: document.getElementById("cfgTelefono2").value.trim(),
    pie:       document.getElementById("cfgPie").value.trim()
  };
  const selZona = document.getElementById("cfgZonaHoraria");
  if (selZona && selZona.value) cfg.zonaHoraria = selZona.value;

  const cbuEl = document.getElementById("cfgCbuTransferencia");
  const titularEl = document.getElementById("cfgNombreTitularCbu");
  if (cbuEl) cfg.cbuTransferencia = cbuEl.value.trim();
  if (titularEl) cfg.nombreTitularCbu = titularEl.value.trim();

  const btn = document.getElementById("btnGuardarConfigNegocio");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar la configuración", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg }; // antes se reemplazaba entera y se perdía el resto (URL del catálogo, etc.)
    if (cfg.zonaHoraria) _fijarZonaNegocio(cfg.zonaHoraria);
    toast("Configuración guardada", "success");

  } catch (error) {
    console.error("Error al guardar la configuración del negocio:", error);
    toast("Error de conexión al guardar la configuración", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/**
 * Modal de confirmación genérico, para reemplazar usos puntuales de
 * window.confirm() nativo — ese diálogo síncrono puede dejar el foco
 * de teclado roto en Electron después de cerrarse (el input de texto
 * deja de responder). mensaje: texto a mostrar. onConfirmar: función
 * que se ejecuta si el usuario confirma.
 */
function confirmarAccion(mensaje, onConfirmar, titulo) {
  window._accionConfirmadaCallback = onConfirmar;

  const backdrop = document.getElementById("modalConfirmGenericoBackdrop");
  const texto = document.getElementById("modalConfirmGenericoTexto");
  const tituloEl = document.getElementById("modalConfirmGenericoTitulo");
  if (texto) texto.textContent = mensaje;
  if (tituloEl) tituloEl.textContent = titulo || "⚠️ Confirmar acción";
  if (backdrop) backdrop.classList.add("show");
}

function cerrarModalConfirmGenerico() {
  const backdrop = document.getElementById("modalConfirmGenericoBackdrop");
  if (backdrop) backdrop.classList.remove("show");
  window._accionConfirmadaCallback = null;
}

function _ejecutarAccionConfirmada() {
  const fn = window._accionConfirmadaCallback;
  cerrarModalConfirmGenerico();
  if (typeof fn === "function") fn();
}

/** Resets the form (and the saved sheet values) back to the generic empty defaults */
function restablecerConfigNegocio() {
  confirmarAccion(
    "¿Restablecer los datos del local a los valores originales?",
    _restablecerConfigNegocioConfirmado
  );
}

async function _restablecerConfigNegocioConfirmado() {
  document.getElementById("cfgNombreLocal").value = CONFIG_NEGOCIO_DEFAULT.nombre;
  document.getElementById("cfgSubtitulo").value   = CONFIG_NEGOCIO_DEFAULT.subtitulo;
  document.getElementById("cfgDireccion").value   = CONFIG_NEGOCIO_DEFAULT.direccion;
  document.getElementById("cfgTelefono1").value   = CONFIG_NEGOCIO_DEFAULT.telefono1;
  document.getElementById("cfgTelefono2").value   = CONFIG_NEGOCIO_DEFAULT.telefono2;
  document.getElementById("cfgPie").value         = CONFIG_NEGOCIO_DEFAULT.pie;

  await guardarConfigNegocioForm();
}

/** Prints a sample receipt using the form's current (possibly unsaved) values, so the admin can check before saving */
function vistaPreviaTicketConfig() {
  const cfgPreview = {
    nombre:    document.getElementById("cfgNombreLocal").value.trim() || CONFIG_NEGOCIO_DEFAULT.nombre,
    subtitulo: document.getElementById("cfgSubtitulo").value.trim(),
    direccion: document.getElementById("cfgDireccion").value.trim(),
    telefono1: document.getElementById("cfgTelefono1").value.trim(),
    telefono2: document.getElementById("cfgTelefono2").value.trim(),
    pie:       document.getElementById("cfgPie").value.trim() || CONFIG_NEGOCIO_DEFAULT.pie
  };

  const itemsEjemplo = [
    { PRODUCTO: "Producto de ejemplo", PRECIO: 1500, cantidad: 2 },
    { PRODUCTO: "Otro producto", PRECIO: 800, cantidad: 1 }
  ];
  const total = itemsEjemplo.reduce((acc, i) => acc + i.PRECIO * i.cantidad, 0);

  const frame = document.getElementById("thermalPrintFrame");
  if (!frame) { toast("Error: frame de impresión no encontrado", "error"); return; }

  frame.innerHTML = buildThermalHTML("PREVIEW", itemsEjemplo, total, "EFECTIVO", new Date(), null, cfgPreview);

  _setPrintPageSize("80mm");
  setTimeout(() => { window.print(); }, 120);
}

/* ===================== APARIENCIA DEL PANEL ADMIN (letra + nombre del sidebar) ===================== */

const SIDEBAR_BRAND_DEFAULT = {
  sidebarMark:  "",
  sidebarTexto: ""
};

/** Applies the saved letter/name to the sidebar in the DOM — runs on every page load, not just inside Configuración */
function aplicarSidebarBrand(cfg) {
  const markEl  = document.getElementById("sidebarMark");
  const textoEl = document.getElementById("sidebarLabelTexto");

  if (markEl)  markEl.textContent  = (cfg.sidebarMark  || SIDEBAR_BRAND_DEFAULT.sidebarMark).toUpperCase();
  if (textoEl) textoEl.textContent = (cfg.sidebarTexto || SIDEBAR_BRAND_DEFAULT.sidebarTexto).toUpperCase();
}

/** Loads the saved letter/name into the "Apariencia del panel admin" form */
function cargarSidebarForm(cfg) {
  document.getElementById("cfgSidebarMark").value  = cfg.sidebarMark  ?? SIDEBAR_BRAND_DEFAULT.sidebarMark;
  document.getElementById("cfgSidebarTexto").value = cfg.sidebarTexto ?? SIDEBAR_BRAND_DEFAULT.sidebarTexto;
}

/** Reads the form and saves the sidebar letter/name to the backend (hoja CONFIGURACION) */
async function guardarSidebarForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const cfg = {
    nombre,
    sidebarMark:  document.getElementById("cfgSidebarMark").value.trim() || SIDEBAR_BRAND_DEFAULT.sidebarMark,
    sidebarTexto: document.getElementById("cfgSidebarTexto").value.trim()
  };

  const btn = document.getElementById("btnGuardarSidebar");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    aplicarSidebarBrand(configNegocioCache);
    toast("Listo, ya se ve en el sidebar", "success");

  } catch (error) {
    console.error("Error al guardar el sidebar:", error);
    toast("Error de conexión al guardar", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Loads the configured Drive folder (for product photo uploads) into the Configuración form */
function cargarDriveProductosForm(cfg) {
  const input = document.getElementById("cfgDriveCarpetaProductos");
  if (input) input.value = cfg.driveCarpetaProductosId || "";
}

/** Saves the Drive folder link/ID used for product photo uploads */
async function guardarDriveProductosForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const cfg = {
    nombre,
    driveCarpetaProductosId: document.getElementById("cfgDriveCarpetaProductos").value.trim()
  };

  const btn = document.getElementById("btnGuardarDriveProductos");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    toast("Carpeta de Drive guardada", "success");

  } catch (error) {
    console.error("Error al guardar la carpeta de Drive:", error);
    toast("Error de conexión al guardar", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Loads the configured Drive folder (for order PDF storage) into the Configuración form */
function cargarDrivePedidosForm(cfg) {
  const carpetaInput = document.getElementById("cfgDriveCarpetaPedidos");
  if (carpetaInput) carpetaInput.value = cfg.driveCarpetaPedidosId || "";

  const subtituloInput = document.getElementById("cfgPedidoPdfSubtitulo");
  if (subtituloInput) subtituloInput.value = cfg.pedidoPdfSubtitulo || "";

  const pieInput = document.getElementById("cfgPedidoPdfPie");
  if (pieInput) pieInput.value = cfg.pedidoPdfPie || "";
}

/** Loads the catalog public URL into the Configuración form, used to build price-QR links */
function cargarUrlCatalogoForm(cfg) {
  const input = document.getElementById("cfgUrlCatalogo");
  if (input) input.value = cfg.urlCatalogo || "";
}

/** Saves the catalog public URL */
async function guardarUrlCatalogoForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const cfg = {
    nombre,
    urlCatalogo: document.getElementById("cfgUrlCatalogo").value.trim().replace(/\/+$/, "") // sin / al final, evita //precio.html
  };

  const btn = document.getElementById("btnGuardarUrlCatalogo");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    toast("URL del catálogo guardada", "success");

  } catch (error) {
    console.error("Error al guardar la URL del catálogo:", error);
    toast("Error de conexión al guardar", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Saves both the Drive folder and the editable texts used by the order PDF */
async function guardarDrivePedidosForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const cfg = {
    nombre,
    driveCarpetaPedidosId: document.getElementById("cfgDriveCarpetaPedidos").value.trim(),
    pedidoPdfSubtitulo:    document.getElementById("cfgPedidoPdfSubtitulo").value.trim(),
    pedidoPdfPie:          document.getElementById("cfgPedidoPdfPie").value.trim()
  };

  const btn = document.getElementById("btnGuardarDrivePedidos");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    toast("Configuración del PDF de pedidos guardada", "success");

  } catch (error) {
    console.error("Error al guardar la carpeta de Drive:", error);
    toast("Error de conexión al guardar", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/**
 * Cambia el usuario/contraseña de acceso al panel. Requiere la
 * contraseña actual — el backend la valida antes de aplicar el cambio.
 * Por seguridad, este formulario nunca precarga el usuario/contraseña
 * vigentes (a diferencia de los demás formularios de Configuración).
 */
async function guardarCredencialesForm() {
  const passwordActual = document.getElementById("credPasswordActual").value;
  const nuevoUsuario = document.getElementById("credNuevoUsuario").value.trim();
  const nuevaPassword = document.getElementById("credNuevaPassword").value;

  if (!passwordActual) {
    toast("Ingresá tu contraseña actual para confirmar el cambio", "error");
    return;
  }

  if (!nuevoUsuario || !nuevaPassword) {
    toast("Completá el nuevo usuario y la nueva contraseña", "error");
    return;
  }

  if (nuevaPassword.length < 6) {
    toast("La nueva contraseña debe tener al menos 6 caracteres", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarCredenciales");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({
      action: "guardarCredencialesAdmin",
      passwordActual,
      nuevoUsuario,
      nuevaPassword,
      dispositivo: (window.posOffline || window.veekpos) ? "1" : ""
    });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo cambiar el acceso", "error");
      return;
    }

    // Cambiar la contraseña cierra las demás sesiones; esta sigue con el token nuevo.
    if (data.token && window.VeekAuth) VeekAuth.guardarSesion(data, !!(window.posOffline || window.veekpos));
    try { sessionStorage.removeItem("debeCambiarPassword"); } catch (e) {}
    const avisoFabrica = document.getElementById("avisoPasswordFabrica");
    if (avisoFabrica) avisoFabrica.remove();

    document.getElementById("credPasswordActual").value = "";
    document.getElementById("credNuevoUsuario").value = "";
    document.getElementById("credNuevaPassword").value = "";
    toast("Usuario y contraseña actualizados", "success");

  } catch (error) {
    console.error("Error al cambiar las credenciales:", error);
    toast("Error de conexión al guardar", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/**
 * Cambia la contraseña de acceso a stock.html. Requiere la contraseña
 * ACTUAL DEL PANEL ADMIN (no la de stock vieja) para autorizar el
 * cambio — quien está acá adentro ya demostró ser admin, no hace
 * falta pedirle que recuerde una segunda contraseña.
 */
async function guardarPasswordStockForm() {
  const passwordActualAdmin = document.getElementById("stockPasswordAdminActual").value;
  const nuevaPasswordStock = document.getElementById("stockPasswordNueva").value;

  if (!passwordActualAdmin) {
    toast("Ingresá tu contraseña del panel admin para confirmar el cambio", "error");
    return;
  }
  if (!nuevaPasswordStock || nuevaPasswordStock.length < 4) {
    toast("La nueva contraseña de stock debe tener al menos 4 caracteres", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarPasswordStock");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({
      action: "guardarPasswordStock",
      passwordActualAdmin,
      nuevaPasswordStock
    });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo cambiar la contraseña de stock", "error");
      return;
    }

    document.getElementById("stockPasswordAdminActual").value = "";
    document.getElementById("stockPasswordNueva").value = "";
    toast("Contraseña de stock actualizada", "success");

  } catch (error) {
    console.error("Error al cambiar la contraseña de stock:", error);
    toast("Error de conexión al guardar", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/* ===================== APARIENCIA DEL CATÁLOGO WEB (banner + tema) ===================== */

const APARIENCIA_DEFAULT = {
  navbarTexto:     "",
  navbarIcono:     "🏬",
  bannerTitulo:    "",
  bannerSubtitulo: "",
  bannerImagen:    "",
  tema:            "navy",
  gradPersonalizado: false,
  gradA: "#241536",
  gradB: "#3a2856",

  urlCatalogo: "",
  nombreCorto: "",

  iconoUrl: "icon-512.png",
  whatsappIconoUrl: "",

  seoTitulo: "",
  seoDescripcion: "",
  seoKeywords: "",

  footerTitulo1: "",
  footerTexto1: "",
  footerTexto2: "",
  footerTexto3: "",
  footerCopyright: ""
};

/**
 * Aplica el tema de color al panel admin (este mismo selector pinta
 * también el catálogo web — es un solo dato compartido en Sheets).
 */
function aplicarTemaAdmin(cfg) {
  const tema = (cfg.tema || APARIENCIA_DEFAULT.tema).toLowerCase();
  document.body.setAttribute("data-tema", tema);
  // Tema de temporada activo: sus colores pisan los del tema (ver temporadas.js)
  let t = null;
  try { if (window.VeekTemporadas) t = window.VeekTemporadas.aplicarPanel(cfg.temporada || null); } catch (e) {}
  const marca = document.getElementById("sidebarTemporada");
  if (t && !marca) {
    const brand = document.querySelector(".sidebar .sidebar-brand, .sidebar-brand");
    if (brand) brand.insertAdjacentHTML("beforeend", `<span id="sidebarTemporada" class="sidebar-temporada" title="Tema de temporada: ${escapeHtml(t.nombre)}">${t.emoji}</span>`);
  } else if (t && marca) {
    marca.textContent = t.emoji; marca.title = "Tema de temporada: " + t.nombre;
  } else if (!t && marca) {
    marca.remove();
  }
}

/* ===================================================================
   TEMA DE TEMPORADA (Configuración → Catálogo web)
=================================================================== */
let _temporadaElegida = "";

function _datosTemporadaForm() {
  return {
    id: _temporadaElegida,
    desde: document.getElementById("tempDesde")?.value || "",
    hasta: document.getElementById("tempHasta")?.value || "",
    cartel: (document.getElementById("tempCartel")?.value || "").trim(),
    animacion: !!document.getElementById("tempAnimacion")?.checked,
    colores: !!document.getElementById("tempColores")?.checked
  };
}

function cargarTemporadaForm(cfg) {
  const T = window.VeekTemporadas;
  const grilla = document.getElementById("tempGrilla");
  if (!T || !grilla) return;
  const tc = cfg.temporada || {};
  _temporadaElegida = tc.id || "";
  const opcion = (id, nombre, emoji, t) => {
    let muestra = '<div class="temp-muestra temp-muestra--ninguno">Estilo normal</div>';
    if (t) {
      const tira = T.adornoTira(t.adorno);
      const c = t.colores;
      muestra = `<div class="temp-muestra" style="background:linear-gradient(135deg, ${c.a}, ${c.b})">
          <div class="temp-muestra-adorno" style="background-image:${tira ? tira.fondo : "none"};background-size:${tira ? tira.ancho / 1.6 + "px " + tira.alto / 1.6 + "px" : "auto"}"></div>
          <div class="temp-muestra-papel" style="background-color:${c.papel};background-image:${T.fondoMotivos(t.fondo, c.acento)};background-size:90px 90px">
            <span class="temp-muestra-btn" style="background:${c.acento}">Agregar</span><span class="temp-muestra-icono">${t.icono}</span>
          </div>
        </div>`;
    }
    return `<button type="button" class="temp-opcion${id === _temporadaElegida ? " activa" : ""}" role="radio" aria-checked="${id === _temporadaElegida}" data-temp="${id}" onclick="elegirTemporada('${id}')">
        ${muestra}<span class="temp-nombre">${emoji} ${escapeHtml(nombre)}</span></button>`;
  };
  grilla.innerHTML = opcion("", "Sin tema", "⚪", null) + T.ORDEN.map(id => opcion(id, T.TEMPORADAS[id].nombre, T.TEMPORADAS[id].emoji, T.TEMPORADAS[id])).join("");
  document.getElementById("tempDesde").value = tc.desde || "";
  document.getElementById("tempHasta").value = tc.hasta || "";
  document.getElementById("tempCartel").value = tc.cartel || "";
  document.getElementById("tempCartel").placeholder = _temporadaElegida && T.TEMPORADAS[_temporadaElegida] ? T.TEMPORADAS[_temporadaElegida].cartel : "Ej: ¡Felices fiestas!";
  document.getElementById("tempAnimacion").checked = tc.animacion !== false;
  document.getElementById("tempColores").checked = tc.colores !== false;
  actualizarEstadoTemporada();
}

function elegirTemporada(id) {
  const T = window.VeekTemporadas;
  const anterior = _temporadaElegida;
  _temporadaElegida = id;
  document.querySelectorAll("#tempGrilla .temp-opcion").forEach(b => {
    const activa = b.dataset.temp === id;
    b.classList.toggle("activa", activa);
    b.setAttribute("aria-checked", activa ? "true" : "false");
  });
  // Cartel sugerido: si estaba vacío o era el sugerido del tema anterior
  const cartel = document.getElementById("tempCartel");
  const sugeridoAnterior = anterior && T.TEMPORADAS[anterior] ? T.TEMPORADAS[anterior].cartel : "";
  if (cartel && (!cartel.value.trim() || cartel.value.trim() === sugeridoAnterior)) cartel.value = id ? T.TEMPORADAS[id].cartel : "";
  if (cartel) cartel.placeholder = id ? T.TEMPORADAS[id].cartel : "Ej: ¡Felices fiestas!";
  actualizarEstadoTemporada();
  previsualizarColoresTemporadaPanel();
}

/** Muestra en el panel cómo quedarían los colores (sin guardar todavía) */
function previsualizarColoresTemporadaPanel() {
  const d = _datosTemporadaForm();
  try { window.VeekTemporadas.aplicarPanel(d.id ? { id: d.id, colores: d.colores } : null); } catch (e) {}
}

function actualizarEstadoTemporada() {
  const el = document.getElementById("tempEstado");
  if (!el || !window.VeekTemporadas) return;
  const d = _datosTemporadaForm();
  const T = window.VeekTemporadas;
  const f = (iso) => iso ? iso.split("-").reverse().join("/") : "";
  if (!d.id) { el.className = "temp-estado"; el.textContent = "Sin tema de temporada"; return; }
  const nombre = T.TEMPORADAS[d.id].emoji + " " + T.TEMPORADAS[d.id].nombre;
  if (d.desde && d.hasta && d.desde > d.hasta) { el.className = "temp-estado error"; el.textContent = "⚠️ \"Desde\" es posterior a \"Hasta\""; return; }
  const hoy = T.hoyAR();
  if (d.desde && hoy < d.desde) { el.className = "temp-estado programado"; el.textContent = `${nombre}: empieza el ${f(d.desde)}`; return; }
  if (d.hasta && hoy > d.hasta) { el.className = "temp-estado vencido"; el.textContent = `${nombre}: terminó el ${f(d.hasta)}`; return; }
  el.className = "temp-estado activo";
  el.textContent = `${nombre}: activo` + (d.hasta ? ` hasta el ${f(d.hasta)}` : "");
}

function vistaPreviaTemporada() {
  const d = _datosTemporadaForm();
  let base = "../index.html";
  if (typeof _esAppEscritorio === "function" && _esAppEscritorio()) {
    base = (configNegocioCache && configNegocioCache.urlCatalogo) || "";
    if (!base) { toast("Cargá la URL del catálogo (Catálogo web → Catálogo online) para ver la vista previa", "error"); return; }
  }
  if (!d.id) { window.open(base, "_blank", "noopener"); return; }
  const url = base + (base.indexOf("?") === -1 ? "?" : "&") + new URLSearchParams({ vistaTemporada: d.id, anim: d.animacion ? "1" : "0", colores: d.colores ? "1" : "0", cartel: d.cartel }).toString();
  window.open(url, "_blank", "noopener");
}

async function guardarTemporada() {
  const d = _datosTemporadaForm();
  if (d.desde && d.hasta && d.desde > d.hasta) { toast("La fecha \"desde\" no puede ser posterior a \"hasta\"", "error"); return; }
  const btn = document.getElementById("btnGuardarTemporada");
  const txt = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando…"; }
  try {
    const params = new URLSearchParams({
      action: "guardarConfiguracionNegocio",
      temporada: d.id, temporadaDesde: d.desde, temporadaHasta: d.hasta,
      temporadaAnimacion: d.animacion ? "SI" : "NO", temporadaColores: d.colores ? "SI" : "NO",
      temporadaCartel: d.cartel
    });
    const res = await fetchAPI(API_URL + "?" + params.toString());
    const data = await res.json();
    if (!data.success) { toast(data.message || "No se pudo guardar el tema", "error"); return; }
    configNegocioCache = { ...configNegocioCache, temporada: d };
    aplicarTemaAdmin(configNegocioCache);
    actualizarEstadoTemporada();
    const T = window.VeekTemporadas;
    toast(d.id ? `${T.TEMPORADAS[d.id].emoji} Tema "${T.TEMPORADAS[d.id].nombre}" guardado — el catálogo lo muestra en unos minutos` : "Tema de temporada quitado — el catálogo vuelve al estilo normal en unos minutos", "success");
  } catch (e) {
    toast("Error de conexión al guardar el tema", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = txt; }
  }
}

/** Muestra u oculta los dos selectores de color del gradiente propio */
function toggleGradPersonalizado() {
  const activo = document.getElementById("cfgGradPersonalizado").checked;
  document.getElementById("cfgGradWrap").style.display = activo ? "" : "none";
}

/** Loads the saved banner/tema config into the "Apariencia" form (called when Configuración opens) */
function cargarAparienciaForm(cfg) {
  try { cargarTemporadaForm(cfg); } catch (e) { console.warn("Tema de temporada:", e); }
  document.getElementById("cfgNavbarTexto").value     = cfg.navbarTexto     ?? APARIENCIA_DEFAULT.navbarTexto;
  document.getElementById("cfgNavbarIcono").value     = cfg.navbarIcono     ?? APARIENCIA_DEFAULT.navbarIcono;
  document.getElementById("cfgBannerTitulo").value    = cfg.bannerTitulo    ?? APARIENCIA_DEFAULT.bannerTitulo;
  document.getElementById("cfgBannerSubtitulo").value = cfg.bannerSubtitulo ?? APARIENCIA_DEFAULT.bannerSubtitulo;
  document.getElementById("cfgBannerImagen").value    = cfg.bannerImagen   ?? APARIENCIA_DEFAULT.bannerImagen;
  document.getElementById("cfgTema").value            = cfg.tema           || APARIENCIA_DEFAULT.tema;
  actualizarPreviewBanner();

  const gradActivo = cfg.gradPersonalizado ?? APARIENCIA_DEFAULT.gradPersonalizado;
  document.getElementById("cfgGradPersonalizado").checked = !!gradActivo;
  document.getElementById("cfgGradA").value = cfg.gradA || APARIENCIA_DEFAULT.gradA;
  document.getElementById("cfgGradB").value = cfg.gradB || APARIENCIA_DEFAULT.gradB;
  document.getElementById("cfgGradWrap").style.display = gradActivo ? "" : "none";

  // Dominio del catálogo, nombre corto y SEO/redes
  const setIfExists = (id, valor) => { const el = document.getElementById(id); if(el) el.value = valor ?? ""; };
  setIfExists("cfgNombreCorto", cfg.nombreCorto ?? APARIENCIA_DEFAULT.nombreCorto);
  setIfExists("cfgIconoUrl", cfg.iconoUrl ?? APARIENCIA_DEFAULT.iconoUrl);
  setIfExists("cfgWhatsappIconoUrl", cfg.whatsappIconoUrl ?? APARIENCIA_DEFAULT.whatsappIconoUrl);
  setIfExists("cfgSeoTitulo", cfg.seoTitulo ?? APARIENCIA_DEFAULT.seoTitulo);
  setIfExists("cfgSeoDescripcion", cfg.seoDescripcion ?? APARIENCIA_DEFAULT.seoDescripcion);
  setIfExists("cfgSeoKeywords", cfg.seoKeywords ?? APARIENCIA_DEFAULT.seoKeywords);
  setIfExists("cfgFooterTitulo1", cfg.footerTitulo1 ?? APARIENCIA_DEFAULT.footerTitulo1);
  setIfExists("cfgFooterTexto1", cfg.footerTexto1 ?? APARIENCIA_DEFAULT.footerTexto1);
  setIfExists("cfgFooterTexto2", cfg.footerTexto2 ?? APARIENCIA_DEFAULT.footerTexto2);
  setIfExists("cfgFooterTexto3", cfg.footerTexto3 ?? APARIENCIA_DEFAULT.footerTexto3);
  setIfExists("cfgFooterCopyright", cfg.footerCopyright ?? APARIENCIA_DEFAULT.footerCopyright);

  // Pedido mínimo
  const cfgPedidoMinimoEl = document.getElementById("cfgPedidoMinimo");
  if (cfgPedidoMinimoEl && cfg.pedidoMinimo !== undefined) cfgPedidoMinimoEl.value = cfg.pedidoMinimo;

  // Dólar de referencia (tipo de cambio de Ingreso de Productos)
  const cfgDolarTipoEl = document.getElementById("cfgDolarTipo");
  if (cfgDolarTipoEl) { cfgDolarTipoEl.value = tcTipoConfigurado(); dolarMostrarValoresHoy(); }

  // Popup promocional
  const popupActivo = document.getElementById("cfgPopupActivo");
  const popupImagen = document.getElementById("cfgPopupImagen");
  if (popupActivo) popupActivo.checked = !!cfg.popupActivo;

  const tipo = cfg.popupTipo === "video" ? "video" : "imagen";
  const radioImagen = document.getElementById("cfgPopupTipoImagen");
  const radioVideo = document.getElementById("cfgPopupTipoVideo");
  if (radioImagen) radioImagen.checked = tipo === "imagen";
  if (radioVideo) radioVideo.checked = tipo === "video";

  if (tipo === "video") {
    if (popupImagen) popupImagen.value = "";
    const videoUrlEl = document.getElementById("cfgPopupVideoUrl");
    if (videoUrlEl) videoUrlEl.value = cfg.popupImagen || "";
  } else {
    if (popupImagen) popupImagen.value = cfg.popupImagen || "";
    const videoUrlEl = document.getElementById("cfgPopupVideoUrl");
    if (videoUrlEl) videoUrlEl.value = "";
  }

  onCambiarTipoPopup();

  // Mostrar preview si ya hay imagen guardada
  const popupPreviewEl = document.getElementById("popupImagenPreview");
  if (popupPreviewEl && tipo === "imagen" && cfg.popupImagen) {
    popupPreviewEl.innerHTML = `<img src="${cfg.popupImagen}" alt="" style="width:100%;height:100%;object-fit:contain;">`;
  }

  // Mostrar preview de video si corresponde
  if (tipo === "video" && cfg.popupImagen) {
    const videoPreview = document.getElementById("popupVideoPreview");
    const videoPreviewTag = document.getElementById("popupVideoPreviewTag");
    if (videoPreview && videoPreviewTag) {
      videoPreviewTag.src = cfg.popupImagen;
      videoPreview.style.display = "block";
    }
  }
}

/** Alterna entre el bloque de "Imagen" y el de "Video" según el radio elegido */
function onCambiarTipoPopup() {
  const esVideo = document.getElementById("cfgPopupTipoVideo")?.checked;
  const wrapImagen = document.getElementById("popupImagenWrap");
  const wrapVideo = document.getElementById("popupVideoWrap");
  if (wrapImagen) wrapImagen.style.display = esVideo ? "none" : "";
  if (wrapVideo) wrapVideo.style.display = esVideo ? "" : "none";
}

/** Previsualiza el video pegado por URL, sin subir nada */
function onCambiarUrlVideoPopup() {
  const url = (document.getElementById("cfgPopupVideoUrl")?.value || "").trim();
  const statusEl = document.getElementById("popupVideoStatus");
  const preview = document.getElementById("popupVideoPreview");
  const previewTag = document.getElementById("popupVideoPreviewTag");
  if (!preview || !previewTag) return;

  if (!url) {
    preview.style.display = "none";
    previewTag.src = "";
    if (statusEl) { statusEl.className = "pm-image-status"; statusEl.textContent = ""; }
    return;
  }

  const pareceVideo = /\.(mp4|webm|ogg|mov|m4v)(\?.*)?$/i.test(url);
  if (!pareceVideo) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ La URL debería terminar en .mp4, .webm, etc."; }
  } else if (statusEl) {
    statusEl.className = "pm-image-status success";
    statusEl.textContent = "✓ URL de video lista";
  }

  previewTag.src = url;
  preview.style.display = "block";
}

async function guardarPedidoMinimoForm() {
  const input = document.getElementById("cfgPedidoMinimo");
  if (!input) return;

  const valor = Number(input.value);

  if (isNaN(valor) || valor < 0) {
    toast("Ingresá un monto válido para el pedido mínimo", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarPedidoMinimo");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({
      action: "guardarConfiguracionNegocio",
      pedidoMinimo: valor
    });
    const response = await fetchConReintento(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar el pedido mínimo", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, pedidoMinimo: valor };
    toast(`Pedido mínimo actualizado a $${valor.toLocaleString("es-AR")}`, "success");

  } catch (error) {
    console.error("Error al guardar el pedido mínimo:", error);
    toast("Error de conexión al guardar el pedido mínimo", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

function previsualizarPopup() {}  // legacy — ya no se usa

/** Lista las impresoras disponibles en Electron y las carga en el selector */
async function cargarListaImpresoras() {
  const sel = document.getElementById("cfgImpresoraSelector");
  if (!sel) return;

  const bridge = window.veekpos || window.posOffline;
  if (!bridge || typeof bridge.listarImpresoras !== "function") {
    sel.innerHTML = '<option value="">Solo disponible en la app de escritorio</option>';
    return;
  }

  sel.innerHTML = '<option value="">Cargando...</option>';
  try {
    const impresoras = await bridge.listarImpresoras();
    const guardada = localStorage.getItem("veekpos_impresora") || "";
    sel.innerHTML = '<option value="">— Predeterminada del sistema —</option>';
    impresoras.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p.name;
      opt.textContent = p.name + (p.isDefault ? " ★" : "");
      if (p.name === guardada) opt.selected = true;
      sel.appendChild(opt);
    });
    if (!guardada) sel.value = "";

    // Si había un nombre guardado pero ya no aparece en la lista actual
    // de impresoras, es la causa más común de que la impresión "en
    // teoría configurada" siga mostrando el diálogo — el sistema ya no
    // reconoce ese nombre exacto. Avisar para que el usuario la
    // vuelva a seleccionar.
    const actual = document.getElementById("cfgImpresoraActual");
    if (actual && guardada && !impresoras.some(p => p.name === guardada)) {
      actual.innerHTML = `⚠️ La impresora guardada ("${escapeHtml(guardada)}") ya no aparece en el sistema — volvé a seleccionarla y guardar.`;
      actual.style.color = "var(--red-600, #dc2626)";
    }
  } catch (e) {
    sel.innerHTML = '<option value="">Error al listar impresoras</option>';
  }
}

function guardarImpresoraSeleccionada() {
  const sel = document.getElementById("cfgImpresoraSelector");
  const nombre = sel ? sel.value : "";
  localStorage.setItem("veekpos_impresora", nombre);
  const actual = document.getElementById("cfgImpresoraActual");
  if (actual) actual.textContent = nombre ? `Impresora guardada: ${nombre}` : "Se usará la predeterminada del sistema.";
  toast(nombre ? `Impresora "${nombre}" guardada` : "Se usará la impresora predeterminada", "success");
}

function quitarImagenPopup() {
  document.getElementById("cfgPopupImagen").value = "";
  const preview = document.getElementById("popupImagenPreview");
  if (preview) preview.innerHTML = `<span class="pm-image-placeholder">Sin imagen</span>`;
  const status = document.getElementById("popupImagenStatus");
  if (status) { status.className = "pm-image-status"; status.textContent = ""; }

  const videoUrlEl = document.getElementById("cfgPopupVideoUrl");
  if (videoUrlEl) videoUrlEl.value = "";
  const videoPreview = document.getElementById("popupVideoPreview");
  if (videoPreview) videoPreview.style.display = "none";
  const videoPreviewTag = document.getElementById("popupVideoPreviewTag");
  if (videoPreviewTag) videoPreviewTag.src = "";
  const videoStatus = document.getElementById("popupVideoStatus");
  if (videoStatus) { videoStatus.className = "pm-image-status"; videoStatus.textContent = ""; }
}

async function onSeleccionarImagenPopup(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById("popupImagenStatus");
  const previewEl = document.getElementById("popupImagenPreview");

  if (!file.type.startsWith("image/")) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "Elegí un archivo de imagen (jpg, png, webp)."; }
    event.target.value = "";
    return;
  }

  // Preview local inmediata
  const localUrl = URL.createObjectURL(file);
  if (previewEl) previewEl.innerHTML = `<img src="${localUrl}" alt="" style="width:100%;height:100%;object-fit:contain;">`;
  if (statusEl) { statusEl.className = "pm-image-status uploading"; statusEl.textContent = "⏳ Optimizando imagen..."; }

  try {
    const { base64, tipoMime } = await comprimirImagenProducto(file);
    const pesoKB = Math.round((base64.length * 0.75) / 1024);

    if (statusEl) { statusEl.className = "pm-image-status uploading"; statusEl.textContent = `⏳ Subiendo a Drive... (${pesoKB}KB)`; }

    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "subirImagenProducto", imagenBase64: base64, tipoMime, codigoProducto: "POPUP_PROMO" })
    }, { timeoutMs: 40000 }); // las imágenes pesan más y tardan más en subir — no es una mutación chica
    const data = await response.json();

    if (!data.success) {
      if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ " + (data.message || "No se pudo subir la imagen."); }
      return;
    }

    document.getElementById("cfgPopupImagen").value = data.url;
    if (statusEl) { statusEl.className = "pm-image-status success"; statusEl.textContent = `✓ Imagen subida (${pesoKB}KB)`; }

  } catch (err) {
    console.error("Error al subir imagen popup:", err);
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ Error de conexión."; }
  } finally {
    URL.revokeObjectURL(localUrl);
  }
}

async function subirImagenPopup() {}  // legacy

async function guardarPopupPromoForm() {
  const activo = document.getElementById("cfgPopupActivo")?.checked ? "SI" : "NO";
  const esVideo = document.getElementById("cfgPopupTipoVideo")?.checked;
  const tipo = esVideo ? "video" : "imagen";

  const imagen = esVideo
    ? (document.getElementById("cfgPopupVideoUrl")?.value || "").trim()
    : (document.getElementById("cfgPopupImagen")?.value || "").trim();

  if (esVideo && imagen && !/\.(mp4|webm|ogg|mov|m4v)(\?.*)?$/i.test(imagen)) {
    toast("La URL del video debería terminar en .mp4, .webm, etc.", "error");
    return;
  }

  try {
    const params = new URLSearchParams({
      action: "guardarConfiguracionNegocio",
      popupImagen: imagen,
      popupActivo: activo,
      popupTipo: tipo
    });
    const res = await fetchAPI(API_URL + "?" + params.toString());
    const data = await res.json();
    if (data.success) toast("Popup guardado correctamente", "success");
    else toast(data.message || "Error al guardar", "error");
  } catch (err) {
    toast("Error de conexión", "error");
  }
}

async function quitarPopupPromo() {
  quitarImagenPopup();
  document.getElementById("cfgPopupActivo").checked = false;
  await guardarPopupPromoForm();
}

/**
 * Saves both the ticket-header fields AND the banner/tema fields together,
 * since both live in the same hoja CONFIGURACION and the backend expects
 * the full set of keys in one call to guardarConfiguracionNegocio.
 * El tema afecta a este mismo panel admin Y al catálogo web a la vez.
 */
async function guardarAparienciaForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const cfg = {
    nombre,
    subtitulo: document.getElementById("cfgSubtitulo").value.trim(),
    direccion: document.getElementById("cfgDireccion").value.trim(),
    telefono1: document.getElementById("cfgTelefono1").value.trim(),
    telefono2: document.getElementById("cfgTelefono2").value.trim(),
    pie:       document.getElementById("cfgPie").value.trim(),

    bannerTitulo:    document.getElementById("cfgBannerTitulo").value.trim(),
    bannerSubtitulo: document.getElementById("cfgBannerSubtitulo").value.trim(),
    bannerImagen:    document.getElementById("cfgBannerImagen").value.trim(),
    tema:            document.getElementById("cfgTema").value,
    gradPersonalizado: document.getElementById("cfgGradPersonalizado").checked,
    gradA:           document.getElementById("cfgGradA").value,
    gradB:           document.getElementById("cfgGradB").value,
    navbarTexto:     document.getElementById("cfgNavbarTexto").value.trim(),
    navbarIcono:     document.getElementById("cfgNavbarIcono").value.trim()
  };

  // Campos opcionales nuevos (URL del catálogo, nombre corto, íconos, SEO, footer)
  const getIfExists = (id) => { const el = document.getElementById(id); return el ? el.value.trim() : undefined; };
  const extra = {
    nombreCorto: getIfExists("cfgNombreCorto"),
    iconoUrl: getIfExists("cfgIconoUrl"),
    whatsappIconoUrl: getIfExists("cfgWhatsappIconoUrl"),
    seoTitulo: getIfExists("cfgSeoTitulo"),
    seoDescripcion: getIfExists("cfgSeoDescripcion"),
    seoKeywords: getIfExists("cfgSeoKeywords"),
    footerTitulo1: getIfExists("cfgFooterTitulo1"),
    footerTexto1: getIfExists("cfgFooterTexto1"),
    footerTexto2: getIfExists("cfgFooterTexto2"),
    footerTexto3: getIfExists("cfgFooterTexto3"),
    footerCopyright: getIfExists("cfgFooterCopyright")
  };
  Object.keys(extra).forEach(k => { if(extra[k] === undefined) delete extra[k]; });
  Object.assign(cfg, extra);

  const btn = document.getElementById("btnGuardarApariencia");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar la apariencia", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    aplicarTemaAdmin(configNegocioCache);
    toast("Apariencia guardada — ya se ve en el panel y en el catálogo", "success");

  } catch (error) {
    console.error("Error al guardar la apariencia:", error);
    toast("Error de conexión al guardar la apariencia", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/* ===================== BENEFICIOS DEL CATÁLOGO WEB (chips bajo el banner) ===================== */

const BENEFICIOS_DEFAULT = {
  beneficioWhatsappNumero: "",
  beneficioInstagramUrl:   "",
  beneficioTelefono1:      "",
  beneficioTelefono2:      "",
  beneficioDireccion:      "",
  beneficioTextoLibre1:    "",
  beneficioTextoLibre2:    ""
};

/** Loads the saved "Beneficios" config into the form (called when Configuración opens) */
function cargarBeneficiosForm(cfg) {
  document.getElementById("cfgBeneficioWhatsapp").value  = cfg.beneficioWhatsappNumero ?? BENEFICIOS_DEFAULT.beneficioWhatsappNumero;
  document.getElementById("cfgBeneficioInstagram").value = cfg.beneficioInstagramUrl   ?? BENEFICIOS_DEFAULT.beneficioInstagramUrl;
  document.getElementById("cfgBeneficioTelefono1").value = cfg.beneficioTelefono1      ?? BENEFICIOS_DEFAULT.beneficioTelefono1;
  document.getElementById("cfgBeneficioTelefono2").value = cfg.beneficioTelefono2      ?? BENEFICIOS_DEFAULT.beneficioTelefono2;
  document.getElementById("cfgBeneficioDireccion").value = cfg.beneficioDireccion      ?? BENEFICIOS_DEFAULT.beneficioDireccion;
  document.getElementById("cfgBeneficioTexto1").value    = cfg.beneficioTextoLibre1    ?? BENEFICIOS_DEFAULT.beneficioTextoLibre1;
  document.getElementById("cfgBeneficioTexto2").value    = cfg.beneficioTextoLibre2    ?? BENEFICIOS_DEFAULT.beneficioTextoLibre2;
}

/** Reads the "Beneficios" form fields and saves them to the backend (hoja CONFIGURACION) */
async function guardarBeneficiosForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const whatsapp = document.getElementById("cfgBeneficioWhatsapp").value.trim();

  const cfg = {
    nombre,

    beneficioWhatsappNumero: whatsapp || BENEFICIOS_DEFAULT.beneficioWhatsappNumero,
    beneficioInstagramUrl:   document.getElementById("cfgBeneficioInstagram").value.trim(),
    beneficioTelefono1:      document.getElementById("cfgBeneficioTelefono1").value.trim(),
    beneficioTelefono2:      document.getElementById("cfgBeneficioTelefono2").value.trim(),
    beneficioDireccion:      document.getElementById("cfgBeneficioDireccion").value.trim(),
    beneficioTextoLibre1:    document.getElementById("cfgBeneficioTexto1").value.trim(),
    beneficioTextoLibre2:    document.getElementById("cfgBeneficioTexto2").value.trim()
  };

  const btn = document.getElementById("btnGuardarBeneficios");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudieron guardar los beneficios", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    toast("Beneficios guardados — ya se ven en el catálogo", "success");

  } catch (error) {
    console.error("Error al guardar los beneficios:", error);
    toast("Error de conexión al guardar los beneficios", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Reads the CBU/titular form fields and saves them to the backend (hoja CONFIGURACION) — usados en el mensaje de WhatsApp */
async function guardarDatosTransferenciaForm() {
  const nombre = document.getElementById("cfgNombreLocal").value.trim();

  if (!nombre) {
    toast("Completá primero el nombre del local, arriba", "error");
    return;
  }

  const cfg = {
    nombre,
    cbuTransferencia: document.getElementById("cfgCbuTransferencia").value.trim(),
    nombreTitularCbu: document.getElementById("cfgNombreTitularCbu").value.trim()
  };

  const btn = document.getElementById("btnGuardarDatosTransferencia");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", ...cfg });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudieron guardar los datos de transferencia", "error");
      return;
    }

    configNegocioCache = { ...configNegocioCache, ...cfg };
    toast("Datos de transferencia guardados", "success");

  } catch (error) {
    console.error("Error al guardar los datos de transferencia:", error);
    toast("Error de conexión al guardar los datos de transferencia", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Saves the rotating top-banner messages — one per line in the textarea, joined with newlines for the backend */
async function guardarBannerTopForm() {
  const bannerTopMensajes = document.getElementById("cfgBannerTopMensajes").value.trim();

  const btn = document.getElementById("btnGuardarBannerTop");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", nombre: configNegocioCache.nombre, bannerTopMensajes });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar el banner", "error");
      return;
    }

    configNegocioCache.bannerTopMensajes = bannerTopMensajes;
    toast("Banner superior guardado — ya se ve en el catálogo", "success");

  } catch (error) {
    console.error("Error al guardar el banner superior:", error);
    toast("Error de conexión al guardar el banner", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Saves the list of blocked shipping carriers — one per line, same pattern as the top-banner messages */
async function guardarTransportesNoDisponiblesForm() {
  const transportesNoDisponibles = document.getElementById("cfgTransportesNoDisponibles").value.trim();

  const btn = document.getElementById("btnGuardarTransportesNoDisponibles");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", nombre: configNegocioCache.nombre, transportesNoDisponibles });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudieron guardar los transportes", "error");
      return;
    }

    configNegocioCache.transportesNoDisponibles = transportesNoDisponibles;
    toast("Transportes bloqueados guardados — ya rige en el catálogo", "success");

  } catch (error) {
    console.error("Error al guardar los transportes no disponibles:", error);
    toast("Error de conexión al guardar los transportes", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/* ===================== TOASTS ===================== */

function toast(mensaje, tipo) {
  const stack = document.getElementById("toastStack");
  if (!stack) return;
  const el = document.createElement("div");
  el.className = "toast-msg" + (tipo ? " " + tipo : "");
  const icon = tipo === "error" ? "⚠️" : tipo === "success" ? "✓" : "ℹ️";
  el.innerHTML = `<span>${icon}</span><span>${escapeHtml(mensaje)}</span>`;
  stack.appendChild(el);
  setTimeout(() => {
    el.style.transition = "opacity .25s";
    el.style.opacity = "0";
    setTimeout(() => el.remove(), 250);
  }, 2600);
}

/**
 * Toca un beep corto y agudo (dos tonos ascendentes) para acompañar el
 * aviso de "transferencia recibida" — sin depender de ningún archivo de
 * audio propio. Si el navegador bloqueó el audio por falta de interacción
 * previa del usuario (autoplay policy), falla en silencio: el toast visual
 * ya avisa igual.
 */
function sonarBeepTransferencia() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    [880, 1160].forEach((freq, i) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = "sine";
      osc.frequency.value = freq;
      gain.gain.setValueAtTime(0.15, ctx.currentTime + i * 0.14);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + i * 0.14 + 0.13);
      osc.connect(gain).connect(ctx.destination);
      osc.start(ctx.currentTime + i * 0.14);
      osc.stop(ctx.currentTime + i * 0.14 + 0.14);
    });
  } catch (error) { /* sin audio disponible, no es crítico */ }
}

/**
 * Solo tiene efecto dentro de la app de escritorio (Electron), que es la
 * única que monitorea transferencias de Mercado Pago (ver main.js →
 * gestionarMonitoreoTransferenciasMP). En la versión web esta función no
 * hace nada porque window.veekpos no existe ahí.
 */
function suscribirseATransferenciasMP() {
  if (!window.veekpos || !window.veekpos.onTransferenciaRecibida) return;

  window.veekpos.onTransferenciaRecibida((t) => {
    const monto = "$" + Number(t.monto).toLocaleString("es-AR");
    const quien = t.pagador ? ` de ${t.pagador}` : "";
    toast(`Transferencia recibida: ${monto}${quien}`, "success");
    sonarBeepTransferencia();
  });
}

/* ===================== METRICAS ===================== */

async function fetchConReintento(url, opciones, intentos) {
  intentos = intentos || 3;
  let ultimoError = null;

  for (let intento = 1; intento <= intentos; intento++) {
    try {
      const response = await fetch(url, opciones);
      if (!response.ok) throw new Error("HTTP " + response.status);
      return response;
    } catch (error) {
      ultimoError = error;
      if (intento < intentos) {
        // Espera creciente: 600ms, 1200ms, 1800ms... da tiempo a que
        // se libere la ejecución concurrente de las otras cajas.
        await new Promise(r => setTimeout(r, 600 * intento));
      }
    }
  }
  throw ultimoError;
}

async function cargarMetricas() {
  const CACHE_KEY = "vpos_cache_metricas";
  const CACHE_TTL = 2 * 60 * 1000; // 2 minutos

  function aplicarMetricas(data) {
    actualizarElemento("pedidosNuevos",   data.pedidosNuevos  || 0);
    actualizarElemento("totalPedidos",    data.totalPedidos   || 0);
    actualizarElemento("productosActivos", data.productosActivos || 0);
    actualizarElemento("stockBajo",       data.stockBajo || 0);
    actualizarElemento("agotados",        data.agotados  || 0);
    actualizarElemento("clientesUnicos",  data.clientesUnicos || 0);
    if (data.masVendidoCantidad > 0) {
      actualizarElemento("cantidadMasVendidos", data.masVendidoCantidad + " uds · " + (data.masVendidoNombre || ""));
    } else {
      actualizarElemento("cantidadMasVendidos", "Ver ranking →");
    }
    const ventasHoyTotal = Number(data.ventasHoy || 0);
    const ventasMesTotal = Number(data.ventasMes || 0);
    actualizarElemento("ventasHoy",     "$" + ventasHoyTotal.toLocaleString("es-AR"));
    actualizarElemento("ventasMes",     "$" + ventasMesTotal.toLocaleString("es-AR"));
    actualizarElemento("ventasTotales", "$" + ventasMesTotal.toLocaleString("es-AR"));
    const tp = "$" + Math.round(data.ticketPromedio || 0).toLocaleString("es-AR");
    actualizarElemento("ticketPromedio", tp);
    const posHoy = Number(data.ventasPOSHoy || 0);
    const posMes = Number(data.ventasPOSMes || 0);
    actualizarElemento("ventasPOSHoy", "$" + posHoy.toLocaleString("es-AR"));
    actualizarElemento("ventasPOSMes", "$" + posMes.toLocaleString("es-AR"));
    const pedHoy = Number(data.ventasPedidosHoy || 0);
    const pedMes = Number(data.ventasPedidosMes || 0);
    actualizarElemento("ventasPedidosHoy", "$" + pedHoy.toLocaleString("es-AR"));
    actualizarElemento("ventasPedidosMes", "$" + pedMes.toLocaleString("es-AR"));
    actualizarElemento("posVentasHoyBanner",  "$" + posHoy.toLocaleString("es-AR"));
    actualizarElemento("posVentasMesBanner",  "$" + posMes.toLocaleString("es-AR"));
    actualizarElemento("posTicketPromBanner", tp);
    actualizarElemento("posTotalPedBanner",   data.totalVentasPOS || 0);
  }

  // Mostrar caché inmediatamente si existe
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (raw) {
      const { ts, data } = JSON.parse(raw);
      aplicarMetricas(data);
      marcarActualizacionEnVivo();
      if (Date.now() - ts < CACHE_TTL) return; // fresco — no recargar
    }
  } catch(e) {}

  // Actualizar en background
  try {
    const response = await fetchAPI(API_URL + "?action=metricas");
    const data = await response.json();
    aplicarMetricas(data);
    marcarActualizacionEnVivo();
    try { localStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), data })); } catch(e) {}
  } catch (error) {
    console.error("Error métricas:", error);
    marcarActualizacionEnVivo(true);
  }
}

/** Updates the "live" badge on the dashboard so users can see the data is fresh */
function marcarActualizacionEnVivo(fallo) {
  const badge = document.getElementById("posLiveBadge");
  if (!badge) return;
  const hora = new Date().toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  if (fallo) {
    badge.classList.add("ps-badge-error");
    badge.innerHTML = `⚠️ Sin conexión`;
    return;
  }
  badge.classList.remove("ps-badge-error");
  badge.innerHTML = `En vivo · ${hora}`;
  badge.classList.remove("pulse");
  void badge.offsetWidth;
  badge.classList.add("pulse");
}

/* ===================== POS VENTAS RECIENTES (dashboard) ===================== */

async function cargarVentasPOS() {
  try {
    const response = await fetchAPI(API_URL + "?action=ventasPOS");
    const data = await response.json();
    ventasPOSGlobal = data.ventas || [];
    renderVentasPOSRecientes(ventasPOSGlobal);
  } catch (error) {
    // Action may not exist yet on Apps Script — silently ignore
    console.warn("ventasPOS action not available:", error);
    renderVentasPOSRecientes([]);
  }
}

function renderVentasPOSRecientes(lista) {
  const tbody = document.getElementById("tablaPOSRecientes");
  if (!tbody) return;

  if (!lista || lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3">No hay ventas del día aún</td></tr>`;
    return;
  }

  let html = "";
  lista.slice(0, 20).forEach(v => {
    const hora = v.FECHA ? new Date(v.FECHA).toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }) : "—";
    const items = v.ITEMS || v.DETALLE || "—";
    const pago  = v.FORMA_PAGO || v.PAGO || "—";
    const total = Number(v.TOTAL || 0).toLocaleString("es-AR");

    const _vidRec = String(v.VENTA_ID || v.ID || "");
    _ventasMapPOS[_vidRec] = v;

    html += `
      <tr>
        <td class="mono">${escapeHtml(String(v.VENTA_ID || v.ID || "—"))}</td>
        <td>${hora}</td>
        <td>${escapeHtml(String(items))}</td>
        <td>${escapeHtml(String(pago))}</td>
        <td class="money">$${total}</td>
        <td>
          <button class="btn btn-sm btn-outline-secondary"
            onclick="imprimirVentaDesdeData(_ventasMapPOS['${escapeJsAttr(_vidRec)}'])">🖨️</button>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

/* ===================== VENTAS POS — HISTORIAL (sección dedicada) ===================== */

let ventasPOSHistorialGlobal = [];

/** Loads POS sales history for the date range selected in the filter inputs (defaults to last 60 days) */
async function cargarVentasPOSHistorial() {
  const tbody = document.getElementById("tablaVentasPOSHistorial");
  if (!tbody) return;

  const desdeInput = document.getElementById("vpDesde");
  const hastaInput = document.getElementById("vpHasta");

  if (desdeInput && !desdeInput.value) {
    const hace30 = new Date();
    hace30.setDate(hace30.getDate() - 30);
    desdeInput.value = _fechaISOEnZona(hace30);
  }
  if (hastaInput && !hastaInput.value) {
    hastaInput.value = _hoyISO();
  }

  const cacheKey = "ventasPOS_" + (desdeInput?.value || "") + "_" + (hastaInput?.value || "");
  const cached = cacheGet("ventasPOS");
  if (cached && cached.data?.key === cacheKey) {
    ventasPOSHistorialGlobal = cached.data.ventas;
    filtrarVentasPOSHistorial(); // respeta el filtro de forma de pago/estado/búsqueda activo, en vez de mostrar siempre todo
    if (!cached.stale) return;
  } else {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-3">Cargando ventas...</td></tr>`;
  }

  try {
    const params = new URLSearchParams({ action: "ventasPOSHistorial" });
    if (desdeInput && desdeInput.value) params.set("desde", desdeInput.value);
    if (hastaInput && hastaInput.value) params.set("hasta", hastaInput.value);

    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    ventasPOSHistorialGlobal = data.ventas || [];
    cacheSet("ventasPOS", { key: cacheKey, ventas: ventasPOSHistorialGlobal });
    filtrarVentasPOSHistorial(); // idem — no pisar el filtro activo con la lista completa

  } catch (error) {
    console.error("Error al cargar historial de ventas POS:", error);
    if (!cached) tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-3">Error al cargar el historial</td></tr>`;
  }
}

function renderVentasPOSHistorial(lista) {
  const tbody = document.getElementById("tablaVentasPOSHistorial");
  if (!tbody) return;

  if (!lista || lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-3">No hay ventas en el rango seleccionado</td></tr>`;
    return;
  }

  const CHUNK = 60;
  let idx = 0;
  tbody.innerHTML = "";
  // Si llega otro render (ej. la copia guardada y enseguida la fresca), los
  // pedazos que quedaban del anterior no se agregan: antes se mezclaban y
  // la tabla mostraba ventas repetidas.
  const turno = (renderVentasPOSHistorial._turno = (renderVentasPOSHistorial._turno || 0) + 1);

  const renderChunk = () => {
    if (turno !== renderVentasPOSHistorial._turno) return;
    const frag = document.createDocumentFragment();
    const fin = Math.min(idx + CHUNK, lista.length);
    for (; idx < fin; idx++) {
      const v = lista[idx];
      const fechaObj = v.FECHA ? new Date(v.FECHA) : null;
      const fecha = fechaObj ? fechaObj.toLocaleDateString("es-AR") : "—";
      const hora  = fechaObj ? fechaObj.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }) : "—";
      const items = v.ITEMS || v.DETALLE || "—";
      const pago  = v.FORMA_PAGO || v.PAGO || "—";
      const total = Number(v.TOTAL || 0).toLocaleString("es-AR");
      const anulada = String(v.ANULADA || "").toUpperCase() === "SI";
      const motivo  = v.MOTIVO_ANULACION || "";
      const _vid = String(v.VENTA_ID || v.ID || "");
      _ventasMapPOS[_vid] = v;
      // Resumen de ventas viejas (ver compactarVentasAntiguas en code.gs): no se reimprime ni se anula
      const esResumen = /^RES-/i.test(_vid);
      const tr = document.createElement("tr");
      if (esResumen) tr.className = "fila-resumen-ventas";
      if (anulada) tr.style.cssText = "background:#fff0f0;color:#b91c1c;";
      const badgeAnulada = anulada
        ? `<span class="badge bg-danger ms-1" style="font-size:10px;text-decoration:none;">ANULADA${motivo ? " · " + escapeHtml(motivo) : ""}</span>`
        : "";
      tr.innerHTML = `
        <td class="mono" style="${anulada ? 'text-decoration:line-through;' : ''}">${escapeHtml(String(v.VENTA_ID || v.ID || "—"))}${badgeAnulada}</td>
        <td>${fecha}</td><td>${hora}</td>
        <td>${escapeHtml(String(items))}</td>
        <td>${escapeHtml(String(pago))}</td>
        <td class="money" style="${anulada ? 'text-decoration:line-through;' : ''}">$${total}</td>
        <td>${esResumen ? `<span class="text-muted" style="font-size:12px;" title="Las ventas de hace más de 6 meses se guardan resumidas por día (el detalle completo quedó en Drive → Respaldo de ventas)">📦 ${escapeHtml(String(v.OBSERVACIONES || "Resumen del día"))}</span>` : `
          <button class="btn btn-sm btn-outline-secondary"
            onclick="imprimirVentaDesdeData(_ventasMapPOS['${escapeJsAttr(_vid)}'])" title="Reimprimir">Reimprimir</button>${typeof esModoDeposito === "function" && esModoDeposito() && !anulada ? `
          <button class="btn btn-sm btn-outline-secondary ms-1"
            onclick="abrirBoletaDesdeVenta(_ventasMapPOS['${escapeJsAttr(_vid)}'])" title="Boleta de entrega A4">Boleta</button>` : ""}
          <button class="btn btn-sm btn-outline-danger ms-1"
            onclick="eliminarVentaPOS(_ventasMapPOS['${escapeJsAttr(_vid)}'])"
            ${anulada ? 'disabled title="Ya anulada"' : 'title="Anular"'}>Anular</button>${window.ArcaUI ? ArcaUI.celdaVenta(v, anulada) : ""}`}
        </td>`;
      frag.appendChild(tr);
    }
    tbody.appendChild(frag);
    if (idx < lista.length) requestAnimationFrame(renderChunk);
  };
  requestAnimationFrame(renderChunk);
}

/**
 * Deletes a sale (and returns its stock) after confirmation. Receives
 * the full sale object (as rendered in the table) so the backend gets
 * the cart directly, without needing to re-read DETALLE_VENTAS.
 */
/** Muestra un modal para ingresar el motivo de anulación */
function eliminarVentaPOS(venta) {
  const ventaId = venta.VENTA_ID || venta.ID;

  // Si ya está anulada, no hacer nada
  if (String(venta.ANULADA || "").toUpperCase() === "SI") {
    toast("Esta venta ya está anulada", "error");
    return;
  }

  // Guardar la venta en variable global para usarla al confirmar
  window._ventaParaAnular = venta;

  // Limpiar y mostrar modal
  document.getElementById("motivoAnulacionInput").value = "";
  document.getElementById("motivoAnulacionVentaId").textContent = ventaId;
  document.getElementById("btnConfirmarAnulacion").disabled = false;
  document.getElementById("btnConfirmarAnulacion").textContent = "Anular venta";
  document.getElementById("modalAnulacionBackdrop").classList.add("show");
  setTimeout(() => document.getElementById("motivoAnulacionInput").focus(), 100);
}

function cerrarModalAnulacion() {
  document.getElementById("modalAnulacionBackdrop").classList.remove("show");
  window._ventaParaAnular = null;
}

async function confirmarAnulacionVenta() {
  const venta = window._ventaParaAnular;
  if (!venta) return;

  const motivoInput = document.getElementById("motivoAnulacionInput");
  const motivoError = document.getElementById("motivoAnulacionError");
  const motivo = motivoInput.value.trim();
  if (!motivo) {
    motivoInput.style.borderColor = "var(--red-500)";
    if (motivoError) motivoError.style.display = "block";
    motivoInput.focus();
    return;
  }
  motivoInput.style.borderColor = "";
  if (motivoError) motivoError.style.display = "none";

  const btn = document.getElementById("btnConfirmarAnulacion");
  btn.disabled = true;
  btn.textContent = "Anulando...";

  try {
    const params = new URLSearchParams({
      action: "eliminarVenta",
      ventaId: venta.VENTA_ID || venta.ID,
      carrito: venta.CARRITO || "[]",
      motivo: motivo
    });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo anular la venta", "error");
      btn.disabled = false;
      btn.textContent = "Anular venta";
      return;
    }

    toast("Venta anulada — sigue en el historial pero ya no cuenta en caja", "success");
    if (window.ArcaUI) ArcaUI.trasAnulacion(venta, motivo); // nota de crédito si estaba facturada
    cerrarModalAnulacion();
    cargarVentasPOSHistorial();
    productosPOS = [];

  } catch (error) {
    console.error("Error al anular venta:", error);
    toast("Error de conexión al anular la venta", "error");
    btn.disabled = false;
    btn.textContent = "Anular venta";
  }
}

/** Client-side filter by sale id or payment method, over the already-loaded historial */
function filtrarVentasPOSHistorial() {
  const termino   = (document.getElementById("vpBuscar")?.value || "").toLowerCase().trim();
  const formaPago = (document.getElementById("vpFormaPago")?.value || "").toUpperCase();
  const estado    = (document.getElementById("vpEstado")?.value || "");

  const filtradas = ventasPOSHistorialGlobal.filter(v => {
    const id     = String(v.VENTA_ID || v.ID || "").toLowerCase();
    const items  = String(v.ITEMS || v.DETALLE || "").toLowerCase();
    const pago   = String(v.FORMA_PAGO || v.PAGO || "").toUpperCase().trim();
    const anulada = String(v.ANULADA || "").toUpperCase() === "SI";

    if (termino && !id.includes(termino) && !items.includes(termino)) return false;
    if (formaPago && pago !== formaPago) return false;
    if (estado === "activas"  &&  anulada) return false;
    if (estado === "anuladas" && !anulada) return false;
    return true;
  });

  renderVentasPOSHistorial(filtradas);
}

/* ===================== NAVEGACION ===================== */

// Recuerda cuándo se cargó cada sección por última vez, para no volver a
// pedirle al backend lo mismo si el cajero entra y sale de una sección
// en pocos segundos (ej: mirar Pedidos, volver al POS, volver a Pedidos).
// El timer de 15s de arriba sigue refrescando la sección activa con
// normalidad; esto solo evita pedidos duplicados al navegar.
const ULTIMA_CARGA_SECCION = {};
const VENCIMIENTO_CACHE_MS = 3 * 60 * 1000; // 3 minutos

function cargarSiVencido(clave, fn) {
  const ahora = Date.now();
  const ultima = ULTIMA_CARGA_SECCION[clave] || 0;
  if (ahora - ultima < VENCIMIENTO_CACHE_MS) return; // todavía fresco, no repetir
  ULTIMA_CARGA_SECCION[clave] = ahora;
  fn();
}

/** Cambia de pestaña dentro de la sección Configuración, sin recargar nada */
function mostrarConfigTab(tabId, btnEl) {
  document.querySelectorAll(".cfg-tab-pane").forEach(p => p.classList.remove("active"));
  const pane = document.getElementById("cfgTab-" + tabId);
  if (pane) pane.classList.add("active");

  document.querySelectorAll(".cfg-tab-btn").forEach(b => b.classList.remove("active"));
  if (btnEl) btnEl.classList.add("active");
  else {
    const btn = document.querySelector(`.cfg-tab-btn[data-cfg-tab="${tabId}"]`);
    if (btn) btn.classList.add("active");
  }

  try { localStorage.setItem("veekpos_cfg_tab", tabId); } catch (e) {}
  if (tabId === "sistema" && typeof cargarEstadoVentasResumidas === "function") cargarEstadoVentasResumidas();
  if (tabId === "sistema") { cargarEstadoActualizacion(); cargarTarjetaPerfiles(); if (typeof depTarjetaTraspasosConfig === "function") depTarjetaTraspasosConfig(); }
  if (tabId === "catalogo") cargarEstadoCatalogoPDF();
}

/**
 * Oculta las pestañas de Configuración que no tienen ninguna tarjeta
 * visible (ej.: en la versión web no hay Mercado Pago ni impresora, que
 * son solo de la app de escritorio). Si la pestaña activa quedó oculta,
 * se pasa a "Negocio".
 */
function actualizarPestanasConfigVisibles() {
  document.querySelectorAll(".cfg-tab-pane").forEach(pane => {
    const tab = pane.id.replace("cfgTab-", "");
    const btn = document.querySelector(`.cfg-tab-btn[data-cfg-tab="${tab}"]`);
    if (!btn) return;
    const hayVisible = [...pane.children].some(c => getComputedStyle(c).display !== "none");
    btn.style.display = (hayVisible && !btn.dataset.ocultoDeposito) ? "" : "none"; // en el depósito no hay catálogo web
  });
  const activo = document.querySelector(".cfg-tab-btn.active");
  if (activo && activo.style.display === "none") mostrarConfigTab("negocio");
}

/* ===================== CATÁLOGO PDF (Configuración → Catálogo web) =====================
   Lo arma la app de escritorio (desktop/src/catalogo-pdf.js) y lo sube al
   backend; acá se muestra el estado y se puede generar a mano. En la
   versión web solo se ve el estado. */

/* =========================================================
   ACTUALIZACIONES DE LA APP (Configuración → Sistema)
   La app consulta al servidor de licencias; actualizador.js baja la
   versión nueva de fondo y la instala al cerrarse. Acá solo se muestra.
========================================================= */
let _avisoActualizacionMostrado = "";

function iniciarAvisosActualizacion() {
  if (!window.veekpos || typeof window.veekpos.alCambiarActualizacion !== "function") return;
  window.veekpos.alCambiarActualizacion((estado) => {
    mostrarEstadoActualizacion(estado);
    avisarActualizacionUnaVez(estado);
  });
}

/** Un toast por versión y situación (no en cada revisión) */
function avisarActualizacionUnaVez(estado) {
  if (!estado || !estado.disponible) return;
  const clave = estado.estado + "|" + estado.disponible.version;
  if (clave === _avisoActualizacionMostrado) return;
  if (estado.estado === "lista") {
    _avisoActualizacionMostrado = clave;
    toast("La versión " + estado.disponible.version + " está lista: se instala sola al cerrar la app.", "success");
  } else if (estado.estado === "esperando-backend" && obtenerRolActual() === "admin") {
    _avisoActualizacionMostrado = clave;
    toast("Hay una versión nueva de VeekPOS (" + estado.disponible.version + "), pero antes hay que actualizar el sistema del negocio. Coordiná con soporte.");
  }
}

async function cargarEstadoActualizacion() {
  const card = document.getElementById("cardActualizaciones");
  if (!card) return;
  if (!window.veekpos || typeof window.veekpos.estadoActualizacion !== "function") { card.style.display = "none"; return; }
  card.style.display = "";
  try { mostrarEstadoActualizacion(await window.veekpos.estadoActualizacion()); }
  catch (e) { document.getElementById("actEstado").textContent = "No se pudo leer el estado de las actualizaciones."; }
}

function mostrarEstadoActualizacion(e) {
  const card = document.getElementById("cardActualizaciones");
  if (!card || !e) return;
  const disp = e.disponible || null;
  const badge = document.getElementById("actBadge");
  const estadoEl = document.getElementById("actEstado");
  const progreso = document.getElementById("actProgreso");
  const notas = document.getElementById("actNotas");
  const btnInstalar = document.getElementById("btnActInstalar");

  document.getElementById("actVersionApp").textContent = e.versionActual || "—";
  document.getElementById("actVersionBackend").textContent = e.versionBackend
    ? e.versionBackend + (e.backendDesactualizado && disp && disp.versionBackend ? " (hay una más nueva: " + disp.versionBackend + ")" : "")
    : "sin dato";

  const v = disp ? disp.version : "";
  const textos = {
    "sin-datos": ["…", "", "Todavía no se buscaron actualizaciones (se buscan solas unos minutos después de abrir la app)."],
    "al-dia": ["Al día", "ok", "Tenés la última versión."],
    "descargando": ["Bajando", "warn", "Bajando la versión " + v + " de fondo… Podés seguir trabajando."],
    "lista": ["Lista", "ok", "La versión " + v + " está lista. Se instala sola la próxima vez que cierres la app (tarda alrededor de un minuto)."],
    "esperando-backend": ["En espera", "warn", "Hay una versión nueva (" + v + "), pero necesita que el sistema del negocio esté en la " + (disp && disp.backendMinimo) + " o más nueva. Se instala sola en cuanto se actualice. Coordiná con soporte."],
    "error": ["Error", "bad", "No se pudo bajar la actualización: " + (e.error || "error desconocido") + ". Se vuelve a intentar sola."],
    "desactivada": ["—", "", "Las actualizaciones automáticas funcionan en la app instalada en Windows."]
  };
  const [txtBadge, clase, mensaje] = textos[e.estado] || textos["sin-datos"];
  badge.textContent = txtBadge;
  badge.className = "licencia-badge" + (clase ? " " + clase : "");
  estadoEl.textContent = mensaje + (e.backendDesactualizado && e.estado === "al-dia" ? " Hay una actualización del sistema del negocio (" + (disp && disp.versionBackend) + "): coordiná con soporte." : "");

  progreso.style.display = e.estado === "descargando" ? "" : "none";
  progreso.firstElementChild.style.width = Math.round(e.progreso || 0) + "%";
  const mostrarNotas = disp && disp.notas && ["descargando", "lista", "esperando-backend"].includes(e.estado);
  notas.style.display = mostrarNotas ? "" : "none";
  notas.textContent = mostrarNotas ? "Qué cambia en la " + v + ":\n" + disp.notas : "";
  btnInstalar.style.display = e.estado === "lista" ? "" : "none";
}

async function revisarActualizacionBtn() {
  const btn = document.getElementById("btnActRevisar");
  btn.disabled = true; btn.textContent = "Buscando…";
  try {
    const e = await window.veekpos.revisarActualizacion();
    mostrarEstadoActualizacion(e);
    if (e && e.errorConsulta) toast("No se pudo consultar: " + e.errorConsulta, "error");
    else if (e && e.estado === "al-dia") toast("Tenés la última versión.", "success");
  } catch (err) {
    toast("No se pudo buscar actualizaciones: " + err.message, "error");
  } finally {
    btn.disabled = false; btn.textContent = "↻ Buscar actualizaciones";
  }
}

function instalarActualizacionBtn() {
  confirmarAccion(
    "La app se cierra, instala la versión nueva y vuelve a abrirse sola (alrededor de un minuto). Terminá la venta en curso antes de seguir.",
    async () => {
      const r = await window.veekpos.instalarActualizacion();
      if (!r || !r.success) toast((r && r.message) || "No se pudo instalar", "error");
    },
    "Reiniciar y actualizar"
  );
}

async function cargarEstadoCatalogoPDF() {
  const box = document.getElementById("catPdfEstado");
  if (!box) return;
  const bridge = window.veekpos;
  const enEscritorio = !!(bridge && typeof bridge.generarCatalogoPDF === "function");
  let local = null;
  if (enEscritorio) { try { local = await bridge.estadoLocalCatalogoPDF(); } catch (e) {} }

  const ver = document.getElementById("btnCatPdfVer");
  const previa = document.getElementById("btnCatPdfPrevia");
  const publicar = document.getElementById("btnCatPdfPublicar");
  const puedeGenerar = enEscritorio; // cualquier caja (principal o secundaria)
  if (previa) previa.style.display = puedeGenerar ? "" : "none";
  if (publicar) publicar.style.display = puedeGenerar ? "" : "none";
  const avisoUrl = document.getElementById("catPdfAvisoUrl");
  if (avisoUrl) avisoUrl.style.display = configNegocioCache && configNegocioCache.urlCatalogo ? "none" : "";

  let info = null;
  try { info = await (await fetchAPI(API_URL + "?action=catalogoPDFInfo")).json(); } catch (e) {}
  const partes = [];
  if (info && info.success) {
    const f = info.fecha ? new Date(info.fecha) : null;
    partes.push(`<div><b>Publicado</b>${f ? " el " + f.toLocaleDateString("es-AR") + " a las " + f.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }) : ""}${info.totalProductos ? " · " + escapeHtml(String(info.totalProductos)) + " productos" : ""}</div>`);
    if (ver) { ver.href = info.url; ver.style.display = ""; }
  } else {
    partes.push(`<div>${info ? "Todavía no hay un catálogo publicado." : "No se pudo consultar el catálogo publicado (¿sin internet?)."}</div>`);
    if (ver) ver.style.display = "none";
  }
  if (!enEscritorio) partes.push(`<div class="text-muted" style="font-size:12.5px;">Se genera desde la app de escritorio (cualquier caja).</div>`);
  // Avisos de la última generación en esta caja (banner o fotos que no se pudieron bajar)
  const av = local && local.ultimosAvisos;
  if (av && av.avisos && av.avisos.length) {
    partes.push(`<div style="color:var(--amber-600, #b45309);font-size:12.5px;margin-top:2px;">Última generación en esta caja (${new Date(av.fecha).toLocaleString("es-AR")}): ${av.avisos.map(escapeHtml).join(" ")}</div>`);
  }
  const publicadoDespues = info && info.success && info.fecha && local && local.ultimoError && new Date(info.fecha) > new Date(local.ultimoError.fecha);
  if (local && local.ultimoError && !publicadoDespues && (!local.ultimoOk || local.ultimoError.fecha > local.ultimoOk)) {
    partes.push(`<div style="color:var(--red-500);font-size:12.5px;margin-top:2px;">Último intento con error (${new Date(local.ultimoError.fecha).toLocaleString("es-AR")}): ${escapeHtml(local.ultimoError.message)}</div>`);
  }
  box.innerHTML = `<div style="width:100%;">${partes.join("")}</div>`;
  if (local && local.generando) _mostrarProgresoCatalogoPDF({ etapa: "datos" });
}

function _mostrarProgresoCatalogoPDF(p) {
  const el = document.getElementById("catPdfProgreso");
  if (!el || !p) return;
  const textos = {
    datos: "Leyendo productos y configuración…",
    fotos: `Preparando fotos (${p.hechas || 0} de ${p.total || 0})…`,
    armando: "Armando las páginas…",
    imprimiendo: `Generando el PDF${p.paginas ? " (" + p.paginas + " páginas)" : ""}…`,
    subiendo: `Publicando (${p.hechas || 0} de ${p.total || 0})…`
  };
  if (p.etapa === "fin") { el.style.display = "none"; return; }
  el.style.display = "";
  el.textContent = textos[p.etapa] || "Generando…";
}

async function generarCatalogoPDFBtn(publicar) {
  const bridge = window.veekpos;
  if (!bridge || typeof bridge.generarCatalogoPDF !== "function") return;
  const botones = ["btnCatPdfPrevia", "btnCatPdfPublicar"].map(id => document.getElementById(id)).filter(Boolean);
  botones.forEach(b => b.disabled = true);
  if (typeof bridge.alProgresoCatalogoPDF === "function") bridge.alProgresoCatalogoPDF(_mostrarProgresoCatalogoPDF);
  _mostrarProgresoCatalogoPDF({ etapa: "datos" });
  try {
    const r = await bridge.generarCatalogoPDF(!!publicar);
    if (!r || !r.success) { toast("No se pudo generar el catálogo: " + ((r && r.message) || "error"), "error"); return; }
    const mb = (r.bytes / 1048576).toFixed(1).replace(".", ",");
    toast(publicar
      ? `Catálogo publicado: ${r.totalProductos} productos, ${r.paginas} páginas (${mb} MB)`
      : (r.aviso || `Vista previa lista: ${r.totalProductos} productos, ${r.paginas} páginas (${mb} MB)`), r.aviso ? "error" : "success");
    if (r.avisos && r.avisos.length) setTimeout(() => toast(r.avisos.join(" "), "error"), 600);
  } catch (e) {
    toast("No se pudo generar el catálogo: " + (e.message || e), "error");
  } finally {
    botones.forEach(b => b.disabled = false);
    _mostrarProgresoCatalogoPDF({ etapa: "fin" });
    cargarEstadoCatalogoPDF();
  }
}

/** Atajos del Inicio: abre Productos ya filtrado (stock bajo, sin stock…) */
function irAProductosFiltrados(estado) {
  const sel = document.getElementById("filtroEstadoProducto");
  const buscar = document.getElementById("buscarProducto");
  const cat = document.getElementById("filtroCategoriaProductos");
  if (buscar) buscar.value = "";
  if (cat) cat.value = "";
  if (sel) sel.value = estado;
  mostrarSeccion("productos");
  if (productosAdminGlobal && productosAdminGlobal.length) filtrarProductos();
}

/** Atajo del Inicio: ranking de más vendidos (Reportes → Productos) */
function irAMasVendidos() {
  if (obtenerRolActual() === "deposito") { toast("Tu usuario no tiene acceso a los reportes de ventas", "error"); return; }
  _repTabActiva = "productos";
  mostrarSeccion("reportes");
}

/**
 * El POS ocupa exactamente el alto de la ventana (escritorio/laptop): se
 * mide dónde empieza y se le da el resto. Así ninguna parte queda abajo
 * del borde en ninguna pantalla, y la página no scrollea; solo la grilla
 * de productos (y la lista del ticket si es muy larga).
 */
function ajustarAltoPOS() {
  const pos = document.getElementById("pos");
  const wrap = pos && pos.querySelector(".pos-wrap");
  if (!wrap) return;
  if (window.innerWidth <= 768 || pos.style.display === "none") { wrap.style.height = ""; wrap.style.maxHeight = ""; return; }
  const content = document.querySelector(".content");
  const abajo = content ? parseFloat(getComputedStyle(content).paddingBottom) || 0 : 0;
  const arriba = wrap.getBoundingClientRect().top + window.scrollY;
  const alto = Math.max(440, Math.floor(window.innerHeight - arriba - abajo));
  wrap.style.height = alto + "px";
  wrap.style.maxHeight = alto + "px";
}
window.addEventListener("resize", () => { if (document.body.classList.contains("en-pos")) ajustarAltoPOS(); });

function mostrarSeccion(id) {
  // "Reportes de Compras" ahora es la pestaña "Compras y reposición" de Reportes
  if (id === "reportesCompras") { _repTabActiva = "compras"; id = "reportes"; }
  if (!seccionPermitidaParaRol(id)) {
    toast("No tenés permiso para acceder a esta sección", "error");
    id = "dashboard";
  }

  document.querySelectorAll(".seccion").forEach(sec => { sec.style.display = "none"; });

  const seccion = document.getElementById(id);
  if (seccion) seccion.style.display = "block";
  document.body.classList.toggle("en-pos", id === "pos");
  if (id === "pos") { requestAnimationFrame(ajustarAltoPOS); setTimeout(ajustarAltoPOS, 250); }

  // Sync sidebar links
  document.querySelectorAll("#navLinks a").forEach(a => {
    a.classList.toggle("active", a.getAttribute("data-target") === id);
  });

  // Sync bottom nav links
  document.querySelectorAll("#bottomNavLinks a").forEach(a => {
    a.classList.toggle("active", a.getAttribute("data-target") === id);
  });

  dolarBurbujaAlCambiarSeccion(id);
  if (id === "dashboard") cargarSiVencido("dashboard", () => { cargarMetricas(); cargarVentasPOS(); });
  if (id === "pedidos")   cargarSiVencido("pedidos", cargarPedidos);
  if (id === "productos") cargarSiVencido("productos", cargarProductos);
  if (id === "ingresoProductos") {
    cargarProductos();
    cargarHistorialIngresos();
    tcAlAbrirSeccion();
    const ipbFecha = document.getElementById("ipbFecha");
    if (ipbFecha && !ipbFecha.value) ipbFecha.value = _hoyISO();
    setTimeout(() => document.getElementById("ipCodigoScan")?.focus(), 100);
  }
  if (id === "ventasPOS") cargarSiVencido("ventasPOS", cargarVentasPOSHistorial);
  if (id === "cotizaciones" && typeof cargarCotizaciones === "function") cargarCotizaciones();
  if (id === "traspasos" && typeof cargarTraspasos === "function") cargarTraspasos();
  if (id === "configuracion") {
    let tabInicial = "negocio";
    try { tabInicial = localStorage.getItem("veekpos_cfg_tab") || "negocio"; } catch (e) {}
    // Nombres de pestañas anteriores a la reorganización
    tabInicial = ({ general: "negocio", productos: "pedidos", pagos: "sistema", hardware: "equipo", avanzado: "sistema" })[tabInicial] || tabInicial;
    if (!document.getElementById("cfgTab-" + tabInicial)) tabInicial = "negocio";
    mostrarConfigTab(tabInicial);

    cargarConfigNegocioForm();
    actualizarEstadoUSBPrint();
    // En Electron: mostrar tarjetas específicas de escritorio
    if (typeof window.veekpos !== "undefined" || typeof window.posOffline !== "undefined") {
      const bridge = window.veekpos || window.posOffline;
      // Tarjeta impresora
      const tImpresora = document.getElementById("cardImpresoraPOS");
      if (tImpresora) {
        tImpresora.style.display = "block";
        if (typeof cargarListaImpresoras === "function") cargarListaImpresoras();
        const guardada = localStorage.getItem("veekpos_impresora") || "";
        const actual = document.getElementById("cfgImpresoraActual");
        if (actual) actual.textContent = guardada
          ? `Impresora guardada: ${guardada}`
          : "Sin impresora guardada — se usará la predeterminada del sistema.";
      }
      // Tarjeta conexión API
      const tConexion = document.getElementById("cardConexionNegocio");
      if (tConexion) {
        tConexion.style.display = "block";
        // bridge.obtenerApiUrl?.() por sí solo solo protege el LLAMADO
        // a la función — si no existe, da undefined, y encadenar
        // .then()/.catch() sin su propio "?." revienta con un error
        // real (Cannot read properties of undefined). Como esta
        // función no está expuesta todavía desde preload.js, esto
        // cortaba la ejecución acá y las tarjetas de Mercado Pago y
        // licencia que venían DESPUÉS en este mismo bloque nunca
        // llegaban a mostrarse.
        const promesaUrl = bridge.obtenerApiUrl?.();
        if (promesaUrl && typeof promesaUrl.then === "function") {
          promesaUrl.then(url => {
            const el = document.getElementById("conexionNegocioUrlActual");
            if (el) el.textContent = url ? "Conectada a: " + url : "No hay URL configurada.";
            const input = document.getElementById("conexionNegocioUrl");
            if (input && url) input.value = url;
          }).catch(() => {});
        }
      }
      // Tarjeta licencia
      if (typeof mostrarEstadoLicenciaEnConfig === "function") mostrarEstadoLicenciaEnConfig();
      // Tarjeta multicaja
      const tRed = document.getElementById("cardMultiCajaRed");
      if (tRed) {
        tRed.style.display = "block";
        cargarConfigRedForm();
      }
      // Tarjeta MercadoPago
      const tMp = document.getElementById("cardMercadoPago");
      if (tMp) {
        tMp.style.display = "block";
        if (typeof mostrarEstadoMercadoPagoEnConfig === "function") mostrarEstadoMercadoPagoEnConfig();
      }
    }
    actualizarPestanasConfigVisibles();
    setTimeout(actualizarPestanasConfigVisibles, 600); // la tarjeta de licencia se muestra async
  }

  if (id === "pos") {
    asegurarProductosPOS().then(renderPosGrid);
    setTimeout(() => {
      const input = document.getElementById("posBusqueda");
      if (input) input.focus();
    }, 80);
  }

  if (id === "cierreCaja") {
    const selector = document.getElementById("ccFechaSelector");
    if (selector && !selector.value) {
      const ayer = new Date();
      ayer.setDate(ayer.getDate() - 1);
      selector.value = _fechaISOEnZona(ayer);
    }
    cargarResumenCierreCaja(selector ? selector.value : null);
  }
  if (id === "movimientosCaja") cargarMovimientosCajaHoy();
  if (id === "cupones") cargarCupones();
  if (id === "pedidos") marcarPedidosVistos();
  if (id === "reportes") {
    // Pasado el tiempo de caché de secciones, se vuelve a pedir todo al entrar
    cargarSiVencido("reportes", () => { Object.keys(_repTabRangoCargado).forEach(k => _repTabRangoCargado[k] = null); });
    mostrarTabReportes(_repTabActiva || _repTabPorDefecto());
  }
}

/* ===================================================================
   BOTONES "↻ ACTUALIZAR": respuesta visible al tocarlos
   Para todos los botones que empiezan con ↻ (o): el ícono gira y
   dice "Actualizando…" mientras haya pedidos al servidor en curso;
   al terminar muestra "✓ Actualizado" (o "Sin conexión") un momento.
=================================================================== */
function _esBotonActualizar(btn) {
  if (!btn || btn.tagName !== "BUTTON" || btn.dataset.sinAnimacion !== undefined) return false;
  return /^\s*↻/.test(btn.textContent || "");
}

function _animarBotonActualizar(btn) {
  if (btn.classList.contains("btn-actualizando")) return;
  const original = btn.innerHTML;
  const anchoOriginal = btn.offsetWidth;
  const soloIcono = (btn.textContent || "").replace(/↻/g, "").trim() === "";
  const fallosAntes = _redEnCurso.fallos;
  const inicio = Date.now();
  let vioPedido = false, quietoDesde = 0;

  btn.classList.add("btn-actualizando");
  btn.setAttribute("aria-busy", "true");
  btn.style.minWidth = anchoOriginal + "px";
  btn.disabled = true;
  btn.innerHTML = '<span class="giro" aria-hidden="true">↻</span>' + (soloIcono ? "" : " Actualizando…");

  const terminar = () => {
    const fallo = _redEnCurso.fallos > fallosAntes;
    btn.classList.remove("btn-actualizando");
    btn.classList.add(fallo ? "btn-actualizado-error" : "btn-actualizado");
    btn.innerHTML = soloIcono ? (fallo ? "⚠️" : "✓") : (fallo ? "⚠️ Sin conexión" : "✓ Actualizado");
    btn.removeAttribute("aria-busy");
    setTimeout(() => {
      btn.classList.remove("btn-actualizado", "btn-actualizado-error");
      btn.innerHTML = original;
      btn.disabled = false;
      btn.style.minWidth = "";
    }, fallo ? 2200 : 1300);
  };

  const revisar = () => {
    const t = Date.now() - inicio;
    if (_redEnCurso.n > 0) { vioPedido = true; quietoDesde = 0; }
    else if (!quietoDesde) quietoDesde = Date.now();
    const quieto = quietoDesde && Date.now() - quietoDesde > 250;
    // Mínimo 600 ms para que se vea; si no hubo pedido al servidor (datos
    // frescos en memoria) se termina igual; tope 60 s.
    if ((t > 600 && quieto && (vioPedido || t > 900)) || t > 60000) { terminar(); return; }
    setTimeout(revisar, 120);
  };
  setTimeout(revisar, 120);
}

document.addEventListener("click", (e) => {
  const btn = e.target && e.target.closest ? e.target.closest("button") : null;
  if (!_esBotonActualizar(btn) || btn.disabled) return;
  const antes = btn.innerHTML;
  // Después de que corra el onclick del botón (que arranca el pedido).
  // Si esa función ya cambió el botón por su cuenta, no se pisa.
  setTimeout(() => { if (btn.innerHTML === antes && !btn.disabled) _animarBotonActualizar(btn); }, 0);
}, true);

/* ===================================================================
   CONFIGURACIÓN → SISTEMA → Historial de ventas en la planilla
   (ver compactarVentasAntiguas en code.gs)
=================================================================== */
let _estadoVentasResumidas = null;

function _fechaDMA(iso) {
  if (!iso) return "";
  const d = /^\d{4}-\d{2}-\d{2}$/.test(iso) ? new Date(iso + "T12:00:00") : new Date(iso);
  return isNaN(d) ? String(iso) : d.toLocaleDateString("es-AR");
}

async function cargarEstadoVentasResumidas() {
  const caja = document.getElementById("ventasResumidasEstado");
  const btn = document.getElementById("btnResumirVentas");
  if (!caja) return;
  caja.textContent = "Cargando…";
  try {
    const res = await fetchAPI(API_URL + "?action=estadoVentasResumidas");
    const d = await res.json();
    if (!d.success) throw new Error(d.message || "No se pudo leer el estado");
    _estadoVentasResumidas = d;
    const num = n => Number(n || 0).toLocaleString("es-AR");
    const lineas = [];
    if (d.desactivado) {
      lineas.push("Desactivado (MESES_DETALLE_VENTAS = 0 en code.gs): se guarda todo el detalle.");
    } else {
      lineas.push(`Se guarda el detalle de los últimos <b>${d.meses} meses</b> (desde el ${escapeHtml(_fechaDMA(d.conservaDesde))}).`);
    }
    lineas.push(`Filas hoy: <b>${num(d.filasVentas)}</b> en VENTAS_LOCAL y <b>${num(d.filasDetalle)}</b> en DETALLE_VENTAS` + (d.filasResumen ? ` (${num(d.filasResumen)} son resúmenes)` : "") + ".");
    lineas.push(d.paraResumir
      ? `Para resumir ahora: <b>${num(d.paraResumir)} venta${d.paraResumir === 1 ? "" : "s"}</b> anteriores al ${escapeHtml(_fechaDMA(d.conservaDesde))}.`
      : "No hay ventas para resumir: todo está dentro del período que se guarda completo.");
    if (d.ultimo && d.ultimo.fecha) {
      const u = d.ultimo;
      const origen = u.origen === "panel" ? "desde el panel" : (u.origen === "editor" ? "desde el editor" : "automático");
      let t = `Último resumen: ${escapeHtml(new Date(u.fecha).toLocaleString("es-AR", { dateStyle: "short", timeStyle: "short" }))} (${origen})`;
      if (u.resumidas) t += ` — ${num(u.resumidas)} ventas: ${num(u.filasVentasAntes)} → ${num(u.filasVentasDespues)} filas de ventas, ${num(u.filasDetalleAntes)} → ${num(u.filasDetalleDespues)} de detalle`;
      else if (u.mensaje) t += ` — ${escapeHtml(u.mensaje)}`;
      if (u.respaldo && u.respaldo.carpeta) t += ` · <a href="${escapeHtml(u.respaldo.carpeta)}" target="_blank" rel="noopener">ver respaldo en Drive</a>`;
      lineas.push(t + ".");
    }
    caja.innerHTML = lineas.map(l => `<div>${l}</div>`).join("");
    if (btn) btn.disabled = d.desactivado || !d.paraResumir;
  } catch (e) {
    caja.textContent = "No se pudo leer el estado: " + (e.message || "error de conexión");
    if (btn) btn.disabled = true;
  }
}

function resumirVentasViejasClick() {
  const d = _estadoVentasResumidas;
  if (!d || !d.paraResumir) return;
  const msg = `Se van a resumir ${Number(d.paraResumir).toLocaleString("es-AR")} ventas anteriores al ${_fechaDMA(d.conservaDesde)}.\n\n` +
    "Los reportes siguen dando lo mismo; esas ventas dejan de verse ticket por ticket (y ya no se pueden anular ni reimprimir). " +
    "Antes, el detalle completo se guarda en Drive → \"Respaldo de ventas\".\n\nPuede tardar hasta un par de minutos. ¿Seguir?";
  confirmarAccion(msg, async () => {
    const btn = document.getElementById("btnResumirVentas");
    const caja = document.getElementById("ventasResumidasEstado");
    if (btn) { btn.disabled = true; btn.textContent = "Resumiendo…"; }
    if (caja) caja.insertAdjacentHTML("afterbegin", '<div class="text-primary">⏳ Resumiendo y guardando el respaldo… no cierres esta ventana.</div>');
    try {
      const res = await fetchAPI(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "resumirVentasViejas" })
      }, { timeoutMs: 330000 });
      const r = await res.json();
      if (!r.success) { toast(r.message || "No se pudo resumir", "error"); return; }
      toast(r.resumidas
        ? `✅ ${Number(r.resumidas).toLocaleString("es-AR")} ventas resumidas: ${r.filasVentasAntes} → ${r.filasVentasDespues} filas de ventas`
        : (r.mensaje || "No había ventas para resumir"), "success");
      try { invalidarCache("ventasPOS"); } catch (e) {}
    } catch (e) {
      toast("No se pudo resumir: " + (e && e.name === "AbortError" ? "tardó demasiado (puede que haya terminado igual; tocá Actualizar)" : "error de conexión"), "error");
    } finally {
      if (btn) btn.textContent = "Resumir ventas viejas ahora";
      cargarEstadoVentasResumidas();
    }
  }, "Resumir ventas viejas");
}

/* ===================== PEDIDOS ===================== */

/* =========================================================
   CACHÉ LOCALSTORAGE GENÉRICO — stale-while-revalidate
   Muestra datos cacheados al instante mientras actualiza
   en segundo plano. TTL configurable por sección.
========================================================= */
const CACHE_TTL_SECCIONES = {
  pedidos:   5 * 60 * 1000,  // 5 min
  clientes:  5 * 60 * 1000,  // 5 min
  ventasPOS: 3 * 60 * 1000,  // 3 min
};

function invalidarCache(...claves) {
  claves.forEach(c => {
    try { localStorage.removeItem("vpos_cache_" + c); } catch(e) {}
    delete ULTIMA_CARGA_SECCION[c];
  });
}

/* =====================================================
   COLA DE VENTAS PENDIENTES (sin conexión)
   Si guardarVenta falla por un corte de conexión, la venta ya se le
   mostró y (si el cajero eligió imprimir) ya se imprimió al cliente
   — no hay vuelta atrás con eso. Lo único que puede fallar es que la
   fila todavía no llegó a VENTAS_LOCAL. Esta cola guarda esa venta en
   el navegador (localStorage) y la reintenta sola en cuanto vuelve la
   conexión, usando el mismo clienteVentaId — que el backend reconoce
   para no duplicarla aunque el primer intento sí hubiera llegado a
   guardarse y solo se haya perdido la respuesta.
===================================================== */

const CLAVE_COLA_VENTAS = "vpos_cola_ventas_pendientes";

function _leerColaVentasPendientes() {
  try { return JSON.parse(localStorage.getItem(CLAVE_COLA_VENTAS) || "[]"); }
  catch(e) { return []; }
}

function _guardarColaVentasPendientes(cola) {
  try { localStorage.setItem(CLAVE_COLA_VENTAS, JSON.stringify(cola)); } catch(e) {}
  actualizarBadgeVentasPendientes(cola.length);
}

function encolarVentaPendiente(venta) {
  const cola = _leerColaVentasPendientes();
  cola.push(venta);
  _guardarColaVentasPendientes(cola);
}

/* ---- Botón de estado de sincronización (barra lateral) ----
   Reemplaza la banda azul de arriba y el aviso flotante de abajo a la
   derecha. Siempre visible, con color según el estado:
     verde  = todo sincronizado
     azul   = hay pendientes / sincronizando
     ámbar  = sin conexión (trabajando local)
     rojo   = hay operaciones con error que no subieron
   En la app de escritorio lo alimenta pos-offline.js (cola SQLite);
   en la web, la cola de ventas de localStorage de acá abajo. */
const _TEXTOS_ESTADO_SYNC = {
  ok: "Sincronizado",
  sincronizando: "Sincronizando…",
  pendiente: "Pendiente de subir",
  offline: "Sin conexión",
  error: "Error al sincronizar"
};

function pintarEstadoSync(info) {
  const estado = (info && info.estado) || "ok";
  const cantidad = Number((info && info.cantidad) || 0);
  const texto = (info && info.texto) || _TEXTOS_ESTADO_SYNC[estado] || "";
  const detalle = (info && info.detalle) || "";

  const btn = document.getElementById("syncEstadoBtn");
  if (btn) {
    btn.dataset.estado = estado;
    btn.title = detalle || texto;
    const t = document.getElementById("syncEstadoTexto");
    const c = document.getElementById("syncEstadoCantidad");
    if (t) t.textContent = texto;
    if (c) c.textContent = cantidad > 0 ? String(cantidad) : "";
  }
  const punto = document.getElementById("syncEstadoDotMovil");
  if (punto) {
    punto.dataset.estado = estado;
    punto.title = (detalle || texto) + (cantidad > 0 ? ` (${cantidad})` : "");
  }
}
window.pintarEstadoSync = pintarEstadoSync;

/** Click en el botón: en escritorio fuerza la sync de la cola SQLite; en la web reintenta la cola local */
function clickBotonEstadoSync() {
  if (typeof window.clickEstadoSyncEscritorio === "function") return window.clickEstadoSyncEscritorio();
  if (_leerColaVentasPendientes().length === 0) { toast("Todo está sincronizado ✓", "success"); return; }
  sincronizarVentasPendientes(true);
}

function _esAppEscritorio() {
  return typeof window.veekpos !== "undefined" || typeof window.posOffline !== "undefined";
}

/** Versión web del indicador (en escritorio lo maneja pos-offline.js) */
function actualizarBadgeVentasPendientes(cantidad) {
  if (_esAppEscritorio()) return;
  if (cantidad > 0) {
    pintarEstadoSync({
      estado: navigator.onLine ? "pendiente" : "offline",
      cantidad,
      detalle: `${cantidad} venta${cantidad === 1 ? "" : "s"} sin sincronizar — tocá para reintentar`
    });
  } else {
    pintarEstadoSync({ estado: navigator.onLine ? "ok" : "offline" });
  }
}

/**
 * Registra una venta ya cobrada. En la app de escritorio va SIEMPRE por
 * el camino local-primero de pos-offline.js (red multi-caja o SQLite +
 * cola, sin esperar a Apps Script). En la web, POST directo y, si falla,
 * a la cola de localStorage. Devuelve { ventaId } con el ID definitivo
 * si se conoce, o null.
 */
async function registrarVentaCobrada(venta) {
  if (typeof window.registrarVentaLocalPrimero === "function") {
    try {
      return await window.registrarVentaLocalPrimero(venta);
    } catch (error) {
      console.error("Falló el registro local de la venta, se usa la cola de respaldo:", error);
    }
  }

  try {
    const response = await fetchAPI(
      API_URL,
      {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "guardarVenta",
          total: venta.total,
          formaPago: venta.formaPago,
          vendedor: venta.vendedor || "",
          observaciones: venta.observaciones || "",
          carrito: venta.carrito,
          clienteVentaId: venta.clienteVentaId,
          montoUsd: venta.montoUsd || "",
          tipoCambio: venta.tipoCambio || ""
        })
      },
      { timeoutMs: 15000 } // mutación: sin reintento automático propio (lo maneja la cola)
    );
    const data = await response.json();
    if (data.success) return { ventaId: data.ventaId || null };
    toast("⚠️ La venta se mostró pero el servidor la rechazó: " + (data.message || "error desconocido"), "error");
    encolarVentaPendiente(venta);
    return null;
  } catch (err) {
    console.error("Error al guardar venta en backend:", err);
    // Se guarda en la cola local con el mismo clienteVentaId: el
    // backend lo reconoce y nunca la duplica al reintentar.
    encolarVentaPendiente(venta);
    toast("Sin conexión — la venta se guardó localmente y se subirá sola al reconectar", "error");
    return null;
  }
}

let _sincronizandoVentasPendientes = false;

/** Reintenta subir cada venta encolada, en orden, una por vez (no en paralelo — si el servidor está lento, mandar 10 a la vez lo empeora).
 *  manual=true cuando lo dispara el usuario (tocando el aviso) — ahí sí se avisa si sigue sin poder subir nada; en los reintentos
 *  automáticos de fondo (cada 30s, o al reconectar) no, para no repetir el mismo error cada rato mientras dure el corte. */
async function sincronizarVentasPendientes(manual) {
  if (_sincronizandoVentasPendientes) return;
  // En escritorio esta cola se migra a SQLite al arrancar (pos-offline.js)
  if (_esAppEscritorio() && typeof window.registrarVentaLocalPrimero === "function") return;
  const cola = _leerColaVentasPendientes();
  if (cola.length === 0) { actualizarBadgeVentasPendientes(0); return; }

  _sincronizandoVentasPendientes = true;
  if (!_esAppEscritorio()) pintarEstadoSync({ estado: "sincronizando", cantidad: cola.length });
  const pendientes = [];

  for (const venta of cola) {
    try {
      const response = await fetchAPI(
        API_URL,
        {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "guardarVenta",
            total: venta.total,
            formaPago: venta.formaPago,
            observaciones: venta.observaciones || "",
            vendedor: venta.vendedor || "",
            carrito: venta.carrito,
            clienteVentaId: venta.clienteVentaId,
            montoUsd: venta.montoUsd || "",
            tipoCambio: venta.tipoCambio || ""
          })
        },
        { timeoutMs: 15000 }
      );
      const data = await response.json();
      if (!data.success) pendientes.push(venta); // el servidor respondió pero rechazó la venta — se mantiene en cola para revisar
    } catch (error) {
      pendientes.push(venta); // seguía sin conexión — se mantiene en cola, se reintenta en el próximo ciclo
    }
  }

  _guardarColaVentasPendientes(pendientes);
  _sincronizandoVentasPendientes = false;

  const subieron = cola.length - pendientes.length;
  if (subieron > 0) {
    toast(`✅ ${subieron} venta${subieron === 1 ? "" : "s"} pendiente${subieron === 1 ? "" : "s"} sincronizada${subieron === 1 ? "" : "s"}`, "success");
    invalidarCache("ventasPOS");
    cargarMetricas();
  } else if (pendientes.length > 0 && manual) {
    // Sigue habiendo ventas sin subir — antes esto no avisaba nada,
    // así que tocar el botón "no hacía nada" visible aunque en
    // realidad sí lo había intentado y seguía fallando. Esto solo se
    // avisa cuando el usuario lo pidió a propósito (no en cada
    // reintento automático de fondo, para no repetir el mismo error
    // cada 30s mientras dure el corte).
    toast(`Todavía sin conexión — ${pendientes.length} venta${pendientes.length === 1 ? "" : "s"} pendiente${pendientes.length === 1 ? "" : "s"} de subir`, "error");
  }
}

// Reintentar apenas el navegador detecta que volvió la conexión, y
// además cada 30s como red de respaldo (el evento "online" del
// navegador no siempre es 100% confiable para saber si hay internet
// real, solo que hay una interfaz de red activa).
window.addEventListener("online", () => sincronizarVentasPendientes());
window.addEventListener("offline", () => actualizarBadgeVentasPendientes(_leerColaVentasPendientes().length));
setInterval(() => sincronizarVentasPendientes(), 30000);
document.addEventListener("DOMContentLoaded", () => actualizarBadgeVentasPendientes(_leerColaVentasPendientes().length));

function cacheGet(clave) {
  try {
    const raw = localStorage.getItem("vpos_cache_" + clave);
    if (!raw) return null;
    const { ts, data } = JSON.parse(raw);
    return { data, stale: Date.now() - ts > (CACHE_TTL_SECCIONES[clave] || 5 * 60 * 1000) };
  } catch(e) { return null; }
}

function cacheSet(clave, data) {
  try {
    localStorage.setItem("vpos_cache_" + clave, JSON.stringify({ ts: Date.now(), data }));
  } catch(e) {}
}

/**
 * Redibuja una lista sin que la pantalla salte al principio: mantiene la
 * altura del contenedor mientras se vacía y se vuelve a llenar (las
 * listas se dibujan en tandas) y devuelve el scroll a donde estaba.
 * Si la persona mueve la rueda o toca la pantalla, se deja de corregir.
 */
function redibujarSinSalto(idContenedor, dibujar) {
  const cont = document.getElementById(idContenedor);
  const y = window.scrollY;
  const seccion = cont ? cont.closest(".seccion") : null;
  const ySeccion = seccion ? seccion.scrollTop : 0;
  if (cont) cont.style.minHeight = cont.offsetHeight + "px";
  let cancelado = false;
  const cancelar = () => { cancelado = true; };
  const eventos = ["wheel", "touchstart", "keydown", "pointerdown"];
  eventos.forEach(ev => window.addEventListener(ev, cancelar, { once: true, passive: true, capture: true }));
  dibujar();
  const fin = Date.now() + 1500;
  const corregir = () => {
    if (!cancelado) {
      if (Math.abs(window.scrollY - y) > 2) window.scrollTo(0, y);
      if (seccion && Math.abs(seccion.scrollTop - ySeccion) > 2) seccion.scrollTop = ySeccion;
    }
    if (!cancelado && Date.now() < fin) { requestAnimationFrame(corregir); return; }
    if (cont) cont.style.minHeight = "";
    eventos.forEach(ev => window.removeEventListener(ev, cancelar, { capture: true }));
  };
  requestAnimationFrame(corregir);
}

let _cargandoPedidos = null; // evita que el polling y la apertura de la sección pidan la lista a la vez

function cargarPedidos() {
  if (_cargandoPedidos) return _cargandoPedidos;
  _cargandoPedidos = _cargarPedidosImpl().finally(() => { _cargandoPedidos = null; });
  return _cargandoPedidos;
}

async function _cargarPedidosImpl() {
  // Mostrar caché al instante si existe
  const cached = cacheGet("pedidos");
  if (cached) {
    // La copia guardada se dibuja solo si todavía no hay nada en pantalla.
    // Antes se redibujaba en cada actualización automática (cuando la
    // copia ya tenía más de 5 min) y eso hacía saltar la lista arriba.
    if (!pedidosGlobal || pedidosGlobal.length === 0) {
      pedidosGlobal = cached.data;
      filtrarPedidos();
    }
    if (!cached.stale) return; // fresco, no hace falta recargar
  }
  try {
    // Un solo intento con más margen: antes eran 2 intentos de 25s, y si
    // Apps Script estaba lento el segundo volvía a pedir la misma lectura
    // pesada mientras la primera todavía corría del lado del servidor.
    const response = await fetchAPI(API_URL + "?action=pedidos", {}, { timeoutMs: 45000, reintentos: 1 });
    const data = await response.json();
    if (!data.pedidos) return;

    // Si no cambió nada respecto de lo que ya se está mostrando, no
    // hace falta destruir y reconstruir toda la lista — eso es lo que
    // hacía "saltar" la pantalla al principio cada vez que el polling
    // de 15s se disparaba, aunque no hubiera pedidos nuevos.
    const firma = p => p.PEDIDO_ID + "|" + p.ESTADO + "|" + p.COBRADO + "|" + p.FORMA_PAGO_COBRO + "|" + p.TOTAL;
    const cambio = !pedidosGlobal
      || pedidosGlobal.length !== data.pedidos.length
      || pedidosGlobal.some((p, i) => firma(p) !== firma(data.pedidos[i]));

    pedidosGlobal = data.pedidos;
    cacheSet("pedidos", pedidosGlobal);
    _avisoPedidosCaido = false; // se pudo actualizar bien — resetea el aviso para el próximo corte
    if (cambio && _pedidosEnProceso.size === 0) redibujarSinSalto("tablaPedidos", filtrarPedidos); // no pisar el "Guardando…" de una acción en curso
  } catch (error) {
    if (error && error.name === "AbortError") {
      console.warn("Pedidos: el servidor tardó demasiado en responder (se reintenta en el próximo ciclo).");
    } else {
      console.error("Error pedidos:", error);
    }
    // Solo se avisa la primera vez que falla, no en cada intento del
    // polling de 15s — si la conexión está mal por un rato, un toast
    // nuevo cada 15s sería más molesto que informativo.
    if (!_avisoPedidosCaido) {
      toast("No se pudieron actualizar los pedidos — revisá tu conexión", "error");
      _avisoPedidosCaido = true;
    }
  }
}

/** Fuerza recarga de pedidos borrando el caché primero */
function recargarPedidos() {
  invalidarCache("pedidos");
  delete ULTIMA_CARGA_SECCION["pedidos"];
  cargarPedidos();
}

/** Botón "Eliminar cancelados" en la sección Pedidos — borra
 *  permanentemente todos los pedidos en estado CANCELADO junto con su
 *  detalle en DETALLE_PEDIDOS. Acción destructiva, por eso pide
 *  confirmación antes de mandar nada al backend. */
function eliminarPedidosCancelados() {
  if (obtenerRolActual() === "vendedor") {
    toast("Tu usuario (Vendedor) no tiene permiso para eliminar pedidos", "error");
    return;
  }
  confirmarAccion(
    "¿Eliminar todos los pedidos cancelados? Se van a borrar de forma permanente, junto con el detalle de sus productos. Esta acción no se puede deshacer.",
    _eliminarPedidosCanceladosConfirmado,
    "Eliminar pedidos cancelados"
  );
}

async function _eliminarPedidosCanceladosConfirmado() {
  try {
    const response = await fetchAPI(
      API_URL,
      {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "eliminarPedidosCancelados", rol: obtenerRolActual() })
      },
      { timeoutMs: 30000 }
    );
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudieron eliminar los pedidos cancelados", "error");
      return;
    }

    if (data.eliminados === 0) {
      toast("No había pedidos cancelados para eliminar", "info");
      return;
    }

    toast(`✓ Se eliminaron ${data.eliminados} pedido(s) cancelado(s)`, "success");
    recargarPedidos();

  } catch (error) {
    console.error("Error al eliminar pedidos cancelados:", error);
    toast("Error de conexión al eliminar los pedidos cancelados", "error");
  }
}

function recargarVentasPOSHistorial() {
  invalidarCache("ventasPOS");
  delete ULTIMA_CARGA_SECCION["ventasPOS"];
  cargarVentasPOSHistorial();
}

/* ---- Feedback visual de acciones sobre pedidos ----
   Antes, al cambiar el estado o marcar un pedido como cobrado no pasaba
   nada visible hasta que respondía Apps Script (a veces varios
   segundos), y no quedaba claro si el click se había registrado.
   Ahora: el control muestra "Guardando…" con spinner apenas se toca,
   la tarjeta queda atenuada, y al terminar se ve ✓ en verde (o el
   error en rojo, volviendo al valor anterior). */
function _tarjetaPedido(pedidoId) {
  return document.querySelector(`[data-pedido-card="${CSS.escape(String(pedidoId))}"]`);
}

function _marcarPedidoProcesando(pedidoId, el, texto) {
  const card = _tarjetaPedido(pedidoId);
  if (card) {
    card.classList.add("pedido-procesando");
    card.querySelectorAll(".pedido-pago-btn, .pedido-card-controls select, .pedido-cobrado-controles button")
      .forEach(b => { b.disabled = true; });
  }
  if (el) {
    el.dataset.textoOriginal = el.tagName === "SELECT" ? "" : el.innerHTML;
    el.classList.add("accion-pendiente");
    if (el.tagName !== "SELECT") el.innerHTML = `<span class="spinner-accion"></span>${texto}`;
    else {
      const aviso = document.createElement("span");
      aviso.className = "accion-aviso accion-aviso-pendiente";
      aviso.innerHTML = `<span class="spinner-accion"></span>${texto}`;
      el.insertAdjacentElement("afterend", aviso);
    }
  }
}

function _terminarPedidoProcesando(pedidoId, el, ok, textoResultado) {
  const card = _tarjetaPedido(pedidoId);
  if (card) {
    card.classList.remove("pedido-procesando");
    if (!ok) {
      card.querySelectorAll(".pedido-pago-btn, .pedido-card-controls select, .pedido-cobrado-controles button")
        .forEach(b => { b.disabled = false; });
    }
  }
  if (!el) return;
  el.classList.remove("accion-pendiente");
  if (el.tagName === "SELECT") {
    const aviso = el.parentElement && el.parentElement.querySelector(".accion-aviso");
    if (aviso) {
      aviso.className = "accion-aviso " + (ok ? "accion-aviso-ok" : "accion-aviso-error");
      aviso.textContent = textoResultado;
      if (!ok) setTimeout(() => aviso.remove(), 2500);
    }
    if (!ok) {
      el.value = el.dataset.estadoPrevio || el.value; // vuelve al estado que tenía
      el.classList.add("accion-error");
      setTimeout(() => el.classList.remove("accion-error"), 600);
    } else {
      el.classList.add("accion-ok");
    }
  } else {
    el.classList.add(ok ? "accion-ok" : "accion-error");
    el.innerHTML = textoResultado;
    if (!ok) {
      setTimeout(() => {
        el.classList.remove("accion-error");
        if (el.dataset.textoOriginal) el.innerHTML = el.dataset.textoOriginal;
      }, 1800);
    }
  }
}

/** Después de un cambio exitoso: deja ver el ✓ un momento, redibuja y resalta la tarjeta */
function _refrescarPedidosConDestello(pedidoId) {
  setTimeout(() => {
    filtrarPedidos();
    requestAnimationFrame(() => {
      const card = _tarjetaPedido(pedidoId);
      if (card) {
        card.classList.add("pedido-recien-actualizado");
        setTimeout(() => card.classList.remove("pedido-recien-actualizado"), 1600);
      }
    });
  }, 900);
}

/** Lee la respuesta de Apps Script avisando claro si no vino JSON (ej. página de error 404 de Google) */
async function _leerRespuestaJSON(response) {
  const texto = await response.text();
  try { return JSON.parse(texto); }
  catch (e) {
    throw new Error(response.ok
      ? "El servidor devolvió una respuesta inválida"
      : `El servidor respondió ${response.status} — revisá la implementación de Apps Script`);
  }
}

async function cambiarEstado(pedidoId, estado, el) {
  pedidoId = String(pedidoId); // los números nuevos (1001…) pueden llegar como número
  if (_pedidosEnProceso.has(pedidoId)) return;
  _pedidosEnProceso.add(pedidoId);
  _marcarPedidoProcesando(pedidoId, el, "Guardando…");
  let ok = false, mensaje = "";
  try {
    const response = await fetchAPI(
      API_URL +
      "?action=actualizarEstado" +
      "&pedidoId=" + encodeURIComponent(pedidoId) +
      "&estado="   + encodeURIComponent(estado),
      {},
      { timeoutMs: 30000 }
    );
    const data = await _leerRespuestaJSON(response);
    if (!data.success) {
      mensaje = data.message || "No se pudo actualizar el pedido";
    } else {
      ok = true;
      const p = _pedidoLocalPorId(pedidoId);
      if (p) { p.ESTADO = estado; invalidarCache("pedidos"); }
    }
  } catch (error) {
    console.error("Error al cambiar el estado del pedido:", error);
    mensaje = error && error.name === "AbortError" ? "El servidor no respondió a tiempo" : (error.message || "Error de conexión");
  } finally {
    _pedidosEnProceso.delete(pedidoId);
  }

  if (ok) {
    _terminarPedidoProcesando(pedidoId, el, true, "✓ Guardado");
    toast(`Pedido ${pedidoId}: estado cambiado a ${estado}`, "success");
    _refrescarPedidosConDestello(pedidoId);
  } else {
    _terminarPedidoProcesando(pedidoId, el, false, "✕ No se guardó");
    toast(`No se pudo cambiar el estado: ${mensaje}`, "error");
  }
}

/**
 * Maneja el checkbox "Cobrado" de un pedido. Solo dispara la llamada
 * al backend cuando YA hay una forma de pago elegida — si todavía no
 * se eligió, simplemente habilita el desplegable y espera a que el
 * usuario la complete (eso lo dispara cambiarFormaPagoPedido). Esto
 * evita mandar "cobrado=SI" sin forma de pago, que el backend
 * rechazaría de todas formas.
 */
async function cambiarCobradoPedido(pedidoId, marcado) {
  const selectFormaPago = document.getElementById(`formaPago-${pedidoId}`);

  if (!marcado) {
    selectFormaPago.disabled = true;
    selectFormaPago.value = "";
    await aplicarCobroPedido(pedidoId, false, "");
    return;
  }

  selectFormaPago.disabled = false;

  const formaPagoElegida = selectFormaPago.value;
  if (!formaPagoElegida) {
    toast("Elegí la forma de pago para terminar de marcarlo como cobrado", "error");
    return; // se queda tildado y con el desplegable habilitado, esperando la forma de pago
  }

  await aplicarCobroPedido(pedidoId, true, formaPagoElegida);
}

/** Maneja el desplegable de forma de pago — si el checkbox ya está tildado, aplica el cambio al elegir */
async function cambiarFormaPagoPedido(pedidoId, formaPago) {
  const checkbox = document.getElementById(`cobrado-${pedidoId}`);
  if (!checkbox.checked) return; // todavía no se tildó "Cobrado", no hace nada hasta que se tilde
  if (!formaPago) return;
  await aplicarCobroPedido(pedidoId, true, formaPago);
}

/** Llama al backend para marcar/desmarcar el pedido como cobrado en caja, y refresca la lista */
const _pedidosEnProceso = new Set(); // evita doble cobro del mismo pedido

async function aplicarCobroPedido(pedidoId, cobrado, formaPago, el) {
  pedidoId = String(pedidoId); // los números nuevos (1001…) pueden llegar como número
  if (_pedidosEnProceso.has(pedidoId)) return; // ya procesando
  _pedidosEnProceso.add(pedidoId);
  _marcarPedidoProcesando(pedidoId, el, cobrado ? "Cobrando…" : "Desmarcando…");

  let ok = false, mensaje = "";
  try {
    // Con timeout (antes era un fetch sin límite: si Apps Script se
    // colgaba, el botón quedaba bloqueado para siempre)
    const response = await fetchAPI(
      API_URL +
      "?action=marcarPedidoCobrado" +
      "&pedidoId=" + encodeURIComponent(pedidoId) +
      "&cobrado=" + (cobrado ? "SI" : "NO") +
      "&formaPago=" + encodeURIComponent(formaPago || ""),
      {},
      { timeoutMs: 30000 }
    );
    const data = await _leerRespuestaJSON(response);

    if (!data.success) {
      mensaje = data.message || "No se pudo actualizar el cobro del pedido";
    } else {
      ok = true;
      const p = _pedidoLocalPorId(pedidoId);
      if (p) {
        p.COBRADO = cobrado ? "SI" : "NO";
        p.FORMA_PAGO_COBRO = formaPago || "";
        invalidarCache("pedidos");
      }
    }
  } catch (error) {
    console.error("Error al marcar el pedido como cobrado:", error);
    mensaje = error && error.name === "AbortError" ? "El servidor no respondió a tiempo" : (error.message || "Error de conexión");
  } finally {
    _pedidosEnProceso.delete(pedidoId);
  }

  if (ok) {
    _terminarPedidoProcesando(pedidoId, el, true, cobrado ? "✓ Cobrado" : "✓ Desmarcado");
    toast(cobrado
      ? `Pedido ${pedidoId} cobrado con ${formaPago} — ya suma al cierre de caja`
      : `Pedido ${pedidoId} desmarcado — ya no suma al cierre de caja`, "success");
    _refrescarPedidosConDestello(pedidoId);
    invalidarCache("ventasPOS");
  } else {
    _terminarPedidoProcesando(pedidoId, el, false, "✕ Error");
    toast(`No se pudo ${cobrado ? "marcar como cobrado" : "desmarcar"} el pedido: ${mensaje}`, "error");
  }
}

/* Pedidos viejos encontrados en el servidor para la búsqueda actual
   (la lista general solo trae los en curso y los resueltos de los
   últimos 30 días). Se muestran junto con los demás, marcados. */
let _pedidosAntiguos = { clave: "", lista: [] };
let _timerBusquedaPedidos = null;
let _genBusquedaPedidos = 0;

function _claveBusquedaPedidos() {
  const q = (document.getElementById("buscarPedido")?.value || "").trim().toLowerCase();
  const desde = document.getElementById("pedidosDesde")?.value || "";
  const hasta = document.getElementById("pedidosHasta")?.value || "";
  return { q, desde, hasta, clave: q + "|" + desde + "|" + hasta, activa: q.length >= 3 || !!desde || !!hasta };
}

/** Pedido por número, en la lista general o entre los viejos encontrados. */
function _pedidoLocalPorId(pedidoId) {
  const id = String(pedidoId);
  return pedidosGlobal.find(x => String(x.PEDIDO_ID) === id) || (_pedidosAntiguos.lista || []).find(x => String(x.PEDIDO_ID) === id);
}

function programarBusquedaPedidosAntiguos(inmediata) {
  clearTimeout(_timerBusquedaPedidos);
  const b = _claveBusquedaPedidos();
  const btnFechas = document.getElementById("pedidosLimpiarFechas");
  if (btnFechas) btnFechas.style.display = (b.desde || b.hasta) ? "" : "none";
  const aviso = document.getElementById("pedidosBusquedaAntiguos");
  if (!b.activa) {
    _pedidosAntiguos = { clave: "", lista: [] };
    if (aviso) aviso.innerHTML = (b.q.length > 0 ? '<span class="text-muted">Escribí al menos 3 letras o números para buscar también en pedidos anteriores.</span>' : "");
    return;
  }
  if (b.clave === _pedidosAntiguos.clave) return;
  _timerBusquedaPedidos = setTimeout(() => buscarPedidosAntiguos(), inmediata ? 0 : 650);
}

async function buscarPedidosAntiguos() {
  const b = _claveBusquedaPedidos();
  if (!b.activa) return;
  const aviso = document.getElementById("pedidosBusquedaAntiguos");
  const miGen = ++_genBusquedaPedidos;
  if (aviso) aviso.innerHTML = '<span class="text-muted"><span class="giro-chico">↻</span> Buscando también en pedidos anteriores…</span>';
  try {
    const qs = new URLSearchParams({ action: "buscarPedidos", q: b.q, desde: b.desde, hasta: b.hasta });
    const res = await fetchAPI(API_URL + "?" + qs.toString(), {}, { timeoutMs: 45000, reintentos: 1 });
    const data = await res.json();
    if (miGen !== _genBusquedaPedidos) return; // ya se escribió otra cosa
    if (!data.success) { if (aviso) aviso.innerHTML = `<span class="text-danger">${escapeHtml(data.message || "No se pudo buscar")}</span>`; return; }
    const enLista = new Set(pedidosGlobal.map(p => String(p.PEDIDO_ID)));
    const extra = (data.pedidos || []).filter(p => !enLista.has(String(p.PEDIDO_ID))).map(p => Object.assign(p, { _antiguo: true }));
    _pedidosAntiguos = { clave: b.clave, lista: extra };
    if (aviso) {
      aviso.innerHTML = extra.length
        ? `🗄️ <b>${extra.length}</b> pedido${extra.length === 1 ? "" : "s"} anterior${extra.length === 1 ? "" : "es"} encontrado${extra.length === 1 ? "" : "s"}` + (data.limitado ? ` (se muestran los 100 más nuevos de ${data.total}; afiná la búsqueda)` : "") + "."
        : '<span class="text-muted">No hay pedidos anteriores con esa búsqueda.</span>';
    }
    filtrarPedidos();
  } catch (e) {
    if (miGen !== _genBusquedaPedidos) return;
    if (aviso) aviso.innerHTML = '<span class="text-danger">No se pudo buscar en pedidos anteriores (sin conexión). Probá de nuevo.</span>';
  }
}

function limpiarBusquedaPedidos() {
  const d = document.getElementById("pedidosDesde"), h = document.getElementById("pedidosHasta");
  if (d) d.value = ""; if (h) h.value = "";
  filtrarPedidos();
  programarBusquedaPedidosAntiguos(true);
}

function filtrarPedidos() {
  const texto = document.getElementById("buscarPedido").value.toLowerCase().trim();
  const textoNum = texto.replace(/\D/g, "");
  const estadoFiltro = (document.getElementById("filtroPedidoEstado")?.value || "").toUpperCase();
  const b = _claveBusquedaPedidos();
  const desde = b.desde ? new Date(b.desde + "T00:00:00") : null;
  const hasta = b.hasta ? new Date(b.hasta + "T23:59:59") : null;
  // Los viejos encontrados en el servidor solo para la búsqueda que los trajo
  const antiguos = (_pedidosAntiguos.clave === b.clave) ? _pedidosAntiguos.lista : [];
  const filtrados = pedidosGlobal.concat(antiguos).filter(p => {
    const coincideTexto = !texto ||
      String(p.PEDIDO_ID || "").toLowerCase().includes(texto) ||
      String(p.CLIENTE   || "").toLowerCase().includes(texto) ||
      String(p.DNI       || "").toLowerCase().includes(texto) ||
      String(p.EMPRESA   || "").toLowerCase().includes(texto) ||
      (textoNum.length >= 4 && String(p.TELEFONO || "").replace(/\D/g, "").includes(textoNum)) ||
      !!p._antiguo; // el servidor ya filtró (sin acentos, DNI con guiones, etc.)
    const coincideEstado = !estadoFiltro || String(p.ESTADO || "").toUpperCase() === estadoFiltro;
    let coincideFecha = true;
    if (desde || hasta) {
      const f = new Date(p.FECHA);
      coincideFecha = !isNaN(f) && (!desde || f >= desde) && (!hasta || f <= hasta);
    }
    return coincideTexto && coincideEstado && coincideFecha;
  });
  renderPedidos(filtrados);
}

const ESTADO_ICONO_PEDIDO = { NUEVO: "🆕", PREPARANDO: "⚙️", ENVIADO: "📦", CANCELADO: "❌" };

let pedidoDetalleActual = null;

async function abrirDetallePedido(pedidoId) {
  pedidoId = String(pedidoId);
  document.getElementById("pedidoDetalleBody").innerHTML = `<div class="text-center text-muted py-4">Cargando...</div>`;
  document.getElementById("pedidoDetalleModalBackdrop").classList.add("show");

  try {
    const response = await fetchAPI(API_URL + "?action=getDetallePedido&pedidoId=" + encodeURIComponent(pedidoId));
    const data = await response.json();
    if (!data.success) { toast(data.message || "No se pudo cargar el pedido", "error"); cerrarModalDetallePedido(); return; }

    const { pedido, detalle } = data;
    pedidoDetalleActual = { pedido, detalle };

    // Botón imprimir: si hay PDF en Drive lo abre, si no genera el A4
    const btnImprimir = document.querySelector("#pedidoDetalleModalBackdrop .btn-success");
    if (btnImprimir) {
      if (pedido.PDF_URL) {
        btnImprimir.textContent = "Ver PDF del pedido";
        btnImprimir.onclick = () => window.open(pedido.PDF_URL, "_blank");
      } else {
        btnImprimir.textContent = "Imprimir pedido (A4)";
        btnImprimir.onclick = imprimirNotaPedidoA4;
      }
    }

    // Mostrar/ocultar botón etiqueta según estado
    const btnEtiqueta = document.getElementById("btnImprimirEtiquetaDetalle");
    if (btnEtiqueta) btnEtiqueta.style.display = pedido.ESTADO === "PREPARANDO" ? "inline-flex" : "none";

    // Editar ítems: solo mientras el pedido sigue en NUEVO — una vez
    // que pasa a Preparando ya se descontó stock, y editar ahí
    // complicaría tener que devolver/redescontar con más margen de error.
    const btnEditarItems = document.getElementById("btnEditarItemsPedido");
    if (btnEditarItems) btnEditarItems.style.display = pedido.ESTADO === "NUEVO" ? "inline-flex" : "none";
    const btnEditarCliente = document.getElementById("btnEditarClientePedido");
    if (btnEditarCliente) btnEditarCliente.style.display = "inline-flex";

    const simbolo = String(pedido.MONEDA || "ARS").toUpperCase() === "USD" ? "US$" : "$";
    const filas = detalle.map(item => `
      <tr>
        <td style="padding:6px 0;">${item.cantidad}x ${escapeHtml(item.PRODUCTO)}</td>
        <td style="padding:6px 0; text-align:right; font-family:var(--font-mono);">${simbolo}${(item.PRECIO * item.cantidad).toLocaleString("es-AR")}</td>
      </tr>`).join("");

    document.getElementById("pedidoDetalleBody").innerHTML = `
      <div class="config-preview-hint mb-3">
        <span class="ic">${ESTADO_ICONO_PEDIDO[pedido.ESTADO] || "📦"}</span>
        Pedido <strong>${escapeHtml(pedido.PEDIDO_ID)}</strong> — Estado: <strong>${escapeHtml(pedido.ESTADO)}</strong>
      </div>
      <div class="mb-2"><strong style="font-size:15px;">${escapeHtml(pedido.CLIENTE)}</strong>${pedido.DNI ? ` · DNI/CUIT: ${escapeHtml(pedido.DNI)}` : ""}</div>
      ${pedido.TELEFONO ? `<div class="mb-1" style="font-size:13px;color:var(--slate-500);">📞 ${escapeHtml(pedido.TELEFONO)}</div>` : ""}
      ${pedido.DIRECCION ? `<div class="mb-1" style="font-size:13px;color:var(--slate-500);">📍 ${escapeHtml(pedido.DIRECCION)}${pedido.LOCALIDAD ? ", " + escapeHtml(pedido.LOCALIDAD) : ""}${pedido.PROVINCIA ? ", " + escapeHtml(pedido.PROVINCIA) : ""}${pedido.CODIGO_POSTAL ? " (CP " + escapeHtml(pedido.CODIGO_POSTAL) + ")" : ""}</div>` : ""}
      ${pedido.EMPRESA ? `<div class="mb-3" style="font-size:13px;color:var(--slate-500);">🚚 ${escapeHtml(pedido.EMPRESA)}</div>` : `<div class="mb-3"></div>`}
      <table style="width:100%; border-collapse:collapse; font-size:13.5px;">
        <tbody>${filas}</tbody>
        <tfoot>
          ${Number(pedido.DESCUENTO || 0) > 0 ? `
          <tr style="border-top:1px solid var(--slate-200);">
            <td style="padding-top:8px; color:var(--slate-500);">Subtotal</td>
            <td style="padding-top:8px; text-align:right; font-family:var(--font-mono); color:var(--slate-500);">${simbolo}${Number(pedido.SUBTOTAL || pedido.TOTAL || 0).toLocaleString("es-AR")}</td>
          </tr>
          <tr>
            <td style="padding-top:4px; color:var(--red-500); font-weight:700;">Descuento${pedido.DESCUENTO_ETIQUETA ? " (" + pedido.DESCUENTO_ETIQUETA + ")" : ""}</td>
            <td style="padding-top:4px; text-align:right; font-family:var(--font-mono); color:var(--red-500); font-weight:700;">-${simbolo}${Number(pedido.DESCUENTO).toLocaleString("es-AR")}</td>
          </tr>` : ""}
          <tr style="border-top:2px solid var(--slate-200);">
            <td style="padding-top:10px;"><strong>Total</strong></td>
            <td style="padding-top:10px; text-align:right; font-family:var(--font-mono);"><strong>${simbolo}${Number(pedido.TOTAL || 0).toLocaleString("es-AR")}</strong></td>
          </tr>
        </tfoot>
      </table>`;
    if (window.ArcaUI) ArcaUI.enDetallePedido(pedido, detalle); // factura del pedido (facturacion-arca.js)

  } catch (error) {
    console.error("Error al abrir detalle de pedido:", error);
    toast("Error al cargar el detalle del pedido", "error");
    cerrarModalDetallePedido();
  }
}

function cerrarModalDetallePedido() {
  document.getElementById("pedidoDetalleModalBackdrop").classList.remove("show");
  pedidoDetalleActual = null;
  _carritoEdicionPedido = null;
  _descuentoEdicionPedido = null;
}

/* ===================== EDITAR DATOS DEL CLIENTE DE UN PEDIDO =====================
   Cualquier estado. El backend (editarDatosClientePedido, en
   PapeleraYClientes.gs) actualiza el pedido, refleja los cambios en la
   hoja CLIENTES (o crea el cliente si no existía) y regenera el PDF. */

function activarEdicionClientePedido() {
  if (!pedidoDetalleActual) return;
  const p = pedidoDetalleActual.pedido;
  const campo = (id, label, valor, col, extra) => `
    <div class="${col}">
      <label class="form-label" style="font-size:12px;">${label}</label>
      <input type="text" class="form-control form-control-sm" id="${id}" value="${escapeHtml(valor || "")}" ${extra || ""}>
    </div>`;

  ["btnEditarItemsPedido", "btnEditarClientePedido"].forEach(id => {
    const b = document.getElementById(id);
    if (b) b.style.display = "none";
  });

  document.getElementById("pedidoDetalleBody").innerHTML = `
    <div class="config-preview-hint mb-3">
      Editando los datos del cliente del pedido <strong>${escapeHtml(p.PEDIDO_ID)}</strong>.
      Los cambios también se guardan en <strong>Clientes</strong>.
    </div>
    <div class="row g-2 mb-2">
      ${campo("edCliNombre", 'Nombre / razón social <span class="text-danger">*</span>', p.CLIENTE, "col-12")}
    </div>
    <div class="row g-2 mb-2">
      ${campo("edCliDni", "DNI/CUIT", p.DNI, "col-6", 'inputmode="numeric"')}
      ${campo("edCliTelefono", "Teléfono", p.TELEFONO, "col-6")}
    </div>
    <div class="row g-2 mb-2">
      ${campo("edCliDireccion", "Dirección", p.DIRECCION, "col-12")}
    </div>
    <div class="row g-2 mb-2">
      ${campo("edCliLocalidad", "Localidad", p.LOCALIDAD, "col-7")}
      ${campo("edCliCP", "Código Postal", p.CODIGO_POSTAL || p.CODIGOPOSTAL, "col-5")}
    </div>
    <div class="row g-2 mb-2">
      ${campo("edCliProvincia", "Provincia", p.PROVINCIA, "col-6")}
      ${campo("edCliEmpresa", "Transporte / empresa", p.EMPRESA, "col-6")}
    </div>
    <div class="d-flex gap-2 justify-content-end mt-3">
      <button type="button" class="btn btn-outline-secondary btn-sm" onclick="cancelarEdicionClientePedido()">Cancelar</button>
      <button type="button" class="btn btn-primary btn-sm" id="btnGuardarClientePedido" onclick="guardarEdicionClientePedido()">Guardar cambios</button>
    </div>`;

  setTimeout(() => document.getElementById("edCliNombre")?.focus(), 50);
}

function cancelarEdicionClientePedido() {
  if (!pedidoDetalleActual) return;
  abrirDetallePedido(pedidoDetalleActual.pedido.PEDIDO_ID);
}

async function guardarEdicionClientePedido() {
  if (!pedidoDetalleActual) return;
  const val = id => (document.getElementById(id)?.value || "").trim();
  const nombre = val("edCliNombre");
  if (!nombre) {
    toast("El nombre del cliente es obligatorio", "error");
    document.getElementById("edCliNombre")?.focus();
    return;
  }

  const pedidoId = pedidoDetalleActual.pedido.PEDIDO_ID;
  const cuerpo = {
    action: "editarDatosClientePedido",
    pedidoId,
    nombre,
    dni: val("edCliDni"),
    telefono: val("edCliTelefono"),
    direccion: val("edCliDireccion"),
    localidad: val("edCliLocalidad"),
    codigoPostal: val("edCliCP"),
    provincia: val("edCliProvincia"),
    empresa: val("edCliEmpresa")
  };

  const btn = document.getElementById("btnGuardarClientePedido");
  if (btn) { btn.disabled = true; btn.textContent = "Guardando..."; }

  try {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(cuerpo)
    });
    const data = await _leerRespuestaJSON(response);

    if (!data || !data.success) {
      let msg = (data && data.message) || "No se pudieron guardar los cambios";
      if (/no v[aá]lida|desconocida|no reconocida|invalid/i.test(msg)) {
        msg = "El backend todavía no tiene esta función: agregá PapeleraYClientes.gs en Apps Script y publicá una nueva versión (ver LEEME).";
      }
      toast(msg, "error");
      if (btn) { btn.disabled = false; btn.textContent = "Guardar cambios"; }
      return;
    }

    const accion = data.cliente && data.cliente.accion === "creado" ? "y se creó el cliente en Clientes" : "y en Clientes";
    toast(`Datos del cliente actualizados en el pedido ${accion}`, "success");
    if (data.pdfError) toast("Ojo: no se pudo regenerar el PDF del pedido", "error");

    // Que Clientes, el selector de clientes y la lista de pedidos vean los datos nuevos
    _clientesPedidoCache = [];
    invalidarCache("clientes", "pedidos");
    try { if (typeof cargarClientesDesdeBackend === "function") cargarClientesDesdeBackend(); } catch (e) {}
    try { recargarPedidos(); } catch (e) {}

    abrirDetallePedido(pedidoId);

  } catch (error) {
    console.error("Error al editar cliente del pedido:", error);
    toast("Error de conexión — revisá tu internet e intentá de nuevo", "error");
    if (btn) { btn.disabled = false; btn.textContent = "Guardar cambios"; }
  }
}

/* ===================== EDITAR ÍTEMS DE UN PEDIDO (solo estado NUEVO) ===================== */

let _carritoEdicionPedido = null; // copia de trabajo mientras se edita — no toca pedidoDetalleActual hasta guardar
let _descuentoEdicionPedido = null; // copia de trabajo del descuento — { monto, etiqueta }

function activarEdicionItemsPedido() {
  if (!pedidoDetalleActual) return;
  // Copia de trabajo — si cancela, pedidoDetalleActual queda intacto
  _carritoEdicionPedido = pedidoDetalleActual.detalle.map(item => ({ ...item }));
  _descuentoEdicionPedido = {
    monto: Number(pedidoDetalleActual.pedido.DESCUENTO || 0),
    etiqueta: pedidoDetalleActual.pedido.DESCUENTO_ETIQUETA || ""
  };
  renderEdicionItemsPedido();
}

function cancelarEdicionItemsPedido() {
  _carritoEdicionPedido = null;
  _descuentoEdicionPedido = null;
  abrirDetallePedido(pedidoDetalleActual.pedido.PEDIDO_ID); // vuelve a la vista normal, con los datos tal cual estaban
}

function renderEdicionItemsPedido() {
  const pedido = pedidoDetalleActual.pedido;
  const simbolo = String(pedido.MONEDA || "ARS").toUpperCase() === "USD" ? "US$" : "$";

  const filas = _carritoEdicionPedido.map((item, idx) => `
    <tr>
      <td style="padding:5px 4px;">
        <input type="number" min="1" step="1" value="${item.cantidad}" class="form-control form-control-sm" style="width:64px;"
          onchange="_edicionPedidoCambiarCampo(${idx}, 'cantidad', this.value)">
      </td>
      <td style="padding:5px 4px; font-size:13px;">${escapeHtml(item.PRODUCTO)}</td>
      <td style="padding:5px 4px;">
        <input type="number" min="0" step="1" value="${item.PRECIO}" class="form-control form-control-sm" style="width:90px;"
          onchange="_edicionPedidoCambiarCampo(${idx}, 'PRECIO', this.value)">
      </td>
      <td style="padding:5px 4px; text-align:right; font-family:var(--font-mono); font-size:13px;">${simbolo}${((Number(item.PRECIO) || 0) * (Number(item.cantidad) || 0)).toLocaleString("es-AR")}</td>
      <td style="padding:5px 4px;">
        <button type="button" class="btn btn-outline-danger btn-sm" title="Quitar" onclick="_edicionPedidoQuitarItem(${idx})">🗑️</button>
      </td>
    </tr>`).join("");

  const subtotal = _carritoEdicionPedido.reduce((acc, i) => acc + (Number(i.PRECIO) || 0) * (Number(i.cantidad) || 0), 0);
  const descuento = Math.min(subtotal, Math.max(0, Number(_descuentoEdicionPedido.monto) || 0));
  const total = Math.max(0, subtotal - descuento);

  document.getElementById("pedidoDetalleBody").innerHTML = `
    <div class="config-preview-hint mb-3">
      Editando ítems del pedido <strong>${escapeHtml(pedido.PEDIDO_ID)}</strong> — solo se puede mientras está en NUEVO.
    </div>
    <table style="width:100%; border-collapse:collapse; font-size:13px;">
      <thead>
        <tr style="color:var(--slate-500); font-size:11.5px;">
          <td style="padding:0 4px 4px;">Cant.</td>
          <td style="padding:0 4px 4px;">Producto</td>
          <td style="padding:0 4px 4px;">Precio</td>
          <td style="padding:0 4px 4px; text-align:right;">Subtotal</td>
          <td></td>
        </tr>
      </thead>
      <tbody>${filas}</tbody>
    </table>

    <div class="mt-3 mb-2" style="position:relative;">
      <div class="d-flex gap-2 align-items-center">
        <input type="text" id="edicionPedidoBuscarProducto" class="form-control form-control-sm" placeholder="Código o nombre del producto a agregar..."
          autocomplete="off" oninput="_edicionPedidoActualizarResultados()" onkeydown="if(event.key==='Enter'){event.preventDefault(); _edicionPedidoAgregarProducto();}">
        <button type="button" class="btn btn-outline-primary btn-sm" style="white-space:nowrap;" onclick="_edicionPedidoAgregarProducto()">+ Agregar</button>
      </div>
      <div id="edicionPedidoResultados" class="edicion-pedido-resultados" style="display:none;"></div>
    </div>

    <table style="width:100%; border-collapse:collapse; font-size:13.5px; margin-top:10px;">
      <tfoot>
        <tr style="border-top:1px solid var(--slate-200);">
          <td colspan="4" style="padding-top:8px; color:var(--slate-500);">Subtotal</td>
          <td style="padding-top:8px; text-align:right; font-family:var(--font-mono); color:var(--slate-500);">${simbolo}${subtotal.toLocaleString("es-AR")}</td>
        </tr>
      </tfoot>
    </table>

    <div class="d-flex gap-2 align-items-center mt-2">
      <label style="font-size:12.5px; color:var(--red-500); font-weight:700; white-space:nowrap;">Descuento (${simbolo})</label>
      <input type="number" min="0" step="1" value="${_descuentoEdicionPedido.monto || 0}" class="form-control form-control-sm" style="width:100px;"
        onchange="_edicionPedidoCambiarDescuento('monto', this.value)">
      <input type="text" placeholder="Motivo / etiqueta (opcional)" value="${escapeHtml(_descuentoEdicionPedido.etiqueta || "")}" class="form-control form-control-sm" style="flex:1;"
        onchange="_edicionPedidoCambiarDescuento('etiqueta', this.value)">
    </div>

    <table style="width:100%; border-collapse:collapse; font-size:13.5px; margin-top:6px;">
      <tfoot>
        ${descuento > 0 ? `
        <tr>
          <td colspan="4" style="padding-top:4px; color:var(--red-500); font-weight:700;">Descuento${_descuentoEdicionPedido.etiqueta ? " (" + escapeHtml(_descuentoEdicionPedido.etiqueta) + ")" : ""}</td>
          <td style="padding-top:4px; text-align:right; font-family:var(--font-mono); color:var(--red-500); font-weight:700;">-${simbolo}${descuento.toLocaleString("es-AR")}</td>
        </tr>` : ""}
        <tr style="border-top:2px solid var(--slate-200);">
          <td colspan="4" style="padding-top:10px;"><strong>Nuevo total</strong></td>
          <td style="padding-top:10px; text-align:right; font-family:var(--font-mono);"><strong>${simbolo}${total.toLocaleString("es-AR")}</strong></td>
        </tr>
      </tfoot>
    </table>

    <div class="d-flex justify-content-end gap-2 mt-3">
      <button type="button" class="btn btn-outline-secondary btn-sm" onclick="cancelarEdicionItemsPedido()">Cancelar</button>
      <button type="button" class="btn btn-success btn-sm" id="btnGuardarEdicionItemsPedido" onclick="guardarEdicionItemsPedido()">Guardar cambios</button>
    </div>`;
}

function _edicionPedidoCambiarDescuento(campo, valor) {
  if (!_descuentoEdicionPedido) return;
  if (campo === "monto") {
    const n = Number(valor);
    _descuentoEdicionPedido.monto = (isNaN(n) || n < 0) ? 0 : n;
  } else if (campo === "etiqueta") {
    _descuentoEdicionPedido.etiqueta = String(valor || "").trim();
  }
  renderEdicionItemsPedido();
}

function _edicionPedidoCambiarCampo(idx, campo, valor) {
  const n = Number(valor);
  if (isNaN(n) || n < 0 || (campo === "cantidad" && n < 1)) { renderEdicionItemsPedido(); return; }
  _carritoEdicionPedido[idx][campo] = n;
  renderEdicionItemsPedido();
}

function _edicionPedidoQuitarItem(idx) {
  if (_carritoEdicionPedido.length <= 1) {
    toast("El pedido no puede quedar sin productos — editalo o cancelalo en vez de vaciarlo", "error");
    return;
  }
  _carritoEdicionPedido.splice(idx, 1);
  renderEdicionItemsPedido();
}

let _productoElegidoEdicionPedido = null; // el producto que quedó seleccionado al hacer clic en un resultado de la lista

function _edicionPedidoActualizarResultados() {
  const input = document.getElementById("edicionPedidoBuscarProducto");
  const texto = input.value.trim().toLowerCase();
  const cont = document.getElementById("edicionPedidoResultados");

  // Si el texto cambió después de haber elegido un producto de la
  // lista, esa selección ya no es válida (el cajero puede estar
  // buscando otra cosa) — se limpia hasta que vuelva a elegir.
  _productoElegidoEdicionPedido = null;

  if (!texto || texto.length < 2) { cont.style.display = "none"; cont.innerHTML = ""; return; }

  const fuente = (productosAdminGlobal && productosAdminGlobal.length ? productosAdminGlobal : productosPOS) || [];
  const coincidencias = fuente
    .filter(p => String(p.CODIGO).toLowerCase().includes(texto) || String(p.PRODUCTO).toLowerCase().includes(texto))
    .slice(0, 15);

  if (coincidencias.length === 0) {
    cont.innerHTML = `<div class="edicion-pedido-resultado-vacio">Sin resultados</div>`;
    cont.style.display = "block";
    return;
  }

  cont.innerHTML = coincidencias.map((p, idx) => `
    <div class="edicion-pedido-resultado-item" onclick="_edicionPedidoElegirResultado(${idx})">
      <span class="epr-codigo">${escapeHtml(p.CODIGO)}</span>
      <span class="epr-nombre">${escapeHtml(p.PRODUCTO)}</span>
      <span class="epr-precio">$${Number(p.PRECIO || 0).toLocaleString("es-AR")}</span>
    </div>`).join("");

  // Se guardan las coincidencias actuales para que el clic (que solo
  // manda el índice) sepa a qué producto corresponde cada fila.
  _resultadosEdicionPedidoActuales = coincidencias;

  cont.style.display = "block";
}

let _resultadosEdicionPedidoActuales = [];

function _edicionPedidoElegirResultado(idx) {
  const producto = _resultadosEdicionPedidoActuales[idx];
  if (!producto) return;

  _productoElegidoEdicionPedido = producto;
  document.getElementById("edicionPedidoBuscarProducto").value = `${producto.CODIGO} — ${producto.PRODUCTO}`;
  document.getElementById("edicionPedidoResultados").style.display = "none";

  _edicionPedidoAgregarProducto();
}

// Cierra la lista de resultados si se toca en cualquier otro lado del modal
document.addEventListener("click", (e) => {
  const cont = document.getElementById("edicionPedidoResultados");
  const input = document.getElementById("edicionPedidoBuscarProducto");
  if (!cont || cont.style.display === "none") return;
  if (e.target === input || cont.contains(e.target)) return;
  cont.style.display = "none";
});

function _edicionPedidoAgregarProducto() {
  const input = document.getElementById("edicionPedidoBuscarProducto");
  const texto = input.value.trim();
  if (!texto) return;

  // Si se hizo clic en un resultado de la lista, ya sabemos exactamente
  // qué producto es (sin depender de volver a parsear el texto escrito).
  let producto = _productoElegidoEdicionPedido;

  if (!producto) {
    const codigoEscrito = texto.split(" — ")[0].trim().toLowerCase();
    const fuente = (productosAdminGlobal && productosAdminGlobal.length ? productosAdminGlobal : productosPOS) || [];
    producto = fuente.find(p => String(p.CODIGO).toLowerCase() === codigoEscrito)
      || fuente.find(p => String(p.PRODUCTO).toLowerCase() === texto.toLowerCase());
  }

  if (!producto) { toast("No se encontró ese producto — elegilo de la lista de resultados", "error"); return; }

  const existente = _carritoEdicionPedido.find(i => String(i.CODIGO) === String(producto.CODIGO));
  if (existente) {
    existente.cantidad++;
  } else {
    _carritoEdicionPedido.push({
      CODIGO: producto.CODIGO,
      PRODUCTO: producto.PRODUCTO,
      cantidad: 1,
      PRECIO: Number(producto.PRECIO) || 0
    });
  }

  input.value = "";
  _productoElegidoEdicionPedido = null;
  renderEdicionItemsPedido();
}

async function guardarEdicionItemsPedido() {
  if (!_carritoEdicionPedido || _carritoEdicionPedido.length === 0) return;

  const pedidoId = pedidoDetalleActual.pedido.PEDIDO_ID;
  const btn = document.getElementById("btnGuardarEdicionItemsPedido");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    // Por POST, con el carrito en el body — mismo motivo que en
    // ventas/pedidos grandes: por GET, con todo en la URL, un pedido
    // con muchos ítems puede superar lo que Apps Script acepta.
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "editarItemsPedido",
        pedidoId,
        carrito: _carritoEdicionPedido,
        descuento: Number(_descuentoEdicionPedido && _descuentoEdicionPedido.monto) || 0,
        descuentoEtiqueta: (_descuentoEdicionPedido && _descuentoEdicionPedido.etiqueta) || ""
      })
    }, { timeoutMs: 45000 });
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudieron guardar los cambios", "error");
      return;
    }

    if (data.pdfError) {
      toast(data.message, "error");
    } else {
      toast("Ítems del pedido actualizados", "success");
    }

    _carritoEdicionPedido = null;
    _descuentoEdicionPedido = null;
    invalidarCache("pedidos");
    cargarPedidos(); // refresca el total en la lista de atrás también
    abrirDetallePedido(pedidoId); // vuelve a la vista normal con los datos ya guardados

  } catch (error) {
    console.error("Error al editar ítems del pedido:", error);
    // OJO: este "error de conexión" no significa necesariamente que el
    // cambio no se guardó. El backend (Apps Script) puede haber
    // terminado de escribir en la hoja después de que el navegador ya
    // se rindió esperando la respuesta (regenerar el PDF del pedido
    // puede tardar bastante). Por eso, en vez de asumir que falló,
    // recargamos el detalle del pedido para mostrar el estado real
    // en vez de dejar al usuario con una duda y un carrito de edición
    // "colgado" que podría reintentar y duplicar ítems.
    toast("Se cortó la conexión esperando la respuesta — revisando si el cambio se guardó igual...", "error");
    try {
      _carritoEdicionPedido = null;
      _descuentoEdicionPedido = null;
      invalidarCache("pedidos");
      await abrirDetallePedido(pedidoId);
    } catch (e2) {
      console.error("Error al revisar el estado real del pedido tras el timeout:", e2);
    }
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

/**
 * Etiqueta de envío MANUAL — misma plantilla A4 que imprimirEtiquetaEnvio
 * (la que ya se usa desde el detalle de un pedido en estado PREPARANDO),
 * pero sin depender de que exista un pedido real: para envíos que no
 * salieron del catálogo web (ej. se tomó el pedido por teléfono o
 * WhatsApp). No crea ningún registro en Pedidos ni toca stock — es
 * pura utilidad de impresión, todo queda en el navegador.
 */
let _etqManualClienteSeleccionado = null; // cliente elegido de la lista (null = destinatario nuevo/escrito a mano)
let _etqManualClientes = [];

function abrirModalEtiquetaManual() {
  ["etqManualCliente", "etqManualDni", "etqManualTelefono", "etqManualDireccion",
   "etqManualLocalidad", "etqManualCP", "etqManualProvincia", "etqManualTransporte",
   "etqManualReferencia", "etqManualBuscarCliente"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  _etqManualClienteSeleccionado = null;
  const chk = document.getElementById("etqManualGuardarCliente");
  if (chk) chk.checked = true;
  const res = document.getElementById("etqManualResultadosClientes");
  if (res) { res.style.display = "none"; res.innerHTML = ""; }
  etqManualActualizarGuardar();

  document.getElementById("etiquetaManualModalBackdrop").classList.add("show");
  setTimeout(() => (document.getElementById("etqManualBuscarCliente") || document.getElementById("etqManualCliente"))?.focus(), 50);
  etqManualCargarClientes();
}

function cerrarModalEtiquetaManual() {
  document.getElementById("etiquetaManualModalBackdrop").classList.remove("show");
}

/** Trae los clientes (con localidad/provincia/CP) — usa el mismo cache que el selector de pedidos */
async function etqManualCargarClientes() {
  let lista = [];
  try { lista = await cargarClientesParaPedido(); } catch (e) { lista = []; }
  // Sin internet en la app de escritorio: se usa la copia local de clientes
  if ((!lista || lista.length === 0) && window.posOffline && typeof window.posOffline.obtenerClientesConCredito === "function") {
    try { lista = await window.posOffline.obtenerClientesConCredito(); } catch (e) { lista = []; }
  }
  if ((!lista || lista.length === 0) && Array.isArray(clientesGlobal)) lista = clientesGlobal;
  _etqManualClientes = lista || [];
  const buscador = document.getElementById("etqManualBuscarCliente");
  if (buscador && buscador.value.trim()) etqManualFiltrarClientes();
}

function etqManualFiltrarClientes() {
  const input = document.getElementById("etqManualBuscarCliente");
  const res = document.getElementById("etqManualResultadosClientes");
  if (!input || !res) return;
  const q = normalizarBusquedaPOS(input.value);
  if (!q) { res.style.display = "none"; res.innerHTML = ""; return; }

  const qDigitos = q.replace(/\D/g, "");
  const coincidencias = _etqManualClientes.filter(c => {
    const texto = normalizarBusquedaPOS([c.NOMBRE, c.ALIAS, c.DNI, c.TELEFONO, c.LOCALIDAD].join(" "));
    if (texto.includes(q)) return true;
    return qDigitos.length >= 3 && String(c.DNI || "").replace(/\D/g, "").includes(qDigitos);
  }).slice(0, 12);

  if (coincidencias.length === 0) {
    res.innerHTML = `<div style="padding:8px 10px; font-size:12.5px; color:var(--slate-500);">
      No hay clientes con "${escapeHtml(input.value.trim())}".
      <a href="#" onclick="event.preventDefault(); etqManualUsarComoNuevo();">Cargarlo como cliente nuevo</a>
    </div>`;
  } else {
    res.innerHTML = coincidencias.map(c => `
      <div class="etq-cli-opcion" style="padding:7px 10px; cursor:pointer; border-bottom:1px solid var(--slate-100);"
           onmousedown="event.preventDefault(); etqManualSeleccionarCliente('${escapeJsAttr(c.CLIENTE_ID)}')"
           onmouseover="this.style.background='var(--slate-100)'" onmouseout="this.style.background=''">
        <div style="font-weight:600; font-size:13px;">${escapeHtml(c.NOMBRE)}${c.ALIAS ? ` <span style="color:var(--slate-500); font-weight:400;">(${escapeHtml(c.ALIAS)})</span>` : ""}</div>
        <div style="font-size:11.5px; color:var(--slate-500);">${[c.DNI ? "DNI " + escapeHtml(c.DNI) : "", escapeHtml(c.LOCALIDAD || ""), escapeHtml(c.TELEFONO || "")].filter(Boolean).join(" · ")}</div>
      </div>`).join("");
  }
  res.style.display = "block";
}

function etqManualTeclaBuscar(event) {
  if (event.key === "Escape") {
    const res = document.getElementById("etqManualResultadosClientes");
    if (res) res.style.display = "none";
  } else if (event.key === "Enter") {
    event.preventDefault();
    const primera = document.querySelector("#etqManualResultadosClientes .etq-cli-opcion");
    if (primera) primera.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
  }
}

function etqManualOcultarResultados() {
  setTimeout(() => {
    const res = document.getElementById("etqManualResultadosClientes");
    if (res) res.style.display = "none";
  }, 150);
}

function etqManualSeleccionarCliente(clienteId) {
  const c = _etqManualClientes.find(x => String(x.CLIENTE_ID) === String(clienteId));
  if (!c) return;
  _etqManualClienteSeleccionado = c;
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v || ""; };
  set("etqManualCliente", c.NOMBRE);
  set("etqManualDni", c.DNI);
  set("etqManualTelefono", c.TELEFONO);
  set("etqManualDireccion", c.DIRECCION);
  set("etqManualLocalidad", c.LOCALIDAD);
  set("etqManualCP", c.CODIGO_POSTAL || c.CODIGOPOSTAL);
  set("etqManualProvincia", c.PROVINCIA);
  set("etqManualTransporte", c.EMPRESA);
  set("etqManualBuscarCliente", c.NOMBRE);
  const res = document.getElementById("etqManualResultadosClientes");
  if (res) res.style.display = "none";
  etqManualActualizarGuardar();
}

/** Pasa lo escrito en el buscador al nombre del destinatario, como cliente nuevo */
function etqManualUsarComoNuevo() {
  const texto = (document.getElementById("etqManualBuscarCliente")?.value || "").trim();
  etqManualLimpiarCliente();
  const nombre = document.getElementById("etqManualCliente");
  if (nombre) { if (texto && !/^\d+$/.test(texto)) nombre.value = texto; nombre.focus(); }
  const dni = document.getElementById("etqManualDni");
  if (dni && /^\d{7,11}$/.test(texto)) dni.value = texto;
}

/** "✕ Nuevo": deja el formulario vacío para cargar un destinatario que no está en la lista */
function etqManualLimpiarCliente() {
  _etqManualClienteSeleccionado = null;
  ["etqManualCliente", "etqManualDni", "etqManualTelefono", "etqManualDireccion",
   "etqManualLocalidad", "etqManualCP", "etqManualProvincia", "etqManualTransporte",
   "etqManualBuscarCliente"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  const res = document.getElementById("etqManualResultadosClientes");
  if (res) res.style.display = "none";
  etqManualActualizarGuardar();
  document.getElementById("etqManualCliente")?.focus();
}

/**
 * Muestra el tilde "Guardar como cliente nuevo" solo cuando el
 * destinatario NO es un cliente existente. Si se elige uno de la lista
 * pero después se cambia el nombre o el DNI a mano, se lo toma como
 * otro destinatario (nuevo).
 */
function etqManualActualizarGuardar() {
  const c = _etqManualClienteSeleccionado;
  if (c) {
    const nombre = (document.getElementById("etqManualCliente")?.value || "").trim().toLowerCase();
    const dni = (document.getElementById("etqManualDni")?.value || "").replace(/\D/g, "");
    const mismoNombre = nombre === String(c.NOMBRE || "").trim().toLowerCase();
    const mismoDni = dni === String(c.DNI || "").replace(/\D/g, "");
    if (!mismoNombre || !mismoDni) _etqManualClienteSeleccionado = null;
  }

  const fila = document.getElementById("etqManualGuardarClienteFila");
  const aviso = document.getElementById("etqManualClienteSeleccionadoAviso");
  if (fila) fila.style.display = _etqManualClienteSeleccionado ? "none" : "block";
  if (aviso) {
    aviso.style.display = _etqManualClienteSeleccionado ? "block" : "none";
    if (_etqManualClienteSeleccionado) aviso.textContent = "✓ Cliente de la lista: " + _etqManualClienteSeleccionado.NOMBRE;
  }
}

/** Busca si el destinatario ya existe en la lista local (por DNI o nombre exacto) */
function _etqManualBuscarExistente(nombre, dni) {
  const dniNum = String(dni || "").replace(/\D/g, "");
  if (dniNum) return _etqManualClientes.find(c => String(c.DNI || "").replace(/\D/g, "") === dniNum) || null;
  const n = String(nombre || "").trim().toLowerCase();
  return _etqManualClientes.find(c => String(c.NOMBRE || "").trim().toLowerCase() === n) || null;
}

/**
 * Guarda el destinatario como cliente nuevo. Usa guardarClienteDesdeRotulo
 * (guarda también localidad/provincia/CP y no duplica por DNI); si el
 * backend todavía no tiene ese archivo, cae a crearCliente; y en la app
 * de escritorio sin internet, lo guarda local y lo encola para Sheets.
 */
async function etqManualGuardarCliente(datos) {
  const existente = _etqManualBuscarExistente(datos.nombre, datos.dni);
  if (existente) return { ok: true, existente: true, nombre: existente.NOMBRE };

  const post = async (cuerpo) => {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(cuerpo)
    });
    return _leerRespuestaJSON(response);
  };

  try {
    let data = await post({ action: "guardarClienteDesdeRotulo", ...datos });
    if (!data || (!data.success && /no v[aá]lida|desconocida|no reconocida|invalid/i.test(String(data.message || "")))) {
      data = await post({ action: "crearCliente", nombre: datos.nombre, alias: "", dni: datos.dni, telefono: datos.telefono,
                          empresa: datos.empresa, direccion: datos.direccion, aCredito: "NO" });
    }
    if (!data || !data.success) return { ok: false, message: (data && data.message) || "No se pudo guardar el cliente" };

    _clientesPedidoCache = [];
    invalidarCache("clientes");
    try { if (typeof cargarClientesDesdeBackend === "function") cargarClientesDesdeBackend(); } catch (e) {}
    return { ok: true, existente: !!data.existente, nombre: data.nombre || datos.nombre };

  } catch (error) {
    // App de escritorio sin conexión: guardar local + encolar para Sheets
    if (window.posOffline && typeof window.posOffline.crearCliente === "function" && typeof window.posOffline.encolarAccion === "function") {
      try {
        const local = { nombre: datos.nombre, alias: "", dni: datos.dni, telefono: datos.telefono,
                        direccion: datos.direccion, empresa: datos.empresa, aCredito: "NO" };
        await window.posOffline.crearCliente(local);
        await window.posOffline.encolarAccion("cliente_upsert", { editando: false, ...local });
        return { ok: true, offline: true, nombre: datos.nombre };
      } catch (e2) {
        console.error("No se pudo guardar el cliente offline:", e2);
      }
    }
    console.error("Error al guardar cliente desde el rótulo:", error);
    return { ok: false, message: "Error de conexión al guardar el cliente" };
  }
}

async function generarEtiquetaManual() {
  const val = id => (document.getElementById(id)?.value || "").trim();

  const cliente = val("etqManualCliente");
  if (!cliente) {
    toast("Ingresá el nombre del destinatario", "error");
    document.getElementById("etqManualCliente")?.focus();
    return;
  }

  // Si no se cargó una referencia a mano, se genera una localmente
  // (timestamp) solo para tener algo que mostrar en el código de
  // barras — no es un ID de pedido real.
  const referencia = val("etqManualReferencia") || ("MANUAL-" + Date.now());

  const datos = {
    pedidoId: referencia,
    cliente,
    telefono: val("etqManualTelefono"),
    direccion: val("etqManualDireccion"),
    localidad: val("etqManualLocalidad"),
    provincia: val("etqManualProvincia"),
    codigoPostal: val("etqManualCP"),
    dni: val("etqManualDni"),
    transporte: val("etqManualTransporte")
  };

  etqManualActualizarGuardar();
  const quiereGuardar = !_etqManualClienteSeleccionado && !!document.getElementById("etqManualGuardarCliente")?.checked;

  // Primero se imprime (window.open tiene que salir del click directo,
  // si no el navegador lo puede bloquear); el guardado va después.
  imprimirEtiquetaEnvio(datos);
  cerrarModalEtiquetaManual();

  if (quiereGuardar) {
    const r = await etqManualGuardarCliente({
      nombre: cliente, dni: datos.dni, telefono: datos.telefono, direccion: datos.direccion,
      localidad: datos.localidad, provincia: datos.provincia, codigoPostal: datos.codigoPostal,
      empresa: datos.transporte
    });
    if (!r.ok) toast("El rótulo se generó, pero no se pudo guardar el cliente: " + r.message, "error");
    else if (r.existente) toast(`"${r.nombre}" ya estaba en Clientes — no se duplicó`, "success");
    else if (r.offline) toast("Cliente guardado offline — se sincroniza con Sheets cuando vuelva la conexión", "success");
    else toast(`Cliente "${r.nombre}" guardado en Clientes`, "success");
  }
}

function imprimirEtiquetaDesdeDetalle() {
  if (!pedidoDetalleActual) return;
  const p = pedidoDetalleActual.pedido;
  imprimirEtiquetaEnvio({
    pedidoId: p.PEDIDO_ID, cliente: p.CLIENTE, telefono: p.TELEFONO || "",
    direccion: p.DIRECCION || "", localidad: p.LOCALIDAD || "", provincia: p.PROVINCIA || "",
    codigoPostal: p.CODIGO_POSTAL || p.CODIGOPOSTAL || "", dni: p.DNI || "", transporte: p.EMPRESA || ""
  });
}

function imprimirNotaPedidoA4() {
  if (!pedidoDetalleActual) return;
  const { pedido, detalle } = pedidoDetalleActual;
  const simbolo = String(pedido.MONEDA || "ARS").toUpperCase() === "USD" ? "US$" : "$";
  const ahora = new Date().toLocaleString("es-AR");
  const fechaPedido = pedido.FECHA ? new Date(pedido.FECHA).toLocaleDateString("es-AR") : "—";

  const filasItems = detalle.map(item => `
    <tr>
      <td style="padding:6px; border:1px solid #ddd; font-family:monospace; font-size:9pt;">${escapeHtml(item.CODIGO || "")}</td>
      <td style="padding:6px; border:1px solid #ddd;">${escapeHtml(item.PRODUCTO)}</td>
      <td style="padding:6px; border:1px solid #ddd; text-align:center;">${item.cantidad}</td>
      <td style="padding:6px; border:1px solid #ddd; text-align:right;">${simbolo}${Number(item.PRECIO).toLocaleString("es-AR")}</td>
      <td style="padding:6px; border:1px solid #ddd; text-align:right;">${simbolo}${(item.PRECIO * item.cantidad).toLocaleString("es-AR")}</td>
    </tr>`).join("");

  const filaDoc = (label, valor) => valor
    ? `<tr><td style="padding:4px 8px; font-size:7pt; font-weight:700; text-transform:uppercase; color:#555; width:32mm; vertical-align:top;">${label}:</td><td style="padding:4px 8px; font-size:11pt; font-weight:700;">${escapeHtml(valor)}</td></tr>`
    : "";

  const cobradoStr = String(pedido.COBRADO || "").toUpperCase() === "SI"
    ? `<div style="margin-top:4mm; font-size:11pt; color:#16a34a; font-weight:700;">Cobrado${pedido.FORMA_PAGO_COBRO ? " — " + escapeHtml(pedido.FORMA_PAGO_COBRO) : ""}</div>`
    : `<div style="margin-top:4mm; font-size:11pt; color:#d32f2f; font-weight:700;">Pendiente de cobro</div>`;

  const html = `
    <div style="font-family:Arial,Helvetica,sans-serif; width:190mm; margin:10mm auto; padding:8mm; box-sizing:border-box; border:2px solid #000; border-radius:4mm;">
      <div style="text-align:center; margin-bottom:6mm;">
        <div style="font-size:20pt; font-weight:900;">${escapeHtml(obtenerConfigNegocio().nombre || "")}</div>
      </div>
      <div style="background:#0b1633; color:#fff; text-align:center; padding:4mm; border-radius:2mm; margin-bottom:6mm;">
        <div style="font-size:14pt; font-weight:900; letter-spacing:2px;">NOTA DE PEDIDO</div>
        <div style="font-size:11pt; font-weight:700;">${escapeHtml(pedido.PEDIDO_ID)} — ${fechaPedido}</div>
      </div>
      <table style="width:100%; border-collapse:collapse; font-size:11pt; margin-bottom:6mm;">
        ${filaDoc("Cliente", pedido.CLIENTE)}
        ${filaDoc("Teléfono", pedido.TELEFONO)}
        ${filaDoc("Dirección", [pedido.DIRECCION, pedido.LOCALIDAD, pedido.PROVINCIA].filter(Boolean).join(", "))}
        ${filaDoc("CP", pedido.CODIGO_POSTAL || pedido.CODIGOPOSTAL)}
        ${filaDoc("DNI/CUIT", pedido.DNI)}
        ${filaDoc("Transporte", pedido.EMPRESA)}
        ${filaDoc("Estado", pedido.ESTADO)}
      </table>
      <table style="width:100%; border-collapse:collapse; font-size:11pt;">
        <thead>
          <tr style="background:#f4f4f4;">
            <th style="padding:6px; text-align:left; border:1px solid #ddd;">Código</th>
            <th style="padding:6px; text-align:left; border:1px solid #ddd;">Producto</th>
            <th style="padding:6px; text-align:center; border:1px solid #ddd;">Cant.</th>
            <th style="padding:6px; text-align:right; border:1px solid #ddd;">Precio</th>
            <th style="padding:6px; text-align:right; border:1px solid #ddd;">Subtotal</th>
          </tr>
        </thead>
        <tbody>${filasItems}</tbody>
        <tfoot>
          ${Number(pedido.DESCUENTO || 0) > 0 ? `
          <tr>
            <td colspan="4" style="padding:6px; text-align:right; border:1px solid #ddd; color:#555;">Subtotal</td>
            <td style="padding:6px; text-align:right; border:1px solid #ddd; color:#555;">${simbolo}${Number(pedido.SUBTOTAL || pedido.TOTAL || 0).toLocaleString("es-AR")}</td>
          </tr>
          <tr>
            <td colspan="4" style="padding:6px; text-align:right; border:1px solid #ddd; color:#d32f2f; font-weight:700;">Descuento${pedido.DESCUENTO_ETIQUETA ? " (" + pedido.DESCUENTO_ETIQUETA + ")" : ""}</td>
            <td style="padding:6px; text-align:right; border:1px solid #ddd; color:#d32f2f; font-weight:700;">-${simbolo}${Number(pedido.DESCUENTO).toLocaleString("es-AR")}</td>
          </tr>` : ""}
          <tr>
            <td colspan="4" style="padding:8px; text-align:right; border:1px solid #ddd;"><strong>TOTAL</strong></td>
            <td style="padding:8px; text-align:right; border:1px solid #ddd;"><strong>${simbolo}${Number(pedido.TOTAL || 0).toLocaleString("es-AR")}</strong></td>
          </tr>
        </tfoot>
      </table>
      ${cobradoStr}
      <div style="text-align:center; font-size:8pt; color:#aaa; margin-top:4mm;">Impreso el ${ahora}</div>
    </div>`;

  // Por si quedó contenido de una impresión de ticket térmico anterior
  // a medio camino, se limpia el otro contenedor de impresión para que
  // el CSS de :empty/:not(:empty) elija el correcto sin ambigüedad.
  const thermalFrameNota = document.getElementById("thermalPrintFrame");
  if (thermalFrameNota) thermalFrameNota.innerHTML = "";

  const area = document.getElementById("etiquetasPrintArea");
  area.innerHTML = html;
  _setPrintPageSize("A4");
  setTimeout(() => window.print(), 120);
}

/* ====================================================================
 * REMITO sobre formulario preimpreso (168x204mm, cod. 091)
 * Imprime SOLO los datos variables en la posición exacta del formulario.
 * Coordenadas en mm medidas desde la esquina superior izquierda de la hoja
 * A4 (la hoja del remito se carga como si fuera A4, pegada a la esquina).
 * y = línea base del texto. Ajuste fino: REMITO_OFFSET (mm) guardado por navegador.
 * ==================================================================== */
const REMITO_POS = {
  dia:        { x: 110.9, y: 28.5 },
  mes:        { x: 129.1, y: 28.5 },
  anio:       { x: 144.1, y: 28.5 },
  cliente:    { x: 27.4,  y: 48.0 },
  domicilio:  { x: 27.4,  y: 53.5 },
  localidad:  { x: 28.0,  y: 58.8 },
  telefono:   { x: 119.3, y: 58.8 },
  provincia:  { x: 27.4,  y: 63.9 },
  cuit:       { x: 102.4, y: 72.0 },
  transporte: { x: 33.3,  y: 175.6 },
  ivaCF:      { x: 43.9,  y: 69.2 },   // centro del casillero
  ivaRI:      { x: 75.0,  y: 68.7 },
  ivaEX:      { x: 43.9,  y: 73.0 },
  ivaMT:      { x: 75.0,  y: 72.9 },
  filasY0: 92.0, filasDy: 5.62, filasN: 13,
  cantX: 20.5, descX: 31.9, descW: 123
};
const REMITO_OFFSET_KEY = "jireh_remito_offset";

function _remitoOffset() {
  try {
    const o = JSON.parse(localStorage.getItem(REMITO_OFFSET_KEY) || "{}");
    return { x: Number(o.x) || 0, y: Number(o.y) || 0 };
  } catch (e) { return { x: 0, y: 0 }; }
}

function _remitoTxt(pos, texto, extra) {
  if (texto === undefined || texto === null || String(texto) === "") return "";
  const o = _remitoOffset();
  return `<div style="position:absolute; left:${(pos.x + o.x).toFixed(2)}mm; top:${(pos.y + o.y - 3.4).toFixed(2)}mm; height:4mm; line-height:4mm; white-space:nowrap; font-size:10pt; ${extra || ""}">${escapeHtml(String(texto))}</div>`;
}

function _remitoMarca(pos) {
  const o = _remitoOffset();
  return `<div style="position:absolute; left:${(pos.x + o.x - 2).toFixed(2)}mm; top:${(pos.y + o.y - 2).toFixed(2)}mm; width:4mm; height:4mm; line-height:4mm; text-align:center; font-size:11pt; font-weight:700;">X</div>`;
}

/** Arma el HTML de la hoja (A4, margen 0) a partir de los datos del formulario. */
function _remitoHTML(d) {
  const P = REMITO_POS, o = _remitoOffset();
  let h = `<div style="position:relative; width:210mm; height:297mm; font-family:Arial,Helvetica,sans-serif; color:#000; overflow:hidden;">`;
  h += _remitoTxt(P.dia, d.dia) + _remitoTxt(P.mes, d.mes) + _remitoTxt(P.anio, d.anio);
  h += _remitoTxt(P.cliente, d.cliente) + _remitoTxt(P.domicilio, d.domicilio);
  h += _remitoTxt(P.localidad, d.localidad) + _remitoTxt(P.telefono, d.telefono);
  h += _remitoTxt(P.provincia, d.provincia) + _remitoTxt(P.cuit, d.cuit);
  h += _remitoTxt(P.transporte, d.transporte);
  const iva = { CF: P.ivaCF, RI: P.ivaRI, EX: P.ivaEX, MT: P.ivaMT }[d.iva];
  if (iva) h += _remitoMarca(iva);
  if (d.bultos) {
    h += `<div style="position:absolute; left:${(P.cantX - 9 + o.x).toFixed(2)}mm; top:${(P.filasY0 + 5 * P.filasDy + o.y - 13).toFixed(2)}mm; width:18mm; text-align:center; font-size:46pt; font-weight:700; line-height:1;">${escapeHtml(String(d.bultos))}</div>`;
  }
  (d.filas || []).slice(0, P.filasN).forEach((f, i) => {
    const y = P.filasY0 + i * P.filasDy;
    if (f.cant !== "" && f.cant != null) {
      h += _remitoTxt({ x: P.cantX - 8, y }, f.cant, "width:16mm; text-align:center;");
    }
    h += _remitoTxt({ x: P.descX, y }, f.desc, `max-width:${P.descW}mm; overflow:hidden;`);
  });
  if (d.calibracion) {
    // Cruces en cada punto de anclaje para superponer contra el formulario
    Object.keys(P).forEach(k => {
      const p = P[k]; if (!p || p.x === undefined) return;
      h += `<div style="position:absolute; left:${(p.x + o.x - 2).toFixed(2)}mm; top:${(p.y + o.y - 0.15).toFixed(2)}mm; width:4mm; height:0.3mm; background:#e00;"></div>`;
      h += `<div style="position:absolute; left:${(p.x + o.x - 0.15).toFixed(2)}mm; top:${(p.y + o.y - 2).toFixed(2)}mm; width:0.3mm; height:4mm; background:#e00;"></div>`;
    });
  }
  return h + `</div>`;
}

function _remitoImprimir(html) {
  const thermal = document.getElementById("thermalPrintFrame");
  if (thermal) thermal.innerHTML = "";
  document.getElementById("etiquetasPrintArea").innerHTML = html;
  _setPrintPageSize("REMITO");
  setTimeout(() => window.print(), 120);
}

let _remitoPedido = null;

function abrirRemitoDesdeDetalle() {
  if (!pedidoDetalleActual) return;
  const { pedido, detalle } = pedidoDetalleActual;
  _remitoPedido = { pedido, detalle };
  const hoy = new Date();
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v ?? ""; };

  _remitoAsegurarModal();
  set("rmDia", String(hoy.getDate()).padStart(2, "0"));
  set("rmMes", String(hoy.getMonth() + 1).padStart(2, "0"));
  set("rmAnio", String(hoy.getFullYear()).slice(-2));
  set("rmCliente", pedido.CLIENTE);
  set("rmDomicilio", pedido.DIRECCION);
  set("rmLocalidad", pedido.LOCALIDAD);
  set("rmProvincia", pedido.PROVINCIA);
  set("rmTelefono", pedido.TELEFONO);
  set("rmCuit", pedido.DNI);
  set("rmTransporte", pedido.EMPRESA);
  set("rmIva", "");
  set("rmValor", "");
  set("rmBultos", 1);
  set("rmDespacho", "");
  const off = _remitoOffset();
  set("rmOffX", off.x); set("rmOffY", off.y);
  remitoCargarResumen();
  document.getElementById("remitoModalBackdrop").classList.add("show");
}

function cerrarModalRemito() {
  document.getElementById("remitoModalBackdrop").classList.remove("show");
}

function _remitoRenderFilas(filas) {
  const cont = document.getElementById("rmFilas");
  let html = "";
  for (let i = 0; i < REMITO_POS.filasN; i++) {
    const f = filas[i] || { cant: "", desc: "" };
    html += `<div style="display:flex; gap:6px; margin-bottom:4px;">
      <input class="form-control rm-cant" style="width:64px; text-align:center;" value="${escapeHtml(String(f.cant ?? ""))}" placeholder="Cant.">
      <input class="form-control rm-desc" style="flex:1;" value="${escapeHtml(String(f.desc ?? ""))}" placeholder="Descripción (línea ${i + 1})">
    </div>`;
  }
  cont.innerHTML = html;
}

function _remitoNumLetras(n) {
  n = parseInt(n, 10);
  const u = ["cero","un","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez","once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte","veintiún","veintidós","veintitrés","veinticuatro","veinticinco","veintiséis","veintisiete","veintiocho","veintinueve"];
  const dec = ["","","","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"];
  if (isNaN(n) || n < 1) return "";
  if (n < 30) return u[n];
  if (n < 100) { const r = n % 10; return dec[Math.floor(n / 10)] + (r ? " y " + (r === 1 ? "un" : u[r]) : ""); }
  return String(n);
}

function remitoCargarResumen() {
  const b = parseInt((document.getElementById("rmBultos") || {}).value, 10) || 1;
  const ico = ((document.getElementById("rmDespacho") || {}).value || "").replace(/\D/g, "");
  const letras = _remitoNumLetras(b);
  const L = letras.charAt(0).toUpperCase() + letras.slice(1);
  const filas = [
    { cant: "", desc: `${L} ${b === 1 ? "bulto" : "bultos"} total de mercadería` },
    { cant: "", desc: "Origen nacional y extranjero" },
    { cant: "", desc: "Rubro: juguetería, regalería y bazar" },
    { cant: "", desc: "Aduana Bs. As." },
    { cant: "", desc: "N° Despacho" }
  ];
  if (ico) filas.push({ cant: "", desc: "ICO: " + (ico.length > 4 ? ico.slice(0, 4) + " " + ico.slice(4) : ico) });
  _remitoRenderFilas(filas);
}

function remitoCargarProductos() {
  if (!_remitoPedido) return;
  const items = _remitoPedido.detalle || [];
  const filas = items.map(it => ({ cant: it.cantidad, desc: it.PRODUCTO }));
  if (filas.length > REMITO_POS.filasN) {
    toast(`El pedido tiene ${filas.length} ítems y el formulario solo ${REMITO_POS.filasN} líneas. Se cargaron las primeras ${REMITO_POS.filasN}; usá "Resumen" o imprimí en 2 hojas.`, "error");
  }
  document.getElementById("rmBultos").value = "";
  _remitoRenderFilas(filas);
}

function _remitoLeerForm() {
  const v = id => (document.getElementById(id).value || "").trim();
  const cants = [...document.querySelectorAll("#rmFilas .rm-cant")].map(e => e.value.trim());
  const descs = [...document.querySelectorAll("#rmFilas .rm-desc")].map(e => e.value.trim());
  let filas = cants.map((c, i) => ({ cant: c, desc: descs[i] })).filter(f => f.cant || f.desc);
  const valor = v("rmValor");
  if (valor) filas.push({ cant: "", desc: "Mercadería en tránsito $" + valor });
  return {
    dia: v("rmDia"), mes: v("rmMes"), anio: v("rmAnio"),
    cliente: v("rmCliente"), domicilio: v("rmDomicilio"), localidad: v("rmLocalidad"),
    provincia: v("rmProvincia"), telefono: v("rmTelefono"), cuit: v("rmCuit"),
    transporte: v("rmTransporte"), iva: v("rmIva"), bultos: v("rmBultos"), filas
  };
}

function _remitoGuardarOffset() {
  const x = parseFloat(document.getElementById("rmOffX").value) || 0;
  const y = parseFloat(document.getElementById("rmOffY").value) || 0;
  try { localStorage.setItem(REMITO_OFFSET_KEY, JSON.stringify({ x, y })); } catch (e) {}
}

function remitoImprimir() {
  _remitoGuardarOffset();
  const d = _remitoLeerForm();
  if (d.filas.length > REMITO_POS.filasN) {
    toast(`Hay más de ${REMITO_POS.filasN} líneas: no entran en el formulario.`, "error");
    return;
  }
  _remitoImprimir(_remitoHTML(d));
}

/** Hoja de prueba: datos de ejemplo + cruces rojas en cada punto de anclaje. */
function remitoImprimirPrueba() {
  _remitoGuardarOffset();
  const d = _remitoLeerForm();
  d.calibracion = true;
  if (!d.cliente) d.cliente = "CLIENTE DE PRUEBA";
  if (!d.filas.length) d.filas = [{ cant: 1, desc: "Línea de prueba 1" }, { cant: 2, desc: "Línea de prueba 2" }];
  if (!d.iva) d.iva = "RI";
  _remitoImprimir(_remitoHTML(d));
}

function _remitoAsegurarModal() {
  if (document.getElementById("remitoModalBackdrop")) return;
  const div = document.createElement("div");
  div.className = "product-modal-backdrop";
  div.id = "remitoModalBackdrop";
  const campo = (id, label, extra) => `<div style="flex:1; min-width:140px;"><label style="font-size:12px; color:#64748b;">${label}</label><input id="${id}" class="form-control" ${extra || ""}></div>`;
  div.innerHTML = `
  <div class="product-modal" style="width:min(640px, 96vw);">
    <div class="product-modal-header"><h6>Imprimir remito (formulario preimpreso)</h6><button onclick="cerrarModalRemito()">✕</button></div>
    <div class="product-modal-body" style="overflow:auto;">
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:8px;">
        ${campo("rmDia", "Día", 'style="max-width:70px;"')}${campo("rmMes", "Mes", 'style="max-width:70px;"')}${campo("rmAnio", "Año", 'style="max-width:70px;"')}
        <div style="flex:2; min-width:200px;"><label style="font-size:12px; color:#64748b;">Condición IVA (casillero)</label>
          <select id="rmIva" class="form-select"><option value="">Sin marcar</option><option value="CF">Cons. Final</option><option value="RI">Resp. Inscripto</option><option value="EX">Exento</option><option value="MT">Resp. Monotributo</option></select></div>
      </div>
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:8px;">
        ${campo("rmCliente", "Señor/es")}${campo("rmCuit", "C.U.I.T. Nº")}
      </div>
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:8px;">
        ${campo("rmDomicilio", "Domicilio")}${campo("rmLocalidad", "Localidad")}
      </div>
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:8px;">
        ${campo("rmProvincia", "Provincia")}${campo("rmTelefono", "Tel.")}${campo("rmTransporte", "Transportista")}
      </div>
      <div style="display:flex; align-items:center; gap:8px; margin:12px 0 6px;">
        <strong style="font-size:13px;">Cantidad / Descripción</strong>
        <button class="btn btn-outline-secondary btn-sm" type="button" onclick="remitoCargarResumen()">Resumen (1 bulto)</button>
        <button class="btn btn-outline-secondary btn-sm" type="button" onclick="remitoCargarProductos()">Cargar productos del pedido</button>
      </div>
      <div id="rmFilas" style="max-height:230px; overflow:auto;"></div>
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:8px;">
        ${campo("rmBultos", "Cantidad de bultos (número grande + en letras)", 'type="number" min="1" oninput="remitoCargarResumen()"')}
        ${campo("rmDespacho", "N° ICO (8 dígitos, se parte en 4 + 4)", 'maxlength="8" placeholder="opcional" oninput="remitoCargarResumen()"')}
      </div>
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:8px;">
        ${campo("rmValor", "Valor declarado ($) — agrega “Mercadería en tránsito $…”")}
      </div>
      <details style="margin-top:12px;"><summary style="cursor:pointer; font-size:13px;">Calibración de impresora (mm)</summary>
        <div style="display:flex; gap:8px; align-items:flex-end; margin-top:8px;">
          ${campo("rmOffX", "Desplazar →  (mm, + derecha)", 'type="number" step="0.5"')}${campo("rmOffY", "Desplazar ↓  (mm, + abajo)", 'type="number" step="0.5"')}
          <button class="btn btn-outline-secondary" type="button" onclick="remitoImprimirPrueba()">Hoja de prueba</button>
        </div>
        <div style="font-size:12px;color:#64748b;margin-top:4px;">Imprimí la prueba en una hoja común, superponela al formulario a contraluz y corregí el desfasaje.</div>
      </details>
    </div>
    <div class="product-modal-actions">
      <button class="btn btn-outline-secondary" onclick="cerrarModalRemito()">Cancelar</button>
      <button class="btn btn-success" onclick="remitoImprimir()">Imprimir remito</button>
    </div>
  </div>`;
  document.body.appendChild(div);
}

/**
 * Normaliza un número de teléfono argentino al formato internacional
 * que requiere wa.me (sin +, sin espacios, con código de país).
 * Maneja: 011XXXXXXXX, 15XXXXXXXX, 9XXXXXXXX, 54911XXXXXXXX, etc.
 */
function normalizarTelefonoWA(tel) {
  if (!tel) return null;

  // Quita todo lo que no sea dígito
  let num = String(tel).replace(/\D/g, "");

  // Si ya empieza con 54 y tiene 13 dígitos → ya está listo
  if (num.startsWith("54") && num.length === 13) return num;

  // Si empieza con 54 pero le falta el 9 de celular (5411XXXXXXXX → 54911XXXXXXXX)
  if (num.startsWith("54") && num.length === 12) return "549" + num.slice(2);

  // Quita el 0 inicial de área si lo tiene (0351... → 351...)
  if (num.startsWith("0")) num = num.slice(1);

  // Quita el 15 de celular si lo tiene después del área (351158... → 35158...)
  // Área argentina: 2-4 dígitos, luego el número. El 15 aparece en posición variable.
  // Estrategia: si el número tiene exactamente 10 dígitos (área 2-4 + número 6-8)
  // lo armamos directo. Si tiene 11, probablemente tenga el 15 incluido.
  if (num.length === 11 && num.charAt(2) === "1" && num.charAt(3) === "5") {
    num = num.slice(0, 2) + num.slice(4); // quita el 15 tras el área de 2 dígitos
  } else if (num.length === 11 && num.charAt(3) === "1" && num.charAt(4) === "5") {
    num = num.slice(0, 3) + num.slice(5); // quita el 15 tras el área de 3 dígitos
  } else if (num.length === 11 && num.charAt(4) === "1" && num.charAt(5) === "5") {
    num = num.slice(0, 4) + num.slice(6); // quita el 15 tras el área de 4 dígitos
  }

  // Agrega código de país Argentina + 9 de celular
  // (WA Argentina requiere 549 + área + número = 13 dígitos total)
  if (num.length === 10) return "549" + num;
  if (num.length === 8)  return "5491" + num; // número corto sin área, asume GBA

  return null; // no se pudo normalizar — el botón no se muestra
}

/** Genera el mensaje de WhatsApp según el estado del pedido */
function mensajeWhatsAppPorEstado(estado, pedidoId, cliente, total, pdfUrl) {
  const totalStr = "$" + Number(total || 0).toLocaleString("es-AR");
  const linkPdf = pdfUrl ? `\n\nComprobante de pedido: ${pdfUrl}` : "";
  const cfg = obtenerConfigNegocio();
  const cbu = cfg.cbuTransferencia || CONFIG_NEGOCIO_DEFAULT.cbuTransferencia;
  const titular = cfg.nombreTitularCbu || CONFIG_NEGOCIO_DEFAULT.nombreTitularCbu;
  switch(estado) {
    case "NUEVO":
      return `Hola ${cliente}! Recibimos tu pedido ${pedidoId} por ${totalStr}.\n\nPara confirmar tu pedido, realizá la transferencia a:\nCBU: ${cbu}\nNombre: ${titular}\n\nUna vez realizado el pago, envianos el comprobante por este medio para que podamos empezar a preparar tu pedido. Gracias por elegirnos!${linkPdf}`;
    case "PREPARANDO":
      return `Hola ${cliente}! Tu pedido ${pedidoId} por ${totalStr} ya esta siendo preparado. En cuanto este listo te avisamos. Gracias por tu compra!${linkPdf}`;
    case "ENVIADO":
      return `Hola ${cliente}! Tu pedido ${pedidoId} por ${totalStr} ya fue enviado y esta en camino. Pronto lo tenes en tus manos!${linkPdf}`;
    default:
      return null;
  }
}

let _renderGenPedidos = 0; // evita que un render viejo (todavía terminando sus chunks) escriba encima de uno más nuevo

function renderPedidos(listaOriginal) {
  const container = document.getElementById("tablaPedidos");
  if (!container) return;

  // Se ordena una COPIA, más nuevo primero — nunca a "lista" en sí ni
  // a pedidosGlobal, porque cargarPedidos() compara el orden crudo del
  // backend contra el anterior para saber si hubo cambios (y así
  // evitar redibujar cuando no hace falta). Si acá se ordenara el
  // array real, esa comparación dejaría de funcionar y volvería el
  // salto de pantalla en cada actualización.
  const lista = [...listaOriginal].sort((a, b) => new Date(b.FECHA) - new Date(a.FECHA));

  const miGen = ++_renderGenPedidos;

  // Guarda dónde estaba parado el usuario para restaurarlo después de
  // reconstruir la lista — sin esto, cada refresco hacía "saltar" la
  // pantalla al principio aunque el usuario estuviera leyendo un
  // pedido más abajo.
  const scrollContainer = container.closest(".seccion") || container.parentElement;
  const scrollTopPrevio = scrollContainer ? scrollContainer.scrollTop : 0;

  if (lista.length === 0) {
    container.innerHTML = `<div class="text-center text-muted py-5" style="font-size:14px;">No se encontraron pedidos</div>`;
    return;
  }

  const estadoClase = { NUEVO:"nuevo", PREPARANDO:"preparando", ENVIADO:"enviado", CANCELADO:"cancelado" };
  const estadoIcono = { NUEVO:"🆕", PREPARANDO:"⚙️", ENVIADO:"📦", CANCELADO:"❌" };

  const CHUNK = 15;
  let idx = 0;
  container.innerHTML = "";

  const renderChunk = () => {
    if (miGen !== _renderGenPedidos) return; // superado por un render más nuevo — no seguir escribiendo
    const frag = document.createDocumentFragment();
    const fin = Math.min(idx + CHUNK, lista.length);
    for (; idx < fin; idx++) {
      const p = lista[idx];
      const claseEstado = estadoClase[p.ESTADO] || "";
      const estaCobrado = String(p.COBRADO || "").toUpperCase() === "SI";
      const formaPagoActual = String(p.FORMA_PAGO_COBRO || "");
      const estadosConWA = ["NUEVO", "PREPARANDO", "ENVIADO"];
      const telefono = String(p.TELEFONO || "").trim();
      const numeroWA = normalizarTelefonoWA(telefono);
      const mensajeWA = mensajeWhatsAppPorEstado(p.ESTADO, p.PEDIDO_ID, p.CLIENTE, p.TOTAL, p.PDF_URL);
      // Si el teléfono no tiene ningún dígito, no hay nada que mandar
      // — ahí sí no tiene sentido mostrar el botón. Pero si HAY
      // dígitos y solo falló el reconocimiento del formato (área rara,
      // como pasa con algunos números fijos o de localidades chicas),
      // se muestra igual con el mejor intento posible (los dígitos tal
      // cual, con el prefijo de Argentina) en vez de ocultarlo del
      // todo sin avisar — antes esto desaparecía en silencio y no
      // había forma de saber por qué.
      const soloDigitos = telefono.replace(/\D/g, "");
      const numeroWAConFallback = numeroWA || (soloDigitos ? "54" + soloDigitos.replace(/^0+/, "") : null);
      const esNumeroDudoso = !numeroWA && !!numeroWAConFallback;
      const btnWA = (estadosConWA.includes(p.ESTADO) && numeroWAConFallback && mensajeWA)
        ? `<a href="https://wa.me/${numeroWAConFallback}?text=${encodeURIComponent(mensajeWA)}" target="_blank" class="btn btn-success btn-sm" title="${esNumeroDudoso ? "Verificá el teléfono antes de mandar — el formato no se reconoció con seguridad" : "Notificar por WhatsApp"}">WhatsApp${esNumeroDudoso ? " ⚠️" : ""}</a>`
        : "";
      const badgeCobrado = estaCobrado
        ? `<span class="pedido-cobrado-badge">✓ Cobrado <span class="pedido-cobrado-forma">${formaPagoActual}</span></span>` : "";
      const selectCobrado = estaCobrado
        ? `<div class="pedido-cobrado-controles">
             <span class="pedido-cobrado-badge">✓ Cobrado</span>
             <span class="pedido-cobrado-forma">${formaPagoActual}</span>
             <button class="btn btn-outline-danger btn-sm" style="font-size:11px;padding:2px 8px;"
               onclick="aplicarCobroPedido('${p.PEDIDO_ID}', false, '', this)">Desmarcar</button>
           </div>`
        : `<div class="pedido-pago-btns" data-pedido="${p.PEDIDO_ID}">
             <span style="font-size:11px;font-weight:600;color:var(--slate-500);">Cobrar con:</span>
             <button class="pedido-pago-btn" onclick="aplicarCobroPedido('${p.PEDIDO_ID}', true, 'EFECTIVO', this)">Efectivo</button>
             <button class="pedido-pago-btn" onclick="aplicarCobroPedido('${p.PEDIDO_ID}', true, 'TRANSFERENCIA', this)">Transfer.</button>
             <button class="pedido-pago-btn" onclick="aplicarCobroPedido('${p.PEDIDO_ID}', true, 'TARJETA', this)">Tarjeta</button>
           </div>`;
      const tmp = document.createElement("div");
      tmp.innerHTML = `<div class="pedido-card estado-${claseEstado}${_pedidosEnProceso.has(String(p.PEDIDO_ID)) ? " pedido-procesando" : ""}" data-pedido-card="${escapeHtml(p.PEDIDO_ID)}">
        <div class="pedido-card-top">
          <div>
            <div class="pedido-card-id">${escapeHtml(p.PEDIDO_ID)}${p._antiguo ? ' <span class="pedido-antiguo-badge" title="Pedido de hace más de 30 días: no está en la lista general, apareció por la búsqueda">Anterior</span>' : ""}</div>
            <div class="pedido-card-cliente">${escapeHtml(p.CLIENTE || p.NOMBRE || "")}${p.EMPRESA ? ` <span style="font-weight:500;color:var(--slate-500);font-size:13px;">· 🚚 ${escapeHtml(p.EMPRESA)}</span>` : ""}</div>
            ${p.DIRECCION ? `<div class="pedido-card-dir">📍 ${escapeHtml(p.DIRECCION)}${p.LOCALIDAD ? ", " + escapeHtml(p.LOCALIDAD) : ""}</div>` : ""}
            <div class="pedido-card-fecha">📅 ${new Date(p.FECHA).toLocaleDateString("es-AR", {day:"2-digit", month:"2-digit", year:"numeric"})}</div>
          </div>
          <div class="text-end">
            <div class="pedido-card-total">$${Number(p.TOTAL || 0).toLocaleString("es-AR")}</div>
            <div class="mt-1"><span class="pedido-estado-badge ${claseEstado}">${estadoIcono[p.ESTADO] || ""} ${escapeHtml(p.ESTADO)}</span></div>
            ${badgeCobrado ? `<div class="mt-1">${badgeCobrado}</div>` : ""}
          </div>
        </div>
        <div class="pedido-card-controls">
          <select class="form-select form-select-sm" style="max-width:160px;" data-estado-previo="${escapeHtml(p.ESTADO)}" onchange="cambiarEstado('${p.PEDIDO_ID}',this.value,this)">
            <option value="NUEVO"      ${p.ESTADO==="NUEVO"?"selected":""}>🆕 Nuevo</option>
            <option value="PREPARANDO" ${p.ESTADO==="PREPARANDO"?"selected":""}>Preparando</option>
            <option value="ENVIADO"    ${p.ESTADO==="ENVIADO"?"selected":""}>Enviado</option>
            <option value="CANCELADO"  ${p.ESTADO==="CANCELADO"?"selected":""}>Cancelado</option>
          </select>
          ${btnWA}
          ${selectCobrado}
          <button class="btn btn-outline-secondary btn-sm ms-auto" onclick="abrirDetallePedido('${p.PEDIDO_ID}')">Ver / Reimprimir</button>
        </div>
      </div>`;
      frag.appendChild(tmp.firstElementChild);
    }
    container.appendChild(frag);
    if (idx < lista.length) {
      requestAnimationFrame(renderChunk);
    } else if (scrollContainer) {
      // Recién al terminar de reconstruir toda la lista se restaura el
      // scroll — hacerlo antes dejaría al contenedor más bajo de lo
      // que todavía existe en el DOM.
      scrollContainer.scrollTop = scrollTopPrevio;
    }
  };

  requestAnimationFrame(renderChunk);
}

/* ===================== PRODUCTOS (tabla admin) ===================== */

async function cargarProductos() {
  // 1. Mostrar desde caché instantáneamente si existe y es fresco
  const CACHE_KEY = "vpos_cache_productosAdmin";
  const CACHE_TTL = 5 * 60 * 1000; // 5 minutos
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (raw) {
      const { ts, data } = JSON.parse(raw);
      if (data && data.length > 0) {
        productosAdminGlobal = data;
        poblarFiltroCategoriasProductos();
        filtrarProductos();
        // Si el caché es fresco, no recargar del server
        if (Date.now() - ts < CACHE_TTL) return;
        // Si está vencido, actualizar en segundo plano sin bloquear
        _actualizarProductosAdminEnBackground(CACHE_KEY);
        return;
      }
    }
  } catch(e) {}

  // 2. Sin caché: mostrar skeleton y cargar del server
  const tbody = document.getElementById("tablaProductos");
  if (tbody) tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-4">Cargando productos...</td></tr>`;
  await _actualizarProductosAdminEnBackground(CACHE_KEY);
}

async function actualizarCatalogoProductosManual() {
  try { localStorage.removeItem("vpos_cache_productosAdmin"); } catch(e) {}
  productosAdminGlobal = [];
  toast("Actualizando productos...", "success");
  await cargarProductos();
}

async function _actualizarProductosAdminEnBackground(cacheKey) {
  try {
    const response = await fetchAPI(API_URL + "?action=productosAdmin");
    const data = await response.json();
    if (!data.productos) return;
    productosAdminGlobal = data.productos;
    // En escritorio el catálogo vive en SQLite (pos-offline.js): no duplicarlo en localStorage
    if (!_esAppEscritorio()) { try { localStorage.setItem(cacheKey, JSON.stringify({ ts: Date.now(), data: data.productos })); } catch(e) {} }
    poblarFiltroCategoriasProductos();
    filtrarProductos();
  } catch (error) {
    console.error("Error productos:", error);
    toast("No se pudieron actualizar los productos — revisá tu conexión", "error");
  }
}

let _renderGenProductos = 0; // evita que un render viejo (todavía completando sus chunks) mezcle filas con uno más nuevo

function renderTablaProductos(lista) {
  const tbody = document.getElementById("tablaProductos");
  if (!tbody) return;

  const miGen = ++_renderGenProductos;

  if (!lista || lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-4">No se encontraron productos</td></tr>`;
    return;
  }

  // El arrastre para reordenar solo tiene sentido si la tabla muestra
  // TODOS los productos, sin ningún filtro/búsqueda activo — si no, el
  // orden visual de un subconjunto filtrado no representa el orden
  // real del catálogo completo, y guardarlo así lo desordenaría.
  const ordenamientoHabilitado = _puedeReordenarProductos();
  const aviso = document.getElementById("avisoOrdenProductos");
  if (aviso) aviso.style.display = (ordenamientoHabilitado && lista.length > 1) ? "flex" : "none";

  // Renderizar en chunks para no bloquear el hilo principal
  // con 829 productos generando DOM de golpe
  const CHUNK = 80;
  let idx = 0;
  tbody.innerHTML = "";

  const renderChunk = () => {
    // Si mientras este render todavía estaba completando sus chunks
    // (varios requestAnimationFrame seguidos) el usuario tipeó de
    // nuevo y arrancó un render más nuevo, este viejo tiene que
    // abandonar en vez de seguir agregando filas — si no, sus filas
    // (de una búsqueda vieja) terminan mezcladas con las del render
    // nuevo, o incluso apareciendo solas después de que el nuevo ya
    // había limpiado y llenado la tabla. Esto era lo que hacía que la
    // búsqueda "no mostrara bien los resultados hasta volver a escribir".
    if (miGen !== _renderGenProductos) return;

    const frag = document.createDocumentFragment();
    const fin = Math.min(idx + CHUNK, lista.length);

    for (; idx < fin; idx++) {
      const p = lista[idx];
      const publicado = String(p.PUBLICADO || "").toUpperCase() === "SI";
      const stock = Number(p.STOCK || 0);
      const stockBadge = stock < 0
        ? `<span class="tile-stock negativo" title="Se vendió sin tener stock cargado">${stock}</span>`
        : stock === 0
        ? `<span class="tile-stock out">Sin stock</span>`
        : (stock <= 5 ? `<span class="tile-stock low">${stock}</span>` : stock);

      const imagenUrl = p.IMAGEN ? String(p.IMAGEN).trim() : "";
      const fotoHtml = imagenUrl
        ? `<img src="${escapeHtml(imagenUrl)}" alt="" loading="lazy" onerror="this.parentElement.innerHTML='🛒';">`
        : "🛒";

      const tr = document.createElement("tr");
      tr.dataset.codigo = p.CODIGO;
      if (ordenamientoHabilitado) {
        tr.draggable = true;
        tr.classList.add("fila-arrastrable");
        tr.addEventListener("dragstart", _dragStartFilaProducto);
        tr.addEventListener("dragover", _dragOverFilaProducto);
        tr.addEventListener("drop", _dropFilaProducto);
        tr.addEventListener("dragend", _dragEndFilaProducto);
      }
      tr.innerHTML = `
        <td class="celda-check-orden">
          <span class="check-orden-inner">
            <span class="drag-handle-producto" title="${ordenamientoHabilitado ? "Arrastrá para reordenar" : "Limpiá la búsqueda y los filtros para poder reordenar"}">${ordenamientoHabilitado ? "⠿" : ""}</span>
            <input type="checkbox" class="check-producto-etiqueta" value="${escapeHtml(p.CODIGO)}" onchange="alCambiarCheckProducto(this)"${_seleccionProductos.has(String(p.CODIGO)) ? " checked" : ""}>
          </span>
        </td>
        <td class="celda-foto-producto"><div class="tabla-producto-thumb">${fotoHtml}</div></td>
        <td class="mono">${escapeHtml(p.CODIGO)}</td>
        <td>${escapeHtml(p.PRODUCTO)}</td>
        <td>${escapeHtml(p.CATEGORIA || "—")}</td>
        <td class="money">${typeof depPrecioVista === "function" ? depPrecioVista(p.PRECIO, p, true) : "$" + Number(p.PRECIO || 0).toLocaleString("es-AR")}</td>
        <td>${stockBadge}</td>
        <td><span class="badge ${publicado ? "bg-success" : "bg-secondary"}">${publicado ? "Publicado" : "Oculto"}</span>${String(p.DESTACADO || "").trim().toUpperCase() === "SI" ? ' <span class="marca-prod" title="Destacado">⭐</span>' : ""}${String(p.OFERTA || "").trim().toUpperCase() === "SI" ? ' <span class="marca-prod" title="En oferta">🔥</span>' : ""}</td>
        <td class="celda-acciones-producto">
          <button class="btn btn-outline-success btn-sm btn-accion-producto" onclick="abrirModalStock('${escapeJsAttr(p.CODIGO)}')" title="Sumar stock">📦 <span class="btn-accion-texto">Stock</span></button>
          <button class="btn btn-primary btn-sm btn-accion-producto ms-2" onclick="editarProducto('${escapeJsAttr(p.CODIGO)}')" title="Editar">✏️ <span class="btn-accion-texto">Editar</span></button>
          <button class="btn btn-danger btn-sm btn-accion-producto ms-2" onclick="eliminarProducto('${escapeJsAttr(p.CODIGO)}')" title="Eliminar">🗑️ <span class="btn-accion-texto">Eliminar</span></button>
        </td>`;
      frag.appendChild(tr);
    }

    tbody.appendChild(frag);

    if (idx < lista.length) {
      // Ceder control al browser entre chunks para que no congele la UI
      requestAnimationFrame(renderChunk);
    } else {
      // Fin del render — la casilla del encabezado refleja los visibles
      actualizarSeleccionEtiquetas();
    }
  };

  requestAnimationFrame(renderChunk);
}

/** True solo si la tabla está mostrando TODOS los productos, sin ninguna búsqueda/filtro activo — condición para poder arrastrar y reordenar filas. */
function _puedeReordenarProductos() {
  const termino = document.getElementById("buscarProducto")?.value.trim() || "";
  const categoria = document.getElementById("filtroCategoriaProductos")?.value || "";
  const estado = document.getElementById("filtroEstadoProducto")?.value || "";
  return !termino && !categoria && !estado;
}

let _filaArrastrada = null;

function _dragStartFilaProducto(ev) {
  _filaArrastrada = ev.currentTarget;
  ev.currentTarget.classList.add("arrastrando");
  ev.dataTransfer.effectAllowed = "move";
  // Firefox necesita datos seteados en dataTransfer para permitir el arrastre
  ev.dataTransfer.setData("text/plain", ev.currentTarget.dataset.codigo || "");
}

function _dragOverFilaProducto(ev) {
  ev.preventDefault();
  if (!_filaArrastrada || _filaArrastrada === ev.currentTarget) return;

  const tbody = ev.currentTarget.parentNode;
  const filas = [...tbody.querySelectorAll("tr")];
  const idxArrastrada = filas.indexOf(_filaArrastrada);
  const idxDestino = filas.indexOf(ev.currentTarget);

  if (idxArrastrada < idxDestino) {
    ev.currentTarget.after(_filaArrastrada);
  } else {
    ev.currentTarget.before(_filaArrastrada);
  }
}

function _dropFilaProducto(ev) {
  ev.preventDefault();
}

async function _dragEndFilaProducto(ev) {
  ev.currentTarget.classList.remove("arrastrando");
  if (!_filaArrastrada) return;
  _filaArrastrada = null;

  const tbody = document.getElementById("tablaProductos");
  const nuevoOrdenCodigos = [...tbody.querySelectorAll("tr")].map(tr => tr.dataset.codigo).filter(Boolean);
  await guardarNuevoOrdenProductos(nuevoOrdenCodigos);
}

/** Persiste en el backend el nuevo orden de productos (array de CODIGO, en el orden deseado) y actualiza la caché local para que no vuelva al orden viejo al recargar. */
async function guardarNuevoOrdenProductos(codigos) {
  try {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "reordenarProductos", codigos })
    }, { timeoutMs: 20000 });
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar el nuevo orden", "error");
      cargarProductos();
      return false;
    }

    // Reordenar en memoria (y en la caché local) según lo recién guardado,
    // así el orden se ve correcto sin esperar un nuevo pedido al servidor.
    const posicion = new Map(codigos.map((c, i) => [c, i]));
    productosAdminGlobal = [...productosAdminGlobal].sort((a, b) => {
      const pa = posicion.has(a.CODIGO) ? posicion.get(a.CODIGO) : Infinity;
      const pb = posicion.has(b.CODIGO) ? posicion.get(b.CODIGO) : Infinity;
      return pa - pb;
    });
    try {
      localStorage.setItem("vpos_cache_productosAdmin", JSON.stringify({ ts: Date.now(), data: productosAdminGlobal }));
    } catch (e) {}

    toast("Orden guardado", "success");
    return true;
  } catch (error) {
    console.error("Error al guardar el orden de productos:", error);
    toast("Error de conexión al guardar el orden", "error");
    cargarProductos();
    return false;
  }
}

/* =====================================================================
   ORDENAR CATÁLOGO — pantalla para elegir el orden de las categorías
   y, dentro de cada una, el de sus productos (con fotos, arrastrando o
   con botones). Al guardar se arma la lista completa: categoría 1 con
   sus productos, categoría 2 con los suyos… y se manda al mismo
   "reordenarProductos" de siempre (columna ORDEN). No hace falta tocar
   el backend. La página pública muestra el cambio en pocos minutos.
===================================================================== */
const OC_SIN_CATEGORIA = "__sin_categoria__";
let _oc = null;          // { cats:[clave], porCat:{clave:[codigo]}, prod:Map, sel, sucio }
let _ocArrastre = null;  // elemento que se está arrastrando

function _ocClaveCat(p) { return String(p.CATEGORIA || "").trim() || OC_SIN_CATEGORIA; }
function _ocNombreCat(clave) { return clave === OC_SIN_CATEGORIA ? "Sin categoría" : clave; }

async function abrirOrdenCatalogo() {
  if (!productosAdminGlobal || !productosAdminGlobal.length) {
    toast("Cargando productos...", "success");
    await cargarProductos();
  }
  if (!productosAdminGlobal || !productosAdminGlobal.length) { toast("No hay productos para ordenar", "error"); return; }
  const cats = [], porCat = {}, prod = new Map();
  productosAdminGlobal.forEach(p => {   // ya viene ordenado por ORDEN
    const clave = _ocClaveCat(p);
    if (!porCat[clave]) { porCat[clave] = []; cats.push(clave); }
    porCat[clave].push(String(p.CODIGO));
    prod.set(String(p.CODIGO), p);
  });
  _oc = { cats, porCat, prod, sel: cats[0], sucio: false };
  document.getElementById("ordenCatalogoBackdrop").classList.add("show");
  _ocRender();
}

function cerrarOrdenCatalogo(forzar) {
  // Modal propio (no confirm() nativo: en Electron deja el teclado sin foco)
  if (!forzar && _oc && _oc.sucio) {
    confirmarAccion("Hay cambios de orden sin guardar. ¿Cerrar igual y descartarlos?", () => cerrarOrdenCatalogo(true), "↕️ Cambios sin guardar");
    return;
  }
  document.getElementById("ordenCatalogoBackdrop").classList.remove("show");
  _oc = null;
}

function _ocMarcarSucio() {
  _oc.sucio = true;
  const est = document.getElementById("ocEstado");
  est.textContent = "Cambios sin guardar";
  est.classList.add("sucio");
  document.getElementById("btnGuardarOrdenCatalogo").disabled = false;
}

function _ocRender() { _ocRenderCats(); _ocRenderProds(); }

function _ocRenderCats() {
  const ul = document.getElementById("ocCategorias");
  const n = _oc.cats.length;
  ul.innerHTML = _oc.cats.map((c, i) => `
    <li class="oc-cat ${c === _oc.sel ? "activa" : ""}" draggable="true" data-cat="${escapeHtml(c)}" onclick="ocElegirCategoria(this.dataset.cat)">
      <span class="oc-handle" title="Arrastrar">⠿</span>
      <span class="oc-num">${i + 1}</span>
      <span class="oc-nombre" title="${escapeHtml(_ocNombreCat(c))}">${escapeHtml(_ocNombreCat(c))}</span>
      <span class="oc-cant">${_oc.porCat[c].length}</span>
      <button type="button" class="oc-mini" title="Subir" ${i === 0 ? "disabled" : ""} onclick="event.stopPropagation(); ocMoverCategoria(${i}, -1)">▲</button>
      <button type="button" class="oc-mini" title="Bajar" ${i === n - 1 ? "disabled" : ""} onclick="event.stopPropagation(); ocMoverCategoria(${i}, 1)">▼</button>
    </li>`).join("");
  ul.querySelectorAll(".oc-cat").forEach(li => _ocHacerArrastrable(li, "y", _ocFinArrastreCats));
}

function _ocRenderProds(destacar) {
  document.getElementById("ocCatNombre").textContent = _ocNombreCat(_oc.sel);
  const grid = document.getElementById("ocProductos");
  const lista = _oc.porCat[_oc.sel] || [];
  if (!lista.length) { grid.innerHTML = `<div class="oc-vacio">Esta categoría no tiene productos.</div>`; return; }
  const n = lista.length;
  grid.innerHTML = lista.map((cod, i) => {
    const p = _oc.prod.get(cod) || {};
    const img = String(p.IMAGEN || "").trim();
    const oculto = String(p.PUBLICADO || "").toUpperCase() !== "SI";
    return `
      <div class="oc-card ${cod === destacar ? "recien-movido" : ""}" draggable="true" data-codigo="${escapeHtml(cod)}">
        <span class="oc-pos" title="Escribir la posición" onclick="ocEditarPosicion(this, ${i})">${i + 1}</span>
        ${oculto ? `<span class="oc-oculto" title="No se muestra en la página">Oculto</span>` : ""}
        <div class="oc-foto">${img ? `<img src="${escapeHtml(img)}" alt="" loading="lazy" onerror="this.parentElement.textContent='🛒'">` : "🛒"}</div>
        <div class="oc-nombre-prod" title="${escapeHtml(p.PRODUCTO || cod)}">${escapeHtml(p.PRODUCTO || cod)}</div>
        <div class="oc-precio">$${Number(p.PRECIO || 0).toLocaleString("es-AR")}</div>
        <div class="oc-acciones">
          <button type="button" class="oc-mini" title="Al principio" ${i === 0 ? "disabled" : ""} onclick="ocMoverProducto(${i}, 0)">⤒</button>
          <button type="button" class="oc-mini" title="Uno antes" ${i === 0 ? "disabled" : ""} onclick="ocMoverProducto(${i}, ${i - 1})">◀</button>
          <button type="button" class="oc-mini" title="Uno después" ${i === n - 1 ? "disabled" : ""} onclick="ocMoverProducto(${i}, ${i + 1})">▶</button>
          <button type="button" class="oc-mini" title="Al final" ${i === n - 1 ? "disabled" : ""} onclick="ocMoverProducto(${i}, ${n - 1})">⤓</button>
        </div>
      </div>`;
  }).join("");
  grid.querySelectorAll(".oc-card").forEach(card => _ocHacerArrastrable(card, "xy", _ocFinArrastreProds));
  if (destacar) {
    const el = grid.querySelector(`.oc-card[data-codigo="${cssEscape(destacar)}"]`);
    if (el) el.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}

function ocElegirCategoria(clave) {
  if (!_oc || clave === _oc.sel) return;
  _oc.sel = clave;
  _ocRender();
  document.getElementById("ocProductos").scrollTop = 0;
}

function ocMoverCategoria(i, delta) {
  const j = i + delta;
  if (j < 0 || j >= _oc.cats.length) return;
  const [c] = _oc.cats.splice(i, 1);
  _oc.cats.splice(j, 0, c);
  _ocMarcarSucio();
  _ocRenderCats();
}

function ocMoverProducto(desde, hasta) {
  const lista = _oc.porCat[_oc.sel];
  hasta = Math.max(0, Math.min(lista.length - 1, hasta));
  if (desde === hasta) return;
  const [cod] = lista.splice(desde, 1);
  lista.splice(hasta, 0, cod);
  _ocMarcarSucio();
  _ocRenderProds(cod);
}

/** Clic en el número: se convierte en un campo para escribir la posición (Enter para mover). */
function ocEditarPosicion(badge, i) {
  const n = _oc.porCat[_oc.sel].length;
  const input = document.createElement("input");
  input.type = "number"; input.min = 1; input.max = n; input.value = i + 1;
  input.className = "oc-pos"; input.style.width = "56px"; input.style.border = "none"; input.style.outline = "2px solid var(--blue-500)";
  let hecho = false;
  const aplicar = () => {
    if (hecho) return; hecho = true;
    const destino = parseInt(input.value, 10);
    if (isFinite(destino) && destino >= 1) ocMoverProducto(i, Math.min(destino, n) - 1);
    else _ocRenderProds();
  };
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); aplicar(); }
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); hecho = true; _ocRenderProds(); }
  });
  input.addEventListener("blur", aplicar);
  input.addEventListener("mousedown", e => e.stopPropagation());
  badge.replaceWith(input);
  input.focus(); input.select();
}

function ocOrdenRapido(criterio) {
  const lista = _oc.porCat[_oc.sel];
  const p = c => _oc.prod.get(c) || {};
  const cmp = {
    nombre: (a, b) => String(p(a).PRODUCTO || "").localeCompare(String(p(b).PRODUCTO || ""), "es", { numeric: true, sensitivity: "base" }),
    precioAsc: (a, b) => Number(p(a).PRECIO || 0) - Number(p(b).PRECIO || 0),
    precioDesc: (a, b) => Number(p(b).PRECIO || 0) - Number(p(a).PRECIO || 0),
    stock: (a, b) => Number(p(b).STOCK || 0) - Number(p(a).STOCK || 0)
  }[criterio];
  if (!cmp) return;
  lista.sort(cmp);
  _ocMarcarSucio();
  _ocRenderProds();
  toast("Ordenado — acordate de tocar Guardar orden", "success");
}

/* Arrastre (mouse): mueve el elemento en vivo y al soltar se lee el orden del DOM. */
function _ocHacerArrastrable(el, eje, alTerminar) {
  el.addEventListener("dragstart", ev => {
    _ocArrastre = el;
    el.classList.add("arrastrando");
    ev.dataTransfer.effectAllowed = "move";
    ev.dataTransfer.setData("text/plain", el.dataset.codigo || el.dataset.cat || "");
  });
  el.addEventListener("dragover", ev => {
    if (!_ocArrastre || _ocArrastre === el || _ocArrastre.parentNode !== el.parentNode) return;
    ev.preventDefault();
    const r = el.getBoundingClientRect();
    const despues = eje === "y"
      ? ev.clientY > r.top + r.height / 2
      : (ev.clientY > r.bottom - r.height * 0.25) || (ev.clientY >= r.top && ev.clientX > r.left + r.width / 2);
    if (despues) el.after(_ocArrastre); else el.before(_ocArrastre);
  });
  el.addEventListener("drop", ev => ev.preventDefault());
  el.addEventListener("dragend", () => {
    el.classList.remove("arrastrando");
    if (_ocArrastre) { _ocArrastre = null; alTerminar(); }
  });
}

function _ocFinArrastreCats() {
  const nuevo = [...document.querySelectorAll("#ocCategorias .oc-cat")].map(li => li.dataset.cat);
  if (nuevo.join("\n") !== _oc.cats.join("\n")) { _oc.cats = nuevo; _ocMarcarSucio(); }
  _ocRenderCats();
}

function _ocFinArrastreProds() {
  const nuevo = [...document.querySelectorAll("#ocProductos .oc-card")].map(c => c.dataset.codigo);
  if (nuevo.join("\n") !== _oc.porCat[_oc.sel].join("\n")) { _oc.porCat[_oc.sel] = nuevo; _ocMarcarSucio(); }
  _ocRenderProds();
}

async function guardarOrdenCatalogo() {
  if (!_oc || !_oc.sucio) return;
  const btn = document.getElementById("btnGuardarOrdenCatalogo");
  btn.disabled = true; btn.textContent = "Guardando...";
  const codigos = [];
  _oc.cats.forEach(c => codigos.push(..._oc.porCat[c]));
  const ok = await guardarNuevoOrdenProductos(codigos);
  btn.textContent = "Guardar orden";
  if (!ok) { btn.disabled = false; return; }
  _oc.sucio = false;
  const est = document.getElementById("ocEstado");
  est.textContent = "✓ Guardado — la página lo muestra en unos minutos";
  est.classList.remove("sucio");
  filtrarProductos();
}

/** Fills the category <select> filter with the distinct categories currently in use */
function poblarFiltroCategoriasProductos() {
  const select = document.getElementById("filtroCategoriaProductos");
  if (!select) return;

  const valorActual = select.value;
  const categorias = [...new Set(
    productosAdminGlobal.map(p => String(p.CATEGORIA || "").trim()).filter(Boolean)
  )].sort();

  select.innerHTML = `<option value="">Todas las categorías</option>` +
    categorias.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");

  // Preserve the user's selection across refreshes when still valid
  if (categorias.includes(valorActual)) select.value = valorActual;
}

/** Client-side filter by name/code (text) and category (select), over the already-loaded product list */
function filtrarProductos() {
  const inputBuscar = document.getElementById("buscarProducto");
  const selectCategoria = document.getElementById("filtroCategoriaProductos");
  const selectEstado = document.getElementById("filtroEstadoProducto");

  const termino = inputBuscar ? inputBuscar.value.trim() : "";
  const categoria = selectCategoria ? selectCategoria.value : "";
  const estado = selectEstado ? selectEstado.value : "";

  let filtrados = productosAdminGlobal;

  if (termino) {
    // Misma búsqueda tolerante a errores de tipeo que usa el POS —
    // también considera el ALIAS del producto, no solo código/nombre/categoría.
    filtrados = filtrados.filter(p => productoCoincideBusquedaPOS(p, termino));
  }

  if (categoria) {
    filtrados = filtrados.filter(p => String(p.CATEGORIA || "").trim() === categoria);
  }

  if (estado === "sin_imagen") {
    filtrados = filtrados.filter(p => !String(p.IMAGEN || "").trim());
  } else if (estado === "stock_bajo") {
    filtrados = filtrados.filter(p => { const st = Number(p.STOCK ?? 0); return st > 0 && st <= 5; });
  } else if (estado === "sin_stock") {
    filtrados = filtrados.filter(p => Number(p.STOCK ?? 0) <= 0);
  } else if (estado === "stock_negativo") {
    // Productos que se vendieron sin tener stock cargado (el POS ya
    // permite hacerlo) — sirve para encontrarlos rápido y actualizar
    // el stock real cuando llegue mercadería o se corrija la carga.
    filtrados = filtrados.filter(p => Number(p.STOCK ?? 0) < 0);
  } else if (estado === "sin_imagen_y_stock") {
    filtrados = filtrados.filter(p => !String(p.IMAGEN || "").trim() && Number(p.STOCK ?? 0) <= 0);
  } else if (estado === "ocultos") {
    filtrados = filtrados.filter(p => String(p.PUBLICADO || "").toUpperCase() !== "SI");
  } else if (estado === "publicados") {
    filtrados = filtrados.filter(p => String(p.PUBLICADO || "").toUpperCase() === "SI");
  } else if (estado === "destacados") {
    filtrados = filtrados.filter(p => String(p.DESTACADO || "").trim().toUpperCase() === "SI");
  } else if (estado === "ofertas") {
    filtrados = filtrados.filter(p => String(p.OFERTA || "").trim().toUpperCase() === "SI");
  }

  renderTablaProductos(filtrados);
}

/**
 * Versión con demora de filtrarProductos, usada SOLO por el tecleo en
 * el buscador (onkeyup) — el cambio de categoría (onchange del
 * <select>) y la carga inicial siguen llamando a filtrarProductos()
 * directo, sin demora, porque ahí no hay riesgo de redibujar la
 * tabla completa en cada tecla.
 */
const filtrarProductosConDemora = debounce(filtrarProductos, 150);

/* ===================== PRODUCTOS — ALTA / EDICIÓN (modal) ===================== */

let productosAdminGlobal = [];

/** Opens the modal in "create" mode */
function nuevoProducto() {
  window._pmStockOriginal = null; // no aplica en modo creación — ver editarProducto
  document.getElementById("productModalTitle").textContent = "+ Nuevo Producto";
  document.getElementById("pmCodigoOriginal").value = "";
  document.getElementById("pmCodigo").value = "";
  document.getElementById("pmCodigo").disabled = false;
  document.getElementById("pmNombre").value = "";
  document.getElementById("pmCategoria").value = "";
  document.getElementById("pmAlias").value = "";
  { const al = document.getElementById("pmAlicuotaIva"); if (al) al.value = ""; }
  document.getElementById("pmPrecio").value = "";
  document.getElementById("pmStock").value = "";
  document.getElementById("pmImagen").value = "";
  document.getElementById("pmImagenArchivo").value = "";
  document.getElementById("pmImagenPreview").innerHTML = "🖼️";
  document.getElementById("pmImagenStatus").textContent = "";
  document.getElementById("pmImagenStatus").className = "pm-image-status";
  document.getElementById("pmColorPrimario").value = "";
  document.getElementById("pmGaleriaWrap").innerHTML = "";
  _pmGaleriaContador = 0;
  document.getElementById("pmPublicado").checked = true;
  document.getElementById("pmDestacado").checked = false;
  document.getElementById("pmOferta").checked = false;

  document.getElementById("pmVentaPorCaja").checked = false;
  document.getElementById("pmCodigoCaja").value = "";
  document.getElementById("pmUnidadesPorCaja").value = "";
  document.getElementById("pmPrecioCaja").value = "";
  document.getElementById("pmCajaWrap").style.display = "none";

  poblarCategoriasDatalist();
  document.getElementById("productModalBackdrop").classList.add("show");
  setTimeout(() => document.getElementById("pmNombre").focus(), 80);
}

/** Opens the modal in "edit" mode, pre-filled with the product's current data */
function editarProducto(codigo) {
  const p = productosAdminGlobal.find(x => String(x.CODIGO) === String(codigo));
  if (!p) { toast("No se encontró el producto para editar", "error"); return; }

  // Se guarda el stock TAL COMO estaba al abrir el formulario — se usa
  // al guardar para mandar solo la DIFERENCIA (delta) en vez de un
  // valor absoluto. Si se mandara el valor absoluto y otra caja vendió
  // este mismo producto mientras el formulario estaba abierto, guardar
  // pisaría ese stock más nuevo con el que se vio acá, desactualizado.
  // Ver guardarProductoForm en pos-offline.js.
  window._pmStockOriginal = Number(p.STOCK || 0);

  document.getElementById("productModalTitle").textContent = "Editar Producto";
  document.getElementById("pmCodigoOriginal").value = p.CODIGO;
  document.getElementById("pmCodigo").value = p.CODIGO;
  document.getElementById("pmNombre").value = p.PRODUCTO || "";
  document.getElementById("pmCategoria").value = p.CATEGORIA || "";
  document.getElementById("pmAlias").value = p.ALIAS || "";
  { const al = document.getElementById("pmAlicuotaIva"); if (al) al.value = String(p.ALICUOTA_IVA === undefined || p.ALICUOTA_IVA === null ? "" : p.ALICUOTA_IVA); }
  document.getElementById("pmPrecio").value = Number(p.PRECIO || 0);
  document.getElementById("pmStock").value = Number(p.STOCK || 0);
  document.getElementById("pmImagen").value = p.IMAGEN || "";
  document.getElementById("pmImagenArchivo").value = "";
  document.getElementById("pmImagenStatus").textContent = "";
  document.getElementById("pmImagenStatus").className = "pm-image-status";
  document.getElementById("pmColorPrimario").value = p.COLOR || "";
  document.getElementById("pmGaleriaWrap").innerHTML = "";
  _pmGaleriaContador = 0;
  parsearGaleriaAdminString(p.IMAGENES).forEach(f => agregarFilaGaleria(f.color, f.url));
  document.getElementById("pmPublicado").checked = String(p.PUBLICADO || "").toUpperCase() === "SI";
  document.getElementById("pmDestacado").checked = String(p.DESTACADO || "").toUpperCase() === "SI";
  document.getElementById("pmOferta").checked = String(p.OFERTA || "").toUpperCase() === "SI";

  const tieneVentaPorCaja = !!(p.CODIGO_CAJA && String(p.CODIGO_CAJA).trim());
  document.getElementById("pmVentaPorCaja").checked = tieneVentaPorCaja;
  document.getElementById("pmCodigoCaja").value = p.CODIGO_CAJA || "";
  document.getElementById("pmUnidadesPorCaja").value = p.UNIDADES_POR_CAJA || "";
  document.getElementById("pmPrecioCaja").value = p.PRECIO_CAJA || "";
  document.getElementById("pmCajaWrap").style.display = tieneVentaPorCaja ? "" : "none";

  actualizarPreviewImagenProducto();
  poblarCategoriasDatalist();
  document.getElementById("productModalBackdrop").classList.add("show");
}

/** Muestra/oculta los campos de código/unidades/precio de caja */
function toggleVentaPorCaja() {
  const activo = document.getElementById("pmVentaPorCaja").checked;
  document.getElementById("pmCajaWrap").style.display = activo ? "" : "none";
}

function cerrarModalProducto() {
  document.getElementById("productModalBackdrop").classList.remove("show");
}

/* ===================== AGREGAR STOCK (entrada de mercadería) ===================== */

/** Opens the "Agregar Stock" modal pre-filled with the product's current data */
function abrirModalStock(codigo) {
  const p = productosAdminGlobal.find(x => String(x.CODIGO) === String(codigo));
  if (!p) { toast("No se encontró el producto", "error"); return; }

  document.getElementById("smCodigo").value = p.CODIGO;
  document.getElementById("smNombre").textContent = p.PRODUCTO || "—";
  document.getElementById("smCodigoLabel").textContent = p.CODIGO;
  const stockActual = Number(p.STOCK || 0);
  document.getElementById("smStockActual").textContent = stockActual.toLocaleString("es-AR");
  document.getElementById("smStockActual").dataset.valor = stockActual;
  document.getElementById("smCantidad").value = "";

  const imagenUrl = p.IMAGEN ? String(p.IMAGEN).trim() : "";
  document.getElementById("smFoto").innerHTML = imagenUrl
    ? `<img src="${escapeHtml(imagenUrl)}" alt="" onerror="this.parentElement.innerHTML='🛒';">`
    : "🛒";

  const resultadoEl = document.getElementById("smResultado");
  resultadoEl.classList.remove("show");

  document.getElementById("stockModalBackdrop").classList.add("show");
  setTimeout(() => document.getElementById("smCantidad").focus(), 80);
}

function cerrarModalStock() {
  document.getElementById("stockModalBackdrop").classList.remove("show");
}

/** Live preview: shows "stock actual + cantidad = stock nuevo" as the user types */
function actualizarPreviewStock() {
  const actual = Number(document.getElementById("smStockActual").dataset.valor || 0);
  const cantidad = Number(document.getElementById("smCantidad").value || 0);
  const resultadoEl = document.getElementById("smResultado");

  if (!cantidad || cantidad <= 0) {
    resultadoEl.classList.remove("show");
    return;
  }

  const nuevo = actual + cantidad;
  resultadoEl.innerHTML = `<span>Nuevo stock</span><strong>${actual.toLocaleString("es-AR")} + ${cantidad.toLocaleString("es-AR")} = ${nuevo.toLocaleString("es-AR")}</strong>`;
  resultadoEl.classList.add("show");
}

/** Sends the stock addition to the backend (sumarStockProducto) */
async function confirmarAgregarStock() {
  const codigo = document.getElementById("smCodigo").value.trim();
  const cantidad = document.getElementById("smCantidad").value;

  if (!cantidad || Number(cantidad) <= 0) {
    toast("Ingresá una cantidad válida, mayor a 0", "error");
    return;
  }

  const btn = document.getElementById("btnConfirmarStock");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    const params = new URLSearchParams({
      action: "sumarStockProducto",
      codigo: codigo,
      cantidad: cantidad,
      rol: obtenerRolActual()
    });

    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo sumar el stock", "error");
      return;
    }

    toast(`Stock actualizado: ${data.stockAnterior.toLocaleString("es-AR")} → ${data.stockNuevo.toLocaleString("es-AR")}`, "success");
    cerrarModalStock();
    cargarProductos();

    // Refresh in-memory POS catalog too, so the new stock shows up right away
    productosPOS = [];

  } catch (error) {
    console.error("Error al sumar stock:", error);
    toast("Error de conexión al sumar el stock", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

/** Live preview of the image URL pasted into the product modal */
function actualizarPreviewImagenProducto() {
  const url = document.getElementById("pmImagen").value.trim();
  const preview = document.getElementById("pmImagenPreview");
  if (!preview) return;
  if (!url) { preview.innerHTML = "🖼️"; return; }
  preview.innerHTML = `<img src="${escapeHtml(url)}" alt="" onerror="this.parentElement.innerHTML='⚠️';">`;
}

/* ===================== GALERÍA DE COLORES / IMÁGENES ADICIONALES ===================== */
/* Cada producto puede tener, además de la foto principal (pmImagen), otras
   fotos asociadas a un color — se muestran como miniaturas en el Quick View
   del catálogo web. Se guardan todas juntas en la columna IMAGENES de la
   hoja PRODUCTOS, con el formato "url1|color1, url2|color2, ...". */

let _pmGaleriaContador = 0;

/** Arma un string "url|color, url|color, ..." a partir de una lista de {url, color} — inverso de parsearGaleriaAdminString() */
function armarGaleriaAdminString(filas) {
  return filas
    .filter(f => f.url)
    .map(f => f.url + (f.color ? "|" + f.color : ""))
    .join(", ");
}

/** Parsea el string guardado en IMAGENES de vuelta a [{url, color}, ...] */
function parsearGaleriaAdminString(texto) {
  return String(texto || "")
    .split(",")
    .map(entrada => entrada.trim())
    .filter(Boolean)
    .map(entrada => {
      const [url, color] = entrada.split("|");
      return { url: (url || "").trim(), color: (color || "").trim() };
    });
}

/** Agrega una fila vacía (o pre-cargada, en modo edición) a la galería de colores */
function agregarFilaGaleria(color, url) {
  const wrap = document.getElementById("pmGaleriaWrap");
  if (!wrap) return;

  const n = ++_pmGaleriaContador;

  const fila = document.createElement("div");
  fila.className = "pm-galeria-row";
  fila.dataset.fila = n;
  fila.innerHTML = `
    <div class="pm-galeria-preview" id="pmGalPreview_${n}">🖼️</div>
    <div style="flex:1; min-width:0;">
      <div class="row g-2">
        <div class="col-6">
          <input type="text" class="form-control form-control-sm" id="pmGalColor_${n}" placeholder="Color (ej: Rojo)" value="${escapeHtml(color || "")}">
        </div>
        <div class="col-6">
          <input type="file" class="form-control form-control-sm" accept="image/*" onchange="onSeleccionarArchivoImagenGaleria(event, ${n})">
        </div>
      </div>
      <input type="text" class="form-control form-control-sm mt-1" id="pmGalUrl_${n}" placeholder="o pegá una URL de imagen" value="${escapeHtml(url || "")}" oninput="actualizarPreviewGaleria(${n})">
      <div class="pm-image-status" id="pmGalStatus_${n}"></div>
    </div>
    <button type="button" class="btn btn-sm btn-outline-danger pm-galeria-quitar" onclick="quitarFilaGaleria(${n})" title="Quitar este color">✕</button>
  `;

  wrap.appendChild(fila);
  actualizarPreviewGaleria(n);
}

function quitarFilaGaleria(n) {
  const fila = document.querySelector(`.pm-galeria-row[data-fila="${n}"]`);
  if (fila) fila.remove();
}

function actualizarPreviewGaleria(n) {
  const url = (document.getElementById(`pmGalUrl_${n}`) || {}).value || "";
  const preview = document.getElementById(`pmGalPreview_${n}`);
  if (!preview) return;
  const limpia = url.trim();
  if (!limpia) { preview.innerHTML = "🖼️"; return; }
  preview.innerHTML = `<img src="${escapeHtml(limpia)}" alt="" onerror="this.parentElement.innerHTML='⚠️';">`;
}

/** Handles the file picker de una fila de la galería: recorta y sube igual que la foto principal */
function onSeleccionarArchivoImagenGaleria(event, n) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById(`pmGalStatus_${n}`);

  if (!file.type.startsWith("image/")) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "Elegí un archivo de imagen (jpg, png, webp)."; }
    event.target.value = "";
    return;
  }

  const TAMANO_ORIGINAL_MAXIMO_MB = 20;
  if (file.size > TAMANO_ORIGINAL_MAXIMO_MB * 1024 * 1024) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = `⚠️ El archivo pesa demasiado (máx. ${TAMANO_ORIGINAL_MAXIMO_MB}MB). Elegí una foto más liviana.`; }
    event.target.value = "";
    return;
  }

  abrirRecorteImagenProducto(file, event.target, {
    urlInputId: `pmGalUrl_${n}`,
    previewId: `pmGalPreview_${n}`,
    statusId: `pmGalStatus_${n}`
  });
}

/** Junta las filas de la galería que están cargadas en el modal en el string que se manda al backend */
function obtenerGaleriaDesdeFormulario() {
  const filas = Array.from(document.querySelectorAll("#pmGaleriaWrap .pm-galeria-row")).map(fila => {
    const n = fila.dataset.fila;
    const url = (document.getElementById(`pmGalUrl_${n}`) || {}).value || "";
    const color = (document.getElementById(`pmGalColor_${n}`) || {}).value || "";
    return { url: url.trim(), color: color.trim() };
  });
  return armarGaleriaAdminString(filas);
}

/**
 * Redimensiona y comprime una imagen en el navegador antes de subirla,
 * para que una foto de celular de varios MB no tarde una eternidad en
 * subir ni ocupe espacio de más en Drive. Devuelve siempre JPEG
 * (excepto si el original es más chico que el límite, en cuyo caso ni
 * vale la pena recomprimir).
 *
 * Estrategia: redimensiona al lado mayor = LADO_MAXIMO_PX, y si con
 * calidad inicial sigue pesando más que PESO_OBJETIVO_KB, vuelve a
 * comprimir con menos calidad (hasta MAX_INTENTOS veces). Nunca rechaza
 * la imagen — en el peor caso, sube la versión más liviana que logró.
 */
async function comprimirImagenProducto(file) {

  const LADO_MAXIMO_PX   = 900;
  const PESO_OBJETIVO_KB = 700;
  const CALIDADES        = [0.82, 0.7, 0.55, 0.4]; // intentos sucesivos, de mejor a peor calidad

  const imagen = await new Promise((resolve, reject) => {
    const img = new Image();
    img.onload  = () => resolve(img);
    img.onerror = () => reject(new Error("No se pudo leer la imagen seleccionada"));
    img.src = URL.createObjectURL(file);
  });

  let { width, height } = imagen;
  if (width > LADO_MAXIMO_PX || height > LADO_MAXIMO_PX) {
    if (width >= height) {
      height = Math.round(height * (LADO_MAXIMO_PX / width));
      width  = LADO_MAXIMO_PX;
    } else {
      width  = Math.round(width * (LADO_MAXIMO_PX / height));
      height = LADO_MAXIMO_PX;
    }
  }

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff"; // si el original tenía transparencia (ej. PNG), JPEG no la soporta
  ctx.fillRect(0, 0, width, height);
  ctx.drawImage(imagen, 0, 0, width, height);
  URL.revokeObjectURL(imagen.src);

  let ultimoBase64 = null;

  for (const calidad of CALIDADES) {
    const base64 = await new Promise(resolve => {
      canvas.toBlob(blob => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      }, "image/jpeg", calidad);
    });

    ultimoBase64 = base64;
    const pesoKB = Math.round((base64.length * 0.75) / 1024);
    if (pesoKB <= PESO_OBJETIVO_KB) break; // ya entra cómodo, no hace falta seguir bajando calidad
  }

  return { base64: ultimoBase64, tipoMime: "image/jpeg" };
}

/** Handles the file picker: compresses the chosen image, then uploads it to Drive */
async function onSeleccionarArchivoImagenProducto(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById("pmImagenStatus");

  if (!file.type.startsWith("image/")) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "Elegí un archivo de imagen (jpg, png, webp)."; }
    event.target.value = "";
    return;
  }

  // Tope de seguridad sobre el archivo original: una foto de celular normal
  // pesa 2-8MB. Más que esto probablemente sea un error de selección
  // (video, RAW, etc.) — se rechaza antes de gastar tiempo procesándolo.
  const TAMANO_ORIGINAL_MAXIMO_MB = 20;
  if (file.size > TAMANO_ORIGINAL_MAXIMO_MB * 1024 * 1024) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = `⚠️ El archivo pesa demasiado (máx. ${TAMANO_ORIGINAL_MAXIMO_MB}MB). Elegí una foto más liviana.`; }
    event.target.value = "";
    return;
  }

  // En vez de subir directo, abrimos el editor de recorte cuadrado.
  // La subida real sigue en subirImagenProductoRecortada(), una vez confirmado el recorte.
  abrirRecorteImagenProducto(file, event.target);
}

/** Sube (comprime + manda a Drive) el blob YA recortado por el editor de recorte */
async function subirImagenProductoRecortada(blob, inputEl, destino) {
  destino = destino || _DESTINO_IMAGEN_PRINCIPAL;
  const statusEl = document.getElementById(destino.statusId);

  // Local preview inmediata, mientras se comprime y sube en segundo plano
  const localUrl = URL.createObjectURL(blob);
  const preview = document.getElementById(destino.previewId);
  if (preview) preview.innerHTML = `<img src="${localUrl}" alt="">`;

  if (statusEl) { statusEl.className = "pm-image-status uploading"; statusEl.textContent = "⏳ Optimizando imagen..."; }

  try {
    const pesoOriginalKB = Math.round(blob.size / 1024);
    const { base64, tipoMime } = await comprimirImagenProducto(blob);
    const pesoFinalKB = Math.round((base64.length * 0.75) / 1024); // estimación: base64 pesa ~33% más que los bytes reales

    if (statusEl) {
      statusEl.className = "pm-image-status uploading";
      statusEl.textContent = pesoFinalKB < pesoOriginalKB
        ? `⏳ Subiendo a Drive... (${pesoOriginalKB}KB → ${pesoFinalKB}KB)`
        : "⏳ Subiendo a Drive...";
    }

    const codigoProducto = document.getElementById("pmCodigo").value.trim() || document.getElementById("pmCodigoOriginal").value.trim();

    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // evita preflight CORS contra Apps Script
      body: JSON.stringify({
        action: "subirImagenProducto",
        imagenBase64: base64,
        tipoMime: tipoMime,
        codigoProducto: codigoProducto
      })
    }, { timeoutMs: 40000 }); // las imágenes pesan más y tardan más en subir — no es una mutación chica

    const data = await response.json();

    if (!data.success) {
      if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ " + (data.message || "No se pudo subir la imagen."); }
      return;
    }

    const urlInput = document.getElementById(destino.urlInputId);
    if (urlInput) urlInput.value = data.url;
    // Ojo: esto solo confirma que la foto quedó en Drive. El producto
    // recién la usa cuando se toca "Guardar" — se aclara en el mensaje.
    if (statusEl) { statusEl.className = "pm-image-status success"; statusEl.textContent = `✓ Imagen subida (${pesoFinalKB}KB) — tocá Guardar para aplicarla al producto`; }

  } catch (error) {
    console.error("Error al subir imagen de producto:", error);
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ Error de conexión al subir la imagen."; }
  } finally {
    URL.revokeObjectURL(localUrl);
    if (inputEl) inputEl.value = "";
  }
}

/* ===================== RECORTE CUADRADO DE FOTO DE PRODUCTO ===================== */
/* Editor de recorte 1:1 nativo (sin librerías externas): el usuario arrastra la
   imagen y ajusta el zoom dentro de un marco cuadrado fijo; al confirmar, se
   genera un canvas cuadrado recortado que después pasa por comprimirImagenProducto(). */

let _cropState = null; // { file, imgEl, natW, natH, vp, scale, minScale, maxScale, tx, ty, inputEl, destino, dragging, dragStartX, dragStartY, txStart, tyStart }
let _cropListenersListos = false;

// Destino por defecto del recorte: la foto principal del producto.
// Las filas de la galería de colores pasan su propio destino (ver
// agregarFilaGaleria / onSeleccionarArchivoImagenGaleria) para que el
// mismo editor de recorte suba a su URL/preview/estado en vez de
// pisar siempre la foto principal.
const _DESTINO_IMAGEN_PRINCIPAL = { urlInputId: "pmImagen", previewId: "pmImagenPreview", statusId: "pmImagenStatus" };

function abrirRecorteImagenProducto(file, inputEl, destino) {
  destino = destino || _DESTINO_IMAGEN_PRINCIPAL;
  const backdrop = document.getElementById("cropImagenModalBackdrop");
  const imgEl = document.getElementById("cropImgEl");
  if (!backdrop || !imgEl) {
    // Fallback de seguridad: si el modal no está en el HTML, subimos sin recortar.
    subirImagenProductoRecortada(file, inputEl, destino);
    return;
  }

  const url = URL.createObjectURL(file);
  imgEl.src = url;

  imgEl.onload = () => {
    // Mostramos el modal PRIMERO: mientras está oculto (display:none) el viewport
    // mide 0px de ancho, así que si midiéramos antes, la imagen quedaría con escala 0 (invisible).
    backdrop.classList.add("show");

    const viewport = document.getElementById("cropViewport");
    const vp = viewport.clientWidth; // el viewport es cuadrado (mismo ancho y alto por CSS)
    const natW = imgEl.naturalWidth;
    const natH = imgEl.naturalHeight;
    const minScale = Math.max(vp / natW, vp / natH); // "cover": la imagen tapa todo el marco

    _cropState = {
      file, imgEl, natW, natH, vp,
      scale: minScale, minScale, maxScale: minScale * 4,
      tx: (vp - natW * minScale) / 2,
      ty: (vp - natH * minScale) / 2,
      inputEl, destino,
      dragging: false, dragStartX: 0, dragStartY: 0, txStart: 0, tyStart: 0
    };

    const slider = document.getElementById("cropZoomSlider");
    if (slider) slider.value = 0;

    _cropAplicarTransform();
    _cropInicializarListeners();
  };
}

function _cropAplicarTransform() {
  if (!_cropState) return;
  const { imgEl, tx, ty, scale } = _cropState;
  imgEl.style.transform = `translate(${tx}px, ${ty}px) scale(${scale})`;
}

/** Evita que se pueda arrastrar/zoomear el marco fuera de los bordes de la imagen */
function _cropClamp() {
  if (!_cropState) return;
  const s = _cropState;
  s.scale = Math.min(s.maxScale, Math.max(s.minScale, s.scale));
  const minTx = s.vp - s.natW * s.scale;
  const minTy = s.vp - s.natH * s.scale;
  s.tx = Math.min(0, Math.max(minTx, s.tx));
  s.ty = Math.min(0, Math.max(minTy, s.ty));
}

function _cropInicializarListeners() {
  if (_cropListenersListos) return;
  _cropListenersListos = true;

  const viewport = document.getElementById("cropViewport");
  const slider = document.getElementById("cropZoomSlider");
  if (!viewport || !slider) return;

  const iniciarDrag = (clientX, clientY) => {
    if (!_cropState) return;
    _cropState.dragging = true;
    _cropState.dragStartX = clientX;
    _cropState.dragStartY = clientY;
    _cropState.txStart = _cropState.tx;
    _cropState.tyStart = _cropState.ty;
    viewport.classList.add("dragging");
  };
  const moverDrag = (clientX, clientY) => {
    if (!_cropState || !_cropState.dragging) return;
    _cropState.tx = _cropState.txStart + (clientX - _cropState.dragStartX);
    _cropState.ty = _cropState.tyStart + (clientY - _cropState.dragStartY);
    _cropClamp();
    _cropAplicarTransform();
  };
  const terminarDrag = () => {
    if (!_cropState) return;
    _cropState.dragging = false;
    viewport.classList.remove("dragging");
  };

  // Mouse
  viewport.addEventListener("mousedown", e => { e.preventDefault(); iniciarDrag(e.clientX, e.clientY); });
  window.addEventListener("mousemove", e => moverDrag(e.clientX, e.clientY));
  window.addEventListener("mouseup", terminarDrag);

  // Touch (celular/tablet)
  viewport.addEventListener("touchstart", e => {
    const t = e.touches[0];
    iniciarDrag(t.clientX, t.clientY);
  }, { passive: true });
  viewport.addEventListener("touchmove", e => {
    const t = e.touches[0];
    moverDrag(t.clientX, t.clientY);
  }, { passive: true });
  viewport.addEventListener("touchend", terminarDrag);

  // Rueda del mouse también hace zoom
  viewport.addEventListener("wheel", e => {
    e.preventDefault();
    const delta = e.deltaY < 0 ? 1 : -1;
    const nuevoValor = Math.min(100, Math.max(0, Number(slider.value) + delta * 4));
    slider.value = nuevoValor;
    slider.dispatchEvent(new Event("input"));
  }, { passive: false });

  // Slider de zoom (0-100) → escala real, manteniendo el centro del marco fijo
  slider.addEventListener("input", () => {
    if (!_cropState) return;
    const s = _cropState;
    const t = Number(slider.value) / 100; // 0..1
    const nuevaEscala = s.minScale + t * (s.maxScale - s.minScale);

    // Punto de la imagen que hoy está en el centro del marco — lo mantenemos fijo al hacer zoom
    const cx = (s.vp / 2 - s.tx) / s.scale;
    const cy = (s.vp / 2 - s.ty) / s.scale;

    s.scale = nuevaEscala;
    s.tx = s.vp / 2 - cx * s.scale;
    s.ty = s.vp / 2 - cy * s.scale;

    _cropClamp();
    _cropAplicarTransform();
  });
}

function cancelarRecorteImagenProducto() {
  const backdrop = document.getElementById("cropImagenModalBackdrop");
  if (backdrop) backdrop.classList.remove("show");
  if (_cropState) {
    if (_cropState.imgEl && _cropState.imgEl.src) URL.revokeObjectURL(_cropState.imgEl.src);
    if (_cropState.inputEl) _cropState.inputEl.value = "";
  }
  _cropState = null;
}

function confirmarRecorteImagenProducto() {
  if (!_cropState) return;
  const s = _cropState;

  // Rectángulo recortado, en píxeles del archivo original (no de la pantalla)
  const sx = (0 - s.tx) / s.scale;
  const sy = (0 - s.ty) / s.scale;
  const sSize = s.vp / s.scale;

  const LADO_SALIDA_PX = 1000; // resolución del cuadrado recortado; comprimirImagenProducto() la ajusta después
  const canvas = document.createElement("canvas");
  canvas.width = LADO_SALIDA_PX;
  canvas.height = LADO_SALIDA_PX;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(s.imgEl, sx, sy, sSize, sSize, 0, 0, LADO_SALIDA_PX, LADO_SALIDA_PX);

  const backdrop = document.getElementById("cropImagenModalBackdrop");
  const urlOriginal = s.imgEl.src;
  const inputEl = s.inputEl;
  const destino = s.destino;

  canvas.toBlob(blob => {
    if (backdrop) backdrop.classList.remove("show");
    URL.revokeObjectURL(urlOriginal);
    _cropState = null;
    if (blob) subirImagenProductoRecortada(blob, inputEl, destino);
  }, "image/jpeg", 0.92);
}

/* ===================== RECORTE PANORÁMICO DE IMAGEN DEL BANNER ===================== */
/* Mismo editor que el recorte de fotos de producto (abrirRecorteImagenProducto),
   pero generalizado a un marco rectangular en vez de cuadrado: el ratio del
   hero de escritorio real (1400:260, ver .hero.hero--imagen en style.css).
   El recorte que el usuario elige acá es el que después el celular recorta
   un poco más angosto de forma automática (background-size:cover), así que
   actualizarPreviewBanner() muestra debajo cómo queda en cada tamaño. */

const BANNER_RATIO_DESKTOP        = 1400 / 260; // debe coincidir con .hero.hero--imagen en style.css
const BANNER_LADO_SALIDA_ANCHO_PX = 1600;        // resolución del recorte panorámico; comprimirImagenBanner() la ajusta después

let _cropBannerState = null; // misma forma que _cropState, pero con vpW/vpH en vez de un solo vp (marco no es cuadrado)
let _cropBannerListenersListos = false;

/** Handles the banner file picker: abre el editor de recorte antes de subir */
async function onSeleccionarArchivoBanner(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById("cfgBannerImagenStatus");

  if (!file.type.startsWith("image/")) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "Elegí un archivo de imagen (jpg, png, webp)."; }
    event.target.value = "";
    return;
  }

  const TAMANO_ORIGINAL_MAXIMO_MB = 20;
  if (file.size > TAMANO_ORIGINAL_MAXIMO_MB * 1024 * 1024) {
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = `⚠️ El archivo pesa demasiado (máx. ${TAMANO_ORIGINAL_MAXIMO_MB}MB). Elegí una foto más liviana.`; }
    event.target.value = "";
    return;
  }

  abrirRecorteBanner(file, event.target);
}

function abrirRecorteBanner(file, inputEl) {
  const backdrop = document.getElementById("cropBannerModalBackdrop");
  const imgEl = document.getElementById("cropBannerImgEl");
  if (!backdrop || !imgEl) {
    // Fallback de seguridad: si el modal no está en el HTML, subimos sin recortar.
    subirImagenBanner(file, inputEl);
    return;
  }

  const url = URL.createObjectURL(file);
  imgEl.src = url;

  imgEl.onload = () => {
    // Igual que en el recorte de producto: mostramos el modal ANTES de medir,
    // porque oculto (display:none) el viewport mide 0px.
    backdrop.classList.add("show");

    const viewport = document.getElementById("cropBannerViewport");
    const vpW = viewport.clientWidth;
    const vpH = viewport.clientHeight; // fijado por CSS vía aspect-ratio (crop-viewport--banner)
    const natW = imgEl.naturalWidth;
    const natH = imgEl.naturalHeight;
    const minScale = Math.max(vpW / natW, vpH / natH); // "cover": la imagen tapa todo el marco

    _cropBannerState = {
      file, imgEl, natW, natH, vpW, vpH,
      scale: minScale, minScale, maxScale: minScale * 4,
      tx: (vpW - natW * minScale) / 2,
      ty: (vpH - natH * minScale) / 2,
      inputEl,
      dragging: false, dragStartX: 0, dragStartY: 0, txStart: 0, tyStart: 0
    };

    const slider = document.getElementById("cropBannerZoomSlider");
    if (slider) slider.value = 0;

    _cropBannerAplicarTransform();
    _cropBannerInicializarListeners();
  };
}

function _cropBannerAplicarTransform() {
  if (!_cropBannerState) return;
  const { imgEl, tx, ty, scale } = _cropBannerState;
  imgEl.style.transform = `translate(${tx}px, ${ty}px) scale(${scale})`;
}

/** Evita arrastrar/zoomear el marco fuera de los bordes de la imagen */
function _cropBannerClamp() {
  if (!_cropBannerState) return;
  const s = _cropBannerState;
  s.scale = Math.min(s.maxScale, Math.max(s.minScale, s.scale));
  const minTx = s.vpW - s.natW * s.scale;
  const minTy = s.vpH - s.natH * s.scale;
  s.tx = Math.min(0, Math.max(minTx, s.tx));
  s.ty = Math.min(0, Math.max(minTy, s.ty));
}

function _cropBannerInicializarListeners() {
  if (_cropBannerListenersListos) return;
  _cropBannerListenersListos = true;

  const viewport = document.getElementById("cropBannerViewport");
  const slider = document.getElementById("cropBannerZoomSlider");
  if (!viewport || !slider) return;

  const iniciarDrag = (clientX, clientY) => {
    if (!_cropBannerState) return;
    _cropBannerState.dragging = true;
    _cropBannerState.dragStartX = clientX;
    _cropBannerState.dragStartY = clientY;
    _cropBannerState.txStart = _cropBannerState.tx;
    _cropBannerState.tyStart = _cropBannerState.ty;
    viewport.classList.add("dragging");
  };
  const moverDrag = (clientX, clientY) => {
    if (!_cropBannerState || !_cropBannerState.dragging) return;
    _cropBannerState.tx = _cropBannerState.txStart + (clientX - _cropBannerState.dragStartX);
    _cropBannerState.ty = _cropBannerState.tyStart + (clientY - _cropBannerState.dragStartY);
    _cropBannerClamp();
    _cropBannerAplicarTransform();
  };
  const terminarDrag = () => {
    if (!_cropBannerState) return;
    _cropBannerState.dragging = false;
    viewport.classList.remove("dragging");
  };

  // Mouse
  viewport.addEventListener("mousedown", e => { e.preventDefault(); iniciarDrag(e.clientX, e.clientY); });
  window.addEventListener("mousemove", e => moverDrag(e.clientX, e.clientY));
  window.addEventListener("mouseup", terminarDrag);

  // Touch (celular/tablet)
  viewport.addEventListener("touchstart", e => {
    const t = e.touches[0];
    iniciarDrag(t.clientX, t.clientY);
  }, { passive: true });
  viewport.addEventListener("touchmove", e => {
    const t = e.touches[0];
    moverDrag(t.clientX, t.clientY);
  }, { passive: true });
  viewport.addEventListener("touchend", terminarDrag);

  // Rueda del mouse también hace zoom
  viewport.addEventListener("wheel", e => {
    e.preventDefault();
    const delta = e.deltaY < 0 ? 1 : -1;
    const nuevoValor = Math.min(100, Math.max(0, Number(slider.value) + delta * 4));
    slider.value = nuevoValor;
    slider.dispatchEvent(new Event("input"));
  }, { passive: false });

  // Slider de zoom (0-100) → escala real, manteniendo el centro del marco fijo
  slider.addEventListener("input", () => {
    if (!_cropBannerState) return;
    const s = _cropBannerState;
    const t = Number(slider.value) / 100; // 0..1
    const nuevaEscala = s.minScale + t * (s.maxScale - s.minScale);

    const cx = (s.vpW / 2 - s.tx) / s.scale;
    const cy = (s.vpH / 2 - s.ty) / s.scale;

    s.scale = nuevaEscala;
    s.tx = s.vpW / 2 - cx * s.scale;
    s.ty = s.vpH / 2 - cy * s.scale;

    _cropBannerClamp();
    _cropBannerAplicarTransform();
  });
}

function cancelarRecorteBanner() {
  const backdrop = document.getElementById("cropBannerModalBackdrop");
  if (backdrop) backdrop.classList.remove("show");
  if (_cropBannerState) {
    if (_cropBannerState.imgEl && _cropBannerState.imgEl.src) URL.revokeObjectURL(_cropBannerState.imgEl.src);
    if (_cropBannerState.inputEl) _cropBannerState.inputEl.value = "";
  }
  _cropBannerState = null;
}

function confirmarRecorteBanner() {
  if (!_cropBannerState) return;
  const s = _cropBannerState;

  // Rectángulo recortado, en píxeles del archivo original (no de la pantalla)
  const sx = (0 - s.tx) / s.scale;
  const sy = (0 - s.ty) / s.scale;
  const sW = s.vpW / s.scale;
  const sH = s.vpH / s.scale;

  const outW = BANNER_LADO_SALIDA_ANCHO_PX;
  const outH = Math.round(outW / BANNER_RATIO_DESKTOP);
  const canvas = document.createElement("canvas");
  canvas.width = outW;
  canvas.height = outH;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(s.imgEl, sx, sy, sW, sH, 0, 0, outW, outH);

  const backdrop = document.getElementById("cropBannerModalBackdrop");
  const urlOriginal = s.imgEl.src;
  const inputEl = s.inputEl;

  canvas.toBlob(blob => {
    if (backdrop) backdrop.classList.remove("show");
    URL.revokeObjectURL(urlOriginal);
    _cropBannerState = null;
    if (blob) subirImagenBanner(blob, inputEl);
  }, "image/jpeg", 0.92);
}

/**
 * Comprime la imagen del banner antes de subirla — misma idea que
 * comprimirImagenProducto(), pero con un lado máximo más grande, porque
 * el banner ocupa mucho más espacio en pantalla que una foto de producto
 * (todo el ancho del catálogo) y se ve pixelado si se lo achica de más.
 */
async function comprimirImagenBanner(file) {
  const LADO_MAXIMO_PX   = 1600;
  const PESO_OBJETIVO_KB = 900;
  const CALIDADES        = [0.82, 0.7, 0.55, 0.4];

  const imagen = await new Promise((resolve, reject) => {
    const img = new Image();
    img.onload  = () => resolve(img);
    img.onerror = () => reject(new Error("No se pudo leer la imagen seleccionada"));
    img.src = URL.createObjectURL(file);
  });

  let { width, height } = imagen;
  if (width > LADO_MAXIMO_PX) {
    height = Math.round(height * (LADO_MAXIMO_PX / width));
    width  = LADO_MAXIMO_PX;
  }

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.drawImage(imagen, 0, 0, width, height);
  URL.revokeObjectURL(imagen.src);

  let ultimoBase64 = null;

  for (const calidad of CALIDADES) {
    const base64 = await new Promise(resolve => {
      canvas.toBlob(blob => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      }, "image/jpeg", calidad);
    });

    ultimoBase64 = base64;
    const pesoKB = Math.round((base64.length * 0.75) / 1024);
    if (pesoKB <= PESO_OBJETIVO_KB) break;
  }

  return { base64: ultimoBase64, tipoMime: "image/jpeg" };
}

/** Sube (comprime + manda a Drive) el blob ya recortado del banner */
async function subirImagenBanner(blob, inputEl) {
  const statusEl = document.getElementById("cfgBannerImagenStatus");

  // Preview local inmediata (miniatura + las dos vistas escritorio/celular),
  // mientras se comprime y sube en segundo plano.
  const localUrl = URL.createObjectURL(blob);
  const preview = document.getElementById("cfgBannerImagenPreview");
  if (preview) preview.innerHTML = `<img src="${localUrl}" alt="">`;
  const mockD = document.getElementById("bannerMockDesktop");
  const mockM = document.getElementById("bannerMockMobile");
  if (mockD) mockD.style.backgroundImage = `url("${localUrl}")`;
  if (mockM) mockM.style.backgroundImage = `url("${localUrl}")`;

  if (statusEl) { statusEl.className = "pm-image-status uploading"; statusEl.textContent = "⏳ Optimizando imagen..."; }

  try {
    const pesoOriginalKB = Math.round(blob.size / 1024);
    const { base64, tipoMime } = await comprimirImagenBanner(blob);
    const pesoFinalKB = Math.round((base64.length * 0.75) / 1024);

    if (statusEl) {
      statusEl.className = "pm-image-status uploading";
      statusEl.textContent = pesoFinalKB < pesoOriginalKB
        ? `⏳ Subiendo a Drive... (${pesoOriginalKB}KB → ${pesoFinalKB}KB)`
        : "⏳ Subiendo a Drive...";
    }

    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // evita preflight CORS contra Apps Script
      body: JSON.stringify({
        action: "subirImagenProducto",
        imagenBase64: base64,
        tipoMime: tipoMime,
        codigoProducto: "BANNER_CATALOGO" // mismo endpoint que las fotos de producto, con un código fijo para el banner
      })
    }, { timeoutMs: 40000 });

    const data = await response.json();

    if (!data.success) {
      if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ " + (data.message || "No se pudo subir la imagen."); }
      return;
    }

    document.getElementById("cfgBannerImagen").value = data.url;
    if (statusEl) { statusEl.className = "pm-image-status success"; statusEl.textContent = `✓ Imagen subida (${pesoFinalKB}KB) — no te olvides de guardar`; }
    actualizarPreviewBanner();

  } catch (error) {
    console.error("Error al subir imagen del banner:", error);
    if (statusEl) { statusEl.className = "pm-image-status error"; statusEl.textContent = "⚠️ Error de conexión al subir la imagen."; }
  } finally {
    URL.revokeObjectURL(localUrl);
    if (inputEl) inputEl.value = "";
  }
}

/**
 * Actualiza la miniatura y las dos vistas previas (escritorio/celular) del
 * banner, a partir de la URL cargada en cfgBannerImagen y los textos de
 * título/subtítulo actuales del formulario. Se llama al escribir, al pegar
 * una URL a mano, al terminar de subir una imagen, y al abrir "Apariencia".
 */
function actualizarPreviewBanner() {
  const urlInput = document.getElementById("cfgBannerImagen");
  if (!urlInput) return;
  const url = urlInput.value.trim();

  const preview = document.getElementById("cfgBannerImagenPreview");
  if (preview) preview.innerHTML = url ? `<img src="${escapeHtml(url)}" alt="" onerror="this.parentElement.innerHTML='⚠️';">` : "🖼️";

  const bgCss = url ? `url("${url.replace(/"/g, '\\"')}")` : "none";
  const mockD = document.getElementById("bannerMockDesktop");
  const mockM = document.getElementById("bannerMockMobile");
  if (mockD) mockD.style.backgroundImage = bgCss;
  if (mockM) mockM.style.backgroundImage = bgCss;

  const tituloEl = document.getElementById("cfgBannerTitulo");
  const subtituloEl = document.getElementById("cfgBannerSubtitulo");
  const titulo    = (tituloEl && tituloEl.value.trim())    || "Tu negocio";
  const subtitulo = (subtituloEl && subtituloEl.value.trim()) || "Catálogo Mayorista Online";
  ["bannerMockTituloD", "bannerMockTituloM"].forEach(id => { const el = document.getElementById(id); if (el) el.textContent = titulo; });
  ["bannerMockSubtituloD", "bannerMockSubtituloM"].forEach(id => { const el = document.getElementById(id); if (el) el.textContent = subtitulo; });
}

/** Fills the category <datalist> with the distinct categories already in use */
function poblarCategoriasDatalist() {
  const datalist = document.getElementById("pmCategoriasList");
  if (!datalist) return;
  const categorias = [...new Set(
    productosAdminGlobal.map(p => String(p.CATEGORIA || "").trim()).filter(Boolean)
  )].sort();
  datalist.innerHTML = categorias.map(c => `<option value="${escapeHtml(c)}"></option>`).join("");
}

/** Saves the product form — creates a new product or updates an existing one */
/* ===================== CÓDIGOS ÚNICOS DE PRODUCTO =====================
   Mismas reglas que el backend (conflictoCodigoProducto en code.gs): un
   código —de producto o de caja— identifica a UN solo producto; se compara
   sin espacios y sin distinguir mayúsculas. Esto avisa al instante (y sin
   internet en la app de escritorio); el backend lo vuelve a validar igual. */

function normalizarCodigoProducto(c) {
  return String(c === undefined || c === null ? "" : c).trim().toUpperCase();
}

/**
 * Devuelve el mensaje de error si `codigo` o `codigoCaja` ya los usa OTRO
 * producto, o null si están libres. `codigoOriginal` = producto que se
 * está editando (vacío si es nuevo). Solo se revisa lo que cambió, igual
 * que el backend. `listaExtra`: productos de otra fuente (caché local).
 */
function conflictoCodigoProductoLocal(codigo, codigoCaja, codigoOriginal, listaExtra) {
  const listas = [
    typeof productosAdminGlobal !== "undefined" ? productosAdminGlobal : [],
    typeof productosPOS !== "undefined" ? productosPOS : [],
    listaExtra || []
  ];
  const propio = normalizarCodigoProducto(codigoOriginal);
  const otros = new Map();
  let actual = null;
  listas.forEach(lista => (lista || []).forEach(p => {
    const c = normalizarCodigoProducto(p && p.CODIGO);
    if (!c) return;
    if (propio && c === propio) { if (!actual) actual = p; return; }
    if (!otros.has(c)) otros.set(c, p);
  }));

  const n = normalizarCodigoProducto(codigo);
  const nCaja = normalizarCodigoProducto(codigoCaja);
  const revisar = (valor, esCaja) => {
    const v = normalizarCodigoProducto(valor);
    for (const p of otros.values()) {
      const nombre = p.PRODUCTO ? `"${p.PRODUCTO}"` : "otro producto";
      if (normalizarCodigoProducto(p.CODIGO) === v) return `${esCaja ? "El código de caja" : "El código"} "${String(valor).trim()}" ya pertenece a ${nombre}.`;
      if (normalizarCodigoProducto(p.CODIGO_CAJA) === v) return `${esCaja ? "El código de caja" : "El código"} "${String(valor).trim()}" ya es el código de caja de ${nombre}.`;
    }
    return null;
  };

  if (n && n !== propio) { const m = revisar(codigo, false); if (m) return m; }
  const cajaActual = normalizarCodigoProducto(actual && actual.CODIGO_CAJA);
  if (nCaja && nCaja !== cajaActual) { const m = revisar(codigoCaja, true); if (m) return m; }
  if (nCaja && nCaja === (n || propio)) return "El código de caja tiene que ser distinto del código del producto.";
  return null;
}
window.conflictoCodigoProductoLocal = conflictoCodigoProductoLocal;

/** Aviso al salir del campo código / código de caja del formulario de producto. */
function avisarCodigoDuplicadoFormulario() {
  const campoCodigo = document.getElementById("pmCodigo");
  const campoCaja = document.getElementById("pmCodigoCaja");
  const ventaPorCaja = document.getElementById("pmVentaPorCaja");
  const msg = conflictoCodigoProductoLocal(
    campoCodigo ? campoCodigo.value : "",
    ventaPorCaja && ventaPorCaja.checked && campoCaja ? campoCaja.value : "",
    (document.getElementById("pmCodigoOriginal") || {}).value || ""
  );
  [campoCodigo, campoCaja].forEach(c => c && c.classList.remove("is-invalid"));
  if (msg) {
    toast(msg, "error");
    const culpable = /de caja/.test(msg.split(" ya ")[0]) || /distinto/.test(msg) ? campoCaja : campoCodigo;
    if (culpable) culpable.classList.add("is-invalid");
  }
  return msg;
}

async function guardarProductoForm() {
  const codigoOriginal = document.getElementById("pmCodigoOriginal").value.trim();
  const codigo   = document.getElementById("pmCodigo").value.trim();
  const nombre   = document.getElementById("pmNombre").value.trim();
  const categoria = document.getElementById("pmCategoria").value.trim();
  const alias     = document.getElementById("pmAlias").value.trim();
  const precio   = document.getElementById("pmPrecio").value;
  const stock    = document.getElementById("pmStock").value;
  const imagen   = document.getElementById("pmImagen").value.trim();
  const colorPrimario = document.getElementById("pmColorPrimario").value.trim();
  const imagenes = obtenerGaleriaDesdeFormulario();
  const publicado = document.getElementById("pmPublicado").checked ? "SI" : "NO";
  const destacado  = document.getElementById("pmDestacado").checked ? "SI" : "NO";
  const oferta     = document.getElementById("pmOferta").checked ? "SI" : "NO";

  const ventaPorCaja = document.getElementById("pmVentaPorCaja").checked;
  const codigoCaja = ventaPorCaja ? document.getElementById("pmCodigoCaja").value.trim() : "";
  const unidadesPorCaja = ventaPorCaja ? document.getElementById("pmUnidadesPorCaja").value : "";
  const precioCaja = ventaPorCaja ? document.getElementById("pmPrecioCaja").value : "";

  if (!nombre) { toast("El nombre del producto es obligatorio", "error"); return; }
  if (precio === "" || Number(precio) < 0) { toast("Ingresá un precio válido", "error"); return; }
  if (ventaPorCaja && (!codigoCaja || !unidadesPorCaja || Number(unidadesPorCaja) < 1 || precioCaja === "")) {
    toast("Completá código, unidades y precio de la caja (o desactivá la venta por caja)", "error");
    return;
  }

  // Códigos únicos (aviso inmediato; el backend lo vuelve a validar)
  if (avisarCodigoDuplicadoFormulario()) return;

  const esEdicion = !!codigoOriginal;
  const btn = document.getElementById("btnGuardarProducto");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    const params = new URLSearchParams({
      action: esEdicion ? "actualizarProducto" : "guardarProducto",
      rol: obtenerRolActual(),
      CODIGO: codigo,
      PRODUCTO: nombre,
      CATEGORIA: categoria,
      ALIAS: alias,
      PRECIO: precio || 0,
      STOCK: stock || 0,
      IMAGEN: imagen,
      COLOR: colorPrimario,
      IMAGENES: imagenes,
      PUBLICADO: publicado,
      DESTACADO: destacado,
      OFERTA: oferta,
      CODIGO_CAJA: codigoCaja,
      UNIDADES_POR_CAJA: unidadesPorCaja || 0,
      PRECIO_CAJA: precioCaja || 0
    });
    if (esEdicion) params.set("codigoOriginal", codigoOriginal);
    // Alícuota de IVA: solo si el campo está a la vista (facturación de RI);
    // si no, no se manda y el backend conserva la que tenía.
    if (_alicuotaIvaFormulario() !== null) params.set("ALICUOTA_IVA", _alicuotaIvaFormulario());

    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar el producto", "error");
      return;
    }

    toast(esEdicion ? "Producto actualizado" : "Producto creado", "success");
    cerrarModalProducto();
    cargarProductos();

    // Refresh in-memory POS catalog too, so the new/edited product shows up right away
    productosPOS = [];

  } catch (error) {
    console.error("Error al guardar producto:", error);
    toast("Error de conexión al guardar el producto", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

/** Deletes a product after confirmation */
/**
 * Pide confirmación con modal propio antes de eliminar un producto.
 * IMPORTANTE: antes esto usaba window.confirm(), pero confirm()/alert()
 * son diálogos nativos síncronos que, en Electron, pueden dejar el
 * foco de teclado roto en la ventana después de cerrarse — el input
 * de texto deja de responder hasta recargar. Por eso en todo el panel
 * se usa este patrón de modal propio en vez de confirm() nativo.
 */
function eliminarProducto(codigo) {
  window._codigoProductoParaEliminar = codigo;

  const backdrop = document.getElementById("modalEliminarProdBackdrop");
  const texto = document.getElementById("modalEliminarProdTexto");
  if (texto) texto.textContent = `¿Mover el producto "${codigo}" a la papelera? Lo vas a poder restaurar desde Productos → Papelera.`;
  if (backdrop) backdrop.classList.add("show");
}

function cerrarModalEliminarProd() {
  const backdrop = document.getElementById("modalEliminarProdBackdrop");
  if (backdrop) backdrop.classList.remove("show");
  window._codigoProductoParaEliminar = null;
}

async function eliminarProductoForm() {
  const codigo = window._codigoProductoParaEliminar;
  cerrarModalEliminarProd();
  if (!codigo) return;

  try {
    const usuario = sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("usuarioLogueado") || "";
    const params = new URLSearchParams({ action: "eliminarProducto", codigo, usuario });
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo eliminar el producto", "error");
      return;
    }

    toast("Producto enviado a la papelera", "success");
    cargarProductos();
    productosPOS = [];

  } catch (error) {
    console.error("Error al eliminar producto:", error);
    toast("Error de conexión al eliminar el producto", "error");
  }
}

/* ===================== PAPELERA DE PRODUCTOS =====================
   Los productos borrados no se pierden: el backend (PapeleraYClientes.gs)
   los mueve a la hoja PAPELERA_PRODUCTOS con todos sus datos. Desde acá
   se pueden restaurar, borrar definitivo o vaciar la papelera.
   En la app de escritorio, un borrado hecho SIN internet aparece en la
   papelera recién cuando se sincroniza con Sheets. */

let _papeleraProductos = [];

async function _postPapelera(cuerpo) {
  const response = await fetchAPI(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(cuerpo)
  });
  const data = await _leerRespuestaJSON(response);
  if (data && data.success === false && /no v[aá]lida|desconocida|no reconocida|invalid/i.test(String(data.message || ""))) {
    data.message = "El backend todavía no tiene la papelera: agregá el archivo PapeleraYClientes.gs en Apps Script y publicá una nueva versión (ver LEEME).";
  }
  return data;
}

function abrirPapeleraProductos() {
  const backdrop = document.getElementById("papeleraProductosModalBackdrop");
  if (!backdrop) return;
  const buscador = document.getElementById("papeleraBuscar");
  if (buscador) buscador.value = "";
  backdrop.classList.add("show");
  cargarPapeleraProductos();
}

function cerrarPapeleraProductos() {
  const backdrop = document.getElementById("papeleraProductosModalBackdrop");
  if (backdrop) backdrop.classList.remove("show");
}

async function cargarPapeleraProductos() {
  const cont = document.getElementById("papeleraProductosLista");
  if (!cont) return;
  cont.innerHTML = `<div class="text-center text-muted py-4">Cargando papelera...</div>`;

  try {
    const data = await _postPapelera({ action: "papeleraProductos" });
    if (!data || !data.success) {
      cont.innerHTML = `<div class="text-center text-danger py-4" style="font-size:13px;">${escapeHtml((data && data.message) || "No se pudo cargar la papelera")}</div>`;
      return;
    }
    _papeleraProductos = data.productos || [];
    renderPapeleraProductos();
  } catch (error) {
    console.error("Error al cargar la papelera:", error);
    cont.innerHTML = `<div class="text-center text-danger py-4" style="font-size:13px;">Error de conexión al cargar la papelera. Revisá tu internet.</div>`;
  }
}

function renderPapeleraProductos() {
  const cont = document.getElementById("papeleraProductosLista");
  const contador = document.getElementById("papeleraContador");
  const btnVaciar = document.getElementById("btnVaciarPapelera");
  if (!cont) return;

  const esVendedor = obtenerRolActual() === "vendedor";
  if (btnVaciar) btnVaciar.style.display = (_papeleraProductos.length > 0 && !esVendedor) ? "inline-flex" : "none";
  if (contador) contador.textContent = _papeleraProductos.length === 1 ? "1 producto" : `${_papeleraProductos.length} productos`;

  const filtro = normalizarBusquedaPOS((document.getElementById("papeleraBuscar") || {}).value || "");
  const lista = filtro
    ? _papeleraProductos.filter(p => normalizarBusquedaPOS(p.PRODUCTO + " " + p.CODIGO + " " + p.CATEGORIA).includes(filtro))
    : _papeleraProductos;

  if (_papeleraProductos.length === 0) {
    cont.innerHTML = `<div class="text-center text-muted py-5" style="font-size:13.5px;">La papelera está vacía</div>`;
    return;
  }
  if (lista.length === 0) {
    cont.innerHTML = `<div class="text-center text-muted py-4" style="font-size:13px;">Ningún producto coincide con la búsqueda</div>`;
    return;
  }

  cont.innerHTML = lista.map(p => `
    <div class="d-flex align-items-center gap-2" style="padding:9px 4px; border-bottom:1px solid var(--slate-200);">
      ${p.IMAGEN && !String(p.IMAGEN).startsWith("data:") ? `<img src="${escapeHtml(p.IMAGEN)}" alt="" style="width:40px;height:40px;object-fit:cover;border-radius:6px;flex-shrink:0;" onerror="this.style.display='none'">` : `<div style="width:40px;height:40px;border-radius:6px;background:var(--slate-100);display:flex;align-items:center;justify-content:center;flex-shrink:0;">📦</div>`}
      <div style="flex:1; min-width:0;">
        <div style="font-weight:600; font-size:13.5px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${escapeHtml(p.PRODUCTO || "(sin nombre)")}</div>
        <div style="font-size:11.5px; color:var(--slate-500);">
          <span style="font-family:var(--font-mono);">${escapeHtml(p.CODIGO)}</span>
          ${p.CATEGORIA ? " · " + escapeHtml(p.CATEGORIA) : ""}
          · $${Number(p.PRECIO || 0).toLocaleString("es-AR")} · stock ${Number(p.STOCK || 0)}
        </div>
        <div style="font-size:11px; color:var(--slate-400);">Eliminado ${escapeHtml(p.FECHA_ELIMINADO || "")}${p.USUARIO ? " por " + escapeHtml(p.USUARIO) : ""}</div>
      </div>
      <button type="button" class="btn btn-outline-success btn-sm" style="white-space:nowrap;" onclick="restaurarProductoPapelera('${escapeJsAttr(p.PAPELERA_ID)}', this)">Restaurar</button>
      ${esVendedor ? "" : `<button type="button" class="btn btn-outline-danger btn-sm" title="Borrar definitivamente" onclick="eliminarDefinitivoPapelera('${escapeJsAttr(p.PAPELERA_ID)}', '${escapeJsAttr(p.PRODUCTO || p.CODIGO)}')">✕</button>`}
    </div>`).join("");
}

async function restaurarProductoPapelera(papeleraId, btn) {
  if (btn) { btn.disabled = true; btn.textContent = "Restaurando..."; }
  try {
    const data = await _postPapelera({ action: "restaurarProductoPapelera", papeleraId });
    if (!data || !data.success) {
      toast((data && data.message) || "No se pudo restaurar el producto", "error");
      if (btn) { btn.disabled = false; btn.textContent = "Restaurar"; }
      return;
    }
    toast(`Producto "${data.producto || data.codigo}" restaurado`, "success");
    _papeleraProductos = _papeleraProductos.filter(p => p.PAPELERA_ID !== papeleraId);
    renderPapeleraProductos();
    await _refrescarCatalogoTrasPapelera();
  } catch (error) {
    console.error("Error al restaurar producto:", error);
    toast("Error de conexión al restaurar el producto", "error");
    if (btn) { btn.disabled = false; btn.textContent = "Restaurar"; }
  }
}

/** Después de restaurar: que Productos y el POS vuelvan a ver el producto */
async function _refrescarCatalogoTrasPapelera() {
  try { invalidarCache("productosAdmin", "productos"); } catch (e) {}
  productosPOS = [];
  try {
    if (typeof window.olvidarVersionCatalogoLocal === "function") await window.olvidarVersionCatalogoLocal();
    if (typeof actualizarCatalogoProductosManual === "function") await actualizarCatalogoProductosManual();
    else if (typeof cargarProductos === "function") await cargarProductos();
  } catch (e) {
    console.warn("No se pudo refrescar el catálogo después de restaurar:", e);
  }
}

function eliminarDefinitivoPapelera(papeleraId, nombre) {
  confirmarAccion(
    `¿Borrar "${nombre}" DEFINITIVAMENTE? Ya no se va a poder recuperar.`,
    async () => {
      try {
        const data = await _postPapelera({ action: "eliminarDefinitivoPapelera", papeleraId, rol: obtenerRolActual() });
        if (!data || !data.success) { toast((data && data.message) || "No se pudo borrar", "error"); return; }
        _papeleraProductos = _papeleraProductos.filter(p => p.PAPELERA_ID !== papeleraId);
        renderPapeleraProductos();
        toast("Producto borrado definitivamente", "success");
      } catch (error) {
        console.error("Error al borrar de la papelera:", error);
        toast("Error de conexión", "error");
      }
    },
    "Borrar definitivamente"
  );
}

function vaciarPapeleraProductos() {
  if (_papeleraProductos.length === 0) return;
  confirmarAccion(
    `¿Vaciar la papelera? Se van a borrar DEFINITIVAMENTE ${_papeleraProductos.length} producto(s) y no se van a poder recuperar.`,
    async () => {
      try {
        const data = await _postPapelera({ action: "vaciarPapeleraProductos", rol: obtenerRolActual() });
        if (!data || !data.success) { toast((data && data.message) || "No se pudo vaciar la papelera", "error"); return; }
        _papeleraProductos = [];
        renderPapeleraProductos();
        toast("Papelera vaciada", "success");
      } catch (error) {
        console.error("Error al vaciar la papelera:", error);
        toast("Error de conexión", "error");
      }
    },
    "Vaciar papelera"
  );
}

/* ===================== CLIENTES ===================== */

let clientesGlobal = [];
let _avisoClientesCaido = false; // evita repetir el toast de error de conexión en cada ciclo del polling

/** Refresca manualmente la tabla de clientes desde el backend, con feedback visual en el botón */
async function actualizarClientesForm() {
  const btn = document.getElementById("btnActualizarClientes");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Actualizando...";

  try {
    invalidarCache("clientes"); // forzar datos frescos, no la copia cacheada
    await cargarClientesDesdeBackend();
    toast("Clientes actualizados", "success");
  } catch (error) {
    console.error("Error al actualizar clientes:", error);
    toast("Error de conexión al actualizar", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

async function cargarClientesDesdeBackend() {
  const cached = cacheGet("clientes");
  if (cached) {
    // Igual que en pedidos: la copia guardada solo si la lista está vacía
    if (!clientesGlobal || clientesGlobal.length === 0) {
      clientesGlobal = cached.data;
      filtrarClientes();
    }
    if (!cached.stale) return;
  }
  try {
    const response = await fetchAPI(API_URL + "?action=clientesConCredito");
    const data = await response.json();
    if (!data.clientes) return;
    // Redibujar solo si algo cambió (antes se redibujaba siempre → salto)
    const cambio = JSON.stringify(clientesGlobal || []) !== JSON.stringify(data.clientes);
    clientesGlobal = data.clientes;
    cacheSet("clientes", clientesGlobal);
    _avisoClientesCaido = false; // se pudo actualizar bien — resetea el aviso para el próximo corte
    if (cambio) redibujarSinSalto("tablaClientes", filtrarClientes);
  } catch (error) {
    console.error("Error clientes:", error);
    if (!_avisoClientesCaido) {
      toast("No se pudieron actualizar los clientes — revisá tu conexión", "error");
      _avisoClientesCaido = true;
    }
  }
}

// Cache separado para el selector de pedidos (incluye dirección y todos los campos)
let _clientesPedidoCache = [];
async function cargarClientesParaPedido() {
  if (_clientesPedidoCache.length > 0) return _clientesPedidoCache;
  try {
    const response = await fetchAPI(API_URL + "?action=todosClientes");
    const data = await response.json();
    _clientesPedidoCache = data.clientes || [];
    return _clientesPedidoCache;
  } catch (e) {
    return [];
  }
}

let _renderGenClientes = 0; // mismo guard que en Productos/Pedidos: evita que un render viejo escriba encima de uno nuevo

function renderTablaClientes(lista) {
  const cont = document.getElementById("tablaClientes");
  if (!cont) return;

  const miGen = ++_renderGenClientes;

  if (!lista || lista.length === 0) {
    cont.innerHTML = `<div class="text-center text-muted py-5" style="font-size:14px;">No se encontraron clientes</div>`;
    return;
  }

  // Render en chunks para no congelar con listas grandes
  const CHUNK_CLI = 25;
  let idxCli = 0;
  cont.innerHTML = "";
  const renderChunkCli = () => {
    if (miGen !== _renderGenClientes) return; // superado por un render más nuevo
    const frag = document.createDocumentFragment();
    const fin = Math.min(idxCli + CHUNK_CLI, lista.length);
    for (; idxCli < fin; idxCli++) {
      const c = lista[idxCli];
    const esCredito = String(c.A_CREDITO || "").toUpperCase() === "SI";
    const saldoArs = Number(c.SALDO_PENDIENTE_ARS || 0);
    const saldoUsd = Number(c.SALDO_PENDIENTE_USD || 0);
    const tieneSaldo = esCredito && (saldoArs > 0 || saldoUsd > 0);

    // Borde color: rojo si tiene deuda, verde si es crédito sin deuda, azul por defecto
    const bordeClase = tieneSaldo ? "cliente-card-deuda" : (esCredito ? "cliente-card-credito" : "");

    // Saldo
    let saldoHtml = "";
    if (esCredito) {
      const partes = [];
      if (saldoArs > 0) partes.push(`<span style="color:var(--red-500);font-weight:700;">$${saldoArs.toLocaleString("es-AR")}</span>`);
      if (saldoUsd > 0) partes.push(`<span style="color:var(--red-500);font-weight:700;">US$${saldoUsd.toLocaleString("es-AR")}</span>`);
      saldoHtml = partes.length > 0
        ? `<span class="cliente-dato-label">Saldo</span> ${partes.join(" · ")}`
        : `<span class="cliente-dato-label">Saldo</span> <span class="text-muted" style="font-size:12px;">Sin deuda</span>`;
    }

    const botones = c.CLIENTE_ID
      ? `<button class="btn btn-outline-secondary btn-sm" onclick="abrirModalDetalleCliente('${escapeJsAttr(c.CLIENTE_ID)}')">Ver cuenta</button>
         <button class="btn btn-warning btn-sm" onclick="abrirModalDeudaExtraDirecto('${escapeJsAttr(c.CLIENTE_ID)}', '${escapeJsAttr(c.NOMBRE || c.CLIENTE)}')">+ Deuda</button>
         <button class="btn btn-primary btn-sm" onclick="abrirModalEditarCliente('${escapeJsAttr(c.CLIENTE_ID)}')">Editar</button>
         <button class="btn btn-danger btn-sm" onclick="eliminarClienteForm('${escapeJsAttr(c.CLIENTE_ID)}', '${escapeJsAttr(c.NOMBRE)}')">Eliminar</button>`
      : `<button class="btn btn-outline-success btn-sm" onclick="marcarClienteDesdeHistorialACredito('${escapeJsAttr(c.DNI)}')">Marcar a crédito</button>`;

      const cardHtml = `
    <div class="pedido-card ${bordeClase}">
      <div class="pedido-card-top">
        <div style="flex:1; min-width:0;">
          <div class="pedido-card-id">${escapeHtml(c.DNI || "")}</div>
          <div class="pedido-card-cliente">
            ${escapeHtml(c.NOMBRE || c.CLIENTE)}
            ${c.ALIAS ? `<span style="font-weight:500;color:var(--slate-500);font-size:13px;">· ${escapeHtml(c.ALIAS)}</span>` : ""}
          </div>
          <div class="pedido-card-dir" style="margin-top:4px; display:flex; flex-wrap:wrap; gap:10px;">
            ${c.EMPRESA   ? `<span>🚚 ${escapeHtml(c.EMPRESA)}</span>` : ""}
            ${c.TELEFONO  ? `<span>📞 ${escapeHtml(c.TELEFONO)}</span>` : ""}
            ${c.DIRECCION ? `<span>📍 ${escapeHtml(c.DIRECCION)}</span>` : ""}
          </div>
        </div>
        <div class="text-end" style="flex-shrink:0;">
          <div class="pedido-card-total" style="font-size:15px;">$${Number(c.TOTAL_COMPRADO || c.TOTAL || 0).toLocaleString("es-AR")}</div>
          <div style="font-size:11px;color:var(--slate-500);margin-top:1px;">${c.PEDIDOS || 0} pedido${c.PEDIDOS !== 1 ? "s" : ""}</div>
          <div class="mt-1">
            ${esCredito
              ? `<span class="pedido-estado-badge preparando">A crédito</span>`
              : `<span class="pedido-estado-badge" style="background:var(--slate-100);color:var(--slate-500);">Contado</span>`}
          </div>
          ${tieneSaldo ? `<div class="mt-1" style="font-size:12px;">${saldoHtml}</div>` : ""}
        </div>
      </div>
      <div class="pedido-card-controls">
        ${botones}
      </div>
    </div>`;
      const div = document.createElement("div");
      div.innerHTML = cardHtml;
      frag.appendChild(div.firstElementChild);
    }
    cont.appendChild(frag);
    if (idxCli < lista.length) requestAnimationFrame(renderChunkCli);
  };
  requestAnimationFrame(renderChunkCli);
}

/** Filters the already-loaded client list by name, alias, empresa, or DNI — no new backend call. Also applies the "solo crédito" checkbox. */
function filtrarClientes() {
  const input = document.getElementById("buscarCliente");
  const soloCredito = document.getElementById("filtroSoloCredito");
  const termino = (input ? input.value : "").toLowerCase().trim();

  let filtrados = clientesGlobal;

  if (termino) {
    filtrados = filtrados.filter(c => {
      const nombre = String(c.NOMBRE || c.CLIENTE || "").toLowerCase();
      const alias = String(c.ALIAS || "").toLowerCase();
      const empresa = String(c.EMPRESA || "").toLowerCase();
      const dni = String(c.DNI || "").toLowerCase();
      return nombre.includes(termino) || alias.includes(termino) || empresa.includes(termino) || dni.includes(termino);
    });
  }

  if (soloCredito && soloCredito.checked) {
    filtrados = filtrados.filter(c => String(c.A_CREDITO || "").toUpperCase() === "SI");
  }

  renderTablaClientes(filtrados);
}

/**
 * Wrapper público: mantiene el nombre que ya usan el menú y el timer de
 * fondo, pero evita pedir de nuevo al backend si se llamó hace menos de
 * VENCIMIENTO_CACHE_MS (ej: el cajero entra y sale de Clientes seguido).
 */
function cargarClientes() {
  cargarSiVencido("clientes", cargarClientesDesdeBackend);
}

/* ===================== CLIENTES — ALTA Y CRÉDITO ===================== */

function abrirModalNuevoCliente() {
  document.getElementById("clClienteIdEditando").value = "";
  document.getElementById("clienteModalTitulo").textContent = "+ Nuevo Cliente";
  document.getElementById("clACreditoWrap").style.display = "block";
  document.getElementById("clNombre").value = "";
  document.getElementById("clAlias").value = "";
  document.getElementById("clDni").value = "";
  document.getElementById("clTelefono").value = "";
  document.getElementById("clEmpresa").value = "";
  document.getElementById("clDireccion").value = "";
  ["clLocalidad", "clProvincia", "clEmail", "clNotas"].forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
  document.getElementById("clACredito").checked = false;
  document.getElementById("clienteModalBackdrop").classList.add("show");
}

/** Opens the same client modal, pre-filled for editing — now also shows the "a crédito" checkbox, pre-checked with the client's current value, so it can be toggled directly from here */
function abrirModalEditarCliente(clienteId) {
  const cliente = clientesGlobal.find(c => String(c.CLIENTE_ID) === String(clienteId));
  if (!cliente) { toast("No se encontró el cliente", "error"); return; }

  document.getElementById("clClienteIdEditando").value = clienteId;
  document.getElementById("clienteModalTitulo").textContent = "Editar Cliente";
  document.getElementById("clACreditoWrap").style.display = "block";
  document.getElementById("clNombre").value = cliente.NOMBRE || "";
  document.getElementById("clAlias").value = cliente.ALIAS || "";
  document.getElementById("clDni").value = cliente.DNI || "";
  document.getElementById("clTelefono").value = cliente.TELEFONO || "";
  document.getElementById("clEmpresa").value = cliente.EMPRESA || "";
  document.getElementById("clDireccion").value = cliente.DIRECCION || "";
  document.getElementById("clACredito").checked = String(cliente.A_CREDITO || "").toUpperCase() === "SI";
  [["clLocalidad", "LOCALIDAD"], ["clProvincia", "PROVINCIA"], ["clEmail", "EMAIL"], ["clNotas", "NOTAS"]].forEach(([id, k]) => {
    const el = document.getElementById(id); if (el) el.value = cliente[k] || "";
  });
  document.getElementById("clienteModalBackdrop").classList.add("show");
}

function cerrarModalNuevoCliente() {
  document.getElementById("clienteModalBackdrop").classList.remove("show");
}

let guardandoCliente = false;

async function guardarNuevoCliente() {
  if (guardandoCliente) return; // evita doble click mientras la petición anterior sigue en vuelo
  guardandoCliente = true;

  const nombre = document.getElementById("clNombre").value.trim();
  if (!nombre) {
    toast("Ingresá el nombre del cliente", "error");
    guardandoCliente = false;
    return;
  }

  const clienteIdEditando = document.getElementById("clClienteIdEditando").value.trim();
  const editando = !!clienteIdEditando;

  const btn = document.getElementById("btnGuardarNuevoCliente");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    const cuerpo = editando
      ? {
          action: "editarCliente",
          clienteId: clienteIdEditando,
          nombre,
          alias: document.getElementById("clAlias").value.trim(),
          dni: document.getElementById("clDni").value.trim(),
          telefono: document.getElementById("clTelefono").value.trim(),
          empresa: document.getElementById("clEmpresa").value.trim(),
          direccion: document.getElementById("clDireccion").value.trim(),
          localidad: (document.getElementById("clLocalidad") || {}).value?.trim() ?? undefined,
          provincia: (document.getElementById("clProvincia") || {}).value?.trim() ?? undefined,
          email: (document.getElementById("clEmail") || {}).value?.trim() ?? undefined,
          notas: (document.getElementById("clNotas") || {}).value?.trim() ?? undefined,
          aCredito: document.getElementById("clACredito").checked ? "SI" : "NO"
        }
      : {
          action: "crearCliente",
          nombre,
          alias: document.getElementById("clAlias").value.trim(),
          dni: document.getElementById("clDni").value.trim(),
          telefono: document.getElementById("clTelefono").value.trim(),
          empresa: document.getElementById("clEmpresa").value.trim(),
          direccion: document.getElementById("clDireccion").value.trim(),
          localidad: (document.getElementById("clLocalidad") || {}).value?.trim() ?? undefined,
          provincia: (document.getElementById("clProvincia") || {}).value?.trim() ?? undefined,
          email: (document.getElementById("clEmail") || {}).value?.trim() ?? undefined,
          notas: (document.getElementById("clNotas") || {}).value?.trim() ?? undefined,
          aCredito: document.getElementById("clACredito").checked ? "SI" : "NO"
        };

    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(cuerpo)
    });
    const data = await response.json();

    if (!data.success) {
      toast(data.message || (editando ? "No se pudo editar el cliente" : "No se pudo crear el cliente"), "error");
      return;
    }

    toast(editando ? "Cliente actualizado" : "Cliente creado", "success");
    cerrarModalNuevoCliente();
    invalidarCache("clientes"); // si no, la caché de 5 min sigue mostrando el dato viejo
    await cargarClientesDesdeBackend();

  } catch (error) {
    console.error("Error al crear cliente:", error);
    toast("Error de conexión al crear el cliente", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
    guardandoCliente = false;
  }
}

/** Used from the "Marcar a crédito" button on a client that only exists in the PEDIDOS-derived ranking (not yet in CLIENTES) */
async function marcarClienteDesdeHistorialACredito(dni) {
  /* confirm eliminado — acción creando un cliente nuevo, recuperable */

  try {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "crearClienteDesdeHistorialYMarcarCredito", dni })
    });
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo marcar a crédito", "error");
      return;
    }

    toast("Cliente marcado a crédito", "success");
    invalidarCache("clientes");
    cargarClientesDesdeBackend();

  } catch (error) {
    console.error("Error al marcar cliente a crédito:", error);
    toast("Error de conexión", "error");
  }
}

/** Deletes a client — only removes the CLIENTES row, never touches their past PEDIDOS or PAGOS_CREDITO history */
async function eliminarClienteForm(clienteId, nombre) {
  const cliente = clientesGlobal.find(c => String(c.CLIENTE_ID) === String(clienteId));
  const tieneSaldo = cliente && (Number(cliente.SALDO_PENDIENTE_ARS || 0) > 0 || Number(cliente.SALDO_PENDIENTE_USD || 0) > 0);

  const mensaje = tieneSaldo
    ? `⚠️ "${nombre}" todavía tiene saldo pendiente. ¿Eliminar igual su ficha de cliente? (sus pedidos y pagos pasados NO se borran, solo deja de aparecer en esta lista como cliente)`
    : `¿Eliminar al cliente "${nombre}"? Sus pedidos y pagos pasados no se borran, solo su ficha de cliente.`;

  /* confirm eliminado */

  try {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "eliminarCliente", clienteId })
    });
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo eliminar el cliente", "error");
      return;
    }

    toast("Cliente eliminado", "success");
    invalidarCache("clientes");
    cargarClientesDesdeBackend();

  } catch (error) {
    console.error("Error al eliminar cliente:", error);
    toast("Error de conexión al eliminar el cliente", "error");
  }
}

let detalleClienteDatosActuales = null;

async function abrirModalDetalleCliente(clienteId) {
  document.getElementById("detalleClienteModalBackdrop").classList.add("show");
  document.getElementById("detalleClienteNombre").textContent = "Cargando...";
  document.getElementById("pagoMonto").value = "";
  document.getElementById("pagoObservaciones").value = "";
  document.getElementById("pagoMonedaPago").value = "ARS";
  document.getElementById("pagoPrioridad").value = "ARS";
  document.getElementById("pagoTipoCambio").value = "";
  document.getElementById("pagoTipoCambioWrap").style.display = "none";
  const btnPago = document.getElementById("btnRegistrarPago");
  if (btnPago) { btnPago.disabled = false; btnPago.textContent = "Registrar pago"; }
  detalleClienteIdActual = clienteId;

  try {
    const response = await fetchAPI(API_URL + "?action=detalleClienteCredito&clienteId=" + encodeURIComponent(clienteId));
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo cargar el cliente", "error");
      cerrarModalDetalleCliente();
      return;
    }

    detalleClienteDatosActuales = data;

    document.getElementById("detalleClienteNombre").textContent =
      data.cliente.ALIAS ? `${data.cliente.NOMBRE} (${data.cliente.ALIAS})` : data.cliente.NOMBRE;

    const esCredito = String(data.cliente.A_CREDITO || "").toUpperCase() === "SI";

    document.getElementById("cardAvisoNoCredito").style.display = esCredito ? "none" : "block";
    document.getElementById("cardFormularioPago").style.display = esCredito ? "block" : "none";

    if (esCredito) {
      // Cliente a crédito: los números son una deuda real, vienen
      // calculados del backend (ya restan los pagos registrados).
      document.getElementById("detalleClienteTotalCompradoArs").textContent = "$" + Number(data.totalCompradoArs).toLocaleString("es-AR");
      document.getElementById("detalleClienteTotalPagadoArs").textContent = "$" + Number(data.totalPagadoArs).toLocaleString("es-AR");
      document.getElementById("detalleClienteSaldoArs").textContent = "$" + Number(data.saldoPendienteArs).toLocaleString("es-AR");

      document.getElementById("detalleClienteTotalCompradoUsd").textContent = "US$" + Number(data.totalCompradoUsd).toLocaleString("es-AR");
      document.getElementById("detalleClienteTotalPagadoUsd").textContent = "US$" + Number(data.totalPagadoUsd).toLocaleString("es-AR");
      document.getElementById("detalleClienteSaldoUsd").textContent = "US$" + Number(data.saldoPendienteUsd).toLocaleString("es-AR");
    } else {
      // Cliente NO a crédito: sus pedidos fueron compras al contado.
      // El total comprado es solo informativo (no es deuda), y el
      // saldo pendiente siempre es $0 — nunca debe nada.
      let totalArsInformativo = 0, totalUsdInformativo = 0;
      (data.pedidos || []).forEach(p => {
        if (String(p.MONEDA || "ARS").toUpperCase() === "USD") totalUsdInformativo += Number(p.TOTAL || 0);
        else totalArsInformativo += Number(p.TOTAL || 0);
      });

      document.getElementById("detalleClienteTotalCompradoArs").textContent = "$" + totalArsInformativo.toLocaleString("es-AR");
      document.getElementById("detalleClienteTotalPagadoArs").textContent = "$" + totalArsInformativo.toLocaleString("es-AR");
      document.getElementById("detalleClienteSaldoArs").textContent = "$0";

      document.getElementById("detalleClienteTotalCompradoUsd").textContent = "US$" + totalUsdInformativo.toLocaleString("es-AR");
      document.getElementById("detalleClienteTotalPagadoUsd").textContent = "US$" + totalUsdInformativo.toLocaleString("es-AR");
      document.getElementById("detalleClienteSaldoUsd").textContent = "US$0";
    }

    const pedidosTbody = document.getElementById("detalleClientePedidosTabla");
    pedidosTbody.innerHTML = data.pedidos.length === 0
      ? `<tr><td colspan="4" class="text-center text-muted">Sin pedidos</td></tr>`
      : data.pedidos.map(p => {
          const moneda = String(p.MONEDA || "ARS").toUpperCase();
          const simbolo = moneda === "USD" ? "US$" : "$";
          return `
          <tr>
            <td>${p.FECHA ? new Date(p.FECHA).toLocaleDateString("es-AR") : "—"}</td>
            <td class="money">${simbolo}${Number(p.TOTAL || 0).toLocaleString("es-AR")}</td>
            <td>${moneda}${moneda === "USD" && p.TIPO_CAMBIO ? ` (TC ${p.TIPO_CAMBIO})` : ""}</td>
            <td>${escapeHtml(p.ESTADO || "")}</td>
          </tr>`;
        }).join("");

    const pagosTbody = document.getElementById("detalleClientePagosTabla");
    pagosTbody.innerHTML = data.pagos.length === 0
      ? `<tr><td colspan="6" class="text-center text-muted">Sin pagos registrados todavía</td></tr>`
      : data.pagos.map(pago => `
          <tr>
            <td>${pago.FECHA ? new Date(pago.FECHA).toLocaleDateString("es-AR") : "—"}</td>
            <td class="money">${String(pago.MONEDA_PAGO || "ARS").toUpperCase() === "USD" ? "US$" : "$"}${Number(pago.MONTO_PAGO || 0).toLocaleString("es-AR")}</td>
            <td class="money">${Number(pago.MONTO_APLICADO_ARS || 0) > 0 ? "$" + Number(pago.MONTO_APLICADO_ARS).toLocaleString("es-AR") : "—"}</td>
            <td class="money">${Number(pago.MONTO_APLICADO_USD || 0) > 0 ? "US$" + Number(pago.MONTO_APLICADO_USD).toLocaleString("es-AR") : "—"}</td>
            <td>${escapeHtml(pago.FORMA_PAGO || "")}</td>
            <td>${escapeHtml(pago.OBSERVACIONES || "")}</td>
          </tr>`).join("");

    // Deudas extra
    const deudasExtra = data.deudasExtra || [];
    const deudasWrap = document.getElementById("deudaExtraTablaWrap");
    const deudasTbody = document.getElementById("detalleClienteDeudasExtraTabla");
    if (deudasExtra.length > 0) {
      deudasWrap.style.display = "block";
      deudasTbody.innerHTML = deudasExtra.map(d => {
        const simbolo = String(d.MONEDA || "ARS").toUpperCase() === "USD" ? "US$" : "$";
        return `<tr>
          <td>${d.FECHA ? new Date(d.FECHA).toLocaleDateString("es-AR") : "—"}</td>
          <td>${escapeHtml(String(d.MONEDA || "ARS").toUpperCase())}</td>
          <td class="money">${simbolo}${Number(d.MONTO || 0).toLocaleString("es-AR")}</td>
          <td>${escapeHtml(d.CONCEPTO || "")}</td>
        </tr>`;
      }).join("");
    } else {
      deudasWrap.style.display = "none";
    }

  } catch (error) {
    console.error("Error al cargar el detalle del cliente:", error);
    toast("Error de conexión al cargar el cliente", "error");
  }
}

function cerrarModalDetalleCliente() {
  document.getElementById("detalleClienteModalBackdrop").classList.remove("show");
  detalleClienteIdActual = null;
  detalleClienteDatosActuales = null;
}

let deudaExtraClienteId = null;

function abrirModalDeudaExtraDirecto(clienteId, nombre) {
  deudaExtraClienteId = clienteId;
  document.getElementById("deudaExtraModalNombre").textContent = nombre;
  document.getElementById("deudaExtraModalMonto").value = "";
  document.getElementById("deudaExtraModalMoneda").value = "ARS";
  document.getElementById("deudaExtraModalConcepto").value = "";
  const btnDeuda = document.getElementById("btnRegistrarDeudaExtra");
  if (btnDeuda) { btnDeuda.disabled = false; btnDeuda.textContent = "Registrar deuda"; }
  document.getElementById("deudaExtraModalBackdrop").classList.add("show");
  setTimeout(() => document.getElementById("deudaExtraModalMonto").focus(), 100);
}

function cerrarModalDeudaExtra() {
  document.getElementById("deudaExtraModalBackdrop").classList.remove("show");
  deudaExtraClienteId = null;
}

async function registrarDeudaExtraForm() {
  if (!deudaExtraClienteId) return;

  const monto = Number(document.getElementById("deudaExtraModalMonto").value);
  if (!monto || monto <= 0) { toast("Ingresá un monto válido", "error"); return; }

  const moneda = document.getElementById("deudaExtraModalMoneda").value;
  const concepto = document.getElementById("deudaExtraModalConcepto").value.trim();
  if (!concepto) { toast("El concepto es obligatorio", "error"); return; }

  const btn = document.getElementById("btnRegistrarDeudaExtra");
  if (btn) { btn.disabled = true; btn.textContent = "Registrando..."; }

  try {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "crearDeudaExtra",
        clienteId: deudaExtraClienteId,
        monto,
        moneda,
        concepto
      })
    });
    const data = await response.json();
    if (!data.success) {
      toast(data.message || "No se pudo registrar la deuda", "error");
      if (btn) { btn.disabled = false; btn.textContent = "Registrar deuda"; }
      return;
    }
    toast("Deuda registrada correctamente", "success");
    cerrarModalDeudaExtra();
    invalidarCache("clientes");
    cargarClientesDesdeBackend();
  } catch (error) {
    console.error("Error al registrar deuda extra:", error);
    toast("Error de conexión al registrar la deuda", "error");
    if (btn) { btn.disabled = false; btn.textContent = "Registrar deuda"; }
  }
}

function toggleFormDeudaExtra() {
  const form = document.getElementById("deudaExtraForm");
  const icon = document.getElementById("deudaExtraToggleIcon");
  const visible = form.style.display !== "none";
  form.style.display = visible ? "none" : "block";
  icon.textContent = visible ? "▼ Expandir" : "▲ Cerrar";
}

/** Shows the exchange-rate field only when the payment currency and the priority currency differ — that's the only case where a conversion is actually needed */
function actualizarVisibilidadTipoCambioPago() {
  const monedaPago = document.getElementById("pagoMonedaPago").value;
  const prioridad = document.getElementById("pagoPrioridad").value;
  document.getElementById("pagoTipoCambioWrap").style.display = monedaPago !== prioridad ? "block" : "none";
  _prellenarTipoCambio("pagoTipoCambio", monedaPago !== prioridad);
}

/** Tipo de cambio sugerido: en el depósito, su dólar fijado a mano; si no, el del ingreso de mercadería */
function _dolarSugerido() {
  if (typeof esModoDeposito === "function" && esModoDeposito() && typeof dolarDelDeposito === "function" && dolarDelDeposito() > 0) return dolarDelDeposito();
  return (typeof ingresoTC !== "undefined" && ingresoTC && ingresoTC.valor > 0) ? ingresoTC.valor : 0;
}
function _prellenarTipoCambio(id, visible) {
  const el = document.getElementById(id);
  if (!el || !visible || el.value) return;
  const v = _dolarSugerido();
  if (v > 0) el.value = v;
}

async function registrarPagoCreditoForm() {
  if (!detalleClienteIdActual) return;

  const monto = Number(document.getElementById("pagoMonto").value);
  if (!monto || monto <= 0) { toast("Ingresá un monto válido", "error"); return; }

  const monedaPago = document.getElementById("pagoMonedaPago").value;
  const prioridad = document.getElementById("pagoPrioridad").value;
  const tipoCambio = document.getElementById("pagoTipoCambio").value.trim();

  if (monedaPago !== prioridad && (!tipoCambio || Number(tipoCambio) <= 0)) {
    toast("Ingresá el tipo de cambio para convertir entre pesos y dólares", "error");
    return;
  }

  const btn = document.getElementById("btnRegistrarPago");
  if (btn) { btn.disabled = true; btn.textContent = "Registrando..."; }

  try {
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "registrarPagoCredito",
        clienteId: detalleClienteIdActual,
        monto,
        monedaPago,
        prioridad,
        tipoCambio: tipoCambio || "",
        formaPago: document.getElementById("pagoFormaPago").value,
        observaciones: document.getElementById("pagoObservaciones").value.trim(),
        cajero: (sessionStorage.getItem("vendedor") || "ADMIN")
      })
    });
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo registrar el pago", "error");
      if (btn) { btn.disabled = false; btn.textContent = "Registrar pago"; }
      return;
    }

    toast("Pago registrado", "success");
    // Invalidar caché para que cierre de caja, movimientos y la lista
    // de clientes (saldo actualizado) muestren datos frescos
    const hoy = _hoyISO();
    try { localStorage.removeItem("vpos_cache_movimientos_" + hoy); } catch(e) {}
    try { localStorage.removeItem("vpos_cache_cierre_" + hoy); } catch(e) {}
    try { localStorage.removeItem("vpos_cache_cierre_ayer"); } catch(e) {}
    invalidarCache("clientes");
    abrirModalDetalleCliente(detalleClienteIdActual);
    cargarClientesDesdeBackend();

  } catch (error) {
    console.error("Error al registrar el pago:", error);
    toast("Error de conexión al registrar el pago", "error");
    if (btn) { btn.disabled = false; btn.textContent = "Registrar pago"; }
  }
}

let detalleClienteIdActual = null;

/**
 * Genera un informe A4 (no térmico) con el resumen de cuenta del
 * cliente que está abierto en el modal de detalle — saldos por
 * moneda, detalle de pedidos, y detalle de pagos. Reutiliza
 * #etiquetasPrintArea, el mismo contenedor que ya usa el sistema
 * para imprimir en A4 (las etiquetas de producto usan el mismo
 * mecanismo, ver imprimirEtiquetas).
 */
function imprimirInformeCliente() {
  if (!detalleClienteDatosActuales) { toast("Esperá a que termine de cargar el cliente", "error"); return; }

  const data = detalleClienteDatosActuales;
  const cliente = data.cliente;
  const fechaHoy = new Date().toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" });

  const filasPedidos = data.pedidos.length === 0
    ? `<tr><td colspan="4" style="text-align:center; color:#888;">Sin pedidos</td></tr>`
    : data.pedidos.map(p => {
        const moneda = String(p.MONEDA || "ARS").toUpperCase();
        const simbolo = moneda === "USD" ? "US$" : "$";
        return `<tr>
          <td>${p.FECHA ? new Date(p.FECHA).toLocaleDateString("es-AR") : "—"}</td>
          <td style="text-align:right;">${simbolo}${Number(p.TOTAL || 0).toLocaleString("es-AR")}</td>
          <td>${moneda}</td>
          <td>${escapeHtml(p.ESTADO || "")}</td>
        </tr>`;
      }).join("");

  const filasPagos = data.pagos.length === 0
    ? `<tr><td colspan="5" style="text-align:center; color:#888;">Sin pagos registrados</td></tr>`
    : data.pagos.map(pago => `<tr>
        <td>${pago.FECHA ? new Date(pago.FECHA).toLocaleDateString("es-AR") : "—"}</td>
        <td style="text-align:right;">${String(pago.MONEDA_PAGO || "ARS").toUpperCase() === "USD" ? "US$" : "$"}${Number(pago.MONTO_PAGO || 0).toLocaleString("es-AR")}</td>
        <td style="text-align:right;">${Number(pago.MONTO_APLICADO_ARS || 0) > 0 ? "$" + Number(pago.MONTO_APLICADO_ARS).toLocaleString("es-AR") : "—"}</td>
        <td style="text-align:right;">${Number(pago.MONTO_APLICADO_USD || 0) > 0 ? "US$" + Number(pago.MONTO_APLICADO_USD).toLocaleString("es-AR") : "—"}</td>
        <td>${escapeHtml(pago.FORMA_PAGO || "")}</td>
      </tr>`).join("");

  const html = `
    <div style="font-family:Arial, sans-serif; color:#1a1a1a; padding:10mm;">
      <h2 style="margin:0 0 4px;">${escapeHtml((configNegocioCache && configNegocioCache.nombre) || "Informe de cuenta")}</h2>
      <div style="color:#666; font-size:12px; margin-bottom:18px;">Informe de cuenta corriente — generado el ${fechaHoy}</div>

      <div style="border:1px solid #ddd; border-radius:8px; padding:12px 16px; margin-bottom:18px;">
        <strong>${escapeHtml(cliente.NOMBRE)}</strong>${cliente.ALIAS ? ` (${escapeHtml(cliente.ALIAS)})` : ""}<br>
        ${cliente.DNI ? `DNI/CUIT: ${escapeHtml(cliente.DNI)}<br>` : ""}
        ${cliente.TELEFONO ? `Tel: ${escapeHtml(cliente.TELEFONO)}<br>` : ""}
        ${cliente.DIRECCION ? `Dirección: ${escapeHtml(cliente.DIRECCION)}` : ""}
      </div>

      <table style="width:100%; border-collapse:collapse; margin-bottom:18px;">
        <tr style="background:#f4f4f4;">
          <th style="padding:8px; text-align:left; border:1px solid #ddd;"></th>
          <th style="padding:8px; text-align:right; border:1px solid #ddd;">Comprado</th>
          <th style="padding:8px; text-align:right; border:1px solid #ddd;">Pagado</th>
          <th style="padding:8px; text-align:right; border:1px solid #ddd;">Saldo</th>
        </tr>
        <tr>
          <td style="padding:8px; border:1px solid #ddd;"><strong>Pesos (ARS)</strong></td>
          <td style="padding:8px; text-align:right; border:1px solid #ddd;">$${Number(data.totalCompradoArs).toLocaleString("es-AR")}</td>
          <td style="padding:8px; text-align:right; border:1px solid #ddd;">$${Number(data.totalPagadoArs).toLocaleString("es-AR")}</td>
          <td style="padding:8px; text-align:right; border:1px solid #ddd; font-weight:bold;">$${Number(data.saldoPendienteArs).toLocaleString("es-AR")}</td>
        </tr>
        <tr>
          <td style="padding:8px; border:1px solid #ddd;"><strong>Dólares (USD)</strong></td>
          <td style="padding:8px; text-align:right; border:1px solid #ddd;">US$${Number(data.totalCompradoUsd).toLocaleString("es-AR")}</td>
          <td style="padding:8px; text-align:right; border:1px solid #ddd;">US$${Number(data.totalPagadoUsd).toLocaleString("es-AR")}</td>
          <td style="padding:8px; text-align:right; border:1px solid #ddd; font-weight:bold;">US$${Number(data.saldoPendienteUsd).toLocaleString("es-AR")}</td>
        </tr>
      </table>

      <h4 style="margin:0 0 8px;">Pedidos</h4>
      <table style="width:100%; border-collapse:collapse; margin-bottom:18px; font-size:13px;">
        <tr style="background:#f4f4f4;">
          <th style="padding:6px; text-align:left; border:1px solid #ddd;">Fecha</th>
          <th style="padding:6px; text-align:right; border:1px solid #ddd;">Total</th>
          <th style="padding:6px; text-align:left; border:1px solid #ddd;">Moneda</th>
          <th style="padding:6px; text-align:left; border:1px solid #ddd;">Estado</th>
        </tr>
        ${filasPedidos}
      </table>

      <h4 style="margin:0 0 8px;">Pagos</h4>
      <table style="width:100%; border-collapse:collapse; font-size:13px;">
        <tr style="background:#f4f4f4;">
          <th style="padding:6px; text-align:left; border:1px solid #ddd;">Fecha</th>
          <th style="padding:6px; text-align:right; border:1px solid #ddd;">Pago</th>
          <th style="padding:6px; text-align:right; border:1px solid #ddd;">Aplicado ARS</th>
          <th style="padding:6px; text-align:right; border:1px solid #ddd;">Aplicado USD</th>
          <th style="padding:6px; text-align:left; border:1px solid #ddd;">Forma</th>
        </tr>
        ${filasPagos}
      </table>
    </div>
  `;

  // Por si quedó contenido de una impresión de ticket térmico anterior
  // a medio camino, se limpia el otro contenedor de impresión para que
  // el CSS de :empty/:not(:empty) elija el correcto sin ambigüedad.
  const thermalFrameInforme = document.getElementById("thermalPrintFrame");
  if (thermalFrameInforme) thermalFrameInforme.innerHTML = "";

  document.getElementById("etiquetasPrintArea").innerHTML = html;

  _setPrintPageSize("A4");
  setTimeout(() => {
    window.print();
  }, 100);
}

/**
 * Imprime una etiqueta de envío A4 para un pedido en estado PREPARANDO.
 * Usa el mismo #etiquetasPrintArea que el resto de las impresiones A4.
 * Incluye un código de barras Code128 con el número de pedido, generado
 * con JsBarcode (ya cargado en el panel).
 */
function imprimirEtiquetaEnvio(datos) {
  const { cliente, telefono, direccion, localidad, provincia, codigoPostal, dni, transporte } = datos;
  // Los números nuevos (1001…) pueden llegar como número: siempre como texto
  const pedidoId = String(datos.pedidoId ?? "");

  const cpStr = codigoPostal ? `  CP:${codigoPostal}` : "";
  const localidadStr = [localidad, cpStr].filter(Boolean).join("").trim();

  const html = `
    <div style="
      font-family: Arial, Helvetica, sans-serif;
      width: 190mm;
      min-height: 130mm;
      margin: 10mm auto;
      padding: 8mm;
      box-sizing: border-box;
      border: 2px solid #000;
      border-radius: 4mm;
      page-break-inside: avoid;
      color: #000;
    ">

      <!-- Franja frágil / cabecera -->
      <div style="background:#d32f2f; color:#fff; text-align:center; padding:6mm 4mm; border-radius:2mm; margin-bottom:6mm;">
        <div style="font-size:12pt; font-weight:900; letter-spacing:2px; text-transform:uppercase;">POR FAVOR</div>
        <div style="font-size:12pt; font-weight:700; letter-spacing:4px; text-transform:uppercase;">MANEJESE CON CUIDADO</div>
        <div style="font-size:32pt; font-weight:900; letter-spacing:6px; margin:4px 0;">FRAGIL</div>
        <div style="font-size:12pt; font-weight:700; letter-spacing:3px;">== GRACIAS ==</div>
      </div>

      <!-- Datos del destinatario -->
      <table style="width:100%; border-collapse:collapse; font-size:13pt;">
        <tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; width:28mm; vertical-align:top; padding-top:4px;">Nombre:</td>
          <td style="font-size:30pt; font-weight:900; text-transform:uppercase; letter-spacing:1px;">${escapeHtml(cliente)}</td>
        </tr>
        ${telefono ? `<tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; padding-top:4px;">Teléfono:</td>
          <td style="font-size:26pt; font-weight:700;">${escapeHtml(telefono)}</td>
        </tr>` : ""}
        ${direccion ? `<tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; padding-top:4px;">Dirección:</td>
          <td style="font-size:26pt; font-weight:700; text-transform:uppercase;">${escapeHtml(direccion)}</td>
        </tr>` : ""}
        ${localidadStr ? `<tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; padding-top:4px;">Localidad:</td>
          <td style="font-size:26pt; font-weight:700; text-transform:uppercase;">${escapeHtml(localidadStr)}</td>
        </tr>` : ""}
        ${provincia ? `<tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; padding-top:4px;">Provincia:</td>
          <td style="font-size:26pt; font-weight:700; text-transform:uppercase;">${escapeHtml(provincia)}</td>
        </tr>` : ""}
        ${dni ? `<tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; padding-top:4px;">DNI:</td>
          <td style="font-size:26pt; font-weight:700;">${escapeHtml(dni)}</td>
        </tr>` : ""}
        ${transporte ? `<tr>
          <td style="font-size:10pt; font-weight:700; text-transform:uppercase; color:#000; padding-top:4px;">Transporte:</td>
          <td style="font-size:26pt; font-weight:700; text-transform:uppercase;">${escapeHtml(transporte)}</td>
        </tr>` : ""}
      </table>

      <!-- Código de barras con número de pedido -->
      <div style="margin-top:6mm; text-align:center; border-top:1px solid #ddd; padding-top:4mm;">
        <svg id="barcode-etiqueta-${pedidoId.replace(/[^a-zA-Z0-9]/g,'_')}"></svg>
      </div>

    </div>
  `;

  // Por si quedó contenido de una impresión de ticket térmico anterior
  // a medio camino, se limpia el otro contenedor de impresión para que
  // el CSS de :empty/:not(:empty) elija el correcto sin ambigüedad.
  const thermalFrameEnvio = document.getElementById("thermalPrintFrame");
  if (thermalFrameEnvio) thermalFrameEnvio.innerHTML = "";

  const area = document.getElementById("etiquetasPrintArea");
  area.innerHTML = html;
  _setPrintPageSize("A4");

  // Generar el código de barras después de que el DOM esté listo
  setTimeout(() => {
    const svgId = `barcode-etiqueta-${pedidoId.replace(/[^a-zA-Z0-9]/g,'_')}`;
    const svgEl = document.getElementById(svgId);
    if (svgEl && typeof JsBarcode !== "undefined") {
      JsBarcode(svgEl, pedidoId, {
        format: "CODE128",
        width: 2.5,
        height: 60,
        displayValue: true,
        fontSize: 14,
        margin: 8,
        textMargin: 4
      });
    }
    window.print();
  }, 150);
}

/* ===================== STOCK BAJO / AGOTADOS / MAS VENDIDOS ===================== */

async function cargarStockBajo() {
  mostrarSeccion("stockBajoProductos");
  try {
    // Reusar productos ya cacheados si están disponibles
    let productos = productosPOS.length > 0 ? productosPOS : null;
    if (!productos) {
      const response = await fetchAPI(API_URL + "?action=productos");
      const data = await response.json();
      productos = data.productos || [];
    }
    let html = "";
    const filtrados = productos.filter(p => { const s = Number(p.STOCK || 0); return s > 0 && s <= 5; });
    if (filtrados.length === 0) {
      html = `<tr><td colspan="3" class="text-center text-muted py-4">Sin productos con stock bajo</td></tr>`;
    }
    filtrados.forEach(p => {
      html += `<tr><td class="mono">${escapeHtml(p.CODIGO)}</td><td>${escapeHtml(p.PRODUCTO)}</td><td>${p.STOCK}</td></tr>`;
    });
    document.getElementById("tablaStockBajo").innerHTML = html;
  } catch (error) {
    console.error("Error stock bajo:", error);
  }
}

async function cargarAgotados() {
  mostrarSeccion("productosAgotados");
  try {
    let productos = productosPOS.length > 0 ? productosPOS : null;
    if (!productos) {
      const response = await fetchAPI(API_URL + "?action=productos");
      const data = await response.json();
      productos = data.productos || [];
    }
    let html = "";
    const filtrados = productos.filter(p => Number(p.STOCK || 0) === 0);
    if (filtrados.length === 0) {
      html = `<tr><td colspan="3" class="text-center text-muted py-4">No hay productos agotados</td></tr>`;
    }
    filtrados.forEach(p => {
      html += `<tr><td class="mono">${escapeHtml(p.CODIGO)}</td><td>${escapeHtml(p.PRODUCTO)}</td><td>0</td></tr>`;
    });
    document.getElementById("tablaAgotados").innerHTML = html;
  } catch (error) {
    console.error("Error agotados:", error);
  }
}

async function cargarMasVendidos() {
  try {
    const response = await fetchAPI(API_URL + "?action=masVendidos");
    const data = await response.json();
    let html = "";
    (data.productos || []).forEach(p => {
      html += `<tr><td class="mono">${p.CODIGO}</td><td>${p.PRODUCTO}</td><td>${p.VENDIDOS}</td></tr>`;
    });
    document.getElementById("tablaMasVendidos").innerHTML = html;
  } catch (error) {
    console.error("Error más vendidos:", error);
  }
}

/* ===================================================================
   PUNTO DE VENTA
=================================================================== */

let productosPOS    = [];
const _ventasMapPOS = {}; // mapa ventaId → objeto venta, para evitar JSON en onclick
let ticketPOS        = [];
let categoriaActivaPOS = "TODAS";
let formaPagoPOS    = "EFECTIVO";
let ultimoCodigoAgregadoPOS = null; // usado por el atajo +/- de cantidad
let posTileFocusIdx  = -1;          // índice de la tarjeta con foco de teclado en el grid

// Descuento / recargo aplicado al ticket actual
let descuentoTipoPOS   = "PORCENTAJE"; // "PORCENTAJE" | "MONTO"
let descuentoValorPOS  = 0;            // valor ingresado (ej: 10 para 10%, o 500 para $500)
let descuentoActivoPOS = false;
let ajusteModoPOS      = "DESCUENTO";  // "DESCUENTO" | "RECARGO" — signo del ajuste aplicado

/**
 * Índice CODIGO -> producto y CODIGO_CAJA -> producto, para no hacer
 * un Array.find() (O(n) sobre ~2000 productos) en cada escaneo. Se
 * reconstruye solo, y una sola vez, la primera vez que se necesita
 * después de que productosPOS cambió de referencia (recarga de
 * catálogo) — se compara por referencia, así que no hace falta tocar
 * cada uno de los puntos donde se reasigna productosPOS.
 */
let _productosPOSMapaCodigo = null;
let _productosPOSMapaCaja = null;
let _productosPOSMapaRef = null;

function _asegurarMapasProductosPOS() {
  if (_productosPOSMapaRef === productosPOS) return;
  _productosPOSMapaCodigo = new Map();
  _productosPOSMapaCaja = new Map();
  for (const p of productosPOS) {
    _productosPOSMapaCodigo.set(String(p.CODIGO).trim().toLowerCase(), p);
    if (p.CODIGO_CAJA && Number(p.UNIDADES_POR_CAJA) > 0) {
      _productosPOSMapaCaja.set(String(p.CODIGO_CAJA).trim().toLowerCase(), p);
    }
  }
  _productosPOSMapaRef = productosPOS;
}

/** Busca un producto por CODIGO (unitario) en O(1) en vez de recorrer todo productosPOS */
function buscarProductoPOSPorCodigo(codigo) {
  _asegurarMapasProductosPOS();
  return _productosPOSMapaCodigo.get(String(codigo).trim().toLowerCase()) || null;
}

/** Busca un producto por CODIGO_CAJA (código de la caja/bulto cerrado) en O(1) */
function buscarProductoPOSPorCodigoCaja(codigo) {
  _asegurarMapasProductosPOS();
  return _productosPOSMapaCaja.get(String(codigo).trim().toLowerCase()) || null;
}

/** Limpia el caché de productos del POS y recarga desde el backend */
async function actualizarCatalogoPOSManual() {
  try { localStorage.removeItem("veekpos_productos_cache"); } catch(e) {}
  productosPOS = [];
  toast("Actualizando catálogo...", "success");
  await asegurarProductosPOS();
  toast("Catálogo actualizado", "success");
}

async function asegurarProductosPOS() {
  if (productosPOS.length > 0) return;

  const CACHE_KEY = "veekpos_productos_cache";
  const CACHE_TTL = 10 * 60 * 1000;
  try {
    const cached = localStorage.getItem(CACHE_KEY);
    if (cached) {
      const { ts, productos } = JSON.parse(cached);
      if (Date.now() - ts < CACHE_TTL && productos.length > 0) {
        productosPOS = productos;
        construirCategoriasPOS();
        renderPosGrid();
        _actualizarCacheProductosPOS(CACHE_KEY);
        return;
      }
    }
  } catch(e) {}

  await _actualizarCacheProductosPOS(CACHE_KEY);
}

async function _actualizarCacheProductosPOS(cacheKey) {
  try {
    const response = await fetchAPI(API_URL + "?action=productos");
    const data = await response.json();
    const productos = data.productos || [];
    if (!_esAppEscritorio()) { try { localStorage.setItem(cacheKey, JSON.stringify({ ts: Date.now(), productos })); } catch(e) {} }
    if (JSON.stringify(productos.map(p => p.CODIGO + p.STOCK)) !==
        JSON.stringify(productosPOS.map(p => p.CODIGO + p.STOCK))) {
      productosPOS = productos;
      construirCategoriasPOS();
      renderPosGrid();
    } else {
      productosPOS = productos;
    }
  } catch(e) {
    console.error("Error actualizando caché de productos POS:", e);
  }
}

function construirCategoriasPOS() {
  const cont = document.getElementById("posCategorias");
  if (!cont) return;

  // Permite desplazar las categorías con la rueda normal del mouse
  // (vertical), sin necesidad de Shift+rueda ni de arrastrar — se
  // agrega una sola vez, porque este <div> no se recrea entre
  // llamadas (solo cambia su innerHTML más abajo).
  if (!cont.dataset.wheelListo) {
    cont.addEventListener("wheel", (e) => {
      if (e.deltaY === 0) return; // ya es un scroll horizontal nativo (trackpad), no interferir
      e.preventDefault();
      cont.scrollLeft += e.deltaY;
    });
    cont.dataset.wheelListo = "1";
  }

  const categorias = new Set();
  productosPOS.forEach(p => { if (p.CATEGORIA) categorias.add(String(p.CATEGORIA).trim()); });
  if (categorias.size === 0) { cont.innerHTML = ""; return; }
  let html = `<div class="cat-chip active" data-cat="TODAS" onclick="filtrarCategoriaPOS('TODAS', this)">Todas</div>`;
  categorias.forEach(c => {
    html += `<div class="cat-chip" data-cat="${escapeHtml(c)}" onclick="filtrarCategoriaPOS('${escapeJsAttr(c)}', this)">${escapeHtml(c)}</div>`;
  });
  cont.innerHTML = html;
}

function filtrarCategoriaPOS(cat, el) {
  categoriaActivaPOS = cat;
  posTileFocusIdx = -1;
  document.querySelectorAll(".cat-chip").forEach(c => c.classList.remove("active"));
  if (el) el.classList.add("active");
  renderPosGrid();
}

/* ---- búsqueda tolerante a errores (código, nombre, categoría y alias) ---- */

/** minúsculas + sin acentos, para comparar sin importar mayúsculas/tildes */
function normalizarBusquedaPOS(t) {
  return String(t || "").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

/**
 * Distancia de Levenshtein entre dos strings, con corte anticipado:
 * si en algún punto ya se superó maxDist, devuelve maxDist + 1 sin
 * terminar el cálculo completo (más rápido cuando se llama muchas
 * veces por tecla tipeada en el buscador del POS).
 */
function distanciaLevenshteinPOS(a, b, maxDist) {
  if (a === b) return 0;
  if (Math.abs(a.length - b.length) > maxDist) return maxDist + 1;

  const n = b.length;
  let prev = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;

  for (let i = 1; i <= a.length; i++) {
    const curr = new Array(n + 1);
    curr[0] = i;
    let filaMin = curr[0];
    for (let j = 1; j <= n; j++) {
      const costo = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + costo);
      if (curr[j] < filaMin) filaMin = curr[j];
    }
    if (filaMin > maxDist) return maxDist + 1;
    prev = curr;
  }
  return prev[n];
}

/** True si "palabraTexto" contiene "palabraBusqueda", o si difiere de ella en pocas letras (typo tolerado según el largo de lo tipeado) */
function palabraCoincideTolerantePOS(palabraBusqueda, palabraTexto) {
  if (!palabraBusqueda) return true;
  if (palabraTexto.includes(palabraBusqueda)) return true;
  // Palabras cortas (<=4 letras) quedan fuera del fuzzy: una distancia
  // de 1 sobre 3-4 letras es 25-33% del texto tipeado y generaba
  // falsos positivos ("pan" encontrando "van", "can", "fan", "pin",
  // etc.). Para estas, solo vale el match literal de arriba.
  if (palabraBusqueda.length <= 4) return false;
  // Un código de barras escaneado (solo dígitos, 6+ caracteres) no
  // tiene "typos" que tolerar — si no matcheó literal arriba, no va a
  // matchear con Levenshtein tampoco. Evita correr la comparación
  // difusa contra cada palabra de cada producto del catálogo en cada
  // tecla del escaneo (el costo que hacía sentir lento el escaneo).
  if (/^\d{6,}$/.test(palabraBusqueda)) return false;
  // A partir de 5 letras se tolera 1 error; recién a partir de 8 se
  // tolera 2 (antes era a partir de 6, demasiado laxo: "cocina" (6)
  // llegaba a matchear "colina" o "cadena" con solo 2 letras en común).
  const maxDist = palabraBusqueda.length <= 7 ? 1 : 2;
  return distanciaLevenshteinPOS(palabraBusqueda, palabraTexto, maxDist) <= maxDist;
}

/**
 * Compara lo tipeado en el buscador contra CODIGO, PRODUCTO,
 * CATEGORIA y ALIAS del producto (el alias no se muestra en ningún
 * lado, solo suma más "palabras" por las que ese producto puede
 * encontrarse). Primero intenta un match literal (rápido y exacto);
 * si no matchea, prueba palabra por palabra con tolerancia a errores
 * de tipeo, para que buscar "cocacol" o "coaca" igual encuentre
 * "Coca Cola".
 */
function productoCoincideBusquedaPOS(producto, filtroTexto) {
  const texto = normalizarBusquedaPOS(filtroTexto);
  if (!texto) return true;

  // El texto normalizado (y sus palabras) de cada producto no cambia
  // entre tecla y tecla — se calcula una sola vez y se guarda en el
  // propio producto. Con ~2000 productos, recalcular esto (lowercase +
  // sacar tildes + unir 4 campos) en CADA letra tipeada era el costo
  // real de la búsqueda por nombre; cacheado, cada tecla solo hace la
  // comparación en sí.
  let textoCompleto = producto._textoBusquedaPOS;
  if (textoCompleto === undefined) {
    textoCompleto = normalizarBusquedaPOS(
      [producto.CODIGO, producto.PRODUCTO, producto.CATEGORIA, producto.ALIAS].join(" ")
    );
    producto._textoBusquedaPOS = textoCompleto;
    producto._palabrasBusquedaPOS = textoCompleto.split(/\s+/).filter(Boolean);
  }

  // Camino rápido: la frase completa tipeada aparece tal cual
  if (textoCompleto.includes(texto)) return true;

  const palabrasTexto = producto._palabrasBusquedaPOS;
  const palabrasBusqueda = texto.split(/\s+/).filter(Boolean);

  // Cada palabra tipeada tiene que encontrar alguna palabra del
  // producto que la contenga o que sea "casi igual" (typo tolerado)
  return palabrasBusqueda.every(pb =>
    palabrasTexto.some(pt => palabraCoincideTolerantePOS(pb, pt))
  );
}

/* ---- product grid ---- */

// Producto que queda "fijado" como primer tile de la grilla después de
// cada escaneo/click, además de aparecer en la tarjeta de "Último
// escaneado" de más arriba — se mantiene ahí hasta el próximo escaneo
// (se reemplaza) o hasta que se vacíe/finalice la venta (ver
// limpiarPineadoPOS, llamado desde los puntos donde ticketPOS se
// resetea a []).
let productoPineadoPOS = null;

/** Arma el HTML de un tile de producto — reutilizado para la grilla normal y para el tile pineado */
function _armarTileProductoHTML(p, dataIdxAttr, esPineado) {
  const stock     = p.STOCK !== undefined ? Number(p.STOCK) : null;
  const agotado   = stock !== null && stock <= 0;
  const stockBajo = stock !== null && stock > 0 && stock <= 5;

  let stockBadge = "";
  if (agotado)        stockBadge = `<span class="tile-stock out">Sin stock</span>`;
  else if (stockBajo) stockBadge = `<span class="tile-stock low">Stock: ${stock}</span>`;
  else if (stock !== null) stockBadge = `<span class="tile-stock ok">Stock: ${stock}</span>`;

  const cat = p.CATEGORIA ? escapeHtml(String(p.CATEGORIA).trim()) : "";
  const imagenUrl = p.IMAGEN ? String(p.IMAGEN).trim() : "";

  return `
    <div
      class="product-tile${esPineado ? " product-tile-pinned" : ""}"
      role="button"
      tabindex="0"
      ${dataIdxAttr}>
      ${esPineado ? `<span class="tile-pinned-badge">Último</span>` : ""}
      <div class="tile-photo">
        ${imagenUrl
          ? `<img src="${escapeHtml(imagenUrl)}" alt="" loading="lazy" onerror="this.parentElement.innerHTML='🛒';">`
          : "🛒"}
      </div>
      <div class="tile-info">
        <span class="tile-code">${escapeHtml(p.CODIGO)}</span>
        <span class="tile-name">${escapeHtml(p.PRODUCTO)}</span>
        ${cat ? `<span class="tile-cat">${cat}</span>` : ""}
      </div>
      <div class="tile-right">
        <span class="tile-price">${typeof depPrecioVista === "function" ? depPrecioVista(p.PRECIO, p, true) : "$" + Number(p.PRECIO || 0).toLocaleString("es-AR")}</span>
        ${stockBadge}
      </div>
      ${obtenerRolActual() === "vendedor" ? "" : `<button type="button" class="tile-edit" title="Editar precio y stock" onclick="event.stopPropagation(); abrirEdicionRapidaPOS('${escapeJsAttr(p.CODIGO)}');">✏️</button>`}
      ${Number(p.UNIDADES_POR_CAJA) > 0 ? `<button type="button" class="tile-caja" title="Agregar 1 caja (${p.UNIDADES_POR_CAJA} uds) a $${Number(p.PRECIO_CAJA || 0).toLocaleString("es-AR")}" onclick="event.stopPropagation(); agregarCajaAlTicket(productosPOS.find(x => String(x.CODIGO)==='${escapeJsAttr(p.CODIGO)}'));">x${p.UNIDADES_POR_CAJA}</button>` : ""}
      <span class="tile-add">+</span>
    </div>`;
}

/** Saca el pineo — se llama al vaciar el ticket o al finalizar la venta, para volver la grilla a su orden por defecto */
function limpiarPineadoPOS() {
  if (!productoPineadoPOS) return;
  productoPineadoPOS = null;
  renderPosGrid(document.getElementById("posBusqueda")?.value.trim() || "");
}

function renderPosGrid(filtroTexto) {
  const grid = document.getElementById("posProductGrid");
  if (!grid) return;

  let lista = productosPOS;
  if (categoriaActivaPOS && categoriaActivaPOS !== "TODAS") {
    lista = lista.filter(p => String(p.CATEGORIA || "").trim() === categoriaActivaPOS);
  }
  if (filtroTexto) {
    lista = lista.filter(p => productoCoincideBusquedaPOS(p, filtroTexto));
  }

  if (lista.length === 0 && !productoPineadoPOS) {
    grid.innerHTML = `
      <div class="pos-empty-state" style="grid-column:1/-1;">
        <div class="ic">🔍</div>
        <strong>Sin resultados</strong>
        <span>Probá con otro código o nombre</span>
      </div>`;
    return;
  }

  // El producto pineado va primero, sin duplicarlo si además calzaba
  // con el filtro/categoría actual — el resto completa hasta el límite
  // de tiles visibles de siempre.
  let visibleList = lista.slice(0, 30);
  if (productoPineadoPOS) {
    visibleList = visibleList.filter(p => String(p.CODIGO).trim() !== String(productoPineadoPOS.CODIGO).trim());
    visibleList = [productoPineadoPOS, ...visibleList].slice(0, 30);
  }

  let html = "";
  visibleList.forEach((p, idx) => {
    const esPineado = !!productoPineadoPOS && String(p.CODIGO).trim() === String(productoPineadoPOS.CODIGO).trim();
    html += _armarTileProductoHTML(p, `data-idx="${idx}"`, esPineado);
  });

  grid.innerHTML = html;

  grid.querySelectorAll(".product-tile[data-idx]").forEach(tile => {
    const idx     = Number(tile.getAttribute("data-idx"));
    const producto = visibleList[idx];
    if (producto) {
      tile.dataset.codigo = producto.CODIGO;
      tile.addEventListener("click", () => agregarProductoPOS(producto.CODIGO, false));
      tile.addEventListener("keydown", e => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); agregarProductoPOS(producto.CODIGO, false); }
      });
    }
  });
}

/** Filtra la grilla del POS en tiempo real mientras se escribe (oninput) */
function onPosInput(e) {
  posTileFocusIdx = -1;
  renderPosGridConDemora(e.target.value.trim());
}

function onPosInputKeyup(e) {
  const input = e.target;
  if (e.key === "Enter") {
    const valor = input.value.trim();
    if (valor !== "") {
      agregarProductoPorCodigo(valor);
      input.value = "";
      // Con debounce en vez de render inmediato: si se están
      // escaneando varios productos seguidos (típico en el POS), esto
      // evita reconstruir hasta 60 tarjetas de la grilla en CADA
      // escaneo — solo se reconstruye una vez, 150ms después del
      // último escaneo de la tanda.
      renderPosGridConDemora("");
    }
    return;
  }
  // Las teclas no-Enter ya las maneja onPosInput (oninput)
  // Solo actualizamos el índice de foco por si se usaron flechas
  if (e.key === "ArrowDown" || e.key === "ArrowUp") {
    posTileFocusIdx = -1;
    renderPosGridConDemora(input.value.trim());
  }
}

const renderPosGridConDemora = debounce(renderPosGrid, 150);

/* ---- adding items ---- */

async function agregarProductoPorCodigo(codigo) {
  codigo = String(codigo).trim();
  if (codigo === "") return;
  await asegurarProductosPOS();

  const producto = buscarProductoPOSPorCodigo(codigo);
  if (producto) { agregarProductoPOS(producto.CODIGO); return; }

  // No matcheó el código unitario — probar si es el código de la caja/bulto cerrado
  const productoPorCaja = buscarProductoPOSPorCodigoCaja(codigo);
  if (productoPorCaja) { agregarCajaAlTicket(productoPorCaja); return; }

  toast(`Producto no encontrado: ${codigo}`, "error");
}

/**
 * Agrega al ticket una "caja" (bulto cerrado) de un producto que se
 * vende también por unidad — mismo CODIGO real (para que el stock se
 * descuente del mismo lugar de siempre), pero como línea aparte
 * (marcada con _esCaja) porque el precio por unidad equivalente es
 * distinto al de venderlo suelto.
 */
function agregarCajaAlTicket(producto) {
  const unidades = Number(producto.UNIDADES_POR_CAJA) || 1;
  const precioCaja = Number(producto.PRECIO_CAJA) || 0;
  const precioUnitarioEquivalente = precioCaja / unidades;

  const existente = ticketPOS.find(item => String(item.CODIGO).trim() === String(producto.CODIGO).trim() && item._esCaja);
  if (existente) {
    existente.cantidad += unidades;
  } else {
    ticketPOS.push({
      CODIGO: producto.CODIGO,
      PRODUCTO: `${producto.PRODUCTO} (caja x${unidades})`,
      cantidad: unidades,
      PRECIO: precioUnitarioEquivalente,
      _esCaja: true
    });
  }
  renderTicketPOS();
  toast(`+1 caja de ${producto.PRODUCTO} (${unidades} uds) — $${precioCaja.toLocaleString("es-AR")}`, "success");
}

/**
 * Edición rápida de precio y stock desde el mismo grid del POS — sin
 * tener que ir a la sección Productos. Se preservan categoría,
 * imagen, publicado, destacado y oferta tal cual estaban: la acción
 * del backend (actualizarProducto) sobreescribe TODO el producto, así
 * que si no se mandaran esos valores existentes quedarían borrados.
 */
let _codigoEdicionRapidaPOS = null;

function abrirEdicionRapidaPOS(codigo) {
  if (obtenerRolActual() === "vendedor") {
    toast("Tu usuario (Vendedor) no tiene permiso para modificar precio o stock", "error");
    return;
  }
  const producto = productosPOS.find(p => String(p.CODIGO).trim() === String(codigo).trim());
  if (!producto) { toast("Producto no encontrado", "error"); return; }

  _codigoEdicionRapidaPOS = codigo;
  document.getElementById("erpNombreProducto").textContent = producto.PRODUCTO;
  document.getElementById("erpCodigo").value = producto.CODIGO;
  document.getElementById("erpPrecio").value = producto.PRECIO ?? 0;
  document.getElementById("erpStock").value = producto.STOCK ?? 0;
  document.getElementById("erpAlias").value = producto.ALIAS || "";
  document.getElementById("modalEdicionRapidaPOSBackdrop").classList.add("show");
}

function cerrarEdicionRapidaPOS() {
  document.getElementById("modalEdicionRapidaPOSBackdrop").classList.remove("show");
  _codigoEdicionRapidaPOS = null;
}

async function guardarEdicionRapidaPOS() {
  const codigoOriginal = _codigoEdicionRapidaPOS;
  if (!codigoOriginal) return;

  const producto = productosPOS.find(p => String(p.CODIGO).trim() === String(codigoOriginal).trim());
  if (!producto) { toast("Producto no encontrado", "error"); cerrarEdicionRapidaPOS(); return; }

  const codigoNuevo = document.getElementById("erpCodigo").value.trim();
  const precio = document.getElementById("erpPrecio").value;
  const stock  = document.getElementById("erpStock").value;
  const alias  = document.getElementById("erpAlias").value.trim();

  if (!codigoNuevo) { toast("El código no puede quedar vacío", "error"); return; }
  if (precio === "" || Number(precio) < 0) { toast("Ingresá un precio válido", "error"); return; }
  if (stock === "") { toast("Ingresá un stock válido", "error"); return; }

  // Si cambiaron el código, chequeo local rápido para avisar antes de
  // mandarlo — el backend igual lo vuelve a validar (por si otra
  // persona creó ese código justo en el medio).
  const conflictoErp = conflictoCodigoProductoLocal(codigoNuevo, "", codigoOriginal);
  if (conflictoErp) {
    toast(conflictoErp, "error");
    return;
  }

  const btn = document.getElementById("btnGuardarEdicionRapidaPOS");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    // Se mandan TODOS los campos del producto (no solo código/precio/
    // stock), preservando lo que ya tenía — incluyendo el precio por
    // caja, que este modal no permite editar pero tampoco debe borrar.
    const params = new URLSearchParams({
      action: "actualizarProducto",
      rol: obtenerRolActual(),
      codigoOriginal: producto.CODIGO,
      CODIGO: codigoNuevo,
      PRODUCTO: producto.PRODUCTO,
      CATEGORIA: producto.CATEGORIA || "",
      ALIAS: alias,
      PRECIO: precio,
      STOCK: stock,
      IMAGEN: producto.IMAGEN || "",
      PUBLICADO: producto.PUBLICADO || "SI",
      DESTACADO: producto.DESTACADO || "NO",
      OFERTA: producto.OFERTA || "NO",
      CODIGO_CAJA: producto.CODIGO_CAJA || "",
      UNIDADES_POR_CAJA: producto.UNIDADES_POR_CAJA || 0,
      PRECIO_CAJA: producto.PRECIO_CAJA || 0
    });

    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo actualizar el producto", "error");
      return;
    }

    // Actualiza en memoria al instante — no hace falta esperar a
    // recargar todo el catálogo para ver el cambio reflejado
    producto.CODIGO = codigoNuevo;
    producto.PRECIO = Number(precio);
    producto.STOCK = Number(stock);
    producto.ALIAS = alias;
    // El código o el alias pueden haber cambiado: invalidar el texto
    // de búsqueda cacheado para que no quede buscando por el dato viejo.
    producto._textoBusquedaPOS = undefined;
    producto._palabrasBusquedaPOS = undefined;
    // Si el producto editado ya estaba en el ticket actual, el código
    // ahí también tiene que actualizarse — si no, quedaría apuntando
    // a un código que ya no existe más en productosPOS.
    const itemEnTicket = ticketPOS.find(i => String(i.CODIGO).trim() === String(codigoOriginal).trim());
    if (itemEnTicket) itemEnTicket.CODIGO = codigoNuevo;

    renderPosGrid(document.getElementById("posBusqueda")?.value.trim() || "");
    if (itemEnTicket) renderTicketPOS();

    toast("Producto actualizado", "success");
    cerrarEdicionRapidaPOS();
    invalidarCache("productosAdmin");

  } catch (error) {
    console.error("Error en edición rápida de producto:", error);
    toast("Error de conexión al actualizar el producto", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

function agregarProductoPOS(codigo, pinear = true) {
  codigo = String(codigo).trim();
  const producto = buscarProductoPOSPorCodigo(codigo);
  if (!producto) { toast("Producto no encontrado", "error"); return; }
  const stock = producto.STOCK !== undefined ? Number(producto.STOCK) : null;
  if (stock !== null && stock <= 0) {
    // Antes esto bloqueaba la venta directamente. Ahora se permite —
    // hay negocios que venden sobre pedido, o el stock cargado no
    // siempre está perfectamente al día — pero se avisa para que
    // quede claro que se está vendiendo sin stock disponible.
    toast(`${producto.PRODUCTO} no tiene stock cargado — se agrega igual`, "error");
  }

  const existente = ticketPOS.find(item => String(item.CODIGO).trim() === codigo);
  if (existente) {
    existente.cantidad++;
  } else {
    ticketPOS.push({ CODIGO: producto.CODIGO, PRODUCTO: producto.PRODUCTO, PRECIO: Number(producto.PRECIO || 0), cantidad: 1 });
  }
  renderTicketPOS();
  flashTile(codigo);
  mostrarUltimoEscaneado(producto, pinear);
  ultimoCodigoAgregadoPOS = codigo;
  scrollTicketAlProducto(codigo);

  const input = document.getElementById("posBusqueda");
  if (input) { input.value = ""; input.focus(); }
}

/** Scrolls the ticket list so the row for this product is visible (used when adding/incrementing it) */
function scrollTicketAlProducto(codigo) {
  const fila = document.querySelector(`.ticket-row[data-codigo="${cssEscape(codigo)}"]`);
  if (fila) fila.scrollIntoView({ block: "end", behavior: "smooth" });
}

/** Updates the "last scanned product" panel with its photo, name, code and price */
function mostrarUltimoEscaneado(producto, pinear = true) {
  const panel = document.getElementById("scanResultPanel");
  if (!panel) return;

  // El pineo (mover este producto al primer tile de la grilla, ver
  // limpiarPineadoPOS) solo se hace cuando el producto se agregó por
  // escaneo/código tipeado (pinear=true, el valor por defecto): ahí
  // tiene sentido, porque el usuario no está mirando/clickeando la
  // grilla, así que reordenarla para tener el último escaneado a mano
  // no interfiere con nada.
  //
  // Cuando se agrega clickeando directamente un tile de la grilla
  // (pinear=false, ver el listener de click más abajo), NO se pinea ni
  // se reordena: si se moviera el producto a la posición 0 en cada
  // click, un segundo click rápido en el mismo lugar de la pantalla
  // terminaba clickeando OTRO producto (el que quedó bajo el dedo/
  // cursor tras el reorden). La grilla debe quedarse quieta para poder
  // clickear el mismo producto varias veces seguidas.
  if (pinear) {
    const yaEstabaPineado = productoPineadoPOS &&
      String(productoPineadoPOS.CODIGO).trim() === String(producto.CODIGO).trim();
    productoPineadoPOS = producto;
    if (!yaEstabaPineado) {
      renderPosGridConDemora(document.getElementById("posBusqueda")?.value.trim() || "");
    }
  }

  const thumb = document.getElementById("scanResultThumb");
  const imagenUrl = producto.IMAGEN ? String(producto.IMAGEN).trim() : "";

  if (thumb) {
    thumb.innerHTML = imagenUrl
      ? `<img src="${escapeHtml(imagenUrl)}" alt="" onerror="this.parentElement.innerHTML='🛒';">`
      : "🛒";
  }

  actualizarElemento("scanResultName", producto.PRODUCTO);
  actualizarElemento("scanResultCode", producto.CODIGO);
  actualizarElemento("scanResultPrice", "$" + Number(producto.PRECIO || 0).toLocaleString("es-AR"));

  panel.classList.add("has-product");
}

function flashTile(codigo) {
  const tile = document.querySelector(`.product-tile[data-codigo="${cssEscape(codigo)}"]`);
  if (!tile) return;
  tile.classList.remove("just-added");
  void tile.offsetWidth;
  tile.classList.add("just-added");
}

function cssEscape(value) {
  return String(value).replace(/["\\]/g, "\\$&");
}

function cambiarCantidadPOS(codigo, delta) {
  const item = ticketPOS.find(i => String(i.CODIGO).trim() === String(codigo).trim());
  if (!item) return;
  item.cantidad += delta;
  if (item.cantidad <= 0) { quitarProductoPOS(codigo); return; }
  renderTicketPOS();
}

/** Lets the cashier type the quantity directly instead of tapping +/- repeatedly */
function actualizarCantidadManualPOS(codigo, valor) {
  const item = ticketPOS.find(i => String(i.CODIGO).trim() === String(codigo).trim());
  if (!item) return;

  let cantidad = parseInt(valor, 10);

  if (isNaN(cantidad) || cantidad <= 0) {
    quitarProductoPOS(codigo);
    return;
  }

  item.cantidad = cantidad;
  renderTicketPOS();
}

/** Selects the whole quantity value on focus, so typing a new number replaces it instead of appending */
function seleccionarCantidadPOS(input) {
  input.select();
}

function quitarProductoPOS(codigo) {
  ticketPOS = ticketPOS.filter(i => String(i.CODIGO).trim() !== String(codigo).trim());
  renderTicketPOS();
}

function vaciarTicketPOS() {
  if (ticketPOS.length === 0) return;
  // Sin confirm() — acción recuperable (el cajero puede volver a agregar productos)
  ticketPOS = [];
  limpiarPineadoPOS();
  resetearDescuentoPOS();
  const inputRecibido = document.getElementById("inputRecibido");
  const cambioValor = document.getElementById("cambioValor");
  if (inputRecibido) inputRecibido.value = "";
  if (cambioValor) { cambioValor.textContent = "—"; cambioValor.classList.remove("negativo"); }
  renderTicketPOS();
}

/* ---- ticket rendering ---- */

function renderTicketPOS() {
  let html = "";
  let subtotal = 0;
  let totalItems = 0;

  ticketPOS.forEach(item => {
    const sub = item.PRECIO * item.cantidad;
    subtotal += sub;
    totalItems += item.cantidad;
    html += `
      <div class="ticket-row" data-codigo="${escapeHtml(item.CODIGO)}">
        <div class="ti-info">
          <div class="ti-name">${escapeHtml(item.PRODUCTO)}</div>
          <div class="ti-price">${typeof depPrecioVista === "function" ? depPrecioVista(item.PRECIO, item, false) : "$" + item.PRECIO.toLocaleString("es-AR")} c/u</div>
        </div>
        <div class="ti-qty">
          <input
            type="number"
            class="qty-input-pos"
            min="1"
            step="1"
            inputmode="numeric"
            value="${item.cantidad}"
            onfocus="seleccionarCantidadPOS(this)"
            onchange="actualizarCantidadManualPOS('${escapeJsAttr(item.CODIGO)}', this.value)">
        </div>
        <div class="ti-sub money">$${sub.toLocaleString("es-AR")}</div>
        <button class="ti-remove" onclick="quitarProductoPOS('${escapeJsAttr(item.CODIGO)}')" title="Quitar">✕</button>
      </div>`;
  });

  const tabla = document.getElementById("ticketPOS");
  if (tabla) tabla.innerHTML = html;

  const emptyState = document.getElementById("ticketEmptyState");
  if (emptyState) emptyState.style.display = ticketPOS.length === 0 ? "flex" : "none";

  const { montoDescuento, total } = calcularDescuentoPOS(subtotal);

  const subtotalEl = document.getElementById("subtotalPOS");
  if (subtotalEl) subtotalEl.innerText = subtotal.toLocaleString("es-AR");

  const rowDescuentoEl = document.getElementById("rowDescuentoPOS");
  const descuentoMontoEl = document.getElementById("descuentoMontoPOS");
  const descuentoLabelEl = rowDescuentoEl ? rowDescuentoEl.querySelector("span:first-child") : null;
  if (rowDescuentoEl && descuentoMontoEl) {
    if (descuentoActivoPOS && montoDescuento !== 0) {
      rowDescuentoEl.style.display = "flex";
      if (montoDescuento > 0) {
        if (descuentoLabelEl) descuentoLabelEl.textContent = "Descuento";
        descuentoMontoEl.style.color = "var(--red-500)";
        descuentoMontoEl.innerText = "-$" + montoDescuento.toLocaleString("es-AR");
      } else {
        if (descuentoLabelEl) descuentoLabelEl.textContent = "Recargo";
        descuentoMontoEl.style.color = "var(--green-600)";
        descuentoMontoEl.innerText = "+$" + Math.abs(montoDescuento).toLocaleString("es-AR");
      }
    } else {
      rowDescuentoEl.style.display = "none";
    }
  }

  const totalEl = document.getElementById("totalPOS");
  if (totalEl) totalEl.innerText = total.toLocaleString("es-AR");
  if (typeof depTotalesTicketUsd === "function") depTotalesTicketUsd(ticketPOS, total);

  const countEl = document.getElementById("ticketCount");
  if (countEl) countEl.innerText = ticketPOS.length;

  const itemsCountEl = document.getElementById("ticketItemsCount");
  if (itemsCountEl) itemsCountEl.innerText = totalItems;

  const finalizeBtn = document.getElementById("btnFinalizarVenta");
  if (finalizeBtn) finalizeBtn.disabled = ticketPOS.length === 0;

  const generarPedidoBtn = document.getElementById("btnGenerarPedido");
  if (generarPedidoBtn) generarPedidoBtn.disabled = ticketPOS.length === 0;

  const printBtn = document.getElementById("btnImprimirTicket");
  if (printBtn) printBtn.disabled = ticketPOS.length === 0;

  actualizarUIDescuentoPOS();
}

/* ---- discount ---- */

/**
 * Computes the discount/surcharge amount and the resulting total.
 * `montoDescuento` is the signed amount SUBTRACTED from the subtotal to get
 * the total: positive for a discount (total = subtotal - montoDescuento),
 * negative for a surcharge/"recargo" (total = subtotal + |montoDescuento|).
 */
function calcularDescuentoPOS(subtotal) {
  if (!descuentoActivoPOS || descuentoValorPOS <= 0) {
    return { montoDescuento: 0, total: subtotal };
  }

  let montoBruto = descuentoTipoPOS === "PORCENTAJE"
    ? subtotal * (descuentoValorPOS / 100)
    : descuentoValorPOS;
  montoBruto = Math.max(0, montoBruto); // nunca negativo

  if (ajusteModoPOS === "RECARGO") {
    // El recargo se suma sobre el subtotal; no tiene tope superior.
    return { montoDescuento: -montoBruto, total: subtotal + montoBruto };
  }

  const monto = Math.min(montoBruto, subtotal); // el descuento no puede superar el subtotal
  return { montoDescuento: monto, total: subtotal - monto };
}

/** Calcula y muestra el cambio a entregar según lo recibido */
function calcularCambio() {
  const totalEl = document.getElementById("totalPOS");
  const recibidoEl = document.getElementById("inputRecibido");
  const cambioEl = document.getElementById("cambioValor");
  if (!totalEl || !recibidoEl || !cambioEl) return;

  const total = Number(String(totalEl.textContent).replace(/\./g, "").replace(",", ".")) || 0;
  const recibido = Number(recibidoEl.value) || 0;

  if (!recibido) {
    cambioEl.textContent = "—";
    cambioEl.classList.remove("negativo");
    return;
  }

  const cambio = recibido - total;
  cambioEl.textContent = "$" + Math.abs(cambio).toLocaleString("es-AR");
  if (cambio < 0) {
    cambioEl.classList.add("negativo");
    cambioEl.textContent = "-$" + Math.abs(cambio).toLocaleString("es-AR");
  } else {
    cambioEl.classList.remove("negativo");
  }
}

/** Devuelve el monto recibido y el cambio para incluirlos en el ticket */
function obtenerDatosCambio() {
  const recibidoEl = document.getElementById("inputRecibido");
  const cambioEl = document.getElementById("cambioValor");
  const recibido = recibidoEl ? Number(recibidoEl.value) || 0 : 0;
  const cambioTexto = cambioEl ? cambioEl.textContent : "—";
  return { recibido, cambioTexto };
}

function toggleDescuentoPOS() {
  const panel = document.getElementById("discountPanel");
  if (!panel) return;
  const abierto = panel.style.display !== "none";
  panel.style.display = abierto ? "none" : "flex";
  if (!abierto) {
    setTimeout(() => {
      const input = document.getElementById("descuentoValorInput");
      if (input) input.focus();
    }, 50);
  }
}

function elegirTipoDescuento(el, tipo) {
  descuentoTipoPOS = tipo;
  document.querySelectorAll(".discount-type-btn").forEach(b => b.classList.remove("active"));
  el.classList.add("active");

  const input = document.getElementById("descuentoValorInput");
  if (input) input.placeholder = tipo === "PORCENTAJE" ? "Ej: 10" : "Ej: 500";

  aplicarDescuentoPOS();
}

function aplicarDescuentoPOS() {
  const input = document.getElementById("descuentoValorInput");
  const valor = Number(input ? input.value : 0) || 0;

  descuentoValorPOS = Math.max(0, valor);
  descuentoActivoPOS = descuentoValorPOS > 0;

  renderTicketPOS();
}

function quitarDescuentoPOS() {
  resetearDescuentoPOS();
  renderTicketPOS();
}

function resetearDescuentoPOS() {
  descuentoActivoPOS = false;
  descuentoValorPOS = 0;
  ajusteModoPOS = "DESCUENTO";
  const input = document.getElementById("descuentoValorInput");
  if (input) input.value = "";
  const motivo = document.getElementById("descuentoMotivoInput");
  if (motivo) motivo.value = "";
}

/** Returns a short human-readable label for the active discount/recargo, e.g. "10% (-$500)" or "+$300" */
function obtenerEtiquetaDescuentoPOS(subtotal) {
  if (!descuentoActivoPOS || descuentoValorPOS <= 0) return "";
  const { montoDescuento } = calcularDescuentoPOS(subtotal);
  const motivo = (document.getElementById("descuentoMotivoInput") || {}).value || "";
  const signo = montoDescuento > 0 ? "-" : "+";
  const montoAbs = Math.round(Math.abs(montoDescuento)).toLocaleString("es-AR");
  let etiqueta = descuentoTipoPOS === "PORCENTAJE"
    ? `${descuentoValorPOS}% (${signo}$${montoAbs})`
    : `${signo}$${montoAbs}`;
  if (motivo.trim()) etiqueta += ` — ${motivo.trim()}`;
  return etiqueta;
}

/** Keeps the toggle button label/style in sync with the current discount/recargo state */
function actualizarUIDescuentoPOS() {
  const btn = document.getElementById("btnToggleDescuento");
  const label = document.getElementById("discountToggleLabel");
  if (!btn || !label) return;

  if (descuentoActivoPOS && descuentoValorPOS > 0) {
    btn.classList.add("has-discount");
    const palabra = ajusteModoPOS === "RECARGO" ? "Recargo" : "Descuento";
    label.innerText = descuentoTipoPOS === "PORCENTAJE"
      ? `${palabra} aplicado: ${descuentoValorPOS}%`
      : `${palabra} aplicado: $${descuentoValorPOS.toLocaleString("es-AR")}`;
  } else {
    btn.classList.remove("has-discount");
    label.innerText = "Agregar descuento";
  }
}

/* ---- payment method ---- */

function elegirFormaPago(el, valor) {
  formaPagoPOS = valor;
  document.querySelectorAll(".pay-method-btn").forEach(b => b.classList.remove("active"));
  el.classList.add("active");
}

/* ---- finalize sale ---- */

/* ---- Modal Finalizar Venta ---- */
let mfvTipoDescuento = "PORCENTAJE";
let mfvModoDescuento = "DESCUENTO"; // "DESCUENTO" | "RECARGO"
let mfvFormaPago = "EFECTIVO";
let mfvMpDisponible = false; // cacheado al abrir el modal — ¿hay MP configurado y usable?
let mfvGenerarQR = false;    // elección del cajero cuando la forma de pago es TRANSFERENCIA:
                              // true = generar QR para que el cliente escanee y pague ahí mismo;
                              // false (default) = el cliente ya transfirió directo al alias/CVU,
                              // registrar la venta sin pasar por Mercado Pago.

function abrirModalFinalizarVenta() {
  if (ticketPOS.length === 0) { toast("El ticket está vacío", "error"); return; }

  try {
  // Pre-cargar valores del ticket actual
  const subtotal = ticketPOS.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
  const { montoDescuento, total } = calcularDescuentoPOS(subtotal);

  document.getElementById("mfvTotal").textContent = "$" + total.toLocaleString("es-AR");

  // Descuento — prellenar con el valor actual del ticket
  document.getElementById("mfvDescuentoValor").value = descuentoValorPOS || "";
  document.getElementById("mfvDescuentoMotivo").value =
    (document.getElementById("descuentoMotivoInput") || {}).value || "";

  // Tipo de descuento
  mfvTipoDescuento = descuentoTipoPOS || "PORCENTAJE";
  document.getElementById("mfvTipoPct").classList.toggle("active", mfvTipoDescuento === "PORCENTAJE");
  document.getElementById("mfvTipoMonto").classList.toggle("active", mfvTipoDescuento === "MONTO");

  // Modo: descuento o recargo
  mfvModoDescuento = ajusteModoPOS || "DESCUENTO";
  document.getElementById("mfvModoDescuento").classList.toggle("active", mfvModoDescuento === "DESCUENTO");
  document.getElementById("mfvModoRecargo").classList.toggle("active", mfvModoDescuento === "RECARGO");

  mfvActualizarDescuento();

  // Forma de pago (en dólares solo en el depósito)
  mfvFormaPago = formaPagoPOS || "EFECTIVO";
  if (mfvFormaPago === "DOLARES" && !(typeof esModoDeposito === "function" && esModoDeposito())) mfvFormaPago = "EFECTIVO";
  const usdRecibidoEl = document.getElementById("mfvUsdRecibido");
  if (usdRecibidoEl) { usdRecibidoEl.value = ""; delete usdRecibidoEl.dataset.tocado; }
  document.querySelectorAll("#mfvPayMethods .pay-method-btn").forEach(b => {
    b.classList.toggle("active", b.dataset.val === mfvFormaPago);
  });
  mfvMostrarRecibido();

  // QR de Mercado Pago para Transferencia — se resetea a "sin QR" cada vez
  // que se abre el modal, y se consulta si MP está disponible en segundo
  // plano (no bloquea la apertura del modal) para decidir si mostrar el
  // selector o no.
  mfvGenerarQR = false;
  mfvMpDisponible = false;
  mfvActualizarVisibilidadQR();
  verificarMpDisponibleParaModal();

  // Recibido
  document.getElementById("mfvRecibido").value = "";
  document.getElementById("mfvCambio").textContent = "—";
  document.getElementById("mfvCambio").classList.remove("negativo");

  document.getElementById("btnConfirmarVenta").disabled = false;
  document.getElementById("btnConfirmarVenta").textContent = "Confirmar venta";
  if (window.ArcaUI) ArcaUI.alAbrirModalVenta(); // "Facturar" (facturacion-arca.js)
  document.getElementById("modalFinalizarVentaBackdrop").classList.add("show");
  setTimeout(() => {
    const el = mfvFormaPago === "EFECTIVO"
      ? document.getElementById("mfvRecibido")
      : document.getElementById("btnConfirmarVenta");
    if (el) el.focus();
  }, 120);
  } catch(err) {
    console.error("Error al abrir modal finalizar:", err);
    // Fallback: abrir el modal igual aunque haya error de prefill
    document.getElementById("modalFinalizarVentaBackdrop").classList.add("show");
  }
}

function cerrarModalFinalizarVenta() {
  document.getElementById("modalFinalizarVentaBackdrop").classList.remove("show");
}

function mfvElegirPago(btn, valor) {
  mfvFormaPago = valor;
  document.querySelectorAll("#mfvPayMethods .pay-method-btn").forEach(b =>
    b.classList.toggle("active", b.dataset.val === valor));
  mfvMostrarRecibido();
  mfvActualizarVisibilidadQR();
  if (valor === "EFECTIVO") setTimeout(() => document.getElementById("mfvRecibido").focus(), 50);
  if (valor === "DOLARES") setTimeout(() => { const el = document.getElementById("mfvUsdRecibido"); if (el) { el.focus(); el.select(); } }, 50);
}

/** Consulta (una sola vez por apertura del modal) si Mercado Pago está configurado y usable, sin bloquear la apertura del modal */
async function verificarMpDisponibleParaModal() {
  const bridge = window.veekpos || window.posOffline;
  if (!bridge || typeof bridge.mercadoPagoDisponible !== "function") {
    mfvMpDisponible = false;
  } else {
    try {
      mfvMpDisponible = await bridge.mercadoPagoDisponible();
    } catch (error) {
      mfvMpDisponible = false;
    }
  }
  mfvActualizarVisibilidadQR();
}

/** Muestra/oculta el selector "Generar QR / Ya me transfirió" según la forma de pago elegida y si MP está disponible */
function mfvActualizarVisibilidadQR() {
  const wrap = document.getElementById("mfvQrWrap");
  if (!wrap) return;

  const mostrar = mfvFormaPago === "TRANSFERENCIA" && mfvMpDisponible;
  wrap.style.display = mostrar ? "block" : "none";
  if (!mostrar) mfvGenerarQR = false; // si se oculta (no hay MP, o cambiaron de forma de pago), nunca queda "generar QR" activo por error

  const btnSi = document.getElementById("mfvQrSi");
  const btnNo = document.getElementById("mfvQrNo");
  if (btnSi) btnSi.classList.toggle("active", mfvGenerarQR);
  if (btnNo) btnNo.classList.toggle("active", !mfvGenerarQR);
}

function mfvElegirGenerarQR(valor) {
  mfvGenerarQR = valor;
  mfvActualizarVisibilidadQR();
}

function mfvMostrarRecibido() {
  const wrap = document.getElementById("mfvRecibidoWrap");
  if (wrap) wrap.style.display = mfvFormaPago === "EFECTIVO" ? "block" : "none";
  const usd = document.getElementById("mfvUsdWrap");
  if (usd) usd.style.display = mfvFormaPago === "DOLARES" ? "block" : "none";
  if (mfvFormaPago === "DOLARES" && typeof depMfvPreparar === "function") depMfvPreparar();
}

function mfvElegirTipo(btn, tipo) {
  mfvTipoDescuento = tipo;
  document.querySelectorAll("#mfvPayMethods .discount-type-btn, .product-modal .discount-type-btn").forEach(b => {
    if (b.dataset.tipo) b.classList.toggle("active", b.dataset.tipo === tipo);
  });
  btn.closest(".product-modal-body").querySelectorAll(".discount-type-btn").forEach(b =>
    b.classList.toggle("active", b.dataset.tipo === tipo));
  mfvActualizarDescuento();
}

function mfvElegirModo(btn, modo) {
  mfvModoDescuento = modo;
  btn.closest(".product-modal-body").querySelectorAll(".discount-mode-btn").forEach(b =>
    b.classList.toggle("active", b.dataset.modo === modo));
  mfvActualizarDescuento();
}

function mfvLimpiarDescuento() {
  document.getElementById("mfvDescuentoValor").value = "";
  document.getElementById("mfvDescuentoMotivo").value = "";
  mfvModoDescuento = "DESCUENTO";
  document.getElementById("mfvModoDescuento").classList.add("active");
  document.getElementById("mfvModoRecargo").classList.remove("active");
  mfvActualizarDescuento();
}

function mfvActualizarDescuento() {
  const subtotal = ticketPOS.reduce((a, i) => a + i.PRECIO * i.cantidad, 0);
  const val = Number(document.getElementById("mfvDescuentoValor").value) || 0;
  let montoBruto = 0;
  if (mfvTipoDescuento === "PORCENTAJE") montoBruto = subtotal * val / 100;
  else montoBruto = val;
  montoBruto = Math.max(0, montoBruto);

  const esRecargo = mfvModoDescuento === "RECARGO";
  const monto = esRecargo ? montoBruto : Math.min(montoBruto, subtotal);
  const total = esRecargo ? subtotal + monto : subtotal - monto;

  document.getElementById("mfvTotal").textContent = "$" + total.toLocaleString("es-AR");
  if (typeof depMfvTituloUsd === "function") depMfvTituloUsd(total);
  const info = document.getElementById("mfvDescuentoInfo");
  if (monto > 0) {
    info.style.display = "block";
    info.style.color = esRecargo ? "var(--green-600)" : "var(--red-500)";
    const signo = esRecargo ? "+" : "-";
    const etiqueta = esRecargo ? "Recargo" : "Descuento";
    info.textContent = `${etiqueta}: ${signo}$${monto.toLocaleString("es-AR")} → Total: $${total.toLocaleString("es-AR")}`;
  } else {
    info.style.display = "none";
  }
  if (mfvFormaPago === "EFECTIVO") mfvCalcularCambio();
  if (mfvFormaPago === "DOLARES" && typeof depMfvActualizar === "function") depMfvActualizar();
}

function mfvCalcularCambio() {
  const totalStr = document.getElementById("mfvTotal").textContent.replace(/\$/g, "").replace(/\./g, "").replace(",", ".");
  const total = Number(totalStr) || 0;
  const recibido = Number(document.getElementById("mfvRecibido").value) || 0;
  const cambioEl = document.getElementById("mfvCambio");
  if (!recibido) { cambioEl.textContent = "—"; cambioEl.classList.remove("negativo"); return; }
  const cambio = recibido - total;
  cambioEl.textContent = "$" + Math.abs(cambio).toLocaleString("es-AR");
  cambioEl.classList.toggle("negativo", cambio < 0);
  if (cambio < 0) cambioEl.textContent = "-$" + Math.abs(cambio).toLocaleString("es-AR");
}

let _ventaEnProceso = false; // evita doble confirmación por clic rápido

async function confirmarFinalizarVenta() {
  if (_ventaEnProceso) return; // ya hay una venta en curso
  _ventaEnProceso = true;
  const btnConfirmar = document.getElementById("btnConfirmarVenta");
  if (btnConfirmar) { btnConfirmar.disabled = true; btnConfirmar.textContent = "Procesando..."; }

  try {
  const valDescuento = document.getElementById("mfvDescuentoValor").value;
  const motivoDescuento = document.getElementById("mfvDescuentoMotivo").value;
  if (valDescuento && Number(valDescuento) > 0) {
    descuentoTipoPOS = mfvTipoDescuento;
    ajusteModoPOS = mfvModoDescuento;
    if (document.getElementById("descuentoValorInput"))
      document.getElementById("descuentoValorInput").value = valDescuento;
    if (document.getElementById("descuentoMotivoInput"))
      document.getElementById("descuentoMotivoInput").value = motivoDescuento;
    aplicarDescuentoPOS();
  } else {
    resetearDescuentoPOS();
  }

  // Cobro en dólares: se valida antes de cerrar el modal
  let pagoUsd = null;
  if (mfvFormaPago === "DOLARES") {
    const subtotalUsd = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
    pagoUsd = typeof depMfvLeerPago === "function" ? depMfvLeerPago(calcularDescuentoPOS(subtotalUsd).total) : null;
    if (!pagoUsd || pagoUsd.error) { toast((pagoUsd && pagoUsd.error) || "Revisá el cobro en dólares", "error"); return; }
  }

  const btnPago = document.querySelector(`#mfvPayMethods .pay-method-btn[data-val="${mfvFormaPago}"]`);
  if (btnPago) elegirFormaPago(btnPago, mfvFormaPago);

  const recibido = Number(document.getElementById("mfvRecibido").value) || 0;
  const inputRec = document.getElementById("inputRecibido");
  if (inputRec) inputRec.value = recibido || "";
  if (recibido) calcularCambio();

  // Factura electrónica: si faltan datos del cliente, el modal sigue abierto
  if (window.ArcaUI) {
    const subtotalArca = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
    if (!ArcaUI.capturarModal(calcularDescuentoPOS(subtotalArca).total)) return;
  }

  cerrarModalFinalizarVenta();

  // Subtotal/total (con descuento o recargo ya aplicado) — se calcula ACÁ,
  // antes de decidir el camino de QR o el directo, porque ambos lo
  // necesitan. Antes estaba declarado más abajo (en la sección
  // "OPTIMISTIC"), y la rama de QR lo usaba desde arriba sin que existiera
  // todavía — un ReferenceError garantizado ("Cannot access 'total' before
  // initialization") cada vez que se intentaba generar el QR.
  const subtotal = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
  const { montoDescuento, total } = calcularDescuentoPOS(subtotal);
  const etiquetaDescuento = obtenerEtiquetaDescuentoPOS(subtotal);

  // Si es TRANSFERENCIA y el cajero eligió generar el QR (ver selector
  // "¿Cómo se cobra la transferencia?" en el modal) → mostrar QR antes de
  // registrar. Si no lo eligió (o MP no está disponible), la venta sigue
  // de largo y se registra directo como transferencia manual — antes esto
  // era obligatorio siempre que MP estuviera configurado, y si la
  // generación del QR fallaba por cualquier motivo, la venta quedaba
  // completamente bloqueada sin forma de simplemente registrarla.
  if (mfvFormaPago === "TRANSFERENCIA" && mfvGenerarQR) {
    const bridge = window.veekpos || window.posOffline;
    if (bridge && typeof bridge.mercadoPagoDisponible === "function") {
      const mpDisponible = await bridge.mercadoPagoDisponible().catch(() => false);
      if (mpDisponible) {
        await iniciarCobroMercadoPago(total);
        return; // iniciarCobroMercadoPago maneja el resto del flujo
      }
    }
  }

  // ── OPTIMISTIC: calcular todo localmente y mostrar el recibo al instante ──
  const itemsSnapshot = [...ticketPOS];
  const fechaVenta = new Date();
  const ventaIdTemp = "VEN-" + Date.now().toString().slice(-6);
  // Generado ACÁ, antes de intentar guardar — así, si hay que
  // reintentar más tarde (ver colaVentasPendientes), el backend puede
  // reconocer que es el mismo intento y no duplicar la venta.
  const clienteVentaId = "CVL-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);

  // Guardar para impresión
  ultimaVentaImprimible = {
    ventaId: ventaIdTemp,
    items: itemsSnapshot,
    total, subtotal,
    descuento: montoDescuento,
    descuentoEtiqueta: etiquetaDescuento,
    formaPago: formaPagoPOS,
    fecha: fechaVenta,
    recibido: formaPagoPOS === "EFECTIVO" ? recibido : 0, // para "Recibido / Cambio" en el ticket
    pagoUsd: pagoUsd || null
  };

  // Mostrar recibo y limpiar ticket INMEDIATAMENTE
  mostrarRecibo(ventaIdTemp, itemsSnapshot, total, subtotal, montoDescuento, pagoUsd ? 0 : recibido, pagoUsd);
  if (window.ArcaUI) ArcaUI.trasVenta({ ref: clienteVentaId, items: itemsSnapshot, total, fecha: fechaVenta });
  ticketPOS = [];
  limpiarPineadoPOS();
  resetearDescuentoPOS();
  if (inputRec) inputRec.value = "";
  const cambioEl = document.getElementById("cambioValor");
  if (cambioEl) { cambioEl.textContent = "—"; cambioEl.classList.remove("negativo"); }
  renderTicketPOS();
  ultimoCodigoAgregadoPOS = null;
  posTileFocusIdx = -1;

  // Actualizar stock optimistamente en el grid
  itemsSnapshot.forEach(item => {
    const p = productosPOS.find(x => String(x.CODIGO) === String(item.CODIGO));
    if (p && p.STOCK !== undefined) p.STOCK = Math.max(0, Number(p.STOCK) - item.cantidad);
  });
  renderPosGrid();

  const btn = document.getElementById("btnFinalizarVenta");
  if (btn) btn.disabled = false;

  // ── GUARDAR: local-primero en escritorio, POST + cola en la web ──
  // (Antes, en la app de escritorio esto iba DIRECTO a Apps Script y se
  // salteaba la SQLite local, la red multi-caja y la cola unificada.)
  const ventaParaGuardar = {
    clienteVentaId,
    ventaIdTemp,
    total, subtotal,
    descuento: montoDescuento,
    formaPago: formaPagoPOS,
    vendedor: sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("usuarioLogueado") || "",
    observaciones: etiquetaDescuento ? (ajusteModoPOS === "RECARGO" ? "Recargo: " : "Descuento: ") + etiquetaDescuento : "",
    carrito: itemsSnapshot
  };
  if (pagoUsd) {
    ventaParaGuardar.montoUsd = pagoUsd.montoUsd;
    ventaParaGuardar.tipoCambio = pagoUsd.tipoCambio;
    const obsUsd = depTextoPagoUsd(pagoUsd);
    ventaParaGuardar.observaciones = ventaParaGuardar.observaciones ? ventaParaGuardar.observaciones + " · " + obsUsd : obsUsd;
  }
  const resultadoGuardado = await registrarVentaCobrada(ventaParaGuardar);
  if (resultadoGuardado && resultadoGuardado.ventaId && resultadoGuardado.ventaId !== ventaIdTemp) {
    ultimaVentaImprimible.ventaId = resultadoGuardado.ventaId;
    const idEl = document.getElementById("reciboVentaId");
    if (idEl) idEl.textContent = resultadoGuardado.ventaId;
  }

  // Métricas en segundo plano
  setTimeout(() => { cargarMetricas(); invalidarCache("ventasPOS"); }, 500);

  } catch(err) {
    console.error("Error en confirmarFinalizarVenta:", err);
    toast("Error al finalizar la venta", "error");
  } finally {
    _ventaEnProceso = false;
    if (btnConfirmar) { btnConfirmar.disabled = false; btnConfirmar.textContent = "Confirmar venta"; }
  }
}

async function finalizarVentaPOS() {
  // Si se llama directo (sin modal), abre el modal
  abrirModalFinalizarVenta();
}

/* ---- generar pedido desde el POS (no es una venta, no descuenta stock) ---- */

/** Valor del campo "Alícuota de IVA" del formulario de producto, o null si no está a la vista */
function _alicuotaIvaFormulario() {
  const wrap = document.getElementById("pmAlicuotaWrap"), sel = document.getElementById("pmAlicuotaIva");
  if (!wrap || !sel || wrap.style.display === "none") return null;
  return sel.value;
}

/** Opens the "Generar Pedido" modal, pre-filling the current ticket total */
function abrirModalPedidoAdmin() {
  if (ticketPOS.length === 0) { toast("El ticket está vacío, agregá productos primero", "error"); return; }

  const subtotal = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
  const { total } = calcularDescuentoPOS(subtotal);

  // Limpiar formulario
  paLimpiarCliente();
  document.getElementById("paTotal").textContent = "$" + total.toLocaleString("es-AR");
  document.getElementById("paTipoCambioWrap").style.display = "none";
  document.getElementById("paMoneda").value = "ARS";
  document.getElementById("paTipoCambio").value = "";

  // Poblar selector con clientes existentes — cargar si no están en memoria
  const selector = document.getElementById("paClienteSelector");
  if (selector) {
    selector.innerHTML = '<option value="">— Nuevo cliente / ingresar datos —</option>';
    const poblarSelector = (lista) => {
      const ordenados = [...lista].sort((a, b) =>
        (a.NOMBRE || a.CLIENTE || "").localeCompare(b.NOMBRE || b.CLIENTE || ""));
      ordenados.forEach(c => {
        const nombre = c.NOMBRE || c.CLIENTE || "";
        const dni = c.DNI ? ` · ${c.DNI}` : "";
        const opt = document.createElement("option");
        opt.value = c.CLIENTE_ID || c.DNI || "";
        opt.textContent = nombre + dni;
        selector.appendChild(opt);
      });
    };
    if (_clientesPedidoCache.length > 0) {
      poblarSelector(_clientesPedidoCache);
    } else {
      // Cargar todos los clientes de la hoja CLIENTES
      cargarClientesParaPedido().then(lista => { if (lista.length > 0) poblarSelector(lista); });
    }
  }

  document.getElementById("pedidoAdminModalBackdrop").classList.add("show");
  setTimeout(() => {
    const sel = document.getElementById("paClienteSelector");
    if (sel) sel.focus(); else document.getElementById("paNombre").focus();
  }, 80);
}

/** Rellena el formulario con los datos del cliente seleccionado */
function paSeleccionarCliente(clienteId) {
  if (!clienteId) { paLimpiarCliente(); return; }
  const lista = _clientesPedidoCache.length > 0 ? _clientesPedidoCache : clientesGlobal;
  const cliente = lista.find(c =>
    String(c.CLIENTE_ID || c.DNI || "") === String(clienteId));
  if (!cliente) return;
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ""; };
  set("paNombre",       cliente.NOMBRE || cliente.CLIENTE);
  set("paEmpresa",      cliente.EMPRESA || "");
  set("paDireccion",    cliente.DIRECCION || "");
  set("paLocalidad",    cliente.LOCALIDAD || "");
  set("paProvincia",    cliente.PROVINCIA || "");
  set("paCodigoPostal", cliente.CODIGO_POSTAL || "");
  set("paTelefono",     cliente.TELEFONO || "");
  set("paDni",          cliente.DNI || "");
  chequearClienteCreditoPedidoAdmin();
}

/** Limpia el formulario para ingresar datos de cliente nuevo */
function paLimpiarCliente() {
  const sel = document.getElementById("paClienteSelector");
  if (sel) sel.value = "";
  ["paNombre","paEmpresa","paDireccion","paLocalidad","paProvincia",
   "paCodigoPostal","paTelefono","paDni"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
  });
  const monedaWrap = document.getElementById("paMonedaWrap");
  if (monedaWrap) monedaWrap.style.display = "none";
  // Limpiar también el ajuste de crédito
  const creditoEl = document.getElementById("paCreditoPct");
  if (creditoEl) creditoEl.value = "";
  const infoEl = document.getElementById("paCreditoInfo");
  if (infoEl) infoEl.style.display = "none";
}

/** Recalcula el total del pedido aplicando el % de crédito/recargo */
function paActualizarTotalConCredito() {
  const subtotal = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
  const { total: totalBase } = calcularDescuentoPOS(subtotal);

  const pct = Number(document.getElementById("paCreditoPct")?.value) || 0;
  const ajuste = Math.round(totalBase * pct / 100);
  const totalFinal = totalBase + ajuste;

  document.getElementById("paTotal").textContent = "$" + totalFinal.toLocaleString("es-AR");
  // Depósito, pedido en dólares: se muestra lo que va a quedar (en dólares)
  const monedaSel = document.getElementById("paMoneda");
  const wrapMoneda = document.getElementById("paMonedaWrap");
  const tc = Number((document.getElementById("paTipoCambio") || {}).value) || 0;
  if (monedaSel && monedaSel.value === "USD" && wrapMoneda && wrapMoneda.style.display !== "none" && tc > 0 &&
      typeof esModoDeposito === "function" && esModoDeposito() && typeof depConvertirPedidoAUsd === "function") {
    const conv = depConvertirPedidoAUsd(ticketPOS, subtotal, totalFinal, tc);
    document.getElementById("paTotal").textContent = depUsd(conv.total) + "  ($" + totalFinal.toLocaleString("es-AR") + ")";
  }

  const info = document.getElementById("paCreditoInfo");
  if (info) {
    if (pct !== 0) {
      const signo = ajuste >= 0 ? "+" : "";
      const tipo = pct > 0 ? "Recargo" : "Descuento";
      info.textContent = `${tipo} ${Math.abs(pct)}%: ${signo}$${ajuste.toLocaleString("es-AR")} → Total: $${totalFinal.toLocaleString("es-AR")}`;
      info.style.color = pct > 0 ? "var(--red-500)" : "var(--green-600)";
      info.style.display = "block";
    } else {
      info.style.display = "none";
    }
  }
}

function cerrarModalPedidoAdmin() {
  document.getElementById("pedidoAdminModalBackdrop").classList.remove("show");
}

/** Checks if the typed DNI belongs to a client already marked as "a crédito" — if so, shows the currency selector */
async function chequearClienteCreditoPedidoAdmin() {
  const dni = document.getElementById("paDni").value.trim();
  const wrap = document.getElementById("paMonedaWrap");

  if (!dni) {
    wrap.style.display = "none";
    document.getElementById("paTipoCambioWrap").style.display = "none";
    return;
  }

  try {
    const response = await fetchAPI(API_URL + "?action=consultarClienteCreditoPorDni&dni=" + encodeURIComponent(dni));
    const data = await response.json();

    wrap.style.display = data.esCredito ? "block" : "none";
    if (!data.esCredito) {
      document.getElementById("paTipoCambioWrap").style.display = "none";
    } else {
      // Depósito mostrando en dólares: el pedido a crédito se propone en dólares (se puede cambiar)
      if (typeof depVistaUsd === "function" && depVistaUsd()) document.getElementById("paMoneda").value = "USD";
      actualizarVisibilidadTipoCambioPedidoAdmin();
    }
  } catch (error) {
    console.error("Error al consultar cliente a crédito:", error);
    // Un error puntual de red no debe bloquear el pedido — simplemente
    // no se muestra el selector de moneda, y el pedido sigue en ARS.
  }
}

function actualizarVisibilidadTipoCambioPedidoAdmin() {
  const moneda = document.getElementById("paMoneda").value;
  document.getElementById("paTipoCambioWrap").style.display = moneda === "USD" ? "block" : "none";
  _prellenarTipoCambio("paTipoCambio", moneda === "USD");
  if (typeof paActualizarTotalConCredito === "function") paActualizarTotalConCredito();
}

async function confirmarPedidoAdmin() {
  if (ticketPOS.length === 0) { toast("El ticket está vacío", "error"); cerrarModalPedidoAdmin(); return; }

  const nombre = document.getElementById("paNombre").value.trim();
  const empresa = document.getElementById("paEmpresa").value.trim();
  const direccion = document.getElementById("paDireccion").value.trim();
  const localidad = document.getElementById("paLocalidad").value.trim();
  const provincia = document.getElementById("paProvincia").value.trim();
  const codigoPostal = document.getElementById("paCodigoPostal").value.trim();
  const telefono = document.getElementById("paTelefono").value.trim();
  const dni = document.getElementById("paDni").value.trim();

  if (!nombre || !direccion || !localidad || !provincia || !telefono || !dni) {
    toast("Completá Nombre, Dirección, Localidad, Provincia, Teléfono y DNI/CUIT", "error");
    return;
  }

  // Moneda y tipo de cambio solo importan si el selector está visible
  // (cliente a crédito) — para cualquier otro pedido, queda en ARS sin
  // tipo de cambio, igual que siempre.
  const monedaWrapVisible = document.getElementById("paMonedaWrap").style.display !== "none";
  const moneda = monedaWrapVisible ? document.getElementById("paMoneda").value : "ARS";
  const tipoCambio = document.getElementById("paTipoCambio").value.trim();

  if (moneda === "USD" && (!tipoCambio || Number(tipoCambio) <= 0)) {
    toast("Ingresá el tipo de cambio para un pedido en dólares", "error");
    return;
  }

  const subtotal = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
  const { total: totalBase } = calcularDescuentoPOS(subtotal);

  // Aplicar ajuste de crédito si fue ingresado
  const creditoPct = Number(document.getElementById("paCreditoPct")?.value) || 0;
  const ajusteCredito = Math.round(totalBase * creditoPct / 100);
  const total = totalBase + ajusteCredito;

  const btn = document.getElementById("btnConfirmarPedidoAdmin");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Creando pedido...";

  try {
    // Por POST, con el carrito en el body — no por GET con todo en la
    // URL. Un pedido con varios cientos de ítems arma una URL enorme
    // que Google/Apps Script rechaza (esto se sentía como "error de
    // conexión" sin serlo realmente). El body no tiene ese límite.
    // Depósito: un pedido en dólares guarda sus precios y su total EN DÓLARES
    // (la deuda del cliente en dólares se lleva aparte de la de pesos).
    let carritoEnviar = ticketPOS, totalEnviar = total;
    if (moneda === "USD" && typeof esModoDeposito === "function" && esModoDeposito() && typeof depConvertirPedidoAUsd === "function") {
      const conv = depConvertirPedidoAUsd(ticketPOS, subtotal, total, Number(tipoCambio));
      carritoEnviar = conv.carrito; totalEnviar = conv.total;
    }
    const response = await fetchAPI(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // evita preflight CORS contra Apps Script
      body: JSON.stringify({
        action: "guardarPedidoAdmin",
        nombre, empresa, direccion, localidad, provincia, codigoPostal, telefono, dni,
        total: totalEnviar,
        moneda: moneda,
        tipoCambio: moneda === "USD" ? tipoCambio : "",
        carrito: carritoEnviar
      })
    }, { timeoutMs: 30000 }); // pedidos grandes tardan más en subir — margen extra
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo crear el pedido", "error");
      return;
    }

    toast(`Pedido ${data.pedidoId} creado`, "success");
    cerrarModalPedidoAdmin();

    // El pedido no es una venta: el ticket se vacía igual, para que el
    // cajero pueda empezar de cero, pero sin tocar stock ni métricas
    // de ventas (a diferencia de finalizarVentaPOS).
    ticketPOS = [];
    limpiarPineadoPOS();
    resetearDescuentoPOS();
    renderTicketPOS();
    ultimoCodigoAgregadoPOS = null;
    posTileFocusIdx = -1;

  } catch (error) {
    console.error("Error al crear el pedido:", error);
    toast("Error de conexión al crear el pedido", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

/* ---- receipt modal ---- */

function mostrarRecibo(ventaId, items, total, subtotal, montoDescuento, recibido, pagoUsd) {
  document.getElementById("receiptId").innerText = "#" + (ventaId || "—");

  let html = "";
  items.forEach(item => {
    const sub = item.PRECIO * item.cantidad;
    html += `
      <tr>
        <td>${item.cantidad}x ${escapeHtml(item.PRODUCTO)}</td>
        <td style="text-align:right;">$${sub.toLocaleString("es-AR")}</td>
      </tr>`;
  });

  if (montoDescuento !== 0) {
    const esRecargo = montoDescuento < 0;
    const color = esRecargo ? "var(--green-600)" : "var(--red-500)";
    const etiqueta = esRecargo ? "Recargo" : "Descuento";
    const signo = esRecargo ? "+" : "-";
    html += `
      <tr>
        <td>Subtotal</td>
        <td style="text-align:right;">$${Number(subtotal).toLocaleString("es-AR")}</td>
      </tr>
      <tr>
        <td style="color:${color};">${etiqueta}</td>
        <td style="text-align:right;color:${color};">${signo}$${Math.round(Math.abs(montoDescuento)).toLocaleString("es-AR")}</td>
      </tr>`;
  }

  document.getElementById("receiptItems").innerHTML = html;
  document.getElementById("receiptTotal").innerText = "$" + total.toLocaleString("es-AR");

  // Mostrar cambio si se ingresó efectivo
  const cambioWrap = document.getElementById("receiptCambioWrap");
  const cambioEl = document.getElementById("receiptCambio");
  if (cambioWrap && cambioEl && recibido > 0) {
    const cambio = recibido - total;
    cambioEl.textContent = "$" + Math.max(0, cambio).toLocaleString("es-AR");
    cambioWrap.style.display = cambio >= 0 ? "flex" : "none";
  } else if (cambioWrap) {
    cambioWrap.style.display = "none";
  }
  // Cobro en dólares: lo que entró en dólares y el resto (o vuelto) en pesos
  if (cambioWrap && pagoUsd) {
    const etiqueta = cambioWrap.querySelector("span");
    if (etiqueta) etiqueta.textContent = "US$ " + Number(pagoUsd.montoUsd).toLocaleString("es-AR") + " · " + (pagoUsd.restoArs > 0 ? "resto en pesos" : "vuelto en pesos");
    cambioEl.textContent = "$" + Math.abs(pagoUsd.restoArs > 0 ? pagoUsd.restoArs : pagoUsd.vueltoArs).toLocaleString("es-AR");
    cambioWrap.style.display = "flex";
  } else if (cambioWrap) {
    const etiqueta = cambioWrap.querySelector("span");
    if (etiqueta) etiqueta.textContent = "Cambio";
  }

  document.getElementById("receiptBackdrop").classList.add("show");
  toast("Venta registrada con éxito", "success");
}

function cerrarRecibo() {
  document.getElementById("receiptBackdrop").classList.remove("show");
}

function nuevaVentaPOS() {
  cerrarRecibo();
  const input = document.getElementById("posBusqueda");
  if (input) input.focus();
}

/* =====================================================
   THERMAL PRINT — POS80 80mm paper
   Renders a receipt into #thermalPrintFrame and calls
   window.print(). The @media print rule hides everything
   except #thermalPrintFrame, and @page sets 80mm width.
===================================================== */

/**
 * Build the thermal HTML string.
 * @param {string}  ventaId
 * @param {Array}   items   — [{PRODUCTO, PRECIO, cantidad}, …]
 * @param {number}  total
 * @param {string}  formaPago
 * @param {Date}    fecha
 * @param {Object}  [descuento] — { monto, etiqueta } opcional, si la venta tuvo descuento
 * @param {Object}  [cfgOverride] — config del negocio a usar en vez de la guardada (solo para vista previa)
 */
/* ---- Formato común de los tickets (venta y cierre, pantalla y térmica) ---- */

/** "$1.250" si es entero, "$1.250,50" si tiene centavos (antes salía "$1.250,5") */
function plataTicket(n) {
  const v = Math.round((Number(n) || 0) * 100) / 100;
  return (v < 0 ? "-" : "") + "$" + money(Math.abs(v));
}
function cantidadTicket(n) {
  const v = Number(n) || 0;
  return Number.isInteger(v) ? String(v) : v.toLocaleString("es-AR", { maximumFractionDigits: 3 });
}
function totalArticulosTicket(items) {
  return (items || []).reduce((a, i) => a + (Number(i.cantidad) || 0), 0);
}
/** El ticket común no es una factura (y el previo ni siquiera es comprobante de pago) */
function leyendaTicketVenta(ventaId) {
  return String(ventaId) === "PREVIO" ? "TICKET PREVIO - NO ES COMPROBANTE DE PAGO" : "DOCUMENTO NO VÁLIDO COMO FACTURA";
}
/** Diferencia de caja: el importe con signo ("-$500", "+$500", "$0")… */
function diferenciaTicket(d) {
  const v = Math.round((Number(d) || 0) * 100) / 100;
  if (!v) return "$0";
  return (v < 0 ? "-" : "+") + plataTicket(Math.abs(v));
}
/** …y la palabra que la explica, en la etiqueta: "Faltante", "Sobrante" o "Diferencia" */
function etiquetaDiferencia(d, prefijo) {
  const v = Math.round((Number(d) || 0) * 100) / 100;
  const palabra = v < 0 ? "Faltante" : v > 0 ? "Sobrante" : "Diferencia";
  return prefijo ? prefijo + " " + palabra.toUpperCase() : palabra;
}
/** Líneas del ticket para un cobro en dólares: [[etiqueta, valor], …] */
function filasLineasUsdTicket(p) {
  const filas = [["Pago en dolares", "US$ " + Number(p.montoUsd).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })],
                 ["Dolar", plataTicket(p.tipoCambio)]];
  if (p.restoArs > 0) filas.push(["Resto en pesos", plataTicket(p.restoArs)]);
  if (p.vueltoArs > 0) filas.push(["Vuelto en pesos", plataTicket(p.vueltoArs)]);
  return filas;
}
/** Cambio a devolver: el del dato si viene, si no recibido − total */
function cambioTicket(cambioData, total) {
  if (!cambioData || !(Number(cambioData.recibido) > 0)) return null;
  const cambio = cambioData.cambio !== undefined ? Number(cambioData.cambio) : Number(cambioData.recibido) - Number(total);
  return { recibido: Number(cambioData.recibido), cambio: Math.max(0, cambio) };
}

function buildThermalHTML(ventaId, items, total, formaPago, fecha, descuento, cfgOverride, cambioData, fiscal) {
  const cfg = cfgOverride || obtenerConfigNegocio();
  if (!(fiscal && window.ArcaUI)) fiscal = null; // factura electrónica (facturacion-arca.js)

  const fechaStr = (fecha || new Date()).toLocaleString("es-AR", {
    day:"2-digit", month:"2-digit", year:"numeric",
    hour:"2-digit", minute:"2-digit"
  });

  let rows = "";
  let subtotal = 0;
  items.forEach(item => {
    const sub = item.PRECIO * item.cantidad;
    subtotal += sub;
    const unitStr = `${cantidadTicket(item.cantidad)} x ${plataTicket(item.PRECIO)}`;
    rows += `
      <tr>
        <td colspan="2" style="padding-bottom:0;">${escapeHtml(item.PRODUCTO)}</td>
      </tr>
      <tr>
        <td>${unitStr}</td>
        <td style="text-align:right;">${plataTicket(sub)}</td>
      </tr>`;
  });

  const tieneDescuento = descuento && Number(descuento.monto) !== 0;
  const esRecargoPrint = tieneDescuento && Number(descuento.monto) < 0;

  let filasTotales = "";
  if (tieneDescuento) {
    const etiquetaFila = esRecargoPrint ? "Recargo" : "Descuento";
    const signoFila = esRecargoPrint ? "+" : "-";
    filasTotales += `
      <tr>
        <td>Subtotal</td>
        <td style="text-align:right;">${plataTicket(subtotal)}</td>
      </tr>
      <tr>
        <td>${etiquetaFila}${descuento.etiqueta ? " (" + escapeHtml(descuento.etiqueta) + ")" : ""}</td>
        <td style="text-align:right;">${signoFila}${plataTicket(Math.abs(descuento.monto))}</td>
      </tr>`;
  }
  filasTotales += `
    <tr class="th-total-row">
      <td><strong>TOTAL</strong></td>
      <td style="text-align:right;"><strong>${plataTicket(total)}</strong></td>
    </tr>`;
  // Factura A: líneas sin IVA y el IVA discriminado (el descuento ya va repartido en las líneas)
  if (fiscal && fiscal.clase === "A") { rows = ArcaUI.htmlFilasA(fiscal); filasTotales = ArcaUI.htmlTotalesA(fiscal); }
  // Cantidad de artículos (útil en ventas por mayor)
  filasTotales = `
    <tr>
      <td>Artículos</td>
      <td style="text-align:right;">${cantidadTicket(totalArticulosTicket(items))}</td>
    </tr>` + filasTotales;

  // Cobro en dólares (depósito)
  if (cambioData && cambioData.usd) {
    filasLineasUsdTicket(cambioData.usd).forEach(([a, b]) => {
      filasTotales += `
    <tr>
      <td>${escapeHtml(a)}</td>
      <td style="text-align:right;">${escapeHtml(b)}</td>
    </tr>`;
    });
  }
  // Recibido y cambio — solo si se ingresó un monto recibido
  const cambioInfo = cambioTicket(cambioData, total);
  if (cambioInfo) {
    filasTotales += `
    <tr>
      <td>Recibido</td>
      <td style="text-align:right;">${plataTicket(cambioInfo.recibido)}</td>
    </tr>
    <tr>
      <td><strong>Cambio</strong></td>
      <td style="text-align:right;"><strong>${plataTicket(cambioInfo.cambio)}</strong></td>
    </tr>`;
  }

  // Encabezado: nombre + subtítulo + dirección + teléfono(s), todos configurables
  let encabezado = `<div class="th-center th-big">${escapeHtml(cfg.nombre)}</div>`;
  if (cfg.subtitulo) {
    encabezado += `<div class="th-center" style="font-size:9.5pt;font-weight:bold;">${escapeHtml(cfg.subtitulo)}</div>`;
  }
  if (cfg.direccion) {
    encabezado += `<div class="th-center" style="font-size:9pt;font-weight:bold;">${escapeHtml(cfg.direccion)}</div>`;
  }
  const telefonos = [cfg.telefono1, cfg.telefono2].filter(Boolean).join(" · ");
  if (telefonos) {
    encabezado += `<div class="th-center" style="font-size:9pt;font-weight:bold;margin-bottom:2mm;">Tel: ${escapeHtml(telefonos)}</div>`;
  } else {
    encabezado += `<div style="margin-bottom:2mm;"></div>`;
  }

  // Pie de página (el texto se configura en Configuración → Negocio → Pie de página)
  let pie = "";
  if (!fiscal) pie += `<div class="th-center" style="font-size:8.5pt;margin-bottom:1.5mm;">${leyendaTicketVenta(ventaId)}</div>`;
  if (cfg.pie) {
    pie += `<div class="th-footer">${escapeHtml(cfg.pie)}</div>`;
  }
  pie += `<div class="th-footer" style="margin-top:1mm;">${escapeHtml(cfg.nombre)} &bull; Sistema POS</div>`;

  if (fiscal) encabezado += `<hr class="th-sep-solid">` + ArcaUI.htmlEncabezado(fiscal);

  return `
    <div class="thermal-receipt">
      ${encabezado}
      <hr class="th-sep-solid">
      <div>Fecha: ${fechaStr}</div>
      <div>${fiscal ? "Comprobante" : "Venta"}: #${escapeHtml(String(ventaId || "—"))}</div>
      ${formaPago || !fiscal ? `<div>Pago: ${escapeHtml(String(formaPago || "—"))}</div>` : ""}
      <hr class="th-sep">
      <table>
        <tbody>${rows}</tbody>
        ${filasTotales}
      </table>
      ${fiscal ? ArcaUI.htmlPie(fiscal) : ""}
      <hr class="th-sep">
      ${pie}
      <br><br>
    </div>`;
}

/* ===================================================================
   IMPRESIÓN DIRECTA USB (Web Serial API — Chrome/Edge, sin diálogo)
   Construye los mismos tickets que buildThermalHTML/buildThermalCierreHTML
   pero como comandos ESC/POS en bytes crudos, enviados directo al
   puerto USB de la impresora. Es un camino alternativo: si está
   desactivado (o el navegador no lo soporta), todo sigue imprimiendo
   con el diálogo normal de Chrome, sin ningún cambio.
=================================================================== */

const USB_PRINT_PREF_KEY = "jireh_usb_print_enabled";
const ANCHO_TICKET_USB = 42; // columnas para fuente normal en 80mm (12 cpl aprox.)

// Inyecta un <style> con el @page correcto justo antes de imprimir.
// No usamos "@page nombreDePagina" + "page: nombreDePagina" en el CSS
// porque el soporte de páginas con nombre es poco confiable en
// Chromium/Electron: sin esto, el navegador cae al @page por defecto
// (80mm, pensado para el ticket térmico) incluso al imprimir una
// etiqueta o un A4, dando como resultado una impresión con forma de
// ticket térmico en vez del tamaño esperado.
function _setPrintPageSize(size, cssPropio) {
  let tag = document.getElementById("dynamicPrintPageSize");
  if (!tag) {
    tag = document.createElement("style");
    tag.id = "dynamicPrintPageSize";
    document.head.appendChild(tag);
  }
  tag.textContent = size === "CUSTOM" && cssPropio
    ? cssPropio
    : size === "REMITO"
    ? "@media print { @page { size:A4; margin:0; } }"
    : size === "A4"
    ? "@media print { @page { size:A4; margin:8mm; } }"
    : "@media print { @page { size:80mm auto; margin:0; } }";
}

let puertoImpresoraUSB = null; // SerialPort activo, o null si no hay conexión

/** Whether the browser supports Web Serial at all */
function soportaImpresionUSB() {
  return "serial" in navigator;
}

/** Reads the saved on/off preference for USB direct printing (per-browser, default off) */
function usbPrintHabilitado() {
  return soportaImpresionUSB() && localStorage.getItem(USB_PRINT_PREF_KEY) === "true";
}

/* ---------------------- ESC/POS byte builders ---------------------- */

const ESC = 0x1B;
const GS  = 0x1D;

const ESCPOS = {
  INIT:          [ESC, 0x40],             // reset printer
  ALIGN_LEFT:    [ESC, 0x61, 0x00],
  ALIGN_CENTER:  [ESC, 0x61, 0x01],
  BOLD_ON:       [ESC, 0x45, 0x01],
  BOLD_OFF:      [ESC, 0x45, 0x00],
  SIZE_NORMAL:   [GS, 0x21, 0x00],
  SIZE_DOUBLE_H: [GS, 0x21, 0x01],         // double height
  SIZE_BIG:      [GS, 0x21, 0x11],         // double width + height
  CUT:           [GS, 0x56, 0x01],         // partial cut
  FEED:          (n) => [ESC, 0x64, n]     // feed n lines
};

/** Builds the raw byte buffer (Uint8Array) for a sale ticket, mirroring buildThermalHTML */
function buildThermalESCPOS(ventaId, items, total, formaPago, fecha, descuento, cfgOverride, fiscal, cambioData) {
  const cfg = cfgOverride || obtenerConfigNegocio();
  const b = new EscPosBuilder();
  if (!(fiscal && window.ArcaUI)) fiscal = null; // factura electrónica (facturacion-arca.js)

  const fechaStr = (fecha || new Date()).toLocaleString("es-AR", {
    day:"2-digit", month:"2-digit", year:"numeric",
    hour:"2-digit", minute:"2-digit"
  });

  b.init();
  b.center(); b.big(); b.text(cfg.nombre); b.feed(1);
  b.normal();
  if (cfg.subtitulo) { b.center(); b.bold(); b.text(cfg.subtitulo); b.feed(1); b.boldOff(); }
  if (cfg.direccion) { b.center(); b.bold(); b.text(cfg.direccion); b.feed(1); b.boldOff(); }
  const telefonos = [cfg.telefono1, cfg.telefono2].filter(Boolean).join(" / ");
  if (telefonos) { b.center(); b.bold(); b.text("Tel: " + telefonos); b.feed(1); b.boldOff(); }
  if (fiscal) { b.left(); b.sepSolid(); ArcaUI.escposEncabezado(b, fiscal); }

  b.left();
  b.sepSolid();
  b.bold();
  b.text("Fecha: " + fechaStr); b.feed(1);
  b.text((fiscal ? "Comprobante" : "Venta") + ": #" + String(ventaId || "—")); b.feed(1);
  if (formaPago || !fiscal) { b.text("Pago: " + String(formaPago || "—")); b.feed(1); }
  b.sep();

  const articulos = totalArticulosTicket(items);
  let subtotal = 0;
  if (fiscal && fiscal.clase === "A") { ArcaUI.escposFilasA(b, fiscal); ArcaUI.escposTotalesA(b, fiscal); items = []; descuento = null; }
  items.forEach(item => {
    const sub = item.PRECIO * item.cantidad;
    subtotal += sub;
    b.bold();
    b.text(item.PRODUCTO); b.feed(1);
    b.row(`${cantidadTicket(item.cantidad)} x ${plataTicket(item.PRECIO)}`, plataTicket(sub));
  });

  b.sep();
  b.row("Articulos", cantidadTicket(articulos));
  const tieneDescuento = descuento && Number(descuento.monto) !== 0;
  if (tieneDescuento) {
    const esRecargoESC = Number(descuento.monto) < 0;
    b.row("Subtotal", plataTicket(subtotal));
    const palabraAjuste = esRecargoESC ? "Recargo" : "Descuento";
    const etiquetaDesc = descuento.etiqueta ? `${palabraAjuste} (${descuento.etiqueta})` : palabraAjuste;
    const signoESC = esRecargoESC ? "+" : "-";
    b.row(etiquetaDesc, signoESC + plataTicket(Math.abs(descuento.monto)));
  }

  b.sepSolid();
  b.doubleH();
  b.row("TOTAL", plataTicket(total));
  b.normal();
  if (cambioData && cambioData.usd) filasLineasUsdTicket(cambioData.usd).forEach(([a, c]) => b.row(a, c));
  const cambioInfo = cambioTicket(cambioData, total);
  if (cambioInfo) {
    b.row("Recibido", plataTicket(cambioInfo.recibido));
    b.bold(); b.row("Cambio", plataTicket(cambioInfo.cambio)); b.boldOff();
  }
  if (fiscal) ArcaUI.escposPie(b, fiscal);
  b.sep();

  b.bold();
  if (!fiscal) { b.center(); b.text(leyendaTicketVenta(ventaId)); b.feed(1); }
  if (cfg.pie) { b.center(); b.text(cfg.pie); b.feed(1); }
  b.center(); b.text(cfg.nombre + " - Sistema POS"); b.feed(1);
  b.boldOff();

  b.feed(6); // margen generoso antes del corte -- con 3 el guillotine cortaba muy cerca de la ultima linea y dejaba texto fantasma del ticket anterior en el siguiente
  b.cut();

  return b.build();
}

/**
 * Alto del papel (en micrones) para la impresión por diálogo, calculado a
 * partir de los renglones REALES del ticket: se cuentan en los mismos bytes
 * ESC/POS (que ya parten las líneas largas), así un nombre de producto que
 * ocupa dos o tres renglones también suma. Antes se suponían 2 renglones
 * por producto, y el cierre de caja pedía siempre 29,7 cm de papel.
 */
function altoTicketDesdeEscPos(bytes, extraMm, mmPorRenglon) {
  let lineas = 0, ultimoFeed = 0;
  for (let i = 0; i < bytes.length; i++) {
    const x = bytes[i], y = bytes[i + 1];
    if (x === 0x1D && y === 0x56) break;                                   // corte
    if (x === 0x1D && y === 0x28 && bytes[i + 2] === 0x6B) { i += 4 + bytes[i + 3] + bytes[i + 4] * 256; continue; } // QR
    if (x === 0x1B && y === 0x64) { lineas += bytes[i + 2]; ultimoFeed = bytes[i + 2]; i += 2; continue; } // avance
    if (x === 0x1B && y === 0x40) { i += 1; continue; }                    // inicializar
    if (x === 0x1B || x === 0x1D) { i += 2; continue; }                    // formato (3 bytes)
    if (x === 0x0A) { lineas++; ultimoFeed = 0; }                          // renglón partido
  }
  lineas -= ultimoFeed; // el avance antes del corte no ocupa papel en el diálogo
  // 9 pt con interlineado + el espacio de las filas de tabla, medido contra
  // tickets reales renderizados (el cierre es casi todo filas de tabla: más alto)
  const altoMm = Math.max(40, lineas * (mmPorRenglon || 5.6) + 14 + (extraMm || 0));
  return Math.round(Math.min(altoMm, 1200) * 1000);
}

/** Bytes ESC/POS del cierre de caja (mismo contenido que buildThermalCierreHTML) */
function buildThermalCierreESCPOS(resumen) {
  const cfg = obtenerConfigNegocio();
  const b = new EscPosBuilder();
  const d = datosTicketCierre(resumen);

  b.init();
  b.center(); b.big(); b.text(cfg.nombre); b.feed(1);
  b.normal();
  b.center(); b.bold(); b.doubleH(); b.text(d.titulo); b.feed(1); b.normal(); b.boldOff();
  if (cfg.direccion) { b.center(); b.text(cfg.direccion); b.feed(1); }
  const telefonos = [cfg.telefono1, cfg.telefono2].filter(Boolean).join(" / ");
  if (telefonos) { b.center(); b.text("Tel: " + telefonos); b.feed(1); }

  b.left();
  b.sepSolid();
  d.datos.forEach(([k, v]) => b.row(k, v));
  b.sep();
  b.bold(); b.row("Ventas del dia (" + d.cantidadVentas + ")", plataTicket(d.totalVentas)); b.boldOff();
  if (d.cantidadAnuladas) b.row("Anuladas (no suman)", String(d.cantidadAnuladas));

  d.metodos.forEach(m => {
    b.sep();
    b.bold(); b.text(m.nombre); b.feed(1); b.boldOff();
    m.lineas.forEach(([k, v]) => b.row("  " + k, v));
    b.bold(); b.row("  Esperado", plataTicket(m.esperado)); b.boldOff();
    b.row("  " + m.etiquetaContado, plataTicket(m.contado));
    b.row("  " + etiquetaDiferencia(m.diferencia), diferenciaTicket(m.diferencia));
  });

  b.sepSolid();
  b.bold();
  b.row("TOTAL ESPERADO", plataTicket(d.total.esperado));
  b.row("TOTAL CONTADO", plataTicket(d.total.contado));
  b.row(etiquetaDiferencia(d.total.diferencia).toUpperCase(), diferenciaTicket(d.total.diferencia));
  b.boldOff();
  b.sep();

  if (resumen.observaciones) {
    b.text("Obs: " + resumen.observaciones); b.feed(1);
    b.sep();
  }
  if (d.esPrevio) { b.center(); b.bold(); b.text("Cierre no guardado: solo para revisar"); b.feed(1); b.boldOff(); }
  b.center(); b.text("Cierre generado por " + cfg.nombre + " POS"); b.feed(1);

  b.feed(6); // margen generoso antes del corte -- con 3 el guillotine cortaba muy cerca de la ultima linea y dejaba texto fantasma del ticket anterior en el siguiente
  b.cut();

  return b.build();
}

/** Número con separador de miles; con centavos solo si los tiene ("1.250" / "1.250,50") */
function money(n) {
  const v = Math.round((Number(n) || 0) * 100) / 100;
  const dec = Number.isInteger(v) ? 0 : 2;
  return v.toLocaleString("es-AR", { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

/** Parte un texto en renglones de `ancho` caracteres, sin cortar palabras (salvo que una sola no entre) */
function partirRenglonesTicket(texto, ancho) {
  const salida = [];
  String(texto).split("\n").forEach(parrafo => {
    let linea = "";
    parrafo.split(/ +/).forEach(palabra => {
      while (palabra.length > ancho) {
        if (linea) { salida.push(linea); linea = ""; }
        salida.push(palabra.slice(0, ancho));
        palabra = palabra.slice(ancho);
      }
      if (!palabra) return;
      if (!linea) linea = palabra;
      else if ((linea + " " + palabra).length <= ancho) linea += " " + palabra;
      else { salida.push(linea); linea = palabra; }
    });
    salida.push(linea);
  });
  return salida;
}

/**
 * Convierte texto a ASCII puro para que imprima bien en cualquier
 * impresora térmica, sin importar su tabla de códigos activa. Ver la
 * nota completa en EscPosBuilder.text() de por qué hace falta esto.
 */
function normalizarParaImpresoraTermica(str) {
  return str
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita tildes/diacríticos: á→a, ñ→n, é→e, etc.
    .replace(/[\u00A0\u2000-\u200B\u202F\u205F\u3000]/g, " ") // espacios "raros" (NBSP, espacio angosto, etc.) → espacio normal
    .replace(/[¡¿]/g, "") // sin equivalente ASCII limpio — se descartan
    .replace(/[""]/g, '"').replace(/['']/g, "'") // comillas tipográficas → comillas rectas
    .replace(/[^\x00-\x7F]/g, "?"); // cualquier otro caracter no-ASCII remanente → "?" en vez de bytes basura
}

/** Small helper that accumulates ESC/POS bytes with simple word-wrap and two-column row support */
class EscPosBuilder {
  constructor() {
    this.bytes = [];
    this.ancho = ANCHO_TICKET_USB; // columnas disponibles con la letra actual
  }
  push(arr) { this.bytes.push(...arr); }
  init()    { this.push(ESCPOS.INIT); this.ancho = ANCHO_TICKET_USB; }
  center()  { this.push(ESCPOS.ALIGN_CENTER); }
  left()    { this.push(ESCPOS.ALIGN_LEFT); }
  bold()    { this.push(ESCPOS.BOLD_ON); }
  boldOff() { this.push(ESCPOS.BOLD_OFF); }
  normal()  { this.push(ESCPOS.SIZE_NORMAL); this.ancho = ANCHO_TICKET_USB; }
  doubleH() { this.push(ESCPOS.SIZE_DOUBLE_H); this.ancho = ANCHO_TICKET_USB; }
  big()     { this.push(ESCPOS.SIZE_BIG); this.ancho = Math.floor(ANCHO_TICKET_USB / 2); } // doble ancho: la mitad de columnas
  cut()     { this.push(ESCPOS.CUT); }
  feed(n)   { this.push(ESCPOS.FEED(n)); }

  /**
   * Codifica el texto para la impresora térmica.
   *
   * Antes esto mandaba los bytes crudos en UTF-8 (vía TextEncoder) —
   * el comentario decía "Latin-1" pero el código hacía otra cosa, y
   * esa es justo la causa del texto corrompido que se ve en tickets
   * reales ("p.m." → "p.τám.", "¡Gracias" → "ᴛíGracias"): un
   * caracter no-ASCII en UTF-8 ocupa 2 o 3 bytes, y la impresora — que
   * lee byte por byte contra SU PROPIA tabla de un solo byte — imprime
   * cada uno de esos bytes como un símbolo suelto en vez de reconocer
   * el caracter original. Pasa con tildes, con "¡"/"¿", y hasta con el
   * espacio angosto que $ toLocaleString mete entre "p." y "m." en
   * formato 12hs — ese espacio también es multi-byte en UTF-8.
   *
   * En vez de adivinar qué tabla de códigos usa esta impresora en
   * particular (varía según el modelo/driver), se transforma el texto
   * a ASCII puro antes de mandarlo: sin tildes (á→a, ñ→n, etc., vía
   * normalize+strip de diacríticos), sin "¡"/"¿" (no tienen
   * equivalente ASCII limpio), y cualquier espacio "raro" (NBSP,
   * espacio angosto, etc.) pasa a espacio normal. El ASCII de 7 bits
   * se imprime igual sin importar la tabla de códigos activa en la
   * impresora, así que esto funciona en cualquier modelo.
   */
  text(str) {
    // Lo que no entra en el ancho del papel pasa al renglón siguiente entre
    // palabras (antes la impresora lo cortaba donde caía, a mitad de palabra)
    const limpio = normalizarParaImpresoraTermica(String(str));
    // Si entra, va tal cual (respeta los espacios de alineación de row())
    const renglones = limpio.length <= this.ancho && limpio.indexOf("\n") === -1 ? [limpio] : partirRenglonesTicket(limpio, this.ancho);
    const bytes = [];
    renglones.forEach((r, k) => {
      if (k > 0) bytes.push(0x0A);
      for (let i = 0; i < r.length; i++) bytes.push(r.charCodeAt(i) & 0xFF);
    });
    this.push(bytes);
  }

  sep() {
    this.bold();
    this.text("-".repeat(ANCHO_TICKET_USB));
    this.feed(1);
    this.boldOff();
  }

  sepSolid() {
    this.bold();
    this.text("=".repeat(ANCHO_TICKET_USB));
    this.feed(1);
    this.boldOff();
  }

  /** Prints a left label and a right-aligned value on the same 42-col line (wraps the label if too long) */
  row(left, right) {
    const rightStr = normalizarParaImpresoraTermica(String(right));
    const maxLeft = ANCHO_TICKET_USB - rightStr.length - 1;
    const leftStr = normalizarParaImpresoraTermica(String(left));
    if (leftStr.length > maxLeft) {
      // Etiqueta larga: va entera (en uno o más renglones) y el importe en el último, a la derecha
      const renglones = partirRenglonesTicket(leftStr, ANCHO_TICKET_USB);
      const ultimo = renglones.pop();
      if (renglones.length) { this.text(renglones.join("\n")); this.feed(1); }
      if (ultimo.length <= maxLeft) { this.text(ultimo + " ".repeat(Math.max(1, ANCHO_TICKET_USB - ultimo.length - rightStr.length)) + rightStr); this.feed(1); return; }
      this.text(ultimo); this.feed(1);
      this.text(" ".repeat(Math.max(0, ANCHO_TICKET_USB - rightStr.length)) + rightStr); this.feed(1);
      return;
    }
    const spaces = Math.max(1, ANCHO_TICKET_USB - leftStr.length - rightStr.length);
    this.text(leftStr + " ".repeat(spaces) + rightStr);
    this.feed(1);
  }

  build() {
    return new Uint8Array(this.bytes);
  }
}

/* ---------------------- Web Serial connection ---------------------- */

/** Prompts the browser's port picker and opens the connection (must be called from a user gesture) */
async function conectarImpresoraUSB() {
  if (!soportaImpresionUSB()) {
    toast("Este navegador no soporta impresión directa", "error");
    return;
  }
  try {
    const port = await navigator.serial.requestPort();
    await port.open({ baudRate: 9600 });
    puertoImpresoraUSB = port;
    actualizarEstadoUSBPrint();
    toast("Impresora conectada", "success");
  } catch (error) {
    if (error.name !== "NotFoundError") { // user just closed the picker without choosing
      console.error("Error al conectar la impresora USB:", error);
      toast("No se pudo conectar con la impresora", "error");
    }
  }
}

async function desconectarImpresoraUSB() {
  if (puertoImpresoraUSB) {
    try { await puertoImpresoraUSB.close(); } catch (e) { /* ignore */ }
  }
  puertoImpresoraUSB = null;
  actualizarEstadoUSBPrint();
  toast("Impresora desconectada", "success");
}

/** Tries to silently reconnect to a previously-authorized port (no picker shown) when the app loads */
async function reconectarImpresoraUSBSiPosible() {
  if (!soportaImpresionUSB()) return;
  try {
    const ports = await navigator.serial.getPorts();
    if (ports.length > 0) {
      const port = ports[0];
      await port.open({ baudRate: 9600 });
      puertoImpresoraUSB = port;
    }
  } catch (error) {
    // Port may be in use or unplugged — just leave it disconnected, the admin can reconnect manually
    puertoImpresoraUSB = null;
  }
  actualizarEstadoUSBPrint();
}

/** Sends a raw byte buffer to the connected printer over Web Serial */
async function enviarBytesAImpresoraUSB(bytes) {
  if (!puertoImpresoraUSB || !puertoImpresoraUSB.writable) {
    throw new Error("La impresora USB no está conectada");
  }
  const writer = puertoImpresoraUSB.writable.getWriter();
  try {
    // Si la impresora está trabada, sin papel, o el cable se desconectó
    // a medias, writer.write() puede quedar esperando para siempre sin
    // resolver ni rechazar — el botón "parece" no responder porque en
    // realidad está esperando una escritura que nunca va a terminar.
    // Con este límite de tiempo, a los 4s se corta y se cae al diálogo
    // de impresión normal en vez de quedar colgado.
    const TIMEOUT_MS = 4000;
    let timeoutId;
    const timeout = new Promise((_, reject) => {
      timeoutId = setTimeout(() => reject(new Error("Tiempo de espera agotado al escribir en la impresora USB")), TIMEOUT_MS);
    });
    try {
      await Promise.race([writer.write(bytes), timeout]);
    } finally {
      clearTimeout(timeoutId);
    }
  } finally {
    try { writer.releaseLock(); } catch(e) { /* puede fallar si la escritura seguía pendiente tras el timeout; no es crítico */ }
  }
}

function probarImpresoraUSB() {
  const b = new EscPosBuilder();
  b.init();
  b.center(); b.big(); b.text("PRUEBA"); b.feed(2);
  b.normal(); b.bold();
  b.text("Si ves esto, la impresión"); b.feed(1);
  b.text("directa por USB funciona bien."); b.feed(1);
  b.boldOff();
  b.feed(6); // margen generoso antes del corte -- con 3 el guillotine cortaba muy cerca de la ultima linea y dejaba texto fantasma del ticket anterior en el siguiente
  b.cut();

  enviarBytesAImpresoraUSB(b.build())
    .then(() => toast("Ticket de prueba enviado", "success"))
    .catch(error => {
      console.error("Error en impresión de prueba:", error);
      toast("No se pudo imprimir: revisá la conexión USB", "error");
    });
}

function onToggleUsbPrint() {
  const checked = document.getElementById("usbPrintToggle").checked;
  if (checked && !puertoImpresoraUSB) {
    toast("Conectá la impresora primero", "error");
    document.getElementById("usbPrintToggle").checked = false;
    return;
  }
  localStorage.setItem(USB_PRINT_PREF_KEY, checked ? "true" : "false");
  toast(checked ? "Impresión directa activada" : "Impresión directa desactivada", "success");
}

/** Refreshes the Configuración card UI to reflect support/connection state */
function actualizarEstadoUSBPrint() {
  const unsupportedEl = document.getElementById("usbPrintUnsupported");
  const supportedEl   = document.getElementById("usbPrintSupported");
  if (!unsupportedEl || !supportedEl) return; // section not in the DOM yet

  if (!soportaImpresionUSB()) {
    unsupportedEl.style.display = "flex";
    supportedEl.style.display = "none";
    return;
  }

  unsupportedEl.style.display = "none";
  supportedEl.style.display = "block";

  const statusEl     = document.getElementById("usbPrintStatus");
  const statusTextEl = document.getElementById("usbPrintStatusText");
  const btnConectar   = document.getElementById("btnConectarImpresora");
  const btnDesconectar = document.getElementById("btnDesconectarImpresora");
  const btnProbar      = document.getElementById("btnProbarImpresora");
  const toggle          = document.getElementById("usbPrintToggle");

  if (puertoImpresoraUSB) {
    statusEl.className = "usb-print-status ok";
    statusTextEl.textContent = "Impresora conectada";
    btnConectar.style.display = "none";
    btnDesconectar.style.display = "inline-block";
    btnProbar.style.display = "inline-block";
  } else {
    statusEl.className = "usb-print-status";
    statusTextEl.textContent = "Impresora no conectada";
    btnConectar.style.display = "inline-block";
    btnDesconectar.style.display = "none";
    btnProbar.style.display = "none";

    // Can't keep "impresión directa" on if there's nothing connected
    if (toggle) toggle.checked = false;
    localStorage.setItem(USB_PRINT_PREF_KEY, "false");
  }

  if (toggle) toggle.checked = usbPrintHabilitado();
}

/** Print the current (unsaved) ticket as a pre-sale receipt */
let _imprimiendoTicket = false; // evita disparar impresiones superpuestas si tocan el botón varias veces

function imprimirTicketThermal() {
  if (ticketPOS.length === 0) { toast("El ticket está vacío", "error"); return; }
  if (_imprimiendoTicket) return; // ya hay una impresión en curso — ignorar toques repetidos
  const subtotal = ticketPOS.reduce((acc, i) => acc + i.PRECIO * i.cantidad, 0);
  const { montoDescuento, total } = calcularDescuentoPOS(subtotal);
  const etiqueta = obtenerEtiquetaDescuentoPOS(subtotal);
  _ejecutarImpresion("PREVIO", ticketPOS, total, formaPagoPOS, new Date(),
    montoDescuento !== 0 ? { monto: montoDescuento, etiqueta } : null);
}

/** Print after a completed sale (from receipt modal) */
async function imprimirUltimoRecibo() {
  if (!ultimaVentaImprimible) { toast("No hay venta para imprimir", "error"); return; }
  const u = ultimaVentaImprimible;
  // Con factura electrónica, el ticket sale con los datos fiscales (CAE y QR)
  const fiscal = window.ArcaUI ? await ArcaUI.facturaParaImprimir(u) : null;
  _ejecutarImpresion(u.ventaId, u.items, u.total, u.formaPago, u.fecha,
    u.descuento !== 0 ? { monto: u.descuento, etiqueta: u.descuentoEtiqueta } : null, fiscal,
    u.pagoUsd ? { usd: u.pagoUsd } : u.recibido > 0 ? { recibido: u.recibido } : null);
}

/** Print a sale from the dashboard recent-sales table or the Ventas POS history */
function imprimirVentaDesdeData(ventaObj) {
  const items = [];
  // Try to parse items if stored as JSON string or array
  if (ventaObj.CARRITO) {
    try {
      const parsed = typeof ventaObj.CARRITO === "string"
        ? JSON.parse(ventaObj.CARRITO)
        : ventaObj.CARRITO;
      if (Array.isArray(parsed)) items.push(...parsed);
    } catch(e) { /* ignore */ }
  }
  // Fallback: show just a total line if no item detail
  if (items.length === 0) {
    items.push({ PRODUCTO: ventaObj.DETALLE || "Venta", PRECIO: Number(ventaObj.TOTAL || 0), cantidad: 1 });
  }

  // El descuento (si lo hubo) quedó guardado como texto en OBSERVACIONES,
  // ej: "Descuento: 10% (-$500) — cliente frecuente". Lo reconstruimos
  // calculando subtotal real de items vs. el TOTAL guardado (ya con descuento).
  const subtotalItems = items.reduce((acc, i) => acc + Number(i.PRECIO) * Number(i.cantidad), 0);
  const totalGuardado = Number(ventaObj.TOTAL || 0);
  const diferencia = subtotalItems - totalGuardado;

  const obs = String(ventaObj.OBSERVACIONES || "");
  let descuentoInfo = null;
  if (diferencia > 0.5 && /descuento/i.test(obs)) {
    descuentoInfo = { monto: diferencia, etiqueta: obs.replace(/^Descuento:\s*/i, "") };
  } else if (diferencia < -0.5 && /recargo/i.test(obs)) {
    descuentoInfo = { monto: diferencia, etiqueta: obs.replace(/^Recargo:\s*/i, "") }; // monto negativo = recargo
  }

  _ejecutarImpresion(
    ventaObj.VENTA_ID || ventaObj.ID || "—",
    items,
    totalGuardado,
    ventaObj.FORMA_PAGO || ventaObj.PAGO || "—",
    ventaObj.FECHA ? new Date(ventaObj.FECHA) : new Date(),
    descuentoInfo,
    null,
    null // reimpresión: no se sabe cuánto entregó el cliente (antes tomaba lo escrito en el POS en ese momento)
  );
}


/**
 * Resuelve qué impresora usar para el camino RAW: el nombre guardado
 * en Configuración si hay uno, o si no, la impresora predeterminada de
 * Windows (consultada una sola vez por sesión y cacheada en memoria).
 *
 * Por qué hace falta esto: antes, sin nada guardado en Configuración,
 * el camino RAW se saltaba directo (el if de más abajo exigía un
 * nombreImpresora no vacío) y toda la impresión dependía del camino
 * viejo — más lento, y con un bug conocido de Electron en Windows
 * donde, si se cancela el diálogo nativo, el callback a veces no
 * llega nunca y el botón de imprimir queda colgado hasta que el
 * watchdog de pos-offline.js lo destraba a la fuerza a los 10s.
 */
let _nombreImpresoraPorDefectoCache = null; // se resuelve una sola vez por sesión de la app

async function resolverNombreImpresoraParaRaw(bridge) {
  const guardado = localStorage.getItem("veekpos_impresora") || "";
  if (guardado) return guardado;

  if (_nombreImpresoraPorDefectoCache !== null) return _nombreImpresoraPorDefectoCache;

  if (typeof bridge.listarImpresoras !== "function") {
    _nombreImpresoraPorDefectoCache = "";
    return "";
  }

  try {
    const impresoras = await bridge.listarImpresoras();
    const porDefecto = (impresoras || []).find(p => p.isDefault) || (impresoras || [])[0];
    _nombreImpresoraPorDefectoCache = (porDefecto && porDefecto.name) || "";
  } catch (error) {
    console.error("No se pudo obtener la impresora predeterminada de Windows:", error);
    _nombreImpresoraPorDefectoCache = "";
  }
  return _nombreImpresoraPorDefectoCache;
}

function _ejecutarImpresion(ventaId, items, total, formaPago, fecha, descuento, fiscal, cambioData) {
  // Sin cambioData explícito (ticket previo, antes de cobrar) se toma lo que
  // está escrito en el POS. Después de la venta esos campos ya se limpiaron:
  // por eso el recibo y las reimpresiones pasan el suyo (o null).
  if (cambioData === undefined) cambioData = obtenerDatosCambio();
  if (!(fiscal && window.ArcaUI)) fiscal = null;

  const btn = document.getElementById("btnImprimirTicket");
  const textoOriginalBtn = btn ? btn.innerHTML : null;
  _imprimiendoTicket = true;
  if (btn) { btn.disabled = true; btn.innerHTML = "⏳"; }

  const liberar = () => {
    _imprimiendoTicket = false;
    if (btn) {
      btn.innerHTML = textoOriginalBtn;
      // El disabled real (según si el ticket tiene items) lo recalcula
      // renderTicketPOS/actualizarBotonesTicketPOS en su próximo render;
      // acá solo se saca el bloqueo temporal de "imprimiendo".
      btn.disabled = ticketPOS.length === 0;
    }
  };

  const bytes = buildThermalESCPOS(ventaId, items, total, formaPago, fecha, descuento, null, fiscal, cambioData);
  _imprimirTicketTermico({
    html: buildThermalHTML(ventaId, items, total, formaPago, fecha, descuento, null, cambioData, fiscal),
    bytes,
    alto: altoTicketDesdeEscPos(bytes, fiscal && fiscal.qrUrl ? 36 : 0), // el QR fiscal mide ~32 mm
    alTerminar: liberar
  });
}

/**
 * Imprime un ticket térmico por el mejor camino disponible:
 * 1. USB directo (Web Serial), si está activado y conectado;
 * 2. impresora de Windows en modo RAW (app de escritorio): los mismos bytes
 *    ESC/POS, sin renderizar la página — mucho más rápido;
 * 3. el diálogo / impresión silenciosa de siempre, con el HTML.
 * Si un camino falla, se cae solo al siguiente.
 */
function _imprimirTicketTermico({ html, bytes, alto, alTerminar }) {
  const fin = () => { if (alTerminar) alTerminar(); };
  const porDialogo = () => _imprimirConDialogo(html, alto);

  if (usbPrintHabilitado() && puertoImpresoraUSB) {
    return enviarBytesAImpresoraUSB(bytes)
      .catch(error => { console.error("Error al imprimir por USB:", error); return porDialogo(); })
      .finally(fin);
  }

  const bridge = window.veekpos || window.posOffline;
  if (bridge && typeof bridge.imprimirDirectoRaw === "function") {
    return resolverNombreImpresoraParaRaw(bridge)
      .then(nombreImpresora => {
        // Sin impresora configurada ni predeterminada en Windows, ni vale la pena intentar RAW
        if (!nombreImpresora) return porDialogo();
        return bridge.imprimirDirectoRaw({ deviceName: nombreImpresora, bytes }).then(resultado => {
          if (resultado && resultado.success) return;
          console.warn("Impresión RAW no disponible o falló, se usa el camino anterior:", resultado && resultado.errorType);
          return porDialogo();
        });
      })
      .catch(error => { console.error("Error al imprimir RAW:", error); return porDialogo(); })
      .finally(fin);
  }

  return Promise.resolve(porDialogo()).finally(fin);
}

/** Falls back to the regular browser print dialog (used when USB printing is off, unsupported, or fails) */
async function _imprimirConDialogo(html, altoMicronesEstimado) {
  const frame = document.getElementById("thermalPrintFrame");
  if (!frame) { toast("Error: frame de impresión no encontrado", "error"); return; }

  const etiquetasArea = document.getElementById("etiquetasPrintArea");
  if (etiquetasArea) etiquetasArea.innerHTML = "";

  frame.innerHTML = html;

  // Esperar a que el contenido del frame esté completamente pintado
  await new Promise(r => requestAnimationFrame(r));

  const imagenes = frame.querySelectorAll("img");
  if (imagenes.length > 0) {
    await Promise.all(Array.from(imagenes).map(img =>
      img.complete ? Promise.resolve() : new Promise(r => {
        img.onload = r; img.onerror = r;
        setTimeout(r, 2000); // máximo 2s de espera por imagen
      })
    ));
  }

  // NOTA: acá hubo dos intentos de calcular el alto real del ticket
  // MIDIENDO EL DOM (scrollHeight) para no pedirle a Windows una
  // página más grande de lo necesario. Los dos terminaron cortando o
  // desalineando tickets reales — medir el alto de un elemento que
  // normalmente está oculto (display:none) es frágil: para que la
  // medición sea correcta hay que "mostrarlo" brevemente, y esa
  // manipulación de estilos resultó nada confiable en la práctica.
  //
  // Esto es distinto: no se mide el DOM, se ESTIMA a partir de los
  // datos del ticket (altoTicketDesdeEscPos, calculado en
  // _ejecutarImpresion antes de armar el HTML) — determinístico, sin
  // tocar estilos ni depender de timing de renderizado. Si por lo que
  // sea no llega un estimado (ej. ticket de cierre de caja, que no lo
  // calcula), se mantiene el alto fijo de 297mm como red de
  // seguridad, igual que antes.
  const altoMicrones = altoMicronesEstimado || 297000;

  // Asegura que, si se cae al diálogo normal (window.print más abajo),
  // la página tenga el tamaño térmico (80mm) y no el A4 que pudo haber
  // quedado configurado por una impresión de etiqueta anterior.
  _setPrintPageSize("80mm");

  // En Electron: impresión silenciosa sin diálogo del sistema
  const bridge = window.veekpos || window.posOffline;
  if (bridge && typeof bridge.imprimirSilencioso === "function") {
    const deviceName = localStorage.getItem("veekpos_impresora") || "";
    let result = await bridge.imprimirSilencioso({ deviceName, altoMicrones });

    // Causa más común de "tengo la impresora configurada y igual sale
    // el diálogo": el nombre guardado (de una lista de impresoras que
    // se cargó una vez) ya no coincide exactamente con cómo el driver
    // reporta la impresora ahora mismo (Windows es especialmente
    // propenso a esto — actualizaciones de driver, reconexión USB,
    // etc.). Antes, un solo fallo con ese nombre hacía caer directo al
    // diálogo visible; ahora, si había un nombre guardado, se reintenta
    // una vez dejando que Electron use la impresora predeterminada del
    // sistema — que suele ser la térmica igual — antes de rendirse.
    if (result && !result.success && deviceName) {
      console.warn(`Impresión silenciosa falló con la impresora guardada ("${deviceName}"): ${result.errorType}. Reintentando con la predeterminada del sistema...`);
      result = await bridge.imprimirSilencioso({ deviceName: "", altoMicrones });
    }

    if (result && !result.success) {
      console.warn("Impresión silenciosa falló:", result.errorType);
      toast(`No se pudo imprimir en silencio (${result.errorType || "motivo desconocido"}) — revisá la impresora en Configuración`, "error");
      window.print();
    }
    return;
  }

  // En el navegador web: diálogo normal
  window.print();
}

/* =====================================================
   THERMAL PRINT — CIERRE DE CAJA (POS80 80mm)
   Reusa el mismo #thermalPrintFrame y las reglas @media
   print ya definidas para el ticket de venta.
===================================================== */

/**
 * Build the thermal HTML string for a cash-register closing (cierre de caja).
 * @param {Object} resumen  — { fecha, cierreId, vendedor, observaciones,
 *                               efectivo:{esperado,contado,diferencia},
 *                               transferencia:{...}, tarjeta:{...}, total:{...} }
 */
/**
 * Contenido del ticket de cierre, igual para pantalla y térmica. Cada forma
 * de pago se arma de modo que las líneas SUMEN el esperado: ventas + cobros a
 * clientes + ingresos de caja − egresos de caja (antes el ticket mostraba solo
 * ventas y egresos, y el esperado no se podía reconstruir).
 */
function datosTicketCierre(resumen) {
  const ahora = new Date();
  const datos = [["Fecha", formatearFechaCierre(resumen.fecha)],
    ["Impreso", ahora.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit" }) + " " + ahora.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit", hour12: false })]];
  if (resumen.cierreId) datos.push(["Cierre", "#" + resumen.cierreId]);
  datos.push(["Responsable", String(resumen.vendedor || "—")]);
  const nombres = { EFECTIVO: "EFECTIVO", TRANSFERENCIA: "TRANSFERENCIA", TARJETA: "TARJETA" };
  const metodos = (resumen.metodos || []).map(m => {
    const lineas = [["Ventas", plataTicket(m.ventas)]];
    if (m.cobros) lineas.push(["Cobros a clientes", plataTicket(m.cobros)]);
    if (m.ingresos) lineas.push(["Ingresos de caja", plataTicket(m.ingresos)]);
    if (m.egresos) lineas.push(["Egresos de caja", "-" + plataTicket(m.egresos)]);
    return {
      nombre: nombres[m.clave] || m.clave, lineas,
      esperado: m.esperado, contado: m.contado, diferencia: m.diferencia,
      // Transferencias y tarjetas no se cuentan: se verifican contra el banco/posnet
      etiquetaContado: m.clave === "EFECTIVO" ? "Contado" : "Verificado"
    };
  });
  return {
    titulo: resumen.esPrevio ? "PRE-CIERRE (NO GUARDADO)" : "CIERRE DE CAJA",
    esPrevio: !!resumen.esPrevio, datos, metodos, total: resumen.total,
    cantidadVentas: Number(resumen.cantidadVentas) || 0,
    cantidadAnuladas: Number(resumen.cantidadAnuladas) || 0,
    totalVentas: Number(resumen.totalVentas) || 0
  };
}

function buildThermalCierreHTML(resumen) {
  const cfg = obtenerConfigNegocio();
  const d = datosTicketCierre(resumen);
  // table-layout:fixed toma los anchos de la primera fila: sin esto la columna
  // de importes quedaba casi en cero y los montos se cortaban
  const COLS_CIERRE = `<colgroup><col><col style="width:40%;"></colgroup>`;
  const fila = (k, v, fuerte) => `<tr><td>${fuerte ? "<strong>" : ""}${escapeHtml(k)}${fuerte ? "</strong>" : ""}</td><td style="text-align:right;">${fuerte ? "<strong>" : ""}${escapeHtml(v)}${fuerte ? "</strong>" : ""}</td></tr>`;

  // Encabezado: mismo tamaño de letra que el ticket de venta (antes la
  // dirección y los teléfonos salían más grandes que el nombre del negocio)
  let encabezado = `<div class="th-center th-big">${escapeHtml(cfg.nombre)}</div>`;
  encabezado += `<div class="th-center" style="font-size:10.5pt;font-weight:bold;margin-top:0.5mm;">${escapeHtml(d.titulo)}</div>`;
  if (cfg.direccion) encabezado += `<div class="th-center" style="font-size:9pt;font-weight:bold;">${escapeHtml(cfg.direccion)}</div>`;
  const telefonos = [cfg.telefono1, cfg.telefono2].filter(Boolean).join(" · ");
  if (telefonos) encabezado += `<div class="th-center" style="font-size:9pt;font-weight:bold;">Tel: ${escapeHtml(telefonos)}</div>`;

  const bloquesMetodos = d.metodos.map(m => `
      <hr class="th-sep">
      <div><strong>${escapeHtml(m.nombre)}</strong></div>
      <table>${COLS_CIERRE}<tbody>
        ${m.lineas.map(([k, v]) => fila("  " + k, v)).join("").replace(/<td>  /g, "<td>&nbsp;&nbsp;")}
        ${fila("Esperado", plataTicket(m.esperado), true)}
        ${fila(m.etiquetaContado, plataTicket(m.contado))}
        ${fila(etiquetaDiferencia(m.diferencia), diferenciaTicket(m.diferencia))}
      </tbody></table>`).join("");

  return `
    <div class="thermal-receipt">
      ${encabezado}
      <hr class="th-sep-solid">
      <table>${COLS_CIERRE}<tbody>${d.datos.map(([k, v]) => fila(k, v)).join("")}</tbody></table>
      <hr class="th-sep">
      <table>${COLS_CIERRE}<tbody>
        ${fila("Ventas del día (" + d.cantidadVentas + ")", plataTicket(d.totalVentas), true)}
        ${d.cantidadAnuladas ? fila("Anuladas (no suman)", String(d.cantidadAnuladas)) : ""}
      </tbody></table>
      ${bloquesMetodos}
      <hr class="th-sep-solid">
      <table>${COLS_CIERRE}<tbody>
        ${fila("TOTAL ESPERADO", plataTicket(d.total.esperado), true)}
        ${fila("TOTAL CONTADO", plataTicket(d.total.contado), true)}
        ${fila(etiquetaDiferencia(d.total.diferencia).toUpperCase(), diferenciaTicket(d.total.diferencia), true)}
      </tbody></table>
      <hr class="th-sep">
      ${resumen.observaciones ? `<div>Obs: ${escapeHtml(resumen.observaciones)}</div><hr class="th-sep">` : ""}
      ${d.esPrevio ? `<div class="th-center">Cierre no guardado: solo para revisar</div>` : ""}
      <div class="th-footer">Cierre generado por ${escapeHtml(cfg.nombre)} POS</div>
      <br><br>
    </div>`;
}

/**
 * Imprime el cierre de caja actualmente visible en pantalla.
 * Usa los montos "esperado" ya cargados (cierreCajaResumenActual)
 * y los montos "contado" tal cual están en los inputs en este momento,
 * sin necesidad de haber guardado el cierre primero.
 */
function imprimirCierreCaja() {
  if (!cierreCajaResumenActual) {
    toast("Esperá a que se calculen las ventas del día", "error");
    return;
  }

  const r = cierreCajaResumenActual;
  const esp = r.esperado || {};
  const mov = r.movimientosCaja || {};
  const pc = r.pagosClientes || {};
  const ingresos = mov.ingresosPorFormaPago || {};
  const egresos = mov.egresosPorFormaPago || {};
  const ventas = { EFECTIVO: r.ventasEfectivo, TRANSFERENCIA: r.ventasTransferencia, TARJETA: r.ventasTarjeta };
  const contados = {
    EFECTIVO:      Number(document.getElementById("ccEfectivoContado").value || 0),
    TRANSFERENCIA: Number(document.getElementById("ccTransferenciaContado").value || 0),
    TARJETA:       Number(document.getElementById("ccTarjetaContado").value || 0)
  };
  const observaciones = document.getElementById("ccObservaciones").value || "";

  const metodos = ["EFECTIVO", "TRANSFERENCIA", "TARJETA"].map(k => {
    const esperado = Number(esp[k]) || 0;
    return {
      clave: k,
      ventas: Number(ventas[k]) || 0, cobros: Number(pc[k]) || 0,
      ingresos: Number(ingresos[k]) || 0, egresos: Number(egresos[k]) || 0,
      esperado, contado: contados[k], diferencia: contados[k] - esperado
    };
  });
  const totalEsperado = metodos.reduce((a, m) => a + m.esperado, 0);
  const totalContado  = metodos.reduce((a, m) => a + m.contado, 0);

  const existente = r.cierreExistente;
  const resumen = {
    fecha: r.fecha,
    cierreId: existente && existente.CIERRE_ID ? existente.CIERRE_ID : "",
    esPrevio: !existente,
    // Antes, sin cierre guardado, salía "ADMIN" aunque hubiera entrado otra persona
    vendedor: (existente && existente.VENDEDOR) || sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("usuarioLogueado") || "—",
    observaciones,
    cantidadVentas: r.cantidadVentas,
    cantidadAnuladas: r.cantidadAnuladas,
    totalVentas: r.totalVentas,
    metodos,
    total: { esperado: totalEsperado, contado: totalContado, diferencia: totalContado - totalEsperado }
  };

  // Mismo camino que el ticket de venta: USB → impresora de Windows directa → diálogo
  // (antes el cierre no usaba la impresión directa y por diálogo pedía 29,7 cm de papel)
  const bytes = buildThermalCierreESCPOS(resumen);
  _imprimirTicketTermico({ html: buildThermalCierreHTML(resumen), bytes, alto: altoTicketDesdeEscPos(bytes, 0, 6.3) });
}

/* ===================================================================
   BARCODE SCANNER SUPPORT + ATAJOS DE TECLADO DEL POS
=================================================================== */

let scanBuffer    = "";
let lastKeyTime   = 0;
const SCAN_KEY_THRESHOLD_MS = 40;

function setupScannerListener() {
  document.addEventListener("keydown", (e) => {
    // ---- Enter/Esc en modal de finalizar venta ----
    const modalFinalizar = document.getElementById("modalFinalizarVentaBackdrop");
    if (modalFinalizar && modalFinalizar.classList.contains("show")) {
      if (e.key === "Enter") { e.preventDefault(); confirmarFinalizarVenta(); return; }
      if (e.key === "Escape") { e.preventDefault(); cerrarModalFinalizarVenta(); return; }
    }

    // ---- Enter/Esc en modal de recibo ----
    const recibo = document.getElementById("receiptBackdrop");
    if (recibo && recibo.classList.contains("show")) {
      if (e.key === "Enter") { e.preventDefault(); imprimirUltimoRecibo(); return; }
      if (e.key === "Escape") { e.preventDefault(); cerrarRecibo(); return; }
    }

    const posSection = document.getElementById("pos");
    const posVisible = posSection && posSection.style.display === "block";
    if (!posVisible) return;

    // Ctrl+K / Cmd+K: foco directo al buscador, funciona incluso si el foco está en otro lado
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
      e.preventDefault();
      const input = document.getElementById("posBusqueda");
      if (input) input.focus();
      return;
    }

    const activeTag  = document.activeElement ? document.activeElement.tagName : "";
    const isOurInput = document.activeElement && document.activeElement.id === "posBusqueda";

    // F2 finaliza la venta desde cualquier lado dentro del POS (incluso con el buscador enfocado)
    if (e.key === "F2") {
      e.preventDefault();
      abrirModalFinalizarVenta();
      return;
    }

    // F6/F7/F8: forma de pago, funcionan también con el foco en el buscador
    // (los números 1/2/3 solo sirven cuando no se está escribiendo).
    if (e.key === "F6" || e.key === "F7" || e.key === "F8") {
      const mapaF = { F6: "EFECTIVO", F7: "TRANSFERENCIA", F8: "TARJETA" };
      const btnF = document.querySelector(`.pay-method-btn[data-val="${mapaF[e.key]}"]`);
      if (btnF) { e.preventDefault(); elegirFormaPago(btnF, mapaF[e.key]); toast("Forma de pago: " + mapaF[e.key].toLowerCase(), "success"); }
      return;
    }

    // F1: ayuda con todos los atajos
    if (e.key === "F1") {
      e.preventDefault();
      mostrarAyudaAtajosPOS();
      return;
    }
    // Esc con la ayuda abierta: solo la cierra (no vacía el ticket)
    if (e.key === "Escape" && cerrarAyudaAtajosPOS()) {
      e.preventDefault();
      return;
    }

    if (!isOurInput && (activeTag === "INPUT" || activeTag === "SELECT" || activeTag === "TEXTAREA")) return;

    // ---- Atajos que solo aplican cuando NO se está escribiendo en el buscador ----
    if (!isOurInput) {
      if (e.key === "ArrowRight" || e.key === "ArrowLeft" || e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        moverFocoPosGrid(e.key);
        return;
      }
      if (e.key === "Enter" && posTileFocusIdx >= 0) {
        e.preventDefault();
        agregarProductoEnFoco();
        return;
      }
      if (e.key === "+" || e.key === "=") {
        if (ultimoCodigoAgregadoPOS) { e.preventDefault(); cambiarCantidadPOS(ultimoCodigoAgregadoPOS, 1); }
        return;
      }
      if (e.key === "-") {
        if (ultimoCodigoAgregadoPOS) { e.preventDefault(); cambiarCantidadPOS(ultimoCodigoAgregadoPOS, -1); }
        return;
      }
      if (e.key === "1" || e.key === "2" || e.key === "3") {
        const mapa = { "1": "EFECTIVO", "2": "TRANSFERENCIA", "3": "TARJETA" };
        const btn = document.querySelector(`.pay-method-btn[data-val="${mapa[e.key]}"]`);
        if (btn) { e.preventDefault(); elegirFormaPago(btn, mapa[e.key]); }
        return;
      }
      if (e.key === "?") {
        e.preventDefault();
        mostrarAyudaAtajosPOS();
        return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        vaciarTicketPOS();
        return;
      }
    }

    const now     = Date.now();
    const elapsed = now - lastKeyTime;
    lastKeyTime   = now;

    if (e.key === "Enter") {
      // Si el foco está en el buscador del POS, onPosInputKeyup (el
      // oninput="" de ese mismo <input>) ya se encarga de procesar
      // este Enter — dejarlo pasar también acá disparaba
      // agregarProductoPorCodigo() DOS VECES para el mismo escaneo
      // (una desde cada listener), una condición de carrera que podía
      // hacer fallar una de las dos búsquedas y mostrar "Producto no
      // encontrado" aunque el producto sí se agregara por la otra vía.
      if (isOurInput) { scanBuffer = ""; return; }

      if (scanBuffer.length >= 3) { agregarProductoPorCodigo(scanBuffer); setScannerStatus("listening"); }
      scanBuffer = "";
      return;
    }

    if (e.key.length === 1) {
      if (elapsed > 250) scanBuffer = "";
      scanBuffer += e.key;
      if (elapsed < SCAN_KEY_THRESHOLD_MS && scanBuffer.length >= 3) setScannerStatus("listening");
    }
  });

  setInterval(() => {
    if (Date.now() - lastKeyTime > 1200) setScannerStatus("idle");
  }, 500);
}

/** Moves the keyboard focus highlight across the visible product tiles in the grid */
function moverFocoPosGrid(tecla) {
  const tiles = Array.from(document.querySelectorAll("#posProductGrid .product-tile:not(.disabled)"));
  if (tiles.length === 0) return;

  // Estimate columns from actual layout so left/right and up/down both feel natural
  const primerTop = tiles[0].getBoundingClientRect().top;
  let columnas = 1;
  for (let i = 1; i < tiles.length; i++) {
    if (Math.abs(tiles[i].getBoundingClientRect().top - primerTop) > 2) { columnas = i; break; }
  }

  let nuevoIdx = posTileFocusIdx;
  if (nuevoIdx < 0) {
    nuevoIdx = 0;
  } else if (tecla === "ArrowRight") {
    nuevoIdx = Math.min(tiles.length - 1, nuevoIdx + 1);
  } else if (tecla === "ArrowLeft") {
    nuevoIdx = Math.max(0, nuevoIdx - 1);
  } else if (tecla === "ArrowDown") {
    nuevoIdx = Math.min(tiles.length - 1, nuevoIdx + columnas);
  } else if (tecla === "ArrowUp") {
    nuevoIdx = Math.max(0, nuevoIdx - columnas);
  }

  tiles.forEach(t => t.classList.remove("kbd-focus"));
  posTileFocusIdx = nuevoIdx;
  const tileActivo = tiles[posTileFocusIdx];
  if (tileActivo) {
    tileActivo.classList.add("kbd-focus");
    tileActivo.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}

/** Adds to the ticket whichever tile currently has keyboard focus */
function agregarProductoEnFoco() {
  const tiles = Array.from(document.querySelectorAll("#posProductGrid .product-tile:not(.disabled)"));
  const tile = tiles[posTileFocusIdx];
  if (!tile || !tile.dataset.codigo) return;
  agregarProductoPOS(tile.dataset.codigo, false);
}

function setScannerStatus(estado) {
  const el = document.getElementById("scannerStatus");
  if (!el) return;
  if (estado === "listening") {
    el.className = "scanner-status listening";
    el.innerHTML = `<span class="dot"></span> Leyendo...`;
  } else if (estado === "camera") {
    el.className = "scanner-status camera";
    el.innerHTML = `<span class="dot"></span> Cámara activa`;
  } else {
    el.className = "scanner-status idle";
    el.innerHTML = `<span class="dot"></span> Listo`;
  }
}

/* ---- shortcuts help modal ---- */

function toggleShortcutsHelp() {
  document.getElementById("shortcutsHelpBackdrop").classList.toggle("show");
}
function cerrarShortcutsHelp(e) {
  if (e.target.id === "shortcutsHelpBackdrop") e.target.classList.remove("show");
}

/* ---- camera-based scanning ----
   iOS Safari (y Firefox) no implementan BarcodeDetector nativo, así que
   usamos ese cuando existe (Chrome/Edge en Android y desktop, más rápido
   por ser nativo) y si no, caemos a ZXing-js, que decodifica en JS puro
   y funciona en cualquier navegador con getUserMedia, incluido iOS Safari. */

let camaraStream = null;
let camaraDetectorTimer = null;   // usado por la rama BarcodeDetector nativo
let zxingReader = null;           // usado por la rama fallback ZXing
let camaraScanActivo = false;     // evita agregar el mismo código dos veces al cerrar

async function abrirCamaraScan() {
  const backdrop  = document.getElementById("scanModalBackdrop");
  const videoWrap = document.getElementById("scanVideoWrap");
  const video     = document.getElementById("scanVideo");

  backdrop.classList.add("show");
  camaraScanActivo = true;

  // getUserMedia requiere HTTPS (o localhost) — en iOS, además, ni siquiera
  // existe el objeto si la página no es segura, así que lo detectamos antes
  // de pedir permiso de cámara para dar un mensaje claro.
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    videoWrap.innerHTML = `
      <div class="scan-unsupported">
        El escaneo por cámara requiere una conexión segura (HTTPS).<br>
        Usá un lector USB o accedé al panel por https.
      </div>`;
    return;
  }

  try {
    camaraStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" } },
      audio: false
    });
    video.srcObject = camaraStream;
    // iOS Safari necesita el play() explícito incluso con autoplay+playsinline
    try { await video.play(); } catch (e) { /* algunos navegadores ya lo reproducen solos */ }

    setScannerStatus("camera");
    document.getElementById("btnCameraScan").classList.add("active");

    if ("BarcodeDetector" in window) {
      iniciarDeteccionNativa(video);
    } else {
      await iniciarDeteccionZXing(video);
    }

  } catch (error) {
    console.error("Error de cámara:", error);
    const permisoDenegado = error && (error.name === "NotAllowedError" || error.name === "PermissionDeniedError");
    videoWrap.innerHTML = permisoDenegado
      ? `<div class="scan-unsupported">
           Se denegó el acceso a la cámara.<br>
           En iOS: Ajustes → Safari (o la app) → Cámara → Permitir.<br>
           En Android: tocá el candado en la barra de direcciones → Permisos → Cámara.
         </div>`
      : `<div class="scan-unsupported">
           No se pudo acceder a la cámara.<br>
           Revisá los permisos del navegador e intentá de nuevo.
         </div>`;
  }
}

function iniciarDeteccionNativa(video) {
  const detector = new BarcodeDetector({
    formats: ["ean_13","ean_8","upc_a","upc_e","code_128","code_39","qr_code","itf"]
  });

  camaraDetectorTimer = setInterval(async () => {
    try {
      const codigos = await detector.detect(video);
      if (codigos.length > 0) {
        const valor = codigos[0].rawValue;
        cerrarCamaraScan();
        agregarProductoPorCodigo(valor);
      }
    } catch (err) { /* frame errors expected */ }
  }, 350);
}

async function iniciarDeteccionZXing(video) {
  if (typeof ZXingBrowser === "undefined" && typeof ZXing === "undefined") {
    document.getElementById("scanVideoWrap").innerHTML = `
      <div class="scan-unsupported">
        No se pudo cargar el motor de escaneo.<br>
        Verificá tu conexión a internet e intentá de nuevo.
      </div>`;
    return;
  }

  try {
    // El paquete UMD de zxing-library expone la librería como `ZXing`
    const ZXingLib = (typeof ZXingBrowser !== "undefined") ? ZXingBrowser : ZXing;
    const hints = new Map();
    const formatos = [
      ZXingLib.BarcodeFormat.EAN_13, ZXingLib.BarcodeFormat.EAN_8,
      ZXingLib.BarcodeFormat.UPC_A, ZXingLib.BarcodeFormat.UPC_E,
      ZXingLib.BarcodeFormat.CODE_128, ZXingLib.BarcodeFormat.CODE_39,
      ZXingLib.BarcodeFormat.QR_CODE, ZXingLib.BarcodeFormat.ITF
    ];
    hints.set(ZXingLib.DecodeHintType.POSSIBLE_FORMATS, formatos);
    hints.set(ZXingLib.DecodeHintType.TRY_HARDER, true);

    zxingReader = new ZXingLib.BrowserMultiFormatReader(hints);

    // decodeFromVideoElementContinuously reutiliza el stream de video ya
    // asignado a <video> (no vuelve a pedir permiso de cámara) y llama al
    // callback en cada frame; seguimos escaneando hasta encontrar un match.
    zxingReader.decodeFromVideoElementContinuously(video, (result, err) => {
      if (result && camaraScanActivo) {
        const valor = result.getText ? result.getText() : result.text;
        cerrarCamaraScan();
        agregarProductoPorCodigo(valor);
      }
      // NotFoundException se dispara en casi todos los frames sin código
      // visible: es esperable, no un error real.
    });
  } catch (error) {
    console.error("Error iniciando ZXing:", error);
    document.getElementById("scanVideoWrap").innerHTML = `
      <div class="scan-unsupported">
        No se pudo iniciar el escaneo en este navegador.<br>
        Usá un lector USB, o Chrome/Edge en Android.
      </div>`;
  }
}

function cerrarCamaraScan() {
  const backdrop = document.getElementById("scanModalBackdrop");
  backdrop.classList.remove("show");
  camaraScanActivo = false;

  if (camaraDetectorTimer) { clearInterval(camaraDetectorTimer); camaraDetectorTimer = null; }
  if (zxingReader) {
    try { zxingReader.reset(); } catch (e) { /* noop */ }
    zxingReader = null;
  }
  if (camaraStream) { camaraStream.getTracks().forEach(t => t.stop()); camaraStream = null; }

  document.getElementById("btnCameraScan").classList.remove("active");
  setScannerStatus("idle");

  document.getElementById("scanVideoWrap").innerHTML = `
    <video id="scanVideo" autoplay playsinline muted></video>
    <div class="scan-reticle"></div>`;
}

/* ===================================================================
   CIERRE DE CAJA — RECONCILIACION DIARIA
   Compara lo esperado (según VENTAS_LOCAL, por forma de pago) contra
   lo contado físicamente al cierre del día. Guarda el resultado en
   la hoja "CIERRES_CAJA" vía el backend.
=================================================================== */

let cierreCajaResumenActual = null; // último resumen "esperado" cargado del backend

/** Loads today's expected totals by payment method and pre-fills the form */
async function cargarResumenCierreCaja(fecha, forzar = false) {
  const estadoEl = document.getElementById("cierreCajaEstado");

  // Caché por fecha — 90 segundos (el cierre rara vez cambia en segundos,
  // pero queremos datos frescos si vuelven a entrar después de registrar movimientos)
  const CACHE_KEY = "vpos_cache_cierre_" + (fecha || "ayer");
  const CACHE_TTL = 90 * 1000;

  try {
    const raw = forzar ? null : localStorage.getItem(CACHE_KEY);
    if (raw) {
      const { ts, data } = JSON.parse(raw);
      if (Date.now() - ts < CACHE_TTL) {
        aplicarDatosCierreCaja(data);
        if (estadoEl) estadoEl.textContent = "";
        return;
      }
    }
  } catch(e) {}

  if (estadoEl) estadoEl.textContent = "Calculando ventas del día...";

  try {
    const url = API_URL + "?action=cierreCajaResumen" + (fecha ? "&fecha=" + encodeURIComponent(fecha) : "");
    const response = await fetch(url);
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo calcular el resumen de caja", "error");
      if (estadoEl) estadoEl.textContent = "";
      return;
    }

    try { localStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), data })); } catch(e) {}
    aplicarDatosCierreCaja(data);
    if (estadoEl) estadoEl.textContent = "";

  } catch (error) {
    console.error("Error cierre de caja:", error);
    toast("Error al calcular el cierre de caja", "error");
    if (estadoEl) estadoEl.textContent = "";
  }
}

function invalidarCacheCierre(fecha) {
  try { localStorage.removeItem("vpos_cache_cierre_" + (fecha || "ayer")); } catch(e) {}
}

function aplicarDatosCierreCaja(data) {
  const estadoEl = document.getElementById("cierreCajaEstado");
  cierreCajaResumenActual = data;

  actualizarElemento("ccFechaLabel", formatearFechaCierre(data.fecha));
  actualizarElemento("ccCantidadVentas", data.cantidadVentas + (data.cantidadVentas === 1 ? " venta" : " ventas"));
  actualizarElemento("ccTotalVentas", "$" + Number(data.totalVentas || 0).toLocaleString("es-AR"));

  actualizarElemento("ccEfectivoEsperado",      "$" + Number(data.esperado.EFECTIVO).toLocaleString("es-AR"));
  actualizarElemento("ccTransferenciaEsperado", "$" + Number(data.esperado.TRANSFERENCIA).toLocaleString("es-AR"));
  actualizarElemento("ccTarjetaEsperado",       "$" + Number(data.esperado.TARJETA).toLocaleString("es-AR"));
  actualizarElemento("ccTotalEsperado",         "$" + Number(data.esperado.TOTAL).toLocaleString("es-AR"));

    // Desglose de movimientos y pagos de clientes — cada forma de pago
    // muestra solo lo que le corresponde a ella.
    const movDetalleEl = document.getElementById("ccMovimientosDetalle");
    const mov = data.movimientosCaja;
    const pfp = mov?.porFormaPago || {};
    const movEfectivo = Number(pfp.EFECTIVO || 0);
    const movTransferencia = Number(pfp.TRANSFERENCIA || 0);
    const movTarjeta = Number(pfp.TARJETA || 0);

    // Ingresos y egresos BRUTOS de cada forma de pago (no el neto). Si
    // un mismo día hay un ingreso y un egreso por el mismo importe en
    // la misma forma de pago, el neto da $0 — pero eso no significa
    // que no hubo ningún movimiento, así que mostrar el neto acá
    // ocultaría un ingreso y un egreso reales. Por eso se usan los
    // valores brutos que ya vienen separados desde el backend.
    const ingPorFormaPago = mov?.ingresosPorFormaPago || {};
    const egrPorFormaPago = mov?.egresosPorFormaPago || {};

    const pc = data.pagosClientes || {};
    const pcEfectivo = Number(pc.EFECTIVO || 0);
    const pcTransferencia = Number(pc.TRANSFERENCIA || 0);
    const pcTarjeta = Number(pc.TARJETA || 0);

    function mostrarLineaPagos(idLinea, idMonto, monto) {
      const lineaEl = document.getElementById(idLinea);
      if (!lineaEl) return;
      if (monto > 0) {
        actualizarElemento(idMonto, "$" + monto.toLocaleString("es-AR"));
        lineaEl.style.display = "flex";
      } else {
        lineaEl.style.display = "none";
      }
    }

    // Desglose de efectivo
    if (movDetalleEl && (movEfectivo !== 0 || pcEfectivo !== 0)) {
      actualizarElemento("ccVentasEfectivo", "$" + Number(data.ventasEfectivo || 0).toLocaleString("es-AR"));
      mostrarLineaPagos("ccPagosClientesEfectivoLinea", "ccPagosClientesEfectivo", pcEfectivo);
      const ingEf = Number(ingPorFormaPago.EFECTIVO || 0);
      const egrEf = Number(egrPorFormaPago.EFECTIVO || 0);
      actualizarElemento("ccMovIngresos", ingEf > 0 ? "+$" + ingEf.toLocaleString("es-AR") : "$0");
      actualizarElemento("ccMovEgresos",  egrEf > 0 ? "−$" + egrEf.toLocaleString("es-AR") : "$0");
      movDetalleEl.style.display = "flex";
    } else if (movDetalleEl) {
      movDetalleEl.style.display = "none";
    }

    // Desglose de transferencia
    const movTransfEl = document.getElementById("ccMovimientosDetalleTransf");
    if (movTransfEl) {
      if (movTransferencia !== 0 || pcTransferencia !== 0) {
        const ingTr = Number(ingPorFormaPago.TRANSFERENCIA || 0);
        const egrTr = Number(egrPorFormaPago.TRANSFERENCIA || 0);
        actualizarElemento("ccVentasTransferencia", "$" + Number(data.ventasTransferencia || 0).toLocaleString("es-AR"));
        mostrarLineaPagos("ccPagosClientesTransfLinea", "ccPagosClientesTransf", pcTransferencia);
        actualizarElemento("ccMovIngresosTransf", ingTr > 0 ? "+$" + ingTr.toLocaleString("es-AR") : "$0");
        actualizarElemento("ccMovEgresosTransf",  egrTr > 0 ? "−$" + egrTr.toLocaleString("es-AR") : "$0");
        movTransfEl.style.display = "flex";
      } else {
        movTransfEl.style.display = "none";
      }
    }

    // Desglose de tarjeta
    const movTarjetaEl = document.getElementById("ccMovimientosDetalleTarjeta");
    if (movTarjetaEl) {
      if (movTarjeta !== 0 || pcTarjeta !== 0) {
        const ingTa = Number(ingPorFormaPago.TARJETA || 0);
        const egrTa = Number(egrPorFormaPago.TARJETA || 0);
        actualizarElemento("ccVentasTarjeta", "$" + Number(data.ventasTarjeta || 0).toLocaleString("es-AR"));
        mostrarLineaPagos("ccPagosClientesTarjetaLinea", "ccPagosClientesTarjeta", pcTarjeta);
        actualizarElemento("ccMovIngresosTarjeta", ingTa > 0 ? "+$" + ingTa.toLocaleString("es-AR") : "$0");
        actualizarElemento("ccMovEgresosTarjeta",  egrTa > 0 ? "−$" + egrTa.toLocaleString("es-AR") : "$0");
        movTarjetaEl.style.display = "flex";
      } else {
        movTarjetaEl.style.display = "none";
      }
    }

    // If a closing already exists for this date, pre-fill counted amounts so the user can review/edit
    const inputEfectivo      = document.getElementById("ccEfectivoContado");
    const inputTransferencia = document.getElementById("ccTransferenciaContado");
    const inputTarjeta       = document.getElementById("ccTarjetaContado");

    if (data.yaCerrado && data.cierreExistente) {
      const c = data.cierreExistente;
      if (inputEfectivo)      inputEfectivo.value      = Number(c.EFECTIVO_CONTADO || 0);
      if (inputTransferencia) inputTransferencia.value = Number(c.TRANSFERENCIA_CONTADO || 0);
      if (inputTarjeta)       inputTarjeta.value        = Number(c.TARJETA_CONTADO || 0);
      const obsEl = document.getElementById("ccObservaciones");
      if (obsEl) obsEl.value = c.OBSERVACIONES || "";
      if (estadoEl) estadoEl.innerHTML = `⚠️ Ya existe un cierre guardado para esta fecha. Guardar de nuevo lo actualizará.`;
    } else {
      if (inputEfectivo)      inputEfectivo.value      = "";
      if (inputTransferencia) inputTransferencia.value = "";
      if (inputTarjeta)       inputTarjeta.value        = "";
      if (estadoEl) estadoEl.textContent = "";
    }

    // Caja de dólares (depósito)
    const dol = data.dolares || null;
    const fmtUsd = (v) => "US$ " + (Number(v) || 0).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    actualizarElemento("ccUsdEsperado", fmtUsd(dol && dol.esperadoUsd));
    actualizarElemento("ccUsdVentas", fmtUsd(dol && dol.ventasUsd));
    actualizarElemento("ccUsdPagos", fmtUsd(dol && dol.pagosClientesUsd));
    const lineaDolEf = document.getElementById("ccEfectivoDolaresLinea");
    if (lineaDolEf) {
      const v = Number(dol && dol.efectivoArs) || 0;
      lineaDolEf.style.display = v !== 0 ? "flex" : "none";
      actualizarElemento("ccEfectivoDolares", (v < 0 ? "−$" : "$") + Math.abs(v).toLocaleString("es-AR"));
      if (v !== 0 && movDetalleEl) {
        actualizarElemento("ccVentasEfectivo", "$" + Number(data.ventasEfectivo || 0).toLocaleString("es-AR"));
        movDetalleEl.style.display = "flex";
      }
    }
    const inputUsd = document.getElementById("ccUsdContado");
    if (inputUsd) inputUsd.value = data.yaCerrado && data.cierreExistente && data.cierreExistente.USD_CONTADO !== undefined && data.cierreExistente.USD_CONTADO !== ""
      ? Number(data.cierreExistente.USD_CONTADO) : "";

    calcularDiferenciasCierreCaja();
    cargarHistorialCierres();
}

function formatearFechaCierre(fechaStr) {
  if (!fechaStr) return "—";
  const [y, m, d] = fechaStr.split("-");
  return `${d}/${m}/${y}`;
}

/** Recalculates differences live as the user types counted amounts */
function calcularDiferenciasCierreCaja() {
  if (!cierreCajaResumenActual) return;

  const efectivoContado      = Number(document.getElementById("ccEfectivoContado").value || 0);
  const transferenciaContado = Number(document.getElementById("ccTransferenciaContado").value || 0);
  const tarjetaContado       = Number(document.getElementById("ccTarjetaContado").value || 0);

  const esp = cierreCajaResumenActual.esperado;

  const difEfectivo      = efectivoContado - esp.EFECTIVO;
  const difTransferencia = transferenciaContado - esp.TRANSFERENCIA;
  const difTarjeta       = tarjetaContado - esp.TARJETA;
  const difTotal         = difEfectivo + difTransferencia + difTarjeta;

  pintarDiferencia("ccEfectivoDif", difEfectivo);
  pintarDiferencia("ccTransferenciaDif", difTransferencia);
  pintarDiferencia("ccTarjetaDif", difTarjeta);
  pintarDiferencia("ccTotalDif", difTotal, true);

  const totalContado = efectivoContado + transferenciaContado + tarjetaContado;
  actualizarElemento("ccTotalContado", "$" + totalContado.toLocaleString("es-AR"));

  // Dólares: diferencia aparte (no se mezcla con los pesos)
  const usdEl = document.getElementById("ccUsdDif");
  if (usdEl) {
    const esperadoUsd = Number((cierreCajaResumenActual.dolares || {}).esperadoUsd) || 0;
    const contadoUsd = Number((document.getElementById("ccUsdContado") || {}).value || 0);
    const dif = Math.round((contadoUsd - esperadoUsd) * 100) / 100;
    usdEl.textContent = (dif > 0 ? "+" : dif < 0 ? "−" : "") + "US$ " + Math.abs(dif).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    usdEl.classList.remove("cc-dif-ok", "cc-dif-sobra", "cc-dif-falta");
    usdEl.classList.add(Math.abs(dif) < 0.01 ? "cc-dif-ok" : dif > 0 ? "cc-dif-sobra" : "cc-dif-falta");
  }
}

/** Parámetros de la caja de dólares para guardar el cierre (solo depósito) */
function paramsCierreUsd() {
  if (!(typeof esModoDeposito === "function" && esModoDeposito()) || !cierreCajaResumenActual) return {};
  const contado = (document.getElementById("ccUsdContado") || {}).value;
  return { usdEsperado: Number((cierreCajaResumenActual.dolares || {}).esperadoUsd) || 0, usdContado: Number(contado) || 0 };
}

function pintarDiferencia(id, valor, esTotal) {
  const el = document.getElementById(id);
  if (!el) return;

  const signo = valor > 0 ? "+" : "";
  el.textContent = signo + "$" + Math.round(valor).toLocaleString("es-AR");

  el.classList.remove("cc-dif-ok", "cc-dif-sobra", "cc-dif-falta");

  if (Math.abs(valor) < 1)      el.classList.add("cc-dif-ok");
  else if (valor > 0)           el.classList.add("cc-dif-sobra");
  else                          el.classList.add("cc-dif-falta");
}

/** Saves the daily reconciliation to the backend (CIERRES_CAJA sheet) */
async function guardarCierreCajaForm() {
  if (!cierreCajaResumenActual) { toast("Esperá a que se calculen las ventas del día", "error"); return; }

  const efectivoContado      = document.getElementById("ccEfectivoContado").value;
  const transferenciaContado = document.getElementById("ccTransferenciaContado").value;
  const tarjetaContado       = document.getElementById("ccTarjetaContado").value;
  const observaciones        = document.getElementById("ccObservaciones").value;

  const usdContadoTxt = (document.getElementById("ccUsdContado") || {}).value || "";
  if (efectivoContado === "" && transferenciaContado === "" && tarjetaContado === "" && usdContadoTxt === "") {
    toast("Ingresá al menos un monto contado", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarCierreCaja");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }

  try {
    const params = new URLSearchParams({
      action: "guardarCierreCaja",
      fecha: cierreCajaResumenActual.fecha,
      efectivoEsperado:      cierreCajaResumenActual.esperado.EFECTIVO,
      efectivoContado:       efectivoContado || 0,
      transferenciaEsperado: cierreCajaResumenActual.esperado.TRANSFERENCIA,
      transferenciaContado:  transferenciaContado || 0,
      tarjetaEsperado:       cierreCajaResumenActual.esperado.TARJETA,
      tarjetaContado:        tarjetaContado || 0,
      observaciones:         observaciones || "",
      ...paramsCierreUsd()
    });

    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar el cierre de caja", "error");
      return;
    }

    toast(data.actualizado ? "Cierre de caja actualizado" : "Cierre de caja guardado", "success");
    cargarHistorialCierres();

  } catch (error) {
    console.error("Error al guardar cierre de caja:", error);
    toast("Error de conexión al guardar el cierre", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Loads the recent reconciliation history table */
async function cargarHistorialCierres() {
  const CACHE_KEY = "vpos_cache_historial_cierres";
  const CACHE_TTL = 5 * 60 * 1000;
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (raw) {
      const { ts, data } = JSON.parse(raw);
      if (Date.now() - ts < CACHE_TTL) { renderHistorialCierres(data); return; }
    }
  } catch(e) {}
  const tbody = document.getElementById("tablaCierresCaja");
  if (!tbody) return;

  try {
    const response = await fetchAPI(API_URL + "?action=historialCierres");
    const data = await response.json();
    const lista = data.cierres || [];

    if (lista.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3">Todavía no hay cierres guardados</td></tr>`;
      return;
    }

    let html = "";
    lista.forEach(c => {
      const fecha = formatearFechaCierre(
        c.FECHA instanceof Date
          ? _fechaISOEnZona(c.FECHA)
          : String(c.FECHA).slice(0, 10)
      );
      const totalDif = Number(c.TOTAL_DIFERENCIA || 0);
      const claseDif = Math.abs(totalDif) < 1 ? "cc-dif-ok" : (totalDif > 0 ? "cc-dif-sobra" : "cc-dif-falta");
      const signo = totalDif > 0 ? "+" : "";

      html += `
        <tr>
          <td>${escapeHtml(fecha)}</td>
          <td class="money">$${Number(c.TOTAL_ESPERADO || 0).toLocaleString("es-AR")}</td>
          <td class="money">$${Number(c.TOTAL_CONTADO || 0).toLocaleString("es-AR")}</td>
          <td class="money ${claseDif}">${signo}$${Math.round(totalDif).toLocaleString("es-AR")}</td>
          <td>${escapeHtml(c.VENDEDOR || "—")}</td>
          <td>${escapeHtml(c.OBSERVACIONES || "—")}</td>
        </tr>`;
    });

    tbody.innerHTML = html;

  } catch (error) {
    console.error("Error al cargar historial de cierres:", error);
    tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3">Error al cargar el historial</td></tr>`;
  }
}

/* ===================== REPORTES ===================== */

/** Lee el filtro de fecha compartido por los 6 reportes. Si está vacío, el backend usa el mes en curso. */
function obtenerRangoReportes() {
  const desde = document.getElementById("repDesde").value;
  const hasta = document.getElementById("repHasta").value;
  let qs = "";
  if (desde) qs += "&desde=" + encodeURIComponent(desde);
  if (hasta) qs += "&hasta=" + encodeURIComponent(hasta);
  return qs;
}

/* ---- Pestañas de Reportes: cada una carga recién cuando se abre, y
   solo si el período cambió desde la última vez (no se piden los 3
   grupos de datos de golpe al entrar a la sección). ---- */
let _repTabActiva = null;
const _repTabRangoCargado = { ventas: null, productos: null, compras: null };

function _repTabPorDefecto() {
  return obtenerRolActual() === "deposito" ? "compras" : "ventas";
}

function mostrarTabReportes(tab, forzar) {
  if (obtenerRolActual() === "deposito") tab = "compras";
  _repTabActiva = tab;
  _repCompletarRangoPorDefecto();
  document.querySelectorAll("#repTabs .rep-tab").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll("#reportes .rep-pane").forEach(p => p.classList.toggle("active", p.id === "repPane-" + tab));

  const rango = obtenerRangoReportes();
  if (!forzar && _repTabRangoCargado[tab] !== null && _repTabRangoCargado[tab] === rango) {
    // Ya cargada con este período: si es la de productos, redibujar los
    // gráficos (estaban ocultos y Chart.js necesita el ancho real)
    if (tab === "productos" && _repProductosDatosActuales.length) renderReporteProductos();
    return;
  }
  _repTabRangoCargado[tab] = rango;

  if (tab === "ventas") {
    cargarReportesVentas(forzar);
  } else if (tab === "productos") {
    cargarReporteProductos(forzar);
  } else if (tab === "compras") {
    cargarReporteCompras(forzar);
  }
}

/* Los reportes leen hojas enteras en Apps Script: con planillas grandes
   pueden tardar bastante más que una lectura común. Más margen de espera
   y UN solo intento (reintentar al vencer el tiempo solo duplicaba la
   carga sobre el servidor y volvía a fallar). */
const _REP_OPC_FETCH = { timeoutMs: 90000, reintentos: 1 };

function _repMensajeError(error) {
  if (error && error.name === "AbortError") return "Google tardó demasiado en responder. Probá de nuevo en un momento o con un período más corto.";
  if (error && /Failed to fetch|NetworkError/i.test(String(error.message))) return "Sin conexión con el servidor. Revisá internet.";
  return (error && error.message) || "Error al cargar el reporte";
}

/** Pestaña Ventas: los 5 reportes en UNA sola llamada (reportesVentas). Con un Code.gs viejo, cae a los 5 pedidos de antes. */
let _repVentasGen = 0;
async function cargarReportesVentas(forzar) {
  const gen = ++_repVentasGen;
  const cuerpos = { repVentasPeriodoTabla: 4, repCategoriasTabla: 3, repFormasPagoTabla: 3, repCierresTabla: 5, repClientesTabla: 4 };
  Object.entries(cuerpos).forEach(([id, cols]) => {
    const tb = document.getElementById(id);
    if (tb) tb.innerHTML = `<tr><td colspan="${cols}" class="text-center text-muted py-3">Cargando…</td></tr>`;
  });
  const btn = document.getElementById("btnAplicarReportes");
  if (btn) btn.disabled = true;

  try {
    const url = API_URL + "?action=reportesVentas" + obtenerRangoReportes() + (forzar ? "&sinCache=1&_=" + Date.now() : "");
    const data = await _leerRespuestaJSON(await fetchAPI(url, {}, _REP_OPC_FETCH));
    if (gen !== _repVentasGen) return;

    if (!data || !data.success) {
      if (/no v[aá]lida/i.test(String((data && data.message) || ""))) {
        // Code.gs viejo: sin el endpoint combinado
        await Promise.all([cargarReporteVentasPeriodo(), cargarReporteCategorias(), cargarReporteFormasPago(), cargarReporteCierres(), cargarReporteClientes()]);
        return;
      }
      throw new Error((data && data.message) || "El servidor no devolvió los reportes");
    }
    const ok = d => (d && d.success) ? d : { success: false };
    await Promise.all([
      cargarReporteVentasPeriodo(ok(data.ventasPeriodo)),
      cargarReporteCategorias(ok(data.categorias)),
      cargarReporteFormasPago(ok(data.formasPago)),
      cargarReporteCierres(ok(data.cierres)),
      cargarReporteClientes(ok(data.clientes))
    ]);
    // Si alguno vino con error del servidor, que no quede "Cargando…"
    Object.entries(cuerpos).forEach(([id, cols]) => {
      const tb = document.getElementById(id);
      if (tb && /Cargando…/.test(tb.textContent)) tb.innerHTML = `<tr><td colspan="${cols}" class="text-center text-muted py-3">No se pudo calcular este reporte</td></tr>`;
    });
  } catch (error) {
    if (gen !== _repVentasGen) return;
    console.error("Error al cargar los reportes de ventas:", error);
    const msg = escapeHtml(_repMensajeError(error));
    Object.entries(cuerpos).forEach(([id, cols]) => {
      const tb = document.getElementById(id);
      if (tb) tb.innerHTML = `<tr><td colspan="${cols}" class="text-center text-muted py-3">${msg} <button class="btn btn-outline-primary btn-sm ms-2" onclick="cargarTodosLosReportes(true)">Reintentar</button></td></tr>`;
    });
    _repTabRangoCargado.ventas = null; // que vuelva a intentar al reabrir la pestaña
  } finally {
    if (btn) btn.disabled = false;
  }
}

/** Sin fechas elegidas: mes en curso (lo mismo que haría el backend), para que todas las pestañas usen el mismo período */
function _repCompletarRangoPorDefecto() {
  const d = document.getElementById("repDesde"), h = document.getElementById("repHasta");
  if (!d || !h) return;
  const hoy = new Date();
  const iso = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`;
  if (!d.value && !h.value) { d.value = iso(new Date(hoy.getFullYear(), hoy.getMonth(), 1)); h.value = iso(hoy); }
  else if (!h.value) h.value = iso(hoy);
  else if (!d.value) d.value = h.value;
}

/** Botón "Aplicar" (y la carga al entrar): recarga la pestaña visible; las otras se recargan al abrirlas */
function cargarTodosLosReportes(forzar) {
  _repTabRangoCargado.ventas = _repTabRangoCargado.productos = _repTabRangoCargado.compras = null;
  mostrarTabReportes(_repTabActiva || _repTabPorDefecto(), !!forzar);
}

/** Atajos de período: Hoy / 7 días / 30 días / Este mes / Mes anterior */
function repRangoRapido(tipo) {
  const hoy = new Date();
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  let desde = new Date(hoy), hasta = new Date(hoy);
  if (tipo === "7") desde.setDate(hoy.getDate() - 6);
  else if (tipo === "30") desde.setDate(hoy.getDate() - 29);
  else if (tipo === "mes") desde = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  else if (tipo === "mesAnterior") { desde = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1); hasta = new Date(hoy.getFullYear(), hoy.getMonth(), 0); }
  document.getElementById("repDesde").value = iso(desde);
  document.getElementById("repHasta").value = iso(hasta);
  cargarTodosLosReportes();
}

/** Sincroniza los inputs de fecha "Desde"/"Hasta" con el rango que devolvió el backend (cuando no se eligió nada, para que el usuario vea qué período se está mostrando). */
function sincronizarRangoReportes(desde, hasta) {
  const inputDesde = document.getElementById("repDesde");
  const inputHasta = document.getElementById("repHasta");
  if (inputDesde && !inputDesde.value) inputDesde.value = desde;
  if (inputHasta && !inputHasta.value) inputHasta.value = hasta;
}

/* ---- Cache de sesión para reportes (no persistente) ---- */
const _reporteCache = {};
function _cacheReporte(key, data) { _reporteCache[key] = { ts: Date.now(), data }; }
function _getCacheReporte(key, ttlMs = 90000) {
  const c = _reporteCache[key];
  return (c && Date.now() - c.ts < ttlMs) ? c.data : null;
}

/* ---- Selector "10 últimos / 20 últimos / todos" para cada uno de los 6 reportes ---- */
const _repLimites = { ventasPeriodo: 10, productos: 10, categorias: 10, formasPago: 10, cierres: 10, clientes: 10 };
const _repDatosActuales = { ventasPeriodo: [], categorias: [], formasPago: [], cierres: [], clientes: [] };

function _limitarReporte(lista, key) {
  const lim = _repLimites[key];
  return (lim === "all") ? lista : lista.slice(0, Number(lim));
}

/** Llamado por el <select> de cada reporte cuando el usuario elige 10 / 20 / todos. */
function cambiarLimiteReporte(key, valor) {
  _repLimites[key] = valor;
  switch (key) {
    case "ventasPeriodo": renderReporteVentasPeriodo(); break;
    case "productos": renderReporteProductos(); break;
    case "categorias": renderReporteCategorias(); break;
    case "formasPago": renderReporteFormasPago(); break;
    case "cierres": renderReporteCierres(); break;
    case "clientes": renderReporteClientes(); break;
  }
}

/* ---- Reporte 1: Ventas por período ---- */
async function cargarReporteVentasPeriodo(datosPrevios) {
  const tbody = document.getElementById("repVentasPeriodoTabla");
  const resumenWrap = document.getElementById("repVentasPeriodoResumen");
  const cacheKey = "ventasPeriodo" + obtenerRangoReportes();
  const cached = datosPrevios ? null : _getCacheReporte(cacheKey);
  if (cached) { _aplicarReporteVentas(cached, tbody, resumenWrap); return; }

  try {
    const data = datosPrevios || await _leerRespuestaJSON(await fetchAPI(API_URL + "?action=reporteVentasPeriodo" + obtenerRangoReportes(), {}, _REP_OPC_FETCH));
    if (!data.success) return;
    _cacheReporte(cacheKey, data);
    _aplicarReporteVentas(data, tbody, resumenWrap);
  } catch (error) {
    console.error("Error reporte ventas:", error);
    tbody.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-3">${escapeHtml(_repMensajeError(error))}</td></tr>`;
  }
}

function _aplicarReporteVentas(data, tbody, resumenWrap) {
  sincronizarRangoReportes(data.desde, data.hasta);
  const r = data.resumen || {};
  resumenWrap.innerHTML = `
    <div class="col-6 col-md-3"><div class="card p-2 text-center"><div class="text-muted" style="font-size:11.5px;">Total POS</div><div class="money fw-bold">$${Number(r.totalPOS || 0).toLocaleString("es-AR")}</div></div></div>
    <div class="col-6 col-md-3"><div class="card p-2 text-center"><div class="text-muted" style="font-size:11.5px;">Total Pedidos</div><div class="money fw-bold">$${Number(r.totalPedidos || 0).toLocaleString("es-AR")}</div></div></div>
    <div class="col-6 col-md-3"><div class="card p-2 text-center"><div class="text-muted" style="font-size:11.5px;">Total general</div><div class="money fw-bold">$${Number(r.totalGeneral || 0).toLocaleString("es-AR")}</div></div></div>
    <div class="col-6 col-md-3"><div class="card p-2 text-center"><div class="text-muted" style="font-size:11.5px;">Ticket promedio</div><div class="money fw-bold">$${Number(r.ticketPromedio || 0).toLocaleString("es-AR")}</div></div></div>`;

  // Más reciente arriba, como en el resto de las tablas de reportes
  const diasOrdenados = data.dias ? [...data.dias].sort((a,b) => String(b.fecha).localeCompare(String(a.fecha))) : [];
  _repDatosActuales.ventasPeriodo = diasOrdenados;
  renderReporteVentasPeriodo();
}

/** Renderiza la tabla de ventas por período respetando el límite elegido (10 / 20 / todos). */
function renderReporteVentasPeriodo() {
  const tbody = document.getElementById("repVentasPeriodoTabla");
  if (!tbody) return;
  const diasOrdenados = _repDatosActuales.ventasPeriodo;

  if (!diasOrdenados || diasOrdenados.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-3">Sin ventas para el rango elegido</td></tr>`;
    return;
  }

  const lista = _limitarReporte(diasOrdenados, "ventasPeriodo");
  tbody.innerHTML = lista.map(d => {
    const fecha = d.fecha ? new Date(d.fecha + "T12:00:00").toLocaleDateString("es-AR") : "—";
    return `
      <tr>
        <td>${escapeHtml(fecha)}</td>
        <td class="money">$${Number(d.pos || 0).toLocaleString("es-AR")}</td>
        <td class="money">$${Number(d.pedidos || 0).toLocaleString("es-AR")}</td>
        <td class="money">$${Number(d.total || 0).toLocaleString("es-AR")}</td>
      </tr>`;
  }).join("");
}

/* ---- Reporte 2: Productos más vendidos ----
   Carga: una sola llamada (el backend ya trae período anterior, stock y
   categoría, y cachea el resultado). Mientras carga se muestra un
   esqueleto; si el usuario cambia de rango a mitad de una carga, la
   respuesta vieja se descarta (contador _repProductosGen). Buscador,
   categoría, orden y Top N filtran en memoria, sin volver a pedir nada. */
let _repProductosDatosActuales = []; // ranking completo del último período cargado
let _repProductosMeta = null;        // { desde, hasta, anterior, totales, totalesAnterior, generado }
let _repProductosGen = 0;
let _repProductosSinVentas = [];     // catálogo sin ventas en el período (solo para el buscador)

/**
 * Búsqueda de Reportes: sin acentos ni mayúsculas, por PALABRAS en
 * cualquier orden ("1kg yerba" encuentra "Yerba Playadito 1kg"), y
 * también por código, código de caja, alias o categoría. Un código
 * escrito con o sin ceros adelante ("0123" / "123") también coincide.
 */
function coincideBusquedaReportes(textoNormalizado, consulta) {
  const q = normalizarBusquedaPOS(consulta);
  if (!q) return true;
  const palabras = q.split(/\s+/).filter(Boolean);
  return palabras.every(w => {
    if (textoNormalizado.includes(w)) return true;
    if (/^\d+$/.test(w)) {
      const sinCeros = w.replace(/^0+/, "");
      return sinCeros.length >= 2 && textoNormalizado.split(/\s+/).some(t => t.replace(/^0+/, "") === sinCeros);
    }
    return false;
  });
}

const _fmtNum = n => Number(n || 0).toLocaleString("es-AR");
const _fmtPlata = n => "$" + Math.round(Number(n || 0)).toLocaleString("es-AR");
const _fmtFechaCorta = s => {
  const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${Number(m[3])}/${Number(m[2])}/${m[1].slice(2)}` : String(s || "");
};

/** Variación porcentual con texto y flecha (nunca solo color) */
function _repVariacion(actual, anterior) {
  actual = Number(actual || 0); anterior = Number(anterior || 0);
  if (!anterior && !actual) return { clase: "na", texto: "—", valor: 0 };
  if (!anterior) return { clase: "nuevo", texto: "● nuevo", valor: Infinity };
  const pct = (actual - anterior) / anterior * 100;
  if (Math.abs(pct) < 0.5) return { clase: "flat", texto: "= 0%", valor: 0 };
  const r = Math.abs(pct) >= 10 ? Math.round(pct) : Math.round(pct * 10) / 10;
  return pct > 0
    ? { clase: "up", texto: `▲ ${r.toLocaleString("es-AR")}%`, valor: pct }
    : { clase: "down", texto: `▼ ${Math.abs(r).toLocaleString("es-AR")}%`, valor: pct };
}

function _repProductosEsqueleto() {
  const tbody = document.getElementById("repProductosTabla");
  if (!tbody) return;
  const fila = w => `<span class="rep-skel" style="width:${w}px;"></span>`;
  tbody.innerHTML = Array.from({ length: 6 }, (_, i) => `
    <tr>
      <td>${fila(14)}</td>
      <td>${fila(160 - i * 12)}<br>${fila(80)}</td>
      <td class="num">${fila(90)}</td>
      <td class="num">${fila(60)}</td>
      <td class="num d-none d-md-table-cell">${fila(34)}</td>
      <td class="num">${fila(44)}</td>
      <td class="num d-none d-md-table-cell">${fila(26)}</td>
    </tr>`).join("");
  const kpis = document.getElementById("repProductosKpis");
  if (kpis && !_repProductosMeta) {
    kpis.innerHTML = Array.from({ length: 4 }, () =>
      `<div class="rep-prod-kpi"><div class="lbl">${fila(70)}</div><div class="val">${fila(90)}</div></div>`).join("");
  }
}

/**
 * Pide reporteProductosVendidos para el período elegido, UNA sola vez:
 * Productos y Compras usan la misma respuesta (caché de sesión de 90 s +
 * la promesa en curso, así dos pestañas que la piden juntas no duplican
 * el pedido). forzar=true saltea ambas cachés (también la del servidor).
 */
const _repProdPedidosEnCurso = {};
async function _obtenerReporteProductosVendidos(forzar) {
  const rango = obtenerRangoReportes();
  const clave = "reporteProductosVendidos" + rango;
  if (!forzar) {
    const enCache = _getCacheReporte(clave);
    if (enCache) return enCache;
    if (_repProdPedidosEnCurso[clave]) return _repProdPedidosEnCurso[clave];
  }
  const promesa = (async () => {
    const response = await fetchAPI(API_URL + "?action=reporteProductosVendidos" + rango + (forzar ? "&sinCache=1&_=" + Date.now() : ""), {}, _REP_OPC_FETCH);
    const data = await _leerRespuestaJSON(response);
    if (data && data.success) {
      data._rango = "&desde=" + encodeURIComponent(data.desde) + "&hasta=" + encodeURIComponent(data.hasta);
      _cacheReporte(clave, data);
      _cacheReporte("reporteProductosVendidos" + data._rango, data);
    }
    return data;
  })();
  _repProdPedidosEnCurso[clave] = promesa;
  try { return await promesa; }
  finally { delete _repProdPedidosEnCurso[clave]; }
}

async function cargarReporteProductos(forzar) {
  const rango = obtenerRangoReportes();
  const estado = document.getElementById("repProductosEstado");
  const gen = ++_repProductosGen;

  if (!forzar) {
    const enCache = _getCacheReporte("reporteProductosVendidos" + rango);
    if (enCache) { _aplicar_cargarReporteProductos(enCache); return; }
  }

  // Si ya hay datos de este mismo rango, se dejan a la vista mientras se
  // actualiza; si es un rango nuevo, esqueleto.
  const mismoRango = _repProductosMeta && _repProductosMeta._rango === rango;
  if (!mismoRango) _repProductosEsqueleto();
  if (estado) estado.textContent = "Cargando…";

  try {
    const data = await _obtenerReporteProductosVendidos(forzar);
    if (gen !== _repProductosGen) return; // llegó tarde: ya se pidió otro rango
    if (!data || !data.success) {
      _repProductosError((data && data.message) || "El servidor no devolvió el reporte");
      return;
    }
    _aplicar_cargarReporteProductos(data);
  } catch (error) {
    if (gen !== _repProductosGen) return;
    console.error("Error al cargar reporte de productos vendidos:", error);
    _repProductosError(_repMensajeError(error));
  }
}

function _repProductosError(mensaje) {
  if (typeof _repTabRangoCargado !== "undefined") _repTabRangoCargado.productos = null; // reintenta al reabrir la pestaña
  const tbody = document.getElementById("repProductosTabla");
  const estado = document.getElementById("repProductosEstado");
  if (estado) estado.textContent = "";
  if (!tbody) return;
  if (_repProductosDatosActuales.length && _repProductosMeta) {
    // Se mantienen los datos anteriores en pantalla, solo se avisa
    toast(mensaje + " — se siguen mostrando los datos anteriores", "error");
    return;
  }
  tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4">
    <div class="text-muted mb-2">⚠️ ${escapeHtml(mensaje)}</div>
    <button class="btn btn-outline-primary btn-sm" onclick="cargarReporteProductos(true)">Reintentar</button>
  </td></tr>`;
}

/** Aplica los datos (del backend o de la caché de sesión) y arma los filtros */
function _aplicar_cargarReporteProductos(data) {
  sincronizarRangoReportes(data.desde, data.hasta);
  _repProductosDatosActuales = (data.productos || []).map(p => ({
    ...p,
    CODIGO: String(p.CODIGO ?? ""),
    CATEGORIA: p.CATEGORIA || "Sin categoría",
    VENDIDOS: Number(p.VENDIDOS || 0),
    INGRESOS: Number(p.INGRESOS || 0),
    _busqueda: normalizarBusquedaPOS([p.PRODUCTO, p.CODIGO, p.CATEGORIA, p.ALIAS, p.CODIGO_CAJA].join(" "))
  }));

  // Productos del catálogo que NO se vendieron en el período: no van en el
  // ranking, pero el buscador los encuentra igual (antes "no aparecían").
  // El backend nuevo los manda en data.sinVentas; con uno viejo se usa el
  // catálogo que ya esté cargado en el panel.
  const vendidos = new Set(_repProductosDatosActuales.map(p => p.CODIGO));
  let sinVentas = [];
  if (Array.isArray(data.sinVentas)) {
    sinVentas = data.sinVentas.map(r => ({ CODIGO: String(r[0] ?? ""), PRODUCTO: r[1] || "", CATEGORIA: r[2] || "Sin categoría", STOCK: r[3], ALIAS: r[4] || "", CODIGO_CAJA: r[5] || "" }));
  } else if (typeof productosAdminGlobal !== "undefined" && Array.isArray(productosAdminGlobal)) {
    sinVentas = productosAdminGlobal.map(p => ({ CODIGO: String(p.CODIGO ?? ""), PRODUCTO: p.PRODUCTO || "", CATEGORIA: p.CATEGORIA || "Sin categoría", STOCK: p.STOCK, ALIAS: p.ALIAS || "", CODIGO_CAJA: p.CODIGO_CAJA || "" }));
  }
  _repProductosSinVentas = sinVentas
    .filter(p => p.CODIGO && !vendidos.has(p.CODIGO))
    .map(p => ({
      ...p, VENDIDOS: 0, INGRESOS: 0, VENDIDOS_ANTERIOR: 0, INGRESOS_ANTERIOR: 0, OPERACIONES: 0,
      EN_CATALOGO: true, SIN_VENTAS: true, DIAS: {},
      _busqueda: normalizarBusquedaPOS([p.PRODUCTO, p.CODIGO, p.CATEGORIA, p.ALIAS, p.CODIGO_CAJA].join(" "))
    }));

  // Totales: los manda el backend nuevo; con uno viejo se calculan acá
  const totales = data.totales || {
    unidades: _repProductosDatosActuales.reduce((s, p) => s + p.VENDIDOS, 0),
    ingresos: _repProductosDatosActuales.reduce((s, p) => s + p.INGRESOS, 0),
    productos: _repProductosDatosActuales.length
  };
  _repProductosMeta = {
    _rango: data._rango, desde: data.desde, hasta: data.hasta,
    anterior: data.anterior || null, totales, totalesAnterior: data.totalesAnterior || null,
    generado: data.generado || null, conComparacion: !!data.anterior,
    porDia: data.porDia || null
  };
  if (_repProductoSel && !_repProductosDatosActuales.some(p => p.CODIGO === _repProductoSel)) _repProductoSel = null;

  // Categorías presentes en el período (respeta la elegida si sigue existiendo)
  const selCat = document.getElementById("repProductosCategoria");
  if (selCat) {
    const actual = selCat.value;
    const cats = [...new Set(_repProductosDatosActuales.map(p => p.CATEGORIA))].sort((a, b) => a.localeCompare(b, "es"));
    selCat.innerHTML = `<option value="">Todas las categorías</option>` +
      cats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
    selCat.value = cats.includes(actual) ? actual : "";
  }

  const periodo = document.getElementById("repProductosPeriodo");
  if (periodo) {
    periodo.textContent = `${_fmtFechaCorta(data.desde)} al ${_fmtFechaCorta(data.hasta)}` +
      (data.anterior ? ` · comparado con ${_fmtFechaCorta(data.anterior.desde)} al ${_fmtFechaCorta(data.anterior.hasta)}` : "");
  }
  const estado = document.getElementById("repProductosEstado");
  if (estado) {
    const hora = data.generado ? new Date(data.generado) : new Date();
    estado.textContent = "Datos de las " + hora.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" });
  }

  _renderKpisReporteProductos();
  renderReporteProductos();
}

function _renderKpisReporteProductos() {
  const cont = document.getElementById("repProductosKpis");
  const m = _repProductosMeta;
  if (!cont || !m) return;
  const t = m.totales, a = m.totalesAnterior;
  const comp = (act, ant) => {
    if (!a) return "";
    const v = _repVariacion(act, ant);
    return `<div class="sub"><span class="rep-tend ${v.clase}">${v.texto}</span> vs. período anterior</div>`;
  };
  // Concentración: cuánto de los ingresos explican los 10 primeros
  const top10 = [..._repProductosDatosActuales].sort((x, y) => y.INGRESOS - x.INGRESOS).slice(0, 10)
    .reduce((s, p) => s + p.INGRESOS, 0);
  const pctTop10 = t.ingresos ? Math.round(top10 / t.ingresos * 100) : 0;

  cont.innerHTML = `
    <div class="rep-prod-kpi"><div class="lbl">Unidades vendidas</div><div class="val">${_fmtNum(t.unidades)}</div>${comp(t.unidades, a && a.unidades)}</div>
    <div class="rep-prod-kpi"><div class="lbl">Ingresos por productos</div><div class="val">${_fmtPlata(t.ingresos)}</div>${comp(t.ingresos, a && a.ingresos)}</div>
    <div class="rep-prod-kpi"><div class="lbl">Productos distintos vendidos</div><div class="val">${_fmtNum(t.productos)}</div>${comp(t.productos, a && a.productos)}</div>
    <div class="rep-prod-kpi"><div class="lbl">Peso del Top 10</div><div class="val">${pctTop10}%</div><div class="sub">de los ingresos del período</div></div>`;
}

/** Lista filtrada y ordenada según buscador, categoría y orden (sin el Top N) */
function _repProductosFiltrados() {
  const texto = normalizarBusquedaPOS((document.getElementById("repProductosBuscador") || {}).value || "");
  const categoria = (document.getElementById("repProductosCategoria") || {}).value || "";
  const orden = (document.getElementById("repProductosOrden") || {}).value || "unidades";

  const consulta = (document.getElementById("repProductosBuscador") || {}).value || "";
  const base = texto ? _repProductosDatosActuales.concat(_repProductosSinVentas) : _repProductosDatosActuales;
  let lista = base.filter(p =>
    (!categoria || p.CATEGORIA === categoria) && (!texto || coincideBusquedaReportes(p._busqueda, consulta)));

  const varUnid = p => _repVariacion(p.VENDIDOS, p.VENDIDOS_ANTERIOR).valor;
  const stockNum = p => (p.STOCK === null || p.STOCK === undefined || p.STOCK === "") ? Infinity : Number(p.STOCK);
  const comparadores = {
    unidades: (x, y) => (y.VENDIDOS - x.VENDIDOS) || (y.INGRESOS - x.INGRESOS),
    ingresos: (x, y) => (y.INGRESOS - x.INGRESOS) || (y.VENDIDOS - x.VENDIDOS),
    suba: (x, y) => (varUnid(y) - varUnid(x)) || (y.VENDIDOS - x.VENDIDOS),
    baja: (x, y) => (varUnid(x) - varUnid(y)) || (y.VENDIDOS - x.VENDIDOS),
    stock: (x, y) => (stockNum(x) - stockNum(y)) || (y.VENDIDOS - x.VENDIDOS)
  };
  return lista.sort(comparadores[orden] || comparadores.unidades);
}

/** Renderiza la tabla (los parámetros viejos se ignoran: todo sale de los controles) */
function renderReporteProductos() {
  const tbody = document.getElementById("repProductosTabla");
  const pie = document.getElementById("repProductosPie");
  if (!tbody) return;

  const buscando = !!normalizarBusquedaPOS((document.getElementById("repProductosBuscador") || {}).value || "");
  if (!_repProductosDatosActuales.length && !buscando) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">Sin ventas de productos en el período elegido</td></tr>`;
    if (pie) pie.textContent = "";
    const g = document.getElementById("repProdGraficos");
    if (g) g.style.display = "none";
    return;
  }

  const filtrados = _repProductosFiltrados();
  if (!filtrados.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">${buscando
      ? "Ningún producto del catálogo coincide con la búsqueda" + ((document.getElementById("repProductosCategoria") || {}).value ? " en esa categoría" : "")
      : "Ningún producto coincide con los filtros"}</td></tr>`;
    if (pie) pie.textContent = "";
    _repDibujarGraficos(filtrados);
    return;
  }

  // Buscando se muestran TODAS las coincidencias (con Top 10 quedaban productos afuera)
  const lista = buscando ? filtrados : _limitarReporte(filtrados, "productos");
  const maxUnid = Math.max(...lista.map(p => p.VENDIDOS), 1);
  const totalIngresos = (_repProductosMeta && _repProductosMeta.totales.ingresos) || 0;
  const conComparacion = _repProductosMeta && _repProductosMeta.conComparacion;

  tbody.innerHTML = lista.map((p, i) => {
    const anchoBarra = Math.max(2, Math.round(p.VENDIDOS / maxUnid * 100));
    const pct = totalIngresos ? (p.INGRESOS / totalIngresos * 100) : 0;
    const v = conComparacion ? _repVariacion(p.VENDIDOS, p.VENDIDOS_ANTERIOR) : { clase: "na", texto: "—" };
    const tituloVar = conComparacion ? `Período anterior: ${_fmtNum(p.VENDIDOS_ANTERIOR)} u.` : "";

    let stockHtml = `<span class="text-muted">—</span>`;
    if (p.EN_CATALOGO === false) stockHtml = `<span class="text-muted" title="El producto ya no está en el catálogo">eliminado</span>`;
    else if (p.STOCK !== null && p.STOCK !== undefined && p.STOCK !== "") {
      const st = Number(p.STOCK);
      // Stock bajo: no alcanza para repetir las ventas de este período
      stockHtml = st < p.VENDIDOS
        ? `<span class="rep-stock-bajo" title="Quedan menos unidades que las vendidas en el período">⚠ ${_fmtNum(st)}</span>`
        : _fmtNum(st);
    }

    if (p.SIN_VENTAS) {
      return `
      <tr class="rep-sin-ventas">
        <td class="rep-prod-rank">—</td>
        <td>
          <div class="rep-prod-nombre">${escapeHtml(p.PRODUCTO || "(sin nombre)")}</div>
          <div class="rep-prod-meta"><span class="mono">${escapeHtml(p.CODIGO)}</span> · ${escapeHtml(p.CATEGORIA)}</div>
        </td>
        <td class="num" colspan="4"><span class="text-muted" style="font-size:12px;">Sin ventas en este período</span></td>
        <td class="num d-none d-md-table-cell">${stockHtml}</td>
      </tr>`;
    }
    const sel = p.CODIGO === _repProductoSel;
    return `
      <tr class="rep-fila${sel ? " rep-sel" : ""}" onclick="repSeleccionarProducto('${escapeJsAttr(p.CODIGO)}')" title="Ver la evolución de este producto">
        <td class="rep-prod-rank">${i + 1}</td>
        <td>
          <div class="rep-prod-nombre">${escapeHtml(p.PRODUCTO || "(sin nombre)")}</div>
          <div class="rep-prod-meta"><span class="mono">${escapeHtml(p.CODIGO)}</span> · ${escapeHtml(p.CATEGORIA)}${p.OPERACIONES ? ` · en ${_fmtNum(p.OPERACIONES)} venta${p.OPERACIONES === 1 ? "" : "s"}` : ""}</div>
        </td>
        <td class="num">
          <div class="rep-prod-barra" title="${_fmtNum(p.VENDIDOS)} unidades">
            <div class="track"><div class="fill" style="width:${anchoBarra}%;"></div></div>
            <span>${_fmtNum(p.VENDIDOS)}</span>
          </div>
        </td>
        <td class="num">${_fmtPlata(p.INGRESOS)}</td>
        <td class="num d-none d-md-table-cell">${pct >= 0.1 ? pct.toLocaleString("es-AR", { maximumFractionDigits: 1 }) + "%" : "<0,1%"}</td>
        <td class="num"><span class="rep-tend ${v.clase}" title="${escapeHtml(tituloVar)}">${v.texto}</span></td>
        <td class="num d-none d-md-table-cell">${stockHtml}</td>
      </tr>`;
  }).join("");

  if (pie) {
    const conVentas = filtrados.filter(p => !p.SIN_VENTAS).length;
    const partes = buscando
      ? [`${_fmtNum(filtrados.length)} coincidencia${filtrados.length === 1 ? "" : "s"}: ${_fmtNum(conVentas)} con ventas y ${_fmtNum(filtrados.length - conVentas)} sin ventas en el período`]
      : [`Mostrando ${_fmtNum(lista.length)} de ${_fmtNum(filtrados.length)} productos vendidos`];
    if (!buscando && filtrados.length !== _repProductosDatosActuales.length) partes.push(`(${_fmtNum(_repProductosDatosActuales.length)} en total)`);
    if (!conComparacion) partes.push("· la comparación con el período anterior aparece al actualizar el Code.gs");
    pie.textContent = partes.join(" ");
  }
  _repDibujarGraficos(filtrados.filter(p => !p.SIN_VENTAS));
}


/* ---- Gráficos del reporte de productos (Chart.js, ya cargado en index.html) ----
   Siguen los mismos filtros que la tabla. Un solo tono (azul) para el
   período actual y gris para el anterior; la categoría o el producto
   elegido se resalta y el resto queda más claro. */
const _REP_AZUL = "#2563eb";
const _REP_AZUL_CLARO = "rgba(37,99,235,.28)";
const _REP_GRIS = "#b8c1cf";
const _REP_GRILLA = "#eef1f6";
const _REP_TEXTO = "#6b7585";
let _repCharts = {};
let _repProductoSel = null; // código del producto elegido para el gráfico diario

const _fmtCompacto = n => {
  const v = Math.abs(Number(n || 0));
  if (v >= 1e6) return (n / 1e6).toLocaleString("es-AR", { maximumFractionDigits: 1 }) + " M";
  if (v >= 1e3) return (n / 1e3).toLocaleString("es-AR", { maximumFractionDigits: 1 }) + " mil";
  return Number(n || 0).toLocaleString("es-AR");
};
const _recortar = (t, n) => { t = String(t || ""); return t.length > n ? t.slice(0, n - 1) + "…" : t; };

function _repOpcionesBase() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 250 },
    plugins: {
      legend: { display: false },
      tooltip: { backgroundColor: "#0b1633", padding: 10, cornerRadius: 8, titleFont: { weight: "600" }, displayColors: true, boxPadding: 4 }
    }
  };
}

function _repDibujarGraficos(filtrados) {
  const cont = document.getElementById("repProdGraficos");
  if (!cont) return;
  if (typeof Chart === "undefined") { cont.style.display = "none"; return; } // sin internet la 1ª vez no carga la librería
  cont.style.display = "";
  _repDibujarGraficoTop(filtrados);
  _repDibujarGraficoCategorias();
  _repDibujarGraficoDias();
}

/* -- 1. Top 10: barras horizontales, período actual vs. anterior -- */
function _repDibujarGraficoTop(filtrados) {
  const canvas = document.getElementById("repGrafTop");
  if (!canvas) return;
  if (_repCharts.top) _repCharts.top.destroy();

  const orden = (document.getElementById("repProductosOrden") || {}).value || "unidades";
  const porIngresos = orden === "ingresos";
  const valor = p => porIngresos ? p.INGRESOS : p.VENDIDOS;
  const valorAnt = p => porIngresos ? Number(p.INGRESOS_ANTERIOR || 0) : Number(p.VENDIDOS_ANTERIOR || 0);
  const top = [...filtrados].sort((a, b) => valor(b) - valor(a)).slice(0, 10);
  const conComparacion = _repProductosMeta && _repProductosMeta.conComparacion;
  const fmt = porIngresos ? _fmtPlata : (n => _fmtNum(n) + " u.");

  const titulo = document.getElementById("repGrafTopTitulo");
  if (titulo) titulo.textContent = `Top ${top.length} por ${porIngresos ? "ingresos" : "unidades"}` + (conComparacion ? " — este período vs. el anterior" : "");

  const resaltar = p => !_repProductoSel || p.CODIGO === _repProductoSel;
  const datasets = [{
    label: "Este período",
    data: top.map(valor),
    backgroundColor: top.map(p => resaltar(p) ? _REP_AZUL : _REP_AZUL_CLARO),
    borderRadius: 4, borderSkipped: "start", barPercentage: .85, categoryPercentage: .75
  }];
  if (conComparacion) datasets.push({
    label: "Período anterior",
    data: top.map(valorAnt),
    backgroundColor: _REP_GRIS,
    borderRadius: 4, borderSkipped: "start", barPercentage: .85, categoryPercentage: .75
  });

  const op = _repOpcionesBase();
  op.indexAxis = "y";
  op.plugins.legend = { display: conComparacion, position: "bottom", labels: {
    boxWidth: 10, boxHeight: 10, font: { size: 11 }, color: _REP_TEXTO,
    // La leyenda usa siempre el azul pleno (las barras pueden estar aclaradas por la selección)
    generateLabels: ch => ch.data.datasets.map((ds, i) => ({
      text: ds.label, datasetIndex: i, hidden: !ch.isDatasetVisible(i),
      fillStyle: i === 0 ? _REP_AZUL : _REP_GRIS, strokeStyle: i === 0 ? _REP_AZUL : _REP_GRIS, lineWidth: 0
    }))
  } };
  op.plugins.tooltip.callbacks = {
    title: items => top[items[0].dataIndex].PRODUCTO,
    label: ctx => ` ${ctx.dataset.label}: ${fmt(ctx.raw)}`
  };
  op.scales = {
    x: { beginAtZero: true, grid: { color: _REP_GRILLA }, border: { display: false }, ticks: { color: _REP_TEXTO, font: { size: 11 }, callback: v => porIngresos ? "$" + _fmtCompacto(v) : _fmtCompacto(v) } },
    y: { grid: { display: false }, border: { display: false }, ticks: { color: "#334155", font: { size: 11.5 }, callback: (v, i) => _recortar(top[i] && top[i].PRODUCTO, 26) } }
  };
  op.onClick = (evt, elementos) => { if (elementos.length) repSeleccionarProducto(top[elementos[0].index].CODIGO); };
  op.onHover = (evt, elementos) => { evt.native.target.style.cursor = elementos.length ? "pointer" : "default"; };

  _repCharts.top = new Chart(canvas, { type: "bar", data: { labels: top.map(p => p.CODIGO), datasets }, options: op });
}

/* -- 2. Ingresos por categoría: barras horizontales ordenadas (8 + "Otras") -- */
function _repDibujarGraficoCategorias() {
  const canvas = document.getElementById("repGrafCategorias");
  if (!canvas) return;
  if (_repCharts.cat) _repCharts.cat.destroy();

  // Respeta el buscador pero NO la categoría (si no, siempre habría una sola barra)
  const texto = normalizarBusquedaPOS((document.getElementById("repProductosBuscador") || {}).value || "");
  const catSel = (document.getElementById("repProductosCategoria") || {}).value || "";
  const sumas = {};
  _repProductosDatosActuales.forEach(p => {
    if (texto && !p._busqueda.includes(texto)) return;
    sumas[p.CATEGORIA] = (sumas[p.CATEGORIA] || 0) + p.INGRESOS;
  });
  let filas = Object.entries(sumas).sort((a, b) => b[1] - a[1]);
  if (filas.length > 9) {
    const otras = filas.slice(8).reduce((s, f) => s + f[1], 0);
    filas = filas.slice(0, 8).concat([["Otras", otras]]);
  }
  const total = filas.reduce((s, f) => s + f[1], 0);

  const op = _repOpcionesBase();
  op.indexAxis = "y";
  op.plugins.tooltip.callbacks = {
    label: ctx => ` ${_fmtPlata(ctx.raw)} (${total ? (ctx.raw / total * 100).toLocaleString("es-AR", { maximumFractionDigits: 1 }) : 0}%)`
  };
  op.scales = {
    x: { beginAtZero: true, grid: { color: _REP_GRILLA }, border: { display: false }, ticks: { color: _REP_TEXTO, font: { size: 11 }, callback: v => "$" + _fmtCompacto(v) } },
    y: { grid: { display: false }, border: { display: false }, ticks: { color: "#334155", font: { size: 11.5 }, callback: (v, i) => _recortar(filas[i] && filas[i][0], 20) } }
  };
  op.onClick = (evt, elementos) => {
    if (!elementos.length) return;
    const cat = filas[elementos[0].index][0];
    if (cat === "Otras") return;
    const sel = document.getElementById("repProductosCategoria");
    if (sel) { sel.value = sel.value === cat ? "" : cat; renderReporteProductos(); }
  };
  op.onHover = (evt, elementos) => {
    const ok = elementos.length && filas[elementos[0].index][0] !== "Otras";
    evt.native.target.style.cursor = ok ? "pointer" : "default";
  };

  _repCharts.cat = new Chart(canvas, {
    type: "bar",
    data: {
      labels: filas.map(f => f[0]),
      datasets: [{
        label: "Ingresos",
        data: filas.map(f => f[1]),
        backgroundColor: filas.map(f => (!catSel || f[0] === catSel) ? _REP_AZUL : _REP_AZUL_CLARO),
        borderRadius: 4, borderSkipped: "start", barPercentage: .8
      }]
    },
    options: op
  });
}

/* -- 3. Evolución por día (o por semana si el período es largo) -- */
function _repDibujarGraficoDias() {
  const canvas = document.getElementById("repGrafDias");
  const aviso = document.getElementById("repGrafDiasAviso");
  const chip = document.getElementById("repGrafDiasChip");
  const titulo = document.getElementById("repGrafDiasTitulo");
  if (!canvas) return;
  if (_repCharts.dias) { _repCharts.dias.destroy(); _repCharts.dias = null; }

  const meta = _repProductosMeta;
  const serieBase = meta && meta.porDia;
  if (!serieBase || !serieBase.length) {
    canvas.parentElement.style.display = "none";
    if (aviso) { aviso.style.display = "block"; aviso.textContent = "Este gráfico aparece al publicar el Code.gs nuevo (necesita las ventas día por día)."; }
    return;
  }
  canvas.parentElement.style.display = "";
  if (aviso) aviso.style.display = "none";

  const metrica = (document.getElementById("repGrafDiasMetrica") || {}).value || "unidades";
  const idx = metrica === "ingresos" ? 1 : 0;
  const catSel = (document.getElementById("repProductosCategoria") || {}).value || "";
  const prod = _repProductoSel ? _repProductosDatosActuales.find(p => p.CODIGO === _repProductoSel) : null;

  // Qué se grafica: un producto, una categoría o el total
  let valores, etiqueta;
  if (prod) {
    valores = serieBase.map(d => { const par = (prod.DIAS || {})[d.fecha]; return par ? Number(Array.isArray(par) ? par[idx] : (idx ? 0 : par)) : 0; });
    etiqueta = prod.PRODUCTO;
  } else if (catSel) {
    const enCat = _repProductosDatosActuales.filter(p => p.CATEGORIA === catSel);
    valores = serieBase.map(d => enCat.reduce((s, p) => { const par = (p.DIAS || {})[d.fecha]; return s + (par ? Number(Array.isArray(par) ? par[idx] : (idx ? 0 : par)) : 0); }, 0));
    etiqueta = catSel;
  } else {
    valores = serieBase.map(d => Number(idx ? d.ingresos : d.unidades) || 0);
    etiqueta = "Todos los productos";
  }

  if (chip) {
    if (prod) {
      chip.style.display = "inline-flex";
      chip.innerHTML = `<span>${escapeHtml(prod.PRODUCTO)}</span><button type="button" title="Ver todos" onclick="repSeleccionarProducto(null)">✕</button>`;
    } else chip.style.display = "none";
  }

  // Más de 2 meses: por semana (lunes a domingo), si no se vuelve ilegible
  let etiquetas = serieBase.map(d => d.fecha);
  let porSemana = false;
  if (serieBase.length > 62) {
    porSemana = true;
    const semanas = [];
    serieBase.forEach((d, i) => {
      const f = new Date(d.fecha + "T12:00:00");
      const lunes = new Date(f); lunes.setDate(f.getDate() - ((f.getDay() + 6) % 7));
      const clave = _fechaISOEnZona(lunes);
      const ult = semanas[semanas.length - 1];
      if (ult && ult.clave === clave) ult.valor += valores[i];
      else semanas.push({ clave, valor: valores[i] });
    });
    etiquetas = semanas.map(s => s.clave);
    valores = semanas.map(s => s.valor);
  }

  if (titulo) titulo.textContent = `${metrica === "ingresos" ? "Ingresos" : "Unidades vendidas"} por ${porSemana ? "semana" : "día"} — ${etiqueta}`;

  const fmt = metrica === "ingresos" ? _fmtPlata : (n => _fmtNum(n) + " u.");
  const fmtEje = f => _fmtFechaCorta(f).replace(/\/\d{2}$/, ""); // "3/10"
  const ctx = canvas.getContext("2d");
  const grad = ctx.createLinearGradient(0, 0, 0, 220);
  grad.addColorStop(0, "rgba(37,99,235,.18)");
  grad.addColorStop(1, "rgba(37,99,235,0)");

  const op = _repOpcionesBase();
  op.interaction = { mode: "index", intersect: false };
  op.plugins.tooltip.callbacks = {
    title: items => (porSemana ? "Semana del " : "") + _fmtFechaCorta(etiquetas[items[0].dataIndex]),
    label: c => ` ${fmt(c.raw)}`
  };
  op.scales = {
    x: { grid: { display: false }, border: { color: _REP_GRILLA }, ticks: { color: _REP_TEXTO, font: { size: 11 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 12, callback: (v, i) => fmtEje(etiquetas[i]) } },
    y: { beginAtZero: true, grid: { color: _REP_GRILLA }, border: { display: false }, ticks: { color: _REP_TEXTO, font: { size: 11 }, maxTicksLimit: 5, precision: 0, callback: v => metrica === "ingresos" ? "$" + _fmtCompacto(v) : _fmtCompacto(v) } }
  };

  _repCharts.dias = new Chart(canvas, {
    type: "line",
    data: {
      labels: etiquetas,
      datasets: [{
        label: etiqueta, data: valores,
        borderColor: _REP_AZUL, borderWidth: 2, backgroundColor: grad, fill: true,
        tension: .3, cubicInterpolationMode: "monotone", pointRadius: valores.length <= 31 ? 2.5 : 0, pointHoverRadius: 5,
        pointBackgroundColor: _REP_AZUL, pointBorderColor: "#fff", pointBorderWidth: 1.5
      }]
    },
    options: op
  });
}

/** Elige (o suelta) un producto: se resalta en la tabla y en el Top 10, y el gráfico diario muestra solo ese */
function repSeleccionarProducto(codigo) {
  _repProductoSel = (codigo === null || codigo === undefined || String(codigo) === _repProductoSel) ? null : String(codigo);
  renderReporteProductos();
  if (_repProductoSel) document.getElementById("repGrafDias")?.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

/** Llamado por el buscador (oninput) */
function filtrarReporteProductos() {
  renderReporteProductos();
}

/** Filas para exportar: TODO lo filtrado (no solo el Top N visible) */
function _repProductosFilasExport() {
  const total = (_repProductosMeta && _repProductosMeta.totales.ingresos) || 0;
  return _repProductosFiltrados().map((p, i) => ({
    "#": i + 1,
    "Código": p.CODIGO,
    "Producto": p.PRODUCTO,
    "Categoría": p.CATEGORIA,
    "Unidades": p.VENDIDOS,
    "Ingresos": Math.round(p.INGRESOS),
    "% ingresos": total ? Math.round(p.INGRESOS / total * 1000) / 10 : 0,
    "Unidades período anterior": p.VENDIDOS_ANTERIOR ?? "",
    "Variación": _repVariacion(p.VENDIDOS, p.VENDIDOS_ANTERIOR).texto.replace(/[▲▼●=]\s?/g, "").trim(),
    "Stock actual": (p.STOCK === null || p.STOCK === undefined) ? "" : p.STOCK
  }));
}

function exportarReporteProductosCSV() {
  const filas = _repProductosFilasExport();
  if (!filas.length) { toast("No hay datos para exportar", "error"); return; }
  const cols = Object.keys(filas[0]);
  const celda = v => {
    const t = typeof v === "number" ? String(v).replace(".", ",") : String(v ?? "");
    return /[";\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t;
  };
  // ";" y BOM: Excel en español lo abre directo en columnas y con acentos
  const csv = "﻿" + [cols.join(";"), ...filas.map(f => cols.map(c => celda(f[c])).join(";"))].join("\r\n");
  const m = _repProductosMeta || {};
  descargarArchivo(`Productos_mas_vendidos_${m.desde || ""}_a_${m.hasta || ""}.csv`, csv, "text/csv;charset=utf-8");
}

function exportarReporteProductosPDF() {
  try {
    const filas = _repProductosFilasExport();
    if (!filas.length) { toast("No hay datos para exportar", "error"); return; }
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    const m = _repProductosMeta || {};
    const t = m.totales || {};
    const nombreLocal = (obtenerConfigNegocio().nombre || "Reporte").toString();

    doc.setFontSize(14);
    doc.text(`${nombreLocal} — Productos más vendidos`, 30, 30);
    doc.setFontSize(10);
    doc.setTextColor(110, 110, 110);
    doc.text(`Período: ${_fmtFechaCorta(m.desde)} al ${_fmtFechaCorta(m.hasta)}  ·  ${_fmtNum(t.unidades)} unidades  ·  ${_fmtPlata(t.ingresos)}  ·  Generado: ${new Date().toLocaleString("es-AR")}`, 30, 46);

    const cols = ["#", "Código", "Producto", "Categoría", "Unidades", "Ingresos", "% ingresos", "Variación", "Stock actual"];
    doc.autoTable({
      head: [cols],
      body: filas.map(f => cols.map(c =>
        c === "Ingresos" ? _fmtPlata(f[c]) :
        c === "% ingresos" ? String(f[c]).replace(".", ",") + "%" :
        c === "Unidades" ? _fmtNum(f[c]) : String(f[c] ?? ""))),
      startY: 58,
      theme: "grid",
      styles: { fontSize: 8.5, cellPadding: 4 },
      headStyles: { fillColor: [18, 32, 71], textColor: [255, 255, 255] },
      columnStyles: { 0: { halign: "right" }, 4: { halign: "right" }, 5: { halign: "right" }, 6: { halign: "right" }, 7: { halign: "right" }, 8: { halign: "right" } }
    });
    doc.save(`Productos_mas_vendidos_${m.desde || ""}_a_${m.hasta || ""}.pdf`);
  } catch (error) {
    console.error("Error al exportar productos a PDF:", error);
    toast("No se pudo generar el PDF", "error");
  }
}

/* ---- Reporte 3: Ventas por categoría ---- */
async function cargarReporteCategorias(datosPrevios) {
  const _ck_repCategorias = "reporteVentasPorCategoria" + obtenerRangoReportes();
  const _cd_repCategorias = null; // (estos reportes no se cachean por separado: vienen todos juntos de reportesVentas)
  if (_cd_repCategorias) { _aplicar_cargarReporteCategorias(_cd_repCategorias); return; }
  const tbody = document.getElementById("repCategoriasTabla");

  try {
    const data = datosPrevios || await _leerRespuestaJSON(await fetchAPI(API_URL + "?action=reporteVentasPorCategoria" + obtenerRangoReportes(), {}, _REP_OPC_FETCH));
    if (!data.success) return;

    sincronizarRangoReportes(data.desde, data.hasta);
    _repDatosActuales.categorias = data.categorias || [];
    renderReporteCategorias();

  } catch (error) {
    console.error("Error al cargar reporte de ventas por categoría:", error);
    tbody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">${escapeHtml(_repMensajeError(error))}</td></tr>`;
  }
}

function renderReporteCategorias() {
  const tbody = document.getElementById("repCategoriasTabla");
  if (!tbody) return;
  const categorias = _repDatosActuales.categorias;

  if (!categorias || categorias.length === 0) {
    tbody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">Sin ventas para el rango elegido</td></tr>`;
    return;
  }

  const lista = _limitarReporte(categorias, "categorias");
  tbody.innerHTML = lista.map(c => `
    <tr>
      <td>${escapeHtml(c.categoria)}</td>
      <td class="money">${Number(c.cantidad || 0).toLocaleString("es-AR")}</td>
      <td class="money">$${Number(c.ingresos || 0).toLocaleString("es-AR")}</td>
    </tr>`).join("");
}

/* ---- Reporte 4: Formas de pago ---- */
async function cargarReporteFormasPago(datosPrevios) {
  const tbody = document.getElementById("repFormasPagoTabla");

  try {
    const data = datosPrevios || await _leerRespuestaJSON(await fetchAPI(API_URL + "?action=reporteFormasPago" + obtenerRangoReportes(), {}, _REP_OPC_FETCH));
    if (!data.success) return;

    sincronizarRangoReportes(data.desde, data.hasta);
    _repDatosActuales.formasPago = data.formas || [];
    renderReporteFormasPago();

  } catch (error) {
    console.error("Error al cargar reporte de formas de pago:", error);
    tbody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">${escapeHtml(_repMensajeError(error))}</td></tr>`;
  }
}

function renderReporteFormasPago() {
  const tbody = document.getElementById("repFormasPagoTabla");
  if (!tbody) return;
  const formas = _repDatosActuales.formasPago;

  if (!formas || formas.length === 0) {
    tbody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">Sin ventas para el rango elegido</td></tr>`;
    return;
  }

  const lista = _limitarReporte(formas, "formasPago");
  tbody.innerHTML = lista.map(f => `
    <tr>
      <td>${escapeHtml(f.forma)}</td>
      <td class="money">${Number(f.cantidad || 0).toLocaleString("es-AR")}</td>
      <td class="money">$${Number(f.total || 0).toLocaleString("es-AR")}</td>
    </tr>`).join("");
}

/* ---- Reporte 5: Historial de cierres de caja ---- */
async function cargarReporteCierres(datosPrevios) {
  const _ck_repCierres = "reporteCierres" + obtenerRangoReportes();
  const _cd_repCierres = null; // (estos reportes no se cachean por separado: vienen todos juntos de reportesVentas)
  if (_cd_repCierres) { _aplicar_cargarReporteCierres(_cd_repCierres); return; }
  const tbody = document.getElementById("repCierresTabla");

  try {
    const data = datosPrevios || await _leerRespuestaJSON(await fetchAPI(API_URL + "?action=reporteCierresCaja" + obtenerRangoReportes(), {}, _REP_OPC_FETCH));
    if (!data.success) return;

    sincronizarRangoReportes(data.desde, data.hasta);
    _repDatosActuales.cierres = data.cierres || [];
    renderReporteCierres();

  } catch (error) {
    console.error("Error al cargar reporte de cierres de caja:", error);
    tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">${escapeHtml(_repMensajeError(error))}</td></tr>`;
  }
}

function renderReporteCierres() {
  const tbody = document.getElementById("repCierresTabla");
  if (!tbody) return;
  const cierres = _repDatosActuales.cierres;

  if (!cierres || cierres.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Sin cierres para el rango elegido</td></tr>`;
    return;
  }

  const lista = _limitarReporte(cierres, "cierres");
  tbody.innerHTML = lista.map(c => {
    const fecha = c.FECHA ? new Date(c.FECHA).toLocaleDateString("es-AR") : "—";
    const totalDif = Number(c.TOTAL_DIFERENCIA || 0);
    const claseDif = Math.abs(totalDif) < 1 ? "cc-dif-ok" : (totalDif > 0 ? "cc-dif-sobra" : "cc-dif-falta");
    const signo = totalDif > 0 ? "+" : "";
    return `
    <tr>
      <td>${escapeHtml(fecha)}</td>
      <td class="money">$${Number(c.TOTAL_ESPERADO || 0).toLocaleString("es-AR")}</td>
      <td class="money">$${Number(c.TOTAL_CONTADO || 0).toLocaleString("es-AR")}</td>
      <td class="money ${claseDif}">${signo}$${Math.round(totalDif).toLocaleString("es-AR")}</td>
      <td>${escapeHtml(c.VENDEDOR || "—")}</td>
    </tr>`;
  }).join("");
}

/* ---- Reporte 6: Clientes que más compran ---- */
async function cargarReporteClientes(datosPrevios) {
  const _ck_repClientes = "reporteClientes" + obtenerRangoReportes();
  const _cd_repClientes = null; // (estos reportes no se cachean por separado: vienen todos juntos de reportesVentas)
  if (_cd_repClientes) { _aplicar_cargarReporteClientes(_cd_repClientes); return; }
  const tbody = document.getElementById("repClientesTabla");

  try {
    const data = datosPrevios || await _leerRespuestaJSON(await fetchAPI(API_URL + "?action=reporteClientes" + obtenerRangoReportes(), {}, _REP_OPC_FETCH));
    if (!data.success) return;

    sincronizarRangoReportes(data.desde, data.hasta);
    _repDatosActuales.clientes = data.clientes || [];
    renderReporteClientes();

  } catch (error) {
    console.error("Error al cargar reporte de clientes:", error);
    tbody.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-3">${escapeHtml(_repMensajeError(error))}</td></tr>`;
  }
}

function renderReporteClientes() {
  const tbody = document.getElementById("repClientesTabla");
  if (!tbody) return;
  const clientes = _repDatosActuales.clientes;

  if (!clientes || clientes.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-3">Sin pedidos para el rango elegido</td></tr>`;
    return;
  }

  const lista = _limitarReporte(clientes, "clientes");
  tbody.innerHTML = lista.map(c => `
    <tr>
      <td>${escapeHtml(c.CLIENTE)}</td>
      <td>${escapeHtml(c.EMPRESA || "—")}</td>
      <td class="money">${Number(c.PEDIDOS || 0).toLocaleString("es-AR")}</td>
      <td class="money">$${Number(c.TOTAL || 0).toLocaleString("es-AR")}</td>
    </tr>`).join("");
}

/* ---- Exportar cualquiera de los 6 reportes a PDF ---- */
/**
 * Toma la tarjeta del reporte (por su id), lee la tabla que tiene
 * adentro, y genera un PDF con jsPDF + autoTable. Funciona igual
 * para los 6 reportes porque todos son <table> dentro de una .card.
 */
function exportarReportePDF(cardId, tituloReporte) {
  try {
    const card = document.getElementById(cardId);
    if (!card) return;

    const tabla = card.querySelector("table");
    if (!tabla) {
      toast("Este reporte no tiene datos para exportar", "error");
      return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

    const nombreLocal = (obtenerConfigNegocio().nombre || "Reporte").toString();
    const desde = document.getElementById("repDesde").value || "—";
    const hasta = document.getElementById("repHasta").value || "—";

    doc.setFontSize(14);
    doc.text(`${nombreLocal} — ${tituloReporte}`, 30, 30);
    doc.setFontSize(10);
    doc.setTextColor(110, 110, 110);
    doc.text(`Período: ${desde} a ${hasta}  ·  Generado: ${new Date().toLocaleString("es-AR")}`, 30, 46);

    doc.autoTable({
      html: tabla,
      startY: 58,
      theme: "grid",
      styles: { fontSize: 9, cellPadding: 5 },
      headStyles: { fillColor: [18, 32, 71], textColor: [255, 255, 255] }
    });

    const nombreArchivo = `${tituloReporte.replace(/\s+/g, "_")}_${desde}_a_${hasta}.pdf`;
    doc.save(nombreArchivo);

  } catch (error) {
    console.error("Error al exportar el reporte a PDF:", error);
    toast("No se pudo generar el PDF", "error");
  }
}


/* ===================================================================
   REPORTE DE COMPRAS — cruza productos vendidos con stock actual para
   guiar qué reponer, cuándo y en qué categoría invertir el presupuesto.
   Reusa el endpoint reporteProductosVendidos y el catálogo ya cargado
   en productosAdminGlobal (o lo pide si todavía no está en memoria).
=================================================================== */

let _rcCharts = {}; // instancias de Chart.js activas, para poder destruirlas antes de re-dibujar

/** Días hacia adelante que "Sugerido reponer" y "Presupuesto sugerido" intentan
 *  cubrir. Es independiente del rango de fechas de arriba: ese rango solo define
 *  con qué venta diaria promedio se calcula, no para cuántos días se compra. */
let _rcDiasCobertura = 30;

/** Compras usa el mismo período que el resto de Reportes */
function _rcRangoFechas() {
  return obtenerRangoReportes();
}

function _rcDiasDelRango(desdeStr, hastaStr) {
  if (!desdeStr || !hastaStr) return 30;
  const d = new Date(desdeStr), h = new Date(hastaStr);
  const dias = Math.round((h - d) / 86400000) + 1;
  return dias > 0 ? dias : 30;
}

async function cargarReporteCompras(forzar) {
  const tbody = document.getElementById("rcSemaforoTabla");
  if (tbody && !_rcProductosActuales.length) tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">Cargando...</td></tr>`;
  try {
    // 1) Productos vendidos del período — la MISMA respuesta que usa la
    //    pestaña Productos (se pide una sola vez y queda en caché)
    const data = await _obtenerReporteProductosVendidos(forzar);
    if (!data || !data.success) { toast((data && data.message) || "No se pudo cargar el reporte de compras", "error"); return; }

    sincronizarRangoReportes(data.desde, data.hasta);
    const dias = _rcDiasDelRango(data.desde, data.hasta);

    // 2) Stock y categoría: el backend nuevo ya los manda en cada producto.
    //    Con un backend viejo se cruzan con el catálogo (que hay que cargar).
    const backendTraeStock = (data.productos || []).some(v => v.STOCK !== undefined);
    const stockPorCodigo = {};
    if (!backendTraeStock) {
      if (!productosAdminGlobal || productosAdminGlobal.length === 0) await cargarProductos();
      (productosAdminGlobal || []).forEach(p => { stockPorCodigo[String(p.CODIGO)] = p; });
    }

    const productos = (data.productos || []).map(v => {
      const info = backendTraeStock ? { STOCK: v.STOCK, CATEGORIA: v.CATEGORIA, ALIAS: v.ALIAS, CODIGO_CAJA: v.CODIGO_CAJA } : (stockPorCodigo[String(v.CODIGO)] || {});
      const p = {
        codigo:    String(v.CODIGO ?? ""),
        nombre:    v.PRODUCTO,
        categoria: info.CATEGORIA || "Sin categoría",
        vendidos:  Number(v.VENDIDOS || 0),
        ingresos:  Number(v.INGRESOS || 0),
        stock:     (info.STOCK !== undefined && info.STOCK !== null && info.STOCK !== "") ? Number(info.STOCK) : 0,
      };
      p._busqueda = normalizarBusquedaPOS([p.nombre, p.codigo, p.categoria, info.ALIAS, info.CODIGO_CAJA].join(" "));
      return p;
    });

    // 4) Pedir la tendencia diaria por categoría (endpoint nuevo — ver nota al final
    //    de este archivo con el código de Apps Script a agregar). Si el backend
    //    todavía no lo tiene, cae automáticamente a una sola línea total.
    const tendenciaPorCategoria = await _rcCargarTendenciaPorCategoria(data.desde, data.hasta);

    _renderReporteCompras(productos, dias, tendenciaPorCategoria);

  } catch (error) {
    console.error("Error al cargar reporte de compras:", error);
    if (tbody) tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">${escapeHtml(_repMensajeError(error))} <button class="btn btn-outline-primary btn-sm ms-2" onclick="cargarReporteCompras(true)">Reintentar</button></td></tr>`;
    toast(_repMensajeError(error), "error");
    _repTabRangoCargado.compras = null;
  }
}

/**
 * Pide la venta diaria desglosada por categoría para el gráfico de tendencia.
 * Requiere una acción nueva en el backend (reporteVentasDiariasPorCategoria) que
 * todavía no existe en el Apps Script actual — ver el comentario al final de este
 * archivo con el código para agregarla. Mientras no exista, devuelve null y el
 * gráfico cae automáticamente a una sola línea con el total de ventas por día
 * (usando ventasPOSHistorial, que sí existe hoy).
 *
 * Formato esperado de una respuesta exitosa:
 *   { success:true, categorias:["Bebidas","Snacks",...],
 *     dias:[{ fecha:"2026-07-01", valores:{ "Bebidas":12000, "Snacks":4500 } }, ...] }
 */
async function _rcCargarTendenciaPorCategoria(desde, hasta) {
  try {
    const params = new URLSearchParams({ action: "reporteVentasDiariasPorCategoria" });
    if (desde) params.set("desde", desde);
    if (hasta) params.set("hasta", hasta);
    const response = await fetchAPI(API_URL + "?" + params.toString(), {}, _REP_OPC_FETCH);
    const data = await response.json();
    if (data && data.success && data.dias) return data;
    return null;
  } catch (error) {
    console.warn("reporteVentasDiariasPorCategoria no disponible todavía en el backend:", error);
    return null;
  }
}

/** Fallback: total de ventas por día (sin desglose de categoría), usando el
 *  historial que sí existe hoy en el backend. */
async function _rcCargarVentasDelRango(desde, hasta) {
  try {
    const params = new URLSearchParams({ action: "ventasPOSHistorial" });
    if (desde) params.set("desde", desde);
    if (hasta) params.set("hasta", hasta);
    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();
    return data.ventas || [];
  } catch (error) {
    console.warn("No se pudo cargar el historial de ventas para la tendencia:", error);
    return [];
  }
}

function sincronizarRangoReportesCompras(desde, hasta) {
  sincronizarRangoReportes(desde, hasta);
}

function _rcVentaDiaria(p, dias) { return p.vendidos / dias; }

function _rcCobertura(p, dias) {
  const vd = _rcVentaDiaria(p, dias);
  if (vd <= 0) return Infinity;
  return p.stock / vd;
}

function _rcEstado(p, dias) {
  if (p.stock <= 0 && p.vendidos > 0) return "sinstock";
  const c = _rcCobertura(p, dias);
  if (c < 7) return "critico";
  if (c < 20) return "atencion";
  return "ok";
}

/* Estado en memoria de la tabla semáforo, para poder re-filtrar sin volver a pedir datos */
let _rcProductosActuales = [];
let _rcDiasActuales = 30;

const _RC_PALETA_CATEGORIAS = ["#2563eb","#16a34a","#d97706","#dc2626","#7c3aed","#0891b2","#db2777","#65a30d"];

function _renderReporteCompras(productos, dias, tendenciaPorCategoria) {
  _rcProductosActuales = productos || [];
  _rcDiasActuales = dias;

  _rcSeleccionados.clear();
  _rcActualizarBarraSeleccion();

  if (!productos || productos.length === 0) {
    document.getElementById("rcSemaforoTabla").innerHTML =
      `<tr><td colspan="9" class="text-center text-muted py-3">Sin ventas para el rango elegido</td></tr>`;
    ["rcKpiSinStock","rcKpiCritico"].forEach(id => actualizarElemento(id, 0));
    ["rcKpiTicket","rcKpiPresupuesto"].forEach(id => actualizarElemento(id, "$0"));
    Object.values(_rcCharts).forEach(c => c && c.destroy());
    _rcCharts = {};
    return;
  }

  /* ---- KPIs ---- */
  const sinStockConDemanda = productos.filter(p => p.stock <= 0 && p.vendidos > 0);
  const criticos = productos.filter(p => _rcEstado(p, dias) === "critico");
  const totalIngresos = productos.reduce((a,p)=>a+p.ingresos,0);
  const totalVendidos = productos.reduce((a,p)=>a+p.vendidos,0);
  const ticketProm = totalVendidos > 0 ? totalIngresos / totalVendidos : 0;

  // presupuesto sugerido: para productos en rojo/amarillo/sin stock, cubrir N días (selector) de venta al precio unitario estimado
  const presupuesto = productos
    .filter(p => ["critico","atencion","sinstock"].includes(_rcEstado(p, dias)))
    .reduce((acc,p) => {
      const precioUnit = p.vendidos > 0 ? p.ingresos / p.vendidos : 0;
      const faltante = Math.max(0, (_rcVentaDiaria(p, dias) * _rcDiasCobertura) - p.stock);
      return acc + faltante * precioUnit;
    }, 0);

  actualizarElemento("rcKpiSinStock",   sinStockConDemanda.length);
  actualizarElemento("rcKpiCritico",    criticos.length);
  actualizarElemento("rcKpiTicket",     "$" + Math.round(ticketProm).toLocaleString("es-AR"));
  actualizarElemento("rcKpiPresupuesto","$" + Math.round(presupuesto).toLocaleString("es-AR"));

  /* ---- Gráficos: los datos se guardan siempre, pero solo se dibujan si el
     acordeón "Ver gráficos y tendencias" ya está abierto — dibujar un canvas
     oculto (details cerrado = ancho 0) deja los charts rotos hasta que se
     redimensiona la ventana. Si está cerrado, _rcDibujarGraficos() se llama
     recién cuando el usuario lo abre (ver onclick del <summary>). ---- */
  _rcDatosGraficosPendientes = { productos, tendenciaPorCategoria };
  const detailsGraficos = document.getElementById("rcGraficosDetails");
  if (detailsGraficos && detailsGraficos.open) {
    _rcDibujarGraficos(productos, tendenciaPorCategoria);
  }

  /* ---- Tabla semáforo ---- */
  _rcAplicarFiltrosTabla();
}

/* Datos en espera de dibujarse la primera vez que se abre el acordeón de gráficos */
let _rcDatosGraficosPendientes = null;

/** Se llama al abrir/cerrar el <details> de gráficos. Si se acaba de abrir y
 *  todavía no se dibujó nada con los datos actuales, dibuja ahora que el
 *  canvas ya tiene ancho real. */
function _rcRedibujarSiEstaAbierto() {
  const detailsGraficos = document.getElementById("rcGraficosDetails");
  if (!detailsGraficos || !detailsGraficos.open || !_rcDatosGraficosPendientes) return;
  _rcDibujarGraficos(_rcDatosGraficosPendientes.productos, _rcDatosGraficosPendientes.tendenciaPorCategoria);
}

function _rcDibujarGraficos(productos, tendenciaPorCategoria) {
  /* ---- Destruir gráficos previos antes de re-dibujar (evita fugas al cambiar de fecha) ---- */
  Object.values(_rcCharts).forEach(c => c && c.destroy());
  _rcCharts = {};

  /* ---- Chart 1: Top productos, barras superpuestas (vendidos vs stock) ---- */
  const top = [...productos].sort((a,b)=>b.vendidos-a.vendidos).slice(0,8);
  _rcCharts.top = new Chart(document.getElementById("rcChartTop"), {
    type: "bar",
    data: {
      labels: top.map(p=>p.nombre),
      datasets: [
        { label:"Vendidos", data: top.map(p=>p.vendidos), backgroundColor:"#2563eb", borderRadius:5 },
        { label:"Stock actual", data: top.map(p=>p.stock), backgroundColor:"#cbd2dd", borderRadius:5 },
      ]
    },
    options: {
      indexAxis: "y",
      plugins:{ legend:{ position:"bottom", labels:{ boxWidth:12, font:{size:11} } } },
      scales:{ x:{ grid:{ color:"#eef1f6" } }, y:{ grid:{ display:false } } }
    }
  });

  /* ---- Chart 2: Donut por categoría ---- */
  const categorias = {};
  productos.forEach(p => { categorias[p.categoria] = (categorias[p.categoria]||0) + p.ingresos; });
  const totalCategorias = Object.values(categorias).reduce((a,b)=>a+b, 0);
  _rcCharts.donut = new Chart(document.getElementById("rcChartDonut"), {
    type:"doughnut",
    data:{
      labels:Object.keys(categorias),
      datasets:[{ data:Object.values(categorias), backgroundColor:_RC_PALETA_CATEGORIAS, borderWidth:2, borderColor:"#fff" }]
    },
    options:{
      plugins:{
        legend:{ position:"bottom", labels:{ boxWidth:12, font:{size:11} } },
        tooltip:{
          callbacks:{
            label: (ctx) => {
              const valor = Number(ctx.raw || 0);
              const pct = totalCategorias > 0 ? (valor / totalCategorias * 100) : 0;
              return `${ctx.label}: $${Math.round(valor).toLocaleString("es-AR")} (${pct.toFixed(1)}%)`;
            }
          }
        }
      },
      cutout:"62%"
    }
  });

  /* ---- Chart 3: Tendencia — apilada por categoría, agrupada y simplificada ---- */
  _rcTendenciaActual = tendenciaPorCategoria; // se guarda para poder redibujar al cambiar el selector Día/Semana
  _rcRedibujarTendencia();

  _rcDatosGraficosPendientes = null;
}

/* Estado en memoria de la tendencia, para redibujar sin re-pedir datos al cambiar el selector */
let _rcTendenciaActual = null;

/** Agrupa los días en semanas (lunes a domingo), sumando cada categoría. */
function _rcAgruparPorSemana(dias) {
  const semanas = {}; // "yyyy-MM-dd" (lunes de esa semana) -> { categoria: total }
  const ordenPrimeraFecha = {};
  dias.forEach(d => {
    const fecha = new Date(d.fecha + "T12:00:00"); // mediodía evita líos de huso horario al restar días
    const diaSemana = fecha.getDay(); // 0=domingo … 6=sábado
    const offsetHastaLunes = diaSemana === 0 ? 6 : diaSemana - 1;
    const lunes = new Date(fecha);
    lunes.setDate(fecha.getDate() - offsetHastaLunes);
    const clave = _fechaISOEnZona(lunes);

    if (!semanas[clave]) { semanas[clave] = {}; ordenPrimeraFecha[clave] = clave; }
    Object.entries(d.valores || {}).forEach(([cat, val]) => {
      semanas[clave][cat] = (semanas[clave][cat] || 0) + Number(val || 0);
    });
  });
  return Object.keys(semanas).sort().map(clave => ({ fecha: clave, valores: semanas[clave] }));
}

/** Se queda con las categorías de mayor ingreso total y agrupa el resto en "Otras",
 *  para no saturar el gráfico con demasiadas líneas/áreas finitas. */
function _rcAplicarTopCategorias(dias, categorias, maxCategorias) {
  const totalPorCategoria = {};
  categorias.forEach(c => totalPorCategoria[c] = 0);
  dias.forEach(d => Object.entries(d.valores || {}).forEach(([c, v]) => {
    totalPorCategoria[c] = (totalPorCategoria[c] || 0) + Number(v || 0);
  }));

  const ordenadas = [...categorias].sort((a,b) => totalPorCategoria[b] - totalPorCategoria[a]);
  if (ordenadas.length <= maxCategorias) return { categoriasFinal: ordenadas, dias };

  const top = ordenadas.slice(0, maxCategorias);
  const resto = ordenadas.slice(maxCategorias);

  const diasFinal = dias.map(d => {
    const valores = {};
    top.forEach(c => { if (d.valores && d.valores[c] !== undefined) valores[c] = d.valores[c]; });
    let otras = 0;
    resto.forEach(c => { otras += Number((d.valores && d.valores[c]) || 0); });
    if (otras > 0) valores["Otras"] = otras;
    return { fecha: d.fecha, valores };
  });

  const categoriasFinal = [...top];
  if (diasFinal.some(d => d.valores["Otras"] !== undefined)) categoriasFinal.push("Otras");
  return { categoriasFinal, dias: diasFinal };
}

function _rcFormatoFechaLabel(fechaStr, esSemanal) {
  const f = new Date(fechaStr + "T12:00:00");
  const dd = String(f.getDate()).padStart(2,"0");
  const mm = String(f.getMonth()+1).padStart(2,"0");
  return esSemanal ? `Sem. ${dd}/${mm}` : `${dd}/${mm}`;
}

/** Redibuja el gráfico de tendencia a partir de los datos ya cargados, aplicando
 *  la agrupación (día/semana/automático) elegida en el selector. No vuelve a
 *  pedir datos al backend — solo re-procesa lo que ya está en memoria. */
function _rcRedibujarTendencia() {
  const canvasLinea = document.getElementById("rcChartLinea");
  const aviso = document.getElementById("rcTendenciaAviso");
  if (!canvasLinea) return;

  const tendenciaPorCategoria = _rcTendenciaActual;

  if (tendenciaPorCategoria && tendenciaPorCategoria.dias && tendenciaPorCategoria.dias.length) {
    if (aviso) aviso.style.display = "none";

    const modo = document.getElementById("rcTendenciaAgrupacion")?.value || "auto";
    const esSemanal = modo === "semana" || (modo === "auto" && tendenciaPorCategoria.dias.length > 45);

    let dias = tendenciaPorCategoria.dias;
    if (esSemanal) dias = _rcAgruparPorSemana(dias);

    const { categoriasFinal, dias: diasFinal } = _rcAplicarTopCategorias(dias, tendenciaPorCategoria.categorias || [], 6);

    const labels = diasFinal.map(d => _rcFormatoFechaLabel(d.fecha, esSemanal));
    const coloresBase = _RC_PALETA_CATEGORIAS;
    const datasets = categoriasFinal.map((cat, i) => {
      const esOtras = cat === "Otras";
      const color = esOtras ? "#94a3b8" : coloresBase[i % coloresBase.length];
      return {
        label: cat,
        data: diasFinal.map(d => Number((d.valores && d.valores[cat]) || 0)),
        borderColor: color,
        backgroundColor: "transparent",
        fill: false, // sin apilar: cada línea muestra su propio valor real, para que una
                     // categoría chica se vea claramente chata frente a una grande, y no
                     // "infle" de tamaño por estar apoyada arriba de otra en un área apilada
        tension: .3, pointRadius: 0, borderWidth: 2
      };
    });

    if (_rcCharts.linea) _rcCharts.linea.destroy();
    _rcCharts.linea = new Chart(canvasLinea, {
      type:"line",
      data:{ labels, datasets },
      options:{
        interaction:{ mode:"index", intersect:false },
        plugins:{
          legend:{ position:"bottom", labels:{ boxWidth:12, font:{size:11} } },
          tooltip:{
            callbacks:{
              label: (ctx) => `${ctx.dataset.label}: $${Math.round(ctx.raw).toLocaleString("es-AR")}`
            }
          }
        },
        scales:{
          x:{ grid:{ display:false }, ticks:{ maxTicksLimit:10 } },
          y:{ grid:{ color:"#eef1f6" }, ticks:{ callback:(v)=>"$"+Number(v).toLocaleString("es-AR") } }
        }
      }
    });
  } else {
    // Fallback: todavía no existe el endpoint por categoría — se avisa y se
    // muestra una sola línea con el total (mejor que no mostrar nada).
    if (aviso) {
      aviso.style.display = "block";
      aviso.textContent = "⚠️ Tu backend todavía no tiene el desglose diario por categoría — mostrando el total general. Pedime el código de Apps Script para agregarlo.";
    }
    _rcCargarVentasDelRango(document.getElementById("repDesde").value, document.getElementById("repHasta").value)
      .then(ventasDelRango => {
        const porDia = {};
        (ventasDelRango || []).forEach(v => {
          if (String(v.ANULADA || "").toUpperCase() === "SI") return;
          if (!v.FECHA) return;
          const f = String(v.FECHA).slice(0, 10);
          porDia[f] = (porDia[f] || 0) + Number(v.TOTAL || 0);
        });
        const fechasOrdenadas = Object.keys(porDia).sort();
        if (_rcCharts.linea) _rcCharts.linea.destroy();
        _rcCharts.linea = new Chart(canvasLinea, {
          type:"line",
          data:{
            labels: fechasOrdenadas.length ? fechasOrdenadas.map(f=>_rcFormatoFechaLabel(f,false)) : ["Sin datos"],
            datasets:[{ label:"Ventas diarias (total)", data: fechasOrdenadas.length ? fechasOrdenadas.map(f=>porDia[f]) : [0],
              borderColor:"#2563eb", backgroundColor:"rgba(37,99,235,.08)", fill:true, tension:.35, pointRadius:0, borderWidth:2 }]
          },
          options:{
            plugins:{ legend:{ display:false } },
            scales:{ x:{ grid:{ display:false }, ticks:{ maxTicksLimit:8 } }, y:{ grid:{ color:"#eef1f6" } } }
          }
        });
      });
  }
}

/** Re-renderiza SOLO la tabla semáforo aplicando los filtros de estado y cantidad,
 *  sin volver a pedir datos ni redibujar los gráficos. */
/** Se llama al cambiar el selector "Cubrir próximos N días". No vuelve a pedir
 *  datos al backend — recalcula con lo que ya está en memoria: el KPI de
 *  presupuesto sugerido, la columna "Sugerido reponer" de la tabla, y el total
 *  estimado de la lista de compra (si hay productos tildados). */
function _rcCambiarDiasCobertura() {
  const select = document.getElementById("rcDiasCobertura");
  _rcDiasCobertura = Number(select?.value || 30);

  const productos = _rcProductosActuales;
  const dias = _rcDiasActuales;
  if (productos && productos.length > 0) {
    const presupuesto = productos
      .filter(p => ["critico","atencion","sinstock"].includes(_rcEstado(p, dias)))
      .reduce((acc,p) => {
        const precioUnit = p.vendidos > 0 ? p.ingresos / p.vendidos : 0;
        const faltante = Math.max(0, (_rcVentaDiaria(p, dias) * _rcDiasCobertura) - p.stock);
        return acc + faltante * precioUnit;
      }, 0);
    actualizarElemento("rcKpiPresupuesto", "$" + Math.round(presupuesto).toLocaleString("es-AR"));
  }

  _rcAplicarFiltrosTabla();
  _rcActualizarBarraSeleccion();
}

function _rcAplicarFiltrosTabla() {
  const productos = _rcProductosActuales;
  const dias = _rcDiasActuales;
  const tbody = document.getElementById("rcSemaforoTabla");
  if (!tbody) return;

  if (!productos || productos.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">Sin ventas para el rango elegido</td></tr>`;
    return;
  }

  const filtroEstado = document.getElementById("rcFiltroEstado")?.value || "todos";
  const filtroCantidad = document.getElementById("rcFiltroCantidad")?.value || "10";
  const busqueda = document.getElementById("rcBuscarProducto")?.value || "";

  // Sin stock se marca oscuro para distinguirlo de un rojo "crítico" pero
  // todavía con algo de stock — son dos urgencias distintas de un vistazo.
  const estadoInfo = {
    sinstock: { clase:"bg-dark",    texto:"Sin stock" },
    critico:  { clase:"bg-danger",  texto:"🔴 Crítico" },
    atencion: { clase:"bg-warning text-dark", texto:"🟡 Atención" },
    ok:       { clase:"bg-success", texto:"🟢 OK" },
  };

  let filtrados = productos.filter(p => filtroEstado === "todos" || _rcEstado(p, dias) === filtroEstado);
  if (busqueda.trim()) {
    filtrados = filtrados.filter(p => coincideBusquedaReportes(p._busqueda || normalizarBusquedaPOS(p.nombre + " " + p.codigo), busqueda));
  }
  let ordenados = filtrados.sort((a,b)=> _rcCobertura(a, dias) - _rcCobertura(b, dias));
  // Buscando, se muestran TODAS las coincidencias (el Top 10/20 escondía resultados)
  if (filtroCantidad !== "todos" && !busqueda.trim()) ordenados = ordenados.slice(0, Number(filtroCantidad));

  if (ordenados.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">Ningún producto en ese estado</td></tr>`;
    _rcActualizarCheckTodos();
    return;
  }

  tbody.innerHTML = ordenados.map(p => {
    const e = _rcEstado(p, dias);
    const info = estadoInfo[e];
    const cob = _rcCobertura(p, dias);
    const cobTxt = cob === Infinity ? "—" : `~${Math.round(cob)} días`;

    // Sugerido reponer: cuánto falta para cubrir _rcDiasCobertura días de venta al ritmo actual.
    // Solo tiene sentido mostrarlo para lo que hay que reponer; en "OK" no se sugiere nada.
    const faltante = ["sinstock","critico","atencion"].includes(e)
      ? Math.max(0, Math.round(_rcVentaDiaria(p, dias) * _rcDiasCobertura - p.stock))
      : 0;
    const sugeridoTxt = faltante > 0 ? `<strong>${faltante.toLocaleString("es-AR")}</strong> uds.` : "—";

    const marcado = _rcSeleccionados.has(String(p.codigo));
    return `
      <tr>
        <td><input type="checkbox" class="rc-check-fila" data-codigo="${escapeHtml(p.codigo)}" ${marcado ? "checked" : ""} onchange="_rcToggleSeleccion('${escapeJsAttr(p.codigo)}', this.checked)"></td>
        <td class="mono" title="${escapeHtml(p.codigo)}">${escapeHtml(p.codigo)}</td>
        <td title="${escapeHtml(p.nombre)}">${escapeHtml(p.nombre)}</td>
        <td title="${escapeHtml(p.categoria)}">${escapeHtml(p.categoria)}</td>
        <td class="money" data-label="Vendidos">${p.vendidos.toLocaleString("es-AR")}</td>
        <td class="money" data-label="Stock">${p.stock.toLocaleString("es-AR")}</td>
        <td class="money" data-label="Cobertura">${cobTxt}</td>
        <td class="money" data-label="Sugerido reponer">${sugeridoTxt}</td>
        <td><span class="badge ${info.clase}">${info.texto}</span></td>
      </tr>`;
  }).join("");

  _rcActualizarCheckTodos();
}

/* ===================================================================
   Lista de compra — tildar productos del semáforo (sin gráficos ni
   modales, solo checkboxes) y copiar un resumen listo para mandar
   al proveedor o pegar en una nota.
=================================================================== */
const _rcSeleccionados = new Set(); // códigos tildados, persiste entre re-renders por filtro/búsqueda

function _rcToggleSeleccion(codigo, marcado) {
  if (marcado) _rcSeleccionados.add(String(codigo));
  else _rcSeleccionados.delete(String(codigo));
  _rcActualizarCheckTodos();
  _rcActualizarBarraSeleccion();
}

/** Tilda/destilda todas las filas actualmente visibles en la tabla (no las que
 *  quedaron ocultas por el filtro), igual que el "seleccionar todos" de Productos. */
function _rcToggleTodos(headerCheckbox) {
  document.querySelectorAll(".rc-check-fila").forEach(chk => {
    chk.checked = headerCheckbox.checked;
    const codigo = chk.dataset.codigo;
    if (headerCheckbox.checked) _rcSeleccionados.add(codigo);
    else _rcSeleccionados.delete(codigo);
  });
  _rcActualizarBarraSeleccion();
}

/** Refleja en el checkbox de cabecera si todas/algunas/ninguna de las filas visibles están tildadas. */
function _rcActualizarCheckTodos() {
  const checkTodos = document.getElementById("rcCheckTodos");
  if (!checkTodos) return;
  const filas = Array.from(document.querySelectorAll(".rc-check-fila"));
  if (filas.length === 0) { checkTodos.checked = false; checkTodos.indeterminate = false; return; }
  const marcadas = filas.filter(f => f.checked).length;
  checkTodos.checked = marcadas === filas.length;
  checkTodos.indeterminate = marcadas > 0 && marcadas < filas.length;
}

function _rcVaciarSeleccion() {
  _rcSeleccionados.clear();
  document.querySelectorAll(".rc-check-fila").forEach(chk => chk.checked = false);
  _rcActualizarCheckTodos();
  _rcActualizarBarraSeleccion();
}

/** Muestra/oculta la barra "N seleccionados" y actualiza el total estimado de compra. */
function _rcActualizarBarraSeleccion() {
  const barra = document.getElementById("rcListaCompraBar");
  const info = document.getElementById("rcListaCompraInfo");
  if (!barra || !info) return;

  const cantidad = _rcSeleccionados.size;
  if (cantidad === 0) { barra.style.display = "none"; return; }

  const dias = _rcDiasActuales;
  let totalEstimado = 0;
  _rcSeleccionados.forEach(codigo => {
    const p = _rcProductosActuales.find(x => String(x.codigo) === codigo);
    if (!p) return;
    const precioUnit = p.vendidos > 0 ? p.ingresos / p.vendidos : 0;
    const faltante = Math.max(0, _rcVentaDiaria(p, dias) * _rcDiasCobertura - p.stock);
    totalEstimado += faltante * precioUnit;
  });

  barra.style.display = "block";
  info.textContent = `${cantidad} producto${cantidad === 1 ? "" : "s"} seleccionado${cantidad === 1 ? "" : "s"} · ≈$${Math.round(totalEstimado).toLocaleString("es-AR")} estimado`;
}

/** Arma un texto simple (código, nombre, cantidad sugerida) con los productos
 *  tildados y lo copia al portapapeles — para pegarlo en un chat/nota al proveedor. */
async function _rcCopiarListaCompra() {
  if (_rcSeleccionados.size === 0) { toast("Tildá al menos un producto", "error"); return; }

  const dias = _rcDiasActuales;
  const lineas = [];
  _rcSeleccionados.forEach(codigo => {
    const p = _rcProductosActuales.find(x => String(x.codigo) === codigo);
    if (!p) return;
    const faltante = Math.max(0, Math.round(_rcVentaDiaria(p, dias) * _rcDiasCobertura - p.stock));
    lineas.push(`${p.codigo} — ${p.nombre} — ${faltante > 0 ? faltante + " uds." : "a definir"}`);
  });

  const texto = `Lista de compra (${lineas.length} productos)\n` + lineas.join("\n");

  try {
    await navigator.clipboard.writeText(texto);
    toast("Lista de compra copiada al portapapeles", "success");
  } catch (error) {
    console.error("No se pudo copiar la lista de compra:", error);
    toast("No se pudo copiar automáticamente — seleccioná y copiá el texto manualmente", "error");
  }
}

/** Filtro rápido desde las tarjetas KPI "Sin stock" / "Cobertura < 7 días":
 *  aplica el filtro de estado correspondiente y lleva la vista a la tabla. */
function _rcFiltrarDesdeKpi(estado) {
  const select = document.getElementById("rcFiltroEstado");
  if (select) { select.value = estado; _rcAplicarFiltrosTabla(); }
  const tabla = document.getElementById("rcSemaforoCard");
  if (tabla) tabla.scrollIntoView({ behavior: "smooth", block: "start" });
}

/* =====================================================================
   Backend: la acción "reporteVentasDiariasPorCategoria" ya está agregada
   en code.gs (misma lógica que reporteVentasPorCategoria, agrupada
   también por día). El fallback de más arriba queda como red de
   seguridad por si alguna vez el deploy del backend queda desactualizado.
===================================================================== */

/* ===================================================================
   MOVIMIENTOS DE CAJA — ingresos y egresos manuales (no son ventas)
=================================================================== */

let tipoMovimientoCajaActivo = "INGRESO";

function elegirTipoMovimientoCaja(el, tipo) {
  tipoMovimientoCajaActivo = tipo;
  document.querySelectorAll(".mc-tipo-btn").forEach(b => b.classList.remove("active"));
  el.classList.add("active");
}

/** Loads today's manual cash movements and the running totals */
async function cargarMovimientosCajaHoy() {
  // Leer fecha del selector — default hoy
  const selector = document.getElementById("mcFechaSelector");
  if (selector && !selector.value) {
    const hoy = _hoyISO();
    selector.value = hoy;
  }
  const fecha = selector ? selector.value : _hoyISO();

  // Actualizar labels de resumen con la fecha seleccionada
  const fechaLabel = fecha === _hoyISO()
    ? "hoy"
    : new Date(fecha + "T12:00:00").toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit" });
  const li = document.getElementById("mcLabelIngresos");
  const le = document.getElementById("mcLabelEgresos");
  if (li) li.textContent = `Ingresos — ${fechaLabel}`;
  if (le) le.textContent = `Egresos — ${fechaLabel}`;

  // Caché 60 segundos por fecha
  // v2: se cambió el nombre de la clave para descartar cualquier caché
  // vieja guardada en localStorage antes de este fix (sobre todo en
  // Electron, donde localStorage persiste entre sesiones y podía
  // quedar con movimientos sin MOVIMIENTO_ID armado correctamente,
  // causando el error "No se encontró el movimiento 'undefined'").
  const CACHE_KEY = "vpos_cache_movimientos_v2_" + fecha;
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (raw) {
      const { ts, data } = JSON.parse(raw);
      if (Date.now() - ts < 60000) {
        _aplicarMovimientos(data);
        return;
      }
    }
  } catch(e) {}

  try {
    const response = await fetchAPI(API_URL + "?action=movimientosCajaHoy&fecha=" + encodeURIComponent(fecha));
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo cargar movimientos de caja", "error");
      return;
    }

    try { localStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), data })); } catch(e) {}
    _aplicarMovimientos(data);

  } catch (error) {
    console.error("Error al cargar movimientos de caja:", error);
    toast("Error de conexión al cargar movimientos de caja", "error");
  }
}

function _aplicarMovimientos(data) {
  actualizarElemento("mcTotalIngresos", "$" + Number(data.totalIngresos).toLocaleString("es-AR"));
  actualizarElemento("mcTotalEgresos",  "$" + Number(data.totalEgresos).toLocaleString("es-AR"));
  const netoEl = document.getElementById("mcSaldoNeto");
  if (netoEl) {
    netoEl.textContent = (data.neto >= 0 ? "$" : "-$") + Math.abs(data.neto).toLocaleString("es-AR");
    netoEl.style.color = data.neto >= 0 ? "var(--green-600)" : "var(--red-500)";
  }
  renderTablaMovimientosCaja(data.movimientos || []);
}

function renderTablaMovimientosCaja(lista) {
  const tbody = document.getElementById("tablaMovimientosCaja");
  if (!tbody) return;

  if (lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-3">Todavía no hay movimientos registrados</td></tr>`;
    return;
  }

  const iconoFormaPago = { EFECTIVO: "💵", TRANSFERENCIA: "📲", TARJETA: "💳" };

  let html = "";
  lista.forEach(m => {
    const esIngreso = String(m.TIPO).toUpperCase() === "INGRESO";
    const hora = m.FECHA_REGISTRO
      ? new Date(m.FECHA_REGISTRO).toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" })
      : "—";
    const fp = String(m.FORMA_PAGO || "EFECTIVO").toUpperCase();
    const fpIcon = iconoFormaPago[fp] || "💵";
    const fpLabel = fp.charAt(0) + fp.slice(1).toLowerCase();

    const esVendedorRol = obtenerRolActual() === "vendedor";
    html += `
    <tr>
      <td>${hora}</td>
      <td><span class="mc-tag ${esIngreso ? "mc-ingreso" : "mc-egreso"}">${esIngreso ? "Ingreso" : "Egreso"}</span></td>
      <td>${escapeHtml(m.MOTIVO || "—")}</td>
      <td><span style="font-size:12.5px;">${fpIcon} ${fpLabel}</span></td>
      <td class="money" style="color:${esIngreso ? "var(--green-600)" : "var(--red-500)"};">
        ${esIngreso ? "+" : "-"}$${Number(m.MONTO || 0).toLocaleString("es-AR")}
      </td>
      <td>${escapeHtml(m.VENDEDOR || "—")}</td>
      <td>${(esVendedorRol || !m.MOVIMIENTO_ID) ? "" : `<button class="btn btn-outline-danger btn-sm"
        onclick="confirmarEliminarMovimiento('${escapeJsAttr(m.MOVIMIENTO_ID)}', '${escapeJsAttr(m.MOTIVO || "")}')">✕</button>`}</td>
    </tr>`;
  });

  tbody.innerHTML = html;
}

/** Pide confirmación con modal propio (confirm() bloquea el input en Electron) */
function confirmarEliminarMovimiento(movimientoId, motivo) {
  // Guardar ID en variable global para usar al confirmar
  window._movimientoParaEliminar = movimientoId;

  const backdrop = document.getElementById("modalEliminarMovBackdrop");
  const texto = document.getElementById("modalEliminarMovTexto");
  if (texto) texto.textContent = motivo
    ? `¿Eliminar el movimiento "${motivo}"? Esta acción no se puede deshacer.`
    : "¿Eliminar este movimiento? Esta acción no se puede deshacer.";
  if (backdrop) backdrop.classList.add("show");
}

function cerrarModalEliminarMov() {
  const backdrop = document.getElementById("modalEliminarMovBackdrop");
  if (backdrop) backdrop.classList.remove("show");
  window._movimientoParaEliminar = null;
}

async function eliminarMovimientoCajaForm(movimientoId) {
  const id = movimientoId || window._movimientoParaEliminar;
  cerrarModalEliminarMov();
  if (!id) return;

  if (obtenerRolActual() === "vendedor") {
    toast("Tu usuario (Vendedor) no tiene permiso para eliminar movimientos de caja", "error");
    return;
  }

  try {
    const response = await fetchAPI(API_URL + "?action=eliminarMovimientoCaja&movimientoId=" + encodeURIComponent(id) + "&rol=" + encodeURIComponent(obtenerRolActual()));
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo eliminar el movimiento", "error");
      return;
    }

    toast("Movimiento eliminado", "success");
    const fechaMov = document.getElementById("mcFechaSelector")?.value || _hoyISO();
    try { localStorage.removeItem("vpos_cache_movimientos_v2_" + fechaMov); } catch(e) {}
    try { localStorage.removeItem("vpos_cache_movimientos_" + fechaMov); } catch(e) {} // limpieza de clave vieja
    try { localStorage.removeItem("vpos_cache_cierre_" + fechaMov); } catch(e) {}
    try { localStorage.removeItem("vpos_cache_cierre_ayer"); } catch(e) {}
    cargarMovimientosCajaHoy();

  } catch (error) {
    console.error("Error al eliminar movimiento de caja:", error);
    toast("Error de conexión al eliminar el movimiento", "error");
  }
}

async function guardarMovimientoCajaForm() {
  const monto = document.getElementById("mcMonto").value;
  const motivo = document.getElementById("mcMotivo").value.trim();
  const formaPago = (document.getElementById("mcFormaPago")?.value || "EFECTIVO").toUpperCase();

  if (!monto || Number(monto) <= 0) { toast("Ingresá un monto válido", "error"); return; }
  if (!motivo) { toast("Ingresá un motivo para el movimiento", "error"); return; }

  const btn = document.getElementById("btnGuardarMovimientoCaja");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    const nombreUsuarioActual = sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("usuarioLogueado") || "ADMIN";
    const params = new URLSearchParams({
      action: "guardarMovimientoCaja",
      tipo: tipoMovimientoCajaActivo,
      monto: monto,
      motivo: motivo,
      formaPago: formaPago,
      vendedor: nombreUsuarioActual
    });

    const response = await fetchAPI(API_URL + "?" + params.toString());
    const data = await response.json();

    if (!data.success) {
      toast(data.message || "No se pudo registrar el movimiento", "error");
      return;
    }

    toast(tipoMovimientoCajaActivo === "INGRESO" ? "Ingreso registrado" : "Egreso registrado", "success");
    document.getElementById("mcMonto").value = "";
    document.getElementById("mcMotivo").value = "";

    // Invalidar caché para que el cierre de caja muestre datos frescos
    const fecha = document.getElementById("mcFechaSelector")?.value || _hoyISO();
    try { localStorage.removeItem("vpos_cache_movimientos_" + fecha); } catch(e) {}
    try { localStorage.removeItem("vpos_cache_cierre_" + fecha); } catch(e) {}
    try { localStorage.removeItem("vpos_cache_cierre_ayer"); } catch(e) {}

    cargarMovimientosCajaHoy();

  } catch (error) {
    console.error("Error al guardar movimiento de caja:", error);
    toast("Error de conexión al registrar el movimiento", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

/* ===================================================================
   ETIQUETAS DE CÓDIGO DE BARRAS (para estantería/góndola)
=================================================================== */

/** Updates the selection counter and enables/disables the three label/QR generation buttons */
/* Selección de productos: se guarda aparte de las casillas que se ven,
   así no se pierde al buscar otro producto o cambiar de categoría
   (antes la tabla se redibujaba y se borraba todo lo tildado). */
const _seleccionProductos = new Set();

function alCambiarCheckProducto(chk) {
  if (chk.checked) _seleccionProductos.add(String(chk.value));
  else _seleccionProductos.delete(String(chk.value));
  actualizarSeleccionEtiquetas();
}

function quitarSeleccionProductos() {
  _seleccionProductos.clear();
  document.querySelectorAll(".check-producto-etiqueta").forEach(c => { c.checked = false; });
  actualizarSeleccionEtiquetas();
}

/** Los tildados que ya no existen (borrados) se sacan de la selección. */
function _limpiarSeleccionInexistente() {
  if (!_seleccionProductos.size || !productosAdminGlobal || !productosAdminGlobal.length) return;
  const existen = new Set(productosAdminGlobal.map(p => String(p.CODIGO)));
  [..._seleccionProductos].forEach(c => { if (!existen.has(c)) _seleccionProductos.delete(c); });
}

function actualizarSeleccionEtiquetas() {
  _limpiarSeleccionInexistente();
  const visibles = [...document.querySelectorAll(".check-producto-etiqueta")];
  const checkTodos = document.getElementById("checkTodosProductos");
  if (checkTodos) {
    const tildadosVisibles = visibles.filter(c => c.checked).length;
    checkTodos.checked = visibles.length > 0 && tildadosVisibles === visibles.length;
    checkTodos.indeterminate = tildadosVisibles > 0 && tildadosVisibles < visibles.length;
  }
  const checks = { length: _seleccionProductos.size };
  const info = document.getElementById("etiquetasSeleccionInfo");
  const btnEtiquetas = document.getElementById("btnGenerarEtiquetas");
  const btnQR = document.getElementById("btnGenerarQR");
  const btnQROffline = document.getElementById("btnGenerarQROffline");
  const btnCambiarLote = document.getElementById("btnCambiarLote");

  const cantidad = checks.length;

  if (info) {
    if (cantidad === 0) {
      info.textContent = "Ningún producto seleccionado";
    } else {
      const visiblesTildados = visibles.filter(c => c.checked).length;
      const fuera = cantidad - visiblesTildados;
      info.innerHTML = `<b>${cantidad === 1 ? "1 producto seleccionado" : cantidad + " productos seleccionados"}</b>` +
        (fuera > 0 ? ` <span class="text-muted">(${fuera} no se ve${fuera === 1 ? "" : "n"} con este filtro)</span>` : "") +
        ` · <a href="#" onclick="quitarSeleccionProductos(); return false;">Quitar selección</a>`;
    }
  }

  if (btnEtiquetas) btnEtiquetas.disabled = cantidad === 0;
  if (btnQR) btnQR.disabled = cantidad === 0;
  if (btnQROffline) btnQROffline.disabled = cantidad === 0;
  if (btnCambiarLote) btnCambiarLote.disabled = cantidad === 0;
  if (cantidad === 0) cerrarMenuLote();
}

/* ===================================================================
   CAMBIOS EN LOTE (Productos): mostrar/ocultar, destacado, oferta para
   todos los tildados de una vez — ej. al cambiar de temporada: filtrar
   la categoría, tildar todos (casilla del encabezado) y elegir la acción.
   Va al backend en UNA sola llamada (actualizarProductosLote), que toca
   solo esas columnas: no pisa stock, precio ni fotos. (Antes "Hacer
   visibles" mandaba el producto entero, uno por uno.)
=================================================================== */
function alternarMenuLote(ev) {
  if (ev) ev.stopPropagation();
  const menu = document.getElementById("menuLoteProductos");
  const btn = document.getElementById("btnCambiarLote");
  if (!menu || !btn || btn.disabled) return;
  const abrir = !menu.classList.contains("abierto");
  menu.classList.toggle("abierto", abrir);
  btn.setAttribute("aria-expanded", abrir ? "true" : "false");
  if (abrir) setTimeout(() => menu.querySelector("button")?.focus(), 30);
}

function cerrarMenuLote() {
  const menu = document.getElementById("menuLoteProductos");
  if (menu) menu.classList.remove("abierto");
  document.getElementById("btnCambiarLote")?.setAttribute("aria-expanded", "false");
}

document.addEventListener("click", (e) => { if (!e.target.closest || !e.target.closest(".lote-menu-wrap")) cerrarMenuLote(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrarMenuLote(); });

function _textoCambioLote(cambios) {
  const partes = [];
  if (cambios.PUBLICADO) partes.push(cambios.PUBLICADO === "SI" ? "mostrar en el catálogo" : "ocultar del catálogo y del POS");
  if (cambios.DESTACADO) partes.push(cambios.DESTACADO === "SI" ? "marcar como destacados" : "quitar destacado");
  if (cambios.OFERTA) partes.push(cambios.OFERTA === "SI" ? "poner en oferta" : "quitar oferta");
  return partes.join(", ");
}

function aplicarCambioLote(cambios) {
  cerrarMenuLote();
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }

  const yaEstan = codigos.filter(cod => {
    const p = productosAdminGlobal.find(x => String(x.CODIGO) === String(cod));
    return p && Object.keys(cambios).every(k => {
      const actual = String(p[k] || "").trim().toUpperCase() || (k === "PUBLICADO" ? "SI" : "NO");
      return actual === cambios[k];
    });
  }).length;
  if (yaEstan === codigos.length) {
    toast(codigos.length === 1 ? "Ese producto ya estaba así" : "Los productos seleccionados ya estaban así", "success");
    return;
  }

  const accion = _textoCambioLote(cambios);
  const mensaje = `¿${accion.charAt(0).toUpperCase() + accion.slice(1)}: ${codigos.length === 1 ? "1 producto" : codigos.length + " productos"}?` +
    (yaEstan ? ` (${yaEstan} ya estaba${yaEstan === 1 ? "" : "n"} así)` : "") +
    (cambios.PUBLICADO === "NO" ? "\n\nLos productos ocultos no se ven en la página ni en el POS, pero no se borran: se vuelven a mostrar con \"Mostrar (publicar)\" (filtro \"Ocultos\")." : "");

  confirmarAccion(mensaje, async () => {
    const btn = document.getElementById("btnCambiarLote");
    const textoOriginal = btn ? btn.innerHTML : "";
    if (btn) { btn.disabled = true; btn.innerHTML = "Guardando…"; }
    try {
      const response = await fetchAPI(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "actualizarProductosLote", rol: obtenerRolActual(), codigos, cambios })
      }, { timeoutMs: 60000 });
      const data = await response.json();
      if (!data.success) { toast(data.message || "No se pudieron guardar los cambios", "error"); return; }

      // Reflejo local inmediato (la caché del catálogo se refresca abajo)
      const set = new Set(codigos.map(String));
      productosAdminGlobal.forEach(p => { if (set.has(String(p.CODIGO))) Object.assign(p, cambios); });
      const n = data.actualizados || 0;
      let msg = n === 0 ? "No hubo nada que cambiar" : (n === 1 ? "1 producto actualizado" : `${n} productos actualizados`);
      if (data.noEncontrados && data.noEncontrados.length) msg += ` · ${data.noEncontrados.length} no se encontraron (¿los borraron?)`;
      toast(msg + (n ? " — la página lo muestra en unos minutos" : ""), data.noEncontrados && data.noEncontrados.length ? "error" : "success");

      // Se limpia la selección: así un próximo cambio no se aplica sin querer a estos mismos
      if (n > 0 || !(data.noEncontrados && data.noEncontrados.length)) _seleccionProductos.clear();
      productosPOS = []; // el POS vuelve a leer el catálogo
      if (typeof window.refrescarCatalogoLocalDesdeSheets === "function") {
        try { await window.refrescarCatalogoLocalDesdeSheets(); } catch (e) {}
      }
      try { localStorage.setItem("vpos_cache_productosAdmin", JSON.stringify({ ts: Date.now(), data: productosAdminGlobal })); } catch (e) {}
      filtrarProductos();
    } catch (error) {
      console.error("Cambio en lote:", error);
      toast(navigator.onLine ? "No se pudo guardar: el servidor no respondió. Probá de nuevo." : "Hace falta internet para cambiar varios productos a la vez.", "error");
    } finally {
      if (btn) { btn.innerHTML = textoOriginal; }
      actualizarSeleccionEtiquetas();
    }
  }, "Cambiar productos seleccionados");
}

/** Compatibilidad: lo llamaba el botón viejo "Hacer visibles". */
function hacerVisiblesProductosSeleccionados() { aplicarCambioLote({ PUBLICADO: "SI" }); }

/** Toggles every visible product checkbox via the header checkbox */
function toggleTodosProductos(checkboxHeader) {
  // Solo los que muestra el filtro actual; los de otras búsquedas quedan como estaban
  document.querySelectorAll(".check-producto-etiqueta").forEach(c => {
    c.checked = checkboxHeader.checked;
    if (c.checked) _seleccionProductos.add(String(c.value)); else _seleccionProductos.delete(String(c.value));
  });
  actualizarSeleccionEtiquetas();
}

/** Códigos seleccionados (también los que no se ven con el filtro actual), en el orden del catálogo */
function obtenerCodigosSeleccionados() {
  _limpiarSeleccionInexistente();
  if (!_seleccionProductos.size) return [];
  const orden = (productosAdminGlobal || []).map(p => String(p.CODIGO)).filter(c => _seleccionProductos.has(c));
  const resto = [..._seleccionProductos].filter(c => orden.indexOf(c) === -1);
  return orden.concat(resto);
}

/* ===================================================================
   ETIQUETAS: tamaños (hoja A4 o rollo de impresora de etiquetas),
   qué datos llevan y estilo. Vista previa en vivo con los mismos
   estilos que la impresión. La última configuración se guarda en esta
   computadora (localStorage "vpos_etq_config").
=================================================================== */
const ETQ_TAMANOS = {
  A4: [
    { id: "a4-38x21",  txt: "Chica — 38 × 21 mm",          w: 38,  h: 21 },
    { id: "a4-48x25",  txt: "Chica — 48 × 25 mm",          w: 48,  h: 25 },
    { id: "a4-52x30",  txt: "Mediana chica — 52 × 30 mm", w: 52,  h: 30 },
    { id: "a4-63x38",  txt: "Mediana — 63 × 38 mm",        w: 63,  h: 38 },
    { id: "a4-99x57",  txt: "Grande — 99 × 57 mm",         w: 99,  h: 57 },
    { id: "a4-95x67",  txt: "Góndola — 95 × 67 mm",        w: 95,  h: 67 },
    { id: "custom",    txt: "Personalizada…" }
  ],
  ROLLO: [
    { id: "r-30x20",  txt: "30 × 20 mm",  w: 30,  h: 20 },
    { id: "r-40x30",  txt: "40 × 30 mm",  w: 40,  h: 30 },
    { id: "r-50x25",  txt: "50 × 25 mm",  w: 50,  h: 25 },
    { id: "r-50x30",  txt: "50 × 30 mm",  w: 50,  h: 30 },
    { id: "r-58x40",  txt: "58 × 40 mm",  w: 58,  h: 40 },
    { id: "r-60x40",  txt: "60 × 40 mm",  w: 60,  h: 40 },
    { id: "r-100x50", txt: "100 × 50 mm", w: 100, h: 50 },
    { id: "custom",   txt: "Personalizada…" }
  ]
};

/** Hojas para imprimir etiquetas en impresora común (medidas en mm; css = tamaño para @page) */
const ETQ_HOJAS = {
  A4:     { nombre: "A4",     w: 210,   h: 297,   css: "A4" },
  CARTA:  { nombre: "Carta",  w: 215.9, h: 279.4, css: "letter" },
  OFICIO: { nombre: "Oficio", w: 216,   h: 340,   css: "216mm 340mm" }
};

/** Medidas de la hoja elegida, o null si es rollo */
function _hojaEtiquetas(cfg) {
  if (cfg.papel === "ROLLO") return null;
  if (cfg.papel === "HOJA") return { nombre: "personalizada", w: cfg.hojaW, h: cfg.hojaH, css: `${cfg.hojaW}mm ${cfg.hojaH}mm` };
  return ETQ_HOJAS[cfg.papel] || ETQ_HOJAS.A4;
}

/** Los tamaños de etiqueta de hoja sirven para cualquier hoja (A4, Carta, Oficio o personalizada) */
const _tamanosDePapel = (papel) => papel === "ROLLO" ? ETQ_TAMANOS.ROLLO : ETQ_TAMANOS.A4;

const ETQ_FABRICA = {
  papel: "A4", hojaAncho: 210, hojaAlto: 297, tamano: "a4-63x38", ancho: 63, alto: 38, separacion: 2, margen: 5, copias: 1, borde: "punteado",
  verNegocio: false, verNombre: true, verCategoria: false, verPrecio: true, verPrecioCaja: false,
  verBarras: true, verCodigo: true, verFecha: false, textoExtra: "",
  tamPrecio: "auto", tamNombre: "1", lineasNombre: "2", altoBarras: "0.26", alineacion: "center", decimales: "0",
  precioResaltado: false, negrita: true
};

const _ETQ_CAMPOS = {
  papel: "etqPapel", hojaAncho: "etqHojaAncho", hojaAlto: "etqHojaAlto", tamano: "etqTamano", ancho: "etqAncho", alto: "etqAlto", separacion: "etqSeparacion", margen: "etqMargen",
  copias: "etqCantidadCopias", borde: "etqBorde", verNegocio: "etqVerNegocio", verNombre: "etqVerNombre", verCategoria: "etqVerCategoria",
  verPrecio: "etqVerPrecio", verPrecioCaja: "etqVerPrecioCaja", verBarras: "etqVerBarras", verCodigo: "etqVerCodigo", verFecha: "etqVerFecha",
  textoExtra: "etqTextoExtra", tamPrecio: "etqTamPrecio", tamNombre: "etqTamNombre", lineasNombre: "etqLineasNombre",
  altoBarras: "etqAltoBarras", alineacion: "etqAlineacion", decimales: "etqDecimales", precioResaltado: "etqPrecioResaltado", negrita: "etqNegrita"
};

function _configEtiquetasGuardada() {
  try { return Object.assign({}, ETQ_FABRICA, JSON.parse(localStorage.getItem("vpos_etq_config") || "{}")); }
  catch (e) { return Object.assign({}, ETQ_FABRICA); }
}

function _llenarTamanosEtiqueta(papel, elegido) {
  const sel = document.getElementById("etqTamano");
  if (!sel) return;
  const lista = _tamanosDePapel(papel);
  sel.innerHTML = lista.map(t => `<option value="${t.id}">${escapeHtml(t.txt)}</option>`).join("");
  sel.value = lista.some(t => t.id === elegido) ? elegido : (papel === "ROLLO" ? "r-50x30" : "a4-63x38");
}

function _ponerConfigEtiquetas(cfg) {
  _llenarTamanosEtiqueta(cfg.papel, cfg.tamano);
  Object.keys(_ETQ_CAMPOS).forEach(k => {
    if (k === "tamano") return;
    const el = document.getElementById(_ETQ_CAMPOS[k]);
    if (!el) return;
    if (el.type === "checkbox") el.checked = !!cfg[k]; else el.value = cfg[k];
  });
}

function leerConfigEtiquetas() {
  const cfg = {};
  Object.keys(_ETQ_CAMPOS).forEach(k => {
    const el = document.getElementById(_ETQ_CAMPOS[k]);
    if (!el) { cfg[k] = ETQ_FABRICA[k]; return; }
    cfg[k] = el.type === "checkbox" ? el.checked : el.value;
  });
  // medidas: del tamaño elegido o personalizadas
  if (!ETQ_HOJAS[cfg.papel] && cfg.papel !== "HOJA" && cfg.papel !== "ROLLO") cfg.papel = "A4";
  const t = _tamanosDePapel(cfg.papel).find(x => x.id === cfg.tamano);
  const num = (v, min, max, def) => { const n = Number(v); return isFinite(n) && n > 0 ? Math.min(max, Math.max(min, n)) : def; };
  cfg.hojaW = num(cfg.hojaAncho, 50, 500, 210);
  cfg.hojaH = num(cfg.hojaAlto, 50, 1000, 297);
  cfg.w = t && t.w ? t.w : num(cfg.ancho, 15, 200, 63);
  cfg.h = t && t.h ? t.h : num(cfg.alto, 10, 280, 38);
  cfg.separacion = num(cfg.separacion, 0, 15, 3) || 0;
  cfg.margen = cfg.papel === "ROLLO" ? 0 : (Number(cfg.margen) >= 0 ? Math.min(25, Number(cfg.margen)) : 8);
  cfg.copias = Math.max(1, Math.min(50, Math.floor(Number(cfg.copias) || 1)));
  return cfg;
}

function _guardarConfigEtiquetas(cfg) {
  const guardar = {};
  Object.keys(ETQ_FABRICA).forEach(k => { guardar[k] = cfg[k]; });
  try { localStorage.setItem("vpos_etq_config", JSON.stringify(guardar)); } catch (e) {}
}

function restablecerConfigEtiquetas() {
  try { localStorage.removeItem("vpos_etq_config"); } catch (e) {}
  _ponerConfigEtiquetas(Object.assign({}, ETQ_FABRICA));
  actualizarPreviewEtiquetas();
}

/** Cuántas entran en la hoja elegida */
function _etiquetasPorHoja(cfg) {
  const hoja = _hojaEtiquetas(cfg);
  if (!hoja) return { cols: 1, filas: 1, total: 1 };
  const anchoUtil = hoja.w - 2 * cfg.margen, altoUtil = hoja.h - 2 * cfg.margen;
  const cols = Math.max(1, Math.floor((anchoUtil + cfg.separacion) / (cfg.w + cfg.separacion)));
  const filas = Math.max(1, Math.floor((altoUtil + cfg.separacion) / (cfg.h + cfg.separacion)));
  return { cols, filas, total: cols * filas };
}

function _precioEtiqueta(valor, cfg) {
  const d = Number(cfg.decimales) || 0;
  return "$" + Number(valor || 0).toLocaleString("es-AR", { minimumFractionDigits: d, maximumFractionDigits: d });
}

/** HTML de una etiqueta (medidas en mm; el código de barras se dibuja después en el <svg>). */
function _htmlEtiqueta(p, cfg, idSvg) {
  const H = cfg.h, W = cfg.w;
  const pt = (mm) => (mm * 2.835).toFixed(1) + "pt"; // 1 mm = 2,835 pt
  const chico = pt(Math.max(1.9, Math.min(H * 0.085, 3.4)));
  const lineas = Number(cfg.lineasNombre) || 2;
  const tamNombreMm = Math.max(2.1, Math.min(H * 0.08, W * 0.06, 4.2)) * (Number(cfg.tamNombre) || 1);
  let tamPrecioMm;
  if (cfg.tamPrecio === "auto") {
    let base = Math.min(H * 0.3, W * 0.16);
    const ocupan = (cfg.verBarras ? 1 : 0) + (cfg.verNombre ? 0.6 : 0);
    tamPrecioMm = base * (ocupan >= 1.5 ? 0.7 : ocupan >= 1 ? 0.8 : 1);
  } else {
    tamPrecioMm = Math.min(H * 0.3, W * 0.16) * Number(cfg.tamPrecio);
  }
  const partes = [];
  const fecha = new Date().toLocaleDateString("es-AR");
  if (cfg.verNegocio && configNegocioCache.nombre) partes.push(`<div class="etq-negocio" style="font-size:${chico}">${escapeHtml(configNegocioCache.nombre)}</div>`);
  if (cfg.verNombre) partes.push(`<div class="etq-nombre" style="font-size:${pt(tamNombreMm)};-webkit-line-clamp:${lineas};font-weight:${cfg.negrita ? 700 : 400}">${escapeHtml(p.PRODUCTO || "")}</div>`);
  if (cfg.verCategoria && p.CATEGORIA) partes.push(`<div class="etq-chico" style="font-size:${chico}">${escapeHtml(p.CATEGORIA)}</div>`);
  if (cfg.textoExtra) partes.push(`<div class="etq-extra" style="font-size:${pt(Math.max(2.2, Math.min(H * 0.09, 4)))}">${escapeHtml(cfg.textoExtra)}</div>`);
  if (cfg.verPrecio) partes.push(`<div class="etq-precio${cfg.precioResaltado ? " resaltado" : ""}" style="font-size:${pt(tamPrecioMm)}">${_precioEtiqueta(p.PRECIO, cfg)}</div>`);
  if (cfg.verPrecioCaja && Number(p.UNIDADES_POR_CAJA) > 0 && Number(p.PRECIO_CAJA) > 0) {
    partes.push(`<div class="etq-chico" style="font-size:${chico}">Caja x${escapeHtml(p.UNIDADES_POR_CAJA)}: <b>${_precioEtiqueta(p.PRECIO_CAJA, cfg)}</b></div>`);
  }
  // Alto del código: proporcional a la etiqueta según la opción elegida, nunca menos de 5 mm (más bajo cuesta leerlo)
  const altoBarrasMm = Math.max(5, H * Number(cfg.altoBarras || 0.26)).toFixed(1);
  if (cfg.verBarras) partes.push(`<svg class="etq-barras" id="${idSvg}" preserveAspectRatio="none" style="height:${altoBarrasMm}mm;flex-basis:${altoBarrasMm}mm"></svg>`);
  if (cfg.verCodigo) partes.push(`<div class="etq-chico etq-codigo" style="font-size:${chico}">${escapeHtml(p.CODIGO)}</div>`);
  if (cfg.verFecha) partes.push(`<div class="etq-chico" style="font-size:${chico}">${fecha}</div>`);
  const borde = cfg.borde === "ninguno" ? "none" : (cfg.borde === "solido" ? "0.2mm solid #555" : "0.25mm dashed #999");
  const pad = Math.max(0.8, Math.min(H, W) * 0.05).toFixed(1);
  return `<div class="etq" style="width:${W}mm;height:${H}mm;padding:${pad}mm;border:${borde};text-align:${cfg.alineacion};align-items:${cfg.alineacion === "left" ? "flex-start" : "center"}">${partes.join("")}</div>`;
}

function _dibujarBarrasEtiquetas(cont, items) {
  items.forEach(({ id, codigo }) => {
    const svg = cont.querySelector("#" + id);
    if (!svg) return;
    // JsBarcode reescribe el atributo style del <svg>: se guarda el alto
    // calculado (según el tamaño de etiqueta y la opción "Alto del código")
    // y se vuelve a poner después. Antes se perdía y el código quedaba
    // siempre en los 60 px de JsBarcode (≈16 mm) sin importar el tamaño.
    const alto = svg.style.height;
    try {
      // Margen blanco a los costados (10 barras finas de cada lado): los
      // lectores necesitan esa "zona libre" para leer el código sin fallar.
      JsBarcode(svg, String(codigo), { format: "CODE128", displayValue: false, width: 2, height: 60, margin: 0, marginLeft: 20, marginRight: 20 });
      svg.setAttribute("preserveAspectRatio", "none");
      svg.removeAttribute("width");
      svg.removeAttribute("height");
      svg.style.transform = "";
      if (alto) { svg.style.height = alto; svg.style.flexBasis = alto; }
    } catch (error) {
      // Código vacío o con caracteres no soportados: no frena el resto
      console.error("No se pudo generar el código de barras para", codigo, error);
      svg.outerHTML = `<div class="etq-chico" style="color:#b91c1c">(sin código de barras)</div>`;
    }
  });
}

/**
 * ¿Las barras de este código quedan demasiado finas para el ancho elegido?
 * Rollo (impresora de etiquetas de 203 ppp): menos de 0,19 mm no se dibuja
 * bien. Hoja A4 (láser/chorro de tinta, mucha más resolución): 0,1 mm.
 */
function _codigoMuyFinoParaEtiqueta(codigo, cfg) {
  try {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    JsBarcode(svg, String(codigo), { format: "CODE128", displayValue: false, width: 2, height: 10, margin: 0, marginLeft: 20, marginRight: 20 });
    const modulos = (parseFloat(svg.getAttribute("width")) || 0) / 2;
    const pad = Math.max(0.8, Math.min(cfg.h, cfg.w) * 0.05);
    return modulos > 0 && (cfg.w - 2 * pad) / modulos < (cfg.papel === "ROLLO" ? 0.19 : 0.1);
  } catch (e) { return false; }
}

/**
 * Si el contenido de una etiqueta no entra en su alto, achica primero el
 * código de barras (hasta 5 mm) y después todas las letras, de a poco,
 * hasta que entre. Así no se corta ninguna línea. Necesita que las
 * etiquetas estén en pantalla (para imprimir se miden fuera de la vista).
 */
function _ajustarEtiquetas(cont) {
  let sinLugar = 0;
  cont.querySelectorAll(".etq").forEach(e => {
    const desborda = () => e.scrollHeight > e.clientHeight + 0.5;
    for (let i = 0; i < 24 && desborda(); i++) {
      const svg = e.querySelector("svg.etq-barras");
      // El alto se lee del estilo (en mm): con la vista previa achicada,
      // el tamaño medido en pantalla no es el real.
      const altoMm = svg ? parseFloat(svg.style.height) || 0 : 0;
      if (svg && altoMm > 5) {
        const nuevo = Math.max(5, altoMm * 0.85);
        svg.style.height = nuevo.toFixed(2) + "mm";
        svg.style.flexBasis = nuevo.toFixed(2) + "mm";
        continue;
      }
      // Las letras: todas un 8% más chicas (mínimo 4,5 pt)
      let achico = false;
      e.querySelectorAll("[style*='font-size']").forEach(t => {
        const m = /([\d.]+)pt/.exec(t.style.fontSize || "");
        if (!m) return;
        const pt = Math.max(4.5, parseFloat(m[1]) * 0.92);
        if (pt < parseFloat(m[1])) { t.style.fontSize = pt.toFixed(1) + "pt"; achico = true; }
      });
      if (!achico) break; // ya no se puede achicar más: lo que sobre queda abajo
    }
    if (desborda()) sinLugar++;
  });
  return sinLugar;
}

function abrirModalEtiquetas() {
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }
  _ponerConfigEtiquetas(_configEtiquetasGuardada());
  const papel = document.getElementById("etqPapel");
  if (papel && !papel.__conCambio) {
    papel.addEventListener("change", () => {
      _llenarTamanosEtiqueta(papel.value, "");
      // En rollo las etiquetas ya vienen troqueladas: sin borde; en hoja, punteado para recortar
      const borde = document.getElementById("etqBorde");
      if (borde) borde.value = papel.value === "ROLLO" ? "ninguno" : "punteado";
      actualizarPreviewEtiquetas();
    });
    papel.__conCambio = true;
  }
  _etqHojaPreview = 0;
  document.getElementById("etiquetasModalBackdrop").classList.add("show");
  actualizarPreviewEtiquetas();
}

function cerrarModalEtiquetas() {
  document.getElementById("etiquetasModalBackdrop").classList.remove("show");
}

/* Vista previa: "hoja" muestra la hoja entera (A4, Carta, Oficio o personalizada) (o un tramo del rollo)
   achicada para que entre, con las etiquetas acomodadas igual que al
   imprimir; "real" muestra una etiqueta a tamaño real (aproximado). */
let _etqModoPreview = "hoja";
let _etqHojaPreview = 0;

function cambiarModoPreviewEtiquetas(modo) {
  _etqModoPreview = modo === "real" ? "real" : "hoja";
  document.querySelectorAll(".etq-modo button").forEach(b => b.classList.toggle("activo", b.dataset.modo === _etqModoPreview));
  actualizarPreviewEtiquetas();
}

function moverHojaPreviewEtiquetas(delta) {
  _etqHojaPreview = Math.max(0, _etqHojaPreview + delta);
  actualizarPreviewEtiquetas();
}

/** Lista de etiquetas a imprimir (producto repetido por cada copia), en orden. */
function _listaEtiquetas(cfg) {
  const codigos = obtenerCodigosSeleccionados();
  const productos = codigos.map(c => productosAdminGlobal.find(p => String(p.CODIGO) === String(c))).filter(Boolean);
  const lista = [];
  productos.forEach(p => { for (let i = 0; i < cfg.copias; i++) lista.push(p); });
  return { productos, lista };
}

function actualizarPreviewEtiquetas() {
  const cfg = leerConfigEtiquetas();
  const personalizada = cfg.tamano === "custom";
  const med = document.getElementById("etqMedidasPersonalizadas");
  if (med) med.style.display = personalizada ? "" : "none";
  const hojaOpc = document.getElementById("etqOpcionesHoja");
  if (hojaOpc) hojaOpc.style.display = cfg.papel === "ROLLO" ? "none" : "";
  const medHoja = document.getElementById("etqMedidasHoja");
  if (medHoja) medHoja.style.display = cfg.papel === "HOJA" ? "" : "none";
  const hoja = _hojaEtiquetas(cfg);

  const { productos, lista } = _listaEtiquetas(cfg);
  const total = lista.length;
  const ph = _etiquetasPorHoja(cfg);
  const hojas = hoja ? Math.max(1, Math.ceil(total / ph.total)) : 1;
  if (_etqHojaPreview > hojas - 1) _etqHojaPreview = hojas - 1;

  const info = document.getElementById("etiquetasPreviewInfo");
  if (info) {
    let txt = `${total} etiqueta${total === 1 ? "" : "s"} (${productos.length} producto${productos.length === 1 ? "" : "s"} × ${cfg.copias} copia${cfg.copias === 1 ? "" : "s"}) de ${cfg.w} × ${cfg.h} mm.`;
    if (hoja) {
      txt += ` Entran ${ph.total} por hoja ${hoja.nombre} (${ph.cols} × ${ph.filas}) → ${hojas} hoja${hojas === 1 ? "" : "s"}.`;
      if (cfg.w > hoja.w - 2 * cfg.margen || cfg.h > hoja.h - 2 * cfg.margen) txt += " ⚠️ La etiqueta no entra en la hoja: achicala o bajá el margen.";
    } else {
      txt += " Una etiqueta por cada salto del rollo: en la impresora elegí ese mismo tamaño de papel.";
    }
    info.textContent = txt;
  }

  const marco = document.getElementById("etqPreview");
  const pag = document.getElementById("etqPaginas");
  if (!marco) return;
  const MM = 96 / 25.4; // px por mm en pantalla
  const disponible = Math.max(140, marco.clientWidth - 16);
  const items = [];
  const etiqueta = (p, i) => { const id = "etqPrev" + i; items.push({ id, codigo: p.CODIGO }); return _htmlEtiqueta(p, cfg, id); };

  let anchoMm, altoMm, contenido;
  if (_etqModoPreview === "real") {
    anchoMm = cfg.w; altoMm = cfg.h;
    contenido = lista.length ? etiqueta(lista[0], 0) : "";
    if (pag) pag.style.display = "none";
  } else if (hoja) {
    anchoMm = hoja.w; altoMm = hoja.h;
    const desde = _etqHojaPreview * ph.total;
    const enHoja = lista.slice(desde, desde + ph.total);
    contenido = `<div class="etq-hoja-a4" style="width:${hoja.w}mm;height:${hoja.h}mm;padding:${cfg.margen}mm">
        <div class="etq-hoja" style="gap:${cfg.separacion}mm">${enHoja.map((p, i) => etiqueta(p, i)).join("")}</div>
      </div>`;
    if (pag) {
      pag.style.display = hojas > 1 ? "" : "none";
      const txt = document.getElementById("etqPaginasTexto");
      if (txt) txt.textContent = `Hoja ${_etqHojaPreview + 1} de ${hojas}`;
      const [ant, sig] = pag.querySelectorAll("button");
      if (ant) ant.disabled = _etqHojaPreview <= 0;
      if (sig) sig.disabled = _etqHojaPreview >= hojas - 1;
    }
  } else {
    // Rollo: tramo con las primeras etiquetas, una debajo de la otra
    const n = Math.min(lista.length, 4);
    const salto = 3; // mm entre etiquetas en el rollo (solo para mostrar)
    anchoMm = cfg.w + 8; altoMm = n * cfg.h + (n + 1) * salto;
    contenido = `<div class="etq-rollo-preview" style="width:${anchoMm}mm;height:${altoMm}mm;padding:${salto}mm 4mm;gap:${salto}mm">
        ${lista.slice(0, n).map((p, i) => etiqueta(p, i)).join("")}
      </div>`;
    if (pag) pag.style.display = "none";
  }

  // Escala: que entre a lo ancho (y en hoja, que no pase de un alto razonable)
  const altoMax = _etqModoPreview === "real" ? Infinity : Math.max(260, Math.min(560, window.innerHeight * 0.62));
  const f = Math.min(1, disponible / (anchoMm * MM), altoMax / (altoMm * MM));
  marco.innerHTML = `<div class="etq-escala" style="width:${(anchoMm * MM * f).toFixed(1)}px;height:${(altoMm * MM * f).toFixed(1)}px">
      <div class="etq-escala-in" style="transform:scale(${f.toFixed(4)})">${contenido}</div>
    </div>`;
  _dibujarBarrasEtiquetas(marco, items);
  // Los códigos se revisan en TODOS los productos elegidos, no solo en los que se ven
  const finos = cfg.verBarras ? [...new Set(productos.map(p => String(p.CODIGO)))].filter(c => _codigoMuyFinoParaEtiqueta(c, cfg)) : [];
  if (finos.length && info) info.textContent += ` ⚠️ ${finos.length === 1 ? "El código " + finos[0] + " es largo" : finos.length + " códigos son largos"} para este ancho: las barras quedan muy finas y puede costar leerlas. Usá una etiqueta más ancha.`;
  const sinLugar = _ajustarEtiquetas(marco);
  if (sinLugar && info) info.textContent += " ⚠️ No entra todo en la etiqueta (se corta abajo): sacá algún dato o elegí un tamaño más grande.";
  const nota = _etqModoPreview === "real"
    ? (f < 1 ? `Achicada al ${Math.round(f * 100)}% para que entre en pantalla.` : "A tamaño real (puede variar según la pantalla).")
    : (hoja ? `Hoja ${hoja.nombre} al ${Math.round(f * 100)}%.` : `Rollo al ${Math.round(f * 100)}%.`);
  marco.insertAdjacentHTML("beforeend", `<div class="etq-escala-nota">${nota}</div>`);
}

/**
 * Arma la grilla de impresión (A4) o una etiqueta por página (rollo),
 * dibuja los códigos de barras y abre el diálogo de impresión.
 */
function imprimirEtiquetas() {
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }
  const cfg = leerConfigEtiquetas();
  _guardarConfigEtiquetas(cfg);

  const productos = codigos.map(c => productosAdminGlobal.find(p => String(p.CODIGO) === String(c))).filter(Boolean);
  if (productos.length === 0) { toast("No se encontraron los productos seleccionados", "error"); return; }

  const printArea = document.getElementById("etiquetasPrintArea");
  const thermalFrame = document.getElementById("thermalPrintFrame");
  if (thermalFrame) thermalFrame.innerHTML = "";

  const items = [];
  let html = "", idx = 0;
  productos.forEach(p => {
    for (let copia = 0; copia < cfg.copias; copia++) {
      const id = "etqSvg" + (idx++);
      items.push({ id, codigo: p.CODIGO });
      html += _htmlEtiqueta(p, cfg, id);
    }
  });
  const clase = cfg.papel === "ROLLO" ? "etq-hoja etq-rollo" : "etq-hoja";
  printArea.innerHTML = `<div class="${clase}" style="gap:${cfg.separacion}mm">${html}</div>`;
  _dibujarBarrasEtiquetas(printArea, items);
  // Para medir, el área de impresión se muestra un instante fuera de la vista
  const estiloPrevio = printArea.getAttribute("style") || "";
  const hojaImp = _hojaEtiquetas(cfg);
  printArea.setAttribute("style", `display:block !important; position:absolute; left:-10000px; top:0; width:${hojaImp ? hojaImp.w - 2 * cfg.margen : cfg.w}mm; visibility:hidden;`);
  try { _ajustarEtiquetas(printArea); } finally { printArea.setAttribute("style", estiloPrevio); if (!estiloPrevio) printArea.removeAttribute("style"); }

  cerrarModalEtiquetas();
  if (cfg.papel === "ROLLO") {
    _setPrintPageSize("CUSTOM", `@media print { @page { size:${cfg.w}mm ${cfg.h}mm; margin:0; } }`);
  } else {
    _setPrintPageSize("CUSTOM", `@media print { @page { size:${hojaImp.css}; margin:${cfg.margen}mm; } }`);
  }
  setTimeout(() => { window.print(); }, 200);
}

/* ===================================================================
   QR DE PRECIO (consulta pública del precio actualizado, sin login)
=================================================================== */

function abrirModalQR() {
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }

  if (!configNegocioCache.urlCatalogo) {
    toast("Configurá primero la URL del catálogo en Configuración → Catálogo online", "error");
    return;
  }

  actualizarPreviewQR();
  document.getElementById("qrModalBackdrop").classList.add("show");
}

/* ---- Pantalla de cobro (web/cobro.html) para el celular que mira el cliente ---- */
function _urlPantallaCobro() {
  let base = (configNegocioCache && configNegocioCache.urlCatalogo) || "";
  if (!base && !(typeof _esAppEscritorio === "function" && _esAppEscritorio())) {
    base = new URL("../", window.location.href).href; // panel web: el catálogo está una carpeta arriba
  }
  base = String(base || "").replace(/\/(index\.html)?$/i, "").replace(/\/+$/, "");
  return base ? base + "/cobro.html" : "";
}

function abrirModalPantallaCobro() {
  const url = _urlPantallaCobro();
  if (!url) { toast("Cargá la URL del catálogo (Catálogo web → Catálogo online) para armar el enlace", "error"); return; }
  document.getElementById("pantallaCobroUrl").value = url;
  const cont = document.getElementById("pantallaCobroQR");
  cont.innerHTML = "";
  try { new QRCode(cont, { text: url, width: 200, height: 200, correctLevel: QRCode.CorrectLevel.M }); }
  catch (e) { cont.textContent = "No se pudo dibujar el QR"; }
  document.getElementById("pantallaCobroModalBackdrop").classList.add("show");
}

function cerrarModalPantallaCobro() {
  document.getElementById("pantallaCobroModalBackdrop").classList.remove("show");
}

function copiarUrlPantallaCobro() {
  const url = document.getElementById("pantallaCobroUrl").value;
  const ok = () => toast("Enlace copiado", "success");
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(url).then(ok).catch(() => { document.getElementById("pantallaCobroUrl").select(); document.execCommand("copy"); ok(); });
  else { document.getElementById("pantallaCobroUrl").select(); document.execCommand("copy"); ok(); }
}

function cerrarModalQR() {
  document.getElementById("qrModalBackdrop").classList.remove("show");
}

function actualizarPreviewQR() {
  const codigos = obtenerCodigosSeleccionados();
  const copias = Math.max(1, Number(document.getElementById("qrCantidadCopias").value || 1));
  const totalQR = codigos.length * copias;

  const info = document.getElementById("qrPreviewInfo");
  if (info) {
    info.textContent = `Se van a imprimir ${totalQR} código${totalQR === 1 ? "" : "s"} QR en total (${codigos.length} producto${codigos.length === 1 ? "" : "s"} × ${copias} copia${copias === 1 ? "" : "s"}).`;
  }
}

/** Builds the public "consultar precio" URL for a given product code */
function construirUrlPrecio(codigo) {
  const base = (configNegocioCache.urlCatalogo || "").replace(/\/+$/, "");
  return `${base}/precio.html?codigo=${encodeURIComponent(codigo)}`;
}

/**
 * Builds the printable QR grid and renders each QR with QRCode.js into
 * its placeholder <div> — then triggers the print dialog. Same
 * client-side approach as the barcode labels: no server round-trip.
 */
function imprimirQR() {
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }

  if (!configNegocioCache.urlCatalogo) {
    toast("Configurá primero la URL del catálogo en Configuración → Catálogo online", "error");
    return;
  }

  const copias = Math.max(1, Number(document.getElementById("qrCantidadCopias").value || 1));
  const columnas = document.getElementById("qrColumnas").value;

  const productos = codigos
    .map(codigo => productosAdminGlobal.find(p => String(p.CODIGO) === String(codigo)))
    .filter(Boolean);

  if (productos.length === 0) {
    toast("No se encontraron los productos seleccionados", "error");
    return;
  }

  const printArea = document.getElementById("etiquetasPrintArea");

  // Mismo cuidado que en las etiquetas de código de barras: se limpia
  // el otro contenedor de impresión para que el CSS :empty/:not(:empty)
  // elija el correcto sin ambigüedad.
  const thermalFrame = document.getElementById("thermalPrintFrame");
  if (thermalFrame) thermalFrame.innerHTML = "";

  let html = `<div class="etiquetas-grid cols-${columnas}">`;

  let idx = 0;
  productos.forEach(p => {
    for (let copia = 0; copia < copias; copia++) {
      html += `
        <div class="qr-item">
          <div class="qr-nombre">${escapeHtml(p.PRODUCTO)}</div>
          <div id="qrCanvas${idx}"></div>
          <div class="qr-leyenda">Escaneá para ver el precio</div>
        </div>`;
      idx++;
    }
  });

  html += `</div>`;
  printArea.innerHTML = html;

  // QRCode.js necesita que el contenedor ya esté en el DOM antes de
  // dibujar adentro — por eso se llena el HTML primero y se recorre después.
  idx = 0;
  productos.forEach(p => {
    const url = construirUrlPrecio(p.CODIGO);
    for (let copia = 0; copia < copias; copia++) {
      try {
        new QRCode(document.getElementById(`qrCanvas${idx}`), {
          text: url,
          width: 120,
          height: 120,
          colorDark: "#000000",
          colorLight: "#ffffff",
          correctLevel: QRCode.CorrectLevel.M
        });
      } catch (error) {
        // Un error puntual generando un QR no debe frenar la
        // impresión del resto.
        console.error("No se pudo generar el QR para", p.CODIGO, error);
      }
      idx++;
    }
  });

  cerrarModalQR();

  _setPrintPageSize("A4");
  setTimeout(() => {
    window.print();
  }, 200);
}

/* ===================================================================
   QR DE IDENTIFICACIÓN (offline, solo nombre del producto)
   A diferencia del QR de Precio, NO apunta a ninguna URL — codifica
   directamente el texto del nombre. Por eso cualquier lector de QR lo
   muestra al instante, sin necesidad de internet ni de que el backend
   esté disponible. Pensado para identificar cajas en el depósito.
=================================================================== */

function abrirModalQROffline() {
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }

  actualizarPreviewQROffline();
  document.getElementById("qrOfflineModalBackdrop").classList.add("show");
}

function cerrarModalQROffline() {
  document.getElementById("qrOfflineModalBackdrop").classList.remove("show");
}

function actualizarPreviewQROffline() {
  const codigos = obtenerCodigosSeleccionados();
  const copias = Math.max(1, Number(document.getElementById("qrOfflineCantidadCopias").value || 1));
  const totalQR = codigos.length * copias;

  const info = document.getElementById("qrOfflinePreviewInfo");
  if (info) {
    info.textContent = `Se van a imprimir ${totalQR} código${totalQR === 1 ? "" : "s"} QR en total (${codigos.length} producto${codigos.length === 1 ? "" : "s"} × ${copias} copia${copias === 1 ? "" : "s"}).`;
  }
}

/**
 * Builds the printable QR grid where each QR encodes just the plain
 * product name as text (not a URL) — works with any QR reader with
 * zero network involved, on either end.
 */
function imprimirQROffline() {
  const codigos = obtenerCodigosSeleccionados();
  if (codigos.length === 0) { toast("Seleccioná al menos un producto", "error"); return; }

  const copias = Math.max(1, Number(document.getElementById("qrOfflineCantidadCopias").value || 1));
  const columnas = document.getElementById("qrOfflineColumnas").value;

  const productos = codigos
    .map(codigo => productosAdminGlobal.find(p => String(p.CODIGO) === String(codigo)))
    .filter(Boolean);

  if (productos.length === 0) {
    toast("No se encontraron los productos seleccionados", "error");
    return;
  }

  const printArea = document.getElementById("etiquetasPrintArea");

  const thermalFrame = document.getElementById("thermalPrintFrame");
  if (thermalFrame) thermalFrame.innerHTML = "";

  let html = `<div class="etiquetas-grid cols-${columnas}">`;

  let idx = 0;
  productos.forEach(p => {
    for (let copia = 0; copia < copias; copia++) {
      html += `
        <div class="qr-item">
          <div class="qr-nombre">${escapeHtml(p.PRODUCTO)}</div>
          <div id="qrOfflineCanvas${idx}"></div>
          <div class="qr-leyenda">Funciona sin internet</div>
        </div>`;
      idx++;
    }
  });

  html += `</div>`;
  printArea.innerHTML = html;

  idx = 0;
  productos.forEach(p => {
    for (let copia = 0; copia < copias; copia++) {
      try {
        new QRCode(document.getElementById(`qrOfflineCanvas${idx}`), {
          text: String(p.PRODUCTO || ""),
          width: 120,
          height: 120,
          colorDark: "#000000",
          colorLight: "#ffffff",
          correctLevel: QRCode.CorrectLevel.M
        });
      } catch (error) {
        console.error("No se pudo generar el QR de identificación para", p.CODIGO, error);
      }
      idx++;
    }
  });

  cerrarModalQROffline();

  _setPrintPageSize("A4");
  setTimeout(() => {
    window.print();
  }, 200);
}


/* =========================================================
   SISTEMA DE LICENCIAS
   Usa window.veekpos (preload.js) → main.js → license-client.js → Apps Script
========================================================= */

let estadoLicenciaActual = { activada: false, modoLimitado: true };

// Servidor de licencias de VeekPOS — fijo: el usuario solo carga correo y clave.
// (main.js tiene la misma URL; se manda también desde acá por si la app
// todavía tiene un main.js viejo que lee la URL guardada en la PC.)
const URL_SERVIDOR_LICENCIAS = "https://script.google.com/macros/s/AKfycbygrGvC481UoNPEfCGgY29TW-CU9L0n4lNO0qKLA5IBVI8eSTkjKZyghIlyLbnB4xRi/exec";

let _relojRevisionLicencia = null;

async function aplicarEstadoLicencia() {
  if (typeof window.veekpos === "undefined" || !window.veekpos.obtenerEstadoLicencia) return;
  const antes = JSON.stringify(estadoLicenciaActual);
  try {
    estadoLicenciaActual = await window.veekpos.obtenerEstadoLicencia();
  } catch (error) { console.error("Error licencia:", error); return; }

  // Mientras esté bloqueada (ej. caja cliente que abrió antes que la caja
  // servidor), se vuelve a revisar cada 5 s: apenas aparece el servidor, el
  // aviso rojo se va y se puede vender sin reiniciar la app. Una vez que
  // queda bien no se revisa sola de nuevo (un corte momentáneo de la red no
  // bloquea las ventas a mitad del día).
  clearTimeout(_relojRevisionLicencia);
  if (estadoLicenciaActual.activada && estadoLicenciaActual.modoLimitado) {
    _relojRevisionLicencia = setTimeout(aplicarEstadoLicencia, 5000);
  }
  const cambio = antes !== JSON.stringify(estadoLicenciaActual);

  const pantallaActivacion = document.getElementById("licenseScreenBackdrop");
  const banner = document.getElementById("licenseLimitedBanner");
  const bannerTexto = document.getElementById("licenseLimitedBannerText");

  if (!estadoLicenciaActual.activada) {
    if (pantallaActivacion) pantallaActivacion.classList.add("show");
    if (banner) banner.style.display = "none";
    return;
  }

  if (pantallaActivacion) pantallaActivacion.classList.remove("show");

  if (estadoLicenciaActual.modoLimitado) {
    if (banner) {
      banner.style.display = "flex";
      banner.className = "license-limited-banner";
      if (bannerTexto) bannerTexto.textContent = (estadoLicenciaActual.motivo || "Licencia no vigente") + " — no se pueden registrar ventas ni editar productos.";
    }
  } else {
    _mostrarAvisoVencimientoLicencia(banner, bannerTexto);
  }
  actualizarBloqueosPorLicencia();
  const configVisible = document.getElementById("configuracion")?.style.display === "block";
  if (configVisible && cambio) mostrarEstadoLicenciaEnConfig();
}

/**
 * Avisos de vencimiento (la licencia sigue funcionando):
 *  - 15 días antes: franja ámbar (se puede posponer hasta mañana) y un aviso al llegar a ese punto.
 *  - 3 días antes y durante los 5 días de gracia: franja naranja siempre visible y aviso al llegar.
 */
function _mostrarAvisoVencimientoLicencia(banner, bannerTexto) {
  const e = estadoLicenciaActual || {};
  const aviso = e.aviso;
  if (!banner) return;
  if (!aviso) { banner.style.display = "none"; return; }
  const fecha = (() => { const d = _fechaLicencia(e.fechaVencimiento); return d ? d.toLocaleDateString("es-AR", { day: "numeric", month: "long" }) : ""; })();
  const dias = Number(e.diasParaVencer);
  let texto, hito;
  if (aviso.nivel === "gracia") {
    const q = aviso.diasGraciaRestantes;
    texto = `La licencia venció el ${fecha}. Estás en los días de gracia: ${q === 0 ? "hoy es el último día" : "quedan " + (q + 1) + " días (contando hoy)"}. Después no se van a poder registrar ventas. Renovala cuanto antes.`;
    hito = "gracia";
  } else {
    const cuando = dias <= 0 ? "hoy (último día)" : dias === 1 ? `mañana, el ${fecha}` : `en ${dias} días, el ${fecha}`;
    texto = `La licencia vence ${cuando}. Renovala para no interrumpir las ventas.`;
    hito = aviso.nivel === "urgente" ? "3" : "15";
  }
  const claveHoy = "vpos_lic_posponer_" + _hoyISO();
  let pospuesto = false;
  try { pospuesto = aviso.nivel === "previo" && localStorage.getItem(claveHoy) === "1"; } catch (err) {}
  banner.className = "license-limited-banner aviso-" + aviso.nivel;
  banner.style.display = pospuesto ? "none" : "flex";
  if (bannerTexto) bannerTexto.textContent = texto;
  let btn = banner.querySelector(".lic-posponer");
  if (aviso.nivel === "previo") {
    if (!btn) {
      btn = document.createElement("button");
      btn.type = "button"; btn.className = "lic-posponer"; btn.textContent = "Recordármelo mañana";
      btn.onclick = () => { try { localStorage.setItem(claveHoy, "1"); } catch (err) {} banner.style.display = "none"; };
      banner.appendChild(btn);
    }
  } else if (btn) btn.remove();
  // Aviso al cruzar cada punto (15 días, 3 días, gracia): una sola vez por vencimiento
  const claveHito = "vpos_lic_hito_" + hito + "_" + (e.fechaVencimiento || "");
  let yaAvisado = false;
  try { yaAvisado = localStorage.getItem(claveHito) === "1"; } catch (err) {}
  if (!yaAvisado) {
    try { localStorage.setItem(claveHito, "1"); } catch (err) {}
    setTimeout(() => confirmarAccion(texto, () => {}, "Aviso de licencia"), 800);
  }
}

function actualizarBloqueosPorLicencia() {
  const bloqueado = estadoLicenciaActual.modoLimitado;
  const btnFinalizar = document.getElementById("btnFinalizarVenta");
  if (btnFinalizar) btnFinalizar.disabled = bloqueado || ticketPOS.length === 0;
  document.querySelectorAll(".btn-accion-producto").forEach(btn => { btn.disabled = bloqueado; });
}

async function activarLicenciaForm() {
  const email = document.getElementById("licenseScreenEmail").value.trim();
  const pin = document.getElementById("licenseScreenPin").value.trim();
  const errorBox = document.getElementById("licenseScreenError");
  if (errorBox) errorBox.style.display = "none";
  if (!email || !pin) {
    if (errorBox) { errorBox.style.display = "block"; errorBox.textContent = "Completá el correo y la clave."; }
    return;
  }
  const btn = document.getElementById("licenseScreenBtn");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true; btn.innerHTML = "Activando...";
  try {
    await window.veekpos.fijarUrlServidorLicencia?.(URL_SERVIDOR_LICENCIAS);
    const resultado = await window.veekpos.activarLicencia(email, pin);
    if (!resultado.success) {
      if (errorBox) { errorBox.style.display = "block"; errorBox.textContent = resultado.message || "No se pudo activar la licencia"; }
      return;
    }
    toast("Licencia activada correctamente", "success");
    const cancelar = document.getElementById("licenseScreenCancelar");
    if (cancelar) cancelar.style.display = "none";
    await aplicarEstadoLicencia();
    mostrarEstadoLicenciaEnConfig();
  } catch (error) {
    if (errorBox) { errorBox.style.display = "block"; errorBox.textContent = "Error: " + String(error.message || error); }
  } finally { btn.disabled = false; btn.innerHTML = textoOriginal; }
}

async function validarLicenciaAhoraBtn() {
  const btn = document.getElementById("btnVerificarLicencia");
  const orig = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Verificando..."; }
  try {
    await window.veekpos?.fijarUrlServidorLicencia?.(URL_SERVIDOR_LICENCIAS);
    estadoLicenciaActual = await window.veekpos?.validarLicenciaAhora?.() || estadoLicenciaActual;
    await aplicarEstadoLicencia();
    mostrarEstadoLicenciaEnConfig();
    toast(estadoLicenciaActual.modoLimitado ? "No vigente: " + (estadoLicenciaActual.motivo||"") : "Licencia válida", estadoLicenciaActual.modoLimitado ? "error" : "success");
  } catch(e) { toast("Error al verificar la licencia", "error"); }
  finally { if (btn) { btn.disabled = false; btn.innerHTML = orig; } }
}

/** Fecha "2027-07-24" (o ISO completa) → Date local, sin correrse un día por zona horaria */
function _fechaLicencia(valor) {
  if (!valor) return null;
  const m = String(valor).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(valor);
  return isNaN(d.getTime()) ? null : d;
}

/** "24 de julio de 2027 (295 días)" */
function _textoVencimientoLicencia(valor) {
  const d = _fechaLicencia(valor);
  if (!d) return valor ? escapeHtml(String(valor)) : "—";
  // Mismo criterio que VeekIQ: días que quedan contando el día del vencimiento
  const finDelDia = new Date(d); finDelDia.setHours(23, 59, 59, 999);
  const dias = Math.ceil((finDelDia - Date.now()) / 86400000);
  const fecha = d.toLocaleDateString("es-AR", { day: "numeric", month: "long", year: "numeric" });
  const vencidaHace = 1 - dias;
  const resto = dias > 1 ? `${dias} días` : dias === 1 ? "último día" : `vencida hace ${vencidaHace} día${vencidaHace === 1 ? "" : "s"}`;
  const color = dias <= 0 ? "var(--red-500)" : dias <= 15 ? "#b45309" : "inherit";
  return `${escapeHtml(fecha)} <span style="color:${color};">(${resto})</span>`;
}

/** "3/10/26, 10:11 a. m." */
function _textoUltimaVerificacionLicencia(valor) {
  const d = _fechaLicencia(valor);
  if (!d) return "—";
  return escapeHtml(d.toLocaleString("es-AR", { day: "numeric", month: "numeric", year: "2-digit", hour: "numeric", minute: "2-digit", hour12: true }));
}

async function mostrarEstadoLicenciaEnConfig() {
  const box = document.getElementById("licenciaEstadoBox");
  const badge = document.getElementById("licenciaBadge");
  const idBox = document.getElementById("licenciaIdInstalacion");
  const acciones = document.getElementById("licenciaAcciones");
  if (!box) return;

  // Datos frescos (la última verificación pudo cambiar en segundo plano)
  try { if (window.veekpos?.obtenerEstadoLicencia) estadoLicenciaActual = await window.veekpos.obtenerEstadoLicencia(); } catch (e) {}
  const e = estadoLicenciaActual || {};

  const fijarBadge = (texto, clase) => { if (badge) { badge.textContent = texto; badge.className = "licencia-badge " + clase; } };

  if (typeof window.veekpos === "undefined") {
    fijarBadge("web", "");
    box.innerHTML = `<div class="text-muted" style="font-size:12.5px;">La licencia se administra desde la app de escritorio.</div>`;
    if (acciones) acciones.style.display = "none";
    if (idBox) idBox.textContent = "";
    return;
  }
  if (acciones) acciones.style.display = "flex";

  // Caja cliente de la red local: no tiene licencia propia
  if (e.activada && !e.email && e.fechaVencimiento === undefined) {
    fijarBadge(e.modoLimitado ? "sin conexión" : "activa", e.modoLimitado ? "bad" : "ok");
    box.innerHTML = `<div style="font-size:13px;">Esta caja usa la licencia de la <strong>caja servidor</strong>.</div>
      ${e.modoLimitado ? `<div style="font-size:12.5px; color:var(--red-500); margin-top:4px;">⚠️ ${escapeHtml(e.motivo || "")}</div>` : ""}`;
    if (acciones) acciones.style.display = "none";
    if (idBox) idBox.textContent = "";
    return;
  }

  if (!e.activada) {
    fijarBadge("sin activar", "bad");
    box.innerHTML = `<div style="font-size:13px; color:var(--red-500);">⚠️ Esta PC no tiene una licencia activada.</div>`;
  } else {
    if (!e.modoLimitado) fijarBadge("activa", "ok");
    else if (e.rechazada) fijarBadge(/venc/i.test(e.motivo || "") ? "vencida" : "suspendida", "bad");
    else fijarBadge("sin verificar", "warn");

    box.innerHTML = `
      <div class="licencia-datos">
        <span class="lbl">Correo</span><span>${escapeHtml(e.email || "—")}</span>
        <span class="lbl">Vence</span><span>${_textoVencimientoLicencia(e.fechaVencimiento)}</span>
        <span class="lbl">Última verificación</span><span>${_textoUltimaVerificacionLicencia(e.ultimaValidacionOk)}</span>
      </div>
      ${e.modoLimitado ? `<div style="font-size:12.5px; color:var(--red-500); margin-top:8px;">⚠️ ${escapeHtml(e.motivo || "Licencia no vigente")}</div>` : ""}`;
  }

  const btnQuitar = document.getElementById("btnQuitarLicencia");
  // Con un main.js/preload.js viejos no existe quitarLicenciaDeEstaPC: no se muestra el botón
  if (btnQuitar) btnQuitar.style.display = (e.activada && typeof window.veekpos.quitarLicenciaDeEstaPC === "function") ? "inline-flex" : "none";
  const btnCambiar = document.getElementById("btnCambiarLicencia");
  if (btnCambiar) btnCambiar.textContent = e.activada ? "Cambiar licencia…" : "Activar licencia…";
  if (idBox) idBox.textContent = e.idInstalacion ? "ID de instalación: VEEKPOS-" + String(e.idInstalacion).toUpperCase() : "";
}

/** Abre la pantalla de activación para cargar otro correo y clave, con opción de cancelar */
async function cambiarLicencia() {
  const pantalla = document.getElementById("licenseScreenBackdrop");
  if (!pantalla) return;
  const email = document.getElementById("licenseScreenEmail");
  if (email) email.value = "";
  const pin = document.getElementById("licenseScreenPin");
  if (pin) pin.value = "";
  const err = document.getElementById("licenseScreenError");
  if (err) err.style.display = "none";
  const cancelar = document.getElementById("licenseScreenCancelar");
  if (cancelar) cancelar.style.display = estadoLicenciaActual.activada ? "block" : "none";
  pantalla.classList.add("show");
  setTimeout(() => email?.focus(), 50);
}

function cancelarCambioLicencia() {
  if (!estadoLicenciaActual.activada) return; // sin licencia no se puede salir de la pantalla
  document.getElementById("licenseScreenBackdrop")?.classList.remove("show");
  const cancelar = document.getElementById("licenseScreenCancelar");
  if (cancelar) cancelar.style.display = "none";
}

function quitarLicenciaDeEstaPC() {
  confirmarAccion(
    "¿Quitar la licencia de esta PC? VeekPOS va a pedir de nuevo el correo y la clave para seguir usándose acá. Tus datos (ventas, productos, clientes) no se tocan.",
    async () => {
      try {
        const r = await window.veekpos?.quitarLicenciaDeEstaPC?.();
        if (!r || !r.success) { toast((r && r.message) || "No se pudo quitar la licencia (actualizá main.js y preload.js)", "error"); return; }
        toast("Licencia quitada de esta PC", "success");
        const cancelar = document.getElementById("licenseScreenCancelar");
        if (cancelar) cancelar.style.display = "none";
        await aplicarEstadoLicencia();
        mostrarEstadoLicenciaEnConfig();
      } catch (e) {
        toast("No se pudo quitar la licencia", "error");
      }
    },
    "Quitar licencia de esta PC"
  );
}

/* =========================================================
   INSTALADOR / GENERADOR DE CLIENTE
   Genera Code.gs, config.json y ZIPs listos para instalar
========================================================= */

function instObtenerDatos() {
  return {
    empresa:       (document.getElementById("instEmpresa")?.value || "").trim(),
    spreadsheetId: (document.getElementById("instSpreadsheetId")?.value || "").trim(),
    apiUrl:        (document.getElementById("instApiUrl")?.value || "").trim(),
    whatsapp:      (document.getElementById("instWhatsapp")?.value || "").trim(),
    direccion:     (document.getElementById("instDireccion")?.value || "").trim(),
    moneda:        (document.getElementById("instMoneda")?.value || "ARS"),
    fecha:         _hoyISO(),
    version:       "1.0"
  };
}

function instValidar() {
  const d = instObtenerDatos();
  if (!d.empresa) { toast("Ingresá el nombre de la empresa", "error"); return null; }
  if (!d.spreadsheetId) { toast("Ingresá el ID de la planilla de Google Sheets", "error"); return null; }
  return d;
}

/** Descarga un archivo de texto desde el browser */
function descargarArchivo(nombre, contenido, tipo = "text/plain") {
  const blob = new Blob([contenido], { type: tipo });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  a.click();
  URL.revokeObjectURL(url);
}

/** Paso 2: Genera y descarga el Code.gs con los datos del cliente */
async function generarCodeGs() {
  const d = instValidar();
  if (!d) return;

  const status = document.getElementById("instStatus");
  if (status) status.innerHTML = `<div class="text-muted">⏳ Cargando template...</div>`;

  try {
    // Cargar el template desde el servidor
    const res = await fetch("../code.gs.template?_=" + Date.now(), { cache: "no-store" });
    if (!res.ok) throw new Error("No se encontró el template code.gs.template");
    let template = await res.text();

    // El template real (el mismo que corre en producción para el
    // negocio actual, así que no puede tener un {{marcador}} literal
    // sin romper ese uso directo) tiene el ID de la planilla escrito
    // como una constante de verdad: const SPREADSHEET_ID = "...".
    // Se busca esa línea puntual con una expresión regular y se le
    // cambia el valor — antes se buscaba un {{SPREADSHEET_ID}} que
    // nunca existió en el archivo, así que el reemplazo no encontraba
    // nada y el .gs se descargaba con el ID de OTRO negocio adentro,
    // sin ningún aviso de que eso había pasado.
    const regexSpreadsheetId = /const\s+SPREADSHEET_ID\s*=\s*"[^"]*"/;
    if (!regexSpreadsheetId.test(template)) {
      throw new Error("No se encontró la línea 'const SPREADSHEET_ID = ...' en el template — no se puede generar el archivo con seguridad.");
    }
    template = template.replace(regexSpreadsheetId, `const SPREADSHEET_ID = "${d.spreadsheetId}"`);

    const nombreArchivo = `Code_${d.empresa.replace(/\s+/g, "_")}.gs`;
    descargarArchivo(nombreArchivo, template, "text/plain");

    if (status) status.innerHTML = `<div class="text-success">✓ ${nombreArchivo} descargado — seguí las instrucciones del Paso 2</div>`;
    toast("Code.gs generado y descargado", "success");

  } catch (err) {
    console.error(err);
    toast("Error al generar Code.gs: " + err.message, "error");
    if (status) status.innerHTML = `<div class="text-danger">⚠️ ${err.message}</div>`;
  }
}

/**
 * Arma el config.json del cliente — el ÚNICO archivo con datos propios
 * de cada negocio en el sitio (ver config.json y scripts/aplicar-cliente.js).
 */
function instArmarConfigCliente(d) {
  let sitioUrl = "";
  try {
    if (/^https?:/.test(location.protocol)) sitioUrl = location.href.replace(/admin\/[^]*$/, "");
  } catch (e) {}
  return {
    _LEEME: "Único archivo con los datos de este cliente. Editalo y subilo: el resto del sitio se adapta solo (scripts/aplicar-cliente.js). Nombre, colores, banner y teléfonos visibles se cambian desde el panel → Configuración.",
    empresa:       d.empresa,
    nombreCorto:   d.empresa.split(/\s+/)[0] || d.empresa,
    descripcion:   "Catálogo online de " + d.empresa + " con envíos a todo el país.",
    sitioUrl:      sitioUrl,
    apiUrl:        d.apiUrl || "PEGAR_URL_API_AQUI",
    apiUrlLectura: "",
    whatsapp:      d.whatsapp,
    direccion:     d.direccion,
    moneda:        d.moneda,
    pais:          "Argentina",
    colorTema:     "#0b1633",
    googleVerificacion: "",
    version:       d.version,
    generadoEn:    d.fecha
  };
}

/** Paso 3a: Genera y descarga config.json */
function generarConfigJson() {
  const d = instValidar();
  if (!d) return;

  const config = instArmarConfigCliente(d);

  const json = JSON.stringify(config, null, 2);
  descargarArchivo("config.json", json, "application/json");
  toast("config.json descargado — colocalo en la raíz del sitio (al lado de index.html)", "success");
}

/** Paso 3b: Genera ZIP del catálogo web completo con config del cliente */
async function generarCatalogoZip() {
  const d = instValidar();
  if (!d) return;
  if (!d.apiUrl) { toast("Ingresá la API URL antes de generar el catálogo", "error"); return; }

  const status = document.getElementById("instStatus");
  if (status) status.innerHTML = `<div class="text-muted">⏳ Generando catálogo web...</div>`;

  try {
    // Cargar JSZip dinámicamente si no está disponible
    if (typeof JSZip === "undefined") {
      await new Promise((res, rej) => {
        const s = document.createElement("script");
        s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
        s.onload = res; s.onerror = () => rej(new Error("No se pudo cargar JSZip"));
        document.head.appendChild(s);
      });
    }

    const zip = new JSZip();

    // Insertar config.json del cliente en el ZIP
    const config = instArmarConfigCliente(d);
    zip.file("config.json", JSON.stringify(config, null, 2));

    // Lista de archivos del catálogo (viven en veekPOS/catalogo/)
    const archivosTexto = [
      "index.html",
      "js/app.js",
      "css/style.css",
      "service-worker.js",
      "manifest.json",
      "manifest-catalogo.json",
      "robots.txt",
      "lector.html",
      "precio.html",
    ];
    const archivosBinarios = ["icon-192.png", "icon-512.png", "favicon.ico"];

    const base = "../"; // el catálogo vive en la raíz del sitio (admin/ está adentro)
    let cargados = 0;

    // Cargar archivos de texto
    await Promise.all(archivosTexto.map(async (archivo) => {
      try {
        const res = await fetch(base + archivo + "?_=" + Date.now(), { cache: "no-store" });
        if (!res.ok) return;
        zip.file(archivo, await res.text());
        cargados++;
        if (status) status.innerHTML = `<div class="text-muted">⏳ Cargando archivos... (${cargados})</div>`;
      } catch(e) { console.warn("Omitido:", archivo); }
    }));

    // Cargar binarios
    await Promise.all(archivosBinarios.map(async (img) => {
      try {
        const res = await fetch(base + img + "?_=" + Date.now(), { cache: "no-store" });
        if (res.ok) { zip.file(img, await res.arrayBuffer()); cargados++; }
      } catch(e) {}
    }));

    // Generar y descargar ZIP
    if (status) status.innerHTML = `<div class="text-muted">⏳ Comprimiendo...</div>`;
    const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Catalogo_${d.empresa.replace(/\s+/g, "_")}_${d.fecha}.zip`;
    a.click();
    URL.revokeObjectURL(url);

    if (status) status.innerHTML = `
      <div class="alert alert-success mt-2" style="font-size:13px;">
        ✓ <strong>Catálogo descargado (${cargados} archivos)</strong><br>
        Subí el contenido del ZIP a tu hosting (Netlify, GitHub Pages, etc.).<br>
        El <code>config.json</code> ya tiene la API URL y los datos del cliente.
      </div>`;
    toast("Catálogo web generado", "success");

  } catch (err) {
    console.error(err);
    toast("Error al generar el catálogo: " + err.message, "error");
    if (status) status.innerHTML = `<div class="text-danger mt-2">⚠️ ${err.message}</div>`;
  }
}

/** Paso 4: Verifica que la API URL responda correctamente */
async function verificarConexionInstalador() {
  const apiUrl = (document.getElementById("instApiUrl")?.value || "").trim();
  const status = document.getElementById("instConexionStatus");

  if (!apiUrl) {
    toast("Ingresá la API URL primero", "error");
    return;
  }

  if (status) status.innerHTML = `<span class="text-muted">⏳ Probando conexión y preparando las hojas...</span>`;

  try {
    // Un solo paso: probar que la API responde Y, de paso, crear en la
    // planilla todas las hojas que el sistema necesita (si ya
    // estaban creadas, no se tocan ni se duplican).
    const res = await fetch(apiUrl + "?action=inicializarHojas", { signal: AbortSignal.timeout(20000) });
    const data = await res.json();

    if (data.success) {
      if (status) status.innerHTML = `<span class="text-success">Conexión exitosa — ${data.message}</span>`;
      toast("Conexión verificada y hojas listas", "success");
    } else {
      if (status) status.innerHTML = `<span class="text-warning">⚠️ La API respondió pero algo falló: ${data.message || "revisá el Code.gs publicado"}</span>`;
    }
  } catch (err) {
    if (status) status.innerHTML = `<span class="text-danger">No se pudo conectar: ${err.message}</span>`;
    toast("Error de conexión: " + err.message, "error");
  }
}

/** Guarda la API URL desde el formulario de configuración (tarjeta Electron) */
async function guardarApiUrlDesdeConfig() {
  const input = document.getElementById("conexionNegocioUrl");
  const url = (input ? input.value : "").trim();
  if (!url) { toast("Ingresá la URL de la API", "error"); return; }

  // Guardar con el método disponible (veekpos o posOffline)
  const bridge = window.veekpos || window.posOffline;
  if (bridge) {
    if (bridge.fijarConfigLocal) await bridge.fijarConfigLocal("api_url", url);
    else if (bridge.guardarApiUrl) await bridge.guardarApiUrl(url);
    API_URL = url;
    const actual = document.getElementById("conexionNegocioUrlActual");
    if (actual) actual.textContent = "Conectada a: " + url;
    toast("URL guardada correctamente", "success");
  }
}

/* =========================================================
   SISTEMA DE LICENCIAS
   Verifica la licencia al iniciar el panel.
   Si no hay conexión, permite continuar por 7 días offline.
========================================================= */

const LICENCIAS_URL_KEY = "veekpos_licencias_url";
const LICENCIA_KEY      = "veekpos_licencia";
const LICENCIA_TS_KEY   = "veekpos_licencia_ts";
const LICENCIA_OFFLINE_DIAS = 7;

async function verificarLicenciaAlIniciar() {
  // Si no hay URL de servidor de licencias configurada, saltar verificación
  const servidorUrl = localStorage.getItem(LICENCIAS_URL_KEY);
  if (!servidorUrl) return; // sin servidor configurado → sin restricción

  const codigoGuardado = localStorage.getItem(LICENCIA_KEY);
  const tsGuardado     = Number(localStorage.getItem(LICENCIA_TS_KEY) || 0);
  const diasDesdeUltima = (Date.now() - tsGuardado) / 86400000;

  // Si verificó recientemente (< 24h), no verificar de nuevo
  if (codigoGuardado && diasDesdeUltima < 1) return;

  if (!codigoGuardado) {
    // Primera vez — pedir código de licencia
    mostrarPantallaActivacion();
    return;
  }

  // Verificar en línea
  try {
    const res = await fetch(`${servidorUrl}?action=validarLicencia&licencia=${codigoGuardado}`,
      { signal: AbortSignal.timeout(8000) });
    const data = await res.json();

    if (!data.valida) {
      mostrarPantallaActivacion(data.mensaje || "Licencia inválida o vencida");
      return;
    }

    // Guardar resultado
    localStorage.setItem(LICENCIA_TS_KEY, Date.now());
    console.log(`✓ Licencia válida — ${data.empresa} (${data.plan})`);

  } catch(err) {
    // Sin conexión al servidor de licencias
    if (diasDesdeUltima > LICENCIA_OFFLINE_DIAS) {
      mostrarPantallaActivacion(
        `Sin conexión al servidor de licencias por ${Math.round(diasDesdeUltima)} días. ` +
        `Conectate a internet para continuar.`
      );
    }
    // Si es menos de 7 días offline, continuar normalmente
  }
}

function mostrarPantallaActivacion(mensaje = "") {
  // Crear overlay de activación si no existe
  if (document.getElementById("licenciaOverlay")) return;

  const overlay = document.createElement("div");
  overlay.id = "licenciaOverlay";
  overlay.style.cssText = `
    position:fixed; inset:0; z-index:99999;
    background:#0b1633; display:flex; flex-direction:column;
    align-items:center; justify-content:center; gap:16px;
  `;
  overlay.innerHTML = `
    <div style="color:#fff; font-size:28px;">🔑</div>
    <div style="color:#fff; font-size:20px; font-weight:700;">Activación de licencia</div>
    ${mensaje ? `<div style="color:#f87171; font-size:13px; max-width:340px; text-align:center;">${mensaje}</div>` : ""}
    <div style="background:#fff; padding:24px; border-radius:12px; width:340px;">
      <label style="font-size:13px; font-weight:600; display:block; margin-bottom:8px;">Código de licencia</label>
      <input type="text" id="inputCodigoLicencia" class="form-control"
        placeholder="Ej: VPK-ABC123" style="text-transform:uppercase; letter-spacing:.1em;">
      <div id="licenciaError" style="color:#ef4444; font-size:12px; margin-top:6px; display:none;"></div>
      <button onclick="activarLicencia()" class="btn btn-success w-100 mt-3">
        Activar
      </button>
    </div>
    <div style="color:#64748b; font-size:12px;">VeekPOS — Sistema de Punto de Venta</div>
  `;
  document.body.appendChild(overlay);
  setTimeout(() => document.getElementById("inputCodigoLicencia")?.focus(), 100);
}

async function activarLicencia() {
  const input = document.getElementById("inputCodigoLicencia");
  const errorEl = document.getElementById("licenciaError");
  const codigo = (input?.value || "").trim().toUpperCase();

  if (!codigo) {
    if (errorEl) { errorEl.textContent = "Ingresá el código de licencia"; errorEl.style.display = "block"; }
    return;
  }

  const servidorUrl = localStorage.getItem(LICENCIAS_URL_KEY);
  if (!servidorUrl) {
    // Sin servidor configurado → aceptar cualquier código localmente
    localStorage.setItem(LICENCIA_KEY, codigo);
    localStorage.setItem(LICENCIA_TS_KEY, Date.now());
    document.getElementById("licenciaOverlay")?.remove();
    return;
  }

  const btn = document.querySelector("#licenciaOverlay button");
  if (btn) { btn.disabled = true; btn.textContent = "Verificando..."; }

  try {
    const res = await fetch(`${servidorUrl}?action=validarLicencia&licencia=${codigo}`,
      { signal: AbortSignal.timeout(10000) });
    const data = await res.json();

    if (!data.valida) {
      if (errorEl) { errorEl.textContent = data.mensaje || "Licencia inválida"; errorEl.style.display = "block"; }
      if (btn) { btn.disabled = false; btn.textContent = "Activar"; }
      return;
    }

    localStorage.setItem(LICENCIA_KEY, codigo);
    localStorage.setItem(LICENCIA_TS_KEY, Date.now());
    document.getElementById("licenciaOverlay")?.remove();
    toast(`✓ Licencia activada — ${data.empresa} (${data.plan})`, "success");

  } catch(err) {
    if (errorEl) { errorEl.textContent = "Error de conexión: " + err.message; errorEl.style.display = "block"; }
    if (btn) { btn.disabled = false; btn.textContent = "Activar"; }
  }
}


/* =========================================================
   MERCADO PAGO — Cobro con QR (integrado desde pos-offline)
========================================================= */

/**
 * Muestra u oculta el campo "Nombre de esta caja" según el rol elegido
 * — solo el servidor necesita un nombre propio (es el que se anuncia
 * por la red para que los clientes lo encuentren, ver local-discovery.js
 * → iniciarAnuncioServidor). Una caja cliente no anuncia nada, así que
 * no necesita nombre.
 */
function actualizarVisibilidadCampoNombreCaja(rol) {
  const campo = document.getElementById("campoNombreCaja");
  if (campo) campo.style.display = rol === "servidor" ? "block" : "none";
  const srv = document.getElementById("bloqueCodigoServidor"), cli = document.getElementById("bloqueCodigoCliente");
  if (srv) srv.style.display = rol === "servidor" ? "block" : "none";
  if (cli) cli.style.display = rol === "cliente" ? "block" : "none";
}

/** Muestra el código (en el servidor) o si ya está cargado (en un cliente) */
async function cargarCodigoRed() {
  const bridge = window.veekpos || window.posOffline;
  if (!bridge || typeof bridge.obtenerCodigoRed !== "function") return;
  try {
    const r = await bridge.obtenerCodigoRed();
    if (r.rol === "servidor") {
      const el = document.getElementById("codigoRedServidor");
      if (el) el.textContent = r.codigo || "";
    } else {
      const ayuda = document.getElementById("cfgCodigoRedAyuda");
      if (ayuda) ayuda.textContent = r.configurado
        ? "Ya hay un código guardado en esta caja (" + r.parcial + "). Escribí uno solo si cambió en la caja servidor."
        : "Todavía no se cargó: sin el código esta caja no puede conectarse a la caja servidor.";
    }
  } catch (e) {}
}

/** Caja servidor: genera un código nuevo (las demás cajas tienen que cargarlo de nuevo) */
function regenerarCodigoRed() {
  confirmarAccion("Se va a generar un código nuevo. Las demás cajas van a dejar de conectarse hasta que les escribas el código nuevo. ¿Seguir?", async () => {
    const bridge = window.veekpos || window.posOffline;
    try {
      const r = await bridge.regenerarCodigoRed();
      if (!r.success) { toast(r.message || "No se pudo generar el código", "error"); return; }
      document.getElementById("codigoRedServidor").textContent = r.codigo;
      toast("Código nuevo generado: escribilo en cada caja cliente", "success");
    } catch (e) { toast("No se pudo generar el código", "error"); }
  }, "Código de emparejamiento");
}

/** Carga el estado actual de la config de red al abrir la tarjeta — rol guardado, nombre de caja, y a quién está viendo (servidor propio o servidor encontrado) */
async function cargarConfigRedForm(soloEstado) {
  const bridge = window.veekpos || window.posOffline;
  if (!bridge) return;
  if (!soloEstado) seguirEstadoRed(false);

  const estadoEl = document.getElementById("estadoRedLocal");

  try {
    const { rol, nombreCaja } = await bridge.obtenerRolRed();
    if (!soloEstado) {
      document.getElementById("cfgRolRed").value = rol || "";
      document.getElementById("cfgNombreCaja").value = nombreCaja || "";
      actualizarVisibilidadCampoNombreCaja(rol || "");
      cargarCodigoRed();
    }

    if (!estadoEl) return;

    if (rol === "servidor") {
      const { ip, puertoHttp } = await bridge.obtenerIpLocal();
      estadoEl.innerHTML = ip
        ? `🟢 Esta caja es el <b>servidor</b>. Las demás cajas (y la pantalla del cliente, si usás una) se conectan solas por red — si necesitás la dirección a mano: <code>${ip}:${puertoHttp}</code>`
        : `🟢 Esta caja es el <b>servidor</b>, pero no se pudo detectar la IP de red local (¿está conectada a wifi/cable?).`;

    } else if (rol === "cliente") {
      const estado = await bridge.obtenerEstadoConexionRed();
      if (estado.conectado && estado.sinCodigo) {
        estadoEl.textContent = "🔴 Se encontró la caja servidor, pero falta el código de emparejamiento: escribilo arriba y guardá.";
      } else if (estado.conectado && estado.bloqueado) {
        estadoEl.textContent = "🔴 La caja servidor bloqueó esta caja unos minutos por intentos con un código incorrecto. Revisá el código y esperá.";
      } else if (estado.conectado && !estado.codigoOk) {
        estadoEl.textContent = "🔴 La caja servidor rechazó el código de emparejamiento: revisá que sea el que muestra la caja servidor.";
      } else if (estado.conectado) {
        estadoEl.textContent = "🟢 Conectada al servidor correctamente.";
      } else if (estado.servidorConocido) {
        estadoEl.textContent = "🟡 Se encontró un servidor en la red, pero no responde ahora mismo.";
      } else {
        estadoEl.textContent = "🟡 Buscando el servidor en la red local… (la caja servidor tiene que estar prendida, con VeekPOS abierto y en la misma red)";
      }

    } else {
      const mpConfigurado = await bridge.mercadoPagoDisponible?.().catch(() => false);
      const { ip, puertoHttp } = await bridge.obtenerIpLocal().catch(() => ({ ip: null }));

      if (mpConfigurado) {
        estadoEl.innerHTML = ip
          ? `⚪ Modo independiente — esta caja no comparte datos con otras por red local. La pantalla del cliente para QR sí está activa (hay Mercado Pago configurado): <code>${ip}:${puertoHttp}</code>`
          : `⚪ Modo independiente. La pantalla del cliente debería estar activa, pero no se pudo detectar la IP de red local (¿está conectada a wifi/cable?).`;
      } else {
        estadoEl.innerHTML = ip
          ? `⚪ Modo independiente — esta caja no comparte datos con otras por red local. La pantalla del cliente para QR no está activa todavía porque no hay Mercado Pago configurado (IP de esta caja, por si la necesitás más adelante: <code>${ip}:${puertoHttp}</code>).`
          : `⚪ Modo independiente — esta caja no comparte datos con otras por red local. La pantalla del cliente para QR no está activa (no hay Mercado Pago configurado).`;
      }
    }
  } catch (error) {
    if (estadoEl) estadoEl.textContent = "No se pudo cargar el estado de la red.";
  }
}

/**
 * Estado de la conexión en vivo: después de guardar se revisa cada 1 s
 * (hasta 20 s o hasta que conecte), y mientras la tarjeta de red esté a la
 * vista, cada 4 s. Antes se miraba una sola vez justo al guardar — cuando
 * todavía no había dado tiempo a encontrar el servidor — y quedaba
 * "Buscando…" hasta salir y volver a entrar a Configuración.
 */
let _relojEstadoRed = null;
function seguirEstadoRed(rapido) {
  clearTimeout(_relojEstadoRed);
  const inicio = Date.now();
  const ciclo = async () => {
    const card = document.getElementById("cardMultiCajaRed");
    const visible = card && card.offsetParent !== null && document.getElementById("configuracion")?.style.display === "block";
    if (!visible) return;
    await cargarConfigRedForm(true);
    const txt = document.getElementById("estadoRedLocal")?.textContent || "";
    const conectado = /Conectada al servidor/.test(txt);
    const enRapido = rapido && !conectado && Date.now() - inicio < 20000;
    _relojEstadoRed = setTimeout(ciclo, enRapido ? 1000 : 4000);
  };
  _relojEstadoRed = setTimeout(ciclo, rapido ? 700 : 4000);
}

/** Guarda el rol de red elegido (botón "Guardar configuración de red") */
async function guardarConfigRed() {
  const bridge = window.veekpos || window.posOffline;
  if (!bridge) return;

  const rol = document.getElementById("cfgRolRed").value;
  const nombreCaja = document.getElementById("cfgNombreCaja").value.trim();

  if (rol === "servidor" && !nombreCaja) {
    toast("Ponele un nombre a esta caja (ej: Caja 1)", "error");
    return;
  }

  const codigo = (document.getElementById("cfgCodigoRed")?.value || "").trim();
  try {
    await bridge.fijarRolRed(rol, nombreCaja);
    if (rol === "cliente" && codigo && typeof bridge.fijarCodigoRed === "function") {
      const r = await bridge.fijarCodigoRed(codigo);
      if (!r.success) { toast(r.message || "El código no es válido", "error"); return; }
      document.getElementById("cfgCodigoRed").value = "";
    }
    toast("Configuración de red guardada", "success");
    await cargarConfigRedForm();
    seguirEstadoRed(true);
  } catch (error) {
    toast("No se pudo guardar la configuración de red", "error");
  }
}


async function mostrarEstadoMercadoPagoEnConfig() {
  const box = document.getElementById("mpEstadoBox");
  const formularioWrap = document.getElementById("mpFormularioWrap");
  if (!box) return;

  try {
    const cfg = await window.veekpos.obtenerConfigMercadoPago();
    const btnQuitar = document.getElementById("btnQuitarMercadoPago");

    if (cfg.esCliente) {
      // El Access Token se configura una sola vez, en la caja
      // servidor — esta caja no tiene nada propio para editar, solo
      // muestra si el servidor ya lo tiene listo.
      if (formularioWrap) formularioWrap.style.display = "none";
      box.innerHTML = cfg.configurado
        ? `<div style="color:var(--green-600); font-weight:600;">✓ Configurado en la caja servidor</div><div class="text-muted" style="font-size:12.5px; margin-top:2px;">El cobro con QR se gestiona desde el servidor — no hace falta configurar nada en esta caja.</div>`
        : `<span class="text-muted">No configurado en la caja servidor. La configuración de Mercado Pago se hace una sola vez, ahí.</span>`;
      return;
    }

    if (formularioWrap) formularioWrap.style.display = "block";

    if (cfg.configurado) {
      box.innerHTML = `<div style="color:var(--green-600); font-weight:600;">✓ Configurado</div><div class="text-muted" style="font-size:12.5px; margin-top:2px;">Token guardado: ${escapeHtml(cfg.tokenParcial)}</div>`;
      if (btnQuitar) btnQuitar.style.display = "inline-block";
    } else if (cfg.hayToken) {
      box.innerHTML = `<span style="color:var(--red-500);">⚠️ Hay un token guardado pero falta terminar de configurar (probá guardarlo de nuevo).</span>`;
      if (btnQuitar) btnQuitar.style.display = "inline-block";
    } else {
      box.innerHTML = `<span class="text-muted">No configurado — "Transferencia" funciona en modo manual.</span>`;
      if (btnQuitar) btnQuitar.style.display = "none";
    }

    // Pre-completa los campos de dirección si ya se habían guardado antes (no pisa lo que el usuario esté escribiendo ahora)
    if (cfg.direccion) {
      const campoCalle = document.getElementById("cfgMpCalle");
      const campoNumero = document.getElementById("cfgMpNumero");
      const campoCiudad = document.getElementById("cfgMpCiudad");
      const campoProvincia = document.getElementById("cfgMpProvincia");
      if (campoCalle && !campoCalle.value) campoCalle.value = cfg.direccion.calle || "";
      if (campoNumero && !campoNumero.value) campoNumero.value = cfg.direccion.numero || "";
      if (campoCiudad && !campoCiudad.value) campoCiudad.value = cfg.direccion.ciudad || "";
      if (campoProvincia && !campoProvincia.value) campoProvincia.value = cfg.direccion.provincia || "";
    }

  } catch (error) {
    console.error("Error al consultar estado de Mercado Pago:", error);
    box.innerHTML = `<span class="text-muted">No se pudo consultar el estado.</span>`;
  }
}

async function guardarConfigMercadoPagoForm() {
  const input = document.getElementById("cfgMpAccessToken");
  const accessToken = input.value.trim();

  const direccion = {
    calle: document.getElementById("cfgMpCalle").value.trim(),
    numero: document.getElementById("cfgMpNumero").value.trim(),
    ciudad: document.getElementById("cfgMpCiudad").value.trim(),
    provincia: document.getElementById("cfgMpProvincia").value.trim()
  };

  if (!accessToken) { toast("Ingresá el Access Token", "error"); return; }
  if (!direccion.calle || !direccion.numero || !direccion.ciudad || !direccion.provincia) {
    toast("Completá la dirección del negocio — Mercado Pago la exige para crear la tienda", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarMercadoPago");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Configurando...";

  try {
    const nombreNegocio = (configNegocioCache && configNegocioCache.nombre) || "VeekPOS";
    const resultado = await window.veekpos.guardarYConfigurarMercadoPago(accessToken, nombreNegocio, direccion);

    if (!resultado.success) {
      toast(resultado.message || "No se pudo configurar Mercado Pago", "error");
      return;
    }

    toast("Mercado Pago configurado correctamente" + (resultado.nombre ? " (" + resultado.nombre + ")" : ""), "success");
    input.value = "";
    mostrarEstadoMercadoPagoEnConfig();

  } catch (error) {
    console.error("Error al configurar Mercado Pago:", error);
    toast("Error al configurar Mercado Pago", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

async function probarCredencialesMercadoPagoBtn() {
  const input = document.getElementById("cfgMpAccessToken");
  const accessToken = input.value.trim();

  if (!accessToken) { toast("Ingresá un Access Token para probar", "error"); return; }

  try {
    const resultado = await window.veekpos.probarCredencialesMercadoPago(accessToken);
    if (resultado.success) {
      toast("Conexión correcta" + (resultado.nombre ? " — cuenta: " + resultado.nombre : ""), "success");
    } else {
      toast(resultado.message || "No se pudo conectar con ese token", "error");
    }
  } catch (error) {
    console.error("Error al probar credenciales de Mercado Pago:", error);
    toast("Error al probar la conexión", "error");
  }
}

async function quitarConfigMercadoPago() {
  if (!await confirmarAccion("¿Quitar la configuración de Mercado Pago? \"Transferencia\" va a volver a funcionar en modo manual.", { textoBoton: "Quitar" })) return;

  try {
    await window.veekpos.borrarConfigMercadoPago();
    toast("Configuración de Mercado Pago eliminada", "success");
    mostrarEstadoMercadoPagoEnConfig();
  } catch (error) {
    console.error("Error al quitar configuración de Mercado Pago:", error);
    toast("Error al quitar la configuración", "error");
  }
}

/* ===================== RED LOCAL MULTI-CAJA (Configuración) ===================== */

/**
 * DIAGNÓSTICO — abre el mismo modal de cobro con QR, pero con una imagen
 * falsa (un SVG armado acá mismo, sin pedirle nada a Mercado Pago) y sin
 * arrancar el polling real. Sirve para probar el comportamiento del modal
 * (por ejemplo, si el botón "Cancelar" responde) sin necesitar tener
 * Mercado Pago configurado ni depender de la red — así se puede descartar
 * rápido si un problema es del modal en sí o de la integración con MP.
 *
 * A propósito NO toca mpReferenciaActual ni arranca mpPollingIntervalId:
 * sin una referencia real no hay nada que consultar, así que el polling
 * ni se inicia — "Cancelar" acá prueba exactamente el mismo camino de
 * limpieza (detenerPollingMercadoPago, sacar la clase "show", resetear
 * los flags) que un cobro real, sin arriesgar nada del carrito actual.
 */
function abrirModalQRDePrueba() {
  const svgFalso = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(`
    <svg xmlns="http://www.w3.org/2000/svg" width="240" height="240">
      <rect width="240" height="240" fill="#eef2ff"/>
      <text x="120" y="112" text-anchor="middle" font-family="monospace" font-size="16" fill="#1e293b">QR DE PRUEBA</text>
      <text x="120" y="136" text-anchor="middle" font-family="monospace" font-size="12" fill="#64748b">(no es un cobro real)</text>
    </svg>
  `)}`;

  document.getElementById("mpQrMontoLabel").textContent = "$0 (prueba)";
  document.getElementById("mpQrError").style.display = "none";
  document.getElementById("mpQrEsperando").style.display = "block";
  document.getElementById("mpQrImagen").src = svgFalso;
  document.getElementById("mpQrBackdrop").classList.add("show");

  // Empuja el mismo QR falso a la pantalla/tablet del cliente (si hay
  // una conectada) — así también se puede probar esa parte sin
  // necesitar Mercado Pago configurado. Si no hay bridge (versión web)
  // o la tablet no está conectada, no pasa nada, es un no-op seguro.
  const bridge = window.veekpos || window.posOffline;
  if (typeof window.enviarTokenAlProcesoPrincipal === "function") window.enviarTokenAlProcesoPrincipal();
  if (bridge && typeof bridge.mostrarQrDePruebaEnPantalla === "function") {
    bridge.mostrarQrDePruebaEnPantalla(svgFalso, 0).catch(() => {});
  }

  toast("Modal de prueba — este QR no cobra nada de verdad", "info");
}

async function iniciarCobroMercadoPago(total) {
  if (typeof window.enviarTokenAlProcesoPrincipal === "function") window.enviarTokenAlProcesoPrincipal(); // para avisar el QR a la pantalla del celular
  document.getElementById("mpQrMontoLabel").textContent = "$" + Number(total).toLocaleString("es-AR");
  document.getElementById("mpQrError").style.display = "none";
  document.getElementById("mpQrEsperando").style.display = "block";
  document.getElementById("mpQrImagen").src = "";
  document.getElementById("mpQrBackdrop").classList.add("show");

  const subtotal = ticketPOS.reduce((acc, item) => acc + (item.PRECIO * item.cantidad), 0);
  const itemsSnapshot = [...ticketPOS]; // snapshot antes de limpiar
  const recibido = Number(document.getElementById("mfvRecibido")?.value) || 0;
  mpVentaEnCurso = { subtotal, total, itemsSnapshot, recibido };

  try {
    const referencia = "veekpos" + Date.now();
    const bridge = window.veekpos || window.posOffline;
    const resultado = await bridge.crearCobroQR(total, referencia);

    if (!resultado.success) {
      mostrarErrorCobroMercadoPago(resultado.message || "No se pudo generar el QR de cobro");
      return;
    }

    mpReferenciaActual = referencia;
    document.getElementById("mpQrImagen").src = resultado.qrImagenDataUrl;
    mpPollingIniciadoEn = Date.now();
    mpProcesandoConfirmacion = false;
    mpPollingIntervalId = setInterval(consultarCobroMercadoPagoPolling, 3000);

  } catch (error) {
    console.error("Error al generar cobro con Mercado Pago:", error);
    mostrarErrorCobroMercadoPago("Error al generar el QR de cobro");
  }
}


/* =========================================================
   MERCADO PAGO — Polling y modal QR
========================================================= */

let mpPollingIntervalId = null;
let mpReferenciaActual  = null;
let mpVentaEnCurso      = null; // { subtotal, total, itemsSnapshot, recibido }

// Dos problemas reales que tenía este polling, arreglados acá:
//
// 1) RACE CONDITION que podía duplicar la venta: setInterval no espera a
//    que termine la vuelta anterior. Si una consulta a Mercado Pago tarda
//    más de 3 segundos (red lenta), podían quedar DOS llamadas a
//    consultarCobroQR "en el aire" al mismo tiempo; si ambas contestaban
//    "pagada: true" (algo totalmente posible ya que consultan lo mismo),
//    las dos ejecutaban el bloque de éxito completo → la venta se
//    guardaba DOS VECES en el backend. `mpProcesandoConfirmacion` corta
//    esto: se chequea antes Y después del await, así que la segunda
//    llamada que ya estaba en vuelo cuando la primera confirmó el pago,
//    al volver de su propio await, ve el flag ya en true y no hace nada.
//
// 2) Sin límite de tiempo: si el cliente nunca paga (se arrepiente, no
//    tiene señal, lo que sea), el polling seguía cada 3 segundos PARA
//    SIEMPRE, sin avisarle nada al cajero, hasta que alguien tocara
//    "Cancelar" a mano. `mpPollingIniciadoEn` + el timeout de abajo lo
//    cortan solos con un aviso, después de 10 minutos sin confirmación.
let mpProcesandoConfirmacion = false;
let mpPollingIniciadoEn = null;
const MP_POLLING_TIMEOUT_MS = 10 * 60 * 1000; // 10 minutos

async function consultarCobroMercadoPagoPolling() {
  if (!mpReferenciaActual || mpProcesandoConfirmacion) return;

  if (mpPollingIniciadoEn && (Date.now() - mpPollingIniciadoEn) > MP_POLLING_TIMEOUT_MS) {
    mostrarErrorCobroMercadoPago("Pasaron 10 minutos sin confirmar el pago. Si el cliente todavía va a pagar, generá el QR de nuevo.");
    return;
  }

  try {
    const bridge = window.veekpos || window.posOffline;
    const resultado = await bridge.consultarCobroQR(mpReferenciaActual);

    // Otra vuelta del intervalo ya se adelantó y confirmó el pago mientras
    // esta esperaba su propia respuesta — no hacer nada, ya se procesó.
    if (mpProcesandoConfirmacion) return;

    if (resultado.success && resultado.pagada) {
      mpProcesandoConfirmacion = true; // a partir de acá, ninguna otra vuelta en vuelo va a duplicar esto (ver comentario arriba)
      detenerPollingMercadoPago();
      document.getElementById("mpQrBackdrop").classList.remove("show");
      toast("Pago de Mercado Pago confirmado ✓", "success");

      if (mpVentaEnCurso) {
        const { subtotal, total, itemsSnapshot, recibido } = mpVentaEnCurso;
        const etiquetaDescuento = obtenerEtiquetaDescuentoPOS(subtotal);

        // Guardar — recién ahora que el pago está confirmado. Mismo
        // camino que una venta normal (local-primero en escritorio; en la
        // web POST + cola). Antes, si este POST fallaba, la venta cobrada
        // por QR se perdía: no se encolaba en ningún lado.
        let ventaId = "VEN-" + Date.now().toString().slice(-6);
        const clienteVentaIdMP = "CVL-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);
        const guardadoMP = await registrarVentaCobrada({
          clienteVentaId: clienteVentaIdMP,
          ventaIdTemp: ventaId,
          total, subtotal,
          descuento: subtotal - total,
          formaPago: "TRANSFERENCIA",
          vendedor: sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("usuarioLogueado") || "",
          observaciones: etiquetaDescuento ? (ajusteModoPOS === "RECARGO" ? "Recargo: " : "Descuento: ") + etiquetaDescuento : "",
          carrito: itemsSnapshot
        });
        if (guardadoMP && guardadoMP.ventaId) ventaId = guardadoMP.ventaId;

        // Mostrar recibo con el ID real
        ultimaVentaImprimible = {
          ventaId, items: itemsSnapshot, total, subtotal,
          descuento: subtotal - total,
          descuentoEtiqueta: etiquetaDescuento,
          formaPago: "TRANSFERENCIA", fecha: new Date(), recibido: 0
        };

        mostrarRecibo(ventaId, itemsSnapshot, total, subtotal, subtotal - total, recibido);
        if (window.ArcaUI) ArcaUI.trasVenta({ ref: clienteVentaIdMP, items: itemsSnapshot, total, fecha: new Date() });

        // Limpiar ticket y actualizar UI
        ticketPOS = [];
        limpiarPineadoPOS();
        resetearDescuentoPOS();
        const inputRec = document.getElementById("inputRecibido");
        const cambioEl = document.getElementById("cambioValor");
        if (inputRec) inputRec.value = "";
        if (cambioEl) { cambioEl.textContent = "—"; cambioEl.classList.remove("negativo"); }
        renderTicketPOS();
        ultimoCodigoAgregadoPOS = null;
        posTileFocusIdx = -1;

        // Stock optimista
        itemsSnapshot.forEach(item => {
          const p = productosPOS.find(x => String(x.CODIGO) === String(item.CODIGO));
          if (p) p.STOCK = Math.max(0, Number(p.STOCK) - item.cantidad);
        });
        renderPosGrid();

        setTimeout(() => { cargarMetricas(); invalidarCache("ventasPOS"); }, 500);
        mpVentaEnCurso = null;
      }
    }
  } catch (error) {
    console.warn("Error consultando estado del cobro MP (se reintenta):", error);
  }
}

function mostrarErrorCobroMercadoPago(mensaje) {
  detenerPollingMercadoPago();
  mpProcesandoConfirmacion = false;
  mpPollingIniciadoEn = null;
  const el = document.getElementById("mpQrEsperando");
  const err = document.getElementById("mpQrError");
  const txt = document.getElementById("mpQrErrorTexto");
  if (el) el.style.display = "none";
  if (err) err.style.display = "block";
  if (txt) txt.textContent = mensaje;
}

function detenerPollingMercadoPago() {
  if (mpPollingIntervalId) { clearInterval(mpPollingIntervalId); mpPollingIntervalId = null; }
}

function cancelarCobroMercadoPago() {
  detenerPollingMercadoPago();
  mpProcesandoConfirmacion = false;
  mpPollingIniciadoEn = null;
  const backdrop = document.getElementById("mpQrBackdrop");
  if (backdrop) backdrop.classList.remove("show");
  mpReferenciaActual = null;
  mpVentaEnCurso = null;

  // Avisa también a la pantalla/tablet del cliente (si hay una conectada)
  // que este cobro se canceló, para que no se quede mostrando un QR
  // muerto — ver local-server.js → limpiarQrActivo. No pasa nada si no
  // hay ninguna pantalla conectada, ni si esto corre en la versión web.
  const bridge = window.veekpos || window.posOffline;
  if (bridge && typeof bridge.cancelarCobroQR === "function") {
    bridge.cancelarCobroQR().catch(() => {});
  }

  toast("Cobro cancelado — el ticket sigue abierto", "info");
}


/* ===================================================================
   INGRESO DE PRODUCTOS — recepción de mercadería con proveedor
   Escanea/tipea un código: si el producto existe recupera nombre y
   categoría (desde productosAdminGlobal, ya en memoria); si no
   existe, permite crearlo. En ambos casos deja registro con fecha,
   precio y proveedor, y permite exportar el historial a PDF.
=================================================================== */

let ingresosProductosGlobal = [];

/* ---------------------------------------------------------------------
   TIPO DE CAMBIO AUTOMÁTICO (sección Ingreso de Productos)
   Trae la cotización del día (valor de VENTA) de DolarApi y completa solo los campos "Tipo de cambio"
   de la sección: el del formulario manual (ipTipoCambio) y el de la revisión por foto/PDF (ocrTc).
   Todo sigue siendo editable: en la barra de arriba (pisa los de abajo) o en cada campo (solo ese).
   Si no hay internet o la API falla, se usa el último valor guardado y se puede cargar a mano.
--------------------------------------------------------------------- */
const TC_API = "https://dolarapi.com/v1/dolares/";
const TC_LS_KEY = "vpos_tc_ingreso";
const TC_CAMPOS = ["ipTipoCambio", "ocrTc"];
const TC_TIPOS = { oficial: "Oficial", blue: "Blue", bolsa: "MEP (Bolsa)", mayorista: "Mayorista", tarjeta: "Tarjeta" };
const TC_TIPO_LS = "vpos_dolar_tipo";
let ingresoTC = { valor: 0, tipo: "oficial", manual: false, dia: "", fecha: "", ultimaAuto: 0, fallo: false };

function tcCargarLocal() {
  try {
    const o = JSON.parse(localStorage.getItem(TC_LS_KEY) || "null");
    if (o && typeof o === "object") Object.assign(ingresoTC, o);
  } catch (e) { /* sin almacenamiento: se trabaja en memoria */ }
}
function tcGuardarLocal() {
  try { localStorage.setItem(TC_LS_KEY, JSON.stringify(ingresoTC)); } catch (e) { /* no pasa nada */ }
}

/** Refleja el estado en la barra de arriba */
function tcRender() {
  const nom = document.getElementById("ingresoTcTipoNombre");
  const inp = document.getElementById("ingresoTcValor");
  const est = document.getElementById("ingresoTcEstado");
  if (!nom || !inp || !est) return;
  nom.textContent = ingresoTC.tipo === "deposito" ? "Depósito" : (TC_TIPOS[ingresoTC.tipo] || TC_TIPOS.oficial);
  if (document.activeElement !== inp) inp.value = ingresoTC.valor || "";
  const cuando = ingresoTC.fecha && !isNaN(new Date(ingresoTC.fecha))
    ? new Date(ingresoTC.fecha).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : "";
  if (ingresoTC.tipo === "deposito" && !ingresoTC.manual) {
    est.textContent = "Dólar del depósito (fijado a mano)" + (cuando ? " · " + cuando : "");
  } else if (ingresoTC.manual) {
    est.innerHTML = '<span class="badge bg-warning text-dark">editado a mano</span> hoy — "Traer cotización" vuelve al valor automático';
  } else if (ingresoTC.fallo) {
    est.innerHTML = '<span class="badge bg-danger">sin conexión</span> ' + (ingresoTC.valor ? "usando el último valor guardado" + (cuando ? " (" + cuando + ")" : "") + " — podés editarlo" : "cargalo a mano");
  } else if (ingresoTC.valor) {
    est.textContent = "Automático · DolarApi" + (cuando ? " · actualizado " + cuando : "");
  } else {
    est.textContent = "Todavía sin cotización";
  }
}

/** ¿Está el campo en uso? (solo se completan los campos cuyo costo está en dólares) */
function tcCampoActivo(id) {
  const sel = document.getElementById(id === "ocrTc" ? "ocrMoneda" : "ipMoneda");
  return !!sel && sel.value === "USD";
}

/** Completa un campo con el valor actual si está vacío o si nunca lo tocaste a mano */
function tcPrellenarCampo(id) {
  const el = document.getElementById(id);
  if (!el || !(ingresoTC.valor > 0) || !tcCampoActivo(id)) return;
  if (!el.value || !el.dataset.tcLocal) el.value = ingresoTC.valor;
}

/** Propaga el valor a los campos de la sección. forzar=true también pisa los que editaste a mano en cada producto */
function tcAplicarACampos(forzar) {
  if (!(ingresoTC.valor > 0)) return;
  TC_CAMPOS.forEach(id => {
    const el = document.getElementById(id);
    if (!el || !tcCampoActivo(id)) return;
    if (forzar) delete el.dataset.tcLocal;
    if (!el.dataset.tcLocal) el.value = ingresoTC.valor;
  });
  if (document.getElementById("ocrTabla")) ocrRenderTabla();   // actualiza los equivalentes en pesos de la revisión
}

/** Trae la cotización de venta del tipo elegido. silencioso=true: sin avisos ni pisar lo que editaste a mano */
async function tcActualizarAutomatico(silencioso) {
  // Depósito: el tipo de cambio es el dólar del depósito, fijado a mano
  if (typeof esModoDeposito === "function" && esModoDeposito() && typeof dolarDelDeposito === "function" && dolarDelDeposito() > 0) {
    depDolarAplicar();
    return;
  }
  const tipo = tcTipoConfigurado();
  const ctrl = new AbortController();
  const corte = setTimeout(() => ctrl.abort(), 8000);
  try {
    const res = await fetch(TC_API + encodeURIComponent(tipo), { signal: ctrl.signal, cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const d = await res.json();
    const venta = Number(d.venta);
    if (!(venta > 0)) throw new Error("cotización inválida");
    if (tipo !== tcTipoConfigurado()) return;   // respuesta vieja: en el medio cambió el tipo de dólar
    ingresoTC = { valor: venta, tipo, manual: false, dia: "", fecha: d.fechaActualizacion || new Date().toISOString(), ultimaAuto: Date.now(), fallo: false };
    tcGuardarLocal();
    tcRender();
    tcAplicarACampos(!silencioso);
  } catch (e) {
    console.warn("Tipo de cambio automático:", e);
    ingresoTC.fallo = true;
    ingresoTC.tipo = tipo;
    tcRender();
    if (!silencioso) toast("No se pudo traer la cotización automática. Se mantiene el último valor; podés cargarlo a mano.", "error");
  } finally {
    clearTimeout(corte);
  }
}

/** Escribiste el tipo de cambio en la barra: vale para hoy y se usa en todos los campos de abajo */
function tcEditarManual(v) {
  ingresoTC.valor = Number(v) || 0;
  ingresoTC.manual = true;
  ingresoTC.fallo = false;
  ingresoTC.dia = _hoyISO();
  tcGuardarLocal();
  tcRender();
  tcAplicarACampos(true);
}

/** Tipo de dólar elegido en Configuración (oficial, blue, MEP...). Si el servidor todavía no lo guarda, se usa el de este equipo */
function tcTipoConfigurado() {
  let t = "";
  try { t = String((obtenerConfigNegocio() || {}).dolarTipo || "").toLowerCase(); } catch (e) { /* sin config cargada */ }
  if (!TC_TIPOS[t]) { try { t = localStorage.getItem(TC_TIPO_LS) || ""; } catch (e) { t = ""; } }
  return TC_TIPOS[t] ? t : "oficial";
}

/** Atajo desde la barra: abre Configuración > Negocio, donde se elige el tipo de dólar */
function tcIrAConfiguracion() {
  mostrarSeccion("configuracion");
  setTimeout(() => {
    const btn = document.querySelector('.cfg-tab-btn[data-cfg-tab="negocio"]');
    if (btn) mostrarConfigTab("negocio", btn);
  }, 60);
}

/** Configuración: muestra cuánto vale hoy cada tipo de dólar para ayudar a elegir */
async function dolarMostrarValoresHoy() {
  const el = document.getElementById("cfgDolarHoy");
  if (!el) return;
  const ctrl = new AbortController();
  const corte = setTimeout(() => ctrl.abort(), 8000);
  try {
    const res = await fetch(TC_API.replace(/\/$/, ""), { signal: ctrl.signal, cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const lista = await res.json();
    clearTimeout(corte);
    const partes = Object.keys(TC_TIPOS).map(casa => {
      const d = (Array.isArray(lista) ? lista : []).find(x => x.casa === casa);
      return d && Number(d.venta) > 0 ? `${TC_TIPOS[casa]} $${Number(d.venta).toLocaleString("es-AR")}` : null;
    }).filter(Boolean);
    el.textContent = partes.length ? "Valor de venta hoy: " + partes.join(" · ") : "";
  } catch (e) {
    clearTimeout(corte);
    el.textContent = "";      // sin conexión: el selector funciona igual
  }
}

/* ---------- Burbuja con los 4 dólares principales (solo en POS y Dashboard) ----------
   Se actualiza al entrar a POS/Dashboard si pasaron más de 10 minutos, cada
   10 minutos mientras estás ahí, y al volver internet. Muestra la hora de la
   COTIZACIÓN (no la de la consulta): fines de semana y feriados el oficial no
   se mueve, y antes decía "Dólar hoy · 00:57" con el valor del viernes.
   Si una actualización falla, se siguen mostrando los últimos valores marcados
   como "sin actualizar" (antes la burbuja desaparecía). */
const DB_TIPOS = ["oficial", "blue", "bolsa", "tarjeta"];
const DB_LS_MIN = "vpos_dolar_burbuja_min";
const DB_REFRESCO_MS = 10 * 60 * 1000;
let dbDatos = null;            // { valores:{oficial:1540,...}, fechas:{oficial:Date,...}, consultado:ms, desactualizado:bool }
let dbSeccion = "";
let dbTimer = null;
let dbPidiendo = false;

/** "hoy 15:56" / "ayer 18:55" / "vie 18:55" / "02/10 18:55" */
function _dbFechaCorta(f) {
  if (!(f instanceof Date) || isNaN(f)) return "";
  const hora = f.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit", hour12: false });
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const dia = new Date(f); dia.setHours(0, 0, 0, 0);
  const dif = Math.round((hoy - dia) / 86400000);
  if (dif === 0) return "hoy " + hora;
  if (dif === 1) return "ayer " + hora;
  if (dif > 1 && dif < 7) return f.toLocaleDateString("es-AR", { weekday: "short" }).replace(".", "") + " " + hora;
  return f.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit" }) + " " + hora;
}

function dolarBurbujaRender() {
  const el = document.getElementById("dolarBurbuja");
  if (!el) return;
  // En el depósito se ve siempre (está su dólar fijado a mano), aunque no haya cotización del mercado
  const dep = typeof esModoDeposito === "function" && esModoDeposito() && typeof depDolarFilasBurbuja === "function";
  const visible = (dbSeccion === "pos" || dbSeccion === "dashboard") && (!!dbDatos || dep);
  el.classList.toggle("visible", visible);
  if (!visible) return;

  const fechas = dbDatos ? DB_TIPOS.map(t => dbDatos.fechas[t]).filter(f => f instanceof Date && !isNaN(f)) : [];
  const masNueva = fechas.length ? new Date(Math.max(...fechas)) : null;
  const filasMercado = dbDatos ? DB_TIPOS.filter(t => dbDatos.valores[t] > 0).map(t => {
    const f = dbDatos.fechas[t];
    const tit = f ? `Venta · cotización de ${f.toLocaleString("es-AR", { weekday: "long", day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}` : "Venta";
    return { tipo: t, valor: dbDatos.valores[t], html: `<div class="db-fila" title="${escapeHtml(tit)}"><span>${TC_TIPOS[t]}</span><b>$${Number(dbDatos.valores[t]).toLocaleString("es-AR", { maximumFractionDigits: 2 })}</b></div>` };
  }) : [];
  document.getElementById("dbFilas").innerHTML = dep ? depDolarFilasBurbuja(filasMercado) : filasMercado.map(f => f.html).join("");
  const titulo = el.querySelector(".db-tit span");
  if (titulo) titulo.textContent = dep
    ? (el.classList.contains("min") && dolarDelDeposito() ? "Dólar dep. $" + dolarDelDeposito().toLocaleString("es-AR", { maximumFractionDigits: 2 }) : "Dólar")
    : "Dólar venta";

  const h = document.getElementById("dbHora");
  if (h) {
    const sinRed = navigator.onLine === false || (dbDatos && dbDatos.desactualizado);
    h.textContent = !dbDatos ? "" : sinRed ? "sin actualizar" : (masNueva ? _dbFechaCorta(masNueva) : "");
    h.style.color = sinRed ? "var(--amber-600, #b45309)" : "";
  }
  el.title = dep ? "Dólar del depósito (fijado a mano) y cotizaciones del mercado. Tocá para minimizar"
    : "Cotización del dólar (valor de venta, DolarApi). Consultado a las " +
      new Date(dbDatos.consultado).toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }) + ". Tocá para minimizar";

  let min = false;
  try { min = localStorage.getItem(DB_LS_MIN) === "1"; } catch (e) {}
  el.classList.toggle("min", min);
}

function dolarBurbujaAlternar() {
  const el = document.getElementById("dolarBurbuja");
  if (!el) return;
  const min = !el.classList.contains("min");
  el.classList.toggle("min", min);
  try { localStorage.setItem(DB_LS_MIN, min ? "1" : "0"); } catch (e) {}
  if (typeof esModoDeposito === "function" && esModoDeposito()) dolarBurbujaRender();
}

async function dolarBurbujaActualizar() {
  if (dbPidiendo) return;
  if (typeof esModoDeposito === "function" && esModoDeposito() && typeof depDolarRefrescar === "function") depDolarRefrescar();
  if (navigator.onLine === false) { dolarBurbujaRender(); return; }
  dbPidiendo = true;
  const ctrl = new AbortController();
  const corte = setTimeout(() => ctrl.abort(), 8000);
  try {
    const res = await fetch(TC_API.replace(/\/$/, ""), { signal: ctrl.signal, cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const lista = await res.json();
    const valores = {}, fechas = {};
    DB_TIPOS.forEach(t => {
      const d = (Array.isArray(lista) ? lista : []).find(x => x.casa === t);
      if (d && Number(d.venta) > 0) {
        valores[t] = Number(d.venta);
        const f = d.fechaActualizacion ? new Date(d.fechaActualizacion) : null;
        fechas[t] = f && !isNaN(f) ? f : null;
      }
    });
    if (!Object.keys(valores).length) throw new Error("respuesta sin cotizaciones");
    dbDatos = { valores, fechas, consultado: Date.now(), desactualizado: false };
  } catch (e) {
    console.warn("Cotización del dólar:", e);
    if (dbDatos) dbDatos.desactualizado = true;   // se quedan los últimos valores, marcados como viejos
  } finally {
    clearTimeout(corte);
    dbPidiendo = false;
    dolarBurbujaRender();
  }
}

/** Se llama en cada cambio de sección: la burbuja solo aparece en POS y Dashboard */
function dolarBurbujaAlCambiarSeccion(id) {
  dbSeccion = id;
  if (!dbTimer) {
    dbTimer = setInterval(() => { if (dbSeccion === "pos" || dbSeccion === "dashboard") dolarBurbujaActualizar(); }, DB_REFRESCO_MS);
    window.addEventListener("online", () => { if (dbSeccion === "pos" || dbSeccion === "dashboard") dolarBurbujaActualizar(); });
    window.addEventListener("offline", dolarBurbujaRender);
  }
  if (id === "pos" || id === "dashboard") {
    const vencido = !dbDatos || dbDatos.desactualizado || (Date.now() - dbDatos.consultado) > DB_REFRESCO_MS;
    dolarBurbujaRender();
    if (vencido) dolarBurbujaActualizar();
  } else dolarBurbujaRender();
}

/** Configuración: guarda el tipo de dólar de referencia (compartido por todos los equipos si el servidor lo soporta) */
async function guardarDolarTipoForm() {
  const sel = document.getElementById("cfgDolarTipo");
  if (!sel) return;
  const tipo = TC_TIPOS[sel.value] ? sel.value : "oficial";
  try { localStorage.setItem(TC_TIPO_LS, tipo); } catch (e) { /* no pasa nada */ }

  const btn = document.getElementById("btnGuardarDolarTipo");
  const textoOriginal = btn ? btn.innerHTML : "";
  if (btn) { btn.disabled = true; btn.innerHTML = "Guardando..."; }
  try {
    const params = new URLSearchParams({ action: "guardarConfiguracionNegocio", dolarTipo: tipo });
    const response = await fetchConReintento(API_URL + "?" + params.toString());
    const data = await response.json();
    if (!data.success) { toast(data.message || "No se pudo guardar el dólar de referencia", "error"); return; }

    configNegocioCache = { ...configNegocioCache, dolarTipo: tipo };
    // Cambió el dólar de referencia: el ingreso de productos vuelve al valor automático de ese tipo
    ingresoTC.tipo = tipo; ingresoTC.manual = false; ingresoTC.ultimaAuto = 0;
    tcGuardarLocal();
    tcRender();
    toast(`Dólar de referencia: ${TC_TIPOS[tipo]}`, "success");
  } catch (error) {
    console.error("Error al guardar el dólar de referencia:", error);
    toast("Error de conexión al guardar el dólar de referencia", "error");
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = textoOriginal; }
  }
}

/** Se llama al entrar a Ingreso de Productos */
function tcAlAbrirSeccion() {
  if (!tcAlAbrirSeccion._listo) {
    tcCargarLocal();
    // Si tocás a mano un campo de tipo de cambio de abajo, esa carga no se pisa con las actualizaciones automáticas
    TC_CAMPOS.forEach(id => document.getElementById(id)?.addEventListener("input", ev => { ev.target.dataset.tcLocal = "1"; }));
    tcAlAbrirSeccion._listo = true;
  }
  // Un valor editado a mano solo vale el día en que se cargó
  if (ingresoTC.manual && ingresoTC.dia !== _hoyISO()) ingresoTC.manual = false;
  // Si en Configuración cambió el tipo de dólar, se vuelve al valor automático del nuevo tipo
  const tipoCfg = tcTipoConfigurado();
  if (ingresoTC.tipo !== tipoCfg) { ingresoTC.tipo = tipoCfg; ingresoTC.manual = false; ingresoTC.ultimaAuto = 0; }
  tcRender();
  const vencida = !ingresoTC.ultimaAuto || (Date.now() - ingresoTC.ultimaAuto) > 30 * 60 * 1000;
  if (!ingresoTC.manual && vencida) tcActualizarAutomatico(true);
}

function onCambioMonedaIngreso() {
  const moneda = document.getElementById("ipMoneda").value;
  const wrap = document.getElementById("ipTipoCambioWrap");
  if (wrap) wrap.style.display = moneda === "USD" ? "" : "none";
  const tc = document.getElementById("ipTipoCambio");
  if (moneda !== "USD") { tc.value = ""; delete tc.dataset.tcLocal; }
  else tcPrellenarCampo("ipTipoCambio");            // en dólares: se completa con el tipo de cambio del día
}

function onCambioOrigenBoleta() {
  const origen = document.getElementById("ipbOrigen").value;
  const esDeposito = origen === "DEPOSITO";
  const proveedorLabel = document.getElementById("ipbProveedorLabel");
  const proveedorInput = document.getElementById("ipbProveedor");
  const aviso = document.getElementById("ipbAvisoSinDeuda");
  const totalLabel = document.getElementById("ipCarritoTotalLabel");

  if (proveedorLabel) proveedorLabel.textContent = esDeposito ? "Sucursal / depósito de origen" : "Proveedor *";
  if (proveedorInput) proveedorInput.placeholder = esDeposito ? "Ej: Depósito central" : "Nombre del proveedor";
  if (aviso) aviso.style.display = esDeposito ? "" : "none";
  if (totalLabel) totalLabel.textContent = esDeposito
    ? "Total de la boleta (interno — no genera deuda)"
    : "Total de la boleta (lo que se le debe al proveedor)";

  renderCarritoBoleta();
}

let ipProductoActualEsNuevo = false;

// Carrito de la boleta que se está armando: cada ítem escaneado se
// acumula acá y recién se manda todo junto al backend al tocar
// "Guardar boleta" — así todos los productos quedan agrupados en UN
// solo documento (BOLETA_ID) por proveedor/día/N° de boleta.
let ipCarritoBoleta = [];

// ID de la boleta que está abierta en el modal de detalle/pago
let boletaProveedorIdActual = null;

// Datos completos (cabecera + items + pagos) de la boleta actualmente
// abierta en el modal de detalle — se guardan acá para poder armar el
// comprobante imprimible sin tener que volver a pedirlos al backend.
let detalleBoletaDatosActual = null;

// true cuando el modal de detalle de boleta se abrió desde "Ver boletas"
// de un proveedor (en vez de desde la sección Ingreso de Productos) —
// controla si hay que volver a mostrar esa ventana al cerrar el detalle.
let detalleBoletaAbiertoDesdeProveedor = false;

function _hoyISO() {
  return _fechaISOEnZona(new Date());
}

/** Refleja en pantalla si ya hay una boleta con este proveedor/fecha/N° (se le sumarían los ítems) o si es una boleta nueva */
function actualizarEstadoBoletaActual() {
  const estado = document.getElementById("ipbEstadoBoleta");
  if (!estado) return;
  const proveedor = document.getElementById("ipbProveedor").value.trim().toLowerCase();
  const fecha = document.getElementById("ipbFecha").value;
  const numero = document.getElementById("ipbNumeroBoleta").value.trim();

  if (!proveedor) { estado.textContent = ""; return; }

  const existente = (ingresosProductosGlobal || []).find(i =>
    String(i.PROVEEDOR || "").trim().toLowerCase() === proveedor &&
    String(i.FECHA || "") === fecha &&
    String(i.NUMERO_BOLETA || "").trim() === numero
  );

  if (existente) {
    estado.innerHTML = `⚠️ Ya existe una boleta con estos datos (Total actual: $${Number(existente.TOTAL || 0).toLocaleString("es-AR")}) — lo que agregues se va a sumar a esa misma boleta.`;
    estado.style.color = "var(--amber-600, #b45309)";
  } else {
    estado.textContent = "✓ Se va a crear una boleta nueva con estos datos.";
    estado.style.color = "var(--green-600, green)";
  }
}

/** Busca el código escaneado/tipeado en productosAdminGlobal; si no está, abre el formulario en modo "producto nuevo" */
function buscarProductoParaIngreso() {
  const input = document.getElementById("ipCodigoScan");
  const codigo = (input.value || "").trim();
  const estado = document.getElementById("ipEstadoBusqueda");

  if (!codigo) {
    toast("Escaneá o escribí un código de barras", "error");
    return;
  }

  const producto = (productosAdminGlobal || []).find(p => String(p.CODIGO) === codigo);

  poblarDatalistCategoriasIngreso();
  poblarDatalistProveedoresIngreso();

  document.getElementById("ipFormularioWrap").style.display = "";
  document.getElementById("ipCodigo").value = codigo;
  document.getElementById("ipCodigoMostrado").value = codigo;
  document.getElementById("ipCantidad").value = "";

  if (producto) {
    ipProductoActualEsNuevo = false;
    document.getElementById("ipTituloProducto").textContent = "✏️ " + (producto.PRODUCTO || codigo);
    document.getElementById("ipBadgeEstado").textContent = "Producto existente";
    document.getElementById("ipBadgeEstado").className = "badge bg-success";
    document.getElementById("ipNombre").value = producto.PRODUCTO || "";
    document.getElementById("ipNombre").readOnly = true;
    document.getElementById("ipCategoria").value = producto.CATEGORIA || "";
    document.getElementById("ipPrecio").value = "";
    document.getElementById("ipPrecioHint").textContent = "Lo que le pagás al proveedor por esta entrega — con esto se calcula el saldo que se le debe.";
    const costoUsdRef = Number(producto.COSTO_USD_REF || 0);
    if (costoUsdRef > 0) {
      document.getElementById("ipMoneda").value = "USD";
      onCambioMonedaIngreso();
      document.getElementById("ipPrecio").value = costoUsdRef;
      document.getElementById("ipPrecioHint").textContent =
        `Último costo cargado en USD: US$${costoUsdRef.toLocaleString("es-AR")} — corregilo si cambió, o el tipo de cambio.`;
    } else {
      document.getElementById("ipMoneda").value = "ARS";
      onCambioMonedaIngreso();
    }
    document.getElementById("ipPrecioVenta").value = Number(producto.PRECIO || 0);
    document.getElementById("ipPrecioVentaHint").textContent = "Dejalo así si no cambió, o corregilo si el precio al público cambió.";
    if (estado) { estado.textContent = "✓ Producto encontrado: " + (producto.PRODUCTO || ""); estado.style.color = "var(--green-600, green)"; }
  } else {
    ipProductoActualEsNuevo = true;
    document.getElementById("ipTituloProducto").textContent = "+ Producto nuevo";
    document.getElementById("ipBadgeEstado").textContent = "No existe — se va a crear";
    document.getElementById("ipBadgeEstado").className = "badge bg-warning text-dark";
    document.getElementById("ipNombre").value = "";
    document.getElementById("ipNombre").readOnly = false;
    document.getElementById("ipCategoria").value = "";
    document.getElementById("ipPrecio").value = "";
    document.getElementById("ipMoneda").value = "ARS";
    onCambioMonedaIngreso();
    document.getElementById("ipPrecioHint").textContent = "Lo que le pagás al proveedor por esta entrega.";
    document.getElementById("ipPrecioVenta").value = "";
    document.getElementById("ipPrecioVentaHint").textContent = "Precio de venta al público con el que se va a publicar en el POS y en la página. Si lo dejás vacío, se usa el precio de ingreso.";
    if (estado) { estado.textContent = "No existe todavía — completá los datos para crearlo"; estado.style.color = "var(--amber-600, #b45309)"; }
    document.getElementById("ipNombre").focus();
  }
}

/** Genera un código automático (pide uno al backend) para un producto sin código de barras */
async function generarCodigoIngresoAutomatico() {
  try {
    const res = await fetchAPI(API_URL + "?action=generarCodigoProducto");
    const data = await res.json();
    if (data.success) {
      document.getElementById("ipCodigoScan").value = data.codigo;
      buscarProductoParaIngreso();
    }
  } catch (e) {
    // Sin conexión: mismo formato que genera el servidor (PRD-tiempo-azar),
    // revisando que no exista en el catálogo cargado.
    const existentes = new Set((productosAdminGlobal || []).map(p => String(p.CODIGO).trim()));
    let codigo;
    do { codigo = "PRD-" + Date.now() + "-" + Math.floor(Math.random() * 900 + 100); } while (existentes.has(codigo));
    document.getElementById("ipCodigoScan").value = codigo;
    buscarProductoParaIngreso();
  }
}

function cancelarIngresoProducto() {
  if (_ipItemEditando) { ipCarritoBoleta.push(_ipItemEditando); _ipItemEditando = null; renderCarritoBoleta(); }
  document.getElementById("ipFormularioWrap").style.display = "none";
  document.getElementById("ipCodigoScan").value = "";
  document.getElementById("ipPrecio").value = "";
  document.getElementById("ipPrecioVenta").value = "";
  document.getElementById("ipMoneda").value = "ARS";
  document.getElementById("ipTipoCambio").value = "";
  document.getElementById("ipTipoCambioWrap").style.display = "none";
  document.getElementById("ipEstadoBusqueda").textContent = "";
  document.getElementById("ipCodigoScan").focus();
}

function poblarDatalistCategoriasIngreso() {
  const dl = document.getElementById("ipCategoriasList");
  if (!dl) return;
  const categorias = [...new Set((productosAdminGlobal || []).map(p => p.CATEGORIA).filter(Boolean))];
  dl.innerHTML = categorias.map(c => `<option value="${escapeHtml(c)}">`).join("");
}

function poblarDatalistProveedoresIngreso() {
  const dl = document.getElementById("ipProveedoresList");
  if (!dl) return;
  const proveedores = [...new Set((ingresosProductosGlobal || []).map(i => i.PROVEEDOR).filter(Boolean))];
  dl.innerHTML = proveedores.map(p => `<option value="${escapeHtml(p)}">`).join("");
}

/** Agrega el producto del formulario al carrito de la boleta (todavía no lo manda al backend) */
function agregarItemABoleta() {
  const origen = (document.getElementById("ipbOrigen")?.value) || "PROVEEDOR";
  const proveedor = document.getElementById("ipbProveedor").value.trim();
  if (!proveedor) {
    toast(origen === "DEPOSITO"
      ? "Completá de qué sucursal/depósito viene la mercadería antes de agregar productos"
      : "Completá el proveedor de la boleta (arriba) antes de agregar productos", "error");
    document.getElementById("ipbProveedor").focus();
    return;
  }

  const codigo = document.getElementById("ipCodigo").value.trim();
  const nombre = document.getElementById("ipNombre").value.trim();
  const categoria = document.getElementById("ipCategoria").value.trim();
  const moneda = document.getElementById("ipMoneda").value === "USD" ? "USD" : "ARS";
  const precio = Number(document.getElementById("ipPrecio").value || 0);
  const tipoCambio = Number(document.getElementById("ipTipoCambio").value || 0);
  const precioVentaInput = document.getElementById("ipPrecioVenta").value;
  const precioVenta = precioVentaInput !== "" ? Number(precioVentaInput) : null;
  const cantidad = Number(document.getElementById("ipCantidad").value || 0);

  if (ipProductoActualEsNuevo && !nombre) {
    toast("Ingresá el nombre del producto para poder crearlo", "error");
    return;
  }
  if (!cantidad || cantidad <= 0) {
    toast("Ingresá una cantidad válida, mayor a 0", "error");
    return;
  }
  if (!precio || precio <= 0) {
    toast("Ingresá el precio de ingreso (costo) — se usa para calcular el saldo al proveedor", "error");
    return;
  }
  if (moneda === "USD" && (!tipoCambio || tipoCambio <= 0)) {
    toast("Ingresá el tipo de cambio para convertir el precio en dólares a pesos", "error");
    document.getElementById("ipTipoCambio").focus();
    return;
  }

  const precioARS = moneda === "USD" ? precio * tipoCambio : precio;

  ipCarritoBoleta.push({
    codigo, producto: nombre || codigo, categoria,
    moneda, precio, tipoCambio: moneda === "USD" ? tipoCambio : null, precioARS,
    precioVenta: precioVenta !== null && !isNaN(precioVenta) ? precioVenta : null,
    cantidad,
    esNuevo: ipProductoActualEsNuevo
  });

  _ipItemEditando = null;
  renderCarritoBoleta();
  toast("Agregado a la boleta", "success");
  cancelarIngresoProducto();
}

function renderCarritoBoleta() {
  const tbody = document.getElementById("ipCarritoBody");
  const totalEl = document.getElementById("ipCarritoTotal");
  if (!tbody) return;

  if (ipCarritoBoleta.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" class="text-center text-muted py-3">Todavía no agregaste productos</td></tr>`;
    if (totalEl) totalEl.textContent = "$0";
    return;
  }

  let total = 0;
  tbody.innerHTML = ipCarritoBoleta.map((it, idx) => {
    // El subtotal (en pesos) se calcula con el precio de INGRESO (costo),
    // convertido a ARS si el producto se cargó en dólares.
    const subtotalARS = it.precioARS * it.cantidad;
    total += subtotalARS;
    const precioVentaTexto = (it.precioVenta !== null && it.precioVenta !== undefined)
      ? "$" + Number(it.precioVenta).toLocaleString("es-AR")
      : "—";
    const precioIngresoTexto = it.moneda === "USD"
      ? "US$" + Number(it.precio).toLocaleString("es-AR") + ` (TC ${it.tipoCambio})`
      : "$" + Number(it.precio).toLocaleString("es-AR");
    return `
      <tr>
        <td class="mono">${escapeHtml(it.codigo || "—")}</td>
        <td>${escapeHtml(it.producto)} ${it.esNuevo ? '<span class="badge bg-warning text-dark">nuevo</span>' : ""}</td>
        <td>${escapeHtml(it.categoria || "—")}</td>
        <td class="money">${Number(it.cantidad).toLocaleString("es-AR")}</td>
        <td class="money">${precioIngresoTexto}</td>
        <td class="money">${precioVentaTexto}</td>
        <td class="money">$${subtotalARS.toLocaleString("es-AR")}</td>
        <td style="white-space:nowrap;"><button class="btn btn-outline-primary btn-sm" onclick="editarItemCarrito(${idx})" title="Revisar / corregir">✏️</button>
          <button class="btn btn-outline-danger btn-sm ms-1" onclick="quitarItemCarrito(${idx})">✕</button></td>
      </tr>`;
  }).join("");

  if (totalEl) totalEl.textContent = "$" + total.toLocaleString("es-AR");
}

/** Vuelve a abrir un renglón de la boleta en el formulario (costo, moneda, precio de venta, cantidad...) para corregirlo */
function editarItemCarrito(idx) {
  const it = ipCarritoBoleta[idx];
  if (!it) return;
  if (_ipItemEditando) { ipCarritoBoleta.push(_ipItemEditando); }  // otro que se estaba corrigiendo vuelve como estaba
  ipCarritoBoleta.splice(ipCarritoBoleta.indexOf(it), 1);
  _ipItemEditando = it;
  renderCarritoBoleta();
  document.getElementById("ipCodigoScan").value = it.codigo || "";
  buscarProductoParaIngreso();
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v === null || v === undefined ? "" : v; };
  if (ipProductoActualEsNuevo) set("ipNombre", it.producto);
  set("ipCategoria", it.categoria || "");
  set("ipMoneda", it.moneda === "USD" ? "USD" : "ARS");
  onCambioMonedaIngreso();
  set("ipPrecio", it.precio || "");
  if (it.moneda === "USD") set("ipTipoCambio", it.tipoCambio || "");
  set("ipPrecioVenta", it.precioVenta);
  set("ipCantidad", it.cantidad);
  const est = document.getElementById("ipEstadoBusqueda");
  if (est) { est.textContent = "Corregí lo que haga falta y tocá “Agregar a la boleta” (si cancelás, queda como estaba)"; est.style.color = "var(--navy-700, #1e3a8a)"; }
  const form = document.getElementById("ipFormularioWrap");
  if (form && form.scrollIntoView) form.scrollIntoView({ behavior: "smooth", block: "center" });
}

let _ipItemEditando = null; // renglón sacado de la boleta para corregirlo (vuelve si se cancela)

function quitarItemCarrito(idx) {
  ipCarritoBoleta.splice(idx, 1);
  renderCarritoBoleta();
}

function vaciarCarritoBoleta() {
  if (!ipCarritoBoleta.length) return;
  // confirmarAccion (modal propio) en vez de confirm() nativo, que en Electron deja el foco roto
  confirmarAccion(
    "¿Vaciar todos los productos cargados en esta boleta?",
    () => { ipCarritoBoleta = []; renderCarritoBoleta(); },
    "Vaciar boleta"
  );
}

/* =====================================================================
   OCR de listas de proveedores -> carrito de "Ingreso de Productos"
   Gratis: Tesseract.js corre en el navegador/Electron, sin servidor ni API key.
   Flujo: foto -> OCR -> tabla editable de revisión -> "Agregar a la boleta"
   (empuja a ipCarritoBoleta, así guardar/stock/deuda funcionan igual que siempre).
   El HTML (tarjeta #ocrCard y modal #ocrModalBackdrop) está en index.html.
===================================================================== */
const OCR_TESSERACT_SRC = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
let ocrFilas = [];
let ocrModoPruebaActivo = false;   // true = "Modo prueba": muestra y compara, pero NO agrega nada a la boleta ni a la base
let ocrUltimaLectura = null;       // { ms, confianza, texto, archivo } de la última foto leída (para stats y JSON)
let ocrImgOriginalURL = null;
// pdf.js: lee el texto de PDFs digitales directamente (sin OCR, sin errores de dígitos)
const OCR_PDFJS_SRC = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
const OCR_PDFJS_WORKER = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
let ocrPdfWorkerListo = false;

/** Minúsculas, sin tildes ni signos: para comparar encabezados y nombres */
function ocrNorm(t) {
  return normalizarBusquedaPOS(t).replace(/[^a-z0-9ñ ]+/g, " ").replace(/\s+/g, " ").trim();
}

/** Carga el motor OCR recién la primera vez que se usa (no pesa en el resto del panel) */
function ocrCargarTesseract() {
  return new Promise((resolve, reject) => {
    if (window.Tesseract) return resolve();
    const s = document.createElement("script");
    s.src = OCR_TESSERACT_SRC;
    s.onload = resolve;
    s.onerror = () => reject(new Error("No se pudo cargar el motor OCR (se necesita internet la primera vez)"));
    document.head.appendChild(s);
  });
}

/** Escala + escala de grises + estirado de contraste: mejora bastante la lectura de fotos de celular */
async function ocrPreprocesarImagen(file) {
  const img = await createImageBitmap(file);
  const escala = Math.max(0.5, Math.min(3, 2400 / Math.max(img.width, img.height)));
  const c = document.createElement("canvas");
  c.width = Math.round(img.width * escala);
  c.height = Math.round(img.height * escala);
  const ctx = c.getContext("2d");
  ctx.drawImage(img, 0, 0, c.width, c.height);
  const id = ctx.getImageData(0, 0, c.width, c.height);
  const d = id.data, hist = new Array(256).fill(0);
  for (let i = 0; i < d.length; i += 4) {
    const g = Math.round(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
    d[i] = d[i + 1] = d[i + 2] = g;
    hist[g]++;
  }
  const total = d.length / 4;
  let acc = 0, lo = 0, hi = 255;
  for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc >= total * 0.02) { lo = v; break; } }
  acc = 0;
  for (let v = 255; v >= 0; v--) { acc += hist[v]; if (acc >= total * 0.02) { hi = v; break; } }
  const rango = Math.max(1, hi - lo);
  for (let i = 0; i < d.length; i += 4) {
    const g = Math.max(0, Math.min(255, ((d[i] - lo) * 255) / rango));
    d[i] = d[i + 1] = d[i + 2] = g;
  }
  const bandas = ocrInvertirFranjasOscuras(d, c.width, c.height);
  c._lineas = ocrBorrarLineasTabla(d, c.width, c.height);
  c._lineas.bandas = bandas;
  ctx.putImageData(id, 0, 0);
  return c;
}

/** Encabezados con fondo oscuro y letra blanca: el OCR no los lee. Se invierten esas franjas para que quede letra negra sobre blanco */
function ocrInvertirFranjasOscuras(d, w, h) {
  const bandas = [];
  const oscura = new Array(h).fill(false);
  for (let y = 0; y < h; y++) {
    let n = 0;
    for (let x = 0; x < w; x += 2) if (d[(y * w + x) * 4] < 110) n++;
    oscura[y] = n > (w / 2) * 0.5;
  }
  let ini = -1;
  for (let y = 0; y <= h; y++) {
    const o = y < h && oscura[y];
    if (o && ini < 0) ini = y;
    if (!o && ini >= 0) {
      if (y - ini >= 10) bandas.push([ini, y]);
      if (y - ini >= 10) for (let yy = ini; yy < y; yy++) for (let x = 0; x < w; x++) { const i = (yy * w + x) * 4; d[i] = d[i + 1] = d[i + 2] = 255 - d[i]; }
      ini = -1;
    }
  }
  return bandas;
}

/** Borra las líneas largas (bordes de la tabla): confunden al OCR ("|", "1", "l") y pegan las columnas */
function ocrBorrarLineasTabla(d, w, h) {
  const oscuro = (x, y) => d[(y * w + x) * 4] < 120;
  const minV = Math.max(50, Math.round(h * 0.05)), minH = Math.max(80, Math.round(w * 0.06));
  const marcar = new Uint8Array(w * h);
  const vx = new Uint8Array(w), hy = new Uint8Array(h);
  for (let x = 0; x < w; x++) {                       // verticales
    let ini = -1;
    for (let y = 0; y <= h; y++) {
      const o = y < h && oscuro(x, y);
      if (o && ini < 0) ini = y;
      if (!o && ini >= 0) { if (y - ini >= minV) { vx[x] = 1; for (let k = ini; k < y; k++) marcar[k * w + x] = 1; } ini = -1; }
    }
  }
  for (let y = 0; y < h; y++) {                       // horizontales
    let ini = -1;
    for (let x = 0; x <= w; x++) {
      const o = x < w && oscuro(x, y);
      if (o && ini < 0) ini = x;
      if (!o && ini >= 0) { if (x - ini >= minH) { hy[y] = 1; for (let k = ini; k < x; k++) marcar[y * w + k] = 1; } ini = -1; }
    }
  }
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {   // se borra la línea y 2 px alrededor (las líneas de foto son gruesas)
    let borrar = false;
    for (let k = -2; k <= 2 && !borrar; k++) {
      if (x + k >= 0 && x + k < w && marcar[y * w + x + k]) borrar = true;
      if (!borrar && y + k >= 0 && y + k < h && marcar[(y + k) * w + x]) borrar = true;
    }
    if (borrar) { const i = (y * w + x) * 4; d[i] = d[i + 1] = d[i + 2] = 255; }
  }
  // posiciones de las líneas (una por grupo de píxeles contiguos): sirven para reconstruir la grilla de la tabla
  const centros = v => { const r = []; let ini = -1; for (let i = 0; i <= v.length; i++) { const o = i < v.length && v[i]; if (o && ini < 0) ini = i; if (!o && ini >= 0) { r.push((ini + i - 1) / 2); ini = -1; } } return r; };
  return { xs: centros(vx), ys: centros(hy) };
}

/** "1.234,50" / "1,234.50" / "1234" -> número */
function ocrNumeroAR(s) {
  s = String(s).replace(/[^\d.,]/g, "");
  if (!s) return NaN;
  const c = s.lastIndexOf(","), d = s.lastIndexOf(".");
  if (c > -1 && d > -1) {
    const dec = c > d ? "," : ".", mil = dec === "," ? "." : ",";
    s = s.split(mil).join("").replace(dec, ".");
  } else if (c > -1) {
    s = /,\d{1,2}$/.test(s) ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (d > -1) {
    s = /^\d{1,3}(\.\d{3})+$/.test(s) ? s.replace(/\./g, "") : s;
  }
  return parseFloat(s);
}
const ocrEsNumero = t => /^\d[\d.,]*$/.test(t);

/** Una línea de texto -> { codigo, nombre, cantidad, precio } (o null si no parece un producto).
    Cubre layouts tipo: [código] [cant] descripción precio [subtotal] */
function ocrParsearLinea(linea) {
  // Tokens limpios: sin bordes de tabla (| _), sin $ y sin corchetes que el OCR mete pegados a las palabras
  // (se sacan los signos sueltos pegados a los bordes de cada palabra: "38.43}" -> "38.43", "‘" -> nada)
  const t = linea.replace(/[|_]/g, " ").replace(/\$/g, " ").trim().split(/\s+/)
    .map(x => x.replace(/^[^0-9A-Za-zÁÉÍÓÚÑáéíóúñ]+|[^0-9A-Za-zÁÉÍÓÚÑáéíóúñ]+$/g, "")).filter(Boolean);
  if (t.length < 2 || !/[a-zA-ZáéíóúñÁÉÍÓÚÑ]{3,}/.test(t.join(" "))) return null;
  // Encabezados y totales (PRODUCTO / CANT / PRECIO / SUBTOTAL / IVA...) no son productos
  if (/^(pedido|producto|descripci[oó]n|art[ií]culo|detalle|c[oó]digo|cant\.?|cantidad|sub\s?total|total|iva|remito|fecha|cliente)\b/i.test(t.join(" "))) return null;

  // Columnas "Pedido" y "Fecha" al principio (ej: PED-581808 11/09/2026 ...): no son parte del producto
  // (la fecha se detecta de forma tolerante porque el OCR suele leerla mal: "/11/09/2026", "111/09/2026")
  const esFecha = x => /\d{1,3}[\/.-]\d{2}[\/.-]\d{4}$/.test(x);
  if (t.length > 3 && esFecha(t[1]) && /^[A-Za-z]{1,4}-?[A-Za-z0-9]{3,}$/.test(t[0])) { t.shift(); t.shift(); }
  else if (t.length > 3 && esFecha(t[0])) t.shift();

  let codigo = "", qty = null;
  if (/^\d{6,14}$/.test(t[0])) codigo = t.shift();
  else if (/^[A-Za-z]{1,4}-?\d{2,}$/.test(t[0]) && t.length > 2) codigo = t.shift();

  if (t.length > 2 && /^\d{1,4}[xX]?$/.test(t[0])) { qty = parseInt(t[0], 10); t.shift(); }

  // Números del final: [cantidad] [precio] [subtotal]  (a veces también el código de la columna "Código")
  const crudos = [];
  while (t.length > 1 && ocrEsNumero(t[t.length - 1]) && crudos.length < 3) crudos.unshift(t.pop());

  // Código al final: muchos proveedores repiten el código al final del nombre Y en su propia columna.
  // Solo se lo toma como código si aparece repetido, para no confundir "Coca Cola 500" con un código.
  const esCod = x => /^\d{3,}$/.test(x);
  const pegado = t.length > 2 ? t[t.length - 1].match(/^(\d{3,})\1$/) : null;   // "53505350" = repetido + código, pegados
  if (!codigo && pegado) { codigo = pegado[1]; t.pop(); }
  else if (!codigo) {
    if (crudos.length === 3 && t.length > 2 && esCod(crudos[0]) && t[t.length - 1] === crudos[0]) {
      codigo = crudos.shift(); t.pop();                       // falta el subtotal: [código, cant, precio]
    } else if (t.length > 2 && esCod(t[t.length - 1]) && t[t.length - 2] === t[t.length - 1]) {
      codigo = t.pop(); t.pop();                              // cant, precio, subtotal completos
    }
  }

  const nums = crudos.map(ocrNumeroAR);
  let precio = null, qtyCruda = null;
  if (qty !== null) precio = nums.length ? nums[0] : null;
  else if (nums.length >= 3) { qty = nums[0]; qtyCruda = crudos[0]; precio = nums[1]; }
  else if (nums.length === 2) {
    // [cantidad, precio] si la 1ª es entera y la 2ª parece precio (mayor o con decimales); si no, [precio, subtotal]
    if (Number.isInteger(nums[0]) && nums[0] <= 999 && (nums[1] > nums[0] || !Number.isInteger(nums[1]))) { qty = nums[0]; qtyCruda = crudos[0]; precio = nums[1]; }
    else precio = nums[0];
  } else if (nums.length === 1) precio = nums[0];
  // Una cantidad con cero adelante ("0007120") o gigante ("11802142") es un código y una cantidad que el OCR pegó:
  // mejor dejarla vacía (queda marcada en amarillo para completar) que cargar un dato equivocado.
  if (qty !== null && ((qtyCruda && /^0\d/.test(qtyCruda)) || qty > 99999)) qty = null;

  const nombre = t.join(" ").replace(/\s+[xX×]$/, "").replace(/[-–—:.,;]+$/, "").trim();
  if (nombre.length < 3) return null;
  return { codigo, nombre, cantidad: qty, precio: (precio === null || isNaN(precio)) ? null : precio };
}

/** Busca el producto existente: por código exacto y, si no, por nombre parecido (>= 70% de coincidencia) */
function ocrBuscarProductoExistente(fila) {
  const prods = productosAdminGlobal || [];
  if (fila.codigo) {
    const p = prods.find(x => String(x.CODIGO).trim() === fila.codigo);
    if (p) return p;
  }
  const toks = ocrNorm(fila.nombre).split(" ").filter(w => w.length > 2);
  if (!toks.length) return null;
  let mejor = null, mejorScore = 0;
  for (const p of prods) {
    // Solo palabras de 3+ letras; la coincidencia parcial (una dentro de la otra) exige 4+ para que
    // "a", "x", "de" o un nombre vacío no den falsos "Existente" con cualquier texto basura.
    const pt = ocrNorm(p.PRODUCTO).split(" ").filter(x => x.length > 2);
    const hit = toks.filter(w => pt.some(x => x === w || (Math.min(x.length, w.length) >= 4 && (x.includes(w) || w.includes(x))))).length;
    const score = hit / toks.length;
    if (hit >= Math.min(2, toks.length) && score > mejorScore) { mejorScore = score; mejor = p; }
  }
  return mejorScore >= 0.7 ? mejor : null;
}

/** Agrega a cada fila: si ya existe en tu base (matchCodigo), si está tildada y el precio de venta (vacío por ahora) */
function ocrEnriquecerFilas(filas) {
  return filas.map(f => {
    const m = ocrBuscarProductoExistente(f);
    return Object.assign(f, { matchCodigo: m ? String(m.CODIGO) : "", incluir: true, precioVenta: f.precioVenta ?? null });
  });
}

/** Texto -> filas. Si el texto viene de un PDF (celdas separadas por " | ") usa las columnas; si no, lee línea por línea */
function ocrTextoAFilas(texto) {
  const lineas = String(texto || "").split(/\r?\n/);
  let filas = [];
  if (lineas.some(l => l.includes(" | "))) {
    const celdas = lineas.filter(l => l.trim()).map(l => l.split(" | ").map(s => ({ str: s.trim() })));
    filas = ocrFilasDesdeCeldas(celdas);
  }
  if (!filas.length) filas = lineas.map(l => ocrParsearLinea(l)).filter(Boolean);
  return ocrEnriquecerFilas(filas);
}

/* ---------------- PDF: lectura directa del texto por columnas (sin OCR) ---------------- */

/** ¿Qué columna es? (producto / código / cantidad / precio / subtotal), según el texto del encabezado */
function ocrRolDeEncabezado(txt) {
  const t = ocrNorm(txt);
  if (!t) return "";
  if (/^(sub ?total|total|importe|monto)\b/.test(t)) return "subtotal";
  if (/^(producto|descripcion|articulo|detalle|item|nombre|mercaderia)\b/.test(t)) return "nombre";
  if (/^(codigo|cod|sku|ean|barras|ref|referencia)\b/.test(t)) return "codigo";
  if (/^(cantidad|cant|unidades|uds|qty|bultos)\b/.test(t)) return "cantidad";
  if (/^(precio|costo|p unit|pu|unitario|valor)\b/.test(t)) return "precio";
  // Fotos: el OCR a veces escribe mal el encabezado ("Cantided", "Preclo"): se acepta con hasta 2 letras distintas
  if (/^[a-z]{4,10}$/.test(t)) {
    const dist = (x, y) => {
      const m = Array.from({ length: x.length + 1 }, (_, i) => [i]);
      for (let j = 1; j <= y.length; j++) m[0][j] = j;
      for (let i = 1; i <= x.length; i++) for (let j = 1; j <= y.length; j++)
        m[i][j] = Math.min(m[i - 1][j] + 1, m[i][j - 1] + 1, m[i - 1][j - 1] + (x[i - 1] === y[j - 1] ? 0 : 1));
      return m[x.length][y.length];
    };
    const cand = { producto: "nombre", codigo: "codigo", cantidad: "cantidad", precio: "precio", subtotal: "subtotal" };
    for (const k of Object.keys(cand)) if (dist(t, k) <= (k.length >= 7 ? 2 : 1)) return cand[k];
  }
  return "otro";
}

/** Items de texto del PDF (con posición) -> filas de celdas [{str,x0,x1}], de arriba hacia abajo y de izquierda a derecha */
function ocrAgruparItemsEnFilas(items) {
  const lim = items.filter(i => i.str && i.str.trim() !== "").sort((a, b) => b.y - a.y);
  const grupos = [];
  lim.forEach(it => {
    const g = grupos.find(g => Math.abs(g.y - it.y) <= Math.max(2, (it.h || 8) * 0.4));
    if (g) g.items.push(it); else grupos.push({ y: it.y, items: [it] });
  });
  grupos.sort((a, b) => b.y - a.y);
  return grupos.map(g => {
    g.items.sort((a, b) => a.x - b.x);
    const celdas = [];
    g.items.forEach(it => {
      const alto = it.h || 8;
      const ult = celdas[celdas.length - 1];
      const hueco = ult ? it.x - ult.x1 : Infinity;
      if (ult && hueco <= alto * 0.35) {              // mismo texto partido en trozos: se une
        ult.str += (hueco > alto * 0.12 ? " " : "") + it.str.trim();
        ult.x1 = it.x + it.w;
      } else celdas.push({ str: it.str.trim(), x0: it.x, x1: it.x + it.w });
    });
    celdas.y0 = Math.min(...g.items.map(i => i.top !== undefined ? i.top : Infinity));   // solo fotos: franja vertical de la fila
    celdas.y1 = Math.max(...g.items.map(i => i.bot !== undefined ? i.bot : -Infinity));
    return celdas;
  });
}

/** Filas de celdas -> productos. Detecta la fila de encabezado (Producto/Código/Cantidad/Precio) y arma cada producto por columna */
function ocrFilasDesdeCeldas(filasCeldas) {
  // 1) Cada fila de encabezado abre un bloque con las filas de datos que le siguen (sirve para PDFs de varias páginas)
  const bloques = [];
  let actual = null;
  for (const celdas of filasCeldas) {
    const roles = celdas.map(c => ocrRolDeEncabezado(c.str));
    const clave = roles.filter(r => r === "nombre" || r === "codigo" || r === "cantidad" || r === "precio").length;
    if (roles.includes("nombre") && clave >= 3) {
      actual = { cols: celdas.map((c, i) => ({ rol: roles[i], x0: c.x0, x1: c.x1, txt: c.str })), filas: [] };
      bloques.push(actual);
    } else if (actual) actual.filas.push(celdas);
  }

  const out = [];
  bloques.forEach(b => {
    // Encabezados pegados en una sola celda ("Codigo Cantidad"): si las filas de datos tienen más celdas que el
    // encabezado, se separa esa celda en una columna por palabra (repartiendo el ancho según las letras).
    const frecuencias = {};
    b.filas.forEach(f => { if (f.length >= 4) frecuencias[f.length] = (frecuencias[f.length] || 0) + 1; });
    const m = Number(Object.keys(frecuencias).sort((p, q) => frecuencias[q] - frecuencias[p])[0] || 0);
    if (m > b.cols.length) {
      const nuevas = [];
      b.cols.forEach(col => {
        const pal = String(col.txt || "").split(/\s+/).filter(Boolean);
        const rolesP = pal.map(ocrRolDeEncabezado);
        const separable = pal.length >= 2 && rolesP.every(r => r && r !== "otro" && r !== "subtotal") && new Set(rolesP).size === pal.length;
        if (!separable) { nuevas.push(col); return; }
        const total = pal.join("").length;
        const conPos = typeof col.x0 === "number" && typeof col.x1 === "number";
        let acum = 0;
        pal.forEach((w, k) => {
          const x0 = conPos ? col.x0 + (col.x1 - col.x0) * acum / total : undefined;
          acum += w.length;
          const x1 = conPos ? col.x0 + (col.x1 - col.x0) * acum / total : undefined;
          nuevas.push({ rol: rolesP[k], x0, x1, txt: w });
        });
      });
      if (nuevas.length === m) b.cols = nuevas;      // solo si con la separación las columnas coinciden con las filas
    }
    const n = b.cols.length;
    // 2) Rango horizontal de cada columna = encabezado + lo que ocupan sus datos en las filas completas.
    //    Así los números alineados a la derecha (o textos a la izquierda) igual caen en su columna.
    const rangos = b.cols.map(c => ({ x0: c.x0, x1: c.x1 }));
    const completas = b.filas.filter(f => f.length === n && f.every(c => c.x0 !== undefined));
    if (completas.length) rangos.forEach((r, i) => {
      completas.forEach(f => { r.x0 = Math.min(r.x0, f[i].x0); r.x1 = Math.max(r.x1, f[i].x1); });
    });

    b.filas.forEach(celdas => {
      // Texto sin posiciones (reprocesado a mano) y celdas de menos: no hay cómo saber cuál falta -> parser de líneas
      if (celdas.length !== n && celdas.every(c => c.x0 === undefined)) {
        const f = ocrParsearLinea(celdas.map(c => c.str).join(" "));
        if (f) out.push(f);
        return;
      }
      const porRol = {};
      if (celdas.length === n) {
        celdas.forEach((c, i) => { porRol[b.cols[i].rol] = c.str; });        // misma cantidad de celdas: por orden
      } else {
        celdas.forEach(c => {                                                 // si no (celda vacía): por posición horizontal
          if (c.x0 === undefined) return;
          let mejor = -1, mejorSolape = 0;
          rangos.forEach((r, i) => {
            const solape = Math.min(c.x1, r.x1) - Math.max(c.x0, r.x0);
            if (solape > mejorSolape) { mejorSolape = solape; mejor = i; }
          });
          if (mejor < 0) {                                                    // sin solape: la columna más cercana
            let dmin = Infinity;
            rangos.forEach((r, i) => { const d = Math.max(r.x0 - c.x1, c.x0 - r.x1, 0); if (d < dmin) { dmin = d; mejor = i; } });
          }
          if (mejor >= 0) { const r = b.cols[mejor].rol; porRol[r] = (porRol[r] ? porRol[r] + " " : "") + c.str; }
        });
      }

      let nombre = String(porRol.nombre || "").trim();
      const cant = ocrNumeroAR(porRol.cantidad || "");
      const precio = ocrNumeroAR(porRol.precio || "");
      if (!/[a-zA-ZáéíóúñÁÉÍÓÚÑ]{3,}/.test(nombre)) return;                 // fila vacía, total, etc.
      if (isNaN(cant) && isNaN(precio)) return;
      let codigo = String(porRol.codigo || "").trim();
      // Si el nombre es tan largo que se mete debajo de la celda del código, el PDF lo entrega todo junto
      // ("…TSV5350 5350 5350" = nombre + código repetido + código): se separa el código del final.
      // (el código puede venir separado por espacio "…5350 5350" o pegado "…53505350" = repetido + código)
      if (!codigo) {
        const mc = nombre.match(/^(.*\S)\s+(\d{3,})\s+\2$/) || nombre.match(/^(.*\S)\s+(\d{3,})\2$/);
        if (mc) { nombre = mc[1]; codigo = mc[2]; }
      }
      // Muchos proveedores repiten el código al final del nombre: se saca de ahí (ya está en su columna)
      if (codigo && nombre.endsWith(" " + codigo)) nombre = nombre.slice(0, -codigo.length).trim();
      out.push({ codigo, nombre, cantidad: isNaN(cant) ? null : cant, precio: isNaN(precio) ? null : precio });
    });
  });
  return out;
}

/** Recorre todas las páginas del PDF y devuelve las filas de celdas */
async function ocrExtraerCeldasPDF(pdf) {
  const filasCeldas = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const tc = await page.getTextContent();
    const items = tc.items.filter(i => typeof i.str === "string").map(i => ({
      str: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: i.height || Math.abs(i.transform[3]) || 8
    }));
    ocrAgruparItemsEnFilas(items).forEach(r => filasCeldas.push(r));
  }
  return filasCeldas;
}

/** Carga pdf.js recién cuando se sube un PDF por primera vez */
async function ocrCargarPdfJs() {
  if (!window.pdfjsLib) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = OCR_PDFJS_SRC;
      s.onload = resolve;
      s.onerror = () => reject(new Error("No se pudo cargar el lector de PDF (se necesita internet la primera vez)"));
      document.head.appendChild(s);
    });
  }
  if (!ocrPdfWorkerListo) {
    try {   // el worker se baja como texto y se levanta desde un blob (los navegadores no permiten workers de otro dominio)
      const r = await fetch(OCR_PDFJS_WORKER);
      const blob = new Blob([await r.text()], { type: "text/javascript" });
      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
    } catch (e) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = OCR_PDFJS_WORKER;
    }
    ocrPdfWorkerListo = true;
  }
}

/** PDF -> { texto, filasCeldas, vista } (vista = imagen de la 1ª página, para el modo prueba) */
async function ocrLeerPDF(file) {
  await ocrCargarPdfJs();
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const filasCeldas = await ocrExtraerCeldasPDF(pdf);
  let vista = "";
  try {
    const page = await pdf.getPage(1);
    const vp = page.getViewport({ scale: 1.3 });
    const cv = document.createElement("canvas");
    cv.width = vp.width; cv.height = vp.height;
    await page.render({ canvasContext: cv.getContext("2d"), viewport: vp }).promise;
    vista = cv.toDataURL("image/jpeg", 0.8);
  } catch (e) { /* la vista previa es opcional */ }
  const texto = filasCeldas.map(r => r.map(c => c.str).join(" | ")).join("\n");
  return { texto, filasCeldas, vista, paginas: pdf.numPages };
}

function ocrRenderTabla() {
  const tb = document.getElementById("ocrTabla");
  if (!tb) return;
  const prods = productosAdminGlobal || [];
  // El costo puede estar en ARS o USD (selector de arriba); el precio de venta es SIEMPRE en pesos
  const usd = document.getElementById("ocrMoneda")?.value === "USD";
  const tc = Number(document.getElementById("ocrTc")?.value || 0);
  const thCosto = document.getElementById("ocrThCosto");
  if (thCosto) thCosto.textContent = usd ? "Costo unit. (USD)" : "Costo unit. (ARS)";
  tb.innerHTML = ocrFilas.map((f, i) => {
    const falta = !f.cantidad || !f.precio;
    const equivARS = (usd && tc > 0 && f.precio)
      ? `<div class="text-muted" style="font-size:11px;">≈ $${(Math.round(f.precio * tc * 100) / 100).toLocaleString("es-AR")} ARS</div>` : "";
    const ex = f.matchCodigo ? prods.find(p => String(p.CODIGO).trim() === String(f.matchCodigo).trim()) : null;
    const phVenta = ex ? "actual: " + ex.PRECIO : "opcional";
    return `<tr style="${falta ? "background:#fff8e1;" : ""}">
      <td><input type="checkbox" ${f.incluir ? "checked" : ""} onchange="ocrEditarFila(${i},'incluir',this.checked)"></td>
      <td><input class="form-control form-control-sm mono" value="${escapeHtml(f.matchCodigo || f.codigo || "")}" onchange="ocrEditarFila(${i},'codigo',this.value)"></td>
      <td><input class="form-control form-control-sm" value="${escapeHtml(f.nombre)}" onchange="ocrEditarFila(${i},'nombre',this.value)"></td>
      <td><input type="number" min="1" class="form-control form-control-sm" value="${f.cantidad ?? ""}" onchange="ocrEditarFila(${i},'cantidad',this.value)"></td>
      <td><input type="number" min="0" step="any" class="form-control form-control-sm" value="${f.precio ?? ""}" onchange="ocrEditarFila(${i},'precio',this.value)">${equivARS}</td>
      <td><input type="number" min="0" step="any" class="form-control form-control-sm" placeholder="${escapeHtml(phVenta)}" value="${f.precioVenta ?? ""}" onchange="ocrEditarFila(${i},'precioVenta',this.value)"></td>
      <td>${f.matchCodigo ? '<span class="badge bg-success">Existente</span>' : '<span class="badge bg-warning text-dark">Nuevo</span>'}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="7" class="text-center text-muted py-3">No se pudo reconocer ningún producto. Probá con una foto más nítida (o un PDF) o corregí el texto de abajo.</td></tr>`;
  if (ocrModoPruebaActivo) ocrActualizarStats();
}

/* ---------------- MODO PRUEBA: comparar lo leído contra lo esperado (no escribe nada) ---------------- */

/** Muestra u oculta el panel de prueba y el botón "Agregar a la boleta" según el modo */
function ocrConfigurarModalSegunModo(prueba) {
  ocrModoPruebaActivo = !!prueba;
  tcPrellenarCampo("ocrTc");     // si la moneda de la revisión ya estaba en dólares, trae el tipo de cambio del día
  document.getElementById("ocrPanelPrueba").style.display = prueba ? "" : "none";
  document.getElementById("ocrBtnAgregar").style.display = prueba ? "none" : "";
  document.getElementById("ocrModalTitulo").textContent = prueba ? "Prueba de lectura (no se guarda nada)" : "Revisá lo que leyó la foto";
  document.getElementById("ocrDetalleTexto").open = !!prueba;
  document.getElementById("ocrComparacion").innerHTML = "";
}

function ocrActualizarStats() {
  const el = document.getElementById("ocrStats");
  if (!el || !ocrUltimaLectura) return;
  const lineas = (document.getElementById("ocrTextoCrudo").value || "").split(/\r?\n/).filter(l => l.trim()).length;
  const completas = ocrFilas.filter(f => f.cantidad && f.precio).length;
  const conf = ocrUltimaLectura.confianza;
  const colorConf = conf >= 85 ? "#15803d" : conf >= 65 ? "#b45309" : "#b91c1c";
  el.innerHTML =
    (conf === null
      ? `<b>Origen:</b> <span style="color:#15803d; font-weight:700;">texto digital del PDF (sin OCR)</span> · `
      : `<b>Confianza del OCR:</b> <span style="color:${colorConf}; font-weight:700;">${Math.round(conf)}%</span> · `) +
    `<b>Tiempo:</b> ${(ocrUltimaLectura.ms / 1000).toFixed(1)} s · ` +
    `<b>Líneas de texto:</b> ${lineas} · ` +
    `<b>Filas interpretadas:</b> ${ocrFilas.length} (${completas} con cantidad y costo)`;
}

/** Texto "esperado" -> filas. Acepta "nombre ; cant ; costo", "código ; nombre ; cant ; costo" o la línea tal cual de la lista */
function ocrParsearEsperado(texto) {
  return String(texto || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean).map(l => {
    if (/[;\t]/.test(l)) {
      const p = l.split(/[;\t]/).map(x => x.trim());
      const conCodigo = p.length >= 4;
      const nombre = conCodigo ? p[1] : p[0];
      const cant = ocrNumeroAR(conCodigo ? p[2] : p[1]);
      const costo = ocrNumeroAR(conCodigo ? p[3] : p[2]);
      return { codigo: conCodigo ? p[0] : "", nombre, cantidad: isNaN(cant) ? null : cant, precio: isNaN(costo) ? null : costo };
    }
    return ocrParsearLinea(l) || { codigo: "", nombre: l, cantidad: null, precio: null };
  });
}

/** Parecido entre dos nombres (0 a 1): tokens en común / tokens totales */
function ocrParecidoNombres(a, b) {
  const ta = new Set(ocrNorm(a).split(" ").filter(Boolean));
  const tb = new Set(ocrNorm(b).split(" ").filter(Boolean));
  if (!ta.size || !tb.size) return 0;
  let comunes = 0;
  ta.forEach(w => { if (tb.has(w)) comunes++; });
  return comunes / new Set([...ta, ...tb]).size;
}

function ocrCompararConEsperado() {
  const cont = document.getElementById("ocrComparacion");
  const esperadas = ocrParsearEsperado(document.getElementById("ocrEsperado").value);
  if (!esperadas.length) { cont.innerHTML = `<div class="text-muted" style="font-size:12.5px;">Pegá arriba lo que esperabas que leyera (una línea por producto).</div>`; return; }

  const usadas = new Set();
  const iguales = (a, b) => a !== null && b !== null && a !== undefined && b !== undefined && Math.abs(Number(a) - Number(b)) < 0.01;
  let okNombre = 0, okCant = 0, okCosto = 0, encontradas = 0;
  const celda = (ok, txt) => `<td style="background:${ok ? "#dcfce7" : "#fee2e2"};">${ok ? "✓" : "✗"} ${txt}</td>`;

  const filasHtml = esperadas.map(e => {
    let mejor = -1, mejorScore = 0;
    ocrFilas.forEach((f, i) => {
      if (usadas.has(i)) return;
      const s = ocrParecidoNombres(e.nombre, f.nombre);
      if (s > mejorScore) { mejorScore = s; mejor = i; }
    });
    if (mejor === -1 || mejorScore < 0.4) {
      return `<tr><td>${escapeHtml(e.nombre)}<br><span class="text-muted">${e.cantidad ?? "?"} × ${e.precio ?? "?"}</span></td>
        <td colspan="4" style="background:#fee2e2;">✗ No apareció en lo leído</td></tr>`;
    }
    usadas.add(mejor);
    encontradas++;
    const f = ocrFilas[mejor];
    const nOk = mejorScore >= 0.8, cOk = iguales(e.cantidad, f.cantidad), pOk = iguales(e.precio, f.precio);
    if (nOk) okNombre++; if (cOk) okCant++; if (pOk) okCosto++;
    return `<tr>
      <td>${escapeHtml(e.nombre)}<br><span class="text-muted">${e.cantidad ?? "?"} × ${e.precio ?? "?"}</span></td>
      <td>${escapeHtml(f.nombre)}<br><span class="text-muted">${f.cantidad ?? "?"} × ${f.precio ?? "?"}</span></td>
      ${celda(nOk, Math.round(mejorScore * 100) + "%")}
      ${celda(cOk, `${e.cantidad ?? "?"} / ${f.cantidad ?? "?"}`)}
      ${celda(pOk, `${e.precio ?? "?"} / ${f.precio ?? "?"}`)}
    </tr>`;
  }).join("");

  const sobrantes = ocrFilas.filter((f, i) => !usadas.has(i));
  const total = esperadas.length;
  cont.innerHTML = `
    <div class="mb-2" style="font-size:13px;">
      <b>Resumen:</b> encontrados <b>${encontradas}/${total}</b> · nombre correcto <b>${okNombre}/${total}</b> ·
      cantidad correcta <b>${okCant}/${total}</b> · costo correcto <b>${okCosto}/${total}</b> ·
      filas de más (leídas pero no esperadas) <b>${sobrantes.length}</b>
    </div>
    <div class="table-responsive"><table class="table table-sm" style="font-size:12.5px;">
      <thead><tr><th>Esperado</th><th>Leído</th><th>Nombre</th><th>Cant. (esp / leída)</th><th>Costo (esp / leído)</th></tr></thead>
      <tbody>${filasHtml}</tbody>
    </table></div>
    ${sobrantes.length ? `<div style="font-size:12.5px;"><b>Filas leídas que no estaban en lo esperado:</b><ul class="mb-0">${sobrantes.map(f => `<li>${escapeHtml(f.nombre)} <span class="text-muted">(${f.cantidad ?? "?"} × ${f.precio ?? "?"})</span></li>`).join("")}</ul></div>` : ""}`;

  console.log("[OCR prueba] esperado vs leído", { esperadas, leidas: ocrFilas });
}

/** Copia al portapapeles todo lo necesario para analizar una lectura (texto crudo, filas, confianza, tiempo) */
function ocrCopiarResultadoPrueba() {
  const payload = {
    archivo: ocrUltimaLectura?.archivo || null,
    confianzaOCR: ocrUltimaLectura ? Math.round(ocrUltimaLectura.confianza) : null,
    tiempoMs: ocrUltimaLectura?.ms ?? null,
    textoCrudo: document.getElementById("ocrTextoCrudo").value,
    filasInterpretadas: ocrFilas.map(f => ({ codigo: f.codigo || f.matchCodigo || "", nombre: f.nombre, cantidad: f.cantidad, costo: f.precio, existente: !!f.matchCodigo })),
    esperado: document.getElementById("ocrEsperado").value
  };
  const texto = JSON.stringify(payload, null, 2);
  const respaldo = () => {
    const ta = document.createElement("textarea");
    ta.value = texto; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); toast("Resultado copiado", "success"); } catch (e) { toast("No se pudo copiar", "error"); }
    ta.remove();
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(texto).then(() => toast("Resultado copiado", "success")).catch(respaldo);
  } else respaldo();
}

function ocrEditarFila(i, campo, valor) {
  const f = ocrFilas[i];
  if (!f) return;
  if (campo === "incluir") f.incluir = !!valor;
  else if (campo === "cantidad" || campo === "precio" || campo === "precioVenta") f[campo] = valor === "" ? null : Number(valor);
  else if (campo === "codigo") {
    f.codigo = String(valor).trim();
    const m = (productosAdminGlobal || []).find(p => String(p.CODIGO).trim() === f.codigo);
    f.matchCodigo = m ? String(m.CODIGO) : "";
  } else f[campo] = valor;
  ocrRenderTabla();
}

function ocrCambioMoneda() {
  const usd = document.getElementById("ocrMoneda").value === "USD";
  document.getElementById("ocrTcWrap").style.display = usd ? "" : "none";
  if (usd) tcPrellenarCampo("ocrTc");   // en dólares: se completa con el tipo de cambio del día (editable)
  ocrRenderTabla();   // actualiza el encabezado (ARS/USD) y el equivalente en pesos de cada costo
}

function ocrCerrar() {
  document.getElementById("ocrModalBackdrop").classList.remove("show");
}

function ocrReprocesarTexto() {
  const texto = document.getElementById("ocrTextoCrudo").value;
  // PDF sin cambios en el texto: se reusan las posiciones originales de cada celda (más preciso que releer el texto)
  if (ocrUltimaLectura?.filasCeldas && texto === ocrUltimaLectura.texto) {
    ocrFilas = ocrEnriquecerFilas(ocrFilasDesdeCeldas(ocrUltimaLectura.filasCeldas));
  } else {
    ocrFilas = ocrTextoAFilas(texto);
  }
  ocrRenderTabla();
}

/** Foto elegida -> OCR -> abre el modal de revisión */
async function ocrProcesarArchivo(ev) {
  const file = ev.target.files && ev.target.files[0];
  if (!file) return;
  const estado = document.getElementById("ocrEstado");
  const wrap = document.getElementById("ocrProgWrap");
  const barra = document.getElementById("ocrProg");
  const modoPrueba = !!document.getElementById("ocrModoPrueba")?.checked;
  wrap.style.display = "";
  barra.style.width = "0%";
  try {
    if (!productosAdminGlobal || !productosAdminGlobal.length) await cargarProductos();

    // ---- PDF: se lee el texto directamente (sin OCR): exacto y casi instantáneo ----
    if (file.type === "application/pdf" || /\.pdf$/i.test(file.name)) {
      estado.textContent = "⏳ Leyendo PDF...";
      const t0pdf = performance.now();
      const pdf = await ocrLeerPDF(file);
      if (!pdf.filasCeldas.length) {
        throw new Error("Este PDF no tiene texto (es una imagen escaneada). Sacale una captura de la tabla y subila como imagen.");
      }
      let filas = ocrEnriquecerFilas(ocrFilasDesdeCeldas(pdf.filasCeldas));
      if (!filas.length) filas = ocrTextoAFilas(pdf.texto);   // sin encabezado reconocible: se prueba línea por línea
      estado.textContent = "✓ Listo";
      ocrUltimaLectura = { ms: performance.now() - t0pdf, confianza: null, texto: pdf.texto, archivo: file.name, filasCeldas: pdf.filasCeldas };
      ocrConfigurarModalSegunModo(modoPrueba);
      if (modoPrueba) {
        document.getElementById("ocrImgOriginal").src = pdf.vista || "";
        document.getElementById("ocrImgProcesada").style.display = "none";
        console.log("[OCR prueba] texto del PDF:\n" + pdf.texto);
      }
      ocrFilas = filas;
      document.getElementById("ocrTextoCrudo").value = pdf.texto;
      ocrRenderTabla();
      document.getElementById("ocrModalBackdrop").classList.add("show");
      return;
    }

    // ---- Foto: OCR con Tesseract ----
    estado.textContent = "⏳ Cargando motor OCR...";
    await ocrCargarTesseract();
    estado.textContent = "⏳ Preparando imagen...";
    const canvas = await ocrPreprocesarImagen(file);

    const worker = await Tesseract.createWorker("spa", 1, {
      logger: m => {
        if (m.status === "recognizing text") {
          const pct = Math.round(m.progress * 100);
          barra.style.width = pct + "%";
          estado.textContent = "⏳ Leyendo... " + pct + "%";
        }
      }
    });
    await worker.setParameters({ tessedit_pageseg_mode: "6", preserve_interword_spaces: "1" });
    const t0 = performance.now();
    const { data } = await worker.recognize(canvas);
    const ms = performance.now() - t0;

    estado.textContent = "✓ Listo";
    ocrUltimaLectura = { ms, confianza: data.confidence || 0, texto: data.text, archivo: file.name };
    ocrConfigurarModalSegunModo(modoPrueba);
    if (modoPrueba) {
      // Fotos lado a lado: la original y la que realmente recibió el OCR (después del preprocesado)
      if (ocrImgOriginalURL) URL.revokeObjectURL(ocrImgOriginalURL);
      ocrImgOriginalURL = URL.createObjectURL(file);
      document.getElementById("ocrImgOriginal").src = ocrImgOriginalURL;
      document.getElementById("ocrImgProcesada").style.display = "";
      document.getElementById("ocrImgProcesada").src = canvas.toDataURL("image/jpeg", 0.85);
      console.log("[OCR prueba] texto crudo:\n" + data.text, { confianza: data.confidence, ms });
    }
    // Se intenta leer la tabla por columnas (nombre, código, cantidad, precio). Si no se reconoce el encabezado, se lee línea por línea.
    let filasFoto = [];
    let textoFoto = data.text;
    try {
      const progreso = p => { estado.textContent = "⏳ Leyendo tabla... " + Math.round(p * 100) + "%"; };
      const grilla = await ocrFotoPorCuadricula(worker, canvas, progreso);
      if (grilla.length) {
        filasFoto = ocrEnriquecerFilas(grilla);
        textoFoto = grilla.map(f => [f.codigo, f.nombre, f.cantidad, f.precio].join(" | ")).join("\n");
      } else {
        const filasCeldas = await ocrFotoAFilasCeldas(worker, canvas, data.words, progreso);
        filasFoto = ocrEnriquecerFilas(ocrFilasDesdeCeldas(filasCeldas));
        if (filasFoto.length) textoFoto = filasCeldas.map(f => f.map(c => c.str).join(" | ")).join("\n");
      }
    } catch (e) { console.warn("OCR por columnas:", e); }
    await worker.terminate();
    ocrFilas = filasFoto.length ? filasFoto : ocrTextoAFilas(data.text);
    ocrUltimaLectura.texto = textoFoto;
    document.getElementById("ocrTextoCrudo").value = textoFoto;
    ocrRenderTabla();
    document.getElementById("ocrModalBackdrop").classList.add("show");
  } catch (e) {
    console.error("OCR:", e);
    estado.textContent = "";
    toast(e.message || "No se pudo leer la imagen", "error");
  } finally {
    wrap.style.display = "none";
    ev.target.value = "";
  }
}

/** Lee una zona del canvas (agrandada) con las letras permitidas indicadas; psm 7 = una sola línea */
async function ocrLeerCelda(worker, canvas, x0, y0, x1, y1, lista, psm, mult) {
  const pad = 8;
  x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
  x1 = Math.min(canvas.width, Math.ceil(x1)); y1 = Math.min(canvas.height, Math.ceil(y1));
  const w = x1 - x0, h = y1 - y0;
  if (w < 4 || h < 4) return "";
  const k = Math.max(1, Math.min(5, (mult || 1) * 100 / h));
  const c = document.createElement("canvas");
  c.width = Math.round(w * k) + pad * 2; c.height = Math.round(h * k) + pad * 2;
  const ctx = c.getContext("2d");
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(canvas, x0, y0, w, h, pad, pad, Math.round(w * k), Math.round(h * k));
  // Si no lee nada (pasa con números cortos como "1"), se reintenta con otros modos de lectura
  for (const modo of [psm || 7, lista ? 8 : 6, 13]) {
    await worker.setParameters({ tessedit_char_whitelist: lista || "", tessedit_pageseg_mode: String(modo) });
    const { data } = await worker.recognize(c);
    const t = String(data.text || "").replace(/\s+/g, " ").trim();
    if (t) return t;
  }
  return "";
}
const OCR_LISTA_NUM = "0123456789$.,", OCR_LISTA_ENT = "0123456789", OCR_LISTA_COD = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_./";
/** Los números se leen 3 veces a distinta escala y gana la lectura que más se repite (una sola lectura confunde "142" con "442") */
async function ocrLeerCeldaNumerica(worker, canvas, x0, y0, x1, y1, soloEnteros) {
  const lista = soloEnteros ? OCR_LISTA_ENT : OCR_LISTA_NUM;
  const lecturas = [];
  for (const m of [1.3, 1, 0.8]) {
    const t = (await ocrLeerCelda(worker, canvas, x0, y0, x1, y1, lista, 7, m)).replace(/\s+/g, "");
    if (t) lecturas.push(t);
  }
  if (!lecturas.length) return "";
  const cuenta = {};
  lecturas.forEach(t => { cuenta[t] = (cuenta[t] || 0) + 1; });
  return lecturas.slice().sort((a, b) => cuenta[b] - cuenta[a])[0];   // sort estable: si empatan, queda la primera
}

/**
 * Foto de una tabla con bordes: con las líneas se arma la grilla y se lee cada celda por separado
 * (nombre, código, cantidad y precio; las demás columnas se ignoran). Mucho más confiable que leer toda la hoja de corrido.
 * Devuelve [] si no parece una tabla con grilla.
 */
async function ocrFotoPorCuadricula(worker, canvas, alProgreso) {
  const L = canvas._lineas;
  if (!L || L.xs.length < 5 || L.ys.length < 4) return [];
  // Bordes que la foto recortada no muestra: el borde superior del encabezado oscuro y el borde derecho de la imagen
  const xs = L.xs.slice(), ys = L.ys.slice();
  const b0 = (L.bandas || [])[0];      // encabezado oscuro: sus dos bordes son límites de fila aunque no se hayan detectado como línea
  if (b0) {
    for (let i = ys.length - 1; i >= 0; i--) if (ys[i] >= b0[0] - 4 && ys[i] <= b0[1] + 6) ys.splice(i, 1);
    ys.unshift(b0[1] + 2); ys.unshift(b0[0]);
  }
  if (canvas.width - xs[xs.length - 1] >= 20) xs.push(canvas.width - 2);
  if (xs[0] >= 20) xs.unshift(1);
  const tramos = (v, min) => { const r = []; for (let i = 0; i < v.length - 1; i++) if (v[i + 1] - v[i] >= min) r.push([v[i] + 3, v[i + 1] - 3]); return r; };
  const cols = tramos(xs, 14), filas = tramos(ys, 14);
  if (cols.length < 4 || filas.length < 2) return [];

  // 1) Encabezado: qué es cada columna
  const enc = filas[0];
  const roles = [];
  for (const [xa, xb] of cols) roles.push(xb - xa < 40 ? "" : ocrRolDeEncabezado(await ocrLeerCelda(worker, canvas, xa, enc[0], xb, enc[1], "", 7)));
  // El nombre es la columna más ancha: si su título salió ilegible, se deduce
  if (roles.indexOf("nombre") < 0) {
    let mejor = -1;
    cols.forEach(([xa, xb], i) => { if ((roles[i] === "" || roles[i] === "otro") && (mejor < 0 || xb - xa > cols[mejor][1] - cols[mejor][0])) mejor = i; });
    if (mejor >= 0) roles[mejor] = "nombre";
  }
  const iNom = roles.indexOf("nombre");
  if (iNom < 0) return [];
  // Una columna ilegible entre el nombre y la cantidad casi seguro es el código
  if (roles.indexOf("codigo") < 0) {
    const iCant = roles.indexOf("cantidad");
    const cand = roles.map((r, i) => (i > iNom && (iCant < 0 || i < iCant) && r === "otro") ? i : -1).filter(i => i >= 0);
    if (cand.length === 1) roles[cand[0]] = "codigo";
  }
  // El precio va pegado a la derecha de la cantidad: si su título salió ilegible, se deduce por la posición
  if (roles.indexOf("precio") < 0 && roles.indexOf("cantidad") >= 0) {
    const i = roles.indexOf("cantidad") + 1;
    if (i < roles.length - 1 && (roles[i] === "" || roles[i] === "otro")) roles[i] = "precio";
  }
  const idx = r => roles.indexOf(r);
  if (idx("cantidad") < 0 || idx("precio") < 0) return [];

  // 2) Cada fila de datos, solo las 4 columnas que importan
  const out = [];
  for (let n = 1; n < filas.length; n++) {
    const [ya, yb] = filas[n];
    const celda = (r, lista) => ocrLeerCelda(worker, canvas, cols[idx(r)][0], ya, cols[idx(r)][1], yb, lista, (yb - ya) > 70 && r === "nombre" ? 6 : 7);
    const nombre = await celda("nombre", "");
    const cantTxt = await ocrLeerCeldaNumerica(worker, canvas, cols[idx("cantidad")][0], ya, cols[idx("cantidad")][1], yb, true);
    const precioTxt = await ocrLeerCeldaNumerica(worker, canvas, cols[idx("precio")][0], ya, cols[idx("precio")][1], yb, false);
    const codigo = idx("codigo") >= 0 ? (await celda("codigo", OCR_LISTA_COD)).replace(/\s/g, "") : "";
    if (alProgreso) alProgreso(n / (filas.length - 1));
    if (!nombre && !cantTxt) continue;                   // fila vacía / total
    if (nombre.length < 3 && !cantTxt) continue;
    let cantidad = cantTxt ? parseInt(cantTxt, 10) : NaN;
    if (!(cantidad > 0) || cantidad > 99999 || /^0\d/.test(cantTxt)) cantidad = NaN;
    const precio = precioTxt ? ocrNumeroAR(precioTxt) : NaN;
    let nom = nombre;
    if (codigo && nom.endsWith(" " + codigo)) nom = nom.slice(0, -codigo.length).trim();
    out.push({ codigo, nombre: nom, cantidad: isNaN(cantidad) ? null : cantidad, precio: isNaN(precio) ? null : precio });
  }
  return out;
}

/** Palabras con posición (OCR de foto) -> filas de celdas, igual que las de un PDF; después se relee cada cantidad/precio celda por celda */
async function ocrFotoAFilasCeldas(worker, canvas, palabras, alProgreso) {
  const items = (palabras || []).filter(p => p.bbox && /[A-Za-z0-9$]/.test(p.text || "")).map(p => ({
    str: p.text.trim(), x: p.bbox.x0, w: p.bbox.x1 - p.bbox.x0,
    y: -(p.bbox.y0 + p.bbox.y1) / 2, h: p.bbox.y1 - p.bbox.y0, top: p.bbox.y0, bot: p.bbox.y1
  }));
  const filas = ocrAgruparItemsEnFilas(items);
  const iEnc = filas.findIndex(celdas => {
    const roles = celdas.map(c => ocrRolDeEncabezado(c.str));
    return roles.includes("nombre") && roles.filter(r => r === "cantidad" || r === "precio" || r === "codigo").length >= 2;
  });
  if (iEnc < 0) return filas;
  const cols = filas[iEnc].map(c => ({ rol: ocrRolDeEncabezado(c.str), x0: c.x0, x1: c.x1 }));
  const numericas = cols.filter(c => c.rol === "cantidad" || c.rol === "precio");
  // los límites de cada columna: desde el borde derecho de la anterior hasta el borde derecho propio (los números van a la derecha)
  const rango = col => {
    const i = cols.indexOf(col);
    const izq = i > 0 ? (cols[i - 1].x1 + col.x0) / 2 : col.x0 - 20;
    const der = i < cols.length - 1 ? (col.x1 + cols[i + 1].x0) / 2 : col.x1 + 30;
    return [Math.min(izq, col.x0 - 10), der];
  };
  const dataFilas = filas.slice(iEnc + 1).filter(f => f.y0 !== undefined && isFinite(f.y0));
  let hechos = 0;
  for (const fila of dataFilas) {
    for (const col of numericas) {
      const [xa, xb] = rango(col);
      const txt = await ocrLeerCeldaNumerica(worker, canvas, xa, fila.y0 - 3, xb, fila.y1 + 3, col.rol === "cantidad");
      // Se saca lo que se había leído mal en esa zona y se pone la lectura celda por celda
      for (let k = fila.length - 1; k >= 0; k--) { const cx = (fila[k].x0 + fila[k].x1) / 2; if (cx >= xa && cx <= xb) fila.splice(k, 1); }
      if (txt) fila.push({ str: txt, x0: col.x0, x1: col.x1 });
    }
    fila.sort((a, b) => a.x0 - b.x0);
    hechos++;
    if (alProgreso) alProgreso(hechos / dataFilas.length);
  }
  return filas;
}

/** Vuelca las filas tildadas al carrito de la boleta (sin tocar el backend todavía) */
function ocrAgregarABoleta() {
  // Red de seguridad: en modo prueba jamás se toca el carrito ni la base
  if (ocrModoPruebaActivo) { toast("Modo prueba: no se agrega nada a la boleta", "error"); return; }
  const moneda = document.getElementById("ocrMoneda").value === "USD" ? "USD" : "ARS";
  const tc = Number(document.getElementById("ocrTc").value || 0);
  if (moneda === "USD" && tc <= 0) { toast("Ingresá el tipo de cambio para precios en dólares", "error"); return; }

  const elegidas = ocrFilas.filter(f => f.incluir);
  if (!elegidas.length) { toast("No hay filas tildadas", "error"); return; }

  for (const f of elegidas) {
    if (!f.nombre || !String(f.nombre).trim()) { toast("Hay una fila tildada sin nombre de producto", "error"); return; }
    if (!f.cantidad || f.cantidad <= 0 || !f.precio || f.precio <= 0) {
      toast(`Falta cantidad o costo en "${f.nombre}"`, "error"); return;
    }
  }

  const prods = productosAdminGlobal || [];
  elegidas.forEach(f => {
    const codigoFila = String(f.codigo || f.matchCodigo || "").trim();
    const existente = codigoFila ? (prods.find(p => String(p.CODIGO).trim() === codigoFila) || null) : null;
    const precioARS = moneda === "USD" ? f.precio * tc : f.precio;
    // Precio de venta: SIEMPRE en pesos. Si lo cargaste, se usa tal cual. Si no lo cargaste y el producto es NUEVO con
    // costo en dólares, se manda el costo ya convertido a pesos como precio provisorio (si no, el sistema usaría el
    // número en dólares como si fueran pesos). En productos existentes, vacío = no se toca su precio actual.
    let precioVentaARS = (f.precioVenta !== null && f.precioVenta !== undefined && f.precioVenta > 0) ? f.precioVenta : null;
    if (precioVentaARS === null && !existente && moneda === "USD") precioVentaARS = Math.round(precioARS * 100) / 100;
    ipCarritoBoleta.push({
      codigo: existente ? String(existente.CODIGO) : codigoFila,
      producto: existente ? existente.PRODUCTO : f.nombre,
      categoria: existente ? (existente.CATEGORIA || "") : "",
      moneda,
      precio: f.precio,
      tipoCambio: moneda === "USD" ? tc : null,
      precioARS,
      precioVenta: precioVentaARS,   // siempre en pesos (ver arriba)
      cantidad: f.cantidad,
      esNuevo: !existente
    });
  });

  renderCarritoBoleta();
  ocrCerrar();
  toast(`${elegidas.length} producto(s) agregados a la boleta`, "success");
}

/** Manda TODA la boleta (cabecera + todos los productos del carrito) al backend en un solo documento */
async function guardarBoletaCompleta() {
  const origen = (document.getElementById("ipbOrigen")?.value) || "PROVEEDOR";
  const sinDeuda = origen === "DEPOSITO";
  const proveedor = document.getElementById("ipbProveedor").value.trim();
  const proveedorContacto = document.getElementById("ipbProveedorContacto").value.trim();
  const numeroBoleta = document.getElementById("ipbNumeroBoleta").value.trim();
  const observaciones = document.getElementById("ipbObservaciones").value.trim();
  const fecha = document.getElementById("ipbFecha").value || _hoyISO();

  if (!proveedor) {
    toast(sinDeuda ? "Ingresá de qué sucursal/depósito viene la mercadería" : "Ingresá el proveedor de esta boleta", "error");
    return;
  }
  if (ipCarritoBoleta.length === 0) {
    toast("Agregá al menos un producto a la boleta antes de guardar", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarBoleta");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  const payload = _payloadBoletaIngreso({ proveedor, proveedorContacto, numeroBoleta, observaciones, fecha, origen, sinDeuda });

  try {
    const res = await fetchAPI(API_URL, {
      method: "POST",
      body: JSON.stringify(Object.assign({ action: "registrarIngresoProducto" }, payload))
    });
    const data = await res.json();

    if (!data.success) {
      toast(data.message || data.error || "No se pudo guardar la boleta", "error");
      return;
    }

    toast(sinDeuda
      ? `Boleta guardada — ${ipCarritoBoleta.length} producto(s), ingreso interno sin deuda`
      : `Boleta guardada — ${ipCarritoBoleta.length} producto(s), total $${Number(data.totalBoleta || 0).toLocaleString("es-AR")}`, "success");

    try { localStorage.removeItem("vpos_cache_productosAdmin"); } catch(e) {}
    _limpiarBoletaIngreso();
    await cargarProductos();
    await cargarHistorialIngresos();

  } catch (error) {
    console.error("Error al guardar la boleta:", error);
    // App de escritorio: si se cortó la conexión, la boleta queda en la cola
    // con la misma clave, así no se duplica si el servidor llegó a guardarla.
    if (typeof window.encolarBoletaIngresoOffline === "function") {
      const ok = await window.encolarBoletaIngresoOffline(payload);
      if (ok) return;
    }
    toast("Error de conexión al guardar la boleta", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

/** Arma lo que se manda al backend para registrar la boleta. La clave
 *  idOperacionSync es la misma mientras no se vacíe la boleta: un
 *  reintento (o la cola offline) no vuelve a sumar el stock. */
let _ipClaveBoleta = null;
function _payloadBoletaIngreso(cab) {
  if (!_ipClaveBoleta) _ipClaveBoleta = "ING-" + Date.now() + "-" + Math.floor(Math.random() * 9000 + 1000);
  return {
    idOperacionSync: _ipClaveBoleta,
    rol: obtenerRolActual(),
    usuario: (sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("vendedor") || "ADMIN"),
    proveedor: cab.proveedor, proveedorContacto: cab.proveedorContacto, numeroBoleta: cab.numeroBoleta,
    observaciones: cab.observaciones, fecha: cab.fecha, origen: cab.origen, sinDeuda: cab.sinDeuda,
    items: ipCarritoBoleta.map(it => ({
      codigo: it.codigo, producto: it.producto, categoria: it.categoria,
      moneda: it.moneda, precio: it.precio, tipoCambio: it.tipoCambio, precioARS: it.precioARS,
      precioVenta: it.precioVenta, cantidad: it.cantidad
    }))
  };
}

/** Deja el formulario de boleta vacío después de guardarla (online u offline) */
function _limpiarBoletaIngreso() {
  _ipClaveBoleta = null;
  _ipItemEditando = null;
  ipCarritoBoleta = [];
  renderCarritoBoleta();
  document.getElementById("ipbOrigen").value = "PROVEEDOR";
  onCambioOrigenBoleta();
  document.getElementById("ipbNumeroBoleta").value = "";
  document.getElementById("ipbObservaciones").value = "";
}

/** Carga el listado de boletas de proveedor (una fila = una boleta, con total/pagado/saldo) desde el backend */
async function cargarHistorialIngresos() {
  const tbody = document.getElementById("tablaIngresosBody");
  try {
    const res = await fetchAPI(API_URL + "?action=boletasProveedores");
    const data = await res.json();
    ingresosProductosGlobal = data.boletas || [];
    try { localStorage.setItem("vpos_cache_boletasProveedores", JSON.stringify(ingresosProductosGlobal)); } catch (e) {}
    poblarDatalistProveedoresIngreso();
    filtrarHistorialIngresos();
  } catch (error) {
    console.error("Error al cargar boletas de proveedores:", error);
    // Sin conexión: se muestra el último historial descargado (los proveedores
    // también quedan para autocompletar la boleta nueva).
    let copia = null;
    try { copia = JSON.parse(localStorage.getItem("vpos_cache_boletasProveedores") || "null"); } catch (e) {}
    if (Array.isArray(copia)) {
      ingresosProductosGlobal = copia;
      poblarDatalistProveedoresIngreso();
      filtrarHistorialIngresos();
      toast("Sin conexión: se muestra el último historial de boletas guardado", "info");
      return;
    }
    if (tbody) tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">Error al cargar el historial</td></tr>`;
  }
}

/** Da formato compacto a un monto que puede tener parte en pesos y parte en dólares — ej: "$5.000 + US$120" — omitiendo la parte que esté en $0 */
function _formatoMontoBiMoneda(ars, usd) {
  const ars_ = Number(ars || 0);
  const usd_ = Number(usd || 0);
  const partes = [];
  if (ars_ || !usd_) partes.push("$" + ars_.toLocaleString("es-AR"));
  if (usd_) partes.push("US$" + usd_.toLocaleString("es-AR"));
  return partes.join(" + ");
}

function _badgeEstadoBoleta(estado) {
  if (estado === "SIN_DEUDA") return `<span class="badge bg-info text-dark">Interno · sin deuda</span>`;
  if (estado === "PAGADA") return `<span class="badge bg-success">Pagada</span>`;
  if (estado === "PARCIAL") return `<span class="badge bg-warning text-dark">Pago parcial</span>`;
  return `<span class="badge bg-danger">Pendiente</span>`;
}

function filtrarHistorialIngresos() {
  const tbody = document.getElementById("tablaIngresosBody");
  if (!tbody) return;

  const texto = (document.getElementById("ipBuscarHistorial")?.value || "").trim().toLowerCase();
  const proveedor = (document.getElementById("ipFiltroProveedor")?.value || "").trim().toLowerCase();
  const desde = document.getElementById("ipFiltroDesde")?.value || "";
  const hasta = document.getElementById("ipFiltroHasta")?.value || "";
  const estadoFiltro = document.getElementById("ipFiltroEstado")?.value || "";

  let lista = ingresosProductosGlobal || [];

  if (texto) {
    lista = lista.filter(i => String(i.NUMERO_BOLETA || "").toLowerCase().includes(texto));
  }
  if (proveedor) {
    lista = lista.filter(i => String(i.PROVEEDOR || "").toLowerCase().includes(proveedor));
  }
  if (desde) lista = lista.filter(i => String(i.FECHA) >= desde);
  if (hasta) lista = lista.filter(i => String(i.FECHA) <= hasta);
  if (estadoFiltro) lista = lista.filter(i => i.ESTADO === estadoFiltro);

  if (lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">Sin boletas registradas para este filtro</td></tr>`;
    return;
  }

  tbody.innerHTML = lista.map(i => `
    <tr>
      <td>${escapeHtml(i.FECHA || "—")}</td>
      <td class="mono">${escapeHtml(i.NUMERO_BOLETA || "—")}</td>
      <td>${escapeHtml(i.PROVEEDOR || "—")}</td>
      <td>${escapeHtml(i.PROVEEDOR_CONTACTO || "—")}</td>
      <td class="money">${_formatoMontoBiMoneda(i.TOTAL_ARS, i.TOTAL_USD)}</td>
      <td class="money" style="color:var(--green-600);">${_formatoMontoBiMoneda(i.PAGADO_ARS, i.PAGADO_USD)}</td>
      <td class="money" style="color:var(--red-500);">${_formatoMontoBiMoneda(i.SALDO_ARS, i.SALDO_USD)}</td>
      <td>${_badgeEstadoBoleta(i.ESTADO)}</td>
      <td><button class="btn btn-outline-primary btn-sm" onclick="abrirModalDetalleBoleta('${i.BOLETA_ID}')">Ver / Pagar</button></td>
    </tr>`).join("");
}

/** Abre el modal de detalle de una boleta: items, pagos, y el formulario para registrar un pago nuevo */
async function abrirModalDetalleBoleta(boletaId) {
  // Si este modal se abrió desde "Ver boletas" de un proveedor, esa
  // ventana queda abierta atrás — se oculta mientras se ve el detalle
  // para que no se superpongan dos modales a la vez, y se vuelve a
  // mostrar al cerrar este (ver cerrarModalDetalleBoleta).
  const modalProveedor = document.getElementById("boletasProveedorModalBackdrop");
  detalleBoletaAbiertoDesdeProveedor = !!(modalProveedor && modalProveedor.classList.contains("show"));
  if (detalleBoletaAbiertoDesdeProveedor) modalProveedor.classList.remove("show");

  document.getElementById("detalleBoletaModalBackdrop").classList.add("show");
  document.getElementById("detalleBoletaTitulo").textContent = "Cargando...";
  document.getElementById("pagoProveedorMonto").value = "";
  document.getElementById("pagoProveedorObservaciones").value = "";
  document.getElementById("pagoProveedorFormaPago").value = "EFECTIVO";
  const selMoneda = document.getElementById("pagoProveedorMoneda");
  if (selMoneda) selMoneda.value = "ARS";
  const btnPago = document.getElementById("btnRegistrarPagoProveedor");
  if (btnPago) { btnPago.disabled = false; btnPago.textContent = "Pagar"; }
  boletaProveedorIdActual = boletaId;

  try {
    const res = await fetchAPI(API_URL + "?action=detalleBoletaProveedor&boletaId=" + encodeURIComponent(boletaId));
    const data = await res.json();
    if (!data.success) {
      toast(data.message || "No se pudo cargar la boleta", "error");
      cerrarModalDetalleBoleta();
      return;
    }

    const b = data.boleta;
    detalleBoletaDatosActual = { boleta: b, items: data.items || [], pagos: data.pagos || [] };
    document.getElementById("detalleBoletaTitulo").textContent =
      `${b.PROVEEDOR}${b.NUMERO_BOLETA ? " — Boleta " + b.NUMERO_BOLETA : ""} (${b.FECHA})`;
    document.getElementById("detalleBoletaTotal").textContent = "$" + Number(b.TOTAL_ARS || 0).toLocaleString("es-AR");
    document.getElementById("detalleBoletaPagado").textContent = "$" + Number(b.PAGADO_ARS || 0).toLocaleString("es-AR");
    document.getElementById("detalleBoletaSaldo").textContent = "$" + Number(b.SALDO_ARS || 0).toLocaleString("es-AR");
    document.getElementById("detalleBoletaTotalUsd").textContent = b.TOTAL_USD > 0 ? "US$" + Number(b.TOTAL_USD).toLocaleString("es-AR") : "";
    document.getElementById("detalleBoletaPagadoUsd").textContent = b.PAGADO_USD > 0 ? "US$" + Number(b.PAGADO_USD).toLocaleString("es-AR") : "";
    document.getElementById("detalleBoletaSaldoUsd").textContent = b.SALDO_USD > 0 ? "US$" + Number(b.SALDO_USD).toLocaleString("es-AR") : "";

    const datosWrap = document.getElementById("detalleBoletaDatos");
    if (datosWrap) {
      const filasDatos = [
        ["N° de boleta / remito", b.NUMERO_BOLETA || "—"],
        ["Proveedor", b.PROVEEDOR || "—"],
        ["Contacto del proveedor", b.PROVEEDOR_CONTACTO || "—"],
        ["Fecha de recepción", b.FECHA || "—"],
        ["Cargada por", b.USUARIO || "—"],
        ["Observaciones", b.OBSERVACIONES || "—"]
      ];
      datosWrap.innerHTML = filasDatos.map(([label, valor]) =>
        `<div class="d-flex justify-content-between gap-2"><span class="text-muted">${escapeHtml(label)}</span><span class="text-end">${escapeHtml(String(valor))}</span></div>`
      ).join("");
    }

    const formWrap = document.getElementById("cardFormularioPagoProveedor");
    if (formWrap) formWrap.style.display = (b.SALDO_ARS > 0 || b.SALDO_USD > 0) ? "" : "none";
    // Si la boleta solo tiene deuda en una de las dos monedas, se arranca
    // con esa moneda seleccionada para no confundir con un saldo en $0.
    if (selMoneda) selMoneda.value = (b.SALDO_ARS <= 0 && b.SALDO_USD > 0) ? "USD" : "ARS";
    onCambioMonedaPagoProveedor();

    const items = data.items || [];
    const resumenEl = document.getElementById("detalleBoletaResumenItems");
    if (resumenEl) {
      const totalUnidades = items.reduce((acc, it) => acc + (Number(it.CANTIDAD) || 0), 0);
      const cantidadNuevos = items.filter(it => String(it.ES_PRODUCTO_NUEVO || "").toUpperCase() === "SI").length;
      resumenEl.textContent = `${items.length} producto${items.length === 1 ? "" : "s"} · ${totalUnidades.toLocaleString("es-AR")} unidad${totalUnidades === 1 ? "" : "es"}${cantidadNuevos > 0 ? ` · ${cantidadNuevos} nuevo${cantidadNuevos === 1 ? "" : "s"}` : ""}`;
    }

    document.getElementById("detalleBoletaItemsBody").innerHTML = items.map(it => {
      // El precio que compone el TOTAL de la boleta (lo que se le paga
      // al proveedor) es siempre PRECIO_INGRESO — nunca el de venta.
      // Los ingresos guardados antes de esta separación no tienen
      // PRECIO_INGRESO cargado; en ese caso se usa el precio de venta
      // como referencia, igual que hacía el sistema antes.
      const tieneCosto = it.PRECIO_INGRESO !== undefined && it.PRECIO_INGRESO !== "" && it.PRECIO_INGRESO !== null;
      const precioCosto = tieneCosto ? Number(it.PRECIO_INGRESO) : Number(it.PRECIO_NUEVO || it.PRECIO_ANTERIOR || 0);
      const precioVenta = Number(it.PRECIO_NUEVO || it.PRECIO_ANTERIOR || 0);
      const subtotalCosto = Number(it.CANTIDAD || 0) * precioCosto;
      return `
      <tr>
        <td class="mono">${escapeHtml(it.CODIGO || "—")}</td>
        <td>${escapeHtml(it.PRODUCTO || "—")}</td>
        <td class="money">${Number(it.CANTIDAD || 0).toLocaleString("es-AR")}</td>
        <td class="money">$${precioCosto.toLocaleString("es-AR")}${!tieneCosto ? ' <span class="text-muted" style="font-size:10px;" title="Este ingreso es de antes de separar costo/venta — se muestra el precio de venta como referencia">(ref.)</span>' : ""}</td>
        <td class="money" style="color:var(--slate-500, #64748b);">$${precioVenta.toLocaleString("es-AR")}</td>
        <td class="money">$${subtotalCosto.toLocaleString("es-AR")}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="6" class="text-center text-muted py-2">Sin productos</td></tr>`;

    const pagos = data.pagos || [];
    document.getElementById("detalleBoletaPagosBody").innerHTML = pagos.length
      ? pagos.map(p => `
        <tr>
          <td>${escapeHtml(p.FECHA || "—")}</td>
          <td class="money">${(p.MONEDA === "USD" ? "US$" : "$")}${Number(p.MONTO || 0).toLocaleString("es-AR")}</td>
          <td>${escapeHtml(p.MONEDA || "ARS")}</td>
          <td>${escapeHtml(p.FORMA_PAGO || "—")}</td>
          <td>${escapeHtml(p.OBSERVACIONES || "—")}</td>
        </tr>`).join("")
      : `<tr><td colspan="5" class="text-center text-muted py-2">Todavía no hay pagos registrados</td></tr>`;

  } catch (error) {
    console.error("Error al cargar el detalle de la boleta:", error);
    toast("Error de conexión al cargar la boleta", "error");
    cerrarModalDetalleBoleta();
  }
}

/** Al cambiar la moneda del formulario de pago, actualiza el hint con el saldo pendiente disponible en esa moneda */
function onCambioMonedaPagoProveedor() {
  const sel = document.getElementById("pagoProveedorMoneda");
  const hint = document.getElementById("pagoProveedorSaldoHint");
  const montoInput = document.getElementById("pagoProveedorMonto");
  if (!sel || !hint || !detalleBoletaDatosActual) return;
  const b = detalleBoletaDatosActual.boleta;
  const esUsd = sel.value === "USD";
  const saldo = esUsd ? Number(b.SALDO_USD || 0) : Number(b.SALDO_ARS || 0);
  const simbolo = esUsd ? "US$" : "$";
  hint.textContent = `Saldo pendiente en ${esUsd ? "dólares" : "pesos"}: ${simbolo}${saldo.toLocaleString("es-AR")}`;
  if (montoInput) montoInput.max = saldo > 0 ? saldo : "";
}

/** Arma e imprime (A4) el comprobante de la boleta de proveedor que está abierta en el modal de detalle */
function imprimirBoletaProveedor() {
  if (!detalleBoletaDatosActual) { toast("No hay una boleta cargada para imprimir", "error"); return; }

  const { boleta: b, items, pagos } = detalleBoletaDatosActual;
  const cfg = obtenerConfigNegocio();

  const printArea = document.getElementById("etiquetasPrintArea");
  const thermalFrame = document.getElementById("thermalPrintFrame");
  if (thermalFrame) thermalFrame.innerHTML = "";

  const filasItems = items.map(it => {
    const tieneCosto = it.PRECIO_INGRESO !== undefined && it.PRECIO_INGRESO !== "" && it.PRECIO_INGRESO !== null;
    const precioCosto = tieneCosto ? Number(it.PRECIO_INGRESO) : Number(it.PRECIO_NUEVO || it.PRECIO_ANTERIOR || 0);
    const subtotal = Number(it.CANTIDAD || 0) * precioCosto;
    return `
      <tr>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0;">${escapeHtml(it.CODIGO || "—")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0;">${escapeHtml(it.PRODUCTO || "—")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0; text-align:right;">${Number(it.CANTIDAD || 0).toLocaleString("es-AR")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0; text-align:right;">$${precioCosto.toLocaleString("es-AR")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0; text-align:right;">$${subtotal.toLocaleString("es-AR")}</td>
      </tr>`;
  }).join("") || `<tr><td colspan="5" style="padding:8px; text-align:center; color:#64748b;">Sin productos</td></tr>`;

  const filasPagos = pagos.length
    ? pagos.map(p => `
      <tr>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0;">${escapeHtml(p.FECHA || "—")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0; text-align:right;">${(p.MONEDA === "USD" ? "US$" : "$")}${Number(p.MONTO || 0).toLocaleString("es-AR")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0;">${escapeHtml(p.MONEDA || "ARS")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0;">${escapeHtml(p.FORMA_PAGO || "—")}</td>
        <td style="padding:5px 6px; border-bottom:1px solid #e2e8f0;">${escapeHtml(p.OBSERVACIONES || "—")}</td>
      </tr>`).join("")
    : `<tr><td colspan="5" style="padding:8px; text-align:center; color:#64748b;">Sin pagos registrados</td></tr>`;

  const totalUnidades = items.reduce((acc, it) => acc + (Number(it.CANTIDAD) || 0), 0);

  const html = `
    <div style="font-family:Arial, sans-serif; color:#1e293b; padding:10mm; font-size:12px;">
      <div style="display:flex; justify-content:space-between; align-items:flex-start; border-bottom:2px solid #1e293b; padding-bottom:10px; margin-bottom:14px;">
        <div>
          <div style="font-size:18px; font-weight:800;">${escapeHtml(cfg.nombre || "")}</div>
          <div style="color:#64748b;">Comprobante de ingreso de mercadería</div>
        </div>
        <div style="text-align:right;">
          <div style="font-size:15px; font-weight:700;">Boleta ${escapeHtml(b.NUMERO_BOLETA || "—")}</div>
          <div style="color:#64748b;">${escapeHtml(b.FECHA || "—")}</div>
        </div>
      </div>

      <table style="width:100%; border-collapse:collapse; margin-bottom:16px;">
        <tr>
          <td style="padding:3px 0; color:#64748b; width:150px;">Proveedor</td>
          <td style="padding:3px 0; font-weight:600;">${escapeHtml(b.PROVEEDOR || "—")}</td>
        </tr>
        <tr>
          <td style="padding:3px 0; color:#64748b;">Contacto</td>
          <td style="padding:3px 0;">${escapeHtml(b.PROVEEDOR_CONTACTO || "—")}</td>
        </tr>
        <tr>
          <td style="padding:3px 0; color:#64748b;">Cargada por</td>
          <td style="padding:3px 0;">${escapeHtml(b.USUARIO || "—")}</td>
        </tr>
        <tr>
          <td style="padding:3px 0; color:#64748b; vertical-align:top;">Observaciones</td>
          <td style="padding:3px 0;">${escapeHtml(b.OBSERVACIONES || "—")}</td>
        </tr>
      </table>

      <table style="width:100%; border-collapse:collapse; margin-bottom:10px;">
        <thead>
          <tr style="background:#f1f5f9;">
            <th style="padding:6px; text-align:left;">Código</th>
            <th style="padding:6px; text-align:left;">Producto</th>
            <th style="padding:6px; text-align:right;">Cant.</th>
            <th style="padding:6px; text-align:right;">Precio costo</th>
            <th style="padding:6px; text-align:right;">Subtotal</th>
          </tr>
        </thead>
        <tbody>${filasItems}</tbody>
      </table>
      <div style="text-align:right; color:#64748b; margin-bottom:16px;">${items.length} producto${items.length === 1 ? "" : "s"} · ${totalUnidades.toLocaleString("es-AR")} unidad${totalUnidades === 1 ? "" : "es"}</div>

      <div style="display:flex; justify-content:flex-end; gap:24px; border-top:2px solid #1e293b; padding-top:10px; margin-bottom:18px;">
        <div style="text-align:right;"><div style="color:#64748b;">Total</div><div style="font-size:15px; font-weight:800;">${_formatoMontoBiMoneda(b.TOTAL_ARS, b.TOTAL_USD)}</div></div>
        <div style="text-align:right;"><div style="color:#64748b;">Pagado</div><div style="font-size:15px; font-weight:800; color:#16a34a;">${_formatoMontoBiMoneda(b.PAGADO_ARS, b.PAGADO_USD)}</div></div>
        <div style="text-align:right;"><div style="color:#64748b;">Saldo</div><div style="font-size:15px; font-weight:800; color:#dc2626;">${_formatoMontoBiMoneda(b.SALDO_ARS, b.SALDO_USD)}</div></div>
      </div>

      <div style="font-weight:700; margin-bottom:6px;">Pagos registrados</div>
      <table style="width:100%; border-collapse:collapse;">
        <thead>
          <tr style="background:#f1f5f9;">
            <th style="padding:6px; text-align:left;">Fecha</th>
            <th style="padding:6px; text-align:right;">Monto</th>
            <th style="padding:6px; text-align:left;">Moneda</th>
            <th style="padding:6px; text-align:left;">Forma de pago</th>
            <th style="padding:6px; text-align:left;">Observaciones</th>
          </tr>
        </thead>
        <tbody>${filasPagos}</tbody>
      </table>

      <div style="margin-top:26px; color:#94a3b8; font-size:10.5px;">Impreso el ${new Date().toLocaleString("es-AR")}</div>
    </div>`;

  printArea.innerHTML = html;
  _setPrintPageSize("A4");
  setTimeout(() => {
    window.print();
  }, 150);
}

function cerrarModalDetalleBoleta() {
  document.getElementById("detalleBoletaModalBackdrop").classList.remove("show");
  boletaProveedorIdActual = null;
  detalleBoletaDatosActual = null;

  // Si este modal se había abierto desde "Ver boletas" de un proveedor,
  // esa ventana se ocultó temporalmente al abrir el detalle — se vuelve
  // a mostrar ahora, para volver al listado tal como estaba.
  if (detalleBoletaAbiertoDesdeProveedor) {
    const modalProveedor = document.getElementById("boletasProveedorModalBackdrop");
    if (modalProveedor) modalProveedor.classList.add("show");
    detalleBoletaAbiertoDesdeProveedor = false;
  }
}

/** Registra un pago contra la boleta abierta en el modal, y refresca el detalle y el listado */
async function registrarPagoProveedorForm() {
  if (!boletaProveedorIdActual) return;

  const monto = Number(document.getElementById("pagoProveedorMonto").value);
  if (!monto || monto <= 0) { toast("Ingresá un monto válido", "error"); return; }

  const btn = document.getElementById("btnRegistrarPagoProveedor");
  if (btn) { btn.disabled = true; btn.textContent = "Registrando..."; }

  try {
    const res = await fetchAPI(API_URL, {
      method: "POST",
      body: JSON.stringify({
        action: "registrarPagoProveedor",
        rol: obtenerRolActual(),
        boletaId: boletaProveedorIdActual,
        monto,
        moneda: document.getElementById("pagoProveedorMoneda")?.value || "ARS",
        formaPago: document.getElementById("pagoProveedorFormaPago").value,
        observaciones: document.getElementById("pagoProveedorObservaciones").value.trim(),
        usuario: (sessionStorage.getItem("nombreUsuario") || sessionStorage.getItem("vendedor") || "ADMIN")
      })
    });
    const data = await res.json();

    if (!data.success) {
      toast(data.message || "No se pudo registrar el pago", "error");
      if (btn) { btn.disabled = false; btn.textContent = "Pagar"; }
      return;
    }

    toast("Pago registrado", "success");
    await abrirModalDetalleBoleta(boletaProveedorIdActual);
    await cargarHistorialIngresos();
    if (document.getElementById("tablaProveedores")) cargarProveedores();

  } catch (error) {
    console.error("Error al registrar el pago a proveedor:", error);
    toast("Error de conexión al registrar el pago", "error");
    if (btn) { btn.disabled = false; btn.textContent = "Pagar"; }
  }
}

/** Exporta el listado de boletas (según los filtros aplicados) a PDF, con jsPDF + autoTable */
function exportarIngresosPDF() {
  try {
    const tabla = document.getElementById("tablaIngresosProductos");
    if (!tabla) return;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

    const nombreLocal = (obtenerConfigNegocio().nombre || "Reporte").toString();
    const desde = document.getElementById("ipFiltroDesde").value || "—";
    const hasta = document.getElementById("ipFiltroHasta").value || "—";

    doc.setFontSize(14);
    doc.text(`${nombreLocal} — Ingreso de Productos y Pagos a Proveedores`, 30, 30);
    doc.setFontSize(10);
    doc.setTextColor(110, 110, 110);
    doc.text(`Período: ${desde} a ${hasta}  ·  Generado: ${new Date().toLocaleString("es-AR")}`, 30, 46);

    doc.autoTable({
      html: tabla,
      startY: 58,
      theme: "grid",
      styles: { fontSize: 9, cellPadding: 5 },
      headStyles: { fillColor: [18, 32, 71], textColor: [255, 255, 255] }
    });

    doc.save(`Boletas_Proveedores_${desde}_a_${hasta}.pdf`);

  } catch (error) {
    console.error("Error al exportar boletas a PDF:", error);
    toast("No se pudo generar el PDF", "error");
  }
}

/* ===================================================================
   PROVEEDORES — resumen de boletas y saldos pendientes por proveedor
   (mismo espíritu que "Clientes a crédito", pero de este lado del
   mostrador: cuánto se le compró a cada proveedor, cuánto se le pagó
   y cuánto se le debe todavía).
=================================================================== */

let proveedoresGlobal = [];

async function cargarProveedores() {
  const cont = document.getElementById("tablaProveedores");
  try {
    const res = await fetchAPI(API_URL + "?action=proveedoresConSaldo");
    const data = await res.json();
    proveedoresGlobal = data.proveedores || [];
    filtrarProveedores();
  } catch (error) {
    console.error("Error al cargar proveedores:", error);
    if (cont) cont.innerHTML = `<div class="text-center text-muted py-5" style="font-size:14px;">Error al cargar los proveedores</div>`;
  }
}

function renderTablaProveedores(lista) {
  const cont = document.getElementById("tablaProveedores");
  if (!cont) return;

  if (!lista || lista.length === 0) {
    cont.innerHTML = `<div class="text-center text-muted py-5" style="font-size:14px;">No se encontraron proveedores</div>`;
    return;
  }

  cont.innerHTML = lista.map(p => {
    const saldoArs = Number(p.SALDO_ARS || 0);
    const saldoUsd = Number(p.SALDO_USD || 0);
    const tieneSaldo = saldoArs > 0.01 || saldoUsd > 0.01;
    const bordeClase = tieneSaldo ? "cliente-card-deuda" : "cliente-card-credito";

    return `
    <div class="pedido-card ${bordeClase}">
      <div class="pedido-card-top">
        <div style="flex:1; min-width:0;">
          <div class="pedido-card-cliente">${escapeHtml(p.PROVEEDOR)}</div>
          <div class="pedido-card-dir" style="margin-top:4px; display:flex; flex-wrap:wrap; gap:10px;">
            ${p.PROVEEDOR_CONTACTO ? `<span>📞 ${escapeHtml(p.PROVEEDOR_CONTACTO)}</span>` : ""}
            ${p.ULTIMA_FECHA ? `<span>Última boleta: ${escapeHtml(p.ULTIMA_FECHA)}</span>` : ""}
          </div>
        </div>
        <div class="text-end" style="flex-shrink:0;">
          <div class="pedido-card-total" style="font-size:15px;">${_formatoMontoBiMoneda(p.TOTAL_ARS, p.TOTAL_USD)}</div>
          <div style="font-size:11px;color:var(--slate-500);margin-top:1px;">${p.BOLETAS || 0} boleta${p.BOLETAS !== 1 ? "s" : ""}${p.BOLETAS_PENDIENTES ? ` · ${p.BOLETAS_PENDIENTES} pendiente${p.BOLETAS_PENDIENTES !== 1 ? "s" : ""}` : ""}</div>
          <div class="mt-1" style="font-size:12px;">
            <span class="cliente-dato-label">Saldo</span>
            ${tieneSaldo
              ? `<span style="color:var(--red-500);font-weight:700;">${_formatoMontoBiMoneda(saldoArs, saldoUsd)}</span>`
              : `<span class="text-muted" style="font-size:12px;">Sin deuda</span>`}
          </div>
        </div>
      </div>
      <div class="pedido-card-controls">
        <button class="btn btn-outline-secondary btn-sm" onclick="abrirModalBoletasProveedor('${escapeJsAttr(p.PROVEEDOR)}')">Ver boletas</button>
      </div>
    </div>`;
  }).join("");
}

function filtrarProveedores() {
  const input = document.getElementById("buscarProveedor");
  const soloConSaldo = document.getElementById("filtroSoloConSaldoProveedor");
  const termino = (input ? input.value : "").toLowerCase().trim();

  let filtrados = proveedoresGlobal;

  if (termino) {
    filtrados = filtrados.filter(p => String(p.PROVEEDOR || "").toLowerCase().includes(termino));
  }
  if (soloConSaldo && soloConSaldo.checked) {
    filtrados = filtrados.filter(p => Number(p.SALDO_ARS || 0) > 0.01 || Number(p.SALDO_USD || 0) > 0.01);
  }

  renderTablaProveedores(filtrados);
}

/** Abre el modal con todas las boletas de un proveedor puntual (con acceso directo a "Ver / Pagar" de cada una) */
async function abrirModalBoletasProveedor(proveedor) {
  document.getElementById("boletasProveedorModalBackdrop").classList.add("show");
  document.getElementById("boletasProveedorTitulo").textContent = proveedor;
  document.getElementById("boletasProveedorBody").innerHTML = `<tr><td colspan="7" class="text-center text-muted py-3">Cargando...</td></tr>`;

  try {
    const res = await fetchAPI(API_URL + "?action=boletasProveedores&proveedorExacto=" + encodeURIComponent(proveedor));
    const data = await res.json();
    const boletas = data.boletas || [];

    const totalArs = boletas.reduce((acc, b) => acc + Number(b.TOTAL_ARS || 0), 0);
    const totalUsd = boletas.reduce((acc, b) => acc + Number(b.TOTAL_USD || 0), 0);
    const pagadoArs = boletas.reduce((acc, b) => acc + Number(b.PAGADO_ARS || 0), 0);
    const pagadoUsd = boletas.reduce((acc, b) => acc + Number(b.PAGADO_USD || 0), 0);
    const saldoArs = boletas.reduce((acc, b) => acc + Number(b.SALDO_ARS || 0), 0);
    const saldoUsd = boletas.reduce((acc, b) => acc + Number(b.SALDO_USD || 0), 0);

    document.getElementById("boletasProveedorTotal").textContent = "$" + totalArs.toLocaleString("es-AR");
    document.getElementById("boletasProveedorPagado").textContent = "$" + pagadoArs.toLocaleString("es-AR");
    document.getElementById("boletasProveedorSaldo").textContent = "$" + saldoArs.toLocaleString("es-AR");
    document.getElementById("boletasProveedorTotalUsd").textContent = totalUsd > 0 ? "US$" + totalUsd.toLocaleString("es-AR") : "";
    document.getElementById("boletasProveedorPagadoUsd").textContent = pagadoUsd > 0 ? "US$" + pagadoUsd.toLocaleString("es-AR") : "";
    document.getElementById("boletasProveedorSaldoUsd").textContent = saldoUsd > 0 ? "US$" + saldoUsd.toLocaleString("es-AR") : "";

    document.getElementById("boletasProveedorBody").innerHTML = boletas.length
      ? boletas.map(b => `
        <tr>
          <td>${escapeHtml(b.FECHA || "—")}</td>
          <td class="mono">${escapeHtml(b.NUMERO_BOLETA || "—")}</td>
          <td class="money">${_formatoMontoBiMoneda(b.TOTAL_ARS, b.TOTAL_USD)}</td>
          <td class="money" style="color:var(--green-600);">${_formatoMontoBiMoneda(b.PAGADO_ARS, b.PAGADO_USD)}</td>
          <td class="money" style="color:var(--red-500);">${_formatoMontoBiMoneda(b.SALDO_ARS, b.SALDO_USD)}</td>
          <td>${_badgeEstadoBoleta(b.ESTADO)}</td>
          <td><button class="btn btn-outline-primary btn-sm" onclick="abrirModalDetalleBoleta('${b.BOLETA_ID}')">Ver / Pagar</button></td>
        </tr>`).join("")
      : `<tr><td colspan="7" class="text-center text-muted py-3">Este proveedor no tiene boletas cargadas</td></tr>`;

  } catch (error) {
    console.error("Error al cargar las boletas del proveedor:", error);
    toast("Error de conexión al cargar las boletas", "error");
  }
}

function cerrarModalBoletasProveedor() {
  document.getElementById("boletasProveedorModalBackdrop").classList.remove("show");
}

/* ===================================================================
   ROLES Y PERMISOS
   El login (login.html) guarda en sessionStorage: "admin"="true" (ya
   existía) y, además, "rol", "nombreUsuario" y "usuarioLogueado" —
   estos tres son nuevos. Instalaciones que todavía tengan un
   login.html viejo (sin mandar el rol) siguen funcionando: sin rol
   guardado, se asume "admin" (acceso total), igual que siempre.
=================================================================== */

/* ===================== AYUDA DE ATAJOS DEL POS =====================
   F1 o "?" abren el mismo panel que el botón "Atajos" del POS. */

function mostrarAyudaAtajosPOS() {
  const fondo = document.getElementById("shortcutsHelpBackdrop");
  if (!fondo) return;
  if (fondo.classList.contains("show")) { cerrarAyudaAtajosPOS(); return; }
  fondo.classList.add("show");
}

/** Cierra la ayuda si estaba abierta. Devuelve true si la cerró. */
function cerrarAyudaAtajosPOS() {
  const fondo = document.getElementById("shortcutsHelpBackdrop");
  if (!fondo || !fondo.classList.contains("show")) return false;
  fondo.classList.remove("show");
  const input = document.getElementById("posBusqueda");
  if (input) input.focus();
  return true;
}

/* ===================== ESTILO DEL PANEL (por computadora) =====================
   Se guarda en el navegador / app de ESTA caja (localStorage), no en la
   planilla: cada caja elige el suyo. El color del negocio (Catálogo Web →
   Color del sitio) se sigue aplicando encima. Ver estilos/estilos.css. */

const ESTILOS_PANEL = [
  { id: "clasico",   nombre: "Clásico",            desc: "El de siempre." },
  { id: "mostrador", nombre: "Mostrador claro",    desc: "Claro y amplio, con botones y buscador grandes." },
  { id: "turno",     nombre: "Turno largo",        desc: "Oscuro, cansa menos la vista. El total se ve como el visor de una caja." },
  { id: "recibo",    nombre: "Escáner y recibo",   desc: "Productos en lista, para escanear rápido. El ticket parece un recibo impreso." },
  { id: "centro",    nombre: "Ticket al centro",   desc: "El ticket ocupa la parte principal; los productos quedan al costado." },
  { id: "pizarron",  nombre: "Pizarrón de almacén",desc: "Pizarra verde con precios a tiza y el ticket en papel madera." },
  { id: "minimal",   nombre: "Minimalista",        desc: "Blanco y gris, mucho aire y un solo color para lo importante." },
  { id: "contraste", nombre: "Alto contraste",     desc: "Letras grandes, bordes marcados y foco amarillo. Colores fijos." },
  { id: "celeste",   nombre: "Celeste y blanco",   desc: "Colores fijos celeste y dorado. El ticket va en una barra abajo." },
  { id: "degrade",   nombre: "Degradé",            desc: "Oscuro con degradé en los colores del negocio." }
];
const ESTILOS_OSCUROS = { turno: true, pizarron: true, degrade: true };

function obtenerEstiloPanel() {
  try { return localStorage.getItem("veekEstiloPanel") || "clasico"; } catch (e) { return "clasico"; }
}

function aplicarEstiloPanel(id) {
  if (!ESTILOS_PANEL.some(e => e.id === id)) id = "clasico";
  const r = document.documentElement;
  r.setAttribute("data-estilo", id);
  if (ESTILOS_OSCUROS[id]) { r.setAttribute("data-oscuro", ""); r.setAttribute("data-bs-theme", "dark"); }
  else { r.removeAttribute("data-oscuro"); r.removeAttribute("data-bs-theme"); }
}

function elegirEstiloPanel(id) {
  aplicarEstiloPanel(id);
  try { localStorage.setItem("veekEstiloPanel", id); } catch (e) {}
  renderSelectorEstilosPanel();
  const nombre = (ESTILOS_PANEL.find(e => e.id === id) || {}).nombre || id;
  toast(`Estilo "${nombre}" aplicado en esta computadora`, "success");
}

function renderSelectorEstilosPanel() {
  const cont = document.getElementById("selectorEstilosPanel");
  if (!cont) return;
  const actual = obtenerEstiloPanel();
  cont.innerHTML = ESTILOS_PANEL.map(e => `
    <button type="button" class="estilo-opcion" aria-pressed="${e.id === actual}" onclick="elegirEstiloPanel('${e.id}')">
      <img src="estilos/vista-${e.id}.jpg" alt="" loading="lazy">
      <b>${escapeHtml(e.nombre)}${e.id === actual ? " ✓" : ""}</b>
      <small>${escapeHtml(e.desc)}</small>
    </button>`).join("");
}
document.addEventListener("DOMContentLoaded", renderSelectorEstilosPanel);

/**
 * Si el admin entró con la contraseña que viene de fábrica, se le avisa
 * arriba de todo hasta que la cambie (cualquiera que conozca VeekPOS la
 * conoce). Lo marca login.html con "debeCambiarPassword".
 */
function mostrarAvisoPasswordDeFabrica() {
  let debe = false;
  try { debe = sessionStorage.getItem("debeCambiarPassword") === "1"; } catch (e) {}
  if (!debe || document.getElementById("avisoPasswordFabrica")) return;
  const div = document.createElement("div");
  div.id = "avisoPasswordFabrica";
  div.setAttribute("role", "alert");
  div.style.cssText = "position:sticky;top:0;z-index:1040;background:#b91c1c;color:#fff;padding:10px 16px;margin:-4px 0 16px;" +
    "border-radius:10px;display:flex;gap:12px;align-items:center;justify-content:center;flex-wrap:wrap;font-weight:600";
  div.innerHTML = '<span>⚠️ Estás usando la contraseña de fábrica. Cualquiera que la conozca puede entrar a tu panel.</span>' +
    '<button type="button" class="btn btn-sm btn-light fw-bold">Cambiarla ahora</button>';
  div.querySelector("button").onclick = () => {
    mostrarSeccion("usuarios");
    if (typeof cargarUsuarios === "function") cargarUsuarios();
    setTimeout(() => {
      const campo = document.getElementById("credPasswordActual");
      if (campo) { campo.scrollIntoView({ behavior: "smooth", block: "center" }); campo.focus(); }
    }, 200);
  };
  // Dentro del área de contenido (no tapa la barra lateral)
  (document.querySelector(".content") || document.body).prepend(div);
}

/* =========================================================
   MODO DEPÓSITO
   El backend del depósito es otro Apps Script con el mismo código
   (MODO_NEGOCIO = DEPOSITO). El panel es el mismo: cambia lo que muestra.
   En la app de escritorio, cada uno tiene su perfil (ver perfiles.js).
========================================================= */
function esModoDeposito() {
  let m = "";
  try { m = sessionStorage.getItem("modoNegocio") || ""; } catch (e) {}
  if (!m && typeof configNegocioCache !== "undefined") m = configNegocioCache.modoNegocio || "";
  return String(m).toUpperCase() === "DEPOSITO";
}

/** Lo que no tiene sentido en el depósito (cupones y catálogo web) se oculta; se marca distinto para que no se confunda con el local */
function aplicarModoNegocio() {
  const dep = esModoDeposito();
  document.body.classList.toggle("modo-deposito", dep);
  let badge = document.getElementById("badgeModoDeposito");
  if (dep && !badge) {
    badge = document.createElement("div");
    badge.id = "badgeModoDeposito";
    badge.className = "badge-modo-deposito";
    badge.innerHTML = "<span>●</span> DEPÓSITO";
    badge.title = "Estás trabajando con el backend del depósito (ventas mayoristas, su stock y su caja)";
    const marca = document.querySelector(".sidebar-brand");
    if (marca) marca.after(badge);
  }
  if (badge) badge.style.display = dep ? "" : "none";

  const ocultar = (el, si) => {
    if (si) { el.dataset.ocultoDeposito = "1"; el.style.display = "none"; }
    else if (el.dataset.ocultoDeposito) { delete el.dataset.ocultoDeposito; el.style.display = ""; }
  };
  document.querySelectorAll('#navLinks a[data-target="cupones"], #bottomNavLinks a[data-target="cupones"]').forEach(a => ocultar(a, dep));
  document.querySelectorAll('.cfg-tab-btn[data-cfg-tab="catalogo"]').forEach(b => ocultar(b, dep));
  document.querySelectorAll('#navLinks a[data-target="pedidos"]').forEach(a => { a.textContent = dep ? "Pedidos" : "Pedidos web"; });
  document.querySelectorAll("[data-solo-deposito]").forEach(el => {
    const destino = el.getAttribute("data-target");
    el.style.display = dep && (!destino || seccionPermitidaParaRol(destino)) ? "" : "none";
  });
  if (dep && !/Depósito/.test(document.title)) document.title = "Depósito · " + document.title;
  if (!dep) document.title = document.title.replace(/^Depósito · /, "");
}

/* ---------- Configuración → Sistema: local y depósito (solo app de escritorio) ---------- */
async function cargarTarjetaPerfiles() {
  const card = document.getElementById("cardPerfiles");
  if (!card) return;
  const puente = window.veekpos;
  if (!puente || typeof puente.perfilInfo !== "function") { card.style.display = "none"; return; }
  card.style.display = "";
  try {
    const info = await puente.perfilInfo();
    const enDeposito = info.activo === "deposito";
    document.getElementById("perfilActualTexto").innerHTML = enDeposito
      ? "Esta PC está ahora en el <b>depósito</b>. Para volver al local, cerrá sesión e ingresá con la clave del local."
      : "Esta PC está ahora en el <b>local</b>. Para entrar al depósito, cerrá sesión e ingresá con la clave del depósito.";
    document.getElementById("perfilOtroLabel").textContent = enDeposito ? "Dirección del backend del local" : "Dirección del backend del depósito";
    document.getElementById("perfilOtroUrl").value = info.urls[info.otro] || "";
    document.getElementById("perfilOtroUrl").dataset.perfil = info.otro;
    document.getElementById("perfilOtroEstado").textContent = "";
  } catch (e) {
    document.getElementById("perfilActualTexto").textContent = "No se pudo leer la configuración de esta PC.";
  }
}

async function guardarUrlOtroPerfil(soloProbar) {
  const input = document.getElementById("perfilOtroUrl");
  const estado = document.getElementById("perfilOtroEstado");
  const url = input.value.trim();
  const perfil = input.dataset.perfil;
  const esperado = perfil === "deposito" ? "DEPOSITO" : "LOCAL";
  if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) {
    estado.innerHTML = '<span style="color:var(--red-500,#ef4444)">La dirección tiene que terminar en /exec (Apps Script → Implementar → Administrar implementaciones).</span>';
    return;
  }
  if (url) {
    estado.textContent = "Probando la conexión…";
    try {
      const r = await fetch(url + "?action=configuracionNegocio", { cache: "no-store" });
      const d = await r.json();
      const modo = (d && d.config && d.config.modoNegocio) || "LOCAL";
      if (modo !== esperado) {
        estado.innerHTML = '<span style="color:var(--red-500,#ef4444)">Ese backend responde, pero está configurado como ' + (modo === "DEPOSITO" ? "depósito" : "local") +
          (esperado === "DEPOSITO" ? '. En su Apps Script ejecutá una vez <b>configurarComoDeposito</b>.' : '.') + '</span>';
        return;
      }
      estado.innerHTML = '<span style="color:var(--green-600,#16a34a)">✓ Conectado con «' + escapeHtml((d.config && d.config.nombre) || "") + '».</span>';
    } catch (e) {
      estado.innerHTML = '<span style="color:var(--red-500,#ef4444)">No responde. Revisá la dirección y que la implementación sea "Cualquier usuario".</span>';
      return;
    }
  }
  if (soloProbar) return;
  const r = await window.veekpos.perfilFijarUrl(perfil, url);
  if (r && r.success) toast(url ? "Guardado. Ya se puede entrar con la clave del " + (perfil === "deposito" ? "depósito" : "local") + "." : "Se quitó la dirección.", "success");
  else toast((r && r.message) || "No se pudo guardar", "error");
}

const PERMISOS_POR_ROL = {
  admin: null, // null = acceso a todas las secciones
  vendedor: ["dashboard", "pos", "ventasPOS", "cierreCaja", "movimientosCaja", "pedidos", "clientes", "cotizaciones", "traspasos"],
  deposito: ["dashboard", "productos", "ingresoProductos", "proveedores", "reportes"] // en Reportes solo ve "Compras y reposición"
};

function obtenerRolActual() {
  return sessionStorage.getItem("rol") || "admin";
}

function seccionPermitidaParaRol(id) {
  const permitidas = PERMISOS_POR_ROL[obtenerRolActual()];
  if (!permitidas) return true; // admin, o rol desconocido -> no restringe
  return permitidas.includes(id);
}

/** Oculta del menú (desktop y mobile) las secciones que el rol logueado no puede ver */
function aplicarPermisosPorRol() {
  const rol = obtenerRolActual();
  const permitidas = PERMISOS_POR_ROL[rol];

  const nombre = sessionStorage.getItem("nombreUsuario");
  if (nombre) {
    const label = document.getElementById("sidebarLabelSub");
    if (label) label.textContent = nombre + (rol !== "admin" ? " · " + rol : "");
  }

  // Depósito: dentro de Reportes solo la pestaña de compras/reposición
  if (rol === "deposito") {
    document.querySelectorAll('#repTabs .rep-tab[data-tab="ventas"], #repTabs .rep-tab[data-tab="productos"]').forEach(b => b.style.display = "none");
  }

  if (!permitidas) return; // admin: ve todo, no se toca el menú

  document.querySelectorAll("#navLinks a[data-target], #bottomNavLinks a[data-target]").forEach(a => {
    const target = a.getAttribute("data-target");
    if (!permitidas.includes(target)) a.style.display = "none";
  });

  // Títulos de grupo del menú que quedaron sin ninguna opción visible: se ocultan
  document.querySelectorAll("#navLinks .nav-grupo").forEach(g => {
    let hayVisible = false;
    for (let el = g.nextElementSibling; el && !el.classList.contains("nav-grupo"); el = el.nextElementSibling) {
      if (el.tagName === "A" && el.style.display !== "none") { hayVisible = true; break; }
    }
    g.classList.toggle("vacio", !hayVisible);
  });

  // El rol vendedor ve la sección Pedidos y Movimientos de Caja, pero
  // no puede eliminar pedidos cancelados ni movimientos de caja — esos
  // botones se ocultan puntualmente (el backend también rechaza estas
  // acciones si llega el rol "vendedor", por si alguien fuerza el llamado).
  if (rol === "vendedor") {
    const btnPed = document.getElementById("btnEliminarPedidosCancelados");
    if (btnPed) btnPed.style.display = "none";
  }
}

/* ===================== GESTIÓN DE USUARIOS (solo admin) ===================== */

let usuariosGlobal = [];

async function cargarUsuarios() {
  const tbody = document.getElementById("tablaUsuarios");
  try {
    const res = await fetchAPI(API_URL + "?action=usuarios");
    const data = await res.json();
    usuariosGlobal = data.usuarios || [];

    if (usuariosGlobal.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Todavía no creaste ningún usuario adicional</td></tr>`;
      return;
    }

    const etiquetaRol = { admin: "Admin", vendedor: "Vendedor", deposito: "Depósito" };

    tbody.innerHTML = usuariosGlobal.map(u => `
      <tr>
        <td>${escapeHtml(u.NOMBRE || "—")}</td>
        <td class="mono">${escapeHtml(u.USERNAME || "—")}</td>
        <td>${etiquetaRol[u.ROL] || escapeHtml(u.ROL || "—")}</td>
        <td>${String(u.ACTIVO || "SI").toUpperCase() === "NO" ? '<span class="badge bg-secondary">Inactivo</span>' : '<span class="badge bg-success">Activo</span>'}</td>
        <td>
          <button class="btn btn-sm btn-outline-primary" onclick="abrirModalEditarUsuario('${u.USUARIO_ID}')">✏️</button>
          <button class="btn btn-sm btn-outline-danger" onclick="eliminarUsuarioClick('${u.USUARIO_ID}', '${escapeJsAttr(u.NOMBRE || "")}')">🗑️</button>
        </td>
      </tr>`).join("");

  } catch (error) {
    console.error("Error al cargar usuarios:", error);
    if (tbody) tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Error al cargar los usuarios</td></tr>`;
  }
}

function nuevoUsuario() {
  document.getElementById("usuarioModalTitle").textContent = "+ Nuevo Usuario";
  document.getElementById("umUsuarioId").value = "";
  document.getElementById("umNombre").value = "";
  document.getElementById("umUsername").value = "";
  document.getElementById("umPassword").value = "";
  document.getElementById("umPasswordHint").textContent = "*";
  document.getElementById("umRol").value = "vendedor";
  document.getElementById("umActivo").checked = true;
  document.getElementById("usuarioModalBackdrop").classList.add("show");
  setTimeout(() => document.getElementById("umNombre").focus(), 80);
}

function abrirModalEditarUsuario(usuarioId) {
  const u = usuariosGlobal.find(x => String(x.USUARIO_ID) === String(usuarioId));
  if (!u) { toast("No se encontró el usuario", "error"); return; }

  document.getElementById("usuarioModalTitle").textContent = "Editar Usuario";
  document.getElementById("umUsuarioId").value = u.USUARIO_ID;
  document.getElementById("umNombre").value = u.NOMBRE || "";
  document.getElementById("umUsername").value = u.USERNAME || "";
  document.getElementById("umPassword").value = "";
  document.getElementById("umPasswordHint").textContent = "(dejar vacío para no cambiarla)";
  document.getElementById("umRol").value = u.ROL || "vendedor";
  document.getElementById("umActivo").checked = String(u.ACTIVO || "SI").toUpperCase() !== "NO";
  document.getElementById("usuarioModalBackdrop").classList.add("show");
}

function cerrarModalUsuario() {
  document.getElementById("usuarioModalBackdrop").classList.remove("show");
}

async function guardarUsuarioForm() {
  const usuarioId = document.getElementById("umUsuarioId").value.trim();
  const nombre = document.getElementById("umNombre").value.trim();
  const username = document.getElementById("umUsername").value.trim();
  const password = document.getElementById("umPassword").value;
  const rol = document.getElementById("umRol").value;
  const activo = document.getElementById("umActivo").checked;

  if (!nombre || !username) {
    toast("Nombre y nombre de usuario son obligatorios", "error");
    return;
  }
  if (!usuarioId && !password) {
    toast("La contraseña es obligatoria para un usuario nuevo", "error");
    return;
  }

  const btn = document.getElementById("btnGuardarUsuario");
  const textoOriginal = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = "Guardando...";

  try {
    const body = usuarioId
      ? { action: "editarUsuario", usuarioId, nombre, username, password, rolNuevo: rol, activo }
      : { action: "crearUsuario", nombre, username, password, rolNuevo: rol };

    const res = await fetchAPI(API_URL, { method: "POST", body: JSON.stringify(body) });
    const data = await res.json();

    if (!data.success) {
      toast(data.message || "No se pudo guardar el usuario", "error");
      return;
    }

    toast(usuarioId ? "Usuario actualizado" : "Usuario creado", "success");
    cerrarModalUsuario();
    cargarUsuarios();

  } catch (error) {
    console.error("Error al guardar usuario:", error);
    toast("Error de conexión al guardar el usuario", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

function eliminarUsuarioClick(usuarioId, nombre) {
  confirmarAccion(
    `¿Eliminar el usuario "${nombre}"? Esta acción no se puede deshacer.`,
    async () => {
      try {
        const res = await fetchAPI(API_URL, { method: "POST", body: JSON.stringify({ action: "eliminarUsuario", usuarioId }) });
        const data = await res.json();
        if (!data.success) { toast(data.message || "No se pudo eliminar el usuario", "error"); return; }
        toast("Usuario eliminado", "success");
        cargarUsuarios();
      } catch (error) {
        console.error("Error al eliminar usuario:", error);
        toast("Error de conexión al eliminar el usuario", "error");
      }
    },
    "Eliminar usuario"
  );
}


/* =====================================================================
   ONDA AL HACER CLIC — en botones, menú y pestañas. Solo visual; si la
   computadora pide "reducir movimiento", no se agrega.
===================================================================== */
(function () {
  if (typeof window === "undefined" || !window.matchMedia) return;
  const reducir = window.matchMedia("(prefers-reduced-motion: reduce)");
  const SELECTOR = ".btn, .sidebar a, .bottom-nav a, .cfg-tab-btn, .rep-tab, .pos-action-btn, .cat-chip";
  document.addEventListener("pointerdown", ev => {
    if (reducir.matches || ev.button !== 0) return;
    const el = ev.target.closest(SELECTOR);
    if (!el || el.disabled || el.classList.contains("disabled")) return;
    // La onda se dibuja en una capa aparte encima del elemento (no dentro
    // de él), así no se mezcla con el contenido de botones que guardan y
    // restauran su texto (ej. "Guardando..." → "Guardar").
    const r = el.getBoundingClientRect();
    const cs = getComputedStyle(el);
    const capa = document.createElement("div");
    capa.className = "onda-capa";
    capa.style.cssText = `left:${r.left}px;top:${r.top}px;width:${r.width}px;height:${r.height}px;border-radius:${cs.borderRadius};color:${cs.color}`;
    const tam = Math.max(r.width, r.height);
    const onda = document.createElement("span");
    onda.className = "onda-clic";
    onda.style.width = onda.style.height = tam + "px";
    onda.style.left = (ev.clientX - r.left - tam / 2) + "px";
    onda.style.top = (ev.clientY - r.top - tam / 2) + "px";
    capa.appendChild(onda);
    document.body.appendChild(capa);
    setTimeout(() => capa.remove(), 600);
  }, { passive: true });
})();


/* =====================================================================
   CUPONES — descuentos con código para el catálogo web.
   Cada cupón: % o monto fijo, cupo de usos, vencimiento y compra mínima
   opcionales. Cada DNI/CUIT lo puede usar una vez (lo controla el
   backend al guardar el pedido, dentro del mismo lock).
===================================================================== */
let cuponesGlobal = [];
let reporteCuponesGlobal = null; // { cupones:[...], totales:{...} } del rango elegido

function _textoDescuentoCupon(c) {
  return c.tipo === "MONTO" ? "$" + Number(c.valor).toLocaleString("es-AR") : c.valor + "%";
}

function _estadoCupon(c) {
  if (!c.activo) return { txt: "Pausado", cls: "bg-secondary" };
  if (c.vencido) return { txt: "Vencido", cls: "bg-secondary" };
  if (c.programado) return { txt: "Programado", cls: "bg-info text-dark" };
  if (c.usosMax > 0 && c.usos >= c.usosMax) return { txt: "Agotado", cls: "bg-danger" };
  return { txt: "Activo", cls: "bg-success" };
}

async function cargarCupones() {
  const tbody = document.getElementById("tablaCupones");
  if (!tbody) return;
  try {
    const res = await fetchAPI(API_URL + "?action=cupones");
    const data = await res.json();
    if (!data.success) throw new Error(data.message || "No se pudieron cargar los cupones");
    cuponesGlobal = data.cupones || [];
    renderCupones();
    cargarReporteCupones();
  } catch (e) {
    console.error("Cupones:", e);
    tbody.innerHTML = `<tr><td colspan="8" class="text-center text-danger py-4">${escapeHtml(e.message || "Error de conexión")}</td></tr>`;
  }
}

function renderCupones() {
  const tbody = document.getElementById("tablaCupones");
  if (!cuponesGlobal.length) {
    tbody.innerHTML = `<tr><td colspan="8" class="text-center text-muted py-5">Todavía no hay cupones.<br><button class="btn btn-success btn-sm mt-2" onclick="abrirModalCupon()">+ Crear el primero</button></td></tr>`;
    return;
  }
  tbody.innerHTML = cuponesGlobal.map(c => {
    const est = _estadoCupon(c);
    const pct = c.usosMax > 0 ? Math.min(100, Math.round(c.usos * 100 / c.usosMax)) : 0;
    const cod = escapeJsAttr(c.codigo);
    return `<tr>
      <td><div class="cup-codigo">${escapeHtml(c.codigo)}</div>${c.nota ? `<div class="cup-nota">${escapeHtml(c.nota)}</div>` : ""}</td>
      <td><strong>${_textoDescuentoCupon(c)}</strong></td>
      <td><div class="cup-usos ${c.usos >= c.usosMax ? "lleno" : ""}">${c.usos} / ${c.usosMax} <span class="text-muted" style="font-size:12px;">· quedan ${Math.max(0, c.usosMax - c.usos)}</span>
        <div class="barra"><span style="width:${pct}%"></span></div></div></td>
      <td class="cup-vendido" data-cup-vendido="${escapeHtml(c.codigo)}">${_celdaVendidoCupon(c.codigo)}</td>
      <td>${_textoVigenciaCupon(c)}</td>
      <td>${c.minimo > 0 ? "$" + Number(c.minimo).toLocaleString("es-AR") : '<span class="text-muted">—</span>'}</td>
      <td><span class="badge ${est.cls}">${est.txt}</span></td>
      <td class="cup-acciones">
        <button class="btn btn-outline-secondary" title="Ver pedidos con este cupón" onclick="abrirReporteCupon('${cod}')">📊</button>
        <button class="btn btn-outline-secondary" title="Copiar código" onclick="copiarCodigoCupon('${cod}')">📋</button>
        <button class="btn btn-outline-secondary" title="${c.activo ? "Pausar" : "Activar"}" onclick="alternarCupon('${cod}')">${c.activo ? "⏸" : "▶"}</button>
        <button class="btn btn-outline-primary" title="Editar" onclick="abrirModalCupon('${cod}')">✏️</button>
        <button class="btn btn-outline-danger" title="Eliminar" onclick="eliminarCuponClick('${cod}')">🗑️</button>
      </td>
    </tr>`;
  }).join("");
}

const _pesos = n => "$" + Math.round(Number(n) || 0).toLocaleString("es-AR");

function _celdaVendidoCupon(codigo) {
  if (!reporteCuponesGlobal) return '<span class="text-muted">…</span>';
  const r = (reporteCuponesGlobal.cupones || []).find(x => x.codigo === codigo);
  if (!r || !r.usos) return '<span class="text-muted">—</span>';
  return `${_pesos(r.vendido)}<div class="sub">−${_pesos(r.descontado)} descontado</div>`;
}

async function cargarReporteCupones() {
  const desde = document.getElementById("cupRepDesde")?.value || "";
  const hasta = document.getElementById("cupRepHasta")?.value || "";
  try {
    const qs = "&desde=" + encodeURIComponent(desde) + "&hasta=" + encodeURIComponent(hasta);
    const res = await fetchAPI(API_URL + "?action=reporteCupones" + qs);
    const data = await res.json();
    if (!data.success) throw new Error(data.message || "No se pudo armar el reporte");
    reporteCuponesGlobal = data;
  } catch (e) {
    console.error("Reporte de cupones:", e);
    reporteCuponesGlobal = { cupones: [], totales: null, error: e.message || "Error de conexión" };
  }
  renderReporteCupones();
}

function limpiarFechasReporteCupones() {
  const d = document.getElementById("cupRepDesde"), h = document.getElementById("cupRepHasta");
  if (d) d.value = ""; if (h) h.value = "";
  cargarReporteCupones();
}

function renderReporteCupones() {
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  const t = reporteCuponesGlobal && reporteCuponesGlobal.totales;
  if (!t) {
    ["cupKpiUsos", "cupKpiVendido", "cupKpiDesc", "cupKpiProm"].forEach(id => set(id, "—"));
    set("cupKpiCanc", reporteCuponesGlobal && reporteCuponesGlobal.error ? reporteCuponesGlobal.error : "");
    set("cupKpiPct", "");
  } else {
    const validos = t.usos - t.cancelados;
    set("cupKpiUsos", String(t.usos));
    set("cupKpiCanc", t.cancelados ? `${t.cancelados} cancelado${t.cancelados === 1 ? "" : "s"} (no suman)` : "en el período elegido");
    set("cupKpiVendido", _pesos(t.vendido));
    set("cupKpiDesc", _pesos(t.descontado));
    const bruto = t.vendido + t.descontado;
    set("cupKpiPct", bruto > 0 ? `${(t.descontado * 100 / bruto).toFixed(1).replace(".", ",")}% de lo vendido sin descuento` : "");
    set("cupKpiProm", validos > 0 ? _pesos(t.vendido / validos) : "—");
  }
  document.querySelectorAll("[data-cup-vendido]").forEach(td => {
    td.innerHTML = _celdaVendidoCupon(td.getAttribute("data-cup-vendido"));
  });
}

function abrirReporteCupon(codigo) {
  const r = reporteCuponesGlobal && (reporteCuponesGlobal.cupones || []).find(x => x.codigo === codigo);
  const desde = document.getElementById("cupRepDesde")?.value, hasta = document.getElementById("cupRepHasta")?.value;
  const rango = desde || hasta ? ` · ${desde ? "desde " + _fechaCorta(desde) : ""} ${hasta ? "hasta " + _fechaCorta(hasta) : ""}` : "";
  document.getElementById("cuponRepTitulo").textContent = "Cupón " + codigo + rango;
  const kpis = document.getElementById("cuponRepKpis");
  const tbody = document.getElementById("cuponRepTabla");
  if (!r || !r.usos) {
    kpis.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-4">${reporteCuponesGlobal ? "Todavía nadie usó este cupón" + (rango ? " en ese período" : "") + "." : "Cargando…"}</td></tr>`;
  } else {
    const validos = r.usos - r.cancelados;
    kpis.innerHTML = `
      <div class="cup-kpi"><div class="lbl">Pedidos</div><div class="val">${r.usos}</div><div class="sub">${r.clientes} cliente${r.clientes === 1 ? "" : "s"}${r.cancelados ? ` · ${r.cancelados} cancelado${r.cancelados === 1 ? "" : "s"}` : ""}</div></div>
      <div class="cup-kpi"><div class="lbl">Vendido</div><div class="val">${_pesos(r.vendido)}</div><div class="sub">ya con descuento</div></div>
      <div class="cup-kpi"><div class="lbl">Descontado</div><div class="val">${_pesos(r.descontado)}</div></div>
      <div class="cup-kpi"><div class="lbl">Ticket promedio</div><div class="val">${validos > 0 ? _pesos(r.vendido / validos) : "—"}</div></div>`;
    tbody.innerHTML = r.pedidos.map(p => {
      const f = p.fecha ? new Date(p.fecha) : null;
      const fecha = f && !isNaN(f) ? f.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "2-digit" }) + " " + f.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" }) : "";
      const cancelado = String(p.estado).toUpperCase() === "CANCELADO";
      const id = escapeJsAttr(p.pedidoId);
      return `<tr class="${cancelado ? "cancelado" : ""}">
        <td style="white-space:nowrap;">${escapeHtml(fecha)}</td>
        <td>${p.pedidoId && p.estado !== "NO ENCONTRADO" ? `<a href="#" onclick="cerrarReporteCupon();abrirDetallePedido('${id}');return false;">${escapeHtml(p.pedidoId)}</a>` : escapeHtml(p.pedidoId || "—")}</td>
        <td>${escapeHtml(p.nombre || "")}<div class="text-muted" style="font-size:11.5px;">DNI/CUIT ${escapeHtml(p.dni || "—")}</div></td>
        <td class="text-end" style="white-space:nowrap;">${_pesos(p.total)}</td>
        <td class="text-end" style="white-space:nowrap;">−${_pesos(p.descuento)}</td>
        <td><span class="badge bg-light text-dark border">${escapeHtml(p.estado || "")}</span></td>
      </tr>`;
    }).join("");
  }
  document.getElementById("cuponRepModalBackdrop").classList.add("show");
}

function cerrarReporteCupon() {
  document.getElementById("cuponRepModalBackdrop").classList.remove("show");
}

const _fechaCorta = (iso) => iso ? iso.split("-").reverse().join("/") : "";

function _textoVigenciaCupon(c) {
  if (!c.desde && !c.vence) return '<span class="text-muted">Siempre</span>';
  const partes = [];
  if (c.desde) partes.push(`<span class="${c.programado ? "text-primary fw-semibold" : ""}">desde ${escapeHtml(_fechaCorta(c.desde))}</span>`);
  if (c.vence) partes.push(`<span class="${c.vencido ? "text-danger" : ""}">hasta ${escapeHtml(_fechaCorta(c.vence))}</span>`);
  return partes.join("<br>");
}

function abrirModalCupon(codigo) {
  const c = codigo ? cuponesGlobal.find(x => x.codigo === codigo) : null;
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v ?? ""; };
  document.getElementById("cuponModalTitulo").textContent = c ? "Editar cupón " + c.codigo : "Nuevo cupón";
  set("cupCodigoOriginal", c ? c.codigo : "");
  set("cupCodigo", c ? c.codigo : "");
  set("cupTipo", c ? c.tipo : "PORCENTAJE");
  set("cupValor", c ? c.valor : "");
  set("cupUsosMax", c ? c.usosMax : "");
  set("cupVence", c ? c.vence : "");
  set("cupDesde", c ? c.desde : "");
  set("cupMinimo", c && c.minimo ? c.minimo : "");
  set("cupNota", c ? c.nota : "");
  document.getElementById("cupActivo").checked = c ? !!c.activo : true;
  // un cupón ya usado no puede cambiar de código (el historial lo referencia)
  document.getElementById("cupCodigo").disabled = !!(c && c.usos > 0);
  document.getElementById("cupUsosInfo").textContent = c && c.usos > 0
    ? `Ya se usó ${c.usos} ${c.usos === 1 ? "vez" : "veces"}. Podés ampliar el cupo.`
    : "Se descuenta uno por cada pedido que lo use.";
  actualizarAyudaCupon();
  document.getElementById("cuponModalBackdrop").classList.add("show");
  setTimeout(() => document.getElementById(c ? "cupValor" : "cupCodigo")?.focus(), 80);
}

function cerrarModalCupon() {
  document.getElementById("cuponModalBackdrop").classList.remove("show");
}

/** Resumen en palabras de lo que va a hacer el cupón (ayuda a no equivocarse). */
function actualizarAyudaCupon() {
  const tipo = document.getElementById("cupTipo").value;
  document.getElementById("cupValorLabel").textContent = tipo === "MONTO" ? "Monto a descontar ($) *" : "Porcentaje *";
  const v = Number(document.getElementById("cupValor").value);
  const r = document.getElementById("cupResumen");
  if (!v) { r.textContent = ""; return; }
  const ej = 100000;
  const desc = tipo === "MONTO" ? Math.min(v, ej) : Math.round(ej * v / 100);
  const desde = document.getElementById("cupDesde")?.value, vence = document.getElementById("cupVence")?.value;
  const vig = desde && vence ? ` Vale del ${_fechaCorta(desde)} al ${_fechaCorta(vence)}.`
            : desde ? ` Empieza a valer el ${_fechaCorta(desde)}.`
            : vence ? ` Vale hasta el ${_fechaCorta(vence)}.` : "";
  r.textContent = `Ejemplo: en un pedido de $${ej.toLocaleString("es-AR")}, descuenta $${desc.toLocaleString("es-AR")} y el cliente paga $${(ej - desc).toLocaleString("es-AR")}.` + vig;
}

async function guardarCuponForm() {
  const val = id => document.getElementById(id).value.trim();
  const datos = {
    action: "guardarCupon",
    codigoOriginal: val("cupCodigoOriginal"),
    codigo: val("cupCodigo"),
    tipo: val("cupTipo"),
    valor: val("cupValor"),
    usosMax: val("cupUsosMax"),
    vence: val("cupVence"),
    desde: val("cupDesde"),
    minimo: val("cupMinimo") || 0,
    nota: val("cupNota"),
    activo: document.getElementById("cupActivo").checked
  };
  if (datos.codigo.length < 3) { toast("El código tiene que tener al menos 3 caracteres", "error"); return; }
  if (!(Number(datos.valor) > 0)) { toast("Poné cuánto descuenta el cupón", "error"); return; }
  if (!(Number(datos.usosMax) >= 1)) { toast("Poné la cantidad de usos (1 o más)", "error"); return; }
  if (datos.desde && datos.vence && datos.desde > datos.vence) { toast("La fecha de inicio no puede ser posterior al vencimiento", "error"); return; }
  const btn = document.getElementById("btnGuardarCupon");
  btn.disabled = true; btn.textContent = "Guardando...";
  try {
    const res = await fetchAPI(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(datos) }, { timeoutMs: 20000 });
    const data = await res.json();
    if (!data.success) { toast(data.message || "No se pudo guardar el cupón", "error"); return; }
    toast(`Cupón ${data.codigo} guardado`, "success");
    cerrarModalCupon();
    cargarCupones();
  } catch (e) {
    toast("Error de conexión al guardar el cupón", "error");
  } finally {
    btn.disabled = false; btn.textContent = "Guardar cupón";
  }
}

async function alternarCupon(codigo) {
  const c = cuponesGlobal.find(x => x.codigo === codigo);
  if (!c) return;
  try {
    const res = await fetchAPI(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({
      action: "guardarCupon", codigoOriginal: c.codigo, codigo: c.codigo, tipo: c.tipo, valor: c.valor, usosMax: c.usosMax,
      vence: c.vence, desde: c.desde, minimo: c.minimo, nota: c.nota, activo: !c.activo }) }, { timeoutMs: 20000 });
    const data = await res.json();
    if (!data.success) { toast(data.message || "No se pudo cambiar el cupón", "error"); return; }
    toast(c.activo ? `Cupón ${codigo} pausado` : `Cupón ${codigo} activado`, "success");
    cargarCupones();
  } catch (e) { toast("Error de conexión", "error"); }
}

async function eliminarCuponClick(codigo) {
  const c = cuponesGlobal.find(x => x.codigo === codigo);
  const extra = c && c.usos > 0 ? ` Ya se usó ${c.usos} ${c.usos === 1 ? "vez" : "veces"}; los pedidos que lo usaron no cambian.` : "";
  // Modal propio (no confirm() nativo: en Electron deja el teclado sin foco)
  confirmarAccion(`¿Eliminar el cupón ${codigo}?${extra}`, () => _eliminarCuponConfirmado(codigo), "Eliminar cupón");
}

async function _eliminarCuponConfirmado(codigo) {
  try {
    const res = await fetchAPI(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ action: "eliminarCupon", codigo }) }, { timeoutMs: 20000 });
    const data = await res.json();
    if (!data.success) { toast(data.message || "No se pudo eliminar", "error"); return; }
    toast(`Cupón ${codigo} eliminado`, "success");
    cargarCupones();
  } catch (e) { toast("Error de conexión", "error"); }
}

function copiarCodigoCupon(codigo) {
  try {
    navigator.clipboard.writeText(codigo).then(() => toast(`Código ${codigo} copiado`, "success"), () => toast(codigo, "success"));
  } catch (e) { toast(codigo, "success"); }
}
