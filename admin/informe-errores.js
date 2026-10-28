/* =====================================================================
   INFORMES DE ERROR (panel web y app de escritorio)
   - Guarda en memoria los últimos errores de la página (errores de
     JavaScript, promesas rechazadas, console.error y llamadas al
     servidor que fallaron) y los últimos PASOS de la persona (secciones,
     botones, teclas F, avisos de error que vio), para adjuntarlos.
   - "Informar un problema": la persona cuenta qué pasó y se manda
     por correo a soporte (backend: action=reportarError). Puede adjuntar
     una captura de pantalla (opcional, desmarcada por defecto).
   - En la app de escritorio se agrega el estado de la cola de
     sincronización (pendientes y operaciones con error).
   - Los errores inesperados de JavaScript se avisan solos (agrupados;
     el backend también los agrupa y tiene tope diario).
   Se carga ANTES de admin.js para capturar también sus errores.
===================================================================== */
// Versión del panel: va en cada informe para saber si ese navegador o
// esa caja tiene el panel al día. Se sube junto con VERSION_BACKEND y
// desktop/package.json (la prueba del backend avisa si no coinciden).
const VERSION_PANEL = "2.4.1";

(function () {
  const MAX_REGISTRO = 30;
  const MAX_PASOS = 25;
  const registro = [];            // { hora, tipo, mensaje, origen, pila }
  const pasos = [];               // { hora, paso }
  const yaAvisados = new Set();   // mensajes automáticos ya mandados en esta sesión
  let automaticosEnviados = 0;
  let pendienteAuto = null;
  let capturaLista = null;        // Promise<string base64 jpeg|null> de la captura actual

  const hora = () => new Date().toLocaleTimeString("es-AR", { hour12: false });
  const corto = (t, n) => { t = String(t == null ? "" : t); return t.length > n ? t.slice(0, n) + "…" : t; };
  const limpio = (t) => String(t || "").replace(/\s+/g, " ").trim();

  // Ruido que no son fallas del sistema (extensiones del navegador, red caída, etc.)
  const RUIDO = /ResizeObserver loop|^Script error\.?$|AbortError|The user aborted|Failed to fetch|NetworkError|Load failed|Network request failed|signal is aborted|chrome-extension:|moz-extension:|_AutofillCallbackHandler|webkit\.messageHandlers|__gCrWeb|instantSearchSDKJSBridge|WeixinJSBridge|zaloJSV2|TTNativeBridge|FBNavigation|fbq is not defined|__firefox__|SymBrowser_|ucbrowser|_pcmBridgeCallbackHandler|vivoNewsDetailPage|getReadModeExtract|feedNativeBridge/i;

  function anotar(tipo, mensaje, origen, pila) {
    registro.push({ hora: hora(), tipo, mensaje: corto(mensaje, 600), origen: corto(origen, 200), pila: corto(pila, 1500) });
    if (registro.length > MAX_REGISTRO) registro.shift();
  }

  function anotarPaso(texto) {
    texto = corto(limpio(texto), 140);
    if (!texto) return;
    const ultimo = pasos[pasos.length - 1];
    if (ultimo && ultimo.paso === texto) { ultimo.veces = (ultimo.veces || 1) + 1; ultimo.hora = hora(); return; }
    pasos.push({ hora: hora(), paso: texto });
    if (pasos.length > MAX_PASOS) pasos.shift();
  }

  /* ---------- Pasos: qué tocó la persona (sin lo que escribe) ---------- */
  document.addEventListener("click", (e) => {
    try {
      const el = e.target && e.target.closest && e.target.closest("button, a, [role=menuitem], [onclick], select, input[type=checkbox], input[type=radio], .cfg-tab-btn");
      if (!el || el.closest("#informeModalBackdrop")) return;
      // Dentro de una tabla o lista de datos no se copia el texto de la
      // fila (puede ser un cliente o un producto): solo qué botón fue.
      const enDatos = !!el.closest("tbody, .sync-err-item, .cart-item-row, .pos-grid, .product-tile, .lista-clientes");
      const nombreBoton = limpio(el.getAttribute("aria-label") || el.getAttribute("title") || el.textContent).slice(0, 60);
      let texto;
      if (el.matches("a[data-target]")) texto = "Menú → " + limpio(el.textContent);
      else if (el.matches("input[type=checkbox], input[type=radio]")) texto = (el.checked ? "Tildó" : "Destildó") + " casilla" + (el.id ? " #" + el.id : (enDatos ? " de una fila" : ""));
      else if (el.matches("select")) texto = "Abrió la lista" + (el.id ? " #" + el.id : "");
      else if (enDatos) texto = "Botón «" + (el.getAttribute("title") || el.getAttribute("data-accion") || nombreBoton.split(" ")[0] || "?") + "» en una fila";
      else texto = (el.tagName === "A" ? "Link" : "Botón") + " «" + (nombreBoton || el.id || "?") + "»";
      anotarPaso(texto);
    } catch (x) {}
  }, true);

  document.addEventListener("change", (e) => {
    const el = e.target;
    if (el && el.tagName === "SELECT" && el.id && !el.closest("#informeModalBackdrop")) {
      const opt = el.options[el.selectedIndex];
      anotarPaso("Eligió en #" + el.id + ": " + corto(limpio(opt ? opt.textContent : el.value), 40));
    }
  }, true);

  document.addEventListener("keydown", (e) => {
    if (/^F\d{1,2}$/.test(e.key)) anotarPaso("Tecla " + e.key);
  }, true);

  // Secciones y avisos de error en pantalla (se engancha cuando admin.js ya cargó)
  document.addEventListener("DOMContentLoaded", () => {
    if (typeof window.mostrarSeccion === "function" && !window.mostrarSeccion.__conPasos) {
      const original = window.mostrarSeccion;
      window.mostrarSeccion = function (id) { anotarPaso("Sección: " + id); return original.apply(this, arguments); };
      window.mostrarSeccion.__conPasos = true;
    }
    if (typeof window.toast === "function" && !window.toast.__conPasos) {
      const originalToast = window.toast;
      window.toast = function (mensaje, tipo) {
        if (tipo === "error") anotarPaso("Aviso de error: " + corto(mensaje, 110));
        return originalToast.apply(this, arguments);
      };
      window.toast.__conPasos = true;
    }
  });

  function programarAutomatico(mensaje, origen, pila) {
    if (RUIDO.test(mensaje) || RUIDO.test(origen || "")) return;
    const clave = String(mensaje).slice(0, 160);
    if (yaAvisados.has(clave) || automaticosEnviados >= 5) return;
    yaAvisados.add(clave);
    // Se espera un poco: si salen varios errores juntos, van en un solo informe.
    if (pendienteAuto) { pendienteAuto.extra++; return; }
    pendienteAuto = { titulo: clave, extra: 0 };
    setTimeout(async () => {
      const p = pendienteAuto; pendienteAuto = null;
      automaticosEnviados++;
      enviarAlServidor({
        automatico: true,
        titulo: p.titulo + (p.extra ? ` (+${p.extra} más)` : ""),
        descripcion: "Error inesperado en el panel (aviso automático).",
        contexto: datosDelEquipo(),
        errores: registro.slice(-12),
        pasos: pasos.slice(-15),
        cola: await estadoColaCaja()
      }).catch(() => {});
    }, 4000);
  }

  window.addEventListener("error", (e) => {
    // Errores de carga de imágenes/recursos: no son de código
    if (!e || (e.target && e.target !== window && !e.error)) return;
    const msg = e.message || (e.error && e.error.message) || "Error";
    const origen = (e.filename ? e.filename.split("/").pop() : "") + (e.lineno ? ":" + e.lineno + ":" + (e.colno || 0) : "");
    const pila = e.error && e.error.stack || "";
    anotar("error", msg, origen, pila);
    programarAutomatico(msg, origen, pila);
  }, true);

  window.addEventListener("unhandledrejection", (e) => {
    const r = e && e.reason;
    const msg = (r && (r.message || r.toString && r.toString())) || "Promesa rechazada";
    anotar("promesa", msg, "", r && r.stack || "");
    programarAutomatico(msg, "", r && r.stack || "");
  });

  // console.error: solo se anota (el sistema lo usa también para avisos esperados)
  const errorOriginal = console.error.bind(console);
  console.error = function (...args) {
    try {
      anotar("console", args.map(a => a instanceof Error ? (a.message + (a.stack ? "\n" + a.stack : "")) : (typeof a === "object" ? (() => { try { return JSON.stringify(a); } catch (x) { return String(a); } })() : String(a))).join(" "), "", "");
    } catch (x) {}
    return errorOriginal(...args);
  };

  /** Lo que se adjunta del equipo (nada de contraseñas ni datos de clientes). */
  function datosDelEquipo() {
    const seccionVisible = [...document.querySelectorAll(".seccion")].find(s => s.style.display === "block");
    const enApp = !!(window.veekpos || window.posOffline);
    let estilo = "";
    try { estilo = localStorage.getItem("veekEstiloPanel") || ""; } catch (e) {}
    return {
      seccion: seccionVisible ? seccionVisible.id : "",
      app: enApp ? "App de escritorio" + (window.__versionAppVeek ? " " + window.__versionAppVeek : "") : "Panel web",
      versionPanel: VERSION_PANEL,
      navegador: navigator.userAgent,
      pantalla: `${window.innerWidth}×${window.innerHeight} (x${window.devicePixelRatio || 1})`,
      conexion: navigator.onLine ? "con internet" : "SIN internet",
      rol: (() => { try { return sessionStorage.getItem("rol") || ""; } catch (e) { return ""; } })(),
      estiloPanel: estilo,
      pagina: location.pathname.split("/").pop() || "index.html",
      servidor: (() => { try { return new URL(API_URL).host; } catch (e) { return ""; } })()
    };
  }

  /** App de escritorio: cuántas operaciones esperan subir y cuáles tienen error. */
  async function estadoColaCaja() {
    const po = window.posOffline;
    if (!po || typeof po.obtenerEstadoSync !== "function") return null;
    try {
      const est = await po.obtenerEstadoSync();
      const salida = { pendientes: est.pendientes || 0, conError: est.conError || 0, ultimaSubida: est.ultimaSyncOk || "" };
      if (typeof po.listarErroresSync === "function" && salida.conError > 0) {
        const lista = await po.listarErroresSync();
        salida.errores = lista.slice(0, 10).map(op => {
          const p = op.payload || {};
          const ref = p.codigoOriginal || p.CODIGO || p.codigo || (p.total !== undefined ? "$" + p.total : "") || p.clienteId || "";
          return { tipo: op.tabla, ref: corto(ref, 60), intentos: op.intentos, desde: op.creadoEn, error: corto(op.error, 300) };
        });
      }
      try { if (typeof po.obtenerRolRed === "function") salida.redLocal = await po.obtenerRolRed(); } catch (e) {}
      return salida;
    } catch (e) {
      return { error: "No se pudo leer la cola: " + (e && e.message) };
    }
  }

  async function enviarAlServidor(cuerpo) {
    if (typeof API_URL === "undefined" || !API_URL) throw new Error("El panel todavía no está conectado al servidor.");
    const f = typeof fetchAPI === "function" ? fetchAPI : (u, o) => fetch(u, o);
    const res = await f(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(Object.assign({ action: "reportarError" }, cuerpo))
    }, { timeoutMs: 45000 });
    return res.json();
  }

  /* ---------- Captura de pantalla (opcional) ---------- */
  /** Achica a lo sumo 1400 px de ancho y pasa a JPEG. Devuelve base64 sin prefijo. */
  function canvasAJpeg(canvas) {
    const max = 1400;
    let c = canvas;
    if (canvas.width > max) {
      c = document.createElement("canvas");
      c.width = max; c.height = Math.round(canvas.height * max / canvas.width);
      c.getContext("2d").drawImage(canvas, 0, 0, c.width, c.height);
    }
    return c.toDataURL("image/jpeg", 0.72).split(",")[1] || null;
  }

  // html2canvas-pro (no el original): entiende color-mix() y las demás
  // funciones de color modernas que usan los estilos del panel.
  const fnHtml2canvas = () => (typeof window.html2canvas === "function" ? window.html2canvas : (window.html2canvas && window.html2canvas.default));

  function cargarHtml2canvas() {
    if (fnHtml2canvas()) return Promise.resolve();
    return new Promise((ok, mal) => {
      const s = document.createElement("script");
      s.src = "vendor/html2canvas-pro.min.js";
      s.onload = () => ok(); s.onerror = () => mal(new Error("no se pudo cargar html2canvas"));
      document.head.appendChild(s);
    });
  }

  /**
   * App de escritorio: foto real de la ventana (antes de abrir esta
   * ventanita). Navegador: html2canvas dibuja la página sin el modal.
   */
  function prepararCaptura(enApp) {
    if (enApp && window.veekpos && typeof window.veekpos.capturarVentana === "function") {
      return window.veekpos.capturarVentana().catch(() => null);
    }
    return cargarHtml2canvas().then(() => fnHtml2canvas()(document.body, {
      backgroundColor: getComputedStyle(document.body).backgroundColor || "#ffffff",
      scale: 1, logging: false, useCORS: true,
      x: window.scrollX, y: window.scrollY, width: window.innerWidth, height: window.innerHeight,
      windowWidth: document.documentElement.clientWidth, windowHeight: window.innerHeight,
      ignoreElements: (el) => el.id === "informeModalBackdrop" || el.id === "toastStack",
      // La copia vuelve a correr las animaciones de entrada desde cero (se
      // vería todo transparente): se apagan solo en la copia.
      onclone: (doc) => {
        const st = doc.createElement("style");
        st.textContent = "*,*::before,*::after{animation:none!important;transition:none!important}";
        doc.head.appendChild(st);
      }
    })).then(canvasAJpeg).catch((e) => { console.warn("captura:", e); return null; });
  }

  function mostrarVistaCaptura(base64) {
    const vista = document.getElementById("informeCapturaVista");
    const estado = document.getElementById("informeCapturaEstado");
    if (!vista) return;
    if (base64) {
      vista.src = "data:image/jpeg;base64," + base64;
      vista.style.display = "block";
      if (estado) estado.textContent = "Así se va a ver (" + Math.round(base64.length * 0.75 / 1024) + " KB).";
    } else {
      vista.style.display = "none";
      if (estado) estado.textContent = "No se pudo sacar la captura; el informe se manda igual sin ella.";
    }
  }

  function alCambiarCaptura() {
    const check = document.getElementById("informeCaptura");
    const caja = document.getElementById("informeCapturaCaja");
    if (!check || !caja) return;
    caja.style.display = check.checked ? "block" : "none";
    if (!check.checked) return;
    if (!capturaLista) capturaLista = prepararCaptura(false);
    const estado = document.getElementById("informeCapturaEstado");
    if (estado) estado.textContent = "Sacando la captura…";
    capturaLista.then(mostrarVistaCaptura);
  }

  /* ---------- Ventana "Informar un problema" ---------- */
  async function abrirInformeProblema() {
    const fondo = document.getElementById("informeModalBackdrop");
    if (!fondo) return;
    const enApp = !!(window.veekpos && typeof window.veekpos.capturarVentana === "function");
    // En la app, la foto se saca ANTES de mostrar la ventanita (queda la pantalla tal cual)
    capturaLista = enApp ? prepararCaptura(true) : null;
    if (enApp) await Promise.race([capturaLista, new Promise(r => setTimeout(r, 1500))]);

    const desc = document.getElementById("informeDescripcion");
    const contacto = document.getElementById("informeContacto");
    const estado = document.getElementById("informeEstado");
    desc.value = "";
    try { contacto.value = localStorage.getItem("vpos_informe_contacto") || ""; } catch (e) {}
    estado.textContent = ""; estado.className = "informe-estado";
    document.getElementById("informeIncluirDatos").checked = true;
    const chkCaptura = document.getElementById("informeCaptura");
    if (chkCaptura) { chkCaptura.checked = false; chkCaptura.onchange = alCambiarCaptura; }
    const cajaCaptura = document.getElementById("informeCapturaCaja");
    if (cajaCaptura) cajaCaptura.style.display = "none";
    const vista = document.getElementById("informeCapturaVista");
    if (vista) { vista.removeAttribute("src"); vista.style.display = "none"; }

    const n = registro.length;
    document.getElementById("informeResumenDatos").textContent =
      `Sección, versión, navegador, últimos pasos` + (window.posOffline ? ", estado de la sincronización" : "") +
      (n ? ` · ${n} error${n === 1 ? "" : "es"} registrado${n === 1 ? "" : "s"} en esta sesión` : " · sin errores registrados en esta sesión");
    const cola = await estadoColaCaja();
    document.getElementById("informeDetalleDatos").textContent = JSON.stringify({ equipo: datosDelEquipo(), pasos: pasos.slice(-15), cola, errores: registro.slice(-12) }, null, 2);
    document.getElementById("btnEnviarInforme").disabled = false;
    fondo.classList.add("show");
    setTimeout(() => desc.focus(), 80);
  }

  function cerrarInformeProblema() {
    document.getElementById("informeModalBackdrop")?.classList.remove("show");
  }

  async function enviarInformeProblema() {
    const desc = document.getElementById("informeDescripcion").value.trim();
    const contacto = document.getElementById("informeContacto").value.trim();
    const incluir = document.getElementById("informeIncluirDatos").checked;
    const conCaptura = !!document.getElementById("informeCaptura")?.checked;
    const estado = document.getElementById("informeEstado");
    const btn = document.getElementById("btnEnviarInforme");
    if (desc.length < 5) {
      estado.textContent = "Contá en pocas palabras qué pasó (qué estabas haciendo y qué salió mal).";
      estado.className = "informe-estado error";
      document.getElementById("informeDescripcion").focus();
      return;
    }
    try { localStorage.setItem("vpos_informe_contacto", contacto); } catch (e) {}
    btn.disabled = true;
    estado.textContent = "Enviando…"; estado.className = "informe-estado";
    try {
      let captura = null;
      if (conCaptura) {
        if (!capturaLista) capturaLista = prepararCaptura(false);
        estado.textContent = "Preparando la captura…";
        captura = await Promise.race([capturaLista, new Promise(r => setTimeout(() => r(null), 15000))]);
        estado.textContent = "Enviando…";
      }
      const r = await enviarAlServidor({
        descripcion: desc,
        contacto,
        contexto: incluir ? datosDelEquipo() : { seccion: datosDelEquipo().seccion, versionPanel: VERSION_PANEL },
        errores: incluir ? registro.slice(-20) : [],
        pasos: incluir ? pasos.slice(-MAX_PASOS) : [],
        cola: incluir ? await estadoColaCaja() : null,
        captura: captura || undefined
      });
      if (!r || !r.success) throw new Error((r && r.message) || "El servidor no pudo recibir el informe.");
      if (r.enviado === false && r.message) {
        estado.textContent = "⚠️ " + r.message;
        estado.className = "informe-estado aviso";
        btn.disabled = false;
        return;
      }
      cerrarInformeProblema();
      capturaLista = null;
      if (typeof toast === "function") toast("Informe enviado a soporte. ¡Gracias!", "success");
    } catch (e) {
      estado.textContent = "No se pudo enviar: " + (e.message || "error de conexión") + ". Probá de nuevo en un rato.";
      estado.className = "informe-estado error";
      btn.disabled = false;
    }
  }

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && document.getElementById("informeModalBackdrop")?.classList.contains("show")) {
      e.preventDefault(); cerrarInformeProblema();
    }
  });

  // Versión de la app de escritorio (para el informe)
  try {
    if (window.veekpos && typeof window.veekpos.versionApp === "function") {
      window.veekpos.versionApp().then(v => { window.__versionAppVeek = v; }).catch(() => {});
    }
  } catch (e) {}

  window.abrirInformeProblema = abrirInformeProblema;
  window.cerrarInformeProblema = cerrarInformeProblema;
  window.enviarInformeProblema = enviarInformeProblema;
  window.anotarErrorParaInforme = anotar; // lo usa fetchAPI para anotar llamadas fallidas
  window.anotarPasoParaInforme = anotarPaso;
})();
