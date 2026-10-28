/* =====================================================================
   VeekAuth — sesión con el backend (token firmado)
   ---------------------------------------------------------------------
   Se carga ANTES que admin.js (y en stock.html / Reportes). Hace tres
   cosas, sin que haga falta tocar cada llamada al backend:

   1. Agrega el token de sesión a toda solicitud al backend (Apps Script
      o Worker): como parámetro "token" en GET, o como campo "token" en
      el body JSON de los POST.
   2. Manda las MODIFICACIONES (acciones que empiezan con guardar/
      actualizar/eliminar/...) por POST aunque el código las arme como
      GET: así no viajan datos ni el token en la URL y no hay límite de
      largo. El backend las atiende igual (ver doPostInterno en code.gs).
   3. Si el backend responde que la sesión venció, muestra un aviso con
      el botón "Volver a iniciar sesión" (sin cortar lo que se está
      haciendo: el POS offline sigue funcionando y la cola espera).
===================================================================== */
(function () {
  "use strict";

  var CLAVES = {
    panel: { token: "veekToken", vence: "veekTokenVence", disp: "veekTokenDispositivo", dispVence: "veekTokenDispositivoVence" },
    stock: { token: "veekTokenStock", vence: "veekTokenStockVence", disp: "veekTokenStockDisp", dispVence: "veekTokenStockDispVence" }
  };

  // Acciones que no necesitan token (catálogo público): no se les agrega,
  // así sus URLs quedan iguales y se aprovecha la caché del Worker/navegador.
  var PUBLICAS = {
    productos: 1, categorias: 1, destacados: 1, ofertas: 1, producto: 1,
    configLanding: 1, configuracionNegocio: 1, catalogoVersion: 1,
    catalogoPDFInfo: 1, imagenProxy: 1, consultarPedido: 1, qrEstado: 1, qrPantalla: 1
  };

  var RE_MUTACION = /^(guardar|actualizar|eliminar|crear|editar|registrar|cambiar|anular|borrar|marcar|activar|fijar|cancelar|confirmar|aplicar|reordenar|subir|migrar|inicializar|sumar|restaurar|vaciar|cerrar)/i;
  var MUTACIONES_EXTRA = { qrPagado: 1, qrLimpiar: 1, qrFijar: 1, login: 1, loginStock: 1 };

  var estado = {
    tipo: "panel",
    urlLogin: "login.html",
    apis: [],
    avisoMostrado: false
  };

  function ss(k, v) { try { if (v === undefined) return sessionStorage.getItem(k); if (v === null) sessionStorage.removeItem(k); else sessionStorage.setItem(k, v); } catch (e) { return null; } }
  function ls(k, v) { try { if (v === undefined) return localStorage.getItem(k); if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) { return null; } }

  function claves() { return CLAVES[estado.tipo] || CLAVES.panel; }

  function vigente(venceTexto) {
    var vence = Number(venceTexto || 0);
    // Sin fecha guardada se asume vigente (el servidor decide igual).
    return !vence || vence - 60000 > Date.now();
  }

  var VeekAuth = {
    /** opciones: { tipo: "panel"|"stock", urlLogin, apis: [url...] } */
    configurar: function (opciones) {
      opciones = opciones || {};
      if (opciones.tipo) estado.tipo = opciones.tipo;
      if (opciones.urlLogin) estado.urlLogin = opciones.urlLogin;
      (opciones.apis || []).forEach(VeekAuth.registrarApi);
      return VeekAuth;
    },

    registrarApi: function (url) {
      url = String(url || "").trim();
      if (url && estado.apis.indexOf(url) === -1) estado.apis.push(url.replace(/[?#].*$/, ""));
    },

    /**
     * Guarda la sesión que devolvió login/loginStock.
     * recordar=true (app de escritorio, app de Reportes): además queda en
     * este equipo hasta que vence, para no pedir login todos los días y
     * para que la cola offline pueda seguir sincronizando.
     */
    guardarSesion: function (data, recordar) {
      if (!data || !data.token) return;
      var c = claves();
      ss(c.token, data.token);
      ss(c.vence, String(data.venceEn || ""));
      if (recordar) {
        ls(c.disp, data.token);
        ls(c.dispVence, String(data.venceEn || ""));
      }
      estado.avisoMostrado = false;
      ocultarAviso();
    },

    token: function () {
      var c = claves();
      var t = ss(c.token);
      if (t && vigente(ss(c.vence))) return t;
      var d = ls(c.disp);
      if (d && vigente(ls(c.dispVence))) {
        ss(c.token, d);
        ss(c.vence, ls(c.dispVence) || "");
        return d;
      }
      return "";
    },

    tieneSesion: function () { return !!VeekAuth.token(); },

    /**
     * mantenerDispositivo=true (app de escritorio): se sale del panel pero
     * el equipo conserva su sesión de dispositivo, así las ventas offline
     * que queden en la cola se siguen subiendo solas.
     */
    cerrarSesion: function (opciones) {
      var c = claves();
      [c.token, c.vence].forEach(function (k) { ss(k, null); });
      if (!(opciones && opciones.mantenerDispositivo)) {
        [c.disp, c.dispVence].forEach(function (k) { ls(k, null); });
      }
    },

    esUrlApi: function (url) {
      url = String(url || "");
      if (!/^https?:/i.test(url)) return false;
      // Solo el backend de ESTE negocio (no el servidor de licencias ni
      // otros servicios): la URL guardada en API_URL o las registradas.
      for (var i = 0; i < estado.apis.length; i++) {
        if (estado.apis[i] && url.indexOf(estado.apis[i]) === 0) return true;
      }
      try { if (typeof API_URL !== "undefined" && API_URL && url.indexOf(String(API_URL).replace(/[?#].*$/, "")) === 0) return true; } catch (e) {}
      return false;
    },

    esMutacion: function (accion) {
      accion = String(accion || "");
      return !!MUTACIONES_EXTRA[accion] || RE_MUTACION.test(accion);
    },

    /** Se llama cuando el backend responde sesionInvalida. */
    sesionVencida: function () {
      if (estado.avisoMostrado) return;
      estado.avisoMostrado = true;
      var c = claves();
      ss(c.token, null); ss(c.vence, null);
      ls(c.disp, null); ls(c.dispVence, null);
      mostrarAviso();
      try { window.dispatchEvent(new CustomEvent("veek:sesion-vencida")); } catch (e) {}
    },

    irAlLogin: function () {
      try { sessionStorage.removeItem("admin"); } catch (e) {}
      window.location.href = estado.urlLogin;
    }
  };

  /* ------------------------------ Aviso ------------------------------ */

  function mostrarAviso() {
    function pintar() {
      if (document.getElementById("veekAvisoSesion")) return;
      var div = document.createElement("div");
      div.id = "veekAvisoSesion";
      div.setAttribute("role", "alert");
      div.style.cssText = "position:fixed;left:50%;bottom:20px;transform:translateX(-50%);z-index:2147483000;" +
        "max-width:min(560px,calc(100vw - 32px));background:#1f2937;color:#fff;border-radius:12px;" +
        "box-shadow:0 10px 30px rgba(0,0,0,.35);padding:14px 16px;display:flex;gap:12px;align-items:center;" +
        "font:500 14px/1.4 system-ui,-apple-system,Segoe UI,Roboto,sans-serif";
      div.innerHTML =
        '<span style="font-size:20px" aria-hidden="true">🔒</span>' +
        '<span style="flex:1">' + (ss("veekEntroSinServidor") === "1"
          ? 'Entraste sin poder validar con el servidor (no había conexión o tardó demasiado). Lo que hagas en el POS se guarda; ' +
            'cuando tengas internet, volvé a iniciar sesión para subirlo.'
          : 'Tu sesión venció. Lo que hagas en el POS se sigue guardando; para enviar cambios al servidor, volvé a iniciar sesión.') + '</span>' +
        '<button type="button" id="veekAvisoSesionBtn" style="background:#fff;color:#111;border:0;border-radius:8px;' +
        'padding:8px 12px;font-weight:700;cursor:pointer;white-space:nowrap">Iniciar sesión</button>' +
        '<button type="button" id="veekAvisoSesionX" aria-label="Cerrar aviso" style="background:transparent;color:#fff;' +
        'border:0;font-size:18px;cursor:pointer;padding:4px 6px">✕</button>';
      document.body.appendChild(div);
      document.getElementById("veekAvisoSesionBtn").onclick = VeekAuth.irAlLogin;
      document.getElementById("veekAvisoSesionX").onclick = function () { ocultarAviso(); };
    }
    if (document.body) pintar();
    else document.addEventListener("DOMContentLoaded", pintar);
  }

  function ocultarAviso() {
    var d = document.getElementById("veekAvisoSesion");
    if (d) d.remove();
  }

  /* ------------------------- fetch con sesión ------------------------- */

  var fetchOriginal = window.fetch ? window.fetch.bind(window) : null;
  if (!fetchOriginal) { window.VeekAuth = VeekAuth; return; }

  window.fetch = function (input, init) {
    var url = typeof input === "string" ? input : (input && typeof input.href === "string" ? input.href : null);
    if (url === null || !VeekAuth.esUrlApi(url)) return fetchOriginal(input, init);

    init = Object.assign({}, init || {});
    var metodo = String(init.method || "GET").toUpperCase();
    var tok = VeekAuth.token();

    try {
      if (metodo === "GET") {
        var u = new URL(url, window.location.href);
        var accion = u.searchParams.get("action") || "";
        if (VeekAuth.esMutacion(accion)) {
          var cuerpo = {};
          u.searchParams.forEach(function (v, k) { cuerpo[k] = v; });
          if (tok && !cuerpo.token) cuerpo.token = tok;
          url = u.origin + u.pathname;
          init.method = "POST";
          init.headers = { "Content-Type": "text/plain;charset=utf-8" };
          init.body = JSON.stringify(cuerpo);
        } else if (tok && !PUBLICAS[accion || "productos"]) {
          u.searchParams.set("token", tok);
          url = u.toString();
        }
      } else if (metodo === "POST" && tok && typeof init.body === "string") {
        var obj = JSON.parse(init.body);
        if (obj && typeof obj === "object" && !Array.isArray(obj) && !obj.token) {
          obj.token = tok;
          init.body = JSON.stringify(obj);
        }
      }
    } catch (e) { /* si algo no se puede interpretar, se manda tal cual */ }

    return fetchOriginal(url, init).then(function (resp) {
      try {
        resp.clone().text().then(function (t) {
          if (t && t.indexOf('"sesionInvalida":true') !== -1) VeekAuth.sesionVencida();
        }).catch(function () {});
      } catch (e) {}
      return resp;
    });
  };

  window.VeekAuth = VeekAuth;
})();
