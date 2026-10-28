/* ╔═══════════════════════════════════════════════════════════════════╗
   ║  VeekPOS — BACKEND (Google Apps Script)                           ║
   ║                                                                   ║
   ║  DATOS DEL CLIENTE: lo ÚNICO que se toca al instalar en un        ║
   ║  negocio nuevo es este bloque. Nombre, teléfonos, WhatsApp,       ║
   ║  colores, banner, etc. se editan desde el panel                   ║
   ║  (Configuración → Negocio / Catálogo web) y viven en la hoja      ║
   ║  CONFIGURACION.                                                   ║
   ╠═══════════════════════════════════════════════════════════════════╣
   ║  AL PEGAR ESTE ARCHIVO (instalación nueva o actualización):       ║
   ║   1. ⚙️ Configuración del proyecto → activar "Mostrar el archivo  ║
   ║      de manifiesto appsscript.json" → pegar ahí el                ║
   ║      appsscript.json que viene al lado de este archivo.           ║
   ║   2. Arriba, en el desplegable de funciones, elegir               ║
   ║      autorizarVeekPOS → ▶ Ejecutar → aceptar los permisos         ║
   ║      ("app no verificada": Configuración avanzada → Ir a…).       ║
   ║      Aprueba TODOS los permisos de una vez y, si la planilla      ║
   ║      está vacía, crea hojas, carpetas de Drive y triggers.        ║
   ║      En "Registro de ejecución" todo tiene que salir con ✓.       ║
   ║   3. Implementar → Nueva implementación (o Administrar            ║
   ║      implementaciones → editar → Versión nueva si ya existía)     ║
   ║      → Aplicación web · Ejecutar como: Yo · Acceso: Cualquier     ║
   ║      usuario. Copiar la URL /exec y pegarla en el instalador.     ║
   ╚═══════════════════════════════════════════════════════════════════╝ */

// ID de la planilla de Google Sheets del negocio (lo que está entre
// /d/ y /edit en la URL de la planilla). Si se deja vacío, se usa la
// propiedad del script "SPREADSHEET_ID" y, si tampoco está, la planilla
// a la que está vinculado este script (Extensiones → Apps Script).
const SPREADSHEET_ID = "";

// INFORMES DE ERROR: adónde llegan (el correo de quien da soporte al
// sistema, no el del negocio). Los manda la cuenta dueña del Apps
// Script. Vacío = no se envían (igual quedan anotados en la hoja ERRORES).
const CORREO_SOPORTE = "vheredia255@gmail.com";
// Además de los que manda la persona con "Informar un problema", se
// avisan solos los errores inesperados del servidor, del panel y del
// catálogo (agrupados: el mismo error se avisa una vez cada 6 horas).
const REPORTES_AUTOMATICOS = true;

// Horario en el que vale la pena mantener la caché del catálogo tibia.
const HORA_APERTURA = 7;
const HORA_CIERRE = 23;

// Archivado automático del historial de ventas POS: cada mes, las ventas
// más viejas que esta cantidad de meses pasan a VENTAS_LOCAL_ARCHIVO /
// DETALLE_VENTAS_ARCHIVO (la planilla se mantiene liviana y rápida).
// OJO: los reportes por período no incluyen lo archivado.
// 0 = desactivado (no archiva nunca solo).
const ARCHIVAR_VENTAS_MESES = 0;

// Ventas POS: se guarda el detalle de cada venta (ticket por ticket) de
// los últimos MESES_DETALLE_VENTAS meses. Lo anterior se resume una vez
// por mes (trigger diario de las 3 AM): una fila por día y forma de
// pago en VENTAS_LOCAL y una por día y producto en DETALLE_VENTAS, que
// es lo que usan los reportes (totales, formas de pago, productos y
// categorías por día). Antes de resumir, el detalle completo se guarda
// en un CSV en Drive. 0 = no resumir nunca.
const MESES_DETALLE_VENTAS = 6;

/* ═════════════════════ fin de los datos del cliente ═════════════════ */

const SHEET_NAME = "PRODUCTOS";

// --- OPTIMIZACION: cache del Spreadsheet dentro de una misma ejecucion. ---
// SpreadsheetApp.openById() es la llamada mas costosa de todo el script
// (hace un viaje de red). Antes se llamaba una vez por cada funcion que
// necesitaba la planilla; si una sola solicitud terminaba disparando
// varias funciones (por ej. dashboard(), que combina datos de varias
// hojas), se reabria la planilla varias veces sin necesidad.
// getSS() guarda el resultado en una variable global de script: la
// primera llamada abre la planilla, todas las siguientes (dentro de la
// misma ejecucion) reutilizan el mismo objeto. No cambia ningun
// comportamiento: sigue siendo la misma planilla, con los mismos datos.
let _ssCache = null;
function getSS() {
  if (!_ssCache) {
    const id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "";
    _ssCache = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
    if (!_ssCache) {
      throw new Error("No hay planilla configurada: completá SPREADSHEET_ID al principio de code.gs.");
    }
  }
  return _ssCache;
}

// --- OPTIMIZACION: cache de hojas por nombre dentro de una misma ejecucion. ---
// ss.getSheetByName("PEDIDOS") se repetia muchas veces con el mismo
// nombre dentro de la misma solicitud (por ej. PEDIDOS y DETALLE_PEDIDOS
// se buscan una y otra vez en varias funciones que se llaman entre si).
// getHoja(nombre) guarda cada hoja ya encontrada en un mapa; si se pide
// el mismo nombre de nuevo en la misma ejecucion, se devuelve la hoja
// ya obtenida en vez de volver a buscarla. Es exactamente la misma hoja
// (mismo objeto que devolvia ss.getSheetByName), asi que no cambia nada
// del comportamiento — solo evita busquedas repetidas.
const _hojaCache = {};
function getHoja(nombre) {
  if (_hojaCache[nombre]) return _hojaCache[nombre];
  // Solo se recuerdan las hojas que existen: si se recordara "no existe",
  // una hoja recién creada en esta misma ejecución (instalación inicial)
  // seguiría apareciendo como inexistente.
  const hoja = getSS().getSheetByName(nombre);
  if (hoja) _hojaCache[nombre] = hoja;
  return hoja;
}

/**
 * Devuelve el número de columna (1-based) de un encabezado en la fila 1
 * de una hoja, buscándolo por NOMBRE en vez de asumir una posición fija
 * — así el código no se rompe ni pisa datos de otra columna si la hoja
 * tiene un orden distinto al esperado, o le faltan columnas.
 *
 * Si el encabezado no existe todavía, lo agrega al final de la fila 1
 * (en la primera columna vacía) y devuelve esa columna nueva.
 */
function obtenerOcrearColumna(sheet, nombreEncabezado) {
  const ultimaColumna = Math.max(sheet.getLastColumn(), 1);
  const encabezados = sheet.getRange(1, 1, 1, ultimaColumna).getValues()[0];

  for (let i = 0; i < encabezados.length; i++) {
    if (String(encabezados[i] || "").trim() === nombreEncabezado) {
      return i + 1; // 1-based
    }
  }

  // No existe: se crea en la primera columna vacía después de la última usada.
  const nuevaColumna = ultimaColumna + 1;
  sheet.getRange(1, nuevaColumna).setValue(nombreEncabezado);
  return nuevaColumna;
}

/* =====================================================================
   CACHÉ DE CATÁLOGO (CacheService) — pensado para que el catálogo
   online y el panel admin carguen rápido.

   Antes, CADA request (cada visita al catálogo, cada apertura del
   panel) volvía a leer TODA la hoja PRODUCTOS con getDataRange() —
   la operación más lenta que existe en un Web App de Apps Script.
   Con muchos usuarios mirando el catálogo a la vez, eso significaba
   leer la misma hoja una y otra vez para devolver, la mayoría de las
   veces, exactamente el mismo resultado.

   CacheService.getScriptCache() es una caché COMPARTIDA entre todas
   las ejecuciones del script (no solo dentro de una misma request,
   como _ssCache/_hojaCache) — dura hasta 6 horas, pero acá se usa con
   un TTL corto (ver CACHE_TTL_SEGUNDOS) porque el catálogo cambia
   seguido. Cada vez que se guarda/edita/borra un producto o cambia el
   stock, se invalida (invalidarCacheProductos()) para que la próxima
   lectura sea siempre fresca — nunca se sirve un dato viejo a
   propósito, solo se evita releer la hoja cuando nada cambió.

   LÍMITE IMPORTANTE: cada key de CacheService admite como máximo
   100KB. Un catálogo grande (muchos productos, descripciones largas)
   puede superar eso fácilmente como un solo JSON. Por eso el texto se
   parte en pedazos (chunks) de 90KB antes de guardarlo, y se
   reensambla al leer — ver cacheGuardarTexto/cacheLeerTexto.
===================================================================== */

const CACHE_TTL_SEGUNDOS = 300; // 5 minutos para el catálogo público
const CACHE_TTL_ADMIN_SEGUNDOS = 300; // 5 minutos (se invalida solo al editar; el trigger de calentamiento lo mantiene tibio)
const CACHE_CHUNK_SIZE = 90000; // 90KB por pedazo, con margen bajo el límite de 100KB de CacheService

function cacheGuardarTexto(clave, texto, ttlSegundos){
  try{
    const cache = CacheService.getScriptCache();
    const partes = [];
    for(let i = 0; i < texto.length; i += CACHE_CHUNK_SIZE){
      partes.push(texto.slice(i, i + CACHE_CHUNK_SIZE));
    }
    // Se guarda primero la cantidad de partes, y recién después las
    // partes en sí — así una lectura concurrente nunca encuentra
    // "cantidad" ya escrita pero alguna parte todavía vieja/faltante.
    const mapaChunks = {};
    partes.forEach((parte, i) => { mapaChunks[clave + ":" + i] = parte; });
    cache.putAll(mapaChunks, ttlSegundos);
    cache.put(clave + ":n", String(partes.length), ttlSegundos);
  } catch(e){
    // Si guardar la caché falla por lo que sea (ej. texto gigante que
    // ni con chunks entra), no debe romper la respuesta real — solo
    // significa que esta vez no se cacheó.
    console.error("No se pudo guardar en caché (" + clave + "): " + e.message);
  }
}

function cacheLeerTexto(clave){
  try{
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(clave + ":n") || 0);
    if(!n) return null;

    const claves = [];
    for(let i = 0; i < n; i++) claves.push(clave + ":" + i);
    const mapa = cache.getAll(claves);

    let texto = "";
    for(let i = 0; i < n; i++){
      const parte = mapa[clave + ":" + i];
      if(parte === undefined || parte === null) return null; // algún pedazo venció/faltó -> se descarta todo, mejor releer fresco
      texto += parte;
    }
    return texto;
  } catch(e){
    console.error("No se pudo leer de caché (" + clave + "): " + e.message);
    return null;
  }
}

/** Se llama después de cualquier cambio en la hoja PRODUCTOS (crear, editar, borrar, sumar stock, reordenar, ventas/pedidos que descuentan stock) para que la próxima lectura sea fresca. */
function invalidarCacheProductos(){
  try{
    const cache = CacheService.getScriptCache();
    ["catalogo_publico", "catalogo_admin", "categorias", "destacados", "ofertas"].forEach(base => {
      // Se borran hasta 20 posibles chunks de cada key — de sobra para
      // un catálogo de hasta ~1.8MB de JSON (20 x 90KB).
      const claves = ["n"];
      for(let i = 0; i < 20; i++) claves.push(String(i));
      cache.removeAll(claves.map(c => base + ":" + c));
    });
  } catch(e){
    console.error("No se pudo invalidar la caché de productos: " + e.message);
  }
  marcarCambioCatalogo(false);
}

/* =====================================================================
   VERSIÓN DEL CATÁLOGO
   Las cajas de escritorio antes descargaban el catálogo COMPLETO cada
   vez que se entraba al POS o a Productos, aunque nada hubiera
   cambiado. Ahora preguntan primero action=catalogoVersion (respuesta
   de pocos bytes, sin tocar la planilla) y solo bajan:
     - el catálogo entero, si cambió "version" (alta/edición/borrado
       de productos, precios, edición manual en la hoja), o
     - solo CODIGO+STOCK (action=stockTodos), si cambió únicamente
       "stockVersion" (ventas, ajustes de stock).
===================================================================== */
const PROP_VERSION_CATALOGO = "CATALOGO_VERSION";
const PROP_VERSION_STOCK = "CATALOGO_STOCK_VERSION";

function marcarCambioCatalogo(soloStock){
  try{
    const ahora = String(Date.now());
    const props = PropertiesService.getScriptProperties();
    const cambios = {};
    cambios[PROP_VERSION_STOCK] = ahora;
    if(!soloStock) cambios[PROP_VERSION_CATALOGO] = ahora;
    props.setProperties(cambios);
    CacheService.getScriptCache().remove("catalogo_versiones");
  }catch(e){
    console.error("No se pudo actualizar la versión del catálogo: " + e.message);
  }
}

function leerVersionesCatalogo(){
  const cache = CacheService.getScriptCache();
  try{
    const c = cache.get("catalogo_versiones");
    if(c) return JSON.parse(c);
  }catch(e){}
  const props = PropertiesService.getScriptProperties();
  let version = props.getProperty(PROP_VERSION_CATALOGO);
  let stockVersion = props.getProperty(PROP_VERSION_STOCK);
  if(!version){
    version = String(Date.now());
    stockVersion = stockVersion || version;
    const ini = {};
    ini[PROP_VERSION_CATALOGO] = version;
    ini[PROP_VERSION_STOCK] = stockVersion;
    props.setProperties(ini);
  }
  const res = { version: version, stockVersion: stockVersion || version };
  try{ cache.put("catalogo_versiones", JSON.stringify(res), 600); }catch(e){}
  return res;
}

function getCatalogoVersion(){
  const v = leerVersionesCatalogo();
  return jsonResponse({ success:true, version: v.version, stockVersion: v.stockVersion });
}

/** CODIGO + STOCK de todo el catálogo, leyendo solo esas dos columnas */
function getStockTodos(){
  const sheet = getSheet();
  if(!sheet) return jsonResponse({ success:false, message:"No se encontró la hoja 'PRODUCTOS'." });
  const ultimaFila = sheet.getLastRow();
  const ultimaCol = sheet.getLastColumn();
  const v = leerVersionesCatalogo();
  if(ultimaFila < 2) return jsonResponse({ success:true, productos: [], stockVersion: v.stockVersion });
  const headers = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h || "").trim());
  const idxCodigo = headers.indexOf("CODIGO");
  const idxStock = headers.indexOf("STOCK");
  if(idxCodigo === -1 || idxStock === -1) return jsonResponse({ success:false, message:"Faltan columnas CODIGO/STOCK." });
  const codigos = sheet.getRange(2, idxCodigo + 1, ultimaFila - 1, 1).getValues();
  const stocks = sheet.getRange(2, idxStock + 1, ultimaFila - 1, 1).getValues();
  const productos = [];
  for(let i = 0; i < codigos.length; i++){
    const c = String(codigos[i][0] || "").trim();
    if(c) productos.push({ CODIGO: c, STOCK: Number(stocks[i][0]) || 0 });
  }
  return jsonResponse({ success:true, productos: productos, stockVersion: v.stockVersion });
}

/**
 * En vez de tirar toda la caché del catálogo en cada venta (lo que
 * obligaba a la siguiente lectura a releer la hoja PRODUCTOS entera),
 * se actualiza el STOCK de los códigos tocados dentro del JSON ya
 * cacheado. Si algo falla, se invalida esa clave puntual.
 */
function parchearStockEnCachesCatalogo(stockPorCodigo){
  const cache = CacheService.getScriptCache();
  const borrarClave = (base) => {
    const claves = ["n"];
    for(let i = 0; i < 20; i++) claves.push(String(i));
    try{ cache.removeAll(claves.map(c => base + ":" + c)); }catch(e){}
  };
  [["catalogo_publico", CACHE_TTL_SEGUNDOS], ["catalogo_admin", CACHE_TTL_ADMIN_SEGUNDOS]].forEach(par => {
    const clave = par[0];
    const texto = cacheLeerTexto(clave);
    if(!texto) return;
    try{
      const obj = JSON.parse(texto);
      let cambio = false;
      (obj.productos || []).forEach(p => {
        const c = String(p.CODIGO || "").trim();
        if(Object.prototype.hasOwnProperty.call(stockPorCodigo, c)){ p.STOCK = stockPorCodigo[c]; cambio = true; }
      });
      if(cambio) cacheGuardarTexto(clave, JSON.stringify(obj), par[1]);
    }catch(e){
      borrarClave(clave);
    }
  });
  // Destacados/ofertas son listas chicas derivadas: se regeneran solas
  borrarClave("destacados");
  borrarClave("ofertas");
}

/** Llamar después de cambiar SOLO stock (ventas, ajustes): parchea la caché y avisa a las cajas */
function notificarCambioStock(stockPorCodigo){
  if(stockPorCodigo && Object.keys(stockPorCodigo).length > 0){
    parchearStockEnCachesCatalogo(stockPorCodigo);
  }
  marcarCambioCatalogo(true);
}

/**
 * Aplica deltas de stock { CODIGO: delta } leyendo SOLO las columnas
 * CODIGO y STOCK (antes se leía la hoja PRODUCTOS completa con
 * getDataRange y se escribía celda por celda). Devuelve
 * { CODIGO: stockNuevo } de los códigos encontrados. El llamador debe
 * tener el lock tomado.
 */
function aplicarDeltasStock(deltasPorCodigo, minimoCero){
  const resultado = {};
  const codigosDelta = Object.keys(deltasPorCodigo || {});
  if(codigosDelta.length === 0) return resultado;

  const sheet = getSheet();
  if(!sheet) return resultado;
  const ultimaFila = sheet.getLastRow();
  const ultimaCol = sheet.getLastColumn();
  if(ultimaFila < 2) return resultado;

  const headers = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h || "").trim());
  const idxCodigo = headers.indexOf("CODIGO");
  const idxStock = headers.indexOf("STOCK");
  if(idxCodigo === -1 || idxStock === -1) return resultado;

  const codigos = sheet.getRange(2, idxCodigo + 1, ultimaFila - 1, 1).getValues();
  const rangoStock = sheet.getRange(2, idxStock + 1, ultimaFila - 1, 1);
  const stocks = rangoStock.getValues();

  const filasCambiadas = [];
  for(let i = 0; i < codigos.length; i++){
    const c = String(codigos[i][0] || "").trim();
    if(c && Object.prototype.hasOwnProperty.call(deltasPorCodigo, c) && !Object.prototype.hasOwnProperty.call(resultado, c)){
      let nuevo = (Number(stocks[i][0]) || 0) + Number(deltasPorCodigo[c] || 0);
      if(minimoCero) nuevo = Math.max(0, nuevo);
      stocks[i][0] = nuevo;
      resultado[c] = nuevo;
      filasCambiadas.push(i);
    }
  }

  if(filasCambiadas.length === 0) return resultado;
  if(filasCambiadas.length <= 15){
    filasCambiadas.forEach(i => sheet.getRange(i + 2, idxStock + 1).setValue(stocks[i][0]));
  } else {
    rangoStock.setValues(stocks); // una sola escritura para lotes grandes
  }
  return resultado;
}

/** Agrupa un carrito en { CODIGO: delta } (signo = -1 para ventas) */
function deltasDesdeCarrito(carrito, signo){
  const deltas = {};
  (carrito || []).forEach(item => {
    const codigo = String((item && item.CODIGO) || "").trim();
    if(!codigo) return;
    deltas[codigo] = (deltas[codigo] || 0) + signo * Number(item.cantidad || 0);
  });
  return deltas;
}

/**
 * Trigger instalable de tipo "On edit" — se dispara automáticamente
 * cuando alguien edita una celda A MANO directamente en la planilla
 * de Google Sheets (no a través del panel admin).
 *
 * SIN ESTO: si editás DESTACADO, PUBLICADO, STOCK, etc. directamente
 * en Sheets, el catálogo público sigue mostrando la versión vieja
 * hasta que venza el caché solo (hasta 5 minutos) — porque
 * invalidarCacheProductos() normalmente solo se llama desde las
 * funciones del panel (guardarProducto, actualizarProducto, etc.),
 * que nunca se ejecutan cuando el cambio se hizo a mano en la hoja.
 *
 * CÓMO INSTALARLO (se hace UNA SOLA VEZ):
 *   1. En el editor de Apps Script, elegí "instalarTriggerOnEdit" en
 *      el desplegable de funciones y tocá ▶️.
 *   2. Autorizá los permisos si te los pide.
 *   3. Listo — a partir de ahora, cualquier edición manual en
 *      PRODUCTOS invalida el caché al instante.
 *
 * (Un simple function onEdit(e){...} SIN instalar como trigger
 * también se ejecuta solo, pero corre con permisos limitados y NO
 * puede usar CacheService en todos los casos — por eso se instala
 * como trigger explícito con ScriptApp.newTrigger, que corre con
 * los mismos permisos que el resto del script.)
 */
function instalarTriggerOnEdit(){
  // Evita duplicar el trigger si ya estaba instalado.
  ScriptApp.getProjectTriggers().forEach(t => {
    if(t.getHandlerFunction() === "onEditInstalable"){
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onEditInstalable")
    .forSpreadsheet(getSS())
    .onEdit()
    .create();

  Logger.log("Trigger onEdit instalado correctamente.");
}

/**
 * Handler del trigger instalable. Se ejecuta en CADA edición de
 * CUALQUIER hoja de la planilla — por eso lo primero que hace es
 * filtrar por nombre de hoja, para no invalidar cachés de más.
 */
function onEditInstalable(e){
  try{
    if(!e || !e.range) return;
    const nombreHoja = e.range.getSheet().getName();

    if(nombreHoja === "PRODUCTOS"){
      invalidarCacheProductos();
    } else if(nombreHoja === "PEDIDOS" || nombreHoja === "VENTAS_LOCAL"){
      invalidarCachePolling();
    }
  } catch(err){
    console.error("Error en onEditInstalable: " + err.message);
  }
}

/**
 * Solo para cambios en VENTAS_LOCAL que no tocan pedidos (ventas del POS):
 * antes cada venta tiraba también la caché de "pedidos" y del índice de
 * pedidos, y el siguiente polling de Pedidos tenía que releer PEDIDOS +
 * DETALLE_PEDIDOS completas — con historial grande eso pasaba los 25s
 * y el panel mostraba "Error pedidos: AbortError".
 */
function invalidarCachePollingVentas(){
  try{
    const cache = CacheService.getScriptCache();
    ["dashboard", "ventas_pos_hoy", "metricas"].forEach(base => {
      const claves = ["n"];
      for(let i = 0; i < 20; i++) claves.push(String(i));
      cache.removeAll(claves.map(c => base + ":" + c));
    });
  } catch(e){
    console.error("No se pudo invalidar la caché de ventas: " + e.message);
  }
}

/** Se llama después de cualquier cambio en PEDIDOS o VENTAS_LOCAL (crear/editar/cobrar/anular pedido o venta) — invalida la caché de los endpoints de polling (dashboard, pedidos, ventasPOS) y el índice de búsqueda de pedidos. */
function invalidarCachePolling(){
  try{
    // Una sola llamada a la caché (antes eran cinco).
    const claves = [];
    ["dashboard", "pedidos", "ventas_pos_hoy", "pedidos_indice", "metricas"].forEach(base => {
      claves.push(base + ":n");
      for(let i = 0; i < 20; i++) claves.push(base + ":" + i);
    });
    CacheService.getScriptCache().removeAll(claves);
  } catch(e){
    console.error("No se pudo invalidar la caché de polling: " + e.message);
  }
}

const CACHE_TTL_PEDIDOS_INDICE_SEGUNDOS = 60; // corto a propósito: los pedidos cambian seguido (nuevos, cambios de estado, ítems editados)

/**
 * Devuelve PEDIDOS y DETALLE_PEDIDOS ya indexados por PEDIDO_ID, listos
 * para una búsqueda puntual O(1) — pensada para getDetallePedido y
 * consultarPedido, que antes hacían esto en CADA llamada:
 *
 *   rowsToObjects(pedidosSheet.getDataRange().getValues()).find(...)
 *
 * Eso lee y convierte a objetos TODA la hoja PEDIDOS (y TODA
 * DETALLE_PEDIDOS) solo para terminar usando una sola fila — con
 * pocos pedidos es instantáneo, pero con meses de historial
 * acumulado el escaneo se va poniendo cada vez más lento, hasta
 * empezar a superar el timeout que tiene el cliente para esa llamada
 * (ver fetchAPI en admin.js — 10s con reintento).
 *
 * Se cachea con el mismo mecanismo por chunks que ya usa el catálogo
 * (cacheGuardarTexto/cacheLeerTexto — ver más arriba), con TTL corto
 * porque los pedidos cambian seguido, y se invalida al toque desde
 * invalidarCachePolling() — la misma función que YA se llama en cada
 * alta, cambio de estado, cobro y edición de ítems de un pedido (ver
 * guardarPedido, actualizarEstado, marcarPedidoCobrado,
 * editarItemsPedido), así que no hizo falta tocar esos puntos: el
 * índice nunca queda desactualizado después de un cambio, solo evita
 * releer las hojas cuando nada cambió.
 */
function obtenerIndicePedidos(){
  const cacheado = cacheLeerTexto("pedidos_indice");
  if(cacheado){
    try{ return JSON.parse(cacheado); }
    catch(e){ /* JSON corrupto/parcial — sigue de largo y relee fresco */ }
  }

  const pedidosPorId = {};
  rowsToObjects(getHoja("PEDIDOS").getDataRange().getValues()).forEach(p => {
    p.PEDIDO_ID = String(p.PEDIDO_ID); // los números nuevos (1001…) la planilla los guarda como número
    pedidosPorId[p.PEDIDO_ID] = p;
  });

  const detallePorId = {};
  const detalleSheet = getHoja("DETALLE_PEDIDOS");
  if(detalleSheet){
    rowsToObjects(detalleSheet.getDataRange().getValues()).forEach(d => {
      const id = String(d.PEDIDO_ID);
      (detallePorId[id] || (detallePorId[id] = [])).push(d);
    });
  }

  const indice = { pedidosPorId, detallePorId };
  cacheGuardarTexto("pedidos_indice", JSON.stringify(indice), CACHE_TTL_PEDIDOS_INDICE_SEGUNDOS);
  return indice;
}

/**
 * Envuelve una función que arma su respuesta con jsonResponse(...) para
 * que quede cacheada un rato corto — pensada para los endpoints que el
 * panel consulta por POLLING cada 15s (dashboard, pedidos, ventasPOS).
 * Con varias cajas abiertas a la vez, todas piden lo mismo casi al
 * mismo tiempo; con TTL corto (10-15s) se evita releer las hojas en
 * cada ciclo sin perder frescura real (cualquier cambio invalida la
 * caché al toque vía invalidarCachePolling(), no hay que esperar a que
 * venza el TTL para ver algo nuevo que uno mismo acaba de guardar).
 */
function conCachePolling(clave, ttlSegundos, fnImpl){
  const cacheado = cacheLeerTexto(clave);
  if(cacheado){
    return ContentService.createTextOutput(cacheado).setMimeType(ContentService.MimeType.JSON);
  }
  const resultado = fnImpl();
  const texto = resultado.getContent();
  cacheGuardarTexto(clave, texto, ttlSegundos);
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
   SEGURIDAD — SESIONES FIRMADAS, ROLES Y CONTRASEÑAS
   ---------------------------------------------------------------------
   Antes, login() solo respondía "ok" y el resto de las acciones se
   ejecutaban sin preguntar quién las pedía: con la URL de Apps Script
   (que es pública, porque el catálogo la usa) cualquiera podía borrar
   ventas o cambiar stock. Ahora:

   1. login / loginStock devuelven un TOKEN firmado (HMAC-SHA256 con un
      secreto que vive solo en las Propiedades del script). No se guarda
      nada por sesión: el token lleva adentro usuario, rol y vencimiento,
      y la firma impide modificarlo.
   2. Toda acción que no esté en ACCIONES_PUBLICAS exige un token válido
      (parámetro "token" en GET, o campo "token" en el body del POST).
   3. El ROL lo decide el servidor (el del token / la hoja USUARIOS), no
      lo que manda el navegador: params.rol se pisa con el rol real.
   4. Las contraseñas se guardan con hash + sal + "pimienta" (otro
      secreto de las Propiedades). Las que todavía estén en texto plano
      se convierten solas la próxima vez que esa persona inicia sesión.

   Si alguna vez se pierde el acceso (por ejemplo, se borraron las
   Propiedades del script), ejecutar a mano resetearAccesoAdmin() desde
   el editor de Apps Script: vuelve el usuario/contraseña del admin a
   los valores por defecto (CREDENCIALES_DEFAULT) y cierra todas las
   sesiones abiertas.
===================================================================== */

// Acciones que el catálogo público, la pantalla del QR y los scripts de
// SEO usan SIN iniciar sesión. Todo lo demás requiere token.
// Versión del backend: aparece en las respuestas de pedidos y en
// ?action=prepararPedido, para saber qué código está publicado.
const VERSION_BACKEND = "2.4.2";

const ACCIONES_PUBLICAS = {
  productos: true, categorias: true, destacados: true, ofertas: true,
  producto: true, configLanding: true, configuracionNegocio: true,
  catalogoVersion: true, catalogoPDFInfo: true, imagenProxy: true,
  guardarPedido: true, consultarPedido: true, prepararPedido: true, validarCupon: true,
  login: true, loginStock: true, inicializarHojas: true,
  qrEstado: true, qrPantalla: true,
  reportarErrorCatalogo: true,
  // entre backends (local ↔ depósito): las valida la clave de traspasos, no una sesión
  traspasosParaLocal: true, confirmarRecepcionTraspaso: true
};

// Solo el rol admin (cambian accesos, configuración o borran en forma definitiva).
const ACCIONES_SOLO_ADMIN = {
  usuarios: true, crearUsuario: true, editarUsuario: true, eliminarUsuario: true,
  guardarCredencialesAdmin: true, guardarPasswordStock: true,
  guardarConfiguracionNegocio: true,
  cupones: true, guardarCupon: true, eliminarCupon: true, reporteCupones: true,
  estadoVentasResumidas: true, resumirVentasViejas: true,
  cerrarTodasLasSesiones: true,
  guardarDolarDeposito: true, guardarPreciosUsd: true, guardarConfigDocumentos: true,
  claveTraspasos: true, vincularDeposito: true, anularTraspaso: true
};

// Lo único que puede hacer la contraseña de stock.html (depósito desde el celular).
const ACCIONES_ROL_STOCK = {
  reportarError: true,
  productosAdmin: true, actualizarProducto: true, sumarStockProducto: true,
  producto: true, productos: true, categorias: true, stockDeCodigos: true,
  generarCodigoProducto: true, sesionActual: true
};

const SESION_TTL_PANEL_SEG       = 12 * 3600;       // navegador: 12 horas
const SESION_TTL_DISPOSITIVO_SEG = 30 * 24 * 3600;  // app de escritorio / app de Reportes: 30 días
const HASH_PREFIJO = "h1$";
const HASH_ITERACIONES = 40;

let _propsMemo = null;
let _sesionActual = null; // { u, r, k, n } de la solicitud en curso (null = pública)

function _props(){
  if(!_propsMemo) _propsMemo = PropertiesService.getScriptProperties().getProperties() || {};
  return _propsMemo;
}

/** Lee una propiedad del script; si no existe y se pasa un generador, la crea. */
function _propScript(clave, generar){
  const p = _props();
  if(p[clave]) return p[clave];
  if(!generar) return "";
  // Primera vez: si dos cajas llegan al mismo tiempo, cada una generaría
  // un secreto distinto y uno pisaría al otro (tokens o contraseñas que
  // dejan de validar). Se escribe, se espera un instante y se usa lo que
  // quedó guardado: todas las ejecuciones terminan con el mismo valor.
  const servicio = PropertiesService.getScriptProperties();
  const yaGuardado = servicio.getProperty(clave);
  if(yaGuardado){ p[clave] = yaGuardado; return yaGuardado; }
  servicio.setProperty(clave, generar());
  Utilities.sleep(700);
  const final = servicio.getProperty(clave);
  p[clave] = final;
  return final;
}

/** Crea de antemano los secretos de seguridad (lo llama instalarTriggersVeekPOS). */
function prepararSeguridad(){
  _propScript("VEEK_SECRETO_SESIONES", _secretoAleatorio);
  _propScript("VEEK_PIMIENTA", _secretoAleatorio);
  _epochSesiones("panel");
  _epochSesiones("stock");
  return "Seguridad lista.";
}

function _secretoAleatorio(){
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
}

function _hex(bytes){
  return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

/** Comparación en tiempo constante (no revela cuántos caracteres coinciden). */
function _iguales(a, b){
  a = String(a); b = String(b);
  let dif = a.length ^ b.length;
  const n = Math.max(a.length, b.length);
  for(let i = 0; i < n; i++) dif |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return dif === 0;
}

/* ---------------------------- Contraseñas ---------------------------- */

function esHashPassword(valor){
  return String(valor || "").indexOf(HASH_PREFIJO) === 0;
}

function hashPassword(password, sal){
  const pimienta = _propScript("VEEK_PIMIENTA", _secretoAleatorio);
  sal = sal || Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  let h = sal + ":" + String(password);
  for(let i = 0; i < HASH_ITERACIONES; i++){
    h = _hex(Utilities.computeHmacSha256Signature(h, pimienta));
  }
  return HASH_PREFIJO + sal + "$" + h;
}

/** true si `password` corresponde a lo guardado (hash nuevo o texto plano viejo). */
function verificarPassword(password, guardado){
  guardado = String(guardado || "");
  if(!guardado || password === undefined || password === null || password === "") return false;
  if(esHashPassword(guardado)){
    const sal = guardado.split("$")[1] || "";
    return _iguales(hashPassword(password, sal), guardado);
  }
  return _iguales(String(password), guardado);
}

/* ------------------------------ Tokens ------------------------------- */

function _epochSesiones(tipo){
  return _propScript("VEEK_EPOCH_" + String(tipo || "panel").toUpperCase(), () => "1");
}

/** Invalida TODOS los tokens de ese tipo ("panel" o "stock"). */
function cerrarSesionesTipo(tipo){
  const clave = "VEEK_EPOCH_" + String(tipo || "panel").toUpperCase();
  const nuevo = String(Date.now());
  PropertiesService.getScriptProperties().setProperty(clave, nuevo);
  _props()[clave] = nuevo;
}

function _b64(bytes){ return Utilities.base64EncodeWebSafe(bytes); }

function crearTokenSesion(info, ttlSeg){
  const tipo = info.tipo || "panel";
  const payload = {
    u: String(info.usuario || ""),
    r: String(info.rol || ""),
    k: tipo,
    a: info.adminClasico ? 1 : 0,
    e: Math.floor(Date.now() / 1000) + ttlSeg,
    v: _epochSesiones(tipo)
  };
  const cuerpo = _b64(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  const firma = _b64(Utilities.computeHmacSha256Signature(cuerpo, _propScript("VEEK_SECRETO_SESIONES", _secretoAleatorio)));
  return cuerpo + "." + firma;
}

/** Devuelve la sesión ({u, r, k, a, e}) o null si el token no sirve. */
function verificarTokenSesion(token){
  token = String(token || "");
  const partes = token.split(".");
  if(partes.length !== 2 || !partes[0] || !partes[1]) return null;

  const firma = _b64(Utilities.computeHmacSha256Signature(partes[0], _propScript("VEEK_SECRETO_SESIONES", _secretoAleatorio)));
  if(!_iguales(firma, partes[1])) return null;

  let p;
  try{ p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString()); }
  catch(e){ return null; }
  if(!p || !p.e || p.e * 1000 < Date.now()) return null;
  if(String(p.v) !== String(_epochSesiones(p.k))) return null;

  // Usuarios de la hoja USUARIOS: se confirma que sigan activos y se toma
  // el rol ACTUAL (si el admin le cambia el rol o lo desactiva, rige en
  // unos minutos sin esperar a que venza el token).
  if(p.k === "panel" && !p.a){
    const rolActual = mapaUsuariosActivos()[String(p.u).toLowerCase()];
    if(!rolActual) return null;
    p.r = rolActual;
  }
  return p;
}

const CACHE_USUARIOS_ACTIVOS = "usuarios_activos_v1";

/** { username_en_minusculas: rol } de los usuarios activos (cacheado 5 min). */
function mapaUsuariosActivos(){
  const cache = CacheService.getScriptCache();
  const guardado = cache.get(CACHE_USUARIOS_ACTIVOS);
  if(guardado){ try{ return JSON.parse(guardado); }catch(e){} }

  const mapa = {};
  const sheet = getHoja("USUARIOS");
  if(sheet && sheet.getLastRow() >= 2){
    rowsToObjects(sheet.getDataRange().getValues()).forEach(u => {
      const user = String(u.USERNAME || "").trim().toLowerCase();
      if(user && String(u.ACTIVO || "SI").toUpperCase() !== "NO"){
        mapa[user] = String(u.ROL || "vendedor").trim().toLowerCase();
      }
    });
  }
  try{ cache.put(CACHE_USUARIOS_ACTIVOS, JSON.stringify(mapa), 300); }catch(e){}
  return mapa;
}

function invalidarCacheUsuarios(){
  try{ CacheService.getScriptCache().remove(CACHE_USUARIOS_ACTIVOS); }catch(e){}
}

/* ------------------------- Control de acceso ------------------------- */

function respuestaNoAutorizado(mensaje, codigo){
  return jsonResponse({
    success: false,
    sesionInvalida: codigo === "SESION",
    sinPermiso: codigo === "PERMISO",
    codigoError: codigo,
    message: mensaje
  });
}

/**
 * Revisa si la acción se puede ejecutar. Devuelve null si está todo bien
 * (y deja la sesión en _sesionActual y el rol real en params.rol), o la
 * respuesta de error lista para devolver.
 */
function autorizarSolicitud(action, params){
  _sesionActual = null;
  params = params || {};

  if(ACCIONES_PUBLICAS[action]) {
    // Aunque sea pública, si viene un token válido se registra la sesión
    // (no cambia nada para el catálogo; sirve para auditoría).
    return null;
  }

  const sesion = verificarTokenSesion(params.token);
  if(!sesion){
    return respuestaNoAutorizado("Tu sesión venció o no es válida. Volvé a iniciar sesión.", "SESION");
  }

  if(sesion.k === "stock"){
    if(!ACCIONES_ROL_STOCK[action]){
      return respuestaNoAutorizado("La contraseña de stock no tiene acceso a esta acción.", "PERMISO");
    }
  } else if(ACCIONES_SOLO_ADMIN[action] && sesion.r !== "admin"){
    return respuestaNoAutorizado("Tu usuario no tiene permiso para esta acción.", "PERMISO");
  } else if(sesion.r === "vendedor" && /^reporte/i.test(action)){
    // Los reportes del negocio (ventas por período, clientes, formas de
    // pago...) no son parte de lo que ve el rol vendedor en el panel.
    return respuestaNoAutorizado("Tu usuario (Vendedor) no tiene acceso a los reportes.", "PERMISO");
  }

  _sesionActual = sesion;
  // El rol lo decide el servidor, no el navegador.
  params.rol = sesion.k === "stock" ? "deposito" : sesion.r;
  if(!params.usuario) params.usuario = sesion.u;

  // Operaciones en lote (cola offline): cada operación hereda el rol de
  // la sesión — salvo que venga marcada como "vendedor" y la sesión sea
  // admin (se respeta la restricción más fuerte).
  if(Array.isArray(params.operaciones)){
    params.operaciones.forEach(op => {
      if(!op || typeof op !== "object") return;
      op.params = op.params || {};
      const rolOp = String(op.params.rol || "").toLowerCase();
      if(!(sesion.r === "admin" && rolOp === "vendedor")) op.params.rol = params.rol;
      if(!op.params.usuario) op.params.usuario = sesion.u;
    });
  }
  return null;
}

/** Error inesperado: se registra completo en los logs y al cliente solo le llega un código. */
function respuestaErrorInterno(error, origen){
  const codigo = Utilities.getUuid().slice(0, 8).toUpperCase();
  console.error("[" + codigo + "] " + origen + ": " + (error && error.stack ? error.stack : error));
  if(REPORTES_AUTOMATICOS){
    try{
      enviarInformeError({
        tipo: "servidor",
        titulo: String(error && error.message || error).slice(0, 140),
        descripcion: "Error inesperado en el servidor al atender: " + origen,
        codigo: codigo,
        errores: [{ mensaje: String(error && error.message || error), pila: String(error && error.stack || "") }],
        automatico: true
      });
    }catch(e){ console.error("informe automático: " + e.message); }
  }
  // Falta aprobar un permiso de Google (típico después de pegar un Code.gs
  // nuevo que usa un servicio más): se dice qué hacer, no "error inesperado".
  if(esErrorDePermisos(error)){
    return jsonResponse({
      success: false,
      faltanPermisos: true,
      codigoError: codigo,
      message: "Falta aprobar permisos de Google en el Apps Script. El dueño de la planilla tiene que abrir Extensiones → Apps Script, elegir la función autorizarVeekPOS, tocar ▶ Ejecutar y aceptar los permisos. Después, volver a intentar."
    });
  }
  return jsonResponse({
    success: false,
    codigoError: codigo,
    message: "Ocurrió un error inesperado en el servidor (código " + codigo + "). Si se repite, pasale este código a soporte."
  });
}

function esErrorDePermisos(error){
  const t = String((error && error.message) || error || "");
  return /authoriz|autoriz|You do not have permission|No tienes permiso|No tenés permiso|Required permissions|Se necesitan permisos|permisos necesarios/i.test(t);
}

/** Para que el panel sepa con qué usuario/rol está la sesión (y si sigue vigente). */
function sesionActual(){
  const s = _sesionActual || {};
  return jsonResponse({ success:true, usuario: s.u || "", rol: s.r || "", tipo: s.k || "", vence: s.e ? s.e * 1000 : 0 });
}

/** Cierra todas las sesiones del panel (por ejemplo, si se perdió un celular). */
function cerrarTodasLasSesiones(){
  cerrarSesionesTipo("panel");
  cerrarSesionesTipo("stock");
  return jsonResponse({ success:true, message:"Se cerraron todas las sesiones. Cada equipo va a pedir iniciar sesión de nuevo." });
}

/**
 * ESCAPE DE EMERGENCIA — ejecutar a mano desde el editor de Apps Script
 * si nadie puede entrar al panel. Vuelve usuario/contraseña del admin a
 * CREDENCIALES_DEFAULT (que después hay que cambiar desde Configuración)
 * y cierra todas las sesiones.
 */
function resetearAccesoAdmin(){
  const sheet = getOrCrearHojaCredenciales();
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    const clave = String(data[i][0] || "").trim().toUpperCase();
    if(clave === "USUARIO")  sheet.getRange(i + 1, 2).setValue(CREDENCIALES_DEFAULT.USUARIO);
    if(clave === "PASSWORD") sheet.getRange(i + 1, 2).setValue(CREDENCIALES_DEFAULT.PASSWORD);
  }
  cerrarSesionesTipo("panel");
  cerrarSesionesTipo("stock");
  return "Acceso reseteado a " + CREDENCIALES_DEFAULT.USUARIO + " / " + CREDENCIALES_DEFAULT.PASSWORD + ". Cambialo apenas entres.";
}

function doGet(e) {
  try {
    e = e || {};
    e.parameter = e.parameter || {};
    const action = e.parameter.action || "productos";
    const denegado = autorizarSolicitud(action, e.parameter);
    if (denegado) return denegado;
    return doGetInterno(e);
  } catch (error) {
    return respuestaErrorInterno(error, "GET " + ((e && e.parameter && e.parameter.action) || ""));
  }
}

function doGetInterno(e) {
  if (e.parameter.action === "metricas") {

  // Cacheada unos segundos (igual que dashboard/pedidos): obtenerMetricas
  // lee cinco hojas completas, y varias cajas/pestañas la piden a la vez.
  // invalidarCachePolling() la limpia en cada alta o cambio.
  return conCachePolling("metricas", 15, function(){
    return ContentService
      .createTextOutput(JSON.stringify(obtenerMetricas()))
      .setMimeType(ContentService.MimeType.JSON);
  });

}
if (e.parameter.action === "guardarPedido") {

  const data = {
    nombre: e.parameter.nombre,
    empresa: e.parameter.empresa,
    direccion: e.parameter.direccion,
    telefono: e.parameter.telefono,
    dni: e.parameter.dni,
    total: e.parameter.total,
    codigoPostal: e.parameter.codigoPostal,
    localidad: e.parameter.localidad,
    provincia: e.parameter.provincia,
    carrito: JSON.parse(e.parameter.carrito || "[]"),
    idOperacion: e.parameter.idOperacion
  };

  return guardarPedidoPublico(data);
}
  const action = e.parameter.action || "productos";

  const respuestaDeposito = manejarAccionesDeposito(e.parameter);
  if (respuestaDeposito) return respuestaDeposito;

  switch (action) {
    case "pedidos":
  return getPedidos();

    case "buscarPedidos":
  return buscarPedidos(e.parameter || {});

    case "facturasPorRefs":
  return facturasPorRefs(e.parameter || {});

  case "login":
return login(e);

case "inicializarHojas": {
  // La llama la configuración inicial de la app de escritorio (setup.html)
  // al pegar la URL: deja la planilla lista (hojas + triggers). Pública e
  // idempotente (nunca borra nada), con un tope para que no se abuse.
  if(!permitirIntento("inicializar_hojas", 30, 600)){
    return jsonResponse({ success:false, message:"Demasiados intentos. Probá de nuevo en unos minutos." });
  }
  const primeraVez = asegurarInstalacionInicial();
  return primeraVez ? jsonResponse(primeraVez) : inicializarHojasNuevoNegocio();
}

case "loginStock":
  return loginStock(e);

case "guardarPasswordStock":
  return guardarPasswordStock(e);

case "dashboard":
return dashboard();

case "masVendidos":
  return getMasVendidos();

case "guardarVenta":
  // guardarVenta NO pasa por conIdempotencia a propósito — ya tiene su
  // propia protección contra duplicados vía clienteVentaId (revisa
  // solo las últimas ~300 filas de VENTAS_LOCAL, con su propio lock),
  // más barata que el ledger genérico para la operación más frecuente
  // de todo el sistema. Ver sincronizarVenta en sync-engine.js.
  return guardarVenta(e);

case "guardarPedidoAdmin":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return guardarPedidoAdmin(e); });

case "ventasPOS":
  return getVentasPOS();

case "ventasPOSHistorial":
  return getVentasPOSHistorial(e);

case "eliminarVenta":
  return eliminarVenta(e);

case "cierreCajaResumen":
  return getCierreCajaResumen(e);

case "guardarCierreCaja":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return guardarCierreCaja(e); });

case "historialCierres":
  return getHistorialCierres();

case "movimientosCajaHoy":
  return getMovimientosCajaHoy(e);

case "guardarMovimientoCaja":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return guardarMovimientoCaja(e); });

case "eliminarMovimientoCaja":
  return eliminarMovimientoCaja(e);

case "clientes":
  return getClientes();

case "todosClientes":
  return getTodosClientes();

case "clientesConCredito":
  return getClientesConCredito();

case "consultarClienteCreditoPorDni":
  return consultarClienteCreditoPorDni(e.parameter.dni);

case "detalleClienteCredito":
  return getDetalleClienteCredito(e.parameter.clienteId);

case "agotados":
  return getAgotados();

case "stockBajo":
  return getStockBajo();

case "actualizarEstado":
return conIdempotencia(e.parameter.idOperacionSync, function(){ return actualizarEstado(e); });

case "editarItemsPedido":
  return editarItemsPedido(e);

case "marcarPedidoCobrado":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return marcarPedidoCobrado(e); });

case "getDetallePedido":
  return getDetallePedido(e);

case "consultarPedido":
  return consultarPedido(e);

case "guardarProducto":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return guardarProducto(e); });

case "actualizarProducto":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return actualizarProducto(e); });

case "sumarStockProducto":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return sumarStockProducto(e); });

case "eliminarProducto":
  return conIdempotencia(e.parameter.idOperacionSync, function(){ return eliminarProductoPorCodigo(e); });

case "ingresosProductos":
  return getIngresosProductos(e);

case "boletasProveedores":
  return getBoletasProveedores(e);

case "detalleBoletaProveedor":
  return getDetalleBoletaProveedor(e);

case "proveedoresConSaldo":
  return getProveedoresConSaldo();

case "generarCodigoProducto":
  return jsonResponse({ success:true, codigo: generarCodigoProductoNuevo() });

case "usuarios":
  return getUsuarios();

case "configuracionNegocio":
  return conCachePolling("config_negocio", 300, getConfiguracionNegocio);

case "configLanding":
  return getConfigLanding();

case "guardarConfiguracionNegocio":
  return guardarConfiguracionNegocio(e);

case "guardarCredencialesAdmin":
  return guardarCredencialesAdmin(e);

case "reporteVentasPeriodo":
  return getReporteVentasPeriodo(e);

case "reporteProductosVendidos":
  return getReporteProductosVendidos(e);

case "reporteVentasPorCategoria":
  return getReporteVentasPorCategoria(e);

case "reporteVentasDiariasPorCategoria":
  return getReporteVentasDiariasPorCategoria(e);

case "reporteFormasPago":
  return getReporteFormasPago(e);

case "reporteCierresCaja":
  return getReporteCierresCaja(e);

case "reporteClientes":
  return getReporteClientes(e);

case "reportesVentas":
  return getReportesVentasJuntos(e);


    case "productos":
      return getProductos();

    case "stockDeCodigos":
      return getStockDeCodigos(e);

    case "catalogoVersion":
      return getCatalogoVersion();

    case "prepararPedido":
      return prepararPedido();

    case "stockTodos":
      return getStockTodos();

    case "productosAdmin":
      return getProductosAdmin();

    case "producto":
      return getProducto(e.parameter.codigo);

    case "validarCupon":
      return validarCuponPublico(e.parameter || {});

    case "cupones":
      return listarCupones();

    case "reporteCupones":
      return reporteCupones(e.parameter || {});

    case "avisoPedidos":
      return avisoPedidos();

    case "estadoVentasResumidas":
      return estadoVentasResumidas();

    case "categorias":
      return getCategorias();

    case "destacados":
      return getDestacados();

    case "ofertas":
      return getOfertas();

    case "imagenProxy":
      return getImagenProxy(e);

    case "catalogoPDFInfo":
      return getCatalogoPDFInfo();

    // QR de cobro por Mercado Pago — para la pantalla/tablet del
    // cliente (ver qrFijar/qrPagado/qrLimpiar/qrEstado arriba, y
    // main.js del lado de la app de escritorio).
    case "qrEstado":
      return qrEstado(e);

    case "qrPagado":
      return qrPagado();

    case "qrLimpiar":
      return qrLimpiar();

    case "qrPantalla":
      return HtmlService.createHtmlOutput(paginaPantallaQR())
        .setTitle("VeekPOS — Cobrá acá");

    case "sesionActual":
      return sesionActual();

    case "cerrarTodasLasSesiones":
      return cerrarTodasLasSesiones();

    default:
      return jsonResponse({
        success: false,
        message: "Acción no válida"
      });
  }
}
function doPost(e) {
  let data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || "{}") || {};
  } catch (err) {
    return jsonResponse({ success: false, message: "La solicitud no tiene un formato válido." });
  }
  try {
    const action = String(data.action || "");
    const denegado = autorizarSolicitud(action, data);
    if (denegado) return denegado;
    return doPostInterno(data);
  } catch (error) {
    return respuestaErrorInterno(error, "POST " + (data && data.action || ""));
  }
}

/**
 * Despacho de las acciones POST. Las acciones que no tienen un manejo
 * POST propio se resuelven con el mismo código del GET (el panel ahora
 * manda TODAS las modificaciones por POST — así no viajan datos ni el
 * token en la URL, y no hay límite de largo).
 */
function doPostInterno(data) {

// Papelera de productos / editar cliente del pedido / cliente desde el rótulo
// (ver sección "PAPELERA Y CLIENTES" al final de este archivo)
const respuestaExtra = manejarAccionesExtraPost(data);
if (respuestaExtra) return respuestaExtra;

// Depósito: dólar, precios en dólares, cotizaciones, traspasos (ver "DEPÓSITO")
const respuestaDeposito = manejarAccionesDeposito(data);
if (respuestaDeposito) return respuestaDeposito;

if (data.action === "guardarPedido") {
  return guardarPedidoPublico(data);
}

// Informes de error (ver "INFORMES DE ERROR")
// Configuración → Sistema → "Resumir ventas viejas"
if (data.action === "resumirVentasViejas") {
  return resumirVentasViejasPanel();
}

// Cambios en lote de Productos (ocultar/mostrar, destacado, oferta)
if (data.action === "actualizarProductosLote") {
  return conIdempotencia(data.idOperacionSync, function(){ return actualizarProductosLote(data); });
}

if (data.action === "reportarError") {
  return reportarErrorPanel(data);
}
if (data.action === "reportarErrorCatalogo") {
  return reportarErrorCatalogo(data);
}

// QR de cobro (ver arriba) — por POST porque la imagen del QR en
// base64 puede superar lo que entra en una URL de GET.
if (data.action === "qrFijar") {
  return qrFijar(data);
}

// Ventas del POS con carritos grandes: por GET, con todo en la URL,
// se puede superar lo que Google/Apps Script acepta (se sentía como
// "error de conexión" sin serlo). Por POST no hay ese límite.
if (data.action === "guardarVenta") {
  return guardarVenta(data);
}

// Pedidos creados desde "Generar Pedido" en el POS (panel admin), NO
// desde el catálogo web. Antes solo existía por GET (con el carrito
// en la URL) — funcionaba para pedidos chicos, pero con varios
// cientos de ítems la URL se vuelve enorme y el pedido fallaba con
// "error de conexión" (en realidad, la URL superaba lo que
// Google/Apps Script acepta, no un problema real de red). Por POST,
// con el carrito en el body en vez de la URL, no hay ese límite.
if (data.action === "guardarPedidoAdmin") {
  // Con idOperacionSync (cola offline): un reintento cuya primera
  // respuesta se perdió no crea un segundo pedido (ni un segundo PDF).
  return conIdempotencia(data.idOperacionSync, function(){ return guardarPedidoAdmin(data); });
}

// Varias operaciones de la cola offline en UNA sola ejecución (ver syncLote)
if (data.action === "syncLote") {
  return syncLote(data);
}

if (data.action === "editarItemsPedido") {
  return editarItemsPedido(data);
}

if (data.action === "subirImagenProducto") {
  return subirImagenProducto(data);
}

if (data.action === "crearCliente") {
  return conIdempotencia(data.idOperacionSync, function(){ return crearCliente(data); });
}

if (data.action === "editarCliente") {
  return conIdempotencia(data.idOperacionSync, function(){ return editarCliente(data); });
}

if (data.action === "eliminarCliente") {
  return conIdempotencia(data.idOperacionSync, function(){ return eliminarCliente(data); });
}

if (data.action === "actualizarCreditoCliente") {
  return actualizarCreditoCliente(data);
}

if (data.action === "crearClienteDesdeHistorialYMarcarCredito") {
  return crearClienteDesdeHistorialYMarcarCredito(data);
}

if (data.action === "registrarPagoCredito") {
  return conIdempotencia(data.idOperacionSync, function(){ return registrarPagoCredito(data); });
}

if (data.action === "crearDeudaExtra") {
  return conIdempotencia(data.idOperacionSync, function(){ return crearDeudaExtra(data); });
}

if (data.action === "eliminarPedidosCancelados") {
  return eliminarPedidosCancelados(data);
}

if (data.action === "registrarIngresoProducto") {
  // Con idOperacionSync (cola offline o reintento tras un corte): si la
  // boleta ya se había guardado, no se vuelve a sumar el stock.
  return conIdempotencia(data.idOperacionSync, function(){ return registrarIngresoProducto(data); });
}

if (data.action === "subirCatalogoPDF") {
  return subirCatalogoPDF(data);
}

if (data.action === "registrarFactura") {
  return registrarFactura(data);
}

if (data.action === "facturasPorRefs") {
  return facturasPorRefs(data);
}

if (data.action === "registrarPagoProveedor") {
  return conIdempotencia(data.idOperacionSync, function(){ return registrarPagoProveedor(data); });
}

if (data.action === "crearUsuario") {
  return crearUsuario(data);
}

if (data.action === "editarUsuario") {
  return editarUsuario(data);
}

if (data.action === "eliminarUsuario") {
  return eliminarUsuario(data);
}

if (data.action === "reordenarProductos") {
  return conIdempotencia(data.idOperacionSync, function(){ return reordenarProductos(data); });
}

// Sin manejo POST propio: mismo código que por GET (los parámetros
// llegan como strings, igual que en la URL).
return doGetInterno({ parameter: data });

}

/**
 * Reconciliación de stock LIVIANA — a diferencia de action=productos
 * (que arma/cachea el catálogo público entero: lee TODA la hoja,
 * filtra por PUBLICADO, ordena y guarda en caché por chunks), esto
 * solo devuelve CODIGO+STOCK de los códigos puntuales que se piden, y
 * no toca la caché de catálogo para nada. Pensado para el panel de
 * escritorio: después de sincronizar una venta o un ajuste de stock
 * offline, no hace falta releer el catálogo completo para saber si el
 * stock local quedó desactualizado — alcanza con preguntar por esos
 * pocos códigos puntuales. Antes esto SIEMPRE pasaba por
 * action=productos, y como cualquier escritura sobre productos invalida
 * esa caché (ver invalidarCacheProductos), cada sincronización — aunque
 * fuera de 1 sola unidad — terminaba forzando una reconstrucción
 * completa del catálogo, la operación más cara de todo el sistema.
 */
function getStockDeCodigos(e){
  const params = (e && e.parameter) || {};
  const codigosPedidos = String(params.codigos || "").split(",").map(c => c.trim()).filter(Boolean);

  if(!codigosPedidos.length){
    return jsonResponse({ success:false, message:"Falta el parámetro 'codigos'." });
  }

  const sheet = getSheet();
  if(!sheet) return jsonResponse({ success:false, message:"No se encontró la hoja 'PRODUCTOS'." });

  // Solo las columnas CODIGO y STOCK, no la hoja entera
  const ultimaFila = sheet.getLastRow();
  const ultimaCol = sheet.getLastColumn();
  if(ultimaFila < 2) return jsonResponse({ success:true, productos: [] });
  const headerRow = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h || "").trim());
  const idxCodigo = headerRow.indexOf("CODIGO");
  const idxStock  = headerRow.indexOf("STOCK");

  if(idxCodigo === -1 || idxStock === -1){
    return jsonResponse({ success:false, message:"La hoja 'PRODUCTOS' no tiene las columnas CODIGO/STOCK esperadas." });
  }

  const colCodigos = sheet.getRange(2, idxCodigo + 1, ultimaFila - 1, 1).getValues();
  const colStock = sheet.getRange(2, idxStock + 1, ultimaFila - 1, 1).getValues();

  const buscados = new Set(codigosPedidos.map(c => c.toUpperCase()));
  const encontrados = [];

  for(let i = 0; i < colCodigos.length; i++){
    const codigoFila = String(colCodigos[i][0] || "").trim();
    if(buscados.has(codigoFila.toUpperCase())){
      encontrados.push({ CODIGO: codigoFila, STOCK: Number(colStock[i][0] || 0) });
    }
  }

  return jsonResponse({ success:true, productos: encontrados });
}

function getSheet() {
  return getHoja(SHEET_NAME);
}

/* ===================================================================
   IDEMPOTENCIA DE SINCRONIZACIÓN
   ===================================================================
   El panel de escritorio (app offline) puede reintentar una operación
   que en realidad YA se aplicó del lado del servidor — no porque el
   servidor haya fallado, sino porque la respuesta se perdió en el
   camino de vuelta (ej. se cortó la conexión justo después de que
   Sheets terminara de escribir, o el cliente tuvo que abortar el
   request porque el LockService tardó más de lo que el cliente estaba
   dispuesto a esperar). Sin esto, ese reintento vuelve a ejecutar la
   MISMA operación — inofensivo para algo como actualizarProducto
   (vuelve a escribir los mismos valores), pero no para algo como
   sumarStockProducto, que SUMA un delta: aplicarlo dos veces duplica
   el ajuste de stock.

   La solución: el cliente manda una clave única por operación
   (idOperacionSync, basada en el id local de la cola — ver
   sync-engine.js) junto con el resto de los parámetros. Antes de
   ejecutar la acción, se busca esa clave en este ledger — si ya está,
   se devuelve la respuesta guardada de la vez anterior SIN volver a
   tocar la planilla. Si no está, se ejecuta normal y, solo si terminó
   con éxito, se guarda la respuesta para la próxima vez.

   Solo se cachean resultados EXITOSOS — un error de negocio (ej.
   "falta el nombre del producto") no muta nada, así que no hace falta
   protegerlo, y cachearlo sería incorrecto si la causa del error se
   corrige y se reintenta con la misma clave.
=================================================================== */

const HOJA_IDEMPOTENCIA = "SYNC_IDEMPOTENCIA";

function getHojaIdempotencia() {
  let hoja = getHoja(HOJA_IDEMPOTENCIA);
  if (!hoja) {
    hoja = getSS().insertSheet(HOJA_IDEMPOTENCIA);
    hoja.appendRow(["CLAVE", "FECHA", "RESPUESTA"]);
    hoja.hideSheet();
    _hojaCache[HOJA_IDEMPOTENCIA] = hoja;
  }
  return hoja;
}

const TTL_CACHE_IDEMPOTENCIA_SEG = 21600; // 6h — máximo de CacheService
let _clavesLedgerCargadas = null;  // Set con TODAS las claves del ledger (solo dentro de syncLote)
let _filasLedgerPendientes = null; // filas a escribir juntas al final de syncLote

/**
 * Devuelve el texto de la respuesta ya guardada para esta clave, o null
 * si nunca se procesó (o no se mandó clave). Orden de búsqueda:
 *   1. CacheService (casi todos los reintentos caen acá: pasan a los
 *      segundos/minutos del intento original) — sin tocar la planilla.
 *   2. Dentro de syncLote: el Set de claves leído UNA vez por lote
 *      (una sola lectura de la columna A, en vez de un TextFinder por fila).
 *   3. TextFinder sobre la hoja (solo si la clave realmente existe o
 *      fuera de un lote).
 */
function buscarOperacionYaProcesada(clave) {
  if (!clave) return null;
  try {
    const enCache = CacheService.getScriptCache().get("idem:" + clave);
    if (enCache) return enCache;
  } catch (e) {}

  if (_clavesLedgerCargadas && !_clavesLedgerCargadas.has(clave)) return null;

  const hoja = getHojaIdempotencia();
  const celda = hoja.createTextFinder(clave).matchEntireCell(true).useRegularExpression(false).findNext();
  if (!celda) return null;
  const respuestaGuardada = hoja.getRange(celda.getRow(), 3).getValue();
  return respuestaGuardada ? String(respuestaGuardada) : null;
}

function cargarClavesLedger() {
  const hoja = getHojaIdempotencia();
  const ultima = hoja.getLastRow();
  const set = new Set();
  if (ultima >= 2) {
    hoja.getRange(2, 1, ultima - 1, 1).getValues().forEach(f => { if (f[0]) set.add(String(f[0])); });
  }
  _clavesLedgerCargadas = set;
}

function registrarOperacionProcesada(clave, textoRespuesta) {
  if (!clave) return;
  try {
    if (String(textoRespuesta).length < 90000) {
      CacheService.getScriptCache().put("idem:" + clave, String(textoRespuesta), TTL_CACHE_IDEMPOTENCIA_SEG);
    }
  } catch (e) {}

  const fila = [clave, new Date().toISOString(), textoRespuesta];
  if (_filasLedgerPendientes) {
    _filasLedgerPendientes.push(fila);
    if (_clavesLedgerCargadas) _clavesLedgerCargadas.add(clave);
    return;
  }
  getHojaIdempotencia().appendRow(fila);
  // La limpieza ya no se hace al voleo acá (deleteRow fila por fila era
  // carísimo): corre en un trigger diario — ver limpiarLedgerIdempotencia.
}

/** Escribe de una sola vez las filas acumuladas por syncLote */
function volcarFilasLedgerPendientes() {
  if (!_filasLedgerPendientes || _filasLedgerPendientes.length === 0) return;
  const hoja = getHojaIdempotencia();
  hoja.getRange(hoja.getLastRow() + 1, 1, _filasLedgerPendientes.length, 3).setValues(_filasLedgerPendientes);
  _filasLedgerPendientes = [];
}

/**
 * Trigger diario (ver instalarTriggersVeekPOS): borra del ledger las
 * entradas de más de 60 días reescribiendo las filas que quedan en UNA
 * sola operación, en vez de deleteRow() por cada fila vieja.
 */
function limpiarLedgerIdempotencia() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    limpiarEntradasViejasIdempotencia(getHojaIdempotencia());
  } finally {
    lock.releaseLock();
  }
  // Mismo trigger diario: la hoja ERRORES guarda solo la última semana.
  try { limpiarHojaErrores(); } catch (e) { console.error("limpiarHojaErrores: " + e.message); }
  // Y una vez por mes: resumir las ventas de hace más de MESES_DETALLE_VENTAS meses.
  try { compactarVentasProgramado(); } catch (e) { console.error("compactarVentasProgramado: " + e.message); }
}

function limpiarEntradasViejasIdempotencia(hoja) {
  const LIMITE_DIAS = 60;
  const limiteMs = LIMITE_DIAS * 24 * 60 * 60 * 1000;
  const ahora = Date.now();
  const ultima = hoja.getLastRow();
  if (ultima < 2) return;
  const datos = hoja.getRange(2, 1, ultima - 1, 3).getValues();

  const conservar = datos.filter(f => {
    const fecha = new Date(f[1]);
    return isNaN(fecha) || (ahora - fecha.getTime()) <= limiteMs;
  });
  if (conservar.length === datos.length) return;

  hoja.getRange(2, 1, ultima - 1, 3).clearContent();
  if (conservar.length > 0) hoja.getRange(2, 1, conservar.length, 3).setValues(conservar);
}

/**
 * Envuelve un handler de acción (guardarVenta, sumarStockProducto,
 * etc.) con el chequeo de idempotencia de arriba. `clave` es el
 * idOperacionSync que mandó el cliente (puede venir undefined/vacío
 * si la acción se llamó en vivo desde el panel online, no desde la
 * cola offline — en ese caso esto no hace nada distinto de llamar al
 * handler directo).
 *
 * IMPORTANTE — por qué todo esto pasa por un lock: la primera versión
 * de esto hacía "buscar en el ledger" y "guardar en el ledger" como
 * dos pasos sueltos, sin nada que los protegiera de ejecutarse pisados
 * entre sí. Si el cliente reintentaba una operación varias veces
 * seguidas porque no le llegaba la respuesta a tiempo (ej. timeout
 * nuestro, aunque el servidor sí la hubiera terminado), esos
 * reintentos podían llegar SUPERPUESTOS: los dos (o tres) pasaban el
 * chequeo "todavía no está en el ledger" casi al mismo tiempo, porque
 * ninguno había terminado de registrarse todavía — y los dos
 * terminaban ejecutando la operación de verdad, aunque fuera la misma.
 * Para un ajuste de stock, eso duplicaba (o triplicaba) el ajuste.
 *
 * Por eso ahora el chequeo, la ejecución del handler, Y el registro en
 * el ledger quedan TODOS adentro del mismo lock — un reintento
 * superpuesto tiene que esperar a que el primero termine de
 * registrarse antes de poder siquiera mirar el ledger, así que
 * siempre lo va a encontrar ya escrito.
 */
/**
 * Bandera de "ya tenemos el lock de script tomado por conIdempotencia".
 * guardarProducto / actualizarProducto / sumarStockProducto (los tres
 * handlers que conIdempotencia envuelve para producto/stock) TAMBIÉN
 * toman su propio LockService.getScriptLock() internamente — eso está
 * bien cuando se llaman directo (panel web, sin idOperacionSync, sin
 * pasar por acá), pero si se llaman DESDE ADENTRO de conIdempotencia
 * (cola offline de Electron, que siempre manda idOperacionSync), el
 * lock de acá ya está tomado por esta misma ejecución — el script lock
 * de Apps Script no es reentrante, así que el segundo tryLock() de
 * adentro del handler siempre termina agotando su propio timeout y
 * devolviendo success:false, sin llegar a escribir nada. Esta bandera
 * le avisa a esos tres handlers que salteen su propio tryLock cuando
 * ya estamos protegidos por el de acá.
 */
let _dentroDeLockIdempotencia = false;

function conIdempotencia(clave, funcionHandler) {
  if (!clave) return funcionHandler();

  const lock = LockService.getScriptLock();
  // Margen generoso: puede tener que esperar tanto a que termine el
  // lock propio del handler (hasta 10-20s, ver los tryLock de más
  // abajo) como a que termine de registrarse en el ledger.
  const obtuvoLock = lock.tryLock(30000);

  if (!obtuvoLock) {
    console.error("No se pudo obtener el lock de idempotencia para: " + clave);
    return jsonResponse({
      success:false,
      message:"El sistema está ocupado procesando esta misma operación, probá de nuevo en unos segundos."
    });
  }

  _dentroDeLockIdempotencia = true;

  try {
    const previa = buscarOperacionYaProcesada(clave);
    if (previa !== null) return ContentService.createTextOutput(previa).setMimeType(ContentService.MimeType.JSON);

    const salida = funcionHandler();

    try {
      const texto = salida.getContent();
      const parseado = JSON.parse(texto);
      if (parseado && parseado.success) registrarOperacionProcesada(clave, texto);
    } catch (err) {
      // Si no se pudo leer o parsear la salida, mejor no guardar nada
      // en el ledger — no cachear es más seguro que cachear mal.
      console.error("No se pudo registrar la operación en el ledger de idempotencia:", err);
    }

    return salida;

  } finally {
    _dentroDeLockIdempotencia = false;
    lock.releaseLock();
  }
}


/* =====================================================================
   SYNC POR LOTES
   La cola offline de cada caja antes subía fila por fila: N filas
   pendientes = N ejecuciones de Apps Script, cada una abriendo la
   planilla, tomando el lock, buscando en el ledger con TextFinder y
   (si era una venta) leyendo PRODUCTOS completa para descontar stock.
   Ahora la caja manda hasta 25 operaciones juntas y acá se procesan
   bajo UN lock, con:
     - una sola lectura de las claves del ledger,
     - escritura del ledger en un solo setValues al final,
     - el stock de todas las ventas aplicado en una sola pasada,
     - la respuesta incluye el stock final de los códigos tocados, así
       la caja no necesita pedirlo aparte.
   Cada operación conserva su propia idempotencia: si el lote se corta
   a mitad, las ya aplicadas no se repiten al reintentar.
===================================================================== */
const SYNC_LOTE_MAX_OPERACIONES = 25;
const SYNC_LOTE_PRESUPUESTO_MS = 40000; // deja margen bajo el timeout del cliente (60s)

const ACCIONES_SYNC_LOTE = {
  guardarPedidoAdmin:     p => guardarPedidoAdmin(p),
  guardarMovimientoCaja:  p => guardarMovimientoCaja({ parameter: p }),
  sumarStockProducto:     p => sumarStockProducto({ parameter: p }),
  guardarProducto:        p => guardarProducto({ parameter: p }),
  actualizarProducto:     p => actualizarProducto({ parameter: p }),
  eliminarProducto:       p => eliminarProductoPorCodigo({ parameter: p }),
  crearCliente:           p => crearCliente(p),
  editarCliente:          p => editarCliente(p),
  eliminarCliente:        p => eliminarCliente(p),
  registrarPagoCredito:   p => registrarPagoCredito(p),
  crearDeudaExtra:        p => crearDeudaExtra(p),
  guardarCierreCaja:      p => guardarCierreCaja({ parameter: p }),
  registrarIngresoProducto: p => registrarIngresoProducto(p),
  registrarPagoProveedor: p => registrarPagoProveedor(p),
  actualizarProductosLote: p => actualizarProductosLote(p),
  reordenarProductos:     p => reordenarProductos(p),
  actualizarEstado:       p => actualizarEstado({ parameter: p }),
  marcarPedidoCobrado:    p => marcarPedidoCobrado({ parameter: p })
};

function syncLote(data){
  const operaciones = Array.isArray(data.operaciones) ? data.operaciones.slice(0, SYNC_LOTE_MAX_OPERACIONES) : [];
  if(operaciones.length === 0) return jsonResponse({ success:true, resultados: [], stock: {} });

  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)){
    return jsonResponse({ success:false, ocupado:true, message:"El servidor está ocupado, se reintenta en el próximo ciclo." });
  }

  const inicio = Date.now();
  const resultados = [];
  let stockFinal = {};

  _dentroDeLockIdempotencia = true;
  _deltasStockDiferidos = {};
  _filasLedgerPendientes = [];

  try{
    cargarClavesLedger();

    for(let i = 0; i < operaciones.length; i++){
      if(Date.now() - inicio > SYNC_LOTE_PRESUPUESTO_MS) break; // el resto queda pendiente para el próximo lote
      const op = operaciones[i] || {};
      const fijaStock = (op.accion === "actualizarProducto" || op.accion === "guardarProducto") &&
        op.params && op.params.STOCK !== undefined && op.params.STOCK !== null && String(op.params.STOCK) !== "";
      if(fijaStock) Object.assign(stockFinal, _aplicarDeltasDiferidosAhora());
      const r = procesarOperacionSyncLote(op);
      resultados.push({ clave: String(op.clave || ""), resultado: r });
      if(r && r.stockNuevo !== undefined && r.codigo) stockFinal[String(r.codigo)] = Number(r.stockNuevo) || 0;
      if(r && r.success && r.stock && typeof r.stock === "object") Object.keys(r.stock).forEach(c => { stockFinal[c] = Number(r.stock[c]) || 0; });
    }

  } finally {
    const deltas = _deltasStockDiferidos || {};
    _deltasStockDiferidos = null;
    try{
      if(Object.keys(deltas).length > 0){
        const nuevos = aplicarDeltasStock(deltas, false);
        Object.keys(nuevos).forEach(c => { stockFinal[c] = nuevos[c]; });
        notificarCambioStock(nuevos);
      }
    }catch(err){
      console.error("syncLote: no se pudo aplicar el stock del lote: " + err + " deltas=" + JSON.stringify(deltas));
    }
    try{ volcarFilasLedgerPendientes(); }catch(err){ console.error("syncLote: no se pudo escribir el ledger: " + err); }
    _filasLedgerPendientes = null;
    _clavesLedgerCargadas = null;
    _dentroDeLockIdempotencia = false;
    lock.releaseLock();
  }

  return jsonResponse({
    success: true,
    resultados: resultados,
    stock: stockFinal,
    versiones: leerVersionesCatalogo()
  });
}

/**
 * Las ventas del lote descuentan stock al final (una sola pasada). Si en
 * el mismo lote viene una edición que FIJA el stock (STOCK explícito), hay
 * que aplicar antes los descuentos ya acumulados: si no, se aplicarían
 * DESPUÉS del valor fijado y la venta se descontaría dos veces.
 */
function _aplicarDeltasDiferidosAhora(){
  if(!_deltasStockDiferidos) return {};
  const d = _deltasStockDiferidos;
  if(Object.keys(d).length === 0) return {};
  _deltasStockDiferidos = {};
  const nuevos = aplicarDeltasStock(d, false);
  try{ notificarCambioStock(nuevos); }catch(e){}
  return nuevos;
}

function procesarOperacionSyncLote(op){
  const accion = String(op.accion || "");
  const params = op.params || {};
  const clave = String(op.clave || "");

  try{
    if(accion === "guardarVenta"){
      // Idempotencia propia por clienteVentaId (más barata que el ledger)
      const p = Object.assign({}, params);
      if(!p.clienteVentaId) p.clienteVentaId = clave;
      return guardarVentaNucleo(p);
    }

    const handler = ACCIONES_SYNC_LOTE[accion];
    if(!handler) return { success:false, message:"Acción no permitida en lote: " + accion };

    if(clave){
      const previa = buscarOperacionYaProcesada(clave);
      if(previa !== null){
        try{ return JSON.parse(previa); }catch(e){ return { success:true, idempotente:true }; }
      }
    }

    const p = Object.assign({}, params, { idOperacionSync: clave });
    const salida = handler(p);
    const texto = salida.getContent();
    let obj;
    try{ obj = JSON.parse(texto); }catch(e){ return { success:false, message:"Respuesta inválida del handler " + accion }; }
    if(obj && obj.success && clave) registrarOperacionProcesada(clave, texto);
    return obj;

  }catch(err){
    return { success:false, message:"Error interno en " + accion + ": " + err };
  }
}

/* =====================================================================
   TRIGGERS — ejecutar UNA vez instalarTriggersVeekPOS() desde el editor
===================================================================== */

/* Horario comercial: ver HORA_APERTURA / HORA_CIERRE al principio del archivo. */

function enHorarioComercial(){
  const hora = Number(Utilities.formatDate(new Date(), zonaNegocio(), "H"));
  return hora >= HORA_APERTURA && hora < HORA_CIERRE;
}

/**
 * Instala todos los triggers de mantenimiento (reemplaza los anteriores
 * con el mismo handler, así se puede volver a ejecutar sin duplicar):
 *  - calentarCacheCatalogoProgramado: cada 30 min, solo en horario
 *    comercial (antes: cada 10 min, las 24hs = 144 ejecuciones/día).
 *  - limpiarLedgerIdempotencia: una vez por día, de madrugada.
 *  - actualizarResumenMetricas: una vez por día, de madrugada.
 */
/**
 * Una sola vez por negocio: crea las hojas con sus encabezados
 * (inicializarHojasNuevoNegocio) e instala los triggers. Se llama al
 * iniciar sesión; después de la primera vez solo lee una propiedad.
 * Si la planilla ya tenía todo (cliente existente), no toca nada.
 */
const PROP_INSTALACION_OK = "VEEK_INSTALACION_OK";
function asegurarInstalacionInicial(){
  try{
    if(_props()[PROP_INSTALACION_OK]) return null;
    const lock = LockService.getScriptLock();
    if(!lock.tryLock(15000)) return null; // otra solicitud la está haciendo
    try{
      const props = PropertiesService.getScriptProperties();
      if(props.getProperty(PROP_INSTALACION_OK)) return null;
      const resultado = JSON.parse(inicializarHojasNuevoNegocio().getContent());
      // Instalación NUEVA (la planilla no tenía ni PRODUCTOS): se arman
      // también las carpetas de Drive. En un negocio que ya funcionaba no
      // se tocan (puede tener carpetas vacías a propósito).
      if((resultado.creadas || []).indexOf("PRODUCTOS") !== -1){
        try{
          resultado.carpetas = crearCarpetasDriveFaltantes();
        }catch(errCarp){
          resultado.carpetas = { error: errCarp.message };
          console.warn("asegurarInstalacionInicial: no se pudieron crear las carpetas de Drive (ejecutar crearCarpetasDriveFaltantes a mano): " + errCarp.message);
        }
      }
      let triggers = "ok";
      try{
        const yaHay = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "calentarCacheCatalogoProgramado");
        if(!yaHay) instalarTriggersVeekPOS();
      }catch(errTrig){
        triggers = "pendiente: " + errTrig.message;
        console.warn("asegurarInstalacionInicial: no se pudieron instalar los triggers (ejecutar instalarTriggersVeekPOS a mano): " + errTrig.message);
      }
      props.setProperty(PROP_INSTALACION_OK, new Date().toISOString() + " · triggers " + triggers);
      if(_propsMemo) _propsMemo[PROP_INSTALACION_OK] = "1";
      resultado.triggers = triggers;
      return resultado;
    } finally {
      lock.releaseLock();
    }
  }catch(err){
    console.error("asegurarInstalacionInicial: " + err);
  }
  return null;
}

function instalarTriggersVeekPOS(){
  _retirarCatalogoPDFViejo();
  const handlers = ["calentarCacheCatalogo", "calentarCacheCatalogoProgramado", "limpiarLedgerIdempotencia", "actualizarResumenMetricas", "archivarVentasProgramado"];
  ScriptApp.getProjectTriggers().forEach(t => {
    if(handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  prepararSeguridad();
  ScriptApp.newTrigger("calentarCacheCatalogoProgramado").timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger("limpiarLedgerIdempotencia").timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger("actualizarResumenMetricas").timeBased().everyDays(1).atHour(4).create();
  if(Number(ARCHIVAR_VENTAS_MESES) > 0){
    ScriptApp.newTrigger("archivarVentasProgramado").timeBased().onMonthDay(1).atHour(5).create();
  }
  calentarCacheCatalogo();
  actualizarResumenMetricas();
  Logger.log("Triggers de VeekPOS instalados.");
}

function calentarCacheCatalogoProgramado(){
  if(!enHorarioComercial()) return;
  calentarCacheCatalogo();
}

/* =====================================================================
   RESUMEN DE MÉTRICAS (snapshot diario)
   obtenerMetricas() leía cinco hojas COMPLETAS (PEDIDOS, VENTAS_LOCAL,
   PRODUCTOS, DETALLE_PEDIDOS, DETALLE_VENTAS) en cada consulta del
   dashboard. Con meses de historial eso crece sin parar. Ahora un
   trigger diario guarda:
     - desde qué fila de VENTAS_LOCAL empieza el mes actual y cuántas
       ventas POS hubo antes (en ScriptProperties),
     - las unidades vendidas por código en DETALLE_VENTAS hasta la
       última fila procesada (hoja oculta RESUMEN_METRICAS),
   y la consulta solo lee lo que se agregó después.
===================================================================== */
const HOJA_RESUMEN_METRICAS = "RESUMEN_METRICAS";
const PROP_RESUMEN_METRICAS = "RESUMEN_METRICAS_META";

function _mesActualClave(){
  return Utilities.formatDate(new Date(), zonaNegocio(), "yyyy-MM");
}

function _columnaAnulada(sheet){
  const ultimaCol = sheet.getLastColumn();
  if(ultimaCol < 1) return -1;
  return sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h || "").trim()).indexOf("ANULADA");
}

/**
 * Cuántas ventas representa una fila de VENTAS_LOCAL: 1, o N si es un
 * resumen de ventas viejas (columna VENTAS_RESUMIDAS, ver
 * compactarVentasAntiguas). idx = índice de esa columna o -1.
 */
function _pesoVenta(fila, idx){
  if(idx === -1 || idx === undefined) return 1;
  const n = Number(fila[idx]);
  return n > 0 ? n : 1;
}

function _columnaPesoVenta(sheet){
  const ultimaCol = sheet.getLastColumn();
  if(ultimaCol < 1) return -1;
  return sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h || "").trim()).indexOf("VENTAS_RESUMIDAS");
}

function actualizarResumenMetricas(){
  const zona = zonaNegocio();
  const hoy = new Date();
  const mesActual = Utilities.formatDate(hoy, zona, "yyyy-MM");

  const ventas = getHoja("VENTAS_LOCAL");
  let ventasDesdeFila = 2;
  let totalVentasPOSAntes = 0;
  if(ventas){
    const ultima = ventas.getLastRow();
    const idxAnulada = _columnaAnulada(ventas);
    const idxPeso = _columnaPesoVenta(ventas);
    const nCols = Math.max(5, idxAnulada + 1, idxPeso + 1);
    ventasDesdeFila = ultima + 1;
    if(ultima >= 2){
      const filas = ventas.getRange(2, 1, ultima - 1, nCols).getValues();
      let encontrado = false;
      for(let i = 0; i < filas.length; i++){
        const f = new Date(filas[i][1]);
        const mesFila = isNaN(f) ? "" : Utilities.formatDate(f, zona, "yyyy-MM");
        if(!encontrado && mesFila === mesActual){ ventasDesdeFila = i + 2; encontrado = true; }
        if(encontrado) break;
        const anulada = idxAnulada !== -1 && String(filas[i][idxAnulada] || "").toUpperCase() === "SI";
        const esPedido = String(filas[i][4] || "").toUpperCase() === "PEDIDO WEB";
        if(!anulada && !esPedido) totalVentasPOSAntes += _pesoVenta(filas[i], idxPeso);
      }
    }
  }

  // Unidades por código en DETALLE_VENTAS (solo 4 columnas)
  const detalle = getHoja("DETALLE_VENTAS");
  const acumulado = {};
  let detalleHastaFila = 1;
  if(detalle){
    detalleHastaFila = detalle.getLastRow();
    if(detalleHastaFila >= 2){
      detalle.getRange(2, 2, detalleHastaFila - 1, 3).getValues().forEach(r => {
        const c = String(r[0] || "").trim();
        if(!c) return;
        if(!acumulado[c]) acumulado[c] = { nombre: r[1], cantidad: 0 };
        acumulado[c].cantidad += Number(r[2]) || 0;
      });
    }
  }

  let hoja = getHoja(HOJA_RESUMEN_METRICAS);
  if(!hoja){
    hoja = getSS().insertSheet(HOJA_RESUMEN_METRICAS);
    hoja.hideSheet();
    _hojaCache[HOJA_RESUMEN_METRICAS] = hoja;
  }
  hoja.clearContents();
  const filasResumen = [["CODIGO", "PRODUCTO", "CANTIDAD"]].concat(
    Object.keys(acumulado).map(c => [c, acumulado[c].nombre, acumulado[c].cantidad])
  );
  hoja.getRange(1, 1, filasResumen.length, 3).setValues(filasResumen);

  const meta = {
    mes: mesActual,
    calculadoEn: hoy.toISOString(),
    ventasDesdeFila: ventasDesdeFila,
    totalVentasPOSAntes: totalVentasPOSAntes,
    detalleHastaFila: detalleHastaFila
  };
  PropertiesService.getScriptProperties().setProperty(PROP_RESUMEN_METRICAS, JSON.stringify(meta));
  return meta;
}

function _leerResumenMetricas(){
  let meta = null;
  try{ meta = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_RESUMEN_METRICAS) || "null"); }catch(e){ meta = null; }
  const ventas = getHoja("VENTAS_LOCAL");
  const detalle = getHoja("DETALLE_VENTAS");
  const invalido = !meta || meta.mes !== _mesActualClave() ||
    (ventas && ventas.getLastRow() < meta.ventasDesdeFila - 1) ||
    (detalle && detalle.getLastRow() < meta.detalleHastaFila);
  if(invalido){
    const lock = LockService.getScriptLock();
    if(lock.tryLock(20000)){
      try{ meta = actualizarResumenMetricas(); } finally { lock.releaseLock(); }
    } else if(!meta){
      return null;
    }
  }
  return meta;
}

/**
 * Versión liviana de obtenerMetricas: mismo resultado, pero leyendo
 * solo el mes en curso de VENTAS_LOCAL, las columnas necesarias de
 * PEDIDOS/DETALLE_PEDIDOS, el catálogo desde la caché, y el ranking
 * de más vendidos desde el resumen + las filas nuevas.
 * Si el resumen no está disponible, cae a obtenerMetricasCompleto().
 */
function obtenerMetricas(){
  let meta = null;
  try{ meta = _leerResumenMetricas(); }catch(e){ meta = null; }
  if(!meta) return obtenerMetricasCompleto();

  const zona = zonaNegocio();
  const hoy = new Date();
  const hoyClave = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");
  const mesClave = Utilities.formatDate(hoy, zona, "yyyy-MM");

  // ---- PEDIDOS (solo las primeras 9 columnas: NOMBRE y ESTADO)
  let pedidosNuevos = 0, totalPedidos = 0;
  const clientes = new Set();
  const hojaPedidos = getHoja("PEDIDOS");
  if(hojaPedidos && hojaPedidos.getLastRow() >= 2){
    hojaPedidos.getRange(2, 1, hojaPedidos.getLastRow() - 1, 9).getValues().forEach(f => {
      totalPedidos++;
      clientes.add(f[2]);
      if(String(f[8] || "").toLowerCase() === "nuevo") pedidosNuevos++;
    });
  }

  // ---- VENTAS_LOCAL: solo desde el inicio del mes
  let ventasPOSHoy = 0, ventasPOSMes = 0, totalVentasPOS = meta.totalVentasPOSAntes || 0;
  let ventasPedidosHoy = 0, ventasPedidosMes = 0;
  const hojaVentas = getHoja("VENTAS_LOCAL");
  if(hojaVentas){
    const ultima = hojaVentas.getLastRow();
    const desde = Math.max(2, meta.ventasDesdeFila);
    if(ultima >= desde){
      const idxAnulada = _columnaAnulada(hojaVentas);
      const idxPeso = _columnaPesoVenta(hojaVentas);
      const nCols = Math.max(5, idxAnulada + 1, idxPeso + 1);
      hojaVentas.getRange(desde, 1, ultima - desde + 1, nCols).getValues().forEach(f => {
        if(idxAnulada !== -1 && String(f[idxAnulada] || "").toUpperCase() === "SI") return;
        const fecha = new Date(f[1]);
        if(isNaN(fecha)) return;
        const total = Number(f[2]) || 0;
        const esPedido = String(f[4] || "").toUpperCase() === "PEDIDO WEB";
        const esHoy = Utilities.formatDate(fecha, zona, "yyyy-MM-dd") === hoyClave;
        const esMes = Utilities.formatDate(fecha, zona, "yyyy-MM") === mesClave;
        if(esPedido){
          if(esHoy) ventasPedidosHoy += total;
          if(esMes) ventasPedidosMes += total;
        } else {
          totalVentasPOS += _pesoVenta(f, idxPeso);
          if(esHoy) ventasPOSHoy += total;
          if(esMes) ventasPOSMes += total;
        }
      });
    }
  }

  // ---- PRODUCTOS: desde la caché del catálogo admin si está tibia
  let productosActivos = 0, stockBajo = 0, agotados = 0;
  let listaProductos = null;
  const cacheAdmin = cacheLeerTexto("catalogo_admin");
  if(cacheAdmin){ try{ listaProductos = JSON.parse(cacheAdmin).productos || null; }catch(e){ listaProductos = null; } }
  if(listaProductos){
    listaProductos.forEach(p => {
      if(String(p.CODIGO || "").trim() === "") return;
      const stock = Number(p.STOCK) || 0;
      if(String(p.PUBLICADO || "").toLowerCase() === "si") productosActivos++;
      if(stock === 0) agotados++;
      if(stock > 0 && stock <= 5) stockBajo++;
    });
  } else {
    const hp = getHoja("PRODUCTOS");
    if(hp && hp.getLastRow() >= 2){
      const enc = hp.getRange(1, 1, 1, hp.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const iC = enc.indexOf("CODIGO"), iP = enc.indexOf("PUBLICADO"), iS = enc.indexOf("STOCK");
      const n = hp.getLastRow() - 1;
      const col = (i) => i === -1 ? null : hp.getRange(2, i + 1, n, 1).getValues();
      const cC = col(iC), cP = col(iP), cS = col(iS);
      for(let i = 0; i < n; i++){
        if(cC && String(cC[i][0] || "").trim() === "") continue;
        const stock = cS ? (Number(cS[i][0]) || 0) : 0;
        if(cP && String(cP[i][0] || "").toLowerCase() === "si") productosActivos++;
        if(stock === 0) agotados++;
        if(stock > 0 && stock <= 5) stockBajo++;
      }
    }
  }

  // ---- MÁS VENDIDO: resumen + DETALLE_VENTAS nuevas + DETALLE_PEDIDOS
  const acumulado = {};
  const sumarFila = (c, nombre, cant) => {
    c = String(c || "").trim();
    if(!c) return;
    if(!acumulado[c]) acumulado[c] = { nombre: nombre, cantidad: 0 };
    acumulado[c].cantidad += Number(cant) || 0;
  };
  const hojaResumen = getHoja(HOJA_RESUMEN_METRICAS);
  if(hojaResumen && hojaResumen.getLastRow() >= 2){
    hojaResumen.getRange(2, 1, hojaResumen.getLastRow() - 1, 3).getValues().forEach(r => sumarFila(r[0], r[1], r[2]));
  }
  const hojaDetalleV = getHoja("DETALLE_VENTAS");
  if(hojaDetalleV){
    const ultima = hojaDetalleV.getLastRow();
    const desde = Math.max(2, (meta.detalleHastaFila || 1) + 1);
    if(ultima >= desde){
      hojaDetalleV.getRange(desde, 2, ultima - desde + 1, 3).getValues().forEach(r => sumarFila(r[0], r[1], r[2]));
    }
  }
  const hojaDetalleP = getHoja("DETALLE_PEDIDOS");
  if(hojaDetalleP && hojaDetalleP.getLastRow() >= 2){
    hojaDetalleP.getRange(2, 2, hojaDetalleP.getLastRow() - 1, 3).getValues().forEach(r => sumarFila(r[0], r[1], r[2]));
  }
  let top = { nombre: "—", cantidad: 0 };
  Object.keys(acumulado).forEach(c => { if(acumulado[c].cantidad > top.cantidad) top = acumulado[c]; });

  const ventasHoyTotal = ventasPOSHoy + ventasPedidosHoy;
  const ventasMesTotal = ventasPOSMes + ventasPedidosMes;
  const totalOperaciones = totalVentasPOS + totalPedidos;

  return {
    pedidosNuevos, totalPedidos,
    ventasPedidosHoy, ventasPedidosMes,
    ventasPOSHoy, ventasPOSMes, totalVentasPOS,
    ventasHoy: ventasHoyTotal,
    ventasMes: ventasMesTotal,
    clientesUnicos: clientes.size,
    ticketPromedio: totalOperaciones > 0 ? ventasMesTotal / totalOperaciones : 0,
    productosActivos, stockBajo, agotados,
    masVendidoNombre: top.nombre,
    masVendidoCantidad: top.cantidad
  };
}

/* =====================================================================
   RESUMEN DE VENTAS VIEJAS (limpieza de VENTAS_LOCAL y DETALLE_VENTAS)
   Las ventas de hace más de MESES_DETALLE_VENTAS meses se reemplazan
   por resúmenes que los reportes leen igual que ventas comunes:
   - VENTAS_LOCAL: una fila por día + forma de pago + origen (mostrador
     o "PEDIDO WEB"), VENTA_ID "RES-aaaammdd-n", VENDEDOR "RESUMEN" (o
     "PEDIDO WEB"), y VENTAS_RESUMIDAS = cuántas ventas representa (para
     cantidad de ventas y ticket promedio).
   - DETALLE_VENTAS: una línea por día + producto (cantidad total, precio
     promedio → el importe da exacto), colgada del resumen de ese día.
   Las ventas anuladas no entran (no cuentan en ningún reporte). Se
   controla que los totales den igual antes de escribir; si no, no se
   toca nada. El detalle original queda en un CSV en Drive.
   Corre sola una vez por mes (desde limpiarLedgerIdempotencia); a mano:
   compactarVentasAhora() desde el editor.
===================================================================== */
const PROP_ULTIMA_COMPACTACION = "VEEK_ULTIMA_COMPACTACION_VENTAS";
const PROP_RESULTADO_COMPACTACION = "VEEK_RESULTADO_COMPACTACION_VENTAS";

function _guardarResultadoCompactacion(r, origen){
  try{
    PropertiesService.getScriptProperties().setProperty(PROP_RESULTADO_COMPACTACION,
      JSON.stringify(Object.assign({ fecha: new Date().toISOString(), origen: origen }, r)));
  }catch(e){}
}

/** Admin (Configuración → Sistema): cuánto ocupa el historial y qué se resumiría ahora. */
function estadoVentasResumidas(){
  const meses = Number(MESES_DETALLE_VENTAS) || 0;
  const hoy = new Date();
  const limite = new Date(hoy.getFullYear(), hoy.getMonth() - (meses || 6), 1, 0, 0, 0);
  const zona = zonaNegocio() || "America/Argentina/Buenos_Aires";
  const ventas = getHoja("VENTAS_LOCAL"), detalle = getHoja("DETALLE_VENTAS");
  let filasVentas = 0, filasDetalle = 0, paraResumir = 0, resumenes = 0, masVieja = "";
  if(ventas && ventas.getLastRow() > 1){
    filasVentas = ventas.getLastRow() - 1;
    ventas.getRange(2, 1, filasVentas, 2).getValues().forEach(f => {
      const id = String(f[0] || "");
      if(/^RES-/i.test(id)){ resumenes++; return; }
      const d = f[1] instanceof Date ? f[1] : new Date(f[1]);
      if(isNaN(d)) return;
      if(d < limite) paraResumir++;
      const t = Utilities.formatDate(d, zona, "yyyy-MM-dd");
      if(!masVieja || t < masVieja) masVieja = t;
    });
  }
  if(detalle && detalle.getLastRow() > 1) filasDetalle = detalle.getLastRow() - 1;
  let ultimo = null;
  try{ ultimo = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_RESULTADO_COMPACTACION) || "null"); }catch(e){}
  return jsonResponse({
    success: true, meses: meses, desactivado: !meses,
    conservaDesde: Utilities.formatDate(limite, zona, "yyyy-MM-dd"),
    filasVentas: filasVentas, filasDetalle: filasDetalle,
    paraResumir: paraResumir, filasResumen: resumenes, ventaMasVieja: masVieja,
    ultimo: ultimo
  });
}

/** Admin: botón "Resumir ahora". */
function resumirVentasViejasPanel(){
  const meses = Number(MESES_DETALLE_VENTAS) || 0;
  if(!meses) return jsonResponse({ success:false, message:"El resumen de ventas viejas está desactivado (MESES_DETALLE_VENTAS = 0 en code.gs)." });
  try{
    const r = compactarVentasAntiguas(meses);
    _guardarResultadoCompactacion(r, "panel");
    return jsonResponse(Object.assign({ success: true }, r));
  }catch(e){
    return jsonResponse({ success:false, message: e.message });
  }
}

function compactarVentasProgramado(){
  if(!(Number(MESES_DETALLE_VENTAS) > 0)) return;
  const zona = zonaNegocio() || "America/Argentina/Buenos_Aires";
  const mes = Utilities.formatDate(new Date(), zona, "yyyy-MM");
  const props = PropertiesService.getScriptProperties();
  if(props.getProperty(PROP_ULTIMA_COMPACTACION) === mes) return; // ya corrió este mes
  try{
    const r = compactarVentasAntiguas(MESES_DETALLE_VENTAS);
    props.setProperty(PROP_ULTIMA_COMPACTACION, mes);
    _guardarResultadoCompactacion(r, "automático");
    console.log("compactarVentas: " + JSON.stringify(r));
  }catch(e){
    console.error("compactarVentasProgramado: " + e.message);
    try{ enviarInformeError({ tipo: "servidor", titulo: "No se pudo resumir las ventas viejas: " + e.message, descripcion: "Limpieza mensual de VENTAS_LOCAL / DETALLE_VENTAS.", automatico: true }); }catch(x){}
  }
}

/** Desde el editor de Apps Script: resume ya lo que tenga más de MESES_DETALLE_VENTAS meses. */
function compactarVentasAhora(){
  const r = compactarVentasAntiguas(Number(MESES_DETALLE_VENTAS) || 6);
  _guardarResultadoCompactacion(r, "editor");
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}

function compactarVentasAntiguas(meses){
  meses = Number(meses) || 6;
  const zona = zonaNegocio() || "America/Argentina/Buenos_Aires";
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) throw new Error("Planilla ocupada, probá de nuevo.");
  try{
    const ventas = getHoja("VENTAS_LOCAL");
    const detalle = getHoja("DETALLE_VENTAS");
    if(!ventas || ventas.getLastRow() < 2) return { resumidas: 0, mensaje: "Nada para resumir." };

    // Se conservan meses enteros: hoy 20/10 con 6 meses → se resume lo anterior al 1/4.
    const hoy = new Date();
    const limite = new Date(hoy.getFullYear(), hoy.getMonth() - meses, 1, 0, 0, 0);

    const datosV = ventas.getDataRange().getValues();
    let encV = datosV.shift().map(h => String(h || "").trim());
    ["CLIENTE_VENTA_ID", "ANULADA", "MOTIVO_ANULACION", "VENTAS_RESUMIDAS"].forEach(k => { if(encV.indexOf(k) === -1) encV.push(k); });
    const iId = encV.indexOf("VENTA_ID") === -1 ? 0 : encV.indexOf("VENTA_ID");
    const iFecha = 1, iTotal = 2, iFP = 3, iVend = 4, iObs = 5;
    const iAnul = encV.indexOf("ANULADA"), iPeso = encV.indexOf("VENTAS_RESUMIDAS");
    const ancho = encV.length;
    const normal = f => { const r = f.slice(0, ancho); while(r.length < ancho) r.push(""); return r; };

    const quedan = [], viejas = [];
    datosV.forEach(f => {
      const d = f[iFecha] instanceof Date ? f[iFecha] : new Date(f[iFecha]);
      (!isNaN(d) && d < limite ? viejas : quedan).push(normal(f));
    });
    if(!viejas.length) return { resumidas: 0, mensaje: "No hay ventas de hace más de " + meses + " meses." };

    // ---- Agrupar ventas viejas (sin anuladas) ----
    const grupos = {};        // dia|origen|fp -> { dia, pedido, fp, total, n }
    const diaDeVenta = {};    // VENTA_ID (mostrador) -> dia
    let totalViejo = 0, ventasViejas = 0, anuladasDescartadas = 0;
    viejas.forEach(f => {
      if(String(f[iAnul] || "").toUpperCase() === "SI"){ anuladasDescartadas++; return; }
      const d = f[iFecha] instanceof Date ? f[iFecha] : new Date(f[iFecha]);
      const dia = Utilities.formatDate(d, zona, "yyyy-MM-dd");
      const pedido = String(f[iVend] || "").toUpperCase() === "PEDIDO WEB";
      const fp = String(f[iFP] || "EFECTIVO").toUpperCase();
      const k = dia + "|" + (pedido ? "P" : "V") + "|" + fp;
      const g = grupos[k] || (grupos[k] = { dia: dia, pedido: pedido, fp: fp, total: 0, n: 0 });
      g.total += Number(f[iTotal]) || 0;
      g.n += _pesoVenta(f, iPeso);
      totalViejo += Number(f[iTotal]) || 0;
      ventasViejas += _pesoVenta(f, iPeso);
      if(!pedido) diaDeVenta[String(f[iId])] = dia;
    });

    // ---- Filas resumen (en orden de fecha) ----
    const idResumenDelDia = {}; // dia -> VENTA_ID del primer resumen de mostrador de ese día
    const contadorDia = {};
    const filasResumen = Object.keys(grupos).sort().map(k => {
      const g = grupos[k];
      contadorDia[g.dia] = (contadorDia[g.dia] || 0) + 1;
      const id = "RES-" + g.dia.replace(/-/g, "") + "-" + contadorDia[g.dia];
      if(!g.pedido && !idResumenDelDia[g.dia]) idResumenDelDia[g.dia] = id;
      const partes = g.dia.split("-").map(Number);
      const fila = new Array(ancho).fill("");
      fila[iId] = id;
      fila[iFecha] = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
      fila[iTotal] = Math.round(g.total * 100) / 100;
      fila[iFP] = g.fp;
      fila[iVend] = g.pedido ? "PEDIDO WEB" : "RESUMEN";
      fila[iObs] = "Resumen de " + g.n + (g.n === 1 ? " venta" : " ventas") + (g.pedido ? " (pedidos cobrados)" : "") + " · " + g.fp;
      fila[iPeso] = g.n;
      return fila;
    });

    // ---- Detalle: una línea por día + producto ----
    let datosD = [], encD = ["VENTA_ID", "CODIGO", "PRODUCTO", "CANTIDAD", "PRECIO", "SUBTOTAL"];
    if(detalle && detalle.getLastRow() >= 1){
      datosD = detalle.getDataRange().getValues();
      encD = datosD.shift().map(h => String(h || "").trim());
    }
    const anchoD = Math.max(encD.length, 6);
    const normalD = f => { const r = f.slice(0, anchoD); while(r.length < anchoD) r.push(""); return r; };
    const idsViejasTodas = new Set(viejas.map(f => String(f[iId])));
    const porProducto = {}; // dia|codigo -> { dia, codigo, nombre, cant, imp }
    const detQueda = [];
    let importeViejo = 0, lineasViejas = 0, diaSinResumen = 0;
    datosD.forEach(f => {
      const id = String(f[0]);
      if(!idsViejasTodas.has(id)){ detQueda.push(normalD(f)); return; }
      lineasViejas++;
      const dia = diaDeVenta[id];
      if(!dia) return; // línea de una venta anulada (o de una fila PEDIDO WEB): no cuenta
      const codigo = String(f[1] === null || f[1] === undefined ? "" : f[1]).trim();
      const cant = Number(f[3]) || 0;
      const imp = cant * (Number(f[4]) || 0);
      const k = dia + "|" + codigo;
      const g = porProducto[k] || (porProducto[k] = { dia: dia, codigo: codigo, nombre: "", cant: 0, imp: 0 });
      if(f[2]) g.nombre = String(f[2]);
      g.cant += cant; g.imp += imp;
      importeViejo += imp;
    });
    const lineasResumen = [];
    Object.keys(porProducto).sort().forEach(k => {
      const g = porProducto[k];
      let id = idResumenDelDia[g.dia];
      if(!id){ diaSinResumen++; return; }
      if(!g.cant && !g.imp) return;
      const fila = new Array(anchoD).fill("");
      fila[0] = id; fila[1] = g.codigo; fila[2] = g.nombre;
      // Precio promedio con decimales: cantidad × precio = importe exacto
      fila[3] = g.cant || 1;
      fila[4] = g.cant ? g.imp / g.cant : Math.round(g.imp * 100) / 100; // sin redondear: la planilla guarda el número completo
      fila[5] = Math.round(g.imp * 100) / 100;
      lineasResumen.push(fila);
    });

    // ---- Controles: los totales tienen que dar igual ----
    const totalResumen = filasResumen.reduce((t, f) => t + (Number(f[iTotal]) || 0), 0);
    const importeResumen = lineasResumen.reduce((t, f) => t + (Number(f[3]) || 0) * (Number(f[4]) || 0), 0);
    if(Math.abs(totalResumen - totalViejo) > 0.05 * Math.max(1, filasResumen.length)){
      throw new Error("Control de totales de ventas no coincide (" + totalViejo + " vs " + totalResumen + "): no se tocó nada.");
    }
    if(Math.abs(importeResumen - importeViejo) > Math.max(1, Math.abs(importeViejo) * 0.0001)){
      throw new Error("Control de importes del detalle no coincide (" + importeViejo + " vs " + importeResumen + "): no se tocó nada.");
    }
    if(diaSinResumen) throw new Error("Hay detalle sin venta de mostrador en " + diaSinResumen + " día(s): no se tocó nada.");

    // ---- Respaldo completo en Drive (CSV) antes de reemplazar ----
    const respaldo = _respaldarVentasEnDrive(encV, viejas, encD, datosD.filter(f => idsViejasTodas.has(String(f[0]))), limite, zona);

    // ---- Escribir: resúmenes primero (son lo más viejo), después lo que queda ----
    const nuevasV = filasResumen.concat(quedan);
    const nuevasD = lineasResumen.concat(detQueda);
    ventas.getRange(1, 1, 1, ancho).setValues([encV]);
    if(datosV.length) ventas.getRange(2, 1, datosV.length, ancho).clearContent();
    if(nuevasV.length) ventas.getRange(2, 1, nuevasV.length, ancho).setValues(nuevasV);
    if(datosV.length > nuevasV.length) ventas.deleteRows(2 + nuevasV.length, datosV.length - nuevasV.length);
    if(detalle){
      if(datosD.length) detalle.getRange(2, 1, datosD.length, anchoD).clearContent();
      if(nuevasD.length) detalle.getRange(2, 1, nuevasD.length, anchoD).setValues(nuevasD);
      if(datosD.length > nuevasD.length) detalle.deleteRows(2 + nuevasD.length, datosD.length - nuevasD.length);
    }

    invalidarCachePolling();
    try{ actualizarResumenMetricas(); }catch(e){ console.error("resumen métricas: " + e.message); }

    return {
      resumidas: ventasViejas, anuladasDescartadas: anuladasDescartadas,
      filasVentasAntes: datosV.length, filasVentasDespues: nuevasV.length,
      filasDetalleAntes: datosD.length, filasDetalleDespues: nuevasD.length,
      hasta: Utilities.formatDate(new Date(limite.getTime() - 1), zona, "yyyy-MM-dd"),
      respaldo: respaldo
    };
  } finally {
    lock.releaseLock();
  }
}

/** Guarda el detalle original que se va a resumir en dos CSV en Drive (carpeta "Respaldo de ventas"). */
function _respaldarVentasEnDrive(encV, filasV, encD, filasD, limite, zona){
  const csv = (enc, filas) => [enc].concat(filas).map(f => f.map(v => {
    let t = v instanceof Date ? Utilities.formatDate(v, zona, "yyyy-MM-dd HH:mm:ss") : String(v === null || v === undefined ? "" : v);
    return /[",\n;]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
  }).join(",")).join("\n");
  const hasta = Utilities.formatDate(new Date(limite.getTime() - 1), zona, "yyyy-MM-dd");
  let carpeta = null;
  try{
    const padres = DriveApp.getFileById(getSS().getId()).getParents();
    const base = padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
    const it = base.getFoldersByName("Respaldo de ventas");
    carpeta = it.hasNext() ? it.next() : base.createFolder("Respaldo de ventas");
  }catch(e){
    throw new Error("No se pudo crear el respaldo en Drive (" + e.message + "): no se tocó nada.");
  }
  const sello = Utilities.formatDate(new Date(), zona, "yyyyMMdd-HHmm");
  const a = carpeta.createFile("ventas-hasta-" + hasta + "-" + sello + ".csv", csv(encV, filasV), "text/csv");
  const b = carpeta.createFile("detalle-ventas-hasta-" + hasta + "-" + sello + ".csv", csv(encD, filasD), "text/csv");
  return { carpeta: carpeta.getUrl(), ventas: a.getName(), detalle: b.getName() };
}

/* =====================================================================
   ARCHIVADO DE HISTORIAL (opcional, manual)
   Mueve las ventas POS (VENTAS_LOCAL + DETALLE_VENTAS) de más de
   `meses` meses a hojas *_ARCHIVO en la misma planilla. Todas las
   lecturas del día a día quedan más chicas. OJO: los reportes por
   período ya no incluyen lo archivado — por eso solo corre solo si se
   configura ARCHIVAR_VENTAS_MESES (al principio del archivo) y se vuelve
   a ejecutar instalarTriggersVeekPOS(). Si no, a mano cuando haga falta.
===================================================================== */
/** Lo dispara el trigger mensual (solo si ARCHIVAR_VENTAS_MESES > 0). */
function archivarVentasProgramado(){
  const meses = Number(ARCHIVAR_VENTAS_MESES) || 0;
  if(meses <= 0) return;
  try{ console.log(archivarVentasAntiguas(meses)); }
  catch(e){ console.error("archivarVentasProgramado: " + e); }
}

function archivarVentasAntiguas(meses){
  meses = Number(meses) || 12;
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) throw new Error("Planilla ocupada, probá de nuevo.");
  try{
    const ss = getSS();
    const limite = new Date();
    limite.setMonth(limite.getMonth() - meses);

    const ventas = getHoja("VENTAS_LOCAL");
    const detalle = getHoja("DETALLE_VENTAS");
    if(!ventas || !detalle || ventas.getLastRow() < 2) return "Nada para archivar.";

    const datosV = ventas.getDataRange().getValues();
    const encV = datosV.shift();
    const viejas = [], quedan = [];
    datosV.forEach(f => { const d = new Date(f[1]); (!isNaN(d) && d < limite ? viejas : quedan).push(f); });
    if(viejas.length === 0) return "Nada para archivar.";

    const idsViejos = new Set(viejas.map(f => String(f[0])));
    const datosD = detalle.getDataRange().getValues();
    const encD = datosD.shift();
    const detViejo = [], detQueda = [];
    datosD.forEach(f => (idsViejos.has(String(f[0])) ? detViejo : detQueda).push(f));

    const asegurarArchivo = (nombre, enc) => {
      let h = ss.getSheetByName(nombre);
      if(!h){ h = ss.insertSheet(nombre); h.getRange(1, 1, 1, enc.length).setValues([enc]); }
      return h;
    };
    const archV = asegurarArchivo("VENTAS_LOCAL_ARCHIVO", encV);
    const archD = asegurarArchivo("DETALLE_VENTAS_ARCHIVO", encD);
    archV.getRange(archV.getLastRow() + 1, 1, viejas.length, encV.length).setValues(viejas);
    if(detViejo.length) archD.getRange(archD.getLastRow() + 1, 1, detViejo.length, encD.length).setValues(detViejo);

    // Reescritura en bloque de lo que queda (sin deleteRow fila por fila)
    ventas.getRange(2, 1, datosV.length, encV.length).clearContent();
    if(quedan.length) ventas.getRange(2, 1, quedan.length, encV.length).setValues(quedan);
    detalle.getRange(2, 1, datosD.length, encD.length).clearContent();
    if(detQueda.length) detalle.getRange(2, 1, detQueda.length, encD.length).setValues(detQueda);

    invalidarCachePolling();
    actualizarResumenMetricas();
    return "Archivadas " + viejas.length + " ventas y " + detViejo.length + " líneas de detalle.";
  } finally {
    lock.releaseLock();
  }
}

/**
 * Crea, en la planilla configurada (SPREADSHEET_ID), todas las hojas
 * que el sistema necesita para funcionar — con sus encabezados
 * correctos — si todavía no existen. Pensada para usarse una sola vez,
 * recién implementado el Code.gs de un negocio nuevo (paso "Probar
 * conexión" del Instalador): antes, varias de estas hojas (PRODUCTOS,
 * PEDIDOS, DETALLE_PEDIDOS, VENTAS_LOCAL, DETALLE_VENTAS,
 * MOVIMIENTOS_CAJA) no se creaban solas en ningún lado — había que
 * armarlas a mano, con los encabezados exactos, antes de poder usar
 * el sistema. Las demás (CONFIGURACION, CLIENTES, PAGOS_CREDITO,
 * DEUDAS_EXTRA, ADMIN_CREDENCIALES) ya se creaban solas la primera vez
 * que algo las necesitaba — acá también se fuerza su creación, para
 * no depender de qué acción se ejecute primero.
 *
 * No toca ninguna hoja que ya exista (nunca borra ni reescribe datos).
 */
function inicializarHojasNuevoNegocio(){
  const ss = getSS();
  const creadas = [];
  const yaExistian = [];

  const asegurarHojaConEncabezados = (nombre, encabezados) => {
    let sheet = getHoja(nombre);
    if(sheet){
      yaExistian.push(nombre);
      return;
    }
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(encabezados);
    sheet.setFrozenRows(1);
    creadas.push(nombre);
  };

  asegurarHojaConEncabezados("PRODUCTOS", [
    "CODIGO", "PRODUCTO", "CATEGORIA", "PRECIO", "STOCK",
    "PUBLICADO", "DESTACADO", "OFERTA", "DESCRIPCION", "IMAGEN",
    "CODIGO_CAJA", "UNIDADES_POR_CAJA", "PRECIO_CAJA", "ORDEN",
    "COLOR", "IMAGENES"
  ]);

  asegurarHojaConEncabezados("PEDIDOS", [
    "PEDIDO_ID", "FECHA", "NOMBRE", "EMPRESA", "DIRECCION", "TELEFONO", "DNI",
    "TOTAL", "ESTADO", "PDF_URL", "CODIGO_POSTAL", "LOCALIDAD", "PROVINCIA",
    "STOCK_DESCONTADO", "MONEDA", "TIPO_CAMBIO", "SUBTOTAL", "DESCUENTO", "DESCUENTO_ETIQUETA"
  ]);

  asegurarHojaConEncabezados("DETALLE_PEDIDOS", [
    "PEDIDO_ID", "CODIGO", "PRODUCTO", "CANTIDAD", "PRECIO", "SUBTOTAL", "COLOR"
  ]);

  asegurarHojaConEncabezados("VENTAS_LOCAL", [
    "VENTA_ID", "FECHA", "TOTAL", "FORMA_PAGO", "VENDEDOR", "OBSERVACIONES",
    "CLIENTE_VENTA_ID", "ANULADA", "MOTIVO_ANULACION"
  ]);

  asegurarHojaConEncabezados("DETALLE_VENTAS", [
    "VENTA_ID", "CODIGO", "PRODUCTO", "CANTIDAD", "PRECIO", "SUBTOTAL"
  ]);

  asegurarHojaConEncabezados("MOVIMIENTOS_CAJA", [
    "MOVIMIENTO_ID", "FECHA", "TIPO", "MONTO", "MOTIVO", "VENDEDOR",
    "FECHA_REGISTRO", "FORMA_PAGO"
  ]);

  // Antes no se creaba y el primer cierre de caja de un negocio nuevo fallaba
  asegurarHojaConEncabezados("CIERRES_CAJA", [
    "CIERRE_ID", "FECHA", "EFECTIVO_ESPERADO", "EFECTIVO_CONTADO", "EFECTIVO_DIFERENCIA",
    "TRANSFERENCIA_ESPERADO", "TRANSFERENCIA_CONTADO", "TRANSFERENCIA_DIFERENCIA",
    "TARJETA_ESPERADO", "TARJETA_CONTADO", "TARJETA_DIFERENCIA",
    "TOTAL_ESPERADO", "TOTAL_CONTADO", "TOTAL_DIFERENCIA", "VENDEDOR", "OBSERVACIONES", "FECHA_CIERRE"
  ]);

  // INGRESOS_PRODUCTOS la crea getOrCrearHojaIngresosProductos() (más abajo)
  // con las columnas actuales (BOLETA_ID, MONEDA, PRECIO_INGRESO...).

  // Estas ya se auto-crean solas al primer uso (getOrCrearHoja*), pero
  // se fuerza acá también para que el instalador las deje todas listas
  // de una — no depende de qué pantalla se abra primero en el panel.
  getOrCrearHojaConfiguracion();
  getOrCrearHojaCredenciales();
  getOrCrearHojaClientes();
  getOrCrearHojaPagosCredito();
  getOrCrearHojaDeudasExtra();
  getOrCrearHojaIngresosProductos();
  getOrCrearHojaUsuarios();
  getOrCrearHojaCupones();
  getOrCrearHojaCuponesUsos();

  return jsonResponse({
    success: true,
    creadas: creadas,
    yaExistian: yaExistian,
    message: creadas.length > 0
      ? `Se crearon ${creadas.length} hojas nuevas: ${creadas.join(", ")}.`
      : "Todas las hojas ya existían — no hizo falta crear ninguna."
  });
}

function rowsToObjects(data) {

  const headers = data.shift();

  return data.map(row => {

    let obj = {};

    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    return obj;

  });
}

function getProductos() {

  const cacheado = cacheLeerTexto("catalogo_publico");
  if(cacheado){
    return ContentService.createTextOutput(cacheado).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getSheet();
  asegurarColumnaOrden(sheet);

  const data = sheet
    .getDataRange()
    .getValues();

  const productos = rowsToObjects(data);

  const publicados = productos
    .filter(p => String(p.PUBLICADO).toUpperCase() === "SI")
    .sort(ordenarPorColumnaOrden);

  const texto = JSON.stringify({
    success: true,
    total: publicados.length,
    productos: publicados
  });

  cacheGuardarTexto("catalogo_publico", texto, CACHE_TTL_SEGUNDOS);

  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Calienta la caché del catálogo y de la config pública. Está pensado
 * para correr solo cada 10 minutos (ver instalarTriggerCalentarCatalogo),
 * así el cliente casi nunca le toca pagar la lectura completa de la
 * planilla (que es lo que tarda decenas de segundos en frío).
 */
function calentarCacheCatalogo(){
  try{ getProductos(); }catch(e){ console.error("calentar catálogo: " + e.message); }
  try{ conCachePolling("config_negocio", 300, getConfiguracionNegocio); }catch(e){ console.error("calentar config: " + e.message); }
  try{ getProductosAdmin(); }catch(e){ console.error("calentar catálogo admin: " + e.message); }
  try{ getOrCrearHojaClientes(); getHoja("PEDIDOS"); getHoja("DETALLE_PEDIDOS"); }catch(e){ console.error("calentar hojas: " + e.message); }
}

/** Ejecutar UNA sola vez a mano desde el editor de Apps Script (▶️). */
function instalarTriggerCalentarCatalogo(){
  ScriptApp.getProjectTriggers().forEach(t => {
    if(t.getHandlerFunction() === "calentarCacheCatalogo") ScriptApp.deleteTrigger(t);
  });
  // Reemplazado por instalarTriggersVeekPOS (cada 30 min, solo en horario comercial)
  instalarTriggersVeekPOS();
}

/**
 * Comparador para Array.sort: ordena por la columna ORDEN (numérica,
 * ascendente). Los productos sin ORDEN cargado (o con un valor no
 * numérico) van al final, en vez de mezclarse al principio con un 0.
 */
function ordenarPorColumnaOrden(a, b){
  const ordenA = a.ORDEN === "" || a.ORDEN === undefined || isNaN(Number(a.ORDEN)) ? Infinity : Number(a.ORDEN);
  const ordenB = b.ORDEN === "" || b.ORDEN === undefined || isNaN(Number(b.ORDEN)) ? Infinity : Number(b.ORDEN);
  return ordenA - ordenB;
}

/**
 * Si la hoja PRODUCTOS todavía no tiene la columna ORDEN (usada para
 * elegir en qué posición aparece cada producto en el catálogo/POS), la
 * agrega al final — y a los productos que ya existían les asigna un
 * ORDEN según la fila en la que ya estaban, así agregar la columna no
 * les cambia el orden que ya tenían de entrada.
 */
function asegurarColumnaOrden(sheet){
  if(!sheet) return;

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if(headerRow.indexOf("ORDEN") !== -1) return;

  const colNueva = sheet.getLastColumn() + 1;
  sheet.getRange(1, colNueva).setValue("ORDEN");

  const totalFilas = sheet.getLastRow() - 1;
  if(totalFilas > 0){
    const ordenInicial = [];
    for(let i = 1; i <= totalFilas; i++) ordenInicial.push([i]);
    sheet.getRange(2, colNueva, totalFilas, 1).setValues(ordenInicial);
  }
}

/** Devuelve el próximo número de ORDEN libre (el mayor ORDEN cargado + 1), para que un producto nuevo se agregue al final del catálogo en vez de al principio. */
function obtenerSiguienteOrden(sheet, headerRow){
  const colOrden = headerRow.indexOf("ORDEN");
  if(colOrden === -1) return 1;

  const totalFilas = sheet.getLastRow() - 1;
  if(totalFilas <= 0) return 1;

  const valores = sheet.getRange(2, colOrden + 1, totalFilas, 1).getValues();
  const maxOrden = valores.reduce((max, fila) => {
    const n = Number(fila[0]);
    return isNaN(n) ? max : Math.max(max, n);
  }, 0);

  return maxOrden + 1;
}

/**
 * Guarda un nuevo orden de productos para el catálogo, a partir de una
 * lista de códigos en el orden deseado (mandada por el panel admin
 * después de arrastrar y soltar filas en la tabla de productos).
 *
 * Los productos que NO vienen en la lista (por ejemplo, si el admin
 * tenía un filtro/búsqueda aplicado al reordenar) conservan su
 * posición relativa, pero se corren al final — detrás de todos los
 * que sí se reordenaron — para que no queden mezclados.
 */
function reordenarProductos(data){
  const codigos = data.codigos;

  if(!Array.isArray(codigos) || codigos.length === 0){
    return jsonResponse({
      success: false,
      message: "No se recibió la lista de productos a ordenar."
    });
  }

  const sheet = getSheet();
  if(!sheet){
    return jsonResponse({
      success: false,
      message: "No se encontró la hoja 'PRODUCTOS'."
    });
  }

  asegurarColumnaOrden(sheet);

  const valores = sheet.getDataRange().getValues();
  const headerRow = valores[0];
  const colOrden = headerRow.indexOf("ORDEN");
  const colCodigo = headerRow.indexOf("CODIGO");
  const totalFilas = valores.length - 1;

  if(totalFilas <= 0){
    return jsonResponse({ success: true });
  }

  const posicionDeseada = {};
  codigos.forEach((codigo, idx) => {
    posicionDeseada[String(codigo).trim()] = idx + 1;
  });

  let siguienteParaElResto = codigos.length;
  const nuevoOrdenPorFila = [];

  for(let i = 1; i < valores.length; i++){
    const codigo = String(valores[i][colCodigo]).trim();
    if(posicionDeseada.hasOwnProperty(codigo)){
      nuevoOrdenPorFila.push([posicionDeseada[codigo]]);
    } else {
      siguienteParaElResto++;
      nuevoOrdenPorFila.push([siguienteParaElResto]);
    }
  }

  sheet.getRange(2, colOrden + 1, totalFilas, 1).setValues(nuevoOrdenPorFila);

  invalidarCacheProductos();

  return jsonResponse({ success: true });
}

/**
 * Devuelve TODOS los productos (publicados o no), para la tabla de
 * administración en el panel — ahí el admin necesita ver y editar
 * también los productos que no están publicados en el catálogo/POS.
 */
function getProductosAdmin() {

  const cacheado = cacheLeerTexto("catalogo_admin");
  if(cacheado){
    return ContentService.createTextOutput(cacheado).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getSheet();
  asegurarColumnaOrden(sheet);

  const data = sheet
    .getDataRange()
    .getValues();

  const productos = rowsToObjects(data)
    // El rango de la hoja a veces incluye filas vacías de más (por
    // formato aplicado más allá de los datos reales, por ejemplo) —
    // se descartan las filas sin código, porque no son productos
    // reales y rompían el guardado de la caché local del panel
    // (CODIGO es clave única ahí, y dos filas vacías chocan entre sí).
    .filter(p => String(p.CODIGO || "").trim() !== "")
    .sort(ordenarPorColumnaOrden);

  const texto = JSON.stringify({
    success: true,
    total: productos.length,
    productos: productos
  });

  // TTL más corto que el catálogo público: el panel admin necesita
  // ver los cambios propios (stock, precio, etc.) más rápido.
  cacheGuardarTexto("catalogo_admin", texto, CACHE_TTL_ADMIN_SEGUNDOS);

  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

function getProducto(codigo) {

  const data = getSheet()
    .getDataRange()
    .getValues();

  const productos = rowsToObjects(data);

  const producto = productos.find(
    p => String(p.CODIGO) === String(codigo)
  );

  return jsonResponse({
    success: !!producto,
    producto
  });
}

function getCategorias() {

  const cacheado = cacheLeerTexto("categorias");
  if(cacheado){
    return ContentService.createTextOutput(cacheado).setMimeType(ContentService.MimeType.JSON);
  }

  const data = getSheet()
    .getDataRange()
    .getValues();

  data.shift();

  const categorias = [
    ...new Set(data.map(row => row[2]))
  ];

  const texto = JSON.stringify({ success: true, categorias });
  cacheGuardarTexto("categorias", texto, CACHE_TTL_SEGUNDOS);

  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

function getDestacados() {

  const cacheado = cacheLeerTexto("destacados");
  if(cacheado){
    return ContentService.createTextOutput(cacheado).setMimeType(ContentService.MimeType.JSON);
  }

  const data = getSheet()
    .getDataRange()
    .getValues();

  const productos = rowsToObjects(data);

  const destacados = productos.filter(
    p => String(p.DESTACADO).toUpperCase() === "SI"
  );

  const texto = JSON.stringify({ success: true, productos: destacados });
  cacheGuardarTexto("destacados", texto, CACHE_TTL_SEGUNDOS);

  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

function getOfertas() {

  const cacheado = cacheLeerTexto("ofertas");
  if(cacheado){
    return ContentService.createTextOutput(cacheado).setMimeType(ContentService.MimeType.JSON);
  }

  const data = getSheet()
    .getDataRange()
    .getValues();

  const productos = rowsToObjects(data);

  const ofertas = productos.filter(
    p => String(p.OFERTA).toUpperCase() === "SI"
  );

  const texto = JSON.stringify({ success: true, productos: ofertas });
  cacheGuardarTexto("ofertas", texto, CACHE_TTL_SEGUNDOS);

  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
   PRODUCTOS — ALTA (formulario "Nuevo Producto" en el panel)
   Lee los encabezados reales de la hoja PRODUCTOS y arma la fila
   por nombre de columna, así no rompe nada si la hoja tiene columnas
   en otro orden o adicionales (ej: DESCRIPCION, IMAGEN, MARCA, etc.).
   Cualquier encabezado no provisto por el formulario queda vacío.
===================================================================== */

/**
 * Si la hoja PRODUCTOS todavía no tiene las columnas para manejar el
 * precio por caja (CODIGO_CAJA / UNIDADES_POR_CAJA / PRECIO_CAJA),
 * las agrega al final — sin tocar ninguna columna ni fila existente.
 * Se llama solo al GUARDAR un producto con estos datos, no en cada
 * lectura del catálogo (para no gastar de más revisando esto en cada
 * carga del catálogo web).
 */
function asegurarColumnasProductoCaja(sheet){
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnasNecesarias = ["CODIGO_CAJA", "UNIDADES_POR_CAJA", "PRECIO_CAJA"];
  const faltantes = columnasNecesarias.filter(c => headerRow.indexOf(c) === -1);
  if(faltantes.length > 0){
    const inicio = sheet.getLastColumn() + 1;
    sheet.getRange(1, inicio, 1, faltantes.length).setValues([faltantes]);
  }
}

/**
 * Si la hoja PRODUCTOS todavía no tiene la columna ALIAS (usada para
 * guardar sinónimos/palabras clave que aceleran la búsqueda en el POS
 * sin tocar el nombre real del producto), la agrega al final — sin
 * tocar ninguna columna ni fila existente. Mismo patrón que
 * asegurarColumnasProductoCaja.
 */
/** Columna ALICUOTA_IVA (opcional, facturación ARCA de Responsable Inscripto): vacía = la alícuota por defecto */
function asegurarColumnaAlicuotaIva(sheet){
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if(headerRow.indexOf("ALICUOTA_IVA") !== -1) return;
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue("ALICUOTA_IVA");
}

function _alicuotaIvaValida(v){
  const n = String(v === undefined || v === null ? "" : v).replace(",", ".").replace("%", "").trim();
  return ["0", "2.5", "5", "10.5", "21", "27"].indexOf(n) !== -1 ? n : "";
}

function asegurarColumnaAlias(sheet){
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if(headerRow.indexOf("ALIAS") !== -1) return;

  const colNueva = sheet.getLastColumn() + 1;
  sheet.getRange(1, colNueva).setValue("ALIAS");
}

/**
 * Si la hoja PRODUCTOS todavía no tiene las columnas COLOR e IMAGENES
 * (usadas por la galería de colores del catálogo web: COLOR es el
 * color de la foto principal, IMAGENES son fotos adicionales con su
 * propio color, formato "url1|color1, url2|color2, ..."), las agrega
 * al final — sin tocar ninguna columna ni fila existente. Mismo
 * patrón que asegurarColumnasProductoCaja.
 */
function asegurarColumnasGaleriaProducto(sheet){
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnasNecesarias = ["COLOR", "IMAGENES"];
  const faltantes = columnasNecesarias.filter(c => headerRow.indexOf(c) === -1);
  if(faltantes.length > 0){
    const inicio = sheet.getLastColumn() + 1;
    sheet.getRange(1, inicio, 1, faltantes.length).setValues([faltantes]);
  }
}

/* =====================================================================
   CÓDIGOS ÚNICOS DE PRODUCTO
   Un código (de producto o de caja) identifica UNA sola cosa: el escáner
   del POS busca en los dos. Se compara sin espacios y sin distinguir
   mayúsculas ("abc " y "ABC" son el mismo código).
===================================================================== */
function normalizarCodigo(c){
  return String(c === undefined || c === null ? "" : c).trim().toUpperCase();
}

/**
 * Revisa si `codigo` ya está usado en `filas` (con encabezado) como
 * CODIGO o como CODIGO_CAJA de otro producto. `filaPropia` (índice en
 * `filas`) se ignora — es el producto que se está editando.
 * Devuelve el mensaje de error, o null si está libre.
 */
function conflictoCodigoProducto(filas, codigo, filaPropia, esCodigoCaja){
  const n = normalizarCodigo(codigo);
  if(!n || !filas || !filas.length) return null;
  const enc = filas[0].map(h => String(h || "").trim());
  const iCod = enc.indexOf("CODIGO") === -1 ? 0 : enc.indexOf("CODIGO");
  const iCaja = enc.indexOf("CODIGO_CAJA");
  const iNom = enc.indexOf("PRODUCTO");
  const queEs = esCodigoCaja ? "El código de caja" : "El código";
  for(let i = 1; i < filas.length; i++){
    if(i === filaPropia) continue;
    const nombre = iNom !== -1 ? String(filas[i][iNom] || "").trim() : "";
    if(normalizarCodigo(filas[i][iCod]) === n){
      return queEs + " '" + String(codigo).trim() + "' ya pertenece al producto " + (nombre ? "\"" + nombre + "\"" : "existente") + ".";
    }
    if(iCaja !== -1 && normalizarCodigo(filas[i][iCaja]) === n){
      return queEs + " '" + String(codigo).trim() + "' ya es el código de caja del producto " + (nombre ? "\"" + nombre + "\"" : "existente") + ".";
    }
  }
  return null;
}

/**
 * Ejecutar a mano desde el editor: lista los códigos repetidos que YA
 * existan en la hoja PRODUCTOS (de antes de este control), para
 * corregirlos. No modifica nada.
 */
function revisarCodigosDuplicados(){
  const filas = getSheet().getDataRange().getValues();
  const enc = filas[0].map(h => String(h || "").trim());
  const iCod = enc.indexOf("CODIGO") === -1 ? 0 : enc.indexOf("CODIGO");
  const iCaja = enc.indexOf("CODIGO_CAJA");
  const iNom = enc.indexOf("PRODUCTO");
  const usos = {};
  for(let i = 1; i < filas.length; i++){
    [[iCod, "código"], [iCaja, "código de caja"]].forEach(([idx, tipo]) => {
      if(idx === -1) return;
      const n = normalizarCodigo(filas[i][idx]);
      if(!n) return;
      (usos[n] = usos[n] || []).push("fila " + (i + 1) + " (" + tipo + ", " + String(filas[i][iNom] || "") + ")");
    });
  }
  const repetidos = Object.keys(usos).filter(k => usos[k].length > 1).map(k => k + ": " + usos[k].join(" / "));
  const texto = repetidos.length ? "Códigos repetidos:\n" + repetidos.join("\n") : "No hay códigos repetidos.";
  Logger.log(texto);
  return texto;
}

function guardarProducto(e){

  const params = (e && e.parameter) || {};

  if(String(params.rol || "").trim().toLowerCase() === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para crear o modificar productos."
    });
  }

  // Evita que dos guardados simultáneos (dos cajas, o un doble clic)
  // lean la hoja al mismo tiempo antes de que ninguno haya hecho el
  // appendRow todavía, y terminen creando dos filas con el mismo
  // CODIGO. Mismo patrón que descontarStockPorVenta / sumarStock.
  //
  // Si ya estamos adentro de conIdempotencia (cola offline, siempre
  // manda idOperacionSync), esa misma ejecución ya tiene el script
  // lock tomado — no volver a pedirlo acá (ver _dentroDeLockIdempotencia).
  let lock = null;
  let obtuvoLock = true;
  if (!_dentroDeLockIdempotencia) {
    lock = LockService.getScriptLock();
    obtuvoLock = lock.tryLock(10000);
  }

  if(!obtuvoLock){
    console.error("No se pudo obtener el lock para guardar producto. CODIGO: " + (params.CODIGO || params.codigo || ""));
    return jsonResponse({
      success:false,
      message:"El sistema está ocupado guardando otro producto, probá de nuevo en unos segundos."
    });
  }

  try{

    const sheet = getSheet();

    if(!sheet){
      return jsonResponse({
        success:false,
        message:"No se encontró la hoja 'PRODUCTOS'."
      });
    }

    asegurarColumnasProductoCaja(sheet);
    asegurarColumnaOrden(sheet);
    asegurarColumnaAlias(sheet);
    asegurarColumnasGaleriaProducto(sheet);
    if(params.ALICUOTA_IVA !== undefined && params.ALICUOTA_IVA !== "") asegurarColumnaAlicuotaIva(sheet);

    const headerRow =
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Validaciones mínimas
    const nombre = (params.PRODUCTO || params.producto || "").toString().trim();
    if(!nombre){
      return jsonResponse({
        success:false,
        message:"El nombre del producto es obligatorio."
      });
    }

    // CODIGO: si no lo manda el formulario, se genera uno
    let codigo = (params.CODIGO || params.codigo || "").toString().trim();
    if(!codigo){
      codigo = "PRD-" + new Date().getTime();
    }

    // Códigos únicos: ni el código ni el de caja pueden estar usados por
    // otro producto (ni como código ni como código de caja).
    {
      const existentes = sheet.getDataRange().getValues();
      const codigoCajaNuevo = (params.CODIGO_CAJA || "").toString().trim();
      const conflicto = conflictoCodigoProducto(existentes, codigo, -1, false) ||
        (codigoCajaNuevo ? conflictoCodigoProducto(existentes, codigoCajaNuevo, -1, true) : null);
      if(conflicto) return jsonResponse({ success:false, codigoDuplicado:true, message: conflicto });
      if(codigoCajaNuevo && normalizarCodigo(codigoCajaNuevo) === normalizarCodigo(codigo)){
        return jsonResponse({ success:false, codigoDuplicado:true, message:"El código de caja tiene que ser distinto del código del producto." });
      }
    }

    // Valores recibidos del formulario, normalizados por nombre de columna.
    // Las claves de este objeto deben coincidir EXACTAMENTE con los
    // encabezados de la hoja PRODUCTOS.
    const valoresPorColumna = {
      CODIGO:     codigo,
      PRODUCTO:   nombre,
      CATEGORIA:  (params.CATEGORIA || "").toString().trim(),
      PRECIO:     Number(params.PRECIO || 0),
      STOCK:      Number(params.STOCK || 0),
      PUBLICADO:  (params.PUBLICADO || "SI").toString().toUpperCase(),
      DESTACADO:  (params.DESTACADO || "NO").toString().toUpperCase(),
      OFERTA:     (params.OFERTA || "NO").toString().toUpperCase(),
      DESCRIPCION:(params.DESCRIPCION || "").toString(),
      IMAGEN:     (params.IMAGEN || "").toString().trim(),
      // Galería de colores: COLOR es el color de la foto principal
      // (IMAGEN); IMAGENES son fotos adicionales con su propio color,
      // formato "url1|color1, url2|color2, ...".
      COLOR:      (params.COLOR || "").toString().trim(),
      IMAGENES:   (params.IMAGENES || "").toString().trim(),
      // Alias/sinónimos para búsqueda rápida en el POS (no se muestran
      // al cliente, no afectan el nombre real del producto).
      ALIAS:      (params.ALIAS || "").toString().trim(),
      // Venta por caja (opcional) — mismo producto, otro código de
      // barras para el bulto/caja cerrada, a un precio total distinto
      // (normalmente más conveniente que N veces el precio unitario).
      CODIGO_CAJA:       (params.CODIGO_CAJA || "").toString().trim(),
      UNIDADES_POR_CAJA: Number(params.UNIDADES_POR_CAJA || 0),
      PRECIO_CAJA:       Number(params.PRECIO_CAJA || 0),
      ALICUOTA_IVA:      _alicuotaIvaValida(params.ALICUOTA_IVA),
      // Producto nuevo -> va al final del orden actual del catálogo,
      // no a un ORDEN vacío/0 que lo mandaría al principio.
      ORDEN: obtenerSiguienteOrden(sheet, headerRow)
    };

    // Armar la fila en el MISMO ORDEN que los encabezados reales de la hoja.
    const fila = headerRow.map(header => {
      const key = String(header || "").trim();
      return valoresPorColumna.hasOwnProperty(key)
        ? valoresPorColumna[key]
        : ""; // columna desconocida -> se deja vacía
    });

    sheet.appendRow(fila);

    invalidarCacheProductos();

    return jsonResponse({
      success:true,
      codigo: codigo
    });

  } finally {
    if (lock) lock.releaseLock();
  }

}

/**
 * Actualiza un producto existente (identificado por su CODIGO original,
 * que puede venir distinto si el admin también cambia el código).
 */
function actualizarProducto(e){

  const params = (e && e.parameter) || {};

  if(String(params.rol || "").trim().toLowerCase() === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para modificar precio o stock de productos."
    });
  }

  // Mismo motivo que en guardarProducto: sin lock, dos actualizaciones
  // simultáneas (ej. dos cambios de código casi al mismo tiempo) podrían
  // leer la hoja antes de que la otra escriba y terminar pisándose o
  // dejando dos filas con el mismo CODIGO nuevo.
  //
  // Si ya estamos adentro de conIdempotencia (cola offline, siempre
  // manda idOperacionSync), esa misma ejecución ya tiene el script
  // lock tomado — no volver a pedirlo acá (ver _dentroDeLockIdempotencia).
  let lock = null;
  let obtuvoLock = true;
  if (!_dentroDeLockIdempotencia) {
    lock = LockService.getScriptLock();
    obtuvoLock = lock.tryLock(10000);
  }

  if(!obtuvoLock){
    console.error("No se pudo obtener el lock para actualizar producto. codigoOriginal: " + (params.codigoOriginal || ""));
    return jsonResponse({
      success:false,
      message:"El sistema está ocupado guardando otro producto, probá de nuevo en unos segundos."
    });
  }

  try{

    const sheet = getSheet();

    if(!sheet){
      return jsonResponse({
        success:false,
        message:"No se encontró la hoja 'PRODUCTOS'."
      });
    }

    asegurarColumnasProductoCaja(sheet);
    asegurarColumnaOrden(sheet);
    asegurarColumnaAlias(sheet);
    asegurarColumnasGaleriaProducto(sheet);
    if(params.hasOwnProperty("ALICUOTA_IVA") && params.ALICUOTA_IVA !== "") asegurarColumnaAlicuotaIva(sheet);

    const codigoOriginal = (params.codigoOriginal || "").toString().trim();

    if(!codigoOriginal){
      return jsonResponse({
        success:false,
        message:"Falta el código original del producto a actualizar."
      });
    }

    const nombre = (params.PRODUCTO || params.producto || "").toString().trim();
    if(!nombre){
      return jsonResponse({
        success:false,
        message:"El nombre del producto es obligatorio."
      });
    }

    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    let filaIndex = -1;
    for(let i = 1; i < data.length; i++){
      if(String(data[i][0]).trim() === codigoOriginal){
        filaIndex = i;
        break;
      }
    }

    // No está con ese código: puede que en otra caja (o en el panel web)
    // le hayan cambiado el código o lo hayan borrado mientras esta
    // edición esperaba en la cola offline. Ver ubicarProductoPerdido.
    let codigoRealEncontrado = "";
    if(filaIndex === -1){
      const ub = ubicarProductoPerdido(data, codigoOriginal, nombre);
      if(ub.fila !== -1){
        filaIndex = ub.fila;
        codigoRealEncontrado = String(data[filaIndex][0]).trim();
        console.log("actualizarProducto: '" + codigoOriginal + "' ya no existe; se aplica a '" + codigoRealEncontrado + "' (" + ub.motivo + ")");
      }else{
        return jsonResponse(respuestaProductoNoEncontrado(codigoOriginal));
      }
    }

    let nuevoCodigo = (params.CODIGO || params.codigo || codigoOriginal).toString().trim();
    if(!nuevoCodigo) nuevoCodigo = codigoOriginal;
    // Si se encontró con otro código y esta edición no cambiaba el código,
    // se conserva el código actual (el que le pusieron en la otra caja).
    if(codigoRealEncontrado && normalizarCodigo(nuevoCodigo) === normalizarCodigo(codigoOriginal)){
      nuevoCodigo = codigoRealEncontrado;
    }

    // Códigos únicos: el código (si cambió) y el de caja (si cambió) no
    // pueden estar usados por OTRO producto, ni como código ni como caja.
    {
      const idxCajaActual = headerRow.map(h => String(h || "").trim()).indexOf("CODIGO_CAJA");
      const cajaActual = idxCajaActual !== -1 ? String(data[filaIndex][idxCajaActual] || "").trim() : "";
      const cajaNueva = params.CODIGO_CAJA !== undefined ? String(params.CODIGO_CAJA || "").trim() : cajaActual;
      let conflicto = null;
      if(normalizarCodigo(nuevoCodigo) !== normalizarCodigo(codigoOriginal)){
        conflicto = conflictoCodigoProducto(data, nuevoCodigo, filaIndex, false);
      }
      if(!conflicto && cajaNueva && normalizarCodigo(cajaNueva) !== normalizarCodigo(cajaActual)){
        conflicto = conflictoCodigoProducto(data, cajaNueva, filaIndex, true);
      }
      if(!conflicto && cajaNueva && normalizarCodigo(cajaNueva) === normalizarCodigo(nuevoCodigo)){
        conflicto = "El código de caja tiene que ser distinto del código del producto.";
      }
      if(conflicto) return jsonResponse({ success:false, codigoDuplicado:true, message: conflicto });
    }

    const valoresPorColumna = {
      CODIGO:     nuevoCodigo,
      PRODUCTO:   nombre,
      CATEGORIA:  (params.CATEGORIA || "").toString().trim(),
      PRECIO:     Number(params.PRECIO || 0),
      PUBLICADO:  (params.PUBLICADO || "SI").toString().toUpperCase(),
      DESTACADO:  (params.DESTACADO || "NO").toString().toUpperCase(),
      OFERTA:     (params.OFERTA || "NO").toString().toUpperCase(),
      DESCRIPCION:(params.DESCRIPCION || "").toString(),
      IMAGEN:     (params.IMAGEN || "").toString().trim(),
      COLOR:      (params.COLOR || "").toString().trim(),
      IMAGENES:   (params.IMAGENES || "").toString().trim(),
      CODIGO_CAJA:       (params.CODIGO_CAJA || "").toString().trim(),
      UNIDADES_POR_CAJA: Number(params.UNIDADES_POR_CAJA || 0),
      PRECIO_CAJA:       Number(params.PRECIO_CAJA || 0)
    };

    // STOCK: solo se pisa si el request lo manda explícito (igual
    // patrón que ALIAS, ver abajo). Si no viene, la fila conserva el
    // stock que ya tenía en la hoja — esto es a propósito: el panel
    // de escritorio (app offline) manda ediciones de producto SIN
    // STOCK cuando el cambio de stock en sí se sube aparte como un
    // delta vía sumarStockProducto (para no pisar con un número
    // desactualizado el stock real si otra caja vendió mientras tanto).
    // Antes esto pisaba con 0 cualquier edición que no incluyera STOCK.
    if(params.hasOwnProperty("STOCK")){
      valoresPorColumna.STOCK = Number(params.STOCK || 0);
    }

    // Alias/sinónimos para búsqueda rápida en el POS (no se muestran al
    // cliente, no afectan el nombre real del producto). Solo se
    // pisa si el request lo manda — si no (ej. un formulario viejo
    // que todavía no conoce este campo), la fila conserva el ALIAS
    // que ya tenía, igual que cualquier otra columna no reconocida.
    if(params.hasOwnProperty("ALIAS")){
      valoresPorColumna.ALIAS = (params.ALIAS || "").toString().trim();
    }

    // Alícuota de IVA (facturación ARCA): mismo criterio, solo si viene
    if(params.hasOwnProperty("ALICUOTA_IVA")){
      valoresPorColumna.ALICUOTA_IVA = _alicuotaIvaValida(params.ALICUOTA_IVA);
    }

    const filaActual = data[filaIndex];

    const filaNueva = headerRow.map((header, idx) => {
      const key = String(header || "").trim();
      return valoresPorColumna.hasOwnProperty(key)
        ? valoresPorColumna[key]
        : filaActual[idx]; // columna desconocida -> conserva el valor que ya tenía
    });

    sheet.getRange(filaIndex + 1, 1, 1, filaNueva.length).setValues([filaNueva]);

    invalidarCacheProductos();

    // Se recuerda el cambio de código: las operaciones de la cola que
    // todavía usan el código viejo (ej. un ajuste de stock) lo encuentran.
    if(normalizarCodigo(nuevoCodigo) !== normalizarCodigo(codigoOriginal)){
      registrarRenombreCodigo(codigoOriginal, nuevoCodigo);
    }

    return jsonResponse({
      success:true,
      codigo: nuevoCodigo,
      codigoAnterior: codigoRealEncontrado ? codigoOriginal : undefined
    });

  } finally {
    if (lock) lock.releaseLock();
  }

}

/**
 * Suma (o resta, si "cantidad" viene negativa) al stock ya cargado de
 * un producto. A diferencia de actualizarProducto (que reemplaza el
 * número de stock), esto AJUSTA lo que ya había — y usa LockService
 * por la misma razón que descontarStockPorVenta: si dos personas
 * cargan/corrigen mercadería del mismo producto casi al mismo tiempo,
 * no se puede perder ninguno de los dos ajustes.
 *
 * Antes esta función solo aceptaba positivos (pensada originalmente
 * para "Ingreso de Productos", que en el panel ya valida que sea
 * positivo del lado del cliente — esa pantalla sigue igual). Ahora
 * también la usa la edición de producto offline para subir el DELTA
 * de un cambio de stock, que puede ser negativo (ej. corregir un
 * error de carga) — por eso se relaja acá a "cualquier valor distinto
 * de cero", dejando la restricción de "solo positivo" donde
 * corresponde: en la UI de Ingreso de Productos, no acá.
 *
 * Parámetros esperados: codigo, cantidad (puede ser negativa)
 */
function sumarStockProducto(e){

  const params = (e && e.parameter) || {};

  if(String(params.rol || "").trim().toLowerCase() === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para modificar el stock de productos."
    });
  }

  const codigo = (params.codigo || "").toString().trim();
  if(!codigo){
    return jsonResponse({ success:false, message:"Falta el código del producto." });
  }

  const cantidad = Number(params.cantidad || 0);
  if(!cantidad){
    return jsonResponse({ success:false, message:"Ingresá una cantidad válida, distinta de 0." });
  }

  // Si ya estamos adentro de conIdempotencia (cola offline, siempre
  // manda idOperacionSync), esa misma ejecución ya tiene el script
  // lock tomado — no volver a pedirlo acá (ver _dentroDeLockIdempotencia).
  let lock = null;
  let obtuvoLock = true;
  if (!_dentroDeLockIdempotencia) {
    lock = LockService.getScriptLock();
    obtuvoLock = lock.tryLock(10000);
  }

  if(!obtuvoLock){
    return jsonResponse({
      success:false,
      message:"El sistema está ocupado actualizando stock en este momento. Probá de nuevo en unos segundos."
    });
  }

  try{

    const sheet = getSheet();
    if(!sheet){
      return jsonResponse({ success:false, message:"No se encontró la hoja 'PRODUCTOS'." });
    }

    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    const idxCodigo = headerRow.indexOf("CODIGO");
    const idxStock  = headerRow.indexOf("STOCK");

    if(idxCodigo === -1 || idxStock === -1){
      return jsonResponse({ success:false, message:"La hoja 'PRODUCTOS' no tiene las columnas CODIGO/STOCK esperadas." });
    }

    let filaIndex = -1;
    for(let i = 1; i < data.length; i++){
      if(String(data[i][idxCodigo]).trim() === codigo){
        filaIndex = i;
        break;
      }
    }

    if(filaIndex === -1){
      const ub = ubicarProductoPerdido(data, codigo, "");
      if(ub.fila === -1) return jsonResponse(respuestaProductoNoEncontrado(codigo));
      filaIndex = ub.fila;
      console.log("sumarStockProducto: '" + codigo + "' ya no existe; se aplica a '" + String(data[filaIndex][idxCodigo]).trim() + "' (" + ub.motivo + ")");
    }

    const stockActual = Number(data[filaIndex][idxStock] || 0);
    // Math.max(0, ...) — un delta negativo (corrección a la baja) no
    // debería poder dejar el stock en negativo si, para cuando esto
    // se sincroniza, el stock real en Sheets ya bajó más de lo que la
    // caja offline sabía (ej. se vendió en otra caja mientras tanto).
    const stockNuevo = Math.max(0, stockActual + cantidad);

    sheet.getRange(filaIndex + 1, idxStock + 1).setValue(stockNuevo);

    // Solo cambió el stock de un código: se parchea la caché en vez de tirarla
    const codigoReal = String(data[filaIndex][idxCodigo]).trim() || codigo;
    const cambio = {};
    cambio[codigoReal] = stockNuevo;
    notificarCambioStock(cambio);

    return jsonResponse({
      success: true,
      codigo: codigoReal,
      stockAnterior: stockActual,
      stockNuevo: stockNuevo
    });

  } finally {
    if (lock) lock.releaseLock();
  }

}

/**
 * Elimina un producto de la hoja PRODUCTOS por código.
 */
function eliminarProductoPorCodigo(e){
  // Ya no borra para siempre: mueve el producto a la hoja PAPELERA_PRODUCTOS
  // (ver eliminarProductoAPapelera, al final de este archivo).
  return eliminarProductoAPapelera(e);
}

/* =====================================================================
   INGRESO DE PRODUCTOS (recepción de mercadería con proveedor)
   Sección nueva del panel: permite escanear/tipear un código de barras.
   - Si el producto YA existe en PRODUCTOS: recupera nombre/categoría/
     precio actual, suma la cantidad recibida al stock y, si se indicó
     un precio nuevo, lo actualiza.
   - Si el producto NO existe: lo crea (con el código escaneado, o uno
     generado automáticamente si no se escaneó nada) usando el nombre,
     categoría, precio y stock inicial que se completen en el formulario.
   En ambos casos deja un registro en la hoja INGRESOS_PRODUCTOS con
   fecha, código, precio y proveedor, para poder armar el informe/PDF.
===================================================================== */

function getOrCrearHojaIngresosProductos(){
  const ss = getSS();
  let sheet = getHoja("INGRESOS_PRODUCTOS");
  if(!sheet){
    sheet = ss.insertSheet("INGRESOS_PRODUCTOS");
    sheet.appendRow([
      "INGRESO_ID", "BOLETA_ID", "FECHA", "CODIGO", "PRODUCTO", "CATEGORIA",
      "CANTIDAD", "MONEDA", "TIPO_CAMBIO", "PRECIO_INGRESO", "PRECIO_ARS", "PRECIO_ANTERIOR", "PRECIO_NUEVO", "PROVEEDOR",
      "PROVEEDOR_CONTACTO", "OBSERVACIONES", "ES_PRODUCTO_NUEVO", "USUARIO"
    ]);
    sheet.setFrozenRows(1);
    _hojaCache["INGRESOS_PRODUCTOS"] = sheet;
    return sheet;
  }
  // Migración: instalaciones que ya tenían esta hoja de antes de que
  // existieran las boletas agrupadas no tienen la columna BOLETA_ID.
  // Se agrega sin tocar las filas viejas — quedan con BOLETA_ID vacío
  // (siguen existiendo como ingresos sueltos, simplemente no forman
  // parte de ninguna boleta con pagos).
  const headerActual = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if(String(headerActual[1] || "") !== "BOLETA_ID"){
    sheet.insertColumnAfter(1);
    sheet.getRange(1, 2).setValue("BOLETA_ID");
  }
  // Migración: instalaciones anteriores a la separación entre precio de
  // ingreso (costo, lo que se le paga al proveedor) y precio de venta
  // (PRECIO_ANTERIOR/PRECIO_NUEVO) no tienen la columna PRECIO_INGRESO
  // — se agrega antes de PRECIO_ANTERIOR. Las filas viejas quedan con
  // PRECIO_INGRESO vacío (no hay forma de reconstruir retroactivamente
  // qué costo se pagó en esos ingresos ya guardados).
  const headerActualizado = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const idxPrecioAnterior = headerActualizado.indexOf("PRECIO_ANTERIOR");
  if(idxPrecioAnterior !== -1 && headerActualizado.indexOf("PRECIO_INGRESO") === -1){
    sheet.insertColumnBefore(idxPrecioAnterior + 1);
    sheet.getRange(1, idxPrecioAnterior + 1).setValue("PRECIO_INGRESO");
  }
  // Migración: instalaciones anteriores a poder cargar el precio de
  // ingreso en dólares no tienen MONEDA/TIPO_CAMBIO/PRECIO_ARS. Se
  // agregan antes de PRECIO_INGRESO. Filas viejas quedan con MONEDA
  // vacía — se interpretan como ARS (ver _monedaComoTexto).
  const headerConPrecioIngreso = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const idxPrecioIngresoActual = headerConPrecioIngreso.indexOf("PRECIO_INGRESO");
  if(idxPrecioIngresoActual !== -1 && headerConPrecioIngreso.indexOf("MONEDA") === -1){
    sheet.insertColumnBefore(idxPrecioIngresoActual + 1);
    sheet.insertColumnBefore(idxPrecioIngresoActual + 2);
    sheet.getRange(1, idxPrecioIngresoActual + 1).setValue("MONEDA");
    sheet.getRange(1, idxPrecioIngresoActual + 2).setValue("TIPO_CAMBIO");
  }
  const headerConMoneda = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const idxPrecioIngresoFinal = headerConMoneda.indexOf("PRECIO_INGRESO");
  if(idxPrecioIngresoFinal !== -1 && headerConMoneda.indexOf("PRECIO_ARS") === -1){
    sheet.insertColumnAfter(idxPrecioIngresoFinal + 1);
    sheet.getRange(1, idxPrecioIngresoFinal + 2).setValue("PRECIO_ARS");
  }
  return sheet;
}

/** Returns the BOLETAS_PROVEEDOR sheet (cabecera de cada recepción de mercadería), creating it the first time it's needed */
function getOrCrearHojaBoletasProveedor(){
  const ss = getSS();
  let sheet = getHoja("BOLETAS_PROVEEDOR");
  if(!sheet){
    sheet = ss.insertSheet("BOLETAS_PROVEEDOR");
    sheet.appendRow([
      "BOLETA_ID", "FECHA", "PROVEEDOR", "PROVEEDOR_CONTACTO",
      "NUMERO_BOLETA", "OBSERVACIONES", "TOTAL", "USUARIO", "FECHA_CREACION",
      "ORIGEN", "SIN_DEUDA"
    ]);
    sheet.setFrozenRows(1);
    _hojaCache["BOLETAS_PROVEEDOR"] = sheet;
    return sheet;
  }
  // Migración: instalaciones que ya tenían esta hoja de antes de que
  // existiera el origen "nuestra sucursal/depósito" no tienen las
  // columnas ORIGEN/SIN_DEUDA. Se agregan al final; las filas viejas
  // quedan con SIN_DEUDA vacío, que se interpreta como "con deuda"
  // (comportamiento anterior, sin cambios).
  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if(header.indexOf("ORIGEN") === -1){
    const ultimaCol = sheet.getLastColumn();
    sheet.getRange(1, ultimaCol + 1).setValue("ORIGEN");
    sheet.getRange(1, ultimaCol + 2).setValue("SIN_DEUDA");
  }
  return sheet;
}

/** Returns the PAGOS_PROVEEDOR sheet, creating it the first time it's needed */
/** Returns the PAGOS_PROVEEDOR sheet, creating it the first time it's needed */
function getOrCrearHojaPagosProveedor(){
  const ss = getSS();
  let sheet = getHoja("PAGOS_PROVEEDOR");
  if(!sheet){
    sheet = ss.insertSheet("PAGOS_PROVEEDOR");
    sheet.appendRow([
      "PAGO_ID", "FECHA", "BOLETA_ID", "PROVEEDOR",
      "MONTO", "MONEDA", "FORMA_PAGO", "OBSERVACIONES", "USUARIO"
    ]);
    sheet.setFrozenRows(1);
    _hojaCache["PAGOS_PROVEEDOR"] = sheet;
    return sheet;
  }
  // Migración: instalaciones que ya tenían esta hoja de antes de poder
  // pagarle al proveedor en dólares no tienen la columna MONEDA. Se
  // agrega después de MONTO; los pagos viejos quedan con MONEDA vacía,
  // que se interpreta como ARS (todos los pagos históricos se hicieron
  // en pesos, ya que la opción de dólares no existía todavía).
  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const idxMonto = header.indexOf("MONTO");
  if(idxMonto !== -1 && header.indexOf("MONEDA") === -1){
    sheet.insertColumnAfter(idxMonto + 1);
    sheet.getRange(1, idxMonto + 2).setValue("MONEDA");
  }
  return sheet;
}

/**
 * Calcula, a partir de las líneas de INGRESOS_PRODUCTOS de una boleta,
 * cuánto se debe en cada moneda — SIN convertir nada entre monedas
 * (a diferencia de TOTAL, que es un equivalente en pesos). Esto es lo
 * que permite que una boleta se pague parte en pesos y parte en
 * dólares: cada moneda tiene su propio total y su propio saldo,
 * independiente del tipo de cambio del día.
 */
function _totalesNativosBoleta(boletaId){
  const hojaIngresos = getOrCrearHojaIngresosProductos();
  const items = rowsToObjects(hojaIngresos.getDataRange().getValues())
    .filter(i => String(i.BOLETA_ID) === String(boletaId));

  let totalArs = 0;
  let totalUsd = 0;
  items.forEach(i => {
    const cantidad = Number(i.CANTIDAD || 0);
    const precioIngreso = Number(i.PRECIO_INGRESO || 0);
    const moneda = String(i.MONEDA || "ARS").trim().toUpperCase() === "USD" ? "USD" : "ARS";
    const subtotal = cantidad * precioIngreso;
    if(moneda === "USD") totalUsd += subtotal;
    else totalArs += subtotal;
  });

  return {
    totalArs: Math.round(totalArs * 100) / 100,
    totalUsd: Math.round(totalUsd * 100) / 100
  };
}

/**
 * Suma los pagos ya registrados contra una boleta, separados por
 * moneda (ARS/USD). Un pago sin MONEDA (registrado antes de que
 * existiera esta opción) se cuenta como ARS.
 */
function _pagadoNativoBoleta(boletaId, pagos){
  let pagadoArs = 0;
  let pagadoUsd = 0;
  pagos.forEach(p => {
    if(String(p.BOLETA_ID) !== String(boletaId)) return;
    const moneda = String(p.MONEDA || "ARS").trim().toUpperCase() === "USD" ? "USD" : "ARS";
    const monto = Number(p.MONTO || 0);
    if(moneda === "USD") pagadoUsd += monto;
    else pagadoArs += monto;
  });
  return {
    pagadoArs: Math.round(pagadoArs * 100) / 100,
    pagadoUsd: Math.round(pagadoUsd * 100) / 100
  };
}

/** Normaliza fechas que Sheets puede devolver como objeto Date a texto "yyyy-MM-dd" */
function _fechaComoTexto(valor){
  const tz = zonaNegocio() || "GMT-3";
  if(valor instanceof Date) return Utilities.formatDate(valor, tz, "yyyy-MM-dd");
  return valor ? String(valor).substring(0, 10) : "";
}

/**
 * Genera un código de producto nuevo, corto y sin repetir contra los
 * códigos ya usados en PRODUCTOS. Se usa cuando en "Ingreso de
 * Productos" el producto no existe y no se escaneó/tipeó ningún
 * código de barras propio.
 */
/**
 * Cambia PUBLICADO / DESTACADO / OFERTA de muchos productos de una vez
 * (ej. al cambiar de temporada). Toca SOLO esas columnas: no pisa
 * stock, precio, fotos ni nada más. Una sola lectura y una escritura
 * por columna, dentro del lock.
 * data: { codigos: [...], cambios: { PUBLICADO?: "SI"|"NO", DESTACADO?: ..., OFERTA?: ... } }
 */
function actualizarProductosLote(data){
  const sesion = _sesionActual || {};
  if(String(sesion.r || "").toLowerCase() === "vendedor" || String(data.rol || "").toLowerCase() === "vendedor"){
    return jsonResponse({ success:false, message:"Tu usuario (Vendedor) no tiene permiso para modificar productos." });
  }
  const COLUMNAS = ["PUBLICADO", "DESTACADO", "OFERTA"];
  const cambios = {};
  const entrada = data.cambios || {};
  COLUMNAS.forEach(k => {
    if(entrada[k] === undefined || entrada[k] === null || entrada[k] === "") return;
    const v = String(entrada[k]).trim().toUpperCase();
    if(v !== "SI" && v !== "NO") return;
    cambios[k] = v;
  });
  if(!Object.keys(cambios).length) return jsonResponse({ success:false, message:"No se indicó qué cambiar." });
  const codigos = (Array.isArray(data.codigos) ? data.codigos : []).map(c => String(c || "").trim()).filter(Boolean);
  if(!codigos.length) return jsonResponse({ success:false, message:"Seleccioná al menos un producto." });
  if(codigos.length > 5000) return jsonResponse({ success:false, message:"Son demasiados productos de una vez (máximo 5000)." });

  const lockPropio = !_dentroDeLockIdempotencia; // desde la cola offline el lock ya está tomado
  const lock = LockService.getScriptLock();
  if(lockPropio && !lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado guardando otros cambios, probá de nuevo en unos segundos." });
  try{
    const sheet = getSheet();
    if(!sheet) return jsonResponse({ success:false, message:"No se encontró la hoja 'PRODUCTOS'." });
    const datos = sheet.getDataRange().getValues();
    const enc = datos[0].map(h => String(h || "").trim());
    const iCod = enc.indexOf("CODIGO") === -1 ? 0 : enc.indexOf("CODIGO");
    // Columnas que falten (planillas viejas sin DESTACADO u OFERTA) se crean
    Object.keys(cambios).forEach(k => {
      if(enc.indexOf(k) === -1){
        sheet.getRange(1, enc.length + 1).setValue(k);
        enc.push(k);
        datos.forEach((fila, i) => { fila.push(i === 0 ? k : ""); });
      }
    });

    const filaPorCodigo = {};
    for(let i = 1; i < datos.length; i++){
      const c = normalizarCodigo(datos[i][iCod]);
      if(c && filaPorCodigo[c] === undefined) filaPorCodigo[c] = i;
    }

    const filasAfectadas = new Set();
    const noEncontrados = [];
    let sinCambios = 0;
    codigos.forEach(cod => {
      let i = filaPorCodigo[normalizarCodigo(cod)];
      if(i === undefined){
        const nuevo = codigoActualDeRenombrado(cod);
        if(nuevo) i = filaPorCodigo[normalizarCodigo(nuevo)];
      }
      if(i === undefined){ noEncontrados.push(cod); return; }
      let cambio = false;
      Object.keys(cambios).forEach(k => {
        const j = enc.indexOf(k);
        const actual = String(datos[i][j] || "").trim().toUpperCase() || (k === "PUBLICADO" ? "SI" : "NO");
        if(actual !== cambios[k]){ datos[i][j] = cambios[k]; cambio = true; }
      });
      if(cambio) filasAfectadas.add(i); else sinCambios++;
    });

    if(filasAfectadas.size){
      // Se escribe cada columna cambiada entera (una llamada por columna)
      Object.keys(cambios).forEach(k => {
        const j = enc.indexOf(k);
        const columna = datos.slice(1).map(f => [f[j] === undefined ? "" : f[j]]);
        if(columna.length) sheet.getRange(2, j + 1, columna.length, 1).setValues(columna);
      });
      invalidarCacheProductos();
    }

    return jsonResponse({ success:true, actualizados: filasAfectadas.size, sinCambios: sinCambios, noEncontrados: noEncontrados });
  } finally {
    if(lockPropio) lock.releaseLock();
  }
}

/* ---------------------------------------------------------------------
   PRODUCTOS QUE CAMBIARON DE CÓDIGO
   La cola offline de una caja puede traer cambios con un código que ya
   no existe porque, mientras tanto, en otra caja o en el panel web le
   pusieron otro (ej. el PRD-... provisorio se reemplazó por el código
   de barras) o lo borraron. Se guarda un registro de cambios de código
   (viejo → nuevo) para seguirle el rastro.
--------------------------------------------------------------------- */
const PROP_CODIGOS_RENOMBRADOS = "VEEK_CODIGOS_RENOMBRADOS";

function _leerRenombres(){
  try{ return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_CODIGOS_RENOMBRADOS) || "{}") || {}; }
  catch(e){ return {}; }
}

function registrarRenombreCodigo(viejo, nuevo){
  try{
    const v = normalizarCodigo(viejo), n = String(nuevo || "").trim();
    if(!v || !n || v === normalizarCodigo(n)) return;
    const mapa = _leerRenombres();
    delete mapa[v];
    mapa[v] = n;
    const claves = Object.keys(mapa);
    if(claves.length > 400) claves.slice(0, claves.length - 400).forEach(k => delete mapa[k]);
    PropertiesService.getScriptProperties().setProperty(PROP_CODIGOS_RENOMBRADOS, JSON.stringify(mapa));
  }catch(e){ console.error("registrarRenombreCodigo: " + e.message); }
}

/** Código actual de un código viejo (sigue la cadena de cambios) o "" si no hubo cambio. */
function codigoActualDeRenombrado(codigo){
  const mapa = _leerRenombres();
  let actual = normalizarCodigo(codigo), salida = "", vueltas = 0;
  while(mapa[actual] && vueltas < 10){
    salida = mapa[actual];
    actual = normalizarCodigo(salida);
    vueltas++;
  }
  return salida;
}

/**
 * Busca la fila de un producto cuyo código ya no aparece:
 * 1) por el registro de cambios de código;
 * 2) por el nombre, si hay UN solo producto con ese nombre exacto
 *    (solo cuando se conoce el nombre, ej. una edición completa).
 * data: getValues() de PRODUCTOS (fila 0 = encabezados).
 */
function ubicarProductoPerdido(data, codigo, nombre){
  const enc = data[0].map(h => String(h || "").trim());
  const iCod = enc.indexOf("CODIGO") === -1 ? 0 : enc.indexOf("CODIGO");
  const iNom = enc.indexOf("PRODUCTO");
  const nuevo = codigoActualDeRenombrado(codigo);
  if(nuevo){
    for(let i = 1; i < data.length; i++){
      if(normalizarCodigo(data[i][iCod]) === normalizarCodigo(nuevo)) return { fila: i, motivo: "le cambiaron el código a " + nuevo };
    }
  }
  const nom = String(nombre || "").trim().toUpperCase().replace(/\s+/g, " ");
  if(nom && iNom !== -1){
    const coincidencias = [];
    for(let i = 1; i < data.length; i++){
      if(String(data[i][iNom] || "").trim().toUpperCase().replace(/\s+/g, " ") === nom) coincidencias.push(i);
    }
    if(coincidencias.length === 1) return { fila: coincidencias[0], motivo: "mismo nombre, otro código" };
  }
  return { fila: -1, motivo: "" };
}

/**
 * Respuesta cuando el producto no está. Si está en la papelera, el
 * cambio ya no tiene sentido: descartar:true le dice a la caja que lo
 * saque de la cola en vez de reintentarlo para siempre.
 */
function respuestaProductoNoEncontrado(codigo){
  let enPapelera = false;
  try{
    const pap = getHoja(HOJA_PAPELERA_PRODUCTOS);
    if(pap && pap.getLastRow() > 1){
      const cods = pap.getRange(2, 3, pap.getLastRow() - 1, 1).getValues();
      const buscado = normalizarCodigo(codigo);
      enPapelera = cods.some(f => normalizarCodigo(f[0]) === buscado);
    }
  }catch(e){}
  if(enPapelera){
    return { success: false, descartar: true, productoEliminado: true,
      message: "El producto '" + codigo + "' fue eliminado (está en la papelera de productos): este cambio ya no se aplica. Si lo necesitás, restauralo desde Productos → Papelera." };
  }
  return { success: false, productoNoEncontrado: true,
    message: "No se encontró un producto con el código '" + codigo + "' (¿le cambiaron el código o lo borraron en otra caja?)." };
}

function generarCodigoProductoNuevo(){
  const sheet = getSheet();
  const data = sheet ? sheet.getDataRange().getValues() : [];
  data.shift();
  const existentes = new Set(data.map(row => String(row[0]).trim()));
  let codigo;
  do {
    codigo = "PRD-" + Date.now() + "-" + Math.floor(Math.random() * 900 + 100);
  } while (existentes.has(codigo));
  return codigo;
}

/**
 * Da de alta o actualiza UN producto dentro de una recepción de
 * mercadería (crea el producto si no existe, o suma stock/actualiza
 * precio si ya existe). Es el mismo criterio que tenía antes
 * registrarIngresoProducto, ahora aplicado a un solo ítem de la
 * boleta — se llama una vez por cada línea de productos que llegó.
 * "filas" es la matriz de PRODUCTOS ya leída (se le va agregando la
 * fila nueva en memoria si corresponde, para que dos ítems nuevos con
 * el mismo código en la misma boleta no se pisen entre sí).
 */
/**
 * Asegura que la hoja PRODUCTOS tenga la columna COSTO_USD_REF —
 * guarda el último precio de ingreso cargado en dólares para ese
 * producto, como referencia para futuros ingresos (no se usa en el
 * catálogo/POS, es solo informativo). Se agrega al final de la hoja
 * sin tocar columnas ni filas existentes.
 */
function _asegurarColumnaCostoUsdRef(sheet){
  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if(header.indexOf("COSTO_USD_REF") === -1){
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("COSTO_USD_REF");
  }
}

function _aplicarItemIngresoProducto(sheet, filas, headerRow, item){

  const idxCodigo = headerRow.indexOf("CODIGO");
  const idxProducto = headerRow.indexOf("PRODUCTO");
  const idxCategoria = headerRow.indexOf("CATEGORIA");
  const idxPrecio = headerRow.indexOf("PRECIO");
  const idxStock = headerRow.indexOf("STOCK");
  const idxCostoUsdRef = headerRow.indexOf("COSTO_USD_REF");

  const nombre = (item.producto || "").toString().trim();
  const cantidad = Number(item.cantidad || 0);

  // Precio de INGRESO (costo pagado al proveedor) — es lo que determina
  // cuánto se le debe al proveedor por esta boleta (ver totalBoleta en
  // registrarIngresoProducto). NUNCA se usa para el catálogo/POS.
  const precioIngresoNuevo = item.precio !== undefined && item.precio !== ""
    ? Number(item.precio) : null;

  // Moneda en la que se cargó el precio de ingreso: "USD" o "ARS"
  // (por defecto ARS si no viene nada, para no romper llamadas viejas).
  const monedaItem = String(item.moneda || "ARS").trim().toUpperCase() === "USD" ? "USD" : "ARS";
  const tipoCambioItem = (monedaItem === "USD" && item.tipoCambio) ? Number(item.tipoCambio) : null;
  // Precio de ingreso convertido a pesos — es lo que realmente compone
  // la deuda con el proveedor. Si viene precalculado desde el
  // frontend se usa ese; si no, se recalcula acá como respaldo.
  const precioARSItem = item.precioARS !== undefined && item.precioARS !== "" && !isNaN(Number(item.precioARS))
    ? Number(item.precioARS)
    : (monedaItem === "USD" && tipoCambioItem && precioIngresoNuevo !== null
        ? precioIngresoNuevo * tipoCambioItem
        : precioIngresoNuevo);

  // Precio de VENTA (al público) — el que ve el cliente en el POS y en
  // la página del catálogo (columna PRECIO de PRODUCTOS). Es un dato
  // aparte del precio de ingreso: lo que se le paga al proveedor no
  // tiene por qué ser lo mismo que lo que se le cobra al cliente.
  const precioVentaNuevo = item.precioVenta !== undefined && item.precioVenta !== ""
    ? Number(item.precioVenta) : null;

  if(!cantidad || cantidad <= 0){
    return { error: "Ingresá una cantidad válida (mayor a 0) para \"" + (nombre || item.codigo || "un producto") + "\"." };
  }

  let codigo = (item.codigo || "").toString().trim();

  let filaIndex = -1;
  if(codigo){
    for(let i = 1; i < filas.length; i++){
      if(normalizarCodigo(filas[i][idxCodigo]) === normalizarCodigo(codigo)){
        filaIndex = i;
        break;
      }
    }
  }

  let esProductoNuevo = false;
  let stockNuevo = cantidad;
  let precioVentaAnterior = "";
  let productoFinal = nombre;
  let categoriaFinal = (item.categoria || "").toString().trim();

  if(filaIndex !== -1){
    // ---- Producto existente: sumar stock y, si corresponde, precio de VENTA ----
    precioVentaAnterior = Number(filas[filaIndex][idxPrecio] || 0);
    const stockActual = Number(filas[filaIndex][idxStock] || 0);
    stockNuevo = stockActual + cantidad;
    sheet.getRange(filaIndex + 1, idxStock + 1).setValue(stockNuevo);
    filas[filaIndex][idxStock] = stockNuevo;

    if(precioVentaNuevo !== null && !isNaN(precioVentaNuevo) && precioVentaNuevo !== precioVentaAnterior){
      sheet.getRange(filaIndex + 1, idxPrecio + 1).setValue(precioVentaNuevo);
      filas[filaIndex][idxPrecio] = precioVentaNuevo;
    }

    // Guardar el precio de ingreso en USD como referencia futura — solo
    // si este ingreso puntual vino en dólares; si vino en pesos se deja
    // el último valor en USD que ya hubiera, sin borrarlo.
    if(idxCostoUsdRef !== -1 && monedaItem === "USD" && precioIngresoNuevo !== null && !isNaN(precioIngresoNuevo)){
      sheet.getRange(filaIndex + 1, idxCostoUsdRef + 1).setValue(precioIngresoNuevo);
      filas[filaIndex][idxCostoUsdRef] = precioIngresoNuevo;
    }

    productoFinal = filas[filaIndex][idxProducto];
    categoriaFinal = filas[filaIndex][idxCategoria];

  } else {
    // ---- Producto nuevo: se crea ----
    if(!nombre){
      return { error: "El producto \"" + codigo + "\" no existe todavía — indicá el nombre para crearlo." };
    }
    if(!codigo){
      codigo = generarCodigoProductoNuevo();
    } else {
      const conflicto = conflictoCodigoProducto(filas, codigo, -1, false);
      if(conflicto){
        return { error: conflicto };
      }
    }

    esProductoNuevo = true;
    precioVentaAnterior = "";

    // Si no se cargó un precio de venta propio para el producto nuevo,
    // se usa el precio de ingreso como respaldo — mejor eso que dejar
    // el producto publicado a $0 en el catálogo.
    const precioVentaInicial = (precioVentaNuevo !== null && !isNaN(precioVentaNuevo))
      ? precioVentaNuevo
      : (precioIngresoNuevo !== null && !isNaN(precioIngresoNuevo) ? precioIngresoNuevo : 0);

    const valoresPorColumna = {
      CODIGO: codigo,
      PRODUCTO: nombre,
      CATEGORIA: categoriaFinal,
      PRECIO: precioVentaInicial,
      STOCK: cantidad,
      PUBLICADO: "SI",
      DESTACADO: "NO",
      OFERTA: "NO",
      DESCRIPCION: "",
      IMAGEN: "",
      COSTO_USD_REF: (monedaItem === "USD" && precioIngresoNuevo !== null && !isNaN(precioIngresoNuevo)) ? precioIngresoNuevo : "",
      // Producto nuevo -> va al final del orden actual del catálogo,
      // igual que en guardarProducto. Sin esto la columna ORDEN queda
      // vacía y el producto no aparece bien ubicado en el catálogo/POS
      // ni se puede reordenar por arrastre junto a los demás.
      ORDEN: obtenerSiguienteOrden(sheet, headerRow)
    };

    const fila = headerRow.map(header => {
      const key = String(header || "").trim();
      return valoresPorColumna.hasOwnProperty(key) ? valoresPorColumna[key] : "";
    });

    sheet.appendRow(fila);
    filas.push(fila); // para que otro ítem de la misma boleta lo vea si se repite el código
    productoFinal = nombre;
  }

  // El precio USADO para calcular cuánto se le debe al proveedor en esta
  // boleta es siempre el precio de INGRESO (costo) convertido a pesos
  // (precioARS) — nunca el de venta. Si por algún motivo no se cargó un
  // precio de ingreso puntual, se usa el de venta como respaldo (mejor
  // eso que sumar $0 a la boleta).
  const precioUsado = (precioARSItem !== null && !isNaN(precioARSItem))
    ? precioARSItem
    : (precioVentaNuevo !== null && !isNaN(precioVentaNuevo) ? precioVentaNuevo : 0);

  return {
    codigo, productoFinal, categoriaFinal, cantidad,
    precioAnterior: precioVentaAnterior,
    precioNuevo: precioVentaNuevo,
    precioIngreso: precioIngresoNuevo,
    moneda: monedaItem,
    tipoCambio: tipoCambioItem,
    precioARS: precioARSItem,
    precioUsado, esProductoNuevo, stockNuevo
  };
}

/**
 * Registra la recepción de mercadería de UNA boleta/remito de un
 * proveedor (POST). Recibe la cabecera —proveedor, proveedorContacto,
 * numeroBoleta, observaciones, fecha, usuario— y "items": un array con
 * cada producto que llegó ({codigo, producto, categoria, precio,
 * cantidad}).
 *
 * Todos los ítems quedan agrupados bajo UN solo BOLETA_ID: si ya
 * existía una boleta con el mismo proveedor + fecha + N° de boleta
 * (por ejemplo, se está completando la carga en varias tandas), se le
 * suman los ítems nuevos en vez de crear una boleta duplicada. Cada
 * ítem, además, actualiza stock/precio del producto correspondiente
 * en PRODUCTOS, igual que antes.
 */
function registrarIngresoProducto(data){

  if(String((data && data.rol) || "").trim().toLowerCase() === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para registrar ingresos de productos."
    });
  }

  const proveedor = ((data && data.proveedor) || "").toString().trim();
  if(!proveedor){
    return jsonResponse({ success:false, message:"Ingresá el proveedor de esta mercadería." });
  }

  const items = Array.isArray(data && data.items) ? data.items : [];
  if(items.length === 0){
    return jsonResponse({ success:false, message:"Agregá al menos un producto a la boleta antes de guardar." });
  }

  // Desde la cola offline (syncLote / conIdempotencia) el lock de script
  // ya está tomado por esta misma ejecución y no es reentrante: tomarlo
  // de nuevo agotaría la espera y la boleta nunca se subiría.
  const lockPropio = !_dentroDeLockIdempotencia;
  const lock = LockService.getScriptLock();
  const obtuvoLock = lockPropio ? lock.tryLock(10000) : true;
  if(!obtuvoLock){
    return jsonResponse({
      success:false,
      message:"El sistema está ocupado en este momento. Probá de nuevo en unos segundos."
    });
  }

  try{

    const sheet = getSheet();
    if(!sheet){
      return jsonResponse({ success:false, message:"No se encontró la hoja 'PRODUCTOS'." });
    }
    _asegurarColumnaCostoUsdRef(sheet);
    asegurarColumnaOrden(sheet);

    const filas = sheet.getDataRange().getValues();
    const headerRow = filas[0];

    const resultados = [];
    let totalBoleta = 0;

    for(let idx = 0; idx < items.length; idx++){
      const resultado = _aplicarItemIngresoProducto(sheet, filas, headerRow, items[idx]);
      if(resultado.error){
        return jsonResponse({ success:false, message: resultado.error });
      }
      resultados.push(resultado);
      totalBoleta += resultado.cantidad * resultado.precioUsado;
    }

    // ---- Cabecera de la boleta: reusar si ya existe (mismo proveedor + fecha + N° de boleta), o crear una nueva ----
    const proveedorContacto = ((data && data.proveedorContacto) || "").toString().trim();
    const numeroBoleta = ((data && data.numeroBoleta) || "").toString().trim();
    const observaciones = ((data && data.observaciones) || "").toString().trim();
    const usuario = ((data && data.usuario) || "").toString().trim();
    const fecha = data.fecha ? data.fecha : Utilities.formatDate(new Date(), zonaNegocio() || "GMT-3", "yyyy-MM-dd");
    // Origen de la mercadería: "PROVEEDOR" (default) o "DEPOSITO" —
    // cuando viene de nuestra propia sucursal/depósito no genera deuda,
    // así que la boleta queda marcada SIN_DEUDA y no aparece en los
    // listados de deuda a proveedores (ver getBoletasProveedores/
    // getProveedoresConSaldo).
    const origen = String((data && data.origen) || "PROVEEDOR").trim().toUpperCase() === "DEPOSITO" ? "DEPOSITO" : "PROVEEDOR";
    const sinDeuda = origen === "DEPOSITO" || (data && data.sinDeuda === true);

    const hojaBoletas = getOrCrearHojaBoletasProveedor();
    const filasBoletas = hojaBoletas.getDataRange().getValues();
    const bHeader = filasBoletas[0];
    const bIdxId = bHeader.indexOf("BOLETA_ID");
    const bIdxFecha = bHeader.indexOf("FECHA");
    const bIdxProveedor = bHeader.indexOf("PROVEEDOR");
    const bIdxNumero = bHeader.indexOf("NUMERO_BOLETA");
    const bIdxTotal = bHeader.indexOf("TOTAL");
    const bIdxSinDeuda = bHeader.indexOf("SIN_DEUDA");

    let filaBoletaIndex = -1;
    for(let i = 1; i < filasBoletas.length; i++){
      const mismaFecha = _fechaComoTexto(filasBoletas[i][bIdxFecha]) === fecha;
      const mismoProveedor = String(filasBoletas[i][bIdxProveedor] || "").trim().toLowerCase() === proveedor.toLowerCase();
      const mismoNumero = String(filasBoletas[i][bIdxNumero] || "").trim() === numeroBoleta;
      if(mismaFecha && mismoProveedor && mismoNumero){
        filaBoletaIndex = i;
        break;
      }
    }

    let boletaId;
    if(filaBoletaIndex !== -1){
      boletaId = filasBoletas[filaBoletaIndex][bIdxId];
      const totalActual = Number(filasBoletas[filaBoletaIndex][bIdxTotal] || 0);
      hojaBoletas.getRange(filaBoletaIndex + 1, bIdxTotal + 1).setValue(totalActual + totalBoleta);
      if(bIdxSinDeuda !== -1){
        hojaBoletas.getRange(filaBoletaIndex + 1, bIdxSinDeuda + 1).setValue(sinDeuda ? "SI" : "NO");
      }
    } else {
      boletaId = "BOL-" + Date.now();
      hojaBoletas.appendRow([
        boletaId, fecha, proveedor, proveedorContacto,
        numeroBoleta, observaciones, totalBoleta, usuario,
        Utilities.formatDate(new Date(), zonaNegocio() || "GMT-3", "yyyy-MM-dd HH:mm:ss"),
        origen, sinDeuda ? "SI" : "NO"
      ]);
    }

    // ---- Una línea por ítem en INGRESOS_PRODUCTOS, todas con el mismo BOLETA_ID ----
    const hojaIngresos = getOrCrearHojaIngresosProductos();
    resultados.forEach(r => {
      const ingresoId = "ING-" + Date.now() + "-" + Math.floor(Math.random() * 900 + 100);
      hojaIngresos.appendRow([
        ingresoId,
        boletaId,
        fecha,
        r.codigo,
        r.productoFinal,
        r.categoriaFinal,
        r.cantidad,
        r.moneda, // MONEDA en la que se cargó el precio de ingreso (ARS/USD)
        r.tipoCambio !== null && r.tipoCambio !== undefined ? r.tipoCambio : "", // TIPO_CAMBIO usado, solo si es USD
        r.precioIngreso, // PRECIO_INGRESO — el costo pagado al proveedor, en su moneda original
        r.precioARS, // PRECIO_ARS — costo convertido a pesos, lo que compone el total de la boleta
        r.precioAnterior,
        (r.precioNuevo !== null && !isNaN(r.precioNuevo)) ? r.precioNuevo : r.precioAnterior,
        proveedor,
        proveedorContacto,
        observaciones,
        r.esProductoNuevo ? "SI" : "NO",
        usuario
      ]);
      r.ingresoId = ingresoId;
    });

    // Stock, precios y quizás productos nuevos: la web, el panel y las demás
    // cajas tienen que leer el catálogo actualizado (antes quedaba el viejo
    // en caché hasta que vencía solo).
    invalidarCacheProductos();

    return jsonResponse({
      success: true,
      boletaId: boletaId,
      totalBoleta: totalBoleta,
      origen: origen,
      sinDeuda: sinDeuda,
      items: resultados.map(r => ({
        codigo: r.codigo, producto: r.productoFinal,
        cantidad: r.cantidad, esProductoNuevo: r.esProductoNuevo
      })),
      // Stock final de cada código tocado (la caja offline lo aplica a su copia local)
      stock: resultados.reduce((acc, r) => { acc[String(r.codigo)] = Number(r.stockNuevo) || 0; return acc; }, {})
    });

  } catch(error){
    return jsonResponse({ success:false, error: error.toString() });
  } finally {
    if(lockPropio) lock.releaseLock();
  }
}

/**
 * Devuelve el listado de boletas de proveedor (una fila = una
 * recepción de mercadería completa), con su total, lo pagado (suma de
 * PAGOS_PROVEEDOR para esa boleta) y el saldo pendiente. Filtros
 * opcionales por querystring: desde, hasta (yyyy-MM-dd), proveedor
 * (contiene) y estado ("PENDIENTE" | "PARCIAL" | "PAGADA").
 */
function getBoletasProveedores(e){

  const hojaBoletas = getOrCrearHojaBoletasProveedor();
  const boletas = rowsToObjects(hojaBoletas.getDataRange().getValues());

  const hojaPagos = getOrCrearHojaPagosProveedor();
  const pagos = rowsToObjects(hojaPagos.getDataRange().getValues());

  // Las boletas de mercadería que vino de nuestra propia sucursal/depósito
  // (SIN_DEUDA = "SI") se siguen listando —para poder revisarlas e
  // imprimirlas igual que cualquier otra— pero no representan una deuda
  // real: se muestran con saldo $0/US$0 y estado especial "SIN_DEUDA".
  let resultado = boletas.map(b => {
    const esSinDeuda = String(b.SIN_DEUDA || "").trim().toUpperCase() === "SI";
    const { totalArs, totalUsd } = _totalesNativosBoleta(b.BOLETA_ID);
    const { pagadoArs, pagadoUsd } = _pagadoNativoBoleta(b.BOLETA_ID, pagos);

    const pagadoArsFinal = esSinDeuda ? totalArs : pagadoArs;
    const pagadoUsdFinal = esSinDeuda ? totalUsd : pagadoUsd;
    const saldoArs = esSinDeuda ? 0 : Math.round((totalArs - pagadoArs) * 100) / 100;
    const saldoUsd = esSinDeuda ? 0 : Math.round((totalUsd - pagadoUsd) * 100) / 100;

    let estado = esSinDeuda ? "SIN_DEUDA" : "PENDIENTE";
    if(!esSinDeuda){
      if(saldoArs <= 0 && saldoUsd <= 0) estado = "PAGADA";
      else if(pagadoArs > 0 || pagadoUsd > 0) estado = "PARCIAL";
    }

    return {
      BOLETA_ID: b.BOLETA_ID,
      FECHA: _fechaComoTexto(b.FECHA),
      PROVEEDOR: b.PROVEEDOR,
      PROVEEDOR_CONTACTO: b.PROVEEDOR_CONTACTO,
      NUMERO_BOLETA: b.NUMERO_BOLETA,
      OBSERVACIONES: b.OBSERVACIONES,
      TOTAL_ARS: totalArs,
      TOTAL_USD: totalUsd,
      PAGADO_ARS: Math.round(pagadoArsFinal * 100) / 100,
      PAGADO_USD: Math.round(pagadoUsdFinal * 100) / 100,
      SALDO_ARS: saldoArs,
      SALDO_USD: saldoUsd,
      // Se conservan TOTAL/PAGADO/SALDO (en pesos) por compatibilidad con
      // pantallas/reportes viejos que todavía no distinguen moneda.
      TOTAL: totalArs,
      PAGADO: Math.round(pagadoArsFinal * 100) / 100,
      SALDO: saldoArs,
      ESTADO: estado,
      ORIGEN: b.ORIGEN || "PROVEEDOR",
      USUARIO: b.USUARIO
    };
  });

  const params = (e && e.parameter) || {};
  const desde = (params.desde || "").toString().trim();
  const hasta = (params.hasta || "").toString().trim();
  const proveedor = (params.proveedor || "").toString().trim().toLowerCase();
  const estado = (params.estado || "").toString().trim().toUpperCase();

  if(desde) resultado = resultado.filter(b => String(b.FECHA) >= desde);
  if(hasta) resultado = resultado.filter(b => String(b.FECHA) <= hasta);
  if(proveedor) resultado = resultado.filter(b => String(b.PROVEEDOR || "").toLowerCase().indexOf(proveedor) !== -1);
  const proveedorExacto = (params.proveedorExacto || "").toString().trim().toLowerCase();
  if(proveedorExacto) resultado = resultado.filter(b => String(b.PROVEEDOR || "").trim().toLowerCase() === proveedorExacto);
  if(estado) resultado = resultado.filter(b => b.ESTADO === estado);

  resultado.sort((a, b) => String(b.FECHA).localeCompare(String(a.FECHA)) || String(b.BOLETA_ID).localeCompare(String(a.BOLETA_ID)));

  return jsonResponse({ success:true, total: resultado.length, boletas: resultado });
}

/**
 * Devuelve un resumen por proveedor (una fila = un proveedor), estilo
 * "Clientes a crédito" pero para lo que se les debe a los
 * proveedores: cuántas boletas tiene, cuánto se le compró en total,
 * cuánto se le pagó y el saldo pendiente — sumando todas sus boletas
 * (BOLETAS_PROVEEDOR) y todos los pagos hechos contra ellas
 * (PAGOS_PROVEEDOR). Se agrupa por nombre de proveedor sin importar
 * mayúsculas/espacios, para que "Coca Cola" y "coca cola " cuenten
 * como el mismo proveedor.
 */
function getProveedoresConSaldo(){

  const hojaBoletas = getOrCrearHojaBoletasProveedor();
  const boletas = rowsToObjects(hojaBoletas.getDataRange().getValues());

  const hojaPagos = getOrCrearHojaPagosProveedor();
  const pagos = rowsToObjects(hojaPagos.getDataRange().getValues());

  const porProveedor = {};

  boletas.forEach(b => {
    // Igual que en getBoletasProveedores: la mercadería que vino de
    // nuestra propia sucursal/depósito no genera deuda, así que no
    // suma al saldo de ningún proveedor.
    if(String(b.SIN_DEUDA || "").trim().toUpperCase() === "SI") return;
    const nombre = String(b.PROVEEDOR || "").trim();
    if(!nombre) return;
    const key = nombre.toLowerCase();

    if(!porProveedor[key]){
      porProveedor[key] = {
        PROVEEDOR: nombre,
        PROVEEDOR_CONTACTO: "",
        BOLETAS: 0,
        BOLETAS_PENDIENTES: 0,
        TOTAL_ARS: 0,
        TOTAL_USD: 0,
        PAGADO_ARS: 0,
        PAGADO_USD: 0,
        ULTIMA_FECHA: ""
      };
    }

    const entrada = porProveedor[key];
    const { totalArs, totalUsd } = _totalesNativosBoleta(b.BOLETA_ID);
    const { pagadoArs, pagadoUsd } = _pagadoNativoBoleta(b.BOLETA_ID, pagos);
    const saldoArs = Math.round((totalArs - pagadoArs) * 100) / 100;
    const saldoUsd = Math.round((totalUsd - pagadoUsd) * 100) / 100;
    const fecha = _fechaComoTexto(b.FECHA);

    entrada.BOLETAS += 1;
    entrada.TOTAL_ARS += totalArs;
    entrada.TOTAL_USD += totalUsd;
    entrada.PAGADO_ARS += pagadoArs;
    entrada.PAGADO_USD += pagadoUsd;
    if(saldoArs > 0 || saldoUsd > 0) entrada.BOLETAS_PENDIENTES += 1;

    // Se queda con el contacto y la fecha de la boleta más reciente de este proveedor
    if(!entrada.ULTIMA_FECHA || fecha > entrada.ULTIMA_FECHA){
      entrada.ULTIMA_FECHA = fecha;
      if(b.PROVEEDOR_CONTACTO) entrada.PROVEEDOR_CONTACTO = b.PROVEEDOR_CONTACTO;
    } else if(!entrada.PROVEEDOR_CONTACTO && b.PROVEEDOR_CONTACTO){
      entrada.PROVEEDOR_CONTACTO = b.PROVEEDOR_CONTACTO;
    }
  });

  const proveedores = Object.values(porProveedor).map(p => {
    const totalArs = Math.round(p.TOTAL_ARS * 100) / 100;
    const totalUsd = Math.round(p.TOTAL_USD * 100) / 100;
    const pagadoArs = Math.round(p.PAGADO_ARS * 100) / 100;
    const pagadoUsd = Math.round(p.PAGADO_USD * 100) / 100;
    const saldoArs = Math.round((totalArs - pagadoArs) * 100) / 100;
    const saldoUsd = Math.round((totalUsd - pagadoUsd) * 100) / 100;
    return {
      PROVEEDOR: p.PROVEEDOR,
      PROVEEDOR_CONTACTO: p.PROVEEDOR_CONTACTO,
      BOLETAS: p.BOLETAS,
      BOLETAS_PENDIENTES: p.BOLETAS_PENDIENTES,
      TOTAL_ARS: totalArs,
      TOTAL_USD: totalUsd,
      PAGADO_ARS: pagadoArs,
      PAGADO_USD: pagadoUsd,
      SALDO_ARS: saldoArs,
      SALDO_USD: saldoUsd,
      // Compatibilidad con pantallas viejas que solo mostraban pesos
      TOTAL: totalArs,
      PAGADO: pagadoArs,
      SALDO: saldoArs,
      ULTIMA_FECHA: p.ULTIMA_FECHA
    };
  });

  // Los que tienen más deuda primero (sumando ambas monedas como
  // referencia de orden); entre los que no deben nada, orden alfabético
  proveedores.sort((a, b) => {
    const deudaA = a.SALDO_ARS + a.SALDO_USD;
    const deudaB = b.SALDO_ARS + b.SALDO_USD;
    if(deudaB !== deudaA) return deudaB - deudaA;
    return String(a.PROVEEDOR).localeCompare(String(b.PROVEEDOR));
  });

  return jsonResponse({ success:true, total: proveedores.length, proveedores: proveedores });
}

/**
 * Devuelve el detalle completo de UNA boleta de proveedor: su
 * cabecera (con total/pagado/saldo), la lista de productos recibidos
 * en ella (de INGRESOS_PRODUCTOS) y el historial de pagos hechos
 * contra ella (de PAGOS_PROVEEDOR).
 */
function getDetalleBoletaProveedor(e){

  const boletaId = ((e && e.parameter && e.parameter.boletaId) || "").toString().trim();
  if(!boletaId){
    return jsonResponse({ success:false, message:"Falta el ID de la boleta." });
  }

  const hojaBoletas = getOrCrearHojaBoletasProveedor();
  const boletas = rowsToObjects(hojaBoletas.getDataRange().getValues());
  const boleta = boletas.find(b => String(b.BOLETA_ID) === boletaId);
  if(!boleta){
    return jsonResponse({ success:false, message:"No se encontró la boleta." });
  }

  const hojaIngresos = getOrCrearHojaIngresosProductos();
  const items = rowsToObjects(hojaIngresos.getDataRange().getValues())
    .filter(i => String(i.BOLETA_ID) === boletaId)
    .map(i => Object.assign({}, i, { FECHA: _fechaComoTexto(i.FECHA) }));

  const hojaPagos = getOrCrearHojaPagosProveedor();
  const pagos = rowsToObjects(hojaPagos.getDataRange().getValues())
    .filter(p => String(p.BOLETA_ID) === boletaId)
    .map(p => Object.assign({}, p, {
      FECHA: _fechaComoTexto(p.FECHA),
      MONEDA: String(p.MONEDA || "ARS").trim().toUpperCase() === "USD" ? "USD" : "ARS"
    }))
    .sort((a, b) => String(b.FECHA).localeCompare(String(a.FECHA)) || String(b.PAGO_ID).localeCompare(String(a.PAGO_ID)));

  const esSinDeuda = String(boleta.SIN_DEUDA || "").trim().toUpperCase() === "SI";
  const { totalArs, totalUsd } = _totalesNativosBoleta(boletaId);
  const { pagadoArs, pagadoUsd } = esSinDeuda
    ? { pagadoArs: totalArs, pagadoUsd: totalUsd }
    : _pagadoNativoBoleta(boletaId, pagos);
  const saldoArs = esSinDeuda ? 0 : Math.round((totalArs - pagadoArs) * 100) / 100;
  const saldoUsd = esSinDeuda ? 0 : Math.round((totalUsd - pagadoUsd) * 100) / 100;

  let estado = "PENDIENTE";
  if(esSinDeuda) estado = "SIN_DEUDA";
  else if(saldoArs <= 0 && saldoUsd <= 0) estado = "PAGADA";
  else if(pagadoArs > 0 || pagadoUsd > 0) estado = "PARCIAL";

  return jsonResponse({
    success: true,
    boleta: {
      BOLETA_ID: boleta.BOLETA_ID,
      FECHA: _fechaComoTexto(boleta.FECHA),
      PROVEEDOR: boleta.PROVEEDOR,
      PROVEEDOR_CONTACTO: boleta.PROVEEDOR_CONTACTO,
      NUMERO_BOLETA: boleta.NUMERO_BOLETA,
      OBSERVACIONES: boleta.OBSERVACIONES,
      USUARIO: boleta.USUARIO,
      FECHA_CREACION: boleta.FECHA_CREACION,
      ORIGEN: boleta.ORIGEN || "PROVEEDOR",
      TOTAL_ARS: totalArs,
      TOTAL_USD: totalUsd,
      PAGADO_ARS: Math.round(pagadoArs * 100) / 100,
      PAGADO_USD: Math.round(pagadoUsd * 100) / 100,
      SALDO_ARS: saldoArs,
      SALDO_USD: saldoUsd,
      // Compatibilidad con pantallas viejas que solo mostraban pesos
      TOTAL: totalArs,
      PAGADO: Math.round(pagadoArs * 100) / 100,
      SALDO: saldoArs,
      ESTADO: estado
    },
    items: items,
    pagos: pagos
  });
}

/**
 * Registra un pago a un proveedor (POST), aplicado contra UNA boleta
 * puntual. Recibe: boletaId, monto, formaPago, observaciones, usuario.
 * No deja registrar un pago que supere el saldo pendiente de la
 * boleta (evita saldos negativos por error de tipeo).
 */
function registrarPagoProveedor(data){

  if(String((data && data.rol) || "").trim().toLowerCase() === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para registrar pagos a proveedores."
    });
  }

  const boletaId = ((data && data.boletaId) || "").toString().trim();
  const monto = Number((data && data.monto) || 0);
  const moneda = String((data && data.moneda) || "ARS").trim().toUpperCase() === "USD" ? "USD" : "ARS";

  if(!boletaId){
    return jsonResponse({ success:false, message:"Falta el ID de la boleta." });
  }
  if(!monto || monto <= 0){
    return jsonResponse({ success:false, message:"Ingresá un monto de pago válido, mayor a 0." });
  }

  const lockPropio = !_dentroDeLockIdempotencia; // desde la cola offline el lock ya está tomado
  const lock = LockService.getScriptLock();
  const obtuvoLock = lockPropio ? lock.tryLock(10000) : true;
  if(!obtuvoLock){
    return jsonResponse({
      success:false,
      message:"El sistema está ocupado en este momento. Probá de nuevo en unos segundos."
    });
  }

  try{

    const hojaBoletas = getOrCrearHojaBoletasProveedor();
    const filasBoletas = hojaBoletas.getDataRange().getValues();
    const bHeader = filasBoletas[0];
    const bIdxId = bHeader.indexOf("BOLETA_ID");
    const bIdxProveedor = bHeader.indexOf("PROVEEDOR");

    let filaBoleta = null;
    for(let i = 1; i < filasBoletas.length; i++){
      if(String(filasBoletas[i][bIdxId]) === boletaId){
        filaBoleta = filasBoletas[i];
        break;
      }
    }
    if(!filaBoleta){
      return jsonResponse({ success:false, message:"No se encontró la boleta." });
    }

    const proveedor = String(filaBoleta[bIdxProveedor] || "");

    const { totalArs, totalUsd } = _totalesNativosBoleta(boletaId);
    const hojaPagos = getOrCrearHojaPagosProveedor();
    const pagosPrevios = rowsToObjects(hojaPagos.getDataRange().getValues());
    const { pagadoArs, pagadoUsd } = _pagadoNativoBoleta(boletaId, pagosPrevios);

    const totalMoneda = moneda === "USD" ? totalUsd : totalArs;
    const pagadoMoneda = moneda === "USD" ? pagadoUsd : pagadoArs;
    const saldoActual = Math.round((totalMoneda - pagadoMoneda) * 100) / 100;
    const simbolo = moneda === "USD" ? "US$" : "$";

    if(monto - saldoActual > 0.01){
      return jsonResponse({
        success:false,
        message:"El monto (" + simbolo + monto.toLocaleString("es-AR") + ") supera el saldo pendiente en " + moneda + " de la boleta (" + simbolo + saldoActual.toLocaleString("es-AR") + ")."
      });
    }

    const fecha = data.fecha ? data.fecha : Utilities.formatDate(new Date(), zonaNegocio() || "GMT-3", "yyyy-MM-dd");
    const pagoId = "PAG-" + Date.now();

    hojaPagos.appendRow([
      pagoId,
      fecha,
      boletaId,
      proveedor,
      monto,
      moneda,
      (data.formaPago || "EFECTIVO").toString().trim(),
      (data.observaciones || "").toString().trim(),
      (data.usuario || "").toString().trim()
    ]);

    const nuevoSaldoMoneda = Math.round((saldoActual - monto) * 100) / 100;
    // El resto de la boleta (la otra moneda) no cambia con este pago;
    // se recalcula para saber si la boleta quedó totalmente pagada.
    const saldoOtraMoneda = moneda === "USD"
      ? Math.round((totalArs - pagadoArs) * 100) / 100
      : Math.round((totalUsd - pagadoUsd) * 100) / 100;

    return jsonResponse({
      success: true,
      pagoId: pagoId,
      moneda: moneda,
      nuevoSaldo: nuevoSaldoMoneda,
      estado: (nuevoSaldoMoneda <= 0 && saldoOtraMoneda <= 0) ? "PAGADA" : "PARCIAL"
    });

  } catch(error){
    return jsonResponse({ success:false, error: error.toString() });
  } finally {
    if(lockPropio) lock.releaseLock();
  }
}

/**
 * Devuelve el historial de ingresos de mercadería, más nuevos primero.
 * Filtros opcionales por querystring: desde, hasta (yyyy-MM-dd),
 * proveedor (contiene, sin importar mayúsculas) y codigo (exacto).
 */
function getIngresosProductos(e){

  const hoja = getOrCrearHojaIngresosProductos();
  const data = hoja.getDataRange().getValues();
  let ingresos = rowsToObjects(data);

  const params = (e && e.parameter) || {};
  const desde = (params.desde || "").toString().trim();
  const hasta = (params.hasta || "").toString().trim();
  const proveedor = (params.proveedor || "").toString().trim().toLowerCase();
  const codigo = (params.codigo || "").toString().trim();

  if(desde){
    ingresos = ingresos.filter(i => String(i.FECHA) >= desde);
  }
  if(hasta){
    ingresos = ingresos.filter(i => String(i.FECHA) <= hasta);
  }
  if(proveedor){
    ingresos = ingresos.filter(i => String(i.PROVEEDOR || "").toLowerCase().indexOf(proveedor) !== -1);
  }
  if(codigo){
    ingresos = ingresos.filter(i => String(i.CODIGO || "").trim() === codigo);
  }

  ingresos.sort((a, b) => String(b.FECHA).localeCompare(String(a.FECHA)) || String(b.INGRESO_ID).localeCompare(String(a.INGRESO_ID)));

  // Google Sheets guarda "2026-08-10" como un valor de fecha real (no
  // texto) — al leerlo, Apps Script lo devuelve como objeto Date, y
  // JSON.stringify lo serializa con hora y zona incluida (ej:
  // "2026-08-10T03:00:00.000Z"). Para que en el panel se muestre solo
  // la fecha, cada FECHA se normaliza a "yyyy-MM-dd" antes de responder.
  const tz = zonaNegocio() || "GMT-3";
  ingresos.forEach(i => {
    if(i.FECHA instanceof Date){
      i.FECHA = Utilities.formatDate(i.FECHA, tz, "yyyy-MM-dd");
    } else if(i.FECHA){
      i.FECHA = String(i.FECHA).substring(0, 10);
    }
  });

  return jsonResponse({
    success: true,
    total: ingresos.length,
    ingresos: ingresos
  });
}

/* =====================================================================
   CONFIGURACIÓN DEL NEGOCIO (encabezado del ticket térmico)
   Se guarda en una hoja "CONFIGURACION" con formato clave/valor:
   CLAVE | VALOR
   Si la hoja no existe, se crea automáticamente con los valores
   por defecto (genéricos) la primera vez que se pide la configuración.
===================================================================== */

const CONFIG_NEGOCIO_DEFAULTS = {
  NOMBRE:     "MI NEGOCIO",
  SUBTITULO:  "Punto de Venta",
  DIRECCION:  "",
  TELEFONO1:  "",
  TELEFONO2:  "",
  PIE:        "¡Gracias por su compra!",

  // ---- Popup promocional del catálogo web ----
  // POPUP_IMAGEN: URL de la imagen o video (Drive o externa). Vacío = sin popup.
  // POPUP_ACTIVO: "SI" para mostrar, "NO" para ocultar sin borrar la imagen.
  // POPUP_TIPO: "imagen" o "video" — indica cómo interpretar POPUP_IMAGEN.
  POPUP_IMAGEN:  "",
  POPUP_ACTIVO:  "NO",
  POPUP_TIPO:    "imagen",
  PEDIDO_MINIMO: 100000,

  // ---- Banner y tema del catálogo web (sección "Apariencia") ----
  BANNER_TITULO:    "",
  BANNER_SUBTITULO: "Catálogo Mayorista Online",
  BANNER_IMAGEN:    "",
  TEMA:             "navy",
  // Gradiente propio del encabezado/banner — pisa la paleta del TEMA
  // de arriba cuando está activado. Pensado para que el negocio pueda
  // hacer que el encabezado combine con los colores de su propio
  // banner (imagen), en vez de quedar un choque de estilos entre el
  // navy del panel y los colores del banner subido.
  GRAD_PERSONALIZADO: "NO",
  GRAD_A:           "#241536",
  GRAD_B:           "#3a2856",
  NAVBAR_TEXTO:     "",
  // Ícono junto al nombre del negocio en el navbar del catálogo. Por
  // defecto NO es un carrito/bolsa — ya hay un botón de carrito al
  // lado, y usar el mismo tipo de ícono para las dos cosas confunde
  // al cliente (parecen dos botones de carrito).
  NAVBAR_ICONO:     "🏬",

  // ---- Tema de temporada del catálogo web (Navidad, Día del Niño…) ----
  // Se elige a mano en el panel; DESDE/HASTA (AAAA-MM-DD) son opcionales:
  // si se cargan, el tema solo se ve entre esas fechas. Los colores del
  // tema también se aplican al panel admin mientras está activo.
  TEMPORADA:           "",
  TEMPORADA_DESDE:     "",
  TEMPORADA_HASTA:     "",
  TEMPORADA_ANIMACION: "SI",
  TEMPORADA_COLORES:   "SI",
  TEMPORADA_CARTEL:    "",

  // ---- Sección "Beneficios" del catálogo web (los 4 chips bajo el banner) ----
  // Son independientes de DIRECCION/TELEFONO1/TELEFONO2 de arriba, que son
  // los del encabezado del ticket impreso — esto es lo que ve el cliente
  // en la página pública.
  BENEFICIO_WHATSAPP_NUMERO:  "",
  BENEFICIO_INSTAGRAM_URL:    "",
  BENEFICIO_TELEFONO1:        "",
  BENEFICIO_TELEFONO2:        "",
  BENEFICIO_DIRECCION:        "",
  BENEFICIO_TEXTO_LIBRE1:     "",
  BENEFICIO_TEXTO_LIBRE2:     "",

  // ---- Datos bancarios para el mensaje de WhatsApp que se manda al
  // cliente cuando un pedido queda en estado "Nuevo" (pedido en línea 2421 de admin.js) ----
  CBU_TRANSFERENCIA:  "",
  NOMBRE_TITULAR_CBU: "",

  // ---- Franja superior de mensajes rotativos (arriba de todo, sobre
  // el navbar) — un mensaje por línea, separados por salto de línea ----
  BANNER_TOP_MENSAJES: "🚚 Envíos a todo el país\n✅ Stock actualizado\n💳 Aceptamos todos los medios de pago",

  // ---- Transportes que NO se trabajan (uno por línea) — el catálogo
  // web no deja que el cliente escriba en el campo "Transporte" del
  // carrito ninguno de estos nombres, para evitar que pida un envío
  // con una empresa que el negocio no usa. La comparación ignora
  // mayúsculas/minúsculas y acentos, y es por "contiene" (si el
  // cliente escribe "Correo Argentino - Sucursal Once" y está
  // bloqueado "Correo Argentino", igual lo detecta). ----
  TRANSPORTES_NO_DISPONIBLES: "",

  // ---- Apariencia del panel admin (letra/ícono y nombre del sidebar) ----
  SIDEBAR_MARK:   "",
  SIDEBAR_TEXTO:  "",

  // ---- Carpeta de Drive donde se guardan las fotos subidas desde "Nuevo Producto" ----
  DRIVE_CARPETA_PRODUCTOS_ID: "",

  // ---- Carpeta de Drive donde se guardan los PDF generados de cada pedido ----
  DRIVE_CARPETA_PEDIDOS_ID: "",

  // ---- Textos del PDF de pedidos (el título y el WhatsApp del pie
  // toman directamente NOMBRE y BENEFICIO_WHATSAPP_NUMERO de arriba,
  // no duplican esos datos acá) ----
  PEDIDO_PDF_SUBTITULO: "",
  PEDIDO_PDF_PIE:       "Gracias por su compra.",

  // ---- URL pública del catálogo online (sin / al final). Se usa para
  // armar el link de los QR de "consultar precio" desde las etiquetas
  // del panel admin (apunta a precio.html en esa misma carpeta) ----
  URL_CATALOGO: "",

  // ---- Nombre corto del negocio, usado en la pantalla de acceso del
  // panel de Reportes (login.html) — separado de SIDEBAR_TEXTO por si
  // el negocio quiere un nombre distinto ahí ----
  NOMBRE_CORTO: "",

  // ---- Íconos usados en el catálogo web y al compartir el link
  // (redes sociales, WhatsApp, buscadores) ----
  ICONO_URL:          "icon-512.png",
  WHATSAPP_ICONO_URL: "https://cdn-icons-png.flaticon.com/512/733/733585.png",

  // ---- SEO del catálogo web (título de pestaña, meta description,
  // palabras clave, y lo que se muestra al compartir el link en
  // WhatsApp/Facebook/X) ----
  SEO_TITULO:       "",
  SEO_DESCRIPCION:  "",
  SEO_KEYWORDS:     "",

  // ---- Footer del catálogo web (pie de página con 3 columnas + copyright) ----
  FOOTER_TITULO_1:   "",
  FOOTER_TEXTO_1:    "",
  FOOTER_TEXTO_2:    "",
  FOOTER_TEXTO_3:    "",
  FOOTER_COPYRIGHT:  "",
  // Documentos A4 del depósito (cotización, boleta de entrega): ver guardarConfigDocumentos
  DOC_CUIT: "", DOC_CONDICION_IVA: "", DOC_EMAIL: "", DOC_CONDICIONES_COTIZACION: "", DOC_VALIDEZ_COTIZACION: "", DOC_LEYENDA_BOLETA: ""
};

// Temas de color válidos para el catálogo web. Si en la hoja se guarda
// un valor que no está en esta lista, el front-end usa "navy" (default).
const TEMAS_VALIDOS = ["navy", "azul", "verde", "rojo"];
// Temas de temporada (los define web/js/temporadas.js). "" = ninguno.
const TEMPORADAS_VALIDAS = ["navidad", "dianino", "diamadre", "diapadre", "halloween", "clases", "blackfriday", "sanvalentin", "pascuas", "anonuevo"];

function _temporadaDeConfig(config){
  const id = String(config.TEMPORADA || "").trim().toLowerCase();
  const fecha = v => /^\d{4}-\d{2}-\d{2}$/.test(String(v || "").trim()) ? String(v).trim() : "";
  return {
    id: TEMPORADAS_VALIDAS.indexOf(id) !== -1 ? id : "",
    desde: fecha(config.TEMPORADA_DESDE),
    hasta: fecha(config.TEMPORADA_HASTA),
    animacion: String(config.TEMPORADA_ANIMACION || "SI").toUpperCase() !== "NO",
    colores: String(config.TEMPORADA_COLORES || "SI").toUpperCase() !== "NO",
    cartel: String(config.TEMPORADA_CARTEL || "").slice(0, 140)
  };
}

/** Returns the CONFIGURACION sheet, creating it with default values if it doesn't exist yet */
function getOrCrearHojaConfiguracion(){

  const ss = getSS();
  let sheet = getHoja("CONFIGURACION");

  if(!sheet){
    sheet = ss.insertSheet("CONFIGURACION");
    sheet.appendRow(["CLAVE", "VALOR"]);
    Object.keys(CONFIG_NEGOCIO_DEFAULTS).forEach(clave => {
      sheet.appendRow([clave, CONFIG_NEGOCIO_DEFAULTS[clave]]);
    });
  }

  return sheet;

}

/**
 * Extrae el ID de carpeta de Drive a partir de lo que el usuario pegó en
 * Configuración: puede ser el ID solo, o el link completo de la carpeta
 * (https://drive.google.com/drive/folders/XXXXX...). Así no hay que
 * explicarle a nadie cómo "sacar el ID", alcanza con copiar el link.
 */
function extraerIdCarpetaDrive(valor){
  const texto = String(valor || "").trim();
  if(!texto) return "";

  const match = texto.match(/folders\/([a-zA-Z0-9_-]+)/);
  if(match) return match[1];

  // Si no es un link reconocible, asumimos que ya es el ID directamente
  return texto;
}

/**
 * Lee el valor de una clave puntual de la hoja CONFIGURACION, sin pasar
 * por getConfiguracionNegocio() (que devuelve una respuesta HTTP
 * completa). Útil cuando una función de backend solo necesita un dato
 * suelto (ej. el ID de una carpeta de Drive) antes de hacer su trabajo.
 */
function leerClaveConfiguracion(clave, valorPorDefecto){

  const sheet = getOrCrearHojaConfiguracion();
  const filas = sheet.getDataRange().getValues();
  filas.shift(); // encabezados CLAVE | VALOR

  let valor = "";
  filas.forEach(row => {
    if(String(row[0] || "").trim().toUpperCase() === clave.toUpperCase()){
      valor = row[1] === undefined || row[1] === null ? "" : String(row[1]);
    }
  });

  return valor || (valorPorDefecto || "");
}

/**
 * Escribe (o actualiza) una clave puntual de CONFIGURACION, sin tocar
 * las demás filas — a diferencia de guardarConfiguracionNegocio, que
 * reescribe varias claves juntas a la vez con su propio formulario.
 * Si la clave no existe todavía, se agrega como fila nueva.
 */
function fijarClaveConfiguracion(clave, valor){

  const sheet = getOrCrearHojaConfiguracion();
  const filas = sheet.getDataRange().getValues();

  for(let i = 1; i < filas.length; i++){
    if(String(filas[i][0] || "").trim().toUpperCase() === clave.toUpperCase()){
      sheet.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }

  sheet.appendRow([clave, valor]);
}

/* =====================================================================
   CARPETAS DE DRIVE — se crean solas en una instalación nueva
   Estructura (al lado de la planilla):
     VeekPOS - <negocio>/
       Fotos de productos              → DRIVE_CARPETA_PRODUCTOS_ID
       PDF de pedidos                  → DRIVE_CARPETA_PEDIDOS_ID
       Catálogo PDF                    → DRIVE_CARPETA_CATALOGO_ID
   Solo completa las claves que estén vacías (o apunten a una carpeta que
   ya no existe); nunca cambia una carpeta que el negocio ya configuró.
   Si se ejecuta dos veces, reutiliza las carpetas por nombre (no duplica).
===================================================================== */
const CARPETAS_DRIVE_VEEKPOS = [
  ["DRIVE_CARPETA_PRODUCTOS_ID",      "Fotos de productos"],
  ["DRIVE_CARPETA_PEDIDOS_ID",        "PDF de pedidos"],
  ["DRIVE_CARPETA_CATALOGO_ID",       "Catálogo PDF"]
];

function _buscarOCrearCarpeta(padre, nombre){
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

/** Ejecutable a mano desde el editor también (por ej. en un negocio que ya funcionaba y quiere las carpetas armadas). */
function crearCarpetasDriveFaltantes(){
  const faltan = CARPETAS_DRIVE_VEEKPOS.filter(par => {
    const id = extraerIdCarpetaDrive(leerClaveConfiguracion(par[0], ""));
    if(!id) return true;
    try{ DriveApp.getFolderById(id); return false; }catch(e){ return true; }
  });
  if(!faltan.length) return { creadas: [], carpeta: "" , message: "Las carpetas de Drive ya estaban configuradas." };

  let padre = null;
  try{
    const padres = DriveApp.getFileById(getSS().getId()).getParents();
    if(padres.hasNext()) padre = padres.next();
  }catch(e){}
  if(!padre) padre = DriveApp.getRootFolder();
  const negocio = String(leerClaveConfiguracion("NOMBRE", "") || "").trim();
  const raiz = _buscarOCrearCarpeta(padre, "VeekPOS" + (negocio ? " - " + negocio : ""));

  const creadas = [];
  faltan.forEach(par => {
    const carpeta = _buscarOCrearCarpeta(raiz, par[1]);
    fijarClaveConfiguracion(par[0], carpeta.getUrl());
    creadas.push(par[1]);
  });
  // la configuración del panel está cacheada: que muestre los links nuevos
  try{ CacheService.getScriptCache().removeAll(["config_negocio:n","config_negocio:0","config_negocio:1"]); }catch(e){}
  const res = { creadas: creadas, carpeta: raiz.getUrl(), message: "Carpetas de Drive creadas en \"" + raiz.getName() + "\": " + creadas.join(", ") + "." };
  Logger.log(res.message + " " + res.carpeta);
  return res;
}

/**
 * Recibe una imagen en base64 (data.imagenBase64), la guarda en la
 * carpeta de Drive configurada en CONFIGURACION (clave
 * DRIVE_CARPETA_PRODUCTOS_ID), la hace pública por link, y devuelve la
 * URL pública lista para guardar en la columna IMAGEN de PRODUCTOS.
 *
 * data esperado: { action, imagenBase64, nombreArchivo, tipoMime, codigoProducto }
 */
function subirImagenProducto(data){

  const carpetaId = extraerIdCarpetaDrive(leerClaveConfiguracion("DRIVE_CARPETA_PRODUCTOS_ID"));

  if(!carpetaId){
    return jsonResponse({
      success:false,
      message:"Todavía no configuraste la carpeta de Drive para fotos de productos. Entrá a Configuración → Imágenes de productos y pegá el link de la carpeta."
    });
  }

  let carpeta;
  try{
    carpeta = DriveApp.getFolderById(carpetaId);
  }catch(error){
    return jsonResponse({
      success:false,
      message:"No se pudo abrir la carpeta de Drive configurada. Revisá que el link/ID en Configuración sea correcto y que esta cuenta tenga acceso."
    });
  }

  const base64Limpio = String(data.imagenBase64 || "").replace(/^data:[^,]+,/, "");
  if(!base64Limpio){
    return jsonResponse({ success:false, message:"No llegó ninguna imagen para subir." });
  }

  const tipoMime = data.tipoMime || "image/jpeg";
  const nombreBase = (data.codigoProducto ? String(data.codigoProducto).trim() + "_" : "") + "producto";
  const extension = tipoMime.indexOf("png") !== -1 ? "png" : (tipoMime.indexOf("webp") !== -1 ? "webp" : "jpg");
  const nombreArchivo = nombreBase + "_" + new Date().getTime() + "." + extension;

  try{

    const bytes = Utilities.base64Decode(base64Limpio);
    const blob = Utilities.newBlob(bytes, tipoMime, nombreArchivo);

    const archivo = carpeta.createFile(blob);

    // Hace el archivo accesible por link para cualquiera (solo lectura)
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const urlPublica = "https://lh3.googleusercontent.com/d/" + archivo.getId();

    return jsonResponse({
      success: true,
      url: urlPublica,
      archivoId: archivo.getId()
    });

  }catch(error){
    return jsonResponse({
      success:false,
      message:"Error al subir la imagen a Drive: " + error.toString()
    });
  }
}

/**
 * Descarga una imagen del lado del servidor (Apps Script, no del
 * navegador del cliente) y la devuelve en base64. Existe porque el
 * catálogo público necesita insertar las fotos de los productos en un
 * PDF generado en el navegador (jsPDF), y eso requiere leer los
 * píxeles de la imagen — algo que el navegador bloquea si el servidor
 * de la imagen (Drive) no responde con headers CORS, o si el
 * navegador ya cacheó esa misma imagen sin esos headers (porque la
 * mostró antes en una tarjeta normal). Acá no hay ese problema: Apps
 * Script no es un navegador, no aplica CORS.
 *
 * Parámetro esperado: ?action=imagenProxy&url=<URL de la imagen>
 */
// Dominios desde los que se permite descargar imágenes. Sin esta lista,
// cualquiera podía usar el servidor para pedir cualquier URL (SSRF / proxy
// abierto). Si guardás fotos de productos en otro sitio (ej. imgur), sumá su
// dominio acá, en minúsculas.
const IMAGEN_PROXY_DOMINIOS = [
  "lh3.googleusercontent.com",
  "drive.google.com",
  "drive.usercontent.google.com",
  "docs.google.com"
];
const IMAGEN_PROXY_MAX_BYTES = 6 * 1024 * 1024;

function urlDeImagenPermitida(url){
  const m = /^https:\/\/([a-z0-9.-]+)(?::443)?(?:[\/?#]|$)/i.exec(String(url || ""));
  if(!m) return false;                       // solo https, sin usuario@ ni puertos raros
  const host = m[1].toLowerCase();
  return IMAGEN_PROXY_DOMINIOS.indexOf(host) !== -1;
}

function getImagenProxy(e){

  const url = (e && e.parameter && e.parameter.url || "").toString().trim();

  if(!url){
    return jsonResponse({ success:false, message:"Falta la URL de la imagen." });
  }

  if(!urlDeImagenPermitida(url)){
    return jsonResponse({ success:false, message:"Dominio de imagen no permitido." });
  }

  if(!permitirIntento("imgproxy", 300, 600)){
    return jsonResponse({ success:false, message:"Demasiadas descargas seguidas. Probá en unos minutos." });
  }

  try{

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if(response.getResponseCode() !== 200){
      return jsonResponse({ success:false, message:"No se pudo descargar la imagen (código " + response.getResponseCode() + ")." });
    }

    const blob = response.getBlob();
    if(blob.getBytes().length > IMAGEN_PROXY_MAX_BYTES){
      return jsonResponse({ success:false, message:"La imagen es demasiado grande." });
    }
    let tipoMime = blob.getContentType() || "image/jpeg";

    // Algunos servidores devuelven el content-type con parámetros extra
    // (ej. "image/jpeg; charset=utf-8"), que rompen el formato de un
    // data URL válido si se los deja tal cual. Se queda solo con el
    // tipo MIME limpio (ej. "image/jpeg").
    tipoMime = tipoMime.split(";")[0].trim();
    if(!tipoMime.startsWith("image/")) tipoMime = "image/jpeg";

    const base64 = Utilities.base64Encode(blob.getBytes());

    return jsonResponse({
      success: true,
      dataUrl: "data:" + tipoMime + ";base64," + base64
    });

  }catch(error){
    return jsonResponse({
      success:false,
      message:"Error al descargar la imagen: " + error.toString()
    });
  }
}

/* =========================================================
   CATÁLOGO PDF
   -----------------------------------------------------------
   El PDF lo arma la app de escritorio (desktop/src/catalogo-pdf.js):
   portada con condiciones e índice, productos agrupados por categoría,
   fotos sin deformar, links "Pedir" a cada producto. Se genera solo una
   vez por día al abrir la caja principal, o con "Generar y publicar" en
   Configuración → Catálogo web.

   Acá solo se RECIBE: la app lo manda en partes (base64, de a ~1,5 MB
   para no chocar con el límite de tamaño de una solicitud), se arma el
   archivo en Drive, se publica por link y se guarda en CONFIGURACION
   (CATALOGO_PDF_URL / _FILE_ID / _FECHA / _TOTAL_PRODUCTOS). El botón
   "Descargar Catálogo PDF" del sitio lee eso con ?action=catalogoPDFInfo.

   Orden "publicar primero, borrar después": el PDF anterior se borra
   recién cuando el nuevo ya está creado y guardado en CONFIGURACION, así
   un error a mitad de camino nunca deja el link roto.

   Antes el PDF se armaba acá mismo como un documento de Google, por
   lotes y con triggers diarios. _retirarCatalogoPDFViejo (más abajo)
   borra esos triggers y restos de las instalaciones que lo tenían.
========================================================= */

const PREFIJO_PARTE_CATALOGO = "catalogo_parte_";
const PARTES_MAX_CATALOGO = 60;          // 60 × 1,5 MB = 90 MB: de sobra
const PROP_CATALOGO_V2 = "VEEK_CATALOGO_PDF_V2";

/** Carpeta de Drive del catálogo PDF: DRIVE_CARPETA_CATALOGO_ID, o si no hay, la de los PDF de pedidos */
function getCarpetaCatalogoId(){
  const propia = extraerIdCarpetaDrive(
    leerClaveConfiguracion("DRIVE_CARPETA_CATALOGO_ID", "")
  );
  if(propia) return propia;

  return extraerIdCarpetaDrive(
    leerClaveConfiguracion("DRIVE_CARPETA_PEDIDOS_ID", "")
  );
}

/**
 * Borra un archivo de Drive de forma definitiva (sin pasar por la
 * papelera) con el servicio avanzado de Drive si está habilitado; si no,
 * lo manda a la papelera.
 */
function eliminarArchivoDriveDefinitivo(archivo){
  try{
    // v3 del servicio avanzado de Drive usa "delete"; v2 usa "remove".
    if(typeof Drive.Files.delete === "function"){
      Drive.Files.delete(archivo.getId());
    } else if(typeof Drive.Files.remove === "function"){
      Drive.Files.remove(archivo.getId());
    } else {
      archivo.setTrashed(true);
    }
  }catch(e){
    archivo.setTrashed(true);
  }
}

/**
 * Link del catálogo PDF ya publicado (lectura rápida de CONFIGURACION)
 * — lo usa el botón "Descargar Catálogo PDF" del sitio y la tarjeta de
 * Configuración del panel.
 */
function getCatalogoPDFInfo(){

  const url = leerClaveConfiguracion("CATALOGO_PDF_URL", "");
  const fecha = leerClaveConfiguracion("CATALOGO_PDF_FECHA", "");
  const totalProductos = leerClaveConfiguracion("CATALOGO_PDF_TOTAL_PRODUCTOS", "");

  if(!url){
    return jsonResponse({
      success: false,
      message: "El catálogo en PDF todavía no se generó. Se arma solo una vez por día, cuando se abre la caja principal."
    });
  }

  return jsonResponse({
    success: true,
    url: url,
    fecha: fecha,
    totalProductos: totalProductos
  });
}

/**
 * Recibe una parte del PDF (POST, con sesión). Datos:
 *   subidaId (lo genera la app, igual para todas las partes),
 *   parte (0…n-1), totalPartes, datos (base64), totalProductos.
 * Cada parte queda como archivo temporal en la carpeta del catálogo; con
 * la última se arma el PDF, se publica y se borran las partes y el PDF
 * anterior. Reenviar una parte (reintento) la reemplaza, no la duplica.
 */
function subirCatalogoPDF(data){

  _retirarCatalogoPDFViejo();

  const subidaId = String((data && data.subidaId) || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 40);
  const parte = Number(data && data.parte);
  const totalPartes = Number(data && data.totalPartes);
  const datos = String((data && data.datos) || "");

  if(!subidaId || !(totalPartes >= 1) || totalPartes > PARTES_MAX_CATALOGO || !(parte >= 0) || parte >= totalPartes || !datos){
    return jsonResponse({ success:false, message:"Parte del catálogo inválida." });
  }
  if(!/^[A-Za-z0-9+/=]+$/.test(datos)){
    return jsonResponse({ success:false, message:"La parte del catálogo no está en base64." });
  }

  let carpeta;
  try{
    carpeta = DriveApp.getFolderById(getCarpetaCatalogoId());
  }catch(e){
    return jsonResponse({ success:false, message:"No se encontró la carpeta del catálogo en Drive (Configuración → DRIVE_CARPETA_CATALOGO_ID o la de PDF de pedidos)." });
  }

  const nombreParte = PREFIJO_PARTE_CATALOGO + subidaId + "_" + ("000" + parte).slice(-3);
  const repetidas = carpeta.getFilesByName(nombreParte);
  while(repetidas.hasNext()) eliminarArchivoDriveDefinitivo(repetidas.next());
  carpeta.createFile(nombreParte, datos, MimeType.PLAIN_TEXT);

  if(parte < totalPartes - 1){
    return jsonResponse({ success:true, parte: parte, recibida: true });
  }

  // Última parte: armar el PDF con todas, en orden
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return jsonResponse({ success:false, message:"El sistema está ocupado, se reintenta solo." });
  try{
    const partes = [];
    for(let i = 0; i < totalPartes; i++){
      const it = carpeta.getFilesByName(PREFIJO_PARTE_CATALOGO + subidaId + "_" + ("000" + i).slice(-3));
      if(!it.hasNext()) return jsonResponse({ success:false, faltaParte: i, message:"Falta la parte " + (i + 1) + " de " + totalPartes + " del catálogo." });
      partes.push(it.next());
    }
    const base64 = partes.map(f => f.getBlob().getDataAsString()).join("");
    const pdf = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.PDF, "catalogo.pdf");
    if(pdf.getBytes().slice(0, 4).map(b => String.fromCharCode(b)).join("") !== "%PDF"){
      partes.forEach(eliminarArchivoDriveDefinitivo);
      return jsonResponse({ success:false, message:"El archivo recibido no es un PDF." });
    }

    const idAnterior = leerClaveConfiguracion("CATALOGO_PDF_FILE_ID", "");
    const archivo = carpeta.createFile(pdf);
    try{
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }catch(eShare){
      // Política del dominio que no deja compartir por link: el archivo
      // existe igual, pero puede no descargarse sin iniciar sesión.
      console.warn("Catálogo PDF: no se pudo hacer público por link: " + eShare);
    }
    const url = "https://drive.google.com/uc?export=download&id=" + archivo.getId();
    fijarClaveConfiguracion("CATALOGO_PDF_URL", url);
    fijarClaveConfiguracion("CATALOGO_PDF_FILE_ID", archivo.getId());
    fijarClaveConfiguracion("CATALOGO_PDF_FECHA", new Date().toISOString());
    fijarClaveConfiguracion("CATALOGO_PDF_TOTAL_PRODUCTOS", String(Number(data.totalProductos) || 0));

    // Recién ahora: borrar las partes y el PDF anterior
    partes.forEach(f => { try{ eliminarArchivoDriveDefinitivo(f); }catch(e){} });
    if(idAnterior && idAnterior !== archivo.getId()){
      try{ eliminarArchivoDriveDefinitivo(DriveApp.getFileById(idAnterior)); }catch(e){ /* ya no existía */ }
    }
    // Partes sueltas de subidas que quedaron por la mitad (más de un día)
    try{
      const viejas = carpeta.searchFiles("title contains '" + PREFIJO_PARTE_CATALOGO + "'");
      const limite = Date.now() - 24 * 3600 * 1000;
      while(viejas.hasNext()){
        const f = viejas.next();
        if(f.getDateCreated().getTime() < limite) eliminarArchivoDriveDefinitivo(f);
      }
    }catch(e){}

    return jsonResponse({ success:true, publicado:true, url: url, fileId: archivo.getId(), bytes: pdf.getBytes().length });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Una sola vez por instalación: borra lo que dejaba el armado viejo del
 * catálogo (triggers diarios y de continuación, progreso guardado y el
 * documento temporal a medio hacer). Sin esto, los triggers viejos
 * seguirían disparándose todos los días contra funciones que ya no existen.
 */
const HANDLERS_CATALOGO_VIEJO = ["generarCatalogoPDFPorLotes", "continuarCatalogoPDFPorLotes", "vigilarCatalogoPDFDesactualizado", "generarCatalogoPDFDiario"];
function _retirarCatalogoPDFViejo(){
  try{
    if(_props()[PROP_CATALOGO_V2]) return;
    ScriptApp.getProjectTriggers().forEach(t => {
      if(HANDLERS_CATALOGO_VIEJO.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
    });
    const props = PropertiesService.getScriptProperties();
    const progreso = props.getProperty("CATALOGO_PDF_PROGRESO");
    if(progreso){
      try{ const p = JSON.parse(progreso); if(p && p.docId) DriveApp.getFileById(p.docId).setTrashed(true); }catch(e){}
      props.deleteProperty("CATALOGO_PDF_PROGRESO");
    }
    props.setProperty(PROP_CATALOGO_V2, new Date().toISOString());
    if(_propsMemo) _propsMemo[PROP_CATALOGO_V2] = "1";
  }catch(e){
    console.warn("No se pudieron retirar los triggers del catálogo viejo: " + e);
  }
}

// Si un trigger viejo llega a dispararse antes de que se retire (por
// ejemplo, a la madrugada justo después de actualizar el código), estas
// funciones existen para que no falle con "función no encontrada": lo
// único que hacen es retirar los triggers viejos.
function generarCatalogoPDFPorLotes(){ _retirarCatalogoPDFViejo(); }
function continuarCatalogoPDFPorLotes(){ _retirarCatalogoPDFViejo(); }
function vigilarCatalogoPDFDesactualizado(){ _retirarCatalogoPDFViejo(); }
function generarCatalogoPDFDiario(){ _retirarCatalogoPDFViejo(); }

/**
 * Devuelve SOLO los datos públicos que necesita la landing tipo
 * Linktree (WhatsApp, redes sociales, dirección para el mapa,
 * catálogo y tema de color) leídos de la hoja CONFIGURACION.
 *
 * A propósito es un endpoint separado de getConfiguracionNegocio(): la
 * landing es pública (la ve cualquiera sin login), así que acá NUNCA
 * se exponen datos sensibles del panel admin (CBU, IDs de carpetas de
 * Drive, etc.) — solo lo estrictamente necesario para los botones, el
 * modal del mapa y el color de la página.
 */
function getConfigLanding(){

  const sheet = getOrCrearHojaConfiguracion();
  const data = sheet.getDataRange().getValues();
  data.shift(); // encabezados CLAVE | VALOR

  const config = { ...CONFIG_NEGOCIO_DEFAULTS };

  data.forEach(row => {
    const clave = String(row[0] || "").trim().toUpperCase();
    if(clave && config.hasOwnProperty(clave)){
      config[clave] = row[1] === undefined || row[1] === null ? "" : String(row[1]);
    }
  });

  // Número de WhatsApp: se limpia a solo dígitos para armar el link
  // wa.me/<numero> directo, sin importar cómo lo haya cargado el
  // negocio en la hoja (con espacios, guiones, "+", etc.)
  const whatsappNumero = String(config.BENEFICIO_WHATSAPP_NUMERO || "").replace(/\D/g, "");

  // Tema de color: mismo criterio que getConfiguracionNegocio — si lo
  // guardado no es un tema válido, se cae a "navy" para no romper el
  // front-end.
  const temaGuardado = String(config.TEMA || "").trim().toLowerCase();
  const temaFinal = TEMAS_VALIDOS.indexOf(temaGuardado) !== -1 ? temaGuardado : "navy";

  return jsonResponse({
    success:true,
    config: {
      nombre:   config.NOMBRE,
      iconoUrl: config.ICONO_URL,

      tema: temaFinal,
      temporada: _temporadaDeConfig(config),
      gradPersonalizado: String(config.GRAD_PERSONALIZADO || "NO").toUpperCase() === "SI",
      gradA: config.GRAD_A,
      gradB: config.GRAD_B,

      whatsappNumero: whatsappNumero,
      // En vez de asumir qué red social vive en qué columna (algunos
      // negocios usan BENEFICIO_INSTAGRAM_URL para Instagram y los de
      // texto libre para Facebook/TikTok, otros al revés), se manda
      // la lista completa de URLs cargadas y el front-end detecta solo
      // de qué red es cada una mirando el dominio.
      enlaces: [
        config.BENEFICIO_INSTAGRAM_URL,
        config.BENEFICIO_TEXTO_LIBRE1,
        config.BENEFICIO_TEXTO_LIBRE2
      ].filter(Boolean),

      direccion:  config.BENEFICIO_DIRECCION || config.DIRECCION,

      urlCatalogo:     config.URL_CATALOGO,
      footerCopyright: config.FOOTER_COPYRIGHT
    }
  });
}

/**
 * Devuelve la configuración del negocio (nombre, subtítulo, dirección,
 * teléfonos, pie de página) leída de la hoja CONFIGURACION. Cualquier
 * clave que falte en la hoja se completa con su valor por defecto.
 */
function getConfiguracionNegocio(){

  const sheet = getOrCrearHojaConfiguracion();
  const data = sheet.getDataRange().getValues();
  data.shift(); // encabezados CLAVE | VALOR

  const config = { ...CONFIG_NEGOCIO_DEFAULTS };

  data.forEach(row => {
    const clave = String(row[0] || "").trim().toUpperCase();
    if(clave && config.hasOwnProperty(clave)){
      config[clave] = row[1] === undefined || row[1] === null ? "" : String(row[1]);
    }
  });

  // El tema se valida: si lo que está guardado en la hoja no es uno
  // de los temas conocidos, se usa "navy" para no romper el front-end.
  const temaGuardado = String(config.TEMA || "").trim().toLowerCase();
  const temaFinal = TEMAS_VALIDOS.indexOf(temaGuardado) !== -1
    ? temaGuardado
    : "navy";

  return jsonResponse({
    success:true,
    config: {
      nombre:    config.NOMBRE,
      subtitulo: config.SUBTITULO,
      direccion: config.DIRECCION,
      telefono1: config.TELEFONO1,
      telefono2: config.TELEFONO2,
      pie:       config.PIE,

      bannerTitulo:    config.BANNER_TITULO,
      bannerSubtitulo: config.BANNER_SUBTITULO,
      bannerImagen:    config.BANNER_IMAGEN,
      tema:            temaFinal,
      temasDisponibles: TEMAS_VALIDOS,
      temporada:       _temporadaDeConfig(config),
      gradPersonalizado: String(config.GRAD_PERSONALIZADO || "NO").toUpperCase() === "SI",
      gradA: config.GRAD_A,
      gradB: config.GRAD_B,
      navbarTexto:     config.NAVBAR_TEXTO,
      navbarIcono:     config.NAVBAR_ICONO,

      popupImagen:     config.POPUP_IMAGEN,
      popupActivo:     String(config.POPUP_ACTIVO || "NO").toUpperCase() === "SI",
      popupTipo:       String(config.POPUP_TIPO || "imagen").toLowerCase() === "video" ? "video" : "imagen",
      pedidoMinimo:    (function(v){ const n = Number(v); return (v === "" || v === null || v === undefined || isNaN(n) || n < 0) ? 100000 : n; })(config.PEDIDO_MINIMO),
      zonaHoraria:     zonaNegocio(),

      beneficioWhatsappNumero: config.BENEFICIO_WHATSAPP_NUMERO,
      beneficioInstagramUrl:   config.BENEFICIO_INSTAGRAM_URL,
      beneficioTelefono1:      config.BENEFICIO_TELEFONO1,
      beneficioTelefono2:      config.BENEFICIO_TELEFONO2,
      beneficioDireccion:      config.BENEFICIO_DIRECCION,
      beneficioTextoLibre1:    config.BENEFICIO_TEXTO_LIBRE1,
      beneficioTextoLibre2:    config.BENEFICIO_TEXTO_LIBRE2,

      cbuTransferencia:  config.CBU_TRANSFERENCIA,
      nombreTitularCbu:  config.NOMBRE_TITULAR_CBU,

      bannerTopMensajes: config.BANNER_TOP_MENSAJES,
      transportesNoDisponibles: config.TRANSPORTES_NO_DISPONIBLES,

      sidebarMark:  config.SIDEBAR_MARK,
      sidebarTexto: config.SIDEBAR_TEXTO,

      driveCarpetaProductosId: config.DRIVE_CARPETA_PRODUCTOS_ID,
      driveCarpetaPedidosId:   config.DRIVE_CARPETA_PEDIDOS_ID,

      pedidoPdfSubtitulo: config.PEDIDO_PDF_SUBTITULO,
      pedidoPdfPie:       config.PEDIDO_PDF_PIE,

      urlCatalogo: config.URL_CATALOGO,
      nombreCorto: config.NOMBRE_CORTO,

      iconoUrl: config.ICONO_URL,
      whatsappIconoUrl: config.WHATSAPP_ICONO_URL,

      seoTitulo: config.SEO_TITULO,
      seoDescripcion: config.SEO_DESCRIPCION,
      seoKeywords: config.SEO_KEYWORDS,

      footerTitulo1: config.FOOTER_TITULO_1,
      footerTexto1: config.FOOTER_TEXTO_1,
      footerTexto2: config.FOOTER_TEXTO_2,
      footerTexto3: config.FOOTER_TEXTO_3,
      footerCopyright: config.FOOTER_COPYRIGHT,

      modoNegocio: modoNegocio(),
      depositoVinculado: (function(){ try{ return !!PropertiesService.getScriptProperties().getProperty("DEPOSITO_URL"); }catch(e){ return false; } })(),
      documentos: {
        cuit: config.DOC_CUIT || "", condicionIva: config.DOC_CONDICION_IVA || "", email: config.DOC_EMAIL || "",
        condicionesCotizacion: config.DOC_CONDICIONES_COTIZACION || "", validezCotizacion: Number(config.DOC_VALIDEZ_COTIZACION) || 7,
        leyendaBoleta: config.DOC_LEYENDA_BOLETA || ""
      },
      dolarDeposito: (function(){ try{ return leerDolarDeposito(); }catch(e){ return null; } })()
    }
  });

}

/**
 * Guarda la configuración del negocio en la hoja CONFIGURACION.
 * Actualiza las filas existentes por clave; si una clave no existe
 * todavía en la hoja, la agrega.
 */
function guardarConfiguracionNegocio(e){

  const params = (e && e.parameter) || {};

  // Cargar config actual primero para poder usar valores guardados como fallback
  const sheetActual = getOrCrearHojaConfiguracion();
  const dataActual = sheetActual.getDataRange().getValues();
  dataActual.shift();

  const configGuardada = {};
  dataActual.forEach(row => {
    const clave = String(row[0] || "").trim().toUpperCase();
    if(clave) configGuardada[clave] = row[1] === undefined || row[1] === null ? "" : String(row[1]);
  });

  const nombre = (params.nombre || configGuardada.NOMBRE || "").toString().trim();
  if(!nombre){
    return jsonResponse({
      success:false,
      message:"El nombre del local es obligatorio."
    });
  }

  /**
   * Devuelve el valor nuevo para `claveParam` (nombre del parámetro que
   * manda el formulario): si vino en la request (aunque sea vacío), se
   * usa ese; si no vino (undefined), se conserva el valor ya guardado.
   */
  function valorOPreservado(claveParam, claveSheet, porDefecto){
    if(params[claveParam] !== undefined){
      return params[claveParam].toString().trim();
    }
    return configGuardada.hasOwnProperty(claveSheet)
      ? configGuardada[claveSheet]
      : (porDefecto || "");
  }

  // Tema: si llega y es válido, se usa. Si llega inválido o no llega,
  // se mantiene el que ya estaba guardado (nunca se guarda un tema inválido).
  let temaNuevo = (params.tema || "").toString().trim().toLowerCase();
  if(TEMAS_VALIDOS.indexOf(temaNuevo) === -1){
    temaNuevo = configGuardada.TEMA || "navy";
    if(TEMAS_VALIDOS.indexOf(temaNuevo) === -1){
      temaNuevo = "navy";
    }
  }

  const nuevosValores = {
    NOMBRE:    nombre,
    SUBTITULO: valorOPreservado("subtitulo", "SUBTITULO"),
    DIRECCION: valorOPreservado("direccion", "DIRECCION"),
    TELEFONO1: valorOPreservado("telefono1", "TELEFONO1"),
    TELEFONO2: valorOPreservado("telefono2", "TELEFONO2"),
    PIE:       valorOPreservado("pie", "PIE"),

    BANNER_TITULO:    valorOPreservado("bannerTitulo", "BANNER_TITULO"),
    BANNER_SUBTITULO: valorOPreservado("bannerSubtitulo", "BANNER_SUBTITULO"),
    BANNER_IMAGEN:    valorOPreservado("bannerImagen", "BANNER_IMAGEN"),
    TEMA:             temaNuevo,
    GRAD_PERSONALIZADO: params.gradPersonalizado !== undefined
      ? (params.gradPersonalizado === "1" || params.gradPersonalizado === "true" ? "SI" : "NO")
      : (configGuardada.GRAD_PERSONALIZADO || "NO"),
    GRAD_A: valorOPreservado("gradA", "GRAD_A", "#241536"),
    GRAD_B: valorOPreservado("gradB", "GRAD_B", "#3a2856"),
    NAVBAR_TEXTO:     valorOPreservado("navbarTexto", "NAVBAR_TEXTO", ""),
    NAVBAR_ICONO:     valorOPreservado("navbarIcono", "NAVBAR_ICONO", "🏬"),

    BENEFICIO_WHATSAPP_NUMERO: valorOPreservado("beneficioWhatsappNumero", "BENEFICIO_WHATSAPP_NUMERO", ""),
    BENEFICIO_INSTAGRAM_URL:   valorOPreservado("beneficioInstagramUrl", "BENEFICIO_INSTAGRAM_URL"),
    BENEFICIO_TELEFONO1:       valorOPreservado("beneficioTelefono1", "BENEFICIO_TELEFONO1"),
    BENEFICIO_TELEFONO2:       valorOPreservado("beneficioTelefono2", "BENEFICIO_TELEFONO2"),
    BENEFICIO_DIRECCION:       valorOPreservado("beneficioDireccion", "BENEFICIO_DIRECCION"),
    BENEFICIO_TEXTO_LIBRE1:    valorOPreservado("beneficioTextoLibre1", "BENEFICIO_TEXTO_LIBRE1"),
    BENEFICIO_TEXTO_LIBRE2:    valorOPreservado("beneficioTextoLibre2", "BENEFICIO_TEXTO_LIBRE2"),

    CBU_TRANSFERENCIA:  valorOPreservado("cbuTransferencia", "CBU_TRANSFERENCIA", ""),
    NOMBRE_TITULAR_CBU: valorOPreservado("nombreTitularCbu", "NOMBRE_TITULAR_CBU", ""),

    BANNER_TOP_MENSAJES: valorOPreservado("bannerTopMensajes", "BANNER_TOP_MENSAJES"),
    TRANSPORTES_NO_DISPONIBLES: valorOPreservado("transportesNoDisponibles", "TRANSPORTES_NO_DISPONIBLES"),

    SIDEBAR_MARK:  valorOPreservado("sidebarMark", "SIDEBAR_MARK", ""),
    SIDEBAR_TEXTO: valorOPreservado("sidebarTexto", "SIDEBAR_TEXTO", ""),

    DRIVE_CARPETA_PRODUCTOS_ID: valorOPreservado("driveCarpetaProductosId", "DRIVE_CARPETA_PRODUCTOS_ID"),
    DRIVE_CARPETA_PEDIDOS_ID:   valorOPreservado("driveCarpetaPedidosId", "DRIVE_CARPETA_PEDIDOS_ID", ""),

    PEDIDO_PDF_SUBTITULO: valorOPreservado("pedidoPdfSubtitulo", "PEDIDO_PDF_SUBTITULO", ""),
    PEDIDO_PDF_PIE:       valorOPreservado("pedidoPdfPie", "PEDIDO_PDF_PIE", "Gracias por su compra."),

    URL_CATALOGO: valorOPreservado("urlCatalogo", "URL_CATALOGO", ""),
    NOMBRE_CORTO: valorOPreservado("nombreCorto", "NOMBRE_CORTO", ""),

    ICONO_URL:          valorOPreservado("iconoUrl", "ICONO_URL", "icon-512.png"),
    WHATSAPP_ICONO_URL: valorOPreservado("whatsappIconoUrl", "WHATSAPP_ICONO_URL", "https://cdn-icons-png.flaticon.com/512/733/733585.png"),

    SEO_TITULO:      valorOPreservado("seoTitulo", "SEO_TITULO"),
    SEO_DESCRIPCION: valorOPreservado("seoDescripcion", "SEO_DESCRIPCION"),
    SEO_KEYWORDS:    valorOPreservado("seoKeywords", "SEO_KEYWORDS"),

    FOOTER_TITULO_1:  valorOPreservado("footerTitulo1", "FOOTER_TITULO_1"),
    FOOTER_TEXTO_1:   valorOPreservado("footerTexto1", "FOOTER_TEXTO_1"),
    FOOTER_TEXTO_2:   valorOPreservado("footerTexto2", "FOOTER_TEXTO_2"),
    FOOTER_TEXTO_3:   valorOPreservado("footerTexto3", "FOOTER_TEXTO_3"),
    FOOTER_COPYRIGHT: valorOPreservado("footerCopyright", "FOOTER_COPYRIGHT"),

    POPUP_IMAGEN: valorOPreservado("popupImagen", "POPUP_IMAGEN", ""),
    POPUP_ACTIVO: valorOPreservado("popupActivo", "POPUP_ACTIVO", "NO"),
    POPUP_TIPO:   (function(){
      const v = valorOPreservado("popupTipo", "POPUP_TIPO", "imagen").toLowerCase();
      return v === "video" ? "video" : "imagen";
    })(),
    PEDIDO_MINIMO: valorOPreservado("pedidoMinimo", "PEDIDO_MINIMO", "100000"),

    TEMPORADA: (function(){
      const v = valorOPreservado("temporada", "TEMPORADA", "").toLowerCase();
      return TEMPORADAS_VALIDAS.indexOf(v) !== -1 ? v : "";
    })(),
    TEMPORADA_DESDE: (function(){ const v = valorOPreservado("temporadaDesde", "TEMPORADA_DESDE", ""); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ""; })(),
    TEMPORADA_HASTA: (function(){ const v = valorOPreservado("temporadaHasta", "TEMPORADA_HASTA", ""); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ""; })(),
    TEMPORADA_ANIMACION: (function(){ const v = valorOPreservado("temporadaAnimacion", "TEMPORADA_ANIMACION", "SI").toUpperCase(); return v === "NO" || v === "FALSE" || v === "0" ? "NO" : "SI"; })(),
    TEMPORADA_COLORES: (function(){ const v = valorOPreservado("temporadaColores", "TEMPORADA_COLORES", "SI").toUpperCase(); return v === "NO" || v === "FALSE" || v === "0" ? "NO" : "SI"; })(),
    TEMPORADA_CARTEL: textoSeguroHoja(valorOPreservado("temporadaCartel", "TEMPORADA_CARTEL", ""), 140)
  };
  if(params.zonaHoraria !== undefined && String(params.zonaHoraria).trim() !== ""){
    const z = String(params.zonaHoraria).trim();
    if(!zonaHorariaValida(z)) return jsonResponse({ success:false, message:"La zona horaria \"" + z + "\" no es válida." });
    nuevosValores.ZONA_HORARIA = z;
  }
  if(nuevosValores.TEMPORADA_DESDE && nuevosValores.TEMPORADA_HASTA && nuevosValores.TEMPORADA_DESDE > nuevosValores.TEMPORADA_HASTA){
    return jsonResponse({ success:false, message:"La fecha \"desde\" del tema de temporada no puede ser posterior a \"hasta\"." });
  }


  const clavesEncontradas = {};

  for(let i = 0; i < dataActual.length; i++){
    const clave = String(dataActual[i][0] || "").trim().toUpperCase();
    if(nuevosValores.hasOwnProperty(clave)){
      // +2: +1 porque dataActual ya tiene el header descartado (shift), +1 por ser 1-based
      sheetActual.getRange(i + 2, 2).setValue(nuevosValores[clave]);
      clavesEncontradas[clave] = true;
    }
  }

  // Cualquier clave que no estuviera todavía en la hoja se agrega al final
  Object.keys(nuevosValores).forEach(clave => {
    if(!clavesEncontradas[clave]){
      sheetActual.appendRow([clave, nuevosValores[clave]]);
    }
  });

  // La config pública está cacheada (ver doGet "configuracionNegocio"):
  // se invalida acá para que un cambio del admin (ej. pedido mínimo)
  // se vea al instante y no recién a los 5 minutos.
  try{ CacheService.getScriptCache().removeAll(["config_negocio:n","config_negocio:0","config_negocio:1"]); }catch(errCache){}
  invalidarCacheReglasPedido();
  if(nuevosValores.ZONA_HORARIA){
    PropertiesService.getScriptProperties().setProperty(PROP_ZONA_HORARIA, nuevosValores.ZONA_HORARIA);
    _zonaNegocioMemo = null;
  }

  return jsonResponse({
    success:true
  });

}

/**
 * Número de pedido corto y correlativo: 1001, 1002, 1003…
 * Se llama dentro del lock de guardarPedido (no se repite aunque entren
 * dos pedidos a la vez). El último número queda en una propiedad, y además
 * se toma el mayor número que haya en la hoja, por si alguien editó o
 * restauró filas a mano. Los pedidos viejos (PED-xxxxxx) quedan como están.
 */
/* =====================================================================
   ZONA HORARIA DEL NEGOCIO
   Se elige en el panel (Configuración → Negocio) y se usa para todos los
   cálculos de fecha y hora ("hoy", cierres de caja, reportes por día…).
   Si no se eligió, la del proyecto de Apps Script (appsscript.json).
===================================================================== */
const PROP_ZONA_HORARIA = "ZONA_HORARIA";
const ZONA_HORARIA_DEFECTO = "America/Argentina/Buenos_Aires";
let _zonaNegocioMemo = null;
function zonaNegocio(){
  if(_zonaNegocioMemo) return _zonaNegocioMemo;
  let z = "";
  try{ z = PropertiesService.getScriptProperties().getProperty(PROP_ZONA_HORARIA) || ""; }catch(e){}
  if(!z){ try{ z = Session.getScriptTimeZone(); }catch(e){} }
  _zonaNegocioMemo = z || ZONA_HORARIA_DEFECTO;
  return _zonaNegocioMemo;
}
/** ¿Es una zona horaria válida (IANA, ej. "America/Argentina/Buenos_Aires")? */
function zonaHorariaValida(z){
  if(!/^[A-Za-z_]+(\/[A-Za-z0-9_+\-]+){0,2}$/.test(String(z || ""))) return false;
  try{ Utilities.formatDate(new Date(), String(z), "yyyy-MM-dd HH:mm"); return true; }catch(e){ return false; }
}

const PROP_ULTIMO_PEDIDO = "PEDIDO_ULTIMO_NUMERO";
const PRIMER_NUMERO_PEDIDO = 1001;
function generarNumeroPedido(sheet, ultimaFilaConocida){
  const props = PropertiesService.getScriptProperties();
  let ultimo = Number(props.getProperty(PROP_ULTIMO_PEDIDO)) || (PRIMER_NUMERO_PEDIDO - 1);
  const existentes = new Set();
  if(sheet){
    const ultima = ultimaFilaConocida !== undefined ? ultimaFilaConocida : sheet.getLastRow();
    if(ultima >= 2){
      sheet.getRange(2, 1, ultima - 1, 1).getValues().forEach(r => {
        const v = String(r[0] === null || r[0] === undefined ? "" : r[0]).trim();
        existentes.add(v);
        if(/^\d{1,9}$/.test(v)) ultimo = Math.max(ultimo, Number(v));
      });
    }
  }
  let numero = ultimo + 1;
  while(existentes.has(String(numero))) numero++;
  props.setProperty(PROP_ULTIMO_PEDIDO, String(numero));
  return String(numero);
}

/**
 * Genera un ID corto del tipo "PREFIJO-123456" (6 dígitos numéricos),
 * verificando que no choque con ningún ID ya existente en la columna
 * indicada de la hoja (por defecto la columna 1 / A).
 */
function generarIdCorto(prefijo, sheet, columnaId, ultimaFilaConocida){

  columnaId = columnaId || 1;

  // Solo la columna del ID (antes se leía la hoja entera con
  // getDataRange en CADA venta, solo para evitar un choque de IDs).
  let existentes = new Set();
  if(sheet){
    const ultima = ultimaFilaConocida !== undefined ? ultimaFilaConocida : sheet.getLastRow();
    if(ultima >= 1){
      sheet.getRange(1, columnaId, ultima, 1).getValues().forEach(row => existentes.add(String(row[0] || "")));
    }
  }

  let id;
  let intentos = 0;
  do {
    const numero = Math.floor(100000 + Math.random() * 900000); // 6 dígitos: 100000–999999
    id = prefijo + "-" + numero;
    intentos++;
  } while(existentes.has(id) && intentos < 50);

  return id;

}

/**
 * Estado del "QR de cobro" activo — para que la pantalla/tablet del
 * cliente lo muestre consultando este mismo backend por internet, SIN
 * depender de que la tablet y la caja estén en la misma red local (el
 * enfoque por red local se rompe con aislamiento de WiFi, subredes
 * separadas, firewalls, etc. — por internet no hay ninguno de esos
 * problemas, ambos dispositivos ya hablan con esta misma URL).
 *
 * Se guarda en PropertiesService (no en la planilla): es información
 * de "ahora mismo", efímera, no tiene sentido como fila de historial.
 */
const QR_PROP_KEY = "qrCobroActivo";
const QR_ABANDONO_MS = 20 * 60 * 1000; // 20 min sin actualizarse → se considera abandonado

function qrFijar(data) {
  const estado = {
    activo: true,
    qrImagenDataUrl: data.qrImagenDataUrl,
    // Texto del QR (lo dibuja web/cobro.html, nítido y liviano)
    qrTexto: data.qrTexto ? String(data.qrTexto).slice(0, 1000) : null,
    monto: Number(data.monto) || 0,
    referencia: data.referencia || null,
    pagada: false,
    ts: Date.now()
  };
  // Límite de 9 KB por propiedad: si la imagen no entra pero está el texto,
  // se guarda solo el texto (la pantalla del celular dibuja el QR igual).
  if (JSON.stringify(estado).length > 8800 && estado.qrTexto) estado.qrImagenDataUrl = null;
  try {
    PropertiesService.getScriptProperties().setProperty(QR_PROP_KEY, JSON.stringify(estado));
    return jsonResponse({ success: true });
  } catch (error) {
    // PropertiesService tiene un límite de 9KB por valor — si algún día
    // el QR generado pesa más que eso, mejor devolver un error claro
    // que romper la venta por algo que no tiene nada que ver con MP.
    return jsonResponse({ success: false, message: "No se pudo guardar el QR: " + error });
  }
}

function qrPagado() {
  const props = PropertiesService.getScriptProperties();
  const actual = props.getProperty(QR_PROP_KEY);
  if (!actual) return jsonResponse({ success: true }); // no había nada activo, no pasa nada

  const estado = JSON.parse(actual);
  estado.pagada = true;
  estado.ts = Date.now();
  props.setProperty(QR_PROP_KEY, JSON.stringify(estado));
  return jsonResponse({ success: true });
}

function qrLimpiar() {
  PropertiesService.getScriptProperties().deleteProperty(QR_PROP_KEY);
  return jsonResponse({ success: true });
}

/**
 * Estado del cobro con QR para la pantalla del cliente. Con ?v=<ts> (lo
 * manda web/cobro.html con la versión que ya está mostrando), si el cobro
 * no cambió se responde solo "sinCambios", sin volver a mandar la imagen
 * del QR en cada consulta (eran varios KB cada 3 segundos).
 */
function qrEstado(e) {
  const vCliente = e && e.parameter && e.parameter.v ? String(e.parameter.v) : "";
  const props = PropertiesService.getScriptProperties();
  const actual = props.getProperty(QR_PROP_KEY);
  if (!actual) return jsonResponse({ activo: false });

  const estado = JSON.parse(actual);

  // Por si main.js se cerró de golpe (se corta la luz, se cierra la app
  // a la fuerza) sin llegar a limpiar el estado: si quedó "activo" hace
  // más de QR_ABANDONO_MS sin que nadie lo actualice, se lo da por
  // abandonado — mejor eso que un QR fantasma mostrado para siempre.
  if (Date.now() - (estado.ts || 0) > QR_ABANDONO_MS) {
    props.deleteProperty(QR_PROP_KEY);
    return jsonResponse({ activo: false });
  }

  if (vCliente && vCliente === String(estado.ts)) {
    return jsonResponse({ activo: true, sinCambios: true, ts: estado.ts });
  }
  return jsonResponse(estado);
}

/**
 * Misma pantalla que ya existe para la red local (ver local-server.js
 * → paginaPantallaQR), pero sirviéndola desde acá y consultando este
 * mismo backend por internet en vez de una IP de red local — así no
 * importa si la tablet y la caja están en la misma red o no.
 */
function paginaPantallaQR() {
  const miUrl = ScriptApp.getService().getUrl();

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<title>VeekPOS — Cobrá acá</title>
<style>
  * { box-sizing:border-box; }
  html,body { height:100%; margin:0; }
  body {
    position:fixed; inset:0; overflow:hidden;
    display:flex; align-items:center; justify-content:center;
    background:linear-gradient(160deg,#0b1220 0%,#1e293b 100%);
    font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;
    color:#fff; text-align:center; padding:24px;
  }
  .card { max-width:96vw; width:100%; }
  .idle .icono { font-size:26vmin; margin-bottom:26px; }
  .idle h1 { font-size:8vmin; font-weight:600; opacity:.85; margin:0; }
  .cobro img { width:82vmin; height:82vmin; border-radius:32px; background:#fff; padding:26px; box-shadow:0 30px 90px rgba(0,0,0,.5); }
  .cobro .monto { font-size:15vmin; font-weight:800; margin:32px 0 8px; font-variant-numeric:tabular-nums; }
  .cobro .sub { opacity:.75; font-size:5vmin; }
  .pagada .icono { font-size:32vmin; margin-bottom:20px; }
  .pagada h1 { font-size:10vmin; font-weight:800; color:#4ade80; margin:0 0 12px; }
  .pagada .monto { font-size:7vmin; opacity:.85; }
  .pulso { animation:pulso 1.6s ease-in-out infinite; }
  @keyframes pulso { 0%,100%{opacity:1;} 50%{opacity:.55;} }
</style>
</head>
<body>
  <div class="card" id="card"></div>
  <script>
    const URL_ESTADO = "${miUrl}?action=qrEstado";
    const card = document.getElementById("card");

    function money(n) {
      return "$" + Number(n || 0).toLocaleString("es-AR");
    }

    function render(estado) {
      if (estado.activo && estado.pagada) {
        card.className = "pagada";
        card.innerHTML = '<div class="icono">✅</div><h1>¡Pago recibido!</h1><div class="monto">' + money(estado.monto) + '</div>';
      } else if (estado.activo) {
        card.className = "cobro";
        card.innerHTML =
          '<img src="' + estado.qrImagenDataUrl + '" alt="QR de cobro">' +
          '<div class="monto">' + money(estado.monto) + '</div>' +
          '<div class="sub pulso">Escaneá con Mercado Pago para pagar</div>';
      } else {
        card.className = "idle";
        card.innerHTML = '<div class="icono">🛒</div><h1>Esperando la próxima venta…</h1>';
      }
    }

    async function actualizar() {
      try {
        const r = await fetch(URL_ESTADO, { cache: "no-store" });
        const data = await r.json();
        render(data);
      } catch (error) {
        // Sin conexión momentánea — no se rompe la pantalla, se
        // reintenta solo en el próximo ciclo.
      }
    }

    render({ activo: false });
    actualizar();
    setInterval(actualizar, 3000);
  </script>
</body>
</html>`;
}

function jsonResponse(obj) {

  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
/**
 * Normaliza texto para comparar transportes ignorando mayúsculas,
 * acentos, y espacios de más — así "Vía Cargo", "via cargo" y
 * "VIA  CARGO" se consideran el mismo texto.
 */
function normalizarTextoTransporte(texto){
  return String(texto || "")
    .trim()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // saca acentos
    .replace(/\s+/g, " ");
}

/**
 * Reglas que se validan en cada pedido del catálogo: pedido mínimo y
 * transportes no disponibles. Antes se leía la hoja CONFIGURACION entera
 * DOS veces por pedido; ahora una sola lectura, guardada 2 minutos en la
 * caché. Guardar la configuración desde el panel borra esta caché, así
 * que un cambio del admin se aplica en el pedido siguiente.
 */
const CLAVE_CACHE_REGLAS_PEDIDO = "reglas_pedido_v1";
let _reglasPedidoMemo = null;
function leerReglasPedido(){
  if(_reglasPedidoMemo) return _reglasPedidoMemo;
  try{
    const t = CacheService.getScriptCache().get(CLAVE_CACHE_REGLAS_PEDIDO);
    if(t){ _reglasPedidoMemo = JSON.parse(t); return _reglasPedidoMemo; }
  }catch(e){}
  const reglas = { minimo: null, transportes: "" };
  try{
    const filas = getOrCrearHojaConfiguracion().getDataRange().getValues();
    for(let i = 1; i < filas.length; i++){
      const clave = String(filas[i][0] || "").trim().toUpperCase();
      if(clave === "PEDIDO_MINIMO"){
        const raw = filas[i][1];
        const n = Number(raw);
        reglas.minimo = (raw === "" || raw === null || raw === undefined || isNaN(n) || n < 0) ? 100000 : n;
      } else if(clave === "TRANSPORTES_NO_DISPONIBLES"){
        reglas.transportes = String(filas[i][1] || "");
      }
    }
    try{ CacheService.getScriptCache().put(CLAVE_CACHE_REGLAS_PEDIDO, JSON.stringify(reglas), 120); }catch(e){}
  }catch(err){
    Logger.log("leerReglasPedido: " + err);
    // reglas.minimo queda en null → no se bloquea el pedido por no poder leerlo
  }
  _reglasPedidoMemo = reglas;
  return reglas;
}

function invalidarCacheReglasPedido(){
  _reglasPedidoMemo = null;
  try{ CacheService.getScriptCache().remove(CLAVE_CACHE_REGLAS_PEDIDO); }catch(e){}
}

/** Lista de nombres de transporte bloqueados, ya normalizados */
function obtenerTransportesNoDisponibles(){
  return String(leerReglasPedido().transportes || "")
    .split("\n")
    .map(normalizarTextoTransporte)
    .filter(Boolean);
}

/**
 * Devuelve un mapa CODIGO -> STOCK actual, leído fresco de la hoja
 * PRODUCTOS, para los códigos pedidos (no hace falta traer todo el
 * catálogo si el carrito tiene pocos ítems).
 */
function obtenerStockDesdeCacheCatalogo(codigos){
  // Devuelve un mapa CODIGO -> STOCK usando la caché del catálogo público
  // (se invalida sola cuando cambia el stock). Devuelve null si no hay caché.
  try{
    const texto = cacheLeerTexto("catalogo_publico");
    if(!texto) return null;
    const obj = JSON.parse(texto);
    const set = new Set(codigos.map(c => String(c).trim()));
    const mapa = {};
    (obj.productos || []).forEach(p => {
      const c = String(p.CODIGO || "").trim();
      if(set.has(c)) mapa[c] = Number(p.STOCK) || 0;
    });
    return mapa;
  }catch(e){ return null; }
}

function obtenerStockActualPorCodigo(codigos){
  const sheet = getSheet();
  if(!sheet) return {};

  const ultimaFila = sheet.getLastRow();
  const ultimaCol = sheet.getLastColumn();
  if(ultimaFila < 2 || ultimaCol < 1) return {};

  const headers = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0]
    .map(h => String(h || "").trim());
  const idxCodigo = headers.indexOf("CODIGO");
  const idxStock = headers.indexOf("STOCK");
  if(idxCodigo === -1 || idxStock === -1) return {};

  // Solo se leen las dos columnas necesarias (no toda la hoja).
  const codigosCol = sheet.getRange(2, idxCodigo + 1, ultimaFila - 1, 1).getValues();
  const stockCol = sheet.getRange(2, idxStock + 1, ultimaFila - 1, 1).getValues();

  const codigosSet = new Set(codigos.map(c => String(c).trim()));
  const mapa = {};
  for(let i = 0; i < codigosCol.length; i++){
    const codigo = String(codigosCol[i][0] || "").trim();
    if(codigosSet.has(codigo)){
      mapa[codigo] = Number(stockCol[i][0]) || 0;
    }
  }
  return mapa;
}

/** Pedido mínimo vigente (null si no se pudo determinar → no se bloquea). */
function leerPedidoMinimoActual(){
  return leerReglasPedido().minimo;
}

/**
 * Si la hoja DETALLE_PEDIDOS todavía no tiene la columna COLOR (usada
 * para guardar el color/variante elegido en el Quick View del
 * catálogo web, cuando el producto tiene más de una imagen), la
 * agrega al final — sin tocar ninguna columna ni fila existente.
 * Mismo patrón que asegurarColumnaAlias.
 */
/** Igual que asegurarColumnaColorDetalle, pero recuerda (6 h) que ya existe. */
function asegurarColumnaColorDetalleRapido(sheet){
  try{ if(CacheService.getScriptCache().get("detalle_color_ok")) return; }catch(e){}
  asegurarColumnaColorDetalle(sheet);
  try{ CacheService.getScriptCache().put("detalle_color_ok", "1", 21600); }catch(e){}
}

function asegurarColumnaColorDetalle(sheet){
  if(!sheet) return;
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if(headerRow.indexOf("COLOR") !== -1) return;

  const colNueva = sheet.getLastColumn() + 1;
  sheet.getRange(1, colNueva).setValue("COLOR");
}

/* =====================================================================
   ENDPOINTS PÚBLICOS — VALIDACIÓN DEL LADO DEL SERVIDOR
   Todo lo que llega desde el catálogo web (navegador de un cliente
   cualquiera) se trata como NO CONFIABLE: el precio, el total, el
   nombre del producto y la moneda se recalculan desde la hoja
   PRODUCTOS; lo que mande el navegador para esos campos se ignora.
===================================================================== */

/**
 * Limitador simple por ventana de tiempo, con CacheService.
 * Devuelve true si todavía se puede seguir; false si ya se pasó del máximo.
 * (No es atómico al 100% entre ejecuciones simultáneas, pero alcanza
 * para frenar fuerza bruta y abuso automatizado.)
 */
/**
 * Igual que permitirIntento pero para varios límites a la vez, con una
 * lectura y una escritura de caché (en vez de una de cada por límite).
 * Cada entrada: [clave, maximo, ventanaSegundos]. Si alguno está en el
 * tope, no se cuenta ninguno y devuelve false.
 */
function permitirIntentosJuntos(limites){
  try{
    const cache = CacheService.getScriptCache();
    const claves = limites.map(l => "rl:" + l[0]);
    const actuales = cache.getAll(claves) || {};
    for(let i = 0; i < limites.length; i++){
      if(Number(actuales[claves[i]] || 0) >= limites[i][1]) return false;
    }
    // Se escribe agrupado por ventana, así cada límite conserva su vencimiento.
    const porVentana = {};
    limites.forEach((l, i) => {
      (porVentana[l[2]] = porVentana[l[2]] || {})[claves[i]] = String(Number(actuales[claves[i]] || 0) + 1);
    });
    Object.keys(porVentana).forEach(v => cache.putAll(porVentana[v], Number(v)));
    return true;
  }catch(e){
    return true; // si la caché falla, no se bloquea a nadie por error
  }
}

function permitirIntento(clave, maximo, ventanaSegundos){
  try{
    const cache = CacheService.getScriptCache();
    const k = "rl:" + clave;
    const actual = Number(cache.get(k) || 0);
    if(actual >= maximo) return false;
    cache.put(k, String(actual + 1), ventanaSegundos);
    return true;
  }catch(e){
    return true; // si la caché falla, no se bloquea a nadie por error
  }
}

/** Texto seguro para escribir en la planilla: recorta y evita que "=..." se ejecute como fórmula */
function textoSeguroHoja(valor, max){
  let s = String(valor === undefined || valor === null ? "" : valor).replace(/[\u0000-\u001f]/g, " ").trim();
  if(s.length > max) s = s.slice(0, max);
  if(/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/**
 * Punto de entrada del pedido del catálogo web (GET y POST). Valida,
 * arma el carrito con precios REALES y recién ahí llama a guardarPedido().
 * Los pedidos del POS siguen entrando por guardarPedidoAdmin (no por acá).
 */
function guardarPedidoPublico(entrada){

  const _t0Pedido = Date.now();
  const d = entrada || {};

  // --- Campos del cliente: obligatorios, con largo máximo ---
  const nombre = textoSeguroHoja(d.nombre, 120);
  const direccion = textoSeguroHoja(d.direccion, 200);
  const localidad = textoSeguroHoja(d.localidad, 80);
  const provincia = textoSeguroHoja(d.provincia, 80);
  const telefono = textoSeguroHoja(d.telefono, 40);
  const dni = textoSeguroHoja(d.dni, 20);

  if(!nombre || !direccion || !localidad || !provincia || !telefono || !dni){
    return jsonResponse({ success:false, message:"Completá Nombre, Dirección, Localidad, Provincia, Teléfono y DNI o CUIT." });
  }

  // --- Reintento de un pedido ya guardado (se perdió la respuesta): se
  // devuelve el pedido existente ANTES de validar nada más — si no, un
  // cupón de "uno por cliente" diría que ya se usó (justamente en ese pedido).
  const idOpPrevio = String(d.idOperacion || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 64);
  if(idOpPrevio){
    const previo = buscarPedidoPorIdOperacionRapido(null, idOpPrevio);
    if(previo) return jsonResponse({ success:true, pedidoId: previo, repetido:true, version: VERSION_BACKEND });
  }

  // --- Frenos contra abuso automatizado ---
  const dniNorm = normalizarDni(dni) || dni;
  if(!permitirIntentosJuntos([["pedido_dni:" + dniNorm, 10, 3600], ["pedido_global", 200, 600]])){
    return jsonResponse({ success:false, message:"Recibimos muchos pedidos seguidos. Probá de nuevo en unos minutos." });
  }

  // --- Carrito: forma y tamaño ---
  let carritoCliente = d.carrito;
  if(typeof carritoCliente === "string"){
    try{ carritoCliente = JSON.parse(carritoCliente); }catch(err){ carritoCliente = null; }
  }
  if(!Array.isArray(carritoCliente) || carritoCliente.length === 0){
    return jsonResponse({ success:false, message:"El pedido no tiene productos." });
  }
  if(carritoCliente.length > 300){
    return jsonResponse({ success:false, message:"El pedido tiene demasiados renglones. Dividilo en dos pedidos." });
  }

  // --- Precio y stock REALES, solo de los códigos del carrito ---
  // Antes se rearmaba el catálogo entero (getProductos) cuando no estaba
  // en caché — y casi nunca lo está, porque el sitio lee el catálogo por
  // el Worker. Eso podía tardar más de 30-45 s y el pedido "no respondía".
  let porCodigo;
  try{
    porCodigo = leerProductosParaPedido(carritoCliente.map(it => String((it || {}).CODIGO === undefined ? "" : it.CODIGO).trim()));
  }catch(err){
    console.error("guardarPedido: no se pudieron leer precios: " + err);
    return jsonResponse({ success:false, message:"No pudimos verificar los precios. Probá de nuevo en unos segundos." });
  }
  console.log("guardarPedido: precios y stock leídos en " + (Date.now() - _t0Pedido) + " ms");

  const carritoReal = [];
  const cantidadPorCodigo = {};
  let totalReal = 0;

  for(let i = 0; i < carritoCliente.length; i++){
    const it = carritoCliente[i] || {};
    const codigo = String(it.CODIGO === undefined ? "" : it.CODIGO).trim();
    const prod = porCodigo[codigo];

    if(!prod){
      return jsonResponse({ success:false, message:"Un producto de tu carrito ya no está disponible. Actualizá la página y revisá el pedido." });
    }

    const cantidad = Number(it.cantidad);
    if(!isFinite(cantidad) || cantidad < 1 || cantidad > 100000 || Math.floor(cantidad) !== cantidad){
      return jsonResponse({ success:false, message:"Hay una cantidad inválida en el pedido." });
    }

    let precioUnitario = Number(prod.PRECIO);
    let nombreProducto = String(prod.PRODUCTO || "");

    if(it._esCaja){
      const unidades = Number(prod.UNIDADES_POR_CAJA) || 0;
      const precioCaja = Number(prod.PRECIO_CAJA) || 0;
      if(unidades <= 0 || precioCaja <= 0 || cantidad % unidades !== 0){
        return jsonResponse({ success:false, message:"El precio por caja de un producto cambió. Actualizá la página y volvé a armar el pedido." });
      }
      precioUnitario = precioCaja / unidades;
      nombreProducto = nombreProducto + " (caja x" + unidades + ")";
    }

    if(!isFinite(precioUnitario) || precioUnitario <= 0){
      return jsonResponse({ success:false, message:"Un producto no tiene precio disponible. Actualizá la página." });
    }

    cantidadPorCodigo[codigo] = (cantidadPorCodigo[codigo] || 0) + cantidad;

    totalReal += precioUnitario * cantidad;
    carritoReal.push({
      CODIGO: prod.CODIGO,
      PRODUCTO: nombreProducto,
      PRECIO: precioUnitario,
      cantidad: cantidad,
      COLOR: textoSeguroHoja(it.COLOR, 60)
    });
  }

  // El stock se valida sumando TODAS las líneas del mismo código
  // (antes cada línea se comparaba sola contra el stock).
  const sinStock = [];
  const faltantes = [];
  Object.keys(cantidadPorCodigo).forEach(codigo => {
    const stock = Number(porCodigo[codigo].STOCK) || 0;
    if(cantidadPorCodigo[codigo] > stock){
      sinStock.push(porCodigo[codigo].PRODUCTO + " (pediste " + cantidadPorCodigo[codigo] + ", disponible " + Math.max(0, stock) + ")");
      faltantes.push({ codigo: String(codigo), producto: porCodigo[codigo].PRODUCTO, pedido: cantidadPorCodigo[codigo], disponible: Math.max(0, stock) });
    }
  });
  if(sinStock.length){
    // stockInsuficiente: el catálogo ajusta el carrito solo y explica cada cambio
    return jsonResponse({ success:false, stockInsuficiente: faltantes, message:"Algunos productos no tienen el stock que pediste: " + sinStock.join("; ") });
  }

  totalReal = Math.round(totalReal * 100) / 100;

  // --- Cupón de descuento (opcional). Acá se valida y se calcula; el
  // uso se descuenta recién dentro del lock de guardarPedido, junto con
  // el pedido, para que dos clientes no se lleven el último uso.
  const subtotalReal = totalReal;
  let cuponAplicado = null;
  const codigoCupon = normalizarCodigoCupon(d.cupon);
  if(codigoCupon){
    const ev = evaluarCupon(codigoCupon, subtotalReal, normalizarDni(dni));
    if(!ev.ok){
      return jsonResponse({ success:false, cuponInvalido:true, message: ev.message });
    }
    cuponAplicado = { codigo: ev.cupon.CODIGO, descuento: ev.descuento, etiqueta: ev.etiqueta };
    totalReal = Math.round((subtotalReal - ev.descuento) * 100) / 100;
  }

  // Si el total que vio el cliente no coincide con el real, los precios
  // cambiaron mientras tenía el carrito abierto: se le avisa, no se guarda.
  const totalCliente = Number(d.total);
  if(!isFinite(totalCliente) || Math.abs(totalCliente - totalReal) > Math.max(1, totalReal * 0.001)){
    return jsonResponse({ success:false, message:"Los precios se actualizaron mientras armabas el pedido. Actualizá la página y revisá el total antes de enviarlo." });
  }

  // --- A guardarPedido: solo campos permitidos, con origen SIEMPRE catálogo web ---
  const _respuestaPedido = guardarPedido({
    idOperacion: d.idOperacion,
    nombre: nombre,
    empresa: textoSeguroHoja(d.empresa, 120),
    direccion: direccion,
    telefono: telefono,
    dni: dni,
    codigoPostal: textoSeguroHoja(d.codigoPostal, 12),
    localidad: localidad,
    provincia: provincia,
    total: totalReal,
    subtotal: subtotalReal,
    descuento: cuponAplicado ? cuponAplicado.descuento : 0,
    descuentoEtiqueta: cuponAplicado ? cuponAplicado.etiqueta : "",
    carrito: carritoReal,
    _cupon: cuponAplicado,
    _stockValidado: true
    // origenPOS, moneda, descuento, subtotal: a propósito NO se aceptan desde el público
  });
  // Aviso de pedido nuevo para las cajas (ver avisoPedidos): solo
  // pedidos nuevos de verdad, no reintentos.
  try{
    const r = JSON.parse(_respuestaPedido.getContent());
    if(r && r.success && !r.repetido && r.pedidoId) marcarPedidoWebNuevo(r.pedidoId, nombre, totalReal);
  }catch(e){ console.error("aviso pedido nuevo: " + e.message); }
  // Visible en Apps Script → Ejecuciones: cuánto tardó cada pedido.
  console.log("guardarPedido: pedido completo en " + (Date.now() - _t0Pedido) + " ms");
  return _respuestaPedido;
}

/* =====================================================================
   INFORMES DE ERROR
   - Panel / app de escritorio: "🐞 Informar un problema" (la persona
     describe qué pasó) + errores inesperados que se mandan solos.
   - Catálogo web: solo los automáticos (errores de la página o un
     pedido que no se pudo enviar), con tope diario: es una acción
     pública.
   - Servidor: respuestaErrorInterno avisa sola.
   Todo queda en la hoja ERRORES; el correo va a CORREO_SOPORTE.
   Topes: el mismo error automático se avisa 1 vez cada 6 h; como
   máximo 40 correos automáticos por día (los manuales no tienen tope
   propio, solo la cuota de Google: 100/día en cuentas gratuitas).
===================================================================== */
const ENC_ERRORES = ["FECHA", "TIPO", "ORIGEN", "USUARIO", "TITULO", "DESCRIPCION", "CONTACTO", "CODIGO", "VERSION", "ENVIADO", "DETALLE"];
const TOPE_INFORMES_AUTOMATICOS_DIA = 40;
const TOPE_INFORMES_CATALOGO_DIA = 15;

function getOrCrearHojaErrores(){
  let sheet = getHoja("ERRORES");
  if(!sheet){
    sheet = getSS().insertSheet("ERRORES");
    sheet.appendRow(ENC_ERRORES);
    sheet.setFrozenRows(1);
    _hojaCache["ERRORES"] = sheet;
  }
  return sheet;
}

function _textoCorto(v, max){
  const t = typeof v === "string" ? v : (v === undefined || v === null ? "" : JSON.stringify(v));
  return t.length > max ? t.slice(0, max) + "…" : t;
}

function _escHtml(v){
  return String(v === undefined || v === null ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Cuenta envíos del día por tipo (en propiedades: sobrevive a la caché). */
function _contarInformeDelDia(clave, tope){
  const props = PropertiesService.getScriptProperties();
  const hoy = Utilities.formatDate(new Date(), zonaNegocio() || "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  const k = "VEEK_INFORMES_" + clave;
  const partes = String(props.getProperty(k) || "").split("|");
  const usados = partes[0] === hoy ? Number(partes[1]) || 0 : 0;
  if(usados >= tope) return false;
  props.setProperty(k, hoy + "|" + (usados + 1));
  return true;
}

/**
 * Anota el informe en ERRORES y lo manda por correo (si corresponde).
 * info: { tipo, titulo, descripcion, contacto, codigo, contexto{}, errores[], automatico }
 * Devuelve { enviado:bool, motivo }.
 */
function enviarInformeError(info){
  info = info || {};
  const automatico = !!info.automatico;
  const tipo = String(info.tipo || "panel");
  const titulo = _textoCorto(info.titulo || info.descripcion || "Sin título", 140);
  const sesion = _sesionActual || null;
  const usuario = sesion ? String(sesion.n || sesion.u || "") + (sesion.r ? " (" + sesion.r + ")" : "") : "";
  const contexto = info.contexto && typeof info.contexto === "object" ? info.contexto : {};
  const errores = Array.isArray(info.errores) ? info.errores.slice(0, 25) : [];
  const pasos = Array.isArray(info.pasos) ? info.pasos.slice(-30) : [];
  const cola = info.cola && typeof info.cola === "object" ? info.cola : null;
  let captura = null;
  if(info.captura && typeof info.captura === "string" && info.captura.length < 3500000 && /^[A-Za-z0-9+/=]+$/.test(info.captura.slice(0, 200))){
    try{ captura = Utilities.newBlob(Utilities.base64Decode(info.captura), "image/jpeg", "captura-" + Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "yyyyMMdd-HHmm") + ".jpg"); }
    catch(e){ console.error("captura inválida: " + e.message); }
  }

  let enviado = false, motivo = "";

  // Agrupar automáticos: el mismo error, una vez cada 6 horas.
  if(automatico){
    const firma = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, tipo + "|" + titulo)).slice(0, 22);
    const cache = CacheService.getScriptCache();
    if(cache.get("inf_" + firma)){
      return { enviado: false, motivo: "repetido" };
    }
    cache.put("inf_" + firma, "1", 6 * 3600);
  }

  let negocio = "";
  try{ negocio = String(leerClaveConfiguracion("NOMBRE", "") || ""); }catch(e){}
  let urlPlanilla = "";
  try{ urlPlanilla = getSS().getUrl(); }catch(e){}

  if(!CORREO_SOPORTE){
    motivo = "sin correo de soporte configurado";
  }else if(automatico && !_contarInformeDelDia("AUTO", TOPE_INFORMES_AUTOMATICOS_DIA)){
    motivo = "tope diario de informes automáticos";
  }else{
    try{
      if(MailApp.getRemainingDailyQuota() < 1){
        motivo = "sin cuota de correo de Google por hoy";
      }else{
        const etiqueta = { servidor: "Error del servidor", panel: automatico ? "Error en el panel" : "Problema informado", catalogo: "Error en el catálogo web" }[tipo] || "Informe";
        const asunto = "[VeekPOS] " + (negocio || "Negocio") + " · " + etiqueta + ": " + titulo;
        const filas = [];
        const fila = (k, v) => { if(v !== undefined && v !== null && String(v) !== "") filas.push("<tr><td style='padding:4px 10px 4px 0;color:#666;vertical-align:top;white-space:nowrap'>" + _escHtml(k) + "</td><td style='padding:4px 0'>" + _escHtml(v) + "</td></tr>"); };
        fila("Negocio", negocio);
        fila("Tipo", etiqueta + (automatico ? " (automático)" : ""));
        fila("Fecha", Utilities.formatDate(new Date(), zonaNegocio() || "America/Argentina/Buenos_Aires", "dd/MM/yyyy HH:mm:ss"));
        fila("Usuario", usuario);
        fila("Contacto", info.contacto);
        fila("Código", info.codigo);
        fila("Backend", VERSION_BACKEND);
        if(contexto.versionPanel && String(contexto.versionPanel) !== VERSION_BACKEND){
          fila("⚠️ Versiones", "el panel es " + contexto.versionPanel + " y el backend " + VERSION_BACKEND + " (¿panel o app sin actualizar, o el navegador con la página vieja guardada?)");
        }
        Object.keys(contexto).slice(0, 30).forEach(k => fila(k, _textoCorto(contexto[k], 400)));
        fila("Planilla", urlPlanilla);
        let html = "<div style='font-family:Arial,sans-serif;font-size:14px;color:#222'>";
        html += "<h2 style='margin:0 0 6px;font-size:18px'>" + _escHtml(etiqueta) + "</h2>";
        if(info.descripcion) html += "<p style='white-space:pre-wrap;background:#f6f6f6;padding:10px;border-radius:6px'>" + _escHtml(_textoCorto(info.descripcion, 4000)) + "</p>";
        html += "<table style='border-collapse:collapse;font-size:13px'>" + filas.join("") + "</table>";
        if(pasos.length){
          html += "<h3 style='font-size:15px;margin:16px 0 6px'>Últimos pasos (" + pasos.length + ")</h3><ol style='font-size:12.5px;margin:0;padding-left:20px;color:#333'>" +
            pasos.map(p => "<li><span style='color:#888'>" + _escHtml(p && p.hora || "") + "</span> " + _escHtml(_textoCorto(p && p.paso || "", 200)) + (p && p.veces > 1 ? " <span style='color:#888'>(×" + _escHtml(p.veces) + ")</span>" : "") + "</li>").join("") + "</ol>";
        }
        if(cola){
          html += "<h3 style='font-size:15px;margin:16px 0 6px'>Sincronización de la caja</h3><p style='font-size:13px;margin:0 0 6px'>" +
            _escHtml((cola.pendientes || 0) + " pendiente(s) · " + (cola.conError || 0) + " con error" + (cola.ultimaSubida ? " · última subida " + cola.ultimaSubida : "") + (cola.redLocal ? " · red local: " + cola.redLocal : "") + (cola.error ? " · " + cola.error : "")) + "</p>";
          if(Array.isArray(cola.errores) && cola.errores.length){
            html += "<ul style='font-size:12.5px;margin:0;padding-left:20px'>" + cola.errores.slice(0, 10).map(op =>
              "<li><b>" + _escHtml(op.tipo) + "</b> " + _escHtml(op.ref || "") + " — " + _escHtml(_textoCorto(op.error || "", 300)) + " <span style='color:#888'>(" + _escHtml(op.intentos || 0) + " intentos, desde " + _escHtml(op.desde || "?") + ")</span></li>").join("") + "</ul>";
          }
        }
        if(captura) html += "<p style='font-size:13px;margin:16px 0 0'>📎 Captura de pantalla adjunta.</p>";
        if(errores.length){
          html += "<h3 style='font-size:15px;margin:16px 0 6px'>Errores registrados (" + errores.length + ")</h3>";
          html += errores.map(er => "<pre style='white-space:pre-wrap;background:#fff4f4;border-left:3px solid #d33;padding:8px;font-size:12px;margin:0 0 8px'>" +
            _escHtml(_textoCorto((er.hora ? "[" + er.hora + "] " : "") + (er.mensaje || "") + (er.origen ? "\n" + er.origen : "") + (er.pila ? "\n" + er.pila : ""), 2500)) + "</pre>").join("");
        }
        html += "</div>";
        const opciones = { name: "VeekPOS · " + (negocio || "informes"), htmlBody: html };
        if(captura) opciones.attachments = [captura];
        if(info.contacto && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(info.contacto).trim())) opciones.replyTo = String(info.contacto).trim();
        MailApp.sendEmail(CORREO_SOPORTE, asunto.slice(0, 250), "Informe de VeekPOS — ver la versión HTML de este correo.", opciones);
        enviado = true;
      }
    }catch(e){
      motivo = "no se pudo enviar el correo: " + e.message;
      console.error("enviarInformeError: " + e.message);
    }
  }

  try{
    const hojaErr = getOrCrearHojaErrores();
    if(hojaErr.getLastRow() > MAX_FILAS_HOJA_ERRORES + 50) limpiarHojaErrores();
    hojaErr.appendRow([
      new Date(), tipo + (automatico ? " (auto)" : ""), _textoCorto(contexto.seccion || contexto.pagina || "", 80), usuario,
      titulo, _textoCorto(info.descripcion || "", 2000), _textoCorto(info.contacto || "", 120), info.codigo || "",
      VERSION_BACKEND, enviado ? "SI" : "NO" + (motivo ? " — " + motivo : ""),
      _textoCorto({ contexto: contexto, pasos: pasos, cola: cola, errores: errores, captura: captura ? "adjunta en el correo" : undefined }, 9000)
    ].map(v => typeof v === "string" ? textoSeguroHoja(v, 9000) : v));
  }catch(e){ console.error("hoja ERRORES: " + e.message); }

  return { enviado: enviado, motivo: motivo };
}

/**
 * La hoja ERRORES guarda solo los últimos DIAS_HOJA_ERRORES días (corre
 * todos los días con el trigger de las 3 AM, ver limpiarLedgerIdempotencia;
 * los correos ya enviados no se tocan). Además, nunca pasa de
 * MAX_FILAS_HOJA_ERRORES filas aunque haya una ráfaga de errores.
 * Las filas están en orden de llegada: se borran las primeras.
 */
const DIAS_HOJA_ERRORES = 7;
const MAX_FILAS_HOJA_ERRORES = 1500;

function limpiarHojaErrores(){
  const hoja = getHoja("ERRORES");
  if(!hoja) return { borradas: 0 };
  const ultima = hoja.getLastRow();
  if(ultima < 2) return { borradas: 0 };
  const fechas = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  const limite = Date.now() - DIAS_HOJA_ERRORES * 24 * 3600 * 1000;
  let viejas = 0;
  while(viejas < fechas.length){
    const f = fechas[viejas][0];
    const t = f instanceof Date ? f.getTime() : new Date(f).getTime();
    if(!isNaN(t) && t >= limite) break;
    viejas++;
  }
  const sobrantes = Math.max(0, fechas.length - MAX_FILAS_HOJA_ERRORES);
  const borrar = Math.max(viejas, sobrantes);
  if(borrar > 0){
    if(borrar >= fechas.length){
      // Google no deja borrar todas las filas no congeladas: la fila 2
      // queda vacía (appendRow la vuelve a usar).
      hoja.getRange(2, 1, fechas.length, hoja.getLastColumn()).clearContent();
      if(fechas.length > 1) hoja.deleteRows(3, fechas.length - 1);
    }else{
      hoja.deleteRows(2, borrar);
    }
  }
  return { borradas: borrar };
}

/** Panel / app de escritorio (con sesión). */
function reportarErrorPanel(d){
  const automatico = !!d.automatico;
  if(automatico && !REPORTES_AUTOMATICOS) return jsonResponse({ success: true, enviado: false });
  const descripcion = String(d.descripcion || "").trim();
  if(!automatico && descripcion.length < 5){
    return jsonResponse({ success: false, message: "Contá en pocas palabras qué pasó." });
  }
  const r = enviarInformeError({
    tipo: "panel",
    titulo: d.titulo || descripcion.split("\n")[0],
    descripcion: descripcion,
    contacto: String(d.contacto || "").slice(0, 120),
    contexto: d.contexto,
    errores: d.errores,
    pasos: d.pasos,
    cola: d.cola,
    captura: automatico ? null : d.captura,
    automatico: automatico
  });
  if(!automatico && !r.enviado && r.motivo !== "repetido"){
    return jsonResponse({ success: true, enviado: false, message: "Quedó anotado en la planilla (hoja ERRORES), pero el correo no salió: " + r.motivo + "." });
  }
  return jsonResponse({ success: true, enviado: r.enviado });
}

/** Catálogo web (público): solo automáticos y con tope diario propio. */
function reportarErrorCatalogo(d){
  if(!REPORTES_AUTOMATICOS) return jsonResponse({ success: true });
  if(!permitirIntento("informe_catalogo", 20, 3600)) return jsonResponse({ success: true });
  if(!_contarInformeDelDia("CATALOGO", TOPE_INFORMES_CATALOGO_DIA)) return jsonResponse({ success: true });
  const contexto = d.contexto && typeof d.contexto === "object" ? d.contexto : {};
  const limpio = {};
  Object.keys(contexto).slice(0, 15).forEach(k => { limpio[String(k).slice(0, 40)] = _textoCorto(contexto[k], 300); });
  enviarInformeError({
    tipo: "catalogo",
    titulo: _textoCorto(d.titulo || "Error en el catálogo", 140),
    descripcion: _textoCorto(d.descripcion || "", 1000),
    contexto: limpio,
    errores: (Array.isArray(d.errores) ? d.errores : []).slice(0, 8).map(er => ({
      hora: _textoCorto(er && er.hora || "", 20), mensaje: _textoCorto(er && er.mensaje || "", 500),
      origen: _textoCorto(er && er.origen || "", 200), pila: _textoCorto(er && er.pila || "", 1200)
    })),
    automatico: true
  });
  return jsonResponse({ success: true });
}

/** Desde el editor: manda un informe de prueba para verificar que el correo llega. */
function probarInformeDeError(){
  const r = enviarInformeError({ tipo: "panel", titulo: "Prueba de informes de error", descripcion: "Correo de prueba enviado desde el editor de Apps Script (probarInformeDeError).", contexto: { origen: "editor" } });
  Logger.log(r.enviado ? "✓ Enviado a " + CORREO_SOPORTE : "✗ No se envió: " + r.motivo);
  return r;
}

/* =====================================================================
   AVISO DE PEDIDO WEB NUEVO
   Cada pedido del catálogo deja en una propiedad del script un número
   que sube (n) y los datos del último. Las cajas preguntan cada ~30 s
   con action=avisoPedidos: no se lee ninguna hoja, así que es liviano
   aunque haya varias cajas abiertas todo el día.
===================================================================== */
const PROP_ULTIMO_PEDIDO_WEB = "VEEK_ULTIMO_PEDIDO_WEB";

function marcarPedidoWebNuevo(pedidoId, cliente, total){
  const props = PropertiesService.getScriptProperties();
  let previo = {};
  try{ previo = JSON.parse(props.getProperty(PROP_ULTIMO_PEDIDO_WEB) || "{}") || {}; }catch(e){}
  const recientes = (Array.isArray(previo.recientes) ? previo.recientes : []).slice(0, 9);
  const nuevo = { id: String(pedidoId), cliente: String(cliente || "").slice(0, 60), total: Number(total) || 0, fecha: Date.now() };
  recientes.unshift(nuevo);
  props.setProperty(PROP_ULTIMO_PEDIDO_WEB, JSON.stringify({ n: (Number(previo.n) || 0) + 1, recientes: recientes }));
}

function avisoPedidos(){
  let d = {};
  try{ d = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_ULTIMO_PEDIDO_WEB) || "{}") || {}; }catch(e){}
  return jsonResponse({ success: true, n: Number(d.n) || 0, recientes: Array.isArray(d.recientes) ? d.recientes : [] });
}

/* =====================================================================
   CUPONES DE DESCUENTO (catálogo web)
   Hojas:
     CUPONES       CODIGO | TIPO (PORCENTAJE/MONTO) | VALOR | USOS_MAX |
                   USOS | VENCE | MINIMO | ACTIVO | NOTA | CREADO | DESDE
     CUPONES_USOS  FECHA | CODIGO | DNI | PEDIDO_ID | DESCUENTO
   Reglas: cada cupón tiene un cupo total de usos y cada DNI/CUIT puede
   usar un mismo cupón UNA vez. El uso se descuenta dentro del lock que
   guarda el pedido (ver guardarPedido), así no se pasa del cupo.
===================================================================== */
const ENC_CUPONES = ["CODIGO", "TIPO", "VALOR", "USOS_MAX", "USOS", "VENCE", "MINIMO", "ACTIVO", "NOTA", "CREADO", "DESDE"];
const ENC_CUPONES_USOS = ["FECHA", "CODIGO", "DNI", "PEDIDO_ID", "DESCUENTO"];

function getOrCrearHojaCupones(){
  let sheet = getHoja("CUPONES");
  if(!sheet){
    sheet = getSS().insertSheet("CUPONES");
    sheet.appendRow(ENC_CUPONES);
    sheet.setFrozenRows(1);
    _hojaCache["CUPONES"] = sheet;
    return sheet;
  }
  // Hojas creadas antes de que existiera una columna (ej. DESDE): se
  // completa el encabezado en su posición (siempre se agregan al final).
  const enc = sheet.getRange(1, 1, 1, ENC_CUPONES.length).getValues()[0];
  ENC_CUPONES.forEach((k, i) => { if(String(enc[i] || "").trim() !== k && !String(enc[i] || "").trim()) sheet.getRange(1, i + 1).setValue(k); });
  return sheet;
}

function getOrCrearHojaCuponesUsos(){
  let sheet = getHoja("CUPONES_USOS");
  if(!sheet){
    sheet = getSS().insertSheet("CUPONES_USOS");
    sheet.appendRow(ENC_CUPONES_USOS);
    sheet.setFrozenRows(1);
    _hojaCache["CUPONES_USOS"] = sheet;
  }
  return sheet;
}

function normalizarCodigoCupon(c){
  return String(c || "").trim().toUpperCase().replace(/\s+/g, "").replace(/[^A-Z0-9_-]/g, "").slice(0, 30);
}

/** Todas las filas de CUPONES como objetos (con _fila = número de fila en la hoja). */
function _leerCupones(){
  const sheet = getOrCrearHojaCupones();
  const ultima = sheet.getLastRow();
  if(ultima < 2) return [];
  const datos = sheet.getRange(1, 1, ultima, ENC_CUPONES.length).getValues();
  const enc = datos[0].map(h => String(h || "").trim());
  const out = [];
  for(let i = 1; i < datos.length; i++){
    const o = { _fila: i + 1 };
    ENC_CUPONES.forEach(k => { const j = enc.indexOf(k); o[k] = j === -1 ? "" : datos[i][j]; });
    o.CODIGO = normalizarCodigoCupon(o.CODIGO);
    if(o.CODIGO) out.push(o);
  }
  return out;
}

/** Día desde el que vale el cupón (a las 00:00), o null si no tiene. */
function _fechaDesdeCupon(v){
  if(!v) return null;
  const f = v instanceof Date ? new Date(v.getTime()) : new Date(String(v) + (/^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? "T00:00:00" : ""));
  if(isNaN(f)) return null;
  f.setHours(0, 0, 0, 0);
  return f;
}

function _fechaVenceCupon(v){
  if(!v) return null;
  const f = v instanceof Date ? new Date(v.getTime()) : new Date(String(v) + (/^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? "T00:00:00" : ""));
  if(isNaN(f)) return null;
  f.setHours(23, 59, 59, 999); // vale todo el día de vencimiento
  return f;
}

function _textoFechaCupon(v){
  const f = _fechaDesdeCupon(v);
  return f ? Utilities.formatDate(f, zonaNegocio(), "yyyy-MM-dd") : "";
}

function calcularDescuentoCupon(cupon, subtotal){
  const valor = Number(cupon.VALOR) || 0;
  let d = String(cupon.TIPO).toUpperCase() === "MONTO" ? valor : subtotal * valor / 100;
  d = Math.min(Math.max(d, 0), subtotal);
  return Math.round(d * 100) / 100;
}

function etiquetaCupon(cupon){
  const valor = Number(cupon.VALOR) || 0;
  return "Cupón " + cupon.CODIGO + " · " + (String(cupon.TIPO).toUpperCase() === "MONTO" ? "$" + valor.toLocaleString("es-AR") : valor + "%");
}

function _dniYaUsoCupon(codigo, dniNorm){
  if(!dniNorm) return false;
  const sheet = getOrCrearHojaCuponesUsos();
  const ultima = sheet.getLastRow();
  if(ultima < 2) return false;
  const filas = sheet.getRange(2, 2, ultima - 1, 2).getValues(); // CODIGO, DNI
  return filas.some(f => normalizarCodigoCupon(f[0]) === codigo && normalizarDni(f[1]) === dniNorm);
}

/**
 * ¿Se puede usar este cupón? Devuelve { ok, message, cupon, descuento, etiqueta }.
 * dniNorm es opcional (al validar desde el carrito todavía puede no estar).
 */
function evaluarCupon(codigo, subtotal, dniNorm){
  codigo = normalizarCodigoCupon(codigo);
  const cupon = _leerCupones().find(c => c.CODIGO === codigo);
  if(!cupon) return { ok:false, message:"El cupón \"" + codigo + "\" no existe. Revisá que esté bien escrito." };
  if(String(cupon.ACTIVO || "SI").toUpperCase() !== "SI") return { ok:false, message:"El cupón " + codigo + " no está disponible en este momento." };
  const desde = _fechaDesdeCupon(cupon.DESDE);
  if(desde && new Date() < desde){
    return { ok:false, message:"El cupón " + codigo + " empieza a valer el " + Utilities.formatDate(desde, zonaNegocio(), "dd/MM") + "." };
  }
  const vence = _fechaVenceCupon(cupon.VENCE);
  if(vence && new Date() > vence) return { ok:false, message:"El cupón " + codigo + " ya venció." };
  const max = Number(cupon.USOS_MAX) || 0, usados = Number(cupon.USOS) || 0;
  if(max > 0 && usados >= max) return { ok:false, message:"El cupón " + codigo + " ya se agotó." };
  const minimo = Number(cupon.MINIMO) || 0;
  if(minimo > 0 && Number(subtotal) < minimo) return { ok:false, message:"El cupón " + codigo + " es para compras desde $" + minimo.toLocaleString("es-AR") + "." };
  if(dniNorm && _dniYaUsoCupon(codigo, dniNorm)) return { ok:false, message:"Ya usaste el cupón " + codigo + " en otro pedido (es uno por cliente)." };
  return { ok:true, cupon: cupon, descuento: calcularDescuentoCupon(cupon, Number(subtotal) || 0), etiqueta: etiquetaCupon(cupon) };
}

/** Se llama DENTRO del lock de guardarPedido, después de escribir el pedido. */
function registrarUsoCupon(ev, dniNorm, pedidoId){
  const sheet = getOrCrearHojaCupones();
  const colUsos = ENC_CUPONES.indexOf("USOS") + 1;
  const enc = sheet.getRange(1, 1, 1, ENC_CUPONES.length).getValues()[0].map(h => String(h || "").trim());
  const col = enc.indexOf("USOS") + 1 || colUsos;
  sheet.getRange(ev.cupon._fila, col).setValue((Number(ev.cupon.USOS) || 0) + 1);
  getOrCrearHojaCuponesUsos().appendRow([new Date(), ev.cupon.CODIGO, dniNorm, pedidoId, ev.descuento]);
}

/** Público: el carrito pregunta si un código sirve y cuánto descuenta. */
function validarCuponPublico(p){
  // Freno a quien pruebe códigos al azar
  if(!permitirIntento("validar_cupon", 300, 600)){
    return jsonResponse({ success:false, message:"Demasiados intentos. Probá de nuevo en unos minutos." });
  }
  const codigo = normalizarCodigoCupon(p.codigo || p.cupon);
  if(!codigo) return jsonResponse({ success:false, message:"Escribí un código de cupón." });
  const ev = evaluarCupon(codigo, Number(p.subtotal) || 0, normalizarDni(p.dni));
  if(!ev.ok) return jsonResponse({ success:false, message: ev.message });
  return jsonResponse({
    success:true,
    codigo: ev.cupon.CODIGO,
    tipo: String(ev.cupon.TIPO).toUpperCase() === "MONTO" ? "MONTO" : "PORCENTAJE",
    valor: Number(ev.cupon.VALOR) || 0,
    minimo: Number(ev.cupon.MINIMO) || 0,
    descuento: ev.descuento,
    etiqueta: ev.etiqueta
  });
}

/** Admin: lista de cupones con sus usos. */
function listarCupones(){
  const cupones = _leerCupones().map(c => ({
    codigo: c.CODIGO,
    tipo: String(c.TIPO).toUpperCase() === "MONTO" ? "MONTO" : "PORCENTAJE",
    valor: Number(c.VALOR) || 0,
    usosMax: Number(c.USOS_MAX) || 0,
    usos: Number(c.USOS) || 0,
    vence: _textoFechaCupon(c.VENCE),
    desde: _textoFechaCupon(c.DESDE),
    programado: !!(_fechaDesdeCupon(c.DESDE) && new Date() < _fechaDesdeCupon(c.DESDE)),
    minimo: Number(c.MINIMO) || 0,
    activo: String(c.ACTIVO || "SI").toUpperCase() === "SI",
    nota: String(c.NOTA || ""),
    vencido: !!(_fechaVenceCupon(c.VENCE) && new Date() > _fechaVenceCupon(c.VENCE))
  }));
  return jsonResponse({ success:true, cupones: cupones });
}

/**
 * Admin: cuánto rindió cada cupón. Cruza CUPONES_USOS con PEDIDOS:
 * usos, clientes, monto vendido (total ya con descuento), monto
 * descontado y la lista de pedidos. Los pedidos CANCELADOS se listan
 * pero no suman al vendido ni al descontado.
 * Parámetros opcionales: desde / hasta (AAAA-MM-DD), codigo.
 */
function reporteCupones(p){
  p = p || {};
  const desde = /^\d{4}-\d{2}-\d{2}$/.test(String(p.desde || "")) ? new Date(p.desde + "T00:00:00") : null;
  const hasta = /^\d{4}-\d{2}-\d{2}$/.test(String(p.hasta || "")) ? new Date(p.hasta + "T23:59:59") : null;
  const soloCodigo = normalizarCodigoCupon(p.codigo);

  const aFecha = v => {
    if(v instanceof Date) return v;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  };

  // Pedidos por ID (solo las columnas que hacen falta)
  const pedidos = {};
  const hojaPed = getHoja("PEDIDOS");
  if(hojaPed && hojaPed.getLastRow() > 1){
    const datos = hojaPed.getDataRange().getValues();
    const enc = datos[0].map(h => String(h || "").trim().toUpperCase());
    const c = k => enc.indexOf(k);
    const iId = c("PEDIDO_ID"), iNom = c("CLIENTE") !== -1 ? c("CLIENTE") : c("NOMBRE"), iTot = c("TOTAL"), iEst = c("ESTADO"), iSub = c("SUBTOTAL");
    for(let i = 1; i < datos.length; i++){
      const id = String(iId === -1 ? "" : datos[i][iId]).trim();
      if(!id) continue;
      pedidos[id] = {
        nombre: iNom === -1 ? "" : String(datos[i][iNom] || ""),
        total: iTot === -1 ? 0 : Number(datos[i][iTot]) || 0,
        subtotal: iSub === -1 ? 0 : Number(datos[i][iSub]) || 0,
        estado: iEst === -1 ? "" : String(datos[i][iEst] || "")
      };
    }
  }

  const porCodigo = {};
  const asegurar = codigo => porCodigo[codigo] || (porCodigo[codigo] = {
    codigo: codigo, usos: 0, cancelados: 0, clientes: 0, vendido: 0, descontado: 0,
    ultimoUso: "", pedidos: [], _dnis: {}
  });

  _leerCupones().forEach(cup => { if(!soloCodigo || cup.CODIGO === soloCodigo) asegurar(cup.CODIGO); });

  const hojaUsos = getHoja("CUPONES_USOS");
  if(hojaUsos && hojaUsos.getLastRow() > 1){
    const datos = hojaUsos.getDataRange().getValues();
    const enc = datos[0].map(h => String(h || "").trim().toUpperCase());
    const c = k => enc.indexOf(k);
    const iF = c("FECHA"), iC = c("CODIGO"), iD = c("DNI"), iP = c("PEDIDO_ID"), iDes = c("DESCUENTO");
    for(let i = 1; i < datos.length; i++){
      const codigo = normalizarCodigoCupon(iC === -1 ? "" : datos[i][iC]);
      if(!codigo || (soloCodigo && codigo !== soloCodigo)) continue;
      const fecha = aFecha(iF === -1 ? "" : datos[i][iF]);
      if(desde && (!fecha || fecha < desde)) continue;
      if(hasta && (!fecha || fecha > hasta)) continue;
      const pedidoId = String(iP === -1 ? "" : datos[i][iP]).trim();
      const ped = pedidos[pedidoId] || null;
      const descuento = Number(iDes === -1 ? 0 : datos[i][iDes]) || 0;
      const dni = String(iD === -1 ? "" : datos[i][iD] || "");
      const cancelado = !!(ped && String(ped.estado).toUpperCase() === "CANCELADO");
      const r = asegurar(codigo);
      r.usos++;
      if(cancelado){
        r.cancelados++;
      }else{
        r.descontado += descuento;
        r.vendido += ped ? ped.total : 0;
      }
      if(dni && !r._dnis[dni]){ r._dnis[dni] = true; r.clientes++; }
      const fechaIso = fecha ? fecha.toISOString() : "";
      if(fechaIso > r.ultimoUso) r.ultimoUso = fechaIso;
      r.pedidos.push({
        fecha: fechaIso, pedidoId: pedidoId, dni: dni, descuento: descuento,
        nombre: ped ? ped.nombre : "", total: ped ? ped.total : 0,
        estado: ped ? ped.estado : "NO ENCONTRADO"
      });
    }
  }

  const cupones = Object.keys(porCodigo).map(k => {
    const r = porCodigo[k];
    delete r._dnis;
    r.pedidos.sort((a, b) => a.fecha < b.fecha ? 1 : -1);
    return r;
  }).sort((a, b) => b.usos - a.usos || (a.codigo < b.codigo ? -1 : 1));

  const totales = cupones.reduce((t, r) => {
    t.usos += r.usos; t.cancelados += r.cancelados; t.vendido += r.vendido; t.descontado += r.descontado;
    return t;
  }, { usos: 0, cancelados: 0, vendido: 0, descontado: 0 });

  return jsonResponse({ success: true, cupones: cupones, totales: totales });
}

/** Admin: crea o edita un cupón (codigoOriginal = el que se está editando). */
function guardarCupon(d){
  const codigo = normalizarCodigoCupon(d.codigo);
  const original = normalizarCodigoCupon(d.codigoOriginal);
  const tipo = String(d.tipo || "").toUpperCase() === "MONTO" ? "MONTO" : "PORCENTAJE";
  const valor = Number(d.valor);
  const usosMax = Math.floor(Number(d.usosMax));
  const minimo = Number(d.minimo || 0);
  const vence = String(d.vence || "").trim();
  const desde = String(d.desde || "").trim();

  if(codigo.length < 3) return jsonResponse({ success:false, message:"El código tiene que tener al menos 3 letras o números (sin espacios)." });
  if(!isFinite(valor) || valor <= 0) return jsonResponse({ success:false, message:"Poné cuánto descuenta el cupón." });
  if(tipo === "PORCENTAJE" && valor > 100) return jsonResponse({ success:false, message:"Un porcentaje no puede ser mayor a 100." });
  if(!isFinite(usosMax) || usosMax < 1) return jsonResponse({ success:false, message:"La cantidad de usos tiene que ser 1 o más." });
  if(!isFinite(minimo) || minimo < 0) return jsonResponse({ success:false, message:"La compra mínima no puede ser negativa." });
  if(vence && !/^\d{4}-\d{2}-\d{2}$/.test(vence)) return jsonResponse({ success:false, message:"La fecha de vencimiento no es válida." });
  if(desde && !/^\d{4}-\d{2}-\d{2}$/.test(desde)) return jsonResponse({ success:false, message:"La fecha de inicio no es válida." });
  if(desde && vence && desde > vence) return jsonResponse({ success:false, message:"La fecha de inicio no puede ser posterior al vencimiento." });

  const lock = LockService.getScriptLock();
  if(!lock.tryLock(15000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const sheet = getOrCrearHojaCupones();
    const cupones = _leerCupones();
    const existente = original ? cupones.find(c => c.CODIGO === original) : null;
    if(original && !existente) return jsonResponse({ success:false, message:"No se encontró el cupón " + original + " (¿lo borraron?)." });
    const choca = cupones.find(c => c.CODIGO === codigo && c !== existente);
    if(choca) return jsonResponse({ success:false, message:"Ya existe un cupón con el código " + codigo + "." });
    const usos = existente ? (Number(existente.USOS) || 0) : 0;
    if(existente && original !== codigo && usos > 0) return jsonResponse({ success:false, message:"Este cupón ya se usó: no se le puede cambiar el código (creá uno nuevo)." });
    const activo = d.activo === false || String(d.activo).toUpperCase() === "NO" ? "NO" : "SI";
    const fila = [codigo, tipo, valor, usosMax, usos, vence, minimo, activo, String(d.nota || "").slice(0, 200), existente ? existente.CREADO : new Date(), desde];
    if(existente) sheet.getRange(existente._fila, 1, 1, fila.length).setValues([fila]);
    else sheet.appendRow(fila);
  } finally {
    lock.releaseLock();
  }
  return jsonResponse({ success:true, codigo: codigo });
}

/** Admin: borra un cupón (el historial de usos queda en CUPONES_USOS). */
function eliminarCupon(d){
  const codigo = normalizarCodigoCupon(d.codigo);
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(15000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const c = _leerCupones().find(x => x.CODIGO === codigo);
    if(!c) return jsonResponse({ success:false, message:"No se encontró el cupón " + codigo + "." });
    getOrCrearHojaCupones().deleteRow(c._fila);
  } finally {
    lock.releaseLock();
  }
  return jsonResponse({ success:true });
}

/**
 * CODIGO -> { CODIGO, PRODUCTO, PRECIO, STOCK, UNIDADES_POR_CAJA, PRECIO_CAJA }
 * de los productos PUBLICADOS pedidos. Usa la caché del catálogo si está;
 * si no, lee de la hoja SOLO las columnas necesarias en una sola lectura
 * (en vez de toda la hoja + ordenar + armar el JSON del catálogo).
 */
function leerProductosParaPedido(codigos){
  const buscados = new Set(codigos.filter(Boolean));
  const mapa = {};

  const texto = cacheLeerTexto("catalogo_publico");
  if(texto){
    (JSON.parse(texto).productos || []).forEach(p => {
      const c = String(p.CODIGO || "").trim();
      if(buscados.has(c)) mapa[c] = p;
    });
    return mapa;
  }

  const sheet = getSheet();
  const ultimaFila = sheet.getLastRow();
  const ultimaCol = sheet.getLastColumn();
  if(ultimaFila < 2) return mapa;
  const enc = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(h => String(h || "").trim());
  const campos = ["CODIGO", "PRODUCTO", "PRECIO", "STOCK", "PUBLICADO", "UNIDADES_POR_CAJA", "PRECIO_CAJA"];
  const idx = {};
  campos.forEach(k => { idx[k] = enc.indexOf(k); });
  if(idx.CODIGO === -1 || idx.PRECIO === -1) throw new Error("Faltan columnas CODIGO/PRECIO en PRODUCTOS");
  const hasta = Math.max.apply(null, campos.map(k => idx[k])) + 1;
  const filas = sheet.getRange(2, 1, ultimaFila - 1, hasta).getValues();
  for(let i = 0; i < filas.length; i++){
    const f = filas[i];
    const c = String(f[idx.CODIGO] || "").trim();
    if(!buscados.has(c)) continue;
    if(idx.PUBLICADO !== -1 && String(f[idx.PUBLICADO]).toUpperCase() !== "SI") continue;
    const p = { CODIGO: f[idx.CODIGO] };
    campos.forEach(k => { if(k !== "PUBLICADO" && idx[k] !== -1) p[k] = f[idx[k]]; });
    mapa[c] = p;
  }
  return mapa;
}

function guardarPedido(data){

  const pedidosSheet = getHoja("PEDIDOS");
  const detalleSheet = getHoja("DETALLE_PEDIDOS");

  const idOperacion = String(data.idOperacion || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 64);

  // Idempotencia rápida (sin lock): si el cliente reintenta un pedido cuya
  // respuesta se perdió, se devuelve el pedido ya guardado.
  if(idOperacion){
    const previo = buscarPedidoPorIdOperacionRapido(pedidosSheet, idOperacion);
    if(previo){
      return jsonResponse({ success:true, pedidoId: previo, repetido:true });
    }
  }

  const transporteElegido = normalizarTextoTransporte(data.empresa);
  if(transporteElegido){
    const bloqueados = obtenerTransportesNoDisponibles();
    if(bloqueados.some(b => transporteElegido.indexOf(b) !== -1)){
      return jsonResponse({
        success:false,
        message:"La empresa de transporte elegida no está disponible. Por favor, seleccioná otra."
      });
    }
  }

  if(!data.origenPOS){

    // Pedido mínimo: se valida contra el valor VIGENTE en la hoja (no el
    // que el navegador tenga en caché) y se devuelve para que el cliente
    // se actualice.
    const minimoVigente = leerPedidoMinimoActual();
    // Se compara contra el TOTAL FINAL (ya con el descuento del cupón):
    // un cupón nunca puede dejar el pedido por debajo del mínimo.
    if(minimoVigente !== null && Number(data.total) < minimoVigente){
      const conCupon = Number(data.descuento || 0) > 0;
      const fmt = n => "$" + Number(n).toLocaleString("es-AR");
      return jsonResponse({
        success:false,
        pedidoMinimo: minimoVigente,
        message: conCupon
          ? "Con el descuento del cupón tu total queda en " + fmt(data.total) + " y el pedido mínimo es de " + fmt(minimoVigente) + ". Agregá " + fmt(minimoVigente - Number(data.total)) + " más en productos para continuar."
          : "El pedido mínimo es de " + fmt(minimoVigente) + ". Agregá más productos para continuar."
      });
    }

    // Stock real (solo catálogo web; el POS puede vender sin stock).
    // Si guardarPedidoPublico ya lo validó contra el catálogo, no se repite.
    if(!data._stockValidado && data.carrito && data.carrito.length > 0){
      const codigosCarrito = data.carrito.map(item => item.CODIGO);
      const alcanza = (mapa) => mapa && data.carrito.every(item => {
        const st = mapa[String(item.CODIGO).trim()];
        return st !== undefined && st >= Number(item.cantidad);
      });

      // Camino rápido: si la caché indica que alcanza el stock, se acepta.
      // Si no alcanza (o no hay caché), se relee la planilla para decidir.
      let stockPorCodigo = obtenerStockDesdeCacheCatalogo(codigosCarrito);
      if(!alcanza(stockPorCodigo)){
        stockPorCodigo = obtenerStockActualPorCodigo(codigosCarrito);
      }

      const faltantes = [];
      data.carrito.forEach(item => {
        const stockActual = stockPorCodigo[String(item.CODIGO).trim()];
        if(stockActual === undefined || stockActual < Number(item.cantidad)){
          faltantes.push({
            codigo: String(item.CODIGO),
            producto: item.PRODUCTO,
            pedido: Number(item.cantidad),
            disponible: Math.max(0, stockActual || 0)
          });
        }
      });

      if(faltantes.length > 0){
        const detalle = faltantes
          .map(f => `${f.producto} (pediste ${f.pedido}, disponible ${f.disponible})`)
          .join("; ");
        return jsonResponse({
          success:false,
          message:"Algunos productos ya no tienen el stock que pediste — es probable que se hayan vendido mientras tenías el carrito guardado. Actualizá la página y ajustá el pedido: " + detalle,
          stockInsuficiente: faltantes
        });
      }
    }
  }

  // Sección crítica: id único + escritura, bajo lock (evita filas pisadas
  // y IDs repetidos cuando entran varios pedidos a la vez).
  let lock = null;
  if(!_dentroDeLockIdempotencia){
    const _t0Lock = Date.now();
    lock = LockService.getScriptLock();
    if(!lock.tryLock(25000)){
      console.warn("guardarPedido: no se obtuvo el lock en 25 s");
      return jsonResponse({
        success:false,
        message:"El sistema está ocupado, probá de nuevo en unos segundos."
      });
    }
    const _espera = Date.now() - _t0Lock;
    if(_espera > 1000) console.warn("guardarPedido: esperó " + _espera + " ms el lock (otra operación lo tenía)");
  }

  let pedidoId;
  try{

    if(idOperacion){
      const previo = buscarPedidoPorIdOperacionRapido(pedidosSheet, idOperacion);
      if(previo){
        return jsonResponse({ success:true, pedidoId: previo, repetido:true });
      }
    }

    // Cupón: se vuelve a validar con los datos frescos (otro pedido pudo
    // haber usado el último cupo mientras este esperaba el lock).
    let cuponEnLock = null;
    if(data._cupon){
      const ev = evaluarCupon(data._cupon.codigo, Number(data.subtotal), normalizarDni(data.dni));
      if(!ev.ok){
        return jsonResponse({ success:false, cuponInvalido:true, message: ev.message });
      }
      if(Math.abs(ev.descuento - Number(data._cupon.descuento)) > 0.01){
        return jsonResponse({ success:false, cuponInvalido:true, message:"El cupón cambió mientras armabas el pedido. Volvé a aplicarlo." });
      }
      cuponEnLock = ev;
    }

    const ultimaFilaPedidos = pedidosSheet.getLastRow();
    pedidoId = generarNumeroPedido(pedidosSheet, ultimaFilaPedidos);
    const fecha = new Date();

    const valores = {
      PEDIDO_ID: pedidoId,
      FECHA: fecha,
      CLIENTE: data.nombre,
      EMPRESA: data.empresa,
      DIRECCION: data.direccion,
      TELEFONO: data.telefono,
      DNI: data.dni,
      TOTAL: data.total,
      ESTADO: "NUEVO",
      PDF_URL: "",
      CODIGO_POSTAL: data.codigoPostal || "",
      LOCALIDAD: data.localidad || "",
      PROVINCIA: data.provincia || "",
      STOCK_DESCONTADO: "NO",
      MONEDA: (data.moneda || "ARS").toString().toUpperCase(),
      SUBTOTAL: Number(data.subtotal || data.total || 0),
      DESCUENTO: Number(data.descuento || 0),
      DESCUENTO_ETIQUETA: data.descuentoEtiqueta || ""
    };
    if(data.moneda === "USD"){
      valores.TIPO_CAMBIO = Number(data.tipoCambio || 0);
    }
    if(idOperacion){
      valores.ID_OPERACION = idOperacion;
    }

    // Encabezados leídos UNA vez; las columnas que falten se crean juntas.
    const ultimaColumna = Math.max(pedidosSheet.getLastColumn(), 1);
    const encabezados = pedidosSheet.getRange(1, 1, 1, ultimaColumna).getValues()[0]
      .map(h => String(h || "").trim());

    const faltan = Object.keys(valores).filter(k => encabezados.indexOf(k) === -1);
    if(faltan.length > 0){
      pedidosSheet.getRange(1, encabezados.length + 1, 1, faltan.length).setValues([faltan]);
      faltan.forEach(k => encabezados.push(k));
    }

    const fila = encabezados.map(h => valores.hasOwnProperty(h) ? valores[h] : "");
    const filaDestino = ultimaFilaPedidos + 1;
    pedidosSheet.getRange(filaDestino, 1, 1, fila.length).setValues([fila]);

    if(data.carrito && data.carrito.length > 0){

      asegurarColumnaColorDetalleRapido(detalleSheet);

      const filas = data.carrito.map(item => [
        pedidoId,
        item.CODIGO,
        item.PRODUCTO,
        item.cantidad,
        item.PRECIO,
        item.PRECIO * item.cantidad,
        item.COLOR || ""
      ]);

      detalleSheet
        .getRange(detalleSheet.getLastRow() + 1, 1, filas.length, filas[0].length)
        .setValues(filas);
    }

    if(cuponEnLock) registrarUsoCupon(cuponEnLock, normalizarDni(data.dni), pedidoId);

    // Un reintento del mismo pedido (respuesta perdida) se resuelve con la caché.
    recordarPedidoPorIdOperacion(idOperacion, pedidoId);

  } finally {
    if(lock) lock.releaseLock();
  }

  // Registro automático del cliente (fuera del lock; solo lee la columna DNI).
  try {
    const dniNormalizado = normalizarDni(data.dni);
    if (dniNormalizado) {
      // Caché de DNIs conocidos (10 min): evita abrir la hoja CLIENTES en
      // cada pedido de un cliente ya registrado.
      let dnisCache = null;
      try{
        const t = CacheService.getScriptCache().get("clientes_dnis");
        if(t) dnisCache = JSON.parse(t);
      }catch(e){}
      let yaExiste = !!(dnisCache && dnisCache.indexOf(dniNormalizado) !== -1);
      if(!yaExiste && dnisCache){
        // La caché tiene la lista completa: no hace falta releer la hoja.
        dnisCache.push(dniNormalizado);
        try{ CacheService.getScriptCache().put("clientes_dnis", JSON.stringify(dnisCache), 600); }catch(e){}
      } else if(!yaExiste){
        const clientesSheet = getOrCrearHojaClientes();
        const ultima = clientesSheet.getLastRow();
        const lista = ultima >= 2
          ? clientesSheet.getRange(2, 4, ultima - 1, 1).getValues().map(f => normalizarDni(f[0]))
          : [];
        yaExiste = lista.indexOf(dniNormalizado) !== -1;
        if(!yaExiste) lista.push(dniNormalizado);
        try{ CacheService.getScriptCache().put("clientes_dnis", JSON.stringify(lista), 600); }catch(e){}
      }
      if (!yaExiste) {
        const clientesSheet = getOrCrearHojaClientes();
        const clienteId = generarIdCorto("CLI", clientesSheet);
        clientesSheet.appendRow([
          clienteId,
          String(data.nombre || "").trim(),
          "",
          dniNormalizado,
          String(data.telefono || "").trim(),
          String(data.direccion || "").trim(),
          String(data.empresa || "").trim(),
          "NO",
          new Date()
        ]);
      }
    }
  } catch(errCliente) {
    Logger.log("guardarPedido: no se pudo registrar cliente automático: " + errCliente.toString());
  }

  invalidarCachePolling();

  return jsonResponse({
    success: true,
    pedidoId: pedidoId,
    version: VERSION_BACKEND
  });

}

/**
 * Busca un pedido reciente (últimas 300 filas) por su ID_OPERACION.
 * Devuelve el PEDIDO_ID o null.
 */
/**
 * El catálogo la llama cuando el cliente abre el carrito para completar
 * sus datos: despierta el servidor y deja en caché todo lo que el pedido
 * va a necesitar (mínimo, transportes, DNIs conocidos), así el
 * "Enviar pedido" no paga esas lecturas. Si ya está todo en caché, solo
 * hace lecturas de caché (no toca la planilla).
 */
function prepararPedido(){
  // (No se toca el catálogo: rearmarlo es lo más pesado y el pedido ya no lo necesita.)
  try{ leerReglasPedido(); }catch(e){}
  try{
    const cache = CacheService.getScriptCache();
    if(!cache.get("clientes_dnis")){
      const hoja = getOrCrearHojaClientes();
      const ultima = hoja.getLastRow();
      const lista = ultima >= 2 ? hoja.getRange(2, 4, ultima - 1, 1).getValues().map(f => normalizarDni(f[0])).filter(Boolean) : [];
      cache.put("clientes_dnis", JSON.stringify(lista), 600);
    }
  }catch(e){}
  return jsonResponse({ success:true, version: VERSION_BACKEND });
}

/** Borra la lista de DNIs conocidos (la usa el alta automática de clientes en los pedidos web). */
function invalidarCacheDnisClientes(){
  try{ CacheService.getScriptCache().remove("clientes_dnis"); }catch(e){}
}

const PREFIJO_CACHE_OP_PEDIDO = "pedido_op:";

/** Recuerda (6 h) qué pedido se creó con este idOperacion. */
function recordarPedidoPorIdOperacion(idOperacion, pedidoId){
  if(!idOperacion) return;
  try{ CacheService.getScriptCache().put(PREFIJO_CACHE_OP_PEDIDO + idOperacion, String(pedidoId), 21600); }catch(e){}
}

/**
 * Busca el pedido de un idOperacion: primero en la caché (una sola
 * llamada, lo normal) y solo si la caché no responde, en la hoja.
 * Devuelve el PEDIDO_ID o null.
 */
function buscarPedidoPorIdOperacionRapido(pedidosSheet, idOperacion){
  try{
    const enCache = CacheService.getScriptCache().get(PREFIJO_CACHE_OP_PEDIDO + idOperacion);
    return enCache ? String(enCache) : null;
  }catch(e){
    const hoja = pedidosSheet || getHoja("PEDIDOS");
    return hoja ? buscarPedidoPorIdOperacion(hoja, idOperacion) : null;
  }
}

function buscarPedidoPorIdOperacion(pedidosSheet, idOperacion){
  const ultimaFila = pedidosSheet.getLastRow();
  const ultimaCol = pedidosSheet.getLastColumn();
  if(ultimaFila < 2 || ultimaCol < 1) return null;

  const encabezados = pedidosSheet.getRange(1, 1, 1, ultimaCol).getValues()[0]
    .map(h => String(h || "").trim());
  const colOp = encabezados.indexOf("ID_OPERACION");
  const colId = encabezados.indexOf("PEDIDO_ID");
  if(colOp === -1 || colId === -1) return null;

  const desde = Math.max(2, ultimaFila - 299);
  const n = ultimaFila - desde + 1;
  const ops = pedidosSheet.getRange(desde, colOp + 1, n, 1).getValues();
  for(let i = ops.length - 1; i >= 0; i--){
    if(String(ops[i][0]) === idOperacion){
      return String(pedidosSheet.getRange(desde + i, colId + 1).getValue());
    }
  }
  return null;
}

/**
 * Crea un pedido desde el panel admin (POS en "modo pedido"), con el
 * mismo formato de carrito que ya usa el POS para ventas normales:
 * [{ CODIGO, PRODUCTO, PRECIO, cantidad }, ...].
 *
 * A propósito, NO descuenta stock — un pedido es una promesa de venta
 * futura, no una venta concretada. El stock se ajusta más adelante,
 * cuando el pedido se entrega (a mano, o desde una venta real si en
 * el futuro se conecta ese flujo).
 *
 * Reutiliza guardarPedido() para no duplicar la lógica de guardado en
 * PEDIDOS/DETALLE_PEDIDOS — solo adapta el formato de carrito del POS
 * (cantidad en minúscula, igual que el que ya usa guardarVenta) al
 * que espera esa función (que es el mismo formato, así que no hace
 * falta transformar nada).
 */
function guardarPedidoAdmin(e){

  // Acepta tanto e.parameter (si llega por GET, como antes) como un
  // objeto plano ya parseado (si llega por POST, con el carrito en el
  // body en vez de la URL — necesario para pedidos grandes, ver nota
  // en el dispatcher de doPost).
  const params = (e && e.parameter) || e || {};

  const nombre = (params.nombre || "").toString().trim();
  if(!nombre){
    return jsonResponse({ success:false, message:"El nombre del cliente es obligatorio." });
  }

  const direccion = (params.direccion || "").toString().trim();
  if(!direccion){
    return jsonResponse({ success:false, message:"La dirección es obligatoria." });
  }

  const localidad = (params.localidad || "").toString().trim();
  if(!localidad){
    return jsonResponse({ success:false, message:"La localidad es obligatoria." });
  }

  const provincia = (params.provincia || "").toString().trim();
  if(!provincia){
    return jsonResponse({ success:false, message:"La provincia es obligatoria." });
  }

  // El carrito llega como string JSON si el pedido vino por GET (URL,
  // como antes), o ya como array si vino por POST con body (necesario
  // para pedidos grandes — ver nota en el dispatcher de doPost).
  const carrito = Array.isArray(params.carrito) ? params.carrito : JSON.parse(params.carrito || "[]");
  if(carrito.length === 0){
    return jsonResponse({ success:false, message:"El pedido no tiene productos." });
  }

  const total = Number(params.total || 0);

  const data = {
    nombre: nombre,
    empresa: (params.empresa || "").toString().trim(),
    direccion: direccion,
    telefono: (params.telefono || "").toString().trim(),
    dni: (params.dni || "").toString().trim(),
    total: total,
    subtotal: Number(params.subtotal || total),
    descuento: Number(params.descuento || 0),
    descuentoEtiqueta: (params.descuentoEtiqueta || "").toString().trim(),
    codigoPostal: (params.codigoPostal || "").toString().trim(),
    localidad: localidad,
    provincia: provincia,
    carrito: carrito,
    moneda: (params.moneda || "ARS").toString().toUpperCase(),
    tipoCambio: Number(params.tipoCambio || 0),
    // Marca que este pedido viene del POS (no del catálogo web), para
    // que guardarPedido() salte la validación de stock — desde el POS
    // siempre se debe poder cargar el pedido aunque falte stock.
    origenPOS: true
  };

  return guardarPedido(data);
}

function testGuardarPedido() {

const data = {
nombre: "Cliente Prueba",
empresa: "Empresa Prueba",
direccion: "Calle Falsa 123",
dni: "12345678",
total: 10000,
carrito: [
{
CODIGO: "TEST001",
PRODUCTO: "Producto Test",
cantidad: 2,
PRECIO: 5000
}
]
};

const resultado = guardarPedido(data);

Logger.log(resultado.getContent());

}
/**
 * Convierte un código de producto a texto plano, sin notación
 * científica — puede llegar como número (si Sheets interpretó el
 * código de barras como número en vez de texto) o ya como un string
 * tipo "6.934112480112E12" (lo que Sheets muestra en pantalla para un
 * número así de largo). En cualquiera de los dos casos, se reconstruye
 * el número completo con toFixed(0), que nunca usa notación
 * científica para números de este tamaño.
 */
function formatearCodigoComoTexto(valor){
  if(valor === null || valor === undefined || valor === "") return "";

  const texto = String(valor).trim();
  const pareceNotacionCientifica = /e\+?\d+/i.test(texto);
  const numero = Number(texto);

  if(pareceNotacionCientifica && !isNaN(numero)){
    return numero.toFixed(0);
  }

  return texto;
}

function generarPDFPedido(
  pedidoId,
  fecha,
  data
){

  // La carpeta es configurable desde Configuración → Imágenes de
  // productos (clave DRIVE_CARPETA_PEDIDOS_ID). Si nunca se configuró,
  // no se usa ninguna carpeta por defecto.
  const carpetaId = extraerIdCarpetaDrive(
    leerClaveConfiguracion("DRIVE_CARPETA_PEDIDOS_ID", "")
  );

  // El encabezado y pie del PDF de pedidos toman el nombre del negocio
  // y el WhatsApp ya configurados en otras partes del panel (no se
  // duplican esos datos en una clave aparte), y dos textos propios de
  // este PDF que sí son editables desde Configuración → PDF de pedidos.
  const nombreNegocio    = leerClaveConfiguracion("NOMBRE", "MI NEGOCIO");
  const whatsappNegocio  = leerClaveConfiguracion("BENEFICIO_WHATSAPP_NUMERO", "");
  const subtituloPdf     = leerClaveConfiguracion("PEDIDO_PDF_SUBTITULO", "Distribuidora Mayorista");
  const piePdf           = leerClaveConfiguracion("PEDIDO_PDF_PIE", "Gracias por su compra.");

  const doc = DocumentApp.create(
    "Pedido " + pedidoId
  );

  const body = doc.getBody();

  // ENCABEZADO

  const titulo =
  body.appendParagraph(
    nombreNegocio
  );

  titulo
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    )
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    );

  body.appendParagraph(
    subtituloPdf
  )
  .setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );

  body.appendParagraph(
    "WhatsApp: " + whatsappNegocio
  )
  .setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );

  body.appendHorizontalRule();

  // DATOS DEL PEDIDO

  body.appendParagraph(
    "PEDIDO"
  )
  .setBold(true);

  body.appendParagraph(
    "Número: " + pedidoId
  );

  body.appendParagraph(
    "Fecha: " +
    Utilities.formatDate(
      fecha,
      zonaNegocio(),
      "dd/MM/yyyy HH:mm"
    )
  );

  body.appendHorizontalRule();

  // CLIENTE

  body.appendParagraph(
    "DATOS DEL CLIENTE"
  )
  .setBold(true);

  body.appendParagraph(
    "Cliente: " + data.nombre
  );

  body.appendParagraph(
    "Transporte: " +
    (data.empresa || "-")
  );

  body.appendParagraph(
    "Dirección: " +
    data.direccion
  );

  body.appendParagraph(
    "Localidad: " +
    (data.localidad || "-") +
    "  -  Provincia: " +
    (data.provincia || "-") +
    (data.codigoPostal ? "  -  CP: " + data.codigoPostal : "")
  );

  body.appendParagraph(
    "Teléfono: " +
    (data.telefono || "-")
  );

  body.appendParagraph(
    "DNI: " +
    data.dni
  );

  body.appendHorizontalRule();

  // TABLA

const tabla = [
  [
    "Código",
    "Producto",
    "Cant",
    "$",
    "Subt"
  ]
];

data.carrito.forEach(item => {

  const subtotal =
  item.PRECIO * item.cantidad;

  const nombreConColor =
    item.PRODUCTO + (item.COLOR ? " (Color: " + item.COLOR + ")" : "");

 tabla.push([
  formatearCodigoComoTexto(item.CODIGO),
  nombreConColor,
  String(parseInt(item.cantidad)),
  "$" + Number(item.PRECIO)
          .toLocaleString("es-AR"),
  "$" + Number(subtotal)
          .toLocaleString("es-AR")
]);
});

const table =
body.appendTable(tabla);

table.setBorderWidth(0);

// Ancho de columnas — 5 columnas: Código, Producto, Cant, $, Subtotal
// (el total suma igual que antes, 515pt, solo redistribuido para que
// entre la columna nueva sin angostar de más a "Producto")

try {

  table.getRow(0)
       .getCell(0)
       .setWidth(90); // Código

  table.getRow(0)
       .getCell(1)
       .setWidth(230); // Producto

  table.getRow(0)
       .getCell(2)
       .setWidth(40); // Cant

  table.getRow(0)
       .getCell(3)
       .setWidth(65); // $

  table.getRow(0)
       .getCell(4)
       .setWidth(80); // Subtotal

} catch(err) {
  Logger.log(err);
}

// Fuente: encabezado compacto, productos más legibles

for (let r = 0; r < table.getNumRows(); r++) {

  const row = table.getRow(r);
  const size = r === 0 ? 7 : 10; // encabezado 7pt, filas de producto 10pt

  for (let c = 0; c < row.getNumCells(); c++) {

    row.getCell(c)
       .editAsText()
       .setFontSize(size);

  }

}

// Encabezados en negrita

for (let c = 0; c < 4; c++) {

  table.getRow(0)
       .getCell(c)
       .editAsText()
       .setBold(true);

}
  // TOTAL

  body.appendParagraph("TOTAL DEL PEDIDO").setBold(true);

  if(Number(data.descuento || 0) > 0){

    const subtotalBase = Number(data.subtotal || data.total) || 0;
    const descuentoMonto = Number(data.descuento) || 0;
    const porcentaje = subtotalBase > 0 ? (descuentoMonto / subtotalBase) * 100 : 0;
    const porcentajeTexto = (Math.round(porcentaje * 10) / 10).toString().replace(".", ",") + "%";

    body.appendParagraph(
      "Total sin descuento: $" + subtotalBase.toLocaleString("es-AR")
    ).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

    const etiqueta = data.descuentoEtiqueta
      ? " (" + data.descuentoEtiqueta + ")"
      : "";

    body.appendParagraph(
      "Descuento" + etiqueta + " (" + porcentajeTexto + "): -$" + descuentoMonto.toLocaleString("es-AR")
    ).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  }

  body.appendParagraph(
    "TOTAL: $" + Number(data.total).toLocaleString("es-AR")
  )
  .setBold(true)
  .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph("");

  body.appendParagraph("");

  body.appendHorizontalRule();

  // PIE

  body.appendParagraph(
    piePdf
  )
  .setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );

  body.appendParagraph(
    nombreNegocio
  )
  .setBold(true)
  .setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );

  doc.saveAndClose();

  // PDF

  const pdfBlob =
  DriveApp
    .getFileById(doc.getId())
    .getAs(MimeType.PDF);

  pdfBlob.setName(
    pedidoId + ".pdf"
  );

  const carpeta =
  DriveApp.getFolderById(
    carpetaId
  );

  const pdfFile =
  carpeta.createFile(
    pdfBlob
  );

  DriveApp
    .getFileById(doc.getId())
    .setTrashed(true);

  return pdfFile.getUrl();

}
function getPedidos(){
  return conCachePolling("pedidos", 10, getPedidosImpl);
}

// Estados que todavía necesitan que alguien haga algo — estos se
// devuelven SIEMPRE, sin importar cuándo se cargó el pedido, porque
// un pedido pendiente nunca debe "desaparecer" de la lista solo por
// ser viejo.
const ESTADOS_PEDIDO_EN_CURSO = ["NUEVO", "PREPARANDO"];

// Pedidos ya resueltos (ENVIADO, CANCELADO, o cualquier otro estado
// que no esté en ESTADOS_PEDIDO_EN_CURSO): no hace falta traerlos
// todos en cada refresco del polling — con la operación del día a
// día alcanza y sobra con los últimos 7 días. Si alguna vez hace
// falta ver uno más viejo puntual, se puede seguir abriendo por
// número (getDetallePedido / consultarPedido siguen sin este límite,
// buscan por ID directo).
const DIAS_HISTORIAL_PEDIDOS_FINALIZADOS = 30;

function getPedidosImpl(){
  // Lee SOLO la hoja PEDIDOS (sin DETALLE_PEDIDOS). Antes reutilizaba
  // obtenerIndicePedidos(), que además arma y cachea el detalle de TODOS
  // los pedidos — con meses de historial eso solo podía tardar más que
  // el timeout del panel (25s), y el JSON resultante a veces ni entraba
  // en CacheService, así que se recalculaba en cada consulta.
  const sheet = getHoja("PEDIDOS");
  if(!sheet || sheet.getLastRow() < 2){
    return jsonResponse({ success: true, pedidos: [] });
  }

  const datos = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const encabezados = datos.shift().map(h => String(h || "").trim());
  const idxEstado = encabezados.indexOf("ESTADO");
  const idxFecha = encabezados.indexOf("FECHA");
  const idxId = encabezados.indexOf("PEDIDO_ID");

  const corteFinalizados = new Date();
  corteFinalizados.setDate(corteFinalizados.getDate() - DIAS_HISTORIAL_PEDIDOS_FINALIZADOS);

  const pedidos = [];
  for(let i = 0; i < datos.length; i++){
    const fila = datos[i];
    if(idxId !== -1 && String(fila[idxId] || "").trim() === "") continue;
    const estado = idxEstado !== -1 ? String(fila[idxEstado]) : "";
    if(!ESTADOS_PEDIDO_EN_CURSO.includes(estado)){
      // ENVIADO, CANCELADO, ...: solo los de los últimos N días
      const fecha = new Date(idxFecha !== -1 ? fila[idxFecha] : "");
      if(isNaN(fecha) || fecha < corteFinalizados) continue;
    }
    const obj = {};
    encabezados.forEach((h, j) => { if(h) obj[h] = fila[j]; });
    // Siempre texto: la planilla guarda los números nuevos (1001…) como
    // número, y el panel los usa como texto (rótulo, búsquedas, etc.)
    if(obj.PEDIDO_ID !== undefined) obj.PEDIDO_ID = String(obj.PEDIDO_ID);
    pedidos.push(obj);
  }

  return jsonResponse({
    success: true,
    pedidos: pedidos
  });
}

/**
 * Busca en TODA la hoja PEDIDOS (también los resueltos de hace más de
 * DIAS_HISTORIAL_PEDIDOS_FINALIZADOS días, que la lista general no trae).
 * q: número de pedido, cliente, DNI/CUIT, teléfono, transporte o localidad
 *    (sin importar mayúsculas ni acentos; DNI y teléfono también solo con números).
 * desde / hasta (AAAA-MM-DD, opcionales): rango de fechas.
 * Devuelve los más nuevos primero, hasta 100.
 */
function buscarPedidos(p){
  const q = String(p.q || "").trim();
  const desde = /^\d{4}-\d{2}-\d{2}$/.test(String(p.desde || "")) ? new Date(p.desde + "T00:00:00") : null;
  const hasta = /^\d{4}-\d{2}-\d{2}$/.test(String(p.hasta || "")) ? new Date(p.hasta + "T23:59:59") : null;
  if(q.length < 2 && !desde && !hasta){
    return jsonResponse({ success:false, message:"Escribí al menos 2 letras o números, o elegí fechas." });
  }
  const sheet = getHoja("PEDIDOS");
  if(!sheet || sheet.getLastRow() < 2) return jsonResponse({ success:true, pedidos:[], total:0 });

  const sinAcentos = t => String(t || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const texto = sinAcentos(q);
  const digitos = q.replace(/\D/g, "");
  const datos = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const enc = datos.shift().map(h => String(h || "").trim());
  const col = k => enc.indexOf(k);
  const iId = col("PEDIDO_ID"), iFecha = col("FECHA");
  const camposTexto = ["PEDIDO_ID", "CLIENTE", "NOMBRE", "EMPRESA", "LOCALIDAD", "PROVINCIA", "DNI", "TELEFONO"].map(col).filter(i => i !== -1);
  const camposNum = ["DNI", "TELEFONO"].map(col).filter(i => i !== -1);

  const encontrados = [];
  for(let i = 0; i < datos.length; i++){
    const f = datos[i];
    if(iId !== -1 && String(f[iId] || "").trim() === "") continue;
    if(desde || hasta){
      const fecha = new Date(iFecha !== -1 ? f[iFecha] : "");
      if(isNaN(fecha) || (desde && fecha < desde) || (hasta && fecha > hasta)) continue;
    }
    if(texto){
      const ok = camposTexto.some(j => sinAcentos(f[j]).indexOf(texto) !== -1) ||
        (digitos.length >= 4 && camposNum.some(j => String(f[j] || "").replace(/\D/g, "").indexOf(digitos) !== -1));
      if(!ok) continue;
    }
    encontrados.push(f);
  }
  encontrados.sort((a, b) => new Date(b[iFecha]) - new Date(a[iFecha]));
  const LIMITE = 100;
  const pedidos = encontrados.slice(0, LIMITE).map(f => {
    const o = {};
    enc.forEach((h, j) => { if(h) o[h] = f[j]; });
    if(o.PEDIDO_ID !== undefined) o.PEDIDO_ID = String(o.PEDIDO_ID);
    return o;
  });
  return jsonResponse({ success:true, pedidos: pedidos, total: encontrados.length, limitado: encontrados.length > LIMITE });
}

function testPDF(){
  generarPDFPedido(
    "TEST-001",
    new Date(),
    {
      nombre:"Cliente Prueba",
      empresa:"Empresa Prueba",
      direccion:"Prueba",
      telefono:"123456",
      dni:"12345678",
      total:10000,
      carrito:[
        {
          CODIGO:"A001",
          PRODUCTO:"Producto Test",
          cantidad:2,
          PRECIO:5000
        }
      ]
    }
  );
}
/* =====================================================================
   CREDENCIALES DE ACCESO AL ADMIN
   Hoja separada "ADMIN_CREDENCIALES" (NO la hoja CONFIGURACION, que es
   pública de lectura para el catálogo web). Las contraseñas se guardan
   con hash (ver SEGURIDAD más arriba); las que estén en texto plano de
   versiones anteriores se convierten solas en el próximo login.
===================================================================== */

const CREDENCIALES_DEFAULT = {
  USUARIO: "admin",
  PASSWORD: "admin2026",
  // Contraseña propia para stock.html — a propósito distinta de la
  // del panel completo, así se le puede dar a alguien de depósito sin
  // darle acceso a todo lo demás (ventas, clientes, configuración...).
  STOCK_PASSWORD: "deposito2026"
};

function getOrCrearHojaCredenciales(){

  const ss = getSS();
  let sheet = getHoja("ADMIN_CREDENCIALES");

  if(!sheet){
    sheet = ss.insertSheet("ADMIN_CREDENCIALES");
    sheet.appendRow(["CLAVE", "VALOR"]);
    sheet.appendRow(["USUARIO", CREDENCIALES_DEFAULT.USUARIO]);
    sheet.appendRow(["PASSWORD", CREDENCIALES_DEFAULT.PASSWORD]);
    sheet.appendRow(["STOCK_PASSWORD", CREDENCIALES_DEFAULT.STOCK_PASSWORD]);
    // La hoja de credenciales no necesita estar visible para operar el
    // negocio día a día; se oculta para que no se edite por error desde
    // la planilla. Se puede mostrar de nuevo desde Sheets si hace falta.
    sheet.hideSheet();
    _hojaCache["ADMIN_CREDENCIALES"] = sheet;
  }

  return sheet;
}

function leerCredencialesAdmin(){

  const sheet = getOrCrearHojaCredenciales();
  const data = sheet.getDataRange().getValues();
  data.shift(); // encabezados CLAVE | VALOR

  const creds = { ...CREDENCIALES_DEFAULT };

  data.forEach(row => {
    const clave = String(row[0] || "").trim().toUpperCase();
    if(clave && creds.hasOwnProperty(clave)){
      creds[clave] = row[1] === undefined || row[1] === null ? "" : String(row[1]);
    }
  });

  return creds;
}

/** Escribe (o agrega) una clave de ADMIN_CREDENCIALES. */
function escribirCredencial(clave, valor){
  const sheet = getOrCrearHojaCredenciales();
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0] || "").trim().toUpperCase() === clave){
      sheet.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  sheet.appendRow([clave, valor]);
}

/** true si la contraseña del admin sigue siendo la de fábrica. */
function adminConPasswordDeFabrica(creds){
  return verificarPassword(CREDENCIALES_DEFAULT.PASSWORD, creds.PASSWORD);
}

function login(e){

  // Primera vez que alguien entra al panel de un negocio nuevo: se crean
  // solas las hojas y los triggers (antes había que ejecutarlo a mano).
  asegurarInstalacionInicial();

  const params = (e && e.parameter) || {};
  const usuario = String(params.usuario || "").trim();
  const password = String(params.password || "");
  // "dispositivo=1": app de escritorio o app de Reportes en el celular,
  // donde la sesión dura 30 días (y en la app de escritorio, además,
  // permite que la cola offline siga sincronizando).
  const ttl = String(params.dispositivo || "") === "1" ? SESION_TTL_DISPOSITIVO_SEG : SESION_TTL_PANEL_SEG;

  // Bloqueo temporal tras varios intentos fallidos del mismo usuario
  // (frena la fuerza bruta; se libera solo a los 15 minutos).
  const claveIntentos = "login_fallos:" + usuario.toLowerCase().slice(0, 60);
  const fallosPrevios = Number(CacheService.getScriptCache().get(claveIntentos) || 0);
  if(fallosPrevios >= 8 || !permitirIntento("login_global", 300, 600)){
    return jsonResponse({ success:false, bloqueado:true, message:"Demasiados intentos. Esperá 15 minutos." });
  }

  if(usuario && password){
    const creds = leerCredencialesAdmin();

    // 1) Usuario/contraseña del admin "clásico" — siempre rol admin.
    if(usuario === creds.USUARIO && verificarPassword(password, creds.PASSWORD)){
      const deFabrica = password === CREDENCIALES_DEFAULT.PASSWORD;
      if(!esHashPassword(creds.PASSWORD)) escribirCredencial("PASSWORD", hashPassword(password));
      try{ CacheService.getScriptCache().remove(claveIntentos); }catch(err){}
      return jsonResponse({
        success: true,
        rol: "admin",
        nombre: "Administrador",
        usuario: creds.USUARIO,
        token: crearTokenSesion({ usuario: creds.USUARIO, rol: "admin", adminClasico: true }, ttl),
        venceEn: Date.now() + ttl * 1000,
        debeCambiarPassword: deFabrica,
        modo: modoNegocio()
      });
    }

    // 2) Usuarios cargados en la hoja USUARIOS, cada uno con su rol.
    const usuarioEncontrado = buscarUsuarioPorCredenciales(usuario, password);
    if(usuarioEncontrado){
      try{ CacheService.getScriptCache().remove(claveIntentos); }catch(err){}
      const rol = String(usuarioEncontrado.ROL || "vendedor").trim().toLowerCase();
      return jsonResponse({
        success: true,
        rol: rol,
        nombre: usuarioEncontrado.NOMBRE,
        usuario: usuarioEncontrado.USERNAME,
        token: crearTokenSesion({ usuario: usuarioEncontrado.USERNAME, rol: rol }, ttl),
        venceEn: Date.now() + ttl * 1000,
        modo: modoNegocio()
      });
    }
  }

  try{
    CacheService.getScriptCache().put(claveIntentos, String(fallosPrevios + 1), 900);
  }catch(err){}

  return jsonResponse({ success:false, message:"Usuario o contraseña incorrectos." });

}


/* =====================================================================
   MODO DEPÓSITO
   ---------------------------------------------------------------------
   El mismo código sirve para el backend del local y para el del
   depósito (otra planilla, otro Apps Script). El del depósito se marca
   UNA vez ejecutando configurarComoDeposito() desde el editor; el panel
   y la app de escritorio cambian lo que muestran según este modo.
===================================================================== */

function modoNegocio(){
  try{
    return String(leerClaveConfiguracion("MODO_NEGOCIO", "") || "").trim().toUpperCase() === "DEPOSITO" ? "DEPOSITO" : "LOCAL";
  }catch(e){
    return "LOCAL";
  }
}

function esModoDeposito(){ return modoNegocio() === "DEPOSITO"; }

/** Ejecutar UNA vez desde el editor de Apps Script del backend del DEPÓSITO. */
function configurarComoDeposito(){
  asegurarInstalacionInicial();
  fijarClaveConfiguracion("MODO_NEGOCIO", "DEPOSITO");
  _invalidarCacheConfigNegocio();
  try{ invalidarCacheProductos(); }catch(e){} // por si se cargan productos a mano en la hoja justo después
  Logger.log("Listo: este backend quedó como DEPÓSITO. Usuario y contraseña iniciales: admin / admin2026 (cambialos desde el panel; tienen que ser distintos de los del local).");
}

/** Por si se marcó por error: vuelve a modo local. */
function configurarComoLocal(){
  fijarClaveConfiguracion("MODO_NEGOCIO", "LOCAL");
  _invalidarCacheConfigNegocio();
  Logger.log("Listo: este backend quedó como LOCAL.");
}

/**
 * Login de stock.html — a propósito una contraseña propia, distinta
 * de la del panel completo (sin usuario, para que sea más rápido de
 * escribir desde el celular en el depósito). Devuelve un token que
 * SOLO sirve para ver/editar productos y stock (ACCIONES_ROL_STOCK).
 */
function loginStock(e){

  const password = String(((e && e.parameter) || {}).password || "");

  if(!permitirIntento("login_stock", 30, 900)){
    return jsonResponse({ success:false, bloqueado:true, message:"Demasiados intentos. Esperá 15 minutos." });
  }

  const creds = leerCredencialesAdmin();
  const ok = verificarPassword(password, creds.STOCK_PASSWORD);

  if(!ok) return jsonResponse({ success:false, message:"Contraseña incorrecta." });

  if(!esHashPassword(creds.STOCK_PASSWORD)) escribirCredencial("STOCK_PASSWORD", hashPassword(password));

  return jsonResponse({
    success: true,
    token: crearTokenSesion({ usuario: "stock", rol: "deposito", tipo: "stock" }, SESION_TTL_PANEL_SEG),
    venceEn: Date.now() + SESION_TTL_PANEL_SEG * 1000
  });

}

/**
 * Cambia la contraseña de stock.html — se pide la contraseña ACTUAL
 * del panel admin para autorizar el cambio. Cierra las sesiones de
 * stock abiertas (el celular del depósito va a pedir la nueva).
 */
function guardarPasswordStock(e){

  const params = (e && e.parameter) || {};

  const passwordActualAdmin = (params.passwordActualAdmin || "").toString();
  const nuevaPasswordStock = (params.nuevaPasswordStock || "").toString();

  const creds = leerCredencialesAdmin();

  if(!verificarPassword(passwordActualAdmin, creds.PASSWORD)){
    return jsonResponse({
      success:false,
      message:"La contraseña del panel admin no es correcta."
    });
  }

  if(!nuevaPasswordStock || nuevaPasswordStock.length < 4){
    return jsonResponse({
      success:false,
      message:"La nueva contraseña de stock debe tener al menos 4 caracteres."
    });
  }

  escribirCredencial("STOCK_PASSWORD", hashPassword(nuevaPasswordStock));
  cerrarSesionesTipo("stock");

  return jsonResponse({ success:true });
}

/* =====================================================================
   USUARIOS Y ROLES
   Hoja "USUARIOS": permite crear varios accesos al panel, cada uno con
   un rol que determina qué secciones puede ver/usar. El usuario/
   contraseña "clásico" (ADMIN_CREDENCIALES) sigue funcionando igual,
   siempre con rol "admin" — esto es un agregado, no un reemplazo.

   Roles disponibles:
   - admin:     acceso total (incluye esta pantalla de Usuarios).
   - vendedor:  Punto de Venta, Ventas POS, Cierre de Caja, Pedidos,
                Clientes. Sin Configuración ni Usuarios.
   - deposito:  Productos, Ingreso de Productos, Reportes de Compras.
                Sin ventas, clientes ni configuración.
===================================================================== */

const ROLES_VALIDOS = ["admin", "vendedor", "deposito"];

function getOrCrearHojaUsuarios(){
  const ss = getSS();
  let sheet = getHoja("USUARIOS");
  if(!sheet){
    sheet = ss.insertSheet("USUARIOS");
    sheet.appendRow(["USUARIO_ID", "NOMBRE", "USERNAME", "PASSWORD", "ROL", "ACTIVO"]);
    sheet.setFrozenRows(1);
    _hojaCache["USUARIOS"] = sheet;
  }
  return sheet;
}

function buscarUsuarioPorCredenciales(username, password){
  if(!username || !password) return null;
  const sheet = getOrCrearHojaUsuarios();
  const filas = sheet.getDataRange().getValues();
  const encabezados = filas[0].map(h => String(h || "").trim());
  const idxPassword = encabezados.indexOf("PASSWORD");
  const data = rowsToObjects(filas);
  const buscado = String(username).trim().toLowerCase();

  for(let i = 0; i < data.length; i++){
    const u = data[i];
    if(String(u.USERNAME || "").trim().toLowerCase() !== buscado) continue;
    if(String(u.ACTIVO || "SI").toUpperCase() === "NO") continue;
    if(!verificarPassword(password, u.PASSWORD)) continue;
    // Migración: contraseña vieja en texto plano -> se guarda con hash.
    if(!esHashPassword(u.PASSWORD) && idxPassword !== -1){
      try{ sheet.getRange(i + 2, idxPassword + 1).setValue(hashPassword(password)); }catch(err){}
    }
    return u;
  }
  return null;
}

/** Lista de usuarios para la pantalla de administración (sin exponer la contraseña) */
function getUsuarios(){
  const sheet = getOrCrearHojaUsuarios();
  const data = rowsToObjects(sheet.getDataRange().getValues());
  const usuarios = data
    .filter(u => String(u.USUARIO_ID || "").trim() !== "")
    .map(u => ({
      USUARIO_ID: u.USUARIO_ID,
      NOMBRE: u.NOMBRE,
      USERNAME: u.USERNAME,
      ROL: u.ROL,
      ACTIVO: u.ACTIVO
    }));
  return jsonResponse({ success:true, usuarios });
}

function crearUsuario(data){
  const nombre = (data.nombre || "").toString().trim();
  const username = (data.username || "").toString().trim();
  const password = (data.password || "").toString();
  // OJO: data.rol lo pisa el control de acceso con el rol de QUIEN hace
  // la solicitud; el rol del usuario NUEVO viaja en rolNuevo (o en
  // rolUsuario, por compatibilidad).
  const rol = (data.rolNuevo || data.rolUsuario || "").toString().trim();

  if(!nombre || !username || !password){
    return jsonResponse({ success:false, message:"Nombre, usuario y contraseña son obligatorios." });
  }
  if(password.length < 4){
    return jsonResponse({ success:false, message:"La contraseña debe tener al menos 4 caracteres." });
  }
  if(ROLES_VALIDOS.indexOf(rol) === -1){
    return jsonResponse({ success:false, message:"Rol no válido." });
  }

  const sheet = getOrCrearHojaUsuarios();
  const existentes = rowsToObjects(sheet.getDataRange().getValues());

  const yaExiste = existentes.some(u => String(u.USERNAME || "").trim().toLowerCase() === username.toLowerCase());
  if(yaExiste){
    return jsonResponse({ success:false, message:"Ya existe un usuario con ese nombre de usuario." });
  }

  const usuarioId = "USR-" + Date.now();
  sheet.appendRow([usuarioId, nombre, username, hashPassword(password), rol, "SI"]);
  invalidarCacheUsuarios();

  return jsonResponse({ success:true, usuarioId });
}

function editarUsuario(data){
  const usuarioId = (data.usuarioId || "").toString().trim();
  if(!usuarioId){
    return jsonResponse({ success:false, message:"Falta el ID del usuario a editar." });
  }

  const sheet = getOrCrearHojaUsuarios();
  const filas = sheet.getDataRange().getValues();
  const headerRow = filas[0];
  const idxId = headerRow.indexOf("USUARIO_ID");
  const idxNombre = headerRow.indexOf("NOMBRE");
  const idxUsername = headerRow.indexOf("USERNAME");
  const idxPassword = headerRow.indexOf("PASSWORD");
  const idxRol = headerRow.indexOf("ROL");
  const idxActivo = headerRow.indexOf("ACTIVO");

  let filaIndex = -1;
  for(let i = 1; i < filas.length; i++){
    if(String(filas[i][idxId]).trim() === usuarioId){ filaIndex = i; break; }
  }
  if(filaIndex === -1){
    return jsonResponse({ success:false, message:"No se encontró ese usuario." });
  }

  const rolNuevo = (data.rolNuevo || data.rolUsuario || "").toString().trim();
  if(rolNuevo && ROLES_VALIDOS.indexOf(rolNuevo) === -1){
    return jsonResponse({ success:false, message:"Rol no válido." });
  }
  if(data.password && String(data.password).length < 4){
    return jsonResponse({ success:false, message:"La contraseña debe tener al menos 4 caracteres." });
  }

  if(data.nombre) sheet.getRange(filaIndex + 1, idxNombre + 1).setValue(String(data.nombre).trim());
  if(data.username) sheet.getRange(filaIndex + 1, idxUsername + 1).setValue(String(data.username).trim());
  if(data.password) sheet.getRange(filaIndex + 1, idxPassword + 1).setValue(hashPassword(String(data.password)));
  if(rolNuevo) sheet.getRange(filaIndex + 1, idxRol + 1).setValue(rolNuevo);
  if(data.activo !== undefined && data.activo !== ""){
    const activo = data.activo === true || String(data.activo).toLowerCase() === "true" || String(data.activo).toUpperCase() === "SI";
    sheet.getRange(filaIndex + 1, idxActivo + 1).setValue(activo ? "SI" : "NO");
  }

  invalidarCacheUsuarios();
  return jsonResponse({ success:true });
}

function eliminarUsuario(data){
  const usuarioId = (data.usuarioId || "").toString().trim();
  if(!usuarioId){
    return jsonResponse({ success:false, message:"Falta el ID del usuario a eliminar." });
  }

  const sheet = getOrCrearHojaUsuarios();
  const filas = sheet.getDataRange().getValues();

  for(let i = 1; i < filas.length; i++){
    if(String(filas[i][0]).trim() === usuarioId){
      sheet.deleteRow(i + 1);
      invalidarCacheUsuarios();
      return jsonResponse({ success:true });
    }
  }

  return jsonResponse({ success:false, message:"No se encontró ese usuario." });
}

/**
 * Cambia usuario y/o contraseña del admin. Exige la contraseña ACTUAL
 * para autorizar el cambio. Al cambiarla se cierran todas las sesiones
 * del panel (los demás equipos vuelven a pedir login).
 */
function guardarCredencialesAdmin(e){

  const params = (e && e.parameter) || {};

  const passwordActual = (params.passwordActual || "").toString();
  const creds = leerCredencialesAdmin();

  if(!verificarPassword(passwordActual, creds.PASSWORD)){
    return jsonResponse({
      success:false,
      message:"La contraseña actual no es correcta."
    });
  }

  const nuevoUsuario = (params.nuevoUsuario || "").toString().trim();
  const nuevaPassword = (params.nuevaPassword || "").toString();

  if(!nuevoUsuario){
    return jsonResponse({
      success:false,
      message:"El usuario no puede quedar vacío."
    });
  }

  if(!nuevaPassword || nuevaPassword.length < 6){
    return jsonResponse({
      success:false,
      message:"La contraseña debe tener al menos 6 caracteres."
    });
  }

  if(nuevaPassword === CREDENCIALES_DEFAULT.PASSWORD){
    return jsonResponse({
      success:false,
      message:"Elegí una contraseña distinta de la que viene de fábrica."
    });
  }

  escribirCredencial("USUARIO", nuevoUsuario);
  escribirCredencial("PASSWORD", hashPassword(nuevaPassword));
  cerrarSesionesTipo("panel");

  // Sesión nueva para quien hizo el cambio (así no lo saca del panel).
  const ttl = String(params.dispositivo || "") === "1" ? SESION_TTL_DISPOSITIVO_SEG : SESION_TTL_PANEL_SEG;
  return jsonResponse({
    success:true,
    token: crearTokenSesion({ usuario: nuevoUsuario, rol: "admin", adminClasico: true }, ttl),
    venceEn: Date.now() + ttl * 1000
  });

}

function dashboard(){
  return conCachePolling("dashboard", 12, dashboardImpl);
}

function dashboardImpl(){

const ss =
getSS();

const sheet =
getHoja("PEDIDOS");

const data =
sheet.getDataRange().getValues();

data.shift();

const hoy =
Utilities.formatDate(
new Date(),
zonaNegocio(),
"yyyyMMdd"
);

let ventasHoy = 0;
let pedidosHoy = 0;
let pendientes = 0;

data.forEach(row => {

const fecha = row[1];
const total = Number(row[7]) || 0;
const estado = String(row[8] || "");

const fechaPedido =
Utilities.formatDate(
  new Date(fecha),
  zonaNegocio(),
  "yyyyMMdd"
);

if(fechaPedido === hoy){
  pedidosHoy++;
  ventasHoy += total;
}

if(
  estado === "NUEVO" ||
  estado === "PENDIENTE"
){
  pendientes++;
}

});

return jsonResponse({
success:true,
ventasHoy,
pedidosHoy,
pendientes
});

}

function actualizarEstado(e){

const pedidoId =
e.parameter.pedidoId;

const estado =
e.parameter.estado;

const ss =
getSS();

const sheet =
getHoja("PEDIDOS");

const data =
sheet.getDataRange().getValues();

for(let i=1;i<data.length;i++){

if(String(data[i][0]) === String(pedidoId)){

  sheet
  .getRange(i + 1, 9)
  .setValue(estado);

  const yaDescontado = String(data[i][13] || "").toUpperCase() === "SI";

  // Al pasar a PREPARANDO por primera vez, se descuenta el stock de
  // los productos del pedido — es el momento en que la mercadería
  // efectivamente sale del depósito, no antes (un pedido NUEVO es
  // solo una promesa de venta). La columna STOCK_DESCONTADO (índice
  // 13) evita descontar dos veces si el pedido pasa por PREPARANDO
  // más de una vez (ej. NUEVO → PREPARANDO → NUEVO → PREPARANDO).
  if(String(estado).toUpperCase() === "PREPARANDO" && !yaDescontado){
    descontarStockPorPedido(pedidoId);
    sheet.getRange(i + 1, 14).setValue("SI");
  }

  // Si se cancela un pedido que ya tenía el stock descontado (porque
  // pasó por PREPARANDO antes de cancelarse), esa mercadería nunca
  // llegó a salir de verdad — se devuelve al stock, y se desmarca
  // STOCK_DESCONTADO para que si en algún momento se reabre el pedido
  // y vuelve a PREPARANDO, se pueda descontar de nuevo correctamente.
  //
  // El mismo criterio aplica si el pedido vuelve a NUEVO (por ejemplo,
  // para poder editarle los ítems — editarItemsPedido solo permite
  // editar en estado NUEVO): si no se devolviera el stock acá, quedaría
  // descontado según los ítems VIEJOS para siempre, porque al volver a
  // PREPARANDO la columna STOCK_DESCONTADO ya diría "SI" y el bloque de
  // arriba no lo descontaría de nuevo con los ítems editados.
  if((String(estado).toUpperCase() === "CANCELADO" || String(estado).toUpperCase() === "NUEVO") && yaDescontado){
    devolverStockPorPedido(pedidoId);
    sheet.getRange(i + 1, 14).setValue("NO");
  }

  invalidarCachePolling();

  return jsonResponse({
    success:true
  });

}

}

return jsonResponse({
success:false,
message:"Pedido no encontrado"
});

}

/**
 * Consulta pública de pedido: el cliente ingresa su número de pedido
 * y su DNI para ver el estado, resumen y datos de envío.
 * Solo devuelve datos si el DNI coincide exactamente con el del pedido
 * — nunca expone datos de otro cliente.
 * Parámetros GET: pedidoId, dni
 */
function consultarPedido(e){
  const pedidoId = String(e.parameter.pedidoId || "").trim();
  const dniIngresado = normalizarDni(String(e.parameter.dni || "").trim());

  if(!pedidoId || !dniIngresado){
    return jsonResponse({ success:false, message:"Ingresá el número de pedido y tu DNI." });
  }

  // Frenos contra fuerza bruta (probar DNIs sobre un pedido, o pedidos sobre un DNI).
  if(pedidoId.length > 40 ||
     !permitirIntento("consulta_pedido:" + pedidoId, 8, 900) ||
     !permitirIntento("consulta_dni:" + dniIngresado, 15, 900) ||
     !permitirIntento("consulta_global", 600, 600)){
    return jsonResponse({ success:false, message:"Demasiados intentos. Esperá unos minutos y volvé a probar." });
  }

  const { pedidosPorId, detallePorId } = obtenerIndicePedidos();
  // Acepta "1001", "#1001" y los números viejos con o sin "PED-"
  const limpio = pedidoId.replace(/^#/, "").trim();
  const candidatos = [limpio, limpio.toUpperCase(), "PED-" + limpio.replace(/^PED-?/i, "")];
  let claveEncontrada = candidatos.find(c => pedidosPorId[c]);
  const pedido = claveEncontrada ? pedidosPorId[claveEncontrada] : null;

  if(!pedido){
    return jsonResponse({ success:false, message:"No encontramos ningún pedido con ese número." });
  }

  const dniPedido = normalizarDni(String(pedido.DNI || ""));
  if(dniPedido !== dniIngresado){
    return jsonResponse({ success:false, message:"El DNI no coincide con el pedido. Verificá los datos ingresados." });
  }

  const items = (detallePorId[claveEncontrada] || [])
    .map(d => ({
      PRODUCTO: d.PRODUCTO || "",
      cantidad: Number(d.CANTIDAD) || 0,
      subtotal: Number(d.SUBTOTAL) || 0,
      COLOR: d.COLOR || ""
    }));

  return jsonResponse({
    success: true,
    pedido: {
      PEDIDO_ID: pedido.PEDIDO_ID,
      FECHA:     pedido.FECHA,
      NOMBRE:    pedido.CLIENTE,
      ESTADO:    pedido.ESTADO,
      TOTAL:     pedido.TOTAL,
      EMPRESA:   pedido.EMPRESA || "",  // transporte
      DIRECCION: pedido.DIRECCION || "",
      LOCALIDAD: pedido.LOCALIDAD || "",
      PROVINCIA: pedido.PROVINCIA || "",
      MONEDA:    pedido.MONEDA || "ARS"
    },
    items
  });
}


/**
 * (desde DETALLE_PEDIDOS). Se usa para el modal de detalle en el panel.
 * Parámetros GET: pedidoId
 */
function getDetallePedido(e){
  const pedidoId = String(e.parameter.pedidoId || "").trim();
  if(!pedidoId) return jsonResponse({ success:false, message:"Falta el pedidoId." });

  const { pedidosPorId, detallePorId } = obtenerIndicePedidos();

  const pedido = pedidosPorId[pedidoId];
  if(!pedido) return jsonResponse({ success:false, message:"Pedido no encontrado." });

  const detalle = (detallePorId[pedidoId] || [])
    .map(d => ({
      CODIGO:   d.CODIGO   || "",
      PRODUCTO: d.PRODUCTO || "",
      cantidad: Number(d.CANTIDAD) || 0,
      PRECIO:   Number(d.PRECIO)   || 0,
      subtotal: Number(d.SUBTOTAL) || 0,
      COLOR:    d.COLOR || ""
    }));

  return jsonResponse({ success:true, pedido, detalle });
}


/**
 * Edita los ítems (productos, cantidades, precios) de un pedido que
 * todavía está en estado NUEVO — antes de pasar a PREPARANDO, que es
 * el momento en que se descuenta stock de verdad. Por eso se
 * restringe a NUEVO: editar después complicaría tener que devolver y
 * volver a descontar stock, con mucho más margen de error.
 *
 * Recalcula el subtotal a partir de los ítems nuevos. El descuento se
 * actualiza si el frontend manda uno nuevo (editado desde el modal);
 * si no llega nada, se mantiene el que ya tenía el pedido. Vuelve a
 * generar el PDF para que quede en sincro con lo editado.
 */
function editarItemsPedido(e){
  const params = (e && e.parameter) || e || {};

  const pedidoId = String(params.pedidoId || "").trim();
  if(!pedidoId) return jsonResponse({ success:false, message:"Falta el pedidoId." });

  const carrito = Array.isArray(params.carrito) ? params.carrito : JSON.parse(params.carrito || "[]");
  if(carrito.length === 0){
    return jsonResponse({ success:false, message:"El pedido no puede quedar sin productos." });
  }

  const ss = getSS();
  const pedidosSheet = getHoja("PEDIDOS");
  const detalleSheet = getHoja("DETALLE_PEDIDOS");

  const dataPedidos = pedidosSheet.getDataRange().getValues();
  let filaPedido = -1;
  for(let i = 1; i < dataPedidos.length; i++){
    if(String(dataPedidos[i][0]) === pedidoId){ filaPedido = i; break; }
  }
  if(filaPedido === -1) return jsonResponse({ success:false, message:"Pedido no encontrado." });

  const estadoActual = String(dataPedidos[filaPedido][8] || "").toUpperCase();
  if(estadoActual !== "NUEVO"){
    return jsonResponse({
      success:false,
      message:`Este pedido ya está en estado ${estadoActual} — los ítems solo se pueden editar mientras está en NUEVO (una vez que pasa a Preparando ya se descontó el stock).`
    });
  }

  // Borra las filas de detalle viejas de este pedido (de atrás para
  // adelante, para no correr los índices al borrar) y escribe las
  // nuevas.
  const dataDetalle = detalleSheet.getDataRange().getValues();
  for(let i = dataDetalle.length - 1; i >= 1; i--){
    if(String(dataDetalle[i][0]) === pedidoId){
      detalleSheet.deleteRow(i + 1);
    }
  }

  asegurarColumnaColorDetalle(detalleSheet);

  const filasNuevas = carrito.map(item => {
    const precio = Number(item.PRECIO) || 0;
    const cantidad = Number(item.cantidad) || 0;
    return [
      pedidoId,
      item.CODIGO,
      item.PRODUCTO,
      cantidad,
      precio,
      precio * cantidad,
      item.COLOR || ""
    ];
  });
  detalleSheet
    .getRange(detalleSheet.getLastRow() + 1, 1, filasNuevas.length, filasNuevas[0].length)
    .setValues(filasNuevas);

  // Recalcula subtotal/total, tomando el descuento editado si vino uno
  // nuevo del frontend (o el que ya tenía el pedido, si no vino nada).
  // Cada PRECIO/cantidad se pasa por "|| 0" ANTES de multiplicar, no
  // después de sumar — si se hiciera después, un solo ítem con un
  // valor inválido (por ejemplo, llegó sin PRECIO por algún motivo)
  // convierte la suma ENTERA en NaN, y ese NaN se guardaba tal cual
  // en la planilla (así quedó el pedido de la captura: "$NaN" en todos
  // lados). Con el "|| 0" en cada ítem, como mucho ese ítem puntual
  // suma $0 al total, pero el resto del cálculo sigue siendo válido.
  const subtotalNuevo = carrito.reduce((acc, item) => acc + (Number(item.PRECIO) || 0) * (Number(item.cantidad) || 0), 0);

  // Columnas de subtotal/descuento: se buscan por NOMBRE de encabezado
  // (y se crean si no existen todavía) en vez de asumir una posición
  // fija — la hoja PEDIDOS puede no tener estas columnas si nunca se
  // usó descuento antes, o puede tener otras columnas intercaladas.
  const colTotal            = obtenerOcrearColumna(pedidosSheet, "TOTAL");
  const colSubtotal         = obtenerOcrearColumna(pedidosSheet, "SUBTOTAL");
  const colDescuento        = obtenerOcrearColumna(pedidosSheet, "DESCUENTO");
  const colDescuentoEtiqueta = obtenerOcrearColumna(pedidosSheet, "DESCUENTO_ETIQUETA");
  const colPdfUrl           = obtenerOcrearColumna(pedidosSheet, "PDF_URL");

  // El descuento existente (si el pedido ya tenía uno) se lee de la
  // misma columna por nombre, no por índice fijo.
  const descuentoExistenteEnHoja = Number(pedidosSheet.getRange(filaPedido + 1, colDescuento).getValue()) || 0;
  const etiquetaExistenteEnHoja  = String(pedidosSheet.getRange(filaPedido + 1, colDescuentoEtiqueta).getValue() || "");

  // Si el frontend mandó un descuento (editado en el modal), se usa ese;
  // si no vino nada, se mantiene el que ya tenía el pedido — así esta
  // función sigue funcionando igual para llamadas viejas que no manden
  // el campo.
  const descuentoRecibido = (params.descuento !== undefined && params.descuento !== null && params.descuento !== "")
    ? Math.max(0, Number(params.descuento) || 0)
    : descuentoExistenteEnHoja;
  const descuentoExistente = Math.min(subtotalNuevo, descuentoRecibido);
  const etiquetaRecibida = (params.descuentoEtiqueta !== undefined)
    ? String(params.descuentoEtiqueta || "").trim()
    : etiquetaExistenteEnHoja;
  const totalNuevo = Math.max(0, subtotalNuevo - descuentoExistente);

  // Última barrera: si por algún motivo el resultado igual no es un
  // número válido, se corta ACÁ, antes de guardar nada — mejor un
  // error visible que un "$NaN" guardado en silencio.
  if(isNaN(subtotalNuevo) || isNaN(totalNuevo)){
    return jsonResponse({
      success:false,
      message:"No se pudo calcular el total del pedido — revisá que todos los productos tengan precio y cantidad cargados, y volvé a intentar."
    });
  }

  pedidosSheet.getRange(filaPedido + 1, colTotal).setValue(totalNuevo);
  pedidosSheet.getRange(filaPedido + 1, colSubtotal).setValue(subtotalNuevo);
  pedidosSheet.getRange(filaPedido + 1, colDescuento).setValue(descuentoExistente);
  pedidosSheet.getRange(filaPedido + 1, colDescuentoEtiqueta).setValue(etiquetaRecibida);

  // Vuelve a generar el PDF con los ítems actualizados, para que no
  // quede desincronizado de lo que realmente se va a preparar/enviar.
  try{
    const fila = dataPedidos[filaPedido];
    const dataPdf = {
      nombre: fila[2],
      empresa: fila[3],
      direccion: fila[4],
      telefono: fila[5],
      dni: fila[6],
      codigoPostal: fila[10],
      localidad: fila[11],
      provincia: fila[12],
      total: totalNuevo,
      subtotal: subtotalNuevo,
      descuento: descuentoExistente,
      descuentoEtiqueta: etiquetaRecibida,
      carrito: carrito
    };
    const pdfUrl = generarPDFPedido(pedidoId, fila[1], dataPdf);
    pedidosSheet.getRange(filaPedido + 1, colPdfUrl).setValue(pdfUrl);
  }catch(error){
    // Si falla la generación del PDF no se revierte la edición de
    // ítems (ya se guardó bien) — solo se avisa que el PDF quedó
    // desactualizado, para no perder el trabajo del cajero por un
    // problema aparte (por ejemplo, Drive lento en ese momento).
    Logger.log("Error regenerando PDF tras editar ítems: " + error);
    invalidarCachePolling();
    return jsonResponse({
      success:true,
      total: totalNuevo,
      pdfError: true,
      message:"Los ítems se guardaron, pero no se pudo regenerar el PDF. Volvé a intentar desde el detalle del pedido."
    });
  }

  invalidarCachePolling();

  return jsonResponse({ success:true, total: totalNuevo });
}

/**
 * Marca un pedido como cobrado ("cobrado" = pagado) o desmarca ese
 * cobro. Actualiza el estado de pago del pedido en la hoja PEDIDOS y
 * lo refleja en VENTAS_LOCAL para que el cierre de caja del día en
 * que se cobró lo tenga en cuenta — sin tocar el stock (eso ya
 * se descontó al pasar el pedido a PREPARANDO, no debe descontarse
 * de nuevo acá).
 *
 * Columnas nuevas en PEDIDOS (se agregan al final de la hoja):
 *   COBRADO            -> "SI" / "NO"
 *   FORMA_PAGO_COBRO    -> "EFECTIVO" / "TRANSFERENCIA" / "TARJETA"
 *   VENTA_ID_VINCULADA  -> ID de la fila que se creó en VENTAS_LOCAL,
 *                          para poder borrarla limpio si se desmarca
 *                          o si se cambia la forma de pago.
 *
 * Parámetros esperados (GET):
 *   pedidoId   (obligatorio)
 *   cobrado    "SI" o "NO"
 *   formaPago  "EFECTIVO" / "TRANSFERENCIA" / "TARJETA" — obligatorio si cobrado="SI"
 */
function marcarPedidoCobrado(e){

  const pedidoId = String(e.parameter.pedidoId || "").trim();
  const cobrado = String(e.parameter.cobrado || "NO").toUpperCase() === "SI";
  const formaPago = String(e.parameter.formaPago || "").toUpperCase().trim();

  if(!pedidoId){
    return jsonResponse({ success:false, message:"Falta el ID del pedido." });
  }
  if(cobrado && METODOS_PAGO_CIERRE.indexOf(formaPago) === -1){
    return jsonResponse({ success:false, message:"Elegí una forma de pago válida (efectivo, transferencia o tarjeta)." });
  }

  const ss = getSS();
  const pedidosSheet = getHoja("PEDIDOS");
  const ventasSheet = getHoja("VENTAS_LOCAL");

  if(!ventasSheet){
    return jsonResponse({ success:false, message:"Falta la hoja 'VENTAS_LOCAL' en la planilla." });
  }

  // Asegura que existan las 4 columnas nuevas al final del encabezado
  // (solo la primera vez que se usa esta función en esta planilla).
  const headerRange = pedidosSheet.getRange(1, 1, 1, pedidosSheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  let colCobrado = headers.indexOf("COBRADO");
  let colFormaPago = headers.indexOf("FORMA_PAGO_COBRO");
  let colVentaId = headers.indexOf("VENTA_ID_VINCULADA");
  let colFechaCobro = headers.indexOf("FECHA_COBRO");

  if(colCobrado === -1){
    const proxima = pedidosSheet.getLastColumn() + 1;
    pedidosSheet.getRange(1, proxima, 1, 4).setValues([["COBRADO", "FORMA_PAGO_COBRO", "VENTA_ID_VINCULADA", "FECHA_COBRO"]]);
    colCobrado = proxima - 1;
    colFormaPago = proxima;
    colVentaId = proxima + 1;
    colFechaCobro = proxima + 2;
  }

  const data = pedidosSheet.getDataRange().getValues();

  for(let i = 1; i < data.length; i++){

    if(String(data[i][0]) !== pedidoId) continue;

    const fila = i + 1;
    const ventaIdExistente = String(data[i][colVentaId] || "").trim();

    // Si ya había una venta vinculada de una marcación anterior (por
    // ejemplo, se desmarca, o se vuelve a marcar con otra forma de
    // pago), se borra primero para no duplicar el total en el cierre.
    if(ventaIdExistente){
      eliminarVentaLocalPorId(ventasSheet, ventaIdExistente);
      eliminarPagoCreditoAutomaticoPorPedido(pedidoId);
    }

    if(!cobrado){
      pedidosSheet.getRange(fila, colCobrado + 1, 1, 4).setValues([["NO", "", "", ""]]);
      invalidarCachePolling();
      return jsonResponse({ success:true, cobrado:false });
    }

    // Vuelve a crear la fila en VENTAS_LOCAL con la fecha de HOY (el
    // día real en que se cobró, no el día en que se hizo el pedido).
    const total = Number(data[i][7] || 0); // columna 8: TOTAL del pedido
    const cliente = String(data[i][2] || ""); // columna 3: CLIENTE
    const nuevoVentaId = generarIdCorto("VEN", ventasSheet);
    const fechaCobro = new Date();

    ventasSheet.appendRow([
      nuevoVentaId,
      fechaCobro,
      total,
      formaPago,
      "PEDIDO WEB",
      "Cobro de pedido " + pedidoId + " — " + cliente
    ]);

    pedidosSheet.getRange(fila, colCobrado + 1, 1, 4).setValues([["SI", formaPago, nuevoVentaId, fechaCobro]]);

    // Si el cliente es a crédito, registrar el pago también en PAGOS_CREDITO
    // para que el saldo de su cuenta corriente se actualice automáticamente.
    try {
      const dniPedido = normalizarDni(String(data[i][6] || "")); // columna 7: DNI
      if(dniPedido){
        const clientesSheet = getOrCrearHojaClientes();
        const clientesData = rowsToObjects(clientesSheet.getDataRange().getValues());
        const clienteObj = clientesData.find(c =>
          normalizarDni(String(c.DNI || "")) === dniPedido);
        if(clienteObj && String(clienteObj.A_CREDITO || "").toUpperCase() === "SI"){
          const pagosSheet = getOrCrearHojaPagosCredito();
          const pagoId = generarIdCorto("PAGO", pagosSheet);
          pagosSheet.appendRow([
            pagoId,
            fechaCobro,
            clienteObj.CLIENTE_ID,
            "ARS",           // moneda del pago
            total,           // monto
            "",              // tipo de cambio (no aplica en ARS)
            "ARS",           // prioridad
            total,           // montoAplicadoArs
            0,               // montoAplicadoUsd
            formaPago,       // forma de pago
            "Cobro automático — pedido " + pedidoId
          ]);
        }
      }
    } catch(errPago){
      Logger.log("marcarPedidoCobrado: no se pudo registrar pago en crédito: " + errPago);
      // El cobro del pedido ya quedó registrado igual — esto es solo un extra
    }

    invalidarCachePolling();
    return jsonResponse({ success:true, cobrado:true, ventaId:nuevoVentaId });
  }

  return jsonResponse({ success:false, message:"Pedido no encontrado." });
}

/**
 * Borra de VENTAS_LOCAL la fila con el VENTA_ID dado. Se usa al
 * desmarcar un pedido como cobrado, o al cambiar su forma de pago,
 * para que el cierre de caja no cuente dos veces el mismo dinero.
 */
function eliminarVentaLocalPorId(ventasSheet, ventaId){
  const data = ventasSheet.getDataRange().getValues();
  for(let i = data.length - 1; i >= 1; i--){
    if(String(data[i][0]) === String(ventaId)){
      ventasSheet.deleteRow(i + 1);
      // No hace falta borrar de DETALLE_VENTAS — estas ventas
      // "vinculadas a un pedido" no tienen detalle de productos ahí,
      // el detalle real ya vive en DETALLE_PEDIDOS.
      return;
    }
  }
}

/**
 * Borra de PAGOS_CREDITO la fila (si existe) que marcarPedidoCobrado
 * generó automáticamente para el pedido dado ("Cobro automático —
 * pedido X"). Se usa junto con eliminarVentaLocalPorId al desmarcar o
 * volver a marcar un pedido como cobrado: si no se borra este pago
 * también, queda un registro fantasma en la cuenta corriente del
 * cliente (parece que pagó de más) Y, en el cierre de caja, ese monto
 * se sumaría una segunda vez el día que se vuelva a marcar cobrado
 * (una vez como venta, otra como "pago de cliente" — doble conteo).
 */
function eliminarPagoCreditoAutomaticoPorPedido(pedidoId){
  const marca = "Cobro automático — pedido " + pedidoId;
  const sheet = getOrCrearHojaPagosCredito();
  const data = sheet.getDataRange().getValues();
  for(let i = data.length - 1; i >= 1; i--){
    if(String(data[i][10] || "").trim() === marca){
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * Resta del stock de PRODUCTOS las cantidades de un pedido, leyendo su
 * detalle desde DETALLE_PEDIDOS. Reutiliza el mismo mecanismo de
 * descontarStockPorVenta (incluyendo el LockService) armando un
 * carrito en el mismo formato {CODIGO, cantidad} que esa función
 * espera, a partir de las filas de DETALLE_PEDIDOS para este pedido.
 */
function descontarStockPorPedido(pedidoId){

  const ss = getSS();
  const detalleSheet = getHoja("DETALLE_PEDIDOS");
  if(!detalleSheet) return;

  const detalles = detalleSheet.getDataRange().getValues();
  detalles.shift(); // encabezados

  const carrito = detalles
    .filter(row => String(row[0]) === String(pedidoId))
    .map(row => ({ CODIGO: row[1], cantidad: Number(row[3] || 0) }));

  descontarStockPorVenta(carrito);
}

/**
 * Inverso de descontarStockPorPedido: suma de vuelta al stock las
 * cantidades de un pedido cancelado que ya había descontado stock.
 * Reutiliza descontarStockPorVenta pasándole cantidades negativas —
 * esa función simplemente resta lo que recibe, así que cantidades
 * negativas terminan sumando, sin duplicar el LockService ni la
 * lógica de búsqueda de fila por código.
 */
function devolverStockPorPedido(pedidoId){

  const ss = getSS();
  const detalleSheet = getHoja("DETALLE_PEDIDOS");
  if(!detalleSheet) return;

  const detalles = detalleSheet.getDataRange().getValues();
  detalles.shift(); // encabezados

  const carrito = detalles
    .filter(row => String(row[0]) === String(pedidoId))
    .map(row => ({ CODIGO: row[1], cantidad: -Number(row[3] || 0) }));

  descontarStockPorVenta(carrito);
}
function obtenerMetricasCompleto() {

  const ss = getSS();

  const hoy = new Date();

  /* =========================================================
     PEDIDOS (ventas online / por encargo — hoja "PEDIDOS")
  ========================================================= */

  const hojaPedidos =
    getHoja("PEDIDOS");

  const pedidos =
    hojaPedidos
      ? hojaPedidos.getDataRange().getValues()
      : [];

  let pedidosNuevos      = 0;
  let totalPedidos       = 0;

  const clientes = new Set();

  for(let i = 1; i < pedidos.length; i++){

    const cliente =
      pedidos[i][2];

    const estado =
      String(pedidos[i][8] || "")
      .toLowerCase();

    totalPedidos++;

    clientes.add(cliente);

    if(estado === "nuevo"){
      pedidosNuevos++;
    }

  }

  /* =========================================================
     VENTAS POS (mostrador) y VENTAS — PEDIDOS, ambas desde la
     hoja "VENTAS_LOCAL". Un pedido cobrado entero al retirarlo
     (ver marcarPedidoCobrado) o pagado en cuotas por un cliente a
     crédito (ver registrarPagoCredito) generan ahí una fila con
     VENDEDOR="PEDIDO WEB"; todo lo demás es venta real de
     mostrador. Mismo criterio que usa el reporte de Ventas por
     Período, para que dashboard y reportes siempre coincidan.
  ========================================================= */

  const hojaVentasPOS =
    getHoja("VENTAS_LOCAL");

  const ventasLocal =
    hojaVentasPOS
      ? hojaVentasPOS.getDataRange().getValues()
      : [];

  let ventasPOSHoy    = 0;
  let ventasPOSMes    = 0;
  let totalVentasPOS  = 0;

  let ventasPedidosHoy = 0;
  let ventasPedidosMes = 0;

  const idxAnuladaPOS = ventasLocal.length > 0 ? ventasLocal[0].indexOf("ANULADA") : -1;
  const idxPesoPOS = ventasLocal.length > 0 ? ventasLocal[0].map(h => String(h || "").trim()).indexOf("VENTAS_RESUMIDAS") : -1;

  for(let i = 1; i < ventasLocal.length; i++){

    // Excluir ventas anuladas, tanto de mostrador como de pedidos.
    const esAnulada = idxAnuladaPOS !== -1 &&
      String(ventasLocal[i][idxAnuladaPOS] || "").toUpperCase() === "SI";
    if(esAnulada) continue;

    const esPedidoCobrado =
      String(ventasLocal[i][4] || "").toUpperCase() === "PEDIDO WEB";

    const fecha =
      new Date(ventasLocal[i][1]);

    const total =
      Number(ventasLocal[i][2]) || 0;

    const esHoy =
      fecha.getDate() === hoy.getDate() &&
      fecha.getMonth() === hoy.getMonth() &&
      fecha.getFullYear() === hoy.getFullYear();

    const esEsteMes =
      fecha.getMonth() === hoy.getMonth() &&
      fecha.getFullYear() === hoy.getFullYear();

    if(esPedidoCobrado){
      if(esHoy) ventasPedidosHoy += total;
      if(esEsteMes) ventasPedidosMes += total;
    } else {
      totalVentasPOS += _pesoVenta(ventasLocal[i], idxPesoPOS);
      if(esHoy) ventasPOSHoy += total;
      if(esEsteMes) ventasPOSMes += total;
    }

  }

  /* =========================================================
     PRODUCTOS (hoja "PRODUCTOS")
  ========================================================= */

  const hojaProductos =
    getHoja("PRODUCTOS");

  const productos =
    hojaProductos
      ? hojaProductos.getDataRange().getValues()
      : [];

  let productosActivos = 0;
  let stockBajo = 0;
  let agotados = 0;

  if(productos.length > 1){

    // Columnas por NOMBRE de encabezado, no por posición fija
    const encProd = productos[0].map(h => String(h || "").trim().toUpperCase());
    const iCodigoP = encProd.indexOf("CODIGO");
    const iPublicadoP = encProd.indexOf("PUBLICADO");
    const iStockP = encProd.indexOf("STOCK");

    for(let i = 1; i < productos.length; i++){

      // Filas vacías que arrastra el rango de la hoja: se ignoran
      if(iCodigoP !== -1 && String(productos[i][iCodigoP] || "").trim() === "") continue;

      const publicado =
        iPublicadoP !== -1
          ? String(productos[i][iPublicadoP] || "").toLowerCase()
          : "";

      const stock =
        iStockP !== -1
          ? (Number(productos[i][iStockP]) || 0)
          : 0;

      if(publicado === "si"){
        productosActivos++;
      }

      if(stock === 0){
        agotados++;
      }

      if(stock > 0 && stock <= 5){
        stockBajo++;
      }

    }
  }

  /* =========================================================
     MAS VENDIDOS (top 1 — combina PEDIDOS + POS)
     Solo se necesita el ranking #1 y su cantidad para la
     tarjeta del dashboard; el ranking completo lo sirve
     getMasVendidos().
  ========================================================= */

  const topVendido = obtenerTopVendidoCombinado(ss);

  /* =========================================================
     TOTALES COMBINADOS (POS + Pedidos)
  ========================================================= */

  const ventasHoyTotal = ventasPOSHoy + ventasPedidosHoy;
  const ventasMesTotal = ventasPOSMes + ventasPedidosMes;
  const totalOperaciones = totalVentasPOS + totalPedidos;

  return {

    // Pedidos
    pedidosNuevos,
    totalPedidos,

    // Ventas — Pedidos
    ventasPedidosHoy,
    ventasPedidosMes,

    // Ventas — POS (mostrador)
    ventasPOSHoy,
    ventasPOSMes,
    totalVentasPOS,

    // Ventas combinadas (compatibilidad con tarjetas genéricas)
    ventasHoy: ventasHoyTotal,
    ventasMes: ventasMesTotal,

    clientesUnicos:
      clientes.size,

    ticketPromedio:
      totalOperaciones > 0
        ? ventasMesTotal / totalOperaciones
        : 0,

    productosActivos,
    stockBajo,
    agotados,

    masVendidoNombre:    topVendido.nombre,
    masVendidoCantidad:  topVendido.cantidad

  };

}

/**
 * Calcula el producto #1 más vendido combinando las unidades
 * vendidas por Pedidos (DETALLE_PEDIDOS) y por POS (DETALLE_VENTAS).
 * Devuelve solo el ganador, para la tarjeta resumen del dashboard.
 */
function obtenerTopVendidoCombinado(ss){

  const acumulado = {}; // codigo -> {nombre, cantidad}

  function sumar(sheetName, idxCodigo, idxNombre, idxCantidad){

    const hoja = getHoja(sheetName);
    if(!hoja) return;

    const data = hoja.getDataRange().getValues();
    data.shift(); // encabezados

    data.forEach(row => {

      const codigo   = row[idxCodigo];
      const nombre    = row[idxNombre];
      const cantidad = Number(row[idxCantidad]) || 0;

      if(!codigo) return;

      if(!acumulado[codigo]){
        acumulado[codigo] = { nombre: nombre, cantidad: 0 };
      }

      acumulado[codigo].cantidad += cantidad;

    });

  }

  // DETALLE_PEDIDOS: [PEDIDO_ID, CODIGO, PRODUCTO, CANTIDAD, ...]
  sumar("DETALLE_PEDIDOS", 1, 2, 3);

  // DETALLE_VENTAS: [VENTA_ID, CODIGO, PRODUCTO, CANTIDAD, PRECIO, SUBTOTAL]
  sumar("DETALLE_VENTAS", 1, 2, 3);

  let top = { nombre: "—", cantidad: 0 };

  Object.keys(acumulado).forEach(codigo => {
    if(acumulado[codigo].cantidad > top.cantidad){
      top = acumulado[codigo];
    }
  });

  return top;

}
/** Returns the CLIENTES sheet, creating it (with headers) the first time it's needed */
function getOrCrearHojaClientes(){

  const ss = getSS();
  let sheet = getHoja("CLIENTES");

  if(!sheet){
    sheet = ss.insertSheet("CLIENTES");
    sheet.appendRow(["CLIENTE_ID", "NOMBRE", "ALIAS", "DNI", "TELEFONO", "DIRECCION", "EMPRESA", "A_CREDITO", "FECHA_ALTA"]);
  }

  return sheet;
}

/** Returns the PAGOS_CREDITO sheet, creating it (with headers) the first time it's needed */
function getOrCrearHojaPagosCredito(){

  const ss = getSS();
  let sheet = getHoja("PAGOS_CREDITO");

  const encabezadosNuevos = ["PAGO_ID", "FECHA", "CLIENTE_ID", "MONEDA_PAGO", "MONTO_PAGO", "TIPO_CAMBIO", "PRIORIDAD", "MONTO_APLICADO_ARS", "MONTO_APLICADO_USD", "FORMA_PAGO", "OBSERVACIONES"];

  if(!sheet){
    sheet = ss.insertSheet("PAGOS_CREDITO");
    sheet.appendRow(encabezadosNuevos);
    return sheet;
  }

  // Si la hoja ya existía con el esquema viejo (sin soporte de moneda,
  // de antes de que se agregara el crédito en dólares), todavía no
  // hay pagos reales registrados — se confirmó con el usuario — así
  // que es seguro reescribir los encabezados sin migrar filas.
  const encabezadosActuales = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  if(String(encabezadosActuales[3] || "") !== "MONEDA_PAGO"){
    sheet.clear();
    sheet.appendRow(encabezadosNuevos);
  }

  return sheet;
}

/**
 * Deja solo los dígitos de un DNI/CUIT, para poder comparar
 * "20.123.456-7" contra "20123456-7" sin que el formato (puntos,
 * guiones, espacios) rompa la coincidencia. Es el mismo criterio que
 * usa toda esta sección para vincular pedidos con clientes.
 */
function normalizarDni(valor){
  return String(valor || "").replace(/\D/g, "");
}

/**
 * Devuelve la lista de clientes de la hoja CLIENTES (no de PEDIDOS),
 * con su alias, si están marcados a crédito, y — solo para los que
 * sí lo están — el saldo pendiente real: suma de sus pedidos (de
 * PEDIDOS, vinculados por DNI normalizado) menos la suma de sus pagos
 * registrados (de PAGOS_CREDITO, vinculados por CLIENTE_ID).
 *
 * Esta es la fuente de datos de la sección Clientes del panel — a
 * diferencia de getClientes() (más vieja, sigue existiendo por si
 * algo más la usa), que solo deriva clientes de PEDIDOS y no sabe
 * nada de alias ni de crédito.
 */
/**
 * Consulta rápida: dado un DNI (de cualquier formato), devuelve si
 * existe un cliente con ese DNI marcado a crédito — la usa el modal
 * de "Generar Pedido" del POS, para saber si mostrar el selector de
 * moneda al tipear el DNI del cliente.
 *
 * Parámetro esperado: dni
 */
function consultarClienteCreditoPorDni(dni){

  const dniNorm = normalizarDni(dni);
  if(!dniNorm){
    return jsonResponse({ success:true, esCredito:false });
  }

  const clientesSheet = getOrCrearHojaClientes();
  const clientesData = rowsToObjects(clientesSheet.getDataRange().getValues());
  const cliente = clientesData.find(c => normalizarDni(c.DNI) === dniNorm);

  if(!cliente || String(cliente.A_CREDITO || "").toUpperCase() !== "SI"){
    return jsonResponse({ success:true, esCredito:false });
  }

  return jsonResponse({ success:true, esCredito:true, clienteId: cliente.CLIENTE_ID, nombre: cliente.NOMBRE });
}

function getClientesConCredito(){

  const ss = getSS();

  const clientesSheet = getOrCrearHojaClientes();
  const clientesData = rowsToObjects(clientesSheet.getDataRange().getValues());

  const pedidosSheet = getHoja("PEDIDOS");
  const pedidosData = pedidosSheet ? rowsToObjects(pedidosSheet.getDataRange().getValues()) : [];

  // OPTIMIZACION: leer PAGOS_CREDITO y DEUDAS_EXTRA una sola vez acá
  // arriba (en vez de dejar que calcularSaldosCliente las releyera por
  // cada cliente a crédito) es lo que más acelera esta sección — con
  // muchos clientes a crédito, antes se leían esas hojas completas una
  // vez por cliente.
  const pagosData = rowsToObjects(getOrCrearHojaPagosCredito().getDataRange().getValues());
  const deudasSheetGlobal = getHoja("DEUDAS_EXTRA");
  const deudasData = deudasSheetGlobal ? rowsToObjects(deudasSheetGlobal.getDataRange().getValues()) : [];

  // Total pedido (en cualquier moneda, solo para la columna informativa
  // "Total Comprado" de la tabla general) y cantidad de pedidos,
  // agrupado por DNI normalizado.
  const totalPorDni = {};
  const cantidadPorDni = {};
  pedidosData.forEach(p => {
    const dniNorm = normalizarDni(p.DNI);
    if(!dniNorm) return;
    totalPorDni[dniNorm] = (totalPorDni[dniNorm] || 0) + Number(p.TOTAL || 0);
    cantidadPorDni[dniNorm] = (cantidadPorDni[dniNorm] || 0) + 1;
  });

  const clientes = clientesData.map(c => {
    const dniNorm = normalizarDni(c.DNI);
    const esCredito = String(c.A_CREDITO || "").toUpperCase() === "SI";
    const totalComprado = totalPorDni[dniNorm] || 0;

    // calcularSaldosCliente separa por moneda y mira PAGOS_CREDITO real
    // — solo se calcula para clientes a crédito, ya que para los demás
    // no representa una deuda real (es solo informativo). Le pasamos
    // los datos ya leídos arriba para que no vuelva a leer las hojas.
    const saldos = esCredito
      ? calcularSaldosCliente(c.CLIENTE_ID, dniNorm, { pedidosData, pagosData, deudasData })
      : null;

    return {
      CLIENTE_ID: c.CLIENTE_ID,
      NOMBRE: c.NOMBRE,
      ALIAS: c.ALIAS,
      DNI: c.DNI,
      TELEFONO: c.TELEFONO,
      DIRECCION: c.DIRECCION,
      EMPRESA: c.EMPRESA,
      A_CREDITO: c.A_CREDITO,
      PEDIDOS: cantidadPorDni[dniNorm] || 0,
      TOTAL_COMPRADO: totalComprado,
      SALDO_PENDIENTE_ARS: saldos ? saldos.saldoArs : 0,
      SALDO_PENDIENTE_USD: saldos ? saldos.saldoUsd : 0
    };
  });

  return jsonResponse({ success:true, clientes: clientes });
}

/**
 * Crea un cliente nuevo en la hoja CLIENTES — para clientes viejos
 * que todavía no hicieron ningún pedido con el catálogo nuevo, pero a
 * los que ya se les quiere empezar a llevar la cuenta.
 *
 * Parámetros esperados: nombre, alias (opcional), dni, telefono
 * (opcional), direccion (opcional), empresa (opcional), aCredito
 * ("SI"/"NO", default "NO")
 */
function crearCliente(data){
  invalidarCacheDnisClientes(); // la lista de DNIs que usan los pedidos web se rearma sola

  const nombre = String(data.nombre || "").trim();
  if(!nombre){
    return jsonResponse({ success:false, message:"Falta el nombre del cliente." });
  }

  const sheet = getOrCrearHojaClientes();
  const clienteId = generarIdCorto("CLI", sheet);

  sheet.appendRow([
    clienteId,
    nombre,
    String(data.alias || "").trim(),
    String(data.dni || "").trim(),
    String(data.telefono || "").trim(),
    String(data.direccion || "").trim(),
    String(data.empresa || "").trim(),
    String(data.aCredito || "NO").toUpperCase() === "SI" ? "SI" : "NO",
    new Date()
  ]);
  _guardarDatosExtraCliente(sheet, sheet.getLastRow(), data);

  return jsonResponse({ success:true, clienteId: clienteId });
}

/**
 * Edita los datos de un cliente existente, incluyendo si es a
 * crédito o no (columna A_CREDITO) — no toca FECHA_ALTA.
 *
 * Parámetros esperados: clienteId, nombre, alias, dni, telefono,
 * direccion, empresa, aCredito ("SI"/"NO" — si no se manda, se
 * conserva el valor que el cliente ya tenía, por compatibilidad con
 * cualquier llamador viejo que no lo incluya)
 */
function editarCliente(data){
  invalidarCacheDnisClientes(); // la lista de DNIs que usan los pedidos web se rearma sola

  const clienteId = String(data.clienteId || "").trim();
  const nombre = String(data.nombre || "").trim();

  if(!clienteId){
    return jsonResponse({ success:false, message:"Falta el ID del cliente." });
  }
  if(!nombre){
    return jsonResponse({ success:false, message:"Falta el nombre del cliente." });
  }

  const sheet = getOrCrearHojaClientes();
  const filas = sheet.getDataRange().getValues();

  for(let i = 1; i < filas.length; i++){
    if(String(filas[i][0]) === clienteId){

      const aCreditoActual = String(filas[i][7] || "NO").toUpperCase() === "SI" ? "SI" : "NO";
      const aCreditoNuevo = data.aCredito === undefined
        ? aCreditoActual
        : (String(data.aCredito || "NO").toUpperCase() === "SI" ? "SI" : "NO");

      sheet.getRange(i + 1, 2, 1, 7).setValues([[
        nombre,
        String(data.alias || "").trim(),
        String(data.dni || "").trim(),
        String(data.telefono || "").trim(),
        String(data.direccion || "").trim(),
        String(data.empresa || "").trim(),
        aCreditoNuevo
      ]]);
      _guardarDatosExtraCliente(sheet, i + 1, data);
      return jsonResponse({ success:true });
    }
  }

  return jsonResponse({ success:false, message:"No se encontró el cliente '" + clienteId + "'." });
}

/**
 * Elimina permanentemente TODOS los pedidos en estado CANCELADO: borra
 * sus filas de PEDIDOS y todas las filas relacionadas (por PEDIDO_ID)
 * en DETALLE_PEDIDOS. Es una limpieza manual — nada llama a esto
 * automáticamente, el usuario la dispara a mano desde el botón
 * "Eliminar cancelados" en la sección Pedidos del panel.
 *
 * No toca el stock: un pedido CANCELADO ya devolvió su stock al
 * cancelarse (ver actualizarEstado / devolverStockPorPedido), así que
 * borrarlo de la planilla ahora no debe volver a tocar el inventario.
 *
 * Devuelve cuántos pedidos se borraron, para poder avisarle al usuario.
 */
function eliminarPedidosCancelados(data){

  const rol = ((data && data.rol) || "").toString().trim().toLowerCase();
  if(rol === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para eliminar pedidos."
    });
  }

  const pedidosSheet = getHoja("PEDIDOS");
  if(!pedidosSheet){
    return jsonResponse({ success:false, message:"No se encontró la hoja 'PEDIDOS'." });
  }

  const dataPedidos = pedidosSheet.getDataRange().getValues();
  const idsACancelar = new Set();

  // Recorrer de abajo hacia arriba para poder borrar filas sin que se
  // corran los índices de las filas que todavía faltan revisar.
  for(let i = dataPedidos.length - 1; i >= 1; i--){
    const estado = String(dataPedidos[i][8] || "").toUpperCase(); // columna 9 = ESTADO
    if(estado === "CANCELADO"){
      idsACancelar.add(String(dataPedidos[i][0]));
      pedidosSheet.deleteRow(i + 1);
    }
  }

  if(idsACancelar.size === 0){
    return jsonResponse({ success:true, eliminados:0, message:"No hay pedidos cancelados para eliminar." });
  }

  // Borrar en DETALLE_PEDIDOS todas las filas de esos pedidos — puede
  // haber varias filas (una por ítem) por cada PEDIDO_ID.
  const detalleSheet = getHoja("DETALLE_PEDIDOS");
  if(detalleSheet){
    const dataDetalle = detalleSheet.getDataRange().getValues();
    for(let i = dataDetalle.length - 1; i >= 1; i--){
      if(idsACancelar.has(String(dataDetalle[i][0]))){
        detalleSheet.deleteRow(i + 1);
      }
    }
  }

  return jsonResponse({ success:true, eliminados: idsACancelar.size });
}

function eliminarCliente(data){
  invalidarCacheDnisClientes(); // la lista de DNIs que usan los pedidos web se rearma sola

  const clienteId = String(data.clienteId || "").trim();
  if(!clienteId){
    return jsonResponse({ success:false, message:"Falta el ID del cliente." });
  }

  const sheet = getOrCrearHojaClientes();
  const filas = sheet.getDataRange().getValues();

  for(let i = 1; i < filas.length; i++){
    if(String(filas[i][0]) === clienteId){
      sheet.deleteRow(i + 1);
      return jsonResponse({ success:true });
    }
  }

  return jsonResponse({ success:false, message:"No se encontró el cliente '" + clienteId + "'." });
}

/**
 * Marca o desmarca a un cliente como "a crédito" — la única forma de
 * activar el crédito de alguien es desde este botón en el panel, a
 * propósito (un cliente nunca puede pedirlo solo desde el catálogo).
 *
 * Parámetros esperados: clienteId, aCredito ("SI"/"NO")
 */
function actualizarCreditoCliente(data){

  const clienteId = String(data.clienteId || "").trim();
  if(!clienteId){
    return jsonResponse({ success:false, message:"Falta el ID del cliente." });
  }

  const sheet = getOrCrearHojaClientes();
  const filas = sheet.getDataRange().getValues();

  for(let i = 1; i < filas.length; i++){
    if(String(filas[i][0]) === clienteId){
      sheet.getRange(i + 1, 8).setValue(String(data.aCredito || "NO").toUpperCase() === "SI" ? "SI" : "NO"); // columna 8 = A_CREDITO
      return jsonResponse({ success:true });
    }
  }

  return jsonResponse({ success:false, message:"No se encontró el cliente '" + clienteId + "'." });
}

/**
 * Crea un cliente nuevo a partir de uno que ya existe en PEDIDOS pero
 * todavía no en CLIENTES — toma sus datos del pedido más reciente que
 * tenga ese DNI, y lo marca a crédito de una. Es el camino que usa el
 * botón "Marcar a crédito" cuando se aplica sobre alguien que viene
 * del ranking derivado de pedidos, no de la hoja CLIENTES.
 *
 * Parámetros esperados: dni
 */
function crearClienteDesdeHistorialYMarcarCredito(data){
  invalidarCacheDnisClientes(); // la lista de DNIs que usan los pedidos web se rearma sola

  const dniBuscado = normalizarDni(data.dni);
  if(!dniBuscado){
    return jsonResponse({ success:false, message:"Falta el DNI del cliente." });
  }

  const clientesSheet = getOrCrearHojaClientes();
  const clientesExistentes = rowsToObjects(clientesSheet.getDataRange().getValues());

  const yaExiste = clientesExistentes.some(c => normalizarDni(c.DNI) === dniBuscado);
  if(yaExiste){
    return jsonResponse({ success:false, message:"Ese cliente ya existe en la lista — usá 'Marcar a crédito' directo desde su fila." });
  }

  const ss = getSS();
  const pedidosSheet = getHoja("PEDIDOS");
  const pedidosData = pedidosSheet ? rowsToObjects(pedidosSheet.getDataRange().getValues()) : [];

  const pedidoDelCliente = pedidosData.reverse().find(p => normalizarDni(p.DNI) === dniBuscado);
  if(!pedidoDelCliente){
    return jsonResponse({ success:false, message:"No se encontró ningún pedido con ese DNI." });
  }

  const clienteId = generarIdCorto("CLI", clientesSheet);

  clientesSheet.appendRow([
    clienteId,
    pedidoDelCliente.CLIENTE || "",
    "",
    pedidoDelCliente.DNI || "",
    pedidoDelCliente.TELEFONO || "",
    pedidoDelCliente.DIRECCION || "",
    pedidoDelCliente.EMPRESA || "",
    "SI",
    new Date()
  ]);

  return jsonResponse({ success:true, clienteId: clienteId });
}

/**
 * Calcula los saldos pendientes de un cliente en ambas monedas, sin
 * armar toda la respuesta JSON — la usan tanto getDetalleClienteCredito
 * (para mostrar) como registrarPagoCredito (para saber cuánto hay que
 * descontar de cada deuda antes de repartir un pago nuevo).
 */
function calcularSaldosCliente(clienteId, dniNorm, datosPrecargados){

  // OPTIMIZACION: antes, esta función releía PEDIDOS, PAGOS_CREDITO y
  // DEUDAS_EXTRA completas cada vez que se llamaba — y
  // getClientesConCredito() la llama UNA VEZ POR CADA cliente a
  // crédito. Con 30 clientes a crédito, eso eran 30 lecturas completas
  // de cada una de esas 3 hojas (90 lecturas de hoja completa en una
  // sola carga de la sección Clientes) — de ahí la demora, y en
  // negocios con historial grande, el riesgo de pasarse del tiempo
  // máximo de ejecución de Apps Script (lo que se veía como que "a
  // veces falla").
  //
  // Ahora, quien llama a esta función puede pasar los datos ya leídos
  // en datosPrecargados = {pedidosData, pagosData, deudasData} para
  // que no se vuelvan a leer. Si no se pasan (como en
  // registrarPagoCredito, que solo calcula para UN cliente por vez),
  // se leen igual que antes — ningún llamador existente se rompe.
  datosPrecargados = datosPrecargados || {};

  const pedidosData = datosPrecargados.pedidosData || (function(){
    const pedidosSheet = getHoja("PEDIDOS");
    return pedidosSheet ? rowsToObjects(pedidosSheet.getDataRange().getValues()) : [];
  })();
  const pedidosDelCliente = pedidosData.filter(p => normalizarDni(p.DNI) === dniNorm);

  const pagosData = datosPrecargados.pagosData ||
    rowsToObjects(getOrCrearHojaPagosCredito().getDataRange().getValues());
  const pagosDelCliente = pagosData.filter(pago => String(pago.CLIENTE_ID) === String(clienteId));

  let compradoArs = 0, compradoUsd = 0;
  pedidosDelCliente.forEach(p => {
    if(String(p.MONEDA || "ARS").toUpperCase() === "USD") compradoUsd += Number(p.TOTAL || 0);
    else compradoArs += Number(p.TOTAL || 0);
  });

  // Deudas extra (cobros fuera de pedidos y POS)
  const deudasData = datosPrecargados.deudasData !== undefined
    ? datosPrecargados.deudasData
    : (function(){
        const deudasSheet = getHoja("DEUDAS_EXTRA");
        return deudasSheet ? rowsToObjects(deudasSheet.getDataRange().getValues()) : [];
      })();
  if(deudasData){
    deudasData
      .filter(d => String(d.CLIENTE_ID) === String(clienteId))
      .forEach(d => {
        if(String(d.MONEDA || "ARS").toUpperCase() === "USD") compradoUsd += Number(d.MONTO || 0);
        else compradoArs += Number(d.MONTO || 0);
      });
  }

  let pagadoArs = 0, pagadoUsd = 0;
  pagosDelCliente.forEach(pago => {
    pagadoArs += Number(pago.MONTO_APLICADO_ARS || 0);
    pagadoUsd += Number(pago.MONTO_APLICADO_USD || 0);
  });

  return {
    compradoArs, compradoUsd, pagadoArs, pagadoUsd,
    saldoArs: Math.max(0, compradoArs - pagadoArs),
    saldoUsd: Math.max(0, compradoUsd - pagadoUsd)
  };
}

/**
 * Registra un pago (parcial o total) de un cliente a crédito, y lo
 * reparte entre la deuda en ARS y la deuda en USD según la prioridad
 * elegida — ver el ejemplo completo más abajo.
 *
 * El pago siempre se hace en UNA moneda (moneda). Si "prioridad" pide
 * cancelar primero la deuda de la OTRA moneda, se necesita un tipo de
 * cambio para convertir — con eso se calcula cuánto de esa deuda se
 * cancela, y el sobrante (si lo hay) se aplica directo a la deuda de
 * la moneda del pago, sin necesitar conversión para esa segunda parte
 * (ya está en su propia moneda). Si el pago no alcanza para cubrir
 * completa la deuda priorizada, se aplica entero ahí y la otra deuda
 * no se toca para nada — así se confirmó con el usuario.
 *
 * Ejemplo real: cliente debe $100.000 ARS y USD 2000. Pago de
 * $3.050.000 ARS, prioridad USD, tipo de cambio 1500.
 *   - Deuda USD (2000) convertida a ARS al TC 1500 = $3.000.000
 *   - Pago $3.050.000 - $3.000.000 = sobran $50.000
 *   - Esos $50.000 (ya en ARS, la moneda del pago) se aplican directo
 *     a la deuda en ARS, sin necesitar otra conversión.
 *   - Resultado: USD queda en 0, ARS pasa de $100.000 a $50.000.
 *
 * Parámetros esperados: clienteId, monto, monedaPago ("ARS"/"USD"),
 * prioridad ("ARS"/"USD" — a cuál deuda aplicar primero),
 * tipoCambio (obligatorio solo si monedaPago !== prioridad),
 * formaPago (opcional), observaciones (opcional)
 */
function registrarPagoCredito(data){

  const clienteId = String(data.clienteId || "").trim();
  const monto = Number(data.monto || 0);
  const monedaPago = String(data.monedaPago || "ARS").toUpperCase();
  const prioridad = String(data.prioridad || monedaPago).toUpperCase();
  const tipoCambio = Number(data.tipoCambio || 0);

  if(!clienteId){
    return jsonResponse({ success:false, message:"Falta el cliente." });
  }
  if(!monto || monto <= 0){
    return jsonResponse({ success:false, message:"Ingresá un monto válido, mayor a 0." });
  }
  if(monedaPago !== prioridad && tipoCambio <= 0){
    return jsonResponse({ success:false, message:"Ingresá el tipo de cambio para convertir entre pesos y dólares." });
  }

  const clientesSheet = getOrCrearHojaClientes();
  const clientesData = rowsToObjects(clientesSheet.getDataRange().getValues());
  const cliente = clientesData.find(c => String(c.CLIENTE_ID) === clienteId);
  if(!cliente){
    return jsonResponse({ success:false, message:"No se encontró el cliente." });
  }

  const dniNorm = normalizarDni(cliente.DNI);
  const saldos = calcularSaldosCliente(clienteId, dniNorm);

  let montoAplicadoArs = 0;
  let montoAplicadoUsd = 0;

  if(monedaPago === prioridad){
    // No hace falta convertir nada — el pago entero se intenta
    // aplicar directo a la deuda de su propia moneda. Si sobra, ese
    // sobrante SÍ necesita convertirse para aplicarse a la otra deuda
    // (y para eso hace falta el tipo de cambio, aunque monedaPago ==
    // prioridad — por eso más abajo se exige tipoCambio>0 igual si
    // hay sobrante real).
    if(prioridad === "ARS"){
      montoAplicadoArs = Math.min(monto, saldos.saldoArs);
      const sobranteArs = monto - montoAplicadoArs;
      if(sobranteArs > 0){
        if(tipoCambio <= 0){
          return jsonResponse({ success:false, message:"El pago cubre toda la deuda en pesos — para aplicar el sobrante a la deuda en dólares, ingresá el tipo de cambio." });
        }
        montoAplicadoUsd = Math.min(sobranteArs / tipoCambio, saldos.saldoUsd);
      }
    }else{
      montoAplicadoUsd = Math.min(monto, saldos.saldoUsd);
      const sobranteUsd = monto - montoAplicadoUsd;
      if(sobranteUsd > 0){
        if(tipoCambio <= 0){
          return jsonResponse({ success:false, message:"El pago cubre toda la deuda en dólares — para aplicar el sobrante a la deuda en pesos, ingresá el tipo de cambio." });
        }
        montoAplicadoArs = Math.min(sobranteUsd * tipoCambio, saldos.saldoArs);
      }
    }
  }else{
    // El pago es en una moneda distinta a la deuda priorizada — se
    // convierte el pago entero a la moneda de esa deuda primero.
    if(prioridad === "ARS"){
      // Pago en USD, prioridad ARS: el pago (en USD) se convierte a pesos.
      const pagoEnArs = monto * tipoCambio;
      montoAplicadoArs = Math.min(pagoEnArs, saldos.saldoArs);
      const sobranteArs = pagoEnArs - montoAplicadoArs;
      // El sobrante ya está expresado en ARS — para aplicarlo a la
      // deuda en USD hay que reconvertirlo con el mismo tipo de cambio.
      montoAplicadoUsd = sobranteArs > 0 ? Math.min(sobranteArs / tipoCambio, saldos.saldoUsd) : 0;
    }else{
      // Pago en ARS, prioridad USD: el pago (en ARS) se convierte a dólares.
      const pagoEnUsd = monto / tipoCambio;
      montoAplicadoUsd = Math.min(pagoEnUsd, saldos.saldoUsd);
      const sobranteUsd = pagoEnUsd - montoAplicadoUsd;
      montoAplicadoArs = sobranteUsd > 0 ? Math.min(sobranteUsd * tipoCambio, saldos.saldoArs) : 0;
    }
  }

  const sheet = getOrCrearHojaPagosCredito();
  const pagoId = generarIdCorto("PAGO", sheet);

  sheet.appendRow([
    pagoId,
    new Date(),
    clienteId,
    monedaPago,
    monto,
    tipoCambio || "",
    prioridad,
    montoAplicadoArs,
    montoAplicadoUsd,
    String(data.formaPago || "EFECTIVO").toUpperCase(),
    String(data.observaciones || "")
  ]);

  // Registrar en MOVIMIENTOS_CAJA el dinero físico que ingresó a la caja,
  // con la forma de pago correcta.
  // El monto que va a caja es lo que el cliente entregó físicamente convertido a ARS:
  // - Si pagó en ARS → monto directamente
  // - Si pagó en USD → monto * tipoCambio
  const montoEnCajaArs = monedaPago === "ARS"
    ? monto
    : (tipoCambio > 0 ? monto * tipoCambio : montoAplicadoArs);

  if (montoEnCajaArs > 0) {
    try {
      const ss = getSS();
      const sheetMov = getHoja("MOVIMIENTOS_CAJA");
      if (sheetMov) {
        const movId = "MOV-" + new Date().getTime();
        const fp = String(data.formaPago || "EFECTIVO").toUpperCase();
        const nombreCliente = String(cliente.NOMBRE || cliente.RAZON_SOCIAL || clienteId);
        const ahora = new Date();
        sheetMov.appendRow([
          movId,
          ahora,
          "INGRESO",
          montoEnCajaArs,
          "Cobro crédito — " + nombreCliente + (data.observaciones ? " (" + data.observaciones + ")" : ""),
          String(data.cajero || "ADMIN"),
          ahora,
          fp
        ]);
      }
    } catch(e) {
      Logger.log("Error registrando movimiento de caja por pago de crédito: " + e.message);
    }
  }

  // Además del movimiento de caja, este pago tiene que sumar a las
  // estadísticas de "Ventas — Pedidos" (dashboard y reportes), igual
  // que cuando un pedido se cobra entero en el momento (ver
  // marcarPedidoCobrado). Por eso se agrega una fila en VENTAS_LOCAL
  // con VENDEDOR="PEDIDO WEB" — es el mismo criterio que ya usan el
  // dashboard (obtenerMetricas) y el reporte de ventas por período
  // para identificar plata real que entró por pedidos/cuenta corriente,
  // sin importar si se cobró todo junto o en cuotas.
  if (montoEnCajaArs > 0) {
    try {
      const ventasSheet = getHoja("VENTAS_LOCAL");
      if (ventasSheet) {
        const nombreCliente = String(cliente.NOMBRE || cliente.RAZON_SOCIAL || clienteId);
        const nuevoVentaId = generarIdCorto("VEN", ventasSheet);
        ventasSheet.appendRow([
          nuevoVentaId,
          new Date(),
          montoEnCajaArs,
          String(data.formaPago || "EFECTIVO").toUpperCase(),
          "PEDIDO WEB",
          "Pago de crédito — " + nombreCliente + (data.observaciones ? " (" + data.observaciones + ")" : "")
        ]);
      }
    } catch(e) {
      Logger.log("Error registrando venta por pago de crédito: " + e.message);
    }
  }

  return jsonResponse({
    success: true,
    pagoId: pagoId,
    montoAplicadoArs: montoAplicadoArs,
    montoAplicadoUsd: montoAplicadoUsd
  });
}

/**
 * Devuelve el detalle completo de un cliente a crédito: sus datos,
 * todos sus pedidos (de PEDIDOS, vinculados por DNI normalizado, más
 * recientes primero), y todos sus pagos (de PAGOS_CREDITO, más
 * recientes primero) — para la pantalla de detalle/pago del panel.
 *
 * Parámetros esperados: clienteId
 */
function getDetalleClienteCredito(clienteId){

  const ss = getSS();

  const clientesSheet = getOrCrearHojaClientes();
  const clientesData = rowsToObjects(clientesSheet.getDataRange().getValues());
  const cliente = clientesData.find(c => String(c.CLIENTE_ID) === String(clienteId));

  if(!cliente){
    return jsonResponse({ success:false, message:"No se encontró el cliente." });
  }

  const dniNorm = normalizarDni(cliente.DNI);

  const pedidosSheet = getHoja("PEDIDOS");
  const pedidosData = pedidosSheet ? rowsToObjects(pedidosSheet.getDataRange().getValues()) : [];
  const pedidosDelCliente = pedidosData
    .filter(p => normalizarDni(p.DNI) === dniNorm)
    .reverse();

  const pagosSheet = getOrCrearHojaPagosCredito();
  const pagosData = rowsToObjects(pagosSheet.getDataRange().getValues());
  const pagosDelCliente = pagosData
    .filter(pago => String(pago.CLIENTE_ID) === String(clienteId))
    .reverse();

  const deudasSheet = getHoja("DEUDAS_EXTRA");
  const deudasData = deudasSheet ? rowsToObjects(deudasSheet.getDataRange().getValues()) : [];
  const deudasDelCliente = deudasData
    .filter(d => String(d.CLIENTE_ID) === String(clienteId))
    .reverse();

  const esCredito = String(cliente.A_CREDITO || "").toUpperCase() === "SI";

  // El saldo pendiente solo es una deuda REAL para clientes marcados
  // a crédito. Para los demás (compra al contado, incluso si llegaron
  // automáticamente desde un pedido del catálogo web), el total de
  // sus pedidos es solo informativo — no representa plata que deban,
  // así que el saldo siempre es $0, sin importar cuántos pedidos tenga.
  // Le pasamos a calcularSaldosCliente los datos que ya leímos arriba
  // (pedidosData, pagosData, deudasData) para que no los vuelva a leer.
  const saldos = esCredito
    ? calcularSaldosCliente(clienteId, dniNorm, { pedidosData, pagosData, deudasData })
    : { compradoArs: 0, compradoUsd: 0, pagadoArs: 0, pagadoUsd: 0, saldoArs: 0, saldoUsd: 0 };

  return jsonResponse({
    success: true,
    cliente: cliente,
    pedidos: pedidosDelCliente,
    pagos: pagosDelCliente,
    deudasExtra: deudasDelCliente,
    totalCompradoArs: saldos.compradoArs,
    totalCompradoUsd: saldos.compradoUsd,
    totalPagadoArs: saldos.pagadoArs,
    totalPagadoUsd: saldos.pagadoUsd,
    saldoPendienteArs: saldos.saldoArs,
    saldoPendienteUsd: saldos.saldoUsd
  });
}

/**
 * Devuelve todos los clientes de la hoja CLIENTES sin filtrar por
 * A_CREDITO — se usa para el selector de cliente al generar pedidos
 * desde el POS, donde todos los clientes son válidos.
 */
/**
 * La hoja CLIENTES nunca tuvo columnas para LOCALIDAD, PROVINCIA ni
 * CODIGO_POSTAL — por eso, aunque un cliente ya haya hecho pedidos
 * con esos datos completos, al seleccionarlo desde "Generar Pedido"
 * en el POS esos campos aparecían vacíos (no es que faltaran para ese
 * cliente puntual, es que no había dónde guardarlos). En vez de
 * migrar el esquema de la hoja (riesgoso con datos reales ya
 * cargados), se completan al vuelo con los datos del pedido MÁS
 * RECIENTE de ese cliente — la fuente real donde esa información sí
 * se guardó siempre.
 *
 * Devuelve un mapa: DNI normalizado → { LOCALIDAD, PROVINCIA, CODIGO_POSTAL }
 */
function obtenerUltimaDireccionPorDni(pedidosData){
  const porDni = {};
  pedidosData.forEach(p => {
    const dniNorm = normalizarDni(p.DNI);
    if(!dniNorm) return;
    const fechaPedido = new Date(p.FECHA).getTime() || 0;
    const anterior = porDni[dniNorm];
    if(!anterior || fechaPedido >= anterior._fecha){
      porDni[dniNorm] = {
        LOCALIDAD: p.LOCALIDAD || "",
        PROVINCIA: p.PROVINCIA || "",
        CODIGO_POSTAL: p.CODIGO_POSTAL || "",
        _fecha: fechaPedido
      };
    }
  });
  return porDni;
}

function getTodosClientes(){
  const ss = getSS();
  const sheet = getOrCrearHojaClientes();
  const data = rowsToObjects(sheet.getDataRange().getValues());

  const pedidosSheet = getHoja("PEDIDOS");
  const pedidosData = pedidosSheet ? rowsToObjects(pedidosSheet.getDataRange().getValues()) : [];
  const direccionesPorDni = obtenerUltimaDireccionPorDni(pedidosData);

  const clientes = data
    .filter(c => c.CLIENTE_ID && c.NOMBRE)
    .map(c => {
      const dniNorm = normalizarDni(c.DNI);
      const direccionPedido = direccionesPorDni[dniNorm];
      return {
        CLIENTE_ID: c.CLIENTE_ID,
        NOMBRE:     c.NOMBRE,
        ALIAS:      c.ALIAS     || "",
        DNI:        c.DNI       || "",
        TELEFONO:   c.TELEFONO  || "",
        DIRECCION:  c.DIRECCION || "",
        EMPRESA:    c.EMPRESA   || "",
        LOCALIDAD:  c.LOCALIDAD || (direccionPedido ? direccionPedido.LOCALIDAD : "") || "",
        PROVINCIA:  c.PROVINCIA || (direccionPedido ? direccionPedido.PROVINCIA : "") || "",
        CODIGO_POSTAL: c.CODIGO_POSTAL || (direccionPedido ? direccionPedido.CODIGO_POSTAL : "") || "",
        A_CREDITO:  c.A_CREDITO || "NO"
      };
    });
  return jsonResponse({ success: true, clientes });
}

function getClientes(){

// Reusa el índice de pedidos cacheado (se invalida solo en cada alta o
// cambio de pedido) en vez de releer y convertir toda la hoja PEDIDOS.
const data =
Object.values(obtenerIndicePedidos().pedidosPorId);

const clientes = {};

data.forEach(p => {

const dni = p.DNI;

if(!clientes[dni]){

  clientes[dni] = {

    CLIENTE: p.CLIENTE,
    EMPRESA: p.EMPRESA,
    DIRECCION: p.DIRECCION,
    TELEFONO: p.TELEFONO,
    DNI: p.DNI,
    PEDIDOS: 0,
    TOTAL: 0

  };

}

clientes[dni].PEDIDOS++;
clientes[dni].TOTAL +=
Number(p.TOTAL || 0);

});

return jsonResponse({
success:true,
clientes:Object.values(clientes)
});

}
function getAgotados(){

const data =
getSheet()
.getDataRange()
.getValues();

const productos =
rowsToObjects(data);

const agotados =
productos.filter(p => {

return Number(
p.STOCK || 0
) === 0;

});

return jsonResponse({
success:true,
productos:agotados
});

}
function getStockBajo(){

const data =
getSheet()
.getDataRange()
.getValues();

const productos =
rowsToObjects(data);

const stockBajo =
productos.filter(p => {

const stock =
Number(p.STOCK || 0);

return stock > 0 && stock <= 5;

});

return jsonResponse({
success:true,
productos:stockBajo
});

}
function getMasVendidos(){

const ss =
getSS();

const productos = {};

function acumular(sheetName, idxCodigo, idxNombre, idxCantidad){

  const sheet = getHoja(sheetName);
  if(!sheet) return;

  const data = sheet.getDataRange().getValues();
  data.shift();

  data.forEach(row => {

    const codigo   = row[idxCodigo];
    const producto = row[idxNombre];
    const cantidad = Number(row[idxCantidad]) || 0;

    if(!codigo) return;

    if(!productos[codigo]){
      productos[codigo] = {
        CODIGO: codigo,
        PRODUCTO: producto,
        VENDIDOS: 0
      };
    }

    productos[codigo].VENDIDOS += cantidad;

  });

}

// Ventas por Pedidos: [PEDIDO_ID, CODIGO, PRODUCTO, CANTIDAD, ...]
acumular("DETALLE_PEDIDOS", 1, 2, 3);

// Ventas por POS: [VENTA_ID, CODIGO, PRODUCTO, CANTIDAD, PRECIO, SUBTOTAL]
acumular("DETALLE_VENTAS", 1, 2, 3);

const ranking =
Object.values(productos)
.sort((a,b) => b.VENDIDOS - a.VENDIDOS)
.slice(0,20);

return jsonResponse({
success:true,
productos: ranking
});

}
function generarPDFPendientes(){

const ss =
getSS();

const pedidosSheet =
getHoja("PEDIDOS");

const detalleSheet =
getHoja("DETALLE_PEDIDOS");

const pedidos =
pedidosSheet.getDataRange().getValues();

const detalles =
detalleSheet.getDataRange().getValues();

for(let i=1;i<pedidos.length;i++){

  if(pedidos[i][9]){
    continue;
  }

  const pedidoId = pedidos[i][0];

  const data = {

    nombre: pedidos[i][2],
    empresa: pedidos[i][3],
    direccion: pedidos[i][4],
    telefono: pedidos[i][5],
    dni: pedidos[i][6],
    total: pedidos[i][7],
    codigoPostal: pedidos[i][10],
    localidad: pedidos[i][11],
    provincia: pedidos[i][12],
    // Columnas 17-19: subtotal, descuento y etiqueta (agregadas cuando
    // el pedido viene del POS con descuento aplicado — pueden estar
    // vacías en pedidos del catálogo web o en pedidos viejos).
    subtotal: Number(pedidos[i][16] || pedidos[i][7] || 0),
    descuento: Number(pedidos[i][17] || 0),
    descuentoEtiqueta: String(pedidos[i][18] || ""),

    carrito: []

  };

  for(let j=1;j<detalles.length;j++){

    if(String(detalles[j][0]) === String(pedidoId)){

      data.carrito.push({

        CODIGO: detalles[j][1],
        PRODUCTO: detalles[j][2],
        cantidad: detalles[j][3],
        PRECIO: detalles[j][4]

      });

    }

  }

  try{

    const pdfUrl =
    generarPDFPedido(
      pedidoId,
      pedidos[i][1],
      data
    );

    pedidosSheet
      .getRange(i + 1, 10)
      .setValue(pdfUrl);

  }catch(error){

    Logger.log(error);

  }

}

}
function guardarVenta(e){
  // Acepta tanto e.parameter (GET, como antes) como un objeto plano ya
  // parseado (POST, con el carrito en el body en vez de la URL).
  const params = (e && e.parameter) || e || {};
  return jsonResponse(guardarVentaNucleo(params));
}

/**
 * Lógica real de guardarVenta, devolviendo un objeto plano (no un
 * TextOutput) para poder reutilizarla dentro de syncLote.
 *
 * Cambios de rendimiento respecto de la versión anterior:
 *  - DETALLE_VENTAS se escribe con UN setValues para todo el carrito
 *    (antes: un appendRow por ítem, cada uno es un viaje a Sheets).
 *  - Dentro de syncLote el stock no se descuenta venta por venta: se
 *    acumula y se aplica una sola vez al final del lote.
 *  - Ya no se tira la caché del catálogo: se parchea el STOCK (ver
 *    notificarCambioStock).
 */
function guardarVentaNucleo(params){

  const ventasSheet = getHoja("VENTAS_LOCAL");
  const detalleSheet = getHoja("DETALLE_VENTAS");

  if(!ventasSheet || !detalleSheet){
    return {
      success:false,
      message:"Faltan las hojas 'VENTAS_LOCAL' y/o 'DETALLE_VENTAS' en la planilla. Creálas antes de usar el Punto de Venta."
    };
  }

  // ID generado en el cliente ANTES de intentar guardar: un reintento
  // con el mismo ID nunca duplica la venta (ver buscarVentaPorClienteVentaId).
  const clienteVentaId = String(params.clienteVentaId || "").trim();

  // Chequeo + inserción atómicos (ver historial de este bug en el
  // comentario de buscarVentaPorClienteVentaId). Dentro de syncLote
  // el lock ya lo tiene tomado esta misma ejecución.
  let lock = null;
  if (clienteVentaId && !_dentroDeLockIdempotencia) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return {
        success:false,
        message:"El servidor está ocupado procesando esta misma venta, reintentá en unos segundos."
      };
    }
  }

  try {

    if(clienteVentaId){
      const yaExistente = buscarVentaPorClienteVentaId(ventasSheet, clienteVentaId);
      if(yaExistente){
        return { success:true, ventaId: yaExistente, idempotente: true };
      }
    }

    const ventaId = generarIdCorto("VEN", ventasSheet);
    const fecha = new Date();
    const total = Number(params.total || 0);
    const formaPago = params.formaPago || "EFECTIVO";
    const vendedor = params.vendedor || "ADMIN";
    const observaciones = params.observaciones || "";
    const carrito = Array.isArray(params.carrito) ? params.carrito : JSON.parse(params.carrito || "[]");

    const filaVenta = [
      ventaId,
      fecha,
      total,
      formaPago,
      vendedor,
      observaciones,
      clienteVentaId
    ];
    // Cobro en dólares (depósito): cuántos dólares entraron a la caja y a qué dólar
    const pagoUsd = _pagoEnDolaresDeParams(params);
    if (pagoUsd) {
      const cUsd = obtenerOcrearColumna(ventasSheet, "MONTO_USD");
      const cTc = obtenerOcrearColumna(ventasSheet, "DOLAR_USADO");
      while (filaVenta.length < Math.max(cUsd, cTc)) filaVenta.push("");
      filaVenta[cUsd - 1] = pagoUsd.montoUsd;
      filaVenta[cTc - 1] = pagoUsd.tipoCambio;
    }
    ventasSheet.appendRow(filaVenta);

    if(carrito.length > 0){
      const filasDetalle = carrito.map(item => [
        ventaId,
        item.CODIGO,
        item.PRODUCTO,
        item.cantidad,
        item.PRECIO,
        Number(item.PRECIO || 0) * Number(item.cantidad || 0)
      ]);
      if(lock || _dentroDeLockIdempotencia){
        // Con el lock tomado nadie más escribe: un solo setValues
        detalleSheet.getRange(detalleSheet.getLastRow() + 1, 1, filasDetalle.length, 6).setValues(filasDetalle);
      } else {
        filasDetalle.forEach(f => detalleSheet.appendRow(f));
      }
    }

    if (lock) { lock.releaseLock(); lock = null; }

    const stock = descontarStockPorVenta(carrito) || {};

    invalidarCachePollingVentas();

    return { success:true, ventaId: ventaId, stock: stock };

  } finally {
    if (lock) lock.releaseLock();
  }

}

/**
 * Busca, entre las ventas más recientes (no hace falta revisar todo
 * el historial: un reintento por corte de conexión siempre pasa a los
 * pocos segundos o minutos del intento original, nunca días después),
 * si ya existe una venta con este clienteVentaId. Devuelve su VENTA_ID
 * si la encuentra, o null si no.
 */
function buscarVentaPorClienteVentaId(ventasSheet, clienteVentaId){
  const ULTIMAS_FILAS_A_REVISAR = 300;
  const totalFilas = ventasSheet.getLastRow();
  if (totalFilas < 2) return null;

  const desde = Math.max(2, totalFilas - ULTIMAS_FILAS_A_REVISAR + 1);
  const cantidad = totalFilas - desde + 1;
  // Columna 7 = CLIENTE_VENTA_ID (índice 6, 0-based)
  const valores = ventasSheet.getRange(desde, 1, cantidad, 7).getValues();

  for (let i = valores.length - 1; i >= 0; i--){
    if (String(valores[i][6] || "").trim() === clienteVentaId){
      return valores[i][0]; // VENTA_ID de esa fila
    }
  }
  return null;
}

/**
 * Resta del stock de PRODUCTOS las cantidades vendidas. Usa LockService
 * para evitar que dos ventas casi simultáneas (ej. dos cajas a la vez)
 * lean el mismo stock viejo y una de las restas se pierda.
 *
 * Si un producto del carrito ya no existe en PRODUCTOS (se borró
 * después de cargarse el catálogo en el POS), simplemente se lo
 * ignora — la venta ya quedó registrada igual, no se interrumpe nada
 * por esto.
 *
 * El stock puede quedar en negativo si se vendió más de lo que había
 * cargado — es intencional: no se bloquean ventas por esto, se
 * corrige el número a mano después si hace falta.
 */
let _deltasStockDiferidos = null; // { CODIGO: delta } acumulado dentro de syncLote

function descontarStockPorVenta(carrito){

  if(!carrito || carrito.length === 0) return {};

  const deltas = deltasDesdeCarrito(carrito, -1);

  // Dentro de syncLote: se acumula y se aplica una sola vez al final
  if(_deltasStockDiferidos){
    Object.keys(deltas).forEach(c => { _deltasStockDiferidos[c] = (_deltasStockDiferidos[c] || 0) + deltas[c]; });
    return {};
  }

  let lock = null;
  if(!_dentroDeLockIdempotencia){
    lock = LockService.getScriptLock();
    if(!lock.tryLock(10000)){
      // La venta ya está guardada; el stock de este carrito queda para
      // corregir a mano. Se loguea para poder diagnosticarlo.
      console.error("No se pudo obtener el lock para descontar stock. Carrito: " + JSON.stringify(carrito));
      return {};
    }
  }

  try{
    const nuevos = aplicarDeltasStock(deltas, false);
    notificarCambioStock(nuevos);
    return nuevos;
  } finally {
    if(lock) lock.releaseLock();
  }

}

/* =====================================================================
   POS — VENTAS DEL DIA (para dashboard en tiempo real)
===================================================================== */

function getVentasPOS(){
  // El dashboard solo muestra ventas activas — las anuladas no cuentan
  return conCachePolling("ventas_pos_hoy", 10, function(){
    return getVentasPOSPorRango(null, null, true);
  });
}

/**
 * Devuelve el historial de ventas POS, opcionalmente filtrado por
 * rango de fechas. Usado por la sección "Ventas POS" del panel.
 * Parámetros (todos opcionales):
 *   desde  — "yyyy-MM-dd"
 *   hasta  — "yyyy-MM-dd"
 * Si no se pasa ninguno, devuelve los últimos 60 días.
 */
function getVentasPOSHistorial(e){

  const zona = zonaNegocio();
  const params = (e && e.parameter) || {};

  let desde = params.desde || null;
  let hasta = params.hasta || null;

  if(!desde && !hasta){
    const hoy = new Date();
    const haceSesentaDias = new Date();
    haceSesentaDias.setDate(hoy.getDate() - 60);
    desde = Utilities.formatDate(haceSesentaDias, zona, "yyyy-MM-dd");
    hasta = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");
  }

  return getVentasPOSPorRango(desde, hasta);

}

/**
 * Lógica compartida: lee VENTAS_LOCAL + DETALLE_VENTAS y devuelve
 * las ventas dentro del rango [desde, hasta] (inclusive, formato
 * "yyyy-MM-dd"). Si ambos son null, devuelve solo las ventas de hoy
 * (comportamiento original de getVentasPOS, usado por el dashboard).
 */
/**
 * OPTIMIZACIÓN — lectura parcial de hojas que solo crecen (VENTAS_LOCAL,
 * MOVIMIENTOS_CAJA...). En vez de getDataRange() sobre toda la historia,
 * se lee primero UNA sola columna (la de fecha), se busca desde el final
 * hasta dónde hay filas con fecha >= desdeStr, y recién ahí se leen las
 * filas completas de ese tramo.
 *
 * Las filas se agregan en orden de llegada, pero una venta offline vieja
 * puede llegar después de otras más nuevas: por eso no se corta en la
 * primera fila vieja, sino recién después de `ventana` filas viejas
 * seguidas. Devuelve { encabezados, filas } (filas sin el encabezado,
 * en el mismo orden que en la hoja). El llamador sigue filtrando por
 * fecha exacta como antes — esto solo evita leer años de historial.
 */
function leerFilasDesdeFecha(sheet, colFecha, desdeStr, ventana){
  const ultimaFila = sheet.getLastRow();
  const ultimaCol = Math.max(sheet.getLastColumn(), 1);
  const encabezados = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0];
  if(ultimaFila < 2) return { encabezados: encabezados, filas: [] };

  const zona = zonaNegocio();
  const fechas = sheet.getRange(2, colFecha, ultimaFila - 1, 1).getValues();
  let inicio = fechas.length;
  let viejasSeguidas = 0;
  ventana = ventana || 500;

  for(let i = fechas.length - 1; i >= 0; i--){
    const v = fechas[i][0];
    let s = "";
    try{ s = v ? Utilities.formatDate(new Date(v), zona, "yyyy-MM-dd") : ""; }catch(e){ s = ""; }
    if(!s || s >= desdeStr){ inicio = i; viejasSeguidas = 0; }
    else if(++viejasSeguidas >= ventana) break;
  }

  if(inicio >= fechas.length) return { encabezados: encabezados, filas: [] };
  const filas = sheet.getRange(2 + inicio, 1, fechas.length - inicio, ultimaCol).getValues();
  return { encabezados: encabezados, filas: filas };
}

/** Igual idea para hojas de detalle: lee desde el final solo el tramo que contiene los IDs pedidos (columna A). */
function leerDetallePorIds(sheet, idsSet, ventana){
  const ultimaFila = sheet.getLastRow();
  if(ultimaFila < 2 || !idsSet || !Object.keys(idsSet).length) return [];
  const ultimaCol = Math.max(sheet.getLastColumn(), 1);
  const ids = sheet.getRange(2, 1, ultimaFila - 1, 1).getValues();
  let inicio = ids.length;
  let ajenasSeguidas = 0;
  ventana = ventana || 3000;
  for(let i = ids.length - 1; i >= 0; i--){
    if(idsSet[String(ids[i][0])]){ inicio = i; ajenasSeguidas = 0; }
    else if(++ajenasSeguidas >= ventana) break;
  }
  if(inicio >= ids.length) return [];
  return sheet.getRange(2 + inicio, 1, ids.length - inicio, ultimaCol).getValues();
}

function getVentasPOSPorRango(desde, hasta, excluirAnuladas){

  const ss =
  getSS();

  const ventasSheet =
  getHoja("VENTAS_LOCAL");

  if(!ventasSheet){
    return jsonResponse({
      success:true,
      ventas:[]
    });
  }

  const detalleSheet =
  getHoja("DETALLE_VENTAS");

  // Lectura parcial: solo el tramo final de la hoja con fechas >= desde
  // (o >= hoy, para el dashboard), en vez de toda la historia.
  const zonaLectura = zonaNegocio();
  const desdeLectura = desde || (hasta ? "0000-00-00" : Utilities.formatDate(new Date(), zonaLectura, "yyyy-MM-dd"));
  const lectura = leerFilasDesdeFecha(ventasSheet, 2, desdeLectura, 500);
  const headers = lectura.encabezados; // VENTA_ID, FECHA, TOTAL, FORMA_PAGO, VENDEDOR, OBSERVACIONES [, ANULADA, MOTIVO_ANULACION]

  const idxAnulada = headers.indexOf("ANULADA");
  const idxMotivo  = headers.indexOf("MOTIVO_ANULACION");
  const idxCliVenta = headers.indexOf("CLIENTE_VENTA_ID"); // une la venta con su factura de ARCA

  const data = lectura.filas;

  const idsBuscados = {};
  data.forEach(r => { idsBuscados[String(r[0])] = true; });

  const detalles =
  detalleSheet
    ? leerDetallePorIds(detalleSheet, idsBuscados, 3000)
    : [];

  // Índice VENTA_ID -> filas de detalle, armado una sola vez: antes se
  // recorría TODA la hoja de detalles por cada venta (O(ventas × detalles)).
  const detallesPorVenta = {};
  for(let j = 0; j < detalles.length; j++){
    const k = String(detalles[j][0]);
    if(!detallesPorVenta[k]) detallesPorVenta[k] = [];
    detallesPorVenta[k].push(detalles[j]);
  }

  const zona = zonaNegocio();
  const hoy = new Date();
  const hoyStr = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");

  const filtrarPorRango = !!(desde || hasta);

  const ventas = [];

  for(let i = data.length - 1; i >= 0; i--){

    const row = data[i];

    const ventaId = row[0];
    const fecha   = new Date(row[1]);
    const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");

    if(filtrarPorRango){
      if(desde && fechaStr < desde) continue;
      if(hasta && fechaStr > hasta) continue;
    } else {
      // Sin rango: comportamiento original — solo hoy
      if(fechaStr !== hoyStr) continue;
    }

    // Excluir anuladas cuando se pide (dashboard) — el historial las incluye igual
    if(excluirAnuladas){
      const esAnulada = idxAnulada !== -1 &&
        String(row[idxAnulada] || "").toUpperCase() === "SI";
      if(esAnulada) continue;
    }

    // Armar el resumen de items para esta venta
    const carrito = [];
    let itemsResumen = [];

    const detallesDeVenta = detallesPorVenta[String(ventaId)] || [];
    for(let j = 0; j < detallesDeVenta.length; j++){
      const d = detallesDeVenta[j];
      carrito.push({
        CODIGO: d[1],
        PRODUCTO: d[2],
        cantidad: d[3],
        PRECIO: d[4]
      });
      itemsResumen.push(d[3] + "x " + d[2]);
    }

    ventas.push({
      VENTA_ID: ventaId,
      ID: ventaId,
      FECHA: fecha,
      TOTAL: Number(row[2]) || 0,
      FORMA_PAGO: row[3],
      VENDEDOR: row[4],
      OBSERVACIONES: row[5],
      ANULADA: idxAnulada !== -1 ? String(row[idxAnulada] || "") : "",
      MOTIVO_ANULACION: idxMotivo !== -1 ? String(row[idxMotivo] || "") : "",
      CLIENTE_VENTA_ID: idxCliVenta !== -1 ? String(row[idxCliVenta] || "") : "",
      ITEMS: itemsResumen.join(", "),
      CARRITO: JSON.stringify(carrito)
    });

  }

  return jsonResponse({
    success:true,
    ventas: ventas
  });

}

/**
 * Elimina una venta del POS: borra la fila de VENTAS_LOCAL, borra las
 * filas correspondientes de DETALLE_VENTAS, y devuelve el stock de
 * los productos vendidos (como si la venta nunca hubiera pasado).
 *
 * El carrito de la venta se recibe directo del frontend (que ya lo
 * tiene en memoria, en el campo CARRITO que devuelve
 * getVentasPOSPorRango) — evita una lectura extra de DETALLE_VENTAS
 * solo para reconstruirlo, ya que el propio borrado de esas filas
 * pasa por código de todas formas.
 *
 * Parámetros esperados: ventaId, carrito (JSON con [{CODIGO, cantidad}, ...])
 */
function eliminarVenta(e){

  const params = (e && e.parameter) || {};

  const ventaId = (params.ventaId || "").toString().trim();
  const motivo  = (params.motivo  || "Sin motivo").toString().trim();

  if(!ventaId){
    return jsonResponse({ success:false, message:"Falta el ID de la venta a eliminar." });
  }
  if(/^RES-/i.test(ventaId)){
    return jsonResponse({ success:false, message:"Esa fila es un resumen de ventas de hace más de " + MESES_DETALLE_VENTAS + " meses (ver hoja VENTAS_LOCAL): no se puede anular." });
  }

  const ss = getSS();
  const ventasSheet = getHoja("VENTAS_LOCAL");

  if(!ventasSheet){
    return jsonResponse({ success:false, message:"No se encontró la hoja VENTAS_LOCAL." });
  }

  // Asegura que existan las columnas ANULADA y MOTIVO_ANULACION al final.
  // Solo la primera vez — si ya existen, las reutiliza.
  const headerRange = ventasSheet.getRange(1, 1, 1, ventasSheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  let colAnulada = headers.indexOf("ANULADA");
  let colMotivo  = headers.indexOf("MOTIVO_ANULACION");

  if(colAnulada === -1){
    const proxima = ventasSheet.getLastColumn() + 1;
    ventasSheet.getRange(1, proxima, 1, 2).setValues([["ANULADA", "MOTIVO_ANULACION"]]);
    colAnulada = proxima - 1;
    colMotivo  = proxima;
  }

  // Marcar la venta como anulada (no se borra, sigue visible en historial)
  const dataVentas = ventasSheet.getDataRange().getValues();

  for(let i = 1; i < dataVentas.length; i++){
    if(String(dataVentas[i][0]) === ventaId){
      const fila = i + 1;
      ventasSheet.getRange(fila, colAnulada + 1, 1, 2).setValues([["SI", motivo]]);

      // Devuelve el stock igual que antes
      try{
        const detalleSheet = getHoja("DETALLE_VENTAS");
        const carrito = JSON.parse(params.carrito || "[]");
        if(carrito.length > 0){
          const carritoInvertido = carrito.map(item => ({
            ...item,
            cantidad: -(Math.abs(Number(item.cantidad) || 1))
          }));
          descontarStockPorVenta(carritoInvertido);
        } else if(detalleSheet){
          // Intentar reconstruir desde DETALLE_VENTAS si no vino carrito
          const detalleData = detalleSheet.getDataRange().getValues();
          const itemsDeLaVenta = [];
          for(let j = 1; j < detalleData.length; j++){
            if(String(detalleData[j][0]) === ventaId){
              itemsDeLaVenta.push({ CODIGO: detalleData[j][1], cantidad: -(Math.abs(Number(detalleData[j][3]) || 1)) });
            }
          }
          if(itemsDeLaVenta.length > 0) descontarStockPorVenta(itemsDeLaVenta);
        }
      }catch(errStock){
        Logger.log("anularVenta: no se pudo devolver stock: " + errStock);
      }

      invalidarCachePolling();

      return jsonResponse({ success:true, anulada:true });
    }
  }

  return jsonResponse({ success:false, message:"No se encontró la venta '" + ventaId + "'." });
}
/* =====================================================================
   CIERRE DE CAJA (RECONCILIACION DIARIA)
   Hoja requerida: "CIERRES_CAJA"
   Columnas (fila 1, encabezados exactos):
   CIERRE_ID | FECHA | EFECTIVO_ESPERADO | EFECTIVO_CONTADO | EFECTIVO_DIFERENCIA |
   TRANSFERENCIA_ESPERADO | TRANSFERENCIA_CONTADO | TRANSFERENCIA_DIFERENCIA |
   TARJETA_ESPERADO | TARJETA_CONTADO | TARJETA_DIFERENCIA |
   TOTAL_ESPERADO | TOTAL_CONTADO | TOTAL_DIFERENCIA |
   VENDEDOR | OBSERVACIONES | FECHA_CIERRE
===================================================================== */

const METODOS_PAGO_CIERRE = ["EFECTIVO", "TRANSFERENCIA", "TARJETA"];

/**
 * Calcula lo esperado por forma de pago a partir de VENTAS_LOCAL,
 * para una fecha dada (por defecto, hoy). Usado para precargar el
 * formulario de cierre de caja antes de que el usuario cuente el efectivo.
 */
function getCierreCajaResumen(e){

  const fechaParam = e.parameter.fecha; // "yyyy-MM-dd", opcional

  const ss =
  getSS();

  const ventasSheet =
  getHoja("VENTAS_LOCAL");

  if(!ventasSheet){
    return jsonResponse({
      success:false,
      message:"Falta la hoja 'VENTAS_LOCAL'."
    });
  }

  const data =
  ventasSheet.getDataRange().getValues();

  const idxAnulada = data.length > 0 ? data[0].indexOf("ANULADA") : -1;
  const idxMontoUsd = data.length > 0 ? data[0].indexOf("MONTO_USD") : -1;
  const idxDolarUsado = data.length > 0 ? data[0].indexOf("DOLAR_USADO") : -1;
  const dolares = { ventasUsd: 0, cantidad: 0, efectivoArs: 0, pagosClientesUsd: 0, esperadoUsd: 0 };
  let ventasDolaresArs = 0;

  data.shift();

  const zona = zonaNegocio();

  const fechaObjetivo =
    fechaParam
      ? fechaParam
      : Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");

  const totales = {
    EFECTIVO: 0,
    TRANSFERENCIA: 0,
    TARJETA: 0
  };

  let cantidadVentas = 0;
  let cantidadAnuladas = 0; // se informan en el ticket del cierre

  data.forEach(row => {

    // Las ventas anuladas no cuentan para el cierre de caja
    const anulada = idxAnulada !== -1 && String(row[idxAnulada] || "").toUpperCase() === "SI";
    if(anulada){
      try{ if(Utilities.formatDate(new Date(row[1]), zona, "yyyy-MM-dd") === fechaObjetivo) cantidadAnuladas++; }catch(e){}
      return;
    }

    // registrarPagoCredito() crea, en VENTAS_LOCAL, una fila
    // VENDEDOR="PEDIDO WEB" con descripción "Pago de crédito — <cliente>"
    // por cada pago que un cliente hace en la sección Clientes — es lo
    // que permite que el dashboard y el reporte de ventas por período
    // (que leen VENTAS_LOCAL) cuenten esa plata como "Ventas — Pedidos".
    // Pero esa misma plata YA se suma más abajo vía "pagosClientes"
    // (leído directo de PAGOS_CREDITO, en getPagosCreditoDelDia). Si
    // también se cuenta acá, en el "esperado" del cierre de caja se
    // sumaría dos veces — una como venta, otra como pago de cliente.
    // Por eso se excluye de este total (sigue contando igual en el
    // dashboard y en los reportes de ventas, que no la suman por
    // separado como "pagosClientes").
    const esPagoCreditoDesdeClientes =
      String(row[4] || "").toUpperCase() === "PEDIDO WEB" &&
      /^Pago de cr[eé]dito\s*—/.test(String(row[5] || "").trim());
    if(esPagoCreditoDesdeClientes) return;

    const fechaVenta = new Date(row[1]);

    const fechaVentaStr =
      Utilities.formatDate(fechaVenta, zona, "yyyy-MM-dd");

    if(fechaVentaStr !== fechaObjetivo) return;

    const total = Number(row[2]) || 0;

    let formaPago =
      String(row[3] || "EFECTIVO").toUpperCase().trim();

    // Cobro en dólares: los dólares van a su propia caja y lo que se
    // completó en pesos (o el vuelto que se dio en pesos) al efectivo
    if(formaPago === "DOLARES"){
      const usd = idxMontoUsd !== -1 ? Number(row[idxMontoUsd]) || 0 : 0;
      const tc = idxDolarUsado !== -1 ? Number(row[idxDolarUsado]) || 0 : 0;
      dolares.ventasUsd += usd;
      dolares.cantidad++;
      dolares.efectivoArs += Math.round(total - usd * tc);
      ventasDolaresArs += total;
      cantidadVentas++;
      return;
    }

    if(METODOS_PAGO_CIERRE.indexOf(formaPago) === -1){
      formaPago = "EFECTIVO";
    }

    totales[formaPago] += total;
    cantidadVentas++;

  });

  const totalEsperado =
    totales.EFECTIVO + totales.TRANSFERENCIA + totales.TARJETA + ventasDolaresArs;

  // Pagos de clientes (sección Clientes → "Registrar pago") hechos ese
  // día: son plata que entra a la caja igual que una venta, con su
  // propia forma de pago, y tienen que sumar al cierre de esa forma
  // de pago — no quedar afuera del cierre.
  const pagosClientesDelDia = getPagosCreditoDelDia(fechaObjetivo);
  const pc = pagosClientesDelDia.porFormaPago;

  // Ingresos/egresos manuales de caja (retiros, gastos chicos, aportes):
  // cada movimiento ya guarda su propia forma de pago (EFECTIVO,
  // TRANSFERENCIA o TARJETA), así que se agrupan por esa forma de pago
  // en lugar de asumir que todos son en efectivo.
  const movimientosDelDia = getMovimientosCajaDelDia(fechaObjetivo);

  // Movimientos de caja separados por forma de pago
  // ANTES: todos sumaban/restaban del efectivo — INCORRECTO
  // AHORA: cada movimiento afecta solo su propia columna (efectivo, transferencia, tarjeta)
  const movimientosPorFormaPago = { EFECTIVO: 0, TRANSFERENCIA: 0, TARJETA: 0 };
  const movIngresosPorFormaPago = { EFECTIVO: 0, TRANSFERENCIA: 0, TARJETA: 0 };
  const movEgresosPorFormaPago  = { EFECTIVO: 0, TRANSFERENCIA: 0, TARJETA: 0 };
  let movimientosIngresos = 0;
  let movimientosEgresos  = 0;

  movimientosDelDia.forEach(m => {
    // registrarPagoCredito() ya crea automáticamente, en Movimientos
    // de Caja, un INGRESO "Cobro crédito — <cliente>" por cada pago
    // que un cliente hace en la sección Clientes — es lo que permite
    // ver en Movimientos de Caja que entró plata física a la caja.
    // Pero esa misma plata YA se suma acá arriba vía "pagosClientes"
    // (leído directo de PAGOS_CREDITO). Sumar además este ingreso
    // automático la contaría dos veces — una vez como "pago de
    // cliente", otra como "ingreso de caja". Por eso se excluye de
    // esta suma (sigue viéndose igual en la pantalla de Movimientos de
    // Caja, que sí necesita mostrar todo el movimiento físico de la
    // caja; solo se excluye del cálculo del cierre).
    const esCobroCreditoAutomatico = /^Cobro cr[eé]dito\s*—/.test(String(m.MOTIVO || "").trim());
    if(esCobroCreditoAutomatico) return;

    const fp = String(m.FORMA_PAGO || "EFECTIVO").toUpperCase().trim();
    const columna = ["EFECTIVO","TRANSFERENCIA","TARJETA"].includes(fp) ? fp : "EFECTIVO";
    const signo = String(m.TIPO).toUpperCase() === "INGRESO" ? 1 : -1;
    movimientosPorFormaPago[columna] += signo * m.MONTO;
    if(signo === 1){
      movimientosIngresos += m.MONTO;
      movIngresosPorFormaPago[columna] += m.MONTO;
    }else{
      movimientosEgresos += m.MONTO;
      movEgresosPorFormaPago[columna] += m.MONTO;
    }
  });

  // Pagos de clientes en dólares (en efectivo) entran a la caja de dólares
  dolares.pagosClientesUsd = _pagosCreditoUsdEfectivoDelDia(fechaObjetivo);
  dolares.esperadoUsd = Math.round((dolares.ventasUsd + dolares.pagosClientesUsd) * 100) / 100;
  dolares.ventasUsd = Math.round(dolares.ventasUsd * 100) / 100;

  const efectivoEsperadoAjustado      = totales.EFECTIVO      + pc.EFECTIVO      + movimientosPorFormaPago.EFECTIVO + dolares.efectivoArs;
  const transferenciaEsperadoAjustada = totales.TRANSFERENCIA + pc.TRANSFERENCIA + movimientosPorFormaPago.TRANSFERENCIA;
  const tarjetaEsperadaAjustada        = totales.TARJETA       + pc.TARJETA       + movimientosPorFormaPago.TARJETA;

  const totalEsperadoAjustado =
    efectivoEsperadoAjustado + transferenciaEsperadoAjustada + tarjetaEsperadaAjustada;

  // Verifica si ya existe un cierre guardado para esa fecha
  const cierreExistente = buscarCierrePorFecha(fechaObjetivo);

  return jsonResponse({
    success: true,
    fecha: fechaObjetivo,
    cantidadVentas: cantidadVentas,
    cantidadAnuladas: cantidadAnuladas,
    totalVentas: totalEsperado,
    esperado: {
      EFECTIVO:      efectivoEsperadoAjustado,
      TRANSFERENCIA: transferenciaEsperadoAjustada,
      TARJETA:       tarjetaEsperadaAjustada,
      TOTAL:         totalEsperadoAjustado
    },
    ventasEfectivo: totales.EFECTIVO,
    ventasTransferencia: totales.TRANSFERENCIA,
    ventasTarjeta: totales.TARJETA,
    ventasDolaresArs: ventasDolaresArs,
    dolares: dolares,
    pagosClientes: {
      EFECTIVO: pc.EFECTIVO,
      TRANSFERENCIA: pc.TRANSFERENCIA,
      TARJETA: pc.TARJETA,
      total: pagosClientesDelDia.total,
      detalle: pagosClientesDelDia.detalle
    },
    movimientosCaja: {
      ingresos: movimientosIngresos,
      egresos: movimientosEgresos,
      neto: movimientosIngresos - movimientosEgresos,
      porFormaPago: {
        EFECTIVO:      movimientosPorFormaPago.EFECTIVO,
        TRANSFERENCIA: movimientosPorFormaPago.TRANSFERENCIA,
        TARJETA:       movimientosPorFormaPago.TARJETA
      },
      ingresosPorFormaPago: movIngresosPorFormaPago,
      egresosPorFormaPago: movEgresosPorFormaPago,
      detalle: movimientosDelDia
    },
    yaCerrado: !!cierreExistente,
    cierreExistente: cierreExistente || null
  });

}

function buscarCierrePorFecha(fechaStr){

  const ss =
  getSS();

  const sheet =
  getHoja("CIERRES_CAJA");

  if(!sheet) return null;

  const data =
  sheet.getDataRange().getValues();

  if(data.length < 2) return null;

  const headers = data[0];
  const zona = zonaNegocio();

  for(let i = 1; i < data.length; i++){

    const filaFecha =
      Utilities.formatDate(new Date(data[i][1]), zona, "yyyy-MM-dd");

    if(filaFecha === fechaStr){

      const obj = {};
      headers.forEach((h, idx) => { obj[h] = data[i][idx]; });
      return obj;

    }

  }

  return null;

}

/**
 * Guarda el cierre de caja del día: lo esperado (según VENTAS_LOCAL)
 * vs. lo contado físicamente, por cada forma de pago, con su diferencia.
 * Si ya existe un cierre para esa fecha, lo actualiza (sobrescribe la fila)
 * en lugar de duplicarlo.
 */
function guardarCierreCaja(e){

  const ss =
  getSS();

  const sheet =
  getHoja("CIERRES_CAJA");

  if(!sheet){
    return jsonResponse({
      success:false,
      message:"Falta la hoja 'CIERRES_CAJA' en la planilla. Creála con los encabezados: CIERRE_ID, FECHA, EFECTIVO_ESPERADO, EFECTIVO_CONTADO, EFECTIVO_DIFERENCIA, TRANSFERENCIA_ESPERADO, TRANSFERENCIA_CONTADO, TRANSFERENCIA_DIFERENCIA, TARJETA_ESPERADO, TARJETA_CONTADO, TARJETA_DIFERENCIA, TOTAL_ESPERADO, TOTAL_CONTADO, TOTAL_DIFERENCIA, VENDEDOR, OBSERVACIONES, FECHA_CIERRE"
    });
  }

  const zona = zonaNegocio();

  const fecha =
    e.parameter.fecha ||
    Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");

  const efectivoEsperado      = Number(e.parameter.efectivoEsperado || 0);
  const efectivoContado       = Number(e.parameter.efectivoContado || 0);
  const transferenciaEsperado = Number(e.parameter.transferenciaEsperado || 0);
  const transferenciaContado  = Number(e.parameter.transferenciaContado || 0);
  const tarjetaEsperado       = Number(e.parameter.tarjetaEsperado || 0);
  const tarjetaContado        = Number(e.parameter.tarjetaContado || 0);

  const vendedor      = e.parameter.vendedor || "ADMIN";
  const observaciones = e.parameter.observaciones || "";

  const efectivoDiferencia      = efectivoContado - efectivoEsperado;
  const transferenciaDiferencia = transferenciaContado - transferenciaEsperado;
  const tarjetaDiferencia       = tarjetaContado - tarjetaEsperado;

  const totalEsperado  = efectivoEsperado + transferenciaEsperado + tarjetaEsperado;
  const totalContado   = efectivoContado + transferenciaContado + tarjetaContado;
  const totalDiferencia = totalContado - totalEsperado;

  const data = sheet.getDataRange().getValues();

  let filaExistente = -1;

  for(let i = 1; i < data.length; i++){
    const filaFecha =
      Utilities.formatDate(new Date(data[i][1]), zona, "yyyy-MM-dd");
    if(filaFecha === fecha){
      filaExistente = i + 1; // +1: índice de hoja (1-based) y +1 por encabezado ya incluido en i
      break;
    }
  }

  const cierreId =
    filaExistente > 0
      ? data[filaExistente - 1][0]
      : "CIERRE-" + new Date().getTime();

  const filaValores = [
    cierreId,
    fecha,
    efectivoEsperado,
    efectivoContado,
    efectivoDiferencia,
    transferenciaEsperado,
    transferenciaContado,
    transferenciaDiferencia,
    tarjetaEsperado,
    tarjetaContado,
    tarjetaDiferencia,
    totalEsperado,
    totalContado,
    totalDiferencia,
    vendedor,
    observaciones,
    new Date()
  ];

  // Caja de dólares (depósito): solo si se usó
  const hayUsd = e.parameter.usdEsperado !== undefined || e.parameter.usdContado !== undefined;
  const usdEsperado = Number(e.parameter.usdEsperado || 0);
  const usdContado = Number(e.parameter.usdContado || 0);
  const usdDiferencia = Math.round((usdContado - usdEsperado) * 100) / 100;
  if(hayUsd){
    const cols = ["USD_ESPERADO", "USD_CONTADO", "USD_DIFERENCIA"].map(n => obtenerOcrearColumna(sheet, n));
    while(filaValores.length < Math.max.apply(null, cols)) filaValores.push("");
    filaValores[cols[0] - 1] = usdEsperado;
    filaValores[cols[1] - 1] = usdContado;
    filaValores[cols[2] - 1] = usdDiferencia;
  }

  if(filaExistente > 0){
    sheet
      .getRange(filaExistente, 1, 1, filaValores.length)
      .setValues([filaValores]);
  } else {
    sheet.appendRow(filaValores);
  }

  return jsonResponse({
    success: true,
    cierreId: cierreId,
    actualizado: filaExistente > 0,
    resumen: {
      USD:            hayUsd ? { esperado: usdEsperado, contado: usdContado, diferencia: usdDiferencia } : null,
      EFECTIVO:       { esperado: efectivoEsperado,      contado: efectivoContado,      diferencia: efectivoDiferencia },
      TRANSFERENCIA:  { esperado: transferenciaEsperado, contado: transferenciaContado, diferencia: transferenciaDiferencia },
      TARJETA:        { esperado: tarjetaEsperado,        contado: tarjetaContado,        diferencia: tarjetaDiferencia },
      TOTAL:          { esperado: totalEsperado,          contado: totalContado,          diferencia: totalDiferencia }
    }
  });

}

function getHistorialCierres(){

  const ss =
  getSS();

  const sheet =
  getHoja("CIERRES_CAJA");

  if(!sheet){
    return jsonResponse({
      success:true,
      cierres:[]
    });
  }

  const data =
  sheet.getDataRange().getValues();

  const cierres = rowsToObjects(data);

  cierres.sort((a, b) => new Date(b.FECHA) - new Date(a.FECHA));

  return jsonResponse({
    success:true,
    cierres: cierres.slice(0, 30)
  });

}

/* =====================================================================
   REPORTES (sección "Reportes" del panel admin)
   Todas las funciones de abajo aceptan desde/hasta opcionales en
   e.parameter (formato "yyyy-MM-dd"). Si no vienen, se usa el mes
   en curso por defecto. Devuelven también el rango efectivamente
   usado, para que el front-end lo muestre en el encabezado del PDF.
===================================================================== */

/** Resuelve desde/hasta: si no vienen, devuelve el 1° del mes en curso hasta hoy. */
function resolverRangoReporte(e){

  const zona = zonaNegocio();
  const hoy = new Date();

  let desde = (e.parameter && e.parameter.desde) || "";
  let hasta = (e.parameter && e.parameter.hasta) || "";

  if(!desde && !hasta){
    const primerDiaMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    desde = Utilities.formatDate(primerDiaMes, zona, "yyyy-MM-dd");
    hasta = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");
  } else if(!hasta){
    hasta = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");
  } else if(!desde){
    desde = hasta;
  }

  return { desde, hasta };
}

/**
 * Lee VENTAS_LOCAL (POS) + PEDIDOS (online) dentro de [desde, hasta] y
 * devuelve una lista plana de ventas con origen, fecha, total y forma
 * de pago. Es la base que reutilizan varios reportes de esta sección.
 */
const _memoVentasCombinadas = {}; // misma ejecución: reportesVentas lo usa en varios reportes
function obtenerVentasCombinadasPorRango(desde, hasta){
  const claveMemo = desde + "|" + hasta;
  if(_memoVentasCombinadas[claveMemo]) return _memoVentasCombinadas[claveMemo];
  const resultado = _obtenerVentasCombinadasPorRangoSinMemo(desde, hasta);
  _memoVentasCombinadas[claveMemo] = resultado;
  return resultado;
}

function _obtenerVentasCombinadasPorRangoSinMemo(desde, hasta){

  const ss = getSS();
  const zona = zonaNegocio();

  const ventas = [];

  // ---- POS: hoja VENTAS_LOCAL: [VENTA_ID, FECHA, TOTAL, FORMA_PAGO, VENDEDOR, OBSERVACIONES]
  // Esta hoja ya contiene TANTO las ventas reales de mostrador COMO los
  // pedidos del catálogo web marcados como "cobrado en caja" (ver
  // marcarPedidoCobrado, que les pone VENDEDOR="PEDIDO WEB") — se
  // distinguen por esa columna para que el reporte 1 pueda seguir
  // mostrando "POS" y "Pedidos" por separado, aunque ambos sumen al
  // mismo total general sin duplicar nada.
  const hojaPOS = getHoja("VENTAS_LOCAL");
  if(hojaPOS){
    const dataPOS = hojaPOS.getDataRange().getValues();
    const encPOS = dataPOS.shift().map(h => String(h || "").trim().toUpperCase());
    const idxAnul = encPOS.indexOf("ANULADA");
    const idxPeso = encPOS.indexOf("VENTAS_RESUMIDAS");

    dataPOS.forEach(row => {
      // Ventas anuladas: no cuentan (mismo criterio que el resto de los reportes)
      if(idxAnul !== -1 && String(row[idxAnul] || "").toUpperCase() === "SI") return;
      const fecha = new Date(row[1]);
      const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
      if(fechaStr < desde || fechaStr > hasta) return;

      const esPedidoCobrado = String(row[4] || "").toUpperCase() === "PEDIDO WEB";

      ventas.push({
        ORIGEN: esPedidoCobrado ? "PEDIDO" : "POS",
        ID: row[0],
        FECHA: fecha,
        FECHA_STR: fechaStr,
        TOTAL: Number(row[2]) || 0,
        FORMA_PAGO: String(row[3] || "EFECTIVO").toUpperCase(),
        CANTIDAD: _pesoVenta(row, idxPeso) // > 1 en los resúmenes de ventas viejas
      });
    });
  }

  // NOTA IMPORTANTE: esta función ya NO lee la hoja PEDIDOS por
  // separado. Un pedido marcado como "cobrado en caja" (ver
  // marcarPedidoCobrado) genera su propia fila en VENTAS_LOCAL con la
  // fecha real del cobro — así que VENTAS_LOCAL ya contiene TANTO las
  // ventas de mostrador COMO los pedidos cobrados, sin necesitar
  // sumarlos de nuevo acá. Sumar PEDIDOS de nuevo duplicaría ese
  // mismo dinero. Los pedidos sin cobrar simplemente no están en
  // ningún lado de esto — son solo una promesa de venta, no plata real.

  return ventas;
}

/**
 * REPORTE 1 — Ventas por período.
 * Devuelve el total día por día (POS y Pedidos por separado, más el
 * combinado), para graficar la evolución de ventas en el rango.
 */
function getReporteVentasPeriodo(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const ventas = obtenerVentasCombinadasPorRango(desde, hasta);

  const porDia = {}; // "yyyy-MM-dd" -> { pos, pedidos, total }

  ventas.forEach(v => {
    if(!porDia[v.FECHA_STR]){
      porDia[v.FECHA_STR] = { fecha: v.FECHA_STR, pos: 0, pedidos: 0, total: 0 };
    }
    if(v.ORIGEN === "POS") porDia[v.FECHA_STR].pos += v.TOTAL;
    else porDia[v.FECHA_STR].pedidos += v.TOTAL;
    porDia[v.FECHA_STR].total += v.TOTAL;
  });

  const dias = Object.values(porDia).sort((a,b) => a.fecha.localeCompare(b.fecha));

  const totalPOS = dias.reduce((s,d) => s + d.pos, 0);
  const totalPedidos = dias.reduce((s,d) => s + d.pedidos, 0);
  const totalGeneral = totalPOS + totalPedidos;
  const cantidadVentas = ventas.reduce((n, v) => n + (v.CANTIDAD || 1), 0);

  return jsonResponse({
    success:true,
    desde, hasta,
    dias,
    resumen:{
      totalPOS,
      totalPedidos,
      totalGeneral,
      cantidadVentas,
      ticketPromedio: cantidadVentas ? Math.round(totalGeneral / cantidadVentas) : 0
    }
  });
}

/**
 * REPORTE 2 — Productos más vendidos, con filtro de fecha (a diferencia
 * de getMasVendidos(), que siempre usa el histórico completo para la
 * tarjeta resumen del dashboard).
 *
 * Además del ranking devuelve, en UNA sola lectura de cada hoja:
 * - el período anterior de la misma duración (para la tendencia),
 * - categoría, stock y precio actuales de cada producto (de PRODUCTOS),
 * - en cuántas ventas/pedidos distintos apareció cada producto,
 * - totales del período y del período anterior.
 * El resultado se cachea (CacheService) por rango + versión del
 * catálogo/stock: cualquier venta, pedido o edición de productos cambia
 * esa versión, así que nunca se sirve un ranking desactualizado, y abrir
 * Reportes / Reportes de Compras varias veces no relee las hojas.
 * Respuesta compatible con la anterior (CODIGO, PRODUCTO, VENDIDOS,
 * INGRESOS siguen igual), así que Reportes de Compras no cambia.
 */
function getReporteProductosVendidos(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const zona = zonaNegocio();

  // ---- Período anterior de la misma cantidad de días ----
  const aFecha = s => { const p = String(s).split("-").map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])); };
  const aTexto = d => d.toISOString().slice(0, 10);
  const DIA = 86400000;
  const cantDias = Math.max(1, Math.round((aFecha(hasta) - aFecha(desde)) / DIA) + 1);
  const antHasta = aTexto(new Date(aFecha(desde).getTime() - DIA));
  const antDesde = aTexto(new Date(aFecha(desde).getTime() - DIA * cantDias));

  // ---- Caché por rango + versión de catálogo/stock ----
  let claveCache = null;
  try{
    const props = PropertiesService.getScriptProperties();
    const version = (props.getProperty(PROP_VERSION_CATALOGO) || "0") + "-" + (props.getProperty(PROP_VERSION_STOCK) || "0");
    claveCache = "repProd3_" + desde + "_" + hasta + "_" + version;
    const enCache = (e && e.parameter && e.parameter.sinCache) ? null : cacheLeerTexto(claveCache);
    if(enCache) return ContentService.createTextOutput(enCache).setMimeType(ContentService.MimeType.JSON);
  } catch(err){
    claveCache = null;
  }

  const enRango = f => f >= desde && f <= hasta;
  const enAnterior = f => f >= antDesde && f <= antHasta;
  // Devuelve { per: "act"|"ant", f: "yyyy-MM-dd" } o null si la fecha cae fuera de los dos períodos
  const periodoDe = fecha => {
    if(!(fecha instanceof Date) || isNaN(fecha.getTime())) return null;
    const f = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
    return enRango(f) ? { per: "act", f: f } : (enAnterior(f) ? { per: "ant", f: f } : null);
  };

  // ---- Ventas de mostrador (VENTAS_LOCAL): id -> período ----
  const periodoPOS = {};
  const hojaPOS = getHoja("VENTAS_LOCAL");
  if(hojaPOS && hojaPOS.getLastRow() > 1){
    const dataPOS = hojaPOS.getDataRange().getValues();
    const headerPOS = dataPOS.shift().map(h => String(h || "").trim().toUpperCase());
    const idxAnulada = headerPOS.indexOf("ANULADA");
    const idxPesoP = headerPOS.indexOf("VENTAS_RESUMIDAS");
    dataPOS.forEach(row => {
      // Ventas anuladas: nunca cuentan en ninguna estadística.
      if(idxAnulada !== -1 && String(row[idxAnulada] || "").toUpperCase() === "SI") return;
      // Filas creadas al cobrar un pedido (VENDEDOR="PEDIDO WEB"): su
      // detalle está en DETALLE_PEDIDOS, se cuentan más abajo.
      if(String(row[4] || "").toUpperCase() === "PEDIDO WEB") return;
      const per = periodoDe(new Date(row[1]));
      if(per){ per.peso = _pesoVenta(row, idxPesoP); periodoPOS[String(row[0])] = per; }
    });
  }

  // ---- Pedidos: cuentan si ya se cobraron (fecha de cobro) o si la
  // mercadería ya salió (STOCK_DESCONTADO=SI, ej. venta a crédito; fecha
  // del pedido). Los CANCELADOS nunca cuentan. ----
  const periodoPedido = {};
  const hojaPedidos = getHoja("PEDIDOS");
  if(hojaPedidos && hojaPedidos.getLastRow() > 1){
    rowsToObjects(hojaPedidos.getDataRange().getValues()).forEach(p => {
      if(String(p.ESTADO || "").toUpperCase() === "CANCELADO") return;
      const yaCobrado = String(p.COBRADO || "").toUpperCase() === "SI";
      const yaEntregado = String(p.STOCK_DESCONTADO || "").toUpperCase() === "SI";
      if(!yaCobrado && !yaEntregado) return;
      const per = periodoDe(new Date(p.FECHA_COBRO || p.FECHA));
      if(per) periodoPedido[String(p.PEDIDO_ID)] = per;
    });
  }

  // ---- Acumular detalle (solo las 5 primeras columnas: ID, CODIGO,
  // PRODUCTO, CANTIDAD, PRECIO — no hace falta traer el resto) ----
  const productos = {}; // codigo (texto) -> datos del período actual
  const anteriores = {}; // codigo -> { unidades, ingresos } del período anterior
  const porDia = {};     // "yyyy-MM-dd" -> { unidades, ingresos } (período actual, para el gráfico)
  const diasProducto = {}; // codigo -> { "yyyy-MM-dd": [unidades, ingresos] } (solo días con ventas)

  function acumular(periodoPorId, nombreHoja, prefijo){
    const hoja = getHoja(nombreHoja);
    if(!hoja || hoja.getLastRow() < 2) return;
    const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
    filas.forEach(row => {
      const op = periodoPorId[String(row[0])];
      if(!op) return;
      const per = op.per;
      const codigo = String(row[1] === null || row[1] === undefined ? "" : row[1]).trim();
      if(!codigo) return;
      const cantidad = Number(row[3]) || 0;
      const ingreso = cantidad * (Number(row[4]) || 0);

      if(per === "ant"){
        const a = anteriores[codigo] || (anteriores[codigo] = { unidades: 0, ingresos: 0 });
        a.unidades += cantidad;
        a.ingresos += ingreso;
        return;
      }
      let p = productos[codigo];
      if(!p){
        p = productos[codigo] = { CODIGO: codigo, PRODUCTO: String(row[2] || ""), VENDIDOS: 0, INGRESOS: 0, _ops: {} };
      }
      p.VENDIDOS += cantidad;
      p.INGRESOS += ingreso;
      p._ops[prefijo + String(row[0])] = 1;

      const dia = porDia[op.f] || (porDia[op.f] = { unidades: 0, ingresos: 0 });
      dia.unidades += cantidad;
      dia.ingresos += ingreso;
      const dp = diasProducto[codigo] || (diasProducto[codigo] = {});
      const par = dp[op.f] || (dp[op.f] = [0, 0]); // [unidades, ingresos]
      par[0] += cantidad;
      par[1] += ingreso;
    });
  }
  acumular(periodoPOS, "DETALLE_VENTAS", "V");
  acumular(periodoPedido, "DETALLE_PEDIDOS", "P");

  // ---- Datos actuales del catálogo (categoría, stock, precio, nombre) ----
  const catalogo = {};
  const hojaProductos = getSheet();
  if(hojaProductos && hojaProductos.getLastRow() > 1){
    const dataProd = hojaProductos.getDataRange().getValues();
    const hp = dataProd.shift().map(h => String(h || "").trim().toUpperCase());
    const iCod = hp.indexOf("CODIGO") === -1 ? 0 : hp.indexOf("CODIGO");
    const iNom = hp.indexOf("PRODUCTO"), iCat = hp.indexOf("CATEGORIA"), iSt = hp.indexOf("STOCK"), iPr = hp.indexOf("PRECIO");
    const iAlias = hp.indexOf("ALIAS"), iCaja = hp.indexOf("CODIGO_CAJA");
    dataProd.forEach(r => {
      const c = String(r[iCod] === null || r[iCod] === undefined ? "" : r[iCod]).trim();
      if(!c) return;
      catalogo[c] = {
        nombre: iNom !== -1 ? String(r[iNom] || "") : "",
        categoria: iCat !== -1 ? String(r[iCat] || "") : "",
        stock: iSt !== -1 && r[iSt] !== "" ? Number(r[iSt]) : null,
        precio: iPr !== -1 && r[iPr] !== "" ? Number(r[iPr]) : null,
        alias: iAlias !== -1 ? String(r[iAlias] || "") : "",
        codigoCaja: iCaja !== -1 ? String(r[iCaja] || "") : ""
      };
    });
  }

  let totalUnidades = 0, totalIngresos = 0;
  const ranking = Object.keys(productos).map(codigo => {
    const p = productos[codigo];
    const cat = catalogo[codigo];
    const ant = anteriores[codigo];
    totalUnidades += p.VENDIDOS;
    totalIngresos += p.INGRESOS;
    return {
      CODIGO: codigo,
      PRODUCTO: (cat && cat.nombre) || p.PRODUCTO,
      CATEGORIA: cat ? (cat.categoria || "Sin categoría") : "Sin categoría",
      VENDIDOS: p.VENDIDOS,
      INGRESOS: Math.round(p.INGRESOS * 100) / 100,
      OPERACIONES: Object.keys(p._ops).length,
      VENDIDOS_ANTERIOR: ant ? ant.unidades : 0,
      INGRESOS_ANTERIOR: ant ? Math.round(ant.ingresos * 100) / 100 : 0,
      STOCK: cat ? cat.stock : null,
      PRECIO_ACTUAL: cat ? cat.precio : null,
      EN_CATALOGO: !!cat,
      ALIAS: cat ? cat.alias : "",
      CODIGO_CAJA: cat ? cat.codigoCaja : "",
      DIAS: diasProducto[codigo] || {}
    };
  }).sort((a, b) => (b.VENDIDOS - a.VENDIDOS) || (b.INGRESOS - a.INGRESOS));

  // Productos del catálogo SIN ventas en el período: el buscador del panel
  // los muestra igual (con 0 vendidos y su stock), así ningún producto
  // "desaparece" por no haberse vendido. Formato compacto para no pesar:
  // [codigo, nombre, categoria, stock, alias, codigoCaja]
  const sinVentas = Object.keys(catalogo)
    .filter(c => !productos[c])
    .map(c => {
      const k = catalogo[c];
      return [c, k.nombre, k.categoria || "Sin categoría", k.stock, k.alias, k.codigoCaja];
    });

  let antUnidades = 0, antIngresos = 0;
  Object.keys(anteriores).forEach(c => { antUnidades += anteriores[c].unidades; antIngresos += anteriores[c].ingresos; });

  const serieDiaria = [];
  if(cantDias <= 400){
    for(let i = 0; i < cantDias; i++){
      const f = aTexto(new Date(aFecha(desde).getTime() + DIA * i));
      const d = porDia[f] || { unidades: 0, ingresos: 0 };
      serieDiaria.push({ fecha: f, unidades: d.unidades, ingresos: Math.round(d.ingresos * 100) / 100 });
    }
  }

  const texto = JSON.stringify({
    success: true,
    desde, hasta,
    anterior: { desde: antDesde, hasta: antHasta },
    porDia: serieDiaria,
    totales: {
      unidades: totalUnidades,
      ingresos: Math.round(totalIngresos * 100) / 100,
      productos: ranking.length,
      operaciones: Object.keys(periodoPOS).filter(id => periodoPOS[id].per === "act").reduce((n, id) => n + (periodoPOS[id].peso || 1), 0) +
                   Object.keys(periodoPedido).filter(id => periodoPedido[id].per === "act").length
    },
    totalesAnterior: {
      unidades: antUnidades,
      ingresos: Math.round(antIngresos * 100) / 100,
      productos: Object.keys(anteriores).length
    },
    generado: new Date().toISOString(),
    productos: ranking,
    sinVentas: sinVentas
  });

  // 10 minutos: la clave ya cambia sola con cada venta/edición de catálogo
  if(claveCache) cacheGuardarTexto(claveCache, texto, 600);
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

/**
 * REPORTE 3 — Ventas por categoría (solo POS, que es donde se conoce
 * el precio real cobrado; cruza DETALLE_VENTAS con la hoja PRODUCTOS
 * para obtener la categoría de cada código vendido).
 */
function getReporteVentasPorCategoria(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const ss = getSS();
  const zona = zonaNegocio();

  // Mapa CODIGO -> CATEGORIA, usando la hoja de productos por nombre de columna
  const productos = rowsToObjects(getHoja(SHEET_NAME).getDataRange().getValues());
  const categoriaPorCodigo = {};
  productos.forEach(p => {
    categoriaPorCodigo[p.CODIGO] = (p.CATEGORIA || "Sin categoría").toString().trim() || "Sin categoría";
  });

  const porCategoria = {}; // categoria -> { categoria, cantidad, ingresos }

  function acumular(idsEnRango, sheetNameDetalle){
    const hojaDetalle = getHoja(sheetNameDetalle);
    if(!hojaDetalle) return;

    const dataDetalle = hojaDetalle.getDataRange().getValues();
    dataDetalle.shift();

    dataDetalle.forEach(row => {
      const idVenta = String(row[0]);
      if(!idsEnRango.has(idVenta)) return;

      const codigo = row[1];
      const cantidad = Number(row[3]) || 0;
      const precio = Number(row[4]) || 0;
      const categoria = categoriaPorCodigo[codigo] || "Sin categoría";

      if(!porCategoria[categoria]){
        porCategoria[categoria] = { categoria, cantidad: 0, ingresos: 0 };
      }
      porCategoria[categoria].cantidad += cantidad;
      porCategoria[categoria].ingresos += cantidad * precio;
    });
  }

  // ---- POS: VENTAS_LOCAL (excluye las filas de pedidos cobrados,
  // cuyo detalle real vive en DETALLE_PEDIDOS, no en DETALLE_VENTAS)
  const idsPOSEnRango = new Set();
  const hojaPOS = getHoja("VENTAS_LOCAL");
  if(hojaPOS){
    const dataPOS = hojaPOS.getDataRange().getValues();
    const idxAnulCat = dataPOS.shift().map(h => String(h || "").trim().toUpperCase()).indexOf("ANULADA");
    dataPOS.forEach(row => {
      const esPedidoCobrado = String(row[4] || "").toUpperCase() === "PEDIDO WEB";
      if(esPedidoCobrado) return;
      // Ventas anuladas: no cuentan (igual que en Productos vendidos)
      if(idxAnulCat !== -1 && String(row[idxAnulCat] || "").toUpperCase() === "SI") return;

      const fecha = new Date(row[1]);
      const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
      if(fechaStr >= desde && fechaStr <= hasta) idsPOSEnRango.add(String(row[0]));
    });
  }
  acumular(idsPOSEnRango, "DETALLE_VENTAS");

  // ---- Pedidos cobrados: PEDIDOS (filtrados por COBRADO=SI, con la
  // fecha real de cobro) + su detalle real en DETALLE_PEDIDOS.
  const idsPedidosCobradosEnRango = new Set();
  const hojaPedidos = getHoja("PEDIDOS");
  if(hojaPedidos){
    const dataPedidos = rowsToObjects(hojaPedidos.getDataRange().getValues());
    dataPedidos.forEach(p => {
      const yaCobrado = String(p.COBRADO || "").toUpperCase() === "SI";
      if(!yaCobrado) return;

      const fecha = new Date(p.FECHA_COBRO || p.FECHA);
      const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
      if(fechaStr >= desde && fechaStr <= hasta) idsPedidosCobradosEnRango.add(String(p.PEDIDO_ID));
    });
  }
  acumular(idsPedidosCobradosEnRango, "DETALLE_PEDIDOS");

  const ranking = Object.values(porCategoria).sort((a,b) => b.ingresos - a.ingresos);

  return jsonResponse({
    success:true,
    desde, hasta,
    categorias: ranking
  });
}

/**
 * REPORTE DE COMPRAS — Ventas diarias por categoría, para el gráfico de
 * tendencia del panel "Reportes de Compras" (una línea de color por
 * categoría, para detectar cambios de temporada por rubro).
 *
 * A diferencia de getReporteVentasPorCategoria (que solo cuenta pedidos
 * ya COBRADOS), este reporte cuenta los pedidos en estado ENVIADO —
 * refleja mejor la demanda real de mercadería por categoría, que es lo
 * que importa para decidir compras de stock, sin esperar a que se cobren.
 */
function getReporteVentasDiariasPorCategoria(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const zona = zonaNegocio();

  // Mapa CODIGO -> CATEGORIA
  const productos = rowsToObjects(getHoja(SHEET_NAME).getDataRange().getValues());
  const categoriaPorCodigo = {};
  productos.forEach(p => {
    categoriaPorCodigo[p.CODIGO] = (p.CATEGORIA || "Sin categoría").toString().trim() || "Sin categoría";
  });

  const porDia = {};        // "yyyy-MM-dd" -> { categoria -> total }
  const categoriasSet = {}; // categorias vistas, para devolver la lista completa

  function acumular(fechaPorIdVenta, sheetNameDetalle){
    const hojaDetalle = getHoja(sheetNameDetalle);
    if(!hojaDetalle) return;

    const dataDetalle = hojaDetalle.getDataRange().getValues();
    dataDetalle.shift();

    dataDetalle.forEach(row => {
      const idVenta = String(row[0]);
      const fechaStr = fechaPorIdVenta[idVenta];
      if(!fechaStr) return; // la venta no cae en el rango elegido

      const codigo = row[1];
      const cantidad = Number(row[3]) || 0;
      const precio = Number(row[4]) || 0;
      const categoria = categoriaPorCodigo[codigo] || "Sin categoría";

      categoriasSet[categoria] = true;
      if(!porDia[fechaStr]) porDia[fechaStr] = {};
      porDia[fechaStr][categoria] = (porDia[fechaStr][categoria] || 0) + (cantidad * precio);
    });
  }

  // ---- POS: VENTAS_LOCAL (excluye las filas de pedidos cobrados,
  // cuyo detalle real vive en DETALLE_PEDIDOS, no en DETALLE_VENTAS)
  const fechaPorIdVentaPOS = {};
  const hojaPOS = getHoja("VENTAS_LOCAL");
  if(hojaPOS){
    const dataPOS = hojaPOS.getDataRange().getValues();
    const idxAnulCat = dataPOS.shift().map(h => String(h || "").trim().toUpperCase()).indexOf("ANULADA");
    dataPOS.forEach(row => {
      const esPedidoCobrado = String(row[4] || "").toUpperCase() === "PEDIDO WEB";
      if(esPedidoCobrado) return;
      // Ventas anuladas: no cuentan (igual que en Productos vendidos)
      if(idxAnulCat !== -1 && String(row[idxAnulCat] || "").toUpperCase() === "SI") return;

      const fecha = new Date(row[1]);
      const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
      if(fechaStr >= desde && fechaStr <= hasta) fechaPorIdVentaPOS[String(row[0])] = fechaStr;
    });
  }
  acumular(fechaPorIdVentaPOS, "DETALLE_VENTAS");

  // ---- Pedidos ENVIADOS: a diferencia de los otros reportes (que usan
  // COBRADO=SI para no contar plata que todavía no entró), este gráfico
  // de tendencia necesita ver el pedido en cuanto se despacha, sin
  // esperar a que se cobre — así refleja mejor la demanda real por
  // categoría para decidir compras. Se usa FECHA (fecha del pedido), ya
  // que no existe una columna de "fecha de envío" separada.
  const fechaPorIdPedidoEnviado = {};
  const hojaPedidos = getHoja("PEDIDOS");
  if(hojaPedidos){
    const dataPedidos = rowsToObjects(hojaPedidos.getDataRange().getValues());
    dataPedidos.forEach(p => {
      const fueEnviado = String(p.ESTADO || "").toUpperCase() === "ENVIADO";
      if(!fueEnviado) return;

      const fecha = new Date(p.FECHA);
      const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
      if(fechaStr >= desde && fechaStr <= hasta) fechaPorIdPedidoEnviado[String(p.PEDIDO_ID)] = fechaStr;
    });
  }
  acumular(fechaPorIdPedidoEnviado, "DETALLE_PEDIDOS");

  const dias = Object.keys(porDia).sort().map(fecha => ({
    fecha: fecha,
    valores: porDia[fecha]
  }));

  return jsonResponse({
    success:true,
    desde, hasta,
    categorias: Object.keys(categoriasSet).sort(),
    dias: dias
  });
}

/**
 * REPORTE 4 — Comparativo de formas de pago (solo POS, que es donde
 * se registra la forma de pago; los pedidos online se muestran
 * agrupados aparte porque hoy no tienen ese dato).
 */
function getReporteFormasPago(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const ventas = obtenerVentasCombinadasPorRango(desde, hasta);

  const porForma = {}; // forma -> { forma, cantidad, total }

  ventas.forEach(v => {
    if(!porForma[v.FORMA_PAGO]){
      porForma[v.FORMA_PAGO] = { forma: v.FORMA_PAGO, cantidad: 0, total: 0 };
    }
    porForma[v.FORMA_PAGO].cantidad += (v.CANTIDAD || 1);
    porForma[v.FORMA_PAGO].total += v.TOTAL;
  });

  const ranking = Object.values(porForma).sort((a,b) => b.total - a.total);

  return jsonResponse({
    success:true,
    desde, hasta,
    formas: ranking
  });
}

/**
 * REPORTE 5 — Historial de cierres de caja con diferencias, filtrado
 * por rango de fechas (a diferencia de getHistorialCierres(), que
 * siempre devuelve los últimos 30 sin filtro, para uso interno del POS).
 */
function getReporteCierresCaja(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const ss = getSS();
  const zona = zonaNegocio();

  const sheet = getHoja("CIERRES_CAJA");
  if(!sheet){
    return jsonResponse({ success:true, desde, hasta, cierres: [] });
  }

  const cierres = rowsToObjects(sheet.getDataRange().getValues());

  const filtrados = cierres.filter(c => {
    const fecha = new Date(c.FECHA);
    const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
    return fechaStr >= desde && fechaStr <= hasta;
  });

  filtrados.sort((a, b) => new Date(b.FECHA) - new Date(a.FECHA));

  return jsonResponse({
    success:true,
    desde, hasta,
    cierres: filtrados
  });
}

/**
 * REPORTE 6 — Clientes que más compran, filtrado por rango de fechas
 * (a diferencia de getClientes(), que siempre devuelve el histórico
 * completo para la sección Clientes del admin).
 */
function getReporteClientes(e){

  const { desde, hasta } = resolverRangoReporte(e);
  const ss = getSS();
  const zona = zonaNegocio();

  const hojaPedidos = getHoja("PEDIDOS");
  if(!hojaPedidos){
    return jsonResponse({ success:true, desde, hasta, clientes: [] });
  }

  const pedidos = rowsToObjects(hojaPedidos.getDataRange().getValues());

  const clientes = {}; // dni -> { CLIENTE, EMPRESA, DNI, PEDIDOS, TOTAL }

  pedidos.forEach(p => {
    // Solo cuentan los pedidos efectivamente cobrados — un pedido sin
    // cobrar todavía no es una compra real, es una promesa de venta.
    const yaCobrado = String(p.COBRADO || "").toUpperCase() === "SI";
    if(!yaCobrado) return;

    const fecha = new Date(p.FECHA_COBRO || p.FECHA);
    const fechaStr = Utilities.formatDate(fecha, zona, "yyyy-MM-dd");
    if(fechaStr < desde || fechaStr > hasta) return;

    const dni = p.DNI;
    if(!clientes[dni]){
      clientes[dni] = {
        CLIENTE: p.CLIENTE,
        EMPRESA: p.EMPRESA,
        DNI: p.DNI,
        PEDIDOS: 0,
        TOTAL: 0
      };
    }
    clientes[dni].PEDIDOS++;
    clientes[dni].TOTAL += Number(p.TOTAL || 0);
  });

  const ranking = Object.values(clientes).sort((a,b) => b.TOTAL - a.TOTAL);

  return jsonResponse({
    success:true,
    desde, hasta,
    clientes: ranking
  });
}


/**
 * Los 5 reportes de la pestaña Ventas en UNA sola llamada. Antes el panel
 * hacía 5 pedidos a la vez: 5 ejecuciones de Apps Script leyendo las
 * mismas hojas en paralelo, que con planillas grandes superaban el
 * tiempo de espera ("AbortError"). Ahora es una sola ejecución (las
 * ventas del rango se leen una vez, ver _memoVentasCombinadas) y el
 * resultado queda en caché unos minutos (se descarta solo con cada
 * venta o cambio de stock, y con ?sinCache=1).
 */
function getReportesVentasJuntos(e){
  const { desde, hasta } = resolverRangoReporte(e);
  let claveCache = null;
  try{
    const props = PropertiesService.getScriptProperties();
    const version = (props.getProperty(PROP_VERSION_CATALOGO) || "0") + "-" + (props.getProperty(PROP_VERSION_STOCK) || "0");
    claveCache = "repVentas1_" + desde + "_" + hasta + "_" + version;
    const enCache = (e && e.parameter && e.parameter.sinCache) ? null : cacheLeerTexto(claveCache);
    if(enCache) return ContentService.createTextOutput(enCache).setMimeType(ContentService.MimeType.JSON);
  } catch(err){ claveCache = null; }

  const ev = { parameter: { desde: desde, hasta: hasta } };
  const leer = (fn, nombre) => {
    try { return JSON.parse(fn(ev).getContent()); }
    catch(err){ return { success:false, message: nombre + ": " + err.message }; }
  };

  const texto = JSON.stringify({
    success: true,
    desde, hasta,
    ventasPeriodo: leer(getReporteVentasPeriodo, "Ventas por período"),
    categorias:    leer(getReporteVentasPorCategoria, "Ventas por categoría"),
    formasPago:    leer(getReporteFormasPago, "Formas de pago"),
    cierres:       leer(getReporteCierresCaja, "Cierres de caja"),
    clientes:      leer(getReporteClientes, "Clientes")
  });

  // 5 minutos: cierres y pedidos cobrados no siempre cambian la versión de stock
  if(claveCache) cacheGuardarTexto(claveCache, texto, 300);
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}


/* =====================================================================
   MOVIMIENTOS DE CAJA — ingresos y egresos manuales (no son ventas)
   Hoja requerida: "MOVIMIENTOS_CAJA" con encabezados:
   MOVIMIENTO_ID, FECHA, TIPO, MONTO, MOTIVO, VENDEDOR, FECHA_REGISTRO, FORMA_PAGO

   TIPO es "INGRESO" o "EGRESO". Se usan junto con las ventas del día
   para calcular el efectivo esperado real en Cierre de Caja (ya que un
   retiro de caja o un gasto chico en efectivo no es una venta, pero sí
   afecta cuánto efectivo debería quedar en el cajón).
===================================================================== */

function getMovimientosCajaDelDia(fecha){

  const ss = getSS();
  const sheet = getHoja("MOVIMIENTOS_CAJA");

  if(!sheet){
    return []; // si no existe la hoja todavía, simplemente no hay movimientos que sumar
  }

  const zona = zonaNegocio();
  // Lectura parcial: solo desde la fecha pedida hacia adelante.
  const data = leerFilasDesdeFecha(sheet, 2, String(fecha || ""), 300).filas;

  return data
    .filter(row => {
      const filaFecha = Utilities.formatDate(new Date(row[1]), zona, "yyyy-MM-dd");
      return filaFecha === fecha;
    })
    .map(row => ({
      MOVIMIENTO_ID: row[0],
      FECHA: Utilities.formatDate(new Date(row[1]), zona, "yyyy-MM-dd"),
      TIPO: row[2],
      MONTO: Number(row[3] || 0),
      MOTIVO: row[4],
      VENDEDOR: row[5],
      FECHA_REGISTRO: row[6],
      FORMA_PAGO: String(row[7] || "EFECTIVO").toUpperCase().trim() || "EFECTIVO"
    }));
}

/**
 * Pagos de crédito (hoja PAGOS_CREDITO) registrados en la fecha dada,
 * agrupados por forma de pago (EFECTIVO / TRANSFERENCIA / TARJETA).
 * Son pagos que los clientes hacen en la sección "Clientes" para
 * cancelar su deuda — es plata que entra a la caja ese día con esa
 * forma de pago, y por lo tanto tiene que sumar al cierre de caja
 * igual que una venta.
 *
 * Nota: se contabiliza el monto tal cual fue pagado (MONTO_PAGO, en
 * la moneda MONEDA_PAGO). El cierre de caja trabaja en pesos, así que
 * solo se suman los pagos hechos en pesos (ARS); un pago en dólares
 * no mueve el cajón de pesos.
 */
function getPagosCreditoDelDia(fecha){

  const sheet = getOrCrearHojaPagosCredito();
  const data = sheet.getDataRange().getValues();
  data.shift(); // encabezados

  const zona = zonaNegocio();

  const detalle = data
    .filter(row => {
      const filaFecha = Utilities.formatDate(new Date(row[1]), zona, "yyyy-MM-dd");
      const monedaPago = String(row[3] || "ARS").toUpperCase().trim();
      // Los pagos "Cobro automático — pedido X" los genera
      // marcarPedidoCobrado, y ese mismo dinero YA quedó sumado como
      // venta en VENTAS_LOCAL en ese momento. Si además se suman acá,
      // el cierre de caja contaría el mismo cobro dos veces. Se
      // excluyen — solo cuentan los pagos hechos aparte, en la
      // sección Clientes, que no tienen fila equivalente en VENTAS_LOCAL.
      const esCobroAutomaticoDePedido = /^Cobro autom[aá]tico/i.test(String(row[10] || "").trim());
      return filaFecha === fecha && monedaPago === "ARS" && !esCobroAutomaticoDePedido;
    })
    .map(row => {
      const formaPago = METODOS_PAGO_CIERRE.indexOf(
        String(row[9] || "").toUpperCase().trim()
      ) !== -1
        ? String(row[9]).toUpperCase().trim()
        : "EFECTIVO";

      return {
        PAGO_ID: row[0],
        FECHA: Utilities.formatDate(new Date(row[1]), zona, "yyyy-MM-dd"),
        CLIENTE_ID: row[2],
        MONTO: Number(row[4] || 0),
        FORMA_PAGO: formaPago,
        OBSERVACIONES: row[10]
      };
    });

  const porFormaPago = { EFECTIVO: 0, TRANSFERENCIA: 0, TARJETA: 0 };
  detalle.forEach(p => { porFormaPago[p.FORMA_PAGO] += p.MONTO; });

  return {
    detalle: detalle,
    porFormaPago: porFormaPago,
    total: porFormaPago.EFECTIVO + porFormaPago.TRANSFERENCIA + porFormaPago.TARJETA
  };
}

/** GET ?action=movimientosCajaHoy[&fecha=YYYY-MM-DD] — lista del día, más el saldo neto */
function getMovimientosCajaHoy(e){

  const zona = zonaNegocio();
  const fecha = (e && e.parameter && e.parameter.fecha) ||
    Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");

  const movimientos = getMovimientosCajaDelDia(fecha)
    .sort((a, b) => new Date(b.FECHA_REGISTRO) - new Date(a.FECHA_REGISTRO));

  let totalIngresos = 0;
  let totalEgresos = 0;

  movimientos.forEach(m => {
    if(String(m.TIPO).toUpperCase() === "INGRESO") totalIngresos += m.MONTO;
    else totalEgresos += m.MONTO;
  });

  return jsonResponse({
    success: true,
    fecha: fecha,
    movimientos: movimientos,
    totalIngresos: totalIngresos,
    totalEgresos: totalEgresos,
    neto: totalIngresos - totalEgresos
  });
}

/** GET ?action=guardarMovimientoCaja&tipo=INGRESO|EGRESO&monto=...&motivo=... */
function guardarMovimientoCaja(e){

  const ss = getSS();

  // Auto-migrar columna FORMA_PAGO si la planilla es anterior al fix
  autoMigrarMovimientosCaja();

  const sheet = getHoja("MOVIMIENTOS_CAJA");

  if(!sheet){
    return jsonResponse({
      success:false,
      message:"Falta la hoja 'MOVIMIENTOS_CAJA' en la planilla. Creála con los encabezados: MOVIMIENTO_ID, FECHA, TIPO, MONTO, MOTIVO, VENDEDOR, FECHA_REGISTRO, FORMA_PAGO"
    });
  }

  const params = (e && e.parameter) || {};

  const tipo = String(params.tipo || "").toUpperCase().trim();
  if(tipo !== "INGRESO" && tipo !== "EGRESO"){
    return jsonResponse({
      success:false,
      message:"El tipo de movimiento debe ser INGRESO o EGRESO."
    });
  }

  const monto = Number(params.monto || 0);
  if(!monto || monto <= 0){
    return jsonResponse({
      success:false,
      message:"Ingresá un monto válido, mayor a 0."
    });
  }

  const motivo = (params.motivo || "").toString().trim();
  if(!motivo){
    return jsonResponse({
      success:false,
      message:"El motivo del movimiento es obligatorio (ej: 'Compra de bolsas', 'Retiro para el banco')."
    });
  }

  const formaPago = ["EFECTIVO","TRANSFERENCIA","TARJETA"].includes(
    String(params.formaPago || "").toUpperCase().trim()
  ) ? String(params.formaPago).toUpperCase().trim() : "EFECTIVO";

  const vendedor = (params.vendedor || "ADMIN").toString();
  const zona = zonaNegocio();
  const ahora = new Date();

  const movimientoId = "MOV-" + ahora.getTime();

  sheet.appendRow([
    movimientoId,
    Utilities.formatDate(ahora, zona, "yyyy-MM-dd"),
    tipo,
    monto,
    motivo,
    vendedor,
    ahora,
    formaPago
  ]);

  return jsonResponse({
    success: true,
    movimientoId: movimientoId
  });
}

/** Verifica y migra automáticamente la columna FORMA_PAGO si no existe */
function autoMigrarMovimientosCaja() {
  try {
    const ss = getSS();
    const sheet = getHoja("MOVIMIENTOS_CAJA");
    if (!sheet || sheet.getLastColumn() === 0) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const tieneFormaPago = headers.some(h => String(h).toUpperCase().trim() === "FORMA_PAGO");
    if (!tieneFormaPago) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue("FORMA_PAGO");
      const ultFila = sheet.getLastRow();
      if (ultFila > 1) sheet.getRange(2, col, ultFila - 1, 1).setValue("EFECTIVO");
      Logger.log("Auto-migración: columna FORMA_PAGO agregada a MOVIMIENTOS_CAJA");
    }
  } catch(e) {
    Logger.log("Error en autoMigrarMovimientosCaja: " + e.message);
  }
}

/** GET ?action=eliminarMovimientoCaja&movimientoId=...&rol=... */
function eliminarMovimientoCaja(e){

  const params = (e && e.parameter) || {};
  const rol = (params.rol || "").toString().trim().toLowerCase();

  if(rol === "vendedor"){
    return jsonResponse({
      success:false,
      message:"Tu usuario (Vendedor) no tiene permiso para eliminar movimientos de caja."
    });
  }

  const ss = getSS();
  const sheet = getHoja("MOVIMIENTOS_CAJA");

  if(!sheet){
    return jsonResponse({
      success:false,
      message:"Falta la hoja 'MOVIMIENTOS_CAJA' en la planilla."
    });
  }

  const movimientoId = (params.movimientoId || "").toString().trim();
  if(!movimientoId){
    return jsonResponse({ success:false, message:"Falta el ID del movimiento a eliminar." });
  }

  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]).trim() === movimientoId){
      sheet.deleteRow(i + 1);
      return jsonResponse({ success:true });
    }
  }

  return jsonResponse({ success:false, message:"No se encontró el movimiento '" + movimientoId + "'." });
}

/**
 * APROBAR PERMISOS Y DEJAR TODO LISTO — ejecutar desde el editor.
 *
 * Cuándo: al instalar un negocio nuevo, y cada vez que se pega un
 * Code.gs nuevo (si trae un servicio de Google que antes no usaba).
 *
 * Cómo: en el editor de Apps Script elegí "autorizarVeekPOS" en el
 * desplegable de funciones y tocá ▶ Ejecutar. La primera vez Google
 * muestra UNA ventana con todos los permisos que usa el sistema:
 * Revisar permisos → tu cuenta → (si avisa "app no verificada")
 * Configuración avanzada → Ir a … → Permitir.
 *
 * Además de pedir los permisos, prueba cada servicio y prepara la
 * planilla (hojas + triggers) si todavía no estaba. El resultado queda
 * en el "Registro de ejecución": ✓ o ✗ por cada servicio.
 *
 * Después de aceptar permisos nuevos: Implementar → Administrar
 * implementaciones → editar → Versión nueva (la URL no cambia).
 *
 * Los permisos que se piden son los de appsscript.json (oauthScopes).
 * Si el código empieza a usar un servicio nuevo de Google, hay que
 * agregar su permiso en ese archivo (la prueba del backend lo avisa).
 */
function autorizarVeekPOS(){
  const lineas = [];
  const probar = (nombre, fn) => {
    try{ const r = fn(); lineas.push("✓ " + nombre + (r ? " — " + r : "")); }
    catch(e){ lineas.push("✗ " + nombre + " — " + e.message); }
  };
  probar("Planilla de Google Sheets", () => getSS().getName());
  probar("Propiedades del script", () => { PropertiesService.getScriptProperties().getProperty("x"); return ""; });
  probar("Caché", () => { CacheService.getScriptCache().put("veek_prueba", "1", 30); return ""; });
  probar("Bloqueo (LockService)", () => { const l = LockService.getScriptLock(); if(l.tryLock(5000)) l.releaseLock(); return ""; });
  probar("Google Drive (fotos y PDF)", () => "carpeta raíz: " + DriveApp.getRootFolder().getName());
  probar("Google Docs (catálogo y PDF)", () => { DocumentApp.getActiveDocument(); return ""; });
  probar("Conexiones externas (UrlFetch)", () => "respuesta " + UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true }).getResponseCode());
  probar("Triggers (tareas programadas)", () => ScriptApp.getProjectTriggers().length + " instalados");
  probar("Correo (informes de error)", () => CORREO_SOPORTE ? "puede enviar " + MailApp.getRemainingDailyQuota() + " correos más hoy · destino " + CORREO_SOPORTE : "sin CORREO_SOPORTE (no se mandan)");
  probar("Hojas y triggers de VeekPOS", () => {
    const r = asegurarInstalacionInicial();
    if(r) return r.message + " · triggers " + r.triggers;
    const yaHay = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "calentarCacheCatalogoProgramado");
    if(!yaHay){ instalarTriggersVeekPOS(); return "planilla ya estaba lista · triggers instalados ahora"; }
    return "planilla y triggers ya estaban listos";
  });
  probar("Carpetas de Drive", () => {
    const vacias = CARPETAS_DRIVE_VEEKPOS.filter(par => !extraerIdCarpetaDrive(leerClaveConfiguracion(par[0], ""))).map(par => par[1]);
    return vacias.length ? "sin configurar: " + vacias.join(", ") + " (para crearlas: ejecutar crearCarpetasDriveFaltantes)" : "todas configuradas";
  });
  const resumen = lineas.join("\n");
  Logger.log("VeekPOS — permisos y preparación:\n" + resumen +
    (resumen.indexOf("✗") === -1 ? "\n\nTodo listo. Si aceptaste permisos nuevos, publicá: Implementar → Administrar implementaciones → Versión nueva."
                                 : "\n\nHay servicios con ✗: revisá el mensaje de cada uno."));
  return resumen;
}

/**
 * HERRAMIENTA DE DIAGNÓSTICO — se mantiene a propósito en el sistema.
 * (Para aprobar TODOS los permisos de una vez, usar autorizarVeekPOS.)
 *
 * Sirve para forzar el popup de autorización de Google cuando se
 * agrega al código una función nueva que usa un servicio que el
 * proyecto nunca usó antes (ej. UrlFetchApp, MailApp, CalendarApp,
 * etc.). Google solo pide autorización para un permiso cuando una
 * ejecución llega de verdad a la línea que lo necesita — ejecutar
 * manualmente la función real (ej. getImagenProxy) sin sus parámetros
 * normales puede "completarse" sin pedir nada, sin haber autorizado
 * en realidad nada nuevo, lo cual confunde mucho a la hora de
 * diagnosticar errores de permisos.
 *
 * Cómo usarla: en el editor de Apps Script, elegí esta función en el
 * desplegable de "Ejecutar" y tocá ▶️. Si aparece el popup de
 * autorización, completalo (Revisar permisos → tu cuenta → Avanzado
 * → Ir al proyecto → Permitir). Si no aparece nada y el "Registro de
 * ejecución" muestra "Código de respuesta: 200", el permiso de
 * UrlFetchApp ya estaba autorizado de antes.
 *
 * Después de autorizar, siempre hay que volver a publicar el script
 * (Implementar → Administrar implementaciones → Nueva versión) para
 * que la implementación Web App use la autorización nueva.
 */
function testAutorizacionUrlFetch(){
  const response = UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true });
  Logger.log("Código de respuesta: " + response.getResponseCode());
}

/**
 * MIGRACIÓN — ejecutar UNA SOLA VEZ desde el editor de Apps Script.
 *
 * Recorre todos los pedidos de la hoja PEDIDOS y crea en CLIENTES
 * una ficha para cada DNI que todavía no tenga una. Toma los datos
 * del pedido más reciente de ese DNI (nombre, teléfono, dirección,
 * empresa). Los clientes que ya existen no se tocan.
 *
 * Cómo usarla:
 *   1. En el editor de Apps Script, elegí "migrarClientesDesdePedidos"
 *      en el desplegable de funciones y tocá ▶️.
 *   2. Revisá el "Registro de ejecución" — va a mostrar cuántas fichas
 *      se crearon y cuáles DNI se saltó porque ya existían.
 *   3. Listo. No hace falta volver a ejecutarla.
 */
function migrarClientesDesdePedidos() {
  invalidarCacheDnisClientes(); // la lista de DNIs que usan los pedidos web se rearma sola

  const ss = getSS();
  const pedidosSheet = getHoja("PEDIDOS");

  if (!pedidosSheet) {
    Logger.log("No se encontró la hoja PEDIDOS.");
    return;
  }

  const clientesSheet = getOrCrearHojaClientes();

  // Construir mapa de DNIs que ya están en CLIENTES
  const clientesData = clientesSheet.getDataRange().getValues();
  const colDni = 3; // columna D = DNI (índice base 0)
  const dniExistentes = new Set(
    clientesData.slice(1)
      .map(fila => normalizarDni(fila[colDni]))
      .filter(dni => dni !== "")
  );

  // Leer todos los pedidos y quedarse con el más reciente por DNI
  const pedidosData = pedidosSheet.getDataRange().getValues();
  const headers = pedidosData[0];

  // Detectar columnas por encabezado (robusto ante cambios de orden)
  const idx = {};
  ["PEDIDO_ID","FECHA","CLIENTE","EMPRESA","DIRECCION","TELEFONO","DNI"].forEach(h => {
    idx[h] = headers.indexOf(h);
  });

  // Agrupar por DNI, quedándose con el pedido de fecha más reciente
  const porDni = {};
  pedidosData.slice(1).forEach(fila => {
    const dni = normalizarDni(fila[idx["DNI"]]);
    if (!dni) return;
    const fecha = new Date(fila[idx["FECHA"]]);
    if (!porDni[dni] || fecha > porDni[dni].fecha) {
      porDni[dni] = {
        fecha,
        nombre:    String(fila[idx["CLIENTE"]]   || "").trim(),
        empresa:   String(fila[idx["EMPRESA"]]   || "").trim(),
        direccion: String(fila[idx["DIRECCION"]] || "").trim(),
        telefono:  String(fila[idx["TELEFONO"]]  || "").trim(),
        dni
      };
    }
  });

  let creados = 0;
  let saltados = 0;

  Object.values(porDni).forEach(c => {
    if (dniExistentes.has(c.dni)) {
      Logger.log("Ya existe: " + c.dni + " (" + c.nombre + ")");
      saltados++;
      return;
    }
    const clienteId = generarIdCorto("CLI", clientesSheet);
    clientesSheet.appendRow([
      clienteId,
      c.nombre,
      "",          // ALIAS
      c.dni,
      c.telefono,
      c.direccion,
      c.empresa,
      "NO",        // A_CREDITO — el admin lo activa a mano
      new Date()   // FECHA_ALTA
    ]);
    dniExistentes.add(c.dni);
    Logger.log("Creado: " + c.dni + " — " + c.nombre);
    creados++;
  });

  Logger.log("=== Migración finalizada: " + creados + " fichas creadas, " + saltados + " ya existían. ===");
}

/* =====================================================================
   DEUDAS EXTRA — cobros fuera de pedidos y POS
   Hoja requerida: "DEUDAS_EXTRA" con encabezados:
   DEUDA_ID, FECHA, CLIENTE_ID, MONEDA, MONTO, CONCEPTO
===================================================================== */

function getOrCrearHojaDeudasExtra(){
  const ss = getSS();
  let sheet = getHoja("DEUDAS_EXTRA");
  if(!sheet){
    sheet = ss.insertSheet("DEUDAS_EXTRA");
    sheet.appendRow(["DEUDA_ID","FECHA","CLIENTE_ID","MONEDA","MONTO","CONCEPTO"]);
  }
  return sheet;
}

/**
 * Registra una deuda extra para un cliente (cobro fuera de pedidos/POS).
 * Parámetros esperados: clienteId, moneda ("ARS"/"USD"), monto, concepto
 */
function crearDeudaExtra(data){

  const clienteId = String(data.clienteId || "").trim();
  const monto = Number(data.monto || 0);
  const moneda = String(data.moneda || "ARS").toUpperCase();
  const concepto = String(data.concepto || "").trim();

  if(!clienteId){
    return jsonResponse({ success:false, message:"Falta el cliente." });
  }
  if(!monto || monto <= 0){
    return jsonResponse({ success:false, message:"Ingresá un monto válido, mayor a 0." });
  }
  if(moneda !== "ARS" && moneda !== "USD"){
    return jsonResponse({ success:false, message:"La moneda debe ser ARS o USD." });
  }
  if(!concepto){
    return jsonResponse({ success:false, message:"El concepto es obligatorio." });
  }

  // Verificar que el cliente existe
  const clientesSheet = getOrCrearHojaClientes();
  const clientesData = rowsToObjects(clientesSheet.getDataRange().getValues());
  const cliente = clientesData.find(c => String(c.CLIENTE_ID) === clienteId);
  if(!cliente){
    return jsonResponse({ success:false, message:"No se encontró el cliente." });
  }

  const sheet = getOrCrearHojaDeudasExtra();
  const deudaId = generarIdCorto("DEU", sheet);

  sheet.appendRow([
    deudaId,
    new Date(),
    clienteId,
    moneda,
    monto,
    concepto
  ]);

  return jsonResponse({ success:true, deudaId:deudaId });
}

/* =====================================================================
   PAPELERA Y CLIENTES — agregado para VeekPOS
   ---------------------------------------------------------------------
   Integrado al final de Code.gs. doPost llama a manejarAccionesExtraPost
   y eliminarProductoPorCodigo delega en eliminarProductoAPapelera.

   Incluye:
   - PAPELERA DE PRODUCTOS: borrar un producto lo mueve a la hoja
     PAPELERA_PRODUCTOS (con todos sus datos) en vez de perderlo; desde
     el panel se puede restaurar, borrar definitivo o vaciar.
   - EDITAR DATOS DEL CLIENTE DE UN PEDIDO y reflejarlo en CLIENTES.
   - GUARDAR CLIENTE DESDE EL RÓTULO MANUAL (sin duplicar por DNI).
===================================================================== */

const HOJA_PAPELERA_PRODUCTOS = "PAPELERA_PRODUCTOS";

/** Datos opcionales del cliente (localidad, provincia, email, notas): solo se tocan los que vienen */
function _guardarDatosExtraCliente(sheet, fila, data){
  const campos = { localidad: "LOCALIDAD", provincia: "PROVINCIA", codigoPostal: "CODIGO_POSTAL", email: "EMAIL", notas: "NOTAS" };
  Object.keys(campos).forEach(k => {
    if(data[k] === undefined || data[k] === null) return;
    const v = String(data[k]).trim().slice(0, 1000);
    const col = obtenerOcrearColumna(sheet, campos[k]);
    if(v || String(sheet.getRange(fila, col).getValue() || "") !== "") sheet.getRange(fila, col).setValue(v);
  });
}
const COLUMNAS_EXTRA_CLIENTES = ["LOCALIDAD", "PROVINCIA", "CODIGO_POSTAL"];

/** Despacha las acciones POST nuevas. Devuelve null si la acción no es de este archivo. */
function manejarAccionesExtraPost(data) {
  switch (data && data.action) {
    case "papeleraProductos":          return getPapeleraProductos();
    case "restaurarProductoPapelera":  return restaurarProductoPapelera(data);
    case "eliminarDefinitivoPapelera": return eliminarDefinitivoPapelera(data);
    case "vaciarPapeleraProductos":    return vaciarPapeleraProductos(data);
    case "editarDatosClientePedido":   return editarDatosClientePedido(data);
    case "guardarClienteDesdeRotulo":  return guardarClienteDesdeRotulo(data);
    case "guardarCupon":               return guardarCupon(data);
    case "eliminarCupon":              return eliminarCupon(data);
    case "validarCupon":               return validarCuponPublico(data);
    default: return null;
  }
}

/* ===================== PAPELERA DE PRODUCTOS ===================== */

/**
 * Toma el script lock, salvo que ya lo tenga esta misma ejecución
 * (conIdempotencia / syncLote marcan _dentroDeLockIdempotencia): el lock
 * de Apps Script NO es reentrante, y pedirlo dos veces haría fallar los
 * borrados que llegan desde la cola offline de la app de escritorio.
 * Devuelve { ok, liberar() }.
 */
function _tomarLockPapelera(ms) {
  if (typeof _dentroDeLockIdempotencia !== "undefined" && _dentroDeLockIdempotencia) {
    return { ok: true, liberar: function () {} };
  }
  const lock = LockService.getScriptLock();
  const ok = lock.tryLock(ms);
  return { ok: ok, liberar: function () { if (ok) lock.releaseLock(); } };
}

function getOrCrearHojaPapeleraProductos() {
  let sheet = getHoja(HOJA_PAPELERA_PRODUCTOS);
  if (!sheet) {
    sheet = getSS().insertSheet(HOJA_PAPELERA_PRODUCTOS);
    sheet.appendRow(["PAPELERA_ID", "FECHA_ELIMINADO", "CODIGO", "PRODUCTO", "CATEGORIA", "PRECIO", "STOCK", "USUARIO", "DATOS_JSON"]);
    sheet.setFrozenRows(1);
    sheet.getRange("C:C").setNumberFormat("@"); // el código siempre como texto (no perder ceros a la izquierda)
    _hojaCache[HOJA_PAPELERA_PRODUCTOS] = sheet;
  }
  return sheet;
}

/** Escribe un valor como TEXTO en una celda (DNI, teléfono, códigos: que Sheets no los convierta en número) */
function _escribirComoTexto(sheet, fila, columna, valor) {
  const celda = sheet.getRange(fila, columna);
  celda.setNumberFormat("@");
  celda.setValue(valor === null || valor === undefined ? "" : String(valor));
}

/**
 * Reemplaza al borrado definitivo: copia la fila completa del producto a
 * PAPELERA_PRODUCTOS (como JSON, con todas sus columnas) y recién después
 * la saca de PRODUCTOS. Acepta e.parameter (GET) o un objeto plano.
 */
function eliminarProductoAPapelera(e) {
  const params = (e && e.parameter) || e || {};
  const codigo = String(params.codigo || "").trim();
  if (!codigo) return jsonResponse({ success: false, message: "Falta el código del producto a eliminar." });

  const lock = _tomarLockPapelera(10000);
  if (!lock.ok) {
    return jsonResponse({ success: false, message: "El sistema está ocupado, probá de nuevo en unos segundos." });
  }

  try {
    const sheet = getSheet();
    if (!sheet) return jsonResponse({ success: false, message: "No se encontró la hoja 'PRODUCTOS'." });

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxCodigo = headers.indexOf("CODIGO") === -1 ? 0 : headers.indexOf("CODIGO");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxCodigo]).trim() !== codigo) continue;

      const producto = {};
      headers.forEach((h, k) => {
        const clave = String(h || "").trim();
        if (!clave) return;
        const v = data[i][k];
        producto[clave] = v instanceof Date ? v.toISOString() : v;
      });

      const papelera = getOrCrearHojaPapeleraProductos();
      const tz = zonaNegocio() || "GMT-3";
      papelera.appendRow([
        "PAP-" + Date.now() + "-" + Math.floor(Math.random() * 900 + 100),
        Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss"),
        codigo,
        String(producto.PRODUCTO || ""),
        String(producto.CATEGORIA || ""),
        Number(producto.PRECIO || 0),
        Number(producto.STOCK || 0),
        String(params.usuario || params.rol || ""),
        JSON.stringify(producto)
      ]);
      _escribirComoTexto(papelera, papelera.getLastRow(), 3, codigo);

      sheet.deleteRow(i + 1);
      invalidarCacheProductos();

      return jsonResponse({ success: true, codigo: codigo, enPapelera: true });
    }

    return jsonResponse({ success: false, message: "No se encontró un producto con el código '" + codigo + "'." });

  } finally {
    lock.liberar();
  }
}

/** Lista de productos en la papelera, el más reciente primero */
function getPapeleraProductos() {
  const sheet = getOrCrearHojaPapeleraProductos();
  const filas = sheet.getDataRange().getValues();
  const tz = zonaNegocio() || "GMT-3";
  const productos = [];

  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    if (!f[0]) continue;
    let obj = {};
    try { obj = JSON.parse(f[8] || "{}"); } catch (err) { obj = {}; }
    productos.push({
      PAPELERA_ID: String(f[0]),
      FECHA_ELIMINADO: f[1] instanceof Date ? Utilities.formatDate(f[1], tz, "yyyy-MM-dd HH:mm") : String(f[1] || ""),
      CODIGO: String(obj.CODIGO !== undefined && obj.CODIGO !== "" ? obj.CODIGO : f[2]),
      PRODUCTO: String(obj.PRODUCTO || f[3] || ""),
      CATEGORIA: String(obj.CATEGORIA || f[4] || ""),
      PRECIO: Number(obj.PRECIO !== undefined ? obj.PRECIO : f[5]) || 0,
      STOCK: Number(obj.STOCK !== undefined ? obj.STOCK : f[6]) || 0,
      IMAGEN: String(obj.IMAGEN || ""),
      USUARIO: String(f[7] || "")
    });
  }

  productos.reverse();
  return jsonResponse({ success: true, total: productos.length, productos: productos });
}

function _buscarFilaPapelera(sheet, papeleraId) {
  const filas = sheet.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(papeleraId)) return { indice: i, fila: filas[i] };
  }
  return null;
}

/** Vuelve a poner el producto en PRODUCTOS con todos sus datos originales */
function restaurarProductoPapelera(data) {
  const papeleraId = String((data && data.papeleraId) || "").trim();
  if (!papeleraId) return jsonResponse({ success: false, message: "Falta el producto a restaurar." });

  const lock = _tomarLockPapelera(10000);
  if (!lock.ok) {
    return jsonResponse({ success: false, message: "El sistema está ocupado, probá de nuevo en unos segundos." });
  }

  try {
    const papelera = getOrCrearHojaPapeleraProductos();
    const encontrado = _buscarFilaPapelera(papelera, papeleraId);
    if (!encontrado) return jsonResponse({ success: false, message: "Ese producto ya no está en la papelera." });

    let producto = {};
    try { producto = JSON.parse(encontrado.fila[8] || "{}"); } catch (err) { producto = {}; }
    const codigo = String(producto.CODIGO !== undefined && producto.CODIGO !== "" ? producto.CODIGO : encontrado.fila[2]).trim();
    if (!codigo) return jsonResponse({ success: false, message: "El producto guardado en la papelera no tiene código." });
    producto.CODIGO = codigo;
    if (!producto.PRODUCTO) producto.PRODUCTO = String(encontrado.fila[3] || "");

    const sheet = getSheet();
    const valores = sheet.getDataRange().getValues();
    let headers = valores[0].map(h => String(h || "").trim());
    const idxCodigo = headers.indexOf("CODIGO") === -1 ? 0 : headers.indexOf("CODIGO");

    const conflictoCod = conflictoCodigoProducto(valores, codigo, -1, false) ||
      (producto.CODIGO_CAJA ? conflictoCodigoProducto(valores, producto.CODIGO_CAJA, -1, true) : null);
    if (conflictoCod) {
      return jsonResponse({
        success: false,
        codigoDuplicado: true,
        message: conflictoCod + " Cambiá o eliminá ese producto antes de restaurar este."
      });
    }

    // Si el producto tenía columnas que la hoja ya no tiene, se agregan (no se pierde nada)
    const faltantes = Object.keys(producto).filter(k => k && headers.indexOf(k) === -1);
    if (faltantes.length > 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1, 1, faltantes.length).setValues([faltantes]);
      headers = headers.concat(faltantes);
    }

    sheet.appendRow(headers.map(h => (h && producto.hasOwnProperty(h)) ? producto[h] : ""));
    _escribirComoTexto(sheet, sheet.getLastRow(), idxCodigo + 1, codigo);

    papelera.deleteRow(encontrado.indice + 1);
    invalidarCacheProductos();

    return jsonResponse({ success: true, codigo: codigo, producto: String(producto.PRODUCTO || "") });

  } finally {
    lock.liberar();
  }
}

/** Borra para siempre un producto de la papelera */
function eliminarDefinitivoPapelera(data) {
  if (String((data && data.rol) || "").toLowerCase() === "vendedor") {
    return jsonResponse({ success: false, message: "Tu usuario (Vendedor) no tiene permiso para borrar definitivamente productos." });
  }
  const papeleraId = String((data && data.papeleraId) || "").trim();
  const papelera = getOrCrearHojaPapeleraProductos();
  const encontrado = _buscarFilaPapelera(papelera, papeleraId);
  if (!encontrado) return jsonResponse({ success: false, message: "Ese producto ya no está en la papelera." });
  papelera.deleteRow(encontrado.indice + 1);
  return jsonResponse({ success: true });
}

/** Vacía toda la papelera */
function vaciarPapeleraProductos(data) {
  if (String((data && data.rol) || "").toLowerCase() === "vendedor") {
    return jsonResponse({ success: false, message: "Tu usuario (Vendedor) no tiene permiso para vaciar la papelera." });
  }
  const papelera = getOrCrearHojaPapeleraProductos();
  const ultima = papelera.getLastRow();
  const cantidad = Math.max(0, ultima - 1);
  if (cantidad > 0) papelera.deleteRows(2, cantidad);
  return jsonResponse({ success: true, eliminados: cantidad });
}

/* ===================== CLIENTES: helpers compartidos ===================== */

/** Agrega LOCALIDAD / PROVINCIA / CODIGO_POSTAL a CLIENTES si no existen (al final, sin tocar nada) */
function asegurarColumnasExtraClientes(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const faltantes = COLUMNAS_EXTRA_CLIENTES.filter(c => headers.indexOf(c) === -1);
  if (faltantes.length > 0) {
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, faltantes.length).setValues([faltantes]);
  }
}

/** Busca un cliente por DNI (normalizado) o, si no hay DNI, por nombre exacto. Devuelve el índice de fila (0-based en valores) o -1 */
function _buscarFilaCliente(valores, headers, dni, nombre) {
  const iDni = headers.indexOf("DNI");
  const iNombre = headers.indexOf("NOMBRE");
  const dniNorm = normalizarDni(dni);
  if (dniNorm && iDni !== -1) {
    for (let i = 1; i < valores.length; i++) {
      if (normalizarDni(valores[i][iDni]) === dniNorm) return i;
    }
    return -1;
  }
  const nombreNorm = String(nombre || "").trim().toLowerCase();
  if (nombreNorm && iNombre !== -1) {
    for (let i = 1; i < valores.length; i++) {
      if (String(valores[i][iNombre] || "").trim().toLowerCase() === nombreNorm) return i;
    }
  }
  return -1;
}

function _camposClienteDesdeDatos(d) {
  return {
    NOMBRE: String(d.nombre || "").trim(),
    DNI: String(d.dni || "").trim(),
    TELEFONO: String(d.telefono || "").trim(),
    DIRECCION: String(d.direccion || "").trim(),
    EMPRESA: String(d.empresa || "").trim(),
    LOCALIDAD: String(d.localidad || "").trim(),
    PROVINCIA: String(d.provincia || "").trim(),
    CODIGO_POSTAL: String(d.codigoPostal || "").trim()
  };
}

const _CAMPOS_CLIENTE_TEXTO = ["DNI", "TELEFONO", "CODIGO_POSTAL"];

/** Actualiza (solo los campos indicados) una fila existente de CLIENTES */
function _actualizarFilaCliente(sheet, headers, indiceFila, campos) {
  Object.keys(campos).forEach(clave => {
    const col = headers.indexOf(clave);
    if (col === -1) return;
    if (_CAMPOS_CLIENTE_TEXTO.indexOf(clave) !== -1) _escribirComoTexto(sheet, indiceFila + 1, col + 1, campos[clave]);
    else sheet.getRange(indiceFila + 1, col + 1).setValue(campos[clave]);
  });
}

/** Crea una fila nueva en CLIENTES (a crédito: NO) y devuelve su CLIENTE_ID */
function _crearFilaCliente(sheet, headers, campos) {
  const clienteId = generarIdCorto("CLI", sheet);
  const completo = Object.assign({ CLIENTE_ID: clienteId, ALIAS: "", A_CREDITO: "NO", FECHA_ALTA: new Date() }, campos);
  sheet.appendRow(headers.map(h => completo.hasOwnProperty(h) ? completo[h] : ""));
  const fila = sheet.getLastRow();
  _CAMPOS_CLIENTE_TEXTO.forEach(clave => {
    const col = headers.indexOf(clave);
    if (col !== -1) _escribirComoTexto(sheet, fila, col + 1, completo[clave]);
  });
  return clienteId;
}

function _leerHojaClientesConExtras() {
  const sheet = getOrCrearHojaClientes();
  asegurarColumnasExtraClientes(sheet);
  const valores = sheet.getDataRange().getValues();
  const headers = valores[0].map(h => String(h || "").trim().toUpperCase());
  return { sheet, valores, headers };
}

/* ===================== EDITAR DATOS DEL CLIENTE DE UN PEDIDO ===================== */

/**
 * Edita los datos de envío/cliente de un pedido (cualquier estado) y
 * refleja el cambio en la hoja CLIENTES:
 * - Si el DNI nuevo ya pertenece a otro cliente, se actualiza ESE cliente.
 * - Si no, se actualiza el cliente que tenía el DNI anterior.
 * - Si no existía ninguno, se crea (a crédito: NO).
 * Si el pedido tenía PDF generado, se regenera con los datos nuevos.
 */
function editarDatosClientePedido(data) {
  const pedidoId = String((data && data.pedidoId) || "").trim();
  const nombre = String((data && data.nombre) || "").trim();
  if (!pedidoId) return jsonResponse({ success: false, message: "Falta el pedido." });
  if (!nombre) return jsonResponse({ success: false, message: "El nombre del cliente es obligatorio." });

  const lock = _tomarLockPapelera(15000);
  if (!lock.ok) {
    return jsonResponse({ success: false, message: "El sistema está ocupado, probá de nuevo en unos segundos." });
  }

  let resultadoCliente, filaPedido, valoresPedido, C;
  const pedidosSheet = getHoja("PEDIDOS");

  try {
    valoresPedido = pedidosSheet.getDataRange().getValues();
    const headers = valoresPedido[0].map(h => String(h || "").trim().toUpperCase());
    const col = (nombres, porDefecto) => {
      for (let k = 0; k < nombres.length; k++) {
        const i = headers.indexOf(nombres[k]);
        if (i !== -1) return i;
      }
      return porDefecto;
    };
    // Los índices por defecto son los mismos que usa guardarPedido() al hacer appendRow
    C = {
      nombre:       col(["CLIENTE", "NOMBRE"], 2),
      empresa:      col(["EMPRESA", "TRANSPORTE"], 3),
      direccion:    col(["DIRECCION"], 4),
      telefono:     col(["TELEFONO"], 5),
      dni:          col(["DNI"], 6),
      pdf:          col(["PDF_URL"], 9),
      codigoPostal: col(["CODIGO_POSTAL", "CODIGOPOSTAL", "CP"], 10),
      localidad:    col(["LOCALIDAD"], 11),
      provincia:    col(["PROVINCIA"], 12)
    };

    filaPedido = -1;
    for (let i = 1; i < valoresPedido.length; i++) {
      if (String(valoresPedido[i][0]) === pedidoId) { filaPedido = i; break; }
    }
    if (filaPedido === -1) return jsonResponse({ success: false, message: "Pedido no encontrado." });

    const dniAnterior = String(valoresPedido[filaPedido][C.dni] || "");
    const nuevos = {
      nombre: nombre,
      empresa: String(data.empresa || "").trim(),
      direccion: String(data.direccion || "").trim(),
      telefono: String(data.telefono || "").trim(),
      dni: String(data.dni || "").trim(),
      codigoPostal: String(data.codigoPostal || "").trim(),
      localidad: String(data.localidad || "").trim(),
      provincia: String(data.provincia || "").trim()
    };

    ["nombre", "empresa", "direccion", "localidad", "provincia"].forEach(k => {
      pedidosSheet.getRange(filaPedido + 1, C[k] + 1).setValue(nuevos[k]);
    });
    ["telefono", "dni", "codigoPostal"].forEach(k => {
      _escribirComoTexto(pedidosSheet, filaPedido + 1, C[k] + 1, nuevos[k]);
    });

    // ---- Reflejar en CLIENTES ----
    const cli = _leerHojaClientesConExtras();
    const campos = _camposClienteDesdeDatos(nuevos);
    const dniNuevoNorm = normalizarDni(nuevos.dni);
    const dniViejoNorm = normalizarDni(dniAnterior);

    let indice = -1;
    if (dniNuevoNorm && dniNuevoNorm !== dniViejoNorm) indice = _buscarFilaCliente(cli.valores, cli.headers, nuevos.dni, "");
    if (indice === -1 && dniViejoNorm) indice = _buscarFilaCliente(cli.valores, cli.headers, dniAnterior, "");
    if (indice === -1) indice = _buscarFilaCliente(cli.valores, cli.headers, nuevos.dni, nuevos.nombre);

    if (indice !== -1) {
      _actualizarFilaCliente(cli.sheet, cli.headers, indice, campos);
      resultadoCliente = { accion: "actualizado", clienteId: String(cli.valores[indice][cli.headers.indexOf("CLIENTE_ID")] || "") };
    } else {
      resultadoCliente = { accion: "creado", clienteId: _crearFilaCliente(cli.sheet, cli.headers, campos) };
    }

    invalidarCachePolling();

  } finally {
    lock.liberar();
  }

  // ---- Regenerar el PDF del pedido (fuera del lock: puede tardar) ----
  let pdfError = false;
  try {
    const fila = valoresPedido[filaPedido];
    if (String(fila[C.pdf] || "")) {
      const detalleSheet = getHoja("DETALLE_PEDIDOS");
      const carrito = detalleSheet
        ? detalleSheet.getDataRange().getValues().slice(1)
            .filter(r => String(r[0]) === pedidoId)
            .map(r => ({ CODIGO: r[1], PRODUCTO: r[2], cantidad: Number(r[3]) || 0, PRECIO: Number(r[4]) || 0, COLOR: r[6] || "" }))
        : [];
      const hp = valoresPedido[0].map(h => String(h || "").trim().toUpperCase());
      const v = (nombreCol, porDefecto) => {
        const i = hp.indexOf(nombreCol);
        return i !== -1 ? fila[i] : fila[porDefecto];
      };
      const total = Number(v("TOTAL", 7)) || 0;
      const pdfUrl = generarPDFPedido(pedidoId, v("FECHA", 1), {
        nombre: nombre,
        empresa: String(data.empresa || "").trim(),
        direccion: String(data.direccion || "").trim(),
        telefono: String(data.telefono || "").trim(),
        dni: String(data.dni || "").trim(),
        codigoPostal: String(data.codigoPostal || "").trim(),
        localidad: String(data.localidad || "").trim(),
        provincia: String(data.provincia || "").trim(),
        total: total,
        subtotal: Number(v("SUBTOTAL", 16)) || total,
        descuento: Number(v("DESCUENTO", 17)) || 0,
        descuentoEtiqueta: String(v("DESCUENTO_ETIQUETA", 18) || ""),
        carrito: carrito
      });
      pedidosSheet.getRange(filaPedido + 1, C.pdf + 1).setValue(pdfUrl);
    }
  } catch (err) {
    Logger.log("editarDatosClientePedido: no se pudo regenerar el PDF: " + err);
    pdfError = true;
  }

  return jsonResponse({
    success: true,
    cliente: resultadoCliente,
    pdfError: pdfError
  });
}

/* ===================== GUARDAR CLIENTE DESDE EL RÓTULO MANUAL ===================== */

/**
 * Crea el cliente con los datos del rótulo (incluye localidad, provincia y
 * código postal). Si ya existe uno con ese DNI (o, sin DNI, con el mismo
 * nombre), NO lo duplica: devuelve el existente.
 */
function guardarClienteDesdeRotulo(data) {
  const nombre = String((data && data.nombre) || "").trim();
  if (!nombre) return jsonResponse({ success: false, message: "Falta el nombre del cliente." });

  const lock = _tomarLockPapelera(10000);
  if (!lock.ok) {
    return jsonResponse({ success: false, message: "El sistema está ocupado, probá de nuevo en unos segundos." });
  }

  try {
    const cli = _leerHojaClientesConExtras();
    const indice = _buscarFilaCliente(cli.valores, cli.headers, data.dni, nombre);

    if (indice !== -1) {
      return jsonResponse({
        success: true,
        existente: true,
        clienteId: String(cli.valores[indice][cli.headers.indexOf("CLIENTE_ID")] || ""),
        nombre: String(cli.valores[indice][cli.headers.indexOf("NOMBRE")] || nombre)
      });
    }

    const campos = _camposClienteDesdeDatos(data);
    const clienteId = _crearFilaCliente(cli.sheet, cli.headers, campos);
    return jsonResponse({ success: true, existente: false, clienteId: clienteId, nombre: nombre });

  } finally {
    lock.liberar();
  }
}


/* =====================================================================
   FACTURACIÓN ELECTRÓNICA (ARCA) — registro de comprobantes
   ---------------------------------------------------------------------
   Las facturas las emite la app de escritorio (tiene el certificado del
   negocio y habla directo con ARCA). Acá solo queda el registro, para
   verlas desde cualquier caja o desde el panel web y para los informes.
   - registrarFactura (POST): alta o actualización por FACTURA_ID
     (idempotente: reintentar no duplica filas).
   - facturasPorRefs: las facturas de una lista de ventas/pedidos.
===================================================================== */
const COLUMNAS_FACTURAS = [
  "FACTURA_ID", "REF", "REF_ORIGINAL", "TIPO_DOC", "TIPO", "NUMERO", "PTO_VTA", "NRO", "FECHA",
  "CAE", "CAE_VTO", "TOTAL", "NETO", "IVA", "DOC_TIPO", "DOC_NRO", "RECEPTOR", "CONDICION_IVA",
  "ESTADO", "AMBIENTE", "QR_URL", "DATOS", "REGISTRADA_EN"
];

function getOrCrearHojaFacturas(){
  let sheet = getHoja("FACTURAS");
  if(!sheet){
    sheet = getSS().insertSheet("FACTURAS");
    sheet.appendRow(COLUMNAS_FACTURAS);
    sheet.setFrozenRows(1);
    // Números largos (CAE, CUIT) como texto: si no, Sheets los redondea
    sheet.getRange("A:W").setNumberFormat("@");
    _hojaCache["FACTURAS"] = sheet;
  }
  return sheet;
}

function _filaFactura(f){
  const r = f.receptor || {};
  const datos = {
    clase: f.clase, tipoNombre: f.tipoNombre, ivas: f.ivas || [], ivaContenido: f.ivaContenido || 0,
    lineas: (f.lineas || []).slice(0, 200), asociado: f.asociado || null, emisor: f.emisor || null,
    receptor: r
  };
  let json = JSON.stringify(datos);
  if(json.length > 45000){ datos.lineas = []; json = JSON.stringify(datos); } // límite de celda
  return [
    String(f.id || ""), String(f.ref || ""), String(f.refOriginal || ""), String(f.tipoDoc || ""),
    String(f.cbteTipo || ""), String(f.numero || ""), String(f.ptoVta || ""), String(f.nro || ""), String(f.fecha || ""),
    String(f.cae || ""), String(f.caeVto || ""), Number(f.total) || 0, Number(f.impNeto) || 0, Number(f.impIVA) || 0,
    String(r.docTipo || ""), String(r.docNro || ""), String(r.nombre || ""), String(r.condicionIva || ""),
    String(f.estado || ""), String(f.ambiente || ""), String(f.qrUrl || ""), json, new Date().toISOString()
  ];
}

function _facturaDesdeFila(header, fila){
  const o = {};
  header.forEach((h, i) => { o[h] = fila[i]; });
  let datos = {};
  try { datos = JSON.parse(o.DATOS || "{}"); } catch(e) {}
  return {
    id: String(o.FACTURA_ID), ref: String(o.REF), refOriginal: String(o.REF_ORIGINAL || ""), tipoDoc: String(o.TIPO_DOC),
    cbteTipo: Number(o.TIPO), numero: String(o.NUMERO), ptoVta: Number(o.PTO_VTA), nro: Number(o.NRO), fecha: String(o.FECHA),
    cae: String(o.CAE), caeVto: String(o.CAE_VTO), total: Number(o.TOTAL), impNeto: Number(o.NETO), impIVA: Number(o.IVA),
    estado: String(o.ESTADO), ambiente: String(o.AMBIENTE), qrUrl: String(o.QR_URL || ""),
    clase: datos.clase || "", tipoNombre: datos.tipoNombre || "", ivas: datos.ivas || [], ivaContenido: datos.ivaContenido || 0,
    lineas: datos.lineas || [], asociado: datos.asociado || null, emisor: datos.emisor || null, receptor: datos.receptor || null
  };
}

function registrarFactura(data){
  const f = data && data.factura;
  if(!f || !f.id || !f.ref) return jsonResponse({ success:false, message:"Faltan los datos de la factura." });
  if(f.estado !== "EMITIDA") return jsonResponse({ success:false, message:"Solo se registran comprobantes con CAE." });

  const lock = LockService.getScriptLock();
  if(!lock.tryLock(15000)) return jsonResponse({ success:false, message:"El sistema está ocupado, se reintenta solo." });
  try{
    const sheet = getOrCrearHojaFacturas();
    const fila = _filaFactura(f);
    const ultima = sheet.getLastRow();
    if(ultima > 1){
      const ids = sheet.getRange(2, 1, ultima - 1, 1).getValues();
      for(let i = 0; i < ids.length; i++){
        if(String(ids[i][0]) === String(f.id)){
          sheet.getRange(i + 2, 1, 1, fila.length).setValues([fila]);
          return jsonResponse({ success:true, actualizada:true });
        }
      }
    }
    sheet.getRange(ultima + 1, 1, 1, fila.length).setNumberFormat("@").setValues([fila]);
    return jsonResponse({ success:true });
  } finally {
    lock.releaseLock();
  }
}

/** refs: lista separada por comas (GET) o array (POST). Devuelve { ref: factura (+ notaCredito) } */
function facturasPorRefs(data){
  let refs = data && data.refs;
  if(typeof refs === "string") refs = refs.split(",");
  refs = (refs || []).map(r => String(r).trim()).filter(Boolean).slice(0, 500);
  const sheet = getHoja("FACTURAS");
  const salida = {};
  if(!sheet || sheet.getLastRow() < 2 || !refs.length) return jsonResponse({ success:true, facturas: salida });
  const buscados = {};
  refs.forEach(r => { buscados[r] = true; buscados["NC:" + r] = true; });
  const valores = sheet.getDataRange().getValues();
  const header = valores[0].map(String);
  const iRef = header.indexOf("REF");
  const ncs = {};
  for(let i = 1; i < valores.length; i++){
    const ref = String(valores[i][iRef]);
    if(!buscados[ref]) continue;
    const f = _facturaDesdeFila(header, valores[i]);
    if(ref.indexOf("NC:") === 0) ncs[ref.slice(3)] = f; else salida[ref] = f;
  }
  Object.keys(ncs).forEach(r => { if(salida[r]) salida[r].notaCredito = ncs[r]; });
  return jsonResponse({ success:true, facturas: salida });
}

/* =====================================================================
   DEPÓSITO: DÓLAR DEL DÍA Y PRECIOS EN DÓLARES
   ---------------------------------------------------------------------
   - El dólar del depósito se fija A MANO (Configuración o el widget del
     dólar) y queda un historial en la hoja DOLAR_HISTORIAL.
   - Un producto puede tener precio en dólares (PRECIO_USD y, si vende
     por caja, PRECIO_CAJA_USD). Su PRECIO en pesos se calcula con el
     dólar del día y se guarda en la columna PRECIO de siempre: así el
     POS, los pedidos, los reportes y la app sin conexión siguen
     trabajando en pesos sin saber nada de esto. Cada vez que cambia el
     dólar se recalculan.
===================================================================== */

const HOJA_DOLAR_HISTORIAL = "DOLAR_HISTORIAL";

function leerDolarDeposito(){
  const valor = Number(String(leerClaveConfiguracion("DOLAR_DEPOSITO", "") || "").replace(",", "."));
  return {
    valor: valor > 0 ? valor : 0,
    fecha: String(leerClaveConfiguracion("DOLAR_DEPOSITO_FECHA", "") || ""),
    usuario: String(leerClaveConfiguracion("DOLAR_DEPOSITO_USUARIO", "") || "")
  };
}

function getDolarDeposito(){
  const actual = leerDolarDeposito();
  const hoja = getHoja(HOJA_DOLAR_HISTORIAL);
  let historial = [];
  if(hoja && hoja.getLastRow() > 1){
    const desde = Math.max(2, hoja.getLastRow() - 59);
    historial = hoja.getRange(desde, 1, hoja.getLastRow() - desde + 1, 4).getValues()
      .map(f => ({ fecha: f[0] instanceof Date ? f[0].toISOString() : String(f[0]), valor: Number(f[1]) || 0, anterior: Number(f[2]) || 0, usuario: String(f[3] || "") }))
      .reverse();
  }
  return jsonResponse({ success:true, valor: actual.valor, fecha: actual.fecha, usuario: actual.usuario, historial: historial });
}

function guardarDolarDeposito(data){
  const valor = Math.round(Number(String((data && data.valor) || "").replace(",", ".")) * 100) / 100;
  if(!(valor > 0) || valor > 10000000) return jsonResponse({ success:false, message:"Ingresá un valor de dólar válido." });
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const anterior = leerDolarDeposito().valor;
    const usuario = (_sesionActual && _sesionActual.u) || String((data && data.usuario) || "");
    const ahora = new Date();
    fijarClaveConfiguracion("DOLAR_DEPOSITO", String(valor));
    fijarClaveConfiguracion("DOLAR_DEPOSITO_FECHA", ahora.toISOString());
    fijarClaveConfiguracion("DOLAR_DEPOSITO_USUARIO", usuario);
    let hoja = getHoja(HOJA_DOLAR_HISTORIAL);
    if(!hoja){
      hoja = getSS().insertSheet(HOJA_DOLAR_HISTORIAL);
      hoja.appendRow(["FECHA", "VALOR", "ANTERIOR", "USUARIO"]);
    }
    hoja.appendRow([ahora, valor, anterior || "", usuario]);
    const actualizados = recalcularPreciosEnDolares(valor);
    _invalidarCacheConfigNegocio();
    return jsonResponse({ success:true, valor: valor, anterior: anterior, fecha: ahora.toISOString(), productosActualizados: actualizados });
  } finally {
    lock.releaseLock();
  }
}

/** Precio en pesos de un precio en dólares (sin centavos) */
function _pesosDeDolares(usd, dolar){
  return Math.round(Number(usd) * Number(dolar));
}

/**
 * Recalcula PRECIO (y PRECIO_CAJA) de los productos con precio en
 * dólares. Escribe solo si algo cambió. Devuelve cuántos productos cambiaron.
 */
function recalcularPreciosEnDolares(dolar){
  if(!(dolar > 0)) return 0;
  const hoja = getHoja("PRODUCTOS");
  if(!hoja || hoja.getLastRow() < 2) return 0;
  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(h => String(h || "").trim());
  const iUsd = enc.indexOf("PRECIO_USD"), iCajaUsd = enc.indexOf("PRECIO_CAJA_USD");
  const iPrecio = enc.indexOf("PRECIO"), iCaja = enc.indexOf("PRECIO_CAJA");
  if(iUsd === -1 && iCajaUsd === -1) return 0;
  let cambiados = 0;
  const precios = [], cajas = [];
  for(let r = 1; r < datos.length; r++){
    const f = datos[r];
    let p = f[iPrecio], c = iCaja !== -1 ? f[iCaja] : "";
    let cambio = false;
    if(iUsd !== -1 && Number(f[iUsd]) > 0){
      const nuevo = _pesosDeDolares(f[iUsd], dolar);
      if(Number(p) !== nuevo){ p = nuevo; cambio = true; }
    }
    if(iCaja !== -1 && iCajaUsd !== -1 && Number(f[iCajaUsd]) > 0){
      const nuevo = _pesosDeDolares(f[iCajaUsd], dolar);
      if(Number(c) !== nuevo){ c = nuevo; cambio = true; }
    }
    if(cambio) cambiados++;
    precios.push([p]); cajas.push([c]);
  }
  if(!cambiados) return 0;
  hoja.getRange(2, iPrecio + 1, precios.length, 1).setValues(precios);
  if(iCaja !== -1 && iCajaUsd !== -1) hoja.getRange(2, iCaja + 1, cajas.length, 1).setValues(cajas);
  invalidarCacheProductos();
  return cambiados;
}

/**
 * Fija (o quita, con 0/vacío) el precio en dólares de uno o varios
 * productos y recalcula su precio en pesos con el dólar del día.
 * data.precios = [{ codigo, precioUsd, precioCajaUsd }]
 */
function guardarPreciosUsd(data){
  let lista = data && data.precios;
  if(typeof lista === "string"){ try{ lista = JSON.parse(lista); }catch(e){ lista = null; } }
  if(!Array.isArray(lista) || !lista.length) return jsonResponse({ success:false, message:"No hay precios para guardar." });
  const dolar = leerDolarDeposito().valor;
  const conDolares = lista.some(x => Number(x.precioUsd) > 0 || Number(x.precioCajaUsd) > 0);
  if(conDolares && !(dolar > 0)) return jsonResponse({ success:false, message:"Primero fijá el dólar del día (Configuración → Negocio o el widget del dólar)." });
  const hoja = getHoja("PRODUCTOS");
  if(!hoja) return jsonResponse({ success:false, message:"No existe la hoja PRODUCTOS." });
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const colUsd = obtenerOcrearColumna(hoja, "PRECIO_USD");
    const colCajaUsd = obtenerOcrearColumna(hoja, "PRECIO_CAJA_USD");
    const datos = hoja.getDataRange().getValues();
    const enc = datos[0].map(h => String(h || "").trim());
    const iCod = enc.indexOf("CODIGO"), iPrecio = enc.indexOf("PRECIO"), iCaja = enc.indexOf("PRECIO_CAJA");
    const filaPorCodigo = {};
    for(let r = 1; r < datos.length; r++) filaPorCodigo[normalizarCodigo(datos[r][iCod])] = r;
    const noEncontrados = [];
    let guardados = 0;
    lista.forEach(x => {
      const r = filaPorCodigo[normalizarCodigo(x.codigo)];
      if(r === undefined){ noEncontrados.push(String(x.codigo)); return; }
      const usd = Math.round(Number(x.precioUsd || 0) * 100) / 100;
      const cajaUsd = Math.round(Number(x.precioCajaUsd || 0) * 100) / 100;
      hoja.getRange(r + 1, colUsd).setValue(usd > 0 ? usd : "");
      hoja.getRange(r + 1, colCajaUsd).setValue(cajaUsd > 0 ? cajaUsd : "");
      if(usd > 0) hoja.getRange(r + 1, iPrecio + 1).setValue(_pesosDeDolares(usd, dolar));
      if(cajaUsd > 0 && iCaja !== -1) hoja.getRange(r + 1, iCaja + 1).setValue(_pesosDeDolares(cajaUsd, dolar));
      guardados++;
    });
    invalidarCacheProductos();
    return jsonResponse({ success: guardados > 0, guardados: guardados, noEncontrados: noEncontrados, dolar: dolar,
      message: guardados ? "" : "No se encontró ningún producto con esos códigos." });
  } finally {
    lock.releaseLock();
  }
}

/** La configuración del negocio está cacheada 5 minutos: que se vea el cambio en el acto */
function _invalidarCacheConfigNegocio(){
  try{ CacheService.getScriptCache().removeAll(["config_negocio:n","config_negocio:0","config_negocio:1"]); }catch(e){}
}

/* ---------- Despacho de las acciones del depósito (GET y POST) ---------- */
function manejarAccionesDeposito(data){
  switch (data && data.action) {
    case "dolarDeposito":         return getDolarDeposito();
    case "guardarDolarDeposito":  return guardarDolarDeposito(data);
    case "guardarPreciosUsd":     return guardarPreciosUsd(data);
    case "guardarCotizacion":     return guardarCotizacion(data);
    case "cotizaciones":          return listarCotizaciones(data);
    case "cotizacion":            return obtenerCotizacion(data);
    case "estadoCotizacion":      return cambiarEstadoCotizacion(data);
    case "traspasosParaLocal":    return traspasosParaLocal(data);
    case "confirmarRecepcionTraspaso": return confirmarRecepcionTraspaso(data);
    case "guardarConfigDocumentos": return guardarConfigDocumentos(data);
    // traspasos — en el depósito
    case "crearTraspaso":           return crearTraspaso(data);
    case "guardarCodigosLocal":     return guardarCodigosLocal(data);
    case "traspasos":               return listarTraspasos(data);
    case "anularTraspaso":          return anularTraspaso(data);
    case "claveTraspasos":          return claveTraspasos(data);
    // traspasos — en el local
    case "vincularDeposito":        return vincularDeposito(data);
    case "traspasosDelDeposito":    return traspasosDelDeposito(data);
    case "recibirTraspaso":         return recibirTraspaso(data);
    default: return null;
  }
}


/* =====================================================================
   DEPÓSITO: COBRO EN DÓLARES
   Una venta "DOLARES" guarda cuántos dólares entraron (MONTO_USD) y el
   dólar usado (DOLAR_USADO). Si los dólares no alcanzan, el resto se
   cobra en pesos; si sobran, el vuelto se da en pesos. Para la caja:
   efectivo en pesos = TOTAL − MONTO_USD × DOLAR_USADO (puede ser negativo).
===================================================================== */

function _pagoEnDolaresDeParams(params){
  if(String(params.formaPago || "").toUpperCase() !== "DOLARES") return null;
  const montoUsd = Math.round(Number(String(params.montoUsd || 0).replace(",", ".")) * 100) / 100;
  const tipoCambio = Number(String(params.tipoCambio || 0).replace(",", "."));
  if(!(montoUsd > 0) || !(tipoCambio > 0)) return null;
  return { montoUsd: montoUsd, tipoCambio: tipoCambio };
}

/** Pagos de clientes hechos en dólares y en efectivo ese día (no los cobros automáticos de pedidos) */
function _pagosCreditoUsdEfectivoDelDia(fecha){
  try{
    const sheet = getOrCrearHojaPagosCredito();
    const data = sheet.getDataRange().getValues();
    data.shift();
    const zona = zonaNegocio();
    let usd = 0;
    data.forEach(row => {
      if(String(row[3] || "").toUpperCase().trim() !== "USD") return;
      if(/^Cobro autom[aá]tico/i.test(String(row[10] || "").trim())) return;
      const fp = String(row[9] || "EFECTIVO").toUpperCase().trim();
      if(fp && fp !== "EFECTIVO" && fp !== "DOLARES") return;
      let f = ""; try{ f = Utilities.formatDate(new Date(row[1]), zona, "yyyy-MM-dd"); }catch(e){}
      if(f === fecha) usd += Number(row[4] || 0);
    });
    return usd;
  }catch(e){ return 0; }
}


/* =====================================================================
   COTIZACIONES (presupuestos impresos en A4)
   Hoja COTIZACIONES, una fila por cotización; los ítems van en ITEMS
   como JSON [{codigo, producto, cantidad, precio}] (precio en la moneda
   de la cotización). Número correlativo COT-0001 (contador con lock).
   Estados: PENDIENTE · ACEPTADA · CONVERTIDA · RECHAZADA · ANULADA.
   "VENCIDA" no se guarda: se calcula al leer (PENDIENTE y pasó la validez).
===================================================================== */

/** Datos que salen en los documentos A4 (cotización, boleta de entrega, estado de cuenta) */
function guardarConfigDocumentos(data){
  const campos = { cuit: "DOC_CUIT", condicionIva: "DOC_CONDICION_IVA", email: "DOC_EMAIL",
    condicionesCotizacion: "DOC_CONDICIONES_COTIZACION", validezCotizacion: "DOC_VALIDEZ_COTIZACION", leyendaBoleta: "DOC_LEYENDA_BOLETA" };
  Object.keys(campos).forEach(k => {
    if(data[k] === undefined) return;
    let v = String(data[k]).trim().slice(0, 1500);
    if(k === "validezCotizacion") v = String(Math.max(0, Math.min(365, Math.round(Number(v) || 0))));
    fijarClaveConfiguracion(campos[k], v);
  });
  _invalidarCacheConfigNegocio();
  return jsonResponse({ success:true });
}

const HOJA_COTIZACIONES = "COTIZACIONES";
const COLS_COTIZACIONES = ["COTIZACION_ID", "FECHA", "CLIENTE_ID", "CLIENTE", "CLIENTE_DATOS", "MONEDA", "TIPO_CAMBIO",
  "SUBTOTAL", "DESCUENTO", "TOTAL", "VALIDEZ_DIAS", "ESTADO", "CONVERTIDA_A", "VENDEDOR", "OBSERVACIONES", "ITEMS", "ACTUALIZADA"];
const ESTADOS_COTIZACION = ["PENDIENTE", "ACEPTADA", "CONVERTIDA", "RECHAZADA", "ANULADA"];

function _hojaCotizaciones(){
  let hoja = getHoja(HOJA_COTIZACIONES);
  if(!hoja){
    hoja = getSS().insertSheet(HOJA_COTIZACIONES);
    hoja.appendRow(COLS_COTIZACIONES);
  }
  return hoja;
}

/** Próximo número de una serie (COT, TRA...), correlativo; llamar con el lock tomado */
function _siguienteNumeroSerie(serie, hoja){
  const props = PropertiesService.getScriptProperties();
  let n = Number(props.getProperty("SERIE_" + serie) || 0);
  if(!n && hoja && hoja.getLastRow() > 1){
    // por si se borró la propiedad: se sigue desde el mayor número existente
    hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().forEach(f => {
      const m = String(f[0] || "").match(/(\d+)$/); if(m) n = Math.max(n, Number(m[1]));
    });
  }
  n++;
  props.setProperty("SERIE_" + serie, String(n));
  return serie + "-" + String(n).padStart(4, "0");
}

function _redondearMoneda(v, moneda){
  return moneda === "USD" ? Math.round(Number(v) * 100) / 100 : Math.round(Number(v));
}

function _filaACotizacion(f, conItems){
  const fecha = f[1] instanceof Date ? f[1] : new Date(f[1]);
  const validez = Number(f[10]) || 0;
  let estado = String(f[11] || "PENDIENTE").toUpperCase();
  const vence = validez > 0 && !isNaN(fecha) ? new Date(fecha.getTime() + validez * 86400000) : null;
  const vencida = estado === "PENDIENTE" && vence && vence < new Date();
  let datos = {}; try{ datos = JSON.parse(f[4] || "{}"); }catch(e){}
  let items = []; try{ items = JSON.parse(f[15] || "[]"); }catch(e){}
  const c = {
    cotizacionId: String(f[0]), fecha: isNaN(fecha) ? String(f[1]) : fecha.toISOString(),
    clienteId: String(f[2] || ""), cliente: String(f[3] || ""), clienteDatos: datos,
    moneda: String(f[5] || "ARS").toUpperCase() === "USD" ? "USD" : "ARS", tipoCambio: Number(f[6]) || 0,
    subtotal: Number(f[7]) || 0, descuento: Number(f[8]) || 0, total: Number(f[9]) || 0,
    validezDias: validez, vence: vence ? vence.toISOString() : "",
    estado: vencida ? "VENCIDA" : estado, convertidaA: String(f[12] || ""),
    vendedor: String(f[13] || ""), observaciones: String(f[14] || ""), cantidadItems: items.length
  };
  if(conItems) c.items = items;
  return c;
}

function guardarCotizacion(data){
  const moneda = String(data.moneda || "ARS").toUpperCase() === "USD" ? "USD" : "ARS";
  const tipoCambio = Number(String(data.tipoCambio || 0).replace(",", ".")) || 0;
  let items = data.items;
  if(typeof items === "string"){ try{ items = JSON.parse(items); }catch(e){ items = []; } }
  items = (Array.isArray(items) ? items : []).map(it => ({
    codigo: String(it.codigo || it.CODIGO || "").trim(),
    producto: String(it.producto || it.PRODUCTO || "").trim(),
    cantidad: Number(it.cantidad) || 0,
    precio: _redondearMoneda(Number(it.precio) || 0, moneda)
  })).filter(it => it.producto && it.cantidad > 0);
  if(!items.length) return jsonResponse({ success:false, message:"La cotización no tiene productos." });
  if(moneda === "USD" && !(tipoCambio > 0)) return jsonResponse({ success:false, message:"Falta el dólar de la cotización." });

  const cliente = data.cliente || {};
  const nombre = String(cliente.nombre || "").trim();
  if(!nombre) return jsonResponse({ success:false, message:"Falta el cliente." });
  const datosCliente = {};
  ["direccion", "localidad", "provincia", "telefono", "cuit", "email", "empresa"].forEach(k => { if(cliente[k]) datosCliente[k] = String(cliente[k]).trim(); });

  const subtotal = _redondearMoneda(items.reduce((a, it) => a + it.precio * it.cantidad, 0), moneda);
  const descuento = Math.max(0, Math.min(subtotal, _redondearMoneda(Number(data.descuento) || 0, moneda)));
  const total = _redondearMoneda(subtotal - descuento, moneda);
  const validez = Math.max(0, Math.min(365, Math.round(Number(data.validezDias)))) || 0;
  const vendedor = String(data.vendedor || "").trim().slice(0, 80) || (_sesionActual && _sesionActual.u) || "";

  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const hoja = _hojaCotizaciones();
    const id = String(data.cotizacionId || "").trim();
    let fila = -1, fechaOriginal = new Date(), estado = "PENDIENTE";
    if(id){
      const ids = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 12).getValues() : [];
      for(let i = 0; i < ids.length; i++) if(String(ids[i][0]) === id){ fila = i + 2; fechaOriginal = ids[i][1]; estado = String(ids[i][11] || "PENDIENTE"); break; }
      if(fila === -1) return jsonResponse({ success:false, message:"No se encontró la cotización " + id });
      if(estado === "CONVERTIDA" || estado === "ANULADA") return jsonResponse({ success:false, message:"Una cotización " + estado.toLowerCase() + " ya no se puede modificar (podés duplicarla)." });
    }
    const nuevoId = id || _siguienteNumeroSerie("COT", hoja);
    const valores = [nuevoId, fila > 0 ? fechaOriginal : new Date(), String(cliente.id || ""), nombre, JSON.stringify(datosCliente), moneda,
      moneda === "USD" ? tipoCambio : (tipoCambio || ""), subtotal, descuento, total, validez, estado, "", vendedor,
      String(data.observaciones || "").slice(0, 2000), JSON.stringify(items), new Date()];
    if(fila > 0) hoja.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else hoja.appendRow(valores);
    return jsonResponse({ success:true, cotizacion: _filaACotizacion(valores, true) });
  } finally {
    lock.releaseLock();
  }
}

function listarCotizaciones(data){
  const hoja = getHoja(HOJA_COTIZACIONES);
  if(!hoja || hoja.getLastRow() < 2) return jsonResponse({ success:true, cotizaciones: [] });
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS_COTIZACIONES.length).getValues();
  const q = String(data.q || "").trim().toLowerCase();
  const estado = String(data.estado || "").toUpperCase();
  const clienteId = String(data.clienteId || "");
  const limite = Math.min(1000, Number(data.limite) || 300);
  const lista = [];
  for(let i = filas.length - 1; i >= 0 && lista.length < limite; i--){
    const f = filas[i];
    if(!f[0]) continue;
    const c = _filaACotizacion(f, false);
    if(estado && c.estado !== estado) continue;
    if(clienteId && c.clienteId !== clienteId) continue;
    if(q && !(c.cotizacionId + " " + c.cliente + " " + JSON.stringify(c.clienteDatos) + " " + c.observaciones + " " + f[15]).toLowerCase().includes(q)) continue;
    lista.push(c);
  }
  return jsonResponse({ success:true, cotizaciones: lista });
}

function _buscarFilaCotizacion(hoja, id){
  if(!hoja || hoja.getLastRow() < 2) return -1;
  const ids = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
  for(let i = 0; i < ids.length; i++) if(String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function obtenerCotizacion(data){
  const hoja = getHoja(HOJA_COTIZACIONES);
  const fila = _buscarFilaCotizacion(hoja, data.cotizacionId);
  if(fila < 0) return jsonResponse({ success:false, message:"No se encontró la cotización." });
  return jsonResponse({ success:true, cotizacion: _filaACotizacion(hoja.getRange(fila, 1, 1, COLS_COTIZACIONES.length).getValues()[0], true) });
}

function cambiarEstadoCotizacion(data){
  const estado = String(data.estado || "").toUpperCase();
  if(ESTADOS_COTIZACION.indexOf(estado) === -1) return jsonResponse({ success:false, message:"Estado inválido." });
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const hoja = getHoja(HOJA_COTIZACIONES);
    const fila = _buscarFilaCotizacion(hoja, data.cotizacionId);
    if(fila < 0) return jsonResponse({ success:false, message:"No se encontró la cotización." });
    const anterior = String(hoja.getRange(fila, 12).getValue() || "").toUpperCase();
    if(anterior === "CONVERTIDA" && estado !== "CONVERTIDA") return jsonResponse({ success:false, message:"La cotización ya se convirtió en " + hoja.getRange(fila, 13).getValue() + "." });
    hoja.getRange(fila, 12).setValue(estado);
    if(estado === "CONVERTIDA") hoja.getRange(fila, 13).setValue(String(data.convertidaA || ""));
    hoja.getRange(fila, 17).setValue(new Date());
    return jsonResponse({ success:true, cotizacion: _filaACotizacion(hoja.getRange(fila, 1, 1, COLS_COTIZACIONES.length).getValues()[0], true) });
  } finally {
    lock.releaseLock();
  }
}


/* =====================================================================
   TRASPASOS DEPÓSITO → LOCAL
   ---------------------------------------------------------------------
   El depósito arma el traspaso: su stock baja en el momento y queda
   "EN_CAMINO" (hoja TRASPASOS, número TRA-0001).
   El local lo ve (le pregunta al backend del depósito) y confirma lo
   que llegó: el depósito lo marca RECIBIDO (una sola vez) y recién ahí
   el local suma el stock, como un ingreso de mercadería de origen
   DEPÓSITO (sin deuda con proveedor). Si llegó otra cantidad, queda
   anotada la diferencia en los dos lados.
   Los dos backends se hablan con una clave compartida:
     - depósito: Configuración → Sistema → "Traspasos al local" (genera la clave)
     - local:    Configuración → Sistema → "Depósito vinculado" (dirección + clave)
===================================================================== */

const HOJA_TRASPASOS = "TRASPASOS";
const COLS_TRASPASOS = ["TRASPASO_ID", "FECHA", "ESTADO", "ITEMS", "UNIDADES", "USUARIO", "OBSERVACIONES",
  "FECHA_RECEPCION", "RECIBIO", "DIFERENCIAS", "UNIDADES_RECIBIDAS", "FECHA_ANULACION"];
const HOJA_TRASPASOS_RECIBIDOS = "TRASPASOS_RECIBIDOS";

function _hojaTraspasos(){
  let hoja = getHoja(HOJA_TRASPASOS);
  if(!hoja){ hoja = getSS().insertSheet(HOJA_TRASPASOS); hoja.appendRow(COLS_TRASPASOS); }
  return hoja;
}

function _filaATraspaso(f){
  let items = []; try{ items = JSON.parse(f[3] || "[]"); }catch(e){}
  let dif = []; try{ dif = JSON.parse(f[9] || "[]"); }catch(e){}
  const iso = (v) => v instanceof Date ? v.toISOString() : String(v || "");
  return { traspasoId: String(f[0]), fecha: iso(f[1]), estado: String(f[2] || ""), items: items, unidades: Number(f[4]) || 0,
    usuario: String(f[5] || ""), observaciones: String(f[6] || ""), fechaRecepcion: iso(f[7]), recibio: String(f[8] || ""),
    diferencias: dif, unidadesRecibidas: f[10] === "" || f[10] === undefined ? null : Number(f[10]), fechaAnulacion: iso(f[11]) };
}

function _usuarioActual(data){
  return String((data && data.usuarioNombre) || "").trim().slice(0, 80) || (_sesionActual && _sesionActual.u) || "";
}

/* ---------- depósito ---------- */

function crearTraspaso(data){
  if(!esModoDeposito()) return jsonResponse({ success:false, message:"Los traspasos se arman desde el depósito." });
  let items = data.items;
  if(typeof items === "string"){ try{ items = JSON.parse(items); }catch(e){ items = []; } }
  // se agrupan por código (el mismo producto escaneado dos veces)
  const porCodigo = {};
  (Array.isArray(items) ? items : []).forEach(it => {
    const codigo = String(it.codigo || it.CODIGO || "").trim();
    const cant = Number(it.cantidad) || 0;
    if(!codigo || !(cant > 0)) return;
    if(!porCodigo[codigo]) porCodigo[codigo] = { codigo: codigo, producto: String(it.producto || it.PRODUCTO || ""), cantidad: 0, codigoLocal: "" };
    porCodigo[codigo].cantidad += cant;
    const cl = String(it.codigoLocal || "").trim();
    if(cl) porCodigo[codigo].codigoLocal = cl;
  });
  const lista = Object.keys(porCodigo).map(k => porCodigo[k]);
  if(!lista.length) return jsonResponse({ success:false, message:"El traspaso no tiene productos." });

  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    // stock y nombres reales del depósito
    const prod = getSheet();
    const datos = prod.getDataRange().getValues();
    const h = datos[0].map(x => String(x || "").trim());
    const iC = h.indexOf("CODIGO"), iP = h.indexOf("PRODUCTO"), iS = h.indexOf("STOCK"), iCat = h.indexOf("CATEGORIA");
    const mapa = {};
    for(let i = 1; i < datos.length; i++){ const c = String(datos[i][iC] || "").trim(); if(c && !mapa[c]) mapa[c] = datos[i]; }
    const noExisten = lista.filter(it => !mapa[it.codigo]).map(it => it.codigo);
    if(noExisten.length) return jsonResponse({ success:false, message:"No existen en el depósito: " + noExisten.join(", ") });
    const sinStock = lista.filter(it => (Number(mapa[it.codigo][iS]) || 0) < it.cantidad)
      .map(it => `${mapa[it.codigo][iP]} (hay ${Number(mapa[it.codigo][iS]) || 0}, querés mandar ${it.cantidad})`);
    if(sinStock.length && String(data.permitirSinStock) !== "true" && data.permitirSinStock !== true){
      return jsonResponse({ success:false, sinStock: true, message:"No alcanza el stock: " + sinStock.join("; ") });
    }
    // Precio del depósito (en pesos y, si tiene, en dólares): el local lo usa como costo al ingresarlo
    const iPr = h.indexOf("PRECIO"), iPu = h.indexOf("PRECIO_USD");
    const dolarDep = (function(){ try{ return leerDolarDeposito().valor; }catch(e){ return 0; } })();
    lista.forEach(it => {
      it.producto = String(mapa[it.codigo][iP] || it.producto);
      if(iCat !== -1) it.categoria = String(mapa[it.codigo][iCat] || "");
      if(iPr !== -1) it.precio = Number(mapa[it.codigo][iPr]) || 0;
      const usd = iPu !== -1 ? Number(mapa[it.codigo][iPu]) || 0 : 0;
      if(usd > 0){ it.precioUsd = usd; if(dolarDep > 0) it.dolar = dolarDep; }
    });
    // Código del mismo producto en el local: el que se escribió ahora (y se recuerda) o el guardado en el producto
    const iCL = h.indexOf("CODIGO_LOCAL");
    const nuevosCodigos = [];
    lista.forEach(it => {
      const guardado = iCL !== -1 ? String(mapa[it.codigo][iCL] || "").trim() : "";
      if(it.codigoLocal && it.codigoLocal !== guardado) nuevosCodigos.push({ codigo: it.codigo, codigoLocal: it.codigoLocal });
      if(!it.codigoLocal) it.codigoLocal = guardado;
    });
    if(nuevosCodigos.length) _escribirCodigosLocal(prod, nuevosCodigos);

    const hoja = _hojaTraspasos();
    const id = _siguienteNumeroSerie("TRA", hoja);
    const unidades = lista.reduce((a, it) => a + it.cantidad, 0);
    const fila = [id, new Date(), "EN_CAMINO", JSON.stringify(lista), unidades, _usuarioActual(data), String(data.observaciones || "").slice(0, 1000), "", "", "", "", ""];
    hoja.appendRow(fila);
    const deltas = {}; lista.forEach(it => { deltas[it.codigo] = -it.cantidad; });
    const nuevos = aplicarDeltasStock(deltas, false);
    try{ notificarCambioStock(nuevos); }catch(e){}
    return jsonResponse({ success:true, traspaso: _filaATraspaso(fila), stock: nuevos });
  } finally {
    lock.releaseLock();
  }
}

function listarTraspasos(data){
  const hoja = getHoja(HOJA_TRASPASOS);
  if(!hoja || hoja.getLastRow() < 2) return jsonResponse({ success:true, traspasos: [] });
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS_TRASPASOS.length).getValues();
  const estado = String(data.estado || "").toUpperCase();
  const limite = Math.min(500, Number(data.limite) || 200);
  const lista = [];
  for(let i = filas.length - 1; i >= 0 && lista.length < limite; i--){
    if(!filas[i][0]) continue;
    const t = _filaATraspaso(filas[i]);
    if(estado && t.estado !== estado) continue;
    lista.push(t);
  }
  return jsonResponse({ success:true, traspasos: lista });
}

function _buscarFilaTraspaso(hoja, id){
  if(!hoja || hoja.getLastRow() < 2) return -1;
  const ids = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
  for(let i = 0; i < ids.length; i++) if(String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** Solo un traspaso EN_CAMINO: el stock vuelve al depósito */
function anularTraspaso(data){
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const hoja = getHoja(HOJA_TRASPASOS);
    const fila = _buscarFilaTraspaso(hoja, data.traspasoId);
    if(fila < 0) return jsonResponse({ success:false, message:"No se encontró el traspaso." });
    const t = _filaATraspaso(hoja.getRange(fila, 1, 1, COLS_TRASPASOS.length).getValues()[0]);
    if(t.estado !== "EN_CAMINO") return jsonResponse({ success:false, message:"Solo se puede anular un traspaso en camino (este está " + t.estado.toLowerCase().replace(/_/g, " ") + ")." });
    hoja.getRange(fila, 3).setValue("ANULADO");
    hoja.getRange(fila, 12).setValue(new Date());
    const deltas = {}; t.items.forEach(it => { deltas[it.codigo] = (deltas[it.codigo] || 0) + Number(it.cantidad); });
    const nuevos = aplicarDeltasStock(deltas, false);
    try{ notificarCambioStock(nuevos); }catch(e){}
    return jsonResponse({ success:true });
  } finally {
    lock.releaseLock();
  }
}

/** Clave que el local usa para hablar con el depósito (admin). nueva=true la regenera. */
function claveTraspasos(data){
  const props = PropertiesService.getScriptProperties();
  let clave = props.getProperty("CLAVE_TRASPASOS");
  if(!clave || data.nueva === true || String(data.nueva) === "true"){
    clave = "trp_" + Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
    props.setProperty("CLAVE_TRASPASOS", clave);
  }
  return jsonResponse({ success:true, clave: clave, url: (function(){ try{ return ScriptApp.getService().getUrl() || ""; }catch(e){ return ""; } })() });
}

function _claveTraspasosValida(clave){
  const guardada = PropertiesService.getScriptProperties().getProperty("CLAVE_TRASPASOS");
  return !!guardada && typeof clave === "string" && clave.length > 10 && clave === guardada;
}

/** El local pregunta: traspasos en camino y los últimos recibidos */
function traspasosParaLocal(data){
  if(!esModoDeposito()) return jsonResponse({ success:false, message:"Esta dirección no es la del depósito." });
  if(!_claveTraspasosValida(data.clave)) return jsonResponse({ success:false, claveInvalida: true, message:"La clave de traspasos no coincide con la del depósito." });
  const hoja = getHoja(HOJA_TRASPASOS);
  const nombre = (function(){ try{ return String(leerClaveConfiguracion("NOMBRE", "") || ""); }catch(e){ return ""; } })();
  if(!hoja || hoja.getLastRow() < 2) return jsonResponse({ success:true, deposito: nombre, enCamino: [], recibidos: [] });
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS_TRASPASOS.length).getValues();
  const enCamino = [], recibidos = [];
  for(let i = filas.length - 1; i >= 0; i--){
    if(!filas[i][0]) continue;
    const t = _filaATraspaso(filas[i]);
    if(t.estado === "EN_CAMINO") enCamino.push(t);
    else if(t.estado.indexOf("RECIBIDO") === 0 && recibidos.length < 30) recibidos.push(t);
  }
  return jsonResponse({ success:true, deposito: nombre, enCamino: enCamino.reverse(), recibidos: recibidos });
}

/** El local confirma lo que llegó. Se acepta UNA sola vez por traspaso. */
function confirmarRecepcionTraspaso(data){
  if(!esModoDeposito()) return jsonResponse({ success:false, message:"Esta dirección no es la del depósito." });
  if(!_claveTraspasosValida(data.clave)) return jsonResponse({ success:false, claveInvalida: true, message:"La clave de traspasos no coincide con la del depósito." });
  let recibidos = data.recibidos;
  if(typeof recibidos === "string"){ try{ recibidos = JSON.parse(recibidos); }catch(e){ recibidos = []; } }
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El depósito está ocupado, probá de nuevo en unos segundos." });
  try{
    const hoja = getHoja(HOJA_TRASPASOS);
    const fila = _buscarFilaTraspaso(hoja, data.traspasoId);
    if(fila < 0) return jsonResponse({ success:false, message:"El depósito no tiene el traspaso " + data.traspasoId + "." });
    const t = _filaATraspaso(hoja.getRange(fila, 1, 1, COLS_TRASPASOS.length).getValues()[0]);
    if(t.estado !== "EN_CAMINO") return jsonResponse({ success:false, yaProcesado: true, estado: t.estado, message:"El traspaso " + t.traspasoId + " ya está " + t.estado.toLowerCase().replace(/_/g, " ") + "." });
    const mapa = {}, codLocal = {};
    (Array.isArray(recibidos) ? recibidos : []).forEach(r => {
      if(r.cantidad !== undefined && r.cantidad !== null && r.cantidad !== "") mapa[String(r.codigo)] = Math.max(0, Number(r.cantidad) || 0);
      const cl = String(r.codigoLocal || "").trim(); if(cl) codLocal[String(r.codigo)] = cl;
    });
    // el local dijo con qué código tiene cada producto: se recuerda en el depósito para los próximos traspasos
    const aprendidos = [];
    t.items.forEach(it => { if(codLocal[it.codigo] && codLocal[it.codigo] !== String(it.codigoLocal || "")) { aprendidos.push({ codigo: it.codigo, codigoLocal: codLocal[it.codigo] }); it.codigoLocal = codLocal[it.codigo]; } });
    if(aprendidos.length){ try{ _escribirCodigosLocal(getSheet(), aprendidos); }catch(e){ console.error("no se pudieron guardar los códigos del local: " + e); } }
    hoja.getRange(fila, 4).setValue(JSON.stringify(t.items));
    const diferencias = [];
    let unidadesRecibidas = 0;
    t.items.forEach(it => {
      const llego = Object.prototype.hasOwnProperty.call(mapa, it.codigo) ? mapa[it.codigo] : Number(it.cantidad);
      unidadesRecibidas += llego;
      if(llego !== Number(it.cantidad)) diferencias.push({ codigo: it.codigo, producto: it.producto, enviado: Number(it.cantidad), recibido: llego });
    });
    const estado = diferencias.length ? "RECIBIDO_CON_DIFERENCIAS" : "RECIBIDO";
    hoja.getRange(fila, 3).setValue(estado);
    hoja.getRange(fila, 8, 1, 4).setValues([[new Date(), String(data.recibio || "").slice(0, 80), JSON.stringify(diferencias), unidadesRecibidas]]);
    return jsonResponse({ success:true, traspaso: _filaATraspaso(hoja.getRange(fila, 1, 1, COLS_TRASPASOS.length).getValues()[0]) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- local ---------- */

function _llamarDeposito(cuerpo){
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty("DEPOSITO_URL"), clave = props.getProperty("DEPOSITO_CLAVE");
  if(!url || !clave) return { success:false, sinVincular: true, message:"Este local todavía no está vinculado con el depósito (Configuración → Sistema → Depósito vinculado)." };
  return _llamarBackend(url, Object.assign({}, cuerpo, { clave: clave }));
}

function _llamarBackend(url, cuerpo){
  try{
    const r = UrlFetchApp.fetch(url, { method: "post", contentType: "text/plain;charset=utf-8", payload: JSON.stringify(cuerpo), muteHttpExceptions: true, followRedirects: true });
    const texto = r.getContentText();
    try{ return JSON.parse(texto); }
    catch(e){ return { success:false, message:"El depósito respondió algo inesperado (código " + r.getResponseCode() + "). Revisá la dirección." }; }
  }catch(e){
    return { success:false, sinConexion: true, message:"No se pudo conectar con el depósito: " + e.message };
  }
}

/** Admin del local: guarda dirección y clave del depósito (y prueba la conexión) */
function vincularDeposito(data){
  const url = String(data.url || "").trim(), clave = String(data.clave || "").trim();
  const props = PropertiesService.getScriptProperties();
  if(!url && !clave){
    props.deleteProperty("DEPOSITO_URL"); props.deleteProperty("DEPOSITO_CLAVE");
    _invalidarCacheConfigNegocio();
    return jsonResponse({ success:true, desvinculado: true });
  }
  if(!/^https?:\/\/\S+$/.test(url)) return jsonResponse({ success:false, message:"La dirección del depósito no es válida." });
  if(clave.length < 10) return jsonResponse({ success:false, message:"Pegá la clave de traspasos que muestra el depósito." });
  const prueba = _llamarBackend(url, { action: "traspasosParaLocal", clave: clave });
  if(!prueba.success) return jsonResponse({ success:false, message: prueba.message || "El depósito no aceptó la conexión." });
  props.setProperty("DEPOSITO_URL", url);
  props.setProperty("DEPOSITO_CLAVE", clave);
  _invalidarCacheConfigNegocio();
  return jsonResponse({ success:true, deposito: prueba.deposito || "", enCamino: (prueba.enCamino || []).length });
}

/** Lo que el local ve en "Traspasos": en camino (del depósito) y lo ya recibido acá */
function traspasosDelDeposito(data){
  const r = _llamarDeposito({ action: "traspasosParaLocal" });
  if(!r.success) return jsonResponse(r);
  // marca los productos que no existen en el local
  const prod = getSheet();
  const codigos = {};
  if(prod && prod.getLastRow() > 1){
    const iC = obtenerOcrearColumna(prod, "CODIGO");
    prod.getRange(2, iC, prod.getLastRow() - 1, 1).getValues().forEach(f => { codigos[normalizarCodigo(f[0])] = true; });
  }
  (r.enCamino || []).forEach(t => t.items.forEach(it => {
    // primero el código del local que tiene anotado el depósito; si no, el mismo código
    const candidato = String(it.codigoLocal || "").trim();
    it.codigoLocalSugerido = candidato && codigos[normalizarCodigo(candidato)] ? candidato : (codigos[normalizarCodigo(it.codigo)] ? it.codigo : candidato);
    it.existeEnLocal = !!codigos[normalizarCodigo(it.codigoLocalSugerido)];
  }));
  return jsonResponse({ success:true, deposito: r.deposito, enCamino: r.enCamino || [], recibidos: r.recibidos || [] });
}

/**
 * El local confirma la recepción: primero el depósito (que no deja
 * recibir dos veces el mismo traspaso), después suma el stock acá y
 * lo anota como ingreso de mercadería con origen DEPÓSITO.
 */
function recibirTraspaso(data){
  let recibidos = data.recibidos;
  if(typeof recibidos === "string"){ try{ recibidos = JSON.parse(recibidos); }catch(e){ recibidos = []; } }
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(25000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    // ¿ya lo sumamos acá? (por si se cortó la conexión después de confirmar)
    let hojaRec = getHoja(HOJA_TRASPASOS_RECIBIDOS);
    if(hojaRec && _buscarFilaTraspaso(hojaRec, data.traspasoId) > 0){
      // Con revisión de precios: si la confirmación llegó pero el ingreso no (se cortó), se completa el ingreso (es idempotente)
      if(data.ingreso){ lock.releaseLock(); return _ingresoDeTraspaso(data); }
      return jsonResponse({ success:false, yaProcesado: true, message:"Este traspaso ya se recibió en el local." });
    }
    // revisar precios = ingreso de mercadería: el vendedor no puede
    if(data.ingreso && String(data.rol || "").toLowerCase() === "vendedor") return jsonResponse({ success:false, message:"Tu usuario (Vendedor) no puede registrar ingresos: recibí el traspaso sin revisar precios o pedíselo a un administrador." });

    const usuario = _usuarioActual(data);
    const r = _llamarDeposito({ action: "confirmarRecepcionTraspaso", traspasoId: data.traspasoId, recibidos: recibidos, recibio: usuario });
    let t = r.success ? r.traspaso : null;
    // El depósito ya lo tenía como recibido pero acá no se sumó (se cortó justo después): se completa con lo que el depósito anotó
    if(!t && r.yaProcesado && String(r.estado || "").indexOf("RECIBIDO") === 0){
      const lista = _llamarDeposito({ action: "traspasosParaLocal" });
      t = ((lista && lista.recibidos) || []).filter(x => x.traspasoId === String(data.traspasoId))[0] || null;
    }
    if(!t) return jsonResponse(r);

    // Revisando precios: el stock, los precios y la boleta los hace el ingreso de mercadería (como cualquier boleta)
    if(data.ingreso){
      if(!hojaRec){
        hojaRec = getSS().insertSheet(HOJA_TRASPASOS_RECIBIDOS);
        hojaRec.appendRow(["TRASPASO_ID", "FECHA_ENVIO", "FECHA_RECEPCION", "ESTADO", "ITEMS", "DIFERENCIAS", "UNIDADES_RECIBIDAS", "RECIBIO", "NO_EXISTEN_EN_LOCAL"]);
      }
      hojaRec.appendRow([t.traspasoId, t.fecha, new Date(), t.estado, JSON.stringify(t.items), JSON.stringify(t.diferencias || []), t.unidadesRecibidas, usuario, "(ingreso de mercadería)"]);
      lock.releaseLock();
      const salida = JSON.parse(_ingresoDeTraspaso(data).getContent());
      salida.traspaso = t;
      if(!salida.success) salida.message = "El depósito ya lo tiene como recibido, pero el ingreso no se guardó: " + (salida.message || "") + ". Tocá Guardar de nuevo.";
      return jsonResponse(salida);
    }

    // stock: lo que realmente llegó, solo de productos que existen en el local
    const prod = getSheet();
    const datos = prod.getDataRange().getValues();
    const h = datos[0].map(x => String(x || "").trim());
    const iC = h.indexOf("CODIGO"), iP = h.indexOf("PRODUCTO"), iCat = h.indexOf("CATEGORIA");
    const enLocal = {}, codigoReal = {};
    for(let i = 1; i < datos.length; i++){ const c = String(datos[i][iC] || "").trim(); const n = normalizarCodigo(c); if(c && !enLocal[n]){ enLocal[n] = datos[i]; codigoReal[n] = c; } }
    // código del depósito → código del local (lo que confirmó el local, lo anotado en el depósito, o el mismo)
    const elegido = {};
    (Array.isArray(recibidos) ? recibidos : []).forEach(x => { if(String(x.codigoLocal || "").trim()) elegido[String(x.codigo)] = String(x.codigoLocal).trim(); });
    const llegado = {};
    t.items.forEach(it => { llegado[it.codigo] = Number(it.cantidad); });
    (t.diferencias || []).forEach(d => { llegado[d.codigo] = Number(d.recibido); });
    const deltas = {}, noExisten = [], origen = {};
    t.items.forEach(it => {
      const c = it.codigo;
      if(!(llegado[c] > 0)) return;
      const opciones = [elegido[c], it.codigoLocal, c].filter(Boolean).map(normalizarCodigo);
      const n = opciones.filter(o => enLocal[o])[0];
      if(!n){ noExisten.push(c); return; }
      const real = codigoReal[n];
      deltas[real] = (deltas[real] || 0) + llegado[c];
      origen[real] = (origen[real] ? origen[real] + ", " : "") + c;
    });
    const nuevos = aplicarDeltasStock(deltas, false);
    try{ notificarCambioStock(nuevos); }catch(e){}

    // ingreso de mercadería (origen DEPÓSITO, sin precio: no genera deuda con proveedores)
    try{
      const ing = getOrCrearHojaIngresosProductos();
      const enc = ing.getRange(1, 1, 1, ing.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
      const ahora = new Date();
      Object.keys(deltas).forEach(c => {
        const fila = enc.map(() => "");
        const pone = (k, v) => { const i = enc.indexOf(k); if(i !== -1) fila[i] = v; };
        pone("INGRESO_ID", "ING-" + t.traspasoId + "-" + c); pone("BOLETA_ID", t.traspasoId); pone("FECHA", ahora);
        const filaLocal = enLocal[normalizarCodigo(c)];
        pone("CODIGO", c); pone("PRODUCTO", String(filaLocal[iP] || "")); pone("CATEGORIA", iCat !== -1 ? String(filaLocal[iCat] || "") : "");
        pone("CANTIDAD", deltas[c]); pone("MONEDA", "ARS"); pone("PROVEEDOR", "DEPÓSITO");
        pone("OBSERVACIONES", "Traspaso " + t.traspasoId + " desde el depósito" + (origen[c] && origen[c] !== c ? " (código en el depósito: " + origen[c] + ")" : "")); pone("ES_PRODUCTO_NUEVO", "NO"); pone("USUARIO", usuario);
        ing.appendRow(fila);
      });
    }catch(e){ console.error("recibirTraspaso: no se pudo anotar el ingreso: " + e); }

    if(!hojaRec){
      hojaRec = getSS().insertSheet(HOJA_TRASPASOS_RECIBIDOS);
      hojaRec.appendRow(["TRASPASO_ID", "FECHA_ENVIO", "FECHA_RECEPCION", "ESTADO", "ITEMS", "DIFERENCIAS", "UNIDADES_RECIBIDAS", "RECIBIO", "NO_EXISTEN_EN_LOCAL"]);
    }
    hojaRec.appendRow([t.traspasoId, t.fecha, new Date(), t.estado, JSON.stringify(t.items), JSON.stringify(t.diferencias || []), t.unidadesRecibidas, usuario, noExisten.join(", ")]);
    return jsonResponse({ success:true, traspaso: t, stock: nuevos, noExisten: noExisten });
  } finally {
    lock.releaseLock();
  }
}


/* ---------- códigos del local en el depósito ----------
   En el depósito cada producto puede tener otro código que en el local.
   La columna CODIGO_LOCAL de PRODUCTOS (del depósito) guarda el código
   del mismo producto en el local. Se carga desde la ficha del producto,
   al armar un traspaso, o la "aprende" sola cuando el local confirma un
   traspaso eligiendo su producto. */

/** Escribe CODIGO_LOCAL de varios productos (como texto: un código "0012" no pierde los ceros) */
function _escribirCodigosLocal(hoja, lista){
  const col = obtenerOcrearColumna(hoja, "CODIGO_LOCAL");
  try{ hoja.getRange(1, col, Math.max(hoja.getMaxRows(), 2), 1).setNumberFormat("@"); }catch(e){}
  const datos = hoja.getDataRange().getValues();
  const iC = datos[0].map(x => String(x || "").trim()).indexOf("CODIGO");
  const filaPorCodigo = {};
  for(let r = 1; r < datos.length; r++) filaPorCodigo[normalizarCodigo(datos[r][iC])] = r + 1;
  const noEncontrados = [];
  lista.forEach(x => {
    const f = filaPorCodigo[normalizarCodigo(x.codigo)];
    if(!f){ noEncontrados.push(String(x.codigo)); return; }
    hoja.getRange(f, col).setValue(String(x.codigoLocal || "").trim());
  });
  invalidarCacheProductos();
  return noEncontrados;
}

function guardarCodigosLocal(data){
  if(!esModoDeposito()) return jsonResponse({ success:false, message:"El código del local se carga en el depósito." });
  let lista = data.codigos;
  if(typeof lista === "string"){ try{ lista = JSON.parse(lista); }catch(e){ lista = null; } }
  if(!Array.isArray(lista) || !lista.length) return jsonResponse({ success:false, message:"No hay códigos para guardar." });
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return jsonResponse({ success:false, message:"El sistema está ocupado, probá de nuevo en unos segundos." });
  try{
    const noEncontrados = _escribirCodigosLocal(getSheet(), lista);
    return jsonResponse({ success: noEncontrados.length < lista.length, noEncontrados: noEncontrados, message: noEncontrados.length === lista.length ? "No se encontró ningún producto con esos códigos." : "" });
  } finally {
    lock.releaseLock();
  }
}


/** La boleta de ingreso (origen DEPÓSITO, sin deuda) de un traspaso recibido revisando precios */
function _ingresoDeTraspaso(data){
  let ing = data.ingreso;
  if(typeof ing === "string"){ try{ ing = JSON.parse(ing); }catch(e){ ing = {}; } }
  const payload = Object.assign({}, ing, {
    rol: data.rol, origen: "DEPOSITO", sinDeuda: true,
    numeroBoleta: String(data.traspasoId), usuario: ing.usuario || _usuarioActual(data)
  });
  const clave = String(ing.idOperacionSync || ("ING-" + data.traspasoId));
  return conIdempotencia(clave, function(){ return registrarIngresoProducto(payload); });
}
