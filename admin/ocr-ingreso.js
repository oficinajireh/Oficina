/* ===================================================================
   OCR de listas de proveedores -> carrito de "Ingreso de Productos"
   Gratis: Tesseract.js corre en el navegador/Electron, sin servidor.
   Requiere admin.js cargado antes (usa ipCarritoBoleta, renderCarritoBoleta,
   productosAdminGlobal, escapeHtml, toast, normalizarBusquedaPOS).
=================================================================== */
(function () {
  const TESS_SRC = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
  let ocrFilas = [];

  /* ---------- carga perezosa del motor OCR ---------- */
  function cargarTesseract() {
    return new Promise((res, rej) => {
      if (window.Tesseract) return res();
      const s = document.createElement("script");
      s.src = TESS_SRC;
      s.onload = res;
      s.onerror = () => rej(new Error("No se pudo cargar el motor OCR (¿hay internet la primera vez?)"));
      document.head.appendChild(s);
    });
  }

  /* ---------- preprocesado: escala + gris + contraste ---------- */
  async function preprocesar(file) {
    const img = await createImageBitmap(file);
    const escala = Math.max(0.5, Math.min(2.5, 2000 / Math.max(img.width, img.height)));
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
    ctx.putImageData(id, 0, 0);
    return c;
  }

  /* ---------- parseo de texto -> filas ---------- */
  function numAR(s) {
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
  const esNum = t => /^\d[\d.,]*$/.test(t);

  function parsearLinea(linea) {
    const t = linea.replace(/[|_]/g, " ").replace(/\$/g, " ").trim().split(/\s+/).filter(Boolean);
    if (t.length < 2 || !/[a-zA-ZáéíóúñÁÉÍÓÚÑ]{3,}/.test(t.join(" "))) return null;

    let codigo = "", qty = null;
    if (/^\d{6,14}$/.test(t[0])) codigo = t.shift();
    else if (/^[A-Za-z]{1,4}-?\d{2,}$/.test(t[0]) && t.length > 2) codigo = t.shift();

    if (t.length > 2 && /^\d{1,4}[xX]?$/.test(t[0])) { qty = parseInt(t[0], 10); t.shift(); }

    const nums = [];
    while (t.length > 1 && esNum(t[t.length - 1]) && nums.length < 3) nums.unshift(numAR(t.pop()));

    let precio = null;
    if (qty !== null) precio = nums.length ? nums[0] : null;
    else if (nums.length >= 3) { qty = nums[0]; precio = nums[1]; }
    else if (nums.length === 2) {
      if (Number.isInteger(nums[0]) && nums[0] <= 999 && nums[1] > nums[0]) { qty = nums[0]; precio = nums[1]; }
      else precio = nums[0];
    } else if (nums.length === 1) precio = nums[0];

    const nombre = t.join(" ").replace(/[-–—:.,;]+$/, "").trim();
    if (nombre.length < 3) return null;
    return { codigo, nombre, cantidad: qty, precio: isNaN(precio) ? null : precio };
  }

  /* ---------- match contra productos existentes ---------- */
  function buscarMatch(fila) {
    const prods = (typeof productosAdminGlobal !== "undefined" && productosAdminGlobal) || [];
    if (fila.codigo) {
      const p = prods.find(x => String(x.CODIGO).trim() === fila.codigo);
      if (p) return p;
    }
    const toks = normalizarBusquedaPOS(fila.nombre).split(/\s+/).filter(w => w.length > 2);
    if (!toks.length) return null;
    let mejor = null, mejorScore = 0;
    for (const p of prods) {
      const pt = normalizarBusquedaPOS(p.PRODUCTO).split(/\s+/);
      const hit = toks.filter(w => pt.some(x => x === w || x.includes(w) || w.includes(x))).length;
      const score = hit / toks.length;
      if (score > mejorScore) { mejorScore = score; mejor = p; }
    }
    return mejorScore >= 0.7 ? mejor : null;
  }

  function textoAFilas(texto) {
    return texto.split(/\r?\n/).map(l => parsearLinea(l)).filter(Boolean).map(f => {
      const m = buscarMatch(f);
      return Object.assign(f, { matchCodigo: m ? String(m.CODIGO) : "", incluir: true });
    });
  }

  /* ---------- UI: tarjeta + modal (se inyectan solas) ---------- */
  function montarUI() {
    const ref = document.getElementById("ipbEstadoBoleta");
    if (!ref || document.getElementById("ocrCard")) return;

    const card = document.createElement("div");
    card.className = "card p-3 mb-3";
    card.id = "ocrCard";
    card.innerHTML = `
      <h6 class="mb-2">📷 Cargar desde foto de la lista del proveedor <span class="badge bg-secondary">OCR gratis</span></h6>
      <div class="d-flex gap-2 flex-wrap align-items-center">
        <input type="file" id="ocrArchivo" accept="image/*" class="form-control" style="max-width:340px" onchange="ocrProcesarArchivo(event)">
        <span id="ocrEstado" class="text-muted" style="font-size:13px;"></span>
      </div>
      <div class="progress mt-2" id="ocrProgWrap" style="height:6px; display:none;"><div class="progress-bar" id="ocrProg" style="width:0%"></div></div>
      <div class="form-text" style="font-size:11.5px;">Sacá la foto de frente, con buena luz y sin sombras. Siempre vas a poder revisar y corregir antes de agregar a la boleta.</div>`;
    ref.closest(".card").after(card);

    const modal = document.createElement("div");
    modal.className = "product-modal-backdrop";
    modal.id = "ocrModalBackdrop";
    modal.innerHTML = `
      <div class="product-modal" style="max-width:980px; width:96vw;">
        <div class="product-modal-header"><h6>📷 Revisá lo que leyó la foto</h6><button onclick="ocrCerrar()">✕</button></div>
        <div class="product-modal-body" style="max-height:72vh; overflow:auto;">
          <div class="row g-2 mb-2">
            <div class="col-md-3"><label class="form-label">Moneda de los precios</label>
              <select class="form-select form-select-sm" id="ocrMoneda" onchange="ocrCambioMoneda()"><option value="ARS">Pesos (ARS)</option><option value="USD">Dólares (USD)</option></select></div>
            <div class="col-md-3" id="ocrTcWrap" style="display:none;"><label class="form-label">Tipo de cambio</label>
              <input type="number" class="form-control form-control-sm" id="ocrTc" placeholder="Ej: 1250"></div>
          </div>
          <div class="text-muted mb-2" style="font-size:12.5px;">Las filas en amarillo tienen datos que faltan. Destildá las que no sean productos (encabezados, totales, etc.).</div>
          <div class="table-responsive"><table class="table table-sm table-striped">
            <thead><tr><th></th><th>Código</th><th>Producto</th><th style="width:90px;">Cant.</th><th style="width:120px;">Costo unit.</th><th>Estado</th></tr></thead>
            <tbody id="ocrTabla"></tbody></table></div>
          <details class="mt-2"><summary style="cursor:pointer; font-size:13px;">Ver / corregir el texto leído</summary>
            <textarea id="ocrTextoCrudo" class="form-control mt-2" rows="8" style="font-family:monospace; font-size:12px;"></textarea>
            <button class="btn btn-outline-secondary btn-sm mt-2" onclick="ocrReprocesarTexto()">↻ Volver a armar la tabla con este texto</button>
          </details>
        </div>
        <div class="product-modal-actions">
          <button class="btn btn-outline-secondary" onclick="ocrCerrar()">Cancelar</button>
          <button class="btn btn-success" id="ocrBtnAgregar" onclick="ocrAgregarABoleta()">➕ Agregar a la boleta</button>
        </div>
      </div>`;
    document.body.appendChild(modal);
  }

  function renderTabla() {
    const tb = document.getElementById("ocrTabla");
    tb.innerHTML = ocrFilas.map((f, i) => {
      const falta = !f.cantidad || !f.precio;
      const existente = (f.matchCodigo || f.codigo) ? true : false;
      return `<tr style="${falta ? "background:#fff8e1;" : ""}">
        <td><input type="checkbox" ${f.incluir ? "checked" : ""} onchange="ocrFilas_set(${i},'incluir',this.checked)"></td>
        <td><input class="form-control form-control-sm mono" value="${escapeHtml(f.matchCodigo || f.codigo || "")}" onchange="ocrFilas_set(${i},'codigo',this.value)"></td>
        <td><input class="form-control form-control-sm" value="${escapeHtml(f.nombre)}" onchange="ocrFilas_set(${i},'nombre',this.value)"></td>
        <td><input type="number" min="1" class="form-control form-control-sm" value="${f.cantidad ?? ""}" onchange="ocrFilas_set(${i},'cantidad',this.value)"></td>
        <td><input type="number" min="0" step="any" class="form-control form-control-sm" value="${f.precio ?? ""}" onchange="ocrFilas_set(${i},'precio',this.value)"></td>
        <td>${f.matchCodigo ? '<span class="badge bg-success">Existente</span>' : '<span class="badge bg-warning text-dark">Nuevo</span>'}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="6" class="text-center text-muted py-3">No se pudo reconocer ningún producto. Probá con una foto más nítida o corregí el texto abajo.</td></tr>`;
  }

  window.ocrFilas_set = function (i, campo, valor) {
    const f = ocrFilas[i];
    if (campo === "incluir") f.incluir = !!valor;
    else if (campo === "cantidad" || campo === "precio") f[campo] = valor === "" ? null : Number(valor);
    else if (campo === "codigo") {
      f.codigo = valor.trim();
      const m = (productosAdminGlobal || []).find(p => String(p.CODIGO).trim() === f.codigo);
      f.matchCodigo = m ? String(m.CODIGO) : "";
    } else f[campo] = valor;
    renderTabla();
  };

  window.ocrCambioMoneda = function () {
    document.getElementById("ocrTcWrap").style.display = document.getElementById("ocrMoneda").value === "USD" ? "" : "none";
  };
  window.ocrCerrar = function () { document.getElementById("ocrModalBackdrop").classList.remove("show"); };

  window.ocrReprocesarTexto = function () {
    ocrFilas = textoAFilas(document.getElementById("ocrTextoCrudo").value);
    renderTabla();
  };

  /* ---------- flujo principal: foto -> OCR -> modal ---------- */
  window.ocrProcesarArchivo = async function (ev) {
    const file = ev.target.files && ev.target.files[0];
    if (!file) return;
    const estado = document.getElementById("ocrEstado");
    const wrap = document.getElementById("ocrProgWrap");
    const barra = document.getElementById("ocrProg");
    wrap.style.display = "";
    barra.style.width = "0%";
    try {
      if (!productosAdminGlobal || !productosAdminGlobal.length) await cargarProductos();
      estado.textContent = "⏳ Cargando motor OCR...";
      await cargarTesseract();
      estado.textContent = "⏳ Preparando imagen...";
      const canvas = await preprocesar(file);

      const worker = await Tesseract.createWorker("spa", 1, {
        logger: m => {
          if (m.status === "recognizing text") {
            barra.style.width = Math.round(m.progress * 100) + "%";
            estado.textContent = "⏳ Leyendo... " + Math.round(m.progress * 100) + "%";
          }
        }
      });
      await worker.setParameters({ tessedit_pageseg_mode: "6", preserve_interword_spaces: "1" });
      const { data } = await worker.recognize(canvas);
      await worker.terminate();

      estado.textContent = "✓ Listo";
      ocrFilas = textoAFilas(data.text);
      document.getElementById("ocrTextoCrudo").value = data.text;
      renderTabla();
      document.getElementById("ocrModalBackdrop").classList.add("show");
    } catch (e) {
      console.error("OCR:", e);
      estado.textContent = "";
      toast(e.message || "No se pudo leer la imagen", "error");
    } finally {
      wrap.style.display = "none";
      ev.target.value = "";
    }
  };

  /* ---------- volcar al carrito de la boleta ---------- */
  window.ocrAgregarABoleta = function () {
    const moneda = document.getElementById("ocrMoneda").value === "USD" ? "USD" : "ARS";
    const tc = Number(document.getElementById("ocrTc").value || 0);
    if (moneda === "USD" && tc <= 0) { toast("Ingresá el tipo de cambio para precios en dólares", "error"); return; }

    const elegidas = ocrFilas.filter(f => f.incluir);
    if (!elegidas.length) { toast("No hay filas tildadas", "error"); return; }

    for (const f of elegidas) {
      if (!f.cantidad || f.cantidad <= 0 || !f.precio || f.precio <= 0) {
        toast(`Falta cantidad o costo en "${f.nombre}"`, "error"); return;
      }
    }

    const prods = productosAdminGlobal || [];
    elegidas.forEach(f => {
      const existente = prods.find(p => String(p.CODIGO).trim() === (f.codigo || f.matchCodigo)) || null;
      ipCarritoBoleta.push({
        codigo: existente ? String(existente.CODIGO) : (f.codigo || ""),
        producto: existente ? existente.PRODUCTO : f.nombre,
        categoria: existente ? (existente.CATEGORIA || "") : "",
        moneda,
        precio: f.precio,
        tipoCambio: moneda === "USD" ? tc : null,
        precioARS: moneda === "USD" ? f.precio * tc : f.precio,
        precioVenta: null,
        cantidad: f.cantidad,
        esNuevo: !existente
      });
    });

    renderCarritoBoleta();
    ocrCerrar();
    toast(`${elegidas.length} producto(s) agregados a la boleta`, "success");
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", montarUI);
  else montarUI();
})();
