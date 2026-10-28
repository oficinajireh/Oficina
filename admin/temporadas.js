/* =====================================================================
   TEMAS DE TEMPORADA — compartido por el catálogo web (web/js) y el
   panel admin (sincronizar.js lo copia a web/admin y desktop/panel:
   se edita SOLO este archivo).

   Cada tema cambia en el catálogo: colores, adorno colgando del
   encabezado, fondo con motivos suaves, animación liviana (se puede
   apagar en el panel), ícono en las tarjetas de productos y borde de
   los destacados. En el panel admin cambia los colores.

   La elección es manual (Configuración → Catálogo web → Tema de
   temporada), con fechas desde/hasta opcionales.
===================================================================== */
(function (global) {
  "use strict";

  const TEMPORADAS = {
    navidad: {
      nombre: "Navidad", emoji: "🎄",
      colores: { a: "#0b3d2c", b: "#14532d", acento: "#c8102e", acentoOsc: "#9e0b22", papel: "#fbf6ee", linea: "#eadcc5", extra: "#d4a017" },
      adorno: { tipo: "luces", colores: ["#ef4444", "#facc15", "#22c55e", "#3b82f6", "#f97316"] },
      fondo: ["❄", "✦", "❅"],
      animacion: { modo: "nieve", cantidad: 42 },
      icono: "🎄",
      cartel: "🎄 ¡Felices fiestas! Hacé tu pedido con tiempo para recibirlo antes de Navidad"
    },
    anonuevo: {
      nombre: "Año Nuevo", emoji: "🥂",
      colores: { a: "#0b1020", b: "#1c2547", acento: "#a16207", acentoOsc: "#854d0e", papel: "#fbf9f3", linea: "#ebe3cd", extra: "#e5c07b" },
      adorno: { tipo: "estrellas", colores: ["#facc15", "#e5e7eb", "#fde68a"] },
      fondo: ["✨", "✦", "⭐"],
      animacion: { modo: "confeti", colores: ["#facc15", "#fde68a", "#e5e7eb", "#d4a017"], extras: ["✨"], cantidad: 40 },
      icono: "🥂",
      cartel: "🥂 ¡Feliz Año Nuevo! Gracias por elegirnos"
    },
    sanvalentin: {
      nombre: "San Valentín", emoji: "💘",
      colores: { a: "#5b0f2a", b: "#8f1d43", acento: "#e11d48", acentoOsc: "#be123c", papel: "#fff6f7", linea: "#f8d7dd", extra: "#f472b6" },
      adorno: { tipo: "corazones", colores: ["#e11d48", "#f472b6", "#fb7185"] },
      fondo: ["♥", "♡"],
      animacion: { modo: "sube", items: ["❤️", "💕", "💖"], cantidad: 18 },
      icono: "💘",
      cartel: "💘 San Valentín: regalos para sorprender"
    },
    clases: {
      nombre: "Regreso a clases", emoji: "✏️",
      colores: { a: "#1e3a5f", b: "#2b5a8c", acento: "#d62839", acentoOsc: "#b21f2e", papel: "#fbfaf4", linea: "#e6e1cc", extra: "#f2b705" },
      adorno: { tipo: "lapices", colores: ["#f2b705", "#ef4444", "#3b82f6", "#22c55e"] },
      fondo: ["✏", "★", "📐"],
      animacion: { modo: "cae", items: ["✏️", "⭐", "📏", "📎"], cantidad: 16 },
      icono: "✏️",
      cartel: "✏️ Regreso a clases: todo para la mochila"
    },
    pascuas: {
      nombre: "Pascuas", emoji: "🐣",
      colores: { a: "#3b4a6b", b: "#5b6fa0", acento: "#7c3aed", acentoOsc: "#6d28d9", papel: "#fbfaf2", linea: "#ebe6cc", extra: "#22c55e" },
      adorno: { tipo: "huevos", colores: ["#f9a8d4", "#a7f3d0", "#fde68a", "#c4b5fd", "#bae6fd"] },
      fondo: ["🥚", "✿"],
      animacion: { modo: "cae", items: ["🥚", "🐣", "🌷"], cantidad: 14 },
      icono: "🐣",
      cartel: "🐣 ¡Felices Pascuas!"
    },
    diamadre: {
      nombre: "Día de la Madre", emoji: "💐",
      colores: { a: "#6b2148", b: "#9d3c72", acento: "#d63c79", acentoOsc: "#b52d64", papel: "#fff5f8", linea: "#f6d6e3", extra: "#a855f7" },
      adorno: { tipo: "flores", colores: ["#f472b6", "#c084fc", "#fb7185", "#fbcfe8"] },
      fondo: ["✿", "❀", "♥"],
      animacion: { modo: "cae", items: ["🌸", "💗", "🌷"], cantidad: 20 },
      icono: "💐",
      cartel: "💐 Día de la Madre: regalos para mamá"
    },
    diapadre: {
      nombre: "Día del Padre", emoji: "👔",
      colores: { a: "#0f2742", b: "#1e4a7a", acento: "#1f74b8", acentoOsc: "#195f97", papel: "#f4f8fc", linea: "#d9e6f2", extra: "#c08a3e" },
      adorno: { tipo: "banderines", colores: ["#1e4a7a", "#c08a3e", "#93c5fd", "#334155"] },
      fondo: ["✦", "★"],
      animacion: { modo: "cae", items: ["✦"], color: "#c08a3e", cantidad: 12 },
      icono: "👔",
      cartel: "👔 Día del Padre: el regalo ideal para papá"
    },
    dianino: {
      nombre: "Día del Niño", emoji: "🧸",
      colores: { a: "#1e3a8a", b: "#2563eb", acento: "#7c3aed", acentoOsc: "#6d28d9", papel: "#fffbeb", linea: "#fde7b0", extra: "#f59e0b" },
      adorno: { tipo: "banderines", colores: ["#ef4444", "#f59e0b", "#22c55e", "#3b82f6", "#a855f7", "#ec4899"] },
      fondo: ["★", "●", "▲"],
      animacion: { modo: "confeti", colores: ["#ef4444", "#f59e0b", "#22c55e", "#3b82f6", "#a855f7", "#ec4899"], extras: ["🎈"], cantidad: 40 },
      icono: "🧸",
      cartel: "🧸 Día del Niño: ¡juguetes para todos!"
    },
    halloween: {
      nombre: "Halloween", emoji: "🎃",
      colores: { a: "#1a0f24", b: "#3b1d4f", acento: "#c2410c", acentoOsc: "#9a3412", papel: "#fbf4ea", linea: "#ecd9bf", extra: "#7c3aed" },
      adorno: { tipo: "telaranas", colores: ["#f97316", "#7c3aed", "#111827"] },
      fondo: ["🦇", "✦"],
      animacion: { modo: "cae", items: ["🍂", "🍁", "🦇"], cantidad: 16 },
      icono: "🎃",
      cartel: "🎃 Halloween: disfraces y golosinas"
    },
    blackfriday: {
      nombre: "Black Friday / Hot Sale", emoji: "🏷️",
      colores: { a: "#0a0a0a", b: "#1f1f1f", acento: "#e11d48", acentoOsc: "#be123c", papel: "#f6f6f4", linea: "#e2e2dc", extra: "#facc15" },
      adorno: { tipo: "cinta", colores: ["#facc15", "#111111"] },
      fondo: ["%", "⚡"],
      animacion: { modo: "cae", items: ["🏷️", "⚡", "🔥"], cantidad: 14 },
      icono: "🔥",
      cartel: "🔥 Ofertas por tiempo limitado"
    }
  };

  // Orden para mostrar en el panel (más o menos el del año)
  const ORDEN = ["navidad", "anonuevo", "sanvalentin", "clases", "pascuas", "diamadre", "diapadre", "dianino", "halloween", "blackfriday"];

  /* ---------- Fechas ---------- */
  function hoyAR() {
    try {
      return new Intl.DateTimeFormat("en-CA", { timeZone: "America/Argentina/Buenos_Aires", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
    } catch (e) {
      const d = new Date();
      return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
    }
  }

  /** ¿Se tiene que ver hoy? (elegido, y dentro de las fechas si las hay) */
  function estaActiva(cfg, hoy) {
    if (!cfg || !cfg.id || !TEMPORADAS[cfg.id]) return false;
    hoy = hoy || hoyAR();
    if (cfg.desde && hoy < cfg.desde) return false;
    if (cfg.hasta && hoy > cfg.hasta) return false;
    return true;
  }

  /* ---------- Dibujos (SVG livianos, sin imágenes) ---------- */
  const svgUri = (svg) => "url(\"data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg) + "\")";

  /** Tira que cuelga del encabezado. Devuelve { fondo, alto, ancho } (se repite a lo ancho). */
  function adornoTira(adorno) {
    const c = adorno.colores;
    switch (adorno.tipo) {
      case "luces": {
        const W = 120, H = 26, n = 5;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><path d="M0 4 Q30 14 60 4 T120 4" fill="none" stroke="#1f2937" stroke-width="1.6"/>`;
        for (let i = 0; i < n; i++) {
          const x = 12 + i * 24, y = 4 + Math.sin((x / W) * Math.PI * 2) * -5 + 5;
          s += `<rect x="${x - 2}" y="${y}" width="4" height="4" fill="#374151"/><ellipse cx="${x}" cy="${y + 10}" rx="4.2" ry="6.5" fill="${c[i % c.length]}"/><ellipse cx="${x - 1.3}" cy="${y + 7.5}" rx="1.2" ry="2" fill="#fff" opacity=".55"/>`;
        }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "banderines": {
        const W = 168, H = 24, n = 6;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><path d="M0 2 Q42 7 84 2 T168 2" fill="none" stroke="#6b7280" stroke-width="1.2"/>`;
        for (let i = 0; i < n; i++) { const x = i * 28 + 3; s += `<path d="M${x} 3 L${x + 22} 3 L${x + 11} 21 Z" fill="${c[i % c.length]}"/>`; }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "flores": {
        const W = 120, H = 22, n = 5;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><path d="M0 4 Q60 10 120 4" fill="none" stroke="#86b39a" stroke-width="1.4"/>`;
        for (let i = 0; i < n; i++) {
          const x = 12 + i * 24, y = 12, col = c[i % c.length];
          for (let p = 0; p < 5; p++) { const a = (p / 5) * Math.PI * 2; s += `<circle cx="${(x + Math.cos(a) * 4).toFixed(1)}" cy="${(y + Math.sin(a) * 4).toFixed(1)}" r="3.3" fill="${col}"/>`; }
          s += `<circle cx="${x}" cy="${y}" r="2.4" fill="#fde68a"/>`;
        }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "corazones": {
        const W = 110, H = 22, n = 5;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><path d="M0 3 Q55 9 110 3" fill="none" stroke="#9f1239" stroke-width="1" opacity=".6"/>`;
        for (let i = 0; i < n; i++) { const x = 11 + i * 22, y = 8, col = c[i % c.length]; s += `<path transform="translate(${x - 7} ${y}) scale(.58)" d="M12 21s-8-5.3-10.6-10C-.3 6.9 2 2 6.5 2 9 2 11 3.6 12 5.5 13 3.6 15 2 17.5 2 22 2 24.3 6.9 22.6 11 20 15.7 12 21 12 21z" fill="${col}"/>`; }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "huevos": {
        const W = 125, H = 24, n = 5;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><path d="M0 3 Q62 8 125 3" fill="none" stroke="#a3a3a3" stroke-width="1"/>`;
        for (let i = 0; i < n; i++) { const x = 12 + i * 25, col = c[i % c.length]; s += `<ellipse cx="${x}" cy="13" rx="6.5" ry="8.5" fill="${col}"/><path d="M${x - 6} 13 q3 -3 6 0 t6 0" fill="none" stroke="#fff" stroke-width="1.4" opacity=".8"/>`; }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "estrellas": {
        const W = 96, H = 22, n = 4;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">`;
        for (let i = 0; i < n; i++) {
          const x = 12 + i * 24, y = 4 + (i % 2) * 6, col = c[i % c.length];
          s += `<line x1="${x}" y1="0" x2="${x}" y2="${y + 3}" stroke="#9ca3af" stroke-width=".8"/><path transform="translate(${x - 6} ${y + 2}) scale(.5)" d="M12 0l3.1 7.6 8.2.6-6.3 5.3 2 8L12 17.2 5 21.5l2-8L.7 8.2l8.2-.6z" fill="${col}"/>`;
        }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "lapices": {
        const W = 180, H = 20, n = 4;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">`;
        for (let i = 0; i < n; i++) {
          const x = 4 + i * 45, y = 5 + (i % 2) * 3, col = c[i % c.length];
          s += `<rect x="${x}" y="${y}" width="6" height="8" rx="1.5" fill="#f9a8d4"/><rect x="${x + 6}" y="${y}" width="3" height="8" fill="#d1d5db"/>` +
               `<rect x="${x + 9}" y="${y}" width="22" height="8" fill="${col}"/><rect x="${x + 9}" y="${y + 2.7}" width="22" height="2.6" fill="#fff" opacity=".25"/>` +
               `<path d="M${x + 31} ${y} L${x + 38} ${y + 4} L${x + 31} ${y + 8} Z" fill="#f5deb3"/><path d="M${x + 36} ${y + 2.9} L${x + 38} ${y + 4} L${x + 36} ${y + 5.1} Z" fill="#374151"/>`;
        }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
      case "renglones": {
        const W = 40, H = 16;
        const s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="${c[2]}"/><line x1="0" y1="5" x2="${W}" y2="5" stroke="${c[0]}" stroke-width="1"/><line x1="0" y1="11" x2="${W}" y2="11" stroke="${c[0]}" stroke-width="1"/><line x1="0" y1="15.5" x2="${W}" y2="15.5" stroke="${c[1]}" stroke-width="1"/></svg>`;
        return { fondo: svgUri(s), alto: H, ancho: W };
      }
      case "cinta": {
        const W = 28, H = 14;
        const s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="${c[0]}"/><path d="M0 14 L7 0 L14 0 L7 14Z M14 14 L21 0 L28 0 L21 14Z" fill="${c[1]}"/></svg>`;
        return { fondo: svgUri(s), alto: H, ancho: W };
      }
      case "telaranas": {
        const W = 150, H = 22, n = 6;
        let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><path d="M0 2 Q37 7 75 2 T150 2" fill="none" stroke="#4b5563" stroke-width="1"/>`;
        for (let i = 0; i < n; i++) { const x = i * 25 + 3; s += `<path d="M${x} 3 L${x + 19} 3 L${x + 9.5} 19 Z" fill="${c[i % 2]}"/>`; }
        return { fondo: svgUri(s + "</svg>"), alto: H, ancho: W };
      }
    }
    return null;
  }

  /** Telaraña para las esquinas (solo Halloween) */
  function telarana(color) {
    let s = `<svg xmlns="http://www.w3.org/2000/svg" width="90" height="90" viewBox="0 0 90 90" fill="none" stroke="${color}" stroke-width="1" opacity=".55">`;
    for (let i = 0; i <= 5; i++) { const a = (i / 5) * (Math.PI / 2); s += `<line x1="0" y1="0" x2="${(Math.cos(a) * 90).toFixed(1)}" y2="${(Math.sin(a) * 90).toFixed(1)}"/>`; }
    for (let r = 18; r <= 90; r += 18) {
      let d = "";
      for (let i = 0; i <= 5; i++) { const a = (i / 5) * (Math.PI / 2); const x = (Math.cos(a) * r).toFixed(1), y = (Math.sin(a) * r).toFixed(1); d += (i ? " Q" + (Math.cos(a - 0.157) * r * 0.86).toFixed(1) + " " + (Math.sin(a - 0.157) * r * 0.86).toFixed(1) + " " : "M") + x + " " + y; }
      s += `<path d="${d}"/>`;
    }
    s += `<line x1="58" y1="0" x2="58" y2="44" stroke-width=".8"/><circle cx="58" cy="48" r="4.5" fill="${color}"/><circle cx="58" cy="42" r="3" fill="${color}"/></svg>`;
    return svgUri(s);
  }

  /** Fondo de página con motivos muy suaves */
  function fondoMotivos(simbolos, color) {
    const s = `<svg xmlns="http://www.w3.org/2000/svg" width="180" height="180"><g fill="${color}" font-family="Arial, sans-serif" opacity=".08" font-size="20">` +
      simbolos.map((t, i) => {
        const pos = [[20, 30], [110, 60], [60, 120], [150, 150], [140, 15], [15, 160]][i % 6];
        return `<text x="${pos[0]}" y="${pos[1]}" transform="rotate(${(i * 23) % 40 - 20} ${pos[0]} ${pos[1]})">${t.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</text>`;
      }).join("") +
      simbolos.map((t, i) => {
        const pos = [[90, 100], [40, 75], [125, 115], [80, 20]][i % 4];
        return `<text x="${pos[0]}" y="${pos[1]}" font-size="14">${t.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</text>`;
      }).join("") + `</g></svg>`;
    return svgUri(s);
  }

  /* ---------- Colores ---------- */
  /** Variables CSS del catálogo (web/css/style.css) */
  function varsCatalogo(t) {
    const c = t.colores;
    return {
      "--navy-950": c.a, "--navy-900": c.a, "--navy-800": c.b,
      "--indigo-600": c.acento, "--indigo-700": c.acentoOsc,
      "--paper": c.papel, "--paper-line": c.linea,
      "--coral-600": c.extra, "--gold-500": c.extra,
      "--grad-a": c.a, "--grad-b": c.b
    };
  }

  /** Variables CSS del panel admin (web/admin/admin.css) */
  function varsPanel(t) {
    const c = t.colores;
    return {
      "--navy-950": c.a,
      "--navy-900": `color-mix(in srgb, ${c.a} 70%, ${c.b})`,
      "--navy-800": c.b,
      "--navy-700": `color-mix(in srgb, ${c.b} 82%, #ffffff)`,
      "--navy-600": `color-mix(in srgb, ${c.b} 65%, ${c.acento})`,
      "--blue-500": c.acento,
      "--blue-50": `color-mix(in srgb, ${c.acento} 10%, #ffffff)`
    };
  }

  /* ---------- Animación (canvas, liviana) ---------- */
  let anim = null;

  function detenerAnimacion() {
    if (!anim) return;
    cancelAnimationFrame(anim.raf);
    window.removeEventListener("resize", anim.alRedimensionar);
    document.removeEventListener("visibilitychange", anim.alVisibilidad);
    anim.canvas.remove();
    anim = null;
  }

  function iniciarAnimacion(def) {
    detenerAnimacion();
    if (!def || typeof document === "undefined") return;
    try { if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return; } catch (e) {}
    const canvas = document.createElement("canvas");
    canvas.className = "temp-animacion";
    canvas.setAttribute("aria-hidden", "true");
    document.body.appendChild(canvas);
    const ctx = canvas.getContext("2d");
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    let W = 0, H = 0;
    const medir = () => { W = window.innerWidth; H = window.innerHeight; canvas.width = W * dpr; canvas.height = H * dpr; ctx.setTransform(dpr, 0, 0, dpr, 0, 0); };
    medir();

    // Emojis pre-dibujados una vez (dibujar texto en cada cuadro es caro)
    const cacheEmoji = {};
    const sprite = (txt, tam, color) => {
      const k = txt + "|" + tam + "|" + (color || "");
      if (cacheEmoji[k]) return cacheEmoji[k];
      const c = document.createElement("canvas"); const px = Math.ceil(tam * 1.4 * dpr);
      c.width = c.height = px; const x = c.getContext("2d");
      x.font = `${tam * dpr}px "Apple Color Emoji","Segoe UI Emoji","Noto Color Emoji",sans-serif`;
      x.textAlign = "center"; x.textBaseline = "middle"; if (color) x.fillStyle = color;
      x.fillText(txt, px / 2, px / 2 + tam * 0.08 * dpr);
      return (cacheEmoji[k] = c);
    };

    const movil = W < 700;
    const cantidad = Math.round((def.cantidad || 20) * (movil ? 0.5 : 1));
    const rnd = (a, b) => a + Math.random() * (b - a);
    const nueva = (alInicio) => {
      const p = { x: rnd(0, W), y: alInicio ? rnd(-H, H) : -30, vy: rnd(0.35, 1.1), vx: rnd(-0.3, 0.3), fase: rnd(0, 6.28), giro: rnd(-0.03, 0.03), ang: rnd(0, 6.28) };
      if (def.modo === "nieve") { p.r = rnd(1.2, 3.6); p.vy = rnd(0.3, 0.9); p.alfa = rnd(0.45, 0.9); }
      else if (def.modo === "confeti") {
        const usarExtra = def.extras && Math.random() < 0.12;
        p.txt = usarExtra ? def.extras[Math.floor(Math.random() * def.extras.length)] : null;
        p.color = def.colores[Math.floor(Math.random() * def.colores.length)];
        p.w = rnd(5, 9); p.h = rnd(3, 6); p.vy = rnd(0.6, 1.5); p.tam = rnd(16, 24);
      } else {
        p.txt = def.items[Math.floor(Math.random() * def.items.length)];
        p.tam = rnd(14, 24) * (movil ? 0.85 : 1); p.alfa = rnd(0.55, 0.9);
        if (def.modo === "sube") { p.y = alInicio ? rnd(0, H * 2) : H + 30; p.vy = -rnd(0.35, 0.8); }
      }
      return p;
    };
    const parts = Array.from({ length: cantidad }, () => nueva(true));
    let visible = !document.hidden;
    let t0 = performance.now();

    const cuadro = (ahora) => {
      anim.raf = requestAnimationFrame(cuadro);
      if (!visible) return;
      const dt = Math.min(3, (ahora - t0) / 16.7); t0 = ahora;
      ctx.clearRect(0, 0, W, H);
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i];
        p.fase += 0.02 * dt; p.ang += p.giro * dt;
        p.x += (p.vx + Math.sin(p.fase) * 0.35) * dt; p.y += p.vy * dt;
        if ((p.vy > 0 && p.y > H + 30) || (p.vy < 0 && p.y < -40) || p.x < -40 || p.x > W + 40) { parts[i] = nueva(false); continue; }
        if (def.modo === "nieve") {
          ctx.globalAlpha = p.alfa; ctx.fillStyle = "#ffffff"; ctx.beginPath(); ctx.arc(p.x, p.y, p.r, 0, 6.283); ctx.fill();
          ctx.globalAlpha = p.alfa * 0.35; ctx.strokeStyle = "#94a3b8"; ctx.lineWidth = 0.6; ctx.stroke();
        } else if (def.modo === "confeti" && !p.txt) {
          ctx.globalAlpha = 0.85; ctx.save(); ctx.translate(p.x, p.y); ctx.rotate(p.ang); ctx.scale(1, Math.abs(Math.cos(p.fase * 2)) + 0.15);
          ctx.fillStyle = p.color; ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h); ctx.restore();
        } else {
          const img = sprite(p.txt, Math.round(p.tam), def.color);
          ctx.globalAlpha = p.alfa || 0.85;
          ctx.save(); ctx.translate(p.x, p.y); ctx.rotate(Math.sin(p.fase) * 0.35);
          ctx.drawImage(img, -img.width / dpr / 2, -img.height / dpr / 2, img.width / dpr, img.height / dpr); ctx.restore();
        }
      }
      ctx.globalAlpha = 1;
    };
    anim = { canvas, raf: 0 };
    anim.alRedimensionar = () => medir();
    anim.alVisibilidad = () => { visible = !document.hidden; t0 = performance.now(); };
    window.addEventListener("resize", anim.alRedimensionar, { passive: true });
    document.addEventListener("visibilitychange", anim.alVisibilidad);
    anim.raf = requestAnimationFrame(cuadro);
  }

  /* ---------- Aplicar en el catálogo ---------- */
  const VARS_CATALOGO = ["--navy-950", "--navy-900", "--navy-800", "--indigo-600", "--indigo-700", "--paper", "--paper-line", "--coral-600", "--gold-500", "--grad-a", "--grad-b"];

  /**
   * cfg: { id, desde, hasta, animacion, colores, cartel } (de Configuración)
   * opciones.forzar: id para la vista previa (?vistaTemporada=…), ignora fechas
   * Devuelve el tema aplicado o null.
   */
  function aplicarCatalogo(cfg, opciones) {
    opciones = opciones || {};
    const body = document.body;
    let id = cfg && cfg.id;
    let activa = estaActiva(cfg);
    if (opciones.forzar && TEMPORADAS[opciones.forzar]) { id = opciones.forzar; activa = true; }
    const t = activa ? TEMPORADAS[id] : null;

    // Limpiar lo anterior
    document.querySelectorAll(".temp-adorno, .temp-fondo, .temp-esquina").forEach(el => el.remove());
    detenerAnimacion();
    body.removeAttribute("data-temporada");
    body.style.removeProperty("--temp-icono");
    if (!t) return null;

    body.setAttribute("data-temporada", id);
    const usarColores = opciones.colores !== undefined ? opciones.colores : (cfg ? cfg.colores !== false : true);
    if (usarColores) {
      const v = varsCatalogo(t);
      VARS_CATALOGO.forEach(k => body.style.setProperty(k, v[k]));
    }
    body.style.setProperty("--temp-icono", JSON.stringify(t.icono));

    // Adorno colgando del encabezado
    const nav = document.querySelector("nav.navbar") || document.querySelector("header");
    const tira = adornoTira(t.adorno);
    if (nav && tira) {
      const ad = document.createElement("div");
      ad.className = "temp-adorno temp-adorno--" + t.adorno.tipo;
      ad.setAttribute("aria-hidden", "true");
      ad.style.backgroundImage = tira.fondo;
      ad.style.height = tira.alto + "px";
      ad.style.backgroundSize = tira.ancho + "px " + tira.alto + "px";
      nav.appendChild(ad);
      if (t.adorno.tipo === "telaranas") {
        ["izq", "der"].forEach(lado => {
          const e = document.createElement("div");
          e.className = "temp-esquina temp-esquina--" + lado;
          e.setAttribute("aria-hidden", "true");
          e.style.backgroundImage = telarana("#e5e7eb");
          nav.appendChild(e);
        });
      }
    }

    // Fondo con motivos
    const fondo = document.createElement("div");
    fondo.className = "temp-fondo";
    fondo.setAttribute("aria-hidden", "true");
    fondo.style.backgroundImage = fondoMotivos(t.fondo, t.colores.acento);
    body.appendChild(fondo);

    // Animación (después de que la página ya cargó lo importante)
    const conAnim = opciones.animacion !== undefined ? opciones.animacion : (cfg ? cfg.animacion !== false : true);
    if (conAnim) {
      const arrancar = () => iniciarAnimacion(t.animacion);
      if (document.readyState === "complete") setTimeout(arrancar, 600);
      else window.addEventListener("load", () => setTimeout(arrancar, 600), { once: true });
    }
    return t;
  }

  /* ---------- Aplicar en el panel admin (solo colores) ---------- */
  const VARS_PANEL = ["--navy-950", "--navy-900", "--navy-800", "--navy-700", "--navy-600", "--blue-500", "--blue-50"];

  function aplicarPanel(cfg) {
    const body = document.body;
    VARS_PANEL.forEach(k => body.style.removeProperty(k));
    body.removeAttribute("data-temporada");
    if (!estaActiva(cfg) || cfg.colores === false) return null;
    const t = TEMPORADAS[cfg.id];
    const v = varsPanel(t);
    VARS_PANEL.forEach(k => body.style.setProperty(k, v[k]));
    body.setAttribute("data-temporada", cfg.id);
    return t;
  }

  global.VeekTemporadas = { TEMPORADAS, ORDEN, hoyAR, estaActiva, aplicarCatalogo, aplicarPanel, adornoTira, fondoMotivos, detenerAnimacion };
})(typeof window !== "undefined" ? window : this);
