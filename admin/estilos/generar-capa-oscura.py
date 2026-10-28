#!/usr/bin/env python3
"""
Genera estilos/capa-oscura.css a partir de admin.css.

Los estilos oscuros del panel (Turno largo, Pizarrón, Degradé) necesitan
cambiar los fondos blancos y los textos oscuros que admin.css tiene
escritos fijos (más de 100). En vez de copiarlos a mano —y que queden
desactualizados cuando cambie admin.css— este script los encuentra y
arma las reglas de reemplazo, todas bajo html[data-oscuro].

Uso (desde web/admin):   python3 estilos/generar-capa-oscura.py
Volver a correrlo cada vez que se cambie admin.css.
"""
import re, os

AQUI = os.path.dirname(os.path.abspath(__file__))
CSS = open(os.path.join(AQUI, "..", "admin.css"), encoding="utf-8").read()
CSS = re.sub(r"/\*.*?\*/", "", CSS, flags=re.S)

def hex_a_rgb(h):
    h = h.lstrip("#")
    if len(h) == 3: h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def luminancia(rgb):
    r, g, b = [c / 255 for c in rgb]
    return 0.2126 * r + 0.7152 * g + 0.0722 * b

def es_claro(valor):
    v = valor.strip().lower()
    if v in ("white", "#fff", "#ffffff"): return True
    m = re.match(r"#([0-9a-f]{3}|[0-9a-f]{6})\b", v)
    if m: return luminancia(hex_a_rgb(m.group(0))) > 0.86
    m = re.match(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+))?", v)
    if m:
        a = float(m.group(4)) if m.group(4) else 1
        return a > 0.6 and luminancia(tuple(int(m.group(i)) for i in (1, 2, 3))) > 0.86
    return bool(re.match(r"var\(--(slate-50|slate-100)\)", v))

def es_tinte_claro(valor):
    return bool(re.match(r"var\(--(blue|green|amber|red|indigo)-50\)", valor.strip()))

def es_texto_oscuro(valor):
    v = valor.strip().lower()
    if re.match(r"var\(--navy-(950|900|800|700)\)", v): return True
    if re.match(r"var\(--slate-(700|600)\)", v): return False  # la escala slate ya se invierte
    m = re.match(r"#([0-9a-f]{3}|[0-9a-f]{6})\b", v)
    return bool(m) and luminancia(hex_a_rgb(m.group(0))) < 0.3

fondos, tintes, textos, bordes = [], [], [], []
for sel, cuerpo in re.findall(r"([^{}@]+)\{([^{}]*)\}", CSS):
    sel = " ".join(sel.split())
    if not sel or sel.startswith(("from", "to", "0%", "100%")) or re.match(r"^[\d.]+%", sel): continue
    if "@" in sel or ":root" in sel or "data-tema" in sel or "print" in sel: continue
    props = dict()
    for decl in cuerpo.split(";"):
        if ":" in decl:
            k, v = decl.split(":", 1)
            props[k.strip().lower()] = v.strip()
    fondo = props.get("background-color") or props.get("background") or ""
    primer = re.split(r"\s+(?![^(]*\))", fondo)[0] if fondo else ""
    if primer and "gradient" not in fondo and es_claro(primer): fondos.append(sel)
    elif primer and es_tinte_claro(primer): tintes.append(sel)
    color = props.get("color", "")
    if color and es_texto_oscuro(color): textos.append(sel)
    for k in ("border", "border-color", "border-bottom", "border-top"):
        v = props.get(k, "")
        mh = re.search(r"#[0-9a-fA-F]{3,6}\b", v)
        if mh and luminancia(hex_a_rgb(mh.group(0))) > 0.75: bordes.append(sel); break

def prefijar(sels):
    salida = []
    for s in sels:
        for parte in s.split(","):
            parte = parte.strip()
            if not parte: continue
            if parte.startswith("body"):
                salida.append("html[data-oscuro] " + parte)
            elif parte.startswith("html"):
                continue
            else:
                salida.append("html[data-oscuro] " + parte)
    return ",\n".join(sorted(set(salida)))

out = ["/* GENERADO por estilos/generar-capa-oscura.py — no editar a mano. */",
       "@media screen{",
       prefijar(fondos) + "{ background-color:var(--e-sup) !important; background-image:none !important; }",
       prefijar(tintes) + "{ background-color:var(--e-sup-2) !important; }",
       prefijar(textos) + "{ color:var(--e-texto) !important; }",
       (prefijar(bordes) + "{ border-color:var(--e-linea) !important; }") if bordes else "",
       "}"]
open(os.path.join(AQUI, "capa-oscura.css"), "w", encoding="utf-8").write("\n".join(out) + "\n")
print(f"fondos={len(fondos)} tintes={len(tintes)} textos={len(textos)} bordes={len(bordes)}")
