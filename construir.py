"""Construye la página de la subasta a partir de Subasta_Datos.xlsx.

Uso:  python construir.py
Lee las pestañas Evento y Obras, crea versiones livianas de las fotos en sitio/img
y escribe sitio/datos.js. El diseño está en sitio/index.html.
Requiere: pip install pillow openpyxl
"""
import json
import re
import sys
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageOps

AQUI = Path(__file__).resolve().parent
PROYECTO = AQUI.parent
EXCEL = PROYECTO / "Subasta_Datos.xlsx"
SITIO = AQUI / "sitio"
IMG = SITIO / "img"
TAMANOS = {"g": 1800, "p": 700}  # grande (visor) y pequeña (galería)


def slug(texto):
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-")[:60]


def buscar_foto(nombre):
    nombre = nombre.strip()
    for p in PROYECTO.rglob(nombre):
        if AQUI not in p.parents:
            return p
    return None


def optimizar(origen, base):
    """Devuelve {g: ruta, p: ruta, w, h}; solo reprocesa si la foto cambió."""
    salida = {}
    with Image.open(origen) as im:
        im.draft("RGB", (TAMANOS["g"] * 2, TAMANOS["g"] * 2))
        im = ImageOps.exif_transpose(im).convert("RGB")
        for clave, lado in TAMANOS.items():
            destino = IMG / f"{base}-{clave}.jpg"
            if not destino.exists() or destino.stat().st_mtime < origen.stat().st_mtime:
                copia = im.copy()
                copia.thumbnail((lado, lado), Image.LANCZOS)
                copia.save(destino, "JPEG", quality=82, progressive=True, optimize=True)
            salida[clave] = f"img/{destino.name}"
        salida["w"], salida["h"] = im.size
    return salida


def main():
    IMG.mkdir(parents=True, exist_ok=True)
    recursos = SITIO / "recursos"
    recursos.mkdir(exist_ok=True)
    textura = PROYECTO / "Para Organizar" / "Textura para fondos.jpeg"
    if textura.exists() and not (recursos / "textura.jpg").exists():
        with Image.open(textura) as t:
            t = ImageOps.grayscale(t)
            t.thumbnail((1400, 1400), Image.LANCZOS)
            t.save(recursos / "textura.jpg", "JPEG", quality=70, progressive=True)

    wb = load_workbook(EXCEL, data_only=True)

    evento = {}
    for campo, valor, *_ in wb["Evento"].iter_rows(min_row=2, values_only=True):
        if campo:
            evento[str(campo).strip()] = "" if valor is None else str(valor).strip()
    evento["como_pujar"] = [evento.pop(k) for k in sorted(evento) if k.startswith("como_pujar_") and evento[k]]
    evento["como_pujar"] = [v for v in evento["como_pujar"] if v]

    hoja = wb["Obras"]
    cols = [str(c.value).strip() for c in hoja[1]]
    obras, avisos, usados = [], [], set()
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        d = dict(zip(cols, fila))
        if not d.get("titulo"):
            continue
        if str(d.get("publicar") or "").strip().lower() not in ("sí", "si", "s", "x", "yes"):
            continue
        base = slug(d["titulo"])
        while base in usados:
            base += "-b"
        usados.add(base)
        fotos = []
        for i, nombre in enumerate(n for n in str(d.get("fotos") or "").split(",") if n.strip()):
            ruta = buscar_foto(nombre)
            if not ruta:
                avisos.append(f"  · No encontré la foto '{nombre.strip()}' de «{d['titulo']}»")
                continue
            fotos.append(optimizar(ruta, f"{base}-{i + 1}"))
        if not fotos:
            avisos.append(f"  · «{d['titulo']}» no tiene fotos; no se publica")
            continue
        precio = d.get("precio_base")
        try:
            precio = int(float(str(precio).replace(".", "").replace("$", "").replace(",", "").strip()))
        except (TypeError, ValueError):
            precio = None
        obras.append({
            "id": base,
            "titulo": str(d["titulo"]).strip(),
            "artista": str(d.get("artista") or "").strip(),
            "tecnica": str(d.get("tecnica") or "").strip(),
            "soporte": str(d.get("soporte") or "").strip(),
            "medidas": str(d.get("medidas") or "").strip(),
            "precio": precio,
            "descripcion": str(d.get("descripcion") or "").strip(),
            "fotos": fotos,
        })
        o = obras[-1]
        o["pendiente"] = precio is None or any("confirmar" in (o[k] or "").lower()
                                               for k in ("titulo", "artista", "tecnica", "medidas"))

    # Borra imágenes de obras que ya no se publican
    vigentes = {f[k] for o in obras for f in o["fotos"] for k in TAMANOS}
    for p in IMG.glob("*.jpg"):
        if f"img/{p.name}" not in vigentes:
            p.unlink()

    datos = {"evento": evento, "obras": obras}
    (SITIO / "datos.js").write_text(
        "window.SUBASTA = " + json.dumps(datos, ensure_ascii=False, indent=1) + ";\n", encoding="utf-8")
    print(f"Listo: {len(obras)} obras publicadas en {SITIO}")
    if avisos:
        print("Revisar:\n" + "\n".join(avisos))


if __name__ == "__main__":
    sys.exit(main())
