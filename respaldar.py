"""Guarda una copia (zip con fecha y hora) de la página y de Subasta_Datos.xlsx.

Uso:  python respaldar.py [nota]
En el PC con Windows guarda en el disco E: (fuera de Google Drive).
Si E: no existe (por ejemplo en el Mac), guarda en 043_Subasta4Diciembre/_Respaldos.
"""
import sys
import zipfile
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
PROYECTO = AQUI.parent
LOCAL = Path(r"E:\Juan David CM (E)\Archivos(E)\Respaldos\043_Subasta4Diciembre")
DESTINO = LOCAL if LOCAL.parent.exists() else PROYECTO / "_Respaldos"


def main():
    nota = "_" + "-".join(sys.argv[1:]) if len(sys.argv) > 1 else ""
    DESTINO.mkdir(parents=True, exist_ok=True)
    archivo = DESTINO / f"{datetime.now():%Y-%m-%d_%H%M}_pagina{nota}.zip"
    with zipfile.ZipFile(archivo, "w", zipfile.ZIP_DEFLATED) as z:
        for p in [PROYECTO / "Subasta_Datos.xlsx", *PROYECTO.glob("*.docx"), *AQUI.rglob("*")]:
            if p.is_file() and "__pycache__" not in p.parts and "_Respaldos" not in p.parts:
                z.write(p, p.relative_to(PROYECTO))
    print(f"Respaldo: {archivo} ({archivo.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
