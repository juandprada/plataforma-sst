"""Copia y normaliza los logos referenciados en data/empresas.json.

Por cada empresa toma el nombre de archivo de la columna LOGO (ruta absoluta de
Windows), lo busca en la carpeta de logos original, lo redimensiona a un alto
comun (para que pesen poco y encajen en la caja fija del encabezado) y lo guarda
como logos/<_id>.png. Luego reescribe empresas.json con la ruta web relativa.

Uso:
    python tools/normalize_logos.py \
        --src "../1.PLANEAR/Input" \
        --empresas data/empresas.json \
        --out logos
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from PIL import Image

TARGET_HEIGHT = 200  # px (a 2x del alto de la caja ~100px en pantalla)


def candidate_names(basename: str) -> list[str]:
    """Nombres a probar: el original y una version sin digito final.

    'LogoARMAR2.png' -> ['LogoARMAR2.png', 'LogoARMAR.png']
    """
    names = [basename]
    stem, ext = Path(basename).stem, Path(basename).suffix
    stripped = re.sub(r"\d+$", "", stem)
    if stripped != stem:
        names.append(stripped + ext)
    return names


def find_source(src_dir: Path, logo_value: str) -> Path | None:
    if not logo_value:
        return None
    basename = Path(logo_value.replace("\\", "/")).name
    for name in candidate_names(basename):
        for f in src_dir.iterdir():
            if f.name.lower() == name.lower():
                return f
    return None


def normalize(img_path: Path, out_path: Path) -> None:
    img = Image.open(img_path)
    # Fondo transparente donde se pueda; convierte a RGBA.
    img = img.convert("RGBA")
    if img.height > TARGET_HEIGHT:
        ratio = TARGET_HEIGHT / img.height
        img = img.resize(
            (max(1, round(img.width * ratio)), TARGET_HEIGHT), Image.LANCZOS
        )
    img.save(out_path, "PNG", optimize=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default="../1.PLANEAR/Input")
    ap.add_argument("--empresas", default="data/empresas.json")
    ap.add_argument("--out", default="logos")
    args = ap.parse_args()

    src_dir = Path(args.src)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    empresas = json.loads(Path(args.empresas).read_text(encoding="utf-8"))

    ok, missing = 0, []
    for emp in empresas:
        eid = emp["_id"]
        source = find_source(src_dir, emp.get("LOGO", ""))
        if source is None:
            missing.append(emp.get("EMPRESA", eid))
            emp["LOGO"] = ""  # UI usara el fallback (nombre en recuadro)
            continue
        out_path = out_dir / f"{eid}.png"
        normalize(source, out_path)
        emp["LOGO"] = f"logos/{eid}.png"
        ok += 1

    Path(args.empresas).write_text(
        json.dumps(empresas, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"OK: {ok} logos normalizados -> {out_dir}")
    if missing:
        print("SIN LOGO (usaran fallback):", ", ".join(missing))


if __name__ == "__main__":
    main()
