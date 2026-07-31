#!/usr/bin/env python3
"""
gtc45_to_json.py — Genera/actualiza data/gtc45.json para la plataforma SST.

Fuentes:
  - reportes/gtc45/gtc45results.json  (resultados de la encuesta SurveyJS)
  - reportes/gtc45/mapeo_peligros.csv  (catálogo de peligros + NC + controles)

Modelo de dos archivos:
  - data/gtc45.json es el STORAGE (acumula todas las empresas procesadas).
  - gtc45results.json es la ENTRADA NUEVA (se consume y se mezcla).

El script hace merge/upsert: empresas nuevas se agregan; existentes se
reemplazan si la encuesta es más reciente (por HappendAt).

Uso:
  python tools/gtc45_to_json.py
  python tools/gtc45_to_json.py --resultados path/to/results.json
"""

import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path

# Rutas por defecto (relativas a la raíz del repo plataforma-sst).
SCRIPT_DIR = Path(__file__).resolve().parent
PLATAFORMA_DIR = SCRIPT_DIR.parent
SST_DIR = PLATAFORMA_DIR.parent  # a:/seagate/sst

DEFAULT_RESULTS = SST_DIR / "reportes" / "gtc45" / "gtc45results.json"
DEFAULT_CSV = SST_DIR / "reportes" / "gtc45" / "mapeo_peligros.csv"
DEFAULT_OUTPUT = PLATAFORMA_DIR / "data" / "gtc45.json"


def normalizar_nombre(nombre):
    """Normaliza un nombre de empresa para matching flexible.

    Quita sufijos societarios (SAS, SA, LTDA…), espacios extra, y pasa a minúsculas.
    Ej: "TE RECUPERAMOS SAS" → "te recuperamos"
        "Te recuperamos "   → "te recuperamos"
    """
    s = nombre.strip().lower()
    # Quitar sufijos societarios comunes (Colombia).
    s = re.sub(r'\b(s\.?a\.?s\.?|s\.?a\.?|ltda\.?|e\.?u\.?)\b', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def parse_dotnet_date(s):
    """Convierte /Date(1750955820168)/ a timestamp en ms (int)."""
    if not s:
        return 0
    m = re.search(r'/Date\((\d+)\)/', str(s))
    return int(m.group(1)) if m else 0


def extraer_peligros(submission):
    """Extrae los peligros evaluados de una sumisión de encuesta.

    Devuelve un dict { panel_id: { deficiencia, exposicion, cargos } }
    solo para los peligros con deficiencia != "No Aplica".
    """
    peligros = {}
    keys = list(submission.keys())
    # Detectar prefijos de peligros (BIO_01, FIS_02, BIO_M_01, SEG_03, NAT_01…).
    prefixes = set()
    for k in keys:
        m = re.match(r'^([A-Z]{2,4}_(?:[A-Z0-9]{1,2}_)?\d{2})_', k)
        if m:
            prefixes.add(m.group(1))

    for prefix in sorted(prefixes):
        def_key = f"{prefix}_deficiencia"
        exp_key = f"{prefix}_exposicion"
        cargos_key = f"{prefix}_cargos"

        deficiencia = submission.get(def_key)
        if not deficiencia or deficiencia == "No Aplica":
            continue  # No aplica → no entra en la matriz

        exposicion = submission.get(exp_key, "")
        cargos_raw = submission.get(cargos_key, [])
        # cargos puede ser string, lista de strings, o lista anidada.
        if isinstance(cargos_raw, str):
            cargos = [cargos_raw]
        elif isinstance(cargos_raw, list):
            cargos = []
            for c in cargos_raw:
                if isinstance(c, list):
                    cargos.extend(c)
                else:
                    cargos.append(str(c))
        else:
            cargos = []

        peligros[prefix] = {
            "deficiencia": deficiencia,
            "exposicion": exposicion,
            "cargos": cargos,
        }

    return peligros


def leer_resultados(ruta):
    """Lee gtc45results.json y devuelve lista de sumisiones."""
    with open(ruta, "r", encoding="utf-8") as f:
        data = json.load(f)
    # Estructura: { ResultCount: N, Data: [...] }
    items = data.get("Data", [])
    if isinstance(items, dict):
        items = [items]
    return items


def leer_mapeo(ruta):
    """Lee mapeo_peligros.csv y devuelve lista de dicts."""
    rows = []
    # El CSV puede tener encoding latin1 (viene de Excel/R en Windows).
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(ruta, "r", encoding=enc) as f:
                f.read()
            break
        except UnicodeDecodeError:
            continue
    with open(ruta, "r", encoding=enc) as f:
        reader = csv.DictReader(f)
        for row in reader:
            # NC a número.
            try:
                row["NC"] = int(row.get("NC", 0))
            except (ValueError, TypeError):
                row["NC"] = 0
            rows.append(row)
    return rows


def leer_storage(ruta):
    """Lee el storage existente (si existe)."""
    if not ruta.exists():
        return {"mapeo_peligros": [], "resultados_por_empresa": {}}
    with open(ruta, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    parser = argparse.ArgumentParser(description="Genera/actualiza data/gtc45.json")
    parser.add_argument("--resultados", type=Path, default=DEFAULT_RESULTS,
                        help="Ruta a gtc45results.json")
    parser.add_argument("--csv", type=Path, default=DEFAULT_CSV,
                        help="Ruta a mapeo_peligros.csv")
    parser.add_argument("--output", "-o", type=Path, default=DEFAULT_OUTPUT,
                        help="Ruta de salida (data/gtc45.json)")
    args = parser.parse_args()

    # 1. Leer mapeo de peligros (catálogo estático).
    if not args.csv.exists():
        print(f"ERROR: No se encontró {args.csv}", file=sys.stderr)
        sys.exit(1)
    mapeo = leer_mapeo(args.csv)
    print(f"Catálogo de peligros: {len(mapeo)} entradas")

    # 2. Leer storage existente.
    storage = leer_storage(args.output)
    existentes = storage.get("resultados_por_empresa", {})
    print(f"Storage existente: {len(existentes)} empresa(s)")

    # 3. Leer resultados nuevos.
    if not args.resultados.exists():
        print(f"AVISO: No se encontró {args.resultados}. Solo se actualiza el catálogo.")
        items = []
    else:
        items = leer_resultados(args.resultados)
        print(f"Resultados nuevos: {len(items)} encuesta(s)")

    # 4. Merge/upsert.
    for submission in items:
        nombre = submission.get("nombre_empresa", "").strip()
        if not nombre:
            print("  AVISO: Encuesta sin nombre_empresa, se omite.")
            continue

        clave = normalizar_nombre(nombre)
        ts = parse_dotnet_date(submission.get("HappendAt", ""))

        # ¿Ya existe y es más reciente?
        if clave in existentes:
            ts_existente = existentes[clave].get("_timestamp", 0)
            if ts <= ts_existente:
                print(f"  {nombre!r} → ya existe con fecha más reciente, se omite.")
                continue
            print(f"  {nombre!r} → ACTUALIZACIÓN (encuesta más reciente).")
        else:
            print(f"  {nombre!r} → NUEVA empresa.")

        peligros = extraer_peligros(submission)

        existentes[clave] = {
            "nombre_original": nombre,
            "_clave": clave,
            "_timestamp": ts,
            "trabajadores": {
                "admin": int(submission.get("question1") or 0),
                "operarios": int(submission.get("question2") or 0),
                "servicios": int(submission.get("question3") or 0),
                "conductores": int(submission.get("question4") or 0),
            },
            "peligros": peligros,
        }

    # 5. Escribir storage.
    # Simplificar el mapeo para el JSON de la plataforma.
    mapeo_limpio = []
    for row in mapeo:
        mapeo_limpio.append({
            "panel_id": row.get("panel_id", ""),
            "proceso": row.get("Proceso", ""),
            "actividades": row.get("Actividades", ""),
            "nc": row["NC"],
            "efectos_posibles": row.get("Efectos_Posibles", ""),
            "peor_consecuencia": row.get("Peor_Consecuencia", ""),
            "control_fuente": row.get("Control_Fuente", "") or "",
            "control_medio": row.get("Control_Medio", ""),
            "control_individuo": row.get("Control_Individuo", ""),
            "control_epp": row.get("Control_EPP", ""),
            "requisito_legal": row.get("Requisito_Legal", ""),
        })

    output = {
        "mapeo_peligros": mapeo_limpio,
        "resultados_por_empresa": existentes,
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\nStorage actualizado: {args.output}")
    print(f"  {len(mapeo_limpio)} peligros en catálogo")
    print(f"  {len(existentes)} empresa(s) con resultados")


if __name__ == "__main__":
    main()
