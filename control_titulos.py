"""
Cruza el reporte de SIU Guaraní (CSV de tramite de titulos) con el reporte
de SICER (xls) usando titulos.xlsx como tabla fija de equivalencia de
denominaciones de titulo.

Genera un Excel de salida con:
  - Alertas: DNI+Titulo en estado "Finalizado" en SICER pero que en SIU
    todavia no figura como "Diplomado" (es decir, SIU no actualizo el
    estado del tramite).
  - Detalle_Match: todos los cruces encontrados entre SIU y SICER.
  - Sin_Match_SICER: solicitudes de SIU para las que no se encontro un
    registro correspondiente en SICER (por DNI + titulo equivalente).
  - Demorados: tramites que todavia no son "Diplomado", ordenados por
    dias_totales_tramite descendente (los mas demorados primero).

Uso:
    python control_titulos.py

Requiere los 3 archivos en la misma carpeta:
    - data-*.csv          (reporte SIU Guarani)
    - "sicer consulta.xls" (reporte SICER)
    - titulos.xlsx          (tabla fija de equivalencia de titulos)
"""
import sys
from pathlib import Path

from titulos_core import build_dataframes, dataframes_to_excel_bytes

BASE_DIR = Path(__file__).resolve().parent
TITULOS_PATH = BASE_DIR / "titulos.xlsx"
DIAS_PATH = BASE_DIR / "Dias.xlsx"
OUTPUT_PATH = BASE_DIR / "control_titulos_sicer_vs_siu.xlsx"


def find_csv_path() -> Path:
    candidates = sorted(BASE_DIR.glob("data-*.csv"))
    if not candidates:
        sys.exit("No se encontro ningun archivo 'data-*.csv' en la carpeta.")
    return candidates[-1]


def find_sicer_path() -> Path:
    candidates = sorted(BASE_DIR.glob("sicer*.xls*"))
    if not candidates:
        sys.exit("No se encontro el reporte de SICER en la carpeta.")
    return candidates[-1]


def build_report():
    csv_path = find_csv_path()
    sicer_path = find_sicer_path()
    alertas, detalle, sin_match, demorados = build_dataframes(
        TITULOS_PATH, csv_path, sicer_path, DIAS_PATH
    )
    excel_bytes = dataframes_to_excel_bytes(alertas, detalle, sin_match, demorados)
    OUTPUT_PATH.write_bytes(excel_bytes)

    print(f"Reporte generado: {OUTPUT_PATH}")
    print(f"  Alertas encontradas: {len(alertas)}")
    print(f"  Cruces totales:      {len(detalle)}")
    print(f"  Sin match en SICER:  {len(sin_match)}")
    print(f"  Demorados (>=75 dias habiles): {int(demorados['alerta_demora'].sum())} de {len(demorados)}")


if __name__ == "__main__":
    build_report()
