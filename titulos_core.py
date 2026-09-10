"""
Logica compartida para cruzar SIU Guarani (CSV) con SICER (xls) usando
titulos.xlsx como tabla fija de equivalencia de denominaciones de titulo.

Este modulo no depende de rutas fijas en disco: todas las funciones aceptan
tanto una ruta (str/Path) como un objeto tipo archivo (por ejemplo, el que
entrega un file uploader de Streamlit), lo que permite reutilizar la misma
logica tanto desde la linea de comandos como desde una interfaz web.
"""
import io
import re

import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill

ESTADO_DIPLOMADO_SIU = "Diplomado"
ESTADO_FINALIZADO_SICER = "Finalizado"

# Orden esperado de avance en SIU (numero de paso). "Diplomado" es el ultimo.
_ORDEN_DIPLOMADO = 15

# Por cada estado de SICER, el paso minimo que SIU deberia haber alcanzado ya.
# Si SIU esta por detras de ese paso, es porque no actualizo el tramite.
REQUISITO_SIU_POR_ESTADO_SICER = {
    "Carga Finalizada": (8, "08° Carga SICER finalizada"),
    "Finalizado": (_ORDEN_DIPLOMADO, ESTADO_DIPLOMADO_SIU),
}

# a partir de cuantos dias sin terminar un tramite se considera "demorado"
UMBRAL_DIAS_DEMORA = 100

COLS_ORDER = [
    "dni",
    "apellido_nombres",
    "propuesta_nombre",
    "titulo_sicer",
    "nro_solicitud",
    "nro_solicitud_sicer",
    "estado_siu",
    "fecha_ultimo_cambio_siu",
    "estado_sicer",
    "fecha_sicer",
    "dias_totales_tramite",
    "dias_habiles_tramite",
    "nro_expediente_descr",
]

COLS_ORDER_ALERTAS = COLS_ORDER + ["estado_siu_esperado"]


def extract_dni(series: pd.Series) -> pd.Series:
    return series.astype(str).str.extract(r"(\d+)\s*$")[0]


def _orden_siu(estado_siu: str) -> int:
    """Numero de paso de un estado SIU, para poder compararlos entre si."""
    if estado_siu == ESTADO_DIPLOMADO_SIU:
        return _ORDEN_DIPLOMADO
    match = re.match(r"(\d+)", str(estado_siu))
    return int(match.group(1)) if match else -1


def load_dias_no_laborables(dias_source) -> np.ndarray:
    dias = pd.read_excel(dias_source)
    dias["Dias"] = pd.to_datetime(dias["Dias"], errors="coerce")
    return dias["Dias"].dropna().dt.strftime("%Y-%m-%d").to_numpy(dtype="datetime64[D]")


def _dias_habiles(
    fecha_inicio: pd.Series, dias_corridos: pd.Series, dias_no_laborables: np.ndarray
) -> pd.Series:
    """Dias habiles (sin contar sabados/domingos) entre el inicio del tramite
    y la fecha de referencia usada para calcular dias_totales_tramite."""
    inicio = pd.to_datetime(fecha_inicio, errors="coerce")
    referencia = inicio + pd.to_timedelta(dias_corridos, unit="D", errors="coerce")

    def _contar(i, r):
        if pd.isna(i) or pd.isna(r):
            return None
        return int(np.busday_count(i.date(), r.date(), holidays=dias_no_laborables))

    return pd.Series(
        [_contar(i, r) for i, r in zip(inicio, referencia)], index=fecha_inicio.index
    )


def load_titulos_map(titulos_source) -> dict:
    """propuesta_nombre (SIU) -> lista de Titulo (SICER) posibles."""
    t = pd.read_excel(titulos_source)
    t["propuesta_nombre"] = t["propuesta_nombre"].str.strip().str.upper()
    t["Título"] = t["Título"].str.strip()
    mapping: dict = {}
    for prop, titulo in zip(t["propuesta_nombre"], t["Título"]):
        mapping.setdefault(prop, []).append(titulo)
    return mapping


def load_siu_estado_actual(
    csv_source, titulos_map: dict, dias_no_laborables: np.ndarray
) -> pd.DataFrame:
    df = pd.read_csv(csv_source, encoding="utf-8")
    df["dni"] = extract_dni(df["tipo_nro_documento"])
    df["propuesta_nombre"] = df["propuesta_nombre"].str.strip().str.upper()

    # solo nos interesan las propuestas que estan en la tabla fija de titulos
    df = df[df["propuesta_nombre"].isin(titulos_map.keys())]

    # el estado actual de cada solicitud es el de la fila con mayor "paso"
    # (algunas solicitudes recien iniciadas todavia no tienen "paso" cargado)
    idx_actual = df.assign(_paso=df["paso"].fillna(-1)).groupby("nro_solicitud")["_paso"].idxmax()
    actual = df.loc[
        idx_actual,
        [
            "nro_solicitud",
            "dni",
            "alumno",
            "apellido_nombres",
            "propuesta_nombre",
            "nro_expediente_descr",
            "fecha",
            "estado_nuevo",
            "fecha_inicio_tramite",
            "dias_totales_tramite",
        ],
    ].rename(columns={"fecha": "fecha_ultimo_cambio_siu", "estado_nuevo": "estado_siu"})
    actual["estado_siu"] = actual["estado_siu"].fillna("Sin estado registrado (recien solicitado)")
    actual["dias_habiles_tramite"] = _dias_habiles(
        actual["fecha_inicio_tramite"],
        actual["dias_totales_tramite"],
        dias_no_laborables,
    )
    actual = actual.drop(columns=["fecha_inicio_tramite"])

    # cada propuesta puede corresponder a mas de un titulo SICER (ambiguo)
    actual["titulo_candidato"] = actual["propuesta_nombre"].map(
        lambda p: titulos_map.get(p, [None])
    )
    actual = actual.explode("titulo_candidato", ignore_index=True)
    return actual


def load_sicer(sicer_source, titulos_validos: set) -> pd.DataFrame:
    s = pd.read_excel(sicer_source)
    s["dni"] = extract_dni(s["Número Documento"])
    s["Título"] = s["Título"].str.strip()
    # solo nos interesan los titulos que estan en la tabla fija de titulos
    s = s[s["Título"].isin(titulos_validos)]
    return s[["dni", "Título", "Nº Solicitud", "Estado", "Fecha"]].rename(
        columns={
            "Título": "titulo_sicer",
            "Nº Solicitud": "nro_solicitud_sicer",
            "Estado": "estado_sicer",
            "Fecha": "fecha_sicer",
        }
    )


def build_dataframes(titulos_source, csv_source, sicer_source, dias_source):
    """Devuelve (alertas, detalle, sin_match, demorados) a partir de las 3 fuentes."""
    titulos_map = load_titulos_map(titulos_source)
    titulos_validos = {t for titulos in titulos_map.values() for t in titulos}
    dias_no_laborables = load_dias_no_laborables(dias_source)
    siu = load_siu_estado_actual(csv_source, titulos_map, dias_no_laborables)
    sicer = load_sicer(sicer_source, titulos_validos)

    merged = siu.merge(
        sicer,
        left_on=["dni", "titulo_candidato"],
        right_on=["dni", "titulo_sicer"],
        how="left",
    )

    # si una solicitud tenia varios titulos candidatos, nos quedamos con el
    # que efectivamente hizo match en SICER (o con una sola fila si ninguno matcheo)
    merged["tiene_match"] = merged["titulo_sicer"].notna()
    merged = merged.sort_values("tiene_match", ascending=False)
    merged = merged.drop_duplicates(subset=["nro_solicitud"], keep="first")

    detalle = merged[merged["tiene_match"]].drop(columns=["titulo_candidato", "tiene_match"])
    sin_match = merged[~merged["tiene_match"]].drop(
        columns=["titulo_candidato", "tiene_match", "titulo_sicer", "nro_solicitud_sicer", "estado_sicer", "fecha_sicer"]
    )

    # una solicitud es alerta si SICER ya avanzo a un paso (ej. "Carga Finalizada"
    # o "Finalizado") que SIU todavia no refleja en su propio estado
    requisito = detalle["estado_sicer"].map(REQUISITO_SIU_POR_ESTADO_SICER)
    detalle["_orden_minimo_esperado"] = requisito.map(lambda r: r[0] if isinstance(r, tuple) else None)
    detalle["estado_siu_esperado"] = requisito.map(lambda r: r[1] if isinstance(r, tuple) else None)
    detalle["_orden_siu_actual"] = detalle["estado_siu"].map(_orden_siu)

    alertas = detalle[
        detalle["_orden_minimo_esperado"].notna()
        & (detalle["_orden_siu_actual"] < detalle["_orden_minimo_esperado"])
    ].copy()

    detalle = detalle.drop(columns=["_orden_minimo_esperado", "_orden_siu_actual"])
    alertas = alertas.drop(columns=["_orden_minimo_esperado", "_orden_siu_actual"])

    detalle = detalle.drop(columns=["estado_siu_esperado"])[COLS_ORDER]
    alertas = alertas[COLS_ORDER_ALERTAS]

    # tramites que todavia no terminaron (no son "Diplomado") ordenados por
    # cuantos dias llevan en tramite, para ver cuales son los mas demorados
    demorados = pd.concat([detalle, sin_match], ignore_index=True)
    demorados = demorados[demorados["estado_siu"] != ESTADO_DIPLOMADO_SIU].copy()
    demorados["alerta_demora"] = demorados["dias_totales_tramite"] >= UMBRAL_DIAS_DEMORA
    demorados = demorados.sort_values("dias_totales_tramite", ascending=False, na_position="last")

    return alertas, detalle, sin_match, demorados


def dataframes_to_excel_bytes(
    alertas: pd.DataFrame, detalle: pd.DataFrame, sin_match: pd.DataFrame, demorados: pd.DataFrame
) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        alertas.to_excel(writer, sheet_name="Alertas", index=False)
        detalle.to_excel(writer, sheet_name="Detalle_Match", index=False)
        sin_match.to_excel(writer, sheet_name="Sin_Match_SICER", index=False)
        demorados.to_excel(writer, sheet_name="Demorados", index=False)

    buffer.seek(0)
    styled = _style_workbook(buffer)
    return styled


def _style_workbook(buffer: io.BytesIO) -> bytes:
    from openpyxl import load_workbook

    wb = load_workbook(buffer)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    alert_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for name in wb.sheetnames:
        ws = wb[name]
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)
        if name == "Alertas":
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.fill = alert_fill
        if name == "Demorados":
            headers = [c.value for c in ws[1]]
            col_idx = headers.index("alerta_demora") + 1 if "alerta_demora" in headers else None
            if col_idx:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    if row[col_idx - 1].value:
                        for cell in row:
                            cell.fill = alert_fill

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
