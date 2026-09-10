"""
Tablero web para que el personal suba los reportes de SIU y SICER y vea
el cruce de estados con alertas, sin necesidad de usar la linea de comandos.

Uso:
    streamlit run app.py
"""
from pathlib import Path

import pandas as pd
import streamlit as st

from titulos_core import UMBRAL_DIAS_DEMORA, build_dataframes, dataframes_to_excel_bytes

BASE_DIR = Path(__file__).resolve().parent
TITULOS_PATH_DEFAULT = BASE_DIR / "titulos.xlsx"
DIAS_PATH_DEFAULT = BASE_DIR / "Dias.xlsx"


def safe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Copia del df con columnas de tipo mixto convertidas a texto.

    st.dataframe usa Arrow para serializar, que falla si una columna
    "object" mezcla tipos (por ej. fechas con NaN, o texto con numeros).
    """
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).replace({"nan": "", "NaT": "", "None": ""})
    return df


def download_csv_button(df: pd.DataFrame, label: str, file_name: str, key: str):
    st.download_button(
        label,
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=file_name,
        mime="text/csv",
        key=key,
    )


def render_kpi(label: str, value: int, detail: str, tone: str):
    st.markdown(
        f"""
        <div class="kpi-card kpi-{tone}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-detail">{detail}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="Control de Titulos SIU vs SICER", layout="wide")
st.title("Control de Titulos: SIU Guaraní vs SICER")
st.caption(
    "Sube el reporte CSV de SIU Guaraní y el reporte de SICER (xls) para ver en que "
    "estado se encuentra cada trámite y detectar casos donde SICER ya marcó "
    "'Finalizado' pero SIU todavía no figura como 'Diplomado'."
)
st.markdown(
    """
    <style>
    .kpi-card {
        min-height: 142px;
        padding: 20px 22px 17px;
        border: 1px solid #dce3ea;
        border-top: 5px solid #52718a;
        border-radius: 8px;
        background: #ffffff;
        box-shadow: 0 3px 10px rgba(27, 49, 65, 0.07);
    }
    .kpi-label {
        color: #52616d;
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        text-transform: uppercase;
    }
    .kpi-value {
        color: #182b39;
        font-size: 2.35rem;
        font-weight: 800;
        line-height: 1.15;
        margin-top: 10px;
    }
    .kpi-detail {
        color: #71808b;
        font-size: 0.78rem;
        margin-top: 7px;
    }
    .kpi-red { border-top-color: #c64b4b; }
    .kpi-red .kpi-value { color: #a63232; }
    .kpi-blue { border-top-color: #3f789d; }
    .kpi-amber { border-top-color: #c8922e; }
    .kpi-amber .kpi-value { color: #9b6a0e; }
    .kpi-orange { border-top-color: #d16c35; }
    .kpi-orange .kpi-value { color: #ad4f1d; }
    .summary-strip {
        margin: 18px 0 24px;
        padding: 13px 17px;
        border-left: 4px solid #3f789d;
        background: #edf5f8;
        color: #294657;
        font-size: 0.91rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Archivos")

    st.subheader("1. Tabla de títulos (fija)")
    usar_default = TITULOS_PATH_DEFAULT.exists()
    titulos_file = st.file_uploader(
        "titulos.xlsx (opcional si ya existe en el servidor)", type=["xlsx"]
    )
    if usar_default:
        st.caption(f"Se usará '{TITULOS_PATH_DEFAULT.name}' si no subís otro archivo.")

    st.subheader("2. Reporte SIU Guaraní")
    csv_file = st.file_uploader("data-*.csv", type=["csv"])

    st.subheader("3. Reporte SICER")
    sicer_file = st.file_uploader("sicer consulta.xls", type=["xls", "xlsx"])

    st.subheader("4. Feriados y vacaciones")
    dias_file = st.file_uploader("Dias.xlsx (opcional)", type=["xlsx", "xls"])
    if DIAS_PATH_DEFAULT.exists():
        st.caption(f"Se usará '{DIAS_PATH_DEFAULT.name}' si no subís otro archivo.")

    procesar = st.button("Procesar", type="primary", width="stretch")

titulos_source = titulos_file or (TITULOS_PATH_DEFAULT if usar_default else None)
dias_source = dias_file or (DIAS_PATH_DEFAULT if DIAS_PATH_DEFAULT.exists() else None)

if procesar:
    if titulos_source is None:
        st.error("Falta la tabla de títulos (titulos.xlsx).")
    elif csv_file is None:
        st.error("Falta el reporte CSV de SIU Guaraní.")
    elif sicer_file is None:
        st.error("Falta el reporte de SICER.")
    elif dias_source is None:
        st.error("Falta el calendario Dias.xlsx con feriados y vacaciones.")
    else:
        with st.spinner("Procesando..."):
            try:
                alertas, detalle, sin_match, demorados = build_dataframes(
                    titulos_source, csv_file, sicer_file, dias_source
                )
            except Exception as exc:  # noqa: BLE001 - mostrar error al usuario
                st.error(f"Ocurrió un error al procesar los archivos: {exc}")
            else:
                st.session_state["alertas"] = alertas
                st.session_state["detalle"] = detalle
                st.session_state["sin_match"] = sin_match
                st.session_state["demorados"] = demorados

if "detalle" in st.session_state:
    alertas: pd.DataFrame = st.session_state["alertas"]
    detalle: pd.DataFrame = st.session_state["detalle"]
    sin_match: pd.DataFrame = st.session_state["sin_match"]
    demorados: pd.DataFrame = st.session_state["demorados"]
    n_demorados = int(demorados["alerta_demora"].sum())

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi("Alertas", len(alertas), "SICER avanzó y SIU no acompañó", "red")
    with col2:
        render_kpi("Cruces totales", len(detalle), "Trámites encontrados en ambos reportes", "blue")
    with col3:
        render_kpi("Sin match en SICER", len(sin_match), "Solicitudes para revisar", "amber")
    with col4:
        render_kpi(
            f"Demorados (≥{UMBRAL_DIAS_DEMORA} días hábiles)",
            n_demorados,
            "Trámites aún no finalizados",
            "orange",
        )

    st.markdown(
        f'<div class="summary-strip"><strong>Lectura rápida:</strong> '
        f'{len(alertas)} alertas de actualización, {n_demorados} trámites demorados y '
        f'{len(sin_match)} solicitudes sin correspondencia en SICER.</div>',
        unsafe_allow_html=True,
    )

    if len(alertas) > 0:
        st.error(
            f"⚠️ Hay {len(alertas)} trámite(s) 'Finalizado' en SICER que SIU todavía "
            "no actualizó como 'Diplomado'."
        )
    else:
        st.success("No se encontraron alertas: SIU está al día con SICER.")

    if n_demorados > 0:
        st.warning(
            f"⏳ Hay {n_demorados} trámite(s) sin finalizar con al menos {UMBRAL_DIAS_DEMORA} días hábiles "
            "en trámite. Revisá la pestaña 'Demorados'."
        )

    excel_bytes = dataframes_to_excel_bytes(alertas, detalle, sin_match, demorados)
    st.download_button(
        "Descargar reporte Excel",
        data=excel_bytes,
        file_name="control_titulos_sicer_vs_siu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab_alertas, tab_detalle, tab_sin_match, tab_demorados = st.tabs(
        ["Alertas", "Detalle_Match", "Sin_Match_SICER", "Demorados"]
    )
    with tab_alertas:
        if len(alertas) > 0:
            st.subheader("En que paso de SIU están frenados los trámites en alerta")
            st.caption(
                "SICER ya dice 'Finalizado' pero SIU sigue mostrando estos estados "
                "(ej. '09° Diploma y CA generados'): hay que actualizar el trámite en SIU."
            )
            st.bar_chart(alertas["estado_siu"].value_counts())

            st.subheader("Alertas por título")
            st.bar_chart(alertas["propuesta_nombre"].value_counts())

        download_csv_button(alertas, "Descargar Alertas (CSV)", "alertas.csv", key="csv_alertas")
        st.dataframe(safe_for_display(alertas), width="stretch")
    with tab_detalle:
        st.subheader("Distribución de estados SIU (todos los cruces)")
        st.bar_chart(detalle["estado_siu"].value_counts())

        st.subheader("Distribución de estados SICER")
        st.bar_chart(detalle["estado_sicer"].value_counts())

        download_csv_button(detalle, "Descargar Detalle_Match (CSV)", "detalle_match.csv", key="csv_detalle")
        st.dataframe(safe_for_display(detalle), width="stretch")
    with tab_sin_match:
        download_csv_button(sin_match, "Descargar Sin_Match_SICER (CSV)", "sin_match_sicer.csv", key="csv_sin_match")
        st.dataframe(safe_for_display(sin_match), width="stretch")
    with tab_demorados:
        st.caption(
            f"Trámites que todavía no llegaron a 'Diplomado', ordenados por días totales "
            f"de trámite (los más demorados primero). Se marcan en rojo los que alcanzan "
            f"{UMBRAL_DIAS_DEMORA} días hábiles."
        )
        top_20 = demorados.head(20).set_index("apellido_nombres")[
            ["dias_totales_tramite", "dias_habiles_tramite"]
        ]
        st.subheader("Top 20 trámites más demorados (corridos vs hábiles)")
        st.bar_chart(top_20)

        download_csv_button(demorados, "Descargar Demorados (CSV)", "demorados.csv", key="csv_demorados")
        st.dataframe(
            safe_for_display(demorados).style.apply(
                lambda row: ["background-color: #ffc7ce" if row["alerta_demora"] else "" for _ in row],
                axis=1,
            ),
            width="stretch",
        )
else:
    st.info("Cargá los archivos en el panel de la izquierda y presioná 'Procesar'.")
