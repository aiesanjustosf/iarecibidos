# ia_afip_recibidos.py
# Conversión de ARCA "Recibidos" -> Formato Holistor
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

# --- Rutas de assets ---
HERE = Path(__file__).parent
LOGO = HERE / "logo_aie.png"
FAVICON = HERE / "favicon-aie.ico"

st.set_page_config(
    page_title="ARCA Recibidos → Formato Holistor",
    page_icon=str(FAVICON) if FAVICON.exists() else None,
    layout="centered",
)

# --- Encabezado con logo ---
if LOGO.exists():
    st.image(str(LOGO), width=180)

st.title("ARCA Recibidos → Formato Holistor")

st.write(
    "Subí el Excel original descargado de **ARCA** (Libro IVA Digital - Compras/Recibidos) "
    "y descargá un archivo listo para importar en **Holistor**."
)

uploaded = st.file_uploader(
    "Subí el archivo de ARCA (.xlsx)",
    type=["xlsx"]
)


def map_tipo_letra(concepto: str):
    """Devuelve (Tipo, Letra) según el texto 'Tipo' de ARCA."""
    concepto = str(concepto).strip()

    # Tipo: F / ND / NC
    if "Nota de Crédito" in concepto:
        tipo = "NC"
    elif "Nota de Débito" in concepto:
        tipo = "ND"
    elif "Factura" in concepto:
        tipo = "F"
    else:
        tipo = ""

    # Letra:
    # - en general, última letra (A/B/C)
    # - caso especial: '8 - Nota de Crédito C' => letra B
    if concepto.startswith("8 "):
        letra = "B"
    else:
        letra = concepto[-1] if concepto else ""

    return tipo, letra


if uploaded is None:
    st.stop()

# --- LECTURA DEL EXCEL DE ARCA ---

# header=1 porque la fila 2 del archivo tiene los encabezados reales
df = pd.read_excel(uploaded, sheet_name=0, header=1)

# Nombres de columnas según ARCA
COL_FECHA = "Fecha"
COL_TIPO_AFIP = "Tipo"
COL_PV = "Punto de Venta"
COL_NRO_DESDE = "Número Desde"
COL_NRO_HASTA = "Número Hasta"
COL_CUIT_EMISOR = "Nro. Doc. Emisor"
COL_NOM_EMISOR = "Denominación Emisor"
COL_TC = "Tipo Cambio"
COL_MON = "Moneda"

COL_IVA_105 = "IVA 10,5%"
COL_NETO_105 = "Neto Grav. IVA 10,5%"
COL_IVA_21 = "IVA 21%"
COL_NETO_21 = "Neto Grav. IVA 21%"
COL_IVA_27 = "IVA 27%"
COL_NETO_27 = "Neto Grav. IVA 27%"
COL_NETO_NG = "Neto No Gravado"
COL_EXENTAS = "Op. Exentas"
COL_OTROS = "Otros Tributos"
COL_TOTAL = "Imp. Total"

registros = []


def get_num(row, col):
    """Devuelve número limpio (NaN -> 0)."""
    v = row.get(col, 0)
    if pd.isna(v):
        return 0.0
    return float(v)


for _, row in df.iterrows():
    concepto = str(row.get(COL_TIPO_AFIP, "")).strip()
    if not concepto:
        continue

    tipo, letra = map_tipo_letra(concepto)
    es_nc = "Nota de Crédito" in concepto

    # Función para aplicar el signo correcto
    def s(valor: float) -> float:
        if valor == 0:
            return 0.0
        return -abs(valor) if es_nc else abs(valor)

    # Base común
    base = {
        "Fecha Emisión": row.get(COL_FECHA),
        "Fecha Recepción": row.get(COL_FECHA),
        "Concepto": concepto,
        "Tipo": tipo,
        "Letra": letra,
        "Punto de Venta": row.get(COL_PV),
        "Número Desde": row.get(COL_NRO_DESDE),
        "Número Hasta": row.get(COL_NRO_HASTA),
        "Tipo Doc. Emisor": 80,
        "Nro. Doc. Emisor": row.get(COL_CUIT_EMISOR),
        "Denominación Emisor": row.get(COL_NOM_EMISOR),
        "Condición Fiscal": "RI" if letra == "A" else "MT",
        "Tipo Cambio": row.get(COL_TC),
        "Moneda": row.get(COL_MON),
    }

    # Exento / No gravado, otros tributos y total, con signo correcto
    exng_val = s(get_num(row, COL_NETO_NG) + get_num(row, COL_EXENTAS))
    otros_val = s(get_num(row, COL_OTROS))
    total_val = s(get_num(row, COL_TOTAL))

    filas_comp = []

    # Alícuotas consideradas: 10,5% / 21% / 27%
    aliquotas = [
        ("10.500", COL_NETO_105, COL_IVA_105),
        ("21.000", COL_NETO_21, COL_IVA_21),
        ("27.000", COL_NETO_27, COL_IVA_27),
    ]

    for aliq_txt, col_neto, col_iva in aliquotas:
        neto = s(get_num(row, col_neto))
        iva = s(get_num(row, col_iva))

        # Si no hay importe, no generamos fila para esa alícuota
        if neto == 0 and iva == 0:
            continue

        rec = base.copy()
        rec["Alicuota"] = aliq_txt
        rec["Neto"] = neto
        rec["IVA"] = iva
        rec["Ex/Ng"] = 0.0
        rec["Otros Conceptos"] = 0.0
        filas_comp.append(rec)

    # Asignar Ex/Ng y Otros en UNA sola fila si hay alícuotas
    if filas_comp:
        if exng_val != 0 or otros_val != 0:
            filas_comp[0]["Ex/Ng"] = exng_val
            filas_comp[0]["Otros Conceptos"] = otros_val
    else:
        # Caso sin alícuotas:
        #   - si hay Ex/Ng u Otros: usamos esos valores
        #   - si no, pero hay Total (típico comprobante C), mandamos Total a Ex/Ng
        if exng_val != 0 or otros_val != 0 or total_val != 0:
            rec = base.copy()
            rec["Alicuota"] = "0.000"
            rec["Neto"] = 0.0
            rec["IVA"] = 0.0

            if exng_val != 0 or otros_val != 0:
                rec["Ex/Ng"] = exng_val
                rec["Otros Conceptos"] = otros_val
            else:
                # Comprobantes C con solo “Imp. Total”: todo el total a No Gravado/Ex
                rec["Ex/Ng"] = total_val
                rec["Otros Conceptos"] = 0.0

            filas_comp.append(rec)

    # Calcular total y acumular registros
    for rec in filas_comp:
        rec["Total"] = (
            float(rec["Neto"])
            + float(rec["IVA"])
            + float(rec["Ex/Ng"])
            + float(rec["Otros Conceptos"])
        )
        registros.append(rec)

if not registros:
    st.error("No se encontraron comprobantes con importes.")
    st.stop()

cols_salida = [
    "Fecha Emisión",
    "Fecha Recepción",
    "Concepto",
    "Tipo",
    "Letra",
    "Punto de Venta",
    "Número Desde",
    "Número Hasta",
    "Tipo Doc. Emisor",
    "Nro. Doc. Emisor",
    "Denominación Emisor",
    "Condición Fiscal",
    "Tipo Cambio",
    "Moneda",
    "Alicuota",
    "Neto",
    "IVA",
    "Ex/Ng",
    "Otros Conceptos",
    "Total",
]

salida = pd.DataFrame(registros)[cols_salida]

st.subheader("Vista previa de la salida")
st.dataframe(salida.head(50))

# --- GENERAR EXCEL PARA DESCARGAR ---

buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    salida.to_excel(writer, sheet_name="Salida", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Salida"]
    num_format = workbook.add_format({"num_format": "#,##0.00"})

    # Columnas de importes
    col_idx = {name: i for i, name in enumerate(salida.columns)}
    for nombre in ["Neto", "IVA", "Ex/Ng", "Otros Conceptos", "Total"]:
        j = col_idx[nombre]
        worksheet.set_column(j, j, 15, num_format)

buffer.seek(0)

st.download_button(
    "📥 Descargar Excel procesado",
    data=buffer,
    file_name="Recibidos_salida.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)

# --- Footer ---
st.markdown(
    "<br><hr style='opacity:0.3'><div style='text-align:center; "
    "font-size:12px; color:#6b7280;'>"
    "© AIE – Herramienta para uso interno | Developer Alfonso Alderete"
    "</div>",
    unsafe_allow_html=True,
)
