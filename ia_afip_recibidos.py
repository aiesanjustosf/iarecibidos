# ia_afip_recibidos.py
# Conversión de AFIP "Recibidos" -> Formato Holistor
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="AFIP Recibidos → Formato Holistor",
    layout="centered"
)

st.title("AFIP Recibidos → Formato Holistor")

st.write(
    "Subí el Excel original descargado de AFIP (Libro IVA Digital - Compras/Recibidos) "
    "y descargá un archivo listo para importar en Holistor."
)

uploaded = st.file_uploader(
    "Subí el archivo de AFIP (.xlsx)",
    type=["xlsx"]
)


def map_tipo_letra(concepto: str):
    """Devuelve (Tipo, Letra) según el texto 'Tipo' de AFIP."""
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

# --- LECTURA DEL EXCEL DE AFIP ---

df = pd.read_excel(uploaded, sheet_name=0)

# Nombres de columnas según el archivo de AFIP
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

registros = []

for _, row in df.iterrows():
    concepto = str(row.get(COL_TIPO_AFIP, "")).strip()
    if not concepto:
        continue

    tipo, letra = map_tipo_letra(concepto)

    # Signo según si es Nota de Crédito
    sign = -1 if "Nota de Crédito" in concepto else 1

    # Campos base comunes a todas las filas de este comprobante
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

    # Exento / No gravado y otros tributos
    exng_val = float(row.get(COL_NETO_NG, 0) or 0) + float(row.get(COL_EXENTAS, 0) or 0)
    otros_val = float(row.get(COL_OTROS, 0) or 0)

    filas_comp = []

    # Alícuotas consideradas: 10,5% / 21% / 27%
    aliquotas = [
        ("10.500", COL_NETO_105, COL_IVA_105),
        ("21.000", COL_NETO_21, COL_IVA_21),
        ("27.000", COL_NETO_27, COL_IVA_27),
    ]

    for aliq_txt, col_neto, col_iva in aliquotas:
        neto = float(row.get(col_neto, 0) or 0)
        iva = float(row.get(col_iva, 0) or 0)

        # Si no hay importe, no generamos fila para esa alícuota
        if neto == 0 and iva == 0:
            continue

        rec = base.copy()
        rec["Alicuota"] = aliq_txt
        rec["Neto"] = sign * neto
        rec["IVA"] = sign * iva
        rec["Ex/Ng"] = 0.0
        rec["Otros Conceptos"] = 0.0
        filas_comp.append(rec)

    # Asignar Ex/Ng y Otros en UNA sola fila
    if filas_comp:
        if exng_val != 0 or otros_val != 0:
            filas_comp[0]["Ex/Ng"] = sign * exng_val
            filas_comp[0]["Otros Conceptos"] = sign * otros_val
    else:
        # Sin alícuotas pero con Ex/Ng u Otros: una única fila 0.000
        if exng_val != 0 or otros_val != 0:
            rec = base.copy()
            rec["Alicuota"] = "0.000"
            rec["Neto"] = 0.0
            rec["IVA"] = 0.0
            rec["Ex/Ng"] = sign * exng_val
            rec["Otros Conceptos"] = sign * otros_val
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

# --- ARMAR DATAFRAME DE SALIDA ---

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

    # Formato de números con miles y 2 decimales
    workbook = writer.book
    worksheet = writer.sheets["Salida"]
    num_format = workbook.add_format({"num_format": "#,##0.00"})

    # Columnas de importes: Neto, IVA, Ex/Ng, Otros, Total
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
