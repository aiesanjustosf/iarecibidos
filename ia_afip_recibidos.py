# ia_arca_recibidos.py
# Conversión de ARCA "Recibidos" -> Formato Holistor
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

# --- Rutas de assets ---
HERE = Path(__file__).parent
LOGO = HERE / "logo_aie.png"
FAVICON = HERE / "favicon_aie.ico"

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
    "Subí el Excel original descargado de **ARCA** "
    "(Libro IVA Digital - Compras/Recibidos) y descargá un archivo "
    "listo para importar en **Holistor**."
)

uploaded = st.file_uploader(
    "Subí el archivo de ARCA (.xlsx)",
    type=["xlsx"],
)


def map_tipo_letra(concepto: str):
    """Devuelve (Tipo, Letra) según el texto 'Tipo' de ARCA."""
    concepto = str(concepto).strip()

    # AJUSTE: caso especial pedido
    # 81 - Tique Factura A => Tipo T / Letra A
    if concepto.startswith("81") and "Tique Factura A" in concepto:
        return "T", "A"

    # Tipo: F / ND / NC / R
    if "Nota de Crédito" in concepto:
        tipo = "NC"
    elif "Nota de Débito" in concepto:
        tipo = "ND"
    elif "Recibo" in concepto:
        tipo = "R"
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

# USD
COL_TC = "Tipo Cambio"
COL_MON = "Moneda"

COL_IVA_105 = "IVA 10,5%"
COL_NETO_105 = "Neto Grav. IVA 10,5%"
COL_IVA_21 = "IVA 21%"
COL_NETO_21 = "Neto Grav. IVA 21%"
COL_IVA_27 = "IVA 27%"
COL_NETO_27 = "Neto Grav. IVA 27%"

# Si hay monto acá, pasarlo como EXENTO en Ex/Ng
COL_NETO_0 = "Neto Grav. IVA 0%"

COL_NETO_NG = "Neto No Gravado"
COL_EXENTAS = "Op. Exentas"
COL_OTROS = "Otros Tributos"
COL_TOTAL = "Imp. Total"

# Si no existieran por variación de layout, fallback mínimo para no romper
if COL_TC not in df.columns and "Tipo de Cambio" in df.columns:
    COL_TC = "Tipo de Cambio"
if COL_MON not in df.columns and "Moneda" in df.columns:
    COL_MON = "Moneda"
if COL_NETO_0 not in df.columns and "Neto Grav. IVA 0 %" in df.columns:
    COL_NETO_0 = "Neto Grav. IVA 0 %"

# Asegurar columnas presentes
for c in [COL_TC, COL_MON, COL_NETO_0]:
    if c not in df.columns:
        df[c] = 0.0 if c != COL_MON else ""

registros = []


def get_num_raw(row, col) -> float:
    """Devuelve número limpio (NaN -> 0)."""
    v = row.get(col, 0)
    if pd.isna(v):
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0


for _, row in df.iterrows():
    concepto = str(row.get(COL_TIPO_AFIP, "")).strip()
    if not concepto:
        continue

    tipo, letra = map_tipo_letra(concepto)
    es_nc = "Nota de Crédito" in concepto
    es_tique_factura_a_81 = concepto.startswith("81") and "Tique Factura A" in concepto

    moneda = str(row.get(COL_MON, "") or "").strip().upper()
    tc = get_num_raw(row, COL_TC)

    # Signo correcto:
    # - NC: negativo
    # - resto: positivo
    def s(valor: float) -> float:
        if valor == 0:
            return 0.0
        return -abs(valor) if es_nc else abs(valor)

    # Monto con conversión USD -> ARS (si Moneda == USD)
    def get_num(row_, col_) -> float:
        v = get_num_raw(row_, col_)
        if moneda == "USD" and tc != 0:
            return v * tc
        return v

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
        # visibles en grilla/salida
        "Tipo Cambio": tc,
        "Moneda": moneda,
    }

    # Exento / No gravado:
    # Neto No Gravado + Op. Exentas + Neto Grav. IVA 0%
    exng_val = s(
        get_num(row, COL_NETO_NG)
        + get_num(row, COL_EXENTAS)
        + get_num(row, COL_NETO_0)
    )
    otros_val = s(get_num(row, COL_OTROS))
    total_val = s(get_num(row, COL_TOTAL))

    filas_comp = []

    # Alícuotas: 10,5% / 21% / 27%
    aliquotas = [
        (10.5, COL_NETO_105, COL_IVA_105),
        (21.0, COL_NETO_21, COL_IVA_21),
        (27.0, COL_NETO_27, COL_IVA_27),
    ]

    for aliq_val, col_neto, col_iva in aliquotas:
        neto = s(get_num(row, col_neto))
        iva = s(get_num(row, col_iva))

        if neto == 0 and iva == 0:
            continue

        rec = base.copy()
        rec["Alicuota"] = aliq_val
        rec["Neto"] = neto
        rec["IVA"] = iva
        rec["Ex/Ng"] = 0.0
        rec["Otros Conceptos"] = 0.0
        filas_comp.append(rec)

    # Ex/Ng y Otros: en una sola fila si hay alícuotas
    if filas_comp:
        if exng_val != 0 or otros_val != 0:
            filas_comp[0]["Ex/Ng"] = exng_val
            filas_comp[0]["Otros Conceptos"] = otros_val
    else:
        # Sin alícuotas:
        # - si hay Ex/Ng u Otros: usar esos
        # - si no, pero hay Total: mandar Total a Ex/Ng
        if exng_val != 0 or otros_val != 0 or total_val != 0:
            rec = base.copy()
            rec["Alicuota"] = 0.0
            rec["Neto"] = 0.0
            rec["IVA"] = 0.0

            if exng_val != 0 or otros_val != 0:
                rec["Ex/Ng"] = exng_val
                rec["Otros Conceptos"] = otros_val
            else:
                rec["Ex/Ng"] = total_val
                rec["Otros Conceptos"] = 0.0

            filas_comp.append(rec)

    # AJUSTE: para 81 - Tique Factura A, si la suma no coincide con el total original,
    # llevar la diferencia a Ex/Ng para que el comprobante salga con el mismo total de ARCA
    if es_tique_factura_a_81 and filas_comp:
        total_calculado = sum(
            float(r["Neto"]) + float(r["IVA"]) + float(r["Ex/Ng"]) + float(r["Otros Conceptos"])
            for r in filas_comp
        )
        diferencia = round(float(total_val) - float(total_calculado), 2)

        if abs(diferencia) >= 0.01:
            filas_comp[0]["Ex/Ng"] = float(filas_comp[0]["Ex/Ng"]) + diferencia

    # Total y acumulación
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

    money_format = workbook.add_format({"num_format": "#,##0.00"})
    col_idx = {name: i for i, name in enumerate(salida.columns)}

    # Importes
    for nombre in ["Neto", "IVA", "Ex/Ng", "Otros Conceptos", "Total"]:
        j = col_idx[nombre]
        worksheet.set_column(j, j, 15, money_format)

    # Tipo Cambio / Moneda visibles
    if "Tipo Cambio" in col_idx:
        j = col_idx["Tipo Cambio"]
        worksheet.set_column(j, j, 12, money_format)
    if "Moneda" in col_idx:
        j = col_idx["Moneda"]
        worksheet.set_column(j, j, 10)

    # Alicuota
    aliq_format = workbook.add_format({"num_format": "00.000"})
    j_aliq = col_idx["Alicuota"]
    worksheet.set_column(j_aliq, j_aliq, 8, aliq_format)

buffer.seek(0)

st.download_button(
    "📥 Descargar Excel procesado",
    data=buffer,
    file_name="Recibidos_salida.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# --- Footer ---
st.markdown(
    "<br><hr style='opacity:0.3'><div style='text-align:center; "
    "font-size:12px; color:#6b7280;'>"
    "© AIE – Herramienta para uso interno | Developer Alfonso Alderete"
    "</div>",
    unsafe_allow_html=True,
)
