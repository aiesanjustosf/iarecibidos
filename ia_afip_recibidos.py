# ia_arca_recibidos.py
# Conversión de ARCA "Recibidos" -> Formato Holistor
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path


# ============================================================
# RUTAS DE ASSETS
# ============================================================

HERE = Path(__file__).parent

LOGO = HERE / "logo_aie.png"
FAVICON = HERE / "favicon_aie.ico"


# ============================================================
# CONFIGURACIÓN DE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="ARCA Recibidos → Formato Holistor",
    page_icon=str(FAVICON) if FAVICON.exists() else None,
    layout="centered",
)


# ============================================================
# ENCABEZADO
# ============================================================

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


# ============================================================
# FUNCIONES DE COMPROBANTES
# ============================================================

def get_codigo_arca(concepto: str) -> str:
    """
    Obtiene el código numérico del comprobante ARCA.

    Ejemplos:
    '051 - Factura M' -> '51'
    '52 - Nota de Débito M' -> '52'
    '063 - Liquidación A' -> '63'
    """
    concepto = str(concepto).strip()

    if not concepto:
        return ""

    codigo = concepto.split("-")[0].strip()

    # Quitar ceros a la izquierda.
    # Ej: 051 -> 51
    codigo = codigo.lstrip("0")

    return codigo if codigo else "0"


def map_tipo_letra(concepto: str):
    """
    Devuelve (Tipo, Letra) según el comprobante de ARCA,
    adaptado al formato esperado por Holistor.
    """

    concepto = str(concepto).strip()
    codigo = get_codigo_arca(concepto)

    # --------------------------------------------------------
    # NUEVOS COMPROBANTES
    # --------------------------------------------------------

    # 051 - Factura M
    if codigo == "51":
        return "F", "M"

    # 052 - Nota de Débito M
    # En Holistor todas las Notas de Débito usan Tipo D
    # independientemente de la letra.
    if codigo == "52":
        return "D", "M"

    # 053 - Nota de Crédito M
    # En Holistor todas las Notas de Crédito usan Tipo C
    # independientemente de la letra.
    if codigo == "53":
        return "C", "M"

    # 063 - Liquidación A
    # En Holistor se vincula con comprobante LB
    if codigo == "63":
        return "LB", "A"

    # --------------------------------------------------------
    # CASO ESPECIAL EXISTENTE
    # --------------------------------------------------------

    # 81 - Tique Factura A
    if codigo == "81" and "Tique Factura A" in concepto:
        return "T", "A"

    # --------------------------------------------------------
    # TIPOS GENERALES HOLISTOR
    # --------------------------------------------------------
    #
    # IMPORTANTE:
    # La letra NO cambia el Tipo de comprobante.
    #
    # Nota de Crédito = C
    # Nota de Débito = D
    # Factura = F
    # Recibo = R
    # --------------------------------------------------------

    if "Nota de Crédito" in concepto:
        tipo = "C"

    elif "Nota de Débito" in concepto:
        tipo = "D"

    elif "Recibo" in concepto:
        tipo = "R"

    elif "Factura" in concepto:
        tipo = "F"

    else:
        tipo = ""

    # --------------------------------------------------------
    # LETRA
    # --------------------------------------------------------

    # Caso especial existente:
    # 8 - Nota de Crédito C => Holistor letra B
    if codigo == "8":
        letra = "B"

    else:
        # En general la letra es el último carácter:
        # A / B / C / M
        letra = concepto[-1] if concepto else ""

    return tipo, letra


# ============================================================
# DETENER SI TODAVÍA NO SE SUBIÓ ARCHIVO
# ============================================================

if uploaded is None:
    st.stop()


# ============================================================
# LECTURA DEL EXCEL DE ARCA
# ============================================================

# header=1 porque la fila 2 del archivo tiene
# los encabezados reales.

df = pd.read_excel(
    uploaded,
    sheet_name=0,
    header=1,
)


# ============================================================
# NOMBRES DE COLUMNAS SEGÚN ARCA
# ============================================================

COL_FECHA = "Fecha"
COL_TIPO_AFIP = "Tipo"

COL_PV = "Punto de Venta"
COL_NRO_DESDE = "Número Desde"
COL_NRO_HASTA = "Número Hasta"

COL_CUIT_EMISOR = "Nro. Doc. Emisor"
COL_NOM_EMISOR = "Denominación Emisor"
COL_COD_AUT = "Cód. Autorización"


# ------------------------------------------------------------
# MONEDA / TIPO DE CAMBIO
# ------------------------------------------------------------

COL_TC = "Tipo Cambio"
COL_MON = "Moneda"


# ------------------------------------------------------------
# IVA
# ------------------------------------------------------------

COL_IVA_105 = "IVA 10,5%"
COL_NETO_105 = "Neto Grav. IVA 10,5%"

COL_IVA_21 = "IVA 21%"
COL_NETO_21 = "Neto Grav. IVA 21%"

COL_IVA_27 = "IVA 27%"
COL_NETO_27 = "Neto Grav. IVA 27%"


# Si hay monto acá, pasarlo como EXENTO en Ex/Ng

COL_NETO_0 = "Neto Grav. IVA 0%"


# ------------------------------------------------------------
# OTROS IMPORTES
# ------------------------------------------------------------

COL_NETO_NG = "Neto No Gravado"
COL_EXENTAS = "Op. Exentas"
COL_OTROS = "Otros Tributos"
COL_TOTAL = "Imp. Total"


# ============================================================
# FALLBACKS POR POSIBLES CAMBIOS DE NOMBRE EN ARCA
# ============================================================

if COL_TC not in df.columns and "Tipo de Cambio" in df.columns:
    COL_TC = "Tipo de Cambio"

if COL_MON not in df.columns and "Moneda" in df.columns:
    COL_MON = "Moneda"

if COL_NETO_0 not in df.columns and "Neto Grav. IVA 0 %" in df.columns:
    COL_NETO_0 = "Neto Grav. IVA 0 %"

# Posibles variantes del encabezado de autorización en ARCA
if COL_COD_AUT not in df.columns:
    for alternativa in [
        "Cod. Autorización",
        "Código Autorización",
        "Código de Autorización",
        "Cod. Autorizacion",
        "Código Autorizacion",
        "Código de Autorizacion",
    ]:
        if alternativa in df.columns:
            COL_COD_AUT = alternativa
            break


# ============================================================
# ASEGURAR COLUMNAS
# ============================================================

for c in [
    COL_TC,
    COL_MON,
    COL_NETO_0,
    COL_COD_AUT,
]:
    if c not in df.columns:

        if c in [COL_MON, COL_COD_AUT]:
            df[c] = ""

        else:
            df[c] = 0.0


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def get_num_raw(row, col) -> float:
    """
    Devuelve un número limpio.
    NaN / vacío / error -> 0
    """

    v = row.get(col, 0)

    if pd.isna(v):
        return 0.0

    try:
        return float(v)

    except Exception:
        return 0.0


# ============================================================
# PROCESAMIENTO
# ============================================================

registros = []


for _, row in df.iterrows():

    concepto = str(
        row.get(COL_TIPO_AFIP, "")
    ).strip()

    if not concepto:
        continue

    codigo_arca = get_codigo_arca(concepto)

    tipo, letra = map_tipo_letra(concepto)


    # ========================================================
    # NOTAS DE CRÉDITO
    # ========================================================
    #
    # Todas las Notas de Crédito deben RESTAR.
    #
    # Incluye:
    # - NC A
    # - NC B
    # - NC C
    # - 053 Nota de Crédito M
    #
    # ========================================================

    es_nc = (
        codigo_arca == "53"
        or "Nota de Crédito" in concepto
    )


    # ========================================================
    # COMPROBANTES CON CONTROL ESPECIAL DE TOTAL
    # 6  - FACTURA B
    # 7  - NOTA DE DÉBITO B
    # 81 - TIQUE FACTURA A
    # 82 - TIQUE FACTURA B
    # ========================================================

    es_factura_b_6 = (
        codigo_arca == "6"
        and "Factura B" in concepto
    )

    es_nota_debito_b_7 = (
        codigo_arca == "7"
        and "Nota de Débito" in concepto
        and concepto.endswith("B")
    )

    es_tique_factura_a_81 = (
        codigo_arca == "81"
        and "Tique Factura A" in concepto
    )

    es_tique_factura_b_82 = (
        codigo_arca == "82"
        and "Tique Factura B" in concepto
    )

    es_comprobante_ajustable = (
        es_factura_b_6
        or es_nota_debito_b_7
        or es_tique_factura_a_81
        or es_tique_factura_b_82
    )


    # ========================================================
    # MONEDA
    # ========================================================

    moneda = str(
        row.get(COL_MON, "") or ""
    ).strip().upper()

    tc = get_num_raw(
        row,
        COL_TC,
    )


    # ========================================================
    # SIGNO
    # ========================================================

    def s(valor: float) -> float:
        """
        Nota de Crédito -> negativo.
        Resto -> positivo.
        """

        if valor == 0:
            return 0.0

        if es_nc:
            return -abs(valor)

        return abs(valor)


    # ========================================================
    # CONVERSIÓN MONEDA
    # ========================================================

    def get_num(row_, col_) -> float:
        """
        Obtiene el importe y, si la moneda es USD,
        lo convierte a pesos utilizando Tipo Cambio.
        """

        v = get_num_raw(
            row_,
            col_,
        )

        if moneda == "USD" and tc != 0:
            return v * tc

        return v


    # ========================================================
    # BASE COMÚN DEL COMPROBANTE
    # ========================================================

    base = {

        "Fecha Emisión":
            row.get(COL_FECHA),

        "Fecha Recepción":
            row.get(COL_FECHA),

        "Concepto":
            concepto,

        "Tipo":
            tipo,

        "Letra":
            letra,

        "Punto de Venta":
            row.get(COL_PV),

        "Número Desde":
            row.get(COL_NRO_DESDE),

        "Número Hasta":
            row.get(COL_NRO_HASTA),

        "Cód. Autorización":
            row.get(COL_COD_AUT),

        "Tipo Doc. Emisor":
            80,

        "Nro. Doc. Emisor":
            row.get(COL_CUIT_EMISOR),

        "Denominación Emisor":
            row.get(COL_NOM_EMISOR),

        "Condición Fiscal":
            "RI" if letra == "A" else "MT",

        # Visibles en grilla / salida

        "Tipo Cambio":
            tc,

        "Moneda":
            moneda,

        # Se completa solo si el comprobante fue ajustado
        "Control IA":
            "",
    }


    # ========================================================
    # EXENTO / NO GRAVADO
    # ========================================================
    #
    # Neto No Gravado
    # + Operaciones Exentas
    # + Neto Gravado IVA 0%
    #
    # ========================================================

    exng_val = s(

        get_num(
            row,
            COL_NETO_NG,
        )

        +

        get_num(
            row,
            COL_EXENTAS,
        )

        +

        get_num(
            row,
            COL_NETO_0,
        )
    )


    # ========================================================
    # OTROS TRIBUTOS
    # ========================================================

    otros_val = s(
        get_num(
            row,
            COL_OTROS,
        )
    )


    # ========================================================
    # TOTAL ORIGINAL ARCA
    # ========================================================

    total_val = s(
        get_num(
            row,
            COL_TOTAL,
        )
    )


    # ========================================================
    # FILAS DEL COMPROBANTE
    # ========================================================

    filas_comp = []


    # ========================================================
    # ALÍCUOTAS
    # ========================================================

    aliquotas = [

        (
            10.5,
            COL_NETO_105,
            COL_IVA_105,
        ),

        (
            21.0,
            COL_NETO_21,
            COL_IVA_21,
        ),

        (
            27.0,
            COL_NETO_27,
            COL_IVA_27,
        ),
    ]


    for aliq_val, col_neto, col_iva in aliquotas:

        neto = s(
            get_num(
                row,
                col_neto,
            )
        )

        iva = s(
            get_num(
                row,
                col_iva,
            )
        )


        # Si no hay ni neto ni IVA para esa alícuota,
        # no generar fila.

        if neto == 0 and iva == 0:
            continue


        rec = base.copy()

        rec["Alicuota"] = aliq_val

        rec["Neto"] = neto

        rec["IVA"] = iva

        rec["Ex/Ng"] = 0.0

        rec["Otros Conceptos"] = 0.0


        filas_comp.append(rec)


    # ========================================================
    # EXENTO / NO GRAVADO / OTROS
    # ========================================================

    if filas_comp:

        # Si existen alícuotas,
        # Ex/Ng y Otros se agregan a la primera fila.

        if exng_val != 0 or otros_val != 0:

            filas_comp[0]["Ex/Ng"] = exng_val

            filas_comp[0]["Otros Conceptos"] = otros_val


    else:

        # ====================================================
        # SIN ALÍCUOTAS
        # ====================================================
        #
        # Si hay Ex/Ng u Otros -> utilizar esos importes.
        #
        # Si no hay nada discriminado pero existe Total,
        # mandar el Total completo a Ex/Ng.
        #
        # ====================================================

        if (
            exng_val != 0
            or otros_val != 0
            or total_val != 0
        ):

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


    # ========================================================
    # AJUSTE ESPECIAL DE TOTAL
    # 6  - FACTURA B
    # 7  - NOTA DE DÉBITO B
    # 81 - TIQUE FACTURA A
    # 82 - TIQUE FACTURA B
    # ========================================================
    #
    # Si la suma discriminada no coincide con el total
    # original de ARCA, enviar la diferencia a Ex/Ng.
    #
    # IMPORTANTE:
    # Este ajuste automático se aplica SOLAMENTE a los
    # códigos 6, 7, 81 y 82. El resto de los comprobantes mantiene
    # su tratamiento habitual sin correcciones automáticas.
    #
    # Cuando se aplica una corrección, se marca el comprobante
    # con la leyenda "AJUSTADO POR IA - CORROBORAR".
    # ========================================================

    if es_comprobante_ajustable and filas_comp:

        total_calculado = sum(

            float(r["Neto"])
            + float(r["IVA"])
            + float(r["Ex/Ng"])
            + float(r["Otros Conceptos"])

            for r in filas_comp
        )

        diferencia = round(
            float(total_val)
            - float(total_calculado),
            2,
        )

        if abs(diferencia) >= 0.01:

            filas_comp[0]["Ex/Ng"] = (
                float(filas_comp[0]["Ex/Ng"])
                + diferencia
            )

            for r in filas_comp:
                r["Control IA"] = (
                    "AJUSTADO POR IA - CORROBORAR"
                )


    # ========================================================
    # TOTAL POR FILA
    # ========================================================

    for rec in filas_comp:

        rec["Total"] = (

            float(rec["Neto"])
            + float(rec["IVA"])
            + float(rec["Ex/Ng"])
            + float(rec["Otros Conceptos"])

        )

        registros.append(rec)


# ============================================================
# VALIDACIÓN
# ============================================================

if not registros:

    st.error(
        "No se encontraron comprobantes con importes."
    )

    st.stop()


# ============================================================
# COLUMNAS DE SALIDA
# ============================================================

cols_salida = [

    "Fecha Emisión",

    "Fecha Recepción",

    "Concepto",

    "Tipo",

    "Letra",

    "Punto de Venta",

    "Número Desde",

    "Número Hasta",

    "Cód. Autorización",

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

    "Control IA",
]


salida = pd.DataFrame(
    registros
)[cols_salida]


# ============================================================
# VISTA PREVIA
# ============================================================

st.subheader(
    "Vista previa de la salida"
)

st.dataframe(
    salida.head(50),
    use_container_width=True,
)


# ============================================================
# GENERAR EXCEL PARA DESCARGAR
# ============================================================

buffer = BytesIO()


with pd.ExcelWriter(
    buffer,
    engine="xlsxwriter",
) as writer:

    salida.to_excel(
        writer,
        sheet_name="Salida",
        index=False,
    )


    workbook = writer.book

    worksheet = writer.sheets["Salida"]


    # ========================================================
    # FORMATOS
    # ========================================================

    money_format = workbook.add_format(
        {
            "num_format": "#,##0.00"
        }
    )


    col_idx = {
        name: i
        for i, name in enumerate(
            salida.columns
        )
    }


    # ========================================================
    # IMPORTES
    # ========================================================

    for nombre in [

        "Neto",

        "IVA",

        "Ex/Ng",

        "Otros Conceptos",

        "Total",

    ]:

        j = col_idx[nombre]

        worksheet.set_column(
            j,
            j,
            15,
            money_format,
        )


    # ========================================================
    # TIPO DE CAMBIO
    # ========================================================

    if "Tipo Cambio" in col_idx:

        j = col_idx[
            "Tipo Cambio"
        ]

        worksheet.set_column(
            j,
            j,
            12,
            money_format,
        )


    # ========================================================
    # MONEDA
    # ========================================================

    if "Moneda" in col_idx:

        j = col_idx[
            "Moneda"
        ]

        worksheet.set_column(
            j,
            j,
            10,
        )


    # ========================================================
    # CÓDIGO DE AUTORIZACIÓN / CONTROL IA
    # ========================================================

    if "Cód. Autorización" in col_idx:
        j = col_idx["Cód. Autorización"]
        worksheet.set_column(j, j, 20)

    if "Control IA" in col_idx:
        j = col_idx["Control IA"]
        worksheet.set_column(j, j, 32)


    # ========================================================
    # ALÍCUOTA
    # ========================================================

    aliq_format = workbook.add_format(
        {
            "num_format": "00.000"
        }
    )


    j_aliq = col_idx[
        "Alicuota"
    ]

    worksheet.set_column(
        j_aliq,
        j_aliq,
        8,
        aliq_format,
    )


# Volver al comienzo del archivo generado

buffer.seek(0)


# ============================================================
# DESCARGA
# ============================================================

st.download_button(

    "📥 Descargar Excel procesado",

    data=buffer,

    file_name="Recibidos_salida.xlsx",

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    "© AIE – Herramienta para uso interno | "
    "Developer Alfonso Alderete"
)
