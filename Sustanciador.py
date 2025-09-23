# app.py — Sustanciador Judicial (Streamlit Cloud ready)
# Autor: Andrés Cruz
# Descripción: Genera documentos Word (.docx) basados en modelos y una base Excel.
# Subetapas: Mandamiento de Pago, Corrección de Sentencia, Calificación de Demanda, Liquidación de Crédito.
# Salida: .docx (idéntico al modelo). El banco puede convertir a PDF externamente.

import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document

# ==========================
# Configuración inicial UI
# ==========================
st.set_page_config(page_title="⚖️ Sustanciador Judicial — COS JudicIA", layout="wide")
st.title("⚖️ Sustanciador Judicial — COS JudicIA")
st.caption("Genera memoriales completos desde plantillas .docx con datos de Excel. Salida .docx 100% fiel al modelo.")

# ==========================
# Utilidades
# ==========================

def strip_accents(s: str) -> str:
    if not isinstance(s, str):
        return s
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def norm_text(s: str) -> str:
    """Minusculas + sin tildes para matching robusto."""
    return strip_accents(s or "").lower()


def parse_ddmmyyyy(s: str):
    """Devuelve datetime si s contiene una fecha dd/mm/yyyy; si no, None."""
    if not isinstance(s, str):
        return None
    m = re.search(r"\b(\d{2})[/-](\d{2})[/-](\d{4})\b", s)
    if not m:
        return None
    dd, mm, yyyy = m.groups()
    try:
        return datetime(int(yyyy), int(mm), int(dd))
    except Exception:
        return None


def format_fecha_dd_de_mm_de_yyyy(dt: datetime) -> str:
    if not dt:
        return ""
    return f"{dt.day:02d} de {dt.month:02d} de {dt.year}"


def sanitize_filename(s: str) -> str:
    s = strip_accents(s or "").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9._-]", "", s)
    return s

# ==========================
# Reglas de búsqueda de fechas en AF
# ==========================
MANDAMIENTO_KEYS = [
    "solicitud correccion del auto",
    "solicitud correccion del mandamiento de pago",
    "correccion del mandamiento de pago",
]

SENTENCIA_KEYS = [
    "solicitud correccion de sentencia",
    "correccion de sentencia",
]


def match_any(text: str, keys) -> bool:
    t = norm_text(text)
    return any(k in t for k in keys)


def extract_fecha_mas_reciente_AF(series_like, keys):
    """Dado un texto (o lista de textos) del Cuaderno Principal (AF),
    encuentra la fecha dd/mm/yyyy más reciente cuyo renglón contenga alguna key."""
    if series_like is None:
        return None
    # AF puede venir como una única celda con multilinea o varias filas. Aquí tratamos por fila/cadena.
    candidatos = []
    if isinstance(series_like, (list, tuple, pd.Series)):
        textos = [str(x) for x in series_like if pd.notna(x)]
    else:
        textos = [str(series_like)] if pd.notna(series_like) else []
    for t in textos:
        if match_any(t, keys):
            dt = parse_ddmmyyyy(t)
            if dt:
                candidatos.append(dt)
    if not candidatos:
        return None
    return max(candidatos)

# ==========================
# Motor de reemplazo en DOCX
# ==========================

def replace_paragraph_exact(doc: Document, starts_with: str, new_text: str):
    """Reemplaza el texto del primer párrafo cuyo texto empiece con starts_with (match exacto inicial)."""
    for p in doc.paragraphs:
        if p.text.strip().startswith(starts_with):
            p.text = new_text
            return True
    return False


def replace_line_contains(doc: Document, contains_text: str, new_text: str):
    """Reemplaza el texto del primer párrafo que contenga contains_text."""
    for p in doc.paragraphs:
        if contains_text in p.text:
            p.text = new_text
            return True
    return False


def replace_after_label(doc: Document, label: str, new_value: str):
    """Reemplaza el párrafo que comience con 'label' dejando 'label: {new_value}'."""
    label = label.rstrip(":")
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.upper().startswith(label.upper() + ":"):
            p.text = f"{label.upper()}: {new_value}"
            return True
    return False


def replace_date_pattern(doc: Document, anchor_contains: str, pattern_regex: str, new_date_str: str):
    """Busca el primer párrafo que contenga anchor_contains y reescribe la subcadena que haga match con pattern_regex por la nueva fecha.
    pattern_regex debe capturar toda la frase tipo 'el pasado XX de XXX de XXXX' o 'radicada el día XX de XXXX de XXXX'."""
    rx = re.compile(pattern_regex, flags=re.IGNORECASE)
    for p in doc.paragraphs:
        if anchor_contains.lower() in p.text.lower():
            new_text = rx.sub(new_date_str, p.text)
            if new_text != p.text:
                p.text = new_text
                return True
    return False


def doc_to_preview_text(doc: Document) -> str:
    parts = []
    for p in doc.paragraphs:
        parts.append(p.text)
    return "\n\n".join(parts)

# ==========================
# Mapeo de plantillas y lógica por subetapa
# ==========================
TEMPLATES = {
    "Mandamiento": "MODELO IMPULSO PROCESAL CORRECCIÓN MP.docx",
    "Sentencia": "MODELO IMPULSO PROCESAL CORRECCIÓN SENTENCIA.docx",
    "Calificacion": "MODELO IMPULSO CALIFICACION DE DEMANDA.docx",
    "Liquidacion de credito": "MODELO IMPULSO LIQUIDACION DE CREDITO.docx",
}


SUBETAPAS = list(TEMPLATES.keys())

# Patrones de fecha para reemplazo en el cuerpo
# Mandamiento y Sentencia: 'el pasado XX de XXX de XXXX'
PATTERN_PASADO = r"el pasado\s+\S+\s+de\s+\S+\s+de\s+\S+"
# Calificación: 'radicada el día XX de XXXX de XXXX'
PATTERN_RADICADA = r"radicada\s+el\s+d[ií]a\s+\S+\s+de\s+\S+\s+de\s+\S+"

# ==========================
# UI — Carga de Excel y selección
# ==========================
st.subheader("1) Carga la base en Excel")
excel_file = st.file_uploader("Sube el archivo Excel con la precarga", type=["xlsx", "xls"])

if excel_file:
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"No se pudo leer el Excel: {e}")
        st.stop()

    st.success(f"Base cargada: {df.shape[0]} filas, {df.shape[1]} columnas")
    with st.expander("Ver primeras filas"):
        st.dataframe(df.head(20))

    # Validación columnas mínimas
    required_cols = {
        'A': None,  # CC (columna A real puede llamarse distinto; aceptamos alias)
        'B': None,  # Nombre
        'H': None,  # Juzgado
        'J': None,  # Correo
        'I': None,  # Radicado
    }
    # Intento mapear por nombres conocidos
    # Permitimos alias comunes
    aliases = {
        'A': ['CC', 'Cedula', 'Cédula', 'Documento', 'Identificacion', 'Identificación'],
        'B': ['NombreTitular', 'Nombre', 'Demandado', 'DemandadoNombre'],
        'H': ['Juzgado'],
        'J': ['Correo', 'CorreoJuzgado', 'EmailJuzgado'],
        'I': ['Radicado', 'RADICADO'],
        'O': ['FECHA DE PRESENTACIÓN DDA', 'Fecha de presentacion dda', 'FechaPresentacionDDA'],
        'AF': ['Cuaderno Principal', 'CUADERNO PRINCIPAL', 'Cuaderno_Principal'],
    }

    col_map = {}
    lower_cols = {norm_text(c): c for c in df.columns}
    for key, names in aliases.items():
        found = None
        for n in names:
            if norm_text(n) in lower_cols:
                found = lower_cols[norm_text(n)]
                break
        col_map[key] = found

    missing = [k for k in ['A','B','H','J','I'] if not col_map.get(k)]
    if missing:
        st.warning(f"Faltan columnas mínimas en tu base: {missing}. Puedes renombrar columnas o actualizar aliases en el código.")

    st.subheader("2) Selecciona la subetapa y el registro")
    subetapa = st.selectbox("Subetapa", SUBETAPAS, index=0)

    # Construimos una clave amigable para seleccionar la fila
    def build_key(row):
        cc = str(row.get(col_map.get('A'), ''))
        nombre = str(row.get(col_map.get('B'), ''))
        rad = str(row.get(col_map.get('I'), ''))
        return f"CC {cc} | {nombre} | RAD {rad}"

    opciones = df.apply(build_key, axis=1).tolist()
    sel = st.selectbox("Registro", opciones)
    idx = opciones.index(sel) if sel in opciones else None

    if idx is not None:
        row = df.iloc[idx]

        # Datos base
        cc = str(row.get(col_map.get('A'), '')).strip()
        nombre = str(row.get(col_map.get('B'), '')).strip()
        juzgado = str(row.get(col_map.get('H'), '')).strip()
        correo = str(row.get(col_map.get('J'), '')).strip()
        radicado = str(row.get(col_map.get('I'), '')).strip()

        # Carga del modelo
        template_path = TEMPLATES[subetapa]
        try:
            doc = Document(template_path)
        except Exception as e:
            st.error(f"No se pudo abrir la plantilla '{template_path}': {e}")
            st.stop()

        # === Reemplazos comunes ===
        # 1) Encabezado: Juzgado y correo (líneas consecutivas)
        #   Reemplazamos la primera línea que comience con 'JUZGADO' y la primera que contenga '@'
        ok_juz = replace_line_contains(doc, "JUZGADO", juzgado)
        ok_correo = replace_line_contains(doc, "@", correo)

        # 2) RAD, DEMANDANTE, DEMANDADO
        replace_after_label(doc, "RAD", radicado)
        replace_after_label(doc, "DEMANDANTE", "BANCO GNB SUDAMERIS S.A")
        replace_after_label(doc, "DEMANDADO", f"CC {cc} {nombre}")

        # === Reemplazos específicos por subetapa ===
        fecha_str_final = None

        if subetapa == "Mandamiento":
            # Buscar fecha más reciente en AF por keywords Mandamiento
            af_col = col_map.get('AF')
            fecha_dt = None
            if af_col:
                fecha_dt = extract_fecha_mas_reciente_AF(row.get(af_col), MANDAMIENTO_KEYS)
            # Patrón 'el pasado XX de XXX de XXXX'
            if fecha_dt:
                fecha_str_final = format_fecha_dd_de_mm_de_yyyy(fecha_dt)
                replace_date_pattern(
                    doc,
                    anchor_contains="el pasado",
                    pattern_regex=PATTERN_PASADO,
                    new_date_str=f"el pasado {fecha_str_final}"
                )

        elif subetapa == "Sentencia":
            af_col = col_map.get('AF')
            fecha_dt = None
            if af_col:
                fecha_dt = extract_fecha_mas_reciente_AF(row.get(af_col), SENTENCIA_KEYS)
            if fecha_dt:
                fecha_str_final = format_fecha_dd_de_mm_de_yyyy(fecha_dt)
                replace_date_pattern(
                    doc,
                    anchor_contains="el pasado",
                    pattern_regex=PATTERN_PASADO,
                    new_date_str=f"el pasado {fecha_str_final}"
                )

        elif subetapa == "Calificacion":
            o_col = col_map.get('O')
            fecha_dt = None
            if o_col:
                raw = row.get(o_col)
                # Puede venir como datetime o str
                if isinstance(raw, datetime):
                    fecha_dt = raw
                else:
                    fecha_dt = parse_ddmmyyyy(str(raw))
            if fecha_dt:
                fecha_str_final = format_fecha_dd_de_mm_de_yyyy(fecha_dt)
                replace_date_pattern(
                    doc,
                    anchor_contains="radicada el",
                    pattern_regex=PATTERN_RADICADA,
                    new_date_str=f"radicada el día {fecha_str_final}"
                )

        elif subetapa == "Liquidacion de credito":
            # No hay fecha dinámica. Solo reemplazos comunes.
            pass

        # ==========================
        # Previsualización
        # ==========================
        st.subheader("3) Previsualización")
        preview = doc_to_preview_text(doc)
        st.text_area("Contenido (solo vista rápida)", value=preview, height=400)

        # ==========================
        # Descarga DOCX
        # ==========================
        st.subheader("4) Descargar documento (.docx)")
        nombre_sub = {
            "Mandamiento": "Mandamiento",
            "Sentencia": "Sentencia",
            "Calificacion": "Calificacion",
            "Liquidacion de credito": "Liquidacion_de_credito",
        }[subetapa]

        out_name = f"{sanitize_filename(cc)}_{sanitize_filename(nombre)}_{nombre_sub}.docx"
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.download_button(
            label=f"⬇️ Descargar {out_name}",
            data=bio.getvalue(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

else:
    st.info("Sube la base en Excel para continuar.")
