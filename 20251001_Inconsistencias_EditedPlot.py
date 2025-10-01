import pandas as pd
from datetime import datetime
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from spellchecker import SpellChecker
from unidecode import unidecode
from difflib import get_close_matches
from rapidfuzz import fuzz, process
import unicodedata

import win32com.client as win32

# Configurar diccionario en español
spell = SpellChecker(language="es")

# ==========================
# Selección archivo CSV
# ==========================

Tk().withdraw()
ruta = askopenfilename(
    filetypes=[("Archivos CSV", "*.csv")],
    title="Seleccione el archivo CSV"
)

if not ruta:
    print("❌ No se seleccionó ningún archivo. Saliendo...")
    raise SystemExit

# ==========================
# Columnas a analizar
# ==========================
columnas_objetivo = [
    "ID",
    "Nombre Proyecto",
    "Fecha Captura",
    "Código Interno",
    "Símbolo",
    "Nombre Predio Jurídico",
    "Escala",
    "Fuente Información",
    "Creado Por",
    "Fecha Última Actualización",
    "Modificado Por",
    "Comentarios",
    "Cód DANE Depto",
    "Cód DANE Mpio",
    "Año Vigencia Insumo Geográfico",
    "Nombre Vereda",
    "RULEID",
    "Código SIG Predio Jurídico",
    "Área Terreno Calculada Mts2",
    "Tipo de Propiedad"
]

# ==========================
# Formatos de fecha válidos
# ==========================
DATE_FORMATS = [
    "%Y-%m-%d",   # 2015-09-22
    "%d/%m/%Y",   # 22/09/2015
]

def parse_date_strict(s):
    """Devuelve datetime si coincide con un formato válido.
       Si contiene hora → 'HORA_ENCONTRADA'.
       Si no coincide → 'FORMATO_INVALIDO'.
    """
    if pd.isna(s):
        return pd.NaT
    s_clean = str(s).strip()
    if s_clean == "":
        return pd.NaT

    # 🚨 detectar si incluye hora
    if re.search(r"\d+:\d+", s_clean) or re.search(r"\b(AM|PM|am|pm)\b", s_clean):
        return "HORA_ENCONTRADA"

    # Intentar parsear con los formatos permitidos
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s_clean, fmt)
        except ValueError:
            continue

    # ❌ Ningún formato válido
    return "FORMATO_INVALIDO"

# ==========================
# Cargar CSV: todo como texto
# ==========================

df_raw = pd.read_csv(
    ruta,
    usecols=columnas_objetivo,
    encoding="utf-8",
    sep=";",
    dtype=str
)

# Copia para análisis
df = df_raw.copy()

# Normalizar: convertir espacios vacíos en <ESPACIO>
def limpiar_valor(x):
    if isinstance(x, str):
        if x.strip() == "" and x != "":
            return "<ESPACIO>"
        return x.strip()
    return x

df = df.applymap(limpiar_valor)
df = df.replace("", pd.NA)  # opcional: reemplazar strings vacíos por NaN

# Preprocesamientio de Columnas Fecha Captura y Año Vigencia.
def extraer_anio(texto):
    if pd.isna(texto):
        return None
    match = re.search(r"\d{4}", str(texto))
    return int(match.group()) if match else None

# Crear columnas adicionales en todo el DataFrame
df_raw["Anio_Captura"] = df_raw["Fecha Captura"].apply(extraer_anio)
df_raw["Anio_Vigencia_Num"] = df_raw["Año Vigencia Insumo Geográfico"].apply(extraer_anio)

registros = []

codigos_dane_deptos = {
    "05": "Antioquia",
    "08": "Atlántico",
    "11": "Bogotá, D.C.",
    "13": "Bolívar",
    "15": "Boyacá",
    "17": "Caldas",
    "18": "Caquetá",
    "19": "Cauca",
    "20": "Cesar",
    "23": "Córdoba",
    "25": "Cundinamarca",
    "27": "Chocó",
    "41": "Huila",
    "44": "La Guajira",
    "47": "Magdalena",
    "50": "Meta",
    "52": "Nariño",
    "54": "Norte de Santander",
    "63": "Quindío",
    "66": "Risaralda",
    "68": "Santander",
    "70": "Sucre",
    "73": "Tolima",
    "76": "Valle del Cauca",
    "81": "Arauca",
    "85": "Casanare",
    "86": "Putumayo",
    "88": "Archipiélago de San Andrés, Providencia y Santa Catalina",
    "91": "Amazonas",
    "94": "Guainía",
    "95": "Guaviare",
    "97": "Vaupés",
    "99": "Vichada"
}
# ==========================
# Construir reporte por columna
# ==========================

# 🔹 Validaciones columna por columna:
for idx, fila in df.iterrows():
    id_val = fila["ID"]

    # ---- Nombre Proyecto ----

    val_raw = df_raw.loc[idx, "Nombre Proyecto"]
    val = fila["Nombre Proyecto"]

    obs_nom_proyect = []  # 👉 aquí guardamos observaciones específicas de Nombre Proyecto

    if pd.isna(val_raw) or str(val_raw).strip() == "":  # 🚨 Vacío real
        obs = {
            "ID": id_val,
            "Columna Analizada": "Nombre Proyecto",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Fondo"
        }
        registros.append(obs)
        obs_nom_proyect.append(obs)

    elif val_raw == "<ESPACIO>":  # solo espacios
        obs = {
            "ID": id_val,
            "Columna Analizada": "Nombre Proyecto",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Nombre Proyecto",
            "Tipología": "Forma"
        }
        registros.append(obs)
        obs_nom_proyect.append(obs)

    elif isinstance(val_raw, str):
        observaciones = []

        # 🚨 Validación de espacios problemáticos
        if val_raw.startswith(" "):
            observaciones.append("Espacio al inicio")
        if val_raw.endswith(" "):
            observaciones.append("Espacio al final")
        if "  " in val_raw:
            observaciones.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            observaciones.append("Saltos de línea")

        # 🚨 Validación de estructura (alfanumérico + guion bajo, incluyendo acentos y ñ/ü)
        if not re.match(r"^[A-Za-z0-9_ÁÉÍÓÚÜÑáéíóúüñ]+$", val_raw):
            observaciones.append("Estructura no cumple con el Diccionario de Datos")
        else:
            partes = val_raw.split("_")

            if len(partes) != 2:
                observaciones.append("Estructura no cumple con el Diccionario de Datos")
            else:
                negocio, proyecto = partes
                siglas_negocio = ["SIS", "VEX", "VAS", "VRC", "VRS", "VRO", "OXY", "VFS", "VPI"]
                if negocio not in siglas_negocio:
                    observaciones.append("La sigla del Negocio no se encuentra de acuerdo con el Diccionario de Datos")

        # 🚨 Si hubo observaciones, se registran todas juntas
        if observaciones:
            obs = {
                "ID": id_val,
                "Columna Analizada": "Nombre Proyecto",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(observaciones),
                "Tipología": "Forma"
            }
            registros.append(obs)
            obs_nom_proyect.append(obs)
    # Guardar las observaciones de Nombre Proyecto en el registro
    fila["Obs_Nom_Proyect"] = "; ".join(o["Observación Específica"] for o in obs_nom_proyect) if obs_nom_proyect else ""

    # ---- Fecha Captura ----
    
    fecha_revision = datetime.today()  # 👈 se usa la fecha actual

    val_raw = df_raw.loc[idx, "Fecha Captura"]   # texto original
    parsed = parse_date_strict(val_raw)

    if pd.isna(parsed):  # vacío real
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Captura",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
            })

    elif parsed == "HORA_ENCONTRADA":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Captura",
            "Dato Analizado": str(val_raw),
            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
            "Observación Específica": "La fecha incluye hora (solo debería tener fecha)",
            "Tipología": "Forma"
        })

    elif parsed == "FORMATO_INVALIDO":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Captura",
            "Dato Analizado": str(val_raw),
            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
            "Observación Específica": "La fecha no corresponde al estándar esperado",
            "Tipología": "Forma"
        })

    elif val_raw == "<ESPACIO>":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Captura",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio",
            "Tipología": "Forma"
        })

    elif isinstance(parsed, datetime):
    # 🚨 Reglas de negocio específicas
        if parsed.strftime("%Y-%m-%d") in ["1900-01-01", "1900-12-12"]:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fecha Captura",
                "Dato Analizado": str(val_raw),
                "Observación General": "Inconsistencia Lógica del Dato",
                "Observación Específica": "Fecha Captura no válida para No Aplica y Sin Información",
                "Tipología": "Forma"
            })
        elif parsed < datetime(2009, 1, 1) or parsed > fecha_revision:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fecha Captura",
                "Dato Analizado": str(val_raw),
                "Observación General": "Inconsistencia Lógica del Dato",
                "Observación Específica": "Fechas no son consistentes de acuerdo a los Periodos de captura",
                "Tipología": "Forma"
            })  

    elif isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
        errores_espacios = []
        if val_raw.startswith(" "):
            errores_espacios.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores_espacios.append("Espacio al final")
        if "  " in val_raw:
            errores_espacios.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            errores_espacios.append("Saltos de línea")
    
        if errores_espacios:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fecha Captura",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(errores_espacios),
                "Tipología": "Forma"
        }) 

    # ---- Código Interno ----
    val_raw = df_raw.loc[idx, "Código Interno"]   # dato original sin modificar
    val = fila["Código Interno"]

    # Traer el Nombre Proyecto asociado
    nombre_proyecto_ref = df_raw.loc[idx, "Nombre Proyecto"]

    # 🚨 Caso 1: Vacíos / Totalidad → prioridad absoluta
    if pd.isna(val):
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Código Interno",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Código Interno",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Código Interno",
            "Tipología": "Forma"
        })

    # 🚨 Caso 2: Tiene valor → se acumulan validaciones
    else:
        errores_forma = set()   # usar set evita duplicados
        es_duplicado = False    # bandera para lógica

        # --- Validación de espacios indebidos
        if isinstance(val_raw, str):
            if val_raw.startswith(" "): errores_forma.add("Espacio al inicio")
            if val_raw.endswith(" "): errores_forma.add("Espacio al final")
            if "  " in val_raw: errores_forma.add("Múltiples espacios")
            if "\n" in val_raw or "\r" in val_raw: errores_forma.add("Saltos de línea")

        # --- Validación de duplicidad
        if (df_raw["Código Interno"] == val_raw).sum() > 1:
            es_duplicado = True
            errores_forma.add("Código Interno duplicado")

        # --- Validación de estructura: SIGLA_PROYECTO_PJXX
        partes = val_raw.split("_")
        if len(partes) != 3:
            errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")
        else:
            sigla, proyecto, pj_consec = partes

            # Validación sigla
            siglas_validas = ["SIS", "VEX", "VAS", "VRC", "VRS", "VRO", "OXY", "VFS", "VPI"]
            if sigla not in siglas_validas:
                errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")

            # Validación segunda parte (Proyecto)
            if "_" in proyecto or "-" in proyecto or " " in proyecto:
                errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")
            elif any(ch.isdigit() for ch in proyecto):
                errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")
            elif not re.match(r"^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+$", proyecto):
                errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")

            # Validación PJ + consecutivo
            if not pj_consec.startswith("PJ"):
                errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")
            else:
                consecutivo = pj_consec.replace("PJ", "")
                if not consecutivo.isdigit():
                    errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")
                else:
                    if len(consecutivo) not in [2, 3]:
                        errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")
                    if consecutivo == "00":
                        errores_forma.add("Código Interno no conserva la estructura definida en el Diccionario de Datos")

        # --- Validar que Nombre Proyecto esté contenido en Código Interno
        if isinstance(nombre_proyecto_ref, str) and nombre_proyecto_ref.strip():
            if nombre_proyecto_ref not in val_raw:
                errores_forma.add("Nombre Proyecto no está contenido en Código Interno")

        # --- Consolidar la observación final
        if errores_forma:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Código Interno",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Lógica del Dato" if es_duplicado else "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(sorted(errores_forma)),
                "Tipología": "Fondo" if es_duplicado else "Forma"
            })


    # ---- Símbolo ----

    val_raw = df_raw.loc[idx, "Símbolo"]
    val = fila["Símbolo"]
    
    if pd.isna(val) or str(val_raw).strip() == "": # Vacio
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Símbolo",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Símbolo",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Símbolo",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
        errores = []
        val_clean = val_raw.strip()

        if val_raw.startswith(" "):
            errores.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores.append("Espacio al final")
        if "  " in val_raw:
            errores.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            errores.append("Saltos de línea")

     # 🚨 Validación "No Aplica"

        if val_clean == "No Aplica":
            pass  # ✅ válido, no genera inconsistencia

        elif val_clean.lower() == "no aplica":
            errores.append("Estandarizar con formato tipo título")

        else:
            errores.append("Diligenciar No Aplica")    
    
        if errores:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Símbolo",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(errores),
                "Tipología": "Forma"
        }) 
                  
    # ---- Nombre Predio Jurídico ----

    val_raw = df_raw.loc[idx, "Nombre Predio Jurídico"]
    val = fila["Nombre Predio Jurídico"]

    if pd.isna(val):  # 🚨 Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Nombre Predio Jurídico",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # 🚨 Solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Nombre Predio Jurídico",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Nombre Predio Jurídico",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):
        errores = []
        val_clean = val_raw.strip()

        # 🚨 Saltos de línea / tabulaciones
        if "\n" in val_raw or "\r" in val_raw or "\t" in val_raw:
            errores.append("Saltos de línea o tabulación")

        # 🚨 Espacios problemáticos → ahora SÍ generan inconsistencia
        if val_raw.startswith(" "):
            errores.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores.append("Espacio al final")
        if "  " in val_raw:
            errores.append("Múltiples espacios")

        # --- Normalización para tokens ---
        val_norm = re.sub(r'^[\s\-\.,;:]+', '', val_clean)  # quitar puntuación inicial
        val_norm = re.sub(r'[;,:\.\-]+$', '', val_norm)     # quitar puntuación final
        val_norm = re.sub(r'\s+', ' ', val_norm).strip()

        tokens = [t.strip(" ,.") for t in val_norm.split(" ") if t.strip() != ""]

        # --- Patrones y listas permitidas ---
        allowed_lower = {
            "de","del","la","las","los","el","y","o","por","en","sin","predio","urbano",
            "vía","via","al","san","santa","corregimiento","vereda","sector","urbanización",
            "urbanizacion","barrio"
        }
        roman_pattern = re.compile(r'^(?:I|II|III|IV|V|VI|VII|VIII|IX|X)$', re.IGNORECASE)
        title_pattern = re.compile(r'^[A-ZÁÉÍÓÚÑÜ][a-záéíóúñü]+(?:-[A-ZÁÉÍÓÚÑÜ][a-záéíóúñü]+)*$')
        num_pattern = re.compile(r'^\d+[A-Z]?$')   # 13, 13A, 04
        single_upper = re.compile(r'^[A-Z]$')     # B, H, etc.

        invalid_tokens = []
        for tok in tokens:
            low = tok.lower()
            if low in allowed_lower:
                continue
            if roman_pattern.match(tok):
                continue
            if num_pattern.match(tok):
                continue
            if single_upper.match(tok):
                continue
            if title_pattern.match(tok):
                continue
            invalid_tokens.append(tok)

        # --- Resultado ---
        if errores or invalid_tokens:
            detalles = []
            if errores:
                detalles.append("; ".join(errores))
            if invalid_tokens:
                detalles.append("Token(es) inválido(s): " + ", ".join(invalid_tokens))
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Nombre Predio Jurídico",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(detalles),
                "Tipología": "Forma"
            })

    # ---- Escala ----

    val_raw = df_raw.loc[idx, "Escala"]
    val = fila["Escala"]

    if pd.isna(val):  # 🚨 Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Escala",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # 🚨 Solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Escala",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Escala",
            "Tipología": "Forma"
        })

    else:
        val_clean = str(val_raw).strip()

        # 🚨 Caso exacto válido
        if val_clean in {"10000", "25000"}:
            pass  # ✅ válido

        # 🚨 Caso 1:10000 o 1:25000 → error de forma
        elif val_clean in {"1:10000", "1:25000"}:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Escala",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "Solo debe diligenciarse el Número de la Escala",
                "Tipología": "Forma"
            })

        # 🚨 Otros casos que empiezan con 1: → error de forma + aclaración IGAC
        elif val_clean.startswith("1:"):
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Escala",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "Solo debe diligenciarse el Número de la Escala, Escala IGAC para predios rurales produce cartografía de 10000 y 25000",
                "Tipología": "Forma"
            })

        # 🚨 Si es texto no numérico → inconsistencia lógica especial
        elif not val_clean.isdigit():
               registros.append({
                "ID": id_val,
                "Columna Analizada": "Escala",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Lógica del Dato",
                "Observación Específica": "Dato no corresponde al valor de una escala",
                "Tipología": "Fondo"
            })

        # 🚨 Otros valores numéricos distintos a 10000/25000 → inconsistencia lógica
        else:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Escala",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Lógica del Dato",
                "Observación Específica": "Escala IGAC para predios rurales produce cartografía de 10000 y 25000",
                "Tipología": "Fondo"
            })
                                    
    # ---- Fuente Información ----
    val_raw = df_raw.loc[idx, "Fuente Información"]
    val = fila["Fuente Información"]

    dominios_permitidos = [
        "VIT - Transporte",
        "ECP - Seguridad Fisica",
        "IGAC",
        "IDEAM",
        "Ministerio de Ambiente",
        "Otra Fuente",
        "ECP - Suministro y Mercadeo",
        "DANE",
        "ECP - Inmobiliario",
        "ECP - Social",
        "ECP - Ambiental",
        "Ministerio de Interior y Justicia",
        "VAS - Asociados",
        "Diseños Obra Civil",
        "ECP - Refinacion y Petroquimica",
        "Informacion de Campo",
        "VEX - Exploracion",
        "VPR - Produccion",
        "P8 - Gestion Documental",
        "Depuracion Poligonos SIGDI",
        "Levantamiento Topografico",
        "Trabajo Campo (GPS)",
        "Poligono Google Earth",
        "Poligono IGAC",
        "ECP - Dato Fundamental"
    ]

    dominios_restringidos_predios = [
        "Diseños Obra Civil",
        "ECP - Dato Fundamental",
        "Poligono Google Earth",
        "VEX - Exploracion",
        "VPR - Produccion"
    ]

    if pd.isna(val):  # 🚨 Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fuente Información",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str) and val_raw.strip() == "":  # 🚨 Solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fuente Información",
            "Dato Analizado": val_raw,
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacios",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
        errores_espacios = []
        if val_raw.startswith(" "):
            errores_espacios.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores_espacios.append("Espacio al final")
        if "  " in val_raw:
            errores_espacios.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            errores_espacios.append("Saltos de línea")

        if errores_espacios:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fuente Información",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(errores_espacios),
                "Tipología": "Forma"
            })

        # 🚨 Validación de dominios permitidos
        if val_raw.strip() not in dominios_permitidos:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fuente Información",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estandar del Diccionario de Datos",
                "Observación Específica": "Valores no se encuentran en los dominios del diccionario de datos",
                "Tipología": "Forma"
            })

        # 🚨 Validación de dominios restringidos para predios
        if val_raw.strip() in dominios_restringidos_predios:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fuente Información",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Logica del Dato",
                "Observación Específica": "Dominio no es válido para captura de predios",
                "Tipología": "Fondo"
            })
                                                 
    # ---- Creado Por ----
    val_raw = df_raw.loc[idx, "Creado Por"]
    val = fila["Creado Por"]

    if pd.isna(val):  # Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Creado Por",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Creado Por",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Creado Por",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):

        # Normalización
        val_limpio = " ".join(val_raw.strip().split())
        val_lower = val_limpio.lower()

        # ------------------------------
        # 🚨 Casos especiales (registro aparte)
        # ------------------------------
        if (
            val_lower == "saneamiento p8 fase i"
            or val_lower == "levadata - saneamiento p8 fase i"
            or val_lower == "no aplica"
            or val_lower == "migracion lci"
            or val_lower in ["sin informacion", "sin información", "sin info"]
            or re.match(r"^c\d{6,}[a-zA-Z]?$", val_limpio.strip(), re.IGNORECASE)   # Códigos tipo C102627Q
            or re.match(r"^usuario con registro c\d{6,}[a-zA-Z]?$", val_limpio.strip(), re.IGNORECASE) # Usuario con registro C101848W
            or len(val_limpio.split()) == 1  # 👈 solo una palabra
        ):
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Creado Por",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Logica del Dato",
                "Observación Específica": "Capturar nombre completo, tener en cuenta que el dato entre el Property y EditPlot debe ser en creación el mismo y debe estar en formato tipo título",
                "Tipología": "Fondo"
            })

        # ------------------------------
        # 🚨 Otras validaciones (Formato, espacios, estandarización)
        # ------------------------------
        else:
            observaciones = []

            # Validación Formato Título (respetando tildes)
            if val_limpio != val_limpio.title():
                observaciones.append("Errores en Formato")

            # Detección de variantes similares para estandarización
            val_normalizada = unidecode(val_limpio).title()
            nombres_existentes = df_raw["Creado Por"].dropna().unique()
            coincidencias = [n for n in nombres_existentes if unidecode(str(n)).title() == val_normalizada]

            if len(coincidencias) > 1:
                observaciones.append("Estandarizar Nombre a un solo registro")

            # Validación de espacios
            if val_raw.startswith(" "):
                observaciones.append("Espacio al inicio")
            if val_raw.endswith(" "):
                observaciones.append("Espacio al final")
            if "  " in val_raw:
                observaciones.append("Múltiples espacios")
            if "\n" in val_raw or "\r" in val_raw:
                observaciones.append("Saltos de línea")

            # Consolidar observaciones en un registro
            if observaciones:
                registros.append({
                    "ID": id_val,
                    "Columna Analizada": "Creado Por",
                    "Dato Analizado": val_raw,
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "; ".join(observaciones),
                    "Tipología": "Forma"
                })

    # ---- Fecha Última Actualización ----
     
    from datetime import datetime
    fecha_revision = datetime.today()  # 👈 se usa la fecha actual

    val_raw = df_raw.loc[idx, "Fecha Última Actualización"]   # texto original
    parsed = parse_date_strict(val_raw)

    if pd.isna(parsed):  # vacío real
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Última Actualización",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
            })

    elif parsed == "HORA_ENCONTRADA":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Última Actualización",
            "Dato Analizado": str(val_raw),
            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
            "Observación Específica": "La fecha incluye hora (solo debería tener fecha)",
            "Tipología": "Forma"
        })

    elif parsed == "FORMATO_INVALIDO":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Última Actualización",
            "Dato Analizado": str(val_raw),
            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
            "Observación Específica": "La fecha no corresponde al estándar esperado (solo %Y-%m-%d o %d/%m/%Y)",
            "Tipología": "Forma"
        })

    elif val_raw == "<ESPACIO>":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Última Actualización",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio",
            "Tipología": "Forma"
        })

    elif isinstance(parsed, datetime):
    # 🚨 Reglas de negocio específicas
       
        fecha_str = parsed.strftime("%Y-%m-%d") 
        
        # Caso especial: 1900-01-01 → válido
        if fecha_str == "1900-01-01":
            pass  # No hacer nada, se considera válido

        # Caso especial: 1900-12-12 → inconsistencia, sugerir estandarizar
        elif fecha_str == "1900-12-12":
            registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Última Actualización",
            "Dato Analizado": str(val_raw),
            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
            "Observación Específica": "Estandarizar a 1900-01-01",
            "Tipología": "Forma"
        })
        elif parsed < datetime(2009, 1, 1) or parsed > fecha_revision:
            registros.append({
            "ID": id_val,
            "Columna Analizada": "Fecha Última Actualización",
            "Dato Analizado": str(val_raw),
            "Observación General": "Inconsistencia Lógica del Dato",
            "Observación Específica": "Fechas no son consistentes de acuerdo a los Periodos de captura",
            "Tipología": "Forma"
        })
      
    elif isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
        errores_espacios = []
        if val_raw.startswith(" "):
            errores_espacios.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores_espacios.append("Espacio al final")
        if "  " in val_raw:
            errores_espacios.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            errores_espacios.append("Saltos de línea")
    
        if errores_espacios:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Fecha Última Actualización",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(errores_espacios),
                "Tipología": "Forma"
        }) 
            
    # ---- Modificado Por ----
    val_raw = df_raw.loc[idx, "Modificado Por"]
    val = fila["Modificado Por"]

    if pd.isna(val):  # Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Modificado Por",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Modificado Por",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Modificado Por",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):

        # Normalización
        val_limpio = " ".join(val_raw.strip().split())

        # ------------------------------
        # ✅ Excepción: permitido solo "No Aplica" (tipo título exacto)
        # ------------------------------
        if val_limpio == "No Aplica":
            pass  # Se acepta, no genera inconsistencia

        # ------------------------------
        # 🚨 Variantes incorrectas de "No Aplica"
        # ------------------------------
        elif val_limpio.lower() == "no aplica" and val_limpio != "No Aplica":
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Modificado Por",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "Errores en Formato",
                "Tipología": "Forma"
            })

        # ------------------------------
        # 🚨 Casos especiales (registro aparte)
        # ------------------------------
        elif (
            val_limpio.lower() == "saneamiento p8 fase i"
            or val_limpio.lower() == "levadata - saneamiento p8 fase i"
            or val_limpio.lower() == "migracion lci"
            or val_limpio.lower() in ["sin informacion", "sin información", "sin info"]
            or re.match(r"^c\d{6,}[a-zA-Z]?$", val_limpio.strip(), re.IGNORECASE)   # Códigos tipo C102627Q
            or re.match(r"^usuario con registro c\d{6,}[a-zA-Z]?$", val_limpio.strip(), re.IGNORECASE) # Usuario con registro C101848W
            or len(val_limpio.split()) == 1  # 👈 solo una palabra
        ):
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Modificado Por",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Logica del Dato",
                "Observación Específica": "Capturar nombre completo, tener en cuenta que el dato entre el Property y EditPlot debe ser en creación el mismo y debe estar en formato tipo título",
                "Tipología": "Fondo"
            })

        # ------------------------------
        # 🚨 Otras validaciones (Formato, espacios, estandarización)
        # ------------------------------
        else:
            observaciones = []

            if val_limpio != val_limpio.title():
                observaciones.append("Errores en Formato")

            val_normalizada = unidecode(val_limpio).title()
            nombres_existentes = df_raw["Modificado Por"].dropna().unique()
            coincidencias = [n for n in nombres_existentes if unidecode(str(n)).title() == val_normalizada]

            if len(coincidencias) > 1:
                observaciones.append("Estandarizar Nombre a un solo registro")

            if val_raw.startswith(" "):
                observaciones.append("Espacio al inicio")
            if val_raw.endswith(" "):
                observaciones.append("Espacio al final")
            if "  " in val_raw:
                observaciones.append("Múltiples espacios")
            if "\n" in val_raw or "\r" in val_raw:
                observaciones.append("Saltos de línea")

            if observaciones:
                registros.append({
                    "ID": id_val,
                    "Columna Analizada": "Modificado Por",
                    "Dato Analizado": val_raw,
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "; ".join(observaciones),
                    "Tipología": "Forma"
                })

        # ---- Comentarios ----

    val_raw = df_raw.loc[idx, "Comentarios"]
    val = fila["Comentarios"]

    if pd.isna(val):  # Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Comentarios",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Comentarios",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con Comentarios",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):

        val_limpio = " ".join(val_raw.strip().split())
        val_lower = val_limpio.lower()
        observaciones = []
        formato_invalido = False

        # ------------------------------
        # 🚨 Casos especiales
        # ------------------------------

        # 1. "Sin Comentarios" exacto → válido
        if val_limpio == "Sin Comentarios":
            pass

        # 2. Variantes que deben estandarizarse a "Sin Comentarios"
        elif val_lower in [
            "no aplica", "n/a",
            "sin observacion", "sin observación",
            "sin informacion", "sin información",
            "sin observaciones", "sin observaciónes"
        ]:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Comentarios",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "Estandarizar a Sin Comentarios",
                "Tipología": "Forma"
            })

        # 3. Solo símbolos, solo números o una sola letra
        elif re.fullmatch(r"[\W_]+", val_limpio) or val_limpio.isdigit() or (len(val_limpio.split()) == 1 and len(val_limpio) == 1):
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Comentarios",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Logica del Dato",
                "Observación Específica": "Comentario no es claro",
                "Tipología": "Forma"
            })

        else:
            # ------------------------------
            # 🚨 Validación de espacios y saltos de línea
            # ------------------------------
            if val_raw.startswith(" "):
                observaciones.append("Espacio al inicio")
            if val_raw.endswith(" "):
                observaciones.append("Espacio al final")
            if "  " in val_raw:
                observaciones.append("Múltiples espacios")
            if "\n" in val_raw or "\r" in val_raw:
                observaciones.append("Saltos de línea")

            # ------------------------------
            # 🚨 Validación formato tipo oración con excepción de comillas
            # ------------------------------
            excepcion_comillas = False
            bloques = re.findall(r'"([^"]*)"', val_limpio)

            if bloques:
                excepcion_comillas = True
                for b in bloques:
                    if b == "" or b != b.title():  # vacío o no está en formato título
                        excepcion_comillas = False
                        break

            if not excepcion_comillas:
                if val_limpio:
                    if not val_limpio[0].isupper():  # debe iniciar en mayúscula
                        formato_invalido = True
                    if len(val_limpio) > 1 and val_limpio[1:].isupper():  # no todo mayúsculas
                        formato_invalido = True
                    if not val_limpio[-1].isalnum():  # debe terminar en letra o número
                        formato_invalido = True

            if formato_invalido:
                observaciones.append("Errores en Formato")

            # ------------------------------
            # 🚨 Consolidar observaciones
            # ------------------------------
            if observaciones:
                registros.append({
                    "ID": id_val,
                    "Columna Analizada": "Comentarios",
                    "Dato Analizado": val_raw,
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "; ".join(observaciones),
                    "Tipología": "Forma"
                })

        # 🚨 Validación ortográfica con Microsoft Word

    ''' word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Add()

        palabras = [p.strip(".,;:¡!¿?()\"'") for p in val_raw.split()]
        mal_escritas = [p for p in palabras if p and not doc.CheckSpelling(p)]

        doc.Close(False)
        word.Quit()

        if mal_escritas:
            errores.append(f"Errores ortográficos detectados: {', '.join(mal_escritas)}")'''

    # ---- Cód DANE Depto ----

    val_raw = df_raw.loc[idx, "Cód DANE Depto"]
    val = fila["Cód DANE Depto"]

    observaciones = []  # <- acumulador de observaciones

    if pd.isna(val):  # 🚨 Vacío
        observaciones.append({
            "ID": id_val,
            "Columna Analizada": "Cód DANE Depto",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # 🚨 solo espacios
        observaciones.append({
            "ID": id_val,
            "Columna Analizada": "Cód DANE Depto",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Cód DANE Depto",
            "Tipología": "Forma"
        })

    else:
        if isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
            errores_espacios = []
            if val_raw.startswith(" "):
                errores_espacios.append("Espacio al inicio")
            if val_raw.endswith(" "):
                errores_espacios.append("Espacio al final")
            if "  " in val_raw:
                errores_espacios.append("Múltiples espacios")
            if "\n" in val_raw or "\r" in val_raw:
                errores_espacios.append("Saltos de línea")

            if errores_espacios:
                observaciones.append({
                    "ID": id_val,
                    "Columna Analizada": "Cód DANE Depto",
                    "Dato Analizado": val_raw,
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "; ".join(errores_espacios),
                    "Tipología": "Forma"
                })

        # 🚨 Validación de longitud (solo 2 dígitos numéricos)
        try:
            val_str = str(val).strip()
            if not (val_str.isdigit() and len(val_str) == 2):
                observaciones.append({
                    "ID": id_val,
                    "Columna Analizada": "Cód DANE Depto",
                    "Dato Analizado": val_str,
                    "Observación General": "Inconsistencia Logica del Dato",
                    "Observación Específica": "Digitar solo 2 dígitos numéricos. Verificar con la fuente",
                    "Tipología": "Fondo"
                })
            else:
                # 🚨 Validación contra listado oficial de DANE
                if val_str not in codigos_dane_deptos:
                    observaciones.append({
                        "ID": id_val,
                        "Columna Analizada": "Cód DANE Depto",
                        "Dato Analizado": val_str,
                        "Observación General": "Inconsistencia Logica del Dato",
                        "Observación Específica": "Dato no corresponde al código DANE, Verificar con la fuente",
                        "Tipología": "Fondo"
                    })
        except Exception:
            pass

    # 🚨 Agregar todas las observaciones encontradas
    if observaciones:
        registros.extend(observaciones)
                 
    # ---- Cód DANE Mpio ----

    val_raw = df_raw.loc[idx, "Cód DANE Mpio"]
    val = fila["Cód DANE Mpio"]

    observaciones = []  # <- acumulador

    if pd.isna(val):  # 🚨 Vacío
        observaciones.append({
            "ID": id_val,
            "Columna Analizada": "Cód DANE Mpio",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # 🚨 solo espacios
        observaciones.append({
            "ID": id_val,
            "Columna Analizada": "Cód DANE Mpio",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Cód DANE Mpio",
            "Tipología": "Forma"
        })

    else:
        if isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
            errores_espacios = []
            if val_raw.startswith(" "):
                errores_espacios.append("Espacio al inicio")
            if val_raw.endswith(" "):
                errores_espacios.append("Espacio al final")
            if "  " in val_raw:
                errores_espacios.append("Múltiples espacios")
            if "\n" in val_raw or "\r" in val_raw:
                errores_espacios.append("Saltos de línea")

            if errores_espacios:
                observaciones.append({
                    "ID": id_val,
                    "Columna Analizada": "Cód DANE Mpio",
                    "Dato Analizado": val_raw,
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "; ".join(errores_espacios),
                    "Tipología": "Forma"
                })

        # 🚨 Validación de longitud y contenido
        try:
            val_str = str(val).strip()

            # Caso: 3 dígitos
            if val_str.isdigit() and len(val_str) == 3:
                pass  # válido, no se reporta nada

            # Caso: 5 dígitos
            elif val_str.isdigit() and len(val_str) == 5:
                depto = val_str[:2]  # primeros dos dígitos
                if depto in codigos_dane_deptos:
                    observaciones.append({
                        "ID": id_val,
                        "Columna Analizada": "Cód DANE Mpio",
                        "Dato Analizado": val_str,
                        "Observación General": "El Dato no guarda el estandar del Diccionario de Datos",
                        "Observación Específica": "Extraer y reemplazar los caracteres desde la posición 3 al 5 del dato Cód DANE Mpio",
                        "Tipología": "Forma"
                    })
                else:
                    observaciones.append({
                        "ID": id_val,
                        "Columna Analizada": "Cód DANE Mpio",
                        "Dato Analizado": val_str,
                        "Observación General": "Inconsistencia Lógica del Dato",
                        "Observación Específica": "Dato no guarda relación con Código DANE, este debe contar con 3 dígitos. Verificar Dato",
                        "Tipología": "Fondo"
                    })

            # Caso: ni 3 ni 5 dígitos, o caracteres no numéricos
            else:
                observaciones.append({
                    "ID": id_val,
                    "Columna Analizada": "Cód DANE Mpio",
                    "Dato Analizado": val_str,
                    "Observación General": "Inconsistencia Lógica del Dato",
                    "Observación Específica": "Dato no guarda relación con Código DANE, Verificar Dato",
                    "Tipología": "Fondo"
                })

        except Exception:
            pass

    # 🚨 Agregar todas las observaciones
    if observaciones:
        registros.extend(observaciones)

    # ---- Año Vigencia Insumo Geográfico ----
    val_raw = df_raw.loc[idx, "Año Vigencia Insumo Geográfico"]
    val = fila["Año Vigencia Insumo Geográfico"]

    anio_captura = df_raw.loc[idx, "Anio_Captura"]
    anio_vigencia = df_raw.loc[idx, "Anio_Vigencia_Num"]

    observaciones = []

    # 1. Totalidad (vacío o NaN reales)
    if pd.isna(val) or str(val).strip() == "":
        observaciones.append({
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Fondo"
        })

    else:
        val = str(val).strip()  # 👈 solo convierto a string si no es NaN

        # 2. Sin Información o -9999
        if val.upper().replace("Ó", "O") in ["SIN INFORMACION", "SIN INFORMACIÓN", "-9999", "1900"]:
            if val != "Sin Información":
                observaciones.append({
                    "Observación General": "El Dato no guarda el estandar del Diccionario de Datos",
                    "Observación Específica": "Estandarizar a Sin Información",
                    "Tipología": "Forma"
                })

        # 3. Estructura DD/MM/AAAA o variaciones D/M/AAAA
        elif re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", val):
            observaciones.append({
                "Observación General": "El Dato no guarda el estandar del Diccionario de Datos",
                "Observación Específica": "Capturar solo el año de vigencia del insumo geográfico",
                "Tipología": "Fondo"
            })

        # 4. Numérico válido de 4 dígitos
        elif val.isdigit() and len(val) == 4:
            anio_val = int(val)
            if anio_val < 2000:
                observaciones.append({
                    "Observación General": "Inconsistencia Logica del Dato",
                    "Observación Específica": "Revisar el año de insumo geográfico",
                    "Tipología": "Fondo"
                })
            if anio_captura and anio_val > anio_captura:
                observaciones.append({
                    "Observación General": "Inconsistencia Logica del Dato",
                    "Observación Específica": "Fecha del Insumo no debe ser superior a la fecha de captura",
                    "Tipología": "Fondo"
                })

        # 5. Texto no válido o valor numérico incorrecto
        else:
            observaciones.append({
                "Observación General": "Inconsistencia Logica del Dato",
                "Observación Específica": "Capturar el año de vigencia del Dato",
                "Tipología": "Fondo"
            })

    # Registrar todas las observaciones
    for obs in observaciones:
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Año Vigencia Insumo Geográfico",
            "Dato Analizado": val_raw,  # mantengo el valor original
            "Anio_Captura": anio_captura,
            "Anio_Vigencia_Num": anio_vigencia,
            **obs
        })


    # ---- Nombre Vereda ----
    val_raw = df_raw.loc[idx, "Nombre Vereda"]
    val = fila["Nombre Vereda"]

    if pd.isna(val) or str(val).strip() == "":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Nombre Vereda",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif str(val).strip() == "" or str(val).upper() == "<ESPACIO>":
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Nombre Vereda",
            "Dato Analizado": val_raw,
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Nombre Vereda",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):
        val_str = str(val_raw).strip()
        observaciones_especificas = []

        # -----------------------------
        # 1. FMI / Según Campo / Divipola / Documentos (prioridad)
        # -----------------------------
        if any(word.lower() in val_str.lower() for word in ["fmi", "según campo", "campo","divipola", "documentos","igac","vur","registro"]):
            observaciones_especificas.append("Diligenciar solo el dato correspondiente a FMI")

        else:
            # -----------------------------
            # 2. Validación Formato Título con excepciones
            # -----------------------------
            palabras = val_str.split()
            excepciones_minuscula = ["de", "del", "y"]
            excepciones_mayuscula = ["la", "las", "los", "el"]
            errores_formato = []

            for i, p in enumerate(palabras):
                if p.lower() in excepciones_minuscula:
                    if i == 0:
                        # Primera palabra -> se permite que empiece en mayúscula
                        continue
                    else:
                        if p != p.lower():
                            errores_formato.append("Errores en Formato")
                elif p.lower() in excepciones_mayuscula:
                    if not (p == p.lower() or p == p.capitalize()):
                        errores_formato.append("Errores en Formato")

            else:
                if p != p.capitalize():
                    "Errores en Formato"
                
            # -----------------------------
            # 3. Reglas adicionales
            # -----------------------------
            if re.search(r"\bvereda\b", val_str.lower()):
                observaciones_especificas.append("Solo capturar el nombre de vereda")

            if "no aplica" in val_str.lower():
                observaciones_especificas.append("Estandarizar a Sin Información")

            if re.search(r"\d+\s*(km|KM|m|M)\b", val_str):
                observaciones_especificas.append("Eliminar datos de metraje")

            if re.search(r"[,;.:]$", val_str) or re.search(r"[^a-zA-ZÀ-ÿ0-9\s,\-/]", val_str):
                observaciones_especificas.append("Eliminar caracteres especiales")
            
            if any(word in val_str.lower() for word in ["urbano", "zona urbana"]):
                observaciones_especificas.append("El dato debe diligenciarse como No Aplica si se sitúa en zona urbana")

            # -----------------------------
            # 4. Similaridades entre veredas (fuzzy matching)
            # -----------------------------
            valores_unicos = df_raw["Nombre Vereda"].dropna().unique()
            similares = process.extract(
                val_str,
                valores_unicos,
                scorer=fuzz.token_sort_ratio,
                limit=5
            )
            for match, score, _ in similares:
                if score >= 85 and match.lower() != val_str.lower():
                    observaciones_especificas.append("Estandarizar Nombre Vereda a un único registro")
                    break

        # -----------------------------
        # 5. Validaciones de Fondo (Inconsistencias Lógicas)
        # -----------------------------
        val_lower = val_str.lower()

        # 1. Caso Sin Información
        if val_lower in ["sin información", "sin informacion"]:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Nombre Vereda",
                "Dato Analizado": val_raw,
                "Observación General": "Inconsistencia Logica del Dato",
                "Observación Específica": (
                    "Capturar Nombre de Vereda como primer insumo se deberá capturar el del folio de matrícula, "
                    "luego catastro y finalmente cruce espacial con la capa de veredas de DANE"
                ),
                "Tipología": "Fondo"
            })

        else:
            # Palabras o expresiones que NO corresponden a nombre de vereda
            palabras_no_vereda = [
                "corregimiento", "inspección", "lote", "sin zona", "sin definir", "por definir",
                "directriz ecopetrol", "área de expansión", "cabecera municipal", "el 6",
                "zona especial", "cgto", "rural", "vereda con centro poblado", "zona fiscal",
                "casa lote", "casa lt"
            ]

            # Construir regex para buscar como palabra/frase completa
            patron_no_vereda = r"\b(" + "|".join(palabras_no_vereda) + r")\b"

            # Coincidencia con lista o valor puramente numérico
            if re.search(patron_no_vereda, val_lower) or re.fullmatch(r"\d+", val_str):
                registros.append({
                    "ID": id_val,
                    "Columna Analizada": "Nombre Vereda",
                    "Dato Analizado": val_raw,
                    "Observación General": "Inconsistencia Logica del Dato",
                    "Observación Específica": "El dato no corresponde a Nombre de Vereda",
                    "Tipología": "Fondo"
                })
        
        # -----------------------------
        # Registrar si hubo observaciones
        # -----------------------------
        if observaciones_especificas:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Nombre Vereda",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estandar del Diccionario de Datos",
                "Observación Específica": "; ".join(observaciones_especificas),
                "Tipología": "Forma"
            })

    # ---- RULEID ----

    val_raw = df_raw.loc[idx, "RULEID"]
    val = fila["RULEID"]

    # 🚨 1. Validación de vacíos
    if pd.isna(val):  
        registros.append({
            "ID": id_val,
            "Columna Analizada": "RULEID",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  
        registros.append({
            "ID": id_val,
            "Columna Analizada": "RULEID",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el RULEID",
            "Tipología": "Forma"
        })

    else:
        try:
            val_num = int(val) if str(val).strip() != "" else None

            # 🚨 2. Validación de dominio (solo se permite 1)
            if val_num != 1:
                registros.append({
                    "ID": id_val,
                    "Columna Analizada": "RULEID",
                    "Dato Analizado": str(val),
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "Valor no hace parte del dominio, Diligenciar con valor 1",
                    "Tipología": "Forma"
                })
            else:
                # 🚨 3. Validación de espacios SOLO si el valor es 1
                if isinstance(val_raw, str):
                    errores_espacios = []
                    if val_raw.startswith(" "):
                        errores_espacios.append("Espacio al inicio")
                    if val_raw.endswith(" "):
                        errores_espacios.append("Espacio al final")
                    if "  " in val_raw:
                        errores_espacios.append("Múltiples espacios")
                    if "\n" in val_raw or "\r" in val_raw:
                        errores_espacios.append("Saltos de línea")

                    if errores_espacios:
                        registros.append({
                            "ID": id_val,
                            "Columna Analizada": "RULEID",
                            "Dato Analizado": val_raw,
                            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                            "Observación Específica": "; ".join(errores_espacios),
                            "Tipología": "Forma"
                        })

        except Exception:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "RULEID",
                "Dato Analizado": str(val),
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "Valor no numérico, Diligenciar con valor 1",
                "Tipología": "Forma"
            })     

    # ---- Código SIG Predio Jurídico ----

    val_raw = df_raw.loc[idx, "Código SIG Predio Jurídico"]
    val = fila["Código SIG Predio Jurídico"]

    observaciones = []  # lista de observaciones para este campo

    # 🚨 1. Vacíos
    if pd.isna(val):
        observaciones.append({
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":
        observaciones.append({
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio",
            "Tipología": "Forma"
        })

    else:
        try:
            val_str = str(val).strip()
            val_up = val_str.upper()  # normaliza para prefijos
            estructura_valida = False  # bandera para validar espacios al final
            tiene_extras = any(c in val_str for c in ["_", "-"]) or not val_str.isalnum()

            # 🚨 2. Unicidad (siempre acumula)
            if (df_raw["Código SIG Predio Jurídico"] == val_str).sum() > 1:
                observaciones.append({
                    "Observación General": "Inconsistencia Lógica del Dato",
                    "Observación Específica": "Cód. SIG se encuentra más de una vez",
                    "Tipología": "Fondo"
                })

            # 🚨 3. Estructura (orden: CLC → CO → L → SC → C (catch-all) → numérico → else)
            if val_up.startswith("CLC"):
                # CLC0 + 4 dígitos → total 8
                if len(val_up) == 8 and val_up[3] == "0" and val_up[4:].isdigit() and not tiene_extras:
                    estructura_valida = True
                else:
                    observaciones.append({
                        "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                        "Observación Específica": "Estructura no cumple con el Diccionario de Datos" if tiene_extras
                                                else ("Valor no válido" if not val_up[4:].isdigit()
                                                    else "Estandarizar de acuerdo al diccionario de Datos (CLC)"),
                        "Tipología": "Forma"
                    })

            elif val_up.startswith("CO"):
                # CO + [31-36] + 5–6 dígitos → total 9 o 10
                if (len(val_up) in [9, 10] and val_up[2:4] in ["31","32","33","34","35","36"]
                    and val_up[4:].isdigit() and not tiene_extras):
                    estructura_valida = True
                else:
                    observaciones.append({
                        "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                        "Observación Específica": "Estructura no cumple con el Diccionario de Datos" if tiene_extras
                                                else ("Valor no válido" if not val_up[4:].isdigit()
                                                    else "Estandarizar de acuerdo al diccionario de Datos (CO)"),
                        "Tipología": "Forma"
                    })

            elif val_up.startswith("L"):
                # ✅ Válido: L0 + 4 dígitos (largo 6) y sin separadores
                if val_up.startswith("L0") and len(val_up) == 6 and val_up[2:].isdigit() and val_str.isalnum():
                    estructura_valida = True
                else:
                    # 1) Tiene L0 pero después aparece cualquier no-dígito (letras, separadores, etc.) → Estructura no cumple
                    if val_up.startswith("L0") and (not val_up[2:].isdigit() or not val_str.isalnum()):
                        observaciones.append({
                            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                            "Observación Específica": "Estructura no cumple con el Diccionario de Datos",
                            "Tipología": "Forma"
                        })
                    # 2) Tiene L0 y solo dígitos, pero el largo no es 6 (faltan/sobran) → Estandarizar (L)
                    elif val_up.startswith("L0") and val_up[2:].isdigit() and len(val_up) != 6:
                        observaciones.append({
                            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                            "Observación Específica": "Estandarizar de acuerdo al diccionario de Datos (L)",
                            "Tipología": "Forma"
                        })
                    # 3) No cumple el prefijo L0 (p.ej., LADESPENSA, L12345) → Valor no válido
                    else:
                        observaciones.append({
                            "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                            "Observación Específica": "Valor no válido",
                            "Tipología": "Forma"
                        })

            elif val_up.startswith("SC"):
                # SC0 + 4 dígitos → total 7
                if len(val_up) == 7 and val_up[2] == "0" and val_up[3:].isdigit() and not tiene_extras:
                    estructura_valida = True
                else:
                    observaciones.append({
                        "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                        "Observación Específica": "Estructura no cumple con el Diccionario de Datos" if tiene_extras
                                                else ("Valor no válido" if not val_up[3:].isdigit()
                                                    else "Estandarizar de acuerdo al diccionario de Datos (SC)"),
                        "Tipología": "Forma"
                    })

            elif val_up.startswith("C"):
                # ⚠️ Catch-all: cualquier 'C...' que no sea CLC ni CO → estandarizar como CO
                observaciones.append({
                    "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                    "Observación Específica": "Estandarizar de acuerdo al diccionario de Datos (CO)",
                    "Tipología": "Forma"
                })

            elif val_up[0].isdigit():
                # Numérico puro entre 4 y 10 dígitos
                if val_up.isdigit() and 4 <= len(val_up) <= 10 and not tiene_extras:
                    estructura_valida = True
                else:
                    observaciones.append({
                        "Observación General": "El Dato no guarda el estándar del Diccionario de Datos" if tiene_extras
                                            else "Inconsistencia Lógica del Dato",
                        "Observación Específica": "Estructura no cumple con el Diccionario de Datos" if tiene_extras
                                                else "Revisar la consistencia del Cód. SIG",
                        "Tipología": "Forma" if tiene_extras else "Fondo"
                    })

            else:
                observaciones.append({
                    "Observación General": "Inconsistencia Lógica del Dato",
                    "Observación Específica": "Revisar la consistencia del Cód. SIG",
                    "Tipología": "Fondo"
                })

            # 🚨 4. Espacios problemáticos SOLO si la estructura es válida
            if estructura_valida and isinstance(val_raw, str):
                errores_espacios = []
                if val_raw.startswith(" "): errores_espacios.append("Espacio al inicio")
                if val_raw.endswith(" "):   errores_espacios.append("Espacio al final")
                if "  " in val_raw:         errores_espacios.append("Múltiples espacios")
                if "\n" in val_raw or "\r" in val_raw: errores_espacios.append("Saltos de línea")

                if errores_espacios:
                    observaciones.append({
                        "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                        "Observación Específica": "; ".join(errores_espacios),
                        "Tipología": "Forma"
                    })

        except Exception:
            observaciones.append({
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "Revisar la consistencia del Cód. SIG",
                "Tipología": "Forma"
            })

    # 🚨 Registrar todas las observaciones acumuladas
    for obs in observaciones:
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Código SIG Predio Jurídico",
            "Dato Analizado": str(val) if not pd.isna(val) else "",
            **obs
        })

    # ---- Área Terreno Calculada Mts2 ----

    val_raw = df_raw.loc[idx, "Área Terreno Calculada Mts2"]
    val = fila["Área Terreno Calculada Mts2"]
    
    if pd.isna(val): # Vacio
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Área Terreno Calculada Mts2",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Área Terreno Calculada Mts2",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Área Terreno Calculada Mts2",
            "Tipología": "Forma"
        })
    
    elif isinstance(val_raw, str):  # 🚨 Validación de espacios problemáticos
        errores_espacios = []
        if val_raw.startswith(" "):
            errores_espacios.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores_espacios.append("Espacio al final")
        if "  " in val_raw:
            errores_espacios.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            errores_espacios.append("Saltos de línea")
    
        if errores_espacios:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Área Terreno Calculada Mts2",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(errores_espacios),
                "Tipología": "Forma"
        })
                             
    # ---- Tipo de Propiedad ----

    val_raw = df_raw.loc[idx, "Tipo de Propiedad"]
    val = fila["Tipo de Propiedad"]

    if pd.isna(val):  # 🚨 Vacío
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Tipo de Propiedad",
            "Dato Analizado": "",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato sin diligenciar",
            "Tipología": "Forma"
        })

    elif val == "<ESPACIO>":  # 🚨 Solo espacios
        registros.append({
            "ID": id_val,
            "Columna Analizada": "Tipo de Propiedad",
            "Dato Analizado": " ",
            "Observación General": "Inconsistencia Totalidad del Dato",
            "Observación Específica": "Dato diligenciado únicamente con espacio, Dato no es coherente con el Tipo de Propiedad",
            "Tipología": "Forma"
        })

    elif isinstance(val_raw, str):  
        val_str = val_raw.strip()

        # 🚨 Validación de espacios problemáticos
        errores_espacios = []
        if val_raw.startswith(" "):
            errores_espacios.append("Espacio al inicio")
        if val_raw.endswith(" "):
            errores_espacios.append("Espacio al final")
        if "  " in val_raw:
            errores_espacios.append("Múltiples espacios")
        if "\n" in val_raw or "\r" in val_raw:
            errores_espacios.append("Saltos de línea")

        if errores_espacios:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Tipo de Propiedad",
                "Dato Analizado": val_raw,
                "Observación General": "El Dato no guarda el estándar del Diccionario de Datos",
                "Observación Específica": "; ".join(errores_espacios),
                "Tipología": "Forma"
            })

        # 🚨 Validación de dominios permitidos
        dominios_permitidos = ["PRESUNTAMENTE BALDIO", "PRIVADA", "SIN INFORMACION"]
        if val_str.upper() not in dominios_permitidos:
            registros.append({
                "ID": id_val,
                "Columna Analizada": "Tipo de Propiedad",
                "Dato Analizado": val_str,
                "Observación General": "Inconsistencia Lógica del Dato",
                "Observación Específica": "Dominio no se encuentra de acuerdo con el Diccionario de Datos",
                "Tipología": "Fondo"
            })
             
# ==========================
# Convertir a DataFrame
# ==========================
reporte = pd.DataFrame(registros)

# Añadir la columna Obs_Nom_Proyect desde df (donde la fuimos guardando en cada fila)
df_merge = df.merge(
    reporte,
    on="ID",
    how="left",
    suffixes=("", "_obs")
)
# ==========================
# Función para limpiar caracteres ilegales de Excel
# ==========================
def limpiar_excel(value):
    if isinstance(value, str):
        # Quita caracteres de control ASCII 0–31 y 127
        return re.sub(r'[\x00-\x1F\x7F]', '', value)
    return value

# ==========================
# Guardar en Excel con hojas separadas
# ==========================
if not reporte.empty:
    with pd.ExcelWriter("Inconsistencias_EditedPlot.xlsx", engine="openpyxl") as writer:
        for columna in columnas_objetivo:
            if columna == "ID":
                continue  # no validamos ID directamente
            df_columna = reporte[reporte["Columna Analizada"] == columna]
            if not df_columna.empty:
                nombre_hoja = columna[:31]  # Excel permite máx 31 caracteres
                df_columna = df_columna.applymap(limpiar_excel)    
                df_columna.to_excel(writer, sheet_name=nombre_hoja, index=False)
    print("✅ Reporte generado: 'Inconsistencias_EditedPlot.xlsx'")
    print(f"📊 Total inconsistencias encontradas: {len(reporte)}")
else:
    print("✅ No se encontraron campos inconsistentes en las columnas seleccionadas.")
input("\nPresiona ENTER para salir...")
