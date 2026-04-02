import pandas as pd
import unicodedata
from rapidfuzz import process, fuzz
import os

# =========================
# 1. RUTAS
# =========================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

ruta_oficial = os.path.join(BASE_DIR, "data", "oficial.xlsx")
ruta_asistencia = os.path.join(BASE_DIR, "data", "asistencia.csv")  # 👈 AHORA CSV REAL
ruta_salida = os.path.join(BASE_DIR, "data", "resultado_final.xlsx")

# =========================
# 2. LIMPIEZA
# =========================

def limpiar_texto(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto).lower().strip()

    if "(" in texto:
        texto = texto.split("(")[0]

    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )

    texto = " ".join(texto.split())

    return texto

# =========================
# 3. CARGAR OFICIAL
# =========================

oficial = pd.read_excel(ruta_oficial)

oficial["nombre_completo"] = oficial["apellidos"] + " " + oficial["nombres"]
oficial["nombre_limpio"] = oficial["nombre_completo"].apply(limpiar_texto)

if "correo" in oficial.columns:
    oficial["correo"] = oficial["correo"].astype(str).str.lower().str.strip()

print("✔ Oficial cargado")

# =========================
# 4. LEER CSV INTELIGENTE
# =========================

def leer_csv_inteligente(ruta):
    for sep in [",", ";", "\t"]:
        try:
            df = pd.read_csv(ruta, sep=sep, encoding="latin1")
            if df.shape[1] > 1:
                print(f"✔ CSV leído con separador: '{sep}'")
                return df
        except:
            continue

    raise Exception("❌ No se pudo leer el CSV correctamente")

asistencia_raw = leer_csv_inteligente(ruta_asistencia)

# =========================
# 5. DURACION TEAMS
# =========================

def convertir_duracion_teams(texto):
    if pd.isna(texto):
        return 0

    texto = str(texto)

    horas = 0
    minutos = 0

    if "h" in texto:
        try:
            horas = int(texto.split("h")[0])
        except:
            pass

    if "min" in texto:
        try:
            minutos = int(texto.split("min")[0].split()[-1])
        except:
            pass

    return horas * 60 + minutos

# =========================
# 6. PROCESAR TEAMS
# =========================

def procesar_teams(df):

    fila_inicio = None

    for i in range(len(df)):
        if str(df.iloc[i, 0]).strip().lower() == "nombre":
            fila_inicio = i
            break

    if fila_inicio is None:
        raise Exception("❌ No se encontró encabezado Teams")

    df = df.iloc[fila_inicio + 1:].copy()
    df = df.iloc[:, :5]

    df.columns = ["nombre", "entrada", "salida", "duracion", "correo"]

    df["nombre_limpio"] = df["nombre"].apply(limpiar_texto)
    df["correo"] = df["correo"].astype(str).str.lower().str.strip()
    df["duracion_min"] = df["duracion"].apply(convertir_duracion_teams)

    return df

# =========================
# 7. PROCESAR
# =========================

asistencia = procesar_teams(asistencia_raw)

print("✔ Asistencia procesada")

# =========================
# 8. MATCH
# =========================

resultado = oficial.merge(
    asistencia[["correo", "duracion_min"]],
    on="correo",
    how="left"
)

resultado["asistencia"] = resultado["duracion_min"].notna()

# =========================
# 9. EXPORTAR
# =========================

resultado.to_excel(ruta_salida, index=False)

print("🔥 Archivo generado:", ruta_salida)
