import pandas as pd
import os

# --- Rutas ---
base_path = os.path.dirname(os.path.abspath(__file__))
ruta_oficial = os.path.join(base_path, "..", "data", "oficial.xlsx")
ruta_asistencia = os.path.join(base_path, "..", "data", "asistencia.csv")

# --- Función para normalizar nombres ---
def normalizar_nombre(nombre):
    return " ".join(str(nombre).strip().lower().split())

# --- Leer oficial ---
def cargar_oficial(ruta):
    try:
        df = pd.read_excel(ruta)
        # Normalizar nombre completo en oficial
        df['nombre_completo'] = df['apellidos'].str.strip() + " " + df['nombres'].str.strip()
        df['nombre_completo'] = df['nombre_completo'].apply(normalizar_nombre)
        df['correo'] = df['correo'].str.strip().str.lower()
        print("✔ Oficial cargado")
        return df
    except Exception as e:
        raise Exception(f"❌ Error cargando oficial: {e}")

# --- Leer Teams ---
def leer_csv_teams(ruta):
    try:
        df = pd.read_csv(ruta, sep='\t', engine='python', encoding='utf-16-le', header=10-1)
        df = df.iloc[:, [0, 3, 4]]  # columnas nombre(A), duracion(D), correo(E)
        df.columns = ['nombre', 'duracion', 'correo']
        df = df.apply(lambda col: col.astype(str).str.strip())
        df['nombre'] = df['nombre'].apply(normalizar_nombre)
        df['correo'] = df['correo'].str.lower()
        return df
    except Exception as e:
        raise Exception(f"❌ Error procesando CSV Teams: {e}")

# --- Leer Zoom ---
def leer_csv_zoom(ruta):
    try:
        df = pd.read_csv(ruta, sep=',', engine='python', encoding='utf-8-sig', header=4-1)
        df = df.iloc[:, [0, 1, 2]]  # columnas nombre(A), correo(B), asistencia(C)
        df.columns = ['nombre', 'correo', 'duracion']
        df = df.apply(lambda col: col.astype(str).str.strip())
        df['nombre'] = df['nombre'].apply(normalizar_nombre)
        df['correo'] = df['correo'].str.lower()
        return df
    except Exception as e:
        raise Exception(f"❌ Error procesando CSV Zoom: {e}")

# --- Detectar tipo de CSV ---
def leer_asistencia(ruta):
    try:
        return leer_csv_teams(ruta)
    except:
        try:
            return leer_csv_zoom(ruta)
        except Exception as e:
            raise Exception(f"❌ Error detectando tipo de CSV: {e}")

# --- Cruce asistencia vs oficial ---
def cruzar_asistencia(oficial, asistencia):
    # Crear columna para marcar si hay match
    asistencia['match'] = asistencia.apply(
        lambda row: 
            any(oficial['nombre_completo'] == row['nombre']) or
            any(oficial['correo'] == row['correo']),
        axis=1
    )
    # Filtrar solo los que tienen match
    df_match = asistencia[asistencia['match']].copy()
    return df_match

# --- Main ---
if __name__ == "__main__":
    oficial = cargar_oficial(ruta_oficial)
    asistencia_raw = leer_asistencia(ruta_asistencia)
    print("✔ Asistencia cargada correctamente")
    
    asistencia_final = cruzar_asistencia(oficial, asistencia_raw)
    print(f"✔ Coincidencias encontradas: {len(asistencia_final)}")
    print(asistencia_final[['nombre','correo','duracion']])