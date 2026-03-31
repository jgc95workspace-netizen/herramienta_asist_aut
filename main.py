import pandas as pd

# Ruta INPUT_OFICIAL
ruta_oficial = "data/oficial.xlsx"

# Leer Excel
oficial = pd.read_excel(ruta_oficial)

# Mostrar contenido
print("Contenido del archivo oficial: ")
print(oficial)

#Crear nombre completo
oficial["nombre_completo"] = oficial["apellidos"] + " " + oficial["nombres"]

#Mostrar resultado
print("\nCon nombre completo:")
print(oficial[["nombre_completo"]])
