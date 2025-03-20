import pandas as pd

# Cargar el archivo Excel desde la fila 3
df = pd.read_excel("base_datos_partidas.xlsx", header=2)

# Mostrar los nombres de las columnas detectados
print("Columnas detectadas:", df.columns)
