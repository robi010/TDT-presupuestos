import pandas as pd

# Cargar el archivo nuevas_partidas.xlsx
df = pd.read_excel("nuevas_partidas.xlsx")

# Mostrar los nombres de las columnas detectadas
print("🔍 Columnas detectadas en nuevas_partidas.xlsx:")
print(df.columns)
