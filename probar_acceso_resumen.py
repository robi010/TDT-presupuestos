import pandas as pd

# Cargar el archivo nuevas_partidas.xlsx
df = pd.read_excel("nuevas_partidas.xlsx")

# Verificar si la columna 'Resumen' está accesible
if 'Resumen' in df.columns:
    print("✅ La columna 'Resumen' está accesible correctamente.")
    print("Primeros 5 valores de la columna 'Resumen':")
    print(df['Resumen'].head())
else:
    print("❌ La columna 'Resumen' NO se encuentra en nuevas_partidas.xlsx.")
