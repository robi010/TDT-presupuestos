import pandas as pd
from rapidfuzz import process, fuzz

# Función para leer la base de datos de partidas
def extraer_partidas_construccion(file_path):
    # Leer desde la fila 3 (header=2)
    df = pd.read_excel(file_path, header=2)
    
    # Extraer columnas relevantes
    base_datos_df = df[['Código', 'Resumen', 'Pres']].copy()
    base_datos_df = base_datos_df.rename(columns={'Pres': 'Precio'})  # Renombrar 'Pres' a 'Precio'
    
    # Limpiar filas vacías
    base_datos_df = base_datos_df.dropna(subset=['Código', 'Resumen', 'Precio'])
    
    return base_datos_df

# Leer la base de datos de partidas
base_datos = extraer_partidas_construccion('base_datos_partidas.xlsx')

# Leer el archivo de nuevas partidas usando header=2
nuevas_partidas = pd.read_excel('nuevas_partidas.xlsx', header=2)

# Verificar que 'Resumen' está correctamente cargado
if 'Resumen' not in nuevas_partidas.columns:
    print("❌ ERROR: La columna 'Resumen' no está en nuevas_partidas.xlsx")
    print("Columnas detectadas:", nuevas_partidas.columns)
    exit()

# Función para encontrar la partida más parecida
def encontrar_partida_parecida(descripcion, base_datos, umbral=80):
    lista_descripciones = base_datos['Resumen'].tolist()
    mejor_coincidencia, puntuacion, indice = process.extractOne(
        descripcion, lista_descripciones, scorer=fuzz.token_sort_ratio
    )
    if puntuacion >= umbral:
        return mejor_coincidencia, base_datos.iloc[indice]['Precio']
    return None, None

# Crear una nueva columna para el precio en las nuevas partidas
nuevas_partidas['Precio Asignado'] = nuevas_partidas['Resumen'].apply(
    lambda desc: encontrar_partida_parecida(desc, base_datos, umbral=80)[1] if pd.notna(desc) else 'No encontrado'
)

# Guardar el resultado en un nuevo archivo Excel
nuevas_partidas.to_excel('nuevas_partidas_con_precios.xlsx', index=False)

print("✅ Proceso completado. Revisa el archivo 'nuevas_partidas_con_precios.xlsx'")

