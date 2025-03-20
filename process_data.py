import pandas as pd
import unicodedata
import re
from rapidfuzz import process, fuzz

# Función para normalizar textos
def normalizar_texto(texto):
    # Convertir a minúsculas
    texto = texto.lower()
    # Eliminar acentos
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    # Eliminar caracteres no alfanuméricos (excepto espacios)
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    # Eliminar espacios extra y recortar
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# Función para leer la base de datos de partidas
def extraer_partidas_construccion(file_path):
    # Leer el Excel asumiendo que el encabezado está en la fila 3 (header=2)
    df = pd.read_excel(file_path, header=2)
    # Extraer las columnas relevantes y renombrar 'Pres' a 'Precio'
    base_datos_df = df[['Código', 'Resumen', 'Pres']].copy()
    base_datos_df = base_datos_df.rename(columns={'Pres': 'Precio'})
    # Eliminar filas con datos nulos en las columnas importantes
    base_datos_df = base_datos_df.dropna(subset=['Código', 'Resumen', 'Precio'])
    # Crear columna de resumen normalizado para comparaciones
    base_datos_df['Resumen_normalizado'] = base_datos_df['Resumen'].apply(normalizar_texto)
    return base_datos_df

# Cargar la base de datos de partidas
base_datos = extraer_partidas_construccion('base_datos_partidas.xlsx')

# Leer el archivo Excel con las nuevas partidas (encabezado en la fila 3)
nuevas_partidas = pd.read_excel('nuevas_partidas.xlsx', header=2)

# Verificar que la columna 'Resumen' existe en el archivo de nuevas partidas
if 'Resumen' not in nuevas_partidas.columns:
    print("❌ ERROR: La columna 'Resumen' no se encuentra en nuevas_partidas.xlsx")
    print("Columnas detectadas:", nuevas_partidas.columns)
    exit()

# Normalizar la columna 'Resumen' en las nuevas partidas
nuevas_partidas['Resumen_normalizado'] = nuevas_partidas['Resumen'].apply(normalizar_texto)

# Función para encontrar la partida más parecida usando fuzzy matching
def encontrar_partida_parecida(descripcion, base_datos, umbral=70):
    # Se utiliza la columna normalizada para la comparación
    lista_descripciones = base_datos['Resumen_normalizado'].tolist()
    mejor_coincidencia, puntuacion, indice = process.extractOne(
        descripcion, lista_descripciones, scorer=fuzz.token_sort_ratio
    )
    if puntuacion >= umbral:
        return mejor_coincidencia, base_datos.iloc[indice]['Precio']
    return None, None

# Asignar precios a las nuevas partidas aplicando la función de matching
nuevas_partidas['Precio Asignado'] = nuevas_partidas['Resumen_normalizado'].apply(
    lambda desc: encontrar_partida_parecida(desc, base_datos, umbral=70)[1] if pd.notna(desc) else 'No encontrado'
)

# Guardar el resultado en un nuevo archivo Excel
nuevas_partidas.to_excel('nuevas_partidas_con_precios.xlsx', index=False)

print("✅ Proceso completado. Revisa el archivo 'nuevas_partidas_con_precios.xlsx'")

