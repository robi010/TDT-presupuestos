import pandas as pd
import unicodedata
import re
from rapidfuzz import process, fuzz

# Función para normalizar textos
def normalizar_texto(texto):
    texto = texto.lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# Función para procesar la plantilla: combinar título y resumen
def procesar_template(df, es_base=True):
    """
    Recorre el DataFrame que sigue la plantilla:
      - Cuando la columna 'Código' tiene dato, se toma esa fila como inicio de la partida.
      - Se espera que en la misma fila la columna 'Resumen' contenga el título.
      - En la fila siguiente (sin código) se encuentra el resumen o descripción.
    Se combina el título y el resumen en un único texto.
    Para la base, se conserva el precio de la columna 'Pres'.
    """
    items = []
    i = 0
    while i < len(df):
        row = df.iloc[i]
        codigo = row['Código']
        if pd.notna(codigo):
            # Se asume que esta fila contiene el código y el título.
            title = row['Resumen'] if pd.notna(row['Resumen']) else ""
            summary = ""
            # Si existe la siguiente fila y no tiene código, se asume que es el resumen.
            if i + 1 < len(df):
                next_row = df.iloc[i + 1]
                if pd.isna(next_row['Código']):
                    summary = next_row['Resumen'] if pd.notna(next_row['Resumen']) else ""
                    i += 1  # Saltamos la fila del resumen, ya que la hemos consumido
            texto_completo = (title + " " + summary).strip()
            # Para la base, obtenemos el precio de la columna 'Pres'; en nuevas partidas puede no haberlo.
            precio = row['Pres'] if es_base and 'Pres' in row else None
            items.append({
                'Código': codigo,
                'Texto_completo': texto_completo,
                'Precio': precio
            })
        i += 1
    return pd.DataFrame(items)

# Cargar y procesar la base de datos de partidas
df_base = pd.read_excel('base_datos_partidas.xlsx', header=2)
base_datos = procesar_template(df_base, es_base=True)
# Crear una columna con el texto normalizado para el matching
base_datos['Texto_completo_normalizado'] = base_datos['Texto_completo'].apply(normalizar_texto)

# Cargar y procesar el Excel de nuevas partidas
df_nuevas = pd.read_excel('nuevas_partidas.xlsx', header=2)
nuevas_partidas = procesar_template(df_nuevas, es_base=False)
nuevas_partidas['Texto_completo_normalizado'] = nuevas_partidas['Texto_completo'].apply(normalizar_texto)

# Función para encontrar la partida más parecida usando fuzzy matching
def encontrar_partida_parecida(descripcion, base_datos, umbral=70):
    lista = base_datos['Texto_completo_normalizado'].tolist()
    mejor_coincidencia, puntuacion, idx = process.extractOne(descripcion, lista, scorer=fuzz.token_sort_ratio)
    if puntuacion >= umbral:
        return mejor_coincidencia, base_datos.iloc[idx]['Precio']
    return None, None

# Aplicar la función de matching para asignar precios
nuevas_partidas['Precio Asignado'] = nuevas_partidas['Texto_completo_normalizado'].apply(
    lambda desc: encontrar_partida_parecida(desc, base_datos, umbral=60)[1] if pd.notna(desc) else 'No encontrado'
)

# Guardar el resultado en un nuevo archivo Excel
nuevas_partidas.to_excel('nuevas_partidas_con_precios_final.xlsx', index=False)

print("✅ Proceso completado. Revisa 'nuevas_partidas_con_precios_final.xlsx'")
