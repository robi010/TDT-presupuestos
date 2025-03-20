import pandas as pd
from difflib import get_close_matches

# Función para extraer las partidas de la base de datos con la estructura específica
def extraer_partidas_construccion(file_path):
    # Leer el archivo Excel original
    df = pd.read_excel(file_path)

    # Crear listas para almacenar la información estructurada
    partidas = []
    descripciones = []
    precios = []

    # Recorrer el archivo fila por fila
    for i, row in df.iterrows():
        # Si encontramos la palabra "Partida" en la columna B
        if row['B'] == 'Partida':
            # Extraer el nombre de la partida (columna D en la misma fila)
            nombre_partida = row['D']
            # Extraer el resumen (columna D en la fila siguiente)
            resumen_partida = df.at[i + 1, 'D'] if (i + 1) in df.index else ''
            # Extraer el precio (columna L en la misma fila)
            precio_partida = row['L']
            
            # Unificar el nombre y resumen como una descripción única
            descripcion = f"{nombre_partida} - {resumen_partida}"
            
            # Agregar la información a las listas
            partidas.append(nombre_partida)
            descripciones.append(descripcion)
            precios.append(precio_partida)

    # Crear un DataFrame con la información estructurada
    base_datos_df = pd.DataFrame({
        'Descripción': descripciones,
        'Precio': precios
    })
    
    return base_datos_df

# Función para encontrar la partida más parecida
def encontrar_partida_parecida(descripcion, base_datos):
    lista_descripciones = base_datos['Descripción'].tolist()
    partidas_parecidas = get_close_matches(descripcion, lista_descripciones, n=1, cutoff=0.6)
    if partidas_parecidas:
        partida = partidas_parecidas[0]
        precio = base_datos[base_datos['Descripción'] == partida]['Precio'].values[0]
        return partida, precio
    else:
        return None, None

# Cargar y estructurar la base de datos original
base_datos = extraer_partidas_construccion('base_datos_partidas.xlsx')

# Cargar el archivo de nuevas partidas
nuevas_partidas = pd.read_excel('nuevas_partidas.xlsx')

# Crear una nueva columna para el precio en las nuevas partidas
nuevas_partidas['Precio Asignado'] = None

# Comparar cada nueva partida con las de la base de datos
for i, row in nuevas_partidas.iterrows():
    descripcion = row['Descripción']
    partida_parecida, precio = encontrar_partida_parecida(descripcion, base_datos)
    
    if partida_parecida:
        nuevas_partidas.at[i, 'Precio Asignado'] = precio
    else:
        nuevas_partidas.at[i, 'Precio Asignado'] = 'No encontrado'

# Guardar el resultado en un nuevo archivo Excel
nuevas_partidas.to_excel('nuevas_partidas_con_precios.xlsx', index=False)

print("Proceso completado. El archivo 'nuevas_partidas_con_precios.xlsx' ha sido creado.")
