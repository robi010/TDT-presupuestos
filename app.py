import streamlit as st
import pandas as pd
import numpy as np
import unicodedata, re
from io import BytesIO
from sentence_transformers import SentenceTransformer
import faiss

st.title("Herramienta de Asignación de Precios para Presupuestos de Obra")
st.write("Sube el archivo Excel de la nueva obra y obtén un Excel con los precios asignados, comparando las partidas con la base de datos.")

# Opcional: permite subir la base de datos. Si no se sube, se cargará desde disco.
base_file = st.file_uploader("Sube la Base de Datos de Partidas (Excel)", type=["xlsx"], key="base")
if base_file is None:
    st.info("No se ha subido la base de datos; se intentará cargar 'base_datos_partidas.xlsx' desde el disco.")
    try:
        base_df = pd.read_excel("base_datos_partidas.xlsx", header=2)
    except Exception as e:
        st.error("Error al cargar 'base_datos_partidas.xlsx': " + str(e))
        st.stop()
else:
    base_df = pd.read_excel(base_file, header=2)

# Cargar el archivo Excel de la nueva obra.
new_file = st.file_uploader("Sube el archivo Excel de la Nueva Obra", type=["xlsx"], key="new")
if new_file is None:
    st.warning("Sube el archivo Excel de la nueva obra para continuar.")
    st.stop()
else:
    new_df = pd.read_excel(new_file, header=2)

# Función para procesar la plantilla: combina la fila con código (que tiene el título y precio)
# y la fila siguiente (con la descripción) en un único texto.
def procesar_template(df, es_base=True):
    items = []
    i = 0
    while i < len(df):
        row = df.iloc[i]
        codigo = row.get("Código")
        if pd.notna(codigo):
            title = row.get("Resumen") if pd.notna(row.get("Resumen")) else ""
            summary = ""
            if i + 1 < len(df):
                next_row = df.iloc[i + 1]
                if pd.isna(next_row.get("Código")):
                    summary = next_row.get("Resumen") if pd.notna(next_row.get("Resumen")) else ""
                    i += 1  # Se consume la fila de resumen
            texto_completo = (title + " " + summary).strip()
            precio = row.get("Pres") if es_base else None
            items.append({
                "Código": codigo,
                "Texto_completo": texto_completo,
                "Precio": precio
            })
        i += 1
    return pd.DataFrame(items)

# Procesar la base de datos y el archivo de nuevas partidas
base_data = procesar_template(base_df, es_base=True)
new_data = procesar_template(new_df, es_base=False)

# Función para normalizar textos: baja a minúsculas, quita acentos, elimina caracteres extra.
def normalizar_texto(texto):
    texto = str(texto)
    texto = texto.lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# Crear columnas normalizadas
base_data["Texto_normalizado"] = base_data["Texto_completo"].apply(normalizar_texto)
new_data["Texto_normalizado"] = new_data["Texto_completo"].apply(normalizar_texto)

st.write("Procesando datos...")

# Cargar el modelo de Sentence Transformers
with st.spinner("Cargando modelo de embeddings..."):
    model = SentenceTransformer('all-MiniLM-L6-v2')

# Generar embeddings para la base de datos
with st.spinner("Generando embeddings para la base de datos..."):
    base_texts = base_data["Texto_normalizado"].tolist()
    base_embeddings = model.encode(base_texts, convert_to_numpy=True)
    # Normalizar para similitud coseno
    faiss.normalize_L2(base_embeddings)

# Crear índice FAISS
d = base_embeddings.shape[1]
index = faiss.IndexFlatIP(d)
index.add(base_embeddings)

# Generar embeddings para las nuevas partidas
with st.spinner("Generando embeddings para las nuevas partidas..."):
    new_texts = new_data["Texto_normalizado"].tolist()
    new_embeddings = model.encode(new_texts, convert_to_numpy=True)
    faiss.normalize_L2(new_embeddings)

# Umbral de similitud ajustable (de 0 a 1)
threshold = st.slider("Umbral de similitud", 0.0, 1.0, 0.7, step=0.01)

# Realizar la búsqueda y asignar precios
assigned_prices = []
for emb in new_embeddings:
    emb = emb.reshape(1, -1)
    D, I = index.search(emb, k=1)
    best_score = D[0][0]
    if best_score >= threshold:
        price = base_data.iloc[I[0][0]]["Precio"]
    else:
        price = "No encontrado"
    assigned_prices.append(price)

new_data["Precio Asignado"] = assigned_prices

st.write("Proceso completado. Vista previa de los resultados:")
st.dataframe(new_data)

# Opción para descargar el resultado en Excel
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    new_data.to_excel(writer, index=False, sheet_name="Resultados")
    writer.save()
output.seek(0)

st.download_button(
    label="Descargar archivo Excel con precios asignados",
    data=output,
    file_name="nuevas_partidas_con_precios.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
