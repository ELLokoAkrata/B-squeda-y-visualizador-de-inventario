import pandas as pd
import streamlit as st

def preparar_datos(datos):
    # Asegúrate de que todas las entradas en las columnas que podrían tener tipos mixtos sean cadenas.
    columnas_a_string = ['DESCRIPCIÓN', 'MARCA', 'MODELO', 'P/N', 'S/N', 'OBSERVACIONES', 'STATUS', 'UBICACIÓN', 'MEDIDA']
    for columna in columnas_a_string:
        datos[columna] = datos[columna].astype(str)
    
    # Reemplaza las cadenas vacías generadas por la conversión a `str` con un valor que indique que no hay datos
    datos.replace({'': 'No disponible'}, inplace=True)

    # Asegúrate de que las columnas numéricas, como 'PRECIO UNIT' y 'TOTAL', no tengan valores no numéricos.
    columnas_numericas = ['CANT.', 'PRECIO UNIT', 'TOTAL']
    for columna in columnas_numericas:
        datos[columna] = pd.to_numeric(datos[columna], errors='coerce').fillna(0)
    
    return datos

def cargar_datos(archivo):
    if archivo is not None:
        if archivo.name.endswith('.xlsx'):
            datos = pd.read_excel(archivo, skiprows=3)  # Ajuste para encabezados en la fila 4
        elif archivo.name.endswith('.csv'):
            datos = pd.read_csv(archivo)
        datos = preparar_datos(datos)
        # Omitir columnas completamente vacías
        datos.dropna(axis=1, how='all', inplace=True)
        return datos
    return None

# Función para resaltar coincidencias en los datos (simplificada para Streamlit)
def resaltar_coincidencias(datos, texto_busqueda):
    if datos is not None and texto_busqueda != "":
        return datos[datos.apply(lambda row: texto_busqueda.lower() in row.to_string().lower(), axis=1)]
    return pd.DataFrame()

# Función para generar el resumen de la búsqueda
def resumen_busqueda(datos_filtrados):
    if datos_filtrados.empty:
        return "No se encontraron resultados."
    
    try:
        # Encuentra la cantidad máxima y mínima y sus respectivos ítems
        max_cant = datos_filtrados['CANT.'].max()
        min_cant = datos_filtrados['CANT.'].min()
        max_item_row = datos_filtrados[datos_filtrados['CANT.'] == max_cant].iloc[0]
        min_item_row = datos_filtrados[datos_filtrados['CANT.'] == min_cant].iloc[0]

        # Extrae los detalles del ítem con la cantidad máxima
        max_item = max_item_row['ITEM']
        max_precio = max_item_row['PRECIO UNIT']
        ubicacion_max = max_item_row['UBICACIÓN']
        status_max = max_item_row['STATUS']

        # Extrae los detalles del ítem con la cantidad mínima
        min_item = min_item_row['ITEM']
        min_precio = min_item_row['PRECIO UNIT']
        ubicacion_min = min_item_row['UBICACIÓN']
        status_min = min_item_row['STATUS']

        # Calcula el precio promedio y el rango de precios
        precio_promedio = datos_filtrados['PRECIO UNIT'].mean()
        rango_precios = (datos_filtrados['PRECIO UNIT'].min(), datos_filtrados['PRECIO UNIT'].max())

        # Crea el string de resumen
        resumen_str = f"""
El item con más productos es {max_item} con una cantidad de {max_cant} y un precio de {max_precio}. Ubicación: {ubicacion_max}, Status: {status_max}
El item con menos productos es {min_item} con una cantidad de {min_cant} y un precio de {min_precio}. Ubicación: {ubicacion_min}, Status: {status_min}
Total de productos encontrados: {len(datos_filtrados)}
Precio promedio: {precio_promedio:.2f}
Rango de precios: {rango_precios[0]} - {rango_precios[1]}
        """
    except Exception as e:
        resumen_str = f"Error al generar resumen: {e}"
    
    return resumen_str

# Streamlit App
def app():
    st.title("Conversor y Visualizador de Archivos")

    archivo = st.file_uploader("Cargar archivo XLSX/CSV", type=['xlsx', 'csv'])
    datos_originales = cargar_datos(archivo)

    if datos_originales is not None:
        st.write("Datos Cargados:")
        st.dataframe(datos_originales)

        texto_busqueda = st.text_input("Buscar texto en los datos:")
        if texto_busqueda:
            datos_filtrados = resaltar_coincidencias(datos_originales, texto_busqueda)
            if not datos_filtrados.empty:
                st.write("Resultados de la búsqueda:")
                st.dataframe(datos_filtrados)
                st.write("Resumen de la búsqueda:")
                st.info(resumen_busqueda(datos_filtrados))
            else:
                st.warning("No se encontraron coincidencias.")

if __name__ == "__main__":
    app()
