import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# Cuántas filas saltar antes de la cabecera real
HEADER_OFFSET = 3

# ---------- LIMPIEZA CENTRALIZADA ---------- #
def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Elimina columnas índice fantasma (Unnamed:*), resetea el índice
    y fuerza columnas object → str para que Arrow no reviente.
    """
    # Borra columnas tipo Unnamed y crea copia segura
    df = df.loc[:, ~df.columns.str.contains(r'^Unnamed')].copy()
    # Índice limpio
    df.reset_index(drop=True, inplace=True)
    # Todas las object como string plano sin applymap
    obj_cols = df.select_dtypes(include=['object']).columns
    df.loc[:, obj_cols] = (
        df[obj_cols]
          .fillna('')
          .astype(str)
    )
    return df

# ---------- PREPARAR DATOS ---------- #
def preparar_datos(datos):
    columnas_a_string = ['DESCRIPCIÓN', 'MARCA', 'MODELO', 'P/N', 'S/N',
                         'OBSERVACIONES', 'STATUS', 'UBICACIÓN', 'MEDIDA']
    for columna in columnas_a_string:
        datos[columna] = datos[columna].astype(str)
    datos.replace({'': 'No disponible'}, inplace=True)

    columnas_numericas = ['CANT.', 'PRECIO UNIT', 'TOTAL']
    for columna in columnas_numericas:
        datos[columna] = pd.to_numeric(datos[columna], errors='coerce').fillna(0)

    return datos

# ---------- CARGAR ARCHIVO ---------- #
def cargar_datos(archivo):
    if archivo is None:
        return None

    if archivo.name.endswith('.xlsx'):
        excel_file = pd.ExcelFile(archivo)
        sheet_name = excel_file.sheet_names[0]
        st.session_state['hoja'] = sheet_name
        datos = pd.read_excel(excel_file, sheet_name=sheet_name,
                             skiprows=HEADER_OFFSET, dtype=str)
    elif archivo.name.endswith('.csv'):
        st.session_state['hoja'] = None
        datos = pd.read_csv(archivo, dtype=str)
    else:
        return None

    datos.columns = datos.columns.str.strip()
    datos = preparar_datos(datos)
    datos.dropna(axis=1, how='all', inplace=True)

    return clean_df(datos)

# ---------- CONVERTIR A EXCEL/CSV ---------- #
def convertir_a_excel(df, original_bytes=None, filas_originales=0, sheet_name=None):
    if original_bytes is not None:
        wb = load_workbook(filename=BytesIO(original_bytes))
        ws = wb[sheet_name] if sheet_name else wb.active

        # Borra filas vacías finales
        ultima_fila = ws.max_row
        for row_idx in range(ws.max_row, 0, -1):
            if any(cell.value not in (None, "") for cell in ws[row_idx]):
                break
            ultima_fila = row_idx - 1
        if ultima_fila < ws.max_row:
            ws.delete_rows(ultima_fila + 1, ws.max_row - ultima_fila)

        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))

        for row in df.iloc[filas_originales:].fillna("").itertuples(index=False):
            ws.append(list(row))
            for col_idx, _ in enumerate(row, start=1):
                ws.cell(row=ws.max_row, column=col_idx).border = border

        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.fillna("").to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

# ---------- BÚSQUEDA Y RESUMEN ---------- #
def resaltar_coincidencias(datos, texto_busqueda):
    if datos is not None and texto_busqueda:
        return datos[datos.apply(lambda row: texto_busqueda.lower() in row.to_string().lower(), axis=1)]
    return pd.DataFrame()

def resumen_busqueda(datos_filtrados):
    if datos_filtrados.empty:
        return "No se encontraron resultados."
    try:
        max_cant = datos_filtrados['CANT.'].max()
        min_cant = datos_filtrados['CANT.'].min()
        max_row = datos_filtrados[datos_filtrados['CANT.'] == max_cant].iloc[0]
        min_row = datos_filtrados[datos_filtrados['CANT.'] == min_cant].iloc[0]

        precio_promedio = datos_filtrados['PRECIO UNIT'].mean()
        rango_precios = (datos_filtrados['PRECIO UNIT'].min(), datos_filtrados['PRECIO UNIT'].max())

        return (
            f"Item con más productos: {max_row['ITEM']} | Cant: {max_cant} | Precio: {max_row['PRECIO UNIT']} "
            f"| Ubicación: {max_row['UBICACIÓN']} | Status: {max_row['STATUS']}\n"
            f"Item con menos productos: {min_row['ITEM']} | Cant: {min_cant} | Precio: {min_row['PRECIO UNIT']} "
            f"| Ubicación: {min_row['UBICACIÓN']} | Status: {min_row['STATUS']}\n"
            f"Total de productos: {len(datos_filtrados)}\n"
            f"Precio promedio: {precio_promedio:.2f}\n"
            f"Rango de precios: {rango_precios[0]} - {rango_precios[1]}"
        )
    except Exception as e:
        return f"Error al generar resumen: {e}"

# ---------- APP PRINCIPAL ---------- #
def app():
    st.title("Visualizador, buscador y generador de resumen de inventario")
    st.write("""
    Sube un inventario (CSV o XLSX), visualiza, filtra y descarga la versión actualizada.
    """)

    if 'datos' not in st.session_state:
        st.session_state.update({
            'datos': None,
            'archivo_bytes': None,
            'extension': None,
            'filas_originales': 0,
            'hoja': None,
        })

    archivo = st.file_uploader("Cargar archivo XLSX/CSV", type=['xlsx', 'csv'])
    if archivo:
        st.session_state['datos'] = cargar_datos(archivo)
        st.session_state['archivo_bytes'] = archivo.getvalue()
        st.session_state['extension'] = '.xlsx' if archivo.name.endswith('.xlsx') else '.csv'
        st.session_state['filas_originales'] = len(st.session_state['datos'])

        with st.expander("Previsualizar archivo original cargado"):
            if st.session_state['extension'] == '.xlsx':
                original_df = pd.read_excel(
                    BytesIO(st.session_state['archivo_bytes']),
                    sheet_name=st.session_state['hoja'],
                    skiprows=HEADER_OFFSET,
                    dtype=str
                )
            else:
                original_df = pd.read_csv(
                    BytesIO(st.session_state['archivo_bytes']), dtype=str
                )

            st.dataframe(clean_df(original_df))

    datos = st.session_state['datos']

    if datos is not None:
        st.write("Datos cargados:")
        st.dataframe(clean_df(datos))  # limpio antes de mostrar

        # ---------- FORM NUEVO ITEM ---------- #
        with st.expander("Añadir nuevo item"):
            with st.form("form-nuevo-item"):
                col1, col2 = st.columns(2)
                nuevo = {
                    'ITEM': col1.text_input('ITEM'),
                    'DESCRIPCIÓN': col1.text_input('DESCRIPCIÓN'),
                    'MARCA': col1.text_input('MARCA'),
                    'MODELO': col1.text_input('MODELO'),
                    'P/N': col1.text_input('P/N'),
                    'S/N': col1.text_input('S/N'),
                    'OBSERVACIONES': col2.text_input('OBSERVACIONES'),
                    'STATUS': col2.text_input('STATUS'),
                    'UBICACIÓN': col2.text_input('UBICACIÓN'),
                    'MEDIDA': col2.text_input('MEDIDA'),
                    'CANT.': col1.number_input('CANT.', value=0, step=1),
                    'PRECIO UNIT': col1.number_input('PRECIO UNIT', value=0.0),
                    'TOTAL': col1.number_input('TOTAL', value=0.0),
                }
                enviado = st.form_submit_button('Añadir')

            if enviado:
                st.session_state['datos'] = clean_df(pd.concat(
                    [datos, pd.DataFrame([nuevo])], ignore_index=True))
                st.success(f'Item añadido (fila {len(st.session_state["datos"])})')
                datos = st.session_state['datos']

        # ---------- BÚSQUEDA ---------- #
        texto_busqueda = st.text_input("Buscar texto:")
        if texto_busqueda:
            filtrado = resaltar_coincidencias(datos, texto_busqueda)
            if filtrado.empty:
                st.warning("No se encontraron coincidencias.")
            else:
                st.dataframe(clean_df(filtrado))
                st.info(resumen_busqueda(filtrado))

        # ---------- DESCARGA ---------- #
        if st.session_state['extension'] == '.xlsx':
            buffer = convertir_a_excel(
                datos,
                st.session_state['archivo_bytes'],
                st.session_state['filas_originales'],
                sheet_name=st.session_state['hoja'],
            )
            fname = "inventario_actualizado.xlsx"
        else:
            buffer = BytesIO()
            clean_df(datos).to_csv(buffer, index=False)
            buffer.seek(0)
            fname = "inventario_actualizado.csv"

        # ---------- PREVISUALIZACIÓN ---------- #
        with st.expander("Previsualizar archivo generado"):
            if st.session_state['extension'] == '.xlsx':
                prev = pd.read_excel(
                    BytesIO(buffer.getvalue()),
                    sheet_name=st.session_state['hoja'],
                    skiprows=HEADER_OFFSET,
                    dtype=str
                )
            else:
                prev = pd.read_csv(BytesIO(buffer.getvalue()), dtype=str)
            st.dataframe(clean_df(prev))

        st.download_button("Descargar inventario actualizado",
                           data=buffer, file_name=fname)

if __name__ == "__main__":
    app()

