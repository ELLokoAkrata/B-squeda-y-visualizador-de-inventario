import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import json
import toml

# Cuántas filas saltar antes de la cabecera real
HEADER_OFFSET = 3

# Columnas requeridas para validación
COLUMNAS_REQUERIDAS = ['ITEM', 'DESCRIPCIÓN', 'CANT.', 'PRECIO UNIT']

# Nuevas columnas extendidas para gestión profesional
COLUMNAS_EXTENDIDAS = ['CATEGORÍA', 'STOCK MÍNIMO', 'FECHA_CREACIÓN', 'FECHA_MODIFICACIÓN']

# ---------- LIMPIEZA CENTRALIZADA ---------- #
def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.str.contains(r'^Unnamed')].copy()
    df.reset_index(drop=True, inplace=True)
    obj_cols = df.select_dtypes(include=['object']).columns
    df.loc[:, obj_cols] = df[obj_cols].fillna('').astype(str)
    return df

# ---------- PREPARAR DATOS ---------- #
def preparar_datos(datos):
    columnas_a_string = ['DESCRIPCIÓN', 'MARCA', 'MODELO', 'P/N', 'S/N',
                         'OBSERVACIONES', 'STATUS', 'UBICACIÓN', 'MEDIDA']
    for columna in columnas_a_string:
        if columna in datos.columns:
            datos[columna] = datos[columna].astype(str)
    datos.replace({'': 'No disponible'}, inplace=True)

    columnas_numericas = ['CANT.', 'PRECIO UNIT', 'TOTAL']
    for columna in columnas_numericas:
        if columna in datos.columns:
            datos[columna] = pd.to_numeric(datos[columna], errors='coerce').fillna(0)

    # Agregar columnas extendidas si no existen
    if 'CATEGORÍA' not in datos.columns:
        datos['CATEGORÍA'] = 'Sin categoría'
    if 'STOCK MÍNIMO' not in datos.columns:
        datos['STOCK MÍNIMO'] = 0
    if 'FECHA_CREACIÓN' not in datos.columns:
        datos['FECHA_CREACIÓN'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if 'FECHA_MODIFICACIÓN' not in datos.columns:
        datos['FECHA_MODIFICACIÓN'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Convertir columnas extendidas a los tipos correctos
    datos['CATEGORÍA'] = datos['CATEGORÍA'].astype(str).replace('', 'Sin categoría')
    datos['STOCK MÍNIMO'] = pd.to_numeric(datos['STOCK MÍNIMO'], errors='coerce').fillna(0)

    return datos

# ---------- VALIDACIONES ---------- #
def validar_item_unico(df, nuevo_item):
    """Valida que el ITEM sea único en el inventario"""
    if nuevo_item in df['ITEM'].values:
        return False, f"⚠️ El ITEM '{nuevo_item}' ya existe en el inventario"
    return True, ""

def validar_campos_requeridos(nuevo_item_dict):
    """Valida que los campos requeridos no estén vacíos"""
    errores = []
    for campo in COLUMNAS_REQUERIDAS:
        if campo in nuevo_item_dict:
            valor = str(nuevo_item_dict[campo]).strip()
            if not valor or valor == '0' or valor == '0.0':
                if campo in ['CANT.', 'PRECIO UNIT']:
                    continue  # Permitimos 0 en campos numéricos
                errores.append(f"El campo '{campo}' es requerido")
    return len(errores) == 0, errores

def validar_valores_positivos(cant, precio_unit, stock_minimo=0):
    """Valida que cantidad, precio y stock mínimo no sean negativos"""
    errores = []
    if cant < 0:
        errores.append("La cantidad no puede ser negativa")
    if precio_unit < 0:
        errores.append("El precio unitario no puede ser negativo")
    if stock_minimo < 0:
        errores.append("El stock mínimo no puede ser negativo")
    return len(errores) == 0, errores

def calcular_total(cant, precio_unit):
    """Calcula automáticamente el TOTAL"""
    return cant * precio_unit

# ---------- CLASIFICACIÓN ABC (PARETO) ---------- #
def clasificacion_abc(df):
    """
    Clasifica items según análisis de Pareto:
    - Clase A: 80% del valor (top items)
    - Clase B: siguiente 15% del valor
    - Clase C: último 5% del valor
    """
    if df.empty:
        return df

    # Calcular valor total por item
    df_sorted = df.copy()
    df_sorted['VALOR_TOTAL_ITEM'] = df_sorted['CANT.'] * df_sorted['PRECIO UNIT']
    df_sorted = df_sorted.sort_values('VALOR_TOTAL_ITEM', ascending=False)

    # Calcular porcentaje acumulado
    total_valor = df_sorted['VALOR_TOTAL_ITEM'].sum()
    if total_valor == 0:
        df_sorted['CLASE_ABC'] = 'C'
        return df_sorted

    df_sorted['PORCENTAJE_ACUMULADO'] = (df_sorted['VALOR_TOTAL_ITEM'].cumsum() / total_valor) * 100

    # Clasificar
    df_sorted['CLASE_ABC'] = df_sorted['PORCENTAJE_ACUMULADO'].apply(
        lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
    )

    return df_sorted

# ---------- ALERTAS DE STOCK BAJO ---------- #
def obtener_items_bajo_stock(df):
    """Retorna items donde CANT. <= STOCK MÍNIMO"""
    if df.empty:
        return pd.DataFrame()
    return df[df['CANT.'] <= df['STOCK MÍNIMO']]

# ---------- CALCULAR KPIs ---------- #
def calcular_kpis(df):
    """Calcula indicadores clave de rendimiento del inventario"""
    if df.empty:
        return {
            'total_items': 0,
            'valor_total': 0,
            'items_bajo_stock': 0,
            'categorias': 0,
            'ubicaciones': 0,
            'precio_promedio': 0,
            'cantidad_total': 0
        }

    return {
        'total_items': len(df),
        'valor_total': (df['CANT.'] * df['PRECIO UNIT']).sum(),
        'items_bajo_stock': len(obtener_items_bajo_stock(df)),
        'categorias': df['CATEGORÍA'].nunique(),
        'ubicaciones': df['UBICACIÓN'].nunique() if 'UBICACIÓN' in df.columns else 0,
        'precio_promedio': df['PRECIO UNIT'].mean(),
        'cantidad_total': df['CANT.'].sum()
    }

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

# ---------- EXPORTAR A JSON ---------- #
def exportar_a_json(df, formato='registros'):
    """
    Exporta el DataFrame a JSON en diferentes formatos

    Args:
        df: DataFrame a exportar
        formato: 'registros' (lista de objetos), 'columnas' (objeto de columnas), 'valores' (solo valores)

    Returns:
        BytesIO con el contenido JSON
    """
    # Convertir valores numéricos de pandas a tipos nativos de Python para serialización JSON
    df_export = df.copy()

    # Convertir tipos numéricos a float/int nativos
    for col in df_export.select_dtypes(include=['number']).columns:
        df_export[col] = df_export[col].astype(float)

    # Opciones de formato
    if formato == 'registros':
        # Formato: [{"ITEM": "001", "DESCRIPCIÓN": "..."}, ...]
        json_data = df_export.to_dict(orient='records')
    elif formato == 'columnas':
        # Formato: {"ITEM": ["001", "002"], "DESCRIPCIÓN": ["...", "..."]}
        json_data = df_export.to_dict(orient='list')
    elif formato == 'valores':
        # Formato: [["001", "...", ...], ["002", "...", ...]]
        json_data = df_export.values.tolist()
    else:
        json_data = df_export.to_dict(orient='records')

    # Crear estructura con metadata
    output = {
        "metadata": {
            "total_items": len(df_export),
            "fecha_exportacion": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "formato": formato,
            "columnas": list(df_export.columns)
        },
        "inventario": json_data
    }

    # Convertir a JSON con formato legible
    json_string = json.dumps(output, ensure_ascii=False, indent=2)

    # Crear buffer
    buffer = BytesIO()
    buffer.write(json_string.encode('utf-8'))
    buffer.seek(0)

    return buffer

# ---------- EXPORTAR A TOML ---------- #
def exportar_a_toml(df):
    """
    Exporta el DataFrame a formato TOML

    Args:
        df: DataFrame a exportar

    Returns:
        BytesIO con el contenido TOML
    """
    # Crear estructura TOML
    toml_data = {
        "metadata": {
            "total_items": len(df),
            "fecha_exportacion": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "columnas": list(df.columns)
        },
        "inventario": {}
    }

    # Convertir cada fila a una sección TOML usando ITEM como clave
    for idx, row in df.iterrows():
        item_key = str(row.get('ITEM', f'item_{idx}')).replace(' ', '_').replace('.', '_')

        item_data = {}
        for col in df.columns:
            valor = row[col]

            # Convertir valores de pandas a tipos nativos de Python
            if pd.isna(valor):
                item_data[col] = None
            elif isinstance(valor, (int, float)):
                item_data[col] = float(valor)
            else:
                item_data[col] = str(valor)

        toml_data["inventario"][item_key] = item_data

    # Convertir a string TOML
    toml_string = toml.dumps(toml_data)

    # Crear buffer
    buffer = BytesIO()
    buffer.write(toml_string.encode('utf-8'))
    buffer.seek(0)

    return buffer

# ---------- FILTROS AVANZADOS ---------- #
def aplicar_filtros(df, categoria=None, ubicacion=None, status=None,
                    precio_min=None, precio_max=None, cant_min=None, cant_max=None):
    """Aplica múltiples filtros al DataFrame"""
    df_filtrado = df.copy()

    if categoria and categoria != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['CATEGORÍA'] == categoria]

    if ubicacion and ubicacion != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['UBICACIÓN'] == ubicacion]

    if status and status != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['STATUS'] == status]

    if precio_min is not None:
        df_filtrado = df_filtrado[df_filtrado['PRECIO UNIT'] >= precio_min]

    if precio_max is not None:
        df_filtrado = df_filtrado[df_filtrado['PRECIO UNIT'] <= precio_max]

    if cant_min is not None:
        df_filtrado = df_filtrado[df_filtrado['CANT.'] >= cant_min]

    if cant_max is not None:
        df_filtrado = df_filtrado[df_filtrado['CANT.'] <= cant_max]

    return df_filtrado

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

# ---------- VISUALIZACIONES ---------- #
def crear_grafico_valor_por_categoria(df):
    """Crea gráfico de valor total por categoría"""
    if df.empty:
        return None

    df_temp = df.copy()
    df_temp['VALOR_TOTAL'] = df_temp['CANT.'] * df_temp['PRECIO UNIT']
    valor_por_categoria = df_temp.groupby('CATEGORÍA')['VALOR_TOTAL'].sum().reset_index()
    valor_por_categoria = valor_por_categoria.sort_values('VALOR_TOTAL', ascending=False)

    fig = px.bar(valor_por_categoria, x='CATEGORÍA', y='VALOR_TOTAL',
                 title='Valor Total del Inventario por Categoría',
                 labels={'VALOR_TOTAL': 'Valor Total', 'CATEGORÍA': 'Categoría'},
                 color='VALOR_TOTAL',
                 color_continuous_scale='Blues')
    fig.update_layout(showlegend=False)
    return fig

def crear_grafico_distribucion_abc(df):
    """Crea gráfico de distribución ABC"""
    df_abc = clasificacion_abc(df)
    if df_abc.empty:
        return None

    distribucion = df_abc['CLASE_ABC'].value_counts().reset_index()
    distribucion.columns = ['Clase', 'Cantidad']

    colores = {'A': '#2ecc71', 'B': '#f39c12', 'C': '#e74c3c'}
    fig = px.pie(distribucion, values='Cantidad', names='Clase',
                 title='Clasificación ABC de Inventario (Análisis de Pareto)',
                 color='Clase',
                 color_discrete_map=colores)
    return fig

def crear_grafico_items_por_ubicacion(df):
    """Crea gráfico de cantidad de items por ubicación"""
    if df.empty or 'UBICACIÓN' not in df.columns:
        return None

    items_por_ubicacion = df.groupby('UBICACIÓN').size().reset_index()
    items_por_ubicacion.columns = ['Ubicación', 'Cantidad']
    items_por_ubicacion = items_por_ubicacion.sort_values('Cantidad', ascending=False)

    fig = px.bar(items_por_ubicacion, x='Ubicación', y='Cantidad',
                 title='Distribución de Items por Ubicación',
                 labels={'Cantidad': 'Número de Items', 'Ubicación': 'Ubicación'},
                 color='Cantidad',
                 color_continuous_scale='Viridis')
    fig.update_layout(showlegend=False)
    return fig

def crear_grafico_top_items_valor(df, top_n=10):
    """Crea gráfico de top N items por valor total"""
    if df.empty:
        return None

    df_temp = df.copy()
    df_temp['VALOR_TOTAL'] = df_temp['CANT.'] * df_temp['PRECIO UNIT']
    top_items = df_temp.nlargest(top_n, 'VALOR_TOTAL')[['ITEM', 'DESCRIPCIÓN', 'VALOR_TOTAL']]

    fig = px.bar(top_items, x='VALOR_TOTAL', y='ITEM',
                 orientation='h',
                 title=f'Top {top_n} Items por Valor Total',
                 labels={'VALOR_TOTAL': 'Valor Total', 'ITEM': 'Item'},
                 color='VALOR_TOTAL',
                 color_continuous_scale='Reds',
                 hover_data=['DESCRIPCIÓN'])
    fig.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
    return fig

# ---------- APP PRINCIPAL ---------- #
def app():
    st.set_page_config(page_title="Sistema de Inventario Profesional", layout="wide")
    st.title("📦 Sistema Profesional de Gestión de Inventario")
    st.markdown("---")

    # Inicializar session state
    if 'datos' not in st.session_state:
        st.session_state.update({
            'datos': None,
            'archivo_bytes': None,
            'extension': None,
            'filas_originales': 0,
            'hoja': None,
        })

    # Sidebar para cargar archivo
    with st.sidebar:
        st.header("📁 Cargar Inventario")
        archivo = st.file_uploader("Cargar archivo XLSX/CSV", type=['xlsx', 'csv'])

        if archivo:
            st.session_state['datos'] = cargar_datos(archivo)
            st.session_state['archivo_bytes'] = archivo.getvalue()
            st.session_state['extension'] = '.xlsx' if archivo.name.endswith('.xlsx') else '.csv'
            st.session_state['filas_originales'] = len(st.session_state['datos'])
            st.success(f"✅ Archivo cargado: {len(st.session_state['datos'])} items")

        st.markdown("---")

        # Mostrar KPIs en sidebar si hay datos
        if st.session_state['datos'] is not None:
            kpis = calcular_kpis(st.session_state['datos'])
            st.metric("Total Items", kpis['total_items'])
            st.metric("Valor Total", f"${kpis['valor_total']:,.2f}")
            st.metric("Items Bajo Stock", kpis['items_bajo_stock'],
                     delta=None if kpis['items_bajo_stock'] == 0 else "⚠️")
            st.metric("Categorías", kpis['categorias'])

    datos = st.session_state['datos']

    # Si no hay datos, mostrar instrucciones
    if datos is None:
        st.info("👈 Por favor, carga un archivo de inventario desde el panel lateral para comenzar.")
        st.markdown("""
        ### 📋 Características del Sistema Profesional de Inventario:

        **✅ Validaciones Automáticas:**
        - Validación de unicidad de items
        - Validación de stocks y precios no negativos
        - Cálculo automático de totales
        - Campos requeridos obligatorios

        **📊 Analíticas Avanzadas:**
        - Dashboard con KPIs en tiempo real
        - Clasificación ABC (Análisis de Pareto)
        - Alertas de stock bajo
        - Visualizaciones interactivas

        **🔍 Búsqueda y Filtros:**
        - Búsqueda global de texto
        - Filtros por categoría, ubicación y status
        - Filtros por rangos de precio y cantidad
        - Exportación de resultados filtrados

        **📈 Reportes Profesionales:**
        - Análisis de valor por categoría
        - Distribución ABC
        - Top items por valor
        - Distribución por ubicación

        **🕐 Auditoría:**
        - Timestamps de creación y modificación
        - Trazabilidad de cambios
        """)
        return

    # TABS PRINCIPALES
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Dashboard",
        "📦 Inventario Completo",
        "🔍 Búsqueda Avanzada",
        "📈 Reportes & Analíticas",
        "➕ Agregar Item"
    ])

    # ==================== TAB 1: DASHBOARD ====================
    with tab1:
        st.header("📊 Dashboard de Inventario")

        # KPIs principales
        kpis = calcular_kpis(datos)
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total de Items", kpis['total_items'])
            st.metric("Cantidad Total", f"{kpis['cantidad_total']:,.0f}")
        with col2:
            st.metric("Valor Total Inventario", f"${kpis['valor_total']:,.2f}")
            st.metric("Precio Promedio", f"${kpis['precio_promedio']:,.2f}")
        with col3:
            st.metric("Items Bajo Stock", kpis['items_bajo_stock'],
                     delta="Crítico" if kpis['items_bajo_stock'] > 0 else "OK",
                     delta_color="inverse")
            st.metric("Categorías", kpis['categorias'])
        with col4:
            st.metric("Ubicaciones", kpis['ubicaciones'])

        st.markdown("---")

        # Alertas de stock bajo
        items_bajo_stock = obtener_items_bajo_stock(datos)
        if not items_bajo_stock.empty:
            st.warning(f"⚠️ **ALERTA:** {len(items_bajo_stock)} items con stock bajo o crítico")
            with st.expander("Ver items con stock bajo", expanded=True):
                st.dataframe(
                    items_bajo_stock[['ITEM', 'DESCRIPCIÓN', 'CANT.', 'STOCK MÍNIMO',
                                     'UBICACIÓN', 'CATEGORÍA']],
                    use_container_width=True
                )
        else:
            st.success("✅ Todos los items tienen stock adecuado")

        st.markdown("---")

        # Gráficos del dashboard
        col_left, col_right = st.columns(2)

        with col_left:
            fig_categoria = crear_grafico_valor_por_categoria(datos)
            if fig_categoria:
                st.plotly_chart(fig_categoria, use_container_width=True)

        with col_right:
            fig_abc = crear_grafico_distribucion_abc(datos)
            if fig_abc:
                st.plotly_chart(fig_abc, use_container_width=True)

    # ==================== TAB 2: INVENTARIO COMPLETO ====================
    with tab2:
        st.header("📦 Inventario Completo")

        # Opciones de visualización
        col_order, col_display = st.columns([2, 1])
        with col_order:
            ordenar_por = st.selectbox(
                "Ordenar por:",
                ['ITEM', 'DESCRIPCIÓN', 'CATEGORÍA', 'CANT.', 'PRECIO UNIT',
                 'TOTAL', 'UBICACIÓN', 'FECHA_MODIFICACIÓN']
            )
            orden_desc = st.checkbox("Orden descendente", value=False)

        with col_display:
            mostrar_abc = st.checkbox("Mostrar clasificación ABC", value=False)

        # Mostrar datos
        datos_display = datos.copy()
        if mostrar_abc:
            datos_display = clasificacion_abc(datos_display)

        datos_display = datos_display.sort_values(by=ordenar_por, ascending=not orden_desc)
        st.dataframe(datos_display, use_container_width=True, height=600)

        # Botón de descarga
        st.markdown("---")
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

        st.download_button(
            "⬇️ Descargar inventario completo actualizado (XLSX/CSV)",
            data=buffer,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Exportar a otros formatos
        st.markdown("### 📤 Exportar a Otros Formatos")

        col_json, col_toml = st.columns(2)

        with col_json:
            st.markdown("**JSON**")
            formato_json = st.selectbox(
                "Formato JSON:",
                ['registros', 'columnas', 'valores'],
                help="registros: lista de objetos | columnas: objeto de arrays | valores: matriz de valores"
            )

            buffer_json = exportar_a_json(datos, formato=formato_json)

            st.download_button(
                "⬇️ Descargar JSON",
                data=buffer_json,
                file_name="inventario.json",
                mime="application/json"
            )

            # Previsualizar JSON
            with st.expander("👁️ Previsualizar JSON (primeros 5 items)"):
                df_preview = datos.head(5)
                json_preview = exportar_a_json(df_preview, formato=formato_json)
                st.code(json_preview.getvalue().decode('utf-8'), language='json')

        with col_toml:
            st.markdown("**TOML**")
            st.info("Formato TOML para archivos de configuración")

            buffer_toml = exportar_a_toml(datos)

            st.download_button(
                "⬇️ Descargar TOML",
                data=buffer_toml,
                file_name="inventario.toml",
                mime="application/toml"
            )

            # Previsualizar TOML
            with st.expander("👁️ Previsualizar TOML (primeros 3 items)"):
                df_preview_toml = datos.head(3)
                toml_preview = exportar_a_toml(df_preview_toml)
                st.code(toml_preview.getvalue().decode('utf-8'), language='toml')

    # ==================== TAB 3: BÚSQUEDA AVANZADA ====================
    with tab3:
        st.header("🔍 Búsqueda y Filtros Avanzados")

        # Búsqueda por texto
        texto_busqueda = st.text_input("🔎 Búsqueda global por texto:",
                                       placeholder="Buscar en todas las columnas...")

        st.markdown("---")
        st.subheader("Filtros Avanzados")

        col1, col2, col3 = st.columns(3)

        with col1:
            categorias = ['Todas'] + sorted(datos['CATEGORÍA'].unique().tolist())
            categoria_filtro = st.selectbox("Categoría:", categorias)

            ubicaciones = ['Todas'] + sorted(datos['UBICACIÓN'].unique().tolist())
            ubicacion_filtro = st.selectbox("Ubicación:", ubicaciones)

        with col2:
            status_list = ['Todos'] + sorted(datos['STATUS'].unique().tolist())
            status_filtro = st.selectbox("Status:", status_list)

            precio_min = st.number_input("Precio mínimo:", min_value=0.0, value=0.0)

        with col3:
            precio_max = st.number_input("Precio máximo:", min_value=0.0,
                                        value=float(datos['PRECIO UNIT'].max()))
            cant_min = st.number_input("Cantidad mínima:", min_value=0, value=0)

        # Aplicar filtros
        datos_filtrados = aplicar_filtros(
            datos,
            categoria=categoria_filtro,
            ubicacion=ubicacion_filtro,
            status=status_filtro,
            precio_min=precio_min if precio_min > 0 else None,
            precio_max=precio_max,
            cant_min=cant_min if cant_min > 0 else None
        )

        # Aplicar búsqueda de texto si existe
        if texto_busqueda:
            datos_filtrados = resaltar_coincidencias(datos_filtrados, texto_busqueda)

        st.markdown("---")

        # Mostrar resultados
        if datos_filtrados.empty:
            st.warning("⚠️ No se encontraron resultados con los filtros aplicados.")
        else:
            st.success(f"✅ Se encontraron {len(datos_filtrados)} items")

            # Resumen estadístico
            if texto_busqueda:
                st.info(resumen_busqueda(datos_filtrados))

            # Mostrar resultados
            st.dataframe(datos_filtrados, use_container_width=True, height=400)

            # Exportar resultados filtrados
            st.markdown("### 📥 Exportar Resultados Filtrados")

            col_excel, col_json_filt, col_toml_filt = st.columns(3)

            with col_excel:
                buffer_filtrado = BytesIO()
                datos_filtrados.to_excel(buffer_filtrado, index=False, engine='openpyxl')
                buffer_filtrado.seek(0)

                st.download_button(
                    "⬇️ Excel",
                    data=buffer_filtrado,
                    file_name="inventario_filtrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_json_filt:
                buffer_json_filt = exportar_a_json(datos_filtrados, formato='registros')
                st.download_button(
                    "⬇️ JSON",
                    data=buffer_json_filt,
                    file_name="inventario_filtrado.json",
                    mime="application/json"
                )

            with col_toml_filt:
                buffer_toml_filt = exportar_a_toml(datos_filtrados)
                st.download_button(
                    "⬇️ TOML",
                    data=buffer_toml_filt,
                    file_name="inventario_filtrado.toml",
                    mime="application/toml"
                )

    # ==================== TAB 4: REPORTES & ANALÍTICAS ====================
    with tab4:
        st.header("📈 Reportes y Analíticas Avanzadas")

        # Clasificación ABC
        st.subheader("📊 Clasificación ABC (Análisis de Pareto)")
        st.markdown("""
        La clasificación ABC categoriza los items del inventario según su valor:
        - **Clase A:** Items que representan el 80% del valor total (alta prioridad)
        - **Clase B:** Items que representan el siguiente 15% del valor (prioridad media)
        - **Clase C:** Items que representan el último 5% del valor (baja prioridad)
        """)

        df_abc = clasificacion_abc(datos)

        col_abc1, col_abc2 = st.columns(2)
        with col_abc1:
            fig_abc = crear_grafico_distribucion_abc(datos)
            if fig_abc:
                st.plotly_chart(fig_abc, use_container_width=True)

        with col_abc2:
            # Estadísticas por clase
            for clase in ['A', 'B', 'C']:
                items_clase = df_abc[df_abc['CLASE_ABC'] == clase]
                if not items_clase.empty:
                    valor_clase = items_clase['VALOR_TOTAL_ITEM'].sum()
                    st.metric(
                        f"Clase {clase}",
                        f"{len(items_clase)} items",
                        f"${valor_clase:,.2f}"
                    )

        with st.expander("Ver detalle de clasificación ABC"):
            st.dataframe(
                df_abc[['ITEM', 'DESCRIPCIÓN', 'CANT.', 'PRECIO UNIT',
                       'VALOR_TOTAL_ITEM', 'CLASE_ABC', 'PORCENTAJE_ACUMULADO']],
                use_container_width=True
            )

        st.markdown("---")

        # Top items por valor
        st.subheader("🏆 Top Items por Valor Total")
        fig_top = crear_grafico_top_items_valor(datos, top_n=10)
        if fig_top:
            st.plotly_chart(fig_top, use_container_width=True)

        st.markdown("---")

        # Distribución por ubicación
        st.subheader("📍 Distribución por Ubicación")
        fig_ubicacion = crear_grafico_items_por_ubicacion(datos)
        if fig_ubicacion:
            st.plotly_chart(fig_ubicacion, use_container_width=True)

    # ==================== TAB 5: AGREGAR ITEM ====================
    with tab5:
        st.header("➕ Agregar Nuevo Item al Inventario")

        st.markdown("### Información del Item")
        st.markdown("*Los campos marcados con * son obligatorios*")

        with st.form("form-nuevo-item", clear_on_submit=True):
            col1, col2 = st.columns(2)

            with col1:
                nuevo_item = st.text_input('ITEM *', help="Identificador único del item")
                descripcion = st.text_input('DESCRIPCIÓN *', help="Descripción detallada")
                marca = st.text_input('MARCA')
                modelo = st.text_input('MODELO')
                pn = st.text_input('P/N', help="Número de parte")
                sn = st.text_input('S/N', help="Número de serie")
                categoria = st.text_input('CATEGORÍA', value='Sin categoría')

            with col2:
                observaciones = st.text_area('OBSERVACIONES', height=100)
                status = st.text_input('STATUS', value='Activo')
                ubicacion = st.text_input('UBICACIÓN')
                medida = st.text_input('MEDIDA', help="Unidad de medida")
                cantidad = st.number_input('CANTIDAD *', min_value=0, value=0, step=1)
                precio_unit = st.number_input('PRECIO UNITARIO *', min_value=0.0, value=0.0, step=0.01)
                stock_minimo = st.number_input('STOCK MÍNIMO', min_value=0, value=0, step=1,
                                              help="Cantidad mínima antes de alerta")

            enviado = st.form_submit_button('✅ Agregar Item al Inventario', type="primary")

        if enviado:
            # Validaciones
            errores_validacion = []

            # Validar campos requeridos
            if not nuevo_item or not descripcion:
                errores_validacion.append("Los campos ITEM y DESCRIPCIÓN son obligatorios")

            # Validar unicidad
            if nuevo_item:
                es_unico, msg_unico = validar_item_unico(datos, nuevo_item)
                if not es_unico:
                    errores_validacion.append(msg_unico)

            # Validar valores positivos
            es_valido, errores_valores = validar_valores_positivos(cantidad, precio_unit, stock_minimo)
            if not es_valido:
                errores_validacion.extend(errores_valores)

            # Si hay errores, mostrarlos
            if errores_validacion:
                for error in errores_validacion:
                    st.error(f"❌ {error}")
            else:
                # Calcular total automáticamente
                total_calculado = calcular_total(cantidad, precio_unit)

                # Crear nuevo item con timestamps
                nuevo = {
                    'ITEM': nuevo_item,
                    'DESCRIPCIÓN': descripcion,
                    'MARCA': marca if marca else 'No disponible',
                    'MODELO': modelo if modelo else 'No disponible',
                    'P/N': pn if pn else 'No disponible',
                    'S/N': sn if sn else 'No disponible',
                    'OBSERVACIONES': observaciones if observaciones else 'No disponible',
                    'STATUS': status,
                    'UBICACIÓN': ubicacion if ubicacion else 'No disponible',
                    'MEDIDA': medida if medida else 'No disponible',
                    'CANT.': cantidad,
                    'PRECIO UNIT': precio_unit,
                    'TOTAL': total_calculado,
                    'CATEGORÍA': categoria,
                    'STOCK MÍNIMO': stock_minimo,
                    'FECHA_CREACIÓN': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'FECHA_MODIFICACIÓN': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                }

                # Agregar al inventario
                st.session_state['datos'] = clean_df(pd.concat(
                    [datos, pd.DataFrame([nuevo])], ignore_index=True))

                st.success(f'✅ Item "{nuevo_item}" agregado exitosamente!')
                st.success(f'📊 Total calculado automáticamente: ${total_calculado:.2f}')
                st.info(f'📝 Total de items en inventario: {len(st.session_state["datos"])}')

                # Mostrar alerta si el stock es bajo
                if cantidad <= stock_minimo:
                    st.warning(f'⚠️ ALERTA: El item fue agregado con stock bajo (Stock: {cantidad}, Mínimo: {stock_minimo})')

                st.balloons()
                datos = st.session_state['datos']

if __name__ == "__main__":
    app()


