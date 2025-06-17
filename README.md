# Búsqueda y visualizador de inventario

Aplicación sencilla en Streamlit para cargar un archivo de inventario en formato
CSV o XLSX, buscar información y ahora también añadir nuevos ítems. Al cargar un
documento XLSX la aplicación copia el archivo para preservar su estructura y
cualquier nuevo producto se inserta en esa copia manteniendo los encabezados y
añadiendo bordes a las nuevas filas. Tras realizar cambios se puede descargar el
inventario actualizado sin sobrescribir el documento original.
