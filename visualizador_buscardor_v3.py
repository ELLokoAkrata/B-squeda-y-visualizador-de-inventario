import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
import tempfile
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog

# Variable global para almacenar los datos originales
datos_originales = None

def seleccionar_xlsx():
    global datos_originales

    archivo_xlsx = filedialog.askopenfilename(title="Seleccionar archivo XLSX", filetypes=[("Archivos XLSX", "*.xlsx;*.xls;*.xlsm")])

    if archivo_xlsx:
        # Si el archivo XLSX ya existe, actualizar la entrada y mostrar los datos
        entry_archivo_xlsx.config(state='normal')
        entry_archivo_xlsx.delete(0, tk.END)
        entry_archivo_xlsx.insert(0, archivo_xlsx)
        entry_archivo_xlsx.config(state='disabled')

        datos_originales = pd.read_excel(archivo_xlsx, skiprows=3)
        mostrar_datos(datos_originales)
        # Actualizar el texto del label_nombre_xlsx
        label_nombre_xlsx.config(text=f"Archivo XLSX actual: {archivo_xlsx}")
    else:
        resultado.delete(1.0, tk.END)
        resultado.insert(tk.END, "Por favor, selecciona un archivo XLSX y especifica un nombre para el archivo CSV.")
        # Limpiar el texto del label_nombre_xlsx si no hay archivo XLSX
        label_nombre_xlsx.config(text="")

def seleccionar_csv():
    global datos_originales

    archivo_csv = filedialog.askopenfilename(title="Seleccionar archivo CSV", filetypes=[("Archivo CSV", "*.csv")])

    if archivo_csv and os.path.exists(archivo_csv):
        # Si el archivo CSV ya existe, actualizar la entrada y mostrar los datos
        entry_archivo_csv.config(state='normal')
        entry_archivo_csv.delete(0, tk.END)
        entry_archivo_csv.insert(0, archivo_csv)
        entry_archivo_csv.config(state='disabled')

        datos_originales = pd.read_csv(archivo_csv)
        mostrar_datos(datos_originales)
        # Actualizar el texto del label_nombre_csv
        label_nombre_csv.config(text=f"Archivo CSV actual: {archivo_csv}")
    else:
        resultado.delete(1.0, tk.END)
        resultado.insert(tk.END, "Por favor, selecciona un archivo CSV existente para visualizar los datos.")
        # Limpiar el texto del label_nombre_csv si no hay archivo CSV
        label_nombre_csv.config(text="")

def convertir_y_visualizar():
    global datos_originales
    
    archivo_xlsx = filedialog.askopenfilename(title="Seleccionar archivo XLSX", filetypes=[("Archivos XLSX", "*.xlsx;*.xls;*.xlsm")])
    archivo_csv = filedialog.asksaveasfilename(title="Guardar como archivo CSV", defaultextension=".csv", filetypes=[("Archivo CSV", "*.csv")])

    if archivo_xlsx:
        if not archivo_csv and os.path.exists(archivo_xlsx.replace('.xlsx', '.csv')):
            # Si el archivo CSV ya existe, solo visualizar los datos
            archivo_csv = archivo_xlsx.replace('.xlsx', '.csv')
            datos_originales = pd.read_csv(archivo_csv)
            mostrar_datos(datos_originales)
            # Actualizar el texto del label_nombre_csv
            label_nombre_csv.config(text=f"Archivo CSV actual: {archivo_csv}")
        else:
            # Convertir el archivo xlsx a csv
            convertir_xlsx_a_csv(archivo_xlsx, archivo_csv)

            # Leer el archivo CSV y almacenar los datos originales
            datos_originales = pd.read_csv(archivo_csv)

            # Visualizar los datos originales
            mostrar_datos(datos_originales)
            # Actualizar el texto del label_nombre_csv
            label_nombre_csv.config(text=f"Archivo CSV actual: {archivo_csv}")
    else:
        resultado.delete(1.0, tk.END)
        resultado.insert(tk.END, "Por favor, selecciona un archivo XLSX y especifica un nombre para el archivo CSV.")

def visualizar():
    global datos_originales

    archivo_csv = filedialog.askopenfilename(title="Seleccionar archivo CSV", filetypes=[("Archivo CSV", "*.csv")])

    if archivo_csv and os.path.exists(archivo_csv):
        # Si el archivo CSV ya existe, solo visualizar los datos
        datos_originales = pd.read_csv(archivo_csv)
        mostrar_datos(datos_originales)
        # Actualizar el texto del label_nombre_csv
        label_nombre_csv.config(text=f"Archivo CSV actual: {archivo_csv}")
    else:
        resultado.delete(1.0, tk.END)
        resultado.insert(tk.END, "Por favor, selecciona un archivo CSV existente para visualizar los datos.")
        # Limpiar el texto del label_nombre_csv si no hay archivo CSV
        label_nombre_csv.config(text="")

def convertir_xlsx_a_csv(archivo_xlsx, archivo_csv):
    # Leer el archivo xlsx
    datos = pd.read_excel(archivo_xlsx, skiprows=3)  # Ignorar las primeras 3 filas
    
    # Eliminar columnas y filas con NaN
    datos = datos.dropna(axis=0, how='all').dropna(axis=1, how='all')
    
    # Guardar como archivo CSV
    datos.to_csv(archivo_csv, index=False)

def mostrar_datos(datos):
    resultado.delete(1.0, tk.END)  # Limpiar el resultado anterior
    resultado.insert(tk.END, datos.to_string(index=False))
    
def mostrar_datos_resaltando(datos, texto_busqueda):
    resultado.delete(1.0, tk.END)  # Limpiar el resultado anterior

    for index, row in datos.iterrows():
        linea_resaltada = resaltar_coincidencias(row, texto_busqueda)
        resultado.insert(tk.END, f"{linea_resaltada}\n")
        
        # Resaltar texto
        inicio = tk.END + f"-{len(linea_resaltada) + 1}c"
        fin = tk.END
        resultado.tag_add("resaltado", inicio, fin)
        resultado.tag_config("resaltado", font="helvetica 12 bold", foreground="blue")


def resaltar_coincidencia(palabra, texto_busqueda):
    inicio = palabra.lower().find(texto_busqueda)
    if inicio != -1:
        fin = inicio + len(texto_busqueda)
        palabra_resaltada = f"{palabra[:inicio]}{palabra[inicio:fin]}{palabra[fin:]}"
        return palabra_resaltada, True
    else:
        return palabra, False

def resaltar_coincidencias(row, texto_busqueda):
    palabras_resaltadas = []
    for cell in row:
        palabra_resaltada, coincidencia = resaltar_coincidencia(str(cell), texto_busqueda)
        palabras_resaltadas.append(palabra_resaltada)
    return ' | '.join(palabras_resaltadas)

def resumen_busqueda(datos_filtrados):
    resumen = datos_filtrados.describe()  # Puedes personalizar esto según tus necesidades
    return resumen

def mostrar_resumen(datos_filtrados, texto_busqueda):
    resumen = resumen_busqueda(datos_filtrados)
    resultado.insert(tk.END, "\n\nResumen de la búsqueda:\n")
    resultado.insert(tk.END, resumen.to_string())

    # Crear un gráfico y guardarlo temporalmente con el término de búsqueda como título
    temp_file = generar_grafico_resumen(resumen, texto_busqueda)

    # Mostrar el gráfico en una ventana emergente
    mostrar_grafico_emergente(temp_file)

def generar_grafico_resumen(resumen, texto_busqueda):
    fig, ax = plt.subplots(figsize=(8, 5))
    resumen.plot(kind='bar', ax=ax)
    ax.set_title(f"Resumen de la Búsqueda: {texto_busqueda}")
    temp_file = os.path.join(tempfile.gettempdir(), 'resumen_plot.png')
    plt.savefig(temp_file)
    plt.close()
    return temp_file

def mostrar_grafico_emergente(temp_file):
    # Cargar la imagen usando PhotoImage
    imagen = tk.PhotoImage(file=temp_file)
    
    # Hacer una referencia al objeto PhotoImage para evitar que sea eliminado por la recolección de basura
    resultado.image = imagen

    # Crear una ventana emergente para mostrar la imagen
    popup = tk.Toplevel(root)
    popup.title("Gráfico Emergente")

    # Mostrar la imagen en la ventana emergente
    label = tk.Label(popup, image=imagen)
    label.pack()

    # Cerrar el popup al hacer clic en la imagen
    label.bind("<Button-1>", lambda event: popup.destroy())

def buscar():
    global datos_originales
    
    texto_busqueda = entry_busqueda.get().lower()
    resultado.delete(1.0, tk.END)  # Limpiar el resultado anterior
    
    if datos_originales is not None:
        # Intentar buscar por número de ítem directo
        try:
            prefix, item_buscado = texto_busqueda.split(' ', 1)
            if prefix.lower() == 'item':
                item_buscado = int(item_buscado)
                datos_filtrados = datos_originales[datos_originales['ITEM'] == item_buscado]
                if not datos_filtrados.empty:
                    mostrar_datos_resaltando(datos_filtrados, texto_busqueda)
                    mostrar_resumen(datos_filtrados, texto_busqueda)
                    return
        except (ValueError, IndexError):
            pass  # Continuar con la búsqueda normal si no es un formato válido
            
        # Filtrar los datos originales basados en la búsqueda
        datos_filtrados = datos_originales[datos_originales.apply(lambda row: any(row.astype(str).str.contains(texto_busqueda, case=False, regex=True)), axis=1)]
        
        if not datos_filtrados.empty:
            # Mostrar el resumen y el gráfico emergente con el título
            mostrar_datos_resaltando(datos_filtrados, texto_busqueda)
            mostrar_resumen(datos_filtrados, texto_busqueda)
            return


        # Filtrar los datos originales basados en la búsqueda en todas las columnas
        datos_filtrados = datos_originales[datos_originales.apply(lambda row: any(row.astype(str).str.contains(texto_busqueda, case=False, regex=True)), axis=1)]
        mostrar_datos_resaltando(datos_filtrados, texto_busqueda)
        mostrar_resumen(datos_filtrados)

def restaurar_vista_normal():
    global datos_originales
    if datos_originales is not None:
        mostrar_datos(datos_originales)
        resultado.delete("end-2c", tk.END)  # Eliminar solo resumen
        
# Configuración de la interfaz tkinter
root = TkinterDnD.Tk()
root.title("Conversor y Visualizador")
# Interfaz
frame_archivos = ttk.Frame(root)
frame_archivos.pack(pady=10)

# Label archivo CSV
label_archivo_csv = ttk.Label(frame_archivos, text="Archivo CSV:")
label_archivo_csv.grid(row=0, column=0, padx=5)

# Label para mostrar el nombre del archivo CSV actual
label_nombre_csv = ttk.Label(frame_archivos, text="")
label_nombre_csv.grid(row=0, column=1, padx=5)

# Entry para mostrar el nombre del archivo CSV
entry_archivo_csv = ttk.Entry(frame_archivos, width=30, state='disabled')
entry_archivo_csv.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

# Agregar botón de selección para el archivo CSV
button_seleccionar_csv = ttk.Button(frame_archivos, text="Seleccionar CSV", command=seleccionar_csv)
button_seleccionar_csv.grid(row=0, column=3, padx=5)

# Label para mostrar el nombre del archivo XLSX actual
label_nombre_xlsx = ttk.Label(frame_archivos, text="")
label_nombre_xlsx.grid(row=1, column=1, padx=5)

# Entry para mostrar el nombre del archivo XLSX
entry_archivo_xlsx = ttk.Entry(frame_archivos, width=30, state='disabled')
entry_archivo_xlsx.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

# Agregar botón de selección para el archivo XLSX
button_seleccionar_xlsx = ttk.Button(frame_archivos, text="Seleccionar XLSX", command=seleccionar_xlsx)
button_seleccionar_xlsx.grid(row=1, column=3, padx=5)

button_convertir_visualizar = ttk.Button(frame_archivos, text="Convertir y Visualizar", command=convertir_y_visualizar)
button_convertir_visualizar.grid(row=2, column=0, padx=5)

button_visualizar = ttk.Button(frame_archivos, text="Visualizar", command=visualizar)
button_visualizar.grid(row=2, column=1, padx=5)

label_busqueda = ttk.Label(frame_archivos, text="Buscar:")
label_busqueda.grid(row=2, column=2, padx=5)

entry_busqueda = ttk.Entry(frame_archivos)
entry_busqueda.grid(row=2, column=3, padx=5)

#Botón buscar
button_buscar = ttk.Button(frame_archivos, text="Buscar", command=buscar)
button_buscar.grid(row=3, column=0, padx=5)

# Agregar botón para restaurar la vista normal
button_restaurar = ttk.Button(frame_archivos, text="Restaurar Vista Normal", command=restaurar_vista_normal)
button_restaurar.grid(row=3, column=1, padx=5)

frame_resultado = ttk.Frame(root)
frame_resultado.pack(padx=10, pady=10)

# Aumentar el tamaño del cuadro de texto y agregar barras de desplazamiento
resultado = tk.Text(frame_resultado, height=20, width=100, relief=tk.GROOVE, borderwidth=2, wrap='none')
resultado.grid(row=0, column=0, sticky="nsew")

scroll_y = ttk.Scrollbar(frame_resultado, orient="vertical", command=resultado.yview)
scroll_y.grid(row=0, column=1, sticky="ns")
resultado.config(yscrollcommand=scroll_y.set)

scroll_x = ttk.Scrollbar(frame_resultado, orient="horizontal", command=resultado.xview)
scroll_x.grid(row=1, column=0, sticky="ew")
resultado.config(xscrollcommand=scroll_x.set)

# Ajustar la expansión de las columnas y filas
frame_resultado.columnconfigure(0, weight=1)
frame_resultado.rowconfigure(0, weight=1)

# Iniciar la interfaz
root.mainloop()
