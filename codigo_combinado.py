# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\ActualVersion\rpa_utilization_nest_202412301814.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"



def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Eliminar comillas dobles alrededor de la ruta
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)
logo_path = config.get('logo_path', None)

if excel_output_path:
    print(f'Output path set: {excel_output_path}')
else:
    print('The path "excel_output_path" was not found in the configuration file.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'The path exists: {excel_output_path}')
else:
    print('The specified path does not exist or is not valid.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Invalid date, make sure to use the MM/DD/YYYY format.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{1,2}_\d{1,2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF",
    "STEEL-0.375": "375 INCH"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Checking files in the folder and subfolders: '{folder_path}' with the date {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluding file:: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"File: {file_path} | Category: {category}")

            return pdf_files
        else:
            print(f"The path '{folder_path}' does not exist or is not a valid folder.")
            return []
    except Exception as e:
        print(f"Error listing files in the folder: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)
        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = round(float(utilization_data[3].replace("Utililization: ", "").strip().replace("%", ""),) / 100, 4)
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Warning", "No lines with 'Utilization' were found in the PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        #def get_material(nesting):
            #for material, values in category_dict.items():
                #if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    #return material

        #merged_df['Material'] = merged_df['NESTING'].apply(get_material)

        def get_material(nesting):
            # Asegurarnos de que nesting sea una cadena y no un valor NaN o None
            if isinstance(nesting, str):
                for material, values in category_dict.items():
                    if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                        return material
            # Si nesting no es una cadena, devolver "Desconocido"
            return "Desconocido"

        # Aplicar la función get_material a la columna 'NESTING' del DataFrame
        merged_df['Material'] = merged_df['NESTING'].apply(get_material)
    


      

        # Agregar la columna "Program"
        #merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        merged_df['Program'] = merged_df['Sheet Name'].apply(
                lambda x: str(x).split('-')[0] if isinstance(x, str) and '-' in x else str(x)
            )

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        #merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        merged_df['Gauge'] = merged_df['NESTING'].apply(
                lambda x: get_gauge_from_nesting(str(x), code_to_gauge)  # Convertir x a cadena
            )

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        #merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(str(x)))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Excel exported successfully: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the Excel file: {e}")
    else:
        messagebox.showerror("Error", "The DataFrames cannot be combined.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Select the input folder")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Invalid date. Make sure to use the MM/DD/YYYY format.")
        log_message("Error: Invalid date.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Please select an input folder.")
        log_message("Error: No input folder selected.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No PDF files found to process.")
        log_message("Error: No PDF files found.")
        return

    log_message(f"{len(pdf_files)} PDF files found to process.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Process finished", "The PDF file processing has been completed.")

    root.quit()  # Esto cierra la interfaz
    root.destroy()  # Esto termina la aplicación completamente

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Select the date:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Select folder", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Process PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\consolidate.py
import os
from PyPDF2 import PdfMerger
from datetime import datetime

# Ruta de la carpeta de entrada (ajusta según sea necesario)
folder_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2024\REWORK"
# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"

# Función para convertir la fecha de entrada al formato MM_dd_YYYY (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a MM_dd_YYYY
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")
        # Devolver la fecha como un string en el formato deseado
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None

# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha en el nombre de la carpeta
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        # Verifica que el directorio existe
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")
            
            # Usamos os.walk para recorrer todas las subcarpetas
            for root, dirs, files in os.walk(folder_path):
                # Comprobar si el nombre del folder (root) contiene la fecha
                if date_filter in os.path.basename(root).replace("\\", "/").lower():  # Compara el nombre de la carpeta con la fecha
                    for file_name in files:
                        file_path = os.path.join(root, file_name)
                        
                        # Excluir los archivos PDF cuyo path contenga "turret"
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue
                        
                        # Verificar que sea un archivo PDF (ignorando mayúsculas y minúsculas en la extensión)
                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            print(f"Archivo encontrado: {file_path}")
            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para consolidar los archivos PDF en un solo archivo
def consolidate_pdfs(folder_path, output_pdf, date_filter):
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    
    if not pdf_files:
        print("No se encontraron archivos PDF con la fecha especificada en la carpeta o sus subcarpetas.")
        return
    
    merger = PdfMerger()  # Crea un objeto PdfMerger para fusionar los PDFs
    
    try:
        # Iterar sobre cada archivo PDF y añadirlo al archivo final
        for pdf in pdf_files:
            print(f"Agregando {pdf} al archivo consolidado...")
            merger.append(pdf)  # Agrega el PDF al archivo final
        
        # Escribir el archivo PDF consolidado
        if os.path.exists(output_pdf):
            print(f"El archivo {output_pdf} ya existe. No se sobrescribirá.")
        else:
            merger.write(output_pdf)
            print(f"Archivos PDF consolidados exitosamente en: {output_pdf}")
        
        merger.close()
    
    except Exception as e:
        print(f"Error al consolidar los archivos PDF: {e}")

# Solicitar la fecha de entrada (ejemplo: 11/14/2024)
input_date = input("Ingresa la fecha (MM/DD/YYYY): ")

# Convertir la fecha al formato MM_dd_YYYY
date_filter = convert_date_format(input_date)

# Si la fecha es válida, proceder con la consolidación
if date_filter:
    # Nombre del archivo PDF de salida
    output_pdf = os.path.join(output_folder_path, f"consolidado_{date_filter}.pdf")

    # Llamada a la función para consolidar los PDFs
    consolidate_pdfs(folder_path, output_pdf, date_filter)  # Asegúrate de pasar la ruta de entrada, no la de salida


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\final_version.py
import os
from PyPDF2 import PdfMerger
from datetime import datetime
import pdfplumber
import pandas as pd

# Ruta de la carpeta de entrada (ajusta según sea necesario)
folder_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2024\ENCL"
# Ruta de la carpeta de salida para los archivos PDF consolidados
output_folder_path_pdf = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
# Ruta de la carpeta de salida para el archivo Excel
output_folder_path_excel = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Función para convertir la fecha de entrada al formato MM_dd_YYYY (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a MM_dd_YYYY
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")
        # Devolver la fecha como un string en el formato deseado
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None

# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha en el nombre de la carpeta
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        # Verifica que el directorio existe
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")
            
            # Usamos os.walk para recorrer todas las subcarpetas
            for root, dirs, files in os.walk(folder_path):
                # Comprobar si el nombre del folder (root) contiene la fecha
                if date_filter in os.path.basename(root).replace("\\", "/").lower():  # Compara el nombre de la carpeta con la fecha
                    for file_name in files:
                        file_path = os.path.join(root, file_name)
                        
                        # Excluir los archivos PDF cuyo path contenga "turret"
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue
                        
                        # Verificar que sea un archivo PDF (ignorando mayúsculas y minúsculas en la extensión)
                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            print(f"Archivo encontrado: {file_path}")
            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para consolidar los archivos PDF en un solo archivo
def consolidate_pdfs(folder_path, output_pdf, date_filter):
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    
    if not pdf_files:
        print("No se encontraron archivos PDF con la fecha especificada en la carpeta o sus subcarpetas.")
        return None
    
    merger = PdfMerger()  # Crea un objeto PdfMerger para fusionar los PDFs
    
    try:
        # Iterar sobre cada archivo PDF y añadirlo al archivo final
        for pdf in pdf_files:
            print(f"Agregando {pdf} al archivo consolidado...")
            merger.append(pdf)  # Agrega el PDF al archivo final
        
        # Escribir el archivo PDF consolidado
        if os.path.exists(output_pdf):
            print(f"El archivo {output_pdf} ya existe. No se sobrescribirá.")
        else:
            merger.write(output_pdf)
            print(f"Archivos PDF consolidados exitosamente en: {output_pdf}")
        
        merger.close()
        return output_pdf  # Devolver el path del archivo consolidado
    
    except Exception as e:
        print(f"Error al consolidar los archivos PDF: {e}")
        return None

# Función para buscar en tablas del PDF
def extract_tables_with_keyword(pdf_path, keyword):
    with pdfplumber.open(pdf_path) as pdf:
        matching_rows = []
        
        # Iterar por cada página
        for page_num, page in enumerate(pdf.pages):
            # Extraer las tablas de la página
            tables = page.extract_tables()
            
            # Si hay tablas en la página
            if tables:
                for table in tables:
                    for row in table:
                        # Buscar la palabra clave en las filas de la tabla
                        if any(keyword in str(cell) for cell in row if cell):  # Verifica que la celda no sea None
                            matching_rows.append(row)  # Agrega la fila completa
        
        return matching_rows

# Función principal para ejecutar el flujo completo
def process_pdfs_and_generate_excel(folder_path, output_folder_path_pdf, output_folder_path_excel, input_date, keyword):
    # Convertir la fecha al formato MM_dd_YYYY
    date_filter = convert_date_format(input_date)

    if date_filter:
        # Nombre del archivo PDF consolidado
        output_pdf = os.path.join(output_folder_path_pdf, f"consolidado_{date_filter}.pdf")

        # Consolidar los PDFs que cumplen con la fecha
        consolidated_pdf_path = consolidate_pdfs(folder_path, output_pdf, date_filter)
        
        if consolidated_pdf_path:
            # Extraer las tablas que contienen la palabra clave
            rows_with_keyword = extract_tables_with_keyword(consolidated_pdf_path, keyword)

            # Crear un DataFrame para almacenar los resultados
            columns = ["Nesting", "Columna 2", "Material", "Columna 4", "Quantity", "Utilization", "Columna 7", "Size"]
            results_df = pd.DataFrame(rows_with_keyword, columns=columns)

            # Eliminar columnas no deseadas
            results_df = results_df.drop(columns=["Columna 2", "Columna 4", "Columna 7"])

            # Renombrar columnas
            results_df.rename(columns={
                "Nesting": "Nesting",
                "Material": "Material",
                "Quantity": "Quantity",
                "Utilization": "Utilization",
                "Size": "Size"
            }, inplace=True)

            # Limpiar datos en las columnas Quantity y Utilization
            results_df["Quantity"] = results_df["Quantity"].str.replace("Stack Qty: ", "", regex=False)
            results_df["Utilization"] = results_df["Utilization"].str.replace("Utililization: ", "", regex=False)

            # Generar la fecha de ejecución para el nombre del archivo
            execution_date = datetime.now().strftime("%Y-%m-%d")
            output_file_name = f"Utilization_{execution_date}.xlsx"
            output_excel_path = os.path.join(output_folder_path_excel, output_file_name)

            # Exportar los resultados a un archivo Excel
            if not results_df.empty:
                results_df.to_excel(output_excel_path, index=False)
                print(f"Resultados exportados a Excel en: {output_excel_path}")
            else:
                print(f"No se encontraron filas con la palabra '{keyword}'.")
        else:
            print("No se pudo consolidar los archivos PDF.")
    else:
        print("Fecha no válida. No se puede continuar.")

# Solicitar la fecha de entrada (ejemplo: 11/14/2024)
input_date = input("Ingresa la fecha (MM/DD/YYYY): ")
# Palabra clave a buscar en las tablas del PDF
keyword = "Utililization"

# Ejecutar el proceso completo
process_pdfs_and_generate_excel(folder_path, output_folder_path_pdf, output_folder_path_excel, input_date, keyword)


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\parts.py
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime
from tqdm import tqdm  # Librería para la barra de progreso

# Función para extraer todas las tablas del PDF y filtrar líneas específicas
def extract_filtered_tables(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_matching_rows = []

        # Configurar la barra de progreso para el número de páginas en el PDF
        total_pages = len(pdf.pages)
        print(f"Procesando {total_pages} páginas...")

        # Iterar por cada página con la barra de progreso
        for page_num, page in enumerate(tqdm(pdf.pages, desc="Procesando páginas", unit="página")):
            # Extraer tablas de la página
            tables = page.extract_tables()
            footer_text = page.extract_text()  # Extraer texto completo de la página para obtener el pie de página

            # Obtener el pie de página
            footer_lines = footer_text.split('\n')[-1] if footer_text else "Sin pie de página"  # Tomar la última línea como pie de página

            # Extraer campos del pie de página según la estructura dada
            footer_parts = footer_lines.split(' ')
            if len(footer_parts) >= 4:
                date = footer_parts[0]  # Fecha
                time = footer_parts[1] + ' ' + footer_parts[2]  # Hora (am/pm)
                nesting_full = ' '.join(footer_parts[3:-2])  # Nesting completo, sin el último número
                # Asegurarse de que nesting tenga tres caracteres después del punto
                try:
                    nesting = nesting_full.split('.')[0] + '.' + nesting_full.split('.')[1][:3]
                except IndexError:
                    nesting = nesting_full
            else:
                date, time, nesting = ["Sin dato"] * 3  # Si no hay datos, colocar "Sin dato"

            # Si hay tablas en la página
            if tables:
                for table in tables:
                    for row in table[1:]:  # Omite la fila de encabezados
                        # Filtrar filas donde al menos una celda comience con dos letras minúsculas
                        if any(re.match(r'^[a-z]{2}', str(cell)) for cell in row if cell):
                            # Eliminar registros que contengan letras en la columna específica (por ejemplo, columna 1)
                            if not any(char.isalpha() for char in str(row[0])):  # Cambia el índice según la columna que desees verificar
                                # Eliminar los dos primeros caracteres de 'Part Name'
                                part_name_full = row[1][2:] if len(row[1]) > 2 else row[1]  # Ajustar índice según la columna

                                # Dividir 'Part Name' en 'Projecto' y 'Part Name' usando el primer espacio
                                split_name = part_name_full.split(' ', 1)
                                project = split_name[0] if len(split_name) > 0 else ''
                                part_name = split_name[1] if len(split_name) > 1 else ''

                                # Construir la nueva fila con los datos del pie de página (sin columna 1)
                                new_row = [project, part_name, row[3], date, time, nesting]
                                all_matching_rows.append(new_row)

        return all_matching_rows

# Nueva ruta del archivo PDF
pdf_file_path = r"C:\Users\User\OneDrive - globalpowercomponents\Documents\Utilization\Consolidated_PDF\consolidado.pdf"

# Extraer y filtrar las tablas del PDF
matching_rows = extract_filtered_tables(pdf_file_path)

# Crear un DataFrame para almacenar los resultados
if matching_rows:
    # Generar encabezados, eliminando la columna 1 y ajustando los nombres de las columnas
    columns = ["Projecto", "Part Name", "Quantity Total", "Fecha", "Hora", "Nesting"]
    
    results_df = pd.DataFrame(matching_rows, columns=columns)

    # Eliminar registros duplicados
    results_df = results_df.drop_duplicates()

    # Obtener la fecha actual en formato mm-dd-yyyy
    current_date = datetime.now().strftime("%m-%d-%Y")

    # Definir la ruta para guardar el archivo Excel con el nombre adecuado
    output_folder_path = r"C:\Users\User\OneDrive - globalpowercomponents\Documents\Utilization\Parts_Request_Reports"
    output_excel_name = f"Nesting_report_{current_date}.xlsx"
    output_excel_path = os.path.join(output_folder_path, output_excel_name)

    # Exportar los resultados a un archivo Excel
    results_df.to_excel(output_excel_path, index=False)
    print(f"Resultados filtrados exportados a Excel en: {output_excel_path}")
else:
    print("No se encontraron tablas que cumplan con el criterio en el PDF.")


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\procces_v1.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas


# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función principal que se ejecuta al presionar el botón
def main():
    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        return

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    canvas.create_window(300, 400, window=progress_bar)

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    canvas.create_window(300, 430, window=percentage_label)

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x450")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo negro
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo
canvas = Canvas(root, width=600, height=450, bg="#2F2F2F")
canvas.pack(fill="both", expand=True)

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior izquierda con un margen pequeño
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
canvas.create_window(300, 60, window=date_label)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
canvas.create_window(300, 120, window=date_calendar)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
canvas.create_window(300, 240, window=folder_path_label)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
canvas.create_window(300, 280, window=select_folder_button)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
canvas.create_window(300, 350, window=process_button)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\procces_v2.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re

# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

  

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función principal que se ejecuta al presionar el botón
def main():
    # Obtener la fecha de entrada desde la caja de texto
    input_date = date_entry.get()
    date_filter = convert_date_format(input_date)

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        return

    # Listar archivos PDF
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        return

    # Configurar la barra de progreso
    progress_var = DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400)
    progress_bar.pack(pady=10)
    percentage_label = Label(root, text="0%")
    percentage_label.pack()

    # Procesar los PDFs y generar el archivo Excel
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = Tk()
root.title("Consolidar y Procesar PDF")

# Fijar el tamaño de la ventana
root.geometry("600x400")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Etiqueta y campo de entrada para la fecha
date_label = Label(root, text="Fecha (MM/DD/YYYY):")
date_label.pack()
date_entry = Entry(root)
date_entry.pack()

# Etiqueta y botón para seleccionar la carpeta
folder_label = Label(root, text="Carpeta de entrada:")
folder_label.pack()
folder_path_label = Label(root, text="")
folder_path_label.pack()
select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder)
select_folder_button.pack()

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main)
process_button.pack()

root.mainloop()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\process_202412301056.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\process_Copy.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"



def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Eliminar comillas dobles alrededor de la ruta
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)
logo_path = config.get('logo_path', None)

if excel_output_path:
    print(f'Output path set: {excel_output_path}')
else:
    print('The path "excel_output_path" was not found in the configuration file.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'The path exists: {excel_output_path}')
else:
    print('The specified path does not exist or is not valid.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Invalid date, make sure to use the MM/DD/YYYY format.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{1,2}_\d{1,2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Checking files in the folder and subfolders: '{folder_path}' with the date {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluding file:: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"File: {file_path} | Category: {category}")

            return pdf_files
        else:
            print(f"The path '{folder_path}' does not exist or is not a valid folder.")
            return []
    except Exception as e:
        print(f"Error listing files in the folder: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)
        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Warning", "No lines with 'Utilization' were found in the PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        #def get_material(nesting):
            #for material, values in category_dict.items():
                #if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    #return material

        #merged_df['Material'] = merged_df['NESTING'].apply(get_material)

        def get_material(nesting):
            # Asegurarnos de que nesting sea una cadena y no un valor NaN o None
            if isinstance(nesting, str):
                for material, values in category_dict.items():
                    if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                        return material
            # Si nesting no es una cadena, devolver "Desconocido"
            return "Desconocido"

        # Aplicar la función get_material a la columna 'NESTING' del DataFrame
        merged_df['Material'] = merged_df['NESTING'].apply(get_material)
    


      

        # Agregar la columna "Program"
        #merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        merged_df['Program'] = merged_df['Sheet Name'].apply(
                lambda x: str(x).split('-')[0] if isinstance(x, str) and '-' in x else str(x)
            )

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        #merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        merged_df['Gauge'] = merged_df['NESTING'].apply(
                lambda x: get_gauge_from_nesting(str(x), code_to_gauge)  # Convertir x a cadena
            )

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        #merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(str(x)))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Excel exported successfully: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the Excel file: {e}")
    else:
        messagebox.showerror("Error", "The DataFrames cannot be combined.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Select the input folder")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Invalid date. Make sure to use the MM/DD/YYYY format.")
        log_message("Error: Invalid date.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Please select an input folder.")
        log_message("Error: No input folder selected.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No PDF files found to process.")
        log_message("Error: No PDF files found.")
        return

    log_message(f"{len(pdf_files)} PDF files found to process.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Process finished", "The PDF file processing has been completed.")

    root.quit()  # Esto cierra la interfaz
    root.destroy()  # Esto termina la aplicación completamente

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Select the date:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Select folder", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Process PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\process_old_version_ok.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x500")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\prueba_extraccion_tablaspdf.py
import pdfplumber
import pandas as pd

# Define la ruta del archivo PDF
pdf_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\11-19-2024-N23140217ENCL.PDF'

# Define la estructura esperada de la primera fila
expected_first_row = ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]

# Abre el archivo PDF
with pdfplumber.open(pdf_path) as pdf:
    all_headers = []
    all_utilization_lines = []  # Para almacenar las líneas extraídas que contienen 'Utililization'

    for page_number, page in enumerate(pdf.pages, start=1):
        # Extrae las tablas de cada página
        tables = page.extract_tables()

        for table_number, table in enumerate(tables, start=1):
            if table and table[0] == expected_first_row:  # Verifica que la primera fila coincida con la estructura esperada
                # Extrae la primera fila de datos después de los encabezados
                data = table[1]  # Extrae la fila 2
                print(f"Fila extraída de la tabla {table_number} en la página {page_number}: {data}")
                # Asegúrate de que los datos se asignan correctamente a las columnas
                machine = data[0]
                schedule = data[2]
                total_cut_time = data[8]  # Ajusta esta posición si es necesario
                headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                all_headers.append(headers_df)

        # Buscar y extraer la línea que contenga "Utililization"
        for table in tables:
            for row in table:
                for cell in row:
                    if cell and "Utililization" in str(cell):  # Busca "Utililization" en cada celda
                        # Si encuentra la palabra, extrae solo los índices 0, 2, 4, 5, 7
                        utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                        
                        # Modificar el valor de los índices según los requisitos
                        # 1. Eliminar salto de línea en el índice 0 y renombrar la columna a 'NESTING'
                        utilization_data[0] = utilization_data[0].replace("\n", "").strip()  # Eliminar salto de línea
                        
                        # 2. Eliminar "Stack Qty: " en el índice 4
                        utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")  # Eliminar "Stack Qty: "
                        
                        # 3. Eliminar "Utililization: " en el índice 5 y convertir a número
                        utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")  # Convertir a número

                        all_utilization_lines.append(utilization_data)  # Almacena la línea filtrada

# Combina las tablas de encabezados en un solo DataFrame
if all_headers:
    headers = pd.concat(all_headers, ignore_index=True)
    # Imprime el DataFrame de encabezados en la consola
    print("Headers DataFrame:")
    print(headers)
else:
    print("No se encontraron tablas con los encabezados esperados en el PDF.")

# Imprime las líneas encontradas con 'Utililization'
if all_utilization_lines:
    utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size'])
    print("\nLíneas con 'Utililization':")
    print(utilization_df)
else:
    print("No se encontraron líneas con 'Utililization' en el PDF.")

# Realizar el merge entre los dos DataFrames usando la columna 'SCHEDULE' y 'NESTING'
if not headers.empty and not utilization_df.empty:
    # Asegúrate de que ambas columnas tengan el mismo tipo de datos (string)
    headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
    utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
    
    # Merge de los DataFrames
    merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')

    # Eliminar la columna 'SCHEDULE' después del merge
    merged_df.drop(columns=['SCHEDULE'], inplace=True)

    # Extraer los valores de 'NESTING' y crear las nuevas columnas 'Material' y 'Gauge'
    # 1. Extraer Material: caracteres después del cuarto guion y antes del último
    def extract_material(nesting):
        # Split the string by '-'
        parts = nesting.split('-')
        
        # Si el valor en el cuarto guion es alfanumérico (ej: 3000-ALU), tomar ese rango
        if len(parts) > 4:
            material = '-'.join(parts[4:6])  # Concatenar partes 4 y 5 para formar '3000-ALU'
            return material
        # Si no hay guiones extras, simplemente devolver el valor después del cuarto guion
        return parts[4] if len(parts) > 4 else ''

    merged_df['Material'] = merged_df['NESTING'].apply(extract_material)

    # 2. Extraer Gauge: caracteres después del último guion
    merged_df['Gauge'] = merged_df['NESTING'].str.split('-').str[-1]  # Extrae la parte después del último guion

    # Imprime el DataFrame combinado con las nuevas columnas
    print("\nDataFrame combinado y modificado:")
    print(merged_df)
else:
    print("No se pueden combinar los DataFrames. Asegúrate de que ambos contengan datos.")


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\prueba_interfaz.py
import customtkinter as ctk

root = ctk.CTk()
root.geometry("400x400")

button = ctk.CTkButton(root, text="Iniciar proceso")
button.pack(pady=20)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\rpa_utilization_nest.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\config.py
#excel_output_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data"
logo_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret\image\Logo.png"
folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\main.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *



# Configuración de rutas
#folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
#folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
#excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'


# Función para obtener la fecha actual en formato m_d_yyyy

def main():
    inicio_proceso = datetime.now()
    input_date = get_today_date()
    _ ,laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date)
    df_consolidado1=process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, input_date)

    turret_files, _ = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date) 
    df_consolidado2 =process_pdfs_turret_and_generate_excel(turret_files, input_date)

    
    if df_consolidado1 is not None and df_consolidado2 is not None:        
        output_filename = f"Utilization_{input_date}.xlsx"
        outputfinal = pd.concat([df_consolidado1,df_consolidado2], axis=0)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is not None and df_consolidado2 is None: 
        outputfinal=pd.DataFrame(df_consolidado1)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is None and df_consolidado2 is not None: 
        outputfinal=pd.DataFrame(df_consolidado2)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}") 
    else:
        print("No se generó ningún archivo Excel.")

    fin_proceso = datetime.now()
    print(f"Tiempo de ejecución: {fin_proceso - inicio_proceso}")
if __name__ == "__main__":
    main()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\ActualVersion\rpa_utilization_nest_202412301814.py
import os
import re
import fitz  # PyMuPDF

def buscar_archivos_con_turret(carpeta):
    archivos_con_turret = []
    
    for raiz, dirs, archivos in os.walk(carpeta):
        for archivo in archivos:
            # Verificamos si la palabra 'turret' está en la ruta completa y el archivo es PDF
            if 'turret' in os.path.join(raiz, archivo).lower() and archivo.lower().endswith('.pdf'):
                archivos_con_turret.append(os.path.join(raiz, archivo))
    
    return archivos_con_turret

def extraer_tabla_machine_pdf(archivo_pdf):
    """
    Busca el encabezado 'MACHINE' en el archivo PDF y extrae los valores correspondientes de la tabla.
    """
    tabla_datos = []
    
    try:
        # Abrir el archivo PDF
        doc = fitz.open(archivo_pdf)
        
        # Iterar sobre todas las páginas del PDF
        for pagina in doc:
            texto = pagina.get_text("text")  # Extraer el texto de la página
            
            # Dividir el texto por líneas
            lineas = texto.split('\n')
            
            # Buscar el encabezado 'MACHINE' y extraer los valores correspondientes de la tabla
            for i, linea in enumerate(lineas):
                if 'MACHINE' in linea:
                    # Aquí buscamos la siguiente línea para extraer los valores
                    if i + 1 < len(lineas):
                        # Obtener la información de la siguiente línea que contiene los valores
                        datos = lineas[i + 1].strip()
                        tabla_datos.append(datos)
                        
                        # Si se encuentran otras filas relacionadas con la tabla, podemos extraerlas también
                        # Esto depende de la estructura del PDF y las líneas que siguen después
                        # Ejemplo: buscar la siguiente línea con los datos
                        for j in range(i + 2, len(lineas)):
                            # Si encontramos un patrón de datos esperado (por ejemplo, "Material Name"), extraemos más información
                            if re.match(r"\S+", lineas[j].strip()):
                                tabla_datos.append(lineas[j].strip())
                    break  # Salir si encontramos la tabla
    except Exception as e:
        print(f"Error al leer el archivo PDF {archivo_pdf}: {e}")
    
    return tabla_datos

# Carpeta donde buscar los archivos
carpeta = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2024\REWORK\12_27_2024_REWORK 5'

# Obtenemos los archivos PDF que contienen 'turret' en la ruta
archivos_pdf = buscar_archivos_con_turret(carpeta)

# Ahora vamos a obtener la información de esos archivos PDF
for archivo in archivos_pdf:
    print(f"Buscando la tabla 'MACHINE' en el archivo PDF: {archivo}")
    tabla_machine = extraer_tabla_machine_pdf(archivo)
    if tabla_machine:
        print("Datos extraídos de la tabla:")
        for linea in tabla_machine:
            print(linea)
    else:
        print("No se encontró la tabla 'MACHINE' en el archivo.")


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\Dictionaries.py
code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK316M2": "4.Citurret"
}

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuncionesFieldsTransform.py
####################################################################################################
# Funciones para transformar los campos de un DataFrame
####################################################################################################
# Las funciones a continuación se utilizan para transformar los campos de un DataFrame de Pandas.
# Estas funciones se pueden aplicar a una columna específica para modificar los valores de las celdas.
# Por ejemplo, se puede aplicar una función para ajustar los nombres de las máquinas o para obtener
# el calibre de una celda que contiene un 'nesting'.
#
def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"



def update_sheet_size(sheet_size, valid_sizes):
    # Eliminar los espacios en la cadena y convertir todo a minúsculas
    cleaned_size = sheet_size.replace(" ", "").lower()
    
    # Invertir la cadena limpia para verificar la variante invertida
    inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
    
    # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
    if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
        return sheet_size  # Conservar el valor original si hay coincidencia
    else:
        return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
    
    
# Función para obtener el material desde el nesting
def get_material(nesting, Material_dict):
    for material, values in Material_dict.items():
        if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
            return material
    return "Desconocido"


def adjust_type_nest(Type, type_dict):
    # Buscar si el tipo está en el diccionario y devolver el valor correspondiente
    return type_dict.get(Type, Type)


def adjust_machine(machine_name, machine_dict):
    # Buscar si el nombre de la máquina está en el diccionario y devolver el valor correspondiente
    return machine_dict.get(machine_name, machine_name)  # Si machine_name no está en el diccionario, retorna el valor original    



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionLaserPlasmaFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

def process_pdfs_laser_plasma_and_generate_excel(pdf_files, input_date):
    all_utilization_lines = []
    all_headers = []  # Lista de headers
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las fechas Jake de cada archivo

    headers = pd.DataFrame()  # Inicializamos headers como un DataFrame vacío para evitar el error

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process = input_date
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake = match.group(1)     

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
    # Si se encontraron headers y lines de utilización
    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        print("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")

    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

        # Agregar la columna "Material"

        merged_df['Material'] = merged_df['NESTING'].apply(lambda x: get_material(x, Material_dict)) 

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))


        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time'
            })

        # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization', 'Cut Time'
        ]

        # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Ajustar la columna 'Type' según las condiciones especificadas
        #merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  
        merged_df['Type'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
        df_consolidado1=pd.DataFrame(merged_df)

        '''# Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            print("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            print("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        print("Error", "No se pueden combinar los DataFrames.")'''
    return df_consolidado1  


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionListFiles.py
import os
from datetime import datetime
import re

def list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter):
    turret_files = []
    laser_plasma_files = []

    # Combina las dos rutas de directorios
    folder_paths = [folder_path_1, folder_path_2]

    # Itera sobre ambos directorios
    for folder_path in folder_paths:
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                file_mod_time = os.path.getmtime(file_path)
                file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                # Verifica si el archivo es PDF y si la fecha coincide
                if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                    # Clasifica según si el nombre contiene "turret"
                    if "turret" in file_path.lower():
                        turret_files.append(file_path)
                    else:
                        laser_plasma_files.append(file_path)

    # Imprime las listas de archivos
    print("Archivos 'turret':", turret_files)
    print("Archivos 'laser_plasma':", laser_plasma_files)
    
    return turret_files, laser_plasma_files

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\Funtionscode.py






# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)

def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []       
    

  


# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

def list_pdf_files_in_folder_combined(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []    

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionsGetdata.py
import os
from datetime import datetime, timedelta
import re

def get_today_date():
    today = datetime.today()
    # Restar un día
    yesterday = today - timedelta(days=5)
    # Retornar la fecha en el formato m_d_yyyy
    return yesterday.strftime("%m_%d_%Y")

def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

def convert_date_format_2(date_str):
    # Convertir la cadena en un objeto datetime usando el formato original
    date_obj = datetime.strptime(date_str, '%m_%d_%Y')
    # Formatear el objeto datetime al nuevo formato
    return date_obj.strftime('%m/%d/%Y')


def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"    


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionTurretFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

folder_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF

def process_pdfs_turret_and_generate_excel(fullpath, input_date):
    #all_dfs2 = []  # Lista para almacenar todos los DataFrames
    consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF
    #print(f"Archivos PDF a procesar: {fullpath}")
    
    for pdf_path in fullpath:
        print(f"Procesando el archivo: {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:    
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    #print(f"Procesando tabla: {table}")  # Esto imprimirá todas las tablas, incluso si no cumplen la condición

                    # Si la tabla cumple con la condición:
                    if table[0] == ['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')
                        print(df_final.columns)

                        df_final['Date process'] = input_date
                        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
                        df_final['Date Jake'] = match.group(1)                        
                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
                        df_final['Sheet Size'] = df_final['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
                        df_final['MACHINE'] = df_final['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))
                        df_final['Utilization'] = ''
                        df_final['Category'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
                        df_final['Material Name'] = df_final['SCHEDULE'].apply(lambda x: get_material(x, Material_dict))
                        df_final['Sheet Name'] = df_final['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        df_final['Date Jake'] = df_final['Date Jake'].apply(lambda x: convert_date_format_2(x))
                        df_final['Date process'] = df_final['Date process'].apply(lambda x: convert_date_format_2(x))


                        df_final.rename(columns={
                            'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                            'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                            'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                            'Cut Time': 'Cut Time'
                        }, inplace=True)

                        #print(f"df_final después de renombrar columnas: {df_final.head()}")

                        # Reordenar columnas
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]
                        #print("df_final después de reordenar columnas: ",df_final)

                        # Guardar el DataFrame en un archivo CSV
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)   


    #print(f"consolidated_df después de concatenar: {len(consolidated_df)}")
    #print(consolidated_df.head())
    #print(consolidated_df.columns)
    df_consolidado2=pd.DataFrame(consolidated_df) 
    '''try:
        output_file = os.path.join(excel_output_path, f"Utilization_T_{input_date}.xlsx")
        consolidated_df.to_excel(output_file, index=False)
        print("Éxito", f"Excel exportado correctamente: {output_file}")
    except Exception as e:
        print("Error", f"Error al guardar el archivo Excel: {e}")'''

    return df_consolidado2  

            


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\Old_version&testing\rpa_utilization_nest.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        '''try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")'''


############################################################################################################
# Función para listar archivos PDF turret en una carpeta y sus subcarpetas que contienen la fecha de modificación
############################################################################################################       
# Listar archivos PDF en la carpeta y subcarpetas con la fecha de modificación deseada
def list_pdf_files_turret_in_folder(folder_path, date_filter):
    pdf_files_turret = []
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            file_mod_time = os.path.getmtime(file_path)
            file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

            if date_filter == file_mod_date and file_name.lower().endswith(".pdf") and 'turret'  in file_path.lower():
                pdf_files_turret.append(file_path)
                category = get_category_from_path(file_path)
                nest = get_nest_from_path(file_path)
                date_process = file_mod_date
                


    return pdf_files_turret

# Función para procesar los PDFs y generar el archivo Excel
def process_pdfs_turret_and_generate_excel(pdf_files_turret, input_date):
     all_dfs = []  # Lista para almacenar todos los DataFrames
     for pdf_path in pdf_files_turret:
        category = get_category_from_path(pdf_path)
        #print(category)
        nest = get_nest_from_path(pdf_path)
        #print(nest)
        date_process = input_date
        #print(date_process)
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', nest)
        date_jake = match.group(1) if match else "MISC"
        date_process=date_process.replace("_","/")


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if   table[0] ==['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]  :
                                            
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')#.drop('key', axis=1)

                        df_final = df_final[
                            ['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME', 'Material Name', 'Sheet Code', 'Sheet Size', 
                            'Sheet Count', 'Sheet Name', 'Stack Qty', 'Cut Time', 'Finish']
                        ]
                        def adjust_machine(machine_name):
                            adjustments = {
                                "Amada_ENSIS_4020AJ": "1. Laser", "Messer_170Amp_Plasm": "2.Plasma", 
                                "Amada_Vipros_358K": "3.Turret", "Amada_EMK316M2": "4.Citurret"
                            }
                            return adjustments.get(machine_name, machine_name)
                        
                        df_final['MACHINE'] = df_final['MACHINE'].apply(adjust_machine)

                        def get_gauge_from_nesting(SCHEDULE, code_to_gauge):
                            for code, gauge in code_to_gauge.items():
                                if code in SCHEDULE:
                                    return gauge
                            return "Desconocido"

                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

                        def update_sheet_size(sheet_size):
                            valid_sizes = [
                                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
                            ]
                            cleaned_size = sheet_size.replace(" ", "").lower()
                            inverted_size = 'x'.join(cleaned_size.split('x')[::-1])  # Reverse size
                            return sheet_size if cleaned_size in valid_sizes or inverted_size in valid_sizes else "REMNANT"
                        

                        

                        df_final['Sheet Sizex'] = df_final['Sheet Size'].apply(update_sheet_size)
                        df_final['Date process'] = date_process
                        df_final['Date Jake'] = date_jake
                        df_final['Utilization']=''
                        def adjust_type_nest(Type):
                            adjustments = {
                                "ENCL": "2.Encl", "ENGR": "6.Engr", "MEC": "3.Mec", "PARTS ORDER": "4.Parts Order", "REWORK": "5.Rework",
                                "SIL": "4.Sil", "TANK": "1.Tank"
                            }
                            return adjustments.get(Type, Type)

                        df_final['Category'] = adjust_type_nest(category)

                        df_final.rename(columns={
                        'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                        'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                        'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                        'Cut Time': 'Cut Time'
                    }, inplace=True)
                        
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]

                        def get_material(nesting):
                            for material, values in category_dict.items():
                                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                                    return material
                            return "Desconocido"          

                        df_final['Material'] = df_final['Nesting'].apply(get_material)

                        df_final['Program'] = df_final['Program'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        
                        all_dfs.append(df_final)

                        # Consolidar todos los DataFrames en uno solo
        consolidated_df = pd.concat(all_dfs, ignore_index=True)

        # Guardar el DataFrame consolidado en un archivo CSV
        output_file = excel_output_path + 'Utilization_consolidated_' + str(input_date) + '.csv'
        consolidated_df.to_csv(output_file, index=False)
        print("Todos los archivos procesados y consolidados exitosamente.")
        print(consolidated_df) 



# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    pdf_files_turret = list_pdf_files_turret_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)
    process_pdfs_turret_and_generate_excel(pdf_files_turret, date_filter)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\config.py
#excel_output_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data"
logo_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret\image\Logo.png"
folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\main.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *



# Configuración de rutas
#folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
#folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
#excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'


# Función para obtener la fecha actual en formato m_d_yyyy

def main():
    inicio_proceso = datetime.now()
    input_date = get_today_date()
    _ ,laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date)
    df_consolidado1=process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, input_date)

    turret_files, _ = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date) 
    df_consolidado2 =process_pdfs_turret_and_generate_excel(turret_files, input_date)

    
    if df_consolidado1 is not None and df_consolidado2 is not None:        
        output_filename = f"Utilization_{input_date}.xlsx"
        outputfinal = pd.concat([df_consolidado1,df_consolidado2], axis=0)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is not None and df_consolidado2 is None: 
        outputfinal=pd.DataFrame(df_consolidado1)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is None and df_consolidado2 is not None: 
        outputfinal=pd.DataFrame(df_consolidado2)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}") 
    else:
        print("No se generó ningún archivo Excel.")

    fin_proceso = datetime.now()
    print(f"Tiempo de ejecución: {fin_proceso - inicio_proceso}")
if __name__ == "__main__":
    main()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\utilizacion.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path, logo_path  # Aquí cargamos las variables de configuración

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")


# Obtener las dimensiones de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Definir el tamaño de la ventana
window_width = 760
window_height = 900

# Calcular las coordenadas para centrar la ventana
position_top = int(screen_height / 2 - window_height / 2)
position_left = int(screen_width / 2 - window_width / 2)

# Configurar la geometría para centrar la ventana
root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

root.resizable(TRUE, TRUE)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#D3D3D3")  # Gris claro

# Eliminar el Canvas y trabajar directamente con un Label para el logo
# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente, aumentando su tamaño un 30%
logo_width = 104  # Aumento del 30% respecto a 80 (80 * 1.3 = 104)
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Usar un Label en lugar del Canvas para colocar el logo en la parte superior izquierda
logo_label = Label(root, image=logo_photo, bg="#D3D3D3")  # Fondo transparente
logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")  # Ubicarlo en la parte superior izquierda


# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#2F2F2F", bg="#D3D3D3", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Crear la barra de progreso
progress_var = tk.IntVar()
style = ttk.Style()
style.configure("TProgressbar",
                thickness=50,  # Aumenta el grosor de la barra
                )
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=600, mode="determinate",style="TProgressbar")
progress_bar.grid(row=6, column=0, columnspan=3, pady=10)

# Función para convertir la fecha de entrada en el formato adecuado
def convert_date_format(date_str):
    try:
        return datetime.strptime(date_str, '%m/%d/%Y').strftime('%m_%d_%Y')
    except ValueError:
        return None

# Función para actualizar la barra de progreso
def update_progress_bar(current, total):
    progress_var.set((current / total) * 100)
    root.update_idletasks()

# Función para actualizar el log con un mensaje
def log_message(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + "\n")
    log_text.yview(tk.END)  # Desplazar el log al final
    log_text.config(state=tk.DISABLED)

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Invalid date.")
        return

    if not folder_path_laser or not folder_path_plasma:
        messagebox.showerror("Error", "Por favor verifica las rutas de las carpetas de láser y plasma en config.py.")
        log_message("Error: Folder paths not found.")
        return

    # Listar archivos PDF de las dos carpetas
    log_message("Listing PDF files from the folders...")
    turret_files, laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser, folder_path_plasma, date_filter)
   
    
    if not laser_plasma_files:
        log_message("No PDF files found in the laser and plasma folders.")
    
    total_files = len(laser_plasma_files)
    print(total_files)

    log_message(f"{total_files} PDF files found in the laser and plasma folders.")

    total_files += len(turret_files)
    log_message(f"{len(turret_files)} PDF files found in the turret folders.")
    print(total_files)
    # Procesar PDF de laser y plasma si hay archivos disponibles
    if laser_plasma_files:
        log_message("Processing PDF files from the laser and plasma folders...")
        for file in laser_plasma_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado1 = process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, date_filter, update_progress_bar)
        if df_consolidado1 is None:
            log_message("The DataFrame for laser and plasma was not generated.")
    else:
        df_consolidado1 = None

    # Actualizar progreso después de cada procesamiento
    progress_var.set((len(laser_plasma_files) / total_files) * 100)
    root.update_idletasks()

    # Procesar PDF de turret si hay archivos disponibles
    if turret_files:
        log_message("Processing PDF files from the turret folders...")
        for file in turret_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado2 = process_pdfs_turret_and_generate_excel(turret_files, date_filter, update_progress_bar)
        if df_consolidado2 is None:
            log_message("The DataFrame for turret was not generated.")
    else:
        df_consolidado2 = None

    # Procesar y guardar el archivo Excel si se generaron ambos DataFrames
    if df_consolidado1 is not None or df_consolidado2 is not None:
        # Concatenar los DataFrames si existen
        outputfinal = pd.concat([df_consolidado1, df_consolidado2], axis=0, ignore_index=True) if df_consolidado1 is not None and df_consolidado2 is not None else df_consolidado1 if df_consolidado1 is not None else df_consolidado2
        output_filename = f"Utilization_{date_filter}.xlsx"
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        log_message(f"Excel file saved at: {os.path.join(excel_output_path, output_filename)}")
    else:
        log_message("No Excel file was generated.")
        messagebox.showerror("Error", "No Excel file was generated.")


    log_message("The processing has been completed.")
    messagebox.showinfo("Process completed", "The PDF file processing has been completed.\n                     😊😊😊😊😊")

        # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Close the application", "Do you want to close the application?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación


# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Ajustar el área del log
log_frame = tk.Frame(root, bg="#FFFFFF", height=600)  # Fondo blanco y altura mayor
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew", padx=(10, 20))  # Añadir margen a ambos lados (izquierda y derecha)

# Barra de desplazamiento del log
log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

# Área de texto del log
log_text = Text(log_frame, height=30, width=120 ,wrap="word", font=("Helvetica", 8), fg="#000000", bg="#FFFFFF", bd=1, insertbackground="black", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)

# Configurar la barra de desplazamiento
log_scrollbar.config(command=log_text.yview)



root.mainloop()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\Dictionaries.py
code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK3612M2": "4.Citurret"
}

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuncionesFieldsTransform.py
####################################################################################################
# Funciones para transformar los campos de un DataFrame
####################################################################################################
# Las funciones a continuación se utilizan para transformar los campos de un DataFrame de Pandas.
# Estas funciones se pueden aplicar a una columna específica para modificar los valores de las celdas.
# Por ejemplo, se puede aplicar una función para ajustar los nombres de las máquinas o para obtener
# el calibre de una celda que contiene un 'nesting'.
#
def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"



def update_sheet_size(sheet_size, valid_sizes):
    # Eliminar los espacios en la cadena y convertir todo a minúsculas
    cleaned_size = sheet_size.replace(" ", "").lower()
    
    # Invertir la cadena limpia para verificar la variante invertida
    inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
    
    # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
    if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
        return sheet_size  # Conservar el valor original si hay coincidencia
    else:
        return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
    
    
# Función para obtener el material desde el nesting
def get_material(nesting, Material_dict):
    for material, values in Material_dict.items():
        if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
            return material
    return "Desconocido"


def adjust_type_nest(Type, type_dict):
    # Buscar si el tipo está en el diccionario y devolver el valor correspondiente
    return type_dict.get(Type, Type)


def adjust_machine(machine_name, machine_dict):
    # Buscar si el nombre de la máquina está en el diccionario y devolver el valor correspondiente
    return machine_dict.get(machine_name, machine_name)  # Si machine_name no está en el diccionario, retorna el valor original    



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionLaserPlasmaFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

def process_pdfs_laser_plasma_and_generate_excel(pdf_files, input_date, update_progress_func):
    all_utilization_lines = []
    all_headers = []  # Lista de headers
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las fechas Jake de cada archivo

    headers = pd.DataFrame()  # Inicializamos headers como un DataFrame vacío para evitar el error

    total_files = len(pdf_files)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados

    for pdf_path in pdf_files:
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        print(f"Categoría para {pdf_path}: {category}")
        nest = get_nest_from_path(pdf_path)
        date_process = input_date
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake = match.group(1)     

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                        
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake) 
                                utilization_data.append(category) 
                                #print(f"utilization_data: {utilization_data}")
                                 # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
                                #print(f"all_utilization_lines: {all_utilization_lines[:5]}")
        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Si se encontraron headers y lines de utilización
    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size','Nest','Date process','Date Jake','category'])
    else:
        print("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")

    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

        # Agregar la columna "Material"
        merged_df['Material'] = merged_df['NESTING'].apply(lambda x: get_material(x, Material_dict)) 

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        merged_df['MACHINE'] = merged_df['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time'
            })

        # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization', 'Cut Time'
        ]

        # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]
        merged_df['Utilization'] = pd.to_numeric(merged_df['Utilization'], errors='coerce')

        # Luego realizar la conversión a decimal (dividir entre 100)
        merged_df['Utilization'] = merged_df['Utilization'] / 100

        # Ajustar la columna 'Type' según las condiciones especificadas
        #merged_df['Type'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
        df_consolidado1 = pd.DataFrame(merged_df)

        

        '''# Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            print("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            print("Error", f"Error al guardar el archivo Excel: {e}")
        '''
    else:
        print("Error", "No se pueden combinar los DataFrames.")
    
    return df_consolidado1


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionListFiles.py

import os
from datetime import datetime
import re

def list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter):
    turret_files = []
    laser_plasma_files = []

    # Combina las dos rutas de directorios
    folder_paths = [folder_path_1, folder_path_2]

    # Ahora iteramos sobre los directorios y procesamos los archivos
 
    for folder_path in folder_paths:
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:                
                file_path = os.path.join(root, file_name)
                file_mod_time = os.path.getmtime(file_path)
                file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")
                # Verifica si el archivo es PDF y si la fecha coincide
                if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                    #print(f"Procesando archivo: {file_path}")    
                    # Clasifica según si el nombre contiene "turret"
                    #if "turret" in file_path.lower():
                    if re.search(r"turret", file_path, re.IGNORECASE): 
                        print("turret",file_path)
                        turret_files.append(file_path)
                    else:
                        laser_plasma_files.append(file_path)
                        print("laser",file_path)

    # Imprime las listas de archivos
    print("Archivos 'turret':", len(turret_files))
    print("Archivos 'laser_plasma':", len(laser_plasma_files))

    return turret_files, laser_plasma_files

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\Funtionscode.py






# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)

def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []       
    

  


# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

def list_pdf_files_in_folder_combined(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []    

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionsGetdata.py
import os
from datetime import datetime, timedelta
import re
from funtions.Dictionaries import type_dict

def get_today_date():
    today = datetime.today()
    # Restar un día
    yesterday = today - timedelta(days=5)
    # Retornar la fecha en el formato m_d_yyyy
    return yesterday.strftime("%m_%d_%Y")

def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

def convert_date_format_2(date_str):
    # Convertir la cadena en un objeto datetime usando el formato original
    date_obj = datetime.strptime(date_str, '%m_%d_%Y')
    # Formatear el objeto datetime al nuevo formato
    return date_obj.strftime('%m/%d/%Y')

'''
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"'''


def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Obtener la categoría encontrada
        category = match.group(0)
        
        # Buscar la categoría en el diccionario y devolver el valor correspondiente
        return type_dict.get(category, "MISC")  # Si no se encuentra, devuelve "MISC"
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"

    
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"    


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionsInterfaz.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *


# Crear la ventana principal
root = ThemedTk(theme="arc")
root.title("Interfaz de Selección de Carpeta")

# Crear y colocar el label para mostrar la ruta de la carpeta seleccionada
folder_path_label = Label(root, text="No se ha seleccionado ninguna carpeta")
folder_path_label.pack(pady=10)

# Crear y colocar el widget de texto para el log
log_text = Text(root, state=tk.DISABLED, height=10, width=50)
log_text.pack(pady=10)

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionTurretFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

folder_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF

def process_pdfs_turret_and_generate_excel(fullpath, input_date, update_progress_func):
    consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF
    
    total_files = len(fullpath)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados
    
    for pdf_path in fullpath:
        print(f"Procesando el archivo: {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:    
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if table[0] == ['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        i = 3
                        while i < len(table):
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')
                        #print(df_final.columns)

                        df_final['Date process'] = input_date
                        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
                        df_final['Date Jake'] = match.group(1)                        
                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
                        #df_final['Sheet Size'] = df_final['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
                        df_final['MACHINE'] = df_final['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))
                        df_final['Utilization'] = ''
                        df_final['Category'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
                        df_final['Material Name'] = df_final['SCHEDULE'].apply(lambda x: get_material(x, Material_dict))
                        df_final['Sheet Name'] = df_final['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        df_final['Date Jake'] = df_final['Date Jake'].apply(lambda x: convert_date_format_2(x))
                        df_final['Date process'] = df_final['Date process'].apply(lambda x: convert_date_format_2(x))

                        df_final.rename(columns={
                            'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                            'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                            'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                            'Cut Time': 'Cut Time'
                        }, inplace=True)

                        # Reordenar columnas
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)

        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Devolver el DataFrame consolidado
    df_consolidado2 = pd.DataFrame(consolidated_df) 
    return df_consolidado2


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        '''try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")'''


############################################################################################################
# Función para listar archivos PDF turret en una carpeta y sus subcarpetas que contienen la fecha de modificación
############################################################################################################       
# Listar archivos PDF en la carpeta y subcarpetas con la fecha de modificación deseada
def list_pdf_files_turret_in_folder(folder_path, date_filter):
    pdf_files_turret = []
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            file_mod_time = os.path.getmtime(file_path)
            file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

            if date_filter == file_mod_date and file_name.lower().endswith(".pdf") and 'turret'  in file_path.lower():
                pdf_files_turret.append(file_path)
                category = get_category_from_path(file_path)
                nest = get_nest_from_path(file_path)
                date_process = file_mod_date
                


    return pdf_files_turret

# Función para procesar los PDFs y generar el archivo Excel
def process_pdfs_turret_and_generate_excel(pdf_files_turret, input_date):
     all_dfs = []  # Lista para almacenar todos los DataFrames
     for pdf_path in pdf_files_turret:
        category = get_category_from_path(pdf_path)
        #print(category)
        nest = get_nest_from_path(pdf_path)
        #print(nest)
        date_process = input_date
        #print(date_process)
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', nest)
        date_jake = match.group(1) if match else "MISC"
        date_process=date_process.replace("_","/")


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if   table[0] ==['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]  :
                                            
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')#.drop('key', axis=1)

                        df_final = df_final[
                            ['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME', 'Material Name', 'Sheet Code', 'Sheet Size', 
                            'Sheet Count', 'Sheet Name', 'Stack Qty', 'Cut Time', 'Finish']
                        ]
                        def adjust_machine(machine_name):
                            adjustments = {
                                "Amada_ENSIS_4020AJ": "1. Laser", "Messer_170Amp_Plasm": "2.Plasma", 
                                "Amada_Vipros_358K": "3.Turret", "Amada_EMK316M2": "4.Citurret"
                            }
                            return adjustments.get(machine_name, machine_name)
                        
                        df_final['MACHINE'] = df_final['MACHINE'].apply(adjust_machine)

                        def get_gauge_from_nesting(SCHEDULE, code_to_gauge):
                            for code, gauge in code_to_gauge.items():
                                if code in SCHEDULE:
                                    return gauge
                            return "Desconocido"

                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

                        def update_sheet_size(sheet_size):
                            valid_sizes = [
                                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
                            ]
                            cleaned_size = sheet_size.replace(" ", "").lower()
                            inverted_size = 'x'.join(cleaned_size.split('x')[::-1])  # Reverse size
                            return sheet_size if cleaned_size in valid_sizes or inverted_size in valid_sizes else "REMNANT"
                        

                        

                        df_final['Sheet Sizex'] = df_final['Sheet Size'].apply(update_sheet_size)
                        df_final['Date process'] = date_process
                        df_final['Date Jake'] = date_jake
                        df_final['Utilization']=''
                        def adjust_type_nest(Type):
                            adjustments = {
                                "ENCL": "2.Encl", "ENGR": "6.Engr", "MEC": "3.Mec", "PARTS ORDER": "4.Parts Order", "REWORK": "5.Rework",
                                "SIL": "4.Sil", "TANK": "1.Tank"
                            }
                            return adjustments.get(Type, Type)

                        df_final['Category'] = adjust_type_nest(category)

                        df_final.rename(columns={
                        'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                        'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                        'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                        'Cut Time': 'Cut Time'
                    }, inplace=True)
                        
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]

                        def get_material(nesting):
                            for material, values in category_dict.items():
                                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                                    return material
                            return "Desconocido"          

                        df_final['Material'] = df_final['Nesting'].apply(get_material)

                        df_final['Program'] = df_final['Program'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        
                        all_dfs.append(df_final)

                        # Consolidar todos los DataFrames en uno solo
        consolidated_df = pd.concat(all_dfs, ignore_index=True)

        # Guardar el DataFrame consolidado en un archivo CSV
        output_file = excel_output_path + 'Utilization_consolidated_' + str(input_date) + '.csv'
        consolidated_df.to_csv(output_file, index=False)
        print("Todos los archivos procesados y consolidados exitosamente.")
        print(consolidated_df) 



# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    pdf_files_turret = list_pdf_files_turret_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)
    process_pdfs_turret_and_generate_excel(pdf_files_turret, date_filter)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

    # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Cerrar la aplicación", "¿Deseas cerrar la aplicación?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)



# Esta es la línea que mantiene la interfaz de Tkinter en ejecución hasta que el usuario decida cerrarla.
root.mainloop() 

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\Old_version&testing\test.py

import os
from datetime import datetime
from funtions.FuntionListFiles import list_all_pdf_files_in_folders

folder_path_1 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_2 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
date_filter = "01_10_2025"

def update_progress_func(processed_files, total_files):
    progress = processed_files / total_files * 100
    print(f"Progreso: {progress:.2f}%")


list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter)

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\codigo_combinado.py
# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\ActualVersion\rpa_utilization_nest_202412301814.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"



def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Eliminar comillas dobles alrededor de la ruta
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)
logo_path = config.get('logo_path', None)

if excel_output_path:
    print(f'Output path set: {excel_output_path}')
else:
    print('The path "excel_output_path" was not found in the configuration file.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'The path exists: {excel_output_path}')
else:
    print('The specified path does not exist or is not valid.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Invalid date, make sure to use the MM/DD/YYYY format.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{1,2}_\d{1,2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF",
    "STEEL-0.375": "375 INCH"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Checking files in the folder and subfolders: '{folder_path}' with the date {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluding file:: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"File: {file_path} | Category: {category}")

            return pdf_files
        else:
            print(f"The path '{folder_path}' does not exist or is not a valid folder.")
            return []
    except Exception as e:
        print(f"Error listing files in the folder: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)
        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = round(float(utilization_data[3].replace("Utililization: ", "").strip().replace("%", ""),) / 100, 4)
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Warning", "No lines with 'Utilization' were found in the PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        #def get_material(nesting):
            #for material, values in category_dict.items():
                #if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    #return material

        #merged_df['Material'] = merged_df['NESTING'].apply(get_material)

        def get_material(nesting):
            # Asegurarnos de que nesting sea una cadena y no un valor NaN o None
            if isinstance(nesting, str):
                for material, values in category_dict.items():
                    if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                        return material
            # Si nesting no es una cadena, devolver "Desconocido"
            return "Desconocido"

        # Aplicar la función get_material a la columna 'NESTING' del DataFrame
        merged_df['Material'] = merged_df['NESTING'].apply(get_material)
    


      

        # Agregar la columna "Program"
        #merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        merged_df['Program'] = merged_df['Sheet Name'].apply(
                lambda x: str(x).split('-')[0] if isinstance(x, str) and '-' in x else str(x)
            )

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        #merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        merged_df['Gauge'] = merged_df['NESTING'].apply(
                lambda x: get_gauge_from_nesting(str(x), code_to_gauge)  # Convertir x a cadena
            )

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        #merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(str(x)))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Excel exported successfully: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the Excel file: {e}")
    else:
        messagebox.showerror("Error", "The DataFrames cannot be combined.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Select the input folder")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Invalid date. Make sure to use the MM/DD/YYYY format.")
        log_message("Error: Invalid date.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Please select an input folder.")
        log_message("Error: No input folder selected.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No PDF files found to process.")
        log_message("Error: No PDF files found.")
        return

    log_message(f"{len(pdf_files)} PDF files found to process.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Process finished", "The PDF file processing has been completed.")

    root.quit()  # Esto cierra la interfaz
    root.destroy()  # Esto termina la aplicación completamente

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Select the date:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Select folder", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Process PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\consolidate.py
import os
from PyPDF2 import PdfMerger
from datetime import datetime

# Ruta de la carpeta de entrada (ajusta según sea necesario)
folder_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2024\REWORK"
# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"

# Función para convertir la fecha de entrada al formato MM_dd_YYYY (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a MM_dd_YYYY
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")
        # Devolver la fecha como un string en el formato deseado
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None

# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha en el nombre de la carpeta
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        # Verifica que el directorio existe
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")
            
            # Usamos os.walk para recorrer todas las subcarpetas
            for root, dirs, files in os.walk(folder_path):
                # Comprobar si el nombre del folder (root) contiene la fecha
                if date_filter in os.path.basename(root).replace("\\", "/").lower():  # Compara el nombre de la carpeta con la fecha
                    for file_name in files:
                        file_path = os.path.join(root, file_name)
                        
                        # Excluir los archivos PDF cuyo path contenga "turret"
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue
                        
                        # Verificar que sea un archivo PDF (ignorando mayúsculas y minúsculas en la extensión)
                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            print(f"Archivo encontrado: {file_path}")
            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para consolidar los archivos PDF en un solo archivo
def consolidate_pdfs(folder_path, output_pdf, date_filter):
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    
    if not pdf_files:
        print("No se encontraron archivos PDF con la fecha especificada en la carpeta o sus subcarpetas.")
        return
    
    merger = PdfMerger()  # Crea un objeto PdfMerger para fusionar los PDFs
    
    try:
        # Iterar sobre cada archivo PDF y añadirlo al archivo final
        for pdf in pdf_files:
            print(f"Agregando {pdf} al archivo consolidado...")
            merger.append(pdf)  # Agrega el PDF al archivo final
        
        # Escribir el archivo PDF consolidado
        if os.path.exists(output_pdf):
            print(f"El archivo {output_pdf} ya existe. No se sobrescribirá.")
        else:
            merger.write(output_pdf)
            print(f"Archivos PDF consolidados exitosamente en: {output_pdf}")
        
        merger.close()
    
    except Exception as e:
        print(f"Error al consolidar los archivos PDF: {e}")

# Solicitar la fecha de entrada (ejemplo: 11/14/2024)
input_date = input("Ingresa la fecha (MM/DD/YYYY): ")

# Convertir la fecha al formato MM_dd_YYYY
date_filter = convert_date_format(input_date)

# Si la fecha es válida, proceder con la consolidación
if date_filter:
    # Nombre del archivo PDF de salida
    output_pdf = os.path.join(output_folder_path, f"consolidado_{date_filter}.pdf")

    # Llamada a la función para consolidar los PDFs
    consolidate_pdfs(folder_path, output_pdf, date_filter)  # Asegúrate de pasar la ruta de entrada, no la de salida


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\final_version.py
import os
from PyPDF2 import PdfMerger
from datetime import datetime
import pdfplumber
import pandas as pd

# Ruta de la carpeta de entrada (ajusta según sea necesario)
folder_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2024\ENCL"
# Ruta de la carpeta de salida para los archivos PDF consolidados
output_folder_path_pdf = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
# Ruta de la carpeta de salida para el archivo Excel
output_folder_path_excel = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Función para convertir la fecha de entrada al formato MM_dd_YYYY (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a MM_dd_YYYY
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")
        # Devolver la fecha como un string en el formato deseado
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None

# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha en el nombre de la carpeta
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        # Verifica que el directorio existe
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")
            
            # Usamos os.walk para recorrer todas las subcarpetas
            for root, dirs, files in os.walk(folder_path):
                # Comprobar si el nombre del folder (root) contiene la fecha
                if date_filter in os.path.basename(root).replace("\\", "/").lower():  # Compara el nombre de la carpeta con la fecha
                    for file_name in files:
                        file_path = os.path.join(root, file_name)
                        
                        # Excluir los archivos PDF cuyo path contenga "turret"
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue
                        
                        # Verificar que sea un archivo PDF (ignorando mayúsculas y minúsculas en la extensión)
                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            print(f"Archivo encontrado: {file_path}")
            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para consolidar los archivos PDF en un solo archivo
def consolidate_pdfs(folder_path, output_pdf, date_filter):
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    
    if not pdf_files:
        print("No se encontraron archivos PDF con la fecha especificada en la carpeta o sus subcarpetas.")
        return None
    
    merger = PdfMerger()  # Crea un objeto PdfMerger para fusionar los PDFs
    
    try:
        # Iterar sobre cada archivo PDF y añadirlo al archivo final
        for pdf in pdf_files:
            print(f"Agregando {pdf} al archivo consolidado...")
            merger.append(pdf)  # Agrega el PDF al archivo final
        
        # Escribir el archivo PDF consolidado
        if os.path.exists(output_pdf):
            print(f"El archivo {output_pdf} ya existe. No se sobrescribirá.")
        else:
            merger.write(output_pdf)
            print(f"Archivos PDF consolidados exitosamente en: {output_pdf}")
        
        merger.close()
        return output_pdf  # Devolver el path del archivo consolidado
    
    except Exception as e:
        print(f"Error al consolidar los archivos PDF: {e}")
        return None

# Función para buscar en tablas del PDF
def extract_tables_with_keyword(pdf_path, keyword):
    with pdfplumber.open(pdf_path) as pdf:
        matching_rows = []
        
        # Iterar por cada página
        for page_num, page in enumerate(pdf.pages):
            # Extraer las tablas de la página
            tables = page.extract_tables()
            
            # Si hay tablas en la página
            if tables:
                for table in tables:
                    for row in table:
                        # Buscar la palabra clave en las filas de la tabla
                        if any(keyword in str(cell) for cell in row if cell):  # Verifica que la celda no sea None
                            matching_rows.append(row)  # Agrega la fila completa
        
        return matching_rows

# Función principal para ejecutar el flujo completo
def process_pdfs_and_generate_excel(folder_path, output_folder_path_pdf, output_folder_path_excel, input_date, keyword):
    # Convertir la fecha al formato MM_dd_YYYY
    date_filter = convert_date_format(input_date)

    if date_filter:
        # Nombre del archivo PDF consolidado
        output_pdf = os.path.join(output_folder_path_pdf, f"consolidado_{date_filter}.pdf")

        # Consolidar los PDFs que cumplen con la fecha
        consolidated_pdf_path = consolidate_pdfs(folder_path, output_pdf, date_filter)
        
        if consolidated_pdf_path:
            # Extraer las tablas que contienen la palabra clave
            rows_with_keyword = extract_tables_with_keyword(consolidated_pdf_path, keyword)

            # Crear un DataFrame para almacenar los resultados
            columns = ["Nesting", "Columna 2", "Material", "Columna 4", "Quantity", "Utilization", "Columna 7", "Size"]
            results_df = pd.DataFrame(rows_with_keyword, columns=columns)

            # Eliminar columnas no deseadas
            results_df = results_df.drop(columns=["Columna 2", "Columna 4", "Columna 7"])

            # Renombrar columnas
            results_df.rename(columns={
                "Nesting": "Nesting",
                "Material": "Material",
                "Quantity": "Quantity",
                "Utilization": "Utilization",
                "Size": "Size"
            }, inplace=True)

            # Limpiar datos en las columnas Quantity y Utilization
            results_df["Quantity"] = results_df["Quantity"].str.replace("Stack Qty: ", "", regex=False)
            results_df["Utilization"] = results_df["Utilization"].str.replace("Utililization: ", "", regex=False)

            # Generar la fecha de ejecución para el nombre del archivo
            execution_date = datetime.now().strftime("%Y-%m-%d")
            output_file_name = f"Utilization_{execution_date}.xlsx"
            output_excel_path = os.path.join(output_folder_path_excel, output_file_name)

            # Exportar los resultados a un archivo Excel
            if not results_df.empty:
                results_df.to_excel(output_excel_path, index=False)
                print(f"Resultados exportados a Excel en: {output_excel_path}")
            else:
                print(f"No se encontraron filas con la palabra '{keyword}'.")
        else:
            print("No se pudo consolidar los archivos PDF.")
    else:
        print("Fecha no válida. No se puede continuar.")

# Solicitar la fecha de entrada (ejemplo: 11/14/2024)
input_date = input("Ingresa la fecha (MM/DD/YYYY): ")
# Palabra clave a buscar en las tablas del PDF
keyword = "Utililization"

# Ejecutar el proceso completo
process_pdfs_and_generate_excel(folder_path, output_folder_path_pdf, output_folder_path_excel, input_date, keyword)


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\parts.py
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime
from tqdm import tqdm  # Librería para la barra de progreso

# Función para extraer todas las tablas del PDF y filtrar líneas específicas
def extract_filtered_tables(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_matching_rows = []

        # Configurar la barra de progreso para el número de páginas en el PDF
        total_pages = len(pdf.pages)
        print(f"Procesando {total_pages} páginas...")

        # Iterar por cada página con la barra de progreso
        for page_num, page in enumerate(tqdm(pdf.pages, desc="Procesando páginas", unit="página")):
            # Extraer tablas de la página
            tables = page.extract_tables()
            footer_text = page.extract_text()  # Extraer texto completo de la página para obtener el pie de página

            # Obtener el pie de página
            footer_lines = footer_text.split('\n')[-1] if footer_text else "Sin pie de página"  # Tomar la última línea como pie de página

            # Extraer campos del pie de página según la estructura dada
            footer_parts = footer_lines.split(' ')
            if len(footer_parts) >= 4:
                date = footer_parts[0]  # Fecha
                time = footer_parts[1] + ' ' + footer_parts[2]  # Hora (am/pm)
                nesting_full = ' '.join(footer_parts[3:-2])  # Nesting completo, sin el último número
                # Asegurarse de que nesting tenga tres caracteres después del punto
                try:
                    nesting = nesting_full.split('.')[0] + '.' + nesting_full.split('.')[1][:3]
                except IndexError:
                    nesting = nesting_full
            else:
                date, time, nesting = ["Sin dato"] * 3  # Si no hay datos, colocar "Sin dato"

            # Si hay tablas en la página
            if tables:
                for table in tables:
                    for row in table[1:]:  # Omite la fila de encabezados
                        # Filtrar filas donde al menos una celda comience con dos letras minúsculas
                        if any(re.match(r'^[a-z]{2}', str(cell)) for cell in row if cell):
                            # Eliminar registros que contengan letras en la columna específica (por ejemplo, columna 1)
                            if not any(char.isalpha() for char in str(row[0])):  # Cambia el índice según la columna que desees verificar
                                # Eliminar los dos primeros caracteres de 'Part Name'
                                part_name_full = row[1][2:] if len(row[1]) > 2 else row[1]  # Ajustar índice según la columna

                                # Dividir 'Part Name' en 'Projecto' y 'Part Name' usando el primer espacio
                                split_name = part_name_full.split(' ', 1)
                                project = split_name[0] if len(split_name) > 0 else ''
                                part_name = split_name[1] if len(split_name) > 1 else ''

                                # Construir la nueva fila con los datos del pie de página (sin columna 1)
                                new_row = [project, part_name, row[3], date, time, nesting]
                                all_matching_rows.append(new_row)

        return all_matching_rows

# Nueva ruta del archivo PDF
pdf_file_path = r"C:\Users\User\OneDrive - globalpowercomponents\Documents\Utilization\Consolidated_PDF\consolidado.pdf"

# Extraer y filtrar las tablas del PDF
matching_rows = extract_filtered_tables(pdf_file_path)

# Crear un DataFrame para almacenar los resultados
if matching_rows:
    # Generar encabezados, eliminando la columna 1 y ajustando los nombres de las columnas
    columns = ["Projecto", "Part Name", "Quantity Total", "Fecha", "Hora", "Nesting"]
    
    results_df = pd.DataFrame(matching_rows, columns=columns)

    # Eliminar registros duplicados
    results_df = results_df.drop_duplicates()

    # Obtener la fecha actual en formato mm-dd-yyyy
    current_date = datetime.now().strftime("%m-%d-%Y")

    # Definir la ruta para guardar el archivo Excel con el nombre adecuado
    output_folder_path = r"C:\Users\User\OneDrive - globalpowercomponents\Documents\Utilization\Parts_Request_Reports"
    output_excel_name = f"Nesting_report_{current_date}.xlsx"
    output_excel_path = os.path.join(output_folder_path, output_excel_name)

    # Exportar los resultados a un archivo Excel
    results_df.to_excel(output_excel_path, index=False)
    print(f"Resultados filtrados exportados a Excel en: {output_excel_path}")
else:
    print("No se encontraron tablas que cumplan con el criterio en el PDF.")


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\procces_v1.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas


# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función principal que se ejecuta al presionar el botón
def main():
    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        return

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    canvas.create_window(300, 400, window=progress_bar)

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    canvas.create_window(300, 430, window=percentage_label)

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x450")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo negro
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo
canvas = Canvas(root, width=600, height=450, bg="#2F2F2F")
canvas.pack(fill="both", expand=True)

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior izquierda con un margen pequeño
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
canvas.create_window(300, 60, window=date_label)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
canvas.create_window(300, 120, window=date_calendar)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
canvas.create_window(300, 240, window=folder_path_label)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
canvas.create_window(300, 280, window=select_folder_button)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
canvas.create_window(300, 350, window=process_button)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\procces_v2.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re

# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

  

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función principal que se ejecuta al presionar el botón
def main():
    # Obtener la fecha de entrada desde la caja de texto
    input_date = date_entry.get()
    date_filter = convert_date_format(input_date)

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        return

    # Listar archivos PDF
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        return

    # Configurar la barra de progreso
    progress_var = DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400)
    progress_bar.pack(pady=10)
    percentage_label = Label(root, text="0%")
    percentage_label.pack()

    # Procesar los PDFs y generar el archivo Excel
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = Tk()
root.title("Consolidar y Procesar PDF")

# Fijar el tamaño de la ventana
root.geometry("600x400")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Etiqueta y campo de entrada para la fecha
date_label = Label(root, text="Fecha (MM/DD/YYYY):")
date_label.pack()
date_entry = Entry(root)
date_entry.pack()

# Etiqueta y botón para seleccionar la carpeta
folder_label = Label(root, text="Carpeta de entrada:")
folder_label.pack()
folder_path_label = Label(root, text="")
folder_path_label.pack()
select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder)
select_folder_button.pack()

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main)
process_button.pack()

root.mainloop()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\process_202412301056.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\process_Copy.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"



def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Eliminar comillas dobles alrededor de la ruta
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)
logo_path = config.get('logo_path', None)

if excel_output_path:
    print(f'Output path set: {excel_output_path}')
else:
    print('The path "excel_output_path" was not found in the configuration file.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'The path exists: {excel_output_path}')
else:
    print('The specified path does not exist or is not valid.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Invalid date, make sure to use the MM/DD/YYYY format.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{1,2}_\d{1,2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Checking files in the folder and subfolders: '{folder_path}' with the date {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluding file:: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"File: {file_path} | Category: {category}")

            return pdf_files
        else:
            print(f"The path '{folder_path}' does not exist or is not a valid folder.")
            return []
    except Exception as e:
        print(f"Error listing files in the folder: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)
        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Warning", "No lines with 'Utilization' were found in the PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        #def get_material(nesting):
            #for material, values in category_dict.items():
                #if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    #return material

        #merged_df['Material'] = merged_df['NESTING'].apply(get_material)

        def get_material(nesting):
            # Asegurarnos de que nesting sea una cadena y no un valor NaN o None
            if isinstance(nesting, str):
                for material, values in category_dict.items():
                    if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                        return material
            # Si nesting no es una cadena, devolver "Desconocido"
            return "Desconocido"

        # Aplicar la función get_material a la columna 'NESTING' del DataFrame
        merged_df['Material'] = merged_df['NESTING'].apply(get_material)
    


      

        # Agregar la columna "Program"
        #merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        merged_df['Program'] = merged_df['Sheet Name'].apply(
                lambda x: str(x).split('-')[0] if isinstance(x, str) and '-' in x else str(x)
            )

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        #merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        merged_df['Gauge'] = merged_df['NESTING'].apply(
                lambda x: get_gauge_from_nesting(str(x), code_to_gauge)  # Convertir x a cadena
            )

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        #merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(str(x)))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Excel exported successfully: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the Excel file: {e}")
    else:
        messagebox.showerror("Error", "The DataFrames cannot be combined.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Select the input folder")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Invalid date. Make sure to use the MM/DD/YYYY format.")
        log_message("Error: Invalid date.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Please select an input folder.")
        log_message("Error: No input folder selected.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No PDF files found to process.")
        log_message("Error: No PDF files found.")
        return

    log_message(f"{len(pdf_files)} PDF files found to process.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Process finished", "The PDF file processing has been completed.")

    root.quit()  # Esto cierra la interfaz
    root.destroy()  # Esto termina la aplicación completamente

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Select the date:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Select folder", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Process PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\process_old_version_ok.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)
output_folder_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\PDF_consolidated"
excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x500")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\prueba_extraccion_tablaspdf.py
import pdfplumber
import pandas as pd

# Define la ruta del archivo PDF
pdf_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\11-19-2024-N23140217ENCL.PDF'

# Define la estructura esperada de la primera fila
expected_first_row = ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]

# Abre el archivo PDF
with pdfplumber.open(pdf_path) as pdf:
    all_headers = []
    all_utilization_lines = []  # Para almacenar las líneas extraídas que contienen 'Utililization'

    for page_number, page in enumerate(pdf.pages, start=1):
        # Extrae las tablas de cada página
        tables = page.extract_tables()

        for table_number, table in enumerate(tables, start=1):
            if table and table[0] == expected_first_row:  # Verifica que la primera fila coincida con la estructura esperada
                # Extrae la primera fila de datos después de los encabezados
                data = table[1]  # Extrae la fila 2
                print(f"Fila extraída de la tabla {table_number} en la página {page_number}: {data}")
                # Asegúrate de que los datos se asignan correctamente a las columnas
                machine = data[0]
                schedule = data[2]
                total_cut_time = data[8]  # Ajusta esta posición si es necesario
                headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                all_headers.append(headers_df)

        # Buscar y extraer la línea que contenga "Utililization"
        for table in tables:
            for row in table:
                for cell in row:
                    if cell and "Utililization" in str(cell):  # Busca "Utililization" en cada celda
                        # Si encuentra la palabra, extrae solo los índices 0, 2, 4, 5, 7
                        utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                        
                        # Modificar el valor de los índices según los requisitos
                        # 1. Eliminar salto de línea en el índice 0 y renombrar la columna a 'NESTING'
                        utilization_data[0] = utilization_data[0].replace("\n", "").strip()  # Eliminar salto de línea
                        
                        # 2. Eliminar "Stack Qty: " en el índice 4
                        utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")  # Eliminar "Stack Qty: "
                        
                        # 3. Eliminar "Utililization: " en el índice 5 y convertir a número
                        utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")  # Convertir a número

                        all_utilization_lines.append(utilization_data)  # Almacena la línea filtrada

# Combina las tablas de encabezados en un solo DataFrame
if all_headers:
    headers = pd.concat(all_headers, ignore_index=True)
    # Imprime el DataFrame de encabezados en la consola
    print("Headers DataFrame:")
    print(headers)
else:
    print("No se encontraron tablas con los encabezados esperados en el PDF.")

# Imprime las líneas encontradas con 'Utililization'
if all_utilization_lines:
    utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size'])
    print("\nLíneas con 'Utililization':")
    print(utilization_df)
else:
    print("No se encontraron líneas con 'Utililization' en el PDF.")

# Realizar el merge entre los dos DataFrames usando la columna 'SCHEDULE' y 'NESTING'
if not headers.empty and not utilization_df.empty:
    # Asegúrate de que ambas columnas tengan el mismo tipo de datos (string)
    headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
    utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
    
    # Merge de los DataFrames
    merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')

    # Eliminar la columna 'SCHEDULE' después del merge
    merged_df.drop(columns=['SCHEDULE'], inplace=True)

    # Extraer los valores de 'NESTING' y crear las nuevas columnas 'Material' y 'Gauge'
    # 1. Extraer Material: caracteres después del cuarto guion y antes del último
    def extract_material(nesting):
        # Split the string by '-'
        parts = nesting.split('-')
        
        # Si el valor en el cuarto guion es alfanumérico (ej: 3000-ALU), tomar ese rango
        if len(parts) > 4:
            material = '-'.join(parts[4:6])  # Concatenar partes 4 y 5 para formar '3000-ALU'
            return material
        # Si no hay guiones extras, simplemente devolver el valor después del cuarto guion
        return parts[4] if len(parts) > 4 else ''

    merged_df['Material'] = merged_df['NESTING'].apply(extract_material)

    # 2. Extraer Gauge: caracteres después del último guion
    merged_df['Gauge'] = merged_df['NESTING'].str.split('-').str[-1]  # Extrae la parte después del último guion

    # Imprime el DataFrame combinado con las nuevas columnas
    print("\nDataFrame combinado y modificado:")
    print(merged_df)
else:
    print("No se pueden combinar los DataFrames. Asegúrate de que ambos contengan datos.")


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\prueba_interfaz.py
import customtkinter as ctk

root = ctk.CTk()
root.geometry("400x400")

button = ctk.CTkButton(root, text="Iniciar proceso")
button.pack(pady=20)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\rpa_utilization_nest.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.0_20241230_T1814\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\config.py
#excel_output_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data"
logo_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret\image\Logo.png"
folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\main.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *



# Configuración de rutas
#folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
#folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
#excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'


# Función para obtener la fecha actual en formato m_d_yyyy

def main():
    inicio_proceso = datetime.now()
    input_date = get_today_date()
    _ ,laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date)
    df_consolidado1=process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, input_date)

    turret_files, _ = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date) 
    df_consolidado2 =process_pdfs_turret_and_generate_excel(turret_files, input_date)

    
    if df_consolidado1 is not None and df_consolidado2 is not None:        
        output_filename = f"Utilization_{input_date}.xlsx"
        outputfinal = pd.concat([df_consolidado1,df_consolidado2], axis=0)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is not None and df_consolidado2 is None: 
        outputfinal=pd.DataFrame(df_consolidado1)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is None and df_consolidado2 is not None: 
        outputfinal=pd.DataFrame(df_consolidado2)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}") 
    else:
        print("No se generó ningún archivo Excel.")

    fin_proceso = datetime.now()
    print(f"Tiempo de ejecución: {fin_proceso - inicio_proceso}")
if __name__ == "__main__":
    main()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\ActualVersion\rpa_utilization_nest_202412301814.py
import os
import re
import fitz  # PyMuPDF

def buscar_archivos_con_turret(carpeta):
    archivos_con_turret = []
    
    for raiz, dirs, archivos in os.walk(carpeta):
        for archivo in archivos:
            # Verificamos si la palabra 'turret' está en la ruta completa y el archivo es PDF
            if 'turret' in os.path.join(raiz, archivo).lower() and archivo.lower().endswith('.pdf'):
                archivos_con_turret.append(os.path.join(raiz, archivo))
    
    return archivos_con_turret

def extraer_tabla_machine_pdf(archivo_pdf):
    """
    Busca el encabezado 'MACHINE' en el archivo PDF y extrae los valores correspondientes de la tabla.
    """
    tabla_datos = []
    
    try:
        # Abrir el archivo PDF
        doc = fitz.open(archivo_pdf)
        
        # Iterar sobre todas las páginas del PDF
        for pagina in doc:
            texto = pagina.get_text("text")  # Extraer el texto de la página
            
            # Dividir el texto por líneas
            lineas = texto.split('\n')
            
            # Buscar el encabezado 'MACHINE' y extraer los valores correspondientes de la tabla
            for i, linea in enumerate(lineas):
                if 'MACHINE' in linea:
                    # Aquí buscamos la siguiente línea para extraer los valores
                    if i + 1 < len(lineas):
                        # Obtener la información de la siguiente línea que contiene los valores
                        datos = lineas[i + 1].strip()
                        tabla_datos.append(datos)
                        
                        # Si se encuentran otras filas relacionadas con la tabla, podemos extraerlas también
                        # Esto depende de la estructura del PDF y las líneas que siguen después
                        # Ejemplo: buscar la siguiente línea con los datos
                        for j in range(i + 2, len(lineas)):
                            # Si encontramos un patrón de datos esperado (por ejemplo, "Material Name"), extraemos más información
                            if re.match(r"\S+", lineas[j].strip()):
                                tabla_datos.append(lineas[j].strip())
                    break  # Salir si encontramos la tabla
    except Exception as e:
        print(f"Error al leer el archivo PDF {archivo_pdf}: {e}")
    
    return tabla_datos

# Carpeta donde buscar los archivos
carpeta = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2024\REWORK\12_27_2024_REWORK 5'

# Obtenemos los archivos PDF que contienen 'turret' en la ruta
archivos_pdf = buscar_archivos_con_turret(carpeta)

# Ahora vamos a obtener la información de esos archivos PDF
for archivo in archivos_pdf:
    print(f"Buscando la tabla 'MACHINE' en el archivo PDF: {archivo}")
    tabla_machine = extraer_tabla_machine_pdf(archivo)
    if tabla_machine:
        print("Datos extraídos de la tabla:")
        for linea in tabla_machine:
            print(linea)
    else:
        print("No se encontró la tabla 'MACHINE' en el archivo.")


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\Dictionaries.py
code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK316M2": "4.Citurret"
}

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuncionesFieldsTransform.py
####################################################################################################
# Funciones para transformar los campos de un DataFrame
####################################################################################################
# Las funciones a continuación se utilizan para transformar los campos de un DataFrame de Pandas.
# Estas funciones se pueden aplicar a una columna específica para modificar los valores de las celdas.
# Por ejemplo, se puede aplicar una función para ajustar los nombres de las máquinas o para obtener
# el calibre de una celda que contiene un 'nesting'.
#
def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"



def update_sheet_size(sheet_size, valid_sizes):
    # Eliminar los espacios en la cadena y convertir todo a minúsculas
    cleaned_size = sheet_size.replace(" ", "").lower()
    
    # Invertir la cadena limpia para verificar la variante invertida
    inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
    
    # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
    if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
        return sheet_size  # Conservar el valor original si hay coincidencia
    else:
        return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
    
    
# Función para obtener el material desde el nesting
def get_material(nesting, Material_dict):
    for material, values in Material_dict.items():
        if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
            return material
    return "Desconocido"


def adjust_type_nest(Type, type_dict):
    # Buscar si el tipo está en el diccionario y devolver el valor correspondiente
    return type_dict.get(Type, Type)


def adjust_machine(machine_name, machine_dict):
    # Buscar si el nombre de la máquina está en el diccionario y devolver el valor correspondiente
    return machine_dict.get(machine_name, machine_name)  # Si machine_name no está en el diccionario, retorna el valor original    



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionLaserPlasmaFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

def process_pdfs_laser_plasma_and_generate_excel(pdf_files, input_date):
    all_utilization_lines = []
    all_headers = []  # Lista de headers
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las fechas Jake de cada archivo

    headers = pd.DataFrame()  # Inicializamos headers como un DataFrame vacío para evitar el error

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process = input_date
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake = match.group(1)     

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
    # Si se encontraron headers y lines de utilización
    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        print("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")

    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

        # Agregar la columna "Material"

        merged_df['Material'] = merged_df['NESTING'].apply(lambda x: get_material(x, Material_dict)) 

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))


        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time'
            })

        # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization', 'Cut Time'
        ]

        # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

        # Ajustar la columna 'Type' según las condiciones especificadas
        #merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  
        merged_df['Type'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
        df_consolidado1=pd.DataFrame(merged_df)

        '''# Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            print("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            print("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        print("Error", "No se pueden combinar los DataFrames.")'''
    return df_consolidado1  


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionListFiles.py
import os
from datetime import datetime
import re

def list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter):
    turret_files = []
    laser_plasma_files = []

    # Combina las dos rutas de directorios
    folder_paths = [folder_path_1, folder_path_2]

    # Itera sobre ambos directorios
    for folder_path in folder_paths:
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                file_mod_time = os.path.getmtime(file_path)
                file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                # Verifica si el archivo es PDF y si la fecha coincide
                if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                    # Clasifica según si el nombre contiene "turret"
                    if "turret" in file_path.lower():
                        turret_files.append(file_path)
                    else:
                        laser_plasma_files.append(file_path)

    # Imprime las listas de archivos
    print("Archivos 'turret':", turret_files)
    print("Archivos 'laser_plasma':", laser_plasma_files)
    
    return turret_files, laser_plasma_files

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\Funtionscode.py






# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)

def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []       
    

  


# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

def list_pdf_files_in_folder_combined(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []    

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionsGetdata.py
import os
from datetime import datetime, timedelta
import re

def get_today_date():
    today = datetime.today()
    # Restar un día
    yesterday = today - timedelta(days=5)
    # Retornar la fecha en el formato m_d_yyyy
    return yesterday.strftime("%m_%d_%Y")

def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

def convert_date_format_2(date_str):
    # Convertir la cadena en un objeto datetime usando el formato original
    date_obj = datetime.strptime(date_str, '%m_%d_%Y')
    # Formatear el objeto datetime al nuevo formato
    return date_obj.strftime('%m/%d/%Y')


def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"    


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\funtions\FuntionTurretFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

folder_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF

def process_pdfs_turret_and_generate_excel(fullpath, input_date):
    #all_dfs2 = []  # Lista para almacenar todos los DataFrames
    consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF
    #print(f"Archivos PDF a procesar: {fullpath}")
    
    for pdf_path in fullpath:
        print(f"Procesando el archivo: {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:    
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    #print(f"Procesando tabla: {table}")  # Esto imprimirá todas las tablas, incluso si no cumplen la condición

                    # Si la tabla cumple con la condición:
                    if table[0] == ['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')
                        print(df_final.columns)

                        df_final['Date process'] = input_date
                        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
                        df_final['Date Jake'] = match.group(1)                        
                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
                        df_final['Sheet Size'] = df_final['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
                        df_final['MACHINE'] = df_final['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))
                        df_final['Utilization'] = ''
                        df_final['Category'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
                        df_final['Material Name'] = df_final['SCHEDULE'].apply(lambda x: get_material(x, Material_dict))
                        df_final['Sheet Name'] = df_final['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        df_final['Date Jake'] = df_final['Date Jake'].apply(lambda x: convert_date_format_2(x))
                        df_final['Date process'] = df_final['Date process'].apply(lambda x: convert_date_format_2(x))


                        df_final.rename(columns={
                            'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                            'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                            'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                            'Cut Time': 'Cut Time'
                        }, inplace=True)

                        #print(f"df_final después de renombrar columnas: {df_final.head()}")

                        # Reordenar columnas
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]
                        #print("df_final después de reordenar columnas: ",df_final)

                        # Guardar el DataFrame en un archivo CSV
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)   


    #print(f"consolidated_df después de concatenar: {len(consolidated_df)}")
    #print(consolidated_df.head())
    #print(consolidated_df.columns)
    df_consolidado2=pd.DataFrame(consolidated_df) 
    '''try:
        output_file = os.path.join(excel_output_path, f"Utilization_T_{input_date}.xlsx")
        consolidated_df.to_excel(output_file, index=False)
        print("Éxito", f"Excel exportado correctamente: {output_file}")
    except Exception as e:
        print("Error", f"Error al guardar el archivo Excel: {e}")'''

    return df_consolidado2  

            


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\Old_version&testing\rpa_utilization_nest.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V1.1_20250115_T1930\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        '''try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")'''


############################################################################################################
# Función para listar archivos PDF turret en una carpeta y sus subcarpetas que contienen la fecha de modificación
############################################################################################################       
# Listar archivos PDF en la carpeta y subcarpetas con la fecha de modificación deseada
def list_pdf_files_turret_in_folder(folder_path, date_filter):
    pdf_files_turret = []
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            file_mod_time = os.path.getmtime(file_path)
            file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

            if date_filter == file_mod_date and file_name.lower().endswith(".pdf") and 'turret'  in file_path.lower():
                pdf_files_turret.append(file_path)
                category = get_category_from_path(file_path)
                nest = get_nest_from_path(file_path)
                date_process = file_mod_date
                


    return pdf_files_turret

# Función para procesar los PDFs y generar el archivo Excel
def process_pdfs_turret_and_generate_excel(pdf_files_turret, input_date):
     all_dfs = []  # Lista para almacenar todos los DataFrames
     for pdf_path in pdf_files_turret:
        category = get_category_from_path(pdf_path)
        #print(category)
        nest = get_nest_from_path(pdf_path)
        #print(nest)
        date_process = input_date
        #print(date_process)
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', nest)
        date_jake = match.group(1) if match else "MISC"
        date_process=date_process.replace("_","/")


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if   table[0] ==['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]  :
                                            
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')#.drop('key', axis=1)

                        df_final = df_final[
                            ['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME', 'Material Name', 'Sheet Code', 'Sheet Size', 
                            'Sheet Count', 'Sheet Name', 'Stack Qty', 'Cut Time', 'Finish']
                        ]
                        def adjust_machine(machine_name):
                            adjustments = {
                                "Amada_ENSIS_4020AJ": "1. Laser", "Messer_170Amp_Plasm": "2.Plasma", 
                                "Amada_Vipros_358K": "3.Turret", "Amada_EMK316M2": "4.Citurret"
                            }
                            return adjustments.get(machine_name, machine_name)
                        
                        df_final['MACHINE'] = df_final['MACHINE'].apply(adjust_machine)

                        def get_gauge_from_nesting(SCHEDULE, code_to_gauge):
                            for code, gauge in code_to_gauge.items():
                                if code in SCHEDULE:
                                    return gauge
                            return "Desconocido"

                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

                        def update_sheet_size(sheet_size):
                            valid_sizes = [
                                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
                            ]
                            cleaned_size = sheet_size.replace(" ", "").lower()
                            inverted_size = 'x'.join(cleaned_size.split('x')[::-1])  # Reverse size
                            return sheet_size if cleaned_size in valid_sizes or inverted_size in valid_sizes else "REMNANT"
                        

                        

                        df_final['Sheet Sizex'] = df_final['Sheet Size'].apply(update_sheet_size)
                        df_final['Date process'] = date_process
                        df_final['Date Jake'] = date_jake
                        df_final['Utilization']=''
                        def adjust_type_nest(Type):
                            adjustments = {
                                "ENCL": "2.Encl", "ENGR": "6.Engr", "MEC": "3.Mec", "PARTS ORDER": "4.Parts Order", "REWORK": "5.Rework",
                                "SIL": "4.Sil", "TANK": "1.Tank"
                            }
                            return adjustments.get(Type, Type)

                        df_final['Category'] = adjust_type_nest(category)

                        df_final.rename(columns={
                        'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                        'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                        'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                        'Cut Time': 'Cut Time'
                    }, inplace=True)
                        
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]

                        def get_material(nesting):
                            for material, values in category_dict.items():
                                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                                    return material
                            return "Desconocido"          

                        df_final['Material'] = df_final['Nesting'].apply(get_material)

                        df_final['Program'] = df_final['Program'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        
                        all_dfs.append(df_final)

                        # Consolidar todos los DataFrames en uno solo
        consolidated_df = pd.concat(all_dfs, ignore_index=True)

        # Guardar el DataFrame consolidado en un archivo CSV
        output_file = excel_output_path + 'Utilization_consolidated_' + str(input_date) + '.csv'
        consolidated_df.to_csv(output_file, index=False)
        print("Todos los archivos procesados y consolidados exitosamente.")
        print(consolidated_df) 



# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    pdf_files_turret = list_pdf_files_turret_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)
    process_pdfs_turret_and_generate_excel(pdf_files_turret, date_filter)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)

root.mainloop()

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\config.py
#excel_output_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data"
logo_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret\image\Logo.png"
folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\main.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *



# Configuración de rutas
#folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
#folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
#excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'


# Función para obtener la fecha actual en formato m_d_yyyy

def main():
    inicio_proceso = datetime.now()
    input_date = get_today_date()
    _ ,laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date)
    df_consolidado1=process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, input_date)

    turret_files, _ = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date) 
    df_consolidado2 =process_pdfs_turret_and_generate_excel(turret_files, input_date)

    
    if df_consolidado1 is not None and df_consolidado2 is not None:        
        output_filename = f"Utilization_{input_date}.xlsx"
        outputfinal = pd.concat([df_consolidado1,df_consolidado2], axis=0)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is not None and df_consolidado2 is None: 
        outputfinal=pd.DataFrame(df_consolidado1)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is None and df_consolidado2 is not None: 
        outputfinal=pd.DataFrame(df_consolidado2)
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}") 
    else:
        print("No se generó ningún archivo Excel.")

    fin_proceso = datetime.now()
    print(f"Tiempo de ejecución: {fin_proceso - inicio_proceso}")
if __name__ == "__main__":
    main()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\utilizacion.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path, logo_path  # Aquí cargamos las variables de configuración

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")


# Obtener las dimensiones de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Definir el tamaño de la ventana
window_width = 760
window_height = 900

# Calcular las coordenadas para centrar la ventana
position_top = int(screen_height / 2 - window_height / 2)
position_left = int(screen_width / 2 - window_width / 2)

# Configurar la geometría para centrar la ventana
root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

root.resizable(TRUE, TRUE)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#D3D3D3")  # Gris claro

# Eliminar el Canvas y trabajar directamente con un Label para el logo
# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente, aumentando su tamaño un 30%
logo_width = 104  # Aumento del 30% respecto a 80 (80 * 1.3 = 104)
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Usar un Label en lugar del Canvas para colocar el logo en la parte superior izquierda
logo_label = Label(root, image=logo_photo, bg="#D3D3D3")  # Fondo transparente
logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")  # Ubicarlo en la parte superior izquierda


# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#2F2F2F", bg="#D3D3D3", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Crear la barra de progreso
progress_var = tk.IntVar()
style = ttk.Style()
style.configure("TProgressbar",
                thickness=50,  # Aumenta el grosor de la barra
                )
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=600, mode="determinate",style="TProgressbar")
progress_bar.grid(row=6, column=0, columnspan=3, pady=10)

# Función para convertir la fecha de entrada en el formato adecuado
def convert_date_format(date_str):
    try:
        return datetime.strptime(date_str, '%m/%d/%Y').strftime('%m_%d_%Y')
    except ValueError:
        return None

# Función para actualizar la barra de progreso
def update_progress_bar(current, total):
    progress_var.set((current / total) * 100)
    root.update_idletasks()

# Función para actualizar el log con un mensaje
def log_message(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + "\n")
    log_text.yview(tk.END)  # Desplazar el log al final
    log_text.config(state=tk.DISABLED)

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Invalid date.")
        return

    if not folder_path_laser or not folder_path_plasma:
        messagebox.showerror("Error", "Por favor verifica las rutas de las carpetas de láser y plasma en config.py.")
        log_message("Error: Folder paths not found.")
        return

    # Listar archivos PDF de las dos carpetas
    log_message("Listing PDF files from the folders...")
    turret_files, laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser, folder_path_plasma, date_filter)
   
    
    if not laser_plasma_files:
        log_message("No PDF files found in the laser and plasma folders.")
    
    total_files = len(laser_plasma_files)
    print(total_files)

    log_message(f"{total_files} PDF files found in the laser and plasma folders.")

    total_files += len(turret_files)
    log_message(f"{len(turret_files)} PDF files found in the turret folders.")
    print(total_files)
    # Procesar PDF de laser y plasma si hay archivos disponibles
    if laser_plasma_files:
        log_message("Processing PDF files from the laser and plasma folders...")
        for file in laser_plasma_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado1 = process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, date_filter, update_progress_bar)
        if df_consolidado1 is None:
            log_message("The DataFrame for laser and plasma was not generated.")
    else:
        df_consolidado1 = None

    # Actualizar progreso después de cada procesamiento
    progress_var.set((len(laser_plasma_files) / total_files) * 100)
    root.update_idletasks()

    # Procesar PDF de turret si hay archivos disponibles
    if turret_files:
        log_message("Processing PDF files from the turret folders...")
        for file in turret_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado2 = process_pdfs_turret_and_generate_excel(turret_files, date_filter, update_progress_bar)
        if df_consolidado2 is None:
            log_message("The DataFrame for turret was not generated.")
    else:
        df_consolidado2 = None

    # Procesar y guardar el archivo Excel si se generaron ambos DataFrames
    if df_consolidado1 is not None or df_consolidado2 is not None:
        # Concatenar los DataFrames si existen
        outputfinal = pd.concat([df_consolidado1, df_consolidado2], axis=0, ignore_index=True) if df_consolidado1 is not None and df_consolidado2 is not None else df_consolidado1 if df_consolidado1 is not None else df_consolidado2
        output_filename = f"Utilization_{date_filter}.xlsx"
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        log_message(f"Excel file saved at: {os.path.join(excel_output_path, output_filename)}")
    else:
        log_message("No Excel file was generated.")
        messagebox.showerror("Error", "No Excel file was generated.")


    log_message("The processing has been completed.")
    messagebox.showinfo("Process completed", "The PDF file processing has been completed.\n                     😊😊😊😊😊")

        # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Close the application", "Do you want to close the application?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación


# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Ajustar el área del log
log_frame = tk.Frame(root, bg="#FFFFFF", height=600)  # Fondo blanco y altura mayor
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew", padx=(10, 20))  # Añadir margen a ambos lados (izquierda y derecha)

# Barra de desplazamiento del log
log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

# Área de texto del log
log_text = Text(log_frame, height=30, width=120 ,wrap="word", font=("Helvetica", 8), fg="#000000", bg="#FFFFFF", bd=1, insertbackground="black", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)

# Configurar la barra de desplazamiento
log_scrollbar.config(command=log_text.yview)



root.mainloop()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\Dictionaries.py
code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK3612M2": "4.Citurret"
}

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuncionesFieldsTransform.py
####################################################################################################
# Funciones para transformar los campos de un DataFrame
####################################################################################################
# Las funciones a continuación se utilizan para transformar los campos de un DataFrame de Pandas.
# Estas funciones se pueden aplicar a una columna específica para modificar los valores de las celdas.
# Por ejemplo, se puede aplicar una función para ajustar los nombres de las máquinas o para obtener
# el calibre de una celda que contiene un 'nesting'.
#
def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"



def update_sheet_size(sheet_size, valid_sizes):
    # Eliminar los espacios en la cadena y convertir todo a minúsculas
    cleaned_size = sheet_size.replace(" ", "").lower()
    
    # Invertir la cadena limpia para verificar la variante invertida
    inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
    
    # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
    if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
        return sheet_size  # Conservar el valor original si hay coincidencia
    else:
        return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
    
    
# Función para obtener el material desde el nesting
def get_material(nesting, Material_dict):
    for material, values in Material_dict.items():
        if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
            return material
    return "Desconocido"


def adjust_type_nest(Type, type_dict):
    # Buscar si el tipo está en el diccionario y devolver el valor correspondiente
    return type_dict.get(Type, Type)


def adjust_machine(machine_name, machine_dict):
    # Buscar si el nombre de la máquina está en el diccionario y devolver el valor correspondiente
    return machine_dict.get(machine_name, machine_name)  # Si machine_name no está en el diccionario, retorna el valor original    



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionLaserPlasmaFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

def process_pdfs_laser_plasma_and_generate_excel(pdf_files, input_date, update_progress_func):
    all_utilization_lines = []
    all_headers = []  # Lista de headers
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las fechas Jake de cada archivo

    headers = pd.DataFrame()  # Inicializamos headers como un DataFrame vacío para evitar el error

    total_files = len(pdf_files)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados

    for pdf_path in pdf_files:
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        print(f"Categoría para {pdf_path}: {category}")
        nest = get_nest_from_path(pdf_path)
        date_process = input_date
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake = match.group(1)     

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                        
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake) 
                                utilization_data.append(category) 
                                #print(f"utilization_data: {utilization_data}")
                                 # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
                                #print(f"all_utilization_lines: {all_utilization_lines[:5]}")
        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Si se encontraron headers y lines de utilización
    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size','Nest','Date process','Date Jake','category'])
    else:
        print("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")

    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

        # Agregar la columna "Material"
        merged_df['Material'] = merged_df['NESTING'].apply(lambda x: get_material(x, Material_dict)) 

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        merged_df['MACHINE'] = merged_df['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time'
            })

        # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization', 'Cut Time'
        ]

        # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]
        merged_df['Utilization'] = pd.to_numeric(merged_df['Utilization'], errors='coerce')

        # Luego realizar la conversión a decimal (dividir entre 100)
        merged_df['Utilization'] = merged_df['Utilization'] / 100

        # Ajustar la columna 'Type' según las condiciones especificadas
        #merged_df['Type'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
        df_consolidado1 = pd.DataFrame(merged_df)

        

        '''# Exportar a Excel
        try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            print("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            print("Error", f"Error al guardar el archivo Excel: {e}")
        '''
    else:
        print("Error", "No se pueden combinar los DataFrames.")
    
    return df_consolidado1


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionListFiles.py

import os
from datetime import datetime
import re

def list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter):
    turret_files = []
    laser_plasma_files = []

    # Combina las dos rutas de directorios
    folder_paths = [folder_path_1, folder_path_2]

    # Ahora iteramos sobre los directorios y procesamos los archivos
 
    for folder_path in folder_paths:
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:                
                file_path = os.path.join(root, file_name)
                file_mod_time = os.path.getmtime(file_path)
                file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")
                # Verifica si el archivo es PDF y si la fecha coincide
                if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                    #print(f"Procesando archivo: {file_path}")    
                    # Clasifica según si el nombre contiene "turret"
                    #if "turret" in file_path.lower():
                    if re.search(r"turret", file_path, re.IGNORECASE): 
                        print("turret",file_path)
                        turret_files.append(file_path)
                    else:
                        laser_plasma_files.append(file_path)
                        print("laser",file_path)

    # Imprime las listas de archivos
    print("Archivos 'turret':", len(turret_files))
    print("Archivos 'laser_plasma':", len(laser_plasma_files))

    return turret_files, laser_plasma_files

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\Funtionscode.py






# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)

def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []       
    

  


# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

def list_pdf_files_in_folder_combined(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []    

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionsGetdata.py
import os
from datetime import datetime, timedelta
import re
from funtions.Dictionaries import type_dict

def get_today_date():
    today = datetime.today()
    # Restar un día
    yesterday = today - timedelta(days=5)
    # Retornar la fecha en el formato m_d_yyyy
    return yesterday.strftime("%m_%d_%Y")

def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

def convert_date_format_2(date_str):
    # Convertir la cadena en un objeto datetime usando el formato original
    date_obj = datetime.strptime(date_str, '%m_%d_%Y')
    # Formatear el objeto datetime al nuevo formato
    return date_obj.strftime('%m/%d/%Y')

'''
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"'''


def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Obtener la categoría encontrada
        category = match.group(0)
        
        # Buscar la categoría en el diccionario y devolver el valor correspondiente
        return type_dict.get(category, "MISC")  # Si no se encuentra, devuelve "MISC"
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"

    
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"    


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionsInterfaz.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *


# Crear la ventana principal
root = ThemedTk(theme="arc")
root.title("Interfaz de Selección de Carpeta")

# Crear y colocar el label para mostrar la ruta de la carpeta seleccionada
folder_path_label = Label(root, text="No se ha seleccionado ninguna carpeta")
folder_path_label.pack(pady=10)

# Crear y colocar el widget de texto para el log
log_text = Text(root, state=tk.DISABLED, height=10, width=50)
log_text.pack(pady=10)

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\funtions\FuntionTurretFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

folder_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF

def process_pdfs_turret_and_generate_excel(fullpath, input_date, update_progress_func):
    consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF
    
    total_files = len(fullpath)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados
    
    for pdf_path in fullpath:
        print(f"Procesando el archivo: {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:    
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if table[0] == ['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        i = 3
                        while i < len(table):
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')
                        #print(df_final.columns)

                        df_final['Date process'] = input_date
                        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
                        df_final['Date Jake'] = match.group(1)                        
                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
                        #df_final['Sheet Size'] = df_final['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
                        df_final['MACHINE'] = df_final['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))
                        df_final['Utilization'] = ''
                        df_final['Category'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
                        df_final['Material Name'] = df_final['SCHEDULE'].apply(lambda x: get_material(x, Material_dict))
                        df_final['Sheet Name'] = df_final['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        df_final['Date Jake'] = df_final['Date Jake'].apply(lambda x: convert_date_format_2(x))
                        df_final['Date process'] = df_final['Date process'].apply(lambda x: convert_date_format_2(x))

                        df_final.rename(columns={
                            'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                            'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                            'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                            'Cut Time': 'Cut Time'
                        }, inplace=True)

                        # Reordenar columnas
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)

        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Devolver el DataFrame consolidado
    df_consolidado2 = pd.DataFrame(consolidated_df) 
    return df_consolidado2


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250117_T1129\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        '''try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")'''


############################################################################################################
# Función para listar archivos PDF turret en una carpeta y sus subcarpetas que contienen la fecha de modificación
############################################################################################################       
# Listar archivos PDF en la carpeta y subcarpetas con la fecha de modificación deseada
def list_pdf_files_turret_in_folder(folder_path, date_filter):
    pdf_files_turret = []
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            file_mod_time = os.path.getmtime(file_path)
            file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

            if date_filter == file_mod_date and file_name.lower().endswith(".pdf") and 'turret'  in file_path.lower():
                pdf_files_turret.append(file_path)
                category = get_category_from_path(file_path)
                nest = get_nest_from_path(file_path)
                date_process = file_mod_date
                


    return pdf_files_turret

# Función para procesar los PDFs y generar el archivo Excel
def process_pdfs_turret_and_generate_excel(pdf_files_turret, input_date):
     all_dfs = []  # Lista para almacenar todos los DataFrames
     for pdf_path in pdf_files_turret:
        category = get_category_from_path(pdf_path)
        #print(category)
        nest = get_nest_from_path(pdf_path)
        #print(nest)
        date_process = input_date
        #print(date_process)
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', nest)
        date_jake = match.group(1) if match else "MISC"
        date_process=date_process.replace("_","/")


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if   table[0] ==['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]  :
                                            
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')#.drop('key', axis=1)

                        df_final = df_final[
                            ['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME', 'Material Name', 'Sheet Code', 'Sheet Size', 
                            'Sheet Count', 'Sheet Name', 'Stack Qty', 'Cut Time', 'Finish']
                        ]
                        def adjust_machine(machine_name):
                            adjustments = {
                                "Amada_ENSIS_4020AJ": "1. Laser", "Messer_170Amp_Plasm": "2.Plasma", 
                                "Amada_Vipros_358K": "3.Turret", "Amada_EMK316M2": "4.Citurret"
                            }
                            return adjustments.get(machine_name, machine_name)
                        
                        df_final['MACHINE'] = df_final['MACHINE'].apply(adjust_machine)

                        def get_gauge_from_nesting(SCHEDULE, code_to_gauge):
                            for code, gauge in code_to_gauge.items():
                                if code in SCHEDULE:
                                    return gauge
                            return "Desconocido"

                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

                        def update_sheet_size(sheet_size):
                            valid_sizes = [
                                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
                            ]
                            cleaned_size = sheet_size.replace(" ", "").lower()
                            inverted_size = 'x'.join(cleaned_size.split('x')[::-1])  # Reverse size
                            return sheet_size if cleaned_size in valid_sizes or inverted_size in valid_sizes else "REMNANT"
                        

                        

                        df_final['Sheet Sizex'] = df_final['Sheet Size'].apply(update_sheet_size)
                        df_final['Date process'] = date_process
                        df_final['Date Jake'] = date_jake
                        df_final['Utilization']=''
                        def adjust_type_nest(Type):
                            adjustments = {
                                "ENCL": "2.Encl", "ENGR": "6.Engr", "MEC": "3.Mec", "PARTS ORDER": "4.Parts Order", "REWORK": "5.Rework",
                                "SIL": "4.Sil", "TANK": "1.Tank"
                            }
                            return adjustments.get(Type, Type)

                        df_final['Category'] = adjust_type_nest(category)

                        df_final.rename(columns={
                        'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                        'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                        'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                        'Cut Time': 'Cut Time'
                    }, inplace=True)
                        
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]

                        def get_material(nesting):
                            for material, values in category_dict.items():
                                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                                    return material
                            return "Desconocido"          

                        df_final['Material'] = df_final['Nesting'].apply(get_material)

                        df_final['Program'] = df_final['Program'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        
                        all_dfs.append(df_final)

                        # Consolidar todos los DataFrames en uno solo
        consolidated_df = pd.concat(all_dfs, ignore_index=True)

        # Guardar el DataFrame consolidado en un archivo CSV
        output_file = excel_output_path + 'Utilization_consolidated_' + str(input_date) + '.csv'
        consolidated_df.to_csv(output_file, index=False)
        print("Todos los archivos procesados y consolidados exitosamente.")
        print(consolidated_df) 



# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    pdf_files_turret = list_pdf_files_turret_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)
    process_pdfs_turret_and_generate_excel(pdf_files_turret, date_filter)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

    # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Cerrar la aplicación", "¿Deseas cerrar la aplicación?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)



# Esta es la línea que mantiene la interfaz de Tkinter en ejecución hasta que el usuario decida cerrarla.
root.mainloop() 

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\config.py
#excel_output_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data"
logo_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting_V2.0_20250117_T1129\image\Logo.png'
folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_plasma = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Issue_20250127\01_24_2025_NEST 01 ENCL\PLASMA'
excel_output_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data'

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\main.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *



# Configuración de rutas
#folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
#folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
#excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'


# Función para obtener la fecha actual en formato m_d_yyyy

def main():
    inicio_proceso = datetime.now()
    input_date = get_today_date()
    _ ,laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date)
    df_consolidado1=process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, input_date)
    df_consolidado1['Cut Time'] = ''

    turret_files, _ = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date) 
    df_consolidado2 =process_pdfs_turret_and_generate_excel(turret_files, input_date)
    df_consolidado2['Cut Time'] = ''
    
    if df_consolidado1 is not None and df_consolidado2 is not None:        
        output_filename = f"Utilization_{input_date}.xlsx"
        outputfinal = pd.concat([df_consolidado1,df_consolidado2], axis=0)
        
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is not None and df_consolidado2 is None: 
        outputfinal=pd.DataFrame(df_consolidado1)
        
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is None and df_consolidado2 is not None: 
        outputfinal=pd.DataFrame(df_consolidado2)
        
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}") 
    else:
        print("No se generó ningún archivo Excel.")

    fin_proceso = datetime.now()
    print(f"Tiempo de ejecución: {fin_proceso - inicio_proceso}")
if __name__ == "__main__":
    main()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\todo.py

import os

# Ruta de la carpeta donde se encuentran los archivos
folder_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization'

# Ruta del archivo de salida donde se guardará el código combinado
output_file = 'codigo_combinado.py'

# Función para recorrer directorios y subdirectorios
def agregar_codigo(directorio, archivo_salida):
    for root, dirs, files in os.walk(directorio):
        for filename in files:
            file_path = os.path.join(root, filename)
            
            # Verifica si es un archivo Python (puedes cambiar la extensión si es otro tipo de archivo)
            if filename.endswith('.py'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as infile:
                        # Escribe el nombre del archivo en el archivo de salida
                        archivo_salida.write(f"# Contenido de {file_path}\n")
                        # Escribe el contenido del archivo en el archivo de salida
                        archivo_salida.write(infile.read())
                        archivo_salida.write("\n\n")
                except UnicodeDecodeError:
                    print(f"Error al leer el archivo: {file_path}. Se omitirá este archivo.")

# Abre el archivo de salida en modo escritura
with open(output_file, 'w', encoding='utf-8') as outfile:
    agregar_codigo(folder_path, outfile)

print(f"Todo el código se ha combinado en {output_file}")



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\utilizacion.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path, logo_path  # Aquí cargamos las variables de configuración

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")


# Obtener las dimensiones de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Definir el tamaño de la ventana
window_width = 760
window_height = 900

# Calcular las coordenadas para centrar la ventana
position_top = int(screen_height / 2 - window_height / 2)
position_left = int(screen_width / 2 - window_width / 2)

# Configurar la geometría para centrar la ventana
root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

root.resizable(TRUE, TRUE)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#D3D3D3")  # Gris claro

# Eliminar el Canvas y trabajar directamente con un Label para el logo
# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente, aumentando su tamaño un 30%
logo_width = 104  # Aumento del 30% respecto a 80 (80 * 1.3 = 104)
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Usar un Label en lugar del Canvas para colocar el logo en la parte superior izquierda
logo_label = Label(root, image=logo_photo, bg="#D3D3D3")  # Fondo transparente
logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")  # Ubicarlo en la parte superior izquierda


# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#2F2F2F", bg="#D3D3D3", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Crear la barra de progreso
progress_var = tk.IntVar()
style = ttk.Style()
style.configure("TProgressbar",
                thickness=50,  # Aumenta el grosor de la barra
                )
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=600, mode="determinate",style="TProgressbar")
progress_bar.grid(row=6, column=0, columnspan=3, pady=10)

# Función para convertir la fecha de entrada en el formato adecuado
def convert_date_format(date_str):
    try:
        return datetime.strptime(date_str, '%m/%d/%Y').strftime('%m_%d_%Y')
    except ValueError:
        return None

# Función para actualizar la barra de progreso
def update_progress_bar(current, total):
    progress_var.set((current / total) * 100)
    root.update_idletasks()

# Función para actualizar el log con un mensaje
def log_message(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + "\n")
    log_text.yview(tk.END)  # Desplazar el log al final
    log_text.config(state=tk.DISABLED)

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Invalid date.")
        return

    if not folder_path_laser or not folder_path_plasma:
        messagebox.showerror("Error", "Por favor verifica las rutas de las carpetas de láser y plasma en config.py.")
        log_message("Error: Folder paths not found.")
        return

    # Listar archivos PDF de las dos carpetas
    log_message("Listing PDF files from the folders...")
    turret_files, laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser, folder_path_plasma, date_filter)
   
    
    if not laser_plasma_files:
        log_message("No PDF files found in the laser and plasma folders.")
    
    total_files = len(laser_plasma_files)
    print(total_files)

    log_message(f"{total_files} PDF files found in the laser and plasma folders.")

    total_files += len(turret_files)
    log_message(f"{len(turret_files)} PDF files found in the turret folders.")
    print(total_files)
    # Procesar PDF de laser y plasma si hay archivos disponibles
    if laser_plasma_files:
        log_message("Processing PDF files from the laser and plasma folders...")
        for file in laser_plasma_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado1 = process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, date_filter, update_progress_bar)
        if df_consolidado1 is None:
            log_message("The DataFrame for laser and plasma was not generated.")
    else:
        df_consolidado1 = None

    # Actualizar progreso después de cada procesamiento
    progress_var.set((len(laser_plasma_files) / total_files) * 100)
    root.update_idletasks()

    # Procesar PDF de turret si hay archivos disponibles
    if turret_files:
        log_message("Processing PDF files from the turret folders...")
        for file in turret_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado2 = process_pdfs_turret_and_generate_excel(turret_files, date_filter, update_progress_bar)
        if df_consolidado2 is None:
            log_message("The DataFrame for turret was not generated.")
    else:
        df_consolidado2 = None

    # Procesar y guardar el archivo Excel si se generaron ambos DataFrames
    if df_consolidado1 is not None or df_consolidado2 is not None:
        # Concatenar los DataFrames si existen
        outputfinal = pd.concat([df_consolidado1, df_consolidado2], axis=0, ignore_index=True) if df_consolidado1 is not None and df_consolidado2 is not None else df_consolidado1 if df_consolidado1 is not None else df_consolidado2
        outputfinal['Cut Time'] = None
        output_filename = f"Utilization_{date_filter}.xlsx"
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        log_message(f"Excel file saved at: {os.path.join(excel_output_path, output_filename)}")
    else:
        log_message("No Excel file was generated.")
        messagebox.showerror("Error", "No Excel file was generated.")


    log_message("The processing has been completed.")
    messagebox.showinfo("Process completed", "The PDF file processing has been completed.\n                     😊😊😊😊😊")

        # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Close the application", "Do you want to close the application?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación


# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Ajustar el área del log
log_frame = tk.Frame(root, bg="#FFFFFF", height=600)  # Fondo blanco y altura mayor
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew", padx=(10, 20))  # Añadir margen a ambos lados (izquierda y derecha)

# Barra de desplazamiento del log
log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

# Área de texto del log
log_text = Text(log_frame, height=30, width=120 ,wrap="word", font=("Helvetica", 8), fg="#000000", bg="#FFFFFF", bd=1, insertbackground="black", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)

# Configurar la barra de desplazamiento
log_scrollbar.config(command=log_text.yview)



root.mainloop()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\Dictionaries.py
code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.0": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "GALV-0.14": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "STEEL-0.25": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.14": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK3612M2": "4.Citurret"
}

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\FuncionesFieldsTransform.py
####################################################################################################
# Funciones para transformar los campos de un DataFrame
####################################################################################################
# Las funciones a continuación se utilizan para transformar los campos de un DataFrame de Pandas.
# Estas funciones se pueden aplicar a una columna específica para modificar los valores de las celdas.
# Por ejemplo, se puede aplicar una función para ajustar los nombres de las máquinas o para obtener
# el calibre de una celda que contiene un 'nesting'.
#
def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"



def update_sheet_size(sheet_size, valid_sizes):
    # Eliminar los espacios en la cadena y convertir todo a minúsculas
    cleaned_size = sheet_size.replace(" ", "").lower()
    
    # Invertir la cadena limpia para verificar la variante invertida
    inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
    
    # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
    if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
        return sheet_size  # Conservar el valor original si hay coincidencia
    else:
        return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
    
    
# Función para obtener el material desde el nesting
def get_material(nesting, Material_dict):
    for material, values in Material_dict.items():
        if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
            return material
    return "Desconocido"


def adjust_type_nest(Type, type_dict):
    # Buscar si el tipo está en el diccionario y devolver el valor correspondiente
    return type_dict.get(Type, Type)


def adjust_machine(machine_name, machine_dict):
    # Buscar si el nombre de la máquina está en el diccionario y devolver el valor correspondiente
    return machine_dict.get(machine_name, machine_name)  # Si machine_name no está en el diccionario, retorna el valor original    



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\FuntionLaserPlasmaFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

def process_pdfs_laser_plasma_and_generate_excel(pdf_files, input_date, update_progress_func):
    all_utilization_lines = []
    all_headers = []  # Lista de headers
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las fechas Jake de cada archivo

    headers = pd.DataFrame()  # Inicializamos headers como un DataFrame vacío para evitar el error

    total_files = len(pdf_files)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados

    for pdf_path in pdf_files:
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process = input_date
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake = match.group(1)     

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[9]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                    elif table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)  

                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[1] = utilization_data[1].replace("Sheet Name: ", "")
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake) 
                                utilization_data.append(category) 
                                all_utilization_lines.append(utilization_data)

        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size','Nest','Date process','Date Jake','category'])
    else:
        print("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")

    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        utilization_df['Sheet Name'] = utilization_df['Sheet Name'].astype(str)
        
        # Crear una lista para almacenar los resultados del join
        merged_rows = []

        # Iterar por cada fila en headers
        for _, header_row in headers.iterrows():
            # Extraer la parte después del primer guion en 'Sheet Name' de utilization_df
            utilization_df['Sheet Name Part'] = utilization_df['Sheet Name'].apply(lambda x: x.split('-')[1] if '-' in x else '')

            # Filtrar filas en utilization_df donde 'NESTING' y la parte extraída de 'Sheet Name' estén contenidos en 'SCHEDULE'
            matching_rows = utilization_df[
                utilization_df['NESTING'].apply(lambda x: x in header_row['SCHEDULE']) &
                utilization_df['Sheet Name Part'].apply(lambda x: x in header_row['SCHEDULE'])
            ]
            
            # Si hay coincidencias, combinar las filas
            for _, util_row in matching_rows.iterrows():
                merged_rows.append({**header_row.to_dict(), **util_row.to_dict()})

        # Convertir la lista de resultados a un DataFrame
        merged_df = pd.DataFrame(merged_rows)

        # Eliminar duplicados
        merged_df = merged_df.drop_duplicates()

        # Eliminar la columna 'SCHEDULE' si es necesario
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

        # Agregar la columna "Material"
        merged_df['Material'] = merged_df['NESTING'].apply(lambda x: get_material(x, Material_dict))

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        merged_df['MACHINE'] = merged_df['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))

        merged_df['Gauge'] = merged_df['Sheet Name'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
            'MACHINE': 'Machine',
            'Nest': 'Nesting',
            'Sheet Size': 'Size',
            'category': 'Type',
            'Material': 'Material',
            'Gauge': 'Gage',
            'Stack Qty': '# Sheets',
            'Program': 'Program',
            'Utililization': 'Utilization',
            'Date process': 'Date-Nesting',
            'Date Jake': 'Date Jake',
            'TOTAL CUT TIME': 'Cut Time'
        })

        # Reordenar las columnas
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization'
        ]

        merged_df = merged_df[new_order]
        merged_df['Utilization'] = pd.to_numeric(merged_df['Utilization'], errors='coerce')

        # Realizar la conversión a decimal (dividir entre 100)
        merged_df['Utilization'] = merged_df['Utilization'] / 100

        # Eliminar duplicados de nuevo, por si hubo alguna duplicidad después de los ajustes
        merged_df = merged_df.drop_duplicates()

        df_consolidado1 = pd.DataFrame(merged_df)
        print(df_consolidado1)

    else:
        print("Error", "No se pueden combinar los DataFrames.")
    
    return df_consolidado1


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\FuntionListFiles.py

import os
from datetime import datetime
import re

def list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter):
    turret_files = []
    laser_plasma_files = []

    # Combina las dos rutas de directorios
    folder_paths = [folder_path_1, folder_path_2]

    # Ahora iteramos sobre los directorios y procesamos los archivos
 
    for folder_path in folder_paths:
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:                
                file_path = os.path.join(root, file_name)
                file_mod_time = os.path.getmtime(file_path)
                file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")
                # Verifica si el archivo es PDF y si la fecha coincide
                if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                    #print(f"Procesando archivo: {file_path}")    
                    # Clasifica según si el nombre contiene "turret"
                    #if "turret" in file_path.lower():
                    if re.search(r"turret", file_path, re.IGNORECASE): 
                        ##print("turret",file_path)
                        turret_files.append(file_path)
                    else:
                        laser_plasma_files.append(file_path)
                        ###print("laser",file_path)

    # Imprime las listas de archivos
    ##print("Archivos 'turret':", len(turret_files))
    ##print("Archivos 'laser_plasma':", len(laser_plasma_files))

    return turret_files, laser_plasma_files

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\Funtionscode.py






# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)

def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []       
    

  


# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

def list_pdf_files_in_folder_combined(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []    

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\FuntionsGetdata.py
import os
from datetime import datetime, timedelta
import re
from funtions.Dictionaries import type_dict

def get_today_date():
    today = datetime.today()
    # Restar un día
    yesterday = today - timedelta(days=5)
    # Retornar la fecha en el formato m_d_yyyy
    return yesterday.strftime("%m_%d_%Y")

def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

def convert_date_format_2(date_str):
    # Convertir la cadena en un objeto datetime usando el formato original
    date_obj = datetime.strptime(date_str, '%m_%d_%Y')
    # Formatear el objeto datetime al nuevo formato
    return date_obj.strftime('%m/%d/%Y')

'''
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"'''


def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Obtener la categoría encontrada
        category = match.group(0)
        
        # Buscar la categoría en el diccionario y devolver el valor correspondiente
        return type_dict.get(category, "MISC")  # Si no se encuentra, devuelve "MISC"
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"

    
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"    


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\FuntionsInterfaz.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *


# Crear la ventana principal
root = ThemedTk(theme="arc")
root.title("Interfaz de Selección de Carpeta")

# Crear y colocar el label para mostrar la ruta de la carpeta seleccionada
folder_path_label = Label(root, text="No se ha seleccionado ninguna carpeta")
folder_path_label.pack(pady=10)

# Crear y colocar el widget de texto para el log
log_text = Text(root, state=tk.DISABLED, height=10, width=50)
log_text.pack(pady=10)

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\funtions\FuntionTurretFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

##folder_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
##excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF

def process_pdfs_turret_and_generate_excel(fullpath, input_date, update_progress_func):
    consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF
    
    total_files = len(fullpath)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados
    
    for pdf_path in fullpath:
        print(f"Procesando el archivo: {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:    
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if table[0] == ['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        i = 3
                        while i < len(table):
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')
                        #print(df_final.columns)

                        df_final['Date process'] = input_date
                        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
                        df_final['Date Jake'] = match.group(1)                        
                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
                        #df_final['Sheet Size'] = df_final['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
                        df_final['MACHINE'] = df_final['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))
                        df_final['Utilization'] = ''
                        df_final['Category'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
                        df_final['Material Name'] = df_final['SCHEDULE'].apply(lambda x: get_material(x, Material_dict))
                        df_final['Sheet Name'] = df_final['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        df_final['Date Jake'] = df_final['Date Jake'].apply(lambda x: convert_date_format_2(x))
                        df_final['Date process'] = df_final['Date process'].apply(lambda x: convert_date_format_2(x))

                        df_final.rename(columns={
                            'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                            'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                            'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                            'Cut Time': 'Cut Time'
                        }, inplace=True)

                        # Reordenar columnas
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization'
                        ]
                        df_final = df_final[new_order]
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)

        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Devolver el DataFrame consolidado
    df_consolidado2 = pd.DataFrame(consolidated_df) 
    return df_consolidado2


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        '''try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")'''


############################################################################################################
# Función para listar archivos PDF turret en una carpeta y sus subcarpetas que contienen la fecha de modificación
############################################################################################################       
# Listar archivos PDF en la carpeta y subcarpetas con la fecha de modificación deseada
def list_pdf_files_turret_in_folder(folder_path, date_filter):
    pdf_files_turret = []
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            file_mod_time = os.path.getmtime(file_path)
            file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

            if date_filter == file_mod_date and file_name.lower().endswith(".pdf") and 'turret'  in file_path.lower():
                pdf_files_turret.append(file_path)
                category = get_category_from_path(file_path)
                nest = get_nest_from_path(file_path)
                date_process = file_mod_date
                


    return pdf_files_turret

# Función para procesar los PDFs y generar el archivo Excel
def process_pdfs_turret_and_generate_excel(pdf_files_turret, input_date):
     all_dfs = []  # Lista para almacenar todos los DataFrames
     for pdf_path in pdf_files_turret:
        category = get_category_from_path(pdf_path)
        #print(category)
        nest = get_nest_from_path(pdf_path)
        #print(nest)
        date_process = input_date
        #print(date_process)
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', nest)
        date_jake = match.group(1) if match else "MISC"
        date_process=date_process.replace("_","/")


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if   table[0] ==['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]  :
                                            
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')#.drop('key', axis=1)

                        df_final = df_final[
                            ['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME', 'Material Name', 'Sheet Code', 'Sheet Size', 
                            'Sheet Count', 'Sheet Name', 'Stack Qty', 'Cut Time', 'Finish']
                        ]
                        def adjust_machine(machine_name):
                            adjustments = {
                                "Amada_ENSIS_4020AJ": "1. Laser", "Messer_170Amp_Plasm": "2.Plasma", 
                                "Amada_Vipros_358K": "3.Turret", "Amada_EMK316M2": "4.Citurret"
                            }
                            return adjustments.get(machine_name, machine_name)
                        
                        df_final['MACHINE'] = df_final['MACHINE'].apply(adjust_machine)

                        def get_gauge_from_nesting(SCHEDULE, code_to_gauge):
                            for code, gauge in code_to_gauge.items():
                                if code in SCHEDULE:
                                    return gauge
                            return "Desconocido"

                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

                        def update_sheet_size(sheet_size):
                            valid_sizes = [
                                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
                            ]
                            cleaned_size = sheet_size.replace(" ", "").lower()
                            inverted_size = 'x'.join(cleaned_size.split('x')[::-1])  # Reverse size
                            return sheet_size if cleaned_size in valid_sizes or inverted_size in valid_sizes else "REMNANT"
                        

                        

                        df_final['Sheet Sizex'] = df_final['Sheet Size'].apply(update_sheet_size)
                        df_final['Date process'] = date_process
                        df_final['Date Jake'] = date_jake
                        df_final['Utilization']=''
                        def adjust_type_nest(Type):
                            adjustments = {
                                "ENCL": "2.Encl", "ENGR": "6.Engr", "MEC": "3.Mec", "PARTS ORDER": "4.Parts Order", "REWORK": "5.Rework",
                                "SIL": "4.Sil", "TANK": "1.Tank"
                            }
                            return adjustments.get(Type, Type)

                        df_final['Category'] = adjust_type_nest(category)

                        df_final.rename(columns={
                        'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                        'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                        'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                        'Cut Time': 'Cut Time'
                    }, inplace=True)
                        
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]

                        def get_material(nesting):
                            for material, values in category_dict.items():
                                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                                    return material
                            return "Desconocido"          

                        df_final['Material'] = df_final['Nesting'].apply(get_material)

                        df_final['Program'] = df_final['Program'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        
                        all_dfs.append(df_final)

                        # Consolidar todos los DataFrames en uno solo
        consolidated_df = pd.concat(all_dfs, ignore_index=True)

        # Guardar el DataFrame consolidado en un archivo CSV
        output_file = excel_output_path + 'Utilization_consolidated_' + str(input_date) + '.csv'
        consolidated_df.to_csv(output_file, index=False)
        print("Todos los archivos procesados y consolidados exitosamente.")
        print(consolidated_df) 



# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    pdf_files_turret = list_pdf_files_turret_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)
    process_pdfs_turret_and_generate_excel(pdf_files_turret, date_filter)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

    # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Cerrar la aplicación", "¿Deseas cerrar la aplicación?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)



# Esta es la línea que mantiene la interfaz de Tkinter en ejecución hasta que el usuario decida cerrarla.
root.mainloop() 

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718\Old_version&testing\test.py

import os
from datetime import datetime
from funtions.FuntionListFiles import list_all_pdf_files_in_folders

folder_path_1 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_2 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
date_filter = "01_10_2025"

def update_progress_func(processed_files, total_files):
    progress = processed_files / total_files * 100
    print(f"Progreso: {progress:.2f}%")


list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter)

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\config.py
#excel_output_path = r"C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data"
logo_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting_V2.0_20250117_T1129\image\Logo.png'
folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_plasma = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Issue_20250127\01_24_2025_NEST 01 ENCL\PLASMA'
excel_output_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\Nesting Dept. Docs\Utilization_Daily_Data'

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\main.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *



# Configuración de rutas
#folder_path_laser = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
#folder_path_plasma = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
#excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'


# Función para obtener la fecha actual en formato m_d_yyyy

def main():
    inicio_proceso = datetime.now()
    input_date = get_today_date()
    _ ,laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date)
    df_consolidado1=process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, input_date)
    df_consolidado1['Cut Time'] = ''

    turret_files, _ = list_all_pdf_files_in_folders(folder_path_laser,folder_path_plasma, input_date) 
    df_consolidado2 =process_pdfs_turret_and_generate_excel(turret_files, input_date)
    df_consolidado2['Cut Time'] = ''
    
    if df_consolidado1 is not None and df_consolidado2 is not None:        
        output_filename = f"Utilization_{input_date}.xlsx"
        outputfinal = pd.concat([df_consolidado1,df_consolidado2], axis=0)
        
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is not None and df_consolidado2 is None: 
        outputfinal=pd.DataFrame(df_consolidado1)
        
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}")
    elif df_consolidado1 is None and df_consolidado2 is not None: 
        outputfinal=pd.DataFrame(df_consolidado2)
        
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        print(f"Archivo Excel guardado en: {os.path.join(excel_output_path, output_filename)}") 
    else:
        print("No se generó ningún archivo Excel.")

    fin_proceso = datetime.now()
    print(f"Tiempo de ejecución: {fin_proceso - inicio_proceso}")
if __name__ == "__main__":
    main()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\utilizacion.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import folder_path_laser, folder_path_plasma, excel_output_path, logo_path  # Aquí cargamos las variables de configuración

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")


# Obtener las dimensiones de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Definir el tamaño de la ventana
window_width = 760
window_height = 900

# Calcular las coordenadas para centrar la ventana
position_top = int(screen_height / 2 - window_height / 2)
position_left = int(screen_width / 2 - window_width / 2)

# Configurar la geometría para centrar la ventana
root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

root.resizable(TRUE, TRUE)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#D3D3D3")  # Gris claro

# Eliminar el Canvas y trabajar directamente con un Label para el logo
# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open(logo_path)  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente, aumentando su tamaño un 30%
logo_width = 104  # Aumento del 30% respecto a 80 (80 * 1.3 = 104)
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Usar un Label en lugar del Canvas para colocar el logo en la parte superior izquierda
logo_label = Label(root, image=logo_photo, bg="#D3D3D3")  # Fondo transparente
logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")  # Ubicarlo en la parte superior izquierda


# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#2F2F2F", bg="#D3D3D3", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Crear la barra de progreso
progress_var = tk.IntVar()
style = ttk.Style()
style.configure("TProgressbar",
                thickness=50,  # Aumenta el grosor de la barra
                )
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=600, mode="determinate",style="TProgressbar")
progress_bar.grid(row=6, column=0, columnspan=3, pady=10)

# Función para convertir la fecha de entrada en el formato adecuado
def convert_date_format(date_str):
    try:
        return datetime.strptime(date_str, '%m/%d/%Y').strftime('%m_%d_%Y')
    except ValueError:
        return None

# Función para actualizar la barra de progreso
def update_progress_bar(current, total):
    progress_var.set((current / total) * 100)
    root.update_idletasks()

# Función para actualizar el log con un mensaje
def log_message(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + "\n")
    log_text.yview(tk.END)  # Desplazar el log al final
    log_text.config(state=tk.DISABLED)

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Starting the process...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Invalid date.")
        return

    if not folder_path_laser or not folder_path_plasma:
        messagebox.showerror("Error", "Por favor verifica las rutas de las carpetas de láser y plasma en config.py.")
        log_message("Error: Folder paths not found.")
        return

    # Listar archivos PDF de las dos carpetas
    log_message("Listing PDF files from the folders...")
    turret_files, laser_plasma_files = list_all_pdf_files_in_folders(folder_path_laser, folder_path_plasma, date_filter)
   
    
    if not laser_plasma_files:
        log_message("No PDF files found in the laser and plasma folders.")
    
    total_files = len(laser_plasma_files)
    print(total_files)

    log_message(f"{total_files} PDF files found in the laser and plasma folders.")

    total_files += len(turret_files)
    log_message(f"{len(turret_files)} PDF files found in the turret folders.")
    print(total_files)
    # Procesar PDF de laser y plasma si hay archivos disponibles
    if laser_plasma_files:
        log_message("Processing PDF files from the laser and plasma folders...")
        for file in laser_plasma_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado1 = process_pdfs_laser_plasma_and_generate_excel(laser_plasma_files, date_filter, update_progress_bar)
        if df_consolidado1 is None:
            log_message("The DataFrame for laser and plasma was not generated.")
    else:
        df_consolidado1 = None

    # Actualizar progreso después de cada procesamiento
    progress_var.set((len(laser_plasma_files) / total_files) * 100)
    root.update_idletasks()

    # Procesar PDF de turret si hay archivos disponibles
    if turret_files:
        log_message("Processing PDF files from the turret folders...")
        for file in turret_files:  # Iterar sobre los archivos
            log_message(f"Processing file: {os.path.basename(file)}")  # Log el nombre del archivo
        df_consolidado2 = process_pdfs_turret_and_generate_excel(turret_files, date_filter, update_progress_bar)
        if df_consolidado2 is None:
            log_message("The DataFrame for turret was not generated.")
    else:
        df_consolidado2 = None

    # Procesar y guardar el archivo Excel si se generaron ambos DataFrames
    if df_consolidado1 is not None or df_consolidado2 is not None:
        # Concatenar los DataFrames si existen
        outputfinal = pd.concat([df_consolidado1, df_consolidado2], axis=0, ignore_index=True) if df_consolidado1 is not None and df_consolidado2 is not None else df_consolidado1 if df_consolidado1 is not None else df_consolidado2
        outputfinal['Cut Time'] = None
        output_filename = f"Utilization_{date_filter}.xlsx"
        outputfinal.to_excel(os.path.join(excel_output_path, output_filename), index=False)
        log_message(f"Excel file saved at: {os.path.join(excel_output_path, output_filename)}")
    else:
        log_message("No Excel file was generated.")
        messagebox.showerror("Error", "No Excel file was generated.")


    log_message("The processing has been completed.")
    messagebox.showinfo("Process completed", "The PDF file processing has been completed.\n                     😊😊😊😊😊")

        # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Close the application", "Do you want to close the application?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación


# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Ajustar el área del log
log_frame = tk.Frame(root, bg="#FFFFFF", height=600)  # Fondo blanco y altura mayor
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew", padx=(10, 20))  # Añadir margen a ambos lados (izquierda y derecha)

# Barra de desplazamiento del log
log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

# Área de texto del log
log_text = Text(log_frame, height=30, width=120 ,wrap="word", font=("Helvetica", 8), fg="#000000", bg="#FFFFFF", bd=1, insertbackground="black", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)

# Configurar la barra de desplazamiento
log_scrollbar.config(command=log_text.yview)



root.mainloop()


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\Dictionaries.py
code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.0": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "GALV-0.14": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "STEEL-0.25": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.14": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}

categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

Material_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

valid_sizes = ["60x86","60x120","60x144","72x120","72x144","48x144","60x145","72x145",
               "72x187","60x157","48x90","48x157","48x120","48x158","48x96","23.94x155",
               "23.94x190","23.94x200","33x121","33x145","24.83x133","24.83x170","24.83x190",
               "24.83x205","24.96x145","24.96x200","28x120","28x157","44.24x128","44.24x151",
               "24.45x145","24.45x205","27.92x120","27.92x157","44.21x120","44.21x157","60x175",
               "60x187","60x164","60x172","60x163","45x174","60x162","23.94x205","23.94x191.75",
               "24.83x206.75","24.83x149","33x149","60x146","170X24.833","190X23.95","157X27.93",
               "145X33.015","205X24.8339","151X44.24","121X33.012","133X24.8339","90X48.1",
               "175X50","144X50"]

type_dict = {
    "ENCL": "2.Encl",
    "ENGR": "6.Engr",
    "MEC": "3.Mec",
    "PARTS ORDER": "4.Parts Order",
    "REWORK": "5.Rework",
    "SIL": "4.Sil",
    "TANK": "1.Tank"
}

machine_dict = {
    "Amada_ENSIS_4020AJ": "1. Laser",
    "Messer_170Amp_Plasm": "2.Plasma",
    "Amada_Vipros_358K": "3.Turret",
    "Amada_EMK3612M2": "4.Citurret"
}

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\FuncionesFieldsTransform.py
####################################################################################################
# Funciones para transformar los campos de un DataFrame
####################################################################################################
# Las funciones a continuación se utilizan para transformar los campos de un DataFrame de Pandas.
# Estas funciones se pueden aplicar a una columna específica para modificar los valores de las celdas.
# Por ejemplo, se puede aplicar una función para ajustar los nombres de las máquinas o para obtener
# el calibre de una celda que contiene un 'nesting'.
#
def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"



def update_sheet_size(sheet_size, valid_sizes):
    # Eliminar los espacios en la cadena y convertir todo a minúsculas
    cleaned_size = sheet_size.replace(" ", "").lower()
    
    # Invertir la cadena limpia para verificar la variante invertida
    inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
    
    # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
    if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
        return sheet_size  # Conservar el valor original si hay coincidencia
    else:
        return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
    
    
# Función para obtener el material desde el nesting
def get_material(nesting, Material_dict):
    for material, values in Material_dict.items():
        if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
            return material
    return "Desconocido"


def adjust_type_nest(Type, type_dict):
    # Buscar si el tipo está en el diccionario y devolver el valor correspondiente
    return type_dict.get(Type, Type)


def adjust_machine(machine_name, machine_dict):
    # Buscar si el nombre de la máquina está en el diccionario y devolver el valor correspondiente
    return machine_dict.get(machine_name, machine_name)  # Si machine_name no está en el diccionario, retorna el valor original    



# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\FuntionLaserPlasmaFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

def process_pdfs_laser_plasma_and_generate_excel(pdf_files, input_date, update_progress_func):
    all_utilization_lines = []
    all_headers = []  # Lista de headers
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las fechas Jake de cada archivo

    headers = pd.DataFrame()  # Inicializamos headers como un DataFrame vacío para evitar el error

    total_files = len(pdf_files)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados

    for pdf_path in pdf_files:
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process = input_date
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake = match.group(1)     

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[9]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                    elif table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)  

                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[1] = utilization_data[1].replace("Sheet Name: ", "")
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake) 
                                utilization_data.append(category) 
                                all_utilization_lines.append(utilization_data)

        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size','Nest','Date process','Date Jake','category'])
    else:
        print("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")

    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        utilization_df['Sheet Name'] = utilization_df['Sheet Name'].astype(str)
        
        # Crear una lista para almacenar los resultados del join
        merged_rows = []

        # Iterar por cada fila en headers
        for _, header_row in headers.iterrows():
            # Extraer la parte después del primer guion en 'Sheet Name' de utilization_df
            utilization_df['Sheet Name Part'] = utilization_df['Sheet Name'].apply(lambda x: x.split('-')[1] if '-' in x else '')

            # Filtrar filas en utilization_df donde 'NESTING' y la parte extraída de 'Sheet Name' estén contenidos en 'SCHEDULE'
            matching_rows = utilization_df[
                utilization_df['NESTING'].apply(lambda x: x in header_row['SCHEDULE']) &
                utilization_df['Sheet Name Part'].apply(lambda x: x in header_row['SCHEDULE'])
            ]
            
            # Si hay coincidencias, combinar las filas
            for _, util_row in matching_rows.iterrows():
                merged_rows.append({**header_row.to_dict(), **util_row.to_dict()})

        # Convertir la lista de resultados a un DataFrame
        merged_df = pd.DataFrame(merged_rows)

        # Eliminar duplicados
        merged_df = merged_df.drop_duplicates()

        # Eliminar la columna 'SCHEDULE' si es necesario
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

        # Agregar la columna "Material"
        merged_df['Material'] = merged_df['NESTING'].apply(lambda x: get_material(x, Material_dict))

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        merged_df['MACHINE'] = merged_df['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))

        merged_df['Gauge'] = merged_df['Sheet Name'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
            'MACHINE': 'Machine',
            'Nest': 'Nesting',
            'Sheet Size': 'Size',
            'category': 'Type',
            'Material': 'Material',
            'Gauge': 'Gage',
            'Stack Qty': '# Sheets',
            'Program': 'Program',
            'Utililization': 'Utilization',
            'Date process': 'Date-Nesting',
            'Date Jake': 'Date Jake',
            'TOTAL CUT TIME': 'Cut Time'
        })

        # Reordenar las columnas
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization'
        ]

        merged_df = merged_df[new_order]
        merged_df['Utilization'] = pd.to_numeric(merged_df['Utilization'], errors='coerce')

        # Realizar la conversión a decimal (dividir entre 100)
        merged_df['Utilization'] = merged_df['Utilization'] / 100

        # Eliminar duplicados de nuevo, por si hubo alguna duplicidad después de los ajustes
        merged_df = merged_df.drop_duplicates()

        df_consolidado1 = pd.DataFrame(merged_df)
        print(df_consolidado1)

    else:
        print("Error", "No se pueden combinar los DataFrames.")
    
    return df_consolidado1


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\FuntionListFiles.py

import os
from datetime import datetime
import re

def list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter):
    turret_files = []
    laser_plasma_files = []

    # Combina las dos rutas de directorios
    folder_paths = [folder_path_1, folder_path_2]

    # Ahora iteramos sobre los directorios y procesamos los archivos
 
    for folder_path in folder_paths:
        for root, dirs, files in os.walk(folder_path):
            for file_name in files:                
                file_path = os.path.join(root, file_name)
                file_mod_time = os.path.getmtime(file_path)
                file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")
                # Verifica si el archivo es PDF y si la fecha coincide
                if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                    #print(f"Procesando archivo: {file_path}")    
                    # Clasifica según si el nombre contiene "turret"
                    #if "turret" in file_path.lower():
                    if re.search(r"turret", file_path, re.IGNORECASE): 
                        ##print("turret",file_path)
                        turret_files.append(file_path)
                    else:
                        laser_plasma_files.append(file_path)
                        ###print("laser",file_path)

    # Imprime las listas de archivos
    ##print("Archivos 'turret':", len(turret_files))
    ##print("Archivos 'laser_plasma':", len(laser_plasma_files))

    return turret_files, laser_plasma_files

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\Funtionscode.py






# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)

def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []       
    

  


# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

def list_pdf_files_in_folder_combined(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato mm_dd_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date and file_name.lower().endswith(".pdf"):
                        # Agregar archivo PDF a la lista sin ninguna exclusión
                        pdf_files.append(file_path)
                        
                        # Imprimir el path y la categoría (opcional)
                        category = get_category_from_path(file_path)
                        nest = get_nest_from_path(file_path)
                        print(f"Archivo: {file_path} | Categoría: {category} | Nido: {nest}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []    

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\FuntionsGetdata.py
import os
from datetime import datetime, timedelta
import re
from funtions.Dictionaries import type_dict

def get_today_date():
    today = datetime.today()
    # Restar un día
    yesterday = today - timedelta(days=5)
    # Retornar la fecha en el formato m_d_yyyy
    return yesterday.strftime("%m_%d_%Y")

def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

def convert_date_format_2(date_str):
    # Convertir la cadena en un objeto datetime usando el formato original
    date_obj = datetime.strptime(date_str, '%m_%d_%Y')
    # Formatear el objeto datetime al nuevo formato
    return date_obj.strftime('%m/%d/%Y')

'''
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"'''


def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Obtener la categoría encontrada
        category = match.group(0)
        
        # Buscar la categoría en el diccionario y devolver el valor correspondiente
        return type_dict.get(category, "MISC")  # Si no se encuentra, devuelve "MISC"
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"

    
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"    


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\FuntionsInterfaz.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *
from config import *


############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *


# Crear la ventana principal
root = ThemedTk(theme="arc")
root.title("Interfaz de Selección de Carpeta")

# Crear y colocar el label para mostrar la ruta de la carpeta seleccionada
folder_path_label = Label(root, text="No se ha seleccionado ninguna carpeta")
folder_path_label.pack(pady=10)

# Crear y colocar el widget de texto para el log
log_text = Text(root, state=tk.DISABLED, height=10, width=50)
log_text.pack(pady=10)

# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\funtions\FuntionTurretFiles.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime, timedelta
import re

############################################################################################################
#importar diccionarios
############################################################################################################
from funtions.Dictionaries import *

############################################################################################################
#importar funciones
############################################################################################################
from funtions.FuncionesFieldsTransform import *
from funtions.FuntionLaserPlasmaFiles import *
from funtions.FuntionListFiles import *
from funtions.FuntionsGetdata import *
from funtions.FuntionTurretFiles import *

##folder_path = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
##excel_output_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting - turret'

consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF

def process_pdfs_turret_and_generate_excel(fullpath, input_date, update_progress_func):
    consolidated_df = pd.DataFrame()  # Inicializar el DataFrame consolidado para cada archivo PDF
    
    total_files = len(fullpath)  # Total de archivos a procesar
    processed_files = 0  # Contador de archivos procesados
    
    for pdf_path in fullpath:
        print(f"Procesando el archivo: {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:    
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if table[0] == ['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        i = 3
                        while i < len(table):
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')
                        #print(df_final.columns)

                        df_final['Date process'] = input_date
                        match = re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
                        df_final['Date Jake'] = match.group(1)                        
                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
                        #df_final['Sheet Size'] = df_final['Sheet Size'].apply(lambda x: update_sheet_size(x, valid_sizes))
                        df_final['MACHINE'] = df_final['MACHINE'].apply(lambda x: adjust_machine(x, machine_dict))
                        df_final['Utilization'] = ''
                        df_final['Category'] = adjust_type_nest(get_category_from_path(pdf_path), type_dict)
                        df_final['Material Name'] = df_final['SCHEDULE'].apply(lambda x: get_material(x, Material_dict))
                        df_final['Sheet Name'] = df_final['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        df_final['Date Jake'] = df_final['Date Jake'].apply(lambda x: convert_date_format_2(x))
                        df_final['Date process'] = df_final['Date process'].apply(lambda x: convert_date_format_2(x))

                        df_final.rename(columns={
                            'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                            'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                            'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                            'Cut Time': 'Cut Time'
                        }, inplace=True)

                        # Reordenar columnas
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization'
                        ]
                        df_final = df_final[new_order]
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)

        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Devolver el DataFrame consolidado
    df_consolidado2 = pd.DataFrame(consolidated_df) 
    return df_consolidado2


# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\Old_version&testing\rpa_utilization_nest_202412301127.py
import os
import pandas as pd
import pdfplumber
from datetime import datetime
from tkinter import filedialog, messagebox, Tk, Label, Button, Entry, ttk
from tkinter import DoubleVar
import re
from tkinter import *
from tkcalendar import Calendar
from PIL import Image, ImageTk  # Para manejar la imagen de fondo
from ttkthemes import ThemedTk  # Importamos ThemedTk para usar temas
import tkinter as tk


# Ruta de la carpeta de salida (ajusta según sea necesario)

##excel_output_path = r"C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilizacion_nesting\Excel_report"


def read_config():
    configuration = {}
    with open('config.txt', 'r') as f:
        for linea in f:
            # Eliminar espacios y saltos de línea
            linea = linea.strip()
            # Evitar líneas vacías
            if not linea:
                continue
            # Asegurarse de que la línea contiene un '='
            if '=' in linea:
                key, value = linea.split('=', 1)  # Separar clave y valor
                # Eliminar posibles espacios alrededor de las claves y valores
                key = key.strip()
                value = value.strip()
                # Guardar la clave y su valor en el diccionario
                configuration[key] = value
            else:
                print(f"Advertencia: La línea no tiene el formato adecuado: {linea}")
    return configuration

# Leer configuración desde el archivo config.txt
config = read_config()

# Acceder a la variable 'excel_output_path' y verificar su valor
excel_output_path = config.get('excel_output_path', None)

if excel_output_path:
    print(f'Ruta de salida configurada: {excel_output_path}')
else:
    print('No se encontró la ruta "excel_output_path" en el archivo de configuración.')

# Verificar si la ruta existe y hacer algo con ella
import os
if excel_output_path and os.path.exists(excel_output_path):
    print(f'La ruta existe: {excel_output_path}')
else:
    print('La ruta especificada no existe o no es válida.')

# Definir las categorías
categories = {
    "ENCL": "ENCL",
    "ENGR": "ENGR",
    "MEC": "MEC",
    "PARTS ORDER": "PARTS ORDER",
    "REWORK": "REWORK",
    "SIL": "SIL",
    "TANK": "TANK"
}

# Diccionario para la asignación de materiales
category_dict = {
    "1. Steel": {"tiene": ["STEEL"], "no_tiene": ["STEEL-PERF"]},
    "2. Galv": {"tiene": ["GALV"], "no_tiene": []},
    "3. Diamond Plate": {"tiene": ["DMND"], "no_tiene": []},
    "4. Alum": {"tiene": ["3000-ALU"], "no_tiene": []},
    "5. Perf": {"tiene": ["STEEL-PERF"], "no_tiene": []},
    "6. Steel 304ss": {"tiene": ["304-SS"], "no_tiene": []},
    "7. Steel 316ss": {"tiene": ["316-SS"], "no_tiene": []},
    "8. Steel Perf": {"tiene": ["SS-PERF"], "no_tiene": []}
}

# Función para convertir la fecha de entrada al formato dd_mm_yyyy (como string)
def convert_date_format(date_str):
    try:
        # Convertir la fecha de formato MM/DD/YYYY a dd_mm_yyyy
        date_obj = datetime.strptime(date_str, "%m/%d/%Y")    
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        print("Fecha no válida, asegúrate de usar el formato MM/DD/YYYY.")
        return None 

# Función para determinar la categoría basada en las palabras clave dentro del path completo del archivo
def get_category_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = r"(ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la categoría encontrada
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"
    
    # Función para determinar la nesting basada en las palabras clave dentro del path completo del archivo
def get_nest_from_path(file_path):
    # Normalizar la ruta y convertirla a mayúsculas
    normalized_path = os.path.normpath(file_path).upper()

    # Expresión regular para buscar las categorías
    pattern = "\d{2}_\d{2}_\d{4}_[A-Za-z]+(?: \d+)?"
    
    # Buscar las coincidencias en la ruta
    match = re.search(pattern, normalized_path)

    if match:
        # Retornar la nest encontrado
        return match.group(0)
    else:
        # Si no se encuentra ninguna coincidencia, retornar "MISC"
        return "MISC"


code_to_gauge = {
    "3000-ALU-0.080": "080 AL 3000",
    "3000-ALU-0.125": "125 AL 3000",
    "GALV-0.062": "16 GA GALV",
    "GALV-0.078": "14 GA GALV",
    "GALV-0.140": "10 GA GALV",
    "STEEL-0.187": "7 GA",
    "STEEL-0.250": "25 INCH",
    "GALV-0.102": "12 GA GALV",
    "STEEL-0.313": "312 INCH",
    "DMND-PLT-0.250": "DIA 1/4",
    "304-SS-0.140": "10 GA 304SS",
    "STEEL-0.750": "750 INCH",
    "STEEL-0.500": "500 INCH",
    "304-SS-0.250": "25 INCH 304SS",
    "304-SS-0.187": "7 GA 304SS",
    "STEEL-PERF-0.125": "11 GA PERF",
    "316-SS-0.078": "14 GA 316SS",
    "316-SS-0.062": "16 GA 316SS",
    "STEEL-0.140": "10 GA",
    "STEEL-0.078": "14 GA",
    "GALV-0.102": "12 GA GALV",
    "STEEL-PERF-0.062": "16 GA PERF"
}



# Función para listar archivos PDF en una carpeta y sus subcarpetas que contienen la fecha de modificación
def list_pdf_files_in_folder(folder_path, date_filter):
    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            pdf_files = []
            print(f"Verificando archivos en la carpeta y subcarpetas: '{folder_path}' con la fecha {date_filter}")

            for root, dirs, files in os.walk(folder_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    # Obtener la fecha de modificación del archivo y transformarla al formato dd_mm_yyyy
                    file_mod_time = os.path.getmtime(file_path)
                    file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

                    # Comparamos las fechas de modificación con la fecha proporcionada
                    if date_filter == file_mod_date:
                        if 'turret' in file_path.lower():
                            print(f"Excluyendo archivo: {file_path}")
                            continue

                        if file_name.lower().endswith(".pdf"):
                            pdf_files.append(file_path)
                            # Imprimir el path y el resultado de la categoría
                            category = get_category_from_path(file_path)
                            print(f"Archivo: {file_path} | Categoría: {category}")

            return pdf_files
        else:
            print(f"La ruta '{folder_path}' no existe o no es una carpeta válida.")
            return []
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
        return []

# Función para procesar el PDF consolidado y generar el archivo Excel
def process_pdfs_and_generate_excel(pdf_files, input_date, progress_var, progress_bar, percentage_label, root):
    all_utilization_lines = []
    all_headers = []
    all_categories = []  # Para almacenar las categorías de cada archivo
    all_nest = []  # Para almacenar los nest de cada archivo
    all_date_jake = []  # Para almacenar las categorías de cada archivo

    for i, pdf_path in enumerate(pdf_files):
        # Obtener la categoría del archivo
        category = get_category_from_path(pdf_path)
        nest = get_nest_from_path(pdf_path)
        date_process=input_date
        match=re.match(r'^([^_]+_[^_]+_[^_]+)', get_nest_from_path(pdf_path))
        date_jake=match.group(1)

        
        

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, '', 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                for table in tables:
                    for row in table:
                        for cell in row:
                            if cell and "Utililization" in str(cell):
                                utilization_data = [row[i] for i in [0, 2, 4, 5, 7] if row[i] is not None]
                                utilization_data[0] = utilization_data[0].replace("\n", "").strip()
                                utilization_data[2] = utilization_data[2].replace("Stack Qty: ", "")
                                utilization_data[3] = utilization_data[3].replace("Utililization: ", "").strip().replace("%", "")
                                utilization_data.append(category)  # Añadir la categoría a cada fila de datos
                                utilization_data.append(nest)  # Añadir la nest a cada fila de datos
                                utilization_data.append(date_process)  # Añadir la fecha de input a cada fila de datos
                                utilization_data.append(date_jake)  # Añadir la fecha de input a cada fila de datos
                                all_utilization_lines.append(utilization_data)
        
        # Actualizar la barra de progreso
        progress_var.set(i + 1)
        progress_bar.update()
        percentage_label.config(text=f"{(i + 1) / len(pdf_files) * 100:.2f}%")
        root.update_idletasks()

    if all_headers:
        headers = pd.concat(all_headers, ignore_index=True)
    if all_utilization_lines:
        utilization_df = pd.DataFrame(all_utilization_lines, columns=['NESTING', 'Sheet Name', 'Stack Qty', 'Utililization', 'Sheet Size', 'Category','Nest','Date process','Date Jake'])
    else:
        messagebox.showwarning("Advertencia", "No se encontraron líneas con 'Utililization' en los PDFs.")
    
    if not headers.empty and not utilization_df.empty:
        headers['SCHEDULE'] = headers['SCHEDULE'].astype(str)
        utilization_df['NESTING'] = utilization_df['NESTING'].astype(str)
        merged_df = pd.merge(headers, utilization_df, left_on='SCHEDULE', right_on='NESTING', how='left')
        merged_df.drop(columns=['SCHEDULE'], inplace=True)

    

        # Agregar la columna "Material"
        def get_material(nesting):
            for material, values in category_dict.items():
                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                    return material
            return "Desconocido"          

        merged_df['Material'] = merged_df['NESTING'].apply(get_material)



      

        # Agregar la columna "Program"
        merged_df['Program'] = merged_df['Sheet Name'].apply(lambda x: x.split('-')[0] if '-' in x else x)

        # Ajustar la columna 'MACHINE' según las condiciones especificadas
        def adjust_machine(machine_name):
            if machine_name == "Amada_ENSIS_4020AJ":
                return "1. Laser"
            elif machine_name == "Messer_170Amp_Plasm":
                return "2.Plasma"
            elif machine_name == "Amada_Vipros_358K":
                return "3.Turret"
            elif machine_name == "Amada_EMK316M2":
                return "4.Citurret"
            else:
                return machine_name   

        merged_df['MACHINE'] = merged_df['MACHINE'].apply(adjust_machine)  # Aquí se aplica el ajuste

        

        def get_gauge_from_nesting(nesting, code_to_gauge):
            # Iteramos sobre las claves (Code) y los valores (Gauge) del diccionario
            for Code, Gauge in code_to_gauge.items():
                # Verificamos si el 'Code' está contenido en 'nesting'
                if Code in nesting:
                    return Gauge  # Retornamos el valor (Gauge) correspondiente
    
            return "Desconocido"  # Si no hay coincidencias, devolvemos "Desconocido"

        merged_df['Gauge'] = merged_df['NESTING'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))
        #     # Normalize the nesting value and convert it to uppercase for case-insensitive comparison  

        def update_sheet_size(sheet_size):
            # Eliminar los espacios en la cadena y convertir todo a minúsculas
            cleaned_size = sheet_size.replace(" ", "").lower()
            
            # Lista de tamaños válidos (en formato "60x144")
            valid_sizes = [
                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
            ]
            
            # Invertir la cadena limpia para verificar la variante invertida
            inverted_cleaned_size = 'x'.join(cleaned_size.split('x')[::-1])  # Volteamos la parte antes y después de 'x'
            
            # Verificar si el tamaño limpio o su versión invertida están en la lista de tamaños válidos
            if cleaned_size in valid_sizes or inverted_cleaned_size in valid_sizes:
                return sheet_size  # Conservar el valor original si hay coincidencia
            else:
                return "REMNANT"  # Marcar como "REMNANT" si no hay coincidencia
            
        merged_df['Sheet Size'] = merged_df['Sheet Size'].apply(update_sheet_size)
        merged_df['Date process'] = pd.to_datetime(merged_df['Date process'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')
        merged_df['Date Jake'] = pd.to_datetime(merged_df['Date Jake'], format='%m_%d_%Y').dt.strftime('%m/%d/%Y')

        merged_df = merged_df.rename(columns={
                'MACHINE': 'Machine',
                'Nest': 'Nesting',
                'Sheet Size': 'Size',
                'Category': 'Type',
                'Material': 'Material',
                'Gauge': 'Gage',
                'Stack Qty': '# Sheets',
                'Program': 'Program',
                'Utililization': 'Utilization',
                'Date process': 'Date-Nesting',
                'Date Jake': 'Date Jake',
                'TOTAL CUT TIME': 'Cut Time by Nest_material'
            })

            # Ahora reordenamos las columnas para que sigan el orden especificado
        new_order = [
            'Date-Nesting', 'Date Jake', 'Nesting',  'Type', 'Material', 
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time by Nest_material'
        ]

            # Asegurarse de que las columnas estén en el orden deseado
        merged_df = merged_df[new_order]

                # Ajustar la columna 'Type' según las condiciones especificadas (ENCL|ENGR|MEC|PARTS ORDER|REWORK|SIL|TANK)
        def adjust_type_nest(Type):
            if Type == "ENCL":
                return "2.Encl"
            elif Type == "ENGR":
                return "6.Engr"
            elif Type == "MEC":
                return "3.Mec"
            elif Type == "PARTS ORDER":
                return "4.Parts Order"
            elif Type == "REWORK":
                return "5.Rework"
            elif Type == "SIL":
                return "4.Sil"
            elif Type == "TANK":
                return "1.Tank"            
            else:
                return Type   

        merged_df['Type'] = merged_df['Type'].apply(adjust_type_nest)  # Aquí se aplica el ajuste

        # Exportar a Excel
        '''try:
            output_file = os.path.join(excel_output_path, f"Utilization_{input_date}.xlsx")
            merged_df.to_excel(output_file, index=False)
            messagebox.showinfo("Éxito", f"Excel exportado correctamente: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo Excel: {e}")
    else:
        messagebox.showerror("Error", "No se pueden combinar los DataFrames.")'''


############################################################################################################
# Función para listar archivos PDF turret en una carpeta y sus subcarpetas que contienen la fecha de modificación
############################################################################################################       
# Listar archivos PDF en la carpeta y subcarpetas con la fecha de modificación deseada
def list_pdf_files_turret_in_folder(folder_path, date_filter):
    pdf_files_turret = []
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            file_mod_time = os.path.getmtime(file_path)
            file_mod_date = datetime.fromtimestamp(file_mod_time).strftime("%m_%d_%Y")

            if date_filter == file_mod_date and file_name.lower().endswith(".pdf") and 'turret'  in file_path.lower():
                pdf_files_turret.append(file_path)
                category = get_category_from_path(file_path)
                nest = get_nest_from_path(file_path)
                date_process = file_mod_date
                


    return pdf_files_turret

# Función para procesar los PDFs y generar el archivo Excel
def process_pdfs_turret_and_generate_excel(pdf_files_turret, input_date):
     all_dfs = []  # Lista para almacenar todos los DataFrames
     for pdf_path in pdf_files_turret:
        category = get_category_from_path(pdf_path)
        #print(category)
        nest = get_nest_from_path(pdf_path)
        #print(nest)
        date_process = input_date
        #print(date_process)
        match = re.match(r'^([^_]+_[^_]+_[^_]+)', nest)
        date_jake = match.group(1) if match else "MISC"
        date_process=date_process.replace("_","/")


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if   table[0] ==['MACHINE', 'SCHEDULE', None, None, None, '', 'TOTAL CUT TIME', None, None, None]  :
                                            
                        tabla_1 = []
                        tabla_2 = []
                        tabla_3 = []

                        # Primera tabla: los dos primeros elementos
                        tabla_1.append(table[0])  # Encabezados
                        tabla_1.append(table[1])  # Valores

                        # Segunda tabla: inicia en la fila 2 y toma los encabezados de la fila 2
                        tabla_2.append(table[2])  # Encabezados
                        # Después tomamos todas las filas siguientes hasta la fila que contiene el encabezado de la tercera tabla
                        i = 3
                        while i < len(table):
                            # Si encontramos el encabezado de la tercera tabla, dejamos de agregar a la segunda tabla
                            if table[i] == ['Sheet Name', None, 'Sheet Size', None, 'Stack Qty', None, None, 'Cut Time', None, 'Finish']:
                                break
                            tabla_2.append(table[i])
                            i += 1

                        # Tercera tabla: desde el encabezado identificado hasta el final
                        tabla_3.append(table[i])  # Encabezados de la tercera tabla
                        i += 1
                        while i < len(table):
                            tabla_3.append(table[i])
                            i += 1

                        # Convertir las tablas a DataFrames usando pandas
                        df_1 = pd.DataFrame(tabla_1[1:], columns=tabla_1[0])
                        df_1['key'] = 1
                        df_2 = pd.DataFrame(tabla_2[1:], columns=tabla_2[0]) 
                        df_2['key'] = 1 
                        df_cross = pd.merge(df_1, df_2, on='key').drop('key', axis=1)
                        df_3 = pd.DataFrame(tabla_3[1:], columns=tabla_3[0])  # La primera fila es el encabezado
                        df_final = pd.merge(df_cross, df_3, on='Sheet Size')#.drop('key', axis=1)

                        df_final = df_final[
                            ['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME', 'Material Name', 'Sheet Code', 'Sheet Size', 
                            'Sheet Count', 'Sheet Name', 'Stack Qty', 'Cut Time', 'Finish']
                        ]
                        def adjust_machine(machine_name):
                            adjustments = {
                                "Amada_ENSIS_4020AJ": "1. Laser", "Messer_170Amp_Plasm": "2.Plasma", 
                                "Amada_Vipros_358K": "3.Turret", "Amada_EMK316M2": "4.Citurret"
                            }
                            return adjustments.get(machine_name, machine_name)
                        
                        df_final['MACHINE'] = df_final['MACHINE'].apply(adjust_machine)

                        def get_gauge_from_nesting(SCHEDULE, code_to_gauge):
                            for code, gauge in code_to_gauge.items():
                                if code in SCHEDULE:
                                    return gauge
                            return "Desconocido"

                        df_final['Gauge'] = df_final['SCHEDULE'].apply(lambda x: get_gauge_from_nesting(x, code_to_gauge))

                        def update_sheet_size(sheet_size):
                            valid_sizes = [
                                "60x86", "60x120", "60x144", "72x120", "72x144", "48x144", "60x145", 
                                "72x145", "72x187", "60x157", "48x90", "48x157", "48x120", "48x158", 
                                "48x96", "23.94x155", "23.94x190", "23.94x200", "33x121", "33x145", 
                                "24.83x133", "24.83x170", "24.83x190", "24.83x205", "24.96x145", 
                                "24.96x200", "28x120", "28x157", "44.24x128", "44.24x151", "24.45x145", 
                                "24.45x205", "27.92x120", "27.92x157", "44.21x120", "44.21x157", 
                                "60x175", "60x187", "60x164", "60x172", "60x163", "45x174", "60x162", 
                                "23.94x205", "23.94x191.75", "24.83x206.75", "24.83x149", "33x149", "60x146"
                            ]
                            cleaned_size = sheet_size.replace(" ", "").lower()
                            inverted_size = 'x'.join(cleaned_size.split('x')[::-1])  # Reverse size
                            return sheet_size if cleaned_size in valid_sizes or inverted_size in valid_sizes else "REMNANT"
                        

                        

                        df_final['Sheet Sizex'] = df_final['Sheet Size'].apply(update_sheet_size)
                        df_final['Date process'] = date_process
                        df_final['Date Jake'] = date_jake
                        df_final['Utilization']=''
                        def adjust_type_nest(Type):
                            adjustments = {
                                "ENCL": "2.Encl", "ENGR": "6.Engr", "MEC": "3.Mec", "PARTS ORDER": "4.Parts Order", "REWORK": "5.Rework",
                                "SIL": "4.Sil", "TANK": "1.Tank"
                            }
                            return adjustments.get(Type, Type)

                        df_final['Category'] = adjust_type_nest(category)

                        df_final.rename(columns={
                        'MACHINE': 'Machine', 'SCHEDULE': 'Nesting', 'Sheet Size': 'Size', 'Category': 'Type',
                        'Material Name': 'Material', 'Gauge': 'Gage', 'Stack Qty': '# Sheets', 'Sheet Name': 'Program',
                        'Utilization': 'Utilization', 'Date process': 'Date-Nesting', 'Date Jake': 'Date Jake',
                        'Cut Time': 'Cut Time'
                    }, inplace=True)
                        
                        new_order = [
                            'Date-Nesting', 'Date Jake', 'Nesting', 'Type', 'Material', 'Gage', 'Size', 'Program', '# Sheets', 'Machine',
                            'Utilization', 'Cut Time'
                        ]
                        df_final = df_final[new_order]

                        def get_material(nesting):
                            for material, values in category_dict.items():
                                if any(item in nesting for item in values["tiene"]) and not any(item in nesting for item in values["no_tiene"]):
                                    return material
                            return "Desconocido"          

                        df_final['Material'] = df_final['Nesting'].apply(get_material)

                        df_final['Program'] = df_final['Program'].apply(lambda x: x.split('-')[0] if '-' in x else x)
                        
                        all_dfs.append(df_final)

                        # Consolidar todos los DataFrames en uno solo
        consolidated_df = pd.concat(all_dfs, ignore_index=True)

        # Guardar el DataFrame consolidado en un archivo CSV
        output_file = excel_output_path + 'Utilization_consolidated_' + str(input_date) + '.csv'
        consolidated_df.to_csv(output_file, index=False)
        print("Todos los archivos procesados y consolidados exitosamente.")
        print(consolidated_df) 



# Función para seleccionar la carpeta de entrada usando un cuadro de diálogo
def select_folder():
    folder_path = filedialog.askdirectory(title="Selecciona la carpeta de entrada")
    folder_path_label.config(text=folder_path)
    return folder_path

# Función para convertir la fecha al formato deseado MM_DD_YYYY
def convert_date_format(input_date):
    try:
        # Convertir la fecha al formato MM_DD_YYYY
        date_obj = datetime.strptime(input_date, "%m/%d/%Y")
        return date_obj.strftime("%m_%d_%Y")
    except ValueError:
        return None

# Función para mostrar mensajes en el log
def log_message(message):
    log_text.config(state=tk.NORMAL)  # Hacer el widget editable
    log_text.insert(tk.END, message + "\n")  # Insertar el mensaje en el log
    log_text.yview(tk.END)  # Desplazar el texto hacia abajo para mostrar los últimos mensajes
    log_text.config(state=tk.DISABLED)  # Volver a desactivar la edición

# Función principal que se ejecuta al presionar el botón
def main():
    log_message("Iniciando el proceso...")

    # Obtener la fecha de entrada desde el calendario
    input_date = date_calendar.get_date()  # Fecha en formato "mm/dd/yyyy"
    date_filter = convert_date_format(input_date)  # Convertir la fecha al formato deseado

    if not date_filter:
        messagebox.showerror("Error", "Fecha inválida. Asegúrate de usar el formato MM/DD/YYYY.")
        log_message("Error: Fecha inválida.")
        return

    # Obtener la ruta de la carpeta de entrada desde el label
    folder_path = folder_path_label.cget("text")

    if not folder_path:
        messagebox.showerror("Error", "Por favor selecciona una carpeta de entrada.")
        log_message("Error: No se seleccionó carpeta de entrada.")
        return

    # Listar archivos PDF (Esta función necesita estar definida en otro lugar de tu código)
    pdf_files = list_pdf_files_in_folder(folder_path, date_filter)
    pdf_files_turret = list_pdf_files_turret_in_folder(folder_path, date_filter)

    if not pdf_files:
        messagebox.showerror("Error", "No se encontraron archivos PDF para procesar.")
        log_message("Error: No se encontraron archivos PDF.")
        return

    log_message(f"Se encontraron {len(pdf_files)} archivos PDF para procesar.")

    # Configurar la barra de progreso
    progress_var = DoubleVar()

    # Configuración del estilo para la barra de progreso
    style = ttk.Style()
    style.configure("TProgressbar",
                    thickness=30,  # Cambiar el grosor de la barra de progreso
                    background="#F1C232",  # Amarillo para la barra de progreso
                    troughcolor="#2F2F2F",  # Fondo negro de la barra
                    )

    # Modificar el color del relleno de la barra de progreso (la parte que se llena)
    style.map("TProgressbar",
              foreground=[("active", "#F1C232")],  # El color amarillo cuando la barra está activa
              )

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(pdf_files), length=400, style="TProgressbar")
    progress_bar.grid(row=6, column=0, columnspan=3, pady=10)  # Colocamos la barra de progreso

    percentage_label = Label(root, text="0%", fg="#F1C232", bg="#2F2F2F", font=("Arial", 12, "bold"))  # Amarillo más suave, gris oscuro
    percentage_label.grid(row=7, column=0, columnspan=3)  # Colocamos el porcentaje de la barra de progreso

    # Procesar los PDFs y generar el archivo Excel (Esta función necesita estar definida en otro lugar)
    process_pdfs_and_generate_excel(pdf_files, date_filter, progress_var, progress_bar, percentage_label, root)
    process_pdfs_turret_and_generate_excel(pdf_files_turret, date_filter)

    log_message("El procesamiento de los archivos PDF ha finalizado.")
    messagebox.showinfo("Proceso finalizado", "El procesamiento de los archivos PDF ha finalizado.")

    # Mostrar mensaje preguntando si desea cerrar la aplicación
    if messagebox.askyesno("Cerrar la aplicación", "¿Deseas cerrar la aplicación?"):
        root.quit()  # Esto cerrará la interfaz de la aplicación
        root.destroy()  # Esto finalizará todos los procesos y recursos de la aplicación

# Configuración de la interfaz gráfica
root = ThemedTk(theme="breeze")  # Aplica el tema 'breeze' para una apariencia moderna

root.title("Usage Report Generator")

# Fijar el tamaño de la ventana
root.geometry("600x700")  # Establece el tamaño de la ventana (ancho x alto)
root.resizable(False, False)  # Desactiva la redimensión de la ventana

# Fondo gris oscuro para la ventana
root.configure(bg="#2F2F2F")  # Gris oscuro

# Crear un canvas para colocar el logo, con fondo blanco solo para el logo
canvas = Canvas(root, width=600, height=100, bg="#FFFFFF")
canvas.grid(row=0, column=0, columnspan=3)  # Aseguramos que el logo esté sobre la parte superior

# Cargar el logo de la app (asegúrate de tener la imagen en el directorio correcto)
logo_image = Image.open("logo.png")  # Aquí coloca la ruta de tu imagen del logo

# Ajustamos el tamaño del logo proporcionalmente
logo_width = 80  # Definimos un tamaño pequeño para el logo
logo_height = int(logo_image.height * (logo_width / logo_image.width))  # Mantener la proporción
logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)  # Redimensionamos la imagen
logo_photo = ImageTk.PhotoImage(logo_image)

# Colocar el logo en la parte superior
canvas.create_image(10, 10, image=logo_photo, anchor="nw")  # "nw" significa "noroeste" (top-left)

# Etiqueta y calendario para seleccionar la fecha
date_label = Label(root, text="Selecciona la fecha:", font=("Helvetica", 16, "bold"), fg="#F1C232", bg="#2F2F2F", pady=10)  # Amarillo más suave
date_label.grid(row=1, column=0, columnspan=3, pady=10)

date_calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy", foreground="#F1C232", background="#2F2F2F", bordercolor="#F1C232", headersbackground="#8E8E8E", normalbackground="#6C6C6C", normalforeground="#F1C232")
date_calendar.grid(row=2, column=0, columnspan=3, pady=10)

# Etiqueta y botón para seleccionar la carpeta
folder_path_label = Label(root, text="", font=("Helvetica", 12), fg="#F1C232", bg="#2F2F2F")
folder_path_label.grid(row=3, column=0, columnspan=3, pady=10)

select_folder_button = Button(root, text="Seleccionar carpeta", command=select_folder, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#A0A0A0", relief="flat", padx=20, pady=10)  # Gris suave, texto negro
select_folder_button.grid(row=4, column=0, columnspan=3, pady=10)

# Botón para ejecutar el procesamiento
process_button = Button(root, text="Procesar PDFs", command=main, font=("Helvetica", 12, "bold"), fg="#2F2F2F", bg="#F1C232", relief="flat", padx=20, pady=10)  # Amarillo más suave, texto gris oscuro
process_button.grid(row=5, column=0, columnspan=3, pady=10)

# Crear el widget de log en la parte inferior
log_frame = tk.Frame(root, bg="#2F2F2F", height=100)
log_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

log_scrollbar = Scrollbar(log_frame, orient="vertical")
log_scrollbar.pack(side="right", fill="y")

log_text = Text(log_frame, height=5, wrap="word", font=("Helvetica", 10), fg="#F1C232", bg="#2F2F2F", bd=0, insertbackground="white", yscrollcommand=log_scrollbar.set)
log_text.pack(side="left", fill="both", expand=True)
log_text.config(state=tk.DISABLED)

log_scrollbar.config(command=log_text.yview)



# Esta es la línea que mantiene la interfaz de Tkinter en ejecución hasta que el usuario decida cerrarla.
root.mainloop() 

# Contenido de C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization\Utilizacion_nesting_V2.0_20250127_T1718 - Copy\Old_version&testing\test.py

import os
from datetime import datetime
from funtions.FuntionListFiles import list_all_pdf_files_in_folders

folder_path_1 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_2 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
date_filter = "01_10_2025"

def update_progress_func(processed_files, total_files):
    progress = processed_files / total_files * 100
    print(f"Progreso: {progress:.2f}%")


list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter)

