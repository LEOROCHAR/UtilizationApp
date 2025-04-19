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