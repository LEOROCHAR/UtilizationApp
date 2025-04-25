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
                        total_cut_time = data[8]
                        print(data[7],data[8],data[9],data[10])
                        headers_df = pd.DataFrame([[machine, schedule, total_cut_time]], columns=['MACHINE', 'SCHEDULE', 'TOTAL CUT TIME'])
                        all_headers.append(headers_df)
                    elif table and table[0] == ['MACHINE', None, 'SCHEDULE', None, None, None, None, None, 'TOTAL CUT TIME', None, None, None]:
                        data = table[1]
                        machine = data[0]
                        schedule = data[2]
                        total_cut_time = data[8]
                        print(data[7],data[8],data[9],data[10])
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
            'Gage', 'Size', 'Program', '# Sheets', 'Machine', 'Utilization','Cut Time'
        ]

        merged_df = merged_df[new_order]
        merged_df['Utilization'] = pd.to_numeric(merged_df['Utilization'], errors='coerce')

        # Realizar la conversión a decimal (dividir entre 100)
        merged_df['Utilization'] = merged_df['Utilization'] / 100

        # Eliminar duplicados de nuevo, por si hubo alguna duplicidad después de los ajustes
        merged_df = merged_df.drop_duplicates()

        df_consolidado1 = pd.DataFrame(merged_df)
        print("est es el df consolidado:" , df_consolidado1)

    else:
        print("Error", "No se pueden combinar los DataFrames.")
    
    return df_consolidado1
