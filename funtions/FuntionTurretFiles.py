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
                            'Utilization','Cut Time'
                        ]
                        df_final = df_final[new_order]
                        consolidated_df = pd.concat([consolidated_df, df_final], ignore_index=True)

        # Actualizar progreso después de procesar cada archivo
        processed_files += 1
        update_progress_func(processed_files, total_files)  # Llamar la función de progreso

    # Devolver el DataFrame consolidado
    df_consolidado2 = pd.DataFrame(consolidated_df) 
    return df_consolidado2
