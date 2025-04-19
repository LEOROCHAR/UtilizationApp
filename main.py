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
