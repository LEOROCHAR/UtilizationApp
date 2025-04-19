
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