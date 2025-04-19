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
