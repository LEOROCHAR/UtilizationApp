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

