





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