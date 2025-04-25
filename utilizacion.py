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
        ##outputfinal['Cut Time'] = None
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
