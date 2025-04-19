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