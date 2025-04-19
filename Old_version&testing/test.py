
import os
from datetime import datetime
from funtions.FuntionListFiles import list_all_pdf_files_in_folders

folder_path_1 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\01 - LASER\2025'
folder_path_2 = r'C:\Users\Lrojas\globalpowercomponents\Nesting - Documents\02 - PLASMA\2025'
date_filter = "01_10_2025"

def update_progress_func(processed_files, total_files):
    progress = processed_files / total_files * 100
    print(f"Progreso: {progress:.2f}%")


list_all_pdf_files_in_folders(folder_path_1, folder_path_2, date_filter)