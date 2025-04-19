
import os

# Ruta de la carpeta donde se encuentran los archivos
folder_path = r'C:\Users\Lrojas\OneDrive - globalpowercomponents\Proyectos\Utilization'

# Ruta del archivo de salida donde se guardará el código combinado
output_file = 'codigo_combinado.py'

# Función para recorrer directorios y subdirectorios
def agregar_codigo(directorio, archivo_salida):
    for root, dirs, files in os.walk(directorio):
        for filename in files:
            file_path = os.path.join(root, filename)
            
            # Verifica si es un archivo Python (puedes cambiar la extensión si es otro tipo de archivo)
            if filename.endswith('.py'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as infile:
                        # Escribe el nombre del archivo en el archivo de salida
                        archivo_salida.write(f"# Contenido de {file_path}\n")
                        # Escribe el contenido del archivo en el archivo de salida
                        archivo_salida.write(infile.read())
                        archivo_salida.write("\n\n")
                except UnicodeDecodeError:
                    print(f"Error al leer el archivo: {file_path}. Se omitirá este archivo.")

# Abre el archivo de salida en modo escritura
with open(output_file, 'w', encoding='utf-8') as outfile:
    agregar_codigo(folder_path, outfile)

print(f"Todo el código se ha combinado en {output_file}")

