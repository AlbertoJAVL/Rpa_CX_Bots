import os
import shutil

def eliminar_archivos():

    # Obtiene la ruta de la carpeta %temp%
    temp_folder = os.environ['TEMP']

    try:
        # Listar los archivos en la carpeta %temp%
        temp_files = os.listdir(temp_folder)

        # Eliminar solo los archivos que no están siendo utilizados por otros procesos
        for temp_file in temp_files:
            temp_file_path = os.path.join(temp_folder, temp_file)
            try:
                # Intentar eliminar el archivo
                if os.path.isfile(temp_file_path):
                    os.remove(temp_file_path)
                elif os.path.isdir(temp_file_path):
                    shutil.rmtree(temp_file_path)  # Si es un directorio, elimínalo recursivamente
            except Exception as e:
                print(f"No se pudo eliminar {temp_file}: {str(e)}")

        print("Archivos no utilizados eliminados correctamente en la carpeta %temp%.")

    except Exception as e:
        print(f"Se produjo un error al eliminar archivos en la carpeta %temp%: {str(e)}")
