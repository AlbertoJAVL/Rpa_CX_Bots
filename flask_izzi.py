from flask import Flask,request, jsonify
import socket
from dotenv import load_dotenv, find_dotenv
import os



# Función para cargar el archivo .env
def load_env_file():
    dotenv_path = find_dotenv()
    if dotenv_path:
        load_dotenv(dotenv_path)

# Función para modificar el archivo .env con los nuevos valores de proceso y status
def update_env_proceso_status(proceso, status):
    dotenv_path = find_dotenv()
    if dotenv_path:
        # Lee el archivo .env y almacena las líneas en una lista
        with open(dotenv_path, 'r') as file:
            lines = file.readlines()

        # Modifica las variables proceso y status si existen o las agrega
        for i, line in enumerate(lines):
            if line.startswith("proceso="):
                lines[i] = f"proceso={proceso}\n"
                break
        else:
            lines.append(f"proceso={proceso}\n")

        for i, line in enumerate(lines):
            if line.startswith("status="):
                lines[i] = f"status={status}\n"
                break
        else:
            lines.append(f"status={status}\n")

        # Escribe las líneas modificadas de nuevo en el archivo .env
        with open(dotenv_path, 'w') as file:
            file.writelines(lines)

# Cargar el archivo .env
load_env_file()

app = Flask(__name__)

with app.app_context():
    print("El servidor Flask se ha iniciado. Este mensaje se mostrará solo una vez!!!")
    # os.system(f"start python main3.py ")
    # Coloca aquí la lógica que deseas ejecutar una sola vez


# Obtiene la dirección IP de la computadora actual
hostname = socket.gethostname()
ip_address = socket.gethostbyname(hostname)

@app.route('/')
def hello():
    return "Hola"


@app.route('/update_env', methods=['GET'])
def update_proceso_status():
    proceso = request.args.get("proceso")
    status = request.args.get("status")
    
    if proceso is not None and status is not None:
        update_env_proceso_status(proceso, status)
        return jsonify({"message": "Variables actualizadas correctamente"})
    else:
        return jsonify({"error": "Se requieren valores para proceso y status"})

if __name__ == '__main__':
    app.run(host=ip_address, port=3000, debug=False)
