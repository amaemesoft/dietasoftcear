import pandas as pd
import os
import tkinter as tk
from tkinter import simpledialog
import logging
from openpyxl import load_workbook
import requests
import io

# URL de la base de datos en GitHub
base_de_datos_url = "https://raw.githubusercontent.com/amaemesoft/dietasoft/main/basededatos.xlsx"

# Obtener la ubicación del script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Configuración del logging
logging.basicConfig(filename=os.path.join(script_dir, 'registro_errores.log'), level=logging.ERROR, format='%(asctime)s %(levelname)s:%(message)s')

def consultar_base_de_datos(url):
    """Consulta la base de datos directamente desde la URL en GitHub."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        df = pd.read_excel(pd.ExcelFile(io.BytesIO(response.content)))
        return df
    except Exception as e:
        logging.error(f"Error al consultar la base de datos: {e}")
        return None

# Consultar la base de datos desde GitHub
base_de_datos = consultar_base_de_datos(base_de_datos_url)

def buscar_datos_trabajador(dni):
    """Busca los datos de un trabajador por DNI en la base de datos."""
    try:
        datos_trabajador = base_de_datos[base_de_datos['DNI'] == dni]
        if datos_trabajador.empty:
            logging.warning(f"No se encontró el DNI {dni} en la base de datos.")
            return None
        else:
            return datos_trabajador.to_dict(orient='records')[0]
    except Exception as e:
        logging.error(f"Error al buscar datos del trabajador: {e}")
        return None

def leer_modelo_desde_github(url):
    """Lee un archivo Excel desde una URL de GitHub."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except Exception as e:
        logging.error(f"Error al leer el modelo de dieta desde GitHub: {e}")
        return None

def seleccionar_modelo_dieta(datos_trabajador):
    # URL base para los modelos de dieta en GitHub
    base_url = "https://raw.githubusercontent.com/amaemesoft/dietasoft/main/"
    
    # Diccionario de modelos de dieta. Asegúrate de que estos nombres coincidan con los de tu repositorio.
    modelos_dieta = {
        '01_PI_ACOGIDA ESTÁNDAR_Dieta': base_url + "01_PI_ACOGIDA%20ESTÁNDAR_Dieta.xlsx",
        '02_PI_ACOGIDA VULNERABLES_Dieta': base_url + "02_PI_ACOGIDA%20VULNERABLES_Dieta.xlsx",
        '03_PI_AUTONOMÍA_Dieta': base_url + "03_PI_AUTONOM%C3%8DA_Dieta.xlsx",
        '04_PI_SERVICIOS DE APOYO, INTERVENCIÓN Y ACOMPAÑAMIENTO_Dieta': base_url + "04_PI_SERVICIOS%20DE%20APOYO%2C%20INTERVENCI%C3%93N%20Y%20ACOMPA%C3%91AMIENTO_Dieta.xlsx",
        '05_PI_VALORACIÓN INICIAL Y DERIVACIÓN_Dieta': base_url + "05_PI_VALORACI%C3%93N%20INICIAL%20Y%20DERIVACI%C3%93N_Dieta.xlsx"
    }
    
    # Obtener el nombre del modelo de la base de datos y construir la URL completa
    nombre_modelo = datos_trabajador.get('MODELO DIETAS')
    url_modelo = modelos_dieta.get(nombre_modelo, base_url + "modelo_por_defecto.xlsx")
    
    return leer_modelo_desde_github(url_modelo)


def rellenar_excel(archivo_modelo, datos_trabajador):
    """Rellena un modelo de Excel con los datos del trabajador, añadiendo a lo existente."""
    try:
        libro = load_workbook(archivo_modelo)
        hoja = libro.active
        
        # Mapeo de celdas en Excel a los datos
        mapeo_celdas_datos = {
            'A9': datos_trabajador['TRABAJADOR/A'],  
            'A10': datos_trabajador['DNI'],           
            'A11': datos_trabajador['DIRECCIÓN CENTRO'],
            'B1': str(datos_trabajador['INMUEBLE DIETAS']),
            'B1': str(datos_trabajador['ANALÍTICA DIETAS']),
            'B1': str(datos_trabajador['AREA DIETAS'])
        }

        # Ruta para la nueva carpeta de documentos
        ruta_documentos = os.path.join(script_dir, 'documentos')

        # Crear la carpeta si no existe
        os.makedirs(ruta_documentos, exist_ok=True)

        # Ruta para el nuevo documento
        ruta_documento_nuevo = os.path.join(ruta_documentos, f"Dieta_{datos_trabajador['DNI']}.xlsx")

        # Concatenar los nuevos datos con el contenido existente en las celdas
        for celda, dato in mapeo_celdas_datos.items():
            contenido_actual = hoja[celda].value or ""
            hoja[celda] = contenido_actual + " " + dato

        # Guardar el documento
        libro.save(ruta_documento_nuevo)
        logging.info(f"Documento Excel guardado en: {ruta_documento_nuevo}")

        return ruta_documento_nuevo
    except Exception as e:
        logging.error(f"Error al rellenar el documento Excel: {e}")
        return None

def abrir_documento(dni, archivo_modelo):
    """Abre el documento Excel con los datos del trabajador."""
    try:
        if archivo_modelo is None:
            return

        ruta_documento_nuevo = rellenar_excel(archivo_modelo, buscar_datos_trabajador(dni))
        if ruta_documento_nuevo is None:
            return

        os.startfile(ruta_documento_nuevo)
    except Exception as e:
        logging.error(f"Error al abrir el documento: {e}")

# Interfaz gráfica para ingresar el DNI
root = tk.Tk()
root.withdraw()
dni = simpledialog.askstring("Input", "Introduce tu DNI:", parent=root)
if dni:
    archivo_modelo = seleccionar_modelo_dieta(buscar_datos_trabajador(dni))
    if archivo_modelo:
        abrir_documento(dni, archivo_modelo)
