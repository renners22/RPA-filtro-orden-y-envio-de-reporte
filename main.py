import os
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def limpiar_archivos(rutas):
    """
    Elimina archivos especificados si existen.
    """
    print("Limpieza inicial de archivos...")
    for ruta in rutas:
        if os.path.exists(ruta):
            try:
                os.remove(ruta)
                print(f"Archivo eliminado: {ruta}")
            except Exception as e:
                print(f"Error al eliminar el archivo {ruta}: {e}")
        else:
            print(f"Archivo no encontrado: {ruta}")

def abrir_navegador(url):
    """
    Abre el navegador y navega a la URL especificada.
    """
    print(f"Abriendo navegador en: {url}")
    driver = webdriver.Chrome()
    driver.get(url)
    return driver

def iniciar_sesion(driver, usuario, clave):
    """
    Realiza el inicio de sesión en el sitio web utilizando Selenium.
    """
    try:
        print("Iniciando sesión...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "userLogin"))
        ).send_keys(usuario)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "password"))
        ).send_keys(clave)

        driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
        time.sleep(3)
        print("Sesión iniciada correctamente.")
    except Exception as e:
        print(f"Error durante el inicio de sesión: {e}")

def descargar_archivo(driver, url_archivo, ruta_destino):
    """
    Descarga un archivo desde el navegador utilizando Selenium.
    """
    try:
        print(f"Navegando a la URL del archivo: {url_archivo}")
        driver.get(url_archivo)
        time.sleep(5)
        archivo_descargado = "C:\\Users\\name\\Downloads\\archivo..." #editar la ruta donde se descarga el archivo (por defecto downloads/descargas)
        if os.path.exists(archivo_descargado):
            os.rename(archivo_descargado, ruta_destino)
            print(f"Archivo descargado correctamente en: {ruta_destino}")
        else:
            print("Error: No se encontró el archivo descargado.")
    except Exception as e:
        print(f"Error durante la descarga del archivo: {e}")
    finally:
        driver.quit()

def filtrar_bomberos(archivo_entrada, archivo_salida):
    """
    Filtra datos del archivo Excel según la fecha actual.
    """
    print("Filtrando archivo de bomberos...")
    df = pd.read_excel(archivo_entrada)
    df['Creado'] = pd.to_datetime(df['Creado'], errors='coerce')
    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    df_filtrado = df[df['Creado'].dt.strftime('%Y-%m-%d') == fecha_actual]
    df_filtrado.to_excel(archivo_salida, index=False, engine='openpyxl')

    wb = load_workbook(archivo_salida)
    ws = wb.active
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column_letter].width = max_length + 2
    wb.save(archivo_salida)
    print(f"Archivo filtrado y ajustado guardado en: {archivo_salida}")

def enviar_correo(destinatario, asunto, cuerpo_base, archivo_adjunto):
    """
    Envía un correo con un archivo adjunto.
    """
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    cuerpo = f"{cuerpo_base} Fecha: {fecha_actual}"
    smtp_servidor = "smtp.gmail.com"
    smtp_puerto = 587
    remitente = "tu_emailgmail@.com" #colocar tu correo o correo que usaras para enviar el email
    contraseña = "" # esta clave desbes generarla en google "clave de aplicacion"

    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto
    mensaje.attach(MIMEText(cuerpo, "plain"))

    if archivo_adjunto and os.path.exists(archivo_adjunto):
        with open(archivo_adjunto, "rb") as archivo:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(archivo.read())
            encoders.encode_base64(parte)
            parte.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(archivo_adjunto)}"
            )
            mensaje.attach(parte)
    else:
        print(f"Archivo no encontrado: {archivo_adjunto}")

    try:
        servidor = smtplib.SMTP(smtp_servidor, smtp_puerto)
        servidor.starttls()
        servidor.login(remitente, contraseña)
        servidor.send_message(mensaje)
        print("Correo enviado con exito.")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")
    finally:
        servidor.quit()

def flujo_principal():
    """
    Ejecuta el flujo completo.
    """
    archivos_a_limpiar = [
        #este flujo genera 2 excel uno sin ordenar y el otro ordenado, este parte borra esos 2 excel antes de ejecutar el flujo para evitar errores
        r"ruta/donde/esten/tus/archivos",
        r"ruta/donde/esten/tus/archivos"
    ]
    limpiar_archivos(archivos_a_limpiar)

    url_login = "https://sitioWeb.com/login"
    url_archivo = "https://sitioWeb.com/reporte"
    ruta_destino = r""#este seria el archivo que descargo y esta para ser filtrado y ordenado
    archivo_salida = r""#este seria el archivo final. (filtrado y ordenado)

    #aqui van las credenciales para logear en el sitio web al que quieres entrar a descargar el archivo
    usuario = ""
    clave = ""

    driver = abrir_navegador(url_login)
    try:
        iniciar_sesion(driver, usuario, clave)
        descargar_archivo(driver, url_archivo, ruta_destino)
    except Exception as e:
        print(f"Error en el flujo: {e}")

    filtrar_bomberos(ruta_destino, archivo_salida)

    enviar_correo(
        # destinatario="test@gmail.com",
        destinatario="email@ente.org",#aqui va el correo a donde quieres enviar el archivo
        asunto="Reporte",
        cuerpo_base="Adjunto reporte de ...",
        archivo_adjunto=archivo_salida
    )

if __name__ == "__main__":
    flujo_principal()
