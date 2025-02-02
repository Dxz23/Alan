import pywhatkit
import time
import os
import pandas as pd
import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox
import random
import logging
import glob
import zipfile
import rarfile
import threading
from tkinter import ttk
import webbrowser

# Importaciones para validación de números con Selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# PyMuPDF para conversión de PDF a PNG
import fitz

# -------------------------------------------------------------------------
# Configuración de logging
# -------------------------------------------------------------------------
logging.basicConfig(filename='registro_envio.log', level=logging.DEBUG, format='%(asctime)s - %(message)s')

# -------------------------------------------------------------------------
# Variables globales para reportes
# -------------------------------------------------------------------------
reporte_exitos = []
reporte_errores = []
reporte_invalidos = []
reporte_verificacion = []  # Reporte para validación de números

# -------------------------------------------------------------------------
# Redefinir webbrowser.open para bloquear ciertas URLs
# -------------------------------------------------------------------------
original_webbrowser_open = webbrowser.open
def custom_webbrowser_open(url, *args, **kwargs):
    if "flaticon.es/icono-gratis/validar_5165349" in url:
        print(f"Bloqueada apertura de URL: {url}")
        return False
    return original_webbrowser_open(url, *args, **kwargs)
webbrowser.open = custom_webbrowser_open

# -------------------------------------------------------------------------
# 1) Función para cargar archivo Excel
# -------------------------------------------------------------------------
def cargar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        archivo_excel.set(archivo)
        status_label.config(text="Archivo cargado con éxito")

# -------------------------------------------------------------------------
# 2) Función para enviar mensaje de texto con pywhatkit
# -------------------------------------------------------------------------
def enviar_texto_pywhatkit(telefono, texto):
    try:
        pywhatkit.sendwhatmsg_instantly(telefono, texto, wait_time=10, tab_close=True, close_time=4)
        time.sleep(15)  # Espera a que se envíe el mensaje
        print(f"Mensaje enviado a {telefono}.")
        return True
    except Exception as e:
        logging.error(f"Error al enviar texto a {telefono}: {e}")
        return False

# -------------------------------------------------------------------------
# 3) Función para enviar imagen con pywhatkit (con reintentos)
# -------------------------------------------------------------------------
def press_esc_after_delay(delay):
    time.sleep(delay)
    pyautogui.press('esc')
    print("Tecla ESC presionada.")

def enviar_imagen_pywhatkit(telefono, ruta_imagen, caption="Aquí tienes tu izzi card 📸📄"):
    try:
        esc_thread = threading.Thread(target=press_esc_after_delay, args=(8,))
        esc_thread.start()
        pywhatkit.sendwhats_image(telefono, ruta_imagen, caption, wait_time=15, tab_close=True, close_time=4)
        time.sleep(7)
        print(f"Imagen enviada a {telefono}.")
        return True
    except Exception as e:
        logging.error(f"Error al enviar imagen a {telefono}: {e}")
        return False

# -------------------------------------------------------------------------
# 4) Función para enviar mensajes en lotes (incluye rpt, mes y estado "Inactivo")
# -------------------------------------------------------------------------
def enviar_mensajes():
    archivo = archivo_excel.get()
    if not archivo:
        messagebox.showerror("Error", "Por favor, carga un archivo Excel.")
        return
    try:
        df = pd.read_excel(archivo, dtype=str)
        print("Archivo Excel cargado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo Excel: {e}")
        return

    # Definir columnas posibles (renombrado a nombres estándar)
    posibles_columnas = {
        'telefono': ['telefono', 'telefono 1', 'celular 1', 'celular', 'Celular '],
        'telefono2': ['telefono 2', 'celular 2', 'Celular 2'],
        'nombre': ['nombre', 'Nombre '],
        'numero_cuenta': ['Num. de contrato', 'No de cuenta', 'número de cuenta'],
        'dinero': ['dinero', 'saldo', 'monto'],
        'fecha': ['fecha', 'fecha límite'],
        'estado': ['Estado'],
        'rpt': ['rpt'],
        'mes': ['mes']
    }
    columnas_renombradas = {}
    for columna_estandar, alternativas in posibles_columnas.items():
        for alternativa in alternativas:
            if alternativa in df.columns:
                columnas_renombradas[alternativa] = columna_estandar
                break
    df.rename(columns=columnas_renombradas, inplace=True)

    # Verificar que existan las columnas esenciales
    columnas_faltantes = []
    for col in ['telefono', 'nombre', 'numero_cuenta', 'dinero', 'fecha']:
        if col not in df.columns:
            columnas_faltantes.append(col)
    if columnas_faltantes:
        messagebox.showerror("Error", f"Faltan columnas en el archivo: {columnas_faltantes}")
        return

    # Definir mensajes base
    mensaje_base = (
        "Hola {nombre} 😊,\n"
        "Nos comunicamos de izzi telecom 📡 y te invitamos a realizar el pago correspondiente de tu mensualidad.\n"
        "Tu número de cuenta es {numero_cuenta}. Lo que tienes que pagar es ${dinero} 💸💳.\n"
        "Tu último día para pagar es el día {fecha}.\n\n"
        "🛑Recuerda que puedes pagarlo en cualquier tienda de autoservicio.\n\n"
        "¡Gracias por ser parte de nuestra familia IZZI! 🎉"
    )
    mensaje_saldo_0 = (
        "👋Hola {nombre},\n"
        "¡Bienvenido a IZZI! 📡🌐\n"
        "Felicidades por completar el proceso de instalación.\n"
        "Le presentamos su número de cuenta {numero_cuenta}.\n\n"
        "Además, lo invitamos a descargar la izziapp📱 donde podrá administrar su servicio y realizar pagos desde su celular."
    )
    mensaje_descuento = (
        "🤩Además puedes recibir un descuento adicional de $50.00 MXN *DE POR VIDA*.\n\n"
        "💳Al domiciliar tu pago de tus servicios izzi 🌐\n\n"
        "Si deseas domiciliar marca al 📲8188808821\n\n"
        "¡Gracias por ser parte de nuestra familia IZZI! 🎉😊"
    )
    mensaje_inactivo = (
        "👋✨ Hola {nombre}\n"
        "Te recordamos que tu servicio de IZZI está vencido, te invitamos a pagar tu servicio para volver a reactivar. 🔥\n\n"
        "💸 Monto a pagar: {dinero}\n"
        "⏰ Fecha límite: {fecha}\n\n"
        "📡 Evita interrupciones y sigue disfrutando de tu servicio. Paga ahora y mantente conectado. 🌐\n\n"
        "Para más información o realizar tu pago, ingresa a nuestra izzi app 🖥️. O marcando al 800 120 5000.📲\n\n"
    )
    carpeta_imagenes = r"C:\Users\OCELOT\Desktop\IZZI_CARDS"
    lote_tamano = 100
    total_mensajes = len(df)
    mensajes_enviados_total = 0

    for inicio in range(0, total_mensajes, lote_tamano):
        df_lote = df[inicio:inicio + lote_tamano]
        for idx, row in df_lote.iterrows():
            try:
                # Extraer datos básicos y adicionales
                telefono1 = str(row['telefono']).strip() if pd.notna(row['telefono']) else ""
                telefono2 = ""
                if 'telefono2' in df.columns:
                    telefono2 = str(row['telefono2']).strip() if pd.notna(row['telefono2']) else ""
                nombre = row['nombre'].strip() if pd.notna(row['nombre']) else "Cliente"
                numero_cuenta = row['numero_cuenta'].strip() if pd.notna(row['numero_cuenta']) else "Sin especificar"
                dinero = row['dinero'].strip() if pd.notna(row['dinero']) else "0"
                fecha_dia = row['fecha'].strip() if pd.notna(row['fecha']) else "Sin especificar"
                estado_cliente = ""
                if 'estado' in df.columns and pd.notna(row['estado']):
                    estado_cliente = row['estado'].strip()
                rpt_val = row['rpt'].strip() if ('rpt' in df.columns and pd.notna(row['rpt'])) else ""
                mes_val = row['mes'].strip() if ('mes' in df.columns and pd.notna(row['mes'])) else ""

                def obtener_mensaje_principal():
                    if estado_cliente == "Inactivo":
                        return mensaje_inactivo.format(nombre=nombre, dinero=dinero, fecha=fecha_dia)
                    elif dinero.isdigit() and int(dinero) == 0:
                        return mensaje_saldo_0.format(nombre=nombre, numero_cuenta=numero_cuenta)
                    else:
                        return mensaje_base.format(nombre=nombre, numero_cuenta=numero_cuenta, dinero=dinero, fecha=fecha_dia)

                # Buscar imagen asociada al número de cuenta
                patrones = os.path.join(carpeta_imagenes, f"*{numero_cuenta}*.png")
                imagenes_encontradas = glob.glob(patrones)
                ruta_imagen = imagenes_encontradas[0] if imagenes_encontradas else None

                # Función interna para enviar mensajes a un número
                def enviar_todo_al_telefono(tel):
                    if not tel:
                        logging.warning(f"Fila {idx + 1}: Teléfono vacío o no válido. Se omite.")
                        reporte_invalidos.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono: {tel} - Teléfono vacío/no válido.")
                        return
                    max_reintentos = 3
                    mensaje = obtener_mensaje_principal()
                    exito_main = False
                    exito_img = False
                    exito_desc = False
                    # Enviar mensaje principal
                    for intento in range(max_reintentos):
                        exito_main = enviar_texto_pywhatkit(tel, mensaje)
                        if exito_main:
                            print(f"Mensaje principal enviado a {tel} en intento {intento + 1}.")
                            break
                        else:
                            logging.warning(f"Intento {intento + 1} fallido para mensaje principal a {tel}.")
                            time.sleep(5)
                    else:
                        logging.error(f"No se pudo enviar mensaje principal a {tel} tras {max_reintentos} intentos.")
                        reporte_errores.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono: {tel} - No se pudo enviar el mensaje principal.")
                    # Enviar imagen si existe
                    if ruta_imagen:
                        for intento in range(max_reintentos):
                            exito_img = enviar_imagen_pywhatkit(tel, ruta_imagen, "Aquí tienes tu izzi card 📸📄")
                            if exito_img:
                                print(f"Imagen enviada a {tel} en intento {intento + 1}.")
                                break
                            else:
                                logging.warning(f"Intento {intento + 1} fallido para imagen a {tel}.")
                                time.sleep(5)
                        else:
                            logging.error(f"No se pudo enviar imagen a {tel} tras {max_reintentos} intentos.")
                            reporte_errores.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono: {tel} - No se pudo enviar la imagen.")
                    else:
                        print(f"No se encontró imagen para la cuenta {numero_cuenta}.")
                    # Enviar mensaje de descuento
                    for intento in range(max_reintentos):
                        exito_desc = enviar_texto_pywhatkit(tel, mensaje_descuento)
                        if exito_desc:
                            print(f"Mensaje de descuento enviado a {tel} en intento {intento + 1}.")
                            break
                        else:
                            logging.warning(f"Intento {intento + 1} fallido para mensaje de descuento a {tel}.")
                            time.sleep(5)
                    else:
                        logging.error(f"No se pudo enviar mensaje de descuento a {tel} tras {max_reintentos} intentos.")
                        reporte_errores.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono: {tel} - No se pudo enviar el mensaje de descuento.")
                    # Reportar éxito o fallo
                    if exito_main and (exito_img or not ruta_imagen) and exito_desc:
                        logging.debug(f"Todos los mensajes enviados correctamente a: {tel}")
                        reporte_exitos.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono: {tel} - OK")
                    else:
                        fallos = []
                        if not exito_main:
                            fallos.append("Mensaje principal")
                        if ruta_imagen and not exito_img:
                            fallos.append("Izzi Card (imagen)")
                        if not exito_desc:
                            fallos.append("Mensaje de descuento")
                        reporte_invalidos.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono: {tel} - Falló(n): {', '.join(fallos)}")
                    time.sleep(random.uniform(8, 15))

                # Enviar a cada teléfono (si existen)
                if telefono1:
                    enviar_todo_al_telefono(telefono1)
                if telefono2:
                    enviar_todo_al_telefono(telefono2)

                mensajes_enviados_total += 1
                porcentaje = (mensajes_enviados_total / total_mensajes) * 100
                progress_var.set(min(porcentaje, 100))
                progress_label.config(text=f"{porcentaje:.2f}%")
                ventana.update()

            except Exception as e:
                logging.error(f"Error en fila {idx + 1}: {e}")
                reporte_errores.append(f"Fila {idx + 1} - Nombre: {nombre}, Número de Cuenta: {numero_cuenta}, rpt: {rpt_val}, mes: {mes_val}, Teléfono(s): {telefono1}/{telefono2} - {e}")
                actualizar_estado(f"Error en fila {idx + 1}: {e}")

        if (inicio + lote_tamano) < total_mensajes:
            actualizar_estado("Pausando 5 minutos antes del siguiente lote...")
            time.sleep(300)

    actualizar_estado("Proceso completado.")
    messagebox.showinfo("Proceso completado", "Se han enviado todos los mensajes del archivo Excel.")

# -------------------------------------------------------------------------
# 5) Función para generar reporte final (TXT)
# -------------------------------------------------------------------------
def generar_reporte():
    archivo_reporte = filedialog.asksaveasfilename(defaultextension=".txt",
                                                   filetypes=[("Text files", "*.txt")],
                                                   title="Guardar reporte como")
    if archivo_reporte:
        with open(archivo_reporte, "w", encoding="utf-8") as reporte:
            reporte.write("Mensajes enviados con éxito:\n")
            for exito in reporte_exitos:
                reporte.write(exito + "\n")
            reporte.write("\nErrores al enviar mensajes:\n")
            for error in reporte_errores:
                reporte.write(error + "\n")
            reporte.write("\nMensajes inválidos:\n")
            for invalido in reporte_invalidos:
                reporte.write(invalido + "\n")
            reporte.write("\nVerificación de Números:\n")
            for verificacion in reporte_verificacion:
                reporte.write(verificacion + "\n")
        messagebox.showinfo("Reporte generado", "El reporte ha sido generado con éxito.")

# -------------------------------------------------------------------------
# 6) Función para convertir PDF a PNG (y eliminar PDF)
# -------------------------------------------------------------------------
def pdf_to_png(input_folder, output_folder, total_pdfs):
    # Se asegura que la carpeta de salida exista
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    # Se recorre recursivamente el directorio de entrada
    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            if filename.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, filename)
                try:
                    pdf_document = fitz.open(pdf_path)
                except Exception as e:
                    logging.error(f"Error al abrir el PDF {pdf_path}: {e}")
                    continue
                for page_num in range(len(pdf_document)):
                    page = pdf_document.load_page(page_num)
                    pix = page.get_pixmap()
                    # Se guarda la imagen en la carpeta de salida
                    output_file = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}_page{page_num + 1}.png")
                    pix.save(output_file)
                    print(f"Guardado: {output_file}")
                pdf_document.close()
                os.remove(pdf_path)
                print(f"Eliminado el archivo PDF: {pdf_path}")

# -------------------------------------------------------------------------
# 7) Función para extraer PDFs de ZIP o RAR y convertirlos a PNG
# -------------------------------------------------------------------------
def extraer_y_convertir():
    archivo_comprimido = filedialog.askopenfilename(filetypes=[("Archivos ZIP", "*.zip"), ("Archivos RAR", "*.rar")])
    if not archivo_comprimido:
        return
    carpeta_destino = filedialog.askdirectory(title="Selecciona una carpeta de destino para los PNGs")
    if not carpeta_destino:
        return
    actualizar_estado("Extrayendo y convirtiendo PDFs...")
    total_pdfs = 0
    if archivo_comprimido.endswith(".zip"):
        with zipfile.ZipFile(archivo_comprimido, 'r') as zip_ref:
            archivos_pdf = [f for f in zip_ref.namelist() if f.endswith('.pdf')]
            total_pdfs = len(archivos_pdf)
            for archivo in archivos_pdf:
                actualizar_estado(f"Extrayendo {archivo}...")
                pdf_path = os.path.join(carpeta_destino, archivo)
                # Se crea la estructura de directorios intermedios si no existe
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
                with zip_ref.open(archivo) as file:
                    with open(pdf_path, 'wb') as output_file:
                        output_file.write(file.read())
                # Se convierte el PDF (se buscarán de forma recursiva)
                pdf_to_png(carpeta_destino, carpeta_destino, total_pdfs)
    elif archivo_comprimido.endswith(".rar"):
        with rarfile.RarFile(archivo_comprimido, 'r') as rar_ref:
            archivos_pdf = [f for f in rar_ref.namelist() if f.endswith('.pdf')]
            total_pdfs = len(archivos_pdf)
            for archivo in archivos_pdf:
                actualizar_estado(f"Extrayendo {archivo}...")
                pdf_path = os.path.join(carpeta_destino, archivo)
                # Se crea la estructura de directorios intermedios si no existe
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
                with rar_ref.open(archivo) as file:
                    with open(pdf_path, 'wb') as output_file:
                        output_file.write(file.read())
                pdf_to_png(carpeta_destino, carpeta_destino, total_pdfs)
    actualizar_estado("Extracción y conversión completadas.")
    messagebox.showinfo("Proceso completado", "Archivos PDF extraídos y convertidos a PNG con éxito.")

# -------------------------------------------------------------------------
# 8) Funciones para ejecutar tareas en hilos
# -------------------------------------------------------------------------
def convertir_en_hilo():
    hilo = threading.Thread(target=extraer_y_convertir)
    hilo.start()

def enviar_mensajes_en_hilo():
    hilo = threading.Thread(target=enviar_mensajes)
    hilo.start()

# -------------------------------------------------------------------------
# 9) Funciones auxiliares de interfaz
# -------------------------------------------------------------------------
def actualizar_estado(texto):
    estado_label.after(0, lambda: estado_label.config(text=texto))

def salir():
    ventana.destroy()

# -------------------------------------------------------------------------
# 10) Funciones de validación de números usando Selenium
# -------------------------------------------------------------------------
def iniciar_whatsapp_con_perfil():
    chrome_profile_path = r"C:\Users\OCELOT\Desktop\selenium_whatsapp_profile"
    if not os.path.exists(chrome_profile_path):
        os.mkdir(chrome_profile_path)
    options = Options()
    options.add_argument(f"--user-data-dir={chrome_profile_path}")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    driver.get("https://web.whatsapp.com/")
    print("Esperando 20 segundos para cargar/escaneo QR.")
    time.sleep(20)
    return driver

def validar_un_telefono(driver, telefono, fila_info):
    if not telefono:
        logging.info(f"{fila_info} - Teléfono vacío, omitiendo...")
        return "Vacío"
    for intento in range(3):
        try:
            print(f"Verificando número: {telefono} (Intento {intento + 1})")
            driver.get(f"https://web.whatsapp.com/send?phone={telefono}")
            wait = WebDriverWait(driver, 30)
            try:
                wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div[contenteditable='true'][aria-placeholder='Escribe un mensaje']")
                ))
                print(f"Número válido: {telefono}")
                return "Válido"
            except:
                mensaje_error = "El número de teléfono compartido a través de la dirección URL no es válido"
                if driver.find_elements(By.XPATH, f"//*[contains(text(),'{mensaje_error}')]"):
                    print(f"Número inválido: {telefono}")
                    return "Inválido"
                else:
                    raise Exception("No se pudo determinar el estado del número.")
        except WebDriverException as e:
            print(f"Error en intento {intento + 1} para {telefono}: {e}")
            time.sleep(5)
        except Exception as e:
            print(f"Intento {intento + 1} fallido para {telefono}: {e}")
            time.sleep(5)
    return "Indeterminado"

def verificar_numeros(driver, archivo_entrada, archivo_salida):
    try:
        posibles_columnas = {
            'telefono': ['telefono', 'telefono 1', 'celular 1', 'celular', 'Celular '],
            'telefono2': ['telefono 2', 'celular 2', 'Celular 2'],
            'nombre': ['nombre', 'Nombre '],
            'numero_cuenta': ['Num. de contrato', 'No de cuenta', 'número de cuenta'],
            'dinero': ['dinero', 'saldo', 'monto'],
            'fecha': ['fecha', 'fecha límite'],
            'rpt': ['rpt'],
            'mes': ['mes']
        }
        df = pd.read_excel(archivo_entrada, dtype=str)
        columnas_renombradas = {}
        for col_estandar, alternativas in posibles_columnas.items():
            for alt in alternativas:
                if alt in df.columns:
                    columnas_renombradas[alt] = col_estandar
                    break
        df.rename(columns=columnas_renombradas, inplace=True)

        def obtener_valor(row, columna):
            return row[columna].strip() if pd.notna(row.get(columna, "")) else ""

        resultado = []
        total_numeros = len(df)

        for i, row in df.iterrows():
            nombre = obtener_valor(row, 'nombre')
            numero_cuenta = obtener_valor(row, 'numero_cuenta')
            rpt_val = obtener_valor(row, 'rpt')
            mes_val = obtener_valor(row, 'mes')
            dinero = obtener_valor(row, 'dinero')
            fecha_val = obtener_valor(row, 'fecha')

            telefono1 = obtener_valor(row, 'telefono')
            telefono2 = obtener_valor(row, 'telefono2')

            if telefono1:
                estado_telefono1 = validar_un_telefono(driver, telefono1, f"Fila {i+1}")
                reporte_verificacion.append(f"{nombre}\t{numero_cuenta}\t{rpt_val}\t{mes_val}\t{dinero}\t{fecha_val}\t{telefono1} {estado_telefono1}")
            if telefono2:
                estado_telefono2 = validar_un_telefono(driver, telefono2, f"Fila {i+1}")
                reporte_verificacion.append(f"{nombre}\t{numero_cuenta}\t{rpt_val}\t{mes_val}\t{dinero}\t{fecha_val}\t{telefono2} {estado_telefono2}")

            resultado.append({
                "fila": i+1,
                "nombre": nombre,
                "numero_cuenta": numero_cuenta,
                "rpt": rpt_val,
                "mes": mes_val,
                "dinero": dinero,
                "fecha": fecha_val,
                "telefono1": telefono1,
                "estado_telefono1": estado_telefono1 if telefono1 else "",
                "telefono2": telefono2,
                "estado_telefono2": estado_telefono2 if telefono2 else ""
            })

            porcentaje = ((i + 1) / total_numeros) * 100
            actualizar_progreso_verificacion(porcentaje)

        df_resultado = pd.DataFrame(resultado)
        df_resultado.to_excel(archivo_salida, index=False)
        print(f"Reporte generado: {archivo_salida}")
        messagebox.showinfo("Proceso completado", f"La verificación ha finalizado. Revisa el reporte generado: {archivo_salida}")

    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"Error durante la verificación: {e}")

def validar_numeros():
    archivo_entrada = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")], 
                                                 title="Selecciona el archivo Excel con los números de teléfono")
    if not archivo_entrada:
        return
    archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                  filetypes=[("Archivos Excel", "*.xlsx")], 
                                                  title="Selecciona la ubicación para guardar el reporte")
    if not archivo_salida:
        return
    actualizar_estado("Iniciando validación de números...")
    progress_var_verificacion.set(0)
    progress_label_verificacion.config(text="0.00%")
    ventana.update()

    def proceso_validacion():
        driver = iniciar_whatsapp_con_perfil()
        verificar_numeros(driver, archivo_entrada, archivo_salida)
        driver.quit()
        actualizar_estado("Validación completada.")

    hilo = threading.Thread(target=proceso_validacion)
    hilo.start()

def actualizar_progreso_verificacion(porcentaje):
    porcentaje = min(porcentaje, 100)
    progress_var_verificacion.set(porcentaje)
    progress_label_verificacion.config(text=f"{porcentaje:.2f}%")
    ventana.update()

# -------------------------------------------------------------------------
# 11) Función para verificar respuestas y envío de mensajes
# -------------------------------------------------------------------------
def verificar_respuestas_y_envio():
    archivo_entrada = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")], 
                                                 title="Selecciona el archivo Excel con los números validados")
    if not archivo_entrada:
        return
    archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                  filetypes=[("Archivos Excel", "*.xlsx")], 
                                                  title="Selecciona la ubicación para guardar el reporte de respuestas y envío")
    if not archivo_salida:
        return
    actualizar_estado("Iniciando verificación de respuestas y envío...")
    progress_var_verificacion.set(0)
    progress_label_verificacion.config(text="0.00%")
    ventana.update()

    def proceso_verificacion_respuestas_y_envio():
        try:
            driver = iniciar_whatsapp_con_perfil()
            df = pd.read_excel(archivo_entrada, dtype=str)
            if 'telefono' not in df.columns:
                raise ValueError("El archivo debe contener una columna llamada 'telefono'.")
            resultados = []
            for i, row in df.iterrows():
                telefono = row['telefono'].strip() if pd.notna(row['telefono']) else ""
                nombre = row.get('nombre', '').strip()
                if not telefono:
                    resultados.append({
                        "telefono": telefono,
                        "nombre": nombre,
                        "respondio": "Vacío",
                        "mensaje": "",
                        "envio_correcto": "No"
                    })
                    continue
                driver.get(f"https://web.whatsapp.com/send?phone={telefono}")
                time.sleep(10)
                try:
                    mensajes_entrantes = driver.find_elements(By.XPATH, "//div[contains(@class,'message-in')]//span[@dir='ltr']")
                    if mensajes_entrantes:
                        ult_respuesta = mensajes_entrantes[-1].text
                        respondio = "Sí"
                        mensaje_respuesta = ult_respuesta
                    else:
                        respondio = "No"
                        mensaje_respuesta = ""
                    mensajes_salientes = driver.find_elements(By.XPATH, "//div[contains(@class,'message-out')]//span[@dir='ltr']")
                    envio_correcto = "Sí" if mensajes_salientes else "No"
                    if respondio == "Sí":
                        print(f"[{telefono}] El cliente SÍ contestó. Último mensaje: {mensaje_respuesta}")
                    else:
                        print(f"[{telefono}] El cliente NO contestó.")
                    if envio_correcto == "Sí":
                        print(f"[{telefono}] Se envió al menos un mensaje.")
                    else:
                        print(f"[{telefono}] No se envió ningún mensaje saliente.")
                    resultados.append({
                        "telefono": telefono,
                        "nombre": nombre,
                        "respondio": respondio,
                        "mensaje": mensaje_respuesta,
                        "envio_correcto": envio_correcto
                    })
                except Exception as e:
                    resultados.append({
                        "telefono": telefono,
                        "nombre": nombre,
                        "respondio": "Error",
                        "mensaje": str(e),
                        "envio_correcto": "Error"
                    })
                    print(f"Error al verificar el número {telefono}: {e}")
                porcentaje = ((i + 1) / len(df)) * 100
                progress_var_verificacion.set(porcentaje)
                progress_label_verificacion.config(text=f"{porcentaje:.2f}%")
                ventana.update()
            df_resultados = pd.DataFrame(resultados)
            df_resultados.to_excel(archivo_salida, index=False)
            print(f"Reporte de respuestas y envío generado: {archivo_salida}")
            messagebox.showinfo("Proceso completado", f"Se ha generado el reporte de respuestas y envío: {archivo_salida}")
        except Exception as e:
            print(f"Error en la verificación de respuestas y envío: {e}")
            messagebox.showerror("Error", f"Error en la verificación de respuestas y envío: {e}")
        finally:
            driver.quit()
            actualizar_estado("Verificación de respuestas y envío completada.")
    hilo = threading.Thread(target=proceso_verificacion_respuestas_y_envio)
    hilo.start()

# -------------------------------------------------------------------------
# 12) Construcción de la interfaz gráfica
# -------------------------------------------------------------------------
ventana = tk.Tk()
ventana.title("Automatización de Envío de Mensajes")
ventana.geometry("700x800")

archivo_excel = tk.StringVar()
frame = tk.Frame(ventana)
frame.pack(pady=20)

verificar_respuestas_envio_btn = tk.Button(ventana, text="Verificar Respuestas y Envío", bg="orange", fg="white", command=verificar_respuestas_y_envio)
verificar_respuestas_envio_btn.pack(pady=10)

# Cargar íconos (si están disponibles)
try:
    icono_cargar = tk.PhotoImage(file="icono_cargar.png")
    icono_extraer = tk.PhotoImage(file="icono_extraer.png")
    icono_enviar = tk.PhotoImage(file="icono_enviar.png")
    icono_reporte = tk.PhotoImage(file="icono_reporte.png")
    icono_salir = tk.PhotoImage(file="icono_salir.png")
except Exception as e:
    logging.error(f"Error al cargar los íconos: {e}")
    messagebox.showerror("Error", f"No se pudieron cargar los íconos: {e}")

cargar_btn = tk.Button(frame, text="Cargar Archivo Excel", image=icono_cargar, compound="left", command=cargar_archivo)
cargar_btn.grid(row=0, column=0, padx=10)
extraer_btn = tk.Button(frame, text="Extraer y Convertir PDFs", image=icono_extraer, compound="left", command=convertir_en_hilo)
extraer_btn.grid(row=0, column=1, padx=10)
validar_btn = tk.Button(frame, text="Validar Números", command=validar_numeros)
validar_btn.grid(row=0, column=2, padx=10)

enviar_btn = tk.Button(ventana, text="Enviar Mensajes", image=icono_enviar, compound="left", command=enviar_mensajes_en_hilo, bg="green", fg="white")
enviar_btn.pack(pady=10)
reporte_btn = tk.Button(ventana, text="Generar Reporte", image=icono_reporte, compound="left", command=generar_reporte, bg="blue", fg="white")
reporte_btn.pack(pady=10)
salir_btn = tk.Button(ventana, text="Salir", image=icono_salir, compound="left", command=salir, bg="red", fg="white")
salir_btn.pack(pady=10)

progress_var = tk.DoubleVar(value=0)
progress_bar = ttk.Progressbar(ventana, variable=progress_var, maximum=100)
progress_bar.pack(pady=10, fill="x", padx=20)
progress_label = tk.Label(ventana, text="0.00%", font=("Arial", 10))
progress_label.pack(pady=5)
progress_var_verificacion = tk.DoubleVar(value=0)
progress_bar_verificacion = ttk.Progressbar(ventana, variable=progress_var_verificacion, maximum=100)
progress_bar_verificacion.pack(pady=10, fill="x", padx=20)
progress_label_verificacion = tk.Label(ventana, text="0.00%", font=("Arial", 10))
progress_label_verificacion.pack(pady=5)

status_label = tk.Label(ventana, text="Cargar archivo Excel para comenzar", font=("Arial", 12))
status_label.pack(pady=5)
estado_label = tk.Label(ventana, text="Estado: Sin iniciar", font=("Arial", 12))
estado_label.pack(pady=10)

ventana.mainloop()
