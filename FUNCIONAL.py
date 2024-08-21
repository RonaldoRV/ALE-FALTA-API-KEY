import json
from difflib import get_close_matches
import re
import openpyxl
from fuzzywuzzy import fuzz
from ctypes import cast, POINTER
import sounddevice as sd
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import random
import unicodedata
import pygame
import threading
from tkinter import ttk
import tkinter as tk
import subprocess
from moviepy.editor import *
from PIL import Image, ImageTk
import threading
from threading import Thread
from email.mime import image
from tkinter import *
from turtle import left
from PIL import ImageTk, Image, ImageDraw  # para colocar imagenes en el code
import pyttsx3
import speech_recognition as sr
import datetime
import webbrowser
import time
import win32com.client as wincl
import openai
from urllib.request import urlopen
import os  # reproducir el codigo desde otro lugar
import requests
from bs4 import BeautifulSoup
from difflib import get_close_matches
from openai import OpenAI
from collections import Counter
import requests
from PyPDF2 import PdfReader
from io import BytesIO
import math
import sys
import time
import pandas as pd
from unicodedata import normalize
from fuzzywuzzy import fuzz
import subprocess
import threading
from tkinter import Label
import pyttsx3

# Inicializar pygame
pygame.init()

# Cargar los sonidos MP3
#sound1 = pygame.mixer.Sound("sounds/escuchar.mp3")
#sound2 = pygame.mixer.Sound("tone2.mp3")

# asignamos la clave de open ai api
openai.api_key = ""

root = tk.Tk()
root.title("ALE")  # set the window title
# Obtener dimensiones de la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Establecer la geometría de la ventana a pantalla completa
root.geometry(f"{screen_width}x{screen_height}")
root.attributes('-fullscreen', True)

# Cargar y redimensionar la imagen de fondo
bg_image = Image.open("img/alereal.png")
bg_image = bg_image.resize((screen_width, screen_height))
background_tk = ImageTk.PhotoImage(bg_image)

# Crear y colocar la etiqueta de fondo
background_label = tk.Label(root, image=background_tk)
background_label.place(x=0, y=0, relwidth=1.0, relheight=1.0)

# Crear un marco para otros elementos (si es necesario)
frame = tk.Frame(master=root)
frame.pack(pady=0, side="bottom", fill="x")

# Crear etiqueta de bienvenida
label = tk.Label(master=frame, text="BIENVENIDO A SU ASISTENTE ALE", font=("Raleway", 17))
label.pack(pady=4, padx=4)

# Carga tu icono
root.iconbitmap("img/logo.ico")  # Reemplaza "ruta_del_archivo.ico" con la ubicación de tu archivo .ico

#OPEN Terminal

# Example usage


# Asignar la clave de la API de OpenAI ESTA TOCA ACTUALIZARLA PPORQUE NO SIRVE XD 
openai.api_key = "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"

# Definir función para abrir una imagen
def open_image():
    # Carga la imagen externa y muestra en una ventana emergente
    image = Image.open("img/terms.png")  # Reemplaza "ruta_de_la_imagen.jpg" con la ubicación de tu imagen
    image.show()

# Establecer atributos de la ventana para permitir minimizar y cerrar
root.attributes("-fullscreen", False)  # Salir del modo de pantalla completa

# asignamos la clave de open ai api
openai.api_key = "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"

def open_image():
    # Carga la imagen externa y muestra en una ventana emergente
    image = Image.open("img/terms.png")  # Reemplaza "ruta_de_la_imagen.jpg" con la ubicación de tu imagen
    image.show()


class ChatBot:
    def __init__(self, api_key):
        self.api_key = api_key
        self.client = OpenAI(api_key=self.api_key)

    def get_response(self, prompt):
        message = {
            "role": "user",
            "content": prompt,
        }
        completion = self.client.chat.completions.create(messages=[message], model="gpt-3.5-turbo")
        
        # Extract response from completion object
        response = None
        if completion.choices:
            response = completion.choices[0].message.content
        
        return response



def compare_strings_l(str1, str2, threshold=0.2):
    # Asegúrate de que ambos sean cadenas
    if not isinstance(str1, str):
        str1 = ' '.join(str1)
    if not isinstance(str2, str):
        str2 = ' '.join(str2)

    # Tokenización de las strings
    tokens1 = str1.split()
    tokens2 = str2.split()

    # Conteo de ocurrencias de palabras
    vector1 = Counter(tokens1)
    vector2 = Counter(tokens2)

    # Intersección de las palabras en ambos strings
    intersection = set(vector1.keys()) & set(vector2.keys())

    # Producto punto de los vectores
    dot_product = sum(vector1[x] * vector2[x] for x in intersection)

    # Magnitud de los vectores
    magnitude1 = math.sqrt(sum(vector1[x] ** 2 for x in vector1.keys()))
    magnitude2 = math.sqrt(sum(vector2[x] ** 2 for x in vector2.keys()))

    # Cálculo de la similitud coseno
    if magnitude1 == 0 or magnitude2 == 0:
        return False
    cosine_similarity = dot_product / (magnitude1 * magnitude2)

    # Imprimir la similitud para depuración
    print(f"Comparando '{str1}' con '{str2}' - Similitud: {cosine_similarity}")

    return cosine_similarity >= threshold


def cosine_similarity(str1, str2):
    tokens1 = str1.split()
    tokens2 = str2.split()

    vector1 = Counter(tokens1)
    vector2 = Counter(tokens2)

    intersection = set(vector1.keys()) & set(vector2.keys())

    dot_product = sum(vector1[x] * vector2[x] for x in intersection)

    magnitude1 = math.sqrt(sum(vector1[x] ** 2 for x in vector1.keys()))
    magnitude2 = math.sqrt(sum(vector2[x] ** 2 for x in vector2.keys()))

    if magnitude1 == 0 or magnitude2 == 0:
        return 0
    else:
        cosine_similarity = dot_product / (magnitude1 * magnitude2)
        return cosine_similarity
    
def compare_strings(str1, str2, threshold=None):
    if threshold is None:
        threshold = 0.4

    similarity = cosine_similarity(str1, str2)

    if similarity >= threshold:
        return True
    else:
        return False
    
def read_pdf_from_url(pdf_url):
    # Descargar el archivo PDF
    response = requests.get(pdf_url)
    
    if response.status_code == 200:
        # Cargar el contenido en un buffer para leerlo como PDF
        pdf_file = BytesIO(response.content)
        
        # Crear el lector de PDF
        pdf_reader = PdfReader(pdf_file)
        
        # Leer y almacenar el texto de cada página
        pdf_text = ""
        for page in pdf_reader.pages:
            pdf_text += page.extract_text()
        
        return pdf_text
    else:
        raise Exception("Error al descargar el PDF.")



def restart_program():
    python = sys.executable
    os.execl(python, python, *sys.argv)

def exit_program():
    exit()

def on_button_click2():
    exit_program()

def pause_program():
    spk_print("Ya vengo")
    time.sleep(15)
    spk_print("Ya volvi")



def web_scrape_and_respond(url, search_string, chatbot):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        # Supongamos que los títulos son etiquetas <h1> y <h2> y los párrafos son <p>
        titles = [title.text.strip() for title in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])]
        paragraphs = [p.text.strip() for p in soup.find_all('p')]
        # Unimos los títulos y párrafos en un solo texto
        extracted_text = '\n'.join(titles + paragraphs)
        search_text = search_string + extracted_text
        prints('Bien hecho')
        response_text = chatbot.get_response(search_text)
        return response_text
    else:
        return "No se pudo acceder al sitio web"

# Correos
# Leer el archivo Excel
def buscar_correo_por_nombre(nombre):
    # Carga el archivo 'nuevo_archivo.xlsx' en un DataFrame de Pandas
    df = pd.read_excel('correos.xlsx')

    max_ratio = -1
    correo_electronico = None
    for index, row in df.iterrows():
        ratio = fuzz.partial_ratio(nombre, row['Nombre'])
        if ratio > max_ratio:
            max_ratio = ratio
            correo_electronico = row['Correo']
    if max_ratio >= 60:
        return correo_electronico
    else:
        return "No se encontró un correo electrónico cercano para este nombre."

# Función que encapsula el proceso de búsqueda y devuelve el correo electrónico
def obtener_correo(nombre):
    correo_electronico = buscar_correo_por_nombre(nombre)
    return correo_electronico

# ENCONTRAR NOMBRES
def encontrar_nombres(texto):
    # Patrón de expresión regular para buscar nombres propios
    patron = r'\b[A-Z][a-z]*\b'
    
    # Buscar coincidencias de nombres en el texto
    nombres_encontrados = re.findall(patron, texto)
    
    return nombres_encontrados

# hacer la misma monda de excel pero para buscar palabras claves y obtener el link para buscarlo
def buscar_url_por_keyword(keyword):
    # Carga el archivo 'nuevo_archivo.xlsx' en un DataFrame de Pandas
    df = pd.read_excel('database.xlsx')

    max_ratio = -1
    url = None
    for index, row in df.iterrows():
        ratio = fuzz.partial_ratio(keyword, row['keyword'])
        if ratio > max_ratio:
            max_ratio = ratio
            url = row['url']
    if max_ratio >= 60:
        return url
    else:
        return "No se encontró un url cercano para este keyword."

# Función que encapsula el proceso de búsqueda y devuelve el correo electrónico
def obtener_url(keyword):
    url = buscar_url_por_keyword(keyword)
    return url

# ENCONTRAR NOMBRES
def encontrar_nombres(texto):
    # Patrón de expresión regular para buscar nombres propios
    patron = r'\b[A-Z][a-z]*\b'
    
    # Buscar coincidencias de nombres en el texto
    nombres_encontrados = re.findall(patron, texto)
    
    return nombres_encontrados


def obtener_valor_mas_similar(query, archivo_excel, nombre_hoja):
    """
    Función para obtener el valor más similar a la entrada de 'query' en la columna 'keyword' del archivo Excel,
    junto con el enlace asociado.
    
    :param query: Entrada de texto para comparar.
    :param archivo_excel: Ruta del archivo Excel.
    :param nombre_hoja: Nombre de la hoja en el archivo Excel.
    :return: Tupla que contiene el valor más similar y el enlace asociado.
    """
    wb = openpyxl.load_workbook(archivo_excel)
    hoja = wb[nombre_hoja]
    
    # Obtener valores de la columna 'keyword' y enlaces asociados
    valores_columna = [(celda.value, celda.offset(column=1).value) for celda in hoja['A']]  
    similitudes = [(valor, enlace, fuzz.ratio(query, valor)) for valor, enlace in valores_columna]  # Calcular similitud con distancia de Levenshtein
    similitudes.sort(key=lambda x: x[2], reverse=True)  # Ordenar por similitud descendente
    wb.close()
    return similitudes[0][:2]  # Devolver el valor y enlace con mayor similitud


#leer excel y keywords busqueda
indice_valor_columna = 0


def leer_excel_y_obtener_siguiente_valor(nombre_archivo, nombre_columna, nombre_columna_enlace, indice_valor_columna):
    # Lee el archivo Excel
    df = pd.read_excel(nombre_archivo, header=0)  # header=0 especifica que la primera fila es el encabezado
    
    # Verifica si la columna especificada existe en el DataFrame
    if nombre_columna in df.columns and nombre_columna_enlace in df.columns:
        # Almacena los valores de la columna especificada en una lista
        valores_columna = list(df[nombre_columna])
        enlaces_columna = list(df[nombre_columna_enlace])
        
        # Verifica si hay valores disponibles en la columna
        if indice_valor_columna < len(valores_columna):
            # Obtiene el siguiente valor de la columna y su enlace correspondiente
            siguiente_valor = valores_columna[indice_valor_columna]
            enlace = enlaces_columna[indice_valor_columna]
            indice_valor_columna += 1
            cadena = siguiente_valor
            lista = cadena.split(", ")
            lista_con_corchetes = [f"{elemento}" for elemento in lista]  # Solo el valor sin enlace
            return lista_con_corchetes, enlace, indice_valor_columna
        else:
            #print("Se han agotado los valores en la columna especificada.")
            # Devuelve un valor predeterminado en caso de que se agoten los valores
            return None, None, None  # O puedes devolver cualquier otro valor que desees
    else:
        print(f"La columna '{nombre_columna}' o '{nombre_columna_enlace}' no existen en el archivo Excel proporcionado.")
        # Devuelve un valor predeterminado en caso de que la columna no exista
        return None, None, None

# Ejemplo de uso
nombre_archivo = "database.xlsx"  # Reemplaza con el nombre de tu archivo Excel
nombre_columna = "keyword"
nombre_columna_enlace = "url"
indice_valor_columna = 0

#Leer excel y extraer tamano
def contar_valores_columna_excel(nombre_archivo, nombre_columna):
    try:
        # Lee el archivo Excel
        df = pd.read_excel(nombre_archivo)
        
        # Verifica si la columna especificada existe en el DataFrame
        if nombre_columna in df.columns:
            # Cuenta la cantidad de valores en la columna
            cantidad_valores = df[nombre_columna].count()
            return cantidad_valores
        else:
            print(f"La columna '{nombre_columna}' no existe en el archivo Excel proporcionado.")
            return None
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return None


# voz de Alee
def speak(audio):
    engine = pyttsx3.init('sapi5')  # sapi 5 para que hable español
    voices = engine.getProperty('voices')
    engine.setProperty('voice', r'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-MX_SABINA_11.0')
    engine.say(audio)
    engine.runAndWait()


# Función para actualizar el label en la GUI
def update_label(message):
    label.config(text=message)

# Función para actualizar el label de estado en la GUI
def update_status(status):
    status_label.config(text=status)

# Función para actualizar lo que dice el usuario
def user_say(user_say):
    status_label.config(text=user_say)


def speak_e(audio):
    engine = pyttsx3.init('sapi5')  # sapi 5 para que hable español
    voices = engine.getProperty('voices')
    engine.setProperty('voice', r'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0')
    engine.say(audio)
    engine.runAndWait()

def spk_print(message):
    # Crear dos hilos, uno para la función prints() y otro para la función speak()
    thread1 = threading.Thread(target=prints, args=(message,))
    thread2 = threading.Thread(target=speak, args=(message,))

    # Iniciar los hilos
    thread1.start()
    thread2.start()

    # Actualizar el label en el hilo principal
    root.after(0, update_label, message)

    # Esperar a que ambos hilos terminen
    thread1.join()
    thread2.join()

def spk_print_e(message):
    # Crear dos hilos, uno para la función prints() y otro para la función speak_spanish()
    thread1 = threading.Thread(target=prints, args=(message,))
    thread2 = threading.Thread(target=speak_e, args=(message,))

    # Iniciar los hilos
    thread1.start()
    thread2.start()

    # Esperar a que ambos hilos terminen
    thread1.join()
    thread2.join()


def prints(message, delay=0.05):
    for char in message:
        sys.stdout.write(char)
        sys.stdout.flush()
        time.sleep(delay)
    # Agregar un salto de línea al final del mensaje
    sys.stdout.write('\n')
    sys.stdout.flush()
# VOZ ALE FIN

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        spk_print("¡Buenos dias!")

    elif hour >= 12 and hour < 18:
        spk_print("¡Buenas tardes!")

    else:
        spk_print("¡Buenas noches!")

def usrname():
    uname = None
    while uname is None:
        spk_print("¿Cómo debo llamarte?")
        uname = takeCommand('es-CO')
        if uname and uname.strip():
            spk_print(f"Bienvenido, {uname}.")
        else:
            spk_print("Lo siento, no entendí tu nombre o no dijiste nada. Por favor, dime tu nombre nuevamente.")
    return uname

def q_open_web(url):
    spk_print("¿Deseas abrir la pagina? Responde con 'si' o 'no'")
    query2: str = takeCommand('es-CO')

    if query2 == 'sí':
        webbrowser.open(url)
    else:
        root.after(0, label, "Escuchando...")

def q_open_web_xd(url):
   
    def on_key_press(event):
        if event.char.lower() == 'y':
            webbrowser.open(url)
            label.config(text="SI")

        elif event.char.lower() == 'n':
            label.config(text="NO")

    spk_print("¿Deseas abrir la pagina? presiona la tecla Y o N ('Y'para sí o 'N' para no)")
    root.after(0, update_status, "PRESIONA LA TECLA...")
    root.bind("<Key>", on_key_press)



# borrar registros chat
def clear(): 
    return os.system('cls')

# funcion para recibir informacion del microfono y pasarlo a texto
def takeCommand(language):
    r = sr.Recognizer()
    error_count = 0  # Contador de errores

    while error_count < 4:  # Intenta hasta 4 veces
        with sr.Microphone() as source:
           #sound1.play()
            time.sleep(1)
            root.after(0, update_status, "Escuchando...")
            r.pause_threshold = 1
            audio = r.listen(source)

        try:
            root.after(0, update_status, "...")
            query = r.recognize_google(audio, language=language)
            root.after(0, user_say.config, {"text": f"Has dicho: {query}"})
            return query.strip() 

        except sr.UnknownValueError:
            root.after(0, update_status, "No pude escucharte. Por favor, repite tu respuesta.")
            spk_print("No pude escucharte. Por favor, repite tu respuesta.")
            error_count += 1  # Incrementa el contador de errores
            if error_count >= 4:
                root.after(0, update_status, "No pude escucharte. Por favor, repite tu respuesta.")
                spk_print("Parece que estás ocupado. Adiós.")
                exit()

        except Exception as e:
            prints(e)
            spk_print("Ocurrió un error. Por favor, inténtalo de nuevo.")
            error_count += 1  # Incrementa el contador de errores
            if error_count >= 4:
                spk_print("Parece que estás ocupado. Adiós.")
                exit()

    # Si se supera el límite de errores, ejecuta un speak y sale del programa
    spk_print("Parece que estás ocupado. Adiós.")
    exit()
# Función para cargar la base de conocimientos desde un archivo JSON
def load_knowledge_base(file_path: str) -> dict:
    with open(file_path, 'r') as file:
        data: dict = json.load(file)
    return data

def unescape_text(text):
    return unicodedata.normalize('NFKD', text).encode('latin-1', 'ignore').decode('utf-8')


# Función para guardar la base de conocimientos en un archivo JSON
def save_knowledge_base(file_path: str, data: dict):
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=2)

# Función para encontrar la mejor coincidencia entre la pregunta del usuario y las preguntas existentes
def find_best_match(user_question: str, question: list[str]) -> str | None:
    matches: list = get_close_matches(user_question, question, n=1, cutoff=0.6)
    return matches[0] if matches else None

# Función para obtener la respuesta a una pregunta de la base de conocimientos
def get_answer_for_question(question: str, knowledge_base: dict) -> str | None:
    for q in knowledge_base["questions"]:
        if q["question"] == question:
            return q["answer"]

def iniciar_comandos():
    # Inicia el bucle en un hilo separado
    threading.Thread(target=comandos, daemon=True).start()

# Función principal del chatbot
def machine_learning():
    if __name__ == '__main__':
        spk_print('Haz activado el modo entrenamiento!')
        # Carga la base de conocimientos desde 'knowledge_base.json'
        knowledge_base: dict = load_knowledge_base('knowledge_base.json')

        while True:
            # El usuario ingresa una pregunta
            query: str = takeCommand('es-CO').lower()

            if query.lower() == 'chao':
                quit()  # Si el usuario escribe "quit", se sale del bucle y termina el programa

            # Busca la mejor coincidencia en las preguntas existentes
            best_match: str | None = find_best_match(query, [q["question"] for q in knowledge_base["questions"]])

            if any(keyword in query for keyword in ('espera un momento', 'espera', 'pausa un momento', 'pausa')):
                spk_print("Vale")
                time.sleep(15)
                spk_print("Ya volvi")

            elif best_match:
                # Si se encuentra una coincidencia, obtiene la respuesta y la muestra al usuario
                answer: str = get_answer_for_question(best_match, knowledge_base)
                prints(f'Bot: {answer}')
                spk_print(answer)

            elif query.lower() == 'chao':
                quit()  # Si el usuario escribe "quit", se sale del bucle y termina el programa
            else:
                # Si no se encuentra una coincidencia, solicita al usuario que proporcione una respuesta
                prints("No conozco la respuesta. ¿Puedes enseñármela?")
                spk_print("No conozco la respuesta. ¿Puedes enseñármela?")
                new_answer: str = takeCommand('es-CO').lower()


                if new_answer.lower() != 'skip':
                    # Agrega la nueva pregunta y respuesta a la base de conocimientos
                    knowledge_base["questions"].append({"question": query, "answer": new_answer})
                    # Guarda la base de conocimientos actualizada en 'knowledge_base.json'
                    save_knowledge_base('knowledge_base.json', knowledge_base)
                    prints('Bot: Gracias, aprendí una nueva respuesta')
                    spk_print('Gracias, aprendí una nueva respuesta')

def iniciar_comandos_e():
    # Inicia el bucle en un hilo separado
    threading.Thread(target=english_question, daemon=True).start()

def english_question():
    # get ready the english voice and motor 
    if __name__ == '__main__':
        chatbot = ChatBot("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")
        clear()
        spk_print_e("Welcome to my english version. How can I assist you today?")

        while True:
            # El usuario ingresa una pregunta
            query: str = takeCommand('en-US')
        
            if compare_strings(query, ['hablame', 'ingles', 'traduce']):
                    spk_print("Espero haberte ayudado. Hasta pronto ")
                    bot_response = chatbot.get_response("respondeme en ingles" + query.lower())
                    spk_print_e(bot_response)

            else:
                    bot_response = chatbot.get_response(query.lower())
                    spk_print_e(bot_response)


def g_question():
    if __name__ == '__main__':
        chatbot = ChatBot("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")
        clear()
        wishMe()
        usrname()
        spk_print("¿En qué puedo asistirte con respecto a asuntos generales?")

        while True:
            # El usuario ingresa una pregunta
            query: str = takeCommand('es-CO')
           
            # --ADIOS BOT ------------------ 
            if compare_strings(query, 'chao'):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break
            
            elif compare_strings(query, 'hasta pronto luego'):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break
            

           
            else:
                 bot_response = chatbot.get_response(query.lower())
                 spk_print(bot_response)

def iniciar_comandos2():
    threading.Thread(target= g_question, daemon=True).start()

def comandos():
    if __name__ == '__main__':
        chatbot = ChatBot("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")
        
        clear()
        wishMe()
        usrname()
      

        spk_print("¿En qué puedo asistirte con respecto a asuntos universitarios?")

        # Carga la base de conocimientos desde 'knowledge_base.json'
        knowledge_base: dict = load_knowledge_base('knowledge_base.json')

        while True:
            # El usuario ingresa una pregunta
            query: str = takeCommand('es-CO')
           
            # Busca la mejor coincidencia en las preguntas existentes
            best_match: str | None = find_best_match(query, [q["question"] for q in knowledge_base["questions"]])

            valor_mas_similar, enlace_asociado = obtener_valor_mas_similar(query, 'database.xlsx', 'Hoja1')

            # CORREO BUSQUEDA EXCEL TOCA ARREGLARLO XD
            if compare_strings(query, 'correo', threshold= 0.2):
                cantidad_valores = contar_valores_columna_excel("correos.xlsx", "Nombre")
                print(cantidad_valores)

                indice_valor_columna = 0

                resultado = query.split("de", 1)[-1].strip().lower()
               
                for indice_valor_columna in range(cantidad_valores):
                    siguiente_valor, correo, _ = leer_excel_y_obtener_siguiente_valor("correos.xlsx", "Nombre", "Correo", indice_valor_columna)
                    print(siguiente_valor)
                    if compare_strings_l(resultado, siguiente_valor):
                        print("Valor de la columna:", indice_valor_columna, siguiente_valor)
                        print("correo correspondiente:", correo)
                        spk_print(correo)
                        break
            
             # Prompts
            elif compare_strings(query, 'hola'):
                saludo = ["Hola, soy Ale y estoy dispuesta a ayudarte en asuntos universitarios.","Que tal, espero estes bien, que asunto universitario deseas saber hoy?","Tiempo sin verte, en que te puedo ayudar?"]
                spk_print(random.choice(saludo))
                break

            # ------------------ CONTROL DEL BOT  ------------------ 
            # --ADIOS BOT ------------------ 
            elif compare_strings(query, 'chao'):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break
            
            elif compare_strings(query, 'hasta pronto luego'):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break

            # --ESPERA BOT  ------------------ 
            elif compare_strings(query, 'espera'):
                spk_print("Vale")
                time.sleep(15)
                spk_print("Ya volvi")

            else:
                cantidad_valores = contar_valores_columna_excel('database.xlsx', 'keyword')
                encontrado = False
                url= ""

                # Ejemplo de uso en el flujo principal
                for indice_valor_columna in range(cantidad_valores + 1):
                    siguiente_valor, enlace, _ = leer_excel_y_obtener_siguiente_valor(nombre_archivo, nombre_columna, nombre_columna_enlace, indice_valor_columna)
                    if siguiente_valor is not None:
                        if compare_strings_l(query, siguiente_valor, threshold=0.2):
                            print("Valor de la columna:", indice_valor_columna, siguiente_valor)
                            print("Enlace correspondiente:", enlace)
                            url = enlace
                            response_text = web_scrape_and_respond(url, query, chatbot)
                            spk_print(response_text)
                            encontrado = True
                            break
                
                if encontrado:
                    q_open_web_xd(url)
                    time.sleep(2)
  
                if not encontrado:
                 spk_print("Recuerda que soy una inteligencia artificial que esta en proceso de desarrollo")
                

            # ------------------ CONTROL DEL BOT  ------------------ 

            #elif best_match:
                #Si se encuentra una coincidencia, obtiene la respuesta y la muestra al usuario
                #answer: str = get_answer_for_question(best_match, knowledge_base)
                #print(f'Bot: {answer}')
                #speak(answer)

# LABELES
# Configuración del widget label
label = Label(root, font=("Helvetica", 18), justify="left", bg="#273475", fg="white", wraplength=1200, anchor="s")  # Ajusta el valor de wraplength según sea necesario
label.pack(pady=(20, 10))  # Añadir un poco de padding para la separación
label.place(relx=0.5, rely=0.7, anchor="center")

# Configuración del widget status_label
status_label = Label(root, font=("Helvetica", 10), justify="left", wraplength=400 )  # Ajusta el valor de wraplength según sea necesario
status_label.pack(pady=(20, 10))
status_label.place(relx=0.5, rely=0.755, anchor="center")

# Configuración del widget user_say
user_say = Label(root, font=("Helvetica", 10), justify="left", bg="#00963F", fg="white", wraplength=400)  # Ajusta el valor de wraplength según sea necesario
user_say.pack(pady=(20, 10))
user_say.place(relx=0.5, rely=0.8, anchor="center")


# INTERFAZ DE LOS BOTONES DE ABAJO WE
# Crear un contenedor para los botones de abajo
button_container = tk.Frame(master=frame)
button_container.pack()

button1 = tk.Button(
    master=button_container, text="UNIPAZ PREGUNTAS", command=iniciar_comandos, font=("Raleway", 9), fg="white")
button1.configure(bg="#273475")
button1.grid(row=0, column=0, padx=10, pady=4)

button2 = tk.Button(
    master=button_container, text="PREGUNTA GENERAL", command=iniciar_comandos2, font=("Raleway", 9), fg="white")
button2.configure(bg="#273475")
button2.grid(row=0, column=1, padx=10, pady=4)

button3 = tk.Button(
    master=button_container, text="ENGLISH VERSION", command=iniciar_comandos_e, font=("Raleway", 9), fg="white")
button3.configure(bg="#273475")
button3.grid(row=0, column=2, padx=10, pady=4)

button4 = tk.Button(
    master=button_container, text="TÉRMINOS Y CONDICIONES", command=open_image, font=("Raleway", 9), fg="white")
button4.configure(bg="#273475")
button4.grid(row=0, column=3, padx=10, pady=4)

button_restart = tk.Button(
    master=button_container, text="REINICIAR PROGRAMA", command=restart_program, font=("Raleway", 9), fg="white")
button_restart.configure(bg="#273475")
button_restart.grid(row=0, column=4, padx=10, pady=4)

button_exit = tk.Button(
    master=button_container, text="SALIR", command=exit_program, font=("Raleway", 9), fg="white")
button_exit.configure(bg="#273475")
button_exit.grid(row=0, column=5, padx=10, pady=4)

button_pause = tk.Button(
    master=button_container, text="PAUSA", command= pause_program, font=("Raleway", 9), fg="white")
button_pause.configure(bg="#273475")
button_pause.grid(row=0, column=6, padx=10, pady=4)


root.mainloop()