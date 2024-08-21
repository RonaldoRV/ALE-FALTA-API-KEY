import json
from difflib import get_close_matches
import re
import openpyxl
from fuzzywuzzy import fuzz
import unicodedata
import pygame
import threading
import tkinter as tk
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
import math
import sys
import time
import pandas as pd
from unicodedata import normalize
from fuzzywuzzy import fuzz
import subprocess


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


# Asignar la clave de la API de OpenAI
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

def cosine_similarity(str1, str2):
    # Tokenización de las strings
    tokens1 = str1.split()
    str2_as_string = ' '.join(str2)
    tokens2 = str2_as_string.split()

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
    cosine_similarity = dot_product / (magnitude1 * magnitude2)

    return cosine_similarity

def compare_strings(str1, str2, threshold= None):

    if threshold == None:
        threshold=0.4
    # threshold "1" indica una alta similitud entre las cadenas, mientras que un valor cercano a 0 indica una baja similitud.
    #es mejor tenerla bajita para que tome mas facil el prompt
    similarity = cosine_similarity(str1, str2)
    if similarity >= threshold:
        return True
    else:
        return False


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
        prints('scrapped')
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

def speak_e(audio):
    engine = pyttsx3.init('sapi5')  # sapi 5 para que hable español
    voices = engine.getProperty('voices')
    engine.setProperty('voice', r'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0')
    engine.say(audio)
    engine.runAndWait()

def spk_print(message):
    # Crear dos hilos, uno para la función prints() y otro para la función speak_spanish()
    thread1 = threading.Thread(target=prints, args=(message,))
    thread2 = threading.Thread(target=speak, args=(message,))

    # Iniciar los hilos
    thread1.start()
    thread2.start()

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
        spk_print("Buenos dias ")

    elif hour >= 12 and hour < 18:
        spk_print("Buenas tardes!")

    else:
        spk_print("Buenas noches!")

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
            prints("Escuchando...")
            r.pause_threshold = 1
            audio = r.listen(source)

        try:
            prints("Reconociendo...")
            query = r.recognize_google(audio, language= language)
            prints(f"Has dicho: {query}\n")
            return query.strip() 

        except sr.UnknownValueError:
            spk_print("No pude escucharte. Por favor, repite tu respuesta.")
            error_count += 1  # Incrementa el contador de errores
            if error_count >= 4:
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
            if compare_strings(query, ['chao']):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break
            
            elif compare_strings(query, ['hasta pronto', 'luego']):
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

            # -------------------------------------------MATRICULAS-----------------------------------------------
            if compare_strings(query,['matricula','inscripcion']): 
                spk_print("Ya te digo como es el proceso")          
                url = 'https://unipaz.edu.co/index.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # CORREO BUSQUEDA EXCEL TOCA ARREGLARLO XD
            elif compare_strings(query, ['correo'], threshold= 0.2):
                cantidad_valores = contar_valores_columna_excel("correos.xlsx", "Nombre")
                print(cantidad_valores)

                indice_valor_columna = 0

                resultado = query.split("de", 1)[-1].strip().lower()
                print(resultado)
               
                for indice_valor_columna in range(cantidad_valores):
                    siguiente_valor, correo, _ = leer_excel_y_obtener_siguiente_valor("correos.xlsx", "Nombre", "Correo", indice_valor_columna)
                    print(siguiente_valor)
                    if compare_strings(resultado, siguiente_valor):
                        print("Valor de la columna:", indice_valor_columna, siguiente_valor)
                        print("correo correspondiente:", correo)
                        spk_print(correo)
                        break

            # -----------------------POSGRADOS------------------
            
            elif compare_strings(query, ['posgrados', 'especializaciones', 'maestrías']):
                speak("¿Cuál de las siguientes buscas: especializaciones tecnológicas, especializaciones profesionales o maestrías?")
                query2 = takeCommand('es-CO')

                if compare_strings(query2, ['especializaciones tecnológicas']):
                    # especializaciones tecnológicas
                    speak("Tenemos la siguiente especialización tecnológica: Especialización Tecnológica en Control de Calidad de Biocombustibles Líquidos - SNIES 104100 (ofertado en Barrancabermeja, Piedecuesta y San Pablo -Sur de Bolívar). Si quieres más información, pregúntame.")

                elif compare_strings(query2, ['especializaciones profesionales']):
                    # especializaciones profesionales
                    speak("""Tenemos las siguientes especializaciones profesionales:          
                        #1) Especialización en Agronegocios - 3SNIES 104724
                        #2) Especialización en Aseguramiento de la Calidad e Inocuidad Agroalimentaria - SNIES 102840
                        #3) Especialización en Gerencia de Proyectos Culturales- SNIES 105553
                        #4) Especialización Mercadeo Global Empresarial- SNIES
                        #5) Especialización Gestión Ambiental- SNIES""")
                        
                elif compare_strings(query2, ['maestrías']):
                    # maestrías
                    speak("""Tenemos las siguientes maestrías profesionales:          
                        #1) Maestría en Logística y Cadena de Suministro 
                        #Si quieres más información, pregúntame por el nombre de la maestría.""")
                else:
                    speak("No entendi tu peticion.")

                    
            # ESPECIALIZACION TECNOLOGICA CONTROL DE CALIDAD
            elif compare_strings(query, ['control', 'calidad biocombustibles', 'especialización']):        

                url = 'https://unipaz.edu.co/especializacion_tecnologica_biocombustibles.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ESPECIALIZACIONES PROFESIONALES
            # Especialización en Agronegocios - SNIES 104724
            elif compare_strings(query, ['especialización','agronegocios']):

                url = 'https://unipaz.edu.co/especializacion_agronegocios.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # Especialización en Aseguramiento de la Calidad e Inocuidad Agroalimentaria - SNIES 102840
            elif compare_strings(query, ['especialización', 'aseguramiento', 'calidad e inocuidad agroalimentaria']):
                
                url = 'https://unipaz.edu.co/especializacion_aseguramiento_calidad_inocuidad_agroalimentaria.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

            # Especialización en Gerencia de Proyectos Culturales- SNIES 105553
            elif compare_strings(query, ['especialización', 'gerencia proyectos culturales']):        
                
                url = 'https://unipaz.edu.co/especializacion_proyectos_culturales.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

            # Especialización Mercadeo Global Empresarial- SNIES
            elif compare_strings(query, ['especialización', 'mercadeo global empresarial']):

                url = 'https://unipaz.edu.co/especializacion_mercadeo_global_empresarial.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # Especialización Gestión Ambiental- SNIES
            elif compare_strings(query, ['especialización', 'gestión ambiental']): 
                
                url = 'https://unipaz.edu.co/especializacion_gestion_ambiental.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

            #  MAESTRIAS
            # Maestría en Lógistica y Cadena de Suministro
            elif compare_strings(query, ['maestría', 'lógistica cadena suministro']):

                url = 'https://unipaz.edu.co/maestria_logistica_cadena_suministro.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)  
            #--------------------END POSGRADOS-----------------
                
            # ------------------ ESTUDIANTES ------------------
            elif compare_strings(query, ['estudiantes', 'matriculados', 'existen', 'hay']):       
                
                url = 'https://unipaz.edu.co/pregrado.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 

            # ------------------ END ESTUDIANTES ------------------     
            # ------------------ UBICACION UNIVERSIDADES  ------------------    
            elif compare_strings(query, ['visitar', 'Llévame', 'quiero ir', 'universidad']):
                spk_print('¿A cuál universidad deseas ir?')
                university = takeCommand('es-CO')
                if compare_strings(university,['unipaz', 'instituto universitario de La Paz']):
                    spk_print("Entendido, ya te muestro la ubicación del Instituto Universitario de La Paz.")
                    webbrowser.open("https://shorturl.at/bwIP5")
                    time.sleep(2)
                    spk_print("El campus universitario de unipaz de barrancabermeja, queda en el Centro de investigación Santa Lucía, km 14 vía Bucaramanga, antes del peaje la lizama, girando a la izquierda en el retorno, luego a 500 metros giras a la derecha y llegarás al lugar de destino.")

                elif compare_strings(university, ['universidad industrial de santander']):
                    spk_print("Entendido, ya te muestro la ubicación de la universidad industrial de santander UIS.")
                    webbrowser.open("https://shorturl.at/enLO9")
                    time.sleep(2)


            elif compare_strings(query, ['llevame', 'unipaz', 'instituto universitario de la paz']):
                    spk_print("Entendido, ya te muestro la ubicación del  instituto universitario de La Paz")
                    webbrowser.open("https://shorturl.at/bwIP5")
                    time.sleep(2)
            # ------------------ FUNDADA, MISION, VISION E HISTORIA UNIPAZ ------------------ 
            # VISION
            elif compare_strings(query, ['visión']):       
                
                url = 'https://www.unipaz.edu.co/mision-vision-historia.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            # MISION
            elif compare_strings(query, ['misión']):       
                
                url = 'https://www.unipaz.edu.co/mision-vision-historia.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            # HISTORIA
            elif compare_strings(query, ['historia']):       
                
                url = 'https://www.unipaz.edu.co/mision-vision-historia.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

            elif compare_strings(query, ['fundada', 'fundacion', 'fundó', 'fundado']):       
                
                url = 'https://www.unipaz.edu.co/mision-vision-historia.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ------------------ END MISION, VISION E HISTORIA UNIPAZ ------------------ 
            # ------------------ PRINCIPIOS Y VALORES UNIPAZ------------------ 
            elif compare_strings(query, ['principios y valores']):       
                url = 'https://unipaz.edu.co/prinicipios_y_valores.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            # ------------------ ePRINCIPIOS Y VALORES UNIPAZ------------------ 
            # ------------------OBJETIVOS INSTITUCIONALES UNIPAZ ------------------ 
            elif compare_strings(query, ['objetivos']):       

                url = 'https://unipaz.edu.co/objetivos.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)    
                
            # ------------------MANUAL DE IDENTIDAD VISUAL ESCUDO Y BANDERA ------------------ 
            #identidad visual
            elif compare_strings(query, ['identidad visual']):       
                url = 'https://unipaz.edu.co/identidad-institucional.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 
            #bandera
            elif compare_strings(query, ['bandera']):       
                url = 'https://unipaz.edu.co/identidad-institucional.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 
            #simbolo
            elif compare_strings(query, ['simbolo']):       
                url = 'https://unipaz.edu.co/identidad-institucional.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 
                
            # ------------------ RECTORIA ------------------ 
            elif compare_strings(query, ['rector', 'nombre']):       
                url = 'https://unipaz.edu.co/rectoria.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 
            
            # ------------------ PROGRAMAS DE PREGADO ------------------
            elif compare_strings(query, ['carreras', 'pregrados', 'programas', 'profesionales']):       
                url = 'https://unipaz.edu.co/pregrado.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 
            
            # ESCUELA DE CIENCIAS 
            # Administración de Negocios Internacionales
            elif compare_strings(query, ['administración negocios internacionales']):       
                url = 'https://unipaz.edu.co/administracion_negocios_internacionales.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 

            # Ingeniería Informática
            elif compare_strings(query, ['ingeniería informática']):       
                url = 'https://unipaz.edu.co/ingenieria_informatica.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 

            #Licenciatura en Artes
            elif compare_strings(query, ['licenciatura en artes']):       
                url = 'https://unipaz.edu.co/licenciatura_en_artes.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text) 

            #Química 
            elif compare_strings(query, ['química']):       
                url = 'https://unipaz.edu.co/quimica.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ESCUELA DE CIENCIAS SOCIALES Y COMUNICACIONES
            #Comunicación Social
            elif compare_strings(query, ['comunicación social']):       
                url = 'https://unipaz.edu.co/comunicacion_social.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            #Trabajo Social
            elif compare_strings(query, ['trabajo social']):       
                url = 'https://unipaz.edu.co/trabajo_social.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

                        
            # ESCUELA DE INGENIERIA AGROINDUSTRIAL
            #Tecnico en extraccion de biomasa
            elif compare_strings(query, ['técnico', 'procesos extracción biomasa energética',]):       
                url = 'https://unipaz.edu.co/tecnico_extraccion_biomasa_energetica.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

            #Tecnico en Procesamiento de Alimentos
            elif compare_strings(query, ['técnico', 'procesamiento alimentos',]):       
                url = 'https://unipaz.edu.co/tecnologia_en_procesamiento_de_alimentos.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            #Ingenieria Agroindustrial
            elif compare_strings(query, ['ingeniería agroindustrial']):       
                url = 'https://unipaz.edu.co/ingenieria_agroindustrial.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ESCUELA INGENIERIA AGRONOMICA
            elif compare_strings(query, ['ingeniería agronómica']):       
                url = 'https://unipaz.edu.co/ingenieria_agronomica.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ESCUELA INGENIERIA AMBIENTAL
            elif compare_strings(query, ['ingeniería ambiental saneamiento']):       
                url = 'https://unipaz.edu.co/ingenieria_ambiental_y_saneamiento.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            #Tecnologia en obras civiles
            elif compare_strings(query, ['obras civiles', 'tecnología']):       
                url = 'https://unipaz.edu.co/tecnologia_obras_civiles.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ESCUELA DE INGENIERIA DE PRODUCCION
            #Tecnologia operación sistemas electromecánicos
            elif compare_strings(query, ['operación sistemas electromecánicos', 'tecnología']):       
                url = 'https://unipaz.edu.co/tecnico_oepracion_sistemas.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # sst
            elif compare_strings(query, ['seguridad y salud trabajo', 'tecnología']):       
                url = 'https://unipaz.edu.co/tecnologia_seguridad_salud_trabajo.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)

            # ing sst
            elif compare_strings(query, ['seguridad y salud trabajo', 'ingeniería']):       
                url = 'https://unipaz.edu.co/ingenieria_seguridad_salud_trabajo.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ing produccion
            elif compare_strings(query, ['producción', 'ingeniería']):       
                url = 'https://unipaz.edu.co/ingenieria_produccion.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
            
            # ESCUELA DE MVZ
            elif compare_strings(query, ['medicina veterinaria zootecnia mvz']):       
                url = 'https://unipaz.edu.co/medicina_veterinaria_y_zootecnia.html'
                response_text = web_scrape_and_respond(url, query, chatbot)
                spk_print(response_text)
                        

            # ------------------ CONTROL DEL BOT  ------------------ 
            # --ADIOS BOT ------------------ 
            elif compare_strings(query, ['chao']):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break
            
            elif compare_strings(query, ['hasta pronto', 'luego']):
                spk_print("Espero haberte ayudado. Hasta pronto ")
                break

            # --ESPERA BOT  ------------------ 
            elif compare_strings(query, ['espera']):
                spk_print("Vale")
                time.sleep(15)
                spk_print("Ya volvi")

            else:
                cantidad_valores = contar_valores_columna_excel('database.xlsx', 'keyword')
                encontrado = False

                for indice_valor_columna in range(cantidad_valores + 1):
                    siguiente_valor, enlace, _ = leer_excel_y_obtener_siguiente_valor(nombre_archivo, nombre_columna, nombre_columna_enlace, indice_valor_columna)
                    if siguiente_valor is not None and compare_strings(query, siguiente_valor, threshold= 0.3):
                        print("Valor de la columna:", indice_valor_columna, siguiente_valor)
                        print("Enlace correspondiente:", enlace)
                        url = enlace
                        response_text = web_scrape_and_respond(url, query, chatbot)
                        spk_print(response_text)
                        encontrado = True
                        break
                    
                if not encontrado:
                 spk_print("Recuerda que soy una inteligencia artificial que esta en proceso de desarrollo")
                

            # ------------------ CONTROL DEL BOT  ------------------ 

            #elif best_match:
                #Si se encuentra una coincidencia, obtiene la respuesta y la muestra al usuario
                #answer: str = get_answer_for_question(best_match, knowledge_base)
                #print(f'Bot: {answer}')
                #speak(answer)



# INTERFAZ DE LOS BOTONES DE ABAJO WE
button_container = tk.Frame(master=frame)
button_container.pack()

button1 = tk.Button(
    master=button_container, text="UNIPAZ PREGUNTAS", command=iniciar_comandos, font=("Raleway", 9), fg="white")
button1.configure(bg="blue")
button1.grid(row=0, column=0, padx=10, pady=4)

button1 = tk.Button(
    master=button_container, text="PREGUNTA GENERAL", command=iniciar_comandos2, font=("Raleway", 9), fg="white")
button1.configure(bg="blue")
button1.grid(row=0, column=1, padx=10, pady=4)

button2 = tk.Button(
    master=button_container, text="ENGLISH VERSION", command=iniciar_comandos_e, font=("Raleway", 9), fg="white")
button2.configure(bg="blue")
button2.grid(row=0, column=2, padx=10, pady=4)

button3 = tk.Button(
    master=button_container, text="TÉRMINOS Y CONDICIONES", command=open_image, font=("Raleway", 9), fg="white")
button3.configure(bg="blue")
button3.grid(row=0, column=3, padx=10, pady=4)

button_restart = tk.Button(
    master=button_container, text="REINICIAR PROGRAMA", command=restart_program, font=("Raleway", 9), fg="white")
button_restart.configure(bg="blue")
button_restart.grid(row=0, column=4, padx=10, pady=4)

button_exit = tk.Button(
    master=button_container, text="SALIR", command=exit_program, font=("Raleway", 9), fg="white")
button_exit.configure(bg="blue")
button_exit.grid(row=0, column=5, padx=10, pady=4)

button_pause = tk.Button(
    master=button_container, text="PAUSA", command= pause_program, font=("Raleway", 9), fg="white")
button_pause.configure(bg="blue")
button_pause.grid(row=0, column=6, padx=10, pady=4)


root.mainloop()