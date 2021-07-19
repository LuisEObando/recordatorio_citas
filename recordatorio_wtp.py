from openpyxl import load_workbook #Importamos librería que deja cargar excel (solo .xlsx)
import time
import os #libreria para cerrar ventanas y .exe
from selenium import webdriver #Importamos selenium
from selenium.webdriver.common.keys import Keys  #Para realizar acciones por teclado
from openpyxl import load_workbook #Importamos librería que deja cargar excel (solo .xlsx)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait #Para hacer las esperas y validar condiciones
from selenium.webdriver.support import expected_conditions as EC #Condiciones, modulo de selenium
import pyttsx3
import pyautogui as pg, webbrowser as web


#Webdriver
driver = webdriver.Chrome(executable_path="C:/Users/luise/Downloads/chromedriver.exe") #Declaramos la función abrir navegador con webdriver
driver.get("https://web.whatsapp.com/")
print("Tiene 30 segundos para validar su sesión...")
time.sleep(30)

#Excel
excel = "./Base de datos.xlsx"
wb = load_workbook(excel)
hojas= wb.get_sheet_names() #Permite saber cuantas hojas tiene el excel 
print(hojas, "Hojas en el excel")
hoja1 = wb.get_sheet_by_name('Hoja1')
#mensaje = hoja1[f'B1'] #Lee el mensaje a escribir
#print (mensaje.value, "Es el Mensaje a enviar")
#ruta_adjunto = hoja2[f'B2'] #Lee la ruta del archivo a adjuntar
#print (ruta_adjunto.value, "Es la ruta del adjunto")

#Preparamos para el ciclo de Whatsapp
contador_excel = 2
ultima_fila = hoja1.max_row
ultima_fila = ultima_fila + 1
print(ultima_fila, "ultima fila")

for i in range((contador_excel), (ultima_fila)):
    telefono = hoja1[f'E{contador_excel}']
    print(telefono.value)
    if str(telefono.value) == 'None':
        print(telefono.value)
    else:
        
        driver.get("https://web.whatsapp.com/send?phone=+57" + str(telefono.value)) #Llamamos la función abrir navegador
        time.sleep(5)
        #leemos las base de datos
        mensaje_g = hoja1[f'G{contador_excel}']
        print (mensaje_g.value, "Es el saludo")
        fecha_h = hoja1[f'H{contador_excel}']
        print (fecha_h.value, "Es la fecha")
        hora_i = hoja1[f'I{contador_excel}']
        print(hora_i.value, "Es la hora")
        direccion_j = hoja1[f'J{contador_excel}']
        print(direccion_j.value, "es el lugar")
        profesional_k = hoja1[f'K{contador_excel}']
        print(profesional_k.value, "Medico")
        recomendaciones_l = hoja1[f'L{contador_excel}']
        print(recomendaciones_l.value, "son las recomendaciones")
        #Concatenamos mensaje
        mensaje = str(mensaje_g.value +" el día " +  fecha_h.value +" A las "+ hora_i.value + " En la " + direccion_j.value +" Con el Médico "+ profesional_k.value +" le recordamos "+ recomendaciones_l.value)
        print(mensaje)

        #Escribimos y enviamos mensaje de texto
        driver.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[1]/div/div[2]").send_keys(mensaje) #Escribimos el mensaje!
        driver.find_element_by_xpath("/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[2]/div/div[2]/button").click() #click en enviar


        #Convertimos mensaje a mp3
        engine = pyttsx3.init()
        engine.setProperty("rate",150)
        telefono = str(telefono.value)
        audio_guardado = "Audio"+(telefono)+".mp3"
        engine.save_to_file(mensaje, audio_guardado)
        #engine.say(mensaje) #reproduce
        engine.runAndWait()

        #Adjuntamos y enviamos el audio
        driver.find_element_by_xpath("/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[1]/div[2]/div/div/span").click()#click en adjunto
        time.sleep(2)
        driver.find_element_by_xpath("/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[1]/div[2]/div/span/div[1]/div/ul/li[1]/button/span").click() #imgicon
        time.sleep(1)
    
        audio_guardado = "\Audio"+(telefono)+".mp3"
        ruta_audio = (r"C:\Users\luise" + str(audio_guardado))
        pg.write(ruta_audio)
        time.sleep(1)
        pg.press('enter') #damos enter en cuadro de dialogo 
        time.sleep(1)
        driver.find_element_by_xpath("/html/body/div/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/span/div/div/span").click()#enviar adj

        time.sleep(8) #esperamos que salga el mensaje y cargue el nuevo número

    contador_excel = contador_excel + 1
        #
        # web.open("https://web.whatsapp.com/send?phone=+" + str(telefono.value))
        #Celdas de excel de hora y fecha deben estar en formato texto