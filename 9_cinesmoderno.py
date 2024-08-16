from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd
from selenium.common.exceptions import NoSuchElementException

browser = webdriver.Firefox()

URL = 'https://www.cinesmodernopanama.com/'

browser.get(URL)

peli = []

def agregar_datos(nombre_cine, titulo, idioma, formato, horarios):
    for hora in horarios:
        datos = {
            "Fecha": date.today(),
            "País": "Panama",
            "Cine": "Cines Moderno",
            "Nombre de cine": nombre_cine,
            "Titulo": titulo,
            "Idioma": idioma,
            "Formato": formato,
            "Hora": hora
        }

        peli.append(datos)


browser.get(URL)

cartelera = browser.find_elements(
    By.XPATH, '//*[@id="cd-lateral-nav"]/ul[1]/li/ul/li/a')

links = [link.get_attribute("href") for link in cartelera]

for link in links:

    browser.get(link)

    try:

        nombre_cine = browser.find_element(
            By.XPATH, '//*[@id="cartelera"]/div/div[1]/div/p[2]').text[2:]

        peliculas = browser.find_elements(
            By.XPATH, '//*[@id="cartelera"]/div/div[5]/*')

        for pelicula in peliculas:
            titulo = pelicula.find_element(
                By.XPATH, './/div[@class="combopelititulo"]/h2').text
            idioma = pelicula.find_element(
                By.XPATH, './/div[@class="combodetallepeli"]//div[2]//div[@class="detallespeli"]//div[1]/p').text
            horarios = pelicula.find_elements(
                By.XPATH, './/div[@class="combodetallepeli"]//div[2]//div[@class="detallespeli"]//div[7]/ul/li/a')

            lista_horarios = []

            horas = len(horarios)

            if horas > 0:
                for horario in horarios:
                    hora = horario.text[0:6]
                    formato = horario.text[7:-1]
                    lista_horarios.append(hora)
            else:
                hora = " "
                formato = horario.text[7:-1]
                lista_horarios.append(hora)

            agregar_datos(nombre_cine, titulo, idioma, formato, lista_horarios)
    except NoSuchElementException: 
        pass

browser.quit

df = pd.DataFrame(peli)

df.to_excel('cinesmoderno - '+str(date.today())+'.xlsx', index=False)
