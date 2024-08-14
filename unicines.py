from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd

browser = webdriver.Firefox()

URL = 'https://unicines.com/'

browser.get(URL)

peli = []


def agregar_datos(nombre_cine, titulo, idioma, formato, horarios):

    datos = {
        "Fecha": date.today(),
        "País": "Honduras",
        "Cine": "Unicines",
        "Nombre de cine": nombre_cine,
        "Titulo": titulo,
        "Idioma": idioma,
        "Formato": formato,
        "Hora": horarios
    }

    peli.append(datos)


browser.get(URL)

cartelera = browser.find_elements(
    By.XPATH, '/html/body/header/div[1]/div/div[1]/div/div/ul/li[3]/ul/li/a')

links = [link.get_attribute("href") for link in cartelera]
# print(len(links))

for link in links:
    browser.get(link)
    nombre_cine = browser.find_element(
        By.XPATH, '/html/body/main/div[2]/div[2]/div/h1').text[13:]
    print(nombre_cine)

    peliculas = browser.find_elements(
        By.XPATH, '//div[contains(@id,"collapse")]/div/*')
    print(len(peliculas))

    for pelicula in peliculas:

        titulo = pelicula.find_element(
            By.XPATH, './div/div[3]/div/h3').text
        print(titulo)

        funciones = pelicula.find_element(
            By.XPATH, './div/div[4]/div/div/span')

        funciones = funciones.text.split('\n')

        print(len(funciones))

        for funcion in funciones:

            formato_idiomas, horarios = funcion.rsplit(' - ')
            formato_idiomas = formato_idiomas[4:]
            formato, idioma = formato_idiomas.rsplit(' ', 1)
            print(horarios)
            print(formato)
            print(idioma)

            agregar_datos(nombre_cine,titulo,idioma,formato,horarios)

browser.quit

df = pd.DataFrame(peli)

df.to_excel('unicines - '+str(date.today())+'.xlsx', index=False)
