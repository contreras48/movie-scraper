from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd

browser = webdriver.Firefox()

URL = 'https://caribbeancinemas.com/location/panama/'

browser.get(URL)

peliculas = []

def fun(fecha, pais, cine, nombre_cine, titulo, idioma, formato, horarios):
    for h in horarios:
        datos = {
            "Fecha": fecha,
            "País": pais,
            "Cine": cine,
            "Nombre de cine": nombre_cine,
            "Titulo": titulo,
            "Idioma": idioma,
            "Formato": formato,
            "Hora": h.text
        }
        peliculas.append(datos)

teatros = browser.find_elements(
    By.XPATH, '//ul[@id="menu-main-menu"]//li[5]//ul[@class="sub-menu "]/li/a')
lista1 = [link.get_attribute('href') for link in teatros]

for i in lista1:
    browser.get(i)
    fecha = date.today()
    pais = 'Panama'
    cine = 'Caribbean Cinemas'
    nombre_cine = browser.find_element(
        By.XPATH, '//div[@id="cineinfo"]//b').text

    cartelera = browser.find_elements(By.XPATH, '//div[@id="horarios"]/*')

    for movie in cartelera:
        titulo = movie.find_element(By.XPATH, './/b')
        idioma = movie.find_element(By.XPATH, './/i').text
        if  "0" in idioma:
            indice = idioma.index("0")
            idioma=idioma[2:]if indice==0 else idioma[0:-2]
        print(idioma)
            
        formato="2D"        
        horario = movie.find_elements(
            By.XPATH, './/div [@class="column three-fourth"]//div[2]/a')

        fun(fecha, pais, cine, nombre_cine,
        titulo.text, idioma, formato, horario)

browser.quit()

df = pd.DataFrame(peliculas)
df.to_excel('horarios_caribean_cinemas_panama.xlsx', index=False)
