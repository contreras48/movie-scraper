from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd

# inicialización del navegador
browser = webdriver.Firefox()

# URL del cine del que se desea extraer la informacion
URL = 'https://caribbeancinemas.com/location/panama/'

# browser.get(URL) navega a esa URL en el navegador Firefox controlado por Selenium
browser.get(URL)

# Se inicializa una lista vacía llamada peliculas, que almacenará la información de cada película
peliculas = []

# funcion que toma varios parámetros como fecha, país, cine, nombre del cine, título de la película, idioma, formato y horarios
def fun(fecha, pais, cine, nombre_cine, titulo, idioma, formato, horarios):
# Se itera sobre cada hora en horarios y crea un diccionario datos que almacena esta información
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
# agrega el diccionario datos a la lista peliculas
        peliculas.append(datos)

# browser.find_elements() encuentra todos los elementos en la página que corresponden a los enlaces de los teatros (cines)
teatros = browser.find_elements(
    By.XPATH, '//ul[@id="menu-main-menu"]//li[5]//ul[@class="sub-menu "]/li/a')

# lista1 almacena los enlaces (href) de esos teatros para poder navegar a ellos posteriormente
lista1 = [link.get_attribute('href') for link in teatros]

# Se itera sobre cada enlace en lista1
for i in lista1:
    # browser.get(i) navega a la página del cine específico
    browser.get(i)
    fecha = date.today()
    pais = 'Panama'
    cine = 'Caribbean Cinemas'

# Se obtiene el nombre del cine de la página web
    nombre_cine = browser.find_element(
        By.XPATH, '//div[@id="cineinfo"]//b').text

# Se almacena todos los elementos que contienen la información de las películas y horarios disponibles en ese cine
    cartelera = browser.find_elements(By.XPATH, '//div[@id="horarios"]/*')
# Se itera sobre cada película en la cartelera
    for movie in cartelera:
        # Se extrae el título de la película y el idioma
        titulo = movie.find_element(By.XPATH, './/b')
        idioma = movie.find_element(By.XPATH, './/i').text

        # Se corrige el formato del idioma si contiene un "0" en algún lugar, eliminando caracteres no deseados
        if "0" in idioma:
            indice = idioma.index("0")
            idioma = idioma[2:]if indice == 0 else idioma[0:-2]
        formato = "2D"
        
# horario almacena todos los horarios disponibles para esa película
        horario = movie.find_elements(
            By.XPATH, './/div [@class="column three-fourth"]//div[2]/a')
        
# Se llama a la función fun pasando toda la información recolectada de la película actual
        fun(fecha, pais, cine, nombre_cine,
            titulo.text, idioma, formato, horario)

# Cerrar el navegador
browser.quit()

# Se convierte la lista peliculas en un DataFrame de pandas
df = pd.DataFrame(peliculas)

# Se exporta el DataFrame a un archivo Excel
df.to_excel('caribean_cinemas_panama.xlsx', index=False)
