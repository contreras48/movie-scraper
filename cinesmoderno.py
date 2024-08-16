from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd
from selenium.common.exceptions import NoSuchElementException

# inicialización del navegador
browser = webdriver.Firefox()

# URL del cine del que se desea extraer la informacion
URL = 'https://www.cinesmodernopanama.com/'

# browser.get(URL) navega a esa URL en el navegador Firefox controlado por Selenium
browser.get(URL)

# Se inicializa una lista vacía para almacenar los datos de las películas
peli = []

# funcion que toma varios parámetros como fecha, país, cine, nombre del cine, título de la película, idioma, formato y horarios
def agregar_datos(nombre_cine, titulo, idioma, formato, horarios):
# Se itera sobre cada hora en horarios y crea un diccionario datos que almacena esta información
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
# agrega el diccionario datos a la lista datos
        peli.append(datos)
        
# browser.find_elements() encuentra todos los elementos en la página que corresponden a los enlaces de los teatros (cines)
cartelera = browser.find_elements(
    By.XPATH, '//*[@id="cd-lateral-nav"]/ul[1]/li/ul/li/a')

# links almacena los enlaces (href) de esos teatros para poder navegar a ellos posteriormente
links = [link.get_attribute("href") for link in cartelera]

# Se itera sobre cada enlace en links
for link in links:
    
# browser.get(linki) navega a la página del cine específico
    browser.get(link)
    
 #intenta ejecutar el siguiente bloque de código para manejar excepciones si no se encuentran elementos
    try:
# Se obtiene el nombre del cine de la página web
        nombre_cine = browser.find_element(
            By.XPATH, '//*[@id="cartelera"]/div/div[1]/div/p[2]').text[2:]
#Encuentra todos los elementos que corresponden a las películas en esa página
        peliculas = browser.find_elements(
            By.XPATH, '//*[@id="cartelera"]/div/div[5]/*')
#Iterar sobre cada película encontrada
        for pelicula in peliculas:
#Extrae el título de la película
            titulo = pelicula.find_element(
                By.XPATH, './/div[@class="combopelititulo"]/h2').text
#Extrae el idioma de la película
            idioma = pelicula.find_element(
                By.XPATH, './/div[@class="combodetallepeli"]//div[2]//div[@class="detallespeli"]//div[1]/p').text
#Encuentra todos los elementos que corresponden a los horarios de esa película     
            horarios = pelicula.find_elements(
                By.XPATH, './/div[@class="combodetallepeli"]//div[2]//div[@class="detallespeli"]//div[7]/ul/li/a')
#Crea una lista vacía para almacenar los horarios de la película
            lista_horarios = []
#Cuenta cuántos horarios hay disponibles
            horas = len(horarios)
#Verifica si hay al menos un horario disponible
            if horas > 0:
#Itera sobre cada horario
                for horario in horarios:
#Extrae la hora de proyección del texto del horario
                    hora = horario.text[0:6]
#Extrae el formato (2D, 3D) del texto del horario
                    formato = horario.text[7:-1]
#Agrega la hora a la lista lista_horarios
                    lista_horarios.append(hora)
#Si no hay horarios disponibles, agrega un horario vacío
            else:
                hora = " "
                formato = horario.text[7:-1]
# Agrega la hora a la lista lista_horarios               
                lista_horarios.append(hora)
#agregar la información de la película a la lista peli.
            agregar_datos(nombre_cine, titulo, idioma, formato, lista_horarios)
#Si algún elemento no se encuentra, la excepción es manejada y se continúa con la siguiente iteración            
    except NoSuchElementException: 
        pass
    
# Cerrar el navegador
browser.quit()

# Se convierte la lista peliculas en un DataFrame de pandas
df = pd.DataFrame(peli)

# Se exporta el DataFrame a un archivo Excel
df.to_excel('cinesmoderno - '+str(date.today())+'.xlsx', index=False)
