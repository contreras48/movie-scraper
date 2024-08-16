from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date, timedelta # Importar timedelta para agregar el dia siguiente a la fecha
import pandas as pd

browser = webdriver.Firefox() # Inicializa el navegador Firefox utilizando Selenium

URL = 'https://cinestar.com.gt/' # URL del sitio web de Cinestar

browser.get(URL) # Navega a la URL especificada

peli = [] # Lista para almacenar la información de las películas

# Función que agrega los datos de cada película a la lista 'peli'
def fun(fecha, pais, cine, nombre_cine, titulo, idioma, formato, horarios):
  for h in horarios:
    datos = {
      "Fecha": fecha,                 # Fecha de la función de la película
      "País": pais,                   # País donde se encuentra el cine
      "Cine": cine,                   # Cadena de cines (en este caso 'Caribbean Cinemas')
      "Nombre de cine": nombre_cine,  # Nombre específico del cine
      "Titulo": titulo,               # Título de la película
      "Idioma": idioma,               # Idioma en que se proyecta la película
      "Formato": formato,             # Formato de la proyección (2D, 3D, etc.)
      "Hora": h.text                  # Horario de la función
    }
    peli.append(datos) # Agrega los datos de la película a la lista 'peli'

# Encuentra los enlaces a los diferentes cines dentro de la página
teatros = browser.find_elements(By.XPATH, '//ul[@id="menu-main-menu"]//li[5]//ul[@class="sub-menu "]/li/a')
lista1 = [link.get_attribute('href') for link in teatros] # Obtiene la URL de cada cine

for i in lista1:         # Itera sobre la lista de URLs de los cines
    browser.get(i)       # Navega a la URL del cine actual
    fecha = date.today() # Navega fecha actual # Modificación: fecha del día siguiente agregar + timedelta(days=1)
    pais = 'Guatemala'
    cine = 'Caribbean Cinemas'
    nombre_cine = browser.find_element(By.XPATH, '//div[@id="cineinfo"]//b').text # Nombre del cine
    
    # Encuentra todos los elementos que contienen las películas y sus horarios
    cartelera = browser.find_elements(By.XPATH, '//div[@id="horarios"]/*')
    
    # Itera sobre cada película en la cartelera
    for movie in cartelera:
        titulo = movie.find_element(By.XPATH, './/b')         # Encuentra el título de la película
        idioma_formato = movie.find_element(By.XPATH, './/i') # Encuentra el idioma y formato
        idioma_formato = idioma_formato.text.split(' ')       # Separa el texto para obtener el idioma y el formato
        if len(idioma_formato)>1:
            formato = idioma_formato[0] # Si hay formato, lo asigna
            idioma = idioma_formato[1]  # Asigna el idioma
        else :
            idioma = idioma_formato[0]  # Si solo hay un elemento, asume que es el idioma
            formato = '2D'              # Asume que el formato es '2D' por defecto
        horario = movie.find_elements(By.XPATH, './/div [@class="column three-fourth"]//div[2]/a') # Encuentra los horarios
        fun(fecha, pais, cine, nombre_cine, titulo.text, idioma, formato, horario) # Llama a la función para agregar los dato
                           

	  
   
# Cierra el navegador       
browser.quit()

# Crea un DataFrame con los datos recopilados y lo exporta a un archivo Excel
df = pd.DataFrame(peli)
df.to_excel('Cinestar - '+str(date.today())+'.xlsx', index = False) # Guarda el DataFrame en un archivo Excel