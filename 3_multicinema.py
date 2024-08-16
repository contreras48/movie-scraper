from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from datetime import date
import pandas as pd

geckodriver_path = './geckodriver' 
service = Service(executable_path=geckodriver_path)
browser = webdriver.Firefox(service=service)

def agregar_datos(nombre_cine, titulo, formato, horarios):
  for hora in horarios:
    datos = {
      "Fecha": date.today(),
      "País": "El Salvador",
      "Cine": "Multicinema",
      "Nombre del cine": nombre_cine,
      "Titulo": titulo,
      "Idioma": "Epañol",
      "Formato": formato,
      "Hora": hora.text
    }

    horarios_pelicula.append(datos)

URL = "https://www.multicinema.com.sv/"
horarios_pelicula = []

browser.get(URL)

cartelera = browser.find_element(By.XPATH, '//ul[@class="nav navbar-nav navbar-right"]//li//a[@href="Cartelera.php"]')
cartelera.click()

cantidad = len(browser.find_elements(By.XPATH, '//select[@name="Complejo"]/option'))


for i in range(1, cantidad):
  combobox = browser.find_element(By.XPATH, '//select[@name="Complejo"]')
  select = Select(combobox)
  select.select_by_index(i)

  nombre_cine = select.first_selected_option.text

  btn = browser.find_element(By.XPATH, '//input[@type="submit" and contains(@value, "Consulta")]')
  btn.click()

  lista_peliculas = browser.find_elements(By.XPATH, '//div[@class="panel panel-info"]//div[@class="panel-body"]')

  for pelicaula in lista_peliculas:
    titulo = pelicaula.find_element(By.XPATH, './/span[@class="media-heading"]').text

    try:
      formato_3d = pelicaula.find_element(By.XPATH, './/font[contains(text(),"Español 3 D")]')
    except NoSuchElementException:
      formato_3d = None

    xpath = './/font[contains(text(),"Español 2D")]/following-sibling::button[following-sibling::font[contains(text(),"Español 3 D")]]' if formato_3d  else './/font[contains(text(),"Español 2D")]/following-sibling::button'

    horarios_2d = pelicaula.find_elements(By.XPATH, xpath)
    horarios_3d = pelicaula.find_elements(By.XPATH, './/font[contains(text(),"Español 3 D")]/following-sibling::button')
    
    agregar_datos(nombre_cine, titulo, "2D", horarios_2d )  
    agregar_datos(nombre_cine, titulo, "3D", horarios_3d )  

browser.quit()

df = pd.DataFrame(horarios_pelicula)
df.to_excel('multicinema - '+str(date.today())+'.xlsx', index=False)