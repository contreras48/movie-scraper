from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from datetime import date
import pandas as pd


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

browser = webdriver.Firefox('geckodriver')
URL = "https://www.multicinema.com.sv/"
peliculas = []

browser.get(URL)

cartelera = browser.find_element(By.XPATH, '//ul[@class="nav navbar-nav navbar-right"]//li//a[@href="Cartelera.php"]')
cartelera.click()

lenght = len(browser.find_elements(By.XPATH, '//select[@name="Complejo"]/option'))


for i in range(1, lenght):
  cine = browser.find_element(By.XPATH, '//select[@name="Complejo"]')
  select = Select(cine)
  select.select_by_index(i)

  cine = select.first_selected_option.text

  btn = browser.find_element(By.XPATH, '//input[@type="submit" and contains(@value, "Consulta")]')
  btn.click()

  lista_peliculas = browser.find_elements(By.XPATH, '//div[@class="panel panel-info"]//div[@class="panel-body"]')

  for mpelicula in lista_peliculas:
    titulo = mpelicula.find_element(By.XPATH, './/span[@class="media-heading"]').text
    
    try:
      formato_3d = mpelicula.find_element(By.XPATH, './/font[contains(text(),"Español 3 D")]')
    except NoSuchElementException:
      formato_3d = None

    xpath = './/font[contains(text(),"Español 2D")]/following-sibling::button[following-sibling::font[contains(text(),"Español 3 D")]]' if formato_3d  else './/font[contains(text(),"Español 2D")]/following-sibling::button'

    horarios_2d = mpelicula.find_elements(By.XPATH, xpath)
    horarios_3d = mpelicula.find_elements(By.XPATH, './/font[contains(text(),"Español 3 D")]/following-sibling::button')
    
    fun(date.today(), "El Salvador", "Multicinema", cine, titulo,  "DOB", "2D", horarios_2d )  
    fun(date.today(), "El Salvador", "Multicinema", cine, titulo, "DOB", "3D", horarios_3d )  

browser.quit()

df = pd.DataFrame(peliculas)
df.to_excel('horarios_multicinema.xlsx', index=False)