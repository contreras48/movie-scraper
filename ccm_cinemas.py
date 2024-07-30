from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from datetime import date
import pandas as pd
from time import sleep

browser = webdriver.Firefox()
URL = 'https://www.ccmcinemas.com/'
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

browser.get(URL)

links_cine_elem = browser.find_elements(By.XPATH, '//div[@class="inner"]//p[1]/a')
links_cine = [link.get_attribute('href') for link in links_cine_elem]

for link_cine in links_cine:
  browser.get(link_cine)
  sleep(3)
  modal = browser.find_element(By.XPATH, '//div[@class="spu-box  spu-centered spu-total- "]')
  if modal.is_displayed() == True:
    btn_modal_close = browser.find_element(By.XPATH, '//i[@class="spu-icon spu-icon-close"]')
    btn_modal_close.click()

  sleep(1)
  cartelera = browser.find_element(By.XPATH, '//ul[@class="nav navbar-nav cactus-main-menu cactus-megamenu"]//li[2]/a')
  cartelera.click()
  sleep(2)

  lista_peliculas = browser.find_elements(By.XPATH, '//div[@class="wpb_column vc_column_container vc_col-sm-12"]//div[@class="pt-cv-ifield"]/a')
  links_pelicula = [link.get_attribute("href") for link in lista_peliculas]
  for link_pelicula in links_pelicula:
    browser.get(link_pelicula)
    sleep(1)

    fecha = date.today()
    pais = "Costa Rica"
    cine = "CCM Cinemas"
    nombre_cine = browser.find_element(By.XPATH, '//div[@id="main-nav"]//img').get_attribute("alt")
    
    iframe_pelicula = browser.find_element(By.XPATH, '//iframe[@id="CCMTANDAS"]')
    browser.switch_to.frame(iframe_pelicula)

    pelicula = browser.find_element(By.XPATH, '//form[@id="form1"]//div[3]')
    
    titulo = pelicula.find_element(By.XPATH, './/div[@id="ContentPlaceHolder1_Nombre_Peli"]/span').text
    
    contenedor_horarios = pelicula.find_elements(By.XPATH, './/div[@id="ContentPlaceHolder1_accordion"]/*')
    info_horarios = []
    temp = []

    for elemento in contenedor_horarios:
      if elemento.tag_name != 'br':
        if elemento.get_attribute("class") != 'clear':
          temp.append(elemento)
      else:
        info_horarios.append(temp)
        temp = []

    for horario in info_horarios:
      formato, idioma = horario[0].find_element(By.XPATH, './/span').text.split(",")
      horas_elements = horario[1].find_elements(By.XPATH, './/span[@class="TandasHoraCalendario"]')
      horas = [hora for hora in horas_elements]

      fun(fecha, pais, cine, nombre_cine, titulo, idioma, formato, horas)




    browser.switch_to.default_content()

browser.quit()

df = pd.DataFrame(peliculas)
df.to_excel('horarios_CCM_Cinemas.xlsx', index=False)