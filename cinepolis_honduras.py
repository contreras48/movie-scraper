from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd
from time import sleep

geckodriver_path = './geckodriver' 
service = Service(executable_path=geckodriver_path)
browser = webdriver.Firefox(service=service)

def agregar_datos(nombre_cine, titulo, idioma, formato, horario):
  for hora in horario:
    datos = {
      "Fecha": date.today(),
      "País": "Honduras",
      "Cine": "Cinépolis",
      "Nombre de cine": nombre_cine,
      "Titulo": titulo,
      "Idioma": idioma,
      "Formato": formato,
      "Hora": hora.text
    }
    horarios_pelicula.append(datos)

URL = 'https://cinepolis.com.hn/'
horarios_pelicula = []

browser.get(URL)

video_popup = browser.find_element(By.XPATH, '//div[@class="welcome-video-popup"]')
if video_popup.is_displayed:
  btn_close = video_popup.find_element(By.XPATH, './/div[@class="button button-close"]')
  btn_close.click()
#//div[@class="Cinema_cinema__3_YUn"]
cinemas = browser.find_elements(By.XPATH, '//div[@id="popup-cinemas"]//div[@class="cinemas"]//div[@class="Cinema_cinema__3_YUn"]/a')
link_cines = [link.get_attribute("href") for link in cinemas]

for link in link_cines:
  browser.get(link)

  cartelera = browser.find_elements(By.XPATH, '//div[@class="movie-projections column"]//div[@class="movie-projection__inner"]/a')
  links_pelicula = []

  for c in cartelera:
    titulo = c.find_element(By.XPATH, './/h2').text
    url = c.get_attribute('href')
    links_pelicula.append([titulo, url])

  for titulo, link_pelicula in links_pelicula:
    browser.get(link_pelicula)

    info_pelicula = browser.find_element(By.XPATH, '//div[@class="movie-projection"]')
    nombre_cine = info_pelicula.find_element(By.XPATH, './/h2[@class="title-cinema"]').text
    cantidad_exhibicion = len(info_pelicula.find_elements(By.XPATH, './/h3[@class="title-attribute"]'))

    for i in range(1, (cantidad_exhibicion + 1)):
      formato_idioma = info_pelicula.find_element(By.XPATH, './/h3['+str(i)+']').text.split(' ')
      formato = ''
      idioma = ''

      for fi in formato_idioma:
        if(fi == 'VIP' or fi == '2D' or fi == '3D' or fi == 'MacroXE'):
          formato += fi + ' '
        else:
          idioma = fi

      horario = info_pelicula.find_elements(By.XPATH, './/ul['+str(i)+']/li/label')

      agregar_datos(nombre_cine, titulo, idioma, formato, horario)
      

browser.quit()

df = pd.DataFrame(horarios_pelicula)
df.to_excel('Cinépolis - Honduras - '+str(date.today())+'.xlsx', index=False)