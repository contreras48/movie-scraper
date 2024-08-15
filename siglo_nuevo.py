from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, date
import locale
import pandas as pd
from time import sleep

#locale.setlocale(locale.LC_TIME, 'es_SV.utf8')
meses = {
    "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
    "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
    "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"
}

geckodriver_path = './geckodriver' 
service = Service(executable_path=geckodriver_path)
browser = webdriver.Firefox(service=service)

def agregar_datos(nombre_cine, titulo, idioma, formato, hora):
    datos = {
      "Fecha": date.today(),
      "País": "Nicaragua",
      "Cine": "Cine Siglo Nuevo",
      "Nombre del cine": nombre_cine,
      "Titulo": titulo,
      "Idioma": idioma,
      "Formato": formato,
      "Hora": hora
    }

    horarios_pelicula.append(datos)

URL = "https://psiglonuevo.com/"
horarios_pelicula = []

browser.get(URL)

cartelera = browser.find_element(By.XPATH, '//div[@class="fusion-post-cards fusion-post-cards-1 fusion-grid-archive"]')
links_pelicula_elem = cartelera.find_elements(By.XPATH, './ul/li/div/div/div/a')
links_pelicula = [link.get_attribute("href") for link in links_pelicula_elem]

for link_pelicula in links_pelicula:
  browser.get(link_pelicula)

  titulo = browser.find_element(By.XPATH, '//h2[@class="title-heading-center fusion-responsive-typography-calculated"]').text
  
  cines = browser.find_elements(By.XPATH, '//div[@class="nav"]/ul/li/a')
  contenidos = browser.find_elements(By.XPATH, '//div[@class="tab-content"]//div[contains(@class,"tab-pane fade fusion-clearfix")]')

  for cine, contenido in zip(cines, contenidos):
    nombre_cine = cine.find_element(By.XPATH, './h4').text
    cine.click()
    
    tanda = contenido.find_element(By.XPATH, './center[1]/div').text
    
    try:
      fecha = contenido.find_element(By.XPATH, './div[1]/span/strong').text[0:-1]
      _, dia, mes = fecha.split()
      fecha_datetime = datetime.strptime(f"{dia}-{meses[mes]}-{datetime.now().year}", "%d-%m-%Y")
      fecha_actual = datetime.now()

      if fecha_datetime.date() == fecha_actual.date():
        
        horarios = contenido.find_elements(By.XPATH, './div[1]/div')
        if len(tanda.split(' | ')) == 2:
          formato, idioma = tanda.split(' | ')
          formato = formato.split(" ")[1]
          horas = [hora.text for hora in horarios]
          
          for hora in horas:
            if hora.lower() != "no disponible":
              #llamar aqui a la funcion
              agregar_datos(nombre_cine, titulo, idioma[0,3].upper(), formato, hora)

        else:
          formato = tanda.split(' | ')[0]
          formato = formato.split(" ")[1]
          horarios = [horario.text.rsplit(" ", 1) for horario in horarios]
          for horario in horarios:
            hora = horario[0]
            idioma = horario[1]
            agregar_datos(nombre_cine, titulo, idioma, formato, hora)
      else:
        continue


    except NoSuchElementException:
      print("No hay funcines disponibles")
      continue

browser.quit()

df = pd.DataFrame(horarios_pelicula)
df.to_excel('Cine Siglo Nuevo - '+str(date.today())+'.xlsx', index = False)