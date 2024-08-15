from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

browser = webdriver.Firefox()

URL = 'https://www.novacinemas.cr/'

browser.get(URL)

horarios_pelicula = []
def agregar_datos(nombre_cine, titulo, idioma, formato, hora):
  
    datos = {
      "Fecha": date.today(),
      "País": "Costa Rica",
      "Cine": "Novacinemas",
      "Nombre del cine": nombre_cine,
      "Titulo": titulo,
      "Idioma": idioma,
      "Formato": formato,
      "Hora": hora

    }

    horarios_pelicula.append(datos)

#Busca Cartelera y hace click
cartelera_link = browser.find_element(By.XPATH, '//ul[@id="menu-menu"]//li[1]/a')
sleep(1)
cartelera_link.click()


for i in range(1,3):
        
    cartelera = browser.find_element(By.XPATH, '//div[@id="cartelera"]')
    sleep(3)
    select_cine = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@id="selectCinema"]/span'))
    )
    select_cine.click()
    #cine = cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-cine"]//div[@id="selectCinema"]//div[@class="sel-box"]/ul/li//a['+str(i)+']')
    cine = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@id="selectCinema"]//div[@class="sel-box"]/ul/li['+str(i)+']/a'))
        )
    id_cine = cine.get_attribute('data-cinema-id')
    
    cine.click()

    select_fecha = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-fecha"]/nav/a'))
    )
    select_fecha.click()
    #cine = cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-cine"]//div[@id="selectCinema"]//div[@class="sel-box"]/ul/li//a['+str(i)+']')
    fecha = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-fecha"]/nav/ul/li[2]/a'))
        )
    fecha.click()
    
    divs = browser.find_elements(By.XPATH, '//div[@id="cinema-'+id_cine+'"][2]/ul/*')
    nombre_cine = browser.find_element(By.XPATH, '//div[@id="cinema-'+id_cine+'"][2]/h2').text 
    
    for div in divs:
        
        if div.is_displayed():
            peliculas = div.find_elements(By.XPATH, './li//div[@class="cols movieDates"]')
            
            print(len(peliculas))
            for pelicula in peliculas:
                titulo = pelicula.find_element(By.XPATH, './/div[@class="col movieDates__right"]/div/h3').text
                horarios = pelicula.find_elements(By.XPATH, './/div[@class="showTimes"]//div[@class="items-list"]/*')
                for horario in horarios:
                   idioma_formato = horario.find_element(By.XPATH, './span').text
                   print(idioma_formato.split(' '))
                   formato, idioma = idioma_formato.split(" ")
                   
                   hora = horario.find_element(By.XPATH, './a').text
                   agregar_datos( nombre_cine, titulo, idioma, formato, hora)
                
            
                   
                 
                
                
                

                

browser.quit()

df = pd.DataFrame(horarios_pelicula)
df.to_excel('Novacinemas - '+str(date.today())+'.xlsx', index = False)

            

            

    
    
    
       
   


    


 