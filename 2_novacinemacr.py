from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import pandas as pd
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

browser = webdriver.Firefox()        # Inicia el navegador Firefox

URL = 'https://www.novacinemas.cr/'  # Establece la URL a visitar

browser.get(URL)                     # Abre la página de Novacinemas

horarios_pelicula = []               # Lista para almacenar la información de los horarios de las películas

# Función para agregar los datos de la película a la lista
def agregar_datos(nombre_cine, titulo, idioma, formato, hora):
  
    datos = {
      "Fecha": date.today(), # Almacena la fecha actual (+ timedelta(days=1),  # Modificación: obtener la fecha del día siguiente)
      "País": "Costa Rica",             # Establece el país como Costa Rica
      "Cine": "Novacinemas",            # Establece el nombre de la cadena de cines
      "Nombre del cine": nombre_cine,   # Nombre del cine específico
      "Titulo": titulo,                 # Título de la película
      "Idioma": idioma,                 # Idioma de la película
      "Formato": formato,               # Formato de la película (2D, 3D, etc.)
      "Hora": hora                      # Hora de la función

    }

    horarios_pelicula.append(datos)     # Agrega el diccionario con los datos a la lista

# Encuentra el enlace de la cartelera y hace clic en él
try:
    cartelera_link = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//ul[@id="menu-menu"]//li[1]/a'))
    )
    cartelera_link.click()
except TimeoutException:
    print("Error: No se pudo encontrar el enlace de la cartelera.")
    browser.quit()
    exit()  

try:
# Itera sobre los cines disponibles (ajusta el rango según el número de cines)
    for i in range(1,4): # Cambia el rango a `range(1, n)` donde n es el número total de cines
            
        cartelera = browser.find_element(By.XPATH, '//div[@id="cartelera"]') # Encuentra la sección de cartelera
        sleep(3)        # Espera 3 segundos para que la cartelera se cargue completamente
        
        # Espera hasta que el menú desplegable de cines sea clicable y luego hace clic
        select_cine = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@id="selectCinema"]/span'))
        )
        select_cine.click()
        # Encuentra y selecciona el cine correspondiente basado en la iteración
        cine = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@id="selectCinema"]//div[@class="sel-box"]/ul/li['+str(i)+']/a'))
            )
        id_cine = cine.get_attribute('data-cinema-id')  # Obtiene el ID del cine seleccionado
        cine.click()                                    # Hace clic para seleccionar el cine

        # Espera hasta que el menú de fechas sea clicable y luego hace clic
        select_fecha = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-fecha"]/nav/a'))
        )
        select_fecha.click()
        
        # Encuentra y selecciona la fecha para la cartelera 
        fecha = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-fecha"]/nav/ul/li[2]/a'))
            )           #Cambiado a li[3] para obtener el segundo día
        fecha.click()
        
        # Encuentra todas las películas para el cine y la fecha seleccionados
        divs = browser.find_elements(By.XPATH, '//div[@id="cinema-'+id_cine+'"][2]/ul/*')
        nombre_cine = browser.find_element(By.XPATH, '//div[@id="cinema-'+id_cine+'"][2]/h2').text  # Obtiene el nombre del cine
        
        for div in divs:            # Recorre cada película disponible en la cartelera
            
            if div.is_displayed():  # Verifica que la película esté visible
                # Encuentra todas las películas en la lista
                peliculas = div.find_elements(By.XPATH, './li//div[@class="cols movieDates"]') 
                
                #print(len(peliculas))
                for pelicula in peliculas:

                    # Obtiene el título de la película
                    titulo = pelicula.find_element(By.XPATH, './/div[@class="col movieDates__right"]/div/h3').text

                    # Encuentra los horarios de la película
                    horarios = pelicula.find_elements(By.XPATH, './/div[@class="showTimes"]//div[@class="items-list"]/*')

                    for horario in horarios:                                            # Recorre cada horario encontrado
                        idioma_formato = horario.find_element(By.XPATH, './span').text   # Obtiene el formato e idioma del horario
                        partes = idioma_formato.split(" ")

                        if len(partes) == 2:
                                formato, idioma = partes
                        else:
                            # En caso de que no se pueda separar en dos partes, asignar valores predeterminados o continuar
                            formato = partes[0] if len(partes) > 0 else "Desconocido"
                            idioma = "Desconocido"

                        hora = horario.find_element(By.XPATH, './a').text    # Obtiene la hora de la función

                        # Agrega los datos obtenidos a la lista
                        agregar_datos( nombre_cine, titulo, idioma, formato, hora)
except TimeoutException:
    print("Error: No se pudo encontrar el enlace del cine.")
    browser.quit()
    exit()                
            
                   
                 
                
              
               
# Cierra el navegador           
browser.quit()

# Crea un DataFrame a partir de la lista de horarios de películas
df = pd.DataFrame(horarios_pelicula)
# Guarda los datos en un archivo Excel con la fecha actual en el nombre
df.to_excel('Novacinemas - '+str(date.today())+'.xlsx', index = False)

            

            

    
    
    
       
   


    


 