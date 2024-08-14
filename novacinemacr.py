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

#Busca Cartelera y hace click
cartelera_link = browser.find_element(By.XPATH, '//ul[@id="menu-menu"]//li[1]/a')
sleep(1)
cartelera_link.click()


for i in range(1,4):
    cartelera = browser.find_element(By.XPATH, '//div[@id="cartelera"]')
    sleep(5)
    select_cine = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@id="selectCinema"]/span'))
    )
    select_cine.click()
    #cine = cartelera.find_element(By.XPATH, './/div[@class="topSelect selector-cine"]//div[@id="selectCinema"]//div[@class="sel-box"]/ul/li//a['+str(i)+']')
    cine = WebDriverWait(browser, 10).until(
        EC.element_to_be_clickable(cartelera.find_element(By.XPATH, './/div[@id="selectCinema"]//div[@class="sel-box"]/ul/li['+str(i)+']/a'))
        )
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
    
            

    
    
    
       
   


    


 