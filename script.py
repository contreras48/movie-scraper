from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

browser = webdriver.Firefox()

URL = 'https://www.multicinema.com.sv/'

browser.get(URL)

link = browser.find_element(By.XPATH, '//ul[@class="nav navbar-nav navbar-right"]//li//a[@href="Cartelera.php"]')
link.click()

combobox = browser.find_element(By.XPATH, '//select[@name="Complejo"]')
select = Select(combobox)
lenght = len(select.options)
movie_list = []

for i in range(1, lenght):
    
    combobox = browser.find_element(By.XPATH, '//select[@name="Complejo"]')
    select = Select(combobox)
    print(select.options[i].text)
    select.select_by_index(i)
    btn = browser.find_element(By.XPATH, '//input[@type="submit" and contains(@value, "Consulta")]')
    btn.click()

    m = []
    movies = browser.find_elements(By.XPATH, '//div[@class="panel panel-info"]//div[@class="panel-body"]')
    
    for movie in movies:
      title = browser.find_element(By.XPATH, '//span[@class="media-heading"]').text
      p_2d = browser.find_elements(By.XPATH, "//font[contains(text(),'Español 2D')]/following-sibling::button")
      p_3d = browser.find_elements(By.XPATH, "//font[contains(text(),'Español 3 D')]/following-sibling::button")
      
      h_2d = []
      h_3d = []
      for h in p_2d:
         h_2d.append(h.text)
      
      for h in p_3d:
         h_3d.append(h.text)

      item = {
         "title": title,
         "2d": h_2d,
         "3d": h_3d
      }

      m.append(item)
    
    movie_list.append({i: m})
    print(movie_list)


#elem = browser.find_element(By.NAME, 'p')  # Find the search box
#elem.send_keys('seleniumhq' + Keys.RETURN)

browser.quit()