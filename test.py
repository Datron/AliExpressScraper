from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import ElementNotInteractableException
import time


options = Options()
options.add_argument("--headless")
driver = webdriver.Firefox(firefox_options=options, executable_path="geckodriver.exe")
driver.get("https://sale.aliexpress.com/__pc/bestselling.htm?spm=a2g01.8005310.2004.101.5d8fde54rH5A6M#701")
ul = driver.find_element_by_class_name("menuitems")
links = ul.find_elements_by_tag_name("a")
i = 0
j = 350
for link in links:
    try:
        # link.click()
        driver.execute_script("window.scrollBy("+str(i)+","+str(j)+");")
        time.sleep(7)
        element = WebDriverWait(driver, 20).until(lambda driver: driver.find_element_by_class_name("detail-box"))
        i = j
        j += 350
    except Exception as e:
        print(e)
elements = driver.find_elements_by_class_name("detail-box")
print(len(elements))
driver.close()
