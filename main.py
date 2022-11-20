import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import openpyxl
import pandas as pd
import numpy
import time
import json

from selenium.webdriver.common.action_chains import ActionChains

def loop():

    driver = webdriver.Chrome()
    driver.maximize_window()

    wrkbk = pd.read_excel('url.xlsx')  

    arr = numpy.array(wrkbk)

    for key in arr:
        print(key[0])
        driver.get(key[0])

        time.sleep(5)

        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")

        time.sleep(5)

        element = driver.find_element(By.TAG_NAME,"head")

        driver.execute_script("return arguments[0].scrollIntoView(true);", element)

        time.sleep(5)

    driver.close()

    time.sleep(5)

    main()

def main():

    chromedriver_autoinstaller.install()

    driver = webdriver.Chrome()
    driver.maximize_window()

     # load excel with its path
    wrkbk = pd.read_excel('url.xlsx')    

    arr = numpy.array(wrkbk)

    for key in arr:
        print(key[0])
        driver.get(key[0])

        time.sleep(10)

        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")

        time.sleep(5)

        element = driver.find_element(By.TAG_NAME,"head")

        driver.execute_script("return arguments[0].scrollIntoView(true);", element)

        time.sleep(5)

    driver.close()

    time.sleep(5)

    loop()



if __name__ == "__main__":
    main()