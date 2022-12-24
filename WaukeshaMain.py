import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from AuthorizationData import epic_password
from AuthorizationData import epic_login
import pandas as pd
import glob

start_time = time.time()

# Get a frame serial number base
file_name = input('Enter the name of file with a serial numbers: ')
result = pd.read_excel(f'{file_name}.xlsx')
serial_numbers = result['sn'].tolist()
total_sn = len(serial_numbers)

# Using Selenium
url = 'https://epic.waukeshaengine.com/EpicLogin.aspx?nologin=Y'

# set webdriver options
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "C:\\Users\\hp\\GitHubProjects\\WaukeshaSparePartsDatabase\\Downloads"}
chromeOptions.add_experimental_option("prefs",prefs)
chromeOptions.add_argument("--disable-blink-features=AutomationControlled")
# headless mode
# chromeOptions.headless = True
driver = webdriver.Chrome(options=chromeOptions)


try:
    print('Login Epic Website...')
    # Open and login  Epic website
    driver.get(url=url)

    # entering a email
    email_input = driver.find_element(By.ID,'TextBox2')
    email_input.clear()
    email_input.send_keys(epic_login)
    driver.implicitly_wait(10)

    # entering a password
    password_input = driver.find_element(By.ID,'TextBox3')
    password_input.clear()
    password_input.send_keys(epic_password)
    driver.implicitly_wait(10)

    # press login button
    login_button_click = driver.find_element(By.ID, 'Button1').click()
    driver.implicitly_wait(20)
    time.sleep(10)




except Exception as ex:
    print(ex)
finally:
    driver.close()
    driver.quit()
