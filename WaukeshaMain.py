import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from AuthorizationData import epic_password
from AuthorizationData import epic_login
import pandas as pd
import glob
from selenium.webdriver.common.keys import Keys


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
    driver.implicitly_wait(30)

    # entering a password
    password_input = driver.find_element(By.ID,'TextBox3')
    password_input.clear()
    password_input.send_keys(epic_password)
    driver.implicitly_wait(30)

    # press login button
    login_button_click = driver.find_element(By.ID, 'Button1').click()
    driver.implicitly_wait(30)

    print('Start downloading the bom files...')
    counter = 0
    not_downloaded_list = []
    for elem in serial_numbers:
        try:
            # fill in the search field with the serial number
            search_field = driver.find_element(By.ID, '40')
            search_field.clear()
            search_field.send_keys(elem)
            search_field.send_keys((Keys.ENTER))
            driver.implicitly_wait(20)

            # press plus button
            plus_button = driver.find_element(By.ID, '30')
            plus_button.click()
            driver.implicitly_wait(20)

            # press  BOM button
            BOM_button = driver.find_element(By.ID, '78')
            BOM_button.click()
            driver.implicitly_wait(10)

            # switch window and press export to excel button
            driver.switch_to.window(driver.window_handles[1])
            excel_button = driver.find_element(By.ID, 'btnExportExcel')
            excel_button.click()
            driver.implicitly_wait(20)

            counter += 1
            current_time = time.time()

            print(f'{counter}/ {total_sn} was downloaded ({elem}). {round(current_time - start_time, 1)} seconds have passed.')
            driver.close()
            driver.switch_to.window(driver.window_handles[1])
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except:
            counter += 1
            current_time = time.time()
            print(f'{counter}/ {total_sn} was not downloaded ({elem}). {round(current_time - start_time, 1)} seconds have passed.')
            not_downloaded_list.append(elem)
        finally:
            continue

except Exception as ex:
    print(ex)
finally:
    driver.close()
    driver.quit()
print('File downloading completed!')

print('These files were not downloaded:')
for el in not_downloaded_list:
    print(el)
