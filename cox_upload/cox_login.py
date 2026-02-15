import time
import traceback
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
def log_message(log_file, message):
    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(message + '\n')
    print(message)
def cox_login(driver, credentials, log_file):
    coxlogin = False
    try:
        driver.get(credentials['cox_url'])
        driver.maximize_window()   
        time.sleep(10)
        try:
            driver.refresh()
            time.sleep(3)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "(//div[@data-ofsc-role='page-title-text'])[1]")))
            log_message(log_file, "COX login was successful (session already active).")
            driver.find_element(By.XPATH, "(//div[@data-ofsc-role='page-title-text'])[1]")
            coxlogin = True
        except:
            print("COX login page  is visibled")
            driver.refresh()
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="username"]')))
            driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(credentials['username'])
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(credentials['password'])
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="remember_username"]').click()
            time.sleep(3)
            driver.find_element(By.XPATH, '//*[@id="second-step-container"]/div[5]').click()
            time.sleep(10)
            # try:
            #     ##Sign in with Mail##
            #     try:
            #         WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '(//input[@type="email" and @name="loginfmt"])[1]'))).click()
            #         driver.find_element(By.XPATH, ' (//input[@type="email" and @name="loginfmt"])[1]').send_keys(credentials['Mail'])
            #         time.sleep(1)
            #         driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
            #         time.sleep(1)
            #     except:
            #         print("Sign in email input field not found")
            #         pass
            #     try:
            #         WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='passwd' and @type='password']"))).click()
            #         driver.find_element(By.XPATH, "//input[@name='passwd' and @type='password'] ").send_keys(credentials['Mail_password'])
            #         time.sleep(1)
            #         driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
            #         time.sleep(1)
            #     except:
            #         print("Password input field not found")
            #         pass
            #     try:
            #         WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@data-bind='text: display' and contains(normalize-space(.),'Text')]")))
            #         driver.find_element(By.XPATH, "//div[@data-bind='text: display' and contains(normalize-space(.),'Text')]").click()
            #     except:
            #         print("Stay signed in prompt not found")
            #         pass
            #     try: 
            #         WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[text()='stephen.davis@cox.com']")))
            #         driver.find_element(By.XPATH, "//div[text()='stephen.davis@cox.com']").click()
            #         time.sleep(3)
            #     except:
            #         print("Account selection not found")
            #         pass
            #     try:
            #         WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='passwd' and @type='password']"))).click()
            #         driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(credentials['Mail_password'])
            #         time.sleep(1)
            #     except:
            #         print("Password input field not found.")
            #         pass
            #     try:
            #         WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@data-bind='text: display' and contains(normalize-space(.),'Text')]")))
            #         driver.find_element(By.XPATH, "//div[@data-bind='text: display' and contains(normalize-space(.),'Text')]").click()
            #     except:
            #         print("Stay signed in prompt not found.")
            #         pass
            WebDriverWait(driver, 200).until(EC.element_to_be_clickable((By.XPATH, "(//div[@data-ofsc-role='page-title-text'])[1]")))
            log_message(log_file, "COX dispatch console login was successful.")
            coxlogin = True
            # except:
            #     log_message(log_file, "COX login was Unsuccessful.")
            #     coxlogin = False
    except Exception as e:
        log_message(log_file, f"COX login was Unsuccessful: {str(e)}")
        log_message(log_file, traceback.format_exc())
    
    return coxlogin
