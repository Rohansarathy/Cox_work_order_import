import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import time
import traceback
def log_message(log_file, message):
    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(message + '\n')
    print(message)
def fuse_login(driver, credentials, log_file):
    FuseLogin = False
    try:
        driver.execute_script(f"window.open('{credentials['FuseURL']}');")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(5)
        
        try:
            WebDriverWait(driver, 50).until(EC.alert_is_present())
            time.sleep(1)
            alert = driver.switch_to.alert
            alert.accept()
            print("Popup clicked")
            time.sleep(6)
        except TimeoutException:
            print("No alert appeared.")
        try:
            time.sleep(8)
            try: 
                driver.execute_script("location.reload(true)")
            except:
                driver.refresh()
            print("Home Page is Visible")
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="myNavbar"]/ul/li[1]/a')))
            FuseLogin = True
        except:
            try: 
                driver.execute_script("location.reload(true)")
            except:
                driver.refresh()
            print("Fuse login is Visibled")
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//input[@name='username']"))).click()
            driver.find_element(By.XPATH, "//input[@name='username']").send_keys(credentials['Fusername'])
            time.sleep(1)
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//input[@name='password']"))).click()
            driver.find_element(By.XPATH, "//input[@name='password']").send_keys(credentials['Fpassword'])
            time.sleep(1)
            driver.find_element(By.XPATH, "//a[normalize-space()='Login']").click()
            time.sleep(3)
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/mobile/mobile-login.php']")))
            log_message(log_file, "Fuse login was successful.")
            FuseLogin = True
            time.sleep(8)
    except Exception as e:
        log_message(log_file, f"Fuse login was Unsuccessful: {str(e)}")
        log_message(log_file, traceback.format_exc())
    log_message(log_file, f"Fuse login:{FuseLogin}")

    return FuseLogin
    