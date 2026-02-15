import os
import time
from datetime import datetime
from Sendmail import Sendmail
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException

def log_message(log_file, message):
    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(message + '\n')
    print(message)

def switch_to_window_by_title_contains(driver, needle, timeout=10):
    WebDriverWait(driver, timeout).until(lambda d: len(d.window_handles) >= 1)
    for header in driver.window_handles:
        driver.switch_to.window(header)
        if needle.lower() in (driver.title or "").lower():
            return header
    raise RuntimeError(f"No window with title containing '{needle}'")
def cox_bulk_import(driver, mode, COX_FILE, target_date, credentials, log_file, previous_date, current_hour):
    log_message(log_file, "going to import the excel file.")
    print("going to import the excel file.")
    upload_success = False
    #driver.switch_to.window(driver.window_handles[1])
    handle = switch_to_window_by_title_contains(driver, "FUSE", timeout=10)
    print(f"Active window: {handle} | Title: {driver. title}")
    log_message(log_file, f"Active window: {handle} | Title: {driver. title}")
 
    try:
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//a[text()="Work Order"]')))
        fuse_login = True
    except Exception:
        fuse_login = False
        
    if fuse_login:
        try:
            # try:
            #     # Try finding Import Type directly
            #     import_type = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Import_type"]')))
            #     print("Import type found directly.")
            # except:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Work Order']"))).click()
            time.sleep(2)
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(text())='Schedule Work Order Import']"))).click()
            time.sleep(2)
            import_type = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="Import_type"]')))
            # ----- Import Type Dropdown -----
            Select(import_type).select_by_value("2")
            time.sleep(5)
            # ----- Import Template Dropdown -----
            template = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="templateid"]')))
            Select(template).select_by_value("5")
            time.sleep(1)
            ####Upload File####
            upload = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="fileToUpload"]')))
            upload.send_keys(COX_FILE)
            print(f"Uploaded file: {COX_FILE}")
            log_message(log_file, f"Uploaded file: {COX_FILE}")
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, "//button[normalize-space(text())='Schedule Task']")))
            schedule_task = driver.find_element(By.XPATH, "//button[normalize-space(text())='Schedule Task']")
            schedule_task.click()
            print(f"\033[92mUploaded {COX_FILE} Work Order Import successfully.\033[0m")
            time.sleep(2)
            # #### Mail Logic ####
            print("It's time to send a mail")
            recipient_emails = credentials['sdavis']
            cc_emails = credentials['ybotID']
            if mode == "previous":
                subject = f"Cox upload completed for previous day - {previous_date}"
                body_message = f"The COX work order import {previous_date} previous day excel  file has been successfully uploaded into Fuse."
            else:
                subject = f"Cox upload completed for current day - {target_date}-{current_hour}"
                body_message = f"The COX work order import {target_date} current day excel file has been successfully uploaded into Fuse."
            body_message1 = ""
            attachment_path = ""
            try:
                Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path)
                log_message(log_file, "Mail sent Successfully for uploaded Work Order bulk file.")
            except Exception:
                log_message(log_file, "Mail not sent for uploaded Work Order bulk file.")
        except TimeoutException:
            print("\033[91mUpload failed due to timeout:\033[0m")
            log_message(log_file, "Upload failed due to timeout")
            print("It's time to send a mail")
            recipient_emails = credentials['ybotID']
            cc_emails = credentials['ybotID']
            subject = f"Fuse Upload failed for {COX_FILE}"
            body_message = f"The output file for the {mode} department failed to upload into Fuse."
            body_message1 = ""
            attachment_path = ""
            try:
                Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path)
                log_message(log_file, "Mail sent Successfully for failed to upload Work Order bulk file.")
            except Exception:
                log_message(log_file, "Mail not sent for failed to upload Work Order bulk file.")
        