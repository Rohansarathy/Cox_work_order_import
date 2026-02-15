import shutil
import time
import os
import re
import json
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

from Sendmail import Sendmail
from cox_login import cox_login
from excel_process import validate_excel


credentials_file = 'cox_work_order.json'
with open(credentials_file, 'r') as file:
    credentials = json.load(file)

current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
print(f"{current_time}")
year = current_time[0:4]
month = current_time[5:7]
day = current_time[8:10]
mint = current_time[11:13]
sec = current_time[14:16]
log_file = f"{credentials['Logfile']}\\cox_work_order_{day}{month}{year}_{mint}_{sec}.txt"

def log_message(log_file, message):
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(message + '\n')
    print(message)




main_file = credentials['Default_Path']
main_excel = os.path.join(main_file, "Cox.xlsx")
main_file_folder = os.path.join(main_file, "main_file")
previous_folder = os.path.join(main_file_folder, "previous_day")
current_folder = os.path.join(main_file_folder, "current_day")
# Ensure folders exist
os.makedirs(previous_folder, exist_ok=True)
os.makedirs(current_folder, exist_ok=True)
if os.path.exists(main_excel):
    # Destination paths
    prev_dest = os.path.join(previous_folder, "Cox.xlsx")
    curr_dest = os.path.join(current_folder, "Cox.xlsx")
    # Remove existing files if present
    if os.path.exists(prev_dest):
        os.remove(prev_dest)

    if os.path.exists(curr_dest):
        os.remove(curr_dest)
    # Copy to both folders
    shutil.copy2(main_excel, prev_dest)
    shutil.copy2(main_excel, curr_dest)
    print("Cox.xlsx copied to both previous_day and current_day folders.")
else:
    print("Cox.xlsx not found in base folder.")

base_folder = r"C:\Users\RohansarathyGoudhama\Downloads\cox"
now = datetime.now()
cutoff_time = now.replace(hour=8, minute=0, second=0, microsecond=0)
before_8am = now < cutoff_time
# previous_folder = os.path.join(credentials['Default_Path'], "previous_day")
# current_folder = os.path.join(credentials['Default_Path'], "current_day")
previous_folder = os.path.join(base_folder, "previous_day")
current_folder = os.path.join(base_folder, "current_day")
os.makedirs(previous_folder, exist_ok=True)
os.makedirs(current_folder, exist_ok=True)
if now < cutoff_time:
    download_folder = previous_folder
    COX_FILE = os.path.join(download_folder, "Cox.xlsx")
    target_date = (now - timedelta(days=1)).date()
    print(f"Before 8 AM - Using yesterday: {target_date}")
else:
    download_folder = current_folder
    COX_FILE = os.path.join(download_folder, "Cox.xlsx")
    target_date = now.date()
    print(f"After 8 AM - Using today: {target_date}")

chrome_options = Options()
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument(r"--user-data-dir=C:\ChromeSeleniumProfiles\CoxWorkOrder")


chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})
chrome_options.add_experimental_option("detach", True)

service = Service(ChromeDriverManager("144.0.7559.132").install())
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.maximize_window()
driver.execute_cdp_cmd("Page.bringToFront", {})

##### Dlete existing files in download folder #####
for filename in os.listdir(download_folder):
    # print(f"Checking existing file: {filename}")
    file_path = os.path.join(download_folder, filename)
    print(f"Deleting existing file: {file_path}")
    try:
        if os.path.isfile(file_path):
            os.remove(file_path)  # delete file
    except Exception as e:
        print(f"Failed to delete {file_path}: {e}")

#### Cox Login #####
cox_login(driver, credentials, log_file)
if cox_login:
    print("COX Login Successful")
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="manage-content"]/div[1]/div[2]/div[2]/div/div[2]/div[3]')))
    time.sleep(5)
    first_area = driver.find_element(By.XPATH, "(//span[@class='rtl-text rtl-prov-name' and text()='AZ_ITG'])[1] ")
    first_area.click()
    time.sleep(3)
    seen_ids = set()
    index = 0
    while True:
        areas = driver.find_element(By.XPATH,"(//*[@id='manage-content']//div[contains(@class,'edt-root')])[1]")
        rows = areas.find_elements(By.XPATH,"./div[contains(@class,'edt-item') and @role='treeitem']")
        if index >= len(rows):
            break
        try:
            row = rows[index]
            row_id = row.get_attribute("data-id")
            if row_id in seen_ids:
                index += 1
                continue

            seen_ids.add(row_id)
            label = row.find_element(By.XPATH,".//span[contains(@class,'rtl-prov-name')]").text.strip()
            print(f"Processing: {label} ({row_id})")
            row.click()
            time.sleep(3)
        except StaleElementReferenceException:
            print("Tree refreshed. Refetching...")
            # index += 1
            continue
        dates = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//button[@data-ofsc-id='toolbar-date-button-value']//span[@class='app-button-title'])[1]")))
        raw_date = dates.text.strip()
        print(f"Current Date Displayed: {raw_date}")

        clean_date = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', raw_date)

        ui_datetime = datetime.strptime(clean_date, "%A, %B %d, %Y")
        ui_date = ui_datetime.date()  
        ui_date_str = ui_datetime.strftime("%m/%d/%Y") 

        print(f"UI Date (logic): {ui_date}")
        print(f"UI Date (display): {ui_date_str}")

        today = datetime.today().date()
        yesterday = today - timedelta(days=1)
        yesterday_str = yesterday.strftime("%m/%d/%Y")
        print(f"yesterday date:{yesterday_str}")
        if before_8am:
            if ui_date == today:
                print("Dates match.")
                click_left_arrow = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//span[contains(@class,'oj-ux-ico-chevron-left')])[1]")))
                click_left_arrow.click()
                if ui_date - timedelta(days=1) == yesterday:
                    print("Yesterday's date confirmed.")
                    # ### View ####
                    time.sleep(5)
                    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//span[@data-ofsc-role='button-text' and normalize-space()='View']/ancestor::button)[1]")))
            elif ui_date == yesterday:
                print("UI already shows yesterday.")
        else:
            print("After 8 AM -> Using today.")
            if ui_date == yesterday:
                print("UI showing yesterday but after 8 AM. No arrow click required (as per your logic).")

            elif ui_date == today:
                print("UI already shows today.")
        time.sleep(2)
        # #### check No items to display ####
        no_items = driver.find_elements(By.XPATH,"//div[contains(@class,'oj-datagrid-empty-text') and normalize-space()='No items to display.']")
        if no_items:
            log_message(log_file, f"No items to display for area: {label} ({row_id}) on date: {yesterday_str}")
            print(f"No items to display for area: {label} ({row_id}) on date: {yesterday_str}")
            print(f"\033[95mNo Items found for area: {label} ({row_id})\033[0m")
            print("---------------------------------------------------")
            index += 1
            continue

        #### View ####
        view_dropdown = driver.find_element(By.XPATH, "(//span[@data-ofsc-role='button-text' and normalize-space()='View']/ancestor::button)[1]")
        view_dropdown.click()
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//*[contains(@id,'dc__top_panel__viewFilter')]/div")))
        apply_label = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,"//label[.//oj-option[normalize-space()='Apply hierarchically']]")))
        apply_hierarchically_checkbox = driver.find_element(By.XPATH,"//label[.//oj-option[normalize-space()='Apply hierarchically']]//preceding-sibling::span//input")
        if not apply_hierarchically_checkbox.is_selected():
            print("Checkbox is not selected. Clicking it now.")
            apply_label.click()
        else:
            print("Checkbox is already selected. No action needed.")
        click_apply = driver.find_element(By.XPATH, "(//button[.//span[@data-ofsc-role='button-text' and normalize-space()='Apply']])[2]")
        click_apply.click()
        time.sleep(2)
        #### Actions #####
        actions = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Actions']]")))
        actions.click()
        time.sleep(2)
        print(("Actions clicked"))
        export_excel = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@role='dialog' and @aria-label='Actions']//button[.//span[normalize-space()='Export']]")))
        export_excel.click()
        time.sleep(2)
        print("Export to Excel clicked")
        print(f"\033[92mFile downloaded for area: {label} ({row_id})\033[0m")
        print("---------------------------------------------------")
        index += 1
    #### Go to Validate Excel Process ####
    validate_excel(driver, COX_FILE, credentials, download_folder, log_file)
else:
    print("COX Work Order Login Failed")
    recipient_emails = credentials['ybotID']
    cc_emails = credentials['ybotID']
    subject = f"COX work order Import login failed"
    body_message = f"COX work order Import login failed or Fuse login fialed."
    body_message1 = ""
    attachment_path = ""
    try:
        Sendmail(recipient_emails, cc_emails, subject, body_message, body_message1, attachment_path)
        log_message(log_file, "Mail sent Successfully for Error Logs.")
        print("\033[96mMail sent Successfully for Error Logs.\033[0m")
    except Exception:
        log_message(log_file, "Mail not sent for Error Logs.")