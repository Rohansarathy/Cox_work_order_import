
import time
import os
import re
import json
from selenium import webdriver
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

from Sendmail import Sendmail
from cox_login import cox_login
from fuselogin import fuse_login
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


base_folder = r"C:\Users\RohansarathyGoudhama\Downloads\cox"
now = datetime.now()
print(f"now:{now}")
cutoff_time = now.replace(hour=8, minute=0, second=0, microsecond=0)
print(f"Cutoff time:{cutoff_time}")
before_8am = now < cutoff_time
# previous_folder = os.path.join(credentials['Default_Path'], "previous_day")
# current_folder = os.path.join(credentials['Default_Path'], "current_day")
previous_folder = os.path.join(base_folder, "previous_day")
current_folder = os.path.join(base_folder, "current_day")
os.makedirs(previous_folder, exist_ok=True)
os.makedirs(current_folder, exist_ok=True)
previous_date = None
current_hour = None
if now < cutoff_time:
    download_folder = previous_folder
    mode = "previous"
    # target_date = (now - timedelta(days=1)).date()
    previous_date = (now - timedelta(days=1)).strftime("%m/%d/%Y")
    target_date = (now - timedelta(days=1)).strftime("%m%d%Y")
    COX_FILE = os.path.join(download_folder, f"Cox{target_date}.xlsx")
    log_message(log_file, "@@@@@@@@@@@@@@@@@@@@@@ COX Dispatch Process Start  for Previous Day @@@@@@@@@@@@@@@@@@@")
    print(f"Before 8 AM - Using yesterday: {target_date}")
    log_message(log_file, f"Before 8 AM - Using yesterday: {target_date}")
    print(f"Before 8 AM → Using yesterday file: {COX_FILE}")
    log_message(log_file, f"Before 8 AM → Using yesterday file: {COX_FILE}")
else:
    download_folder = current_folder
    # target_date = now.date()
    mode = "current"
    target_date = now.strftime("%m/%d/%Y")
    # current_hour = now.strftime("%H")
    # current_hour = now.strftime("%I")
    current_hour = now.strftime("%I%p")
    COX_FILE = os.path.join(download_folder,  f"Cox{current_hour}.xlsx")
    log_message(log_file, "@@@@@@@@@@@@@@@@@@@@@@ COX Dispatch Process Start for Current Day @@@@@@@@@@@@@@@@@@@")
    print(f"After 8 AM - Using today: {target_date}")
    log_message(log_file, f"After 8 AM - Using today: {target_date}")
    print(f"After 8 AM → Using current file: {COX_FILE}")
    log_message(log_file, f"After 8 AM → Using current file: {COX_FILE}")

chrome_options = Options()
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument(r"--user-data-dir=C:\ChromeSeleniumProfiles\CoxWorkOrder")


chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})
chrome_options.add_experimental_option("detach", True)

service = Service(ChromeDriverManager("144.0.7559.133").install())
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.maximize_window()
driver.execute_cdp_cmd("Page.bringToFront", {})

##### Delete existing files in download folder #####
for filename in os.listdir(download_folder):
    # print(f"Checking existing file: {filename}")
    file_path = os.path.join(download_folder, filename)
    print(f"Deleting existing file: {file_path}")
    log_message(log_file, f"Deleting existing file: {file_path}")
    try:
        if os.path.isfile(file_path):
            os.remove(file_path)  # delete file
    except Exception as e:
        print(f"Failed to delete {file_path}: {e}")
        log_message(log_file, f"Failed to delete {file_path}: {e}")

#### Cox Login #####
coxlogin = cox_login(driver, credentials, log_file)
log_message(log_file, f"cox login:{coxlogin}")
FuseLogin = fuse_login(driver, credentials, log_file)
log_message(log_file, f"Fuse login:{FuseLogin}")
if coxlogin and FuseLogin:
    log_message(log_file, "COX & Fuse Login Successful")
    driver.switch_to.window(driver.window_handles[0])
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
            log_message(log_file, f"Processing: {label} ({row_id})")
            row.click()
            time.sleep(3)
        except StaleElementReferenceException:
            print("Tree refreshed. Refetching...")
            log_message(log_file, "Tree refreshed. Refetching...")
            # index += 1
            continue
        dates = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//button[@data-ofsc-id='toolbar-date-button-value']//span[@class='app-button-title'])[1]")))
        raw_date = dates.text.strip()
        print(f"Current Date Displayed: {raw_date}")
        log_message(log_file, f"Current Date Displayed: {raw_date}")

        clean_date = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', raw_date)

        ui_datetime = datetime.strptime(clean_date, "%A, %B %d, %Y")
        ui_date = ui_datetime.date()  
        ui_date_str = ui_datetime.strftime("%m/%d/%Y") 

        print(f"UI Date (logic): {ui_date}")
        log_message(log_file, f"UI Date (logic): {ui_date}")
        print(f"UI Date (display): {ui_date_str}")
        log_message(log_file, f"UI Date (display): {ui_date_str}")

        today = datetime.today().date()
        yesterday = today - timedelta(days=1)
        yesterday_str = yesterday.strftime("%m/%d/%Y")
        print(f"yesterday date:{yesterday_str}")
        log_message(log_file, f"yesterday date:{yesterday_str}")
        if before_8am:
            if ui_date == today:
                print("Dates match.")
                log_message(log_file, "Today Dates match.")
                click_left_arrow = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//span[contains(@class,'oj-ux-ico-chevron-left')])[1]")))
                click_left_arrow.click()
                if ui_date - timedelta(days=1) == yesterday:
                    print("Yesterday's date confirmed.")
                    log_message(log_file, "Yesterday's date confirmed.")
                    # ### View ####
                    time.sleep(5)
                    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//span[@data-ofsc-role='button-text' and normalize-space()='View']/ancestor::button)[1]")))
            elif ui_date == yesterday:
                print("UI already shows yesterday.")
                log_message(log_file, "UI already shows yesterday.")
        else:
            print("After 8 AM -> Using today.")
            log_message(log_file, "After 8 AM -> Using today.")
            if ui_date == yesterday:
                print("UI showing yesterday but after 8 AM. No arrow click required (as per your logic).")
                log_message(log_file, "UI showing yesterday but after 8 AM. No arrow click required (as per your logic).")

            elif ui_date == today:
                print("UI already shows today.")
                log_message(log_file, "UI already shows today.")
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
            log_message(log_file, "Apply Hierarchically Checkbox is not selected. Clicking it now.")
            apply_label.click()
        else:
            print("Checkbox is already selected. No action needed.")
            log_message(log_file, "Apply Hierarchically Checkbox is is already selected. No action needed.")
        click_apply = driver.find_element(By.XPATH, "(//button[.//span[@data-ofsc-role='button-text' and normalize-space()='Apply']])[2]")
        click_apply.click()
        time.sleep(2)
        #### Actions #####
        actions = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Actions']]")))
        actions.click()
        time.sleep(2)
        print("Actions clicked")
        log_message(log_file, "Actions clicked")
        export_excel = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@role='dialog' and @aria-label='Actions']//button[.//span[normalize-space()='Export']]")))
        export_excel.click()
        time.sleep(2)
        print("Export to Excel clicked")
        log_message(log_file, "Export to Excel clicked")
        print(f"\033[92mFile downloaded for area: {label} ({row_id})\033[0m")
        log_message(log_file, f"\033[92mFile downloaded for area: {label} ({row_id})\033[0m")
        print("---------------------------------------------------")
        log_message(log_file, "---------------------------------------------------")
        index += 1
    #### Go to Validate Excel Process ####
    validate_excel(driver, mode, COX_FILE, target_date, credentials, download_folder, log_file, previous_date, current_hour)
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
        log_message(log_file, "Mail sent Successfully for login failed.")
        print("\033[96mMail sent Successfully for login failed.\033[0m")
    except Exception:
        log_message(log_file, "Mail not sent for login failed.")