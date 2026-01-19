from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import time
import os

load_dotenv()

username = os.getenv("USER")
password = os.getenv("PASSWORD")
download_path = os.getenv("DOWNLOAD_PATH")
login_url = os.getenv("LOGIN_URL")
download_url = os.getenv("DOWNLOAD_URL")

options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

prefs = {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)

driver.get(login_url)
driver.find_element(By.ID, "login_id").send_keys(username)
driver.find_element(By.ID, "userpwd_v").send_keys(password)
driver.find_element(By.CLASS_NAME, "login_button").click()

time.sleep(3)

driver.get(download_url)

wait = WebDriverWait(driver, 5)

wait.until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe")))
box = wait.until(EC.presence_of_element_located((By.ID, "recent-box-0")))
driver.execute_script("arguments[0].click();", box)

time.sleep(3)

driver.switch_to.default_content()

wait.until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe")))
btn = wait.until(EC.presence_of_element_located((By.ID, "js-btn-run")))
driver.execute_script("arguments[0].click();", btn)

time.sleep(5)

xls_btn = driver.find_element(By.ID, "XLSbtn")
driver.execute_script("arguments[0].click();", xls_btn)

time.sleep(5)

for f in os.listdir(download_path):
    if f.startswith("BRANCH STOCK") and f.endswith(".xlsx"):
        new_name = "BRANCH STOCK.xlsx"
        os.replace(
            os.path.join(download_path, f), os.path.join(download_path, new_name)
        )

driver.close()
