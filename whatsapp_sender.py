import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

# ------------------------
# CONFIGURATION
# ------------------------
CHROMEDRIVER_PATH = r"C:\Users\ahtis\OneDrive\Documents\WebDriver\chromedriver-win64\chromedriver.exe"
EXCEL_FILE_PATH = r"C:\Users\ahtis\OneDrive\Documents\whatsapp_automation\contacts.xlsx"
MESSAGE_TEXT = "Hello! This is a test message from automation."

# ------------------------
# SETUP CHROME OPTIONS
# ------------------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
input("📱 Please scan QR code in browser and press ENTER to continue...")

# ------------------------
# LOAD CONTACTS
# ------------------------
df = pd.read_excel(EXCEL_FILE_PATH)
contacts = df['Name'].tolist()  # Assuming 'Name' column in Excel

# ------------------------
# HELPER FUNCTION WITH RETRY
# ------------------------
def send_whatsapp_message(contact, message, retries=2):
    for attempt in range(retries):
        try:
            # SEARCH CONTACT
            search_box = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
            )
            search_box.clear()
            search_box.click()
            time.sleep(0.5)
            search_box.send_keys(contact)
            time.sleep(2)
            search_box.send_keys(Keys.ENTER)

            # RE-FIND MESSAGE BOX
            message_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )

            # FORCE FOCUS & SEND MESSAGE
            driver.execute_script("arguments[0].focus();", message_box)
            message_box.clear()
            message_box.send_keys(message)
            message_box.send_keys(Keys.ENTER)

            print(f"✅ Message sent to: {contact}")
            time.sleep(1.5)  # mimic human behavior
            return True

        except (StaleElementReferenceException, TimeoutException) as e:
            print(f"⚠️ Attempt {attempt + 1} failed for {contact} — retrying...")
            time.sleep(1)
            continue
        except Exception as e:
            print(f"❌ Could not send message to: {contact} — Reason: {e}")
            return False
    print(f"❌ Failed to send message to: {contact} after {retries} attempts")
    return False

# ------------------------
# SEND MESSAGES TO ALL CONTACTS
# ------------------------
total_sent = 0
total_failed = 0

for contact in contacts:
    success = send_whatsapp_message(contact, MESSAGE_TEXT)
    if success:
        total_sent += 1
    else:
        total_failed += 1

# ------------------------
# SUMMARY
# ------------------------
print("\n📊 --- SUMMARY ---")
print(f"✅ Total messages sent: {total_sent}")
print(f"❌ Total messages failed: {total_failed}")

driver.quit()
