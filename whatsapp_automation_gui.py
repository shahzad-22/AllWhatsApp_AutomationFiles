import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

# ------------------------
# FUNCTIONS
# ------------------------

def browse_file():
    global EXCEL_FILE_PATH
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        EXCEL_FILE_PATH = file_path
        file_label.config(text=f"Selected: {file_path}")

def send_messages():
    if not EXCEL_FILE_PATH:
        messagebox.showwarning("No File", "Please select an Excel file first!")
        return

    # ------------------------
    # Load contacts from Excel
    # ------------------------
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, dtype=str)
        contacts = df['Name'].tolist()  # Assuming 'Name' column
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel: {e}")
        return

    # ------------------------
    # Set up Selenium
    # ------------------------
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

    service = Service(r"C:\Users\ahtis\OneDrive\Documents\WebDriver\chromedriver-win64\chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get("https://web.whatsapp.com")
    input("📱 Please scan QR code in browser and press ENTER to continue...")

    # ------------------------
    # Send messages
    # ------------------------
    MESSAGE_TEXT = "Hello! This is a test message from automation."
    total_sent = 0
    total_failed = 0

    def send_whatsapp_message(contact, message, retries=2):
        for attempt in range(retries):
            try:
                search_box = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
                )
                search_box.clear()
                search_box.click()
                time.sleep(0.5)
                search_box.send_keys(contact)
                time.sleep(2)
                search_box.send_keys(Keys.ENTER)

                message_box = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
                )
                driver.execute_script("arguments[0].focus();", message_box)
                message_box.clear()
                message_box.send_keys(message)
                message_box.send_keys(Keys.ENTER)

                print(f"✅ Message sent to: {contact}")
                time.sleep(1.5)
                return True

            except (StaleElementReferenceException, TimeoutException):
                print(f"⚠️ Attempt {attempt + 1} failed for {contact} — retrying...")
                time.sleep(1)
                continue
            except Exception as e:
                print(f"❌ Could not send message to: {contact} — Reason: {e}")
                return False
        print(f"❌ Failed to send message to: {contact} after {retries} attempts")
        return False

    for contact in contacts:
        success = send_whatsapp_message(contact, MESSAGE_TEXT)
        if success:
            total_sent += 1
        else:
            total_failed += 1

    # ------------------------
    # Summary
    # ------------------------
    print("\n📊 --- SUMMARY ---")
    print(f"✅ Total messages sent: {total_sent}")
    print(f"❌ Total messages failed: {total_failed}")
    driver.quit()
    messagebox.showinfo("Done", f"Messages sent: {total_sent}\nFailed: {total_failed}")

# ------------------------
# GUI
# ------------------------
EXCEL_FILE_PATH = None

root = tk.Tk()
root.title("WhatsApp Automation")

tk.Label(root, text="Step 1: Upload your contacts Excel file").pack(pady=5)
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)
file_label = tk.Label(root, text="No file selected")
file_label.pack(pady=5)

tk.Label(root, text="Step 2: Send WhatsApp Messages").pack(pady=20)
tk.Button(root, text="Send Messages", command=send_messages, bg="green", fg="white").pack(pady=5)

root.mainloop()
