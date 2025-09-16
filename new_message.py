import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk, filedialog, messagebox, Button, Label

# ------------------------
# Global variables
# ------------------------
CHROMEDRIVER_PATH = r"C:\Users\ahtis\OneDrive\Documents\WebDriver\chromedriver.exe"
EXCEL_FILE_PATH = None

# ------------------------
# GUI Functions
# ------------------------
def browse_file():
    global EXCEL_FILE_PATH
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        EXCEL_FILE_PATH = file_path
        file_label.config(text=f"Selected: {file_path}")
        send_button.config(state="normal")

def send_whatsapp_messages():
    if not EXCEL_FILE_PATH:
        messagebox.showwarning("No file", "Please select an Excel file first!")
        return

    # Read contacts from Excel
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, dtype=str)
        df.fillna('', inplace=True)
        contacts = df.to_dict(orient='records')
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel: {e}")
        return

    # Setup Selenium Chrome
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # Open WhatsApp Web
    driver.get("https://web.whatsapp.com/")
    messagebox.showinfo("QR Scan", "Please scan the QR code in WhatsApp Web and then click OK here.")

    sent_count = 0
    failed_numbers = []

    for contact in contacts:
        name = contact.get("Name", "")
        number = contact.get("Phone Number", "")
        message = f"Hello {name}, this is a test message."

        if not number:
            continue

        try:
            # Open chat using wa.me link
            url = f"https://web.whatsapp.com/send?phone={number}&text={message}"
            driver.get(url)
            # Wait for chat box to appear
            send_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="compose-btn-send"]'))
            )
            send_btn.click()
            sent_count += 1
            time.sleep(2)  # small delay between messages
        except Exception as e:
            failed_numbers.append(number)
            continue

    driver.quit()

    # Summary
    summary = f"Messages sent: {sent_count}\nFailed numbers: {len(failed_numbers)}"
    if failed_numbers:
        summary += "\nFailed numbers:\n" + "\n".join(failed_numbers)

    messagebox.showinfo("WhatsApp Messaging Completed", summary)

# ------------------------
# GUI Layout
# ------------------------
root = Tk()
root.title("WhatsApp Bulk Messaging")
root.geometry("500x250")

Label(root, text="Step 1: Upload your contacts Excel file").pack(pady=10)
Button(root, text="Browse Excel", command=browse_file).pack(pady=5)
file_label = Label(root, text="No file selected")
file_label.pack(pady=5)

Label(root, text="Step 2: Send WhatsApp Messages").pack(pady=20)
send_button = Button(root, text="Start Sending Messages", command=send_whatsapp_messages, state="disabled", bg="green", fg="white")
send_button.pack(pady=5)

root.mainloop()
