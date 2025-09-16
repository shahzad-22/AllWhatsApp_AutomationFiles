import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Load contact list from Excel
contacts = pd.read_excel("contacts.xlsx", dtype=str)

# Set up Selenium
driver = webdriver.Chrome()
driver.get("https://web.whatsapp.com")

print("🔒 Please scan the QR code with your phone to log in to WhatsApp Web.")
input("✅ Press ENTER once you’ve scanned the QR code...")

# Loop through each contact
for index, row in contacts.iterrows():
    name = row['Name']
    number = row['Number']
    message = f"Hi {name}, if you're buying or selling a plot, contact us today."

    # Open chat with the number
    link = f"https://web.whatsapp.com/send?phone={number}&text={message}"
    driver.get(link)

    time.sleep(8)  # Wait for chat to load

    try:
        # Click the send button
        send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_button.click()
        print(f"✅ Message sent to {name} ({number})")
    except Exception as e:
        print(f"❌ Failed to send message to {name} ({number}) — {e}")

    time.sleep(5)  # Short delay between messages

print("🎉 All messages sent.")
driver.quit()
