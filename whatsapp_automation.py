from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import os
import pyautogui

# Constants
SCREENSHOT_FOLDER = r"D:\automation\screenshots"
CHROME_DRIVER_PATH = "D:/automation/chromedriver.exe"

def setup_driver():
    """Sets up and returns a Selenium WebDriver instance with required options."""
    chrome_options = Options()
    chrome_options.add_argument("--user-data-dir=D:/automation/Chrome")
    chrome_options.add_argument("--profile-directory=Profile 9")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    service = Service(CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # Open WhatsApp Web
    driver.get("https://web.whatsapp.com/")
    time.sleep(10)  # Initial loading time
    return driver

def search_contact(driver, contact_name):
    """Searches and selects a WhatsApp contact or group efficiently."""
    search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
    search_box.click()
    search_box.send_keys(Keys.CONTROL + "a")  # Clear existing text
    search_box.send_keys(Keys.BACKSPACE)
    time.sleep(1)  # Allow UI update
    
    search_box.send_keys(contact_name)
    time.sleep(2)  # Reduced wait time
    search_box.send_keys(Keys.ENTER)
    time.sleep(2)

def open_attachment_menu(driver):
    """Opens the attachment menu (paperclip icon)."""
    attachment_xpath = "/html/body/div[1]/div/div/div[3]/div/div[4]/div/footer/div[1]/div/span/div/div[2]/div/div[1]/button"
    attachment_box = driver.find_element(By.XPATH, attachment_xpath)
    attachment_box.click()
    time.sleep(1.5)

def select_photos_videos(driver):
    """Clicks the 'Photos & Videos' option in the attachment menu."""
    photos_videos_xpath = "/html/body/div[1]/div/div/span[6]/div/ul/div/div/div[2]"
    photos_videos_button = driver.find_element(By.XPATH, photos_videos_xpath)
    photos_videos_button.click()
    time.sleep(2)  # Ensures the file selection window opens

def select_file(driver, keyword):
    """Selects a file based on the keyword using PyAutoGUI to simulate human behavior."""
    screenshots = [f for f in os.listdir(SCREENSHOT_FOLDER) if keyword.lower() in f.lower()]
    if not screenshots:
        print(f"No matching screenshots found for '{keyword}'.")
        return False

    latest_screenshot = max(screenshots, key=lambda f: os.path.getmtime(os.path.join(SCREENSHOT_FOLDER, f)))
    screenshot_path = os.path.join(SCREENSHOT_FOLDER, latest_screenshot)
    
    print(f"Attempting to send: {screenshot_path}")  # Debug log

    time.sleep(1.5)  # Ensure file selection window is ready
    pyautogui.write(screenshot_path)
    pyautogui.press('enter')
    time.sleep(2.5)  # Allow preview to load

    # Wait for send button and click it
    try:
    # Wait for the send button to be present and clickable
        send_button_xpath = "//div[@role='button' and @aria-label='Send']"
        send_button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, send_button_xpath))
    )
        time.sleep(1)
        send_button.click()
        time.sleep(2.5)
        print(f"Screenshot '{latest_screenshot}' sent successfully!")
        return True
    except Exception as e:
        print(f"[ERROR] Send button not found or not clickable for '{latest_screenshot}'. Error: {e}")
    return False



def send_state_screenshot(driver, state_name, contact_name):
    """Sends a screenshot to a specific WhatsApp group."""
    search_contact(driver, contact_name)
    open_attachment_menu(driver)
    select_photos_videos(driver)
    file_sent = select_file(driver, state_name)
    
    if not file_sent:
        print(f"Skipping {state_name} - No matching file found.")

# Main execution
driver = setup_driver()

# Example usage for multiple groups

# send_state_screenshot(driver, "Mohit", "Mohit Ho")

send_state_screenshot(driver, "Andhra Pradesh", "RAKSHA_AP_")

send_state_screenshot(driver, "Bihar", "RAKSHA_BR_")
send_state_screenshot(driver, "Chhattisgarh", "RAKSHA_CG_")
# send_state_screenshot(driver, "Daman", "RAKSHA_GJ_")
send_state_screenshot(driver, "Gujarat", "RAKSHA_GJ_")

send_state_screenshot(driver, "Jharkhand", "RAKSHA_JH_")
send_state_screenshot(driver, "Karnataka", "RAKSHA_KA_")
send_state_screenshot(driver, "Kerala", "RAKSHA_KL_")
send_state_screenshot(driver, "Madhya Pradesh", "RAKSHA_MP_")
send_state_screenshot(driver, "Maharashtra", "RAKSHA_MH_")

send_state_screenshot(driver, "Punjab", "RAKSHA_PB_")
send_state_screenshot(driver, "Rajasthan", "RAKSHA_RJ_")
send_state_screenshot(driver, "Tamil Nadu", "RAKSHA_TN_")
send_state_screenshot(driver, "Telangana", "HO - Telangana")

send_state_screenshot(driver, "Uttar Pradesh", "RAKSHA_UP_")
send_state_screenshot(driver, "Uttarakhand", "RAKSHA_UK_")
send_state_screenshot(driver, "West Bengal", "RAKSHA_WB_")


print("Keeping the browser open...")
input("Press Enter to close the browser...")
driver.quit()