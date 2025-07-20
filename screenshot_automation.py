import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
from PIL import Image
from io import BytesIO
from selenium.common.exceptions import StaleElementReferenceException

class PowerBILoginAutomation:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--no-default-browser-check")

        # Disable the "Chrome is being controlled" banner
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)

        # Use correct Chrome profile
        chrome_options.add_argument("--user-data-dir=D:/automation/Chrome")
        chrome_options.add_argument("--profile-directory=Profile 9")

        service = Service("D:/automation/chromedriver.exe")
        self.driver = webdriver.Chrome(service=service, options=chrome_options)

        # Open Power BI report
        self.driver.get("https://app.powerbi.com/groups/me/reports/33083153-ca75-4fbf-895f-f3f52c4c983b/aaabcac832e850de1711?ctid=83c91c40-3ff2-40f5-b552-47fd1bcfd47f&experience=power-bi")
        
        # Wait for the page to load
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

    def click_reset_button(self):
        reset_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-role-button-fill"))
        )
        reset_button.click()
        time.sleep(2)
        print("Reset Button Clicked!")

    def select_checkbox(self, state_name):
        try:
            print(f"Selecting checkbox for {state_name}...")

            # Find the state container dynamically using title attribute
            state_xpath = f"//div[@class='slicerItemContainer' and @title='{state_name}']"
            state_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, state_xpath))
            )

            # Scroll into view
            self.driver.execute_script("arguments[0].scrollIntoView();", state_element)
            time.sleep(1)

            # Find and click the checkbox inside the state container
            checkbox = state_element.find_element(By.CLASS_NAME, "slicerCheckbox")
            checkbox.click()
            time.sleep(2)

            print(f"Checkbox for {state_name} selected successfully!")
            self.take_screenshot(state_name)

        except Exception as e:
            print(f"Error selecting checkbox for {state_name}: {e}")

        

    def select_latest_date(self):
        try:
            print("Selecting latest date...")

            # Let slicer render properly
            time.sleep(5)

            # XPath to open the date slicer dropdown
            date_slicer_arrow_xpath = "/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[2]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/report/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[1]/transform/div/div[3]/div/div/visual-modern/div/div/div[2]/div"
            
            # Open the date slicer dropdown
            slicer = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, date_slicer_arrow_xpath))
            )
            self.driver.execute_script("arguments[0].click();", slicer)

            time.sleep(2)

            # XPath for the first date option (top-most row)
            first_date_xpath = "/html/body/div[8]/div[1]/div/div[2]/div/div[1]/div/div/div[1]/div"

            first_date = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, first_date_xpath))
            )
            self.driver.execute_script("arguments[0].scrollIntoView();", first_date)
            self.driver.execute_script("arguments[0].click();", first_date)

            print("First (latest) date selected.")

            time.sleep(2)

        except Exception as e:
            print(f"Error selecting latest date: {e}")



    
    def close_date_slicer(self):
        try:
            date_slicer_xpath = "/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/report/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[1]/transform/div/div[3]/div/div/visual-modern/div/div/div[2]/div/i"
            
            # Wait until the element is present and clickable
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, date_slicer_xpath))
            )

            # Use JavaScript to click the date slicer element
            element = self.driver.find_element(By.XPATH, date_slicer_xpath)
            self.driver.execute_script("arguments[0].click();", element)
            
            time.sleep(2)
            print("Date slicer dropdown closed.")
        
        except Exception as e:
            print(f"Error closing the date slicer dropdown: {e}")

    def take_screenshot(self, state_name):
        try:
            screenshot_folder = "D:/automation/screenshots"
            date_stamp = time.strftime("%d-%m-%Y")
            screenshot_path = f"{screenshot_folder}/{state_name}_DCR_{date_stamp}.png"

            x, y, width, height = 470, 170, 1110, 660
            screenshot = self.driver.get_screenshot_as_png()
            image = Image.open(BytesIO(screenshot))
            cropped_image = image.crop((x, y, x + width, y + height))
            cropped_image.save(screenshot_path)

            print(f"Screenshot saved: {screenshot_path}")
    
        except Exception as e:
            print(f"Error while taking screenshot: {e}")

    def keep_browser_open(self):
        print("Browser will remain open for inspection.")
        input("Press Enter to close the browser...")

if __name__ == "__main__":
    automation = PowerBILoginAutomation()
    automation.click_reset_button()
    automation.select_latest_date()
    automation.close_date_slicer()
    
    states = ["Andhra Pradesh","Bihar","Chhattisgarh","Daman","Gujarat",
              "Jharkhand","Karnataka","Kerala","Madhya Pradesh",
              "Maharashtra","Punjab","Rajasthan","Tamil Nadu","Telangana",
              "Uttar Pradesh","Uttarakhand","West Bengal"]  # Modify this list to include more states if needed
    
    for state in states:
        automation.select_checkbox(state)
    
    automation.keep_browser_open()