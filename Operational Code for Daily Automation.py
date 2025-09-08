import pandas as pd
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

# -----------------------------------
# Setup
# -----------------------------------

# Path to Excel file
excel_file = "tickets_data.xlsx"

# Load existing data if available
if os.path.exists(excel_file):
    existing_df = pd.read_excel(excel_file)
else:
    existing_df = pd.DataFrame()

# Start browser
driver = webdriver.Chrome()
driver.get("https:")
#address has been redsacted for security purposes
# Login
driver.find_element(By.ID, "email").send_keys("insert email address")
#email address and password have been redacted for security purposes
driver.find_element(By.ID, "password").send_keys("insert password")
driver.find_element(By.XPATH, "//button[text()='Sign In']").click()

# Wait for dashboard to load
wait = WebDriverWait(driver, 10)
wait.until(EC.invisibility_of_element((By.CLASS_NAME, "loading-overlay")))

# Navigate to Tickets > All
tickets_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Tickets']")))
tickets_menu.click()

all_tickets_link = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//a[@href='https://maxprof.web.za/admin/tickets/all']/span[text()='All']"
)))
all_tickets_link.click()

# -----------------------------------
# Scrape First 2 Pages
# -----------------------------------
scraped_data = []

for page in range(2):  # Check only first 2 pages
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#DataTables-SS tbody tr")))
        table = driver.find_element(By.ID, "DataTables-SS")
        html = table.get_attribute("outerHTML")
        df = pd.read_html(html)[0]

        # Format 'Submitter' column to "Name Surname"
        if 'Submitter' in df.columns:
            df['Submitter'] = df['Submitter'].str.extract(r'^([A-Za-z]+\s[A-Za-z]+)')

        scraped_data.append(df)

        # Click 'Next' to go to next page
        next_button = driver.find_element(By.LINK_TEXT, "Next")
        if "disabled" in next_button.get_attribute("class"):
            break
        next_button.click()
        time.sleep(3)

    except Exception as e:
        print(f"❌ Error on page {page+1}: {e}")
        break

driver.quit()

# -----------------------------------
# Combine and Update Excel File
# -----------------------------------
if scraped_data:
    new_df = pd.concat(scraped_data, ignore_index=True)

    # Remove duplicates (compare by all columns)
    if not existing_df.empty:
        combined_df = pd.concat([new_df, existing_df], ignore_index=True)
        combined_df.drop_duplicates(inplace=True)
    else:
        combined_df = new_df

    # Save with new entries at the top
    combined_df.to_excel(excel_file, index=False)
    print(f" Excel file updated with new rows (top of sheet).")
else:
    print("No new data scraped.")

