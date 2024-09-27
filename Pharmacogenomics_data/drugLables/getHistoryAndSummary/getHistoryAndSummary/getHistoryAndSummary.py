import csv
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from openpyxl import Workbook, load_workbook
from selenium.webdriver.support import expected_conditions as EC

# Load the TSV file
tsv_file_path = r"C:\Users\vieir\data_playground\Pharmacogenomics_data\drugLables\drugLabels\drugLabels.tsv"

# Setup Selenium WebDriver
options = Options()
options.add_argument('--headless')  # Run in headless mode if you don't want a browser window
service = Service('path_to_your_chromedriver')  # Replace with your ChromeDriver path
driver = webdriver.Chrome()

# Prepare a list to hold the updated data
updated_rows = []

# Read the TSV file
with open(tsv_file_path, mode='r', encoding='utf-8') as file:
    reader = csv.reader(file, delimiter='\t')
    headers = next(reader)  # Get the headers
    
    # Only add 'Extracted Summary' header if it doesn't already exist
    if 'Extracted Summary' not in headers:
        headers.append('Extracted Summary')
    
    updated_rows.append(headers)  # Adding updated headers

        # Process each row
    for row in reader:
        # If the row already contains an 'Extracted Summary', skip extraction
        if len(row) > len(headers) - 1:
            updated_rows.append(row)  # Already has the summary, append without modification
            continue

        try:
            url = row[0]
            print(f"{url} + \n")
            driver.get(url)
        except Exception as e:
            print("Empty URL")
            summary = "N/A"  # Set summary to a default value if not found
            break

        # Extract the summary from the webpage
        try:
            summary = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#label-summary .summary-text"))
            ).text
        except Exception as e:
            print(f"Error processing {url}: {e}")
            summary = "N/A"  # Use default if there is an issue

        # Append extracted summary to the row
        row.append(summary)
        updated_rows.append(row)


# Close the driver
driver.quit()

# Write updated data back to a new TSV file
output_file_path = r"C:\Users\vieir\data_playground\Pharmacogenomics_data\drugLables\drugLabels\updated_drugLabels.tsv"
with open(output_file_path, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file, delimiter='\t')
    writer.writerows(updated_rows)

print(f"Data extraction complete. Updated data saved to {output_file_path}.")
