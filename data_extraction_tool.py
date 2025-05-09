import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

def init_excel(file_name):
    try:
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Company Name", "Phone Number", "Email ID"])
    return wb, ws

def extract_data_from_google_maps(gmap_url):
    # Set up headless Chrome
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    driver.get(gmap_url)
    time.sleep(5)  # wait for data to load

    try:
        company_name = driver.find_element(By.CLASS_NAME, 'DUwDvf').text
    except:
        company_name = "N/A"

    try:
        phone = driver.find_element(By.XPATH, "//button[@data-tooltip='Copy phone number']").text
    except:
        phone = "N/A"

    try:
        # Most businesses don't list email on Google Maps; kept for structure
        email = "N/A"
    except:
        email = "N/A"

    driver.quit()
    return company_name, phone, email

def save_data_to_excel(file_name, data):
    wb, ws = init_excel(file_name)
    ws.append(data)
    wb.save(file_name)

def main():
    print("=== Google Maps Data Extractor ===")
    gmap_url = input("Enter the Google Maps business URL: ").strip()
    file_name = "company_records.xlsx"

    print("\nExtracting data...")
    company_name, phone, email = extract_data_from_google_maps(gmap_url)
    print(f"Name: {company_name}\nPhone: {phone}\nEmail: {email}")

    save_data_to_excel(file_name, (company_name, phone, email))
    print(f"Data saved to {file_name} successfully.")

if __name__ == "__main__":
    main()
