from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

# Function to scrape data from a single page
def scrape_data(floor_code, page_num):
    # Initialize the Chrome driver
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (no GUI)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    table_data = []
    try:
        # Access the URL with specified floor_code and page number
        url = f"https://24hmoney.vn/companies?industry_code=all&floor_code={floor_code}&com_type=all&letter=all&page={page_num}"
        driver.get(url)

        # Wait for the page to load
        driver.implicitly_wait(10)

        # Find the table rows (excluding the header)
        rows = driver.find_elements(By.XPATH, '//div[@class="vue-table"]//tbody/tr')

        # Loop through rows and collect cell data
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, 'td')
            row_data = [
                cells[0].text,  # Company Code
                cells[1].text,  # Company Name
                cells[2].text,  # Sector
                cells[3].text,  # Floor
                cells[4].text,  # Volume
            ]
            table_data.append(row_data)

        # Print success message
        print(f"Successfully scraped page {page_num}")

    except Exception as e:
        print(f"Error on page {page_num}: {e}")
    finally:
        driver.quit()

    return table_data

# Function to scrape multiple pages and concatenate into a DataFrame
def scrape_multiple_pages(floor_code, num_pages):
    all_data = []

    for i in range(1, num_pages + 1):  # Loop through page numbers
        page_data = scrape_data(floor_code, i)
        all_data.extend(page_data)  # Add page data to the list

    # Convert to DataFrame
    columns = ['Company Code', 'Company Name', 'Sector', 'Floor', 'Volume']
    df = pd.DataFrame(all_data, columns=columns)
    return df

# Save DataFrame to Excel
def save_to_excel(df, filename):
    df.to_excel(filename, index=False)

# Example usage
floor_code = "HNX"  # Replace with your desired floor code
num_pages = 16  # Number of pages to scrape

# Scrape the data from multiple pages
df = scrape_multiple_pages(floor_code, num_pages)

# Save the concatenated data to an Excel file
save_to_excel(df, "HNX_stocks.xlsx")

print("Data successfully saved")
