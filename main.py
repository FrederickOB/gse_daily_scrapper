import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import StaleElementReferenceException
import os



BASE_URL='https://gse.com.gh'
PAGE_URL=f'{BASE_URL}/trading-and-data/'



def get_previous_trading_date():
    #Calculate yesterday's date in appropriate format
    today = datetime.now().date()
    if today.weekday() == 0:  # Monday
        previous_trading_day = today - timedelta(days=3)
    elif today.weekday() >=5:  # Weekend
        previous_trading_day = today - timedelta(days=today.weekday() - 4)
    else:
        previous_trading_day = today - timedelta(days=1)
    return previous_trading_day



def wait_for_table_update(driver, target_date, timeout=25):
    """Re-find the table on each poll — DataTables replaces the DOM after filter changes."""
    def table_contains_date(d):
        try:
            table = d.find_element(By.ID, "table_1")
            return target_date in table.text
        except StaleElementReferenceException:
            return False

    WebDriverWait(driver, timeout).until(table_contains_date)
    return driver.find_element(By.ID, "table_1")


def start_driver():
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service)



def get_table(driver, yesterday):
    wait = WebDriverWait(driver, 10)
    from_input = wait.until(EC.presence_of_element_located((By.ID, "table_1_range_from_1")))
    to_input = driver.find_element(By.ID, "table_1_range_to_1")
    show_length = driver.find_element(By.NAME, "table_1_length")

    for inp in (from_input, to_input):
        inp.clear()
        inp.send_keys(yesterday)

    Select(show_length).select_by_visible_text("All")
    return wait_for_table_update(driver, yesterday)

def extract_table_data(driver, table):
    for _ in range(3):
        try:
            headers = table.find_elements(by=By.TAG_NAME, value="th")
            column_names = [header.text for header in headers if header.text]

            rows = table.find_elements(by=By.TAG_NAME, value="tr")
            data = []
            for row in rows:
                cells = row.find_elements(by=By.TAG_NAME, value="td")
                if cells:
                    data.append([cell.text for cell in cells])
            return data, column_names
        except StaleElementReferenceException:
            table = driver.find_element(By.ID, "table_1")
    raise StaleElementReferenceException("table_1 became stale while reading rows")

def save_data_to_excel(data, column_names, yesterday):
    filename = '/Users/mac/Desktop/data/gse_trading_data_latest.xlsx'
    new_df = pd.DataFrame(data, columns=column_names)
    
    if os.path.exists(filename):
        existing_df = pd.read_excel(filename)
        if yesterday not in existing_df['Daily Date'].values:
            df = pd.concat([existing_df, new_df], ignore_index=True)
            print(f'Appended {len(new_df)} new rows to existing data')
        else:
            df = existing_df
            print(f'Data for {yesterday} already exists. File unchanged.')
    else:
        df = new_df
        print(f'Created new file with {len(new_df)} rows')
    
    df.to_excel(filename, index=False)
    print(f'Data saved to {filename}')
    print(f'Total rows: {len(df)}')



def main():
    driver=start_driver()
    if driver:
        try:
            driver.get(PAGE_URL)
            target_date = get_previous_trading_date()
            while True:
                yesterday = target_date.strftime('%d/%m/%Y')
                table = get_table(driver, yesterday)
                data, column_names = extract_table_data(driver, table)

                if len(data) > 0:
                    save_data_to_excel(data, column_names, yesterday)
                    break

                print(f'No data found for {yesterday}. Retrying...')
                target_date -= timedelta(days=1)
    
        finally:
            driver.quit()
        

if __name__ == "__main__":
    main()
