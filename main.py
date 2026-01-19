import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
import os



BASE_URL='https://gse.com.gh'
PAGE_URL=f'{BASE_URL}/trading-and-data/'



def get_yesterday_date():
    #Calculate yesterday's date in appropriate format
    yesterday = datetime.now() - timedelta(days=1)
    return yesterday


def wait_for_table_update(driver, table, target_date, timeout=25):
    WebDriverWait(driver, timeout).until(lambda d: target_date in table.text)


def start_driver():
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service)



def get_table(driver,yesterday):
    # get scrapped table after setting filter inputs
    get_from_date_input = driver.find_element(by=By.ID, value="table_1_range_from_1")
    get_to_date_input = driver.find_element(by=By.ID, value="table_1_range_to_1")
    get_show_length = driver.find_element(by=By.NAME, value="table_1_length")
    select = Select(get_show_length)
    get_to_date_input.send_keys(yesterday)
    get_from_date_input.send_keys(yesterday)
    select.select_by_visible_text('All')
    return  driver.find_element(by=By.ID, value="table_1")

def extract_table_data(table):
    headers = table.find_elements(by=By.TAG_NAME, value="th")
    column_names = [header.text for header in headers if header.text]

    # Extract row data
    rows = table.find_elements(by=By.TAG_NAME, value="tr")
    data = []
    for row in rows:
        cells = row.find_elements(by=By.TAG_NAME, value="td")
        if cells:
            row_data = [cell.text for cell in cells]
            data.append(row_data)
    return [data, column_names]

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
    yesterday = get_yesterday_date().strftime('%d/%m/%Y')
    driver=start_driver()
    if driver:
        wait = WebDriverWait(driver, 30)
        try:
            driver.get(PAGE_URL)
            table = get_table(driver,yesterday)
            wait_for_table_update(driver, table, yesterday)
            data, column_names = extract_table_data(table)
            save_data_to_excel(data, column_names, yesterday)
        finally:
            driver.quit()
        

if __name__ == "__main__":
    main()
