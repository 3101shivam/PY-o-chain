import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

import write_to_g_sheet

from selenium.webdriver.support.ui import Select
def nse_url():
    driver = get_chrome_driver()
    sleep_program_for(6)

    # action = ActionChains(driver)
    # expiry_selection = driver.find_element_by_xpath('//*[@id="expirySelect"]/option[03]').click()
    # sleep_program_for(10)
    # action.move_to_element(expiry_selection).perform()
    #action = ActionChains(driver)
    driver.find_element_by_xpath('//*[@id="expirySelect"]/option[03]').click()
    #action.move_to_element(expiry_selection).perform()

    # select = Select(driver.find_element_by_id('expirySelect'))
    # select.select_by_index(3)



    sleep_program_for(20)
    table_rows = get_table_from_html(driver)

    row_list = get_row_list(table_rows)

    equity_derivatives_table = get_data_in_df(row_list)
    # print(equity_derivatives_table)
    file_name = "Equity.xlsx"
    delete_if_file_exists(file_name)
    df = equity_derivatives_table["CHNG IN OI"]
    df_oi = equity_derivatives_table["OI"]
    df = df.replace("-", 0)
    df = df.replace(",", "", regex=True)
    df_oi = df_oi.replace("-", 0)
    df_oi = df_oi.replace(",", "", regex=True)
    current_row_number = add_to_excel(df, df_oi)
    return current_row_number


def add_to_excel(df, df_oi):
    import math
    first_column_sum = pd.to_numeric(df.iloc[:, 0]).sum()
    second_column_sum = pd.to_numeric(df.iloc[:, 1]).sum()
    if math.isnan(first_column_sum):
        first_column_sum = 0
    if math.isnan(second_column_sum):
        second_column_sum = 0
    first_oi_column_sum = pd.to_numeric(df_oi.iloc[:, 0]).sum()
    second_oi_column_sum = pd.to_numeric(df_oi.iloc[:, 1]).sum()
    if math.isnan(first_oi_column_sum):
        first_oi_column_sum = 0
    if math.isnan(second_oi_column_sum):
        second_oi_column_sum = 0
    excel_dict = dict = {'Date': [todays_date()],
                         'Time': [current_time().time()],
                         'Index / Stock': 'NIFTY',
                         'Put Write (Teji Wale)': second_column_sum,
                         'Call Write ( Mandi Wale)': first_column_sum,
                         'Diff (P-C)': second_column_sum - first_column_sum,
                         'Diff (C-P)': first_column_sum - second_column_sum,
                         'P/C (Teji Jada)': second_column_sum / first_column_sum,
                         'c/p (Mandi Jyada)': first_column_sum / second_column_sum,
                         'Put Total OI': second_oi_column_sum,
                         'Call Total OI': first_oi_column_sum,
                         'P-C OI': second_oi_column_sum - first_oi_column_sum,
                         'C-P OI': first_oi_column_sum - second_oi_column_sum,
                         'P/C OI': second_oi_column_sum / first_oi_column_sum,
                         'c/p OI': first_oi_column_sum / second_oi_column_sum
                         }
    excel_df = pd.DataFrame(excel_dict)
    excel_df.columns = ['Date', 'time', 'Index', 'Put Write (Teji Wale)', 'Call Write ( Mandi Wale)', 'Diff (P-C)',
                        'Diff (C-P)', 'P/C (Teji Jada)', 'c/p (Mandi Jyada)', 'Put Total OI', 'Call Total OI', 'P-C OI',
                        'C-P OI', 'P/C OI', 'c/p OI']
    from tabulate import tabulate
    print(tabulate(excel_df, headers='firstrow'))

    current_row_number = write_to_g_sheet.main(row_num, excel_df)

    return current_row_number


def delete_if_file_exists(file_name):
    import os
    if os.path.exists(file_name):
        os.remove(file_name)
        print("existing file deleted")
        return


def get_data_in_df(row_list):
    # Creating the table using pandas
    equity_derivatives_table = pd.DataFrame(row_list, columns=["", "OI", "CHNG IN OI", "VOLUME", "IV", "LTP",
                                                               "CHNG", "BID QTY", "BID PRICE", "ASK PRICE",
                                                               "ASK QTY", "STRIKE PRICE", "BID QTY", "BID PRICE",
                                                               "ASK PRICE", "ASK QTY", "CHNG", "LTP", "IV", "VOLUME",
                                                               "CHNG IN OI", "OI", ""])
    equity_derivatives_table = equity_derivatives_table.drop([0], axis=0)
    return equity_derivatives_table


def get_row_list(table_rows):
    l = []
    for tr in table_rows:
        td = tr.find_all('td')
        row = [tr.text for tr in td]
        l.append(row)
    return l


def get_table_from_html(driver):
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    # soup = BeautifulSoup(html,'lxml')
    result = soup.find('table', {'class': 'common_table'})
    table_rows = result.find_all('tr')
    return table_rows


def sleep_program_for(time_in_sec):
    import time
    time.sleep(time_in_sec)


def get_chrome_driver():
    url = "https://www.nseindia.com/option-chain"

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument(
        f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36')
    driver = webdriver.Chrome('./chromedriver', options=chrome_options)
    driver.get(url)
    return driver


def todays_date():
    from datetime import date
    today = date.today()
    return today.strftime("%d/%m/%Y")


def current_time():
    from datetime import datetime
    now = datetime.now()
    return now


row_num = 2

if __name__ == "__main__":
    count = 1
    current_row_number = 0;
    for i in range(100):
        while True:
            try:
                current_row_number = nse_url()
            except Exception as e:
                print("something went wrong retry initiated ... " + str(i))
                print(e)
                #row_num = current_row_number + 1
                i = i + 1
                if i > 100:
                    exit(0)
                continue
            sleep_program_for(150)
            row_num += count
