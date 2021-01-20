import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options

import time
import datetime

import pandas as pd
import openpyxl


def init_driver():
    """
    =======================================================================
    Def Decription          : 셀레니움 드라이버 이니셜라이징
    =======================================================================
    """
    try:
        # path
        ROOT_PATH = "C:/Users/JLK/web_project/AS21/data_scrapying"
        DRIVER_PATH = "/driver/chromedriver.exe"

        # option
        WINDOW_SIZE = "1920,1080"
        chrome_options = Options()
        # chrome_options.add_argument( "--headless" )     # 크롬창이 열리지 않음
        # GUI를 사용할 수 없는 환경에서 설정, linux, docker 등
        chrome_options.add_argument("--no-sandbox")
        # GUI를 사용할 수 없는 환경에서 설정, linux, docker 등
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument(f"--window-size={ WINDOW_SIZE }")
        chrome_options.add_argument(
            'Content-Type=application/json; charset=utf-8')

        driver = webdriver.Chrome(
            executable_path=ROOT_PATH + DRIVER_PATH, chrome_options=chrome_options)

    except Exception as ex:
        response_object = {
            "status": "fail",
            "message": str(ex),
            "driver_object": None
        }
        return response_object

    response_object = {
        "status": "success",
        "message": "driver_initialized",
        "driver_object": driver
    }
    return response_object


def close_driver(driver):
    """
    =======================================================================
    Def Decription          : 셀레니움 드라이버 클로징
    =======================================================================
    """
    try:
        driver.close()

    except Exception as ex:
        response_object = {
            "status": "fail",
            "message": str(ex)
        }
        return response_object

    response_object = {
        "status": "success",
        "message": "driver_closed"
    }
    return response_object

# * should find more pattern


def main():
    # init driver
    response_object = init_driver()
    if response_object["status"] != "success":
        print(response_object)

    driver = response_object["driver_object"]

    # go lucid
    driver.get("https://lucidtracking.com/#/dashboard")
    time.sleep(5)

    # login lucid
    # id
    id_xpath = "/html/body/div[1]/div/div[1]/div/div[1]/input"
    my_id = "sdi654321"
    driver.find_element_by_xpath(id_xpath).send_keys(my_id)

    # pw
    pw_xpath = "/html/body/div[1]/div/div[1]/div/div[2]/input"
    my_pw = "HLmpBQZ0ERfcd3xw"
    driver.find_element_by_xpath(pw_xpath).send_keys(my_pw)

    # login
    login_xpath = "/html/body/div[1]/div/div[1]/div/button"
    driver.find_element_by_xpath(login_xpath).click()
    time.sleep(10)

    # click tsla
    tsla_xpath = "/html/body/div[1]/div/div[4]/div[2]/div[1]/div[2]/div/div[1]/div/div/table/tbody[23]/tr/td[1]/b"
    element = driver.find_element_by_xpath(tsla_xpath)
    driver.execute_script("arguments[0].click();", element)
    time.sleep(10)

    # click collapsed
    collapsed_path = "/html/body/div[1]/div/div[2]/modal-directive/div/div[2]/div[2]/div[5]/div/h4"
    driver.find_element_by_xpath(collapsed_path).click()
    time.sleep(1)

    # find ARRK
    ARRK_index_list = []
    ARRK_data_list = []
    columns = ['date', 'share']

    for i in range(1, 31):

        # date
        ARRK_date_xpath = "/html/body/div[1]/div/div[2]/modal-directive/div/div[2]/div[2]/div[5]/div/div/table/tbody/tr[%s]/td[1]" % (
            i)
        ARRK_date = driver.find_element_by_xpath(ARRK_date_xpath).text

        # share
        ARRK_share_xpath = "/html/body/div[1]/div/div[2]/modal-directive/div/div[2]/div[2]/div[5]/div/div/table/tbody/tr[%s]/td[2]" % (
            i)
        ARRK_share = driver.find_element_by_xpath(ARRK_share_xpath).text

        ARRK_index_list.append(i)

        # make data row
        ARRK_data_row = []
        ARRK_data_row.append(ARRK_date)
        ARRK_data_row.append(ARRK_share)

        # put to list
        ARRK_data_list.append(ARRK_data_row)

    # save to excel
    df = pd.DataFrame(ARRK_data_list,
                      index=ARRK_index_list, columns=columns)

    now = str(datetime.datetime.now())
    now = now.replace("-", "_")
    now = now.replace(" ", "_")
    now = now.replace(":", "_")
    now = now.replace(".", "_")

    file_path = 'C:/Users/JLK/web_project/AS21/data_scrapying/excel_file/tsla_ARKK_%s.xlsx' % (
        now)

    df.to_excel(
        file_path, sheet_name='tsla')

    # close driver
    response_object = close_driver(driver)
