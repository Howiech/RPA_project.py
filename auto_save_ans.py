from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import pygetwindow as gw
import os
import openpyxl


# 於 cmd 輸入 "你的chrome 執行路徑" --remote-debugging-port=9222 如下
# "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Default" --remote-debugging-port=9222
#rr
# Windows 系統中於桌面捷徑點擊右鍵 - 內容 - 捷徑 - 目標, 目標欄位內即是你的chrome執行路徑
# 或是於chrome中輸入 chrome://version/  裡面的可執行檔路徑即是chrome的執行路徑

def initialize_driver():
    options = Options()
    options.add_experimental_option("debuggerAddress", "localhost:9222")
    s = Service(r"C:\Users\underdog\PycharmProjects\chromedriver-win64\chromedriver.exe")
    driver = webdriver.Chrome(service=s, options=options)
    return driver


def navigate_and_ask_openai(driver):
    time.sleep(2)
    driver.get('https://chat.openai.com/')
    time.sleep(5)
    length_q = len(question_list)

    for i in range(length_q):
        input_element = driver.find_element(By.XPATH, '//*[@id="prompt-textarea"]')
        input_element.send_keys(question_list[i])
        time.sleep(1)

        # button_xpath may change
        button = driver.find_element(By.XPATH,
                                     '//*[@id="__next"]/div/div/div[1]/div[2]/div/main/div/div[2]/form/div/div[2]/div/button')
        button.click()
        time.sleep(10)
        answer_xpath = f"//*[@id='__next']/div/div/div[1]/div[2]/div/main/div/div[1]/div/div/div/div[{(i + 1) * 2}]/div/div[2]/div[1]/div/div"
        # answer_xpath may change, but the pattern is                       "......./div/div/div/div[{(i + 1) * 2}]........"

        try:
            answer = driver.find_element(By.XPATH, answer_xpath).text
            print(answer)
            ans_list.append(answer)
            print(f'流程{i + 1}完成')
            time.sleep(5)
        except NoSuchElementException:
            print('再等15秒')
            time.sleep(15)
            answer = driver.find_element(By.XPATH, answer_xpath).text
            print(answer)
            ans_list.append(answer)
            print(f'流程{i + 1}完成')
            time.sleep(5)
    return question_list, ans_list


def dict_xl_output():
    # 建立dict目的為增加此框架可利用空間
    result_dict = dict(zip(question_list, ans_list))

    # create dataframe
    df_gpt = pd.DataFrame(list(result_dict.items()), columns=["question", "answer"])

    # write on Excel
    excel_filename = 'gpt4_interaction.xlsx'
    df_gpt.to_excel(excel_filename, Index=False)

    # open
    os.startfile(excel_filename)


def activate_chrome_9222():  # for test purpose if needed, not necessarily require
    # 獲取所有窗口
    windows = gw.getWindowsWithTitle('')

    for window in windows:
        if "新分頁 - Google Chrome" in window.title:  # base on your window's title
            window.activate()
            break


question_list = ['Hi', 'How are you', 'Have a good day']  # customize
ans_list = []


def main():
    driver = initialize_driver()
    # activate_chrome_9222()
    navigate_and_ask_openai(driver)
    dict_xl_output()


if __name__ == '__main__':
    main()

