import jieba
import pandas as pd
import os
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui


# "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Default" --remote-debugging-port=9222

with open(r"C:\Users\underdog\PycharmProjects\fami711\fami_goods.txt", 'r', encoding='utf-8') as f:
    content = f.read()

jieba.load_userdict(r"C:\Users\underdog\PycharmProjects\fami711\fami_goods.txt")


def initialize_article_list():
    return []


def initialize_driver():
    options = Options()
    options.add_experimental_option("debuggerAddress", "localhost:9222")
    s = Service(r"C:\Users\underdog\PycharmProjects\chromedriver-win64\chromedriver.exe")
    driver = webdriver.Chrome(service=s, options=options)
    return driver


# https://www.facebook.com/groups/261728378573604
def fami711_crawler(driver, article_list):
    time.sleep(3)
    driver.get('https://www.facebook.com/groups/261728378573604')
    time.sleep(5)
    i = 1
    counter = 0
    max_count = 30
    while True:
        i += 1
        xpath_locator = (f'//*[@id="mount_0_0_O1"]/div/div[1]/div/div[3]/div/div/div/div[1]/div[1]/div[2]'
                         f'/div/div/div/div[4]/div/div[2]/div/div/div/div[2]/div[2]/div[{i}]'
                         f'/div/div/div/div/div/div/div/div/div/div[8]/div/div/div[3]/div[1]')
        while True:
            try:
                element = WebDriverWait(driver, 0.5).until(
                    EC.presence_of_element_located((By.XPATH, xpath_locator))
                )
                break

            except:
                # 滾動並再次檢查
                pyautogui.click(1908, 1018)
                time.sleep(0.5)

        try:
            Text = element.text
            article_list.append(Text)

            counter += 1
            if counter >= max_count:
                print("Reached the maximum count. Exiting the loop.")
                break

        except Exception as e:
            print("Error encountered:", e)
            break


def segment_articles(articles):
    segmented_articles = []
    for article in articles:
        seg_list = jieba.cut(article, cut_all=False)
        segmented_articles.append(list(seg_list))
    return segmented_articles


def print_segmented_results(segmented_articles):
    for i, seg_article in enumerate(segmented_articles):
        print(f"文章 {i+1} 分詞結果：", seg_article)


def save_to_excel(segmented_articles):
    max_length = max(len(article) for article in segmented_articles)
    columns = [f"詞 {i + 1}" for i in range(max_length)]
    df = pd.DataFrame(segmented_articles, columns=columns)
    excel_filename = "segmented_results.xlsx"
    df.to_excel(excel_filename, index=False, startcol=1)

    # A1命名
    book = openpyxl.load_workbook(excel_filename)
    sheet = book.active
    sheet['A1'] = "貼文編號"
    book.save(excel_filename)

    print(f"分詞結果已保存到 {excel_filename}")
    os.startfile(excel_filename)


def main():
    driver = initialize_driver()
    articles = initialize_article_list()
    fami711_crawler(driver, articles)
    segmented_articles = segment_articles(articles)
    print_segmented_results(segmented_articles)
    save_to_excel(segmented_articles)


if __name__ == "__main__":
    main()