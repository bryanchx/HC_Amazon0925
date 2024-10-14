from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time


def search_amazon(keywords):
    results = []

    # 初始化 ChromeDriver
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')  # 如果您希望无头模式，可以启用这一行
    driver = webdriver.Chrome(options=options)

    for keyword in keywords:
        url = f'https://www.amazon.co.uk/s?k={keyword}'
        driver.get(url)
        time.sleep(5)  # 等待页面加载

        # 查找产品标题
        product_titles = driver.find_elements(By.CSS_SELECTOR, 'span.a-size-medium.a-color-base.a-text-normal')

        for title in product_titles:
            results.append({'Keyword': keyword, 'Title': title.text})

    driver.quit()  # 关闭浏览器
    return results


# 示例关键词列表
keywords = ['FREE','DEALS','AMAZON','KIPLING','PANDORA','CHEAP','PRIME']  # 替换为您的关键词
results = search_amazon(keywords)

# 将结果保存到Excel
df = pd.DataFrame(results)
df.to_excel('amazon_search_results.xlsx', index=False)
