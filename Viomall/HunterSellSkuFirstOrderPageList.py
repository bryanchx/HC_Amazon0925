import os
import random

import requests
import pandas as pd
import time


class HunterSellSkuFirstOrderPageList:
    def __init__(self, session, page, page_size=100, file_name='product_skus.xlsx'):
        self.session = session  # API 基础URL
        self.page = page  # 页数
        self.page_size = page_size  # 每页请求数量
        self.file_name = file_name  # Excel 文件名

    # 获取 赏金首单25%抵扣列表
    def _hunterSellSkuPageList(session, page, pageSize):
        url = "https://dms.viomall.com/v3/hunter/sell/hunterSellSkuPageList.do"

        payload = f'page={page}&pageSize={pageSize}&sortField=&sortOrder='
        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Cookie': 'isUseV3=true; checkTianjinDeclarationRequired_22843=false; checkPointsRequired=false; JSESSIONID=9F1186D35D76A4911B681FDE635481B9; JSESSIONID=89B1B75330976383EA7DC4FE5E33B185; isUseV3=true',
            'Origin': 'https://dms.viomall.com',
            'Pragma': 'no-cache',
            'Referer': 'https://dms.viomall.com/v3/hunter/home.do',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Chromium";v="130", "Google Chrome";v="130", "Not?A_Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"'
        }
        response = requests.request("POST", url, headers=headers, data=payload)

        if response.status_code != 200:
            print(f"获取 赏金首单25%抵扣列表 产品列表失败 status code {response.status_code}")
            return None, None

        data = response.json()
        if response.status_code != 200:
            print(f"Request failed with status code {response.status_code}")
            return None, None

        # 从返回的数据中获取分页总数和产品列表
        total = data['value']['pagination']['total']
        product_skus = data['value']['data']

        return total, product_skus


    def _save_to_excel(self, product_skus):

        # Excel 文件的路径
        excel_file = 'hunter_product_skus.xlsx'

        # 提取 productSku 的数据
        product_skus = [item['productSku'] for item in product_skus]
        df_new = pd.DataFrame(product_skus, columns=['productSku'])

        # 检查文件是否存在
        file_exists = os.path.exists(excel_file)

        if file_exists:
            # 如果文件存在，读取现有数据
            df_existing = pd.read_excel(excel_file, sheet_name='HuntProductSKUs')

            # 合并现有数据和新数据
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            # 如果文件不存在，直接使用新数据
            df_combined = df_new

        # 将合并后的数据写入到 Excel
        with pd.ExcelWriter(excel_file, mode='w', engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False, header=True, sheet_name='HuntProductSKUs')

        print("productSku 已成功追加到 Excel 文件！")

    def fetch_and_save_all_skus(self):
        """ 获取所有产品SKU数据并保存到Excel """

        # 第一次请求，获取总数
        total, product_skus = self._hunterSellSkuPageList(1, self.page_size)

        if total == 0:
            print("No data found.")
            return

        # 保存第一页的数据到Excel
        self._save_to_excel(product_skus)

        # 计算需要多少次请求
        total_pages = (total // self.page_size) + (1 if total % self.page_size != 0 else 0)

        # 获取其他页的数据并保存到Excel
        for page in range(2, total_pages + 1):
            print(f"Fetching page {page}/{total_pages}")
            total, product_skus = self._hunterSellSkuPageList(page = page, pageSize= self.page_size)

            if total == 0:
                break

            self._save_to_excel(product_skus)

            # 随机休息 5 到 15 秒
            sleep_time = random.randint(5, 15)
            print(f"休息 {sleep_time} 秒...")
            time.sleep(sleep_time)

            print("休息的时间是", sleep_time, "S")

        print(f"Data fetching completed. All data saved to {self.file_name}.")

