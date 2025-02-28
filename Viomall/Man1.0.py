import os
import time
import random
from enum import Enum

import pandas as pd
import requests

from Viomall.HunterSellSkuFirstOrderPageList import HunterSellSkuFirstOrderPageList


# 登录
def login(username, password):
    login_url = "https://dms.viomall.com/v3/user/userLogin.do"
    session = requests.Session()

    payload = f"userName={username}&passWord={password}"
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Origin': 'https://dms.viomall.com',
        'Referer': 'https://dms.viomall.com/v3/users/login.do',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }

    # 发送登录请求
    login_response = session.post(login_url, headers=headers, data=payload)

    # 检查登录是否成功
    if login_response.ok:
        print("登录成功！")
        print(login_response.text)  # 可以选择移除此行，避免泄露敏感信息
        return session  # 返回会话对象以供后续请求使用
    else:
        print("登录失败！")
        print("响应内容:", login_response.text)
        return None

# 获取Api刊登工具已完成列表数据
def get_amazon_listing(session, page, pageSize):
    url = "https://dms.viomall.com/v3/amazonListing/getAmazonListingBatchList.do"
    payload = f'display=completed&pageSize={pageSize}&page={page}&sortField=&sortOrder='
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://dms.viomall.com',
        'Referer': 'https://dms.viomall.com/v3/amazonListing/amazonListingBatchList.do',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }

    # 发送请求
    response = session.post(url, headers=headers, data=payload)
    if response.ok:
        print('getAmazonListingBatchList= ', response.text)
        return response.json()
    else:
        print("获取 getAmazonListingBatchList 失败！")
        print("响应内容:", response.text)
        return None

# 获取 Api刊登工具 里面的SKU列表
def get_amazon_listing_product_list(session, product_id, page=1, page_size=10):
    url = f"https://dms.viomall.com/v3/amazonListing/getAmazonListingBatchProductList.do?id={product_id}"

    payload = f'display=all&page={page}&pageSize={page_size}&sortField=&sortOrder='
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://dms.viomall.com',
        'Referer': f'https://dms.viomall.com/v3/amazonListing/editAmazonListingBatch.do?id={product_id}',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
    }

    # 发送请求
    response = session.post(url, headers=headers, data=payload)

    if response.ok:
        print('获取 Amazon Listing 产品列表成功！')
        return response.json()  # 返回 JSON 数据
    else:
        print("获取 Amazon Listing 产品列表失败！")
        print("响应内容:", response.text)
        return None

# 定义一个枚举类
class RequestType(Enum):
    WeiShangJia = 1
    ShangJinShouDan = 2

# 主程序
if __name__ == "__main__":
    username = 'bryanchx@163.com'  # 替换为你的用户名
    password = '48ncmTHF'           # 替换为你的密码

    # 登录并获取会话对象
    session = login(username, password)
    request_type = RequestType.WeiShangJia

    if session:
        if request_type == RequestType.WeiShangJia:
            # 使用登录的会话请求其他接口
            complete_api_list = get_amazon_listing(session, 1, 150)
            if complete_api_list:  # 确保有数据
                # 解析返回结果
                if complete_api_list['success']:
                    # 提取 pagination 信息
                    pagination = complete_api_list['value']['pagination']
                    current_page = pagination['current']
                    page_size = pagination['pageSize']
                    total = pagination['total']

                    print(f"当前页: {current_page}, 每页大小: {page_size}, 总数: {total}")

                    # 提取数据列表
                    listings = complete_api_list['value']['data']
                    for listing in listings:

                        org_id = listing['orgId']
                        item_id = listing['id']
                        title = listing['title']
                        print(f"ID: {item_id}, Org ID: {org_id}, 标题: {title}")
                        listing_product_list = get_amazon_listing_product_list(session, item_id, 1, 1000)

                        if listing_product_list and listing_product_list['success']:
                            # 提取 productSku
                            product_skus = [item['productSku'] for item in listing_product_list['value']['data']]
                            print("product_skus=",product_skus)

                            # Excel 文件的路径
                            excel_file = 'product_skus.xlsx'

                            # 提取 productSku 的数据
                            product_skus = [item['productSku'] for item in listing_product_list['value']['data']]
                            df_new = pd.DataFrame(product_skus, columns=['productSku'])

                            # 检查文件是否存在
                            file_exists = os.path.exists(excel_file)

                            if file_exists:
                                # 如果文件存在，读取现有数据
                                df_existing = pd.read_excel(excel_file, sheet_name='ProductSKUs')

                                # 合并现有数据和新数据
                                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                            else:
                                # 如果文件不存在，直接使用新数据
                                df_combined = df_new

                            # 将合并后的数据写入到 Excel
                            with pd.ExcelWriter(excel_file, mode='w', engine='openpyxl') as writer:
                                df_combined.to_excel(writer, index=False, header=True, sheet_name='ProductSKUs')

                            print("productSku 已成功追加到 Excel 文件！")

                            # 随机休息 5 到 15 秒
                            sleep_time = random.randint(5, 15)
                            print(f"休息 {sleep_time} 秒...")
                            time.sleep(sleep_time)

                            print("休息的时间是", sleep_time, "S")

                        else:
                            print("没有获取到有效的产品列表。")
                else:
                    print("请求失败，未能获取数据。")

        elif request_type == RequestType.ShangJinShouDan:
            # 创建 ProductSkuFetcher 类的实例
            sku_fetcher = HunterSellSkuFirstOrderPageList(session, page=1, page_size=100,
                                                          file_name='hunter_product_skus.xlsx')

            # 调用方法来获取并保存所有产品SKU
            sku_fetcher.fetch_and_save_all_skus()
