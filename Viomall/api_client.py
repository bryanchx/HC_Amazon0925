# api_client.py

import requests
from config import LOGIN_URL, LISTING_URL, PRODUCT_LIST_URL

class APIClient:
    def __init__(self):
        self.session = requests.Session()

    def login(self, username, password):
        payload = f"userName={username}&passWord={password}"
        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
        }

        response = self.session.post(LOGIN_URL, headers=headers, data=payload)

        if response.ok:
            print("登录成功！")
            return True
        else:
            print("登录失败！", response.text)
            return False

    def login1(self,username,password):
        return True

    def get_amazon_listing(self, page, page_size):
        payload = f'display=completed&pageSize={page_size}&page={page}'
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

        response = self.session.post(LISTING_URL, headers=headers, data=payload)
        return response.json() if response.ok else None

    def get_amazon_listing_product_list(self, product_id, page=1, page_size=10):
        payload = f'display=all&page={page}&pageSize={page_size}'
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
        url = f"{PRODUCT_LIST_URL}?id={product_id}"
        response = self.session.post(url,headers= headers, data=payload)

        return response.json() if response.ok else None
