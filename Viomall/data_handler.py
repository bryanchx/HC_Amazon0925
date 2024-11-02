# data_handler.py

import pandas as pd
import os
from config import EXCEL_FILE

class DataHandler:
    def __init__(self):
        self.file_exists = os.path.exists(EXCEL_FILE)

    def save_product_skus(self, product_skus):
        df_new = pd.DataFrame(product_skus, columns=['productSku'])

        if self.file_exists:
            df_existing = pd.read_excel(EXCEL_FILE, sheet_name='ProductSKUs')
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        with pd.ExcelWriter(EXCEL_FILE, mode='w', engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False, header=True, sheet_name='ProductSKUs')

        print("productSku 已成功追加到 Excel 文件！")
