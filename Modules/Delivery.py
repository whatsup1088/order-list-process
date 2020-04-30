#!/usr/bin/env python
# coding: utf-8

from pandas import pandas as pd
from datetime import datetime

import logging
import os


today=datetime.today().strftime('%Y-%m-%d')
log_filename=f'{os.path.expanduser("~/Desktop")}/{today}_總表無對應商品代號記錄.txt'

class Delivery:
    def __init__(self):
        self._sheet_name = '總表'
        self._source_df = None
        self._logfile = False


    def read_source(self, file_path):
        try:
            self._source_df = pd.read_excel(
                io=file_path,
                sheet_name=self._sheet_name,
                dtype=str
            )
        except Exception as e:
            raise TypeError(f'總表讀取失敗，請確認總表中的工作表名稱設為: "{self._sheet_name}"\n{e}')

        if(not all(col in self._source_df.columns.tolist() for col in self._source_header)):
            raise TypeError(f'總表欄位有缺少，請檢查欄位是否含有 {self._source_header}')

        # To make index same as file row
        self._source_df.index = self._source_df.index + 2

    def get_logfile(self):
        if self._logfile:
            return

        logging.basicConfig(
            filename=log_filename,
            filemode='w',
            level=logging.DEBUG,
            format='%(message)s'
        )
        self._logfile = True

    def get_delivery_data(self):
        return pd.DataFrame(self._data, columns=self._output_header)

    def get_colnum_string_by_index(self, index):
        n = index + 1
        string = ""
        while n > 0:
            n, remainder = divmod(n -1, 26)
            string = chr(65 + remainder) + string
        return string

    def save_to(self, output_file_path):
        result = self.get_delivery_data()

        writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        result.to_excel(writer, index=False, sheet_name='出貨表單')

        # Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets['出貨表單']

        # Add cell formats.
        cell_format = workbook.add_format({'num_format': '@'})

        # Set the format
        for n in range(0, len(self._output_header)):
            colnum_string = self.get_colnum_string_by_index(n)
            worksheet.set_column(f'{colnum_string}:{colnum_string}', None, cell_format)

        writer.save()


class NormalDelivery(Delivery):
    def __init__(self):
        super(NormalDelivery, self).__init__()
        self._source_header = ['訂單編號', '商品代號', '賣場名稱', '商品規格', '訂購數量', '收件者姓名', '收件者電話', '收件者郵編', '收件者地址', '訂單留言']
        self._output_header = ['作業類別','客戶編號','出貨單號','客單編號','到運倉','指定出貨日','指定送達日','第二備註','移動類型','送貨類型','訂貨人姓名','訂貨人電話號碼','訂貨人行動電話','提貨人姓名','提貨人郵遞區號','提貨人地址','提貨人日間電話','提貨人夜間電話','提貨人行動電話','提貨人身分證字號','代收金額','配送件數','才積','指定時間','指定時段','供應商編號','產品編號','群品','批號','產品名稱','數量','單位','單價','品級','贈品欄','入庫欄位(良品)', '訂購數量', '業務倉別']
        self._output_file_path = None
        self._data = pd.DataFrame()


    def set_default(self):
        data = {value: '' for value in self._output_header}
        data['作業類別'] = '1'
        data['送貨類型'] = '6'
        data['群品'] = '1'
        data['品級'] = '0'

        return data


    def proccess(self, mapping):
        mapping_error_flag=False
        for column, row in self._source_df.iterrows():
            proccess_data = self.set_default()
            proccess_data['客戶編號'] = proccess_data['客單編號'] = row['訂單編號']
            proccess_data['訂貨人姓名'] = proccess_data['提貨人姓名'] = row['收件者姓名']
            proccess_data['提貨人日間電話'] = row['收件者電話']
            proccess_data['提貨人郵遞區號'] = row['收件者郵編']
            proccess_data['提貨人地址'] = row['收件者地址']
            proccess_data['第二備註'] = row['訂單留言']
            proccess_data['贈品欄'] = row['賣場名稱']
            proccess_data['入庫欄位(良品)'] = row['商品規格']
            proccess_data['訂購數量'] = row['訂購數量']

            mapping_df = mapping.get_df_by_item_code(row['商品代號'])

            if mapping_df.empty:
                self.get_logfile()
                logging.debug(f'常溫訂單總表第 {column} 行，沒有對應的商品代號')
                mapping_error_flag=True
                continue

            for mp_column, mp_row in mapping_df.iterrows():
                proccess_data['產品編號'] = mp_row['倉庫出貨產品編號']
                proccess_data['產品名稱'] = mp_row['倉庫出貨產品名稱']
                proccess_data['數量'] = str(int(row['訂購數量']) * int(mp_row['數量']))
                self._data = self._data.append(proccess_data , ignore_index=True)

        if mapping_error_flag:
            raise TypeError(f'常溫訂單總表含有對應不到的商品代號，詳細資訊請參考此檔案\n {log_filename}')


class FreezeDelivery(Delivery):
    def __init__(self):
        super(FreezeDelivery, self).__init__()
        self._source_header = ['通路名稱', '訂單編號', '商品代號', '賣場名稱', '商品規格', '訂購數量', '收件者姓名', '收件者電話', '收件者郵編', '收件者地址', '訂單留言']
        self._output_header = ['通路名稱', '訂單編號', '訂單項次', '聯絡人', '電話', '單據號碼', '銷貨日期', '姓名', '客戶編號', '派送路線', '產品編號', '產品名稱', '單價', '數量', '贈品', '賣場名稱', '商品規格', '訂購數量']
        self._output_file_path = None
        self._data = pd.DataFrame()


    def set_default(self):
        data = {value: '' for value in self._output_header}
        data['姓名'] = '安心電商'
        data['客戶編號'] = 'SZA001'
        data['派送路線'] = '99'
        data['贈品'] = '0'

        return data


    def proccess(self, mapping):
        mapping_error_flag=False
        for column, row in self._source_df.iterrows():
            proccess_data = self.set_default()
            proccess_data['通路名稱'] = row['通路名稱']
            proccess_data['訂單編號'] = row['訂單編號']
            proccess_data['聯絡人'] = row['收件者姓名']
            proccess_data['電話'] = row['收件者電話']
            proccess_data['賣場名稱'] = row['賣場名稱']
            proccess_data['商品規格'] = row['商品規格']
            proccess_data['訂購數量'] = row['訂購數量']

            mapping_df = mapping.get_df_by_item_code(row['商品代號'])

            if mapping_df.empty:
                self.get_logfile()
                logging.debug(f'冷凍訂單總表第 {column} 行，沒有對應的商品代號')
                mapping_error_flag=True
                continue

            for mp_column, mp_row in mapping_df.iterrows():
                proccess_data['產品編號'] = mp_row['倉庫出貨產品編號']
                proccess_data['產品名稱'] = mp_row['倉庫出貨產品名稱']
                proccess_data['數量'] = str(int(row['訂購數量']) * int(mp_row['數量']))
                self._data = self._data.append(proccess_data , ignore_index=True)

        if mapping_error_flag:
            raise TypeError(f'冷凍訂單總表含有對應不到的商品代號，詳細資訊請參考此檔案\n {log_filename}')

