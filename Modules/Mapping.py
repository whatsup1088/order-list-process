import pandas as pd
class Mapping:
    def __init__(self, mapping_file_path: str):

        self._header = ['訂單總表商品代號', '訂單總表 商品名稱', '倉庫出貨產品編號', '倉庫出貨產品名稱', '數量']
        self._mappging_df = None
        self._sheet_name = '商品對照表'

        try:
            self._mappging_df = pd.read_excel(
                io=mapping_file_path,
                sheet_name=self._sheet_name,
                dtype=str
            )
        except Exception as e:
            raise TypeError(f'商品對照表讀取失敗，請確認是否匯入商品對照表且工作表名稱設為: "{self._sheet_name}"\n{e}')

        if self._mappging_df.columns.tolist() != self._header:
            raise TypeError(f'對照表欄位有缺少，請檢查欄位是否含有 {self._header}')

        self.preproccess()

        # To make index same as file row
        self._mappging_df.index = self._mappging_df.index + 2


    def get_df_by_item_code(self, code: str):
        filter = (self._mappging_df['訂單總表商品代號'] == code)
        target = self._mappging_df[filter]

        return target


    def show(self):
        return self._mappging_df


    def preproccess(self):
        for col, row in self._mappging_df.iterrows():
            if pd.isnull(row['訂單總表商品代號']):
                row['訂單總表商品代號'] = tmp['訂單總表商品代號']
                row['訂單總表 商品名稱'] = tmp['訂單總表 商品名稱']
            tmp = row

        self._mappging_df = self._mappging_df.drop_duplicates()
