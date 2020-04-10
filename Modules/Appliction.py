from tkinter import filedialog
from functools import partial
from shutil import copyfile

import os
import tkinter as tk
import tkinter.messagebox
from tempfile import gettempdir

from Modules.Delivery import NormalDelivery as nd, FreezeDelivery as fd
from Modules.Mapping import Mapping as mp

class Appliction:
    def __init__(self):
        self._window = tk.Tk()
        self._source = {'常溫': None, '冷凍': None}
        self._export_type = '常溫'

        self._mapping = None
        self._mapping_file_path = os.path.join(gettempdir(), 'mapping.xlsx')

        self._mapping_file_notice = ''
        if os.path.isfile(self._mapping_file_path):
            self._mapping = mp(self._mapping_file_path)
            self._mapping_file_notice = '(已匯入)'


    def get_source_file(self, type=''):
        source_path = filedialog.askopenfilename(
            title="請選擇總表",
            filetypes=[("Excel files", ".xlsx .xls")]
        )

        if not source_path:
            return

        try:
            if '常溫' == type:
                self._source[type] = nd()
                self._source[type].read_source(source_path)
                self.show_success(f'{type}訂單總表匯入成功')
            if '冷凍' == type:
                self._source[type] = fd()
                self._source[type].read_source(source_path)
                self.show_success(f'{type}訂單總表匯入成功')
        except Exception as e:
                self.show_error(e)


    def import_mapping_file(self):
        file_path = os.path.normpath(filedialog.askopenfilename(
            title="請選擇對照表",
            filetypes=[("Excel files", ".xlsx .xls")])
        )

        if file_path and '.' != file_path:
            try:
                self._mapping = mp(f"{file_path}")
                copyfile(file_path, self._mapping_file_path)
                self.show_success(f'對照表匯入成功')

                self._mapping_file_notice = '(已匯入)'
                self._menubar.entryconfigure(1, label=f'對照表 {self._mapping_file_notice}')
            except Exception as e:
                self.show_error(e)


    def export_file(self):
        if not self._mapping:
            self.show_error('請先匯入對照表')
            return

        if not self._source[self._export_type]:
            self.show_error(f'請先匯入{self._export_type}訂單總表')
            return

        try:
            self._source[self._export_type].proccess(self._mapping)
            eport_file_path = filedialog.asksaveasfilename(
                title="請選擇匯出位置",
                defaultextension='.xlsx',
                filetypes=[("Excel files", ".xlsx .xls")]
            )

            if eport_file_path:
                self._source[self._export_type].save_to(eport_file_path)
                self.show_success(f'{self._export_type}出貨格式資料匯出成功')
        except Exception as e:
                self.show_error(f'匯出檔案失敗\n{e}')


    def set_export_type(self, type=''):
        self._export_type = type.get()


    def create_radiobutton(self):
        type = tk.StringVar(None, '常溫')

        r_normal = tk.Radiobutton(
            self._window,
            text='常溫',
            variable=type,
            value='常溫',
            command=partial(self.set_export_type, type)
        ).grid(row=2, column=1, padx=0, pady=10, ipadx=10, ipady=10)

        r_freeze = tk.Radiobutton(
            self._window,
            text='冷凍',
            variable=type,
            value='冷凍',
            command=partial(self.set_export_type, type)
        ).grid(row=2, column=2, padx=0, pady=10, ipadx=10, ipady=10)


    def create_menu(self):
        self._menubar = tk.Menu(self._window)
        mapping_menu = tk.Menu(self._menubar, tearoff=0)

        self._menubar.add_cascade(label=f'對照表 {self._mapping_file_notice}', menu=mapping_menu)

        mapping_menu.add_command(label=f'新增對照表', command=self.import_mapping_file)

        self._window.config(menu=self._menubar)


    def create_buttons(self):
        source_normal_btn = tk.Button(
            self._window,
            text='匯入\n常溫訂單總表',
            font=('Arial', 20),
            width=10,
            height=5,
            command=partial(self.get_source_file, '常溫')
        ).grid(row=1, column=1, padx=25, pady=25, ipadx=10, ipady=10)

        source_freeze_btn = tk.Button(
            self._window,
            text='匯入\n冷凍訂單總表',
            font=('Arial', 20),
            width=10,
            height=5,
            command=partial(self.get_source_file, '冷凍')
        ).grid(row=1, column=2, padx=25, pady=25, ipadx=10, ipady=10)

        export_btn = tk.Button(
            self._window,
            text='匯出出貨格式',
            font=('Arial', 26),
            width=15,
            height=1,
            command=self.export_file
        ).grid(row=3, column=1, columnspan=2, padx=10, pady=10, ipadx=10, ipady=10)

    def show_success(self, msg):
        tkinter.messagebox.showinfo(title='系統通知', message=str(msg))


    def show_error(self, msg):
        tkinter.messagebox.showerror(title='系統通知', message=str(msg))


    def run(self):
        self._window.title('出貨訂單小幫手')
        self._window.geometry('500x450')
        self.create_menu()
        self.create_buttons()
        self.create_radiobutton()

        self._window['bg'] = 'powderblue'
        self._window.mainloop()
