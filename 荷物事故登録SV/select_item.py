import PySimpleGUI as sg
import sqlite3

sg.theme("SystemDefault")
DATABASE = "database/database.db"


class SelectItem:
    def __init__(self, str_title, list):
        self.items = list
        self.title = str_title
     
    def make_window_sub(self):
        # 検索ウインドウを作成
        # ヘッダー
        header = ["ID", "NAME"]
        # レイアウト
        layout = [[sg.Table(self.items, header, font=("Yu Gothic UI", 10), justification="left", col_widths=(10, 32, 5),
                            auto_size_columns=False, select_mode=sg.TABLE_SELECT_MODE_BROWSE, num_rows=None,
                            k="slc", enable_events=True)]]
        window = sg.Window(self.title, layout, font=("Yu Gothic UI", 10), size=(360, 210), disable_close=False)
        window.finalize()
        return window
    
    def get_select_item(self):
        # 検索ウィンドウを開いて選択したアイテムのコードと名称を取得
        win = self.make_window_sub()
        while True:
            e, v = win.read()
            if e == "slc":
                row = v["slc"]
                row_num = row[0]
                break
            if e == None:
                row_num = None
                break
        win.close()
        return row_num