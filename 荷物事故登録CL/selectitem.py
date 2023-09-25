import PySimpleGUI as sg
import sqlite3
from insert_mst import MasterInsert
import sys

sg.theme("SystemDefault")
DATABASE = "database/database.db"


def get_db_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


class SelectItem:
    def __init__(self, id):
        self.id = id
        self.cd = ""
    
    def get_syntax(self, kw):
        # テーブルの明細データを取得する構文を取得
        if self.id in(1, 3):
            sql = f"""SELECT コード, 名称 FROM M部署 WHERE 名称 like '%{kw}%'"""
        elif self.id == 2:
            sql = f"SELECT ID, 名称 FROM M荷主 WHERE 名称 like '%{kw}%' AND 所管店 = '{sys.argv[1]}'"
        elif self.id == 4:
            sql = f"SELECT 郵便番号, 住所 FROM M住所 WHERE 住所 like '%{kw}%'"
        elif self.id == 5:
            sql = f"SELECT ID, 発生状況 FROM M発生状況 WHERE 発生状況 like '%{kw}%'"
        elif self.id == 6:
            sql = f"SELECT ID, 名称, 他社 FROM M事故惹起者 WHERE 名称 like '%{kw}%'"
        return sql

    def check_exist(self, kw):
        # 検索結果が存在するかチェック
        if self.id in(1, 3):
            sql = f"""SELECT * FROM M部署 WHERE 名称 like '%{kw}%'"""
        elif self.id == 2:
            sql = f"SELECT * FROM M荷主 WHERE 名称 like '%{kw}%' AND 所管店 = '{sys.argv[1]}'"
        elif self.id == 4:
            sql = f"SELECT * FROM M住所 WHERE 住所 like '%{kw}%'"
        elif self.id == 5:
            sql = f"SELECT * FROM M発生状況 WHERE 発生状況 like '%{kw}%'"
        elif self.id == 6:
            sql = f"SELECT * FROM M事故惹起者 WHERE 名称 like '%{kw}%'"
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(sql)
        if len(cur.fetchall()) == 0:
            result = True
        else:
            result = False
        return result

    def make_window_sub(self):
        # 検索ウインドウを作成
        # ヘッダー
        if self.id == 6:
            header = ["ID", "NAME", "FLG"]
        else:
            header = ["ID", "NAME", ""]
        # # 明細（初期値）
        # data = []

        # レイアウト
        layout = [[sg.T("キーワード"), sg.I(font=("Yu Gothic UI", 9), size=(35, 0), k="in_kw"),
                   sg.B("検索", font=("Yu Gothic UI", 9), k="btn1"), sg.Push(),
                   sg.B("マスタ", k="btn2", font=("Yu Gothic UI", 9), disabled=False if self.id in(2, 5, 6) else True)],
                   [sg.Table([[]], header, font=("Yu Gothic UI", 10), justification="left", col_widths=(10, 32, 5),
                             auto_size_columns=False, select_mode=sg.TABLE_SELECT_MODE_BROWSE, num_rows=None,
                             k="slc", enable_events=True)],
                    [sg.Push(), sg.B("キャンセル", font=("Yu Gothic UI", 9), size=(10, 0), k="btn_cancel"),
                     sg.B("選択", k="btn3", font=("Yu Gothic UI", 9), size=(10, 0), disabled=True)]]
        win = sg.Window(self.title, layout, font=("Yu Gothic UI", 10), size=(420, 270), disable_close=True)
        win.finalize()
        return win
    
    def open_window_sub(self):
        win = self.make_window_sub()
        win["in_kw"].bind("<Return>", "_rtn")
        win["btn1"].bind("<Return>", "_rtn")
        win["btn2"].bind("<Return>", "_rtn")
        win["btn3"].bind("<Return>", "_rtn")
        # データベースに接続
        conn = get_db_connection()
        cur = conn.cursor()
        # ウインドウを開く
        slc_click = False   # テーブルクリックアクション判定用変数
        while True:
            e, v = win.read()
            if e == "in_kw_rtn":
                win["btn1"].SetFocus(True)
            if e == "btn1" or e == "btn1_rtn":
                kw = v["in_kw"]
                if self.check_exist(kw):
                    sg.popup("該当する登録がありません", font=("Yu Gothic UI", 8), title="確認")
                    self.cd = 0
                else:
                    sql = self.get_syntax(kw)
                    cur.execute(sql)
                    result = [list(row) for row in cur.fetchall()]
                    win["slc"].update(result)
            if e == "slc":
                row = v["slc"]
                # リストから行が取得されているか ※エラー対応
                if not row:
                    win["btn3"].update(disabled=False)
                else:
                    slc_click = True
                    item = result[row[0]]
                    self.cd = item[0]
                    win["in_kw"].update(item[1])
                    win["btn3"].update(disabled=False)
            if e == "btn2" or e == "btn2_rtn":
                if not slc_click:
                    self.cd = 0
                slc_click = False
                win["btn3"].update(disabled=True)
                mi = MasterInsert(self.id, self.title, v["in_kw"], self.cd)
                mi.open_window_sub2()
            if e == "btn3" or e == "btn3_rtn":
                break
            if e == "btn_cancel":
                item = []
                break
        win.close()
        return item


if __name__ == "__main__":
    si = SelectItem()
    si.open_window_sub()
