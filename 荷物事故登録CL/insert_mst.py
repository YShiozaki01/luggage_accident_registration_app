import PySimpleGUI as sg
import sqlite3
import sys
import datetime

sg.theme("SystemDefault")
DATABASE = "database/database.db"


def get_db_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


def get_now_time():
    # 現在日時を取得
    dt_now = datetime.datetime.now()
    str_now_time = dt_now.strftime('%Y/%m/%d %H:%M:%S')
    return str_now_time


class MasterInsert:
    def __init__(self, id, title, kw, cd_num):
        self.id = id
        self.title = title
        self.keyword = kw
        self.cd = cd_num
        self.table_dct = {2: {"T": "M荷主", "F1": "ID", "F2": "名称"},
                          5: {"T": "M発生状況", "F1": "ID", "F2": "発生状況"},
                          6: {"T": "M事故惹起者", "F1": "ID", "F2": "名称", "F3": "他社"}}

    def make_window_sub2(self):
        layout = [[sg.T(self.table_dct[self.id]["F1"], font=("Yu Gothic UI", 8), size=(10, 0), k="txt1"),
                   sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in1"),
                   sg.T("新規登録", text_color="#FFFFFF", background_color="#FF0000", visible=False, k="txt_new")],
                   [sg.T(self.table_dct[self.id]["F2"], font=("Yu Gothic UI", 8), size=(10, 0), k="txt2"),
                   sg.I(default_text=self.keyword, font=("Yu Gothic UI", 9), size=(35, 0), k="in2")],
                   [sg.Checkbox("他社", font=("Yu Gothic UI", 9), default=False, disabled=True, k="chk1"),
                   sg.Push(), sg.B("キャンセル", font=("Yu Gothic UI", 9), k="btn_cancel"),
                   sg.B("削除", font=("Yu Gothic UI", 9), k="btn2", disabled=False),
                   sg.B("登録", font=("Yu Gothic UI", 9), k="btn1")]]
        win = sg.Window(f"{self.title}マスタ", layout, font=("Yu Gothic UI", 10), size=(380, 90), disable_close=True)
        win.finalize()
        return win
    
    def open_window_sub2(self):
        win_sub2 = self.make_window_sub2()
        mst_name = self.table_dct[self.id]["T"]
        # データベースへ接続
        conn = get_db_connection()
        cur = conn.cursor()
        # IDを取得
        # 新規登録なら
        if self.cd == 0:
            # 「新規登録」表示
            win_sub2["txt_new"].update(visible=True)
            # 削除ボタン無効
            win_sub2["btn2"].update(disabled=True)
            # 新規登録のIDを取得
            sql = f"SELECT * FROM {mst_name};"
            cur.execute(sql)
            if len(cur.fetchall()) == 0:
                cd = 1
            else:
                sql = f"SELECT max(ID) lastID FROM {mst_name};"
                cur.execute(sql)
                result = cur.fetchone()
                cd = result["lastID"] + 1
        else:
            cd = self.cd
        # ウィンドウの初期値を設定
        win_sub2["in1"].update(cd)
        win_sub2["in2"].update(self.keyword)
        # もし、事故惹起者の新規ならchk1を入力可にする
        if self.id == 6:
            win_sub2["chk1"].update(disabled=False)
        while True:
            e, v = win_sub2.read()
            if e == "btn1":
                id_num = v["in1"]
                name = v["in2"]
                dt = get_now_time()
                # IDが登録済みか
                sql = f"""SELECT * FROM {mst_name} WHERE {self.table_dct[self.id]["F1"]} = {self.cd};"""
                cur.execute(sql)
                exist = False if len(cur.fetchall()) == 0 else True
                # マスタ別のINSERT文を取得
                if self.id == 6:
                    flg = "1" if v["chk1"] else ""
                    sql_inst = f"""INSERT INTO {mst_name}(
                            {self.table_dct[self.id]["F1"]},
                            {self.table_dct[self.id]["F2"]},
                            {self.table_dct[self.id]["F3"]},
                            更新日, 更新区分
                            )
                            VALUES({id_num}, '{name}', '{flg}', '{dt}', 'I'
                            )"""
                if self.id == 2:
                    sql_inst = f"""INSERT INTO {mst_name}(
                            {self.table_dct[self.id]["F1"]},
                            {self.table_dct[self.id]["F2"]},
                            更新日, 更新区分, 所管店
                            )
                            VALUES({id_num}, '{name}', '{dt}', 'I', '{sys.argv[1]}'
                            )"""
                if self.id == 5:
                    sql_inst = f"""INSERT INTO {mst_name}(
                            {self.table_dct[self.id]["F1"]},
                            {self.table_dct[self.id]["F2"]},
                            更新日, 更新区分
                            )
                            VALUES({id_num}, '{name}', '{dt}', 'I'
                            )"""
                answ = sg.popup_ok_cancel("この内容で登録しますか？", font=("Yu Gothic UI", 8), title="確認")
                if answ == "OK":
                    # もし当該IDがテーブルに既存なら、同レコードを削除
                    if exist:
                        sql_del = f"""DELETE FROM {mst_name} WHERE {self.table_dct[self.id]["F1"]} = {self.cd};"""                        
                        cur.execute(sql_del)
                        conn.commit()
                    # テーブルに入力値を追加
                    cur.execute(sql_inst)
                    conn.commit()
                    sg.popup("登録しました。", font=("Yu Gothic UI", 8), title="完了")
                    exit = True
                    break
                else:
                    sg.popup("取り消しました。", font=("Yu Gothic UI", 8), title="中止")
                    exit = True
                    break
            if e == "btn2":
                answ = sg.popup_ok_cancel("このレコードを削除します。よろしいですか？", font=("Yu Gothic UI", 8), title="確認")
                if answ == "OK":
                    sql_del = f"""DELETE FROM {mst_name} WHERE {self.table_dct[self.id]["F1"]} = {self.cd};"""                        
                    cur.execute(sql_del)
                    conn.commit()
                    sg.popup("削除しました。", font=("Yu Gothic UI", 8), title="完了")
                    exit = True
                    break
                else:
                    sg.popup("取り消しました。", font=("Yu Gothic UI", 8), title="中止")
                    exit = True
                    break
            if e == "btn_cancel":
                exit = True
                break
        win_sub2.close()
        return exit
