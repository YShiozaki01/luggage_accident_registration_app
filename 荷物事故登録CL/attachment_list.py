import PySimpleGUI as sg
import sqlite3
import os
from pathlib import Path

sg.theme("SystemDefault")

DATABASE = fr"{os.getcwd()}\database\database.db"
FOLDER = "attachment"

class AttachmentList:
    def __init__(self, control_num, detail_num):
        self.control_number = control_num
        self.detail_number = detail_num

    def make_list_window(self):
        header = ("添付資料", "拡張子")
        layout = [[sg.Table([[]], header, font=("Yu Gothic UI", 10), justification="left", col_widths=(30, 10),
                             auto_size_columns=False, select_mode=sg.TABLE_SELECT_MODE_BROWSE, num_rows=None,
                             k="slc", enable_events=True)],
                  [sg.B("開く", k="btn_open", font=("Yu Gothic UI", 8), size=(10, 0)),
                   sg.B("削除", k="btn_del", font=("Yu Gothic UI", 8), size=(10, 0)),
                   sg.Push(), sg.B("閉じる", k="btn_close", font=("Yu Gothic UI", 8), size=(10, 0))]]
        window = sg.Window("添付資料リスト", layout, size=(370, 240), disable_close=True)
        window.finalize()
        return window
    
    def check_empty(self):
        sql = f"""SELECT * FROM T添付資料 WHERE 管理番号 = {self.control_number} and 明細番号 = {self.detail_number}"""
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        result = len(cur.fetchall())
        return result

    def get_attachment_data(self):
        sql = f"""SELECT ファイル名称, 拡張子 FROM T添付資料
                WHERE 管理番号 = {self.control_number} and 明細番号 = {self.detail_number}"""
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        result = cur.fetchall()
        return result
    
    def open_list_window(self):
        if self.check_empty() == 0:
            sg.popup("資料は添付されていません。", font=("Yu Gothic UI", 8))
        else:
            win = self.make_list_window()
            win["slc"].update(self.get_attachment_data())
            while True:
                e, v = win.read()
                if e == "slc":
                    row = v["slc"]
                    row_num = row[0]
                    file_list = self.get_attachment_data()
                    record = file_list[row_num]
                    filename = record[0]
                    currentdir = os.getcwd()
                    folder = os.path.dirname(filename)
                    name = os.path.basename(filename)
                    filepath = os.path.join(currentdir, folder, name)
                    # print(filepath)
                    # os.startfile(filepath)
                if e == "btn_open":
                    os.startfile(filepath)
                if e == "btn_del":
                    fname = os.path.basename(filepath)
                    answ = sg.popup_ok_cancel(f"{fname}を削除します。よろしいですか？", font=("Yu Gothic UI", 8))
                    # もしYESなら以下を実行
                    if answ == "OK":
                        # T添付資料のレコードを削除
                        sql = f"""
                            DELETE FROM T添付資料 WHERE ファイル名称 like '%{fname}';
                            """
                        conn = sqlite3.connect(DATABASE)
                        cur = conn.cursor()
                        cur.execute(sql)
                        conn.commit()
                        # 削除したファイルをT削除番号Fに書き込み
                        sql = f"""
                            INSERT INTO T削除番号F(ファイル名) VALUES('{fname}');
                            """
                        cur.execute(sql)
                        conn.commit()
                        sg.popup(f"{fname}を削除しました。", font=("Yu Gothic UI", 8))
                    else:
                        sg.popup(f"{fname}の削除を中止します。", font=("Yu Gothic UI", 8))
                if e == "btn_close":
                    break
            win.close()


if __name__ == "__main__":
    al = AttachmentList(2, 6)
    al.open_list_window()