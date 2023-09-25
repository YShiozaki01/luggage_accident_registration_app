import PySimpleGUI as sg
from PIL import Image
import io
import os
import shutil
import datetime
import sqlite3

sg.theme("SystemDefault")

DATABASE = "database/database.db"
TABLE = "T添付資料"
DESTINATION = "attachment"
SIZE = (800, 800)
EXTLIST = [".jpg", ".JPG", ".jpeg", ".png", ".PNG"]

def get_db_connection():
    # データベースへのコネクションを取得
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


def get_now_time():
    # 現在日時を取得
    dt_now = datetime.datetime.now()
    str_now_time = dt_now.strftime('%Y/%m/%d %H:%M:%S')
    return str_now_time


def keepAspectResize(path):
    # 画像ファイルを指定矩形内に収まるようにリサイズ
    # 1)画像の読み込み
    image = Image.open(path)
    # 2)サイズを幅と高さにアンパック
    width, height = SIZE
    # 3)矩形と画像の幅・高さの比率の小さい方を拡大率とする
    ratio = min(width / image.width, height / image.height)
    # 4)画像の幅と高さに拡大率を掛けてリサイズ後の画像サイズを算出
    resize_size = (round(ratio * image.width), round(ratio * image.height))
    # 5)リサイズ後の画像サイズにリサイズ
    resized_image = image.resize(resize_size)
    return resized_image


class LoadImage:
    def __init__(self):
        pass

    def get_new_number(self):
        # 最新連番を取得
        # テーブルにレコードがあるか、又は同一管理番号、明細番号のレコードがあるか
        sql = f"""SELECT * FROM {TABLE} WHERE 管理番号 = {self.control_number} 
                and 明細番号 = {self.detail_number};"""
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(sql)
        if len(cur.fetchall()) == 0:
            result = 1
        else:
            sql = f"""SELECT max(資料連番号) as 最終番号 FROM {TABLE}
                    WHERE 管理番号 = {self.control_number}
                        and 明細番号 = {self.detail_number};"""
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute(sql)
            last_num = cur.fetchone()
            result = last_num["最終番号"] + 1
        return result
    
    def insert_table(self, ext, dest, serial_number):
        # T添付資料テーブルに添付資料情報を挿入
        sql = f"""INSERT INTO T添付資料(管理番号, 明細番号, 拡張子, ファイル名称, 資料連番号,
                元ファイル, 更新日時)
                VALUES({self.control_number}, {self.detail_number}, '{ext}',
                '{dest}', {serial_number}, '{self.file_path}', '{get_now_time()}'
                );"""
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()

    def make_img_window(self):
        layout = [[sg.T(k="txt", font=("Yu Gothic UI", 8))],
                  [sg.Im(k="img")],
                  [sg.B("登録", k="btn_register", font=("Yu Gothic UI", 8))]]
        window = sg.Window("画像ファイルを表示", layout, size=(320, 380))
        window.finalize()
        return window

    def load_img(self):
        win = self.make_img_window()
        try:
            img = Image.open(self.file_path)
            img.thumbnail((300, 300))
            bio = io.BytesIO()
            img.save(bio, format="PNG")
            win["img"].update(data=bio.getvalue())
            win["txt"].update(self.file_path)
        except:
            win["img"].update()
            win["txt"].update("No Image")
        while True:
            e, v = win.read()
            if e == "btn_register":
                # 確認メッセージ
                result = sg.popup_ok_cancel("このファイルを登録します。\nよろしいですか？", font=("Yu Gothic UI", 8), title="確認")
                if result == "OK":
                    # 当該ファイルの拡張子を取得
                    ext_pair = os.path.splitext(self.file_path)
                    ext = ext_pair[1]
                    # 最新の添付連番号（同一管理番号、明細番号内）を取得
                    serial_number = self.get_new_number()
                    # 保存名用の事故番号を生成
                    case_number = self.control_number.zfill(6) + self.detail_number.zfill(3)
                    # 保存名称を生成（報告店コード＋事故番号＋添付連番号）
                    str_number = str(serial_number)
                    save_name = f"{self.department_code}_{case_number}_{str_number.zfill(2)}"
                    # 生成した保存名称で当該ファイルを保存
                    src = self.file_path
                    dest = f"{DESTINATION}/{save_name}{ext}"
                    if ext in EXTLIST:
                        resize_img = keepAspectResize(src)
                        resize_img.save(dest)
                    else:
                        shutil.copy(src, dest)
                    # 添付ファイルの保存情報をテーブルに書き込み
                    self.insert_table(ext, dest, serial_number)
                    # 確認メッセージ
                    sg.popup("登録しました。", font=("Yu Gothic UI", 8), title="完了")
                else:
                    sg.popup("登録を中止しました。", font=("Yu Gothic UI", 8), title="取消")
                # ループを抜ける
                break
            if e == None:
                break
        win.close()
