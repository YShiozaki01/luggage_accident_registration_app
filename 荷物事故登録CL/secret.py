import shutil
import sqlite3
import os

FOLDER1 = "attachment"
DATABASE = "database/database.db"
TABLE1 = "T削除番号"
TABLE2 = "T添付資料"
TABLE3 = "T荷物事故"
TABLE4 = "T送信番号"
TABLE5 = "T削除番号F"


class Secret:
    def __init__(self):
        pass

    def clear_folder(self):
        # 添付資料ファイル保存用フォルダを一旦フォルダごと削除して、再度フォルダを作成
        shutil.rmtree(FOLDER1)
        os.mkdir(FOLDER1)

    def clear_table1(self):
        # T荷物事故テーブルと、それに関連するテーブルをクリア
        sql = f"""
            DELETE FROM {TABLE1};
            DELETE FROM {TABLE2};
            DELETE FROM {TABLE3};
            DELETE FROM {TABLE4};
            DELETE FROM {TABLE5};
            """
        conn = sqlite3.Connection(DATABASE)
        cur = conn.cursor()
        cur.executescript(sql)
        conn.commit()
