import sqlite3
import os
import shutil

DATABASE = "database/database.db"
STORAGE1 = "output"
STORAGE2 = "static/attachment"
SQLSCRIPT1 = """DELETE FROM TS削除番号;
                DELETE FROM TS削除番号F;
                DELETE FROM TS荷物事故;
                DELETE FROM TS添付資料;
                DELETE FROM T最終取込ファイル名称;"""

cwd = os.getcwd()


class SecretSV:
    def __init__(self):
        pass
    
    def bankruptcy(self):
        shutil.rmtree(os.path.join(cwd, STORAGE1))
        shutil.rmtree(os.path.join(cwd, STORAGE2))
        os.mkdir(os.path.join(cwd, STORAGE1))
        os.mkdir(os.path.join(cwd, STORAGE2))
        conn = sqlite3.connect(os.path.join(cwd, DATABASE))
        cur = conn.cursor()
        cur.executescript(SQLSCRIPT1)
        conn.commit()
