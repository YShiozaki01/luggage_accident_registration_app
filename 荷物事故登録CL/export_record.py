import sqlite3
import csv
import shutil
import os
import sys

DATABASE = "database/database.db"
TABLE1 = "T添付資料"
TABLE2 = "T送信番号"
SQL1 = "SELECT * FROM T削除番号 WHERE 送信日時 is NULL;"
SQL2 = """
        SELECT 管理番号, 明細番号, max(削除日時) FROM T削除番号
        WHERE 送信日時 is NULL GROUP by 管理番号, 明細番号;
        """
SQL3 = f"""SELECT * FROM T荷物事故 WHERE 送信日時 is NULL;"""
SQL4 = f"""
        SELECT 
            a.管理番号,
            a.所管店,
            b.名称 as 報告部署名,
            a.荷主,
            c.名称 as 荷主名,
            a.締日,
            a.明細番号,
            a.責任店,
            d.名称 as 責任部署名,
            a.発生日,
            a.発生地,
            a.住所,
            e.住所 as 自治体名,
            a.発生状況,
            f.発生状況 as 状況,
            a.惹起者,
            g.名称 as 事故惹起者名,
            a.損害額,
            a.詳細,
            a.自他,
            a.更新日時
        FROM T荷物事故 as a
        LEFT JOIN M部署 as b on a.所管店 = b.コード
        LEFT JOIN M荷主 as c on a.荷主 = c.ID
        LEFT JOIN M部署 as d on a.責任店 = d.コード
        LEFT JOIN M住所 as e on a.住所 = e.郵便番号
        LEFT JOIN M発生状況 as f on a.発生状況 = f.ID
        LEFT JOIN M事故惹起者 as g on a.惹起者 = g.ID
        WHERE a.送信日時 is NULL;
        """
SQL5 = "SELECT * FROM T削除番号F WHERE 送信日時 is NULL;"
SQL6 = """
        SELECT ファイル名 FROM T削除番号F WHERE 送信日時 is NULL;
        """
FILENAME1 = "t_delnum.csv"
FILENAME2 = "t_main.csv"
FILENAME3 = "t_delnum_f.csv"


class ExportRecord:
    def __init__(self, send_time):
        self.send_time = send_time

    def export_table1(self):
        # T削除番号テーブルの、未送信のレコード（送信日時がNull）をcsv形式で出力
        # 送信日時がNullのレコードを抽出
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(SQL1)
        # 抽出結果が空でなければ
        if not len(cur.fetchall()) == 0:
            cur.execute(SQL2)
            result = cur.fetchall()
            # 抽出したレコードをCSVファイルに出力
            with open(f"{self.folder}/{FILENAME1}", "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(result)

    def export_table2(self):
        # T荷物事故テーブルの、未送信のレコード（送信日時がNull）をcsv形式で出力
        # 送信日時がNullのレコードを抽出
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(SQL3)
        # 抽出結果が空でなければ
        if not len(cur.fetchall()) == 0:
            cur.execute(SQL4)
            result = cur.fetchall()
            # 抽出したレコードをCSVファイルに出力
            with open(f"{self.folder}/{FILENAME2}", "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(result)

    def export_table3(self):
        # T削除番号Fテーブルの、未送信のレコード（送信日時がNull）をcsv形式で出力
        # 送信日時がNullのレコードを抽出
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(SQL5)
        # 抽出結果が空でなければ
        if not len(cur.fetchall()) == 0:
            cur.execute(SQL6)
            result = cur.fetchall()
            # 抽出したレコードをCSVファイルに出力
            with open(f"{self.folder}/{FILENAME3}", "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(result)

    def copy_file(self):
        # attachmentフォルダ内の添付資料のうち、T添付資料の送信日時が空白のファイルをフォルダへコピペ
        # 1)コピー元ファイルリストを作成
        sql = f"""SELECT ファイル名称 FROM {TABLE1} WHERE 送信日時 is NULL;"""
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        result = cur.fetchall()
        # 2)抽出されたファイルのみコピー
        for name in result:
            name_dst = os.path.basename(name[0])
            shutil.copyfile(name[0], f"{self.folder}/{name_dst}")

    def make_zip(self):
        # zipファイルを作成
        shutil.make_archive(self.zip_path, 'zip', root_dir=self.folder)

    def set_send_time(self):
        # テーブルの送信日時に日時をセット
        sql = f"""UPDATE T削除番号 SET 送信日時 = '{self.send_time}' WHERE 送信日時 is NULL;
                UPDATE T荷物事故 SET 送信日時 = '{self.send_time}' WHERE 送信日時 is NULL;
                UPDATE T添付資料 SET 送信日時 = '{self.send_time}' WHERE 送信日時 is NULL;
                UPDATE T削除番号F SET 送信日時 = '{self.send_time}' WHERE 送信日時 is NULL;"""
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.executescript(sql)
        conn.commit()

    def gen_zipname(self):
        # 添付ファイル（zip）用の名称を取得
        sql = f"SELECT * FROM {TABLE2};"
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        if len(cur.fetchall()) == 0:
            num = 1
        else:
            cur.execute(f"SELECT max(ID) FROM {TABLE2};")
            result = cur.fetchone()
            num = result[0] + 1
        zipname = f"{sys.argv[1]}_CASE_{num}"
        return zipname

    def input_send_history(self, name):
        # 送信日時を記録
        ID = int(name[8:])
        sql = f"""INSERT INTO {TABLE2}(ID, ファイル名, 送信日時)
                VALUES({ID}, '{name}', '{self.send_time}')"""
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()


if __name__ == "__main__":
    er = ExportRecord("")
    print(er.gen_zipname())
