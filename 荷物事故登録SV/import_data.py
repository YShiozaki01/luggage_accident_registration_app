import sqlite3
import csv
import os
import zipfile
import shutil
import datetime
from sql_tool import SqlTool

DATABASE = "database/database.db"
SOURCE1 = "t_delnum.csv"
SOURCE2 = "t_main.csv"
SOURCE3 = "t_delnum_f.csv"
STORAGE2 = "static/attachment"
STORAGE3 = "storage"
TEMPORARY = "temp"
SCRIPTSQL1 = """
            DELETE FROM TS荷物事故 WHERE (管理番号, 明細番号, 送信部署コード)
            in (SELECT 管理番号, 明細番号, 送信部署コード FROM TS削除番号);
            """
TABLE1 = "T最終取込ファイル名称"


def insert_source1(dep_cd):
    # 削除番号データをテーブルに挿入
    with open(f"{TEMPORARY}/{SOURCE1}", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        data = []
        for row in reader:
            data.append(row)
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        sql = f"""INSERT INTO TS削除番号 
            VALUES(?,?,?,'{dep_cd}')"""
        cur.executemany(sql, data)
        conn.commit()


def insert_source2(dep_cd):
    # 荷物事故データをテーブルに挿入
    with open(f"{TEMPORARY}/{SOURCE2}", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        data = []
        for row in reader:
            data.append(row)
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        sql1 = f"""INSERT INTO TS荷物事故(
            管理番号,
            報告部署コード,
            報告部署名,
            荷主コード,
            荷主名,
            締日,
            明細番号,
            責任部署コード,
            責任部署名,
            発生日,
            発生地,
            郵便番号,
            住所,
            発生状況コード,
            発生状況,
            惹起者コード,
            事故惹起者名,
            損害額,
            詳細,
            自他,
            更新日時,
            送信部署コード
            ) 
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'{dep_cd}')"""
        sql2 = "UPDATE TS荷物事故 SET 荷物事故コード = substr('000000' || 管理番号, -6, 6) || substr('000' || 明細番号, -3, 3)"
        cur.executemany(sql1, data)
        conn.commit()
        cur.execute(sql2)
        conn.commit()


def insert_source3():
    # 削除番号Fデータをテーブルに挿入
    with open(f"{TEMPORARY}/{SOURCE3}", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        data = []
        for row in reader:
            data.append(row)
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        sql = f"""INSERT INTO TS削除番号F(ファイル名) VALUES(?)"""
        cur.executemany(sql, data)
        conn.commit()


def delete_atch_file():
    # 添付ファイルを削除
    # 未処理のファイルがあれば以下を実行
    sql = """
        SELECT * FROM TS削除番号F WHERE 削除日時 is NULL;
        """
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    if len(cur.fetchall()) != 0:
        # 1)TS削除番号Fから削除するファイル名の一覧を取得
        sql = """
            SELECT ファイル名 FROM TS削除番号F WHERE 削除日時 is NULL;
            """
        cur.execute(sql)
        result = cur.fetchall()
        for c in result:
            # 2)TS添付資料からレコードを削除
            fname = c[0]
            name = os.path.splitext(fname)[0]
            sql = f"""
                DELETE FROM TS添付資料 WHERE ファイル名 = '{name}';
                """
            cur.execute(sql)
            conn.commit()
            # 3)attachmentフォルダの同ファイルを削除
            filepath = os.path.join(STORAGE2, fname)
            os.remove(filepath)
        # 4)削除日時を挿入
        dt_now = datetime.datetime.now()
        str_now_time = dt_now.strftime('%Y/%m/%d %H:%M:%S')
        sql = f"""
            UPDATE TS削除番号F SET 削除日時 = '{str_now_time}' WHERE 削除日時 is NULL;
            """
        cur.execute(sql)
        conn.commit()


# def copy_file(dep_cd): # 送信元の報告店コードを固定したため、添付ファイル名に送信元コードをつけるのを廃止
def copy_file():
    # 画像等添付ファイルを保存フォルダにコピー
    filelist = os.listdir(TEMPORARY)
    for file in filelist:
        # 荷物事故履歴CSV、削除番号CSV, 削除番号FCSVでなければフォルダへコピー
        if not file in (SOURCE1, SOURCE2, SOURCE3):
            # save_file_name = f"{STORAGE2}/{dep_cd}_{file}"
            save_file_name = f"{STORAGE2}/{file}"
            shutil.copy(f"{TEMPORARY}/{file}", save_file_name)
            # TS添付資料テーブルに添付資料ファイルの情報を書き込み
            # 1)テーブルの最新ID番号を取得
            # 1-1)もし添付資料テーブルが空ならIDは1 
            st = SqlTool(DATABASE, "TS添付資料")
            if st.dec_empt_table():
                id_num = 1
            # 1-2)最終番号の次
            else:
                id_num = st.get_new_number("ID")
            # 2)テーブル書き込み用パスを取得
            fpath = f"attachment/{file}"
            # 3)コピー先ファイルフルパスからファイル名（拡張子なし）と拡張子を取得
            fname, ext = os.path.splitext(os.path.basename(save_file_name))
            # 4)取得した情報をTS添付資料テーブルに書き込み
            sql = f"""
                    INSERT INTO TS添付資料 VALUES(
                    {id_num}, '{fname}', '{ext}', '{fpath}'
                    )
                """
            conn = sqlite3.connect(DATABASE)
            cur = conn.cursor()
            cur.execute(sql)
            conn.commit()


def update_name(name, number):
    # T最終取込ファイル名称テーブルを更新する
    # 1)今回取り込んだファイル名と前8文字が同じレコードを削除
    sql = f"""DELETE FROM {TABLE1} WHERE substr(ファイル名, 1, 8) = substr('{name}', 1, 8);
            INSERT INTO {TABLE1} VALUES('{name}', {number});"""
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.executescript(sql)
    conn.commit()


def get_last_filename(filename):
    # 前回取り込んだzipファイルの名前を取得
    sql = f"SELECT * FROM {TABLE1} WHERE substr(ファイル名, 1, 8) = substr('{filename}', 1, 8);"
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    # テーブルに該当するレコードの有無
    if len(cur.fetchall()) == 0:
        ret = "None"
    else:
        cur.execute(sql)
        result = cur.fetchone()
        ret = result[0]
    return ret


# ファイルの取込順をチェック
def check_number(filename):
    name = os.path.basename(filename)
    number_new = int(os.path.splitext(name)[0][8:])
    # 最終連番号を取得
    sql = f"SELECT * FROM {TABLE1} WHERE substr(ファイル名, 1, 8) = substr('{name}', 1, 8);"
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    # テーブルに該当するレコードの有無
    if len(cur.fetchall()) == 0:
        number_last = 0
    else:
        cur.execute(sql)
        result = cur.fetchone()
        number_last = result[1]
    # 最終連番号と新連番号を比較
    diff = number_new - number_last
    if diff == 1:
        txt_warning = ""
    elif diff < 1:
        txt_warning = "取込み済みの可能性があります。"
    elif diff > 1:
        txt_warning = "取り込んでいないデータがある可能性があります。"
    return txt_warning


def import_data(filename):
    # 選択した荷物事故データファイル（zip）を処理
    # 1)ファイル名から送信部署コードを取得
    name = os.path.basename(filename)
    depcd = name[:2]
    # 2)ZIPファイルをtempフォルダに展開
    with zipfile.ZipFile(filename, "r") as zf:
        zf.extractall(TEMPORARY)
    # 3)tempフォルダにt_delnum.csvがあれば
    if SOURCE1 in os.listdir(TEMPORARY):
        # 4)t_delnum.csvをTS削除番号テーブルにインポート
        insert_source1(depcd)
        # 5)TS荷物事故テーブルのレコードのうち、TS削除番号のレコードと一致するものを削除
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(SCRIPTSQL1)
        conn.commit()
    # 6)tempフォルダにt_main.csvがあれば
    if SOURCE2 in os.listdir(TEMPORARY):
        # 7)t_main.csvをTS荷物事故テーブルにインポート
        insert_source2(depcd)
    # 8)tempフォルダにt_delnum_f.csvがあれば
    if SOURCE3 in os.listdir(TEMPORARY):
        # 9)t_delnum_f.csvをTS削除番号Fテーブルにインポート
        insert_source3()
        # 10)t_delnum_fファイルに記載された既存の添付ファイルを削除
        delete_atch_file()
    # 11)画像等の添付ファイルを保存フォルダにコピー
    # copy_file(depcd)
    copy_file()
    # 12)インポートしたCSVファイルをtempフォルダごと削除
    shutil.rmtree(TEMPORARY)
    # 13)T最終取込ファイル名称テーブルを更新
    nametext = os.path.splitext(name)
    number = int(nametext[0][8:])
    update_name(name, number)


if __name__ == "__main__":
    # import_data()
    filename = "09_CASE_11.zip"
    lastname = get_last_filename(filename)
    print(f"前回: {lastname} → 今回: {filename}")
