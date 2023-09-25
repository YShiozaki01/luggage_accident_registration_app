import sqlite3

DATABASE = "database/database.db"
TABLE = "T荷物事故"


def get_db_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


class InsertRecord:
    def __init__(self):
        pass


    def get_new_number(self):
        # 入力データに付加する管理番号を取得
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"SELECT * FROM {TABLE};"
        cur.execute(sql)
        # もし、T荷物事故テーブルが空なら、管理番号＝1
        if len(cur.fetchall()) == 0:
            control_number = 1
        else:
            sql = f"""SELECT * FROM {TABLE} WHERE 所管店 = '{self.cel1}'
                    AND 荷主 = {self.cel2} AND 締日 = '{self.cel3}';"""
            cur.execute(sql)
            # もし、入力中のデータと同じ報告店且つ同じ荷主、締日のレコードかあれば、それと同じ番号
            if len(cur.fetchall()) != 0:
                sql = f"""SELECT max(管理番号) as 最終番号 FROM {TABLE} WHERE 所管店 = '{self.cel1}'
                        AND 荷主 = {self.cel2} AND 締日 = '{self.cel3}';"""
                cur.execute(sql)
                result = cur.fetchone()
                control_number = result["最終番号"]
            else:
                sql = f"""SELECT max(管理番号) as 最終番号 FROM {TABLE};"""
                cur.execute(sql)
                result = cur.fetchone()
                control_number = result["最終番号"] + 1
        return control_number


    def get_new_dtlnum(self):
        # 入力データに付加する管理番号を取得
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"SELECT * FROM {TABLE};"
        cur.execute(sql)
        # もし、T荷物事故テーブルが空なら、明細番号＝1
        if len(cur.fetchall()) == 0:
            detail_number = 1
        else:
            sql = f"""SELECT * FROM {TABLE} WHERE 所管店 = '{self.cel1}'
                    AND 荷主 = {self.cel2} AND 締日 = '{self.cel3}';"""
            cur.execute(sql)
            # もし、入力中のデータと同じ報告店且つ同じ荷主、締日のレコードかあれば、同条件の最大の明細番号の次の番号
            if len(cur.fetchall()) != 0:
                sql = f"""SELECT max(明細番号) as 最終番号 FROM {TABLE} WHERE 所管店 = '{self.cel1}'
                        AND 荷主 = {self.cel2} AND 締日 = '{self.cel3}';"""
                cur.execute(sql)
                result = cur.fetchone()
                detail_number = result["最終番号"] + 1
            else:
                detail_number = 1
        return detail_number
    

    def insert_table(self, control_number, detail_number, update_time):
        # T荷物事故テーブルに入力値を書き込み
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"""
                INSERT INTO T荷物事故(
                            管理番号,
                            所管店,
                            荷主,
                            締日,
                            明細番号,
                            責任店,
                            発生日,
                            発生地,
                            住所,
                            発生状況,
                            惹起者,
                            損害額,
                            詳細,
                            自他,
                            更新日時
                            )
                            VALUES(
                            {control_number},
                            '{self.cel1}',
                            '{self.cel2}',
                            '{self.cel3}',
                            {detail_number},
                            '{self.cel4}',
                            '{self.cel5}',
                            '{self.cel6}',
                            '{self.cel7}',
                            '{self.cel8}',
                            '{self.cel9}',
                            {self.cel10},
                            '{self.cel11}',
                            '{self.cel12}',
                            '{update_time}'
                            )"""
        cur.execute(sql)
        conn.commit()

    def input_history(self, control_number):
        # テーブル表示用の入力履歴データを取得する
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"""SELECT substr('000000' || a.管理番号, -6, 6) || substr('000' || a.明細番号, -3, 3) as 事故番号,
            b.名称 as 荷主名, a.発生日, a.発生地
            FROM T荷物事故 as a
            LEFT JOIN M荷主 as b
            on a.荷主 = b.ID
            WHERE a.管理番号 = {control_number}
            ORDER by a.明細番号 DESC;"""
        cur.execute(sql)
        result = [list(row) for row in cur.fetchall()]
        return result

    def get_input_data(self):
        # 管理番号を取得する
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"""SELECT * FROM T荷物事故 WHERE 所管店 = '{self.cel1}'
                AND 荷主 = {self.cel2} AND 締日 = '{self.cel3}'
                GROUP by 管理番号"""
        cur.execute(sql)
        if len(cur.fetchall()) == 0:
            data_list = "no_data"
        else:
            sql = f"""SELECT 管理番号 FROM T荷物事故 WHERE 所管店 = '{self.cel1}'
                    AND 荷主 = {self.cel2} AND 締日 = '{self.cel3}'
                    GROUP by 管理番号"""
            cur.execute(sql)
            result = cur.fetchone()
            control_number = result["管理番号"]
            data_list = self.input_history(control_number)
        return data_list

    def set_redo(self):
        # リストから選択した入力済みデータを再取得        
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"""SELECT 事故番号, 管理番号, 明細番号, 責任店, 責任店名称, 発生日, 発生地, 郵便番号, 住所, 発生状況,
                発生状況名称, 惹起者, 惹起者名称, 自他, 損害額, 詳細
                FROM
                (SELECT substr('000000' || a.管理番号, -6, 6) || substr('000' || a.明細番号, -3, 3) as 事故番号,
                a.管理番号, a.明細番号, a.責任店, b.名称 as 責任店名称, a.発生日, a.発生地, a.住所 as 郵便番号, c.住所,
                a.発生状況, d.発生状況 as 発生状況名称, a.惹起者, e.名称 as 惹起者名称, a.自他, a.損害額, a.詳細
                FROM T荷物事故 as a
                LEFT JOIN M部署 as b on a.責任店 = b.コード
                LEFT JOIN M住所 as c on a.住所 = c.郵便番号
                LEFT JOIN M発生状況 as d on a.発生状況 = d.ID
                LEFT JOIN M事故惹起者 as e on a.惹起者 = e.ID)
                WHERE 事故番号 = '{self.case_number}';"""
        cur.execute(sql)
        result = cur.fetchone()
        return result


    def delete_record(self, control_number, detail_number, update_date):
        # 選択されたレコードを削除
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"""DELETE FROM T荷物事故 WHERE 管理番号 = {control_number}
                AND 明細番号 = {detail_number};"""
        cur.execute(sql)
        conn.commit()
        # 削除した管理番号、明細番号をT削除番号へ書き込む
        sql = f"""INSERT INTO T削除番号(管理番号, 明細番号, 削除日時) VALUES(
                {control_number}, {detail_number}, '{update_date}');"""
        cur.execute(sql)
        conn.commit()


    def search_record(self, control_number, detail_number):
        # 指定の管理番号且つ明細番号のレコードが存在するか判定
        c = control_number if control_number else 0
        d = detail_number if detail_number else 0
        conn = get_db_connection()
        cur = conn.cursor()
        sql = f"""SELECT * FROM T荷物事故 WHERE 管理番号 = {c} AND 明細番号 = {d};"""
        cur.execute(sql)
        result = bool(len(cur.fetchall()))
        return result

