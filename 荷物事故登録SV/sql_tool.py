import sqlite3


class SqlTool:
    def __init__(self, database, table):
        self.database = database
        self.table = table

    # テーブルが空か判定する
    def dec_empt_table(self):
        sql = f"SELECT * FROM {self.table};"
        conn = sqlite3.connect(self.database)
        cur = conn.cursor()
        cur.execute(sql)
        if len(cur.fetchall()) == 0:
            result = True
        else:
            result = False
        return result
    
    # 最新の番号を取得
    def get_new_number(self, field_name):
        sql = f"SELECT max({field_name}) FROM {self.table};"
        conn = sqlite3.connect(self.database)
        cur = conn.cursor()
        cur.execute(sql)
        result = cur.fetchone()
        new_num = result[0] + 1
        return new_num
