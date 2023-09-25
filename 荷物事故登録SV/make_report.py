import sqlite3
import datetime
import os

today = datetime.date.today()

WORKBOOK = "運送保険事故連絡票兼現認書.xlsm"
WORKSHEET = "Sheet1"
DATABASE = "database/database.db"


# 「運送保険事故連絡票兼現認書」に転記するデータを生成する
class PostingRecord:
    def __init__(self, wb, record):
        self.wb = wb
        self.department = record["責任部署名"]
        self.control_number = record["荷物事故コード"]
        self.accrual_date = record["発生日"]
        self.consignor = record["荷主名"]
        self.place_of_occurrence = record["発生地"]
        self.address = record["住所"]
        self.occurrence_status = record["発生状況"]
        self.initiator = record["事故惹起者名"]
        self.detail = record["詳細"]
        self.amount_of_damages = record["損害額"]
        self.today = f"{today.year}年{today.month}月{today.day}日"
        self.snd_dept = record["送信部署コード"]

    # 「運送保険事故連絡票兼現認書」に値を転記
    def loading_worksheet(self):
        wb = self.wb
        ws_org = wb[WORKSHEET]
        ws = wb.copy_worksheet(ws_org)
        ws["O3"] = self.today
        ws["E4"] = self.department
        ws["N4"] = self.control_number
        ws["E5"] = self.accrual_date
        ws["E6"] = self.consignor
        ws["E7"] = f"{self.place_of_occurrence}（{self.address}）"
        ws["E8"] = self.occurrence_status
        ws["E9"] = self.initiator
        ws["E10"] = self.detail
        strprice = '¥' + '{:,.0f}'.format(self.amount_of_damages)
        ws["E17"] = strprice
        ws["O21"] = self.today
        ws.title = self.control_number
        # 「添付資料」欄に添付資料のファイルパスを転記して、ハイパーリンクに変換
        f_name = f"{self.snd_dept}_{self.control_number}"
        sql = f"""
            SELECT ファイル名 || 拡張子 FROM TS添付資料 WHERE substr(ファイル名, 1, 12) = '{f_name}';
            """
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        result = cur.fetchall()
        i = 3
        for c in result:
            i += 1
            filepath = os.path.join(os.getcwd(), "static", "attachment", c[0])
            ws[f"U{i}"] = f"資料{i-3}"
            ws[f"U{i}"].hyperlink = filepath
        # 印刷範囲を設定（「添付資料」欄を印刷しない）
        print_area = 'A1:S26'
        ws.print_area = print_area
