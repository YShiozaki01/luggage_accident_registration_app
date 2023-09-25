from flask import Flask, render_template, request
import sqlite3

app = Flask(__name__)

# DATABASE = "../荷物事故登録SV/database/database.db"
DATABASE = "./database/database.db"


def get_db_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


# TS荷物事故テーブルに登録されている報告部署の一覧を取得する
def get_reporting_department_list():
    # データベース接続
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    sql = """
        SELECT 報告部署コード, max(報告部署名) FROM TS荷物事故 GROUP by 報告部署コード;
        """
    cur.execute(sql)
    reporting_dept_list = cur.fetchall()
    return reporting_dept_list


# TS荷物事故テーブルに登録されている責任部署の一覧を取得する
def get_cause_department_list():
    # データベース接続
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    sql = """
        SELECT 責任部署コード, max(責任部署名) FROM TS荷物事故 GROUP by 責任部署コード;
        """
    cur.execute(sql)
    cause_dept_list = cur.fetchall()
    return cause_dept_list


# TS荷物事故テーブルに登録されている発生状況の一覧を取得する
def get_situation_list():
    # データベース接続
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    sql = """
        SELECT 発生状況コード, max(発生状況) FROM TS荷物事故 GROUP by 発生状況コード;  
        """
    cur.execute(sql)
    situation_list = cur.fetchall()
    return situation_list


# TS荷物事故テーブルに登録されている惹起者の一覧を取得する
def get_initiator_list():
    # データベース接続
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    sql = """
        SELECT 惹起者コード, max(事故惹起者名) FROM TS荷物事故 GROUP by 惹起者コード;
        """
    cur.execute(sql)
    initiator_list = cur.fetchall()
    return initiator_list


# TS荷物事故テーブルから入力した検索キーで検索したレコードを抽出する
def get_search_record(search_keys):
    key1 = search_keys["reporting_dept_cd"]     # 報告部署コード
    key2 = search_keys["consignor"]             # 荷主名
    # 締め日
    key3_1 = "0000/00/00" if search_keys["closing_date_from"] == "%" else search_keys["closing_date_from"]
    key3_2 = "9999/99/99" if search_keys["closing_date_to"] == "%" else search_keys["closing_date_to"]
    key4 = search_keys["cause_dept_cd"]         # 責任部署コード
    key5_1 = search_keys["place_of_occurrence"] # 発生地
    # 発生日
    key5_2 = "0000/00/00" if search_keys["accrual_date_from"] == "%" else search_keys["accrual_date_from"]
    key5_3 = "9999/99/99" if search_keys["accrual_date_to"] == "%" else search_keys["accrual_date_to"]
    key6 = search_keys["situation_cd"]      # 発生状況コード
    key7 = search_keys["initiator_cd"]      # 事故惹起者コード
    key8 = search_keys["sign"]
    # 損害額
    if search_keys["amount"] == "%":
        key9 = ""
    else:
        key9 = f"損害額 {key8} {search_keys['amount']} and"
    key10 = search_keys["key_word"]          # 詳細検索キーワード
    sql = f"""
        SELECT * FROM TS荷物事故
        WHERE 報告部署コード like '%{key1}%' and
        荷主名 like '%{key2}%' and
        締日 between '{key3_1}' and '{key3_2}' and
        責任部署コード like '%{key4}%' and
        発生地 like '%{key5_1}%' and
        発生日 between '{key5_2}' and '{key5_3}' and 
        発生状況コード like '%{key6}%' and
        惹起者コード like '%{key7}%' and
        {key9} 
        詳細 like '%{key10}%';
        """
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    return result


# テーブルから選択した履歴の詳細を取得
def search_record_one(no0, no1, no2):
    sql = f"""
        SELECT 荷物事故コード, 報告部署名, 発生日, 荷主名, 責任部署名, 発生地, 発生状況,
        事故惹起者名 as 事故惹起者, 損害額, 詳細 as 事故の詳細
        FROM TS荷物事故
        WHERE 送信部署コード = '{no0}' and 管理番号 = {no1} and 明細番号 = {no2}
        """
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchone()
    return result


@app.route("/")
def index():
    dept_list1 = get_reporting_department_list()
    dept_list2 = get_cause_department_list()
    situation_list = get_situation_list()
    initiator_list = get_initiator_list()
    return render_template("main.html", dept_list1=dept_list1,
                           dept_list2=dept_list2,
                           situation_list=situation_list,
                           initiator_list=initiator_list)


@app.route("/result", methods=["GET", "POST"])
def result():
    if request.method == "POST":
        keys = request.form.to_dict()
        # 報告部署コード
        if not "reporting_dept" in keys:
            reporting_dept_cd = "%"
        else:
            reporting_dept_cd = keys["reporting_dept"]
        # 荷主名（一部）
        if keys["consignor"] == "":
            consignor = "%"
        else:
            consignor = keys["consignor"]
        # 締め日(from)
        if keys["closing_date_from"] == "":
            closing_date_from = "%"
        else:
            closing_date_from = "{0}/{1}/{2}".format(
                keys["closing_date_from"][:4], keys["closing_date_from"][4:6], keys["closing_date_from"][6:])
        # 締め日(to)
        if keys["closing_date_to"] == "":
            closing_date_to = "%"
        else:
            closing_date_to = "{0}/{1}/{2}".format(
                keys["closing_date_to"][:4], keys["closing_date_to"][4:6], keys["closing_date_to"][6:])
        # 責任部署コード
        if not "cause_dept" in keys:
            cause_dept_cd = "%"
        else:
            cause_dept_cd = keys["cause_dept"]
        # 発生地（一部）
        if keys["place_of_occurrence"] == "":
            place_of_occurrence = "%"
        else:
            place_of_occurrence = keys["place_of_occurrence"]
        # 発生日(from)
        if keys["accrual_date_from"] == "":
            accrual_date_from = "%"
        else:
            accrual_date_from  = "{0}/{1}/{2}".format(
                keys["accrual_date_from"][:4], keys["accrual_date_from"][4:6], keys["accrual_date_from"][6:])
        # 発生日(to)
        if keys["accrual_date_to"] == "":
            accrual_date_to = "%"
        else:
            accrual_date_to  = "{0}/{1}/{2}".format(
                keys["accrual_date_to"][:4], keys["accrual_date_to"][4:6], keys["accrual_date_to"][6:])
        # 発生状況コード
        if not "situation" in keys:
            situation_cd = "%"
        else:
            situation_cd = keys["situation"]
        # 惹起者
        if not "initiator" in keys:
            initiator_cd = "%"
        else:
            initiator_cd = keys["initiator"]
        # 損害額
        if keys["amount"] == "":
            amount = "%"
        else:
            amount = keys["amount"]
        # 不等号記号
        if keys["sign"] == "0":
            sign = "="
        elif keys["sign"] == "1":
            sign = ">="
        else:
            sign = "<="
        # 詳細（一部）
        if keys["key_word"] == "":
            key_word = "%"
        else:
            key_word = keys["key_word"]
        search_keys = {
            "reporting_dept_cd": reporting_dept_cd,
            "consignor": consignor,
            "closing_date_from": closing_date_from,
            "closing_date_to": closing_date_to,
            "cause_dept_cd": cause_dept_cd,
            "place_of_occurrence": place_of_occurrence,
            "accrual_date_from": accrual_date_from,
            "accrual_date_to": accrual_date_to,
            "situation_cd": situation_cd,
            "initiator_cd": initiator_cd,
            "amount": amount,
            "sign": sign,
            "key_word": key_word
        }
        records = get_search_record(search_keys)
        return render_template("main.html", records=records)


@app.route("/show_detail/<no0>/<no1>/<no2>")
def show_detail(no0, no1, no2):
    rec_one = search_record_one(no0, no1, no2)
    # 添付ファイルのパスの一覧を取得
    # 1)ファイルパス検索用のコードを生成
    key_cd = f"{no0}_{no1.zfill(6)}{no2.zfill(3)}"
    # 2)TS添付資料のファイルパスの一覧を取得
    sql = f"""
            SELECT ファイルパス FROM TS添付資料
            WHERE substr(ファイル名, 1, 12) = '{key_cd}';
        """
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    path_list = []
    for c in result:
        path_list.append(c[0])
    return render_template("detail_record.html", rec_one=rec_one,
                           path_list=path_list)


if __name__ == "__main__":
    app.debug = True
    app.run(host="0.0.0.0", port="5000")
