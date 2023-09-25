import PySimpleGUI as sg
import openpyxl
import os
import import_data
import sqlite3
from secret import SecretSV
from select_item import SelectItem
from make_report import PostingRecord

sg.theme("SystemDefault")
DATABASE = "database/database.db"
WORKBOOK = "運送保険事故連絡票兼現認書.xlsm"
TARGETAMOUNT = 20000
OUTPUTDIR = "output/"


def get_db_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


# select_itemウインドウ用の部署のリストを取得
def get_department_list():
    sql = "SELECT 報告部署コード, 報告部署名 FROM TS荷物事故 GROUP by 報告部署コード;"
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = list(cur.fetchall())
    return result


# select_itemウインドウ用の荷主のリストを取得
def get_consignor_list(dept_cd):
    sql = f"""
        SELECT 荷主コード, 荷主名 FROM TS荷物事故 WHERE 報告部署コード like '%{dept_cd}%' GROUP by 荷主コード;
        """
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = list(cur.fetchall())
    return result


# 締め日コンボボックス用の締め日のリストを取得
def get_date_list(cons_cd):
    sql = f"""
        SELECT 締日 FROM TS荷物事故 WHERE 荷主コード = '{cons_cd}' GROUP by 締日 ORDER by 締日 DESC;
        """
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    date_list = []
    for c in result:
        date_list.append(c[0])
    return date_list


# Excelへ転記する保険適用対象データを取得
def get_posting_data(cons_cd, close_date):
    sql = f"""
        SELECT * FROM TS荷物事故 WHERE 荷主コード = '{cons_cd}' and 締日 = '{close_date}'
        and 損害額 > {TARGETAMOUNT} and 自他 <> '1';
        """
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    return result


# 作成したExcelファイルの保存名を生成
def get_save_fname(cons_cd, close_date):
    sql = f"""
        SELECT 報告部署コード, 荷主コード, 締日 FROM TS荷物事故
        WHERE 荷主コード = '{cons_cd}' and 締日 = '{close_date}'
        GROUP by 報告部署コード, 荷主コード, 締日;
        """
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchone()
    name1 = result["報告部署コード"]
    name2 = result["荷主コード"]
    name3 = f"{result['締日'][0:4]}{result['締日'][5:7]}{result['締日'][8:10]}"
    save_name = f"{name1}_{name2}_{name3}"
    return save_name


# 添付ファイル検索用のキーコードを取得
def get_key_code(cons_cd, close_date):
    sql = f"""
            SELECT 送信部署コード, 荷物事故コード FROM TS荷物事故
            WHERE 荷主コード = '{cons_cd}' and 締日 = '{close_date}';
        """
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    keys = []
    for c in result:
        key = f"{c[0]}_{c[1]}"
        keys.append(key)
    return keys


frame_layout = [[sg.Push()],
                [sg.T("部署", font=("Yu Gothic UI", 8), size=(6, 0)), sg.B("検索", font=("Yu Gothic UI", 8), size=(5, 0), k="btn_src1"),
                 sg.I(font=("Yu Gothic UI", 8), size=(
                     35, 0), k="in_dept", readonly=True),
                 sg.I(font=("Yu Gothic UI", 8), size=(5, 0), k="in_cd1", visible=False)],
                [sg.T("荷主", font=("Yu Gothic UI", 8), size=(6, 0)), sg.B("検索", font=("Yu Gothic UI", 8), size=(5, 0), k="btn_src2"),
                 sg.I(font=("Yu Gothic UI", 8), size=(
                     35, 0), k="in_cons", readonly=True),
                 sg.I(font=("Yu Gothic UI", 8), size=(5, 0), k="in_cd2", visible=False)],
                [sg.T("締日", font=("Yu Gothic UI", 8), size=(6, 0)),
                 sg.Combo(values=[""], font=("Yu Gothic UI", 8), size=(20, 0), k="cmb_date")],
                [sg.Push()],
                [sg.B("キャンセル", font=("Yu Gothic UI", 8), size=(10, 0), k="btn_cancel"),
                 sg.Push(), sg.B("印刷", font=("Yu Gothic UI", 8), size=(10, 0), k="btn_prt")]]
layout = [[sg.T("受信ファイル取込み", font=("Yu Gothic UI", 11), k="txt_title")],
          [sg.InputText(font=("Yu Gothic UI", 8), size=(50, 0), k="intt1", readonly=True),
           sg.FileBrowse("選択", font=("Yu Gothic UI", 8),
                         size=(6, 0), key="fb1", target="intt1"),
           sg.B("取込み", font=("Yu Gothic UI", 8), size=(10, 0), k="btn_opn")],
          [sg.Push()],
          [sg.Push()],
          [sg.Frame(title="運送保険事故連絡票印刷", font=(
              "Yu Gothic UI", 8), layout=frame_layout, size=(430, 140))]]
win = sg.Window("受信ファイル取込み処理", layout, font=(
    "Yu Gothic UI", 11), size=(460, 235), disable_close=False)
win.finalize()

# タイトルの文字をダブルクリック
win["txt_title"].bind("<Double-ButtonPress-1>", "_secret")


# 稟議書印刷用に入力した値を全てクリア
def clear_input():
    win["in_dept"].update("")
    win["in_cd1"].update("")
    win["in_cons"].update("")
    win["in_cd2"].update("")
    win["cmb_date"].update("")


while True:
    e, v = win.read()
    if e == "btn_opn":
        # ダイアログで選択したファイルのパスを取得
        filepath = v["intt1"]
        # パスからファイル名を取得
        zipfilename = os.path.basename(filepath)
        # 拡張子なしのファイル名を取得
        filename = os.path.splitext(zipfilename)[0]
        # 前回取込み処理したzipファイルの名称を取得
        lastfilename = import_data.get_last_filename(filename)
        # 間違ったデータである可能性があるかチェック
        msg = import_data.check_number(filepath)
        # ポップアップで表示し、処理を継続するか確認
        answ = sg.popup_ok_cancel(f"""今回：　{zipfilename}\n前回：　{lastfilename}
                                    \n\n{msg}このファイルを取り込みますか？""",
                                  font=("Yu Gothic UI", 8), title="確認")
        # 継続
        if answ == "OK":
            import_data.import_data(filepath)
            win["intt1"].update("")
            sg.popup("Zipファイルを取り込みました。", font=("Yu Gothic UI", 8))
        # 中止
        else:
            sg.popup("取込み処理を中止しました。", font=("Yu Gothic UI", 8))
    if e == "btn_src1":
        item_list = get_department_list()
        si1 = SelectItem("部署選択", item_list)
        row_num = si1.get_select_item()
        if not row_num == None:
            item = item_list[row_num]
            win["in_dept"].update(item[1])
            win["in_cd1"].update(item[0])
    if e == "btn_src2":
        if v["in_cd1"] == "":
            item_list = get_consignor_list("%")
        else:
            item_list = get_consignor_list(v["in_cd1"])
        si2 = SelectItem("荷主選択", item_list)
        row_num = si2.get_select_item()
        if not row_num == None:
            item = item_list[row_num]
            win["in_cons"].update(item[1])
            win["in_cd2"].update(item[0])
            date_list = get_date_list(item[0])
            win.find_element("cmb_date").Update(values=date_list)
    if e == "btn_cancel":
        clear_input()
    if e == "btn_prt":
        if v["cmb_date"]:
            wb = openpyxl.load_workbook(WORKBOOK)
            ws_org = wb["Sheet1"]
            # 運送保険適用該当データを取得
            records = get_posting_data(v["in_cd2"], v["cmb_date"])
            if not records:
                sg.popup("保険請求対象データがありません。", font=("Yu Gothic UI", 8))
            else:
                # 運送保険事故連絡票兼現認書に転記
                for posting_data in records:
                    pr = PostingRecord(wb, posting_data)
                    pr.loading_worksheet()
                wb.remove(ws_org)
                # 保存用ファイル名を取得
                name = get_save_fname(v["in_cd2"], v["cmb_date"])
                # 保存用のディレクトリを作成
                f_name = f"rp_{name}.xlsx"
                save_name = f"{OUTPUTDIR}/{f_name}"
                wb.save(save_name)
                sg.popup("運送保険事故連絡票兼現認書を作成しました。", font=("Yu Gothic UI", 8), title="確認")
        else:
            sg.popup("締日を入力してください。", font=("Yu Gothic UI", 8))
    if e == "txt_title_secret":
        answ = sg.popup_ok_cancel("このアプリを初期化します。\nよろしいですか？", font=("Yu Gothic UI", 8))
        if answ == "OK":
            srs = SecretSV()
            srs.bankruptcy()
        else:
            continue
    if e == None:
        break
win.close()
