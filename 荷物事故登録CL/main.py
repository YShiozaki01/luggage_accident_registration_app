import PySimpleGUI as sg
import pandas as pd
import datetime
import os
import shutil
import sys
import sqlite3
import glob
from selectitem import SelectItem
from insert_record import InsertRecord
from loadimage import LoadImage
from send_mail import SendMail
from export_record import ExportRecord
from secret import Secret
from attachment_list import AttachmentList

sg.theme("SystemDefault")
DATABASE = "database/database.db"
AFOLDER = "attachment"
# WINLOCATION = ()
MAILTO = "shiozaki@egw.co.jp" #事故履歴データ送信先メールアドレス
# MAILTO = "yukio_s@rb3.so-net.ne.jp"


# バッチファイルに記録した部署コートの部署名を取得する
def get_department_name():
    cd = sys.argv[1]
    sql = f"SELECT 名称 FROM M部署 WHERE コード = '{cd}';"
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchone()
    name = result[0]
    header = {"code": cd, "name": name}
    return header


def get_now_time():
    # 現在日時を取得
    dt_now = datetime.datetime.now()
    str_now_time = dt_now.strftime('%Y/%m/%d %H:%M:%S')
    return str_now_time


def change_date(str_date):
    # 8桁数字を日付に変換
    d = str_date
    date = pd.to_datetime(d)
    format_date = format(date, "%Y/%m/%d")
    return format_date


# 事故管理番号を取得
def get_control_number(cnum, dnum):
    sql = f"""
        SELECT 所管店, 管理番号, 明細番号 FROM T荷物事故
        WHERE 管理番号 = {cnum} and 明細番号 = {dnum};
        """
    conn = sqlite3.connect(DATABASE)
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchone()
    control_number = f"{result[0]}_{str(result[1]).zfill(6)}{str(result[2]).zfill(3)}"
    return control_number


def insert_data():
    # T荷物事故テーブルに入力した事故履歴を書き込む
    ir = InsertRecord()
    # T荷物事故テーブルに同じ管理番号＋明細番号（＝事故番号）のレコードがあれば削除
    if ir.search_record(v["in_cnum"], v["in_dnum"]):
        # 削除レコードの管理番号（添付資料名用）を取得
        control_number_old = get_control_number(v["in_cnum"], v["in_dnum"])
        # レコードを削除
        ir.delete_record(v["in_cnum"], v["in_dnum"], get_now_time())
    # 管理番号を取得
    ir.cel1 = v["in_cd1"]
    ir.cel2 = v["in_cd2"]
    ir.cel3 = v["in1"]
    control_number = ir.get_new_number()
    # 明細番号を取得
    detail_number = ir.get_new_dtlnum()
    ir.cel4 = v["in_cd3"]
    ir.cel5 = v["in2"]
    ir.cel6 = v["in3"]
    ir.cel7 = v["in_cd4"]
    ir.cel8 = v["in_cd5"]
    ir.cel9 = v["in_cd6"]
    ir.cel10 = v["in4"] if v["in4"] else 0
    ir.cel11 = v["ml1"]
    ir.cel12 = v["in_cd7"]
    # 更新日を取得
    update_time = get_now_time()
    # T荷物事故テーブルに入力値を書き込み
    ir.insert_table(control_number, detail_number, update_time)
    # 入力履歴テーブルを更新
    data_dct = ir.input_history(control_number)
    win_main["hlst"].update(data_dct)
    # もしin_cnumに値があれば（＝既存の履歴を修正削除）以下を実行
    if v["in_cnum"]:
        # T添付資料に旧管理番号且つ旧明細番号の履歴があれば以下を実行
        sql = f"""
            SELECT * FROM T添付資料 WHERE 管理番号 = {v["in_cnum"]} and 明細番号 = {v["in_dnum"]};
            """
        conn = sqlite3.connect(DATABASE)
        cur = conn.cursor()
        cur.execute(sql)
        if len(cur.fetchall()) != 0:
            # 新管理番号を取得
            control_number_new = get_control_number(control_number, detail_number)
            # T添付資料に該当する管理番号且つ明細番号のレコードが有れば新番号に書き換え
            sql = f"""
                UPDATE T添付資料 SET 管理番号 = {control_number}, 明細番号 = {detail_number},
                ファイル名称 = '{AFOLDER}' || '/' || {v['in_cd1']} || '_' || substr('000000' || {control_number}, -6, 6)
                || substr('000' || {detail_number}, -3, 3) || '_' || substr('00' || 資料連番号, -2, 2) || 拡張子,
                送信日時 = NULL 
                WHERE 管理番号 = {v["in_cnum"]} and 明細番号 ={v["in_dnum"]};
                """
            cur.execute(sql)
            conn.commit()
            # 「attachiment」フォルダ内のファイルを新番号に名称を書き換え
            files = glob.glob(f"./{AFOLDER}/*")
            for file in files:
                filename = os.path.basename(file)
                name, ext = os.path.splitext(filename)
                if name[:12] == control_number_old:
                    print(name)
                    new_fname = f"{control_number_new}{name[12:]}{ext}"
                    new_name_full = os.path.join(os.path.dirname(file), new_fname)
                    os.rename(file, new_name_full)
    # 入力欄をクリア
    clear_field()
    # 削除ボタンを無効化
    win_main["btn_delete"].update(disabled=True)
    # フォーカスをbtn3に移動
    win_main["btn3"].SetFocus(True)


def clear_field():
    # データ登録後やキャンセルで入力欄をクリアする
    num_in = ["2", "3", "4"]
    num_ro_in = ["3", "4", "5", "6"]
    num_in_cd = ["3", "4", "5", "6", "7"]
    # 入力欄をクリア1
    for i in num_in:
        win_main[f"in{i}"].update("")
    # 入力欄をクリア2
    for i in num_ro_in:
        win_main[f"ro_in{i}"].update("")
    # 入力欄をクリア3
    for i in num_in_cd:
        win_main[f"in_cd{i}"].update("")
    # 「詳細」欄をクリア
    win_main["ml1"].update("")
    # 管理番号欄、明細番号欄をクリア
    win_main["in_cnum"].update("")
    win_main["in_dnum"].update("")
    # 添付ファイルパス欄をクリア
    win_main["intt1"].update("")


def attachment_valid(flg=False):
    # 添付資料登録欄の有効・無効切り替え
    win_main["fb1"].update(disabled=flg)
    win_main["btn_opn"].update(disabled=flg)


def check_input():
    checklist = {}
    checklist["in_cd1"] = 1 if v["in_cd1"] else 0
    checklist["in_cd2"] = 1 if v["in_cd2"] else 0
    checklist["in1"] = 1 if v["in1"] else 0
    checklist["in_cd3"] = 1 if v["in_cd3"] else 0
    checklist["in2"] = 1 if v["in2"] else 0
    checklist["in3"] = 1 if v["in3"] else 0
    checklist["in_cd4"] = 1 if v["in_cd4"] else 0
    checklist["in_cd5"] = 1 if v["in_cd5"] else 0
    checklist["in_cd6"] = 1 if v["in_cd6"] else 0
    return checklist


def make_window_main():
    # 報告店コード・名称を取得
    header = get_department_name()
    # メイン画面
    frame_layout = [[sg.InputText(font=("Yu Gothic UI", 8), size=(50, 0), k="intt1", readonly=True),
                    sg.FileBrowse("選択", font=("Yu Gothic UI", 8), size=(6, 0), key="fb1", target="intt1", disabled=True),
                    sg.B("開く", font=("Yu Gothic UI", 8), size=(10, 0), k="btn_opn", disabled=True)]]
    layout = [[sg.T("報告店", size=(8, 0), font=("Yu Gothic UI", 9)),
             sg.I(default_text=header["name"], font=("Yu Gothic UI", 9), size=(30, 0), k="ro_in1", readonly=True),
             sg.I(default_text=header["code"], font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd1", visible=False, readonly=True)],
            [sg.T("荷主", size=(8, 0), font=("Yu Gothic UI", 9)), sg.B("検索", k="btn2", font=("Yu Gothic UI", 8), size=(4, 0)),
             sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="ro_in2", readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd2", visible=False, readonly=True)],
            [sg.T("締日", size=(8, 0), font=("Yu Gothic UI", 9), k="txt_closedate"), sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="in1"),
             sg.Push(), sg.B("入力履歴検索", font=("Yu Gothic UI", 8), k="btn_search")],
            [sg.T("責任店", size=(8, 0), font=("Yu Gothic UI", 9)), sg.B("検索", k="btn3", font=("Yu Gothic UI", 8), size=(4, 0)),
             sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="ro_in3", readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd3", visible=False, readonly=True)],
            [sg.T("発生日", size=(8, 0), font=("Yu Gothic UI", 9)), sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="in2")],
            [sg.T("発生地", size=(8, 0), font=("Yu Gothic UI", 9)), sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="in3")],
            [sg.T("住所", size=(8, 0), font=("Yu Gothic UI", 9)), sg.B("検索", k="btn4", font=("Yu Gothic UI", 8), size=(4, 0)), 
             sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="ro_in4", readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd4", visible=False, readonly=True)],
            [sg.T("発生状況", size=(8, 0), font=("Yu Gothic UI", 9)), sg.B("検索", k="btn5", font=("Yu Gothic UI", 8), size=(4, 0)), 
             sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="ro_in5", readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd5", visible=False, readonly=True)],
            [sg.T("惹起者", size=(8, 0), font=("Yu Gothic UI", 9)), sg.B("検索", k="btn6", font=("Yu Gothic UI", 8), size=(4, 0)),
             sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="ro_in6", readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd6", visible=False, readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cd7", visible=False, readonly=True)],
            [sg.T("損害額", size=(8, 0), font=("Yu Gothic UI", 9)), sg.I(font=("Yu Gothic UI", 9), size=(30, 0), k="in4")],
            [sg.T("詳細", size=(8, 0), font=("Yu Gothic UI", 9)), sg.ML(font=("Yu Gothic UI", 9), size=(60, 5), k="ml1")],
            [sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_cnum", visible=False, readonly=True),
             sg.I(font=("Yu Gothic UI", 9), size=(10, 0), k="in_dnum", visible=False, readonly=True),
             sg.Push(), sg.B("キャンセル", k="btn_cancel", font=("Yu Gothic UI", 8), size=(10, 0), disabled=True),
             sg.B("登録", k="btn_register", font=("Yu Gothic UI", 8), size=(10, 0), disabled=True),
             sg.B("削除", k="btn_delete", disabled=True, font=("Yu Gothic UI", 8), size=(10, 0))],
            [sg.Table([""], ["№", "荷主", "発生日", "発生地"], font=("Yu Gothic UI", 10), justification="left", col_widths=(10, 20, 10, 20),
                      auto_size_columns=False, select_mode=sg.TABLE_SELECT_MODE_BROWSE, k="hlst", enable_events=True,
                      vertical_scroll_only=False)],
            [sg.Frame(title="添付ファイル", font=("Yu Gothic UI", 8), layout=frame_layout)],
            [sg.B("送信", k="btn_submit", font=("Yu Gothic UI", 8), size=(10, 0), disabled=False),
             sg.Push(), sg.B("閉じる", k="btn_close", font=("Yu Gothic UI", 8), size=(10, 0))]]
    win = sg.Window("荷物事故登録", layout, font=("Yu Gothic UI", 11), size=(480, 680), disable_close=True)
    win.finalize()
    return win

ir = InsertRecord()

win_main = make_window_main()
# win_main["btn1"].bind("<Return>", "_rtn")
win_main["btn2"].bind("<Return>", "_rtn")
win_main["in1"].bind("<Return>", "_rtn")
win_main["btn3"].bind("<Return>", "_rtn")
win_main["in2"].bind("<Return>", "_rtn")
win_main["in3"].bind("<Return>", "_rtn")
win_main["btn4"].bind("<Return>", "_rtn")
win_main["btn5"].bind("<Return>", "_rtn")
win_main["btn6"].bind("<Return>", "_rtn")
win_main["in4"].bind("<Return>", "_rtn")
# 「締日」の文字をダブルクリック
win_main["txt_closedate"].bind("<Double-ButtonPress-1>", "_secret")

record = {}
while True:
    e, v = win_main.read()
    # if e == "btn1" or e == "btn1_rtn":
    #     si = SelectItem(1)
    #     si.title = "報告店"
    #     result = si.open_window_sub()
    #     if not len(result) == 0:
    #         win_main["ro_in1"].update(result[1])
    #         win_main["in_cd1"].update(result[0])
    #         win_main["btn2"].SetFocus(True)
    #     else:
    #         continue
    if e == "btn2" or e == "btn2_rtn":
        si = SelectItem(2)
        si.title = "荷主"
        result = si.open_window_sub()
        if not len(result) == 0:
            win_main["ro_in2"].update(result[1])
            win_main["in_cd2"].update(result[0])
            win_main["in1"].SetFocus(True)
        else:
            continue
    if e == "in1_rtn":
        # 入力された値が8桁数字かチェック（数値ではなく整数文字か判定）
        target = v["in1"]
        if target.isdigit() and len(target) == 8:
            format_date = change_date(target)
            win_main["in1"].update(format_date)
            win_main["btn3"].SetFocus(True)
        else:
            sg.popup("8桁数字で入力してください", font=("Yu Gothic UI", 8))
            continue
    if e == "btn3" or e == "btn3_rtn":
        si = SelectItem(3)
        si.title = "責任店"
        result = si.open_window_sub()
        if not len(result) == 0:
            win_main["ro_in3"].update(result[1])
            win_main["in_cd3"].update(result[0])
            win_main["in2"].SetFocus(True)
            # （入力開始）登録ボタンを有効
            win_main["btn_register"].update(disabled=False)
            # （入力開始）キャンセルボタンを有効
            win_main["btn_cancel"].update(disabled=False)
            # （入力開始）送信ボタンを無効
            win_main["btn_submit"].update(disabled=True)
        else:
            continue
    if e == "in2_rtn":
        target = v["in2"]
        if target.isdigit() and len(target) == 8:
            format_date = change_date(target)
            win_main["in2"].update(format_date)
            win_main["in3"].set_focus(True)
        else:
            sg.popup("8桁数字で入力してください", font=("Yu Gothic UI", 8))
            continue
    if e == "in3_rtn":
        win_main["btn4"].SetFocus(True)
    if e == "btn4" or e == "btn4_rtn":
        si = SelectItem(4)
        si.title = "住所"
        result = si.open_window_sub()
        if not len(result) == 0:
            win_main["ro_in4"].update(result[1])
            win_main["in_cd4"].update(result[0])
            win_main["btn5"].SetFocus(True)
        else:
            continue
    if e == "btn5" or e == "btn5_rtn":
        si = SelectItem(5)
        si.title = "発生状況"
        result = si.open_window_sub()
        if not len(result) == 0:
            win_main["ro_in5"].update(result[1])
            win_main["in_cd5"].update(result[0])
            win_main["btn6"].SetFocus(True)
        else:
            continue
    if e == "btn6" or e == "btn6_rtn":
        si = SelectItem(6)
        si.title = "惹起者"
        result = si.open_window_sub()
        if not len(result) == 0:
            win_main["ro_in6"].update(result[1])
            win_main["in_cd6"].update(result[0])
            win_main["in_cd7"].update(result[2])
            win_main["in4"].SetFocus(True)
        else:
            continue
    if e == "in4_rtn":
        win_main["ml1"].SetFocus(True)
    if e == "btn_register":
        # 未入力がないかチェック
        checklist = check_input()
        result_chk = 0 in checklist.values()
        # 未入力があれば警告して処理中止
        if result_chk:
            sg.popup("入力の抜けている項目があります", font=("Yu Gothic UI", 8), title="確認")
            continue
        else:
            insert_data()
            # （入力終了）登録ボタンを無効
            win_main["btn_register"].update(disabled=True)
            # （入力終了）キャンセルボタンを無効
            win_main["btn_cancel"].update(disabled=True)
            # （入力終了）送信ボタンを有効
            win_main["btn_submit"].update(disabled=False)
            # 添付ファイル登録欄を使用不可にする
            attachment_valid(True)
    if e == "btn_search":
        ir.cel1 = v["in_cd1"]
        ir.cel2 = v["in_cd2"]
        ir.cel3 = v["in1"]
        data = ir.get_input_data()
        if data == "no_data":
            sg.popup("該当する履歴はありません", font=("Yu Gothic UI", 8), title="該当なし")
            win_main["in1"].update("")
        else:
            win_main["hlst"].update(data)
    # 入力履歴テーブルがクリックされたら
    if e == "hlst":
        # 選択された行番号を取得
        if v["hlst"] == []:
            # もし、イベント「hlst」が空なら以下の処理は行わない
            continue
        else:
            # 添付資料登録欄を入力可にする
            attachment_valid()
            # 選択されたレコードの行番号を取得
            row = v["hlst"][0]
            # テーブルのデータを再取得
            ir.cel1 = v["in_cd1"]
            ir.cel2 = v["in_cd2"]
            ir.cel3 = v["in1"]
            data_list = ir.get_input_data()
            # 選択された履歴データを取得
            select_data = data_list[row]
            # 選択された履歴の事故番号を取得
            ir.case_number = select_data[0]
            # T荷物事故テーブルの該当データを取得して、入力フィールドへ再セット
            case_data = ir.set_redo()
            win_main["ro_in3"].update(case_data["責任店名称"])
            win_main["in_cd3"].update(case_data["責任店"])
            win_main["in2"].update(case_data["発生日"])
            win_main["in3"].update(case_data["発生地"])
            win_main["ro_in4"].update(case_data["住所"])
            win_main["in_cd4"].update(case_data["郵便番号"])
            win_main["ro_in5"].update(case_data["発生状況名称"])
            win_main["in_cd5"].update(case_data["発生状況"])
            win_main["ro_in6"].update(case_data["惹起者名称"])
            win_main["in_cd6"].update(case_data["惹起者"])
            win_main["in_cd7"].update(case_data["自他"])
            win_main["in4"].update(case_data["損害額"])
            win_main["ml1"].update(case_data["詳細"])
            win_main["in_cnum"].update(case_data["管理番号"])
            win_main["in_dnum"].update(case_data["明細番号"])
            # 入力履歴テーブル更新用の管理番号を再取得
            control_number = case_data["管理番号"]
            # イベントフォーカスを責任店検索ボタンに移動
            win_main["btn3"].SetFocus(True)
            # 登録ボタンを有効化
            win_main["btn_register"].update(disabled=False)
            # 削除ボタンを有効化
            win_main["btn_delete"].update(disabled=False)
            # キャンセルボタンを有効化
            win_main["btn_cancel"].update(disabled=False)
            # 送信ボタンを無効
            win_main["btn_submit"].update(disabled=True)
    if e == "btn_delete":
        result = sg.popup_ok_cancel("この入力を削除します。\nよろしいですか？", font=("Yu Gothic UI", 9), title="削除")
        if result == "OK":
            # T荷物事故テーブルの該当レコードを削除
            ir.delete_record(v["in_cnum"], v["in_dnum"], get_now_time())
            # 入力フィールドをクリア
            clear_field()
            # 入力履歴テーブルを更新
            data_dct = ir.input_history(control_number)
            win_main["hlst"].update(data_dct)
            # 削除ボタンを無効化
            win_main["btn_delete"].update(disabled=True)
            # メッセージ
            sg.popup("削除しました。", font=("Yu Gothic UI", 9), title="完了")
        else:
            # 入力フィールドをクリア
            clear_field()
            # メッセージ
            sg.popup("取消しました。", font=("Yu Gothic UI", 9), title="取消")
        # 添付ファイル登録欄を使用不可にする
        attachment_valid(True)
        # 送信ボタンを有効化
        win_main["btn_submit"].update(disabled=False)
    if e == "btn_opn":
        if v["intt1"]:
            # もし、添付ファイルが選択されていたら
            # 画像ファイルなど履歴に添付するファイルを選択して、保存する。
            li = LoadImage()
            li.file_path = v["fb1"]
            # 報告店コード、管理番号、明細番号を取得
            li.department_code = v["in_cd1"]
            li.control_number = v["in_cnum"]
            li.detail_number = v["in_dnum"]
            # 資料ファイルプレビュー画面を表示
            li.load_img()
        else:
            # もし、添付ファイルが未選択なら、添付ファイルリストウィンドウを表示
            al = AttachmentList(v["in_cnum"], v["in_dnum"])
            al.open_list_window()
        # 入力フィールドをクリア
        clear_field()
        # 添付ファイル登録欄を入力不可にする
        attachment_valid(True)
        # 送信ボタンを有効化
        win_main["btn_submit"].update(disabled=False)
    if e == "btn_submit":
        # 処理開始（送信ボタン押下）時の日時を取得
        send_now = get_now_time()
        # サーバへ送信するデータを生成してメール送信
        # 1)圧縮するファイルをまとめるための空のディレクトリを作成
        er = ExportRecord(send_now)
        zip_name = er.gen_zipname()
        er.folder = "upload/fcomp"
        er.zip_path = f"upload/{zip_name}"
        os.mkdir(er.folder)
        # 2)T削除番号とT荷物事故のそれぞれ差分レコードをCSV出力して作成して保存
        er.export_table1()
        er.export_table2()
        er.export_table3()
        # 3)T添付資料の送信日時が空白のファイル名称を基にディレクトリへファイルをコピー
        er.copy_file()
        # 4)fcompディレクトリが空かどうかチェック
        if not len(os.listdir(er.folder)) == 0:
            # fcompが空でなければ
            # 5)ファイル入りのディレクトリをzipファイルにする
            er.make_zip()
            # 6)メール送信のためのパラメータ
            sm = SendMail(
                smtp_server="smtp.office365.com",
                port_number=587,
                my_account=sys.argv[2],
                password=sys.argv[3],
                to_address=MAILTO,
                subject="#CASE メール送信テスト",
                attach_zip=f"{er.zip_path}.zip"
            )
            # 7)メール送信
            sm.send_mail()
            # 8)zipファイルを削除
            os.remove(f"{er.zip_path}.zip")
            # 9)T削除番号、T荷物事故、T添付資料の送信日時がNullのレコードに現在日時をセット
            er.set_send_time()
            # 9-2)T送信番号に添付ファイルの情報を書き込み
            er.input_send_history(zip_name)
            # 10)完了メッセージ
            sg.popup("事故データの送信を完了しました", font=("Yu Gothic UI", 8), title="完了")
        else:
            # fcompが空でなければ
            # 11)確認メッセージ
            sg.popup("送信するデータがありません", font=("Yu Gothic UI", 8), title="確認")
        # 圧縮元を削除
        shutil.rmtree(er.folder)
    if e == "btn_cancel":
        clear_field()
        win_main["intt1"].update("")
        # 添付ファイル登録欄を使用不可にする
        attachment_valid(True)
        # 削除ボタンを無効
        win_main["btn_delete"].update(disabled=True)
        # 登録ボタンを無効
        win_main["btn_register"].update(disabled=True)
        # キャンセルボタンを無効
        win_main["btn_cancel"].update(disabled=True)
        # 送信ボタンを有効化
        win_main["btn_submit"].update(disabled=False)
    if e == "txt_closedate_secret":
        answ = sg.popup_ok_cancel("このアプリを初期化します。\nよろしいですか？", font=("Yu Gothic UI", 8))
        if answ == "OK":
            sc = Secret()
            sc.clear_folder()
            sc.clear_table1()
        else:
            continue
    if not e or e == "btn_close":
        break
win_main.close()
