import FreeSimpleGUI as sg
from smartcard.System import readers
from smartcard.util import toHexString
from smartcard.Exceptions import NoCardException
import gspread
from datetime import datetime, timedelta

# 認証
try:
    gc = gspread.service_account(filename=r"replace_with_your_service_account_json_file.json")
    spreadsheet = gc.open("Equipment_Manager")
except Exception as e:
    custom_popup_ok(f"認証に失敗しました。プログラムを終了します。\nAuthentication failed. Exiting program.\n\n{e}")
    exit()

# 関数定義

def read_nfc_id():
    GET_UID_COMMAND = [0xFF, 0xCA, 0x00, 0x00, 0x00]
    """NFCカードを最大10回まで検知を試みる。"""
    for i in range(10):
        sg.popup_no_buttons(f"タッチしてください...\nPlease touch NFC item...", non_blocking=True, auto_close=True, auto_close_duration=1)
        try:
            r = readers()
            if len(r) == 0:
                custom_popup_ok("カードリーダーが見つかりません。\nCard reader not found.")
                return None
            reader = r[0]
            connection = reader.createConnection()
            try:
                connection.connect()
                data, sw1, sw2 = connection.transmit(GET_UID_COMMAND)
                # 処理が成功したかチェック (0x90 0x00 は成功を意味する)
                if sw1 == 0x90 and sw2 == 0x00:
                # バイトデータを16進数文字列に変換
                    idm = toHexString(data)
                    return idm
            except Exception:
                import time
                time.sleep(1)
        except Exception as e:
            custom_popup_ok(f"NFCリーダーエラー: {e}\nNFC reader error: {e}")
            return None
    custom_popup_ok("カードが検知できませんでした。\nCard not detected.")
    return None

def get_employee_name_by_id(idm, employee_ids, employee_list):
    if idm in employee_ids:
        idx = employee_ids.index(idm)
        return employee_list[idx]
    return None

def get_item_name_by_id(idm, item_ids, item_name_list):
    if idm in item_ids:
        idx = item_ids.index(idm)
        return item_name_list[idx]
    return None

def get_all_ids():
    try:
        # 社員マスタからIDと氏名を取得（ヘッダー除く全データ）
        employee_sheet = spreadsheet.worksheet("社員マスタ")
        employee_ids = employee_sheet.col_values(2)[1:]  # 2列目（IDm）
        employee_list = employee_sheet.col_values(1)[1:]  # 1列目（氏名）
        # 物品マスタからIDと物品名を取得
        item_sheet = spreadsheet.worksheet("物品マスタ")
        item_ids = item_sheet.col_values(2)[1:]  # 2列目（IDm）
        item_name_list = item_sheet.col_values(1)[1:]  # 1列目（物品名）
        return employee_ids, item_ids, employee_list, item_name_list
    except Exception as e:
        custom_popup_ok(f"IDリストの取得エラー: {e}\nError retrieving ID lists: {e}")
        return [], [], [], []

def register_employee(idm, name, email):
    """新しい社員をスプレッドシートに登録する"""
    try:
        worksheet = spreadsheet.worksheet("社員マスタ")
        new_row = [name, idm, email]
        worksheet.append_row(new_row)
        custom_popup_ok("社員証の登録が完了しました。\nEmployee card registration completed.")
    except Exception as e:
        custom_popup_ok(f"社員登録エラー: {e}\nError registering employee: {e}")

def register_item(idm, item_name):
    """新しい物品をスプレッドシートに登録する"""
    try:
		#wip tourokutyuuni hyouji
		#sg.popup_no_buttons(f"ただいま登録中です…\nRegistaring...", non_blocking=True, auto_close=True, auto_close_duration=1)
        worksheet = spreadsheet.worksheet("物品マスタ")
        new_row = [item_name, idm]
        worksheet.append_row(new_row)
        custom_popup_ok("物品の登録が完了しました。\nItem registration completed.")
    except Exception as e:
        custom_popup_ok(f"物品登録エラー: {e}\nError registering item: {e}")

def appllication_submit(employee_name, item_name, calendar_date):
    """申請内容をスプレッドシートに保存する"""
    try:
        worksheet = spreadsheet.worksheet("貸出中一覧")
        today = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        new_row = [str(today),str(employee_name), str(item_name), str(calendar_date)]
        worksheet.append_row(new_row)

        worksheet_master = spreadsheet.worksheet("物品マスタ")
        cell = worksheet_master.find(item_name, in_column=1)
        if cell:
            worksheet_master.update_cell(cell.row, 3, employee_name)
            worksheet_master.update_cell(cell.row, 4, today)

        custom_popup_ok(f"登録が終了しました。\n申請者:{employee_name}\n物品:{item_name}\n返却日:{calendar_date}\n\nRegistration completed.\nApplicant: {employee_name}\nItem: {item_name}\nReturn Date: {calendar_date}")
        return_to_main()
    except Exception as e:
        custom_popup_ok(f"申請保存エラー: {e}\nError saving application: {e}")

def application_submit_bug(reporter_name, bug_description):
    try:
        worksheet = spreadsheet.worksheet("不具合報告")
        today = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        new_row = ["未対応", str(today), str(reporter_name), str(bug_description)]
        worksheet.append_row(new_row)
        custom_popup_ok("不具合報告が完了しました。ご協力ありがとうございます。\nBug report submitted. Thank you for your cooperation.")
        return_to_main()
    except Exception as e:
        custom_popup_ok(f"不具合報告エラー: {e}\nError reporting bug: {e}")

def calendar(window):
    window['-VIEW_MAIN-'].update(visible=False)
    window['-VIEW_CALENDAR-'].update(visible=True)
    global current_view
    current_view = 'CALENDAR'
    window["今日まで\nUntil Today"].set_focus()

    selected_date = None
    while True:
        event, values = window.read(timeout=10)
        event = handle_common_events(event) #文字の色変更とキー操作
        if event in (sg.WIN_CLOSED, None):
            break
        if event == "今日まで\nUntil Today":
            selected_date = datetime.now().strftime('%Y-%m-%d')
            break
        elif event == "明日まで\nUntil Tomorrow":
            tomorrow = datetime.now() + timedelta(days=1)
            selected_date = tomorrow.strftime('%Y-%m-%d')
            break
        elif event == "カレンダーを選択\nSelect from Calendar":
            window["カレンダーを選択\nSelect from Calendar"].click()
        elif event == "登録\nRegister":
            if values['-DATE-']:
                selected_date = values['-DATE-']
                break
            else:
                custom_popup_ok("日付を入力してください。\nPlease enter a date.")
                
    return selected_date

def check_employee_borrowed(second_idm, employee_name):
    try:
        worksheet = spreadsheet.worksheet("貸出中一覧")
        all_data = worksheet.get_all_records()
        borrowed_list = []
        for row in all_data:
            if row.get('申請者', '') == employee_name:
                borrowed_list.append(row)
        if borrowed_list:
            msg = f"{employee_name}さんが現在借りている物品一覧:\nYour borrowed items:\n"
            for info in borrowed_list:
                item_name = info.get('物品名', '')
                calendar_date = info.get('返却予定日', '')
                msg += f"\n・物品: {item_name}\n・返却日: {calendar_date}\n"
            result = custom_popup_yes_no(msg + "\n返却の場合は、初めに物品をタッチしてください。\n追加で貸出登録しますか？\nIf you are returning an item, please touch the item first. \nWould you like to register an additional item for loan?")
            if result == "はい\nYes":
                return False
            else:
                return_to_main()
                return True
    except Exception as e:
        custom_popup_ok(f"貸出確認エラー: {e}\nError checking borrowings: {e}")
        return_to_main()
        return True

def check_item_borrowed(item_name):
    try:
        worksheet = spreadsheet.worksheet("貸出中一覧")
        all_data = worksheet.get_all_records()
        borrowed_info = None
        for row in all_data:
            if row['物品名'] == item_name:
                borrowed_info = row
                break
        if borrowed_info:
            employee_name = borrowed_info.get('申請者', '')
            calendar_date = borrowed_info.get('返却予定日', '')
            borrower_name = borrowed_info.get('申請者', '')
            scheduled_date = borrowed_info.get('返却予定日', '')
            msg = (f"{item_name}は既に貸出中です。返却しますか？\n{item_name} is currently borrowed. Return it?\n\n"
                    f"[持ち出し情報]\n申請者:{employee_name}\n物品:{item_name}\n返却日:{calendar_date}")
            result = custom_popup_yes_no(msg)
            if result == "はい\nYes":
                return_item(item_name, borrower_name, scheduled_date)
                custom_popup_ok(f"{item_name}の返却が完了しました。借りる場合はもう一度貸出登録をしてください。\n\n{item_name} return completed. Please register again if you want to borrow it.")
                return_to_main()
                return True  # 返却した
            elif result == "いいえ\nNo":
                return True  # 返却しなかった
        return_to_main()
        return False  # 返却しなかった
    except Exception as e:
        custom_popup_ok(f"貸出確認エラー: {e}\nError checking borrowings: {e}")
        return_to_main()
        return True  # エラー時もスキップ

def return_item(item_name, borrower_name, scheduled_date):
    """スキャンしたアイテムを貸出中一覧から削除する"""
    try:
        add_return_record(item_name, borrower_name, scheduled_date)  # 返却履歴に追加
        worksheet = spreadsheet.worksheet("貸出中一覧")
        cell = worksheet.find(item_name, in_column=3)
        if cell:
            worksheet.delete_rows(cell.row)
    except Exception as e:
        custom_popup_ok(f"返却処理エラー: {e}\nError during return processing: {e}")


def return_to_main():
    """メインメニューに戻る"""
    global current_view
    window[f'-VIEW_{current_view}-'].update(visible=False)
    window['-VIEW_MAIN-'].update(visible=True)
    current_view = 'MAIN'

def add_return_record(item_name, borrower_name, scheduled_date):
    try:
        worksheet = spreadsheet.worksheet("返却履歴")
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        new_row = [timestamp, item_name, borrower_name, scheduled_date]
        worksheet.append_row(new_row)
    except Exception as e:
        custom_popup_ok(f"返却履歴の記録エラー: {e}\nError recording return history: {e}")

def get_borrowed_list_data():
    try:
        worksheet = spreadsheet.worksheet("貸出中一覧")
        all_data = worksheet.get_all_values()
        if len(all_data) <= 1:
            return [["現在、貸出中の物品はありません", "", "", ""]]
        
        # 1行目を除く
        data = all_data[1:]
        data.reverse()
        return data
    except Exception as e:
        print(f"データ取得エラー: {e}")
        return [["エラー", "データを取得できませんでした", "", ""]]

def get_returned_list_data():
    try:
        worksheet = spreadsheet.worksheet("返却履歴")
        all_data = worksheet.get_all_values()
        if len(all_data) <= 1:
            return [["現在、返却済みの物品はありません", "", "", ""]]
        
        # 1行目を除く
        data = all_data[1:] 
        data.reverse()
        return data
    except Exception as e:
        print(f"データ取得エラー: {e}")
        return [["エラー", "データを取得できませんでした", "", ""]]

def get_employee_list_data():
    try:
        worksheet = spreadsheet.worksheet("社員マスタ")
        all_data = worksheet.get_all_values()
        if len(all_data) <= 1:
            return [["現在、登録されている社員はいません", "", ""]]
        return [[row[0], row[2]] for row in all_data[1:]]
    except Exception as e:
        print(f"データ取得エラー: {e}")
        return [["エラー", "データを取得できませんでした", ""]]

def get_item_list_data():
    try:
        worksheet = spreadsheet.worksheet("物品マスタ")
        all_data = worksheet.get_all_values()
        if len(all_data) <= 1:
            return [["現在、登録されている物品はありません", ""]]
        return [[row[0], row[2], row[3]] for row in all_data[1:]]
    except Exception as e:
        print(f"データ取得エラー: {e}")
        return [["エラー", "データを取得できませんでした"]]

def bug_list_data():
    try:
        worksheet = spreadsheet.worksheet("不具合報告")
        all_data = worksheet.get_all_values()
        if len(all_data) <= 1:
            return [["現在、報告されている不具合はありません", "", "", ""]]
        data = all_data[1:]
        data.reverse()
        return data
    except Exception as e:
        print(f"データ取得エラー: {e}")
        return [["エラー", "データを取得できませんでした", "", ""]]

last_view = None
def handle_common_events(event): #Enter,Focus時に色を変える
    global current_view, last_view
    if current_view != last_view:
        default_keys = FOCUS_MAP.get(current_view, [])
        if default_keys:
            # 最初の要素にフォーカスを当てる
            window[default_keys[0]].set_focus()
        last_view = current_view

    if event == "-ENTER-":
        focused_element = window.find_element_with_focus()
        if focused_element:
            # keyがない場合は「はい\nYES」などのテキストを取得し，特定する
            return focused_element.key if focused_element.key else focused_element.get_text()
    
    if event in ("Up", "Down", "Left", "Right") or any(event.startswith(k) for k in ("Up", "Down", "Left", "Right")):
        current_keys = FOCUS_MAP.get(current_view, [])
        if current_keys:
            focused_element = window.find_element_with_focus()
            # 同じく判定には key を優先、なければ text を使う
            current_key = focused_element.key if (focused_element and focused_element.key) else (focused_element.get_text() if focused_element else None)
            
            try:
                idx = current_keys.index(current_key) if current_key in current_keys else 0
                if event in ("Down", "Right") or "116" in event or "114" in event:
                    next_idx = (idx + 1) % len(current_keys)
                else:
                    next_idx = (idx - 1) % len(current_keys)
                window[current_keys[next_idx]].set_focus()
            except Exception as e:
                print(f"Focus move error: {e}")

    focused_element = window.find_element_with_focus()
    # 表示されている画面のキーだけをループ対象にする（エラー防止）
    current_keys = FOCUS_MAP.get(current_view, [])
    for k in current_keys:
        if k in window.AllKeysDict:
            el = window[k]
            # Table要素は色の更新対象外にする
            if isinstance(el, sg.Table): continue
            
            if focused_element == el:
                if isinstance(el, sg.Button):
                    el.update(button_color=('black', 'lightgreen'))
                else:
                    el.update(background_color='lightgreen')
            else:
                if isinstance(el, sg.Button):
                    el.update(button_color=sg.theme_button_color())
                else:
                    el.update(background_color='white')
    
    return event


def custom_popup_yes_no(message):
    global current_view
    # 以前のビューを隠してポップアップ用ビューを表示
    old_view = current_view
    window[f'-VIEW_{old_view}-'].update(visible=False)
    window['-POPUP_YESNO-'].update(visible=True)
    window['-POPUP_MSG-'].update(message)
    current_view = 'POPUP_YESNO'
    window["はい\nYes"].set_focus() # 初期フォーカス

    result = "いいえ\nNo" 
    while True:
        event, values = window.read(timeout=10)
        if event == sg.WIN_CLOSED:
            exit()
        event = handle_common_events(event) #文字の色変更とキー操作

        if event == "はい\nYes":
            result = "はい\nYes"
            break
        if event == "いいえ\nNo":
            result = "いいえ\nNo"
            break
    window["-POPUP_YESNO-"].update(visible=False)
    window[f'-VIEW_{old_view}-'].update(visible=True)
    current_view = old_view
    return result

def custom_popup_ok(message):
    global current_view
    old_view = current_view
    window[f'-VIEW_{old_view}-'].update(visible=False)
    window['-POPUP_OK-'].update(visible=True)
    window['OK-MSG-'].update(message)
    current_view = 'POPUP_OK'
    window["了解\nOK"].set_focus()

    while True:
        event, values = window.read(timeout=10)
        if event == sg.WIN_CLOSED:
            exit()
        event = handle_common_events(event) #文字の色変更とキー操作

        if event == "了解\nOK":
            break
    window["-POPUP_OK-"].update(visible=False)
    window[f'-VIEW_{old_view}-'].update(visible=True)
    current_view = old_view

# GUIレイアウト定義
FOCUS_MAP = {
    'MAIN': [
        "貸出 / 返却 / 登録\nBorrow / Return / Register", 
        "貸出状況一覧を見る\nView Current Borrowed Items", 
        "登録されている物品一覧を見る\nView Registered Items",
        "返却履歴一覧を見る\nView Returned Items History",
        "登録されている社員一覧を見る\nView Registered Employees",
        "不具合報告\nBug Reports"
    ],
    'REG_SELECT': [  # VIEW_ を取って現在の状態名に合わせる
        "社員証として登録\nRegister as employee card",
        "物品として登録\nRegister as Item"
    ],
    'REG_EMP': ['-EMP_REG_NAME-', '-EMP_REG_EMAIL-', "この内容で登録\nRegister with this information"], # 入力欄も追加
    'REG_ITEM': ['-ITEM_REG_NAME-', "この内容で登録\nRegister with this information"], # 入力欄も追加
    'CALENDAR': ["今日まで\nUntil Today", "明日まで\nUntil Tomorrow", "-DATE-", "カレンダーを選択\nSelect from Calendar", "登録\nRegister"],
    'BORROW_LIST': ["-BACK_BORROW-"],
    'RETURN_LIST': ["-BACK_RETURN-"],
    'EMPLOYEE_LIST': ["-BACK_EMPLOYEE-"],
    'ITEM_LIST': ["-BACK_ITEM-"],
    'BUG_LIST': ["-BACK_BUG-", '-BUG_REPORTER-', '-BUG_DESCRIPTION-', "-SUBMIT_BUG-"],
    'POPUP_YESNO': ["はい\nYes", "いいえ\nNo"],
    'POPUP_OK': ["了解\nOK"],
}

# 各画面のレイアウトをsg.Columnで定義
layout_main = [
    [sg.Btn("貸出 / 返却 / 登録\nBorrow / Return / Register", size=(25, 3))],
    [sg.Btn("貸出状況一覧を見る\nView Current Borrowed Items", size=(25, 3))],
    [sg.Btn("登録されている物品一覧を見る\nView Registered Items", size=(25, 3))],
    [sg.Btn("返却履歴一覧を見る\nView Returned Items History", size=(25, 3))],
    [sg.Btn("登録されている社員一覧を見る\nView Registered Employees", size=(25, 3))],
    [sg.Btn("不具合報告\nBug Reports", size=(25, 3))],]

layout_register_select = [
    [sg.Txt("未登録のIDです。どちらを登録しますか？\nThis ID is not registered. Which type do you want to register?")],
    [sg.Btn("社員証として登録\nRegister as employee card", size=(25, 3)), sg.Btn("物品として登録\nRegister as Item",size=(15, 3))],
]

layout_register_employee = [
    [sg.Txt("新しい社員証を登録します。\nRegister a new employee card.")],
    [sg.Txt("IDm:", size=(8,1)), sg.Txt("", key='-EMP_REG_IDM-')],
    [sg.Txt("Your name:", size=(8,1)), sg.In(key='-EMP_REG_NAME-')],
    [sg.Txt("eMail(ugo):", size=(8,1)), sg.In(key='-EMP_REG_EMAIL-')],
    [sg.Btn("この内容で登録\nRegister with this information")],
]

layout_register_item = [
    [sg.Txt("新しい物品を登録します。\nRegister a new item.")],
    [sg.Txt("IDm:", size=(8,1)), sg.Txt("", key='-ITEM_REG_IDM-')],
    [sg.Txt("Item name:", size=(8,1)), sg.In(key='-ITEM_REG_NAME-')],
    [sg.Btn("この内容で登録\nRegister with this information")],
]

layout_calendar = [
    [sg.Btn("今日まで\nUntil Today", size=(10, 2)), sg.Btn("明日まで\nUntil Tomorrow", size=(10, 2))],
    [sg.Txt("それ以外の場合はカレンダーから選択してください。\nFor other dates, please select from the calendar.")],
    [
        sg.In(key='-DATE-', size=(15, 1)),
        sg.CalendarButton("カレンダーを選択\nSelect from Calendar", target='-DATE-', format='%Y-%m-%d', size=(20,2))
    ],
    [sg.Button("登録\nRegister", size=(10, 2))],
]

header = ["申請日時(Application Date)", "申請者(Applicant)", "物品名(Item Name)", "返却予定日(Scheduled Return Date)"]
layout_borrow_list = [
    [sg.Txt("現在の貸出状況一覧 (Current Borrowed Items)")],
    [sg.Table(values=get_borrowed_list_data(), 
              headings=header, 
              auto_size_columns=False,
              col_widths=[5, 5, 10, 5],
              justification='left',
              key='-BORROW_TABLE-',
              row_height=30,
              num_rows=10)], # 表示する行数
    [sg.Btn("メインに戻る", key='-BACK_BORROW-')]
]

header = ["返却日時(Return Date)", "物品名(Item Name)", "返却者(Returner)", "予定返却日(Scheduled Return Date)"]
layout_return_list = [
    [sg.Txt("返却履歴一覧 (Returned Items History)")],
    [sg.Table(values=get_returned_list_data(), 
              headings=header, 
              auto_size_columns=False,
              col_widths=[10, 10, 5, 5],
              justification='left',
              key='-RETURN_TABLE-',
              row_height=30,
              num_rows=10)], # 表示する行数
    [sg.Btn("メインに戻る", key='-BACK_RETURN-')]
]

header = ["氏名(Name)", "eMail(ugo)"]
layout_employee_list = [
    [sg.Txt("登録されている社員一覧 (Registered Employees)")],
    [sg.Table(values=get_employee_list_data(), 
              headings=header,
              auto_size_columns=False,
              col_widths=[10, 20],
              justification='left',
              key='-EMPLOYEE_TABLE-',
              row_height=30,
              num_rows=10)], # 表示する行数
    [sg.Btn("メインに戻る", key='-BACK_EMPLOYEE-')]
]

header = ["物品名(Item Name)", "現在の貸出者(Current Borrower)", "最終貸出日時(Last Borrowed Date)"]
layout_item_list = [
    [sg.Txt("登録されている物品一覧 (Registered Items)")],
    [sg.Table(values=get_item_list_data(), 
              headings=header,
              auto_size_columns=False,
              col_widths=[15, 10, 10],
              justification='left',
              key='-ITEM_TABLE-',
              row_height=30,
              num_rows=10)], # 表示する行数
    [sg.Btn("メインに戻る", key='-BACK_ITEM-')]
]

header = [ "対応状況(Status)", "報告日時(Report Date)", "報告者(Reporter)", "不具合内容(Bug Description)"]
layout_bug_list = [
    [sg.Txt("不具合報告一覧 (Bug Reports)")],
    [sg.Table(values=bug_list_data(), 
              headings=header,
              auto_size_columns=False,
              col_widths=[5, 10, 10, 30],
              justification='left',
              key='-BUG_TABLE-',
              row_height=30,
              num_rows=10)], # 表示する行数
    [sg.Btn("メインに戻る\nReturn to Main", size=(15,2), key='-BACK_BUG-')],
    [sg.Txt("名前(Your Name):"), sg.In(key='-BUG_REPORTER-', size=(20,1))],
    [sg.Txt("不具合内容(Bug Description):")],
    [sg.Multiline(key='-BUG_DESCRIPTION-', size=(60,5))],
    [sg.Btn("不具合報告を送信\nSubmit Bug Report", size=(20,2), key='-SUBMIT_BUG-')]
]

layout = [[sg.Txt("", key='-POPUP_MSG-', font=('Helvetica', 12))],
    [sg.Btn("はい\nYes", size=(10, 2)), sg.Btn("いいえ\nNo", size=(10, 2))],
]

layout_popup_ok = [
    [sg.Txt("", key='OK-MSG-')],
    [sg.Btn("了解\nOK", size=(10, 2))]
]


# 全てのColumnを一つのレイアウトにまとめる
layout = [
    [sg.Column(layout_main, key='-VIEW_MAIN-'),
     sg.Column(layout_register_select, visible=False, key='-VIEW_REG_SELECT-'),
     sg.Column(layout_register_employee, visible=False, key='-VIEW_REG_EMP-'),
     sg.Column(layout_register_item, visible=False, key='-VIEW_REG_ITEM-'),
	 sg.Column(layout_calendar, visible=False, key='-VIEW_CALENDAR-'),
     sg.Column(layout_borrow_list, visible=False, key='-VIEW_BORROW_LIST-'),
     sg.Column(layout_return_list, visible=False, key='-VIEW_RETURN_LIST-'),
     sg.Column(layout_employee_list, visible=False, key='-VIEW_EMPLOYEE_LIST-'),
     sg.Column(layout_item_list, visible=False, key='-VIEW_ITEM_LIST-'),
     sg.Column(layout_bug_list, visible=False, key='-VIEW_BUG_LIST-'),
     sg.Column(layout, visible=False, key='-POPUP_YESNO-'),
     sg.Column(layout_popup_ok, visible=False, key='-POPUP_OK-')]
]

# ウィンドウ作成とイベントループ
window = sg.Window("Equipment Manager", layout, finalize=True, return_keyboard_events=True)
current_view = 'MAIN'
unregistered_idm = None

window.bind("<Return>", "-ENTER-")
window.bind("<Escape>", "-ESCAPE-")


while True: 
    event, values = window.read(timeout=10)
    if event == sg.WIN_CLOSED:
        break
    
    # ESCキーでメインメニューに戻る
    if event in "-ESCAPE-":
        if current_view != 'MAIN':
            return_to_main()
        continue
    
    event = handle_common_events(event) #文字の色変更とキー操作

    # メイン画面の処理
    if current_view == 'MAIN':
        if event == "貸出 / 返却 / 登録\nBorrow / Return / Register":
            idm = read_nfc_id()
            if idm:
                employee_ids, item_ids, E_name, I_names = get_all_ids()
                employee_name = get_employee_name_by_id(idm, employee_ids, E_name)
                item_name = get_item_name_by_id(idm, item_ids, I_names)
                if idm in employee_ids:
                    if check_employee_borrowed(idm, employee_name):
                        continue  # 返却した場合は以降の処理をスキップしてメインメニューへ
                    custom_popup_ok(f"社員証を確認しました: { employee_name }\n借りる物品をタッチしてください。\nI checked your employee ID:{ employee_name }\nPlease touch the item you want to borrow.")
                    second_idm = read_nfc_id()  # 物品のIDを読み取る
                    employee_name = get_employee_name_by_id(idm, employee_ids, E_name)
                    item_name = get_item_name_by_id(second_idm, item_ids, I_names)
                    if second_idm in item_ids:
                        if check_item_borrowed(item_name):
                            continue  # 返却した場合は以降の処理をスキップしてメインメニューへ
                        custom_popup_ok(f"物品を確認しました: { item_name }\n返却日を登録してください\nI checked the item: { item_name }\nPlease register the return date.")
                        calendar_date = calendar(window)
                        if calendar_date:
                            appllication_submit(employee_name, item_name, calendar_date)
                            return_to_main()
                    elif second_idm in employee_ids:
                        custom_popup_ok(f"Error:社員証をタッチしています。物品のタッチを行ってください。\nError: You are touching the employee ID. Please touch the item.")
                    else:
                        unregistered_idm = idm
                        window['-VIEW_MAIN-'].update(visible=False)
                        window['-VIEW_REG_SELECT-'].update(visible=True)
                        current_view = 'REG_SELECT'
                    

                elif idm in item_ids:
                    if check_item_borrowed(item_name):
                        continue  # 返却した場合は以降の処理をスキップしてメインメニューへ
                    custom_popup_ok(f"物品を確認しました: { item_name }\n社員証をタッチしてください。\nI checked the item: { item_name }\nPlease touch your employee ID.")
                    second_idm = read_nfc_id()
                    employee_name = get_employee_name_by_id(second_idm, employee_ids, E_name)
                    item_name = get_item_name_by_id(idm, item_ids, I_names)
                    if second_idm in employee_ids:
                        employee_name = get_employee_name_by_id(second_idm, employee_ids, E_name)
                        item_name = get_item_name_by_id(idm, item_ids, I_names)
                        custom_popup_ok(f"社員証を確認しました: { employee_name }\n返却日を登録してください\nI checked your employee ID: { employee_name }\nPlease register the return date.")
                        calendar_date = calendar(window)
                        if calendar_date:
                            appllication_submit(employee_name, item_name, calendar_date)
                            return_to_main()
                    elif second_idm in item_ids:
                        custom_popup_ok(f"Error:物品をタッチしています。社員証のタッチを行ってください。\nError: You are touching the item. Please touch your employee ID.")
                    else:
                        unregistered_idm = idm
                        window['-VIEW_MAIN-'].update(visible=False)
                        window['-VIEW_REG_SELECT-'].update(visible=True)
                        current_view = 'REG_SELECT'
                else:
                    unregistered_idm = idm
                    window['-VIEW_MAIN-'].update(visible=False)
                    window['-VIEW_REG_SELECT-'].update(visible=True)
                    current_view = 'REG_SELECT'
        if event == "貸出状況一覧を見る\nView Current Borrowed Items":
            window['-VIEW_MAIN-'].update(visible=False)
            window['-VIEW_BORROW_LIST-'].update(visible=True)
            window['-BORROW_TABLE-'].update(values=get_borrowed_list_data())
            current_view = 'BORROW_LIST'
        if event == "返却履歴一覧を見る\nView Returned Items History":
            window['-VIEW_MAIN-'].update(visible=False)
            window['-VIEW_RETURN_LIST-'].update(visible=True)
            window['-RETURN_TABLE-'].update(values=get_returned_list_data())
            current_view = 'RETURN_LIST'
        if event == "登録されている社員一覧を見る\nView Registered Employees":
            window['-VIEW_MAIN-'].update(visible=False)
            window['-VIEW_EMPLOYEE_LIST-'].update(visible=True)
            current_view = 'EMPLOYEE_LIST'
        if event == "登録されている物品一覧を見る\nView Registered Items":
            window['-VIEW_MAIN-'].update(visible=False)
            window['-VIEW_ITEM_LIST-'].update(visible=True)
            current_view = 'ITEM_LIST'
        if event == "不具合報告\nBug Reports":
            window['-VIEW_MAIN-'].update(visible=False)
            window['-VIEW_BUG_LIST-'].update(visible=True)
            window['-BUG_TABLE-'].update(values=bug_list_data())
            current_view = 'BUG_LIST'
    # 貸出状況一覧画面の処理
    elif current_view == 'BORROW_LIST':
        if event == "-BACK_BORROW-":
            return_to_main()
    # 返却履歴一覧画面の処理
    elif current_view == 'RETURN_LIST':
        if event == "-BACK_RETURN-":
            return_to_main()
    # 社員一覧画面の処理
    elif current_view == 'EMPLOYEE_LIST':
        if event == "-BACK_EMPLOYEE-":
            return_to_main()
    # 物品一覧画面の処理
    elif current_view == 'ITEM_LIST':
        if event == "-BACK_ITEM-":
            return_to_main()
    # 不具合報告一覧画面の処理
    elif current_view == 'BUG_LIST':
        if event == "-BACK_BUG-":
            return_to_main()
        elif event == "-SUBMIT_BUG-":
            reporter_name = values['-BUG_REPORTER-']
            bug_description = values['-BUG_DESCRIPTION-']
            if reporter_name and bug_description:
                application_submit_bug(reporter_name, bug_description)
                # フォームをクリア
                window['-BUG_REPORTER-'].update('')
                window['-BUG_DESCRIPTION-'].update('')
            else:
                custom_popup_ok("名前と不具合内容を入力してください。\nPlease enter your name and bug description.")
    

    # 登録種別選択画面の処理
    elif current_view == 'REG_SELECT':
        if event == "社員証として登録\nRegister as employee card":
            window['-VIEW_REG_SELECT-'].update(visible=False)
            window['-VIEW_REG_EMP-'].update(visible=True)
            window['-EMP_REG_IDM-'].update(unregistered_idm) # IDmを表示
            current_view = 'REG_EMP'
        elif event == "物品として登録\nRegister as Item":
            window['-VIEW_REG_SELECT-'].update(visible=False)
            window['-VIEW_REG_ITEM-'].update(visible=True)
            window['-ITEM_REG_IDM-'].update(unregistered_idm) # IDmを表示
            current_view = 'REG_ITEM'

    # 社員証登録画面の処理
    elif current_view == 'REG_EMP':
        if "この内容で登録\nRegister" in event:
            name = values['-EMP_REG_NAME-']
            email = values['-EMP_REG_EMAIL-']
            if name and email:
                register_employee(unregistered_idm, name, email)
                # メインメニューに戻る
                window['-VIEW_REG_EMP-'].update(visible=False)
                window['-VIEW_MAIN-'].update(visible=True)
                current_view = 'MAIN'
            else:
                custom_popup_ok("氏名とメールアドレスを入力してください。\nPlease enter your name and email address.")

    # 物品登録画面の処理
    elif current_view == 'REG_ITEM':
        if "この内容で登録\nRegister" in event:
            item_name = values['-ITEM_REG_NAME-']
            if item_name:
                register_item(unregistered_idm, item_name)
                return_to_main() 
            else:
                custom_popup_ok("物品名を入力してください。\nPlease enter the item name.")

window.close()