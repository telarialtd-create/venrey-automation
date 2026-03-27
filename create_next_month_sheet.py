#!/usr/bin/env python3
"""
来月のシートを自動作成するスクリプト

使い方:
  python3 create_next_month_sheet.py

実行すると現在の月のシートを複製して来月のシートを作成します。
（例: 4月に実行すると「2026年5月」シートが作成される）
"""

import json
from datetime import date, datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

SPREADSHEET_ID = '10siqLe6B9A7uvNWgRUdHb462RqxCxkGEGMEKTPhY-S8'
CREDS_PATH     = '/Users/hiraokawashin/.config/gdrive-server-credentials.json'
OAUTH_PATH     = '/Users/hiraokawashin/.config/gcp-oauth.keys.json'

def get_service():
    creds_data = json.load(open(CREDS_PATH))
    oauth_data = json.load(open(OAUTH_PATH))['installed']
    creds = Credentials(
        token         = creds_data['access_token'],
        refresh_token = creds_data['refresh_token'],
        token_uri     = oauth_data['token_uri'],
        client_id     = oauth_data['client_id'],
        client_secret = oauth_data['client_secret'],
        scopes        = creds_data['scope'].split() if isinstance(creds_data['scope'], str) else creds_data['scope']
    )
    # トークンが期限切れなら自動更新
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        data = {
            'access_token':  creds.token,
            'refresh_token': creds.refresh_token,
            'scope':         ' '.join(creds.scopes),
            'token_type':    'Bearer',
            'expiry_date':   9999999999999
        }
        json.dump(data, open(CREDS_PATH, 'w'))
    return build('sheets', 'v4', credentials=creds)


def main():
    service = get_service()

    # 今月・来月を計算
    now        = datetime.now()
    this_year  = now.year
    this_month = now.month
    if this_month == 12:
        next_year, next_month = this_year + 1, 1
    else:
        next_year, next_month = this_year, this_month + 1

    src_name = f'{this_year}年{this_month}月'
    dst_name = f'{next_year}年{next_month}月'
    days_in_next_month = (date(next_year, next_month % 12 + 1, 1) - date(next_year, next_month, 1)).days \
        if next_month != 12 else 31

    print(f'コピー元: {src_name}')
    print(f'作成先:   {dst_name}（{days_in_next_month}日）')

    # シートID一覧を取得
    meta   = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = {s['properties']['title']: s['properties']['sheetId'] for s in meta['sheets']}

    if src_name not in sheets:
        print(f'エラー: 「{src_name}」シートが見つかりません')
        return

    if dst_name in sheets:
        print(f'「{dst_name}」は既に存在します。スキップします。')
        return

    src_id = sheets[src_name]

    # ── Step1: 今月シートを完全複製（書式・関数・色を保持）──
    print('シートを複製中...')
    res = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': [{'duplicateSheet': {
            'sourceSheetId':  src_id,
            'insertSheetIndex': 1,
            'newSheetName':   dst_name
        }}]}
    ).execute()
    dst_id = res['replies'][0]['duplicateSheet']['properties']['sheetId']
    print(f'複製完了 (sheetId={dst_id})')

    # ── Step2: 今月シートのデータを取得（来月分のシフトデータを探す）──
    print('今月シートのデータを読み込み中...')
    src_data = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{src_name}'!A1:BV200"
    ).execute().get('values', [])

    # 1行目から「来月データ開始列」を特定（2個目の '1' の位置）
    row1 = src_data[0] if src_data else []
    found_first = False
    next_month_col = None
    for i, v in enumerate(row1):
        if str(v).strip() == '1':
            if not found_first:
                found_first = True
            else:
                next_month_col = i
                break

    # ── Step3: 来月シートのシフトデータをクリア（書式は残す）──
    print('シフトデータをクリア中...')
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{dst_name}'!E3:BV200"
    ).execute()

    # ── Step4: 月名・日付ヘッダーを来月用に更新 ──
    print('ヘッダーを更新中...')
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{dst_name}'!A1",
        valueInputOption='RAW',
        body={'values': [[f'{next_month}月']]}
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{dst_name}'!E1",
        valueInputOption='RAW',
        body={'values': [[str(d) for d in range(1, days_in_next_month + 1)]]}
    ).execute()
    # 今月の最終日列以降をクリア（例: 来月が30日なら31日列を消す）
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{dst_name}'!AI1:BV1"
    ).execute()

    # ── Step5: セルの色を整理 ──
    #   来月シートのE列以降の色をリセット → 今月シートの「来月列」の色をコピー
    print('セルの色を整理中...')
    requests = [
        # E列以降の背景色をすべて白にリセット
        {
            'repeatCell': {
                'range': {
                    'sheetId': dst_id,
                    'startRowIndex': 0,
                    'endRowIndex': 200,
                    'startColumnIndex': 4,   # E列
                    'endColumnIndex': 80
                },
                'cell': {'userEnteredFormat': {
                    'backgroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0}
                }},
                'fields': 'userEnteredFormat.backgroundColor'
            }
        }
    ]
    # 来月データがあれば色もコピー
    if next_month_col is not None:
        next_month_days_in_src = len([v for v in row1[next_month_col:] if str(v).strip().isdigit()])
        if next_month_days_in_src > 0:
            requests.append({
                'copyPaste': {
                    'source': {
                        'sheetId': src_id,
                        'startRowIndex': 0,
                        'endRowIndex': 200,
                        'startColumnIndex': next_month_col,
                        'endColumnIndex': next_month_col + next_month_days_in_src
                    },
                    'destination': {
                        'sheetId': dst_id,
                        'startRowIndex': 0,
                        'endRowIndex': 200,
                        'startColumnIndex': 4,
                        'endColumnIndex': 4 + next_month_days_in_src
                    },
                    'pasteType': 'PASTE_FORMAT'
                }
            })
            print(f'  今月シートの来月カラム（{next_month_days_in_src}日分）の色をコピー')

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': requests}
    ).execute()

    # ── Step6: 今月シートに入力済みの来月シフトデータを移行 ──
    if next_month_col is not None:
        print('来月シフトデータを移行中...')
        shift_rows = []
        for row in src_data[2:]:
            new_row = []
            for d in range(days_in_next_month):
                idx = next_month_col + d
                new_row.append(row[idx] if idx < len(row) else '')
            shift_rows.append(new_row)
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{dst_name}'!E3",
            valueInputOption='RAW',
            body={'values': shift_rows}
        ).execute()
        filled = sum(1 for r in shift_rows if any(v for v in r))
        print(f'  {filled} 人分のシフトデータを移行しました')
    else:
        print('今月シートに来月データはありませんでした（シフトデータは空のままです）')

    print(f'\n✅ 完了！ 「{dst_name}」シートを作成しました')


if __name__ == '__main__':
    main()
