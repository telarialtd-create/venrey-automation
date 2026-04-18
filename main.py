#!/usr/bin/env python3
"""
Mr.Venrey 週間スケジュール自動更新スクリプト

スプレッドシートから今週の出勤データを読み込み、
Venrey管理画面の週間スケジュールを自動更新する。

使い方:
  python3 main.py
"""

import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, date, timedelta

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ============================================================
# 設定
# ============================================================
LOGIN_URL = "https://mrvenrey.jp/"

# 店舗ごとのログイン情報
# スプレッドシートの「ふわもこSPA」行より上 → STORES[0]、以降 → STORES[1]
STORES = [
    {"id": "GRP001121", "password": "hj6bf3fwck"},  # 店舗1
    {"id": "rd67",       "password": "52a4et7"},      # ふわもこSPA
]

# スプレッドシートで店舗を区切るセル値（この行以降が次の店舗）
STORE_SEPARATOR = "ふわもこSPA"

# ふわもこSPA のスタッフ読み込みをこの名前の行の手前で停止する
STORE2_STOP_AT = "たくま"

# スプレッドシート ID（URL の /d/ と /edit の間の文字列）
SPREADSHEET_ID = "10siqLe6B9A7uvNWgRUdHb462RqxCxkGEGMEKTPhY-S8"

# 今月のシート名（例: "2026年4月"）を自動生成して URL を組み立てる
# シート名でアクセスするため、毎月 GID を変更する必要なし
_now = datetime.now()
_current_month_sheet = f"{_now.year}年{_now.month}月"
SHEET_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    f"/gviz/tq?tqx=out:csv&sheet={_current_month_sheet}"
)

# False: ブラウザを表示して動作確認（最初はこちら推奨）
# True : ブラウザ非表示で高速実行（動作確認後に変更）
HEADLESS = True
# ============================================================


def parse_time_cell(cell_value):
    """
    スプレッドシートのセル値を (開始時間, 終了時間) にパースする。

    対応フォーマット:
      "11-15上"      → ("11:00", "15:00")
      "1130-1930上"  → ("11:30", "19:30")
      "12-2030上"    → ("12:00", "20:30")
      "19-25上"      → ("19:00", "01:00")  # 25 - 24 = 翌1時
      "14-24*130"    → ("14:00", "01:30")  # *130 → 0130 → 01:30
      "*1800"        → None                # 開始時間なしはスキップ

    Returns:
        (start_str, end_str) のタプル、または None（出勤なし・解析不能）
    """
    if cell_value is None:
        return None
    s = str(cell_value).strip()
    if not s or s == "nan":
        return None
    # 数字を含まないセル（例: "体調不良", "確認中", "作成"）はスキップ
    if not re.search(r"\d", s):
        return None

    def raw_to_hhmm(raw):
        """数字文字列を (hour, minute) に変換"""
        raw = raw.strip().lstrip("*")
        if len(raw) <= 2:
            return int(raw), 0
        # 末尾 2 桁が分、それ以前が時
        return int(raw[:-2]), int(raw[-2:])

    # パターン: START-END上 または START-END*OVERTIME（末尾に送迎などの注記も許容）
    m = re.match(r"^(\d{2,4})-(\d{2,4})[上]?(?:\*(\d{1,4}))?", s)
    if m:
        sh, sm = raw_to_hhmm(m.group(1))

        if m.group(3):
            # *NNN 形式の終了時間（例: *130 → 25:30, *1800 → 18:00）
            overtime_raw = m.group(3).zfill(4)
            eh, em = raw_to_hhmm(overtime_raw)
            # 深夜帯（6時未満）は Venrey の 25時表記に変換（01:30 → 25:30）
            if eh < 6:
                eh += 24
        else:
            # END 数字がそのまま使われる場合（例: 25上 → 25:00）
            eh, em = raw_to_hhmm(m.group(2))

        # 開始時間は基本的に24時未満
        if sh >= 24:
            sh -= 24

        return f"{sh:02d}:{sm:02d}", f"{eh:02d}:{em:02d}"

    # ハイフンなし = 休み（数字を含むがハイフンのないセルはすべて休みとして扱う）
    # 例: "1215", "123020", "1216上", "ロビー確認1120上" など
    return "休み"


def _get_access_token():
    """GDrive OAuth認証情報からアクセストークンを取得（リフレッシュ）する。
    環境変数（GitHub Actions）があればそちらを優先し、なければローカルファイルを使用。
    """
    client_id     = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN")

    if not all([client_id, client_secret, refresh_token]):
        oauth_path = os.path.expanduser("~/.config/gcp-oauth.keys.json")
        creds_path = os.path.expanduser("~/.config/gdrive-server-credentials.json")
        with open(oauth_path) as f:
            oauth = json.load(f)
        with open(creds_path) as f:
            creds = json.load(f)
        client_id     = oauth["installed"]["client_id"]
        client_secret = oauth["installed"]["client_secret"]
        refresh_token = creds["refresh_token"]

    data = urllib.parse.urlencode({
        "client_id":     client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type":    "refresh_token",
    }).encode()
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token", data=data, method="POST"
    )
    with urllib.request.urlopen(req) as r:
        return json.load(r)["access_token"]


def _normalize_name(raw):
    """スタッフ名を正規化する（シートの注釈を除去してVenrey表示名に合わせる）。
    例: '川口 るか3/20K' → '川口るか'
        '星野 るな(吉沢)' → '星野るな'
        'たくま 1' → 'たくま'
    """
    name = raw.strip().replace(" ", "").replace("\u3000", "")
    # 括弧と内容を除去（例: 星野るな(吉沢) → 星野るな）
    name = re.sub(r'[（(][^）)]*[）)]', '', name)
    # 末尾の日付メモを除去（例: 3/20K, 1/15, 3/20 など）
    name = re.sub(r'\d+/\d+[A-Za-z]*$', '', name)
    # 末尾の数字・アルファベットを除去（複数回繰り返して完全除去）
    for _ in range(3):
        name = re.sub(r'\d+$', '', name)
        name = re.sub(r'[A-Za-z]+$', '', name)
    return name


def _fetch_sheet_df(sheet_name):
    """指定シート名の全行（隠し行含む）を Sheets API v4 で取得して DataFrame を返す。取得失敗時は None。"""
    try:
        access_token = _get_access_token()
        encoded_range = urllib.parse.quote(sheet_name)
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}"
            f"/values/{encoded_range}"
        )
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {access_token}"})
        with urllib.request.urlopen(req) as r:
            data = json.load(r)
        rows = data.get("values", [])
        if not rows:
            return None
        # 行ごとに列数を最大列数に揃えて DataFrame 化
        max_cols = max(len(row) for row in rows)
        padded = [row + [""] * (max_cols - len(row)) for row in rows]
        return pd.DataFrame(padded, dtype=str)
    except Exception as e:
        print(f"シート取得エラー ({sheet_name}): {e}")
        return None


def _build_date_map(df, base_year, base_month):
    """
    1行目の数値セルから {col_index: date} マッピングを作る。
    日付番号がリセット（減少）したら翌月に繰り上げる。
    これにより同一シート内の「来月プレビュー列」も正しく翌月日付として扱われる。
    日付列は整数または小数（例: 1.0）でも検出する。
    """
    date_map = {}
    cur_year, cur_month = base_year, base_month
    prev_day = 0
    for col_idx in range(3, df.shape[1]):
        val = df.iloc[0, col_idx]
        if not pd.notna(val):
            continue
        s = str(val).strip()
        # 「1 日」「31 火」など先頭の数字を抽出（曜日付き形式に対応）
        m = re.match(r'^(\d+)', s)
        if not m:
            continue
        day = int(m.group(1))
        if not (1 <= day <= 31):
            continue
        if day < prev_day:  # 日付がリセット → 翌月
            if cur_month == 12:
                cur_year, cur_month = cur_year + 1, 1
            else:
                cur_month += 1
        try:
            date_map[col_idx] = date(cur_year, cur_month, day)
        except ValueError:
            pass
        prev_day = day
    return date_map


def _parse_staff_rows(df, date_map):
    """スタッフ行をパースして [{名前: {日付: シフト}}, ...] を返す。"""
    schedules = [{}, {}]
    store_idx = 0
    for row_idx in range(2, df.shape[0]):
        raw_name = str(df.iloc[row_idx, 0]).strip()
        if not raw_name or raw_name == "nan":
            continue
        if STORE_SEPARATOR in raw_name:
            store_idx = min(store_idx + 1, len(schedules) - 1)
            continue
        name = _normalize_name(raw_name)
        if not name:
            continue
        # ふわもこSPA はSTORE2_STOP_AT の行の手前で読み込み停止
        if store_idx == 1 and name.startswith(STORE2_STOP_AT):
            break
        schedules[store_idx].setdefault(name, {})
        for col_idx, d in date_map.items():
            parsed = parse_time_cell(df.iloc[row_idx, col_idx])
            if parsed:
                schedules[store_idx][name][d] = parsed
    return schedules


def load_schedule():
    """
    Google Sheets からスケジュールを読み込む。

    ・来月シートが存在すれば来月分はそちらを参照
    ・来月シートが未作成の場合は今月シートのAJ列以降（来月プレビュー）を参照

    Returns:
        [{スタッフ名: {日付: シフト}}, ...]  店舗ごとのリスト
    """
    print("スプレッドシートを読み込み中...")

    today = datetime.now()

    # 最新の「YYYY年M月」シートを探す（直近6ヶ月を遡って検索）
    base_year, base_month = None, None
    df_base = None
    for offset in range(6):
        y = today.year if (today.month - offset) > 0 else today.year - 1
        m = (today.month - offset - 1) % 12 + 1
        sheet_name = f"{y}年{m}月"
        df = _fetch_sheet_df(sheet_name)
        if df is not None:
            base_year, base_month = y, m
            df_base = df
            print(f"  参照シート: 「{sheet_name}」")
            break

    if df_base is None:
        print("エラー: 「YYYY年M月」形式のシートが見つかりませんでした。")
        sys.exit(1)

    # 来月シートを確認
    next_year  = base_year if base_month < 12 else base_year + 1
    next_month = base_month + 1 if base_month < 12 else 1
    next_sheet = f"{next_year}年{next_month}月"
    df_next = _fetch_sheet_df(next_sheet)

    if df_next is not None:
        # 来月シートが存在するがデータが空の場合もあるため、まず内容を確認する
        print(f"  来月シート「{next_sheet}」も検出 → 内容を確認します")
        date_map_next = _build_date_map(df_next, next_year, next_month)
        next_schedules = _parse_staff_rows(df_next, date_map_next)
        # 「休み」は除外して実際の出勤シフトのみカウント
        next_total = sum(
            sum(1 for shift in dates.values() if shift != "休み")
            for s in next_schedules for dates in s.values()
        )

        if next_total > 0:
            # 来月シートに実際の出勤データあり → 当月分のみ + 来月シートを使う
            print(f"  来月シートにデータあり（{next_total}件） → 両方を読み込みます")
            date_map_base = _build_date_map(df_base, base_year, base_month)
            date_map_base = {c: d for c, d in date_map_base.items() if d.month == base_month}
            schedules = _parse_staff_rows(df_base, date_map_base)
            for i in range(2):
                for name, dates in next_schedules[i].items():
                    schedules[i].setdefault(name, {}).update(dates)
        else:
            # 来月シートは空 → 基準シートのAJ列以降（来月プレビュー）も含めて読む
            print(f"  来月シートは空 → AJ列以降の来月プレビューも参照します")
            date_map = _build_date_map(df_base, base_year, base_month)
            schedules = _parse_staff_rows(df_base, date_map)
    else:
        # 来月シート未作成 → 基準シートのAJ列以降（来月プレビュー）も含めて読む
        print(f"  来月シートなし → AJ列以降の来月プレビューも参照します")
        date_map = _build_date_map(df_base, base_year, base_month)
        schedules = _parse_staff_rows(df_base, date_map)

    for i, s in enumerate(schedules):
        total = sum(len(v) for v in s.values())
        print(f"店舗{i+1}: {len(s)} 人 / {total} 件のシフトデータ")
    return schedules


def get_current_week_dates():
    """
    今日から7日間（今日〜6日後）を対象期間として返す。
    """
    today = datetime.now().date()
    week_start = today
    week_end = today + timedelta(days=6)
    return week_start, week_end


def get_staff_id_map(page):
    """
    現在のページに表示されているスタッフ名 → data-id のマッピングを取得する。

    Step1: 位置ベースで全スタッフを取得（全員分のIDを確保）
    Step2: label[for] で取得できるものは上書き補正（位置ずれ修正）

    このハイブリッド方式により、全員分のIDが取れつつ
    label[for]で正確に特定できるスタッフは確実に正しいIDが割り当てられる。
    """
    result = page.evaluate("""
        () => {
            const map = {};
            const debugLines = [];

            // Step1: 位置ベースで全スタッフを取得
            const nameEls = [...document.querySelectorAll('.listGirl_name')];
            const seen = new Set();
            const schBoxIds = [];
            document.querySelectorAll('.schBox[data-id]').forEach(el => {
                const id = el.getAttribute('data-id');
                if (!seen.has(id)) { seen.add(id); schBoxIds.push(id); }
            });

            debugLines.push('[位置ベース] nameEls=' + nameEls.length + ' schBoxIds=' + schBoxIds.length);

            const minLen = Math.min(nameEls.length, schBoxIds.length);
            for (let i = 0; i < minLen; i++) {
                const name = nameEls[i].textContent.trim();
                if (name) map[name] = schBoxIds[i];
            }

            // Step2: label[for] で正確に特定できるスタッフを上書き補正
            let labelCount = 0;
            document.querySelectorAll('label[for]').forEach(label => {
                const nameEl = label.querySelector('.listGirl_name');
                if (!nameEl) return;
                const name = nameEl.textContent.trim();
                const forId = label.getAttribute('for');
                if (!name || !forId) return;
                // schBox[data-id=forId] が実際に存在するか確認してから使う
                const schBox = document.querySelector('.schBox[data-id="' + forId + '"]');
                if (schBox) {
                    if (map[name] && map[name] !== forId) {
                        debugLines.push('[補正] ' + name + ': ' + map[name] + ' → ' + forId);
                    }
                    map[name] = forId;
                    labelCount++;
                }
            });

            debugLines.push('[label補正] ' + labelCount + '件補正');
            map['__debug__'] = debugLines.join(' | ');

            return map;
        }
    """)
    # デバッグ情報を取り出してログ出力
    debug_msg = result.pop('__debug__', '')
    if debug_msg:
        print(f"  [IDマッピング] {debug_msg}")
    return result


def _get_search_input(page):
    """検索ボックスのロケーターを返す。"""
    return page.locator('input[placeholder="女性名で検索"], input[placeholder*="検索"]').first


def _clear_search(page):
    """検索ボックスをクリアして全スタッフ表示に戻す。"""
    try:
        inp = _get_search_input(page)
        if inp.count() > 0:
            inp.fill("")
            inp.press("Enter")
            time.sleep(0.8)
    except Exception:
        pass


def search_and_activate_staff(page, staff_name):
    """
    「女性名で検索」ボックスで名前を絞り込み、data-id を返す。
    ★検索フィルターはかけたまま（クリアしない）★
    呼び出し元が update_cell を済ませた後に _clear_search() を呼ぶこと。
    見つからない場合は None を返す（その場合も検索はクリアしない）。
    """
    try:
        inp = _get_search_input(page)
        if inp.count() == 0:
            return None
        inp.click(timeout=3000)
        inp.fill("")
        inp.type(staff_name, delay=50)
        time.sleep(1.5)  # Angular フィルタリングを待つ

        filtered_map = page.evaluate("""
            () => {
                const map = {};
                const nameEls = [...document.querySelectorAll('.listGirl_name')];
                const seen = new Set();
                const schBoxIds = [];
                document.querySelectorAll('.schBox[data-id]').forEach(el => {
                    const id = el.getAttribute('data-id');
                    if (!seen.has(id)) { seen.add(id); schBoxIds.push(id); }
                });
                const minLen = Math.min(nameEls.length, schBoxIds.length);
                for (let i = 0; i < minLen; i++) {
                    const name = nameEls[i].textContent.trim();
                    if (name) map[name] = schBoxIds[i];
                }
                document.querySelectorAll('label[for]').forEach(label => {
                    const nameEl = label.querySelector('.listGirl_name');
                    if (!nameEl) return;
                    const name = nameEl.textContent.trim();
                    const forId = label.getAttribute('for');
                    if (!name || !forId) return;
                    const schBox = document.querySelector('.schBox[data-id="' + forId + '"]');
                    if (schBox) map[name] = forId;
                });
                return map;
            }
        """)
        # 1. 完全一致
        if staff_name in filtered_map:
            return filtered_map[staff_name]

        # 2. 部分一致: Venrey名がシート名のprefixまたは逆（例: 川口るか ↔ 川口るか3/20）
        for venrey_name, vid in filtered_map.items():
            if staff_name.startswith(venrey_name) or venrey_name.startswith(staff_name):
                print(f"    [部分一致] シート名:{staff_name} ↔ Venrey名:{venrey_name}")
                return vid

        # 3. 検索結果が1件のみなら候補として使用
        if len(filtered_map) == 1:
            venrey_name, vid = next(iter(filtered_map.items()))
            print(f"    [1件マッチ] シート名:{staff_name} → Venrey名:{venrey_name}")
            return vid

        print(f"    [検索不一致] {staff_name}: 候補={list(filtered_map.keys())}")
        return None

    except Exception as e:
        print(f"    [検索エラー] {staff_name}: {e}")
        return None


def set_status_to_holiday(schbox_locator):
    """
    schBox のステータスを「休み」(off) に設定する。
    すでに off の場合は何もしない。
    pend（未設定）→ on（出勤）→ off（休み）の順でサイクルすると想定。
    """
    try:
        cls = schbox_locator.get_attribute("class", timeout=2000)
        if cls and "off" in cls:
            return  # すでに休み
        btn = schbox_locator.locator(".schBox_states")
        # 最大3回クリックして off になるまで試みる
        for _ in range(3):
            btn.click(timeout=3000)
            time.sleep(0.4)
            cls = schbox_locator.get_attribute("class", timeout=2000)
            if cls and "off" in cls:
                break
    except PlaywrightTimeout:
        pass


def set_status_to_working(page, schbox_locator):
    """
    schBox のステータスを「出勤」に設定する。
    pend（未設定）または off（休み）の場合に on（出勤）へ変更する。
    サイクル: pend → on → off → pend ...
    """
    try:
        cls = schbox_locator.get_attribute("class", timeout=2000)
        if cls and "on" in cls:
            return  # すでに出勤
        btn = schbox_locator.locator(".schBox_states")
        # 最大3回クリックして on になるまで試みる
        for i in range(3):
            btn.click(timeout=3000)
            time.sleep(0.6)
            cls = schbox_locator.get_attribute("class", timeout=2000)
            if cls and "on" in cls:
                break
    except PlaywrightTimeout:
        pass


def update_cell(page, data_id, target_date, start_time, end_time):
    """
    schBox[data-id][data-date] を特定して時間を入力する。

    Returns:
        True: 成功 / False: 失敗
    """
    # ISO 形式の日付文字列（+09:00 タイムゾーン）
    date_str = f"{target_date.strftime('%Y-%m-%d')}T00:00:00+09:00"
    selector = f'.schBox[data-id="{data_id}"][data-date="{date_str}"]'

    try:
        schbox = page.locator(selector)
        # セルがページ上に存在するか確認
        if schbox.count() == 0:
            return False

        # 必要に応じてスクロールして表示させる
        try:
            schbox.scroll_into_view_if_needed(timeout=3000)
        except PlaywrightTimeout:
            # 休み状態など scroll_into_view_if_needed が失敗する場合は JS でスクロール
            try:
                schbox.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
            except Exception:
                return False
        time.sleep(0.2)

        # 休みの場合はステータスを「休み」に設定して終了
        if start_time == "休み":
            set_status_to_holiday(schbox)
            time.sleep(0.2)
            return True

        # 出勤の場合: ステータスを「出勤」に設定して時間を入力
        was_off = "off" in (schbox.get_attribute("class", timeout=2000) or "")
        set_status_to_working(page, schbox)

        # 状態変更の確認（offのままなら失敗）
        cls_after = schbox.get_attribute("class", timeout=2000) or ""
        if "on" not in cls_after:
            print(f"    状態変更失敗 (class={cls_after})")
            return False

        # off→on の場合、Angularが時間入力フィールドをDOMに追加するまで待つ
        if was_off:
            try:
                schbox.locator("input.schBox_inputTime").first.wait_for(state="visible", timeout=5000)
            except PlaywrightTimeout:
                pass
            time.sleep(1.5)
        else:
            time.sleep(0.3)

        # 時間をJavaScriptで直接セットしAngularのchange検知を発火させる
        # keyboard.type()はコロンが特殊動作しAngularフィールドで誤動作するため使わない
        def set_input_value(locator, value):
            locator.evaluate("""(el, val) => {
                // Angularのchange検知をトリガーするためネイティブsetterを使う
                const setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                setter.call(el, val);
                el.dispatchEvent(new Event('input',  { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }""", value)

        # 開始時間
        start_input = schbox.locator("input.schBox_inputTime").first
        start_input.click(timeout=3000)
        time.sleep(0.15)
        set_input_value(start_input, start_time)
        time.sleep(0.3)

        # 終了時間
        end_input = schbox.locator('input[data-role="end-time"]')
        if end_input.count() > 0:
            end_input.click(timeout=3000)
            time.sleep(0.15)
            set_input_value(end_input, end_time)
        else:
            end_input2 = schbox.locator("input.schBox_inputTime").nth(1)
            end_input2.click(timeout=3000)
            time.sleep(0.15)
            set_input_value(end_input2, end_time)
        time.sleep(0.3)

        # フォーカスを外して確定（Tab キー）
        page.keyboard.press("Tab")
        time.sleep(0.5)
        return True

    except PlaywrightTimeout:
        print(f"    タイムアウト (id={data_id}, date={target_date})")
        return False
    except Exception as e:
        print(f"    エラー: {e}")
        return False




def main():
    # ── Step 1: スプレッドシートを読み込む ──
    schedules = load_schedule()

    # ── Step 1b: SELECTED_STAFF 環境変数で絞り込み（UIからの選択）──
    selected_raw = os.environ.get("SELECTED_STAFF", "").strip()
    if selected_raw:
        selected_normalized = set(
            _normalize_name(n) for n in selected_raw.split(",") if n.strip()
        )
        print(f"選択スタッフ絞り込み: {sorted(selected_normalized)}")
        schedules = [
            {k: v for k, v in sched.items() if k in selected_normalized}
            for sched in schedules
        ]

    # ── Step 2: 今週の日付範囲を特定 ──
    week_start, week_end = get_current_week_dates()
    print(f"今週: {week_start.strftime('%Y/%m/%d')} 〜 {week_end.strftime('%Y/%m/%d')}")

    # ── Step 3: ブラウザ起動 → 店舗ごとにログイン → 更新 ──
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)

        for store_idx, store in enumerate(STORES):
            label_store = f"店舗{store_idx + 1}"

            # 今週分のデータだけ抽出
            schedule = schedules[store_idx]
            this_week = {
                name: {d: t for d, t in dates.items() if week_start <= d <= week_end}
                for name, dates in schedule.items()
            }
            this_week = {name: dates for name, dates in this_week.items() if dates}


            total_entries = sum(len(v) for v in this_week.values())
            print(f"\n{label_store}: 今週の更新対象 {len(this_week)} 人 / {total_entries} 件")

            # 各スタッフの対象日程を表示（デバッグ用）
            for nm, dts in this_week.items():
                dates_str = ", ".join(d.strftime("%m/%d") for d in sorted(dts))
                print(f"  {nm}: {dates_str}")

            if not this_week:
                print(f"{label_store}: 今週の出勤データがありません。スキップします。")
                continue

            context = browser.new_context()
            page = context.new_page()
            page.set_viewport_size({"width": 1600, "height": 900})

            # ログイン
            print(f"{label_store} にログイン中...")
            page.goto(LOGIN_URL)
            page.wait_for_load_state("networkidle", timeout=20000)

            page.locator("input").first.fill(store["id"])
            page.locator('input[type="password"]').fill(store["password"])
            page.locator('button[type="submit"], button:has-text("ログイン")').first.click()
            page.wait_for_load_state("networkidle", timeout=20000)
            # Angular SPA のメニュー描画を待つ
            page.wait_for_selector("text=週間スケジュール", state="visible", timeout=30000)
            time.sleep(2)
            print("ログイン完了")

            # 週間スケジュールへ移動
            page.locator("text=週間スケジュール").first.click()
            page.wait_for_load_state("networkidle", timeout=20000)
            print("週間スケジュール画面を開きました")

            # 「今週」ボタンで現在の週に移動
            try:
                page.locator('button:has-text("今週"), a:has-text("今週")').first.click(timeout=3000)
                page.wait_for_load_state("networkidle", timeout=10000)
            except PlaywrightTimeout:
                pass

            # Angular がスケジュールセルを完全に初期化するまで待つ
            try:
                page.wait_for_selector('.schBox[data-id]', state='visible', timeout=15000)
            except PlaywrightTimeout:
                pass
            time.sleep(3)

            # ── 全スタッフを一括更新（今週・来週の2画面をカバー）──
            updated = 0
            failed  = 0

            # 未更新の (staff_name, target_date, shift) を管理
            pending = []
            for staff_name, date_times in this_week.items():
                for target_date, shift in sorted(date_times.items()):
                    if week_start <= target_date <= week_end:
                        pending.append((staff_name, target_date, shift))

            # Venreyに存在するか確認するためのセット（3画面で見つかった名前を追跡）
            found_in_venrey = set()

            # 先週・今週・来週の最大3画面を試みる
            for screen in ["今週", "来週", "先週"]:
                if not pending:
                    break

                if screen == "来週":
                    # 翌週ボタンをクリック
                    print("\n来週画面に移動して残りを更新します...")
                    try:
                        page.locator(
                            'button:has-text("翌週"), button:has-text("次週"), '
                            'button:has-text("次の週"), a:has-text("翌週"), '
                            'a:has-text("次週"), .next, [aria-label="next"]'
                        ).first.click(timeout=5000)
                        page.wait_for_load_state("networkidle", timeout=15000)
                        time.sleep(1)
                    except PlaywrightTimeout:
                        print("  翌週ボタンが見つかりませんでした。スキップします。")
                        break

                elif screen == "先週":
                    # 今週ボタンで戻ってから前週ボタンで先週へ
                    print("\n先週画面に移動して残りを更新します...")
                    try:
                        page.locator('button:has-text("今週"), a:has-text("今週")').first.click(timeout=5000)
                        page.wait_for_load_state("networkidle", timeout=15000)
                        time.sleep(0.5)
                    except PlaywrightTimeout:
                        pass
                    try:
                        page.locator(
                            'button:has-text("前週"), button:has-text("先週"), '
                            'button:has-text("前の週"), a:has-text("前週"), '
                            'a:has-text("先週"), .prev, [aria-label="prev"]'
                        ).first.click(timeout=5000)
                        page.wait_for_load_state("networkidle", timeout=15000)
                        time.sleep(1)
                    except PlaywrightTimeout:
                        print("  前週ボタンが見つかりませんでした。スキップします。")
                        break

                staff_id_map = get_staff_id_map(page)
                if screen == "今週":
                    print(f"\n管理画面のスタッフ数: {len(staff_id_map)} 人")
                    print("  [シート側の名前]:", list(this_week.keys())[:5])
                    print("  [管理画面の名前 → data-id]:")
                    for n, did in list(staff_id_map.items())[:10]:
                        print(f"    {n} → {did}")
                found_in_venrey.update(staff_id_map.keys())

                still_pending = []
                for (staff_name, target_date, shift) in pending:
                    used_search = False
                    if staff_name not in staff_id_map:
                        # 仮想スクロールで見えていない可能性 → 検索ボックスで絞り込んで再試行
                        # ★フィルターはかけたまま → update_cell まで DOM に残す★
                        found_id = search_and_activate_staff(page, staff_name)
                        if found_id:
                            staff_id_map[staff_name] = found_id
                            found_in_venrey.add(staff_name)
                            used_search = True
                            print(f"    [検索で発見] {staff_name} → {found_id}")
                        else:
                            _clear_search(page)
                            still_pending.append((staff_name, target_date, shift))
                            continue

                    data_id = staff_id_map[staff_name]

                    if shift == "休み":
                        shift_label, start, end = "休み", "休み", ""
                    else:
                        start, end = shift
                        shift_label = f"{start}〜{end}"

                    print(f"  更新: {staff_name} / {target_date.strftime('%m/%d')} {shift_label}")
                    success = update_cell(page, data_id, target_date, start, end)

                    # 検索フィルターを使った場合はここでクリアして全表示に戻す
                    if used_search:
                        _clear_search(page)

                    if success:
                        updated += 1
                        print("    → 完了")
                    else:
                        still_pending.append((staff_name, target_date, shift))
                        if screen != "先週":
                            print("    → この画面にないため次の画面で再試行")
                        else:
                            print("    → 更新失敗")

                pending = still_pending

            # 最終失敗を「Venrey未登録」と「更新失敗」に分類
            not_in_venrey = [(n, d, s) for n, d, s in pending if n not in found_in_venrey]
            real_failures = [(n, d, s) for n, d, s in pending if n in found_in_venrey]

            if not_in_venrey:
                unregistered = sorted({n for n, _, _ in not_in_venrey})
                print(f"  Venrey未登録スタッフ（スキップ）: {', '.join(unregistered)}")
            if real_failures:
                for staff_name, target_date, _ in real_failures:
                    print(f"  最終失敗: {staff_name} / {target_date.strftime('%m/%d')}")

            failed = len(real_failures)

            # ── 店舗ごとの完了報告 ──
            print(f"\n{'=' * 40}")
            print(f"{label_store} 完了！  更新: {updated} 件 / 失敗: {failed} 件")
            if failed > 0:
                print("失敗したセルは手動で確認・入力してください。")
            print("=" * 40)

            # Angular の保存リクエストが完了するまで待つ（context.close() でキャンセルされるのを防ぐ）
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except PlaywrightTimeout:
                pass
            time.sleep(5)

            context.close()

        time.sleep(3)
        browser.close()


if __name__ == "__main__":
    main()






