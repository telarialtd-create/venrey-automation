#!/usr/bin/env python3
"""
Mr.Venrey 週間スケジュール自動更新スクリプト

スプレッドシートから今週の出勤データを読み込み、
Venrey管理画面の週間スケジュールを自動更新する。

使い方:
  python3 main.py
"""

import io
import re
import sys
import time
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

# スプレッドシート ID（URL の /d/ と /edit の間の文字列）
SPREADSHEET_ID = "10siqLe6B9A7uvNWgRUdHb462RqxCxkGEGMEKTPhY-S8"

# 今月のシート名（例: "4月"）を自動生成して URL を組み立てる
# シート名でアクセスするため、毎月 GID を変更する必要なし
_current_month_sheet = f"{datetime.now().month}月"
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


def load_schedule():
    """
    Google Sheets の CSV エクスポートからスケジュールを読み込む。

    Returns:
        {スタッフ名(str): {日付(date): (開始時間, 終了時間)}}
    """
    print("スプレッドシートを読み込み中...")
    try:
        resp = requests.get(SHEET_CSV_URL, timeout=15)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.content.decode("utf-8")), header=None, dtype=str)
    except Exception as e:
        print(f"エラー: スプレッドシートの取得に失敗しました: {e}")
        print("スプレッドシートを「リンクを知っている全員が閲覧可」に設定してください。")
        sys.exit(1)

    # 対象年月を実行時に自動取得
    today = datetime.now()
    year, month = today.year, today.month

    # スプレッドシート構造:
    #   Row 0: "3月😊" | "" | "" | "1" | "2" | ... | "31"   ← 日付番号
    #   Row 1: 名前ヘッダ | "確認日" | "出勤日数" | "日" | "月" | ...
    #   Row 2+: スタッフ名 | 確認日 | 出勤日数 | シフト時間 | ...

    # 日付カラムマッピング: {col_index: date}
    date_map = {}
    for col_idx in range(3, df.shape[1]):
        val = df.iloc[0, col_idx]
        if pd.notna(val) and str(val).strip().isdigit():
            try:
                day = int(str(val).strip())
                date_map[col_idx] = date(year, month, day)
            except ValueError:
                pass

    # スタッフデータ読み込み（Row 2 以降）
    # STORE_SEPARATOR 行を境に店舗ごとに分ける
    schedules = [{}, {}]  # [店舗1, 店舗2]
    store_idx = 0

    for row_idx in range(2, df.shape[0]):
        raw_name = str(df.iloc[row_idx, 0]).strip()
        if not raw_name or raw_name == "nan":
            continue

        # 店舗区切り行を検出したら次の店舗へ切り替え
        if STORE_SEPARATOR in raw_name:
            store_idx = min(store_idx + 1, len(schedules) - 1)
            continue

        # スペースと末尾の数字を除去してVenrey管理画面の名前と合わせる
        # 例: "桜餅 ねる 121" → "桜餅ねる"
        name = raw_name.replace(" ", "").replace("\u3000", "")
        name = re.sub(r'\d+$', '', name)
        if not name:
            continue

        schedules[store_idx][name] = {}
        for col_idx, d in date_map.items():
            parsed = parse_time_cell(df.iloc[row_idx, col_idx])
            if parsed:
                schedules[store_idx][name][d] = parsed

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

    HTML構造:
      <label for="3934032">
        ...
        <p class="listGirl_name">柚月のあ</p>
        ...
      </label>
    """
    result = page.evaluate("""
        () => {
            const map = {};
            // listGirl_name を含む label[for] 要素を探す
            const labels = document.querySelectorAll('label[for]');
            for (const label of labels) {
                const nameEl = label.querySelector('.listGirl_name');
                if (nameEl) {
                    const name = nameEl.textContent.trim();
                    const id = label.getAttribute('for');
                    if (name && id) {
                        map[name] = id;
                    }
                }
            }
            return map;
        }
    """)
    return result


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
    現在の状態が pend（未設定）の場合のみクリックする。
    """
    try:
        cls = schbox_locator.get_attribute("class", timeout=2000)
        if cls and "pend" in cls:
            # 未設定 → 出勤 にするため schBox_states ボタンをクリック
            btn = schbox_locator.locator(".schBox_states")
            btn.click(timeout=3000)
            time.sleep(0.4)
            # クリック後に「出勤」になっていなければもう一度クリック（サイクルする場合）
            cls2 = schbox_locator.get_attribute("class", timeout=2000)
            if cls2 and "on" not in cls2:
                btn.click(timeout=3000)
                time.sleep(0.3)
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
        schbox.scroll_into_view_if_needed(timeout=3000)
        time.sleep(0.2)

        # 休みの場合はステータスを「休み」に設定して終了
        if start_time == "休み":
            set_status_to_holiday(schbox)
            time.sleep(0.2)
            return True

        # 出勤の場合: ステータスを「出勤」に設定して時間を入力
        set_status_to_working(page, schbox)

        # 開始時間の入力（data-role 属性のない 1 つ目の schBox_inputTime）
        start_input = schbox.locator("input.schBox_inputTime").first
        start_input.click(timeout=3000)
        time.sleep(0.2)
        page.keyboard.press("Control+a")
        start_input.fill(start_time)
        time.sleep(0.2)

        # 終了時間の入力（data-role="end-time" の input）
        end_input = schbox.locator('input[data-role="end-time"]')
        if end_input.count() > 0:
            end_input.click(timeout=3000)
            time.sleep(0.2)
            page.keyboard.press("Control+a")
            end_input.fill(end_time)
        else:
            # data-role がない場合は 2 つ目の input を使う
            end_input2 = schbox.locator("input.schBox_inputTime").nth(1)
            end_input2.click(timeout=3000)
            time.sleep(0.2)
            page.keyboard.press("Control+a")
            end_input2.fill(end_time)
        time.sleep(0.2)

        # フォーカスを外して確定（Tab キー）
        page.keyboard.press("Tab")
        time.sleep(0.3)
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

            if not this_week:
                print(f"{label_store}: 今週の出勤データがありません。スキップします。")
                continue

            page = browser.new_page()
            page.set_viewport_size({"width": 1600, "height": 900})

            # ログイン
            print(f"{label_store} にログイン中...")
            page.goto(LOGIN_URL)
            page.wait_for_load_state("networkidle", timeout=20000)

            page.locator("input").first.fill(store["id"])
            page.locator('input[type="password"]').fill(store["password"])
            page.locator('button[type="submit"], button:has-text("ログイン")').first.click()
            page.wait_for_load_state("networkidle", timeout=20000)
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

            # pager2 の表示人数セレクトを 400 に変更して全スタッフを 1 ページに表示
            try:
                page.locator("pager2 select").first.select_option(value="400")
                page.wait_for_load_state("networkidle", timeout=15000)
                print("表示人数を 400 人に変更しました")
            except PlaywrightTimeout:
                print("表示人数の変更がタイムアウトしました（続行します）")
            except Exception as e:
                print(f"表示人数の変更に失敗しました: {e}（続行します）")

            # ── 全スタッフを一括更新 ──
            updated = 0
            failed = 0

            # スタッフ名 → data-id マッピングを取得
            staff_id_map = get_staff_id_map(page)
            print(f"\n管理画面のスタッフ数: {len(staff_id_map)} 人")
            print("  [シート側の名前]:", list(this_week.keys())[:5])
            print("  [管理画面の名前]:", list(staff_id_map.keys())[:5])

            for staff_name, date_times in this_week.items():
                if staff_name not in staff_id_map:
                    print(f"  スキップ（管理画面に見つかりません）: {staff_name}")
                    continue

                data_id = staff_id_map[staff_name]

                for target_date, shift in sorted(date_times.items()):
                    if not (week_start <= target_date <= week_end):
                        continue

                    # shift は ("開始", "終了") または "休み"
                    if shift == "休み":
                        shift_label = "休み"
                        start, end = "休み", ""
                    else:
                        start, end = shift
                        shift_label = f"{start}〜{end}"

                    print(f"  更新: {staff_name} / {target_date.strftime('%m/%d')} {shift_label}")
                    success = update_cell(page, data_id, target_date, start, end)

                    if success:
                        updated += 1
                        print("    → 完了")
                    else:
                        failed += 1
                        print("    → 失敗（スキップ）")

            # ── 店舗ごとの完了報告 ──
            print(f"\n{'=' * 40}")
            print(f"{label_store} 完了！  更新: {updated} 件 / 失敗: {failed} 件")
            if failed > 0:
                print("失敗したセルは手動で確認・入力してください。")
            print("=" * 40)

            page.close()

        time.sleep(3)
        browser.close()


if __name__ == "__main__":
    main()
