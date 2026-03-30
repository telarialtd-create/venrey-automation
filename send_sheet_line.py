#!/usr/bin/env python3
"""
スプレッドシートの指定範囲をLINEに自動送信するスクリプト
当月シート(YYYY年M月)を自動検出し、今日から1週間分を送信
"""
import json
import os
import subprocess
import tempfile
import urllib.request
import urllib.parse
from datetime import date

# 設定
SHEET_ID = "10siqLe6B9A7uvNWgRUdHb462RqxCxkGEGMEKTPhY-S8"
LINE_TOKEN = "ifKMFJwttgSGoWsmSEx0WTWETYx+pauurDW4cFjO/cyszJ7Pxa1ahQg2BFaQ6TFzMqzXTX5U+Xrl0T58bVSumOVOvMnj4e3AjP9NIOv+o3xYJUTqdRG+gIOR0YYhEv7KrJVVslDy+r23STaPvSwRMwdB04t89/1O/w1cDnyilFU="
LINE_USER_ID = "Ufa7625fe16c66fd60aebc14b32a74220"
LINE_GROUP_ID = "Cfc2a40e8199d10556f696690d5885964"
OAUTH_KEYS_PATH = os.environ.get("OAUTH_KEYS_PATH", "/Users/hiraokawashin/.config/gcp-oauth.keys.json")
CREDENTIALS_PATH = os.environ.get("CREDENTIALS_PATH", "/Users/hiraokawashin/.config/gdrive-server-credentials.json")

ROW_START = 1
WEEK_COLS = 7   # 1週間 = 7列


def col_num_to_letter(n):
    """列番号をアルファベットに変換 (1→A, 34→AH)"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_sheet_info(access_token):
    """当月シート(YYYY年M月)のタイトルとGIDを取得"""
    today = date.today()
    target = f"{today.year}年{today.month}月"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}?fields=sheets.properties"
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {access_token}")
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    sheets = {s["properties"]["title"].strip(): (str(s["properties"]["sheetId"]), s["properties"]["title"].strip())
              for s in data["sheets"]}
    if target in sheets:
        gid, title = sheets[target]
        print(f"    シート検出: {title}")
        return gid, title
    for m in range(today.month - 1, 0, -1):
        key = f"{today.year}年{m}月"
        if key in sheets:
            gid, title = sheets[key]
            print(f"    シート検出(フォールバック): {title}")
            return gid, title
    raise RuntimeError(f"シートが見つかりません: {target}")


def get_date_column(access_token, sheet_title):
    """行1を検索して今日の日付の列番号を返す（隠し列も考慮）"""
    today = date.today()
    encoded = urllib.parse.quote(sheet_title)
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/"
           f"{encoded}!1:1?majorDimension=ROWS")
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {access_token}")
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    values = data.get("values", [[]])[0]
    # 今月のシートなら最初の一致、前月フォールバックなら最後の一致
    matches = [i + 1 for i, v in enumerate(values)
               if str(v).strip() == str(today.day)]
    if not matches:
        raise RuntimeError(f"行1に {today.day} 日が見つかりません")
    # 当月シート: 最初の一致が当月の日付
    # フォールバック(前月): 最後の一致が翌月(=今月)の日付
    sheet_month = int(sheet_title.replace("年", "月").split("月")[1]) if "年" in sheet_title else today.month
    if sheet_month == today.month:
        return matches[0]
    else:
        return matches[-1]  # 前月シート内の今月分


def get_last_row(access_token, gid):
    """A列の「講習」行番号を動的に取得"""
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/A1:A200"
           f"?majorDimension=ROWS&ranges=A1:A200")
    # GID指定でシートを特定
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/A1:A200"
           f"?majorDimension=ROWS")
    req = urllib.request.Request(url + f"&gid={gid}")
    req.add_header("Authorization", f"Bearer {access_token}")
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read())
    except Exception:
        # gidパラメータが効かない場合はsheetId形式で試みる
        data = {"values": []}
    for i, row in enumerate(data.get("values", []), 1):
        if row and "講習" in row[0]:
            return i
    return 108  # フォールバック


def get_range(access_token, gid, sheet_title):
    """行1を検索して今日の列を特定し、1週間分の範囲を返す"""
    start_col = get_date_column(access_token, sheet_title)
    end_col = start_col + WEEK_COLS - 1
    last_row = get_last_row(access_token, gid)
    start = f"{col_num_to_letter(start_col)}{ROW_START}"
    end = f"{col_num_to_letter(end_col)}{last_row}"
    print(f"    今日の列: {col_num_to_letter(start_col)}（列{start_col}）")
    return f"{start}:{end}", last_row, start_col, end_col


def get_access_token():
    """OAuth refresh tokenを使ってアクセストークンを取得"""
    with open(CREDENTIALS_PATH) as f:
        creds = json.load(f)
    with open(OAUTH_KEYS_PATH) as f:
        keys = json.load(f)["installed"]

    data = urllib.parse.urlencode({
        "client_id": keys["client_id"],
        "client_secret": keys["client_secret"],
        "refresh_token": creds["refresh_token"],
        "grant_type": "refresh_token"
    }).encode()

    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=data,
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        token_data = json.loads(resp.read())

    return token_data["access_token"]


def export_range_as_png(access_token, range_str, out_path, gid):
    """指定範囲をPDF→PNGで取得"""
    params = urllib.parse.urlencode({
        "format": "pdf",
        "range": range_str,
        "gid": gid,
        "portrait": "true",
        "size": "6",       # A3
        "fith": "true",    # 高さ基準でスケール→両エクスポートで行高さが同一になる
        "top_margin": "0",
        "bottom_margin": "0",
        "left_margin": "0",
        "right_margin": "0",
        "sheetnames": "false",
        "printtitle": "false",
        "pagenumbers": "false",
        "gridlines": "true",
        "notes": "false",  # セルのメモ脚注を除外
    })
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?{params}"
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {access_token}")
    with urllib.request.urlopen(req) as resp:
        pdf_data = resp.read()

    pdf_path = out_path.replace(".png", ".pdf")
    with open(pdf_path, "wb") as f:
        f.write(pdf_data)

    # pymupdfで全ページを変換して縦結合（脚注ページ除外・繰り返しヘッダー除去）
    import fitz
    from PIL import Image
    import numpy as np

    doc = fitz.open(pdf_path)
    page_imgs = []
    for page in doc:
        mat = fitz.Matrix(2, 2)  # 2x解像度
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        # グリッド線（全幅にわたる暗い行）があるページだけ採用
        gray = np.array(img.convert("L"))
        dark_per_row = (gray < 80).sum(axis=1)
        has_gridline = bool((dark_per_row > img.width * 0.15).any())
        if has_gridline:
            page_imgs.append((img, dark_per_row))
    doc.close()

    if not page_imgs:
        raise RuntimeError("有効なページが見つかりません")

    if len(page_imgs) == 1:
        combined = page_imgs[0][0]
    else:
        # 2ページ目以降の繰り返しヘッダー高さを検出して除去
        strips = [page_imgs[0][0]]
        for img, dark_per_row in page_imgs[1:]:
            # 上部200px以内で最初の全幅罫線を探す（凍結ヘッダー1行の下枠）
            # 動的閾値 = 上部200px内の最大暗ピクセル数の80%（コンテンツ幅に依存しない）
            top_rows = dark_per_row[:min(200, len(dark_per_row))]
            gridline_thresh = max(top_rows) * 0.8 if max(top_rows) > 0 else img.width * 0.5
            header_end = 0
            for y in range(len(top_rows)):
                if top_rows[y] >= gridline_thresh:
                    header_end = y + 1
                    break  # 最初の全幅罫線のみ（ヘッダー行末）
            strips.append(img.crop((0, header_end, img.width, img.height)))

        total_h = sum(i.height for i in strips)
        max_w = max(i.width for i in strips)
        combined = Image.new("RGB", (max_w, total_h), "white")
        y = 0
        for img in strips:
            combined.paste(img, (0, y))
            y += img.height

    combined.save(out_path)
    return out_path


def autocrop(img):
    """白い余白を自動クロップ（上下左右）"""
    from PIL import ImageOps
    # グレースケールで閾値処理して余白を検出
    gray = img.convert("L")
    # 253以上（ほぼ白）を背景とみなす
    import PIL.Image as PILImage
    mask = gray.point(lambda p: 0 if p >= 253 else 255)
    bbox = mask.getbbox()
    if bbox:
        # 少しだけ余白を残す
        pad = 4
        w, h = img.size
        bbox = (
            max(0, bbox[0] - pad),
            max(0, bbox[1] - pad),
            min(w, bbox[2] + pad),
            min(h, bbox[3] + pad),
        )
        return img.crop(bbox)
    return img


def get_col_positions(access_token, sheet_title, start_col, end_col):
    """列A〜end_colの幅を取得し、表示列の境界ピクセル比率を返す"""
    from_col = "A"
    to_col = col_num_to_letter(end_col)
    encoded = urllib.parse.quote(sheet_title)
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
           f"?includeGridData=true&ranges={encoded}!{from_col}1:{to_col}1"
           f"&fields=sheets.data.columnMetadata")
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {access_token}")
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    meta = data["sheets"][0]["data"][0].get("columnMetadata", [])
    widths = [m.get("pixelSize", 100) for m in meta]

    # 累積幅（0始まり）
    cum = [0]
    for w in widths:
        cum.append(cum[-1] + w)
    total = cum[-1]

    # A〜C右端（=D左端）のパーセント
    fixed_right_pct  = cum[3] / total          # 列D左端 = 列A〜C右端
    # 週の開始列左端〜終了列右端のパーセント
    week_left_pct    = cum[start_col - 1] / total
    week_right_pct   = cum[end_col]     / total

    return fixed_right_pct, week_left_pct, week_right_pct


def download_sheet_as_png(access_token, range_str, gid, last_row, sheet_title, start_col, end_col):
    """左固定列(A:C)と週列を別々にfith=trueでエクスポート→同一行高さで横結合"""
    from PIL import Image

    tmp_dir = tempfile.mkdtemp()
    week_col = col_num_to_letter(start_col)
    week_end = col_num_to_letter(end_col)

    left_range = f"A1:C{last_row}"
    right_range = f"{week_col}1:{week_end}{last_row}"

    left_png  = export_range_as_png(access_token, left_range,
                                    os.path.join(tmp_dir, "left.png"), gid)
    right_png = export_range_as_png(access_token, right_range,
                                    os.path.join(tmp_dir, "right.png"), gid)

    def crop_whitespace(img, pad=4):
        """上下左右の白い余白を除去"""
        gray = img.convert("L")
        mask = gray.point(lambda p: 0 if p >= 253 else 255)
        bbox = mask.getbbox()
        if bbox:
            w, h = img.size
            return img.crop((
                max(0, bbox[0] - pad),
                max(0, bbox[1] - pad),
                min(w, bbox[2] + pad),
                min(h, bbox[3] + pad),
            ))
        return img

    left_img  = crop_whitespace(Image.open(left_png))
    right_img = crop_whitespace(Image.open(right_png))

    # 両画像の高さを揃えて横結合（高さが同じになるはずだが念のため）
    H = max(left_img.height, right_img.height)
    if left_img.height < H:
        tmp = Image.new("RGB", (left_img.width, H), "white")
        tmp.paste(left_img, (0, 0))
        left_img = tmp
    if right_img.height < H:
        tmp = Image.new("RGB", (right_img.width, H), "white")
        tmp.paste(right_img, (0, 0))
        right_img = tmp

    # 両側にコンテンツがある高さまでに制限（短い方に合わせる）
    H = min(left_img.height, right_img.height)
    combined = Image.new("RGB", (left_img.width + right_img.width, H), "white")
    combined.paste(left_img.crop((0, 0, left_img.width, H)),  (0, 0))
    combined.paste(right_img.crop((0, 0, right_img.width, H)), (left_img.width, 0))

    out_path = os.path.join(tmp_dir, "combined.jpg")
    combined.save(out_path, "JPEG", quality=90)
    return out_path


def upload_image(image_path):
    """画像をアップロードして公開URLを取得（複数サービスでフォールバック）"""
    services = [
        # catbox.moe
        lambda: subprocess.run(
            ["curl", "-s", "-F", "reqtype=fileupload", "-F", f"fileToUpload=@{image_path}",
             "https://catbox.moe/user/api.php"],
            capture_output=True, text=True
        ).stdout.strip(),
        # litterbox.catbox.moe (72時間)
        lambda: subprocess.run(
            ["curl", "-s", "-F", "reqtype=fileupload", "-F", "time=72h",
             "-F", f"fileToUpload=@{image_path}", "https://litterbox.catbox.moe/resources/internals/api.php"],
            capture_output=True, text=True
        ).stdout.strip(),
        # transfer.sh
        lambda: subprocess.run(
            ["curl", "-s", "--upload-file", image_path, "https://transfer.sh/sheet.jpg"],
            capture_output=True, text=True
        ).stdout.strip(),
    ]
    for fn in services:
        try:
            url = fn()
            if url.startswith("https://"):
                return url
        except Exception:
            continue
    raise RuntimeError("全アップロードサービスが失敗しました")


def push_message(to_id, image_url, range_str):
    """指定IDにLINE送信"""
    message_data = json.dumps({
        "to": to_id,
        "notificationDisabled": True,
        "messages": [
            {
                "type": "image",
                "originalContentUrl": image_url,
                "previewImageUrl": image_url
            }
        ]
    }).encode()

    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/push",
        data=message_data,
        headers={
            "Authorization": f"Bearer {LINE_TOKEN}",
            "Content-Type": "application/json"
        }
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def send_to_line(image_url, range_str):
    """個人に送信（テスト用）"""
    return push_message(LINE_USER_ID, image_url, range_str)


def main():
    print("[1] アクセストークン取得中...")
    access_token = get_access_token()

    gid, sheet_title = get_sheet_info(access_token)
    range_str, last_row, start_col, end_col = get_range(access_token, gid, sheet_title)
    print(f"    対象範囲: {range_str}（講習行: {last_row}）")

    print("[2] スプレッドシートをPDF→PNG変換中...")
    image_path = download_sheet_as_png(access_token, range_str, gid, last_row,
                                       sheet_title, start_col, end_col)
    print(f"    保存先: {image_path}")

    print("[4] 画像アップロード中...")
    image_url = upload_image(image_path)
    print(f"    URL: {image_url}")

    print("[5] LINE送信中...")
    result = send_to_line(image_url, range_str)
    print(f"[完了] {result}")


if __name__ == "__main__":
    main()
