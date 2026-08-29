"""
senseki_sheets.py
戦績データをGoogleスプレッドシートへ同期する

senseki.py から呼び出される補助モジュール。
このファイル単体ではDiscordのコマンドを持たない。

【必要な環境変数】
  GOOGLE_SERVICE_ACCOUNT_JSON … サービスアカウントのJSONキー（全文をそのまま貼る）
  SENSEKI_SPREADSHEET_ID       … 対象スプレッドシートのID
                                  URLの /d/ と /edit の間の文字列

【必要なライブラリ（requirements.txt）】
  gspread
  google-auth

【事前準備】
  1. Google Cloud でプロジェクトを作り、Google Sheets API を有効化する
  2. サービスアカウントを作成し、JSONキーをダウンロードする
  3. 対象のスプレッドシートを、そのサービスアカウントのメールアドレスに
     「編集者」として共有する（これを忘れると 403 になる）

環境変数が未設定でも import は成功し、is_configured() が False を返すだけ。
BOT全体の起動を止めないための作り。
"""

import json
import os

try:
    import gspread
    from google.oauth2.service_account import Credentials
    SHEETS_IMPORT_ERROR = None
except Exception as _e:  # ライブラリ未導入時に起動を止めない
    gspread = None
    Credentials = None
    SHEETS_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

RAW_SHEET_NAME = "生データ"
AGG_SHEET_NAME = "集計"


def is_configured() -> bool:
    """同期に必要なものが揃っているか"""
    if gspread is None:
        return False
    if not os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
        return False
    if not os.environ.get("SENSEKI_SPREADSHEET_ID"):
        return False
    return True


def describe_config() -> str:
    """診断用。何が足りないかを返す"""
    lines = []
    if gspread is None:
        lines.append(f"ライブラリ: 未読込（{SHEETS_IMPORT_ERROR}）")
    else:
        lines.append(f"ライブラリ: 読込済（gspread {gspread.__version__}）")

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        lines.append("GOOGLE_SERVICE_ACCOUNT_JSON: 未設定")
    else:
        try:
            info = json.loads(raw)
            email = info.get("client_email", "(client_email が見つかりません)")
            lines.append(f"サービスアカウント: {email}")
        except json.JSONDecodeError as e:
            lines.append(f"GOOGLE_SERVICE_ACCOUNT_JSON: JSONとして読めません（{e}）")

    sheet_id = os.environ.get("SENSEKI_SPREADSHEET_ID")
    lines.append(f"SENSEKI_SPREADSHEET_ID: {sheet_id or '未設定'}")
    return "\n".join(lines)


def _get_client():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def _write_sheet(spreadsheet, title: str, rows: list[list]):
    """シートを作り直して全行を書き込む"""
    try:
        ws = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=title, rows=max(len(rows) + 10, 100), cols=max(len(rows[0]) if rows else 1, 12)
        )

    ws.clear()
    if rows:
        # update() は gspread 5.x と 6.x で引数の順序が違うため、
        # 仕様が安定している append_rows を使う（clear 直後なので1行目から入る）
        ws.append_rows(rows, value_input_option="RAW")
    return ws


def sync_to_sheets(matches: list[dict], format_labels: dict, resolve_name) -> dict:
    """
    戦績データをスプレッドシートへ全面上書きで同期する。

    matches       : matches テーブルの全行（dictのリスト）
    format_labels : {"rotation": "ローテーション", ...}
    resolve_name  : user_id を表示名に変換する関数

    戻り値: {"url": スプレッドシートURL, "raw": 生データ件数, "users": ユーザー数}

    同期は毎回すべて消して書き直す方式。差分更新にすると、取消された
    レコードがシートに残り続けて実データと食い違うため。
    """
    client = _get_client()
    sheet_id = os.environ["SENSEKI_SPREADSHEET_ID"].strip()
    sh = client.open_by_key(sheet_id)

    # ---- 生データ ----
    raw_rows = [[
        "記録日時", "ユーザー", "ユーザーID", "環境ID", "フォーマット",
        "自分のクラス", "自分のデッキ", "ランク帯", "グレード",
        "相手クラス", "相手デッキ", "先攻/後攻", "勝敗",
    ]]
    for m in matches:
        raw_rows.append([
            m["recorded_at"],
            resolve_name(m["user_id"]),
            # 数値として解釈されて丸められるのを防ぐため文字列前置
            "'" + str(m["user_id"]),
            m["env_id"],
            format_labels.get(m["format"], m["format"]),
            m["my_class"],
            m.get("my_deck") or "",
            m.get("rank_tier") or "",
            m.get("cr_grade") or "",
            m["opp_class"],
            m.get("opp_deck") or "",
            "先攻" if m["is_first"] else "後攻",
            "勝ち" if m["is_win"] else "負け",
        ])

    # ---- 集計（ユーザー別） ----
    per_user: dict[str, dict] = {}
    for m in matches:
        row = per_user.setdefault(m["user_id"], {"total": 0, "wins": 0})
        row["total"] += 1
        row["wins"] += m["is_win"]

    agg_rows = [["ユーザー", "ユーザーID", "試合数", "勝数", "敗数", "勝率(%)"]]
    for uid, row in sorted(per_user.items(), key=lambda kv: -kv[1]["total"]):
        total = row["total"]
        wins = row["wins"]
        win_rate = round(wins / total * 100, 1) if total else 0.0
        agg_rows.append([
            resolve_name(uid), "'" + str(uid), total, wins, total - wins, win_rate,
        ])

    _write_sheet(sh, RAW_SHEET_NAME, raw_rows)
    _write_sheet(sh, AGG_SHEET_NAME, agg_rows)

    return {
        "url": sh.url,
        "raw": len(matches),
        "users": len(per_user),
    }
