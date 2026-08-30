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

【シート構成】
  生データ       … matches の全行（全環境・全期間）
  集計           … ユーザー別の試合数・勝率（全環境・全期間）
  クラス別集計   … 現環境のみ。クラスごとの試合数・シェア・勝率
  対面マトリクス … 現環境のみ。自分クラス×相手クラスの勝率（母数が少ないマスは空欄）
  先攻後攻集計   … 現環境のみ

  クラス別集計・対面マトリクス・先攻後攻集計は、初回作成時のみグラフ／
  条件付き書式（ヒートマップ）を付ける。以降の同期は値の上書きだけなので、
  グラフ・書式はそのまま参照範囲の更新後の値を反映し続ける
  （worksheet.clear() は値だけを消し、グラフや条件付き書式は消えないため）。
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
CLASS_SHEET_NAME = "クラス別集計"
MATCHUP_SHEET_NAME = "対面マトリクス"
FIRST_SECOND_SHEET_NAME = "先攻後攻集計"

# ヒートマップ・グラフの色（Discordのグラフ機能と揃えたRdYlGn系）
COLOR_RED = {"red": 0.84, "green": 0.19, "blue": 0.15}
COLOR_YELLOW = {"red": 1.0, "green": 0.93, "blue": 0.55}
COLOR_GREEN = {"red": 0.20, "green": 0.60, "blue": 0.29}


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
    """シートを作り直して全行を書き込む。(worksheet, is_new) を返す。

    is_new は「今回の呼び出しでシートを新規作成したか」。グラフ／条件付き書式は
    新規作成時にしか付けないので、呼び出し側はこれを見て判断する。
    """
    is_new = False
    try:
        ws = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=title, rows=max(len(rows) + 10, 100), cols=max(len(rows[0]) if rows else 1, 12)
        )
        is_new = True

    ws.clear()
    if rows:
        # update() は gspread 5.x と 6.x で引数の順序が違うため、
        # 仕様が安定している append_rows を使う（clear 直後なので1行目から入る）
        ws.append_rows(rows, value_input_option="RAW")
    return ws, is_new


def _safe(label: str, warnings: list, fn, *args, **kwargs):
    """グラフ／条件付き書式などの追加処理を個別に保護する。

    ここで失敗しても、他のシートの同期やこの後の処理は止めない。
    失敗内容は warnings に積んで呼び出し元へ返す（診断コマンドで確認できるように）。
    """
    try:
        return fn(*args, **kwargs)
    except Exception as e:
        warnings.append(f"{label}: {type(e).__name__}: {e}")
        return None


# ==============================
# 生データ・ユーザー別集計（既存・全環境・全期間）
# ==============================

def _build_raw_rows(matches: list[dict], format_labels: dict, resolve_name) -> list[list]:
    rows = [[
        "記録日時", "ユーザー", "ユーザーID", "環境ID", "フォーマット",
        "自分のクラス", "自分のデッキ", "ランク帯", "グレード",
        "相手クラス", "相手デッキ", "先攻/後攻", "勝敗",
    ]]
    for m in matches:
        rows.append([
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
    return rows


def _build_agg_rows(matches: list[dict], resolve_name) -> list[list]:
    per_user: dict[str, dict] = {}
    for m in matches:
        row = per_user.setdefault(m["user_id"], {"total": 0, "wins": 0})
        row["total"] += 1
        row["wins"] += m["is_win"]

    rows = [["ユーザー", "ユーザーID", "試合数", "勝数", "敗数", "勝率(%)"]]
    for uid, row in sorted(per_user.items(), key=lambda kv: -kv[1]["total"]):
        total = row["total"]
        wins = row["wins"]
        win_rate = round(wins / total * 100, 1) if total else 0.0
        rows.append([resolve_name(uid), "'" + str(uid), total, wins, total - wins, win_rate])
    return rows


# ==============================
# クラス別集計（現環境のみ）
# ==============================

def _build_class_summary_rows(class_summary: list[dict], env_id: str) -> list[list]:
    """class_summary: [{"my_class":..., "total":..., "wins":...}, ...]（get_global_summaryのby_my_class）"""
    grand_total = sum(r["total"] for r in class_summary) or 1
    rows = [[f"クラス（現環境：{env_id}）", "試合数", "シェア(%)", "勝数", "敗数", "勝率(%)"]]
    for r in class_summary:
        total = r["total"]
        wins = r["wins"] or 0
        share = round(total / grand_total * 100, 1)
        win_rate = round(wins / total * 100, 1) if total else 0.0
        rows.append([r["my_class"], total, share, wins, total - wins, win_rate])
    return rows


def _add_class_charts(spreadsheet, ws, n_rows: int):
    """クラス使用率（円）・クラス別勝率（横棒）を追加する。初回作成時のみ呼ぶ。"""
    sheet_id = ws.id
    end_row = 1 + n_rows  # ヘッダー行(0)を除いた実データの終端（0-index, 排他的）
    requests = [
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "クラス使用率（現環境）",
                        "pieChart": {
                            "legendPosition": "RIGHT_LEGEND",
                            "domain": {"sourceRange": {"sources": [{
                                "sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": end_row,
                                "startColumnIndex": 0, "endColumnIndex": 1,
                            }]}},
                            "series": {"sourceRange": {"sources": [{
                                "sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": end_row,
                                "startColumnIndex": 1, "endColumnIndex": 2,
                            }]}},
                        },
                    },
                    "position": {"overlayPosition": {"anchorCell": {
                        "sheetId": sheet_id, "rowIndex": 0, "columnIndex": 8,
                    }}},
                },
            },
        },
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "クラス別勝率（現環境）",
                        "basicChart": {
                            "chartType": "BAR",
                            "legendPosition": "NO_LEGEND",
                            "axis": [{"position": "BOTTOM_AXIS", "title": "勝率(%)"}],
                            "domains": [{"domain": {"sourceRange": {"sources": [{
                                "sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": end_row,
                                "startColumnIndex": 0, "endColumnIndex": 1,
                            }]}}}],
                            "series": [{
                                "series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": end_row,
                                    "startColumnIndex": 5, "endColumnIndex": 6,
                                }]}},
                                "targetAxis": "BOTTOM_AXIS",
                            }],
                        },
                    },
                    "position": {"overlayPosition": {"anchorCell": {
                        "sheetId": sheet_id, "rowIndex": 20, "columnIndex": 8,
                    }}},
                },
            },
        },
    ]
    spreadsheet.batch_update({"requests": requests})


# ==============================
# 対面マトリクス（現環境のみ）
# ==============================

def _build_matchup_grids(matchup_rows: list[dict], classes: list[str], min_sample: int):
    """
    matchup_rows: [{"my_class":..., "opp_class":..., "total":..., "wins":...}, ...]

    戻り値: (シート全体の rows, 勝率マスの開始行index, 終了行index)
    勝率マトリクスの後ろに、参考として試合数マトリクスも続けて書く。
    母数が min_sample 未満のマスは、勝率側は空欄（信頼できる数字ではないため）、
    試合数側はそのまま件数を出す（何もないわけではないことが分かるように）。
    """
    lookup = {(r["my_class"], r["opp_class"]): r for r in matchup_rows}

    rate_header = ["自分＼相手"] + classes
    rate_rows = [rate_header]
    count_header = ["自分＼相手"] + classes
    count_rows = [count_header]

    for mc in classes:
        rate_row = [mc]
        count_row = [mc]
        for oc in classes:
            r = lookup.get((mc, oc))
            total = r["total"] if r else 0
            wins = r["wins"] if r else 0
            count_row.append(total)
            if total >= min_sample:
                rate_row.append(round(wins / total * 100, 1))
            else:
                rate_row.append("")  # 母数不足は空欄（条件付き書式が数値のみに反応するため未入力にする）
        rate_rows.append(rate_row)
        count_rows.append(count_row)

    rate_start = 0
    rate_end = len(rate_rows)  # ヘッダー込みの行数
    blank_row = [""] * len(rate_header)
    note_row = [f"（参考）対戦数　※上の勝率マスが空欄のところも含む　母数のしきい値：{min_sample}戦"]

    all_rows = rate_rows + [blank_row, note_row] + count_rows
    return all_rows, rate_start, rate_end


def _add_matchup_heatmap(spreadsheet, ws, n_classes: int, rate_end_row: int):
    """勝率マトリクスのマス（ヘッダー行・見出し列を除く）に色スケールを付ける。初回作成時のみ呼ぶ。"""
    sheet_id = ws.id
    requests = [{
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": sheet_id,
                    "startRowIndex": 1, "endRowIndex": rate_end_row,
                    "startColumnIndex": 1, "endColumnIndex": 1 + n_classes,
                }],
                "gradientRule": {
                    "minpoint": {"type": "NUMBER", "value": "0", "color": COLOR_RED},
                    "midpoint": {"type": "NUMBER", "value": "50", "color": COLOR_YELLOW},
                    "maxpoint": {"type": "NUMBER", "value": "100", "color": COLOR_GREEN},
                },
            },
            "index": 0,
        },
    }]
    spreadsheet.batch_update({"requests": requests})


# ==============================
# 先攻後攻集計（現環境のみ）
# ==============================

def _build_first_second_rows(first_second_summary: list[dict], env_id: str) -> list[list]:
    """first_second_summary: [{"is_first":0/1, "total":..., "wins":...}, ...]（get_global_summaryのby_first）"""
    rows = [[f"区分（現環境：{env_id}）", "試合数", "勝数", "敗数", "勝率(%)"]]
    # 先攻→後攻の順で見せる（表示の慣習に合わせる）
    ordered = sorted(first_second_summary, key=lambda r: 0 if r["is_first"] else 1)
    for r in ordered:
        total = r["total"]
        wins = r["wins"] or 0
        win_rate = round(wins / total * 100, 1) if total else 0.0
        label = "先攻" if r["is_first"] else "後攻"
        rows.append([label, total, wins, total - wins, win_rate])
    return rows


def _add_first_second_chart(spreadsheet, ws, n_rows: int):
    sheet_id = ws.id
    end_row = 1 + n_rows
    requests = [{
        "addChart": {
            "chart": {
                "spec": {
                    "title": "先攻/後攻の勝率（現環境）",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "NO_LEGEND",
                        "axis": [{"position": "LEFT_AXIS", "title": "勝率(%)"}],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": end_row,
                            "startColumnIndex": 0, "endColumnIndex": 1,
                        }]}}}],
                        "series": [{
                            "series": {"sourceRange": {"sources": [{
                                "sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": end_row,
                                "startColumnIndex": 4, "endColumnIndex": 5,
                            }]}},
                            "targetAxis": "LEFT_AXIS",
                        }],
                    },
                },
                "position": {"overlayPosition": {"anchorCell": {
                    "sheetId": sheet_id, "rowIndex": 0, "columnIndex": 6,
                }}},
            },
        },
    }]
    spreadsheet.batch_update({"requests": requests})


# ==============================
# メイン
# ==============================

def sync_to_sheets(
    matches: list[dict], format_labels: dict, resolve_name,
    class_summary: list[dict] = None, first_second_summary: list[dict] = None,
    matchup_rows: list[dict] = None, classes: list[str] = None,
    min_sample: int = 30, env_id: str = "",
) -> dict:
    """
    戦績データをスプレッドシートへ全面上書きで同期する。

    matches               : matches テーブルの全行（dictのリスト、全環境・全期間）
    format_labels         : {"rotation": "ローテーション", ...}
    resolve_name          : user_id を表示名に変換する関数
    class_summary         : 現環境のクラス別集計（get_global_summaryのby_my_class）。省略時は集計シートを更新しない
    first_second_summary  : 現環境の先攻後攻別集計（get_global_summaryのby_first）。省略時は更新しない
    matchup_rows          : 現環境のクラス対クラス集計。省略時は更新しない
    classes               : クラスの一覧（CLASS_CHOICES）。matchup_rows使用時は必須
    min_sample            : 対面マトリクスで勝率を出す最低試合数
    env_id                : シートの見出しに出す現在の環境ID（表示用）

    戻り値: {"url":..., "raw":..., "users":..., "warnings": [...]}

    生データ・集計シートは毎回すべて消して書き直す方式（取消レコードがシートに
    残り続けるのを防ぐため）。クラス別集計・対面マトリクス・先攻後攻集計も同様に
    値は毎回全消去→再書き込みだが、グラフ・条件付き書式は初回作成時にしか
    付けない（worksheet.clear() は値だけを消すため、グラフ・書式は残る）。

    新しいグラフ／条件付き書式の追加（_safe 経由の処理）は個別に例外を握りつぶし、
    失敗しても生データ・集計など他の同期は止めない。失敗内容は戻り値の
    "warnings" に入るので、/戦績管理 シート同期 などで確認できる。
    """
    warnings: list = []
    client = _get_client()
    sheet_id_env = os.environ["SENSEKI_SPREADSHEET_ID"].strip()
    sh = client.open_by_key(sheet_id_env)

    # ---- 生データ・ユーザー別集計（既存・全環境） ----
    _write_sheet(sh, RAW_SHEET_NAME, _build_raw_rows(matches, format_labels, resolve_name))
    _, agg_is_new = _write_sheet(sh, AGG_SHEET_NAME, _build_agg_rows(matches, resolve_name))

    per_user_count = len({m["user_id"] for m in matches})

    # ---- クラス別集計（現環境） ----
    if class_summary is not None:
        rows = _build_class_summary_rows(class_summary, env_id)
        ws, is_new = _write_sheet(sh, CLASS_SHEET_NAME, rows)
        if is_new:
            _safe("クラス別集計グラフ", warnings, _add_class_charts, sh, ws, len(rows) - 1)

    # ---- 対面マトリクス（現環境） ----
    if matchup_rows is not None and classes:
        all_rows, rate_start, rate_end = _build_matchup_grids(matchup_rows, classes, min_sample)
        ws, is_new = _write_sheet(sh, MATCHUP_SHEET_NAME, all_rows)
        if is_new:
            _safe(
                "対面マトリクス ヒートマップ", warnings,
                _add_matchup_heatmap, sh, ws, len(classes), rate_end,
            )

    # ---- 先攻後攻集計（現環境） ----
    if first_second_summary is not None:
        rows = _build_first_second_rows(first_second_summary, env_id)
        ws, is_new = _write_sheet(sh, FIRST_SECOND_SHEET_NAME, rows)
        if is_new:
            _safe("先攻後攻集計グラフ", warnings, _add_first_second_chart, sh, ws, len(rows) - 1)

    return {
        "url": sh.url,
        "raw": len(matches),
        "users": per_user_count,
        "warnings": warnings,
    }
