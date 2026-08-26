"""
senseki.py
シャドバ戦績ツール（フェーズ1：記録できる状態）

詳細仕様は `戦績ツール_実装仕様書.md` を参照。
このファイルはフェーズ1のみを実装しています。

フェーズ1で実装する範囲
  1. SQLiteのテーブル作成（matches / user_settings / deck_names）
  2. /戦績設定（全項目必須）／ /デッキ登録・/デッキ切替・/ランク更新
  3. /戦績（ボタン形式のみ）
  4. /戦績取消
  5. /戦績確認（自分の勝率のみ。クラス別・対面別の内訳はフェーズ2）
  6. /戦績データ抽出（管理者専用）：生データ＋ユーザー別集計をExcel（xlsx）で出力
     毎週月曜9:00（JST）に現環境分を管理者DMへ自動送信もする
  7. /戦績全体（管理者専用）：その瞬間のDBから全体集計を出す
     クラス別・対面別・先後別。シートを見なくてもDiscord上で完結する
  8. /戦績シート同期・/戦績シート診断（管理者専用）：Googleスプレッドシートへ同期
     記録・取消があると SHEETS_DEBOUNCE_SECONDS 後にまとめて自動反映。
     加えて毎日9:30（JST）にも同期する（環境変数が未設定なら何もしない）
     詳細は senseki_sheets.py を参照
  9. /戦績板設置（管理者専用）：全員の使用デッキ＆ランクを1メッセージに
     常時表示する掲示板を作る。/戦績設定・/デッキ切替・/ランク更新の
     どれかが実行されるたびに自動で書き換わる（新規投稿ではなく編集）
  10. /弱点対面（メンバー限定）：自分の対面別勝率の低い相手を抽出し、
      登録済みの攻略記事があれば一緒に表示する（3戦以上・参考値扱い）
  11. /攻略記事登録・/攻略記事一覧・/攻略記事削除（管理者専用）：
      対面ごとのおすすめ記事URLを管理する

  12. /戦績パネル設置（管理者専用）：チャンネルに「戦績を記録する」ボタンを
      1つ常設する。押した人にだけ /戦績 と同じ入力画面が開く（永続ビュー、
      custom_id固定でBOT再起動後も動作し続ける）

【メンバー限定機能について】
  MEMBER_ROLE_IDS（カンマ区切りの環境変数）で判定する。飯テロBOTと同じ方式。
  未設定だと /弱点対面 は誰にも使えない（is_member が常にFalseを返すため）。

フェーズ2以降で追加するもの（このファイルには未実装）
  - 環境ID管理と /環境切替（現状は CURRENT_ENV_ID を固定値で使用）
  - デッキタイプのオートコンプリート・表記揺れ統合（/デッキ名整理）
  - /戦績比較（メンバー限定・全体勝率との比較）
  - ランク帯別の集計（データは記録済みなので後から出せる）

main.py への統合方法は末尾のコメントを参照。

【デプロイ前に必ず確認すること】
Railway の shirokusa-bot サービスに永続ボリューム（マウントパス /data）が
設定されているか確認してください。未設定のままデプロイすると、
再デプロイのたびに戦績データが消えます。
"""

import asyncio
import io
import os
import sqlite3
from datetime import datetime, time as dtime, timedelta, timezone

import discord
from discord import app_commands
from discord.ext import commands, tasks

try:
    from openpyxl import Workbook
    OPENPYXL_IMPORT_ERROR = None
except Exception as _e:  # ライブラリ未導入時に起動を止めない
    Workbook = None
    OPENPYXL_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"

try:
    import senseki_sheets
    SHEETS_MODULE_ERROR = None
except Exception as _e:  # モジュール未配置でも起動を止めない
    senseki_sheets = None
    SHEETS_MODULE_ERROR = f"{type(_e).__name__}: {_e}"

# ==============================
# 設定
# ==============================

GUILD_ID = 1194515135071539210

# Railway永続ボリューム（/data）に置く。飯テロBOTと同じ方式。
DB_PATH = os.environ.get("SENSEKI_DB_PATH", "/data/senseki.db")

JST = timezone(timedelta(hours=9))

# クラスの選択肢（仕様書 2-4）
CLASS_CHOICES = ["エルフ", "ロイヤル", "ウィッチ", "ドラゴン", "ナイトメア", "ビショップ", "ネメシス"]

# フォーマットの選択肢（仕様書 2-5）
FORMAT_LABELS = {"rotation": "ローテーション", "unlimited": "アンリミテッド"}

# ランク帯の選択肢
# ゲーム内では各帯が0〜3に細分化されているが（D0〜D3など）、
# 集計の粒度としては細かすぎるため帯単位でまとめている。
# マスターの上に「グランドマスター」があり、ここに到達すると
# クラス別レーティング（CR）による細分化に切り替わる。
RANK_TIER_CHOICES = [
    "ビギナー",
    "D帯",
    "C帯",
    "B帯",
    "A帯",
    "AA帯",
    "マスター",
    "グランドマスター",
]

# グランドマスター帯のみで使うCRグレード
# EPIC 1650〜1749 / ULTIMATE 1750〜1849 / LEGEND 1850以上 /
# BEYOND 1850以上かつCRランキング100位以内
# CRはクラス別に算出されるため、/戦績設定 のクラスと対応する値を入れる。
CR_GRADE_CHOICES = ["EPIC", "ULTIMATE", "LEGEND", "BEYOND"]

# 「グレードなし」を表す選択肢の値
# グランドマスター昇格直後のCR初期値は1100〜1600で、EPICの下限1650に届かない。
# つまりグレードなしは例外ではなく通常の状態なので、明示的に選べるようにする。
# 未入力での代用はしない（入力し忘れと区別できなくなるため）。
CR_GRADE_NONE_VALUE = "__none__"
CR_GRADE_NONE_LABEL = "グレードなし（CR1650未満）"

# CRグレードを入力できるランク帯
GRAND_MASTER_TIER = "グランドマスター"

# 環境ID（暫定固定値）
# フェーズ2で /環境切替 と data/environment.json による管理に置き換える。
# それまでは記録される全レコードがこの環境IDになる。
# 新弾リリース（2026-08-27）に合わせて設定。新弾切り替わり後に変更する場合は
# ここを書き換えてから再デプロイすること。
CURRENT_ENV_ID = "2026-08-beyond-2"

# 集計データの定期抽出先（管理者DM）
# しろくさのユーザーID（koken_rank.py の ADOPT_JUDGE_USER_ID と同じ値）
SENSEKI_ADMIN_USER_ID = 346475730458378240

# 定期抽出のタイミング（JST）
WEEKLY_EXPORT_WEEKDAY = 0   # 0=月曜
WEEKLY_EXPORT_HOUR = 9
WEEKLY_EXPORT_MINUTE = 0

# Googleスプレッドシートへの日次同期のタイミング（JST）
# 環境変数が未設定なら同期はスキップされる（senseki_sheets.py 参照）
SHEETS_SYNC_HOUR = 9
SHEETS_SYNC_MINUTE = 30

# 統計として扱える最低試合数（仕様書 8-3）
# フェーズ3の /戦績全体 ではこの値を下回る対面は数字を出さない
MIN_SAMPLE_FOR_STATS = 30

# 個人の弱点対面表示の最低試合数
# 30戦は運用初期には現実的でないため、個人向けは緩めにする。
# その代わり画面上に「参考値」であることを明記する。
MIN_SAMPLE_PERSONAL = 3

# メンバー限定機能の判定に使うロールID（カンマ区切り、環境変数）
# 飯テロBOTの MEMBER_ROLE_IDS と同じ方式
def _parse_role_ids(raw: str) -> set:
    ids = set()
    for part in (raw or "").split(","):
        part = part.strip()
        if part.isdigit():
            ids.add(int(part))
    return ids

MEMBER_ROLE_IDS = _parse_role_ids(os.environ.get("MEMBER_ROLE_IDS", ""))


def is_member(user: discord.Member) -> bool:
    if not MEMBER_ROLE_IDS:
        return False
    if not isinstance(user, discord.Member):
        return False
    return any(r.id in MEMBER_ROLE_IDS for r in user.roles)

# 記録・取消のあと、何秒待ってからシートへ反映するか
# 0にすると毎試合すぐ書きに行くが、連戦時にGoogle側のAPI制限に触れる。
# 待つ間に複数の記録が入れば1回の同期にまとめられる（まとめ書き）。
SHEETS_DEBOUNCE_SECONDS = 60

# データが変わったことを同期タスクに知らせるフラグ
# insert_match / delete_last_match から立てる
_sheets_dirty = False


def mark_dirty():
    """戦績データが変わったことを記録する"""
    global _sheets_dirty
    _sheets_dirty = True


def take_dirty() -> bool:
    """フラグを読んで下ろす"""
    global _sheets_dirty
    was = _sheets_dirty
    _sheets_dirty = False
    return was


# ==============================
# DB初期化・アクセス
# ==============================

def _connect() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _init_db_sync():
    conn = _connect()
    try:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS matches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                recorded_at TEXT NOT NULL,
                env_id TEXT NOT NULL,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                my_deck TEXT,
                rank_tier TEXT,
                cr_grade TEXT,
                opp_class TEXT NOT NULL,
                opp_deck TEXT,
                is_first INTEGER NOT NULL,
                is_win INTEGER NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS user_settings (
                user_id TEXT PRIMARY KEY,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                my_deck TEXT,
                rank_tier TEXT,
                cr_grade TEXT,
                updated_at TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS deck_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                deck_name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(user_id, format, my_class, deck_name)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS deck_names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                class_name TEXT NOT NULL,
                deck_name TEXT NOT NULL,
                use_count INTEGER NOT NULL DEFAULT 0,
                is_official INTEGER NOT NULL DEFAULT 0,
                env_id TEXT NOT NULL
            )
        """)
        # 既に旧スキーマでテーブルが作られている場合に備えた追加（初回は何もしない）
        for table in ("matches", "user_settings"):
            cols = {r["name"] for r in cur.execute(f"PRAGMA table_info({table})")}
            if "cr_grade" not in cols:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN cr_grade TEXT")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS guide_links (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                my_class TEXT NOT NULL DEFAULT '',
                opp_class TEXT NOT NULL,
                url TEXT NOT NULL,
                note TEXT,
                updated_at TEXT NOT NULL,
                UNIQUE(my_class, opp_class)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bot_state (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS idx_matches_user_env ON matches(user_id, env_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_matches_env_format ON matches(env_id, format)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_matches_env_classes ON matches(env_id, my_class, opp_class)")
        conn.commit()
    finally:
        conn.close()


async def init_db():
    await asyncio.to_thread(_init_db_sync)


def _get_user_settings_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM user_settings WHERE user_id = ?", (user_id,)
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


async def get_user_settings(user_id: str):
    return await asyncio.to_thread(_get_user_settings_sync, user_id)


def _upsert_user_settings_sync(user_id: str, format_: str, my_class: str, my_deck, rank_tier, cr_grade):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO user_settings (user_id, format, my_class, my_deck, rank_tier, cr_grade, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                format = excluded.format,
                my_class = excluded.my_class,
                my_deck = excluded.my_deck,
                rank_tier = excluded.rank_tier,
                cr_grade = excluded.cr_grade,
                updated_at = excluded.updated_at
        """, (user_id, format_, my_class, my_deck, rank_tier, cr_grade, now))
        conn.commit()
    finally:
        conn.close()


async def upsert_user_settings(user_id: str, format_: str, my_class: str, my_deck, rank_tier, cr_grade):
    await asyncio.to_thread(
        _upsert_user_settings_sync, user_id, format_, my_class, my_deck, rank_tier, cr_grade
    )


def _register_deck_name(conn, class_name: str, deck_name: str):
    """デッキ名プールに登録・使用件数を増やす（オートコンプリート候補用）"""
    if not deck_name:
        return
    row = conn.execute(
        "SELECT id, use_count FROM deck_names "
        "WHERE class_name = ? AND deck_name = ? AND env_id = ?",
        (class_name, deck_name, CURRENT_ENV_ID),
    ).fetchone()
    if row is None:
        conn.execute(
            "INSERT INTO deck_names (class_name, deck_name, use_count, is_official, env_id) "
            "VALUES (?, ?, 1, 0, ?)",
            (class_name, deck_name, CURRENT_ENV_ID),
        )
    else:
        conn.execute(
            "UPDATE deck_names SET use_count = use_count + 1 WHERE id = ?", (row["id"],)
        )


def _insert_match_sync(user_id: str, settings: dict, opp_class: str, opp_deck,
                       is_first: bool, is_win: bool):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO matches (
                user_id, recorded_at, env_id, format, my_class, my_deck, rank_tier, cr_grade,
                opp_class, opp_deck, is_first, is_win
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            user_id, now, CURRENT_ENV_ID, settings["format"], settings["my_class"],
            settings.get("my_deck"), settings.get("rank_tier"), settings.get("cr_grade"),
            opp_class, opp_deck, 1 if is_first else 0, 1 if is_win else 0,
        ))
        # 自分のデッキ・相手のデッキの両方を候補プールに貯める
        _register_deck_name(conn, settings["my_class"], settings.get("my_deck"))
        _register_deck_name(conn, opp_class, opp_deck)
        conn.commit()
    finally:
        conn.close()


async def insert_match(user_id: str, settings: dict, opp_class: str, opp_deck,
                       is_first: bool, is_win: bool):
    await asyncio.to_thread(
        _insert_match_sync, user_id, settings, opp_class, opp_deck, is_first, is_win
    )
    mark_dirty()


def _get_deck_names_sync(class_name: str, limit: int = 20):
    """そのクラスで使われているデッキ名を使用件数の多い順に返す"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT deck_name FROM deck_names
            WHERE class_name = ? AND env_id = ?
            ORDER BY is_official DESC, use_count DESC, deck_name
            LIMIT ?
        """, (class_name, CURRENT_ENV_ID, limit)).fetchall()
        return [r["deck_name"] for r in rows]
    finally:
        conn.close()


async def get_deck_names(class_name: str, limit: int = 20):
    return await asyncio.to_thread(_get_deck_names_sync, class_name, limit)


# ---- デッキテンプレート ----

def _add_template_sync(user_id: str, format_: str, my_class: str, deck_name: str) -> bool:
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        try:
            conn.execute(
                "INSERT INTO deck_templates (user_id, format, my_class, deck_name, created_at) "
                "VALUES (?, ?, ?, ?, ?)",
                (user_id, format_, my_class, deck_name, now),
            )
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False  # 同じ組み合わせが登録済み
    finally:
        conn.close()


async def add_template(user_id: str, format_: str, my_class: str, deck_name: str) -> bool:
    return await asyncio.to_thread(_add_template_sync, user_id, format_, my_class, deck_name)


def _list_templates_sync(user_id: str):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM deck_templates WHERE user_id = ? ORDER BY id", (user_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def list_templates(user_id: str):
    return await asyncio.to_thread(_list_templates_sync, user_id)


def _delete_template_sync(user_id: str, template_id: int) -> bool:
    conn = _connect()
    try:
        cur = conn.execute(
            "DELETE FROM deck_templates WHERE id = ? AND user_id = ?", (template_id, user_id)
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


async def delete_template(user_id: str, template_id: int) -> bool:
    return await asyncio.to_thread(_delete_template_sync, user_id, template_id)


def _apply_template_sync(user_id: str, template_id: int):
    """テンプレートを現在の設定に反映する。ランク帯・グレードは維持する"""
    conn = _connect()
    try:
        t = conn.execute(
            "SELECT * FROM deck_templates WHERE id = ? AND user_id = ?", (template_id, user_id)
        ).fetchone()
        if t is None:
            return None
        cur = conn.execute("SELECT * FROM user_settings WHERE user_id = ?", (user_id,)).fetchone()
        rank_tier = cur["rank_tier"] if cur else None
        cr_grade = cur["cr_grade"] if cur else None
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO user_settings (user_id, format, my_class, my_deck, rank_tier, cr_grade, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                format = excluded.format,
                my_class = excluded.my_class,
                my_deck = excluded.my_deck,
                updated_at = excluded.updated_at
        """, (user_id, t["format"], t["my_class"], t["deck_name"], rank_tier, cr_grade, now))
        conn.commit()
        return dict(t)
    finally:
        conn.close()


async def apply_template(user_id: str, template_id: int):
    return await asyncio.to_thread(_apply_template_sync, user_id, template_id)


def _update_rank_sync(user_id: str, rank_tier: str, cr_grade) -> bool:
    """ランク帯とグレードだけを更新する。デッキ情報は触らない"""
    conn = _connect()
    try:
        row = conn.execute("SELECT user_id FROM user_settings WHERE user_id = ?", (user_id,)).fetchone()
        if row is None:
            return False  # 先に /戦績設定 が必要
        conn.execute(
            "UPDATE user_settings SET rank_tier = ?, cr_grade = ?, updated_at = ? WHERE user_id = ?",
            (rank_tier, cr_grade, datetime.now(JST).isoformat(), user_id),
        )
        conn.commit()
        return True
    finally:
        conn.close()


async def update_rank(user_id: str, rank_tier: str, cr_grade) -> bool:
    return await asyncio.to_thread(_update_rank_sync, user_id, rank_tier, cr_grade)


def _delete_last_match_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT id, opp_class, is_win, recorded_at FROM matches "
            "WHERE user_id = ? ORDER BY id DESC LIMIT 1",
            (user_id,),
        ).fetchone()
        if row is None:
            return None
        conn.execute("DELETE FROM matches WHERE id = ?", (row["id"],))
        conn.commit()
        return dict(row)
    finally:
        conn.close()


async def delete_last_match(user_id: str):
    result = await asyncio.to_thread(_delete_last_match_sync, user_id)
    if result is not None:
        mark_dirty()
    return result


def _get_summary_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute("""
            SELECT
                COUNT(*) AS total,
                SUM(is_win) AS wins
            FROM matches
            WHERE user_id = ? AND env_id = ?
        """, (user_id, CURRENT_ENV_ID)).fetchone()
        total = row["total"] or 0
        wins = row["wins"] or 0
        return {"total": total, "wins": wins, "losses": total - wins}
    finally:
        conn.close()


async def get_summary(user_id: str):
    return await asyncio.to_thread(_get_summary_sync, user_id)


def _fetch_all_matches_sync(env_only: bool):
    conn = _connect()
    try:
        if env_only:
            rows = conn.execute(
                "SELECT * FROM matches WHERE env_id = ? ORDER BY id", (CURRENT_ENV_ID,)
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM matches ORDER BY id").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def fetch_all_matches(env_only: bool = True):
    return await asyncio.to_thread(_fetch_all_matches_sync, env_only)


def _fetch_all_user_settings_sync():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM user_settings").fetchall()
        return {r["user_id"]: dict(r) for r in rows}
    finally:
        conn.close()


async def fetch_all_user_settings():
    return await asyncio.to_thread(_fetch_all_user_settings_sync)


# ---- BOTの小さな永続状態（掲示板のチャンネルID・メッセージIDなど） ----

def _get_state_sync(key: str):
    conn = _connect()
    try:
        row = conn.execute("SELECT value FROM bot_state WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else None
    finally:
        conn.close()


async def get_state(key: str):
    return await asyncio.to_thread(_get_state_sync, key)


def _set_state_sync(key: str, value: str):
    conn = _connect()
    try:
        conn.execute(
            "INSERT INTO bot_state (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
        conn.commit()
    finally:
        conn.close()


async def set_state(key: str, value: str):
    await asyncio.to_thread(_set_state_sync, key, value)


# ---- 攻略記事リンク ----
# my_class は空文字列("")で「自分のクラスを問わない汎用記事」を表す

def _upsert_guide_link_sync(my_class: str, opp_class: str, url: str, note):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO guide_links (my_class, opp_class, url, note, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(my_class, opp_class) DO UPDATE SET
                url = excluded.url, note = excluded.note, updated_at = excluded.updated_at
        """, (my_class, opp_class, url, note, now))
        conn.commit()
    finally:
        conn.close()


async def upsert_guide_link(my_class: str, opp_class: str, url: str, note=None):
    await asyncio.to_thread(_upsert_guide_link_sync, my_class or "", opp_class, url, note)


def _delete_guide_link_sync(my_class: str, opp_class: str) -> bool:
    conn = _connect()
    try:
        cur = conn.execute(
            "DELETE FROM guide_links WHERE my_class = ? AND opp_class = ?", (my_class, opp_class)
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


async def delete_guide_link(my_class: str, opp_class: str) -> bool:
    return await asyncio.to_thread(_delete_guide_link_sync, my_class or "", opp_class)


def _list_guide_links_sync():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM guide_links ORDER BY opp_class, my_class").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def list_guide_links():
    return await asyncio.to_thread(_list_guide_links_sync)


def _find_guide_link_sync(my_class: str, opp_class: str):
    conn = _connect()
    try:
        # 1. クラス指定の記事を優先、2. なければ汎用（my_class=""）にフォールバック
        row = conn.execute(
            "SELECT * FROM guide_links WHERE my_class = ? AND opp_class = ?",
            (my_class, opp_class),
        ).fetchone()
        if row is None:
            row = conn.execute(
                "SELECT * FROM guide_links WHERE my_class = '' AND opp_class = ?",
                (opp_class,),
            ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


async def find_guide_link(my_class: str, opp_class: str):
    return await asyncio.to_thread(_find_guide_link_sync, my_class, opp_class)


# ---- 個人の対面別集計（弱点対面の抽出用） ----

def _get_personal_matchups_sync(user_id: str):
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT my_class, opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE user_id = ? AND env_id = ?
            GROUP BY my_class, opp_class
        """, (user_id, CURRENT_ENV_ID)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_personal_matchups(user_id: str):
    return await asyncio.to_thread(_get_personal_matchups_sync, user_id)


def _get_global_summary_sync():
    """現環境の全体集計。クラス別・先後別まで一度に出す"""
    conn = _connect()
    try:
        total_row = conn.execute("""
            SELECT COUNT(*) AS total, SUM(is_win) AS wins,
                   COUNT(DISTINCT user_id) AS users
            FROM matches WHERE env_id = ?
        """, (CURRENT_ENV_ID,)).fetchone()

        by_first = conn.execute("""
            SELECT is_first, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ? GROUP BY is_first
        """, (CURRENT_ENV_ID,)).fetchall()

        by_my_class = conn.execute("""
            SELECT my_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ?
            GROUP BY my_class ORDER BY total DESC
        """, (CURRENT_ENV_ID,)).fetchall()

        by_opp_class = conn.execute("""
            SELECT opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ?
            GROUP BY opp_class ORDER BY total DESC
        """, (CURRENT_ENV_ID,)).fetchall()

        return {
            "total": total_row["total"] or 0,
            "wins": total_row["wins"] or 0,
            "users": total_row["users"] or 0,
            "by_first": [dict(r) for r in by_first],
            "by_my_class": [dict(r) for r in by_my_class],
            "by_opp_class": [dict(r) for r in by_opp_class],
        }
    finally:
        conn.close()


async def get_global_summary():
    return await asyncio.to_thread(_get_global_summary_sync)


# ==============================
# /戦績 のボタン形式フロー
# ==============================

class OpponentClassSelect(discord.ui.Select):
    def __init__(self):
        options = [discord.SelectOption(label=c, value=c) for c in CLASS_CHOICES]
        super().__init__(placeholder="相手クラスを選択", options=options, min_values=1, max_values=1)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.opp_class = self.values[0]
        # そのクラスで既に使われているデッキ名を候補に出す
        known = await get_deck_names(view.opp_class)
        view.show_opponent_deck(known)
        await interaction.response.edit_message(
            content=view.screen(
                "相手のデッキが分かれば選んでください。分からなければ「分からない」でOKです。"
            ),
            view=view,
        )


class OpponentDeckSelect(discord.ui.Select):
    """既知のデッキ名から選ぶ。候補になければモーダルで新規入力"""

    SKIP_VALUE = "__skip__"
    NEW_VALUE = "__new__"

    def __init__(self, known_decks: list[str]):
        options = [
            discord.SelectOption(label="分からない / 入力しない", value=self.SKIP_VALUE),
            discord.SelectOption(label="新しく入力する…", value=self.NEW_VALUE),
        ]
        # Discordのセレクトは25件が上限。既定の2件を除いた分だけ載せる
        for name in known_decks[:23]:
            options.append(discord.SelectOption(label=name[:100], value=name[:100]))
        super().__init__(
            placeholder="相手のデッキ（任意）", options=options, min_values=1, max_values=1
        )

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        value = self.values[0]

        if value == self.NEW_VALUE:
            await interaction.response.send_modal(OpponentDeckModal(view))
            return

        view.opp_deck = None if value == self.SKIP_VALUE else value
        view.show_first_second()
        await interaction.response.edit_message(
            content=view.screen("先攻／後攻を選んでください。"),
            view=view,
        )


class OpponentDeckModal(discord.ui.Modal, title="相手のデッキ名を入力"):
    deck_name = discord.ui.TextInput(
        label="デッキ名",
        placeholder="例：進化エルフ",
        required=True,
        max_length=50,
    )

    def __init__(self, flow_view: "SensekiFlowView"):
        super().__init__()
        self.flow_view = flow_view

    async def on_submit(self, interaction: discord.Interaction):
        view = self.flow_view
        view.opp_deck = str(self.deck_name.value).strip() or None
        view.show_first_second()
        await interaction.response.edit_message(
            content=view.screen("先攻／後攻を選んでください。"),
            view=view,
        )


class FirstButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="先攻", style=discord.ButtonStyle.primary)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.is_first = True
        view.show_win_lose()
        await interaction.response.edit_message(
            content=view.screen("結果を選んでください。"),
            view=view,
        )


class SecondButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="後攻", style=discord.ButtonStyle.secondary)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.is_first = False
        view.show_win_lose()
        await interaction.response.edit_message(
            content=view.screen("結果を選んでください。"),
            view=view,
        )


class WinButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="勝ち", style=discord.ButtonStyle.success)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        await view.finalize(interaction, is_win=True)


class LoseButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="負け", style=discord.ButtonStyle.danger)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        await view.finalize(interaction, is_win=False)


class RepeatButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="同じ設定でもう1試合", style=discord.ButtonStyle.primary)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.reset()
        await interaction.response.edit_message(
            content=view.screen("相手クラスを選択してください。"),
            view=view,
        )


class SensekiFlowView(discord.ui.View):
    def __init__(self, user_id: str, settings: dict):
        super().__init__(timeout=300)
        self.user_id = user_id
        self.settings = settings
        self.opp_class = None
        self.opp_deck = None
        self.is_first = None
        self.clear_items()
        self.add_item(OpponentClassSelect())

    def describe_settings(self) -> str:
        """今の自分の設定。全ステップで出し続けて、古いまま記録するのを防ぐ"""
        s = self.settings
        fmt = FORMAT_LABELS.get(s["format"], s["format"])
        deck = s.get("my_deck") or "デッキ未設定"
        rank = s.get("rank_tier") or "ランク未設定"
        if s.get("cr_grade"):
            rank += f"・{s['cr_grade']}"
        elif s.get("rank_tier") == GRAND_MASTER_TIER:
            rank += "・グレードなし"
        return f"🧑 あなた：**{deck}**（{s['my_class']} / {fmt}） / ランク：**{rank}**"

    def screen(self, body: str) -> str:
        """設定ヘッダー＋進捗＋操作案内をまとめた画面テキスト"""
        parts = [self.describe_settings()]
        progress = self.describe_progress()
        if progress:
            parts.append(progress)
        parts.append("")
        parts.append(body)
        return "\n".join(parts)

    def describe_progress(self) -> str:
        """ここまでに選んだ内容を1行で"""
        if self.opp_class is None:
            return ""
        label = f"🆚 相手：**{self.opp_class}**"
        if self.opp_deck:
            label += f"（{self.opp_deck}）"
        if self.is_first is not None:
            label += f" / **{'先攻' if self.is_first else '後攻'}**"
        return label

    def show_opponent_deck(self, known_decks: list[str]):
        self.clear_items()
        self.add_item(OpponentDeckSelect(known_decks))

    def show_first_second(self):
        self.clear_items()
        self.add_item(FirstButton())
        self.add_item(SecondButton())

    def show_win_lose(self):
        self.clear_items()
        self.add_item(WinButton())
        self.add_item(LoseButton())

    def reset(self):
        self.opp_class = None
        self.opp_deck = None
        self.is_first = None
        self.clear_items()
        self.add_item(OpponentClassSelect())

    async def finalize(self, interaction: discord.Interaction, is_win: bool):
        await insert_match(
            self.user_id, self.settings, self.opp_class, self.opp_deck,
            self.is_first, is_win,
        )
        result_label = "勝ち" if is_win else "負け"
        self.clear_items()
        self.add_item(RepeatButton())
        await interaction.response.edit_message(
            content=self.screen(
                f"✅ **{result_label}** で記録しました。\n"
                f"続けて記録する場合は下のボタンを押してください。"
            ),
            view=self,
        )

    async def on_timeout(self):
        # メッセージ編集はできないため何もしない（ephemeralなので放置で問題なし）
        pass


def _resolve_grade(rank_value: str, raw_grade):
    """
    ランク帯とグレードの整合を取る。
    戻り値: (保存するグレード, 補足メッセージ, エラーメッセージ)

    グランドマスターはグレード指定を必須にするが、「グレードなし」も
    正規の選択肢として認める（昇格直後はCRが1650に届かないため）。
    """
    if rank_value == GRAND_MASTER_TIER:
        if raw_grade is None:
            return None, None, (
                f"{GRAND_MASTER_TIER}を選んだ場合は `グレード` も指定してください。\n"
                f"CRが1650未満でグレードが付いていない場合は"
                f"「{CR_GRADE_NONE_LABEL}」を選んでください。"
            )
        if raw_grade == CR_GRADE_NONE_VALUE:
            return None, None, None
        return raw_grade, None, None

    # グランドマスター以外にCRは存在しない
    if raw_grade is not None and raw_grade != CR_GRADE_NONE_VALUE:
        return None, (
            f"※グレードは{GRAND_MASTER_TIER}帯でのみ記録されるため、今回は保存していません。"
        ), None
    return None, None, None


# ==============================
# /デッキ切替 のUI
# ==============================

def _template_label(t: dict) -> str:
    fmt = "ロテ" if t["format"] == "rotation" else "アンリミ"
    return f"{t['deck_name']}（{t['my_class']} / {fmt}）"


class TemplateSelect(discord.ui.Select):
    def __init__(self, templates: list[dict], mode: str):
        self.mode = mode  # "switch" or "delete"
        options = [
            discord.SelectOption(label=_template_label(t)[:100], value=str(t["id"]))
            for t in templates[:25]
        ]
        placeholder = "使うデッキを選択" if mode == "switch" else "削除するデッキを選択"
        super().__init__(placeholder=placeholder, options=options, min_values=1, max_values=1)

    async def callback(self, interaction: discord.Interaction):
        user_id = str(interaction.user.id)
        template_id = int(self.values[0])

        if self.mode == "delete":
            ok = await delete_template(user_id, template_id)
            msg = "🗑️ 削除しました。" if ok else "削除できませんでした。"
            await interaction.response.edit_message(content=msg, view=None)
            return

        t = await apply_template(user_id, template_id)
        if t is None:
            await interaction.response.edit_message(
                content="そのデッキが見つかりませんでした。", view=None
            )
            return
        settings = await get_user_settings(user_id)
        rank = settings.get("rank_tier") or "未設定"
        if settings.get("cr_grade"):
            rank += f"（{settings['cr_grade']}）"
        await interaction.response.edit_message(
            content=(
                f"✅ デッキを切り替えました\n"
                f"{_template_label(t)}\n"
                f"ランク帯：{rank}（変更していません）"
            ),
            view=None,
        )
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()


class TemplateView(discord.ui.View):
    def __init__(self, templates: list[dict], mode: str = "switch"):
        super().__init__(timeout=180)
        self.add_item(TemplateSelect(templates, mode))


# ==============================
# 常設パネル（/戦績パネル設置 で1回置く。ボタンは押した人にだけ反応する）
# ==============================

PANEL_BUTTON_CUSTOM_ID = "senseki_panel_record_v1"


class SensekiPanelButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="戦績を記録する",
            style=discord.ButtonStyle.primary,
            emoji="⚔️",
            custom_id=PANEL_BUTTON_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        cog = interaction.client.get_cog("SensekiCog")
        if cog is None:
            await interaction.response.send_message(
                "内部エラー：SensekiCogが見つかりません。", ephemeral=True
            )
            return
        # ここでのレスポンスは押した本人にだけ見える（ephemeral）。
        # パネル本体（誰でも見えるメッセージ）は書き換えない。
        await cog._start_record_flow(interaction)


class SensekiPanelView(discord.ui.View):
    """timeout=None＋固定custom_idで永続化。BOT再起動後もボタンが反応し続ける"""

    def __init__(self):
        super().__init__(timeout=None)
        self.add_item(SensekiPanelButton())


# ==============================
# Excel（生データ＋集計）の生成
# ==============================

def _resolve_display_name(guild, user_id: str) -> str:
    if guild is not None:
        member = guild.get_member(int(user_id))
        if member is not None:
            return member.display_name
    return f"不明なユーザー（{user_id}）"


def build_workbook(matches: list[dict], guild) -> bytes:
    """生データシートと集計シートを持つxlsxをバイト列で返す"""
    wb = Workbook()

    # ---- 生データ ----
    ws_raw = wb.active
    ws_raw.title = "生データ"
    ws_raw.append([
        "記録日時", "ユーザー", "ユーザーID", "環境ID", "フォーマット",
        "自分のクラス", "自分のデッキ", "ランク帯", "グレード",
        "相手クラス", "相手デッキ", "先攻/後攻", "勝敗",
    ])
    for m in matches:
        ws_raw.append([
            m["recorded_at"],
            _resolve_display_name(guild, m["user_id"]),
            m["user_id"],
            m["env_id"],
            FORMAT_LABELS.get(m["format"], m["format"]),
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
    ws_agg = wb.create_sheet("集計")
    ws_agg.append(["ユーザー", "ユーザーID", "試合数", "勝数", "敗数", "勝率(%)"])

    per_user: dict[str, dict] = {}
    for m in matches:
        uid = m["user_id"]
        row = per_user.setdefault(uid, {"total": 0, "wins": 0})
        row["total"] += 1
        row["wins"] += m["is_win"]

    for uid, row in sorted(per_user.items(), key=lambda kv: -kv[1]["total"]):
        total = row["total"]
        wins = row["wins"]
        losses = total - wins
        win_rate = round(wins / total * 100, 1) if total else 0.0
        ws_agg.append([
            _resolve_display_name(guild, uid), uid, total, wins, losses, win_rate,
        ])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================
# Cog本体
# ==============================

class SensekiCog(commands.Cog):
    def __init__(self, bot: commands.Bot):
        self.bot = bot

    async def cog_load(self):
        await init_db()
        if not self.weekly_export.is_running():
            self.weekly_export.start()
        if not self.daily_sheets_sync.is_running():
            self.daily_sheets_sync.start()
        if not self.debounced_sheets_sync.is_running():
            self.debounced_sheets_sync.start()
        # 永続ビューの登録。custom_idが一致すれば、BOT再起動後も
        # 過去に送信済みのパネルのボタンが反応し続ける
        self.bot.add_view(SensekiPanelView())
        # BOT再起動をまたいでも掲示板が最新になるよう一度更新しておく
        try:
            await self._update_board()
        except Exception as e:
            print(f"⚠️ 起動時の戦績掲示板更新に失敗しました: {type(e).__name__}: {e}")

    def cog_unload(self):
        self.weekly_export.cancel()
        self.daily_sheets_sync.cancel()
        self.debounced_sheets_sync.cancel()

    @app_commands.command(
        name="戦績パネル設置",
        description="このチャンネルに「戦績を記録する」ボタンを常設します（管理者専用）",
    )
    @app_commands.default_permissions(administrator=True)
    async def senseki_panel_setup(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True, thinking=True)
        message = await interaction.channel.send(
            content=(
                "⚔️ **戦績記録パネル**\n"
                "下のボタンから対戦結果を記録できます。押した人にだけ入力画面が表示されます。\n"
                "-# 初回は `/戦績設定` で先にデッキ・ランクを登録してください。"
            ),
            view=SensekiPanelView(),
        )
        try:
            await message.pin()
        except discord.Forbidden:
            await interaction.followup.send(
                "パネルは作成しましたが、ピン留めの権限がなかったため手動で留めてください。",
                ephemeral=True,
            )
        await interaction.followup.send("✅ パネルを設置しました。", ephemeral=True)

    # ---- 現状掲示板（全プレイヤーのデッキ＆ランク一覧） ----

    BOARD_CHANNEL_KEY = "board_channel_id"
    BOARD_MESSAGE_KEY = "board_message_id"

    async def _build_board_embed(self) -> discord.Embed:
        all_settings = await fetch_all_user_settings()
        guild = self.bot.get_guild(GUILD_ID)
        resolve = self._name_resolver()

        by_class: dict[str, list[str]] = {}
        for uid, s in all_settings.items():
            deck = s.get("my_deck") or "デッキ未設定"
            rank = s.get("rank_tier") or "ランク未設定"
            if s.get("cr_grade"):
                rank += f"・{s['cr_grade']}"
            elif s.get("rank_tier") == GRAND_MASTER_TIER:
                rank += "・グレードなし"
            line = f"・{resolve(uid)}：{deck}（{rank}）"
            by_class.setdefault(s["my_class"], []).append(line)

        embed = discord.Embed(
            title="📋 現在の使用デッキ＆ランク",
            description="`/戦績設定` `/デッキ切替` `/ランク更新` のいずれかを実行すると自動で更新されます。",
            color=0x2E6DA4,
        )
        if not by_class:
            embed.add_field(name="―", value="まだ誰も設定していません。", inline=False)
        else:
            for cls in CLASS_CHOICES:
                if cls in by_class:
                    value = "\n".join(by_class[cls])[:1024]
                    embed.add_field(name=f"【{cls}】", value=value, inline=False)
        embed.timestamp = datetime.now(JST)
        return embed

    async def _update_board(self):
        """設定変更のたびに呼ぶ。掲示板が未設置なら何もしない"""
        channel_id = await get_state(self.BOARD_CHANNEL_KEY)
        message_id = await get_state(self.BOARD_MESSAGE_KEY)
        if not channel_id or not message_id:
            return

        channel = self.bot.get_channel(int(channel_id))
        if channel is None:
            print(f"⚠️ 戦績掲示板のチャンネルが見つかりません（ID: {channel_id}）")
            return

        try:
            message = await channel.fetch_message(int(message_id))
            embed = await self._build_board_embed()
            await message.edit(embed=embed)
        except discord.NotFound:
            print("⚠️ 戦績掲示板のメッセージが見つかりません。`/戦績板設置` をやり直してください。")
        except discord.Forbidden:
            print("⚠️ 戦績掲示板を編集する権限がありません。")
        except Exception as e:
            print(f"⚠️ 戦績掲示板の更新に失敗しました: {type(e).__name__}: {e}")

    @app_commands.command(
        name="戦績板設置",
        description="このチャンネルに、全員の使用デッキ＆ランクを常時表示する掲示板を作ります（管理者専用）",
    )
    @app_commands.default_permissions(administrator=True)
    async def senseki_board_setup(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True, thinking=True)
        embed = await self._build_board_embed()
        message = await interaction.channel.send(embed=embed)
        try:
            await message.pin()
        except discord.Forbidden:
            await interaction.followup.send(
                "掲示板は作成しましたが、ピン留めの権限がなかったため手動で留めてください。",
                ephemeral=True,
            )
        await set_state(self.BOARD_CHANNEL_KEY, str(interaction.channel.id))
        await set_state(self.BOARD_MESSAGE_KEY, str(message.id))
        await interaction.followup.send("✅ 掲示板を設置しました。", ephemeral=True)

    # ---- Googleスプレッドシート同期 ----

    def _name_resolver(self):
        guild = self.bot.get_guild(GUILD_ID)

        def resolve(user_id: str) -> str:
            if guild is not None:
                member = guild.get_member(int(user_id))
                if member is not None:
                    return member.display_name
            return f"不明なユーザー（{user_id}）"

        return resolve

    async def _run_sheets_sync(self) -> dict:
        matches = await fetch_all_matches(env_only=False)
        return await asyncio.to_thread(
            senseki_sheets.sync_to_sheets, matches, FORMAT_LABELS, self._name_resolver()
        )

    @tasks.loop(time=dtime(hour=SHEETS_SYNC_HOUR, minute=SHEETS_SYNC_MINUTE, tzinfo=JST))
    async def daily_sheets_sync(self):
        if senseki_sheets is None or not senseki_sheets.is_configured():
            return  # 未設定なら静かにスキップ（毎日ログを汚さない）
        try:
            result = await self._run_sheets_sync()
            print(f"✅ 戦績スプレッドシート同期完了: {result['raw']}件 / {result['users']}名")
        except Exception as e:
            print(f"⚠️ 戦績スプレッドシート同期に失敗しました: {type(e).__name__}: {e}")

    @daily_sheets_sync.before_loop
    async def _before_daily_sheets_sync(self):
        await self.bot.wait_until_ready()

    @tasks.loop(seconds=SHEETS_DEBOUNCE_SECONDS)
    async def debounced_sheets_sync(self):
        """記録・取消があった時だけシートへ反映する（まとめ書き）"""
        if senseki_sheets is None or not senseki_sheets.is_configured():
            return
        if not take_dirty():
            return  # 変更がなければAPIを呼ばない
        try:
            await self._run_sheets_sync()
        except Exception as e:
            print(f"⚠️ 戦績スプレッドシートの自動同期に失敗しました: {type(e).__name__}: {e}")
            mark_dirty()  # 次の周回で再試行する

    @debounced_sheets_sync.before_loop
    async def _before_debounced_sheets_sync(self):
        await self.bot.wait_until_ready()

    @app_commands.command(
        name="戦績シート同期",
        description="戦績データをGoogleスプレッドシートへ即時同期します（管理者専用）",
    )
    @app_commands.default_permissions(administrator=True)
    async def senseki_sheets_sync(self, interaction: discord.Interaction):
        if senseki_sheets is None:
            await interaction.response.send_message(
                f"senseki_sheets.py を読み込めませんでした。\n```\n{SHEETS_MODULE_ERROR}\n```",
                ephemeral=True,
            )
            return

        if not senseki_sheets.is_configured():
            await interaction.response.send_message(
                "スプレッドシート連携の設定が未完了です。\n"
                f"```\n{senseki_sheets.describe_config()}\n```\n"
                "Railway の Variables と requirements.txt を確認してください。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            result = await self._run_sheets_sync()
            await interaction.followup.send(
                f"✅ 同期しました（{result['raw']}件 / {result['users']}名）\n{result['url']}",
                ephemeral=True,
            )
        except Exception as e:
            print(f"⚠️ 戦績シート同期でエラー: {type(e).__name__}: {e}")
            hint = ""
            if "403" in str(e) or "PERMISSION" in str(e).upper():
                hint = (
                    "\n\nスプレッドシートがサービスアカウントに共有されていない可能性があります。"
                    "`/戦績シート診断` でメールアドレスを確認し、そのアドレスに"
                    "「編集者」権限で共有してください。"
                )
            await interaction.followup.send(
                f"同期に失敗しました。\n```\n{type(e).__name__}: {e}\n```{hint}",
                ephemeral=True,
            )

    @app_commands.command(
        name="戦績シート診断",
        description="スプレッドシート連携の設定状況を確認します（管理者専用）",
    )
    @app_commands.default_permissions(administrator=True)
    async def senseki_sheets_diag(self, interaction: discord.Interaction):
        if senseki_sheets is None:
            await interaction.response.send_message(
                f"senseki_sheets.py を読み込めませんでした。\n```\n{SHEETS_MODULE_ERROR}\n```",
                ephemeral=True,
            )
            return
        await interaction.response.send_message(
            "```\n" + senseki_sheets.describe_config() + "\n```\n"
            "サービスアカウントのメールアドレスに、対象スプレッドシートを"
            "「編集者」で共有しておく必要があります。",
            ephemeral=True,
        )

    async def _build_export(self, env_only: bool) -> tuple[bytes, str, int]:
        matches = await fetch_all_matches(env_only=env_only)
        guild = self.bot.get_guild(GUILD_ID)
        data = await asyncio.to_thread(build_workbook, matches, guild)
        scope = "現環境" if env_only else "全期間"
        filename = f"senseki_{scope}_{datetime.now(JST).strftime('%Y%m%d_%H%M')}.xlsx"
        return data, filename, len(matches)

    @tasks.loop(time=dtime(hour=WEEKLY_EXPORT_HOUR, minute=WEEKLY_EXPORT_MINUTE, tzinfo=JST))
    async def weekly_export(self):
        if datetime.now(JST).weekday() != WEEKLY_EXPORT_WEEKDAY:
            return
        if SENSEKI_ADMIN_USER_ID == 0:
            print("⚠️ SENSEKI_ADMIN_USER_ID が未設定のため、戦績データの週次送信をスキップしました")
            return
        if Workbook is None:
            print(f"⚠️ openpyxl が読み込めないため、戦績データの週次送信をスキップしました: {OPENPYXL_IMPORT_ERROR}")
            return
        try:
            data, filename, count = await self._build_export(env_only=True)
            admin = self.bot.get_user(SENSEKI_ADMIN_USER_ID) or await self.bot.fetch_user(SENSEKI_ADMIN_USER_ID)
            await admin.send(
                content=f"📊 戦績データ週次抽出（現環境・{count}件）",
                file=discord.File(io.BytesIO(data), filename=filename),
            )
        except Exception as e:
            print(f"⚠️ 戦績データの週次送信に失敗しました: {type(e).__name__}: {e}")

    @weekly_export.before_loop
    async def _before_weekly_export(self):
        await self.bot.wait_until_ready()

    # ---- /戦績データ抽出 ----
    @app_commands.command(name="戦績データ抽出", description="戦績データを生データ＋集計のExcelで出力します（管理者専用）")
    @app_commands.describe(全期間="オンにすると環境をまたいだ全データを出力します（既定は現環境のみ）")
    @app_commands.default_permissions(administrator=True)
    async def senseki_export(self, interaction: discord.Interaction, 全期間: bool = False):
        if Workbook is None:
            await interaction.response.send_message(
                f"openpyxlが読み込めませんでした。\n```\n{OPENPYXL_IMPORT_ERROR}\n```\n"
                "requirements.txt に `openpyxl` を追加してください。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            data, filename, count = await self._build_export(env_only=not 全期間)
            if count == 0:
                await interaction.followup.send("記録がまだありません。", ephemeral=True)
                return

            summary = f"📊 戦績データ抽出（{'全期間' if 全期間 else '現環境'}・{count}件）"
            try:
                await interaction.user.send(
                    content=summary,
                    file=discord.File(io.BytesIO(data), filename=filename),
                )
                await interaction.followup.send("DMにExcelを送付しました。", ephemeral=True)
            except discord.Forbidden:
                await interaction.followup.send(
                    content=f"DMを送信できなかったため、こちらに添付します。\n\n{summary}",
                    file=discord.File(io.BytesIO(data), filename=filename),
                    ephemeral=True,
                )
        except Exception as e:
            print(f"⚠️ 戦績データ抽出で予期しないエラー: {type(e).__name__}: {e}")
            await interaction.followup.send(
                f"処理中にエラーが発生しました。\n```\n{type(e).__name__}: {e}\n```",
                ephemeral=True,
            )

    # ---- /戦績全体 ----
    @app_commands.command(
        name="戦績全体",
        description="サーバー全体の集計を今この瞬間のデータで表示します（管理者専用）",
    )
    @app_commands.default_permissions(administrator=True)
    async def senseki_global(self, interaction: discord.Interaction):
        g = await get_global_summary()
        total = g["total"]
        if total == 0:
            await interaction.response.send_message(
                "まだ記録がありません。", ephemeral=True
            )
            return

        def rate(wins, n):
            return f"{wins / n * 100:.1f}%" if n else "-"

        lines = [
            "📊 **全体集計（現環境）**",
            f"{total}戦 / {g['wins']}勝{total - g['wins']}敗 / 勝率 {rate(g['wins'], total)}",
            f"記録者：{g['users']}名",
        ]

        first_map = {r["is_first"]: r for r in g["by_first"]}
        parts = []
        for key, label in ((1, "先攻"), (0, "後攻")):
            r = first_map.get(key)
            if r:
                parts.append(f"{label} {rate(r['wins'], r['total'])}（{r['total']}戦）")
        if parts:
            lines.append("\n**先後別**\n" + " / ".join(parts))

        lines.append("\n**使用クラス別**")
        for r in g["by_my_class"]:
            lines.append(f"{r['my_class']}：{rate(r['wins'], r['total'])}（{r['total']}戦）")

        lines.append("\n**対面クラス別**")
        for r in g["by_opp_class"]:
            lines.append(f"vs {r['opp_class']}：{rate(r['wins'], r['total'])}（{r['total']}戦）")

        lines.append(
            f"\n※母数が少ない項目は参考値です（{MIN_SAMPLE_FOR_STATS}戦未満は特に）。"
        )
        await interaction.response.send_message("\n".join(lines), ephemeral=True)

    # ---- /デッキ登録 ----
    @app_commands.command(name="デッキ登録", description="よく使うデッキを登録しておきます（複数登録できます）")
    @app_commands.describe(
        フォーマット="このデッキで遊ぶフォーマット",
        クラス="このデッキのクラス",
        デッキ名="デッキ名（例：進化ネメシス）",
    )
    @app_commands.choices(フォーマット=[
        app_commands.Choice(name="ローテーション", value="rotation"),
        app_commands.Choice(name="アンリミテッド", value="unlimited"),
    ])
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    async def deck_register(
        self,
        interaction: discord.Interaction,
        フォーマット: app_commands.Choice[str],
        クラス: app_commands.Choice[str],
        デッキ名: str,
    ):
        name = (デッキ名 or "").strip()
        if not name:
            await interaction.response.send_message("デッキ名を入力してください。", ephemeral=True)
            return

        ok = await add_template(
            str(interaction.user.id), フォーマット.value, クラス.value, name
        )
        if not ok:
            await interaction.response.send_message(
                "同じ組み合わせのデッキがすでに登録されています。", ephemeral=True
            )
            return
        await interaction.response.send_message(
            f"✅ 登録しました\n{name}（{クラス.name} / {フォーマット.name}）\n\n"
            "`/デッキ切替` で選ぶだけで使うデッキを変更できます。",
            ephemeral=True,
        )

    # ---- /デッキ切替 ----
    @app_commands.command(name="デッキ切替", description="登録済みのデッキから、今使うデッキを選びます")
    @app_commands.describe(削除="オンにすると、選んだデッキを登録から削除します")
    async def deck_switch(self, interaction: discord.Interaction, 削除: bool = False):
        templates = await list_templates(str(interaction.user.id))
        if not templates:
            await interaction.response.send_message(
                "登録済みのデッキがありません。`/デッキ登録` で登録してください。",
                ephemeral=True,
            )
            return

        settings = await get_user_settings(str(interaction.user.id))
        if settings is None and not 削除:
            await interaction.response.send_message(
                "先に `/戦績設定` でランク帯まで登録してください。", ephemeral=True
            )
            return

        mode = "delete" if 削除 else "switch"
        header = "削除するデッキを選んでください。" if 削除 else "使うデッキを選んでください。"
        await interaction.response.send_message(
            header, view=TemplateView(templates, mode), ephemeral=True
        )

    # ---- /ランク更新 ----
    @app_commands.command(name="ランク更新", description="ランク帯とグレードだけを変更します（デッキ設定はそのまま）")
    @app_commands.describe(
        ランク帯="現在のランク帯",
        グレード="グランドマスターの方のみ必須。グレードが付いていない場合は「グレードなし」を選択",
    )
    @app_commands.choices(ランク帯=[app_commands.Choice(name=r, value=r) for r in RANK_TIER_CHOICES])
    @app_commands.choices(グレード=(
        [app_commands.Choice(name=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
        + [app_commands.Choice(name=g, value=g) for g in CR_GRADE_CHOICES]
    ))
    async def rank_update(
        self,
        interaction: discord.Interaction,
        ランク帯: app_commands.Choice[str],
        グレード: app_commands.Choice[str] = None,
    ):
        rank_value = ランク帯.value
        raw_grade = グレード.value if グレード else None

        grade_value, note, error = _resolve_grade(rank_value, raw_grade)
        if error:
            await interaction.response.send_message(error, ephemeral=True)
            return

        ok = await update_rank(str(interaction.user.id), rank_value, grade_value)
        if not ok:
            await interaction.response.send_message(
                "先に `/戦績設定` で初期設定をしてください。", ephemeral=True
            )
            return

        label = f"{rank_value}（{grade_value}）" if grade_value else rank_value
        lines = [f"✅ ランク帯を更新しました：{label}"]
        if note:
            lines.append(note)
        await interaction.response.send_message("\n".join(lines), ephemeral=True)
        await self._update_board()

    # ---- /戦績設定 ----
    @app_commands.command(name="戦績設定", description="フォーマット・自分のクラス・デッキタイプ・ランク帯を登録します")
    @app_commands.describe(
        フォーマット="使用するフォーマット",
        クラス="自分の使用クラス",
        デッキタイプ="デッキ名（例：進化ネメシス）",
        ランク帯="現在のランク帯",
        グレード="グランドマスターの方のみ必須。CRのグレードを選択してください",
    )
    @app_commands.choices(フォーマット=[
        app_commands.Choice(name="ローテーション", value="rotation"),
        app_commands.Choice(name="アンリミテッド", value="unlimited"),
    ])
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(ランク帯=[app_commands.Choice(name=r, value=r) for r in RANK_TIER_CHOICES])
    @app_commands.choices(グレード=(
        [app_commands.Choice(name=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
        + [app_commands.Choice(name=g, value=g) for g in CR_GRADE_CHOICES]
    ))
    async def senseki_settings(
        self,
        interaction: discord.Interaction,
        フォーマット: app_commands.Choice[str],
        クラス: app_commands.Choice[str],
        デッキタイプ: str,
        ランク帯: app_commands.Choice[str],
        グレード: app_commands.Choice[str] = None,
    ):
        rank_value = ランク帯.value
        raw_grade = グレード.value if グレード else None
        deck_value = (デッキタイプ or "").strip()

        # 空白だけの入力を弾く
        if not deck_value:
            await interaction.response.send_message(
                "デッキタイプを入力してください。", ephemeral=True
            )
            return

        grade_value, note, error = _resolve_grade(rank_value, raw_grade)
        if error:
            await interaction.response.send_message(error, ephemeral=True)
            return

        await upsert_user_settings(
            str(interaction.user.id), フォーマット.value, クラス.value,
            deck_value, rank_value, grade_value,
        )
        # ここで指定したデッキはテンプレートにも入れておく（次回から選ぶだけで済む）
        await add_template(
            str(interaction.user.id), フォーマット.value, クラス.value, deck_value
        )
        rank_label = f"{rank_value}（{grade_value}）" if grade_value else rank_value
        lines = [
            "✅ 設定を保存しました",
            f"フォーマット：{フォーマット.name}",
            f"クラス：{クラス.name}",
            f"デッキタイプ：{deck_value}",
            f"ランク帯：{rank_label}",
        ]
        if note:
            lines.append(note)
        lines.append("\n`/戦績` で記録できます。相手クラス・先後・勝敗だけ入力すればOKです。")
        lines.append("ランクが変わったら `/ランク更新`、デッキを変えるときは `/デッキ切替` が早いです。")
        await interaction.response.send_message("\n".join(lines), ephemeral=True)
        await self._update_board()

    async def _start_record_flow(self, interaction: discord.Interaction):
        """記録の入り口。/戦績 とパネルのボタンの両方から呼ぶ"""
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None:
            await interaction.response.send_message(
                "先に `/戦績設定` でフォーマット・クラスを登録してください。",
                ephemeral=True,
            )
            return

        view = SensekiFlowView(str(interaction.user.id), settings)
        await interaction.response.send_message(
            view.screen(
                "相手クラスを選択してください。\n"
                "-# 上の内容が違う場合は `/デッキ切替` `/ランク更新` で変更してください。"
            ),
            view=view,
            ephemeral=True,
        )

    # ---- /戦績 ----
    @app_commands.command(name="戦績", description="対戦結果を記録します")
    async def senseki_record(self, interaction: discord.Interaction):
        await self._start_record_flow(interaction)

    # ---- /戦績取消 ----
    @app_commands.command(name="戦績取消", description="直前に記録した1件を取り消します")
    async def senseki_undo(self, interaction: discord.Interaction):
        deleted = await delete_last_match(str(interaction.user.id))
        if deleted is None:
            await interaction.response.send_message(
                "取り消せる記録が見つかりませんでした。", ephemeral=True
            )
            return
        result_label = "勝ち" if deleted["is_win"] else "負け"
        await interaction.response.send_message(
            f"🗑️ 直前の記録を取り消しました\n相手クラス：{deleted['opp_class']} / {result_label}",
            ephemeral=True,
        )

    # ---- /弱点対面（メンバー限定） ----
    @app_commands.command(
        name="弱点対面",
        description="【メンバー限定】あなたの対面別の勝率が低い相手と、おすすめの攻略記事を表示します",
    )
    async def weak_matchups(self, interaction: discord.Interaction):
        if not is_member(interaction.user):
            await interaction.response.send_message(
                "この機能はメンバー限定です。メンバーシップについては固定メッセージをご確認ください。",
                ephemeral=True,
            )
            return

        rows = await get_personal_matchups(str(interaction.user.id))
        candidates = [r for r in rows if r["total"] >= MIN_SAMPLE_PERSONAL]
        if not candidates:
            await interaction.response.send_message(
                f"まだ判定できるだけのデータがありません"
                f"（同じ対面が{MIN_SAMPLE_PERSONAL}戦以上たまると表示されます）。",
                ephemeral=True,
            )
            return

        for r in candidates:
            r["win_rate"] = r["wins"] / r["total"] * 100
        candidates.sort(key=lambda r: r["win_rate"])
        worst = candidates[:3]

        lines = ["📉 **あなたの弱点対面**（現環境・参考値）"]
        for r in worst:
            lines.append(
                f"\n**{r['my_class']}** vs **{r['opp_class']}**："
                f"勝率 {r['win_rate']:.1f}%（{r['wins']}勝{r['total']-r['wins']}敗・{r['total']}戦）"
            )
            link = await find_guide_link(r["my_class"], r["opp_class"])
            if link:
                lines.append(f"📖 {link['url']}")
                if link.get("note"):
                    lines.append(f"　{link['note']}")

        lines.append(
            f"\n-# {MIN_SAMPLE_PERSONAL}戦以上あれば表示されますが、"
            f"母数が少ないうちは参考程度に見てください。"
        )
        await interaction.response.send_message("\n".join(lines), ephemeral=True)

    # ---- 攻略記事の管理（管理者専用） ----
    @app_commands.command(name="攻略記事登録", description="対面ごとのおすすめ攻略記事を登録します（管理者専用）")
    @app_commands.describe(
        相手クラス="この相手クラスへの対策記事",
        URL="記事のURL",
        自分のクラス="任意：自分のクラスを指定すると、その組み合わせでのみ表示されます（省略時は相手クラス共通の記事）",
        メモ="任意：一言コメント",
    )
    @app_commands.choices(相手クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(自分のクラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.default_permissions(administrator=True)
    async def register_guide_link(
        self,
        interaction: discord.Interaction,
        相手クラス: app_commands.Choice[str],
        URL: str,
        自分のクラス: app_commands.Choice[str] = None,
        メモ: str = None,
    ):
        my_class = 自分のクラス.value if 自分のクラス else ""
        await upsert_guide_link(my_class, 相手クラス.value, URL, メモ)
        scope = f"{自分のクラス.name} vs {相手クラス.name}" if 自分のクラス else f"vs {相手クラス.name}（汎用）"
        await interaction.response.send_message(f"✅ 登録しました：{scope}\n{URL}", ephemeral=True)

    @app_commands.command(name="攻略記事一覧", description="登録済みの攻略記事を一覧表示します（管理者専用）")
    @app_commands.default_permissions(administrator=True)
    async def list_guide_links_cmd(self, interaction: discord.Interaction):
        links = await list_guide_links()
        if not links:
            await interaction.response.send_message("まだ登録がありません。", ephemeral=True)
            return
        lines = ["📚 **登録済み攻略記事**"]
        for l in links:
            scope = f"{l['my_class']} vs {l['opp_class']}" if l["my_class"] else f"vs {l['opp_class']}（汎用）"
            lines.append(f"・{scope}：{l['url']}")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    @app_commands.command(name="攻略記事削除", description="登録済みの攻略記事を削除します（管理者専用）")
    @app_commands.describe(
        相手クラス="削除する記事の相手クラス",
        自分のクラス="任意：クラス指定ありで登録した記事を消す場合のみ指定",
    )
    @app_commands.choices(相手クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(自分のクラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.default_permissions(administrator=True)
    async def delete_guide_link_cmd(
        self,
        interaction: discord.Interaction,
        相手クラス: app_commands.Choice[str],
        自分のクラス: app_commands.Choice[str] = None,
    ):
        my_class = 自分のクラス.value if 自分のクラス else ""
        ok = await delete_guide_link(my_class, 相手クラス.value)
        msg = "🗑️ 削除しました。" if ok else "該当する記事が見つかりませんでした。"
        await interaction.response.send_message(msg, ephemeral=True)

    # ---- /戦績確認 ----
    @app_commands.command(name="戦績確認", description="今使っているデッキ・ランクと、現環境での自分の勝率を確認します")
    async def senseki_check(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        lines = []
        if settings is not None:
            deck = settings.get("my_deck") or "デッキ未設定"
            rank = settings.get("rank_tier") or "ランク未設定"
            if settings.get("cr_grade"):
                rank += f"・{settings['cr_grade']}"
            elif settings.get("rank_tier") == GRAND_MASTER_TIER:
                rank += "・グレードなし"
            lines.append(f"🧑 現在の設定：**{deck}**（{settings['my_class']}） / ランク：**{rank}**")
            lines.append("")

        summary = await get_summary(str(interaction.user.id))
        total = summary["total"]
        if total == 0:
            lines.append("まだ記録がありません。`/戦績` で記録してみてください。")
            await interaction.response.send_message("\n".join(lines), ephemeral=True)
            return

        wins = summary["wins"]
        losses = summary["losses"]
        win_rate = wins / total * 100
        lines.append(
            f"📊 現在の環境での戦績\n{wins}勝{losses}敗（{total}戦）\n勝率：{win_rate:.1f}%"
        )
        lines.append("\n対面別の弱点は `/弱点対面` で確認できます（メンバー限定）。")
        await interaction.response.send_message("\n".join(lines), ephemeral=True)


async def setup(bot: commands.Bot):
    await bot.add_cog(SensekiCog(bot), guild=discord.Object(id=GUILD_ID))


# ==============================
# main.py への統合方法
# ==============================
#
#   from senseki import SensekiCog
#
#   # on_ready 内、他のCog登録と同じ場所に追加：
#   if bot.get_cog("SensekiCog") is None:
#       await bot.add_cog(SensekiCog(bot), guild=guild_obj)
#
# 必要な環境変数（任意）:
#   SENSEKI_DB_PATH … 既定値 /data/senseki.db。ボリュームのマウントパスが
#                      違う場合のみ設定する
#
# 必要なライブラリ（requirements.txt に追加）:
#   openpyxl
#   gspread        … スプレッドシート連携を使う場合のみ
#   google-auth    … 同上
#
# デプロイ前提条件:
#   Railway の shirokusa-bot サービスに永続ボリューム（/data）が必要。
#   未設定の場合は Volumes タブから追加してから反映すること。
#
# SENSEKI_ADMIN_USER_ID は設定済み（週次抽出のDM送付先）。
