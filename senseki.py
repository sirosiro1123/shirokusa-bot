"""
senseki.py
シャドバ戦績ツール（フェーズ1：記録できる状態）

詳細仕様は `戦績ツール_実装仕様書.md` を参照。
このファイルはフェーズ1のみを実装しています。

フェーズ1で実装する範囲
  1. SQLiteのテーブル作成（matches / user_settings / deck_names）
  2. /戦績設定
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
# フェーズ3の /戦績比較 ではこの値を下回る対面は数字を出さない
MIN_SAMPLE_FOR_STATS = 30

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


def _insert_match_sync(user_id: str, settings: dict, opp_class: str, is_first: bool, is_win: bool):
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
            opp_class, None, 1 if is_first else 0, 1 if is_win else 0,
        ))
        conn.commit()
    finally:
        conn.close()


async def insert_match(user_id: str, settings: dict, opp_class: str, is_first: bool, is_win: bool):
    await asyncio.to_thread(_insert_match_sync, user_id, settings, opp_class, is_first, is_win)
    mark_dirty()


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
        view.show_first_second()
        await interaction.response.edit_message(
            content=f"相手クラス：**{view.opp_class}**\n先攻／後攻を選んでください。",
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
            content=f"相手クラス：**{view.opp_class}** / **先攻**\n結果を選んでください。",
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
            content=f"相手クラス：**{view.opp_class}** / **後攻**\n結果を選んでください。",
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
            content="相手クラスを選択してください。",
            view=view,
        )


class SensekiFlowView(discord.ui.View):
    def __init__(self, user_id: str, settings: dict):
        super().__init__(timeout=300)
        self.user_id = user_id
        self.settings = settings
        self.opp_class = None
        self.is_first = None
        self.clear_items()
        self.add_item(OpponentClassSelect())

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
        self.is_first = None
        self.clear_items()
        self.add_item(OpponentClassSelect())

    async def finalize(self, interaction: discord.Interaction, is_win: bool):
        await insert_match(self.user_id, self.settings, self.opp_class, self.is_first, is_win)
        result_label = "勝ち" if is_win else "負け"
        first_label = "先攻" if self.is_first else "後攻"
        self.clear_items()
        self.add_item(RepeatButton())
        await interaction.response.edit_message(
            content=(
                f"✅ 記録しました\n"
                f"相手クラス：**{self.opp_class}** / {first_label} / **{result_label}**\n\n"
                f"続けて記録する場合は下のボタンを押してください。"
            ),
            view=self,
        )

    async def on_timeout(self):
        # メッセージ編集はできないため何もしない（ephemeralなので放置で問題なし）
        pass


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
        "相手クラス", "先攻/後攻", "勝敗",
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

    def cog_unload(self):
        self.weekly_export.cancel()
        self.daily_sheets_sync.cancel()
        self.debounced_sheets_sync.cancel()

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

    # ---- /戦績設定 ----
    @app_commands.command(name="戦績設定", description="フォーマット・自分のクラス・デッキタイプ・ランク帯を登録します")
    @app_commands.describe(
        フォーマット="使用するフォーマット",
        クラス="自分の使用クラス",
        デッキタイプ="任意：デッキ名（例：進化ネメシス）。分かる範囲で入力してください",
        ランク帯="任意：現在のランク帯",
        グレード="任意：グランドマスターの方のみ。CRのグレードを選択してください",
    )
    @app_commands.choices(フォーマット=[
        app_commands.Choice(name="ローテーション", value="rotation"),
        app_commands.Choice(name="アンリミテッド", value="unlimited"),
    ])
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(ランク帯=[app_commands.Choice(name=r, value=r) for r in RANK_TIER_CHOICES])
    @app_commands.choices(グレード=[app_commands.Choice(name=g, value=g) for g in CR_GRADE_CHOICES])
    async def senseki_settings(
        self,
        interaction: discord.Interaction,
        フォーマット: app_commands.Choice[str],
        クラス: app_commands.Choice[str],
        デッキタイプ: str = None,
        ランク帯: app_commands.Choice[str] = None,
        グレード: app_commands.Choice[str] = None,
    ):
        rank_value = ランク帯.value if ランク帯 else None
        grade_value = グレード.value if グレード else None

        # CRはグランドマスター帯でのみ発生するため、それ以外では保存しない
        note = None
        if grade_value and rank_value != GRAND_MASTER_TIER:
            grade_value = None
            note = (
                f"※グレードは{GRAND_MASTER_TIER}帯でのみ記録されるため、今回は保存していません。"
            )

        await upsert_user_settings(
            str(interaction.user.id), フォーマット.value, クラス.value,
            デッキタイプ, rank_value, grade_value,
        )
        lines = [
            "✅ 設定を保存しました",
            f"フォーマット：{フォーマット.name}",
            f"クラス：{クラス.name}",
        ]
        if デッキタイプ:
            lines.append(f"デッキタイプ：{デッキタイプ}")
        if rank_value:
            if grade_value:
                lines.append(f"ランク帯：{rank_value}（{grade_value}）")
            else:
                lines.append(f"ランク帯：{rank_value}")
        if note:
            lines.append(note)
        lines.append("\n`/戦績` で記録できます。相手クラス・先後・勝敗だけ入力すればOKです。")
        lines.append("ランクが変わったら `/戦績設定` をもう一度実行してください。")
        await interaction.response.send_message("\n".join(lines), ephemeral=True)

    # ---- /戦績 ----
    @app_commands.command(name="戦績", description="対戦結果を記録します")
    async def senseki_record(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None:
            await interaction.response.send_message(
                "先に `/戦績設定` でフォーマット・クラスを登録してください。",
                ephemeral=True,
            )
            return

        format_label = FORMAT_LABELS.get(settings["format"], settings["format"])
        view = SensekiFlowView(str(interaction.user.id), settings)
        await interaction.response.send_message(
            f"現在の設定：{format_label} / {settings['my_class']}\n\n相手クラスを選択してください。",
            view=view,
            ephemeral=True,
        )

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

    # ---- /戦績確認 ----
    @app_commands.command(name="戦績確認", description="現在の環境での自分の勝率を確認します")
    async def senseki_check(self, interaction: discord.Interaction):
        summary = await get_summary(str(interaction.user.id))
        total = summary["total"]
        if total == 0:
            await interaction.response.send_message(
                "まだ記録がありません。`/戦績` で記録してみてください。", ephemeral=True
            )
            return
        wins = summary["wins"]
        losses = summary["losses"]
        win_rate = wins / total * 100
        await interaction.response.send_message(
            (
                f"📊 現在の環境での戦績\n"
                f"{wins}勝{losses}敗（{total}戦）\n"
                f"勝率：{win_rate:.1f}%\n\n"
                "※クラス別・対面別の内訳はフェーズ2で追加予定です。"
            ),
            ephemeral=True,
        )


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
