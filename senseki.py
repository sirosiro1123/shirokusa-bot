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
import os
import sqlite3
from datetime import datetime, timedelta, timezone

import discord
from discord import app_commands
from discord.ext import commands

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
    return await asyncio.to_thread(_delete_last_match_sync, user_id)


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
# Cog本体
# ==============================

class SensekiCog(commands.Cog):
    def __init__(self, bot: commands.Bot):
        self.bot = bot

    async def cog_load(self):
        await init_db()

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
# デプロイ前提条件:
#   Railway の shirokusa-bot サービスに永続ボリューム（/data）が必要。
#   未設定の場合は Volumes タブから追加してから反映すること。
