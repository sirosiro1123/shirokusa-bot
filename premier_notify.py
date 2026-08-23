"""
premier_notify.py
Shadowverse Premier Series の試合日程を Discord イベント化＋リマインド通知する Cog

main.py から以下で読み込む:
    from premier_notify import PremierSeriesNotifier
    await bot.add_cog(PremierSeriesNotifier(bot), guild=discord.Object(id=GUILD_ID))
"""

import os
import json
import asyncio
from datetime import datetime, timedelta, timezone

import discord
from discord import app_commands
from discord.ext import tasks, commands

# ==============================
# 設定
# ==============================
GUILD_ID = 1194515135071539210
CHANNEL_ID = 1540616850293923991       # 📺-プレミアシリーズ
ROLE_ID = 1540617363777388585          # @プレミアシリーズ観戦者

SCHEDULE_FILE = "data/premier_schedule.json"
OFFICIAL_URL = "https://ps.shadowverse-wb.com/26-27/schedule-results/"

# 同時視聴に使うイベントVC（Discordイベントの開催場所）
EVENT_VC_ID = 1510620128146882610      # 🎤イベントVC

JST = timezone(timedelta(hours=9))

# 一括作成の対象期間（日数）
CREATE_WINDOW_DAYS = 30
# イベントの想定所要時間
EVENT_DURATION_HOURS = 4
# 時刻が TBD のときに使う既定の開始時刻
DEFAULT_TIME = "12:00"


class PremierSeriesNotifier(commands.Cog):
    """プレミアシリーズのイベント自動作成・リマインド"""

    def __init__(self, bot: commands.Bot):
        self.bot = bot
        self.monthly_create.start()
        self.daily_remind.start()

    def cog_unload(self):
        self.monthly_create.cancel()
        self.daily_remind.cancel()

    # ==============================
    # スケジュール読み込み
    # ==============================
    def load_schedule(self) -> list | None:
        """JSONから試合一覧を読み込む。失敗時は None"""
        if not os.path.exists(SCHEDULE_FILE):
            print(f"⚠️ {SCHEDULE_FILE} が見つかりません（cwd={os.getcwd()}）")
            return None
        try:
            with open(SCHEDULE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            matches = data.get("matches", [])
            print(f"✅ スケジュール読み込み: {len(matches)} 件")
            return matches
        except json.JSONDecodeError as e:
            print(f"❌ JSONの形式が不正です: {e}")
            return None
        except Exception as e:
            print(f"❌ スケジュール読み込み失敗: {type(e).__name__}: {e}")
            return None

    @staticmethod
    def parse_start(round_data: dict) -> datetime | None:
        """date + time から開始日時（JST）を作る"""
        date_str = round_data.get("date")
        if not date_str:
            return None
        time_str = round_data.get("time") or DEFAULT_TIME
        if time_str.upper() == "TBD":
            time_str = DEFAULT_TIME
        try:
            naive = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
            return naive.replace(tzinfo=JST)
        except ValueError:
            return None

    @staticmethod
    def iter_cards(round_data: dict):
        """match_1, match_2, ... を順に返す"""
        i = 1
        while f"match_{i}" in round_data:
            m = round_data[f"match_{i}"]
            yield i, m.get("team_a", "TBD"), m.get("team_b", "TBD")
            i += 1

    def event_name(self, round_data: dict) -> str:
        return f"📺 プレミアシリーズ {round_data.get('title', '')}".strip()

    # ==============================
    # 対象試合の抽出
    # ==============================
    def upcoming_in_window(self, days: int = CREATE_WINDOW_DAYS) -> list:
        """今日から指定日数以内の試合"""
        matches = self.load_schedule()
        if not matches:
            return []
        now = datetime.now(JST)
        limit = now + timedelta(days=days)
        result = []
        for r in matches:
            start = self.parse_start(r)
            if start is None:
                continue
            if now.date() <= start.date() <= limit.date():
                result.append((r, start))
        return sorted(result, key=lambda x: x[1])

    def next_match(self):
        """本日以降で最も近い試合"""
        matches = self.load_schedule()
        if not matches:
            return None
        now = datetime.now(JST)
        candidates = []
        for r in matches:
            start = self.parse_start(r)
            if start and start.date() >= now.date():
                candidates.append((r, start))
        if not candidates:
            return None
        return sorted(candidates, key=lambda x: x[1])[0]

    def today_and_tomorrow(self) -> list:
        """本日・明日の試合（リマインド用）"""
        matches = self.load_schedule()
        if not matches:
            return []
        now = datetime.now(JST)
        today = now.date()
        tomorrow = (now + timedelta(days=1)).date()
        result = []
        for r in matches:
            start = self.parse_start(r)
            if start is None:
                continue
            if start.date() == today:
                result.append((r, start, "当日"))
            elif start.date() == tomorrow:
                result.append((r, start, "前日"))
        return sorted(result, key=lambda x: x[1])

    # ==============================
    # Embed生成
    # ==============================
    def build_embed(self, round_data: dict, start: datetime, label: str = "") -> discord.Embed:
        title = self.event_name(round_data)
        if label:
            title += f" 【{label}】"

        time_str = round_data.get("time") or "時刻未定"
        if time_str.upper() == "TBD":
            time_str = "時刻未定"

        embed = discord.Embed(
            title=title,
            description=(
                f"📅 **{start.strftime('%Y-%m-%d')}** ⏰ **{time_str}**\n"
                f"🔗 [公式ページで配信を確認]({OFFICIAL_URL})"
            ),
            color=0x1E90FF,
        )
        for i, a, b in self.iter_cards(round_data):
            embed.add_field(name=f"ROUND {i}", value=f"**{a}** vs **{b}**", inline=False)
        embed.set_footer(text="Shadowverse Premier Series 26-27")
        return embed

    def build_description(self, round_data: dict) -> str:
        lines = [f"ROUND {i}: {a} vs {b}" for i, a, b in self.iter_cards(round_data)]
        lines.append("")
        lines.append("🎧 イベントVCに集まって同時視聴します。出入り自由です。")
        lines.append(f"配信は公式ページから: {OFFICIAL_URL}")
        return "\n".join(lines)

    # ==============================
    # イベント作成（Discord上の既存イベントと照合して二重作成を防ぐ）
    # ==============================
    async def create_events(self) -> tuple[int, int, list]:
        """
        今後30日分のイベントを作成する。
        戻り値: (作成数, スキップ数, エラーメッセージのリスト)
        """
        guild = self.bot.get_guild(GUILD_ID)
        if guild is None:
            return 0, 0, ["サーバーが取得できませんでした"]

        targets = self.upcoming_in_window()
        if not targets:
            return 0, 0, []

        # 同時視聴用のイベントVCを取得（イベントの開催場所になる）
        event_vc = guild.get_channel(EVENT_VC_ID)
        if event_vc is None:
            return 0, 0, [
                f"イベントVC（ID: {EVENT_VC_ID}）が見つかりません。"
                "IDが正しいか、BOTにそのチャンネルの閲覧権限があるか確認してください"
            ]
        if not isinstance(event_vc, discord.VoiceChannel):
            return 0, 0, [f"ID {EVENT_VC_ID} はボイスチャンネルではありません"]

        # 既存のスケジュールイベントを取得（再起動しても正しく判定できる）
        try:
            existing = await guild.fetch_scheduled_events()
        except Exception as e:
            return 0, 0, [f"既存イベントの取得に失敗: {type(e).__name__}: {e}"]

        existing_keys = set()
        for ev in existing:
            if ev.start_time:
                existing_keys.add((ev.name, ev.start_time.astimezone(JST).strftime("%Y-%m-%d")))

        created = 0
        skipped = 0
        errors = []

        for round_data, start in targets:
            name = self.event_name(round_data)
            key = (name, start.strftime("%Y-%m-%d"))
            if key in existing_keys:
                skipped += 1
                continue

            # 開始時刻が過去だとDiscordが拒否するため、少し先にずらす
            start_time = start
            now = datetime.now(JST)
            if start_time <= now:
                start_time = now + timedelta(minutes=10)

            try:
                await guild.create_scheduled_event(
                    name=name,
                    description=self.build_description(round_data),
                    start_time=start_time,
                    end_time=start_time + timedelta(hours=EVENT_DURATION_HOURS),
                    entity_type=discord.EntityType.voice,
                    channel=event_vc,
                    privacy_level=discord.PrivacyLevel.guild_only,
                )
                created += 1
                existing_keys.add(key)
                await asyncio.sleep(1)  # レート制限対策
            except discord.Forbidden:
                errors.append("イベント作成の権限がありません（BOTに「イベントの管理」権限が必要）")
                break
            except Exception as e:
                errors.append(f"{name}: {type(e).__name__}: {e}")

        return created, skipped, errors

    # ==============================
    # リマインド送信
    # ==============================
    async def send_reminders(self):
        targets = self.today_and_tomorrow()
        if not targets:
            print("📺 プレミアシリーズ: 本日・明日に試合なし")
            return

        channel = self.bot.get_channel(CHANNEL_ID)
        if channel is None:
            print(f"❌ チャンネル {CHANNEL_ID} が見つかりません")
            return

        guild = channel.guild
        role = guild.get_role(ROLE_ID)
        mention = role.mention if role else ""

        for round_data, start, label in targets:
            label_text = "本日開催" if label == "当日" else "明日開催"
            embed = self.build_embed(round_data, start, label_text)
            try:
                await channel.send(mention, embed=embed)
                print(f"✅ リマインド送信: {self.event_name(round_data)} 【{label}】")
            except discord.Forbidden:
                print(f"❌ リマインド送信の権限がありません（チャンネル {CHANNEL_ID}）")
            except Exception as e:
                print(f"❌ リマインド送信エラー: {type(e).__name__}: {e}")

    # ==============================
    # スラッシュコマンド
    # ==============================
    @app_commands.command(
        name="プレミア日程",
        description="直近のプレミアシリーズの試合日程を表示します",
    )
    async def premier_schedule(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)

        result = self.next_match()
        if result is None:
            await interaction.followup.send(
                f"予定されている試合が見つかりませんでした。\n{OFFICIAL_URL}",
                ephemeral=True,
            )
            return

        round_data, start = result
        await interaction.followup.send(embed=self.build_embed(round_data, start), ephemeral=True)

    @app_commands.command(
        name="プレミアイベント作成",
        description="今後30日分のプレミアシリーズのイベントを作成します（管理者専用）",
    )
    @app_commands.default_permissions(administrator=True)
    async def manual_create(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)

        created, skipped, errors = await self.create_events()

        if created == 0 and skipped == 0 and not errors:
            await interaction.followup.send(
                "今後30日以内に試合がありませんでした。\n"
                f"（スケジュールファイル: `{SCHEDULE_FILE}`）",
                ephemeral=True,
            )
            return

        embed = discord.Embed(
            title="プレミアシリーズ イベント作成",
            description=f"新規作成: **{created}** 件\n既に作成済みのためスキップ: **{skipped}** 件",
            color=0x1E90FF if not errors else 0xE67E22,
        )
        if errors:
            embed.add_field(
                name="⚠️ 発生した問題",
                value="\n".join(errors[:5])[:1000],
                inline=False,
            )
        embed.set_footer(text="Shadowverse Premier Series 26-27")

        await interaction.followup.send(embed=embed, ephemeral=True)

    # ==============================
    # 定期実行
    # ==============================
    @tasks.loop(hours=24)
    async def monthly_create(self):
        """毎月1日・15日にイベントを一括作成"""
        if datetime.now(JST).day not in (1, 15):
            return

        created, skipped, errors = await self.create_events()
        print(f"📅 月次イベント作成: 新規{created}件 / スキップ{skipped}件 / エラー{len(errors)}件")

        if created == 0:
            return

        channel = self.bot.get_channel(CHANNEL_ID)
        if channel is None:
            return
        try:
            embed = discord.Embed(
                title="📅 プレミアシリーズのイベントを追加しました",
                description=f"新規作成: **{created}** 件\n\n「イベント」タブから参加登録できます。",
                color=0x1E90FF,
            )
            embed.set_footer(text="Shadowverse Premier Series 26-27")
            await channel.send(embed=embed)
        except Exception as e:
            print(f"⚠️ 月次作成の報告に失敗: {type(e).__name__}: {e}")

    @tasks.loop(hours=24)
    async def daily_remind(self):
        """毎日リマインドを送信"""
        await self.send_reminders()

    @monthly_create.before_loop
    async def before_monthly_create(self):
        await self.bot.wait_until_ready()
        await self._sleep_until_next(9)

    @daily_remind.before_loop
    async def before_daily_remind(self):
        await self.bot.wait_until_ready()
        await self._sleep_until_next(9)

    @staticmethod
    async def _sleep_until_next(hour: int):
        """次の指定時刻（JST）まで待機"""
        now = datetime.now(JST)
        target = now.replace(hour=hour, minute=0, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        wait = (target - now).total_seconds()
        print(f"🔔 プレミアシリーズ: 次回実行は {target.strftime('%Y-%m-%d %H:%M')} JST")
        await asyncio.sleep(wait)