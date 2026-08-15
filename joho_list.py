"""
joho_list.py
情報提供者ロール保持者の一覧をCSVで出力し、実行者にDM送付するコマンド

main.py に統合する場合は、setup(bot) を呼び出してください。
単体で動かす場合は、末尾のコメントアウト部分を参照。
"""

import csv
import io
import os
import asyncio
from datetime import datetime, timedelta, timezone

import discord
from discord import app_commands
from discord.ext import commands

# ==============================
# 設定
# ==============================

GUILD_ID = 1194515135071539210

# 対象ロール名（オンボーディング「攻略情報を提供したい！」に紐づくロール）
TARGET_ROLE_NAME = "情報提供者"

# 実績ロール名 → スコア
ACHIEVEMENT_ROLES = {
    "BEYOND達成！！！！【ローテ】": 30,
    "BEYOND達成！！！！【アンリミ】": 30,
    "BEYOND達成！！！！（ロ）": 30,
    "BEYOND達成！！！！（ア）": 30,
    "LEGEND達成！！！": 25,
    "ULTIMATE達成！！": 20,
    "EPIC達成！": 15,
}

# クラス専門ロール名（参考情報として記録）
CLASS_ROLES = [
    "エルフ専門", "ロイヤル専門", "ウィッチ専門", "ドラゴン専門",
    "ナイトメア専門", "ビショップ専門", "ネメシス専門",
]

# 発言履歴を遡る日数
LOOKBACK_DAYS = 60

# 1チャンネルあたりの最大取得メッセージ数（負荷対策）
MAX_MESSAGES_PER_CHANNEL = 5000

# 1日あたりのDM送信人数（送信予定日の割り振りに使用）
DM_PER_DAY = 10

JST = timezone(timedelta(hours=9))


# ==============================
# スコア計算
# ==============================

def calc_activity_score(last_post: datetime | None, now: datetime) -> int:
    """最終発言日からアクティブ度スコアを算出"""
    if last_post is None:
        return 0
    days = (now - last_post).days
    if days <= 7:
        return 40
    if days <= 14:
        return 32
    if days <= 30:
        return 25
    if days <= 60:
        return 12
    return 0


def calc_tenure_score(joined_at: datetime | None, now: datetime) -> int:
    """在籍期間からスコアを算出"""
    if joined_at is None:
        return 0
    days = (now - joined_at).days
    if days >= 365:
        return 30
    if days >= 180:
        return 20
    if days >= 90:
        return 10
    return 5


def calc_achievement_score(member: discord.Member) -> tuple[int, str]:
    """実績ロールからスコアと保持ロール名を算出"""
    held = []
    score = 0
    for role in member.roles:
        if role.name in ACHIEVEMENT_ROLES:
            held.append(role.name)
            score = max(score, ACHIEVEMENT_ROLES[role.name])
    return score, " / ".join(held)


# ==============================
# 発言履歴スキャン
# ==============================

async def scan_last_activity(
    guild: discord.Guild,
    target_ids: set[int],
    lookback_days: int = LOOKBACK_DAYS,
) -> dict[int, datetime]:
    """
    ギルド内のテキストチャンネル・フォーラムを遡り、
    対象ユーザーごとの最終発言日時を返す。
    """
    now = datetime.now(timezone.utc)
    after = now - timedelta(days=lookback_days)
    last_seen: dict[int, datetime] = {}

    def record(author_id: int, ts: datetime) -> None:
        if author_id not in target_ids:
            return
        prev = last_seen.get(author_id)
        if prev is None or ts > prev:
            last_seen[author_id] = ts

    channels: list[discord.abc.Messageable] = []

    for ch in guild.text_channels:
        perms = ch.permissions_for(guild.me)
        if perms.read_message_history:
            channels.append(ch)

    # フォーラム内のスレッドも対象に含める
    for forum in guild.forums:
        perms = forum.permissions_for(guild.me)
        if not perms.read_message_history:
            continue
        for thread in forum.threads:
            channels.append(thread)

    # アクティブなスレッド全般
    for thread in guild.threads:
        if thread not in channels:
            perms = thread.permissions_for(guild.me)
            if perms.read_message_history:
                channels.append(thread)

    for ch in channels:
        try:
            count = 0
            async for msg in ch.history(limit=MAX_MESSAGES_PER_CHANNEL, after=after):
                if msg.author.bot:
                    continue
                record(msg.author.id, msg.created_at)
                count += 1
                # 過負荷防止のため定期的に譲る
                if count % 500 == 0:
                    await asyncio.sleep(0.1)
        except (discord.Forbidden, discord.HTTPException):
            continue

    return last_seen


# ==============================
# CSV生成
# ==============================

def build_csv(rows: list[dict]) -> bytes:
    """CSVをUTF-8 BOM付きで生成（Excelで文字化けしないため）"""
    buf = io.StringIO()
    fieldnames = [
        "No", "優先度", "スコア", "ユーザーID", "ユーザー名", "表示名",
        "参加日", "在籍日数", "最終発言日", "発言からの経過日数",
        "実績ロール", "クラスロール",
        "送信予定日", "送信済", "返信あり", "備考",
    ]
    writer = csv.DictWriter(buf, fieldnames=fieldnames)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    return buf.getvalue().encode("utf-8-sig")


# ==============================
# Cog
# ==============================

class JohoListCog(commands.Cog):
    def __init__(self, bot: commands.Bot):
        self.bot = bot

    @app_commands.command(
        name="情報提供者一覧",
        description="情報提供者ロール保持者の一覧をCSVでDM送付します（管理者専用）",
    )
    @app_commands.describe(
        遡る日数="発言履歴を何日分遡るか（既定60日 / 最大180日）",
    )
    @app_commands.default_permissions(administrator=True)
    async def joho_list(
        self,
        interaction: discord.Interaction,
        遡る日数: app_commands.Range[int, 7, 180] = LOOKBACK_DAYS,
    ):
        if interaction.guild is None:
            await interaction.response.send_message(
                "サーバー内で実行してください。", ephemeral=True
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)

        guild = interaction.guild
        role = discord.utils.get(guild.roles, name=TARGET_ROLE_NAME)
        if role is None:
            await interaction.followup.send(
                f"ロール「{TARGET_ROLE_NAME}」が見つかりませんでした。\n"
                f"joho_list.py の TARGET_ROLE_NAME を確認してください。",
                ephemeral=True,
            )
            return

        members = [m for m in role.members if not m.bot]
        if not members:
            await interaction.followup.send(
                f"「{TARGET_ROLE_NAME}」ロールの保持者が見つかりませんでした。\n"
                f"BOTにMembers Intentが有効か確認してください。",
                ephemeral=True,
            )
            return

        await interaction.followup.send(
            f"対象 {len(members)} 名を検出しました。\n"
            f"過去 {遡る日数} 日分の発言履歴をスキャンしています。"
            f"数分かかる場合があります。",
            ephemeral=True,
        )

        target_ids = {m.id for m in members}
        last_seen = await scan_last_activity(guild, target_ids, 遡る日数)

        now = datetime.now(timezone.utc)
        entries = []

        for m in members:
            last_post = last_seen.get(m.id)
            act = calc_activity_score(last_post, now)
            ten = calc_tenure_score(m.joined_at, now)
            ach, ach_names = calc_achievement_score(m)
            total = act + ten + ach

            class_names = " / ".join(
                r.name for r in m.roles if r.name in CLASS_ROLES
            )

            entries.append({
                "member": m,
                "score": total,
                "last_post": last_post,
                "ach_names": ach_names,
                "class_names": class_names,
            })

        # スコア降順、同点なら参加日が古い順
        entries.sort(
            key=lambda e: (
                -e["score"],
                e["member"].joined_at or datetime.max.replace(tzinfo=timezone.utc),
            )
        )

        # 送信予定日を割り振り（翌日から1日DM_PER_DAY名ずつ）
        base_date = datetime.now(JST).date() + timedelta(days=1)

        rows = []
        for i, e in enumerate(entries, start=1):
            m: discord.Member = e["member"]
            last_post = e["last_post"]

            joined_local = m.joined_at.astimezone(JST) if m.joined_at else None
            tenure_days = (now - m.joined_at).days if m.joined_at else ""

            if last_post:
                last_local = last_post.astimezone(JST).strftime("%Y-%m-%d")
                elapsed = (now - last_post).days
            else:
                last_local = f"{遡る日数}日以内になし"
                elapsed = ""

            if e["score"] >= 70:
                priority = "A"
            elif e["score"] >= 45:
                priority = "B"
            elif e["score"] >= 25:
                priority = "C"
            else:
                priority = "D"

            send_date = base_date + timedelta(days=(i - 1) // DM_PER_DAY)

            rows.append({
                "No": i,
                "優先度": priority,
                "スコア": e["score"],
                "ユーザーID": str(m.id),
                "ユーザー名": m.name,
                "表示名": m.display_name,
                "参加日": joined_local.strftime("%Y-%m-%d") if joined_local else "不明",
                "在籍日数": tenure_days,
                "最終発言日": last_local,
                "発言からの経過日数": elapsed,
                "実績ロール": e["ach_names"],
                "クラスロール": e["class_names"],
                "送信予定日": send_date.strftime("%Y-%m-%d"),
                "送信済": "",
                "返信あり": "",
                "備考": "",
            })

        data = build_csv(rows)
        filename = f"joho_teikyousha_{datetime.now(JST).strftime('%Y%m%d_%H%M')}.csv"

        a_count = sum(1 for r in rows if r["優先度"] == "A")
        b_count = sum(1 for r in rows if r["優先度"] == "B")
        c_count = sum(1 for r in rows if r["優先度"] == "C")
        d_count = sum(1 for r in rows if r["優先度"] == "D")

        summary = (
            f"**情報提供者一覧**\n"
            f"対象 {len(rows)} 名 / 過去 {遡る日数} 日分をスキャン\n\n"
            f"優先度A（スコア70以上）: {a_count} 名\n"
            f"優先度B（45〜69）: {b_count} 名\n"
            f"優先度C（25〜44）: {c_count} 名\n"
            f"優先度D（24以下）: {d_count} 名\n\n"
            f"送信予定日は1日{DM_PER_DAY}名で割り振っています。\n"
            f"優先度A・Bから着手することをおすすめします。"
        )

        try:
            await interaction.user.send(
                content=summary,
                file=discord.File(io.BytesIO(data), filename=filename),
            )
            await interaction.followup.send(
                "DMにCSVを送付しました。", ephemeral=True
            )
        except discord.Forbidden:
            # DMが閉じている場合はエフェメラルで直接返す
            await interaction.followup.send(
                content=f"DMを送信できなかったため、こちらに添付します。\n\n{summary}",
                file=discord.File(io.BytesIO(data), filename=filename),
                ephemeral=True,
            )


async def setup(bot: commands.Bot):
    await bot.add_cog(JohoListCog(bot), guild=discord.Object(id=GUILD_ID))


# ==============================
# main.py への統合方法
# ==============================
#
# main.py の on_ready や setup_hook 内で以下を呼び出してください。
#
#   from joho_list import JohoListCog
#
#   async def setup_hook(self):
#       await self.add_cog(JohoListCog(self), guild=discord.Object(id=GUILD_ID))
#       await self.tree.sync(guild=discord.Object(id=GUILD_ID))
#
# 必要なIntents:
#   intents = discord.Intents.default()
#   intents.members = True          # 必須（Developer Portalでも有効化）
#   intents.message_content = False # 本文は読まないため不要
#
# 必要な権限:
#   - メッセージ履歴を読む
#   - チャンネルを見る
