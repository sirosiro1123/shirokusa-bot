"""
koken_rank.py
攻略情報の貢献度を投稿者ごとに集計し、ランキングとCSVを出力するコマンド

kouryaku_mining.py の判定ロジック（Claudeによる0〜5のスコアリング）を再利用し、
発言単位のスコアを投稿者単位に集計する。

main.py への統合:
    from koken_rank import KokenRankCog
    await bot.add_cog(KokenRankCog(bot), guild=discord.Object(id=GUILD_ID))
"""

import csv
import io
import asyncio
import traceback
from datetime import datetime, timedelta, timezone

import discord
from discord import app_commands
from discord.ext import commands

# 既存モジュールの判定ロジックを再利用
from kouryaku_mining import (
    GUILD_ID,
    TARGET_CHANNELS,
    JST,
    clean_content,
    is_noise,
    judge_all,
)

# ==============================
# 設定
# ==============================

# 「記事に採用した」印として付けるリアクション
ADOPT_EMOJI = "📌"

# このユーザーが ADOPT_EMOJI を付けた投稿のみ「採用」として加点する
# 0 のままなら、誰が付けたかを問わず加点する
ADOPT_JUDGE_USER_ID = 346475730458378240

# 遡る日数の既定値（アマギフ企画を月次で回す想定）
DEFAULT_LOOKBACK_DAYS = 30

# ---- スコアの重み ----
# 意図的に複数の要素を混ぜている。単一指標に寄せると、そこに最適化した
# 投稿（連投・水増し）を招くため、比重は公開しない前提で運用する。

# 記事への採用（最重視）
W_ADOPTED = 50

# AI判定スコアごとの加点
W_BY_SCORE = {5: 12, 4: 6, 3: 2}

# 他メンバーからのリアクション1件あたり（他人の評価なので水増ししにくい）
W_REACTION = 1
CAP_REACTION = 20          # リアクション由来の加点上限

# 投稿があった週数 × この値（単発の連投より継続を評価）
W_ACTIVE_WEEK = 5
CAP_ACTIVE_WEEK = 20

# 1人あたり、質による加点の上限（連投対策）
# ここを緩くすると、低スコアの投稿を大量に出すだけで上位に来られてしまう
CAP_QUALITY = 60


# ==============================
# 履歴収集（リアクション情報つき）
# ==============================

async def collect_with_reactions(
    guild: discord.Guild, lookback_days: int
) -> tuple[list[dict], list[str]]:
    """
    対象チャンネルから候補メッセージを収集する。
    kouryaku_mining.collect_messages とほぼ同じだが、
    リアクション数と採用フラグを追加で取得する。
    """
    after = datetime.now(timezone.utc) - timedelta(days=lookback_days)
    collected: list[dict] = []
    diagnostics: list[str] = []
    seq = 0
    total_scanned = 0

    for ch_id, ch_label in TARGET_CHANNELS.items():
        channel = guild.get_channel(ch_id)
        if channel is None:
            diagnostics.append(f"{ch_label}: チャンネルが見つかりません")
            continue

        targets: list[discord.abc.Messageable] = [channel]
        if isinstance(channel, discord.TextChannel):
            targets.extend(channel.threads)

        ch_scanned = 0
        ch_kept = 0
        ch_error = None

        for target in targets:
            if isinstance(target, (discord.TextChannel, discord.Thread)):
                perms = target.permissions_for(guild.me)
                if not perms.read_message_history:
                    ch_error = "履歴の閲覧権限がありません"
                    continue

            try:
                count = 0
                async for msg in target.history(limit=None, after=after):
                    if msg.author.bot:
                        continue
                    ch_scanned += 1
                    total_scanned += 1

                    text = clean_content(msg)
                    if is_noise(text):
                        continue

                    reaction_count = 0
                    adopted = False
                    for r in msg.reactions:
                        if str(r.emoji) == ADOPT_EMOJI:
                            # 📌 が付いている投稿だけ、誰が付けたかを確認する
                            if ADOPT_JUDGE_USER_ID == 0:
                                adopted = True
                            else:
                                try:
                                    async for u in r.users():
                                        if u.id == ADOPT_JUDGE_USER_ID:
                                            adopted = True
                                            break
                                except discord.HTTPException:
                                    pass
                            continue
                        reaction_count += r.count

                    seq += 1
                    ch_kept += 1
                    collected.append({
                        "id": seq,
                        "channel": ch_label,
                        "author": msg.author.display_name,
                        "author_id": str(msg.author.id),
                        "created_at": msg.created_at,
                        "date": msg.created_at.astimezone(JST).strftime("%Y-%m-%d %H:%M"),
                        "text": text,
                        "url": msg.jump_url,
                        "reactions": reaction_count,
                        "adopted": adopted,
                    })

                    count += 1
                    if count % 300 == 0:
                        await asyncio.sleep(0.1)
            except discord.Forbidden:
                ch_error = "アクセスが拒否されました"
            except discord.HTTPException as e:
                ch_error = f"通信エラー ({e.status})"

        if ch_error:
            diagnostics.append(f"{ch_label}: {ch_error}")
        else:
            diagnostics.append(f"{ch_label}: {ch_scanned}件中 {ch_kept}件が候補")

    diagnostics.append(f"合計 {total_scanned} 件をスキャン")
    return collected, diagnostics


# ==============================
# 投稿者ごとの集計
# ==============================

def aggregate_by_author(
    messages: list[dict], judged: dict[int, dict]
) -> list[dict]:
    """発言単位の判定結果を投稿者ごとに集計する"""
    people: dict[str, dict] = {}

    for m in messages:
        j = judged.get(m["id"])
        if not j:
            continue
        score = int(j.get("score", 0))
        aid = m["author_id"]

        p = people.setdefault(aid, {
            "author_id": aid,
            "author": m["author"],
            "adopted": 0,
            "by_score": {5: 0, 4: 0, 3: 0},
            "reactions": 0,
            "weeks": set(),
            "categories": {},
            "best": [],       # 上位の投稿（確認用）
        })

        # 採用は判定スコアに関わらず加点（実際に記事で使った実績が最重視）
        if m["adopted"]:
            p["adopted"] += 1

        if score >= 3:
            p["by_score"][score] = p["by_score"].get(score, 0) + 1
            p["reactions"] += m["reactions"]
            iso = m["created_at"].astimezone(JST).isocalendar()
            p["weeks"].add((iso[0], iso[1]))
            cat = j.get("category", "その他")
            p["categories"][cat] = p["categories"].get(cat, 0) + 1
            p["best"].append((score, m["url"], j.get("summary", "")))

    rows = []
    for p in people.values():
        quality = sum(
            W_BY_SCORE.get(s, 0) * n for s, n in p["by_score"].items()
        )
        quality = min(quality, CAP_QUALITY)

        reaction_pt = min(p["reactions"] * W_REACTION, CAP_REACTION)
        week_pt = min(len(p["weeks"]) * W_ACTIVE_WEEK, CAP_ACTIVE_WEEK)
        adopt_pt = p["adopted"] * W_ADOPTED

        total = adopt_pt + quality + reaction_pt + week_pt

        p["best"].sort(key=lambda x: -x[0])
        top_cat = sorted(p["categories"].items(), key=lambda x: -x[1])
        cat_text = " / ".join(f"{c}({n})" for c, n in top_cat[:3])

        rows.append({
            "author_id": p["author_id"],
            "author": p["author"],
            "total": total,
            "adopted": p["adopted"],
            "s5": p["by_score"].get(5, 0),
            "s4": p["by_score"].get(4, 0),
            "s3": p["by_score"].get(3, 0),
            "reactions": p["reactions"],
            "weeks": len(p["weeks"]),
            "categories": cat_text,
            "best": p["best"][:3],
        })

    rows.sort(key=lambda r: (-r["total"], -r["adopted"], -r["s5"]))
    return rows


# ==============================
# CSV
# ==============================

def build_rank_csv(rows: list[dict]) -> bytes:
    buf = io.StringIO()
    fieldnames = [
        "順位", "合計スコア", "投稿者", "投稿者ID",
        "記事採用", "スコア5", "スコア4", "スコア3",
        "リアクション", "投稿週数", "主な分類",
        "代表投稿1", "代表投稿2", "代表投稿3",
        "アマギフ対象", "備考",
    ]
    writer = csv.DictWriter(buf, fieldnames=fieldnames)
    writer.writeheader()
    for i, r in enumerate(rows, start=1):
        best = r["best"] + [(0, "", "")] * 3
        writer.writerow({
            "順位": i,
            "合計スコア": r["total"],
            "投稿者": r["author"],
            "投稿者ID": r["author_id"],
            "記事採用": r["adopted"],
            "スコア5": r["s5"],
            "スコア4": r["s4"],
            "スコア3": r["s3"],
            "リアクション": r["reactions"],
            "投稿週数": r["weeks"],
            "主な分類": r["categories"],
            "代表投稿1": best[0][1],
            "代表投稿2": best[1][1],
            "代表投稿3": best[2][1],
            "アマギフ対象": "",
            "備考": "",
        })
    return buf.getvalue().encode("utf-8-sig")


# ==============================
# Cog
# ==============================

class KokenRankCog(commands.Cog):
    def __init__(self, bot: commands.Bot):
        self.bot = bot
        self._running = False

    @app_commands.command(
        name="攻略貢献ランキング",
        description="攻略情報の貢献度を投稿者ごとに集計します（管理者専用）",
    )
    @app_commands.describe(
        遡る日数="集計対象期間（既定30日 / 最大180日）",
        表示人数="ランキングに表示する人数（既定10名）",
    )
    @app_commands.default_permissions(administrator=True)
    async def koken_rank(
        self,
        interaction: discord.Interaction,
        遡る日数: app_commands.Range[int, 7, 180] = DEFAULT_LOOKBACK_DAYS,
        表示人数: app_commands.Range[int, 3, 25] = 10,
    ):
        if interaction.guild is None:
            await interaction.response.send_message(
                "サーバー内で実行してください。", ephemeral=True
            )
            return

        if self._running:
            await interaction.response.send_message(
                "現在ほかの集計を実行中です。完了までお待ちください。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        self._running = True

        try:
            guild = interaction.guild

            await interaction.followup.send(
                f"過去 {遡る日数} 日分を収集しています。\n"
                f"判定にAPIを使うため、数分かかる場合があります。",
                ephemeral=True,
            )

            messages, diagnostics = await collect_with_reactions(guild, 遡る日数)

            if not messages:
                diag_text = "\n".join(f"・{d}" for d in diagnostics)
                await interaction.followup.send(
                    f"対象となる投稿がありませんでした。\n\n```\n{diag_text}\n```",
                    ephemeral=True,
                )
                return

            judged, errors = await judge_all(messages)

            if not judged:
                err_text = "\n".join(f"・{e}" for e in errors[:5]) or "・原因不明"
                await interaction.followup.send(
                    f"判定に失敗しました。\n\n```\n{err_text}\n```",
                    ephemeral=True,
                )
                return

            rows = aggregate_by_author(messages, judged)
            if not rows:
                await interaction.followup.send(
                    "集計対象となる投稿（スコア3以上）がありませんでした。",
                    ephemeral=True,
                )
                return

            # ---- ランキング本文 ----
            lines = []
            for i, r in enumerate(rows[:表示人数], start=1):
                mark = "📌" if r["adopted"] else "　"
                lines.append(
                    f"{i:>2}. {mark} **{r['author']}**　{r['total']}pt\n"
                    f"　　 採用{r['adopted']} / S5:{r['s5']} S4:{r['s4']} S3:{r['s3']}"
                    f" / 反応{r['reactions']} / {r['weeks']}週"
                )
            rank_text = "\n".join(lines)

            total_adopted = sum(r["adopted"] for r in rows)
            warn = ""
            if errors:
                warn = (
                    f"\n\n※一部のバッチで判定に失敗しました（{len(errors)}件）。"
                    f"結果が少ない場合は再実行してください。"
                )

            summary = (
                f"**攻略貢献ランキング**\n"
                f"対象期間: 過去 {遡る日数} 日 / 判定 {len(messages)} 件 / "
                f"対象者 {len(rows)} 名 / 記事採用 {total_adopted} 件\n\n"
                f"{rank_text}\n\n"
                f"上位から内容を確認したうえで、最終的な選定はご自身の判断で行ってください。"
                f"{warn}"
            )

            data = build_rank_csv(rows)
            filename = (
                f"koken_rank_{datetime.now(JST).strftime('%Y%m%d_%H%M')}.csv"
            )

            # Discordの文字数制限に配慮
            if len(summary) > 1900:
                summary = summary[:1900] + "\n…（以下省略。CSVを参照してください）"

            try:
                await interaction.user.send(
                    content=summary,
                    file=discord.File(io.BytesIO(data), filename=filename),
                )
                await interaction.followup.send(
                    "DMにランキングとCSVを送付しました。", ephemeral=True
                )
            except discord.Forbidden:
                await interaction.followup.send(
                    content=summary,
                    file=discord.File(io.BytesIO(data), filename=filename),
                    ephemeral=True,
                )

        except Exception as e:
            print(f"⚠️ 攻略貢献ランキングで予期しないエラー: {type(e).__name__}: {e}")
            traceback.print_exc()
            try:
                await interaction.followup.send(
                    f"処理中にエラーが発生しました。\n```\n{type(e).__name__}: {e}\n```",
                    ephemeral=True,
                )
            except discord.HTTPException:
                pass
        finally:
            self._running = False


async def setup(bot: commands.Bot):
    await bot.add_cog(KokenRankCog(bot), guild=discord.Object(id=GUILD_ID))
