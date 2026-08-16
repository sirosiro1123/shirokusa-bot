"""
kouryaku_mining.py
クラス別雑談チャンネル等の過去ログから、攻略情報として使えそうな発言を抽出する

- 直近30日分をスキャン
- ノイズを事前除去したうえで Claude API で価値判定
- スコア順にCSV化し、実行者にDM送付

main.py に統合する場合は、末尾の統合方法を参照してください。
"""

import asyncio
import csv
import io
import json
import os
import re
from datetime import datetime, timedelta, timezone

import discord
from discord import app_commands
from discord.ext import commands

try:
    from anthropic import AsyncAnthropic
except ImportError:  # ライブラリ未導入時に起動を止めない
    AsyncAnthropic = None


# ==============================
# 設定
# ==============================

GUILD_ID = 1194515135071539210

# スキャン対象チャンネル
TARGET_CHANNELS = {
    1384171023766781982: "エルフ",
    1384171025151037631: "ロイヤル",
    1384171181090930811: "ウィッチ",
    1384171197004255325: "ドラゴン",
    1384171212716113930: "ナイトメア",
    1384171225345036338: "ビショップ",
    1384171237919428659: "ネメシス",
    1401026533434458264: "シャドバ相談所",
    1458372480795414569: "今日の1判断",
    1521414016491065454: "2pick",
}

# 遡る日数の既定値
DEFAULT_LOOKBACK_DAYS = 30

# 事前フィルタ
MIN_LENGTH = 25           # これ未満の文字数は除外
MAX_LENGTH = 1500         # 長すぎる投稿は切り詰め

# API設定
MODEL = "claude-haiku-4-5"
BATCH_SIZE = 15           # 1回のAPI呼び出しで判定する件数
MAX_API_CALLS = 200       # 暴走防止の上限
API_CONCURRENCY = 3       # 同時実行数

# 出力する最低スコア
MIN_OUTPUT_SCORE = 3

JST = timezone(timedelta(hours=9))

# 明らかなノイズを弾く正規表現
NOISE_PATTERNS = [
    re.compile(r"^[\s\W_]+$"),                       # 記号・絵文字のみ
    re.compile(r"^(おはよ|こんにち|こんばん|おつ|乙|ありがと|了解|把握|草|www|それな)"),
    re.compile(r"^(参加|募集|やります|入ります|抜けます|落ちます)"),
]

SYSTEM_PROMPT = """あなたはシャドウバース：ワールズビヨンドのコミュニティ運営を補助するアシスタントです。
Discordの雑談チャンネルの発言を読み、攻略記事の материал として使えるかを判定してください。

各発言に対して、以下を出力してください。

score: 0〜5の整数
  5 = 具体的なプレイング判断、構築の採用理由、対面別の立ち回りなど、そのまま記事の核になる内容
  4 = カード評価や環境認識として有用だが、補足が必要な内容
  3 = 攻略に関係するが断片的で、他の情報と組み合わせる必要がある内容
  2 = 攻略の話題ではあるが、質問や感想が中心で情報価値が低い
  1 = わずかに攻略に触れているだけ
  0 = 攻略と無関係な雑談

category: 次のいずれか
  デッキ構築 / カード評価 / プレイング / 対面相性 / 環境考察 / 2Pick / その他

summary: その発言の要点を1文で（30文字以内、日本語）

判定は厳しめに行ってください。scoreが4以上になるのは全体の1割程度が妥当です。

必ず以下のJSON配列のみを出力してください。前置きやMarkdownのコードブロックは不要です。
[{"id": 発言ID, "score": 数値, "category": "分類", "summary": "要点"}]
"""


# ==============================
# 事前フィルタ
# ==============================

def is_noise(content: str) -> bool:
    """API に渡す前に明らかなノイズを除外"""
    text = content.strip()
    if len(text) < MIN_LENGTH:
        return True
    for pat in NOISE_PATTERNS:
        if pat.match(text):
            return True
    # URLのみの投稿
    if re.fullmatch(r"https?://\S+", text):
        return True
    return False


def clean_content(msg: discord.Message) -> str:
    """メンションやカスタム絵文字を読みやすい形に整形"""
    text = msg.clean_content
    text = re.sub(r"<a?:(\w+):\d+>", r":\1:", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > MAX_LENGTH:
        text = text[:MAX_LENGTH] + "…"
    return text


# ==============================
# 履歴収集
# ==============================

async def collect_messages(
    guild: discord.Guild, lookback_days: int
) -> list[dict]:
    """対象チャンネルから候補メッセージを収集"""
    after = datetime.now(timezone.utc) - timedelta(days=lookback_days)
    collected: list[dict] = []
    seq = 0

    for ch_id, ch_label in TARGET_CHANNELS.items():
        channel = guild.get_channel(ch_id)
        if channel is None:
            continue

        targets: list[discord.abc.Messageable] = [channel]
        # チャンネル内スレッドも対象に含める
        if isinstance(channel, discord.TextChannel):
            targets.extend(channel.threads)

        for target in targets:
            perms = None
            if isinstance(target, (discord.TextChannel, discord.Thread)):
                perms = target.permissions_for(guild.me)
            if perms and not perms.read_message_history:
                continue

            try:
                count = 0
                async for msg in target.history(limit=None, after=after):
                    if msg.author.bot:
                        continue
                    text = clean_content(msg)
                    if is_noise(text):
                        continue

                    seq += 1
                    collected.append({
                        "id": seq,
                        "channel": ch_label,
                        "author": msg.author.display_name,
                        "author_id": str(msg.author.id),
                        "date": msg.created_at.astimezone(JST).strftime("%Y-%m-%d %H:%M"),
                        "text": text,
                        "url": msg.jump_url,
                    })

                    count += 1
                    if count % 300 == 0:
                        await asyncio.sleep(0.1)
            except (discord.Forbidden, discord.HTTPException):
                continue

    return collected


# ==============================
# API 判定
# ==============================

def parse_json_array(raw: str) -> list[dict]:
    """コードブロック混入に耐えるJSONパース"""
    text = raw.strip()
    text = re.sub(r"^```(?:json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1:
        return []
    try:
        data = json.loads(text[start:end + 1])
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


async def judge_batch(
    client, batch: list[dict], sem: asyncio.Semaphore
) -> list[dict]:
    """1バッチ分を判定"""
    lines = []
    for item in batch:
        lines.append(
            f'{{"id": {item["id"]}, "channel": "{item["channel"]}", '
            f'"text": {json.dumps(item["text"], ensure_ascii=False)}}}'
        )
    user_content = "以下の発言を判定してください。\n\n" + "\n".join(lines)

    async with sem:
        for attempt in range(3):
            try:
                resp = await client.messages.create(
                    model=MODEL,
                    max_tokens=2000,
                    system=SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": user_content}],
                )
                raw = "".join(
                    b.text for b in resp.content if getattr(b, "type", "") == "text"
                )
                results = parse_json_array(raw)
                if results:
                    return results
except Exception as e:
                print(f"⚠️ API判定エラー (attempt {attempt+1}): {type(e).__name__}: {e}")
                await asyncio.sleep(2 ** attempt)
    return []


async def judge_all(messages: list[dict]) -> dict[int, dict]:
    """全メッセージを判定し、id をキーにした辞書で返す"""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
if AsyncAnthropic is None:
        print("⚠️ anthropicライブラリが導入されていません。requirements.txtを確認してください。")
        return {}
    if not api_key:
        print("⚠️ ANTHROPIC_API_KEYが取得できません。")
        return {}

    client = AsyncAnthropic(api_key=api_key)
    sem = asyncio.Semaphore(API_CONCURRENCY)

    batches = [
        messages[i:i + BATCH_SIZE]
        for i in range(0, len(messages), BATCH_SIZE)
    ][:MAX_API_CALLS]

    tasks = [judge_batch(client, b, sem) for b in batches]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    judged: dict[int, dict] = {}
    for res in results:
        if isinstance(res, Exception) or not res:
            continue
        for item in res:
            try:
                judged[int(item["id"])] = {
                    "score": int(item.get("score", 0)),
                    "category": str(item.get("category", "")),
                    "summary": str(item.get("summary", "")),
                }
            except (ValueError, TypeError, KeyError):
                continue
    return judged


# ==============================
# CSV生成
# ==============================

def build_csv(rows: list[dict]) -> bytes:
    buf = io.StringIO()
    fieldnames = [
        "No", "スコア", "分類", "要点", "チャンネル",
        "発言者", "発言者ID", "日時", "本文", "リンク",
        "採用", "本人確認", "備考",
    ]
    writer = csv.DictWriter(buf, fieldnames=fieldnames)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    return buf.getvalue().encode("utf-8-sig")


# ==============================
# Cog
# ==============================

class KouryakuMiningCog(commands.Cog):
    def __init__(self, bot: commands.Bot):
        self.bot = bot
        self._running = False

    @app_commands.command(
        name="攻略情報抽出",
        description="クラス別チャンネル等から攻略情報の候補を抽出します（管理者専用）",
    )
    @app_commands.describe(
        遡る日数="何日分を対象にするか（既定30日 / 最大90日）",
        最低スコア="CSVに出力する最低スコア（既定3）",
    )
    @app_commands.default_permissions(administrator=True)
    async def mining(
        self,
        interaction: discord.Interaction,
        遡る日数: app_commands.Range[int, 1, 90] = DEFAULT_LOOKBACK_DAYS,
        最低スコア: app_commands.Range[int, 0, 5] = MIN_OUTPUT_SCORE,
    ):
        if interaction.guild is None:
            await interaction.response.send_message(
                "サーバー内で実行してください。", ephemeral=True
            )
            return

        if self._running:
            await interaction.response.send_message(
                "現在別の抽出処理が実行中です。完了までお待ちください。",
                ephemeral=True,
            )
            return

        if not os.environ.get("ANTHROPIC_API_KEY"):
            await interaction.response.send_message(
                "環境変数 ANTHROPIC_API_KEY が設定されていません。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        self._running = True

        try:
            guild = interaction.guild

            await interaction.followup.send(
                f"過去 {遡る日数} 日分の発言を収集しています。", ephemeral=True
            )

            messages = await collect_messages(guild, 遡る日数)

            if not messages:
                await interaction.followup.send(
                    "対象となる発言が見つかりませんでした。\n"
                    "Message Content Intent が有効か確認してください。",
                    ephemeral=True,
                )
                return

            capped = len(messages) > MAX_API_CALLS * BATCH_SIZE
            note = ""
            if capped:
                note = (
                    f"\n※上限に達したため、新しい順 "
                    f"{MAX_API_CALLS * BATCH_SIZE} 件のみ判定します。"
                )
                messages = messages[-(MAX_API_CALLS * BATCH_SIZE):]

            await interaction.followup.send(
                f"候補 {len(messages)} 件を判定しています。"
                f"数分かかります。{note}",
                ephemeral=True,
            )

            judged = await judge_all(messages)

            if not judged:
                await interaction.followup.send(
                    "判定に失敗しました。APIキーと残高を確認してください。",
                    ephemeral=True,
                )
                return

            rows = []
            for m in messages:
                j = judged.get(m["id"])
                if not j or j["score"] < 最低スコア:
                    continue
                rows.append({
                    "スコア": j["score"],
                    "分類": j["category"],
                    "要点": j["summary"],
                    "チャンネル": m["channel"],
                    "発言者": m["author"],
                    "発言者ID": m["author_id"],
                    "日時": m["date"],
                    "本文": m["text"],
                    "リンク": m["url"],
                    "採用": "",
                    "本人確認": "",
                    "備考": "",
                })

            if not rows:
                await interaction.followup.send(
                    f"スコア {最低スコア} 以上の発言はありませんでした。\n"
                    f"最低スコアを下げて再実行してみてください。",
                    ephemeral=True,
                )
                return

            rows.sort(key=lambda r: (-r["スコア"], r["チャンネル"]))
            for i, r in enumerate(rows, start=1):
                r["No"] = i

            data = build_csv(rows)
            filename = (
                f"kouryaku_candidates_"
                f"{datetime.now(JST).strftime('%Y%m%d_%H%M')}.csv"
            )

            by_score = {}
            by_channel = {}
            for r in rows:
                by_score[r["スコア"]] = by_score.get(r["スコア"], 0) + 1
                by_channel[r["チャンネル"]] = by_channel.get(r["チャンネル"], 0) + 1

            score_lines = "\n".join(
                f"スコア{s}: {by_score[s]} 件" for s in sorted(by_score, reverse=True)
            )
            ch_lines = "\n".join(
                f"{c}: {n} 件"
                for c, n in sorted(by_channel.items(), key=lambda x: -x[1])
            )

            summary = (
                f"**攻略情報 抽出結果**\n"
                f"対象期間: 過去 {遡る日数} 日 / 判定 {len(messages)} 件\n\n"
                f"**スコア別**\n{score_lines}\n\n"
                f"**チャンネル別**\n{ch_lines}\n\n"
                f"スコア5から順に確認することをおすすめします。\n"
                f"記事化する際は、発言をそのまま引用せず"
                f"内容を参考にして書き直してください。"
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
                await interaction.followup.send(
                    content=f"DMを送信できなかったため、こちらに添付します。\n\n{summary}",
                    file=discord.File(io.BytesIO(data), filename=filename),
                    ephemeral=True,
                )

        finally:
            self._running = False


async def setup(bot: commands.Bot):
    await bot.add_cog(KouryakuMiningCog(bot), guild=discord.Object(id=GUILD_ID))


# ==============================
# main.py への統合方法
# ==============================
#
#   from kouryaku_mining import KouryakuMiningCog
#
#   async def setup_hook(self):
#       await self.add_cog(KouryakuMiningCog(self), guild=discord.Object(id=GUILD_ID))
#       await self.tree.sync(guild=discord.Object(id=GUILD_ID))
#
# 必要なIntents:
#   intents.message_content = True   # 必須（Developer Portalでも有効化）
#
# 必要な環境変数:
#   ANTHROPIC_API_KEY
#
# 必要なライブラリ:
#   anthropic
