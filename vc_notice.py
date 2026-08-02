"""
VC入室時 注意事項 + 配信ルール確認 + 違反報告 モジュール（しろくさDiscord用）
============================================================================
main.py から読み込まれて動きます。
設定や文面を変えたいときは「設定」欄だけ書き換えてください。

【機能】
  1. VC入室時に注意事項を自動表示（5分で自動削除）
  2. 同じVCで2人目が配信を始めたときに、許可確認を表示
  3. 注意事項に付く報告ボタン
       ticket方式 → 意見箱（チケット制）へ誘導【初期設定】
       form方式   → BOT内フォーム → 運営専用チャンネルへ送信
"""

import time
import datetime
import discord

# =====================================================================
# 参照チャンネル
# =====================================================================
GUILD_ID = 1194515135071539210

VC_ANNOUNCE_URL = f"https://discord.com/channels/{GUILD_ID}/1467189822350295213"
RULES_URL       = f"https://discord.com/channels/{GUILD_ID}/1401016513061720084"
IDEA_BOX_URL    = f"https://discord.com/channels/{GUILD_ID}/1401393478348443788"

VC_MAIN_ID  = 1208804697801236564
VC_SUB_ID   = 1407176497084629122
VC_EVENT_ID = 1510620128146882610

# =====================================================================
# 設定（この範囲だけ書き換えればOK）
# =====================================================================

# ---- 入室時の注意事項 ----
ENABLED = True

# どこに届けるか
#   "channel"        → VC内テキストチャット（全員に見える）
#   "dm"             → 本人へDM（本人だけが見える。DM拒否の人には届かない）
#   "dm_then_channel"→ まずDM。届かなかった人にだけVC内チャットへ表示【おすすめ】
DELIVERY_MODE = "channel"

# DMで送るときのクールダウン秒数（DELIVERY_MODEにdmを含む場合のみ有効）
#   0      → 入室のたびに毎回DM（通知が鳴り続けるため非推奨）
#   86400  → 24時間に1回まで【おすすめ】
#   -1     → 各ユーザーに一度だけ（BOT再起動でリセットされます）
DM_COOLDOWN_SECONDS = 86400

# 何秒後にメッセージを自動削除するか（300秒 = 5分 / 0にすると削除しない）
# ※VC内チャットに出したときのみ適用。DMは削除しません
DELETE_AFTER = 300

# 同じ人への再表示を抑える秒数
#   0    → 入室のたびに毎回表示（現在の設定）
#   1800 → 30分以内の再入室では表示しない（うるさく感じたらこれに変更）
COOLDOWN_SECONDS = 0

# 新しく表示するとき、そのVCに残っている前回の注意書きを削除するか
REPLACE_PREVIOUS = True

# VC間の移動（別のVCへ引っ越し）でも表示するか
NOTIFY_ON_MOVE = True

# このVCでは表示しない（例: イベントVCを外すなら [VC_EVENT_ID]）
EXCLUDE_CHANNEL_IDS = []

# 指定した場合、ここに書いたVCでのみ表示する（空リストなら全VCが対象）
ONLY_CHANNEL_IDS = []


# ---- 配信人数の自動チェック ----
STREAM_CHECK_ENABLED = True

# 許可なしで配信できる人数（1VC1配信なので 1）
STREAM_LIMIT = 1

# 同じVCで連続して確認を出さない秒数（連投防止）
STREAM_CHECK_COOLDOWN = 600

# 配信確認メッセージを何秒後に削除するか（0で削除しない）
STREAM_CHECK_DELETE_AFTER = 180


# ---- ルール違反の報告 ----
# 報告方式を選びます
#   "ticket" → 意見箱（チケット制）へ誘導するボタンを表示【初期設定】
#   "form"   → BOT内のフォームに入力させ、運営専用チャンネルへ送信
#   "none"   → 報告ボタンを表示しない
REPORT_MODE = "ticket"

# 誘導先チャンネル（ticket方式で使用）。意見箱以外にしたい場合はURLを差し替え
REPORT_TICKET_URL = IDEA_BOX_URL
REPORT_TICKET_LABEL = "意見箱で運営に報告する"

# --- 以下は form 方式を選んだときだけ使う設定 ---
# 報告の送信先チャンネルID（★運営だけが見える非公開チャンネル★）
REPORT_CHANNEL_ID = 0

# 報告があったときにメンションするロールID（不要なら 0）
REPORT_MENTION_ROLE_ID = 0

# 報告者の名前を運営に伝えるか（False で完全匿名）
REPORT_SHOW_REPORTER = True

# 同じ人が連続で報告するのを防ぐ秒数
REPORT_COOLDOWN = 60

# =====================================================================
# 表示する内容
# =====================================================================

NOTICE_TITLE = "🎙️ 通話へようこそ！"

# 本文の長さを選びます
#   "short" → 要点3行のみ。詳細は /ガイド へ誘導【おすすめ・毎回表示向き】
#   "full"  → 全文を表示【初回のみ表示する運用向き】
NOTICE_STYLE = "short"

# 短縮版のタイトル（"short" のときに使われます）
NOTICE_TITLE_SHORT = "🎙️ 通話のお約束"

# 報告についての案内文（REPORT_MODE によって自動で切り替わります）
_REPORT_TEXT = {
    "ticket": (
        "**■ 困ったときは**\n"
        "ルール違反を見かけたり、居心地の悪さを感じたときは、\n"
        "下の**ボタンから意見箱**で運営に知らせてください。\n"
        "意見箱は**チケット制**なので、やり取りは運営とあなたにしか見えません。"
    ),
    "form": (
        "**■ 困ったときは**\n"
        "ルール違反を見かけたら、下の**「ルール違反を報告」**ボタンから知らせてください。\n"
        "内容は運営にだけ届き、**この場には一切表示されません**。\n"
        f"※ 改善提案や要望は[意見箱]({IDEA_BOX_URL})へどうぞ"
    ),
    "none": (
        "**■ 困ったときは**\n"
        f"ルール違反を見かけた場合や、改善提案・要望は[意見箱]({IDEA_BOX_URL})へどうぞ。"
    ),
}

NOTICE_BODY_MAIN = (
    "**■ はじめての方へ**\n"
    "・入ったら軽く挨拶してもらえると嬉しいです（スタンプでもOK）\n"
    "・聞き専・作業通話も歓迎です。無言のままでも問題ありません\n"
    "・途中の入退室は自由です。抜けるときの挨拶も不要です\n"
    "\n"
    "**■ 配信について（重要）**\n"
    "・**1つのVCにつき配信は1人まで**です\n"
    "・2人目が配信したい場合は、**先に配信している人の許可**を取ってください\n"
    "・許可が取れれば2配信までOKです\n"
    "・配信・録画をする際は、参加者に一声かけてください\n"
    "・通話内容の無断での切り抜き・拡散は禁止です\n"
    "\n"
    "**■ お願い**\n"
    f"・初めて通話を開催する場合は[VC告知チャンネル]({VC_ANNOUNCE_URL})に投稿してね\n"
    "・他の参加者への誹謗中傷、晒し行為はやめてください\n"
    "・特定の人を輪から外すような立ち回りもNGです\n"
    f"・詳しくは[サーバールール]({RULES_URL})を確認してください\n"
    "\n"
)

NOTICE_COLOR = 0x27AE60

# ---- 短縮版の本文（NOTICE_STYLE = "short" のとき使用） ----
NOTICE_BODY_SHORT = (
    "📺 **配信は1VCにつき1人まで。**2人目は先に配信している人の許可を取ってください\n"
    f"📣 初めて通話を開催するときは[VC告知チャンネル]({VC_ANNOUNCE_URL})へ\n"
    "🚫 誹謗中傷・晒し・通話内容の無断切り抜きは禁止です\n"
    "\n"
    f"聞き専・作業通話も歓迎です。詳しくは `/ガイド` か[サーバールール]({RULES_URL})をどうぞ。"
)

NOTICE_FOOTER = "このメッセージは5分後に自動で消えます"
NOTICE_FOOTER_DM = "しろくさDiscord ｜ この案内はあなたにだけ送られています"
NOTICE_IMAGE_URL = None

# ---- 配信人数オーバー時のメッセージ ----
STREAM_CHECK_TITLE = "📺 配信人数の確認"
STREAM_CHECK_BODY = (
    "このVCで**{count}人**が同時に配信しています。\n"
    "\n"
    "しろくさDiscordでは**1VCにつき配信は1人まで**がルールです。\n"
    "2人目以降の方は、先に配信していた方の**許可を取ってから**お願いします。\n"
    "\n"
    "※ すでに許可が取れている場合は、このメッセージは無視してください🙏"
)
STREAM_CHECK_COLOR = 0xF39C12

# =====================================================================
# ここから下は編集不要
# =====================================================================

_last_shown = {}
_last_message = {}
_last_stream_check = {}
_last_report = {}
_dm_sent = {}      # {user_id: 最後にDMを送った時刻}


def _effective_mode() -> str:
    """設定不足のまま form を選んでいる場合は none に落とす（事故防止）。"""
    if REPORT_MODE == "form" and not REPORT_CHANNEL_ID:
        return "none"
    if REPORT_MODE == "ticket" and not REPORT_TICKET_URL:
        return "none"
    return REPORT_MODE if REPORT_MODE in _REPORT_TEXT else "none"


# ---------------------------------------------------------------------
# 違反報告フォーム（form方式のときだけ使用）
# ---------------------------------------------------------------------
class ReportModal(discord.ui.Modal, title="ルール違反の報告"):
    place = discord.ui.TextInput(
        label="どこで起きましたか",
        placeholder="例: メインVC / シャドバ雑談チャンネル",
        required=True,
        max_length=100,
    )
    detail = discord.ui.TextInput(
        label="何があったか、できるだけ具体的に",
        placeholder="誰が・いつ・どんな発言や行動をしたか。分かる範囲で構いません。",
        style=discord.TextStyle.paragraph,
        required=True,
        max_length=1000,
    )
    note = discord.ui.TextInput(
        label="運営に伝えたいこと（任意）",
        placeholder="例: 名前は本人に伝えないでほしい / 継続的に起きている など",
        style=discord.TextStyle.paragraph,
        required=False,
        max_length=500,
    )

    async def on_submit(self, interaction: discord.Interaction):
        now = time.time()
        if now - _last_report.get(interaction.user.id, 0) < REPORT_COOLDOWN:
            await interaction.response.send_message(
                "少し時間をおいてから、もう一度お願いします。", ephemeral=True
            )
            return

        channel = interaction.client.get_channel(REPORT_CHANNEL_ID)
        if channel is None:
            await interaction.response.send_message(
                "送信先が正しく設定されていないため、報告を届けられませんでした。\n"
                "お手数ですが、しろくさへ直接ご連絡ください。",
                ephemeral=True,
            )
            print(f"[vc_notice] REPORT_CHANNEL_ID({REPORT_CHANNEL_ID}) が見つかりません")
            return

        embed = discord.Embed(
            title="🚨 ルール違反の報告が届きました",
            color=0xE74C3C,
            timestamp=datetime.datetime.now(datetime.timezone.utc),
        )
        embed.add_field(name="発生場所", value=str(self.place), inline=False)
        embed.add_field(name="内容", value=str(self.detail), inline=False)
        if str(self.note).strip():
            embed.add_field(name="運営への補足", value=str(self.note), inline=False)

        if REPORT_SHOW_REPORTER:
            embed.add_field(
                name="報告者",
                value=f"{interaction.user.mention}（{interaction.user}｜ID: {interaction.user.id}）",
                inline=False,
            )
        else:
            embed.add_field(name="報告者", value="匿名", inline=False)

        content = f"<@&{REPORT_MENTION_ROLE_ID}>" if REPORT_MENTION_ROLE_ID else None

        try:
            await channel.send(content=content, embed=embed)
        except discord.HTTPException as e:
            await interaction.response.send_message(
                "送信に失敗しました。お手数ですが、しろくさへ直接ご連絡ください。",
                ephemeral=True,
            )
            print(f"[vc_notice] 報告の送信失敗: {e}")
            return

        _last_report[interaction.user.id] = now
        await interaction.response.send_message(
            "報告を運営に送信しました。伝えてくれてありがとうございます。\n"
            "内容はこの場には表示されていないので、安心してください。",
            ephemeral=True,
        )


class ReportView(discord.ui.View):
    """form方式の報告ボタン。"""

    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="ルール違反を報告",
        emoji="🚨",
        style=discord.ButtonStyle.secondary,
        custom_id="vc_report_button",
    )
    async def report_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_modal(ReportModal())


class TicketLinkView(discord.ui.View):
    """ticket方式。意見箱へジャンプするだけのリンクボタン。"""

    def __init__(self):
        super().__init__(timeout=None)
        self.add_item(
            discord.ui.Button(
                label=REPORT_TICKET_LABEL,
                emoji="📮",
                style=discord.ButtonStyle.link,
                url=REPORT_TICKET_URL,
            )
        )


# ---------------------------------------------------------------------
# 表示物の生成
# ---------------------------------------------------------------------
def build_notice_embed(in_dm: bool = False) -> discord.Embed:
    if NOTICE_STYLE == "short":
        title = NOTICE_TITLE_SHORT
        body = NOTICE_BODY_SHORT
        if _effective_mode() == "none":
            body += f"\n困ったことがあれば[意見箱]({IDEA_BOX_URL})までご相談ください。"
    else:
        title = NOTICE_TITLE
        body = NOTICE_BODY_MAIN + _REPORT_TEXT[_effective_mode()]

    embed = discord.Embed(title=title, description=body, color=NOTICE_COLOR)
    if NOTICE_IMAGE_URL:
        embed.set_image(url=NOTICE_IMAGE_URL)
    if in_dm:
        if NOTICE_FOOTER_DM:
            embed.set_footer(text=NOTICE_FOOTER_DM)
    elif NOTICE_FOOTER and DELETE_AFTER > 0:
        embed.set_footer(text=NOTICE_FOOTER)
    return embed


def build_notice_view():
    mode = _effective_mode()
    if mode == "ticket":
        return TicketLinkView()
    if mode == "form":
        return ReportView()
    return None


def build_stream_check_embed(count: int) -> discord.Embed:
    return discord.Embed(
        title=STREAM_CHECK_TITLE,
        description=STREAM_CHECK_BODY.format(count=count),
        color=STREAM_CHECK_COLOR,
    )


def _is_target_channel(channel) -> bool:
    if ONLY_CHANNEL_IDS:
        return channel.id in ONLY_CHANNEL_IDS
    return channel.id not in EXCLUDE_CHANNEL_IDS


def _can_send(channel) -> bool:
    perms = channel.permissions_for(channel.guild.me)
    return perms.view_channel and perms.send_messages


def _count_streamers(channel) -> int:
    count = 0
    for m in channel.members:
        if m.bot:
            continue
        if m.voice is not None and m.voice.self_stream:
            count += 1
    return count


# ---------------------------------------------------------------------
# 本体
# ---------------------------------------------------------------------
def setup_vc_notice(bot):
    """既存のbotにVC入室時の注意事項・配信チェック機能を登録する。"""

    async def _send_dm(member) -> bool:
        """本人へDM。送れたらTrue、拒否設定などで送れなければFalse。"""
        now = time.time()
        last = _dm_sent.get(member.id)

        if DM_COOLDOWN_SECONDS == -1:
            if last is not None:
                return True          # 一度送った人には二度と送らない
        elif DM_COOLDOWN_SECONDS > 0:
            if last is not None and now - last < DM_COOLDOWN_SECONDS:
                return True          # クールダウン中。届いた扱いにする

        try:
            kwargs = {"embed": build_notice_embed(in_dm=True)}
            view = build_notice_view()
            if view is not None:
                kwargs["view"] = view
            await member.send(**kwargs)
            _dm_sent[member.id] = now
            return True
        except discord.Forbidden:
            return False             # DMを受け取らない設定
        except discord.HTTPException as e:
            print(f"[vc_notice] DM送信失敗（{member}）: {e}")
            return False

    async def _send_channel(channel):
        if not _can_send(channel):
            print(f"[vc_notice] 権限不足のため送信できません: {channel.name}")
            return

        if REPLACE_PREVIOUS:
            old = _last_message.pop(channel.id, None)
            if old is not None:
                try:
                    await old.delete()
                except discord.HTTPException:
                    pass

        try:
            kwargs = {"embed": build_notice_embed()}
            view = build_notice_view()
            if view is not None:
                kwargs["view"] = view
            if DELETE_AFTER > 0:
                kwargs["delete_after"] = DELETE_AFTER
            _last_message[channel.id] = await channel.send(**kwargs)
        except discord.HTTPException as e:
            print(f"[vc_notice] 注意事項の送信失敗: {e}")

    async def _send_notice(member, channel):
        if DELIVERY_MODE == "dm":
            await _send_dm(member)
        elif DELIVERY_MODE == "dm_then_channel":
            if not await _send_dm(member):
                await _send_channel(channel)
        else:
            await _send_channel(channel)

    async def _check_streams(channel):
        if not STREAM_CHECK_ENABLED:
            return
        if not _is_target_channel(channel) or not _can_send(channel):
            return

        count = _count_streamers(channel)
        if count <= STREAM_LIMIT:
            return

        now = time.time()
        if now - _last_stream_check.get(channel.id, 0) < STREAM_CHECK_COOLDOWN:
            return
        _last_stream_check[channel.id] = now

        try:
            kwargs = {"embed": build_stream_check_embed(count)}
            if STREAM_CHECK_DELETE_AFTER > 0:
                kwargs["delete_after"] = STREAM_CHECK_DELETE_AFTER
            await channel.send(**kwargs)
        except discord.HTTPException as e:
            print(f"[vc_notice] 配信確認の送信失敗: {e}")

    @bot.event
    async def on_voice_state_update(member, before, after):
        if member.bot:
            return

        moved = before.channel != after.channel

        # ---- 入室時の注意事項 ----
        if ENABLED and moved and after.channel is not None:
            if before.channel is None or NOTIFY_ON_MOVE:
                if _is_target_channel(after.channel):
                    now = time.time()
                    if COOLDOWN_SECONDS <= 0 or now - _last_shown.get(member.id, 0) >= COOLDOWN_SECONDS:
                        _last_shown[member.id] = now
                        await _send_notice(member, after.channel)

        # ---- 配信の開始を検知 ----
        started_stream = after.self_stream and not before.self_stream
        if after.channel is not None and (started_stream or (moved and after.self_stream)):
            await _check_streams(after.channel)
