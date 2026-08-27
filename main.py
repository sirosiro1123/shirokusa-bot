import os
import discord
from discord.ext import commands
from discord import app_commands

from vc_notice import (
    setup_vc_notice,
    build_notice_embed,
    build_notice_view,
    ReportModal,
    ReportView,
    REPORT_MODE,
    REPORT_TICKET_URL,
    _effective_mode,
)

# 管理者用コマンド（追加）
from joho_list import JohoListCog
from kouryaku_mining import KouryakuMiningCog
from koken_rank import KokenRankCog
from senseki import SensekiCog
from premier_notify import PremierSeriesNotifier

# ==============================
# サーバーID
# ==============================
GUILD_ID = 1194515135071539210

# ==============================
# チャンネルID設定
# ==============================
BASE = "https://discord.com/channels/1194515135071539210"

# VC関連
VC_MAIN = f"{BASE}/1208804697801236564"
VC_SUB = f"{BASE}/1407176497084629122"
VC_EVENT = f"{BASE}/1510620128146882610"
VC_ANNOUNCE = f"{BASE}/1467189822350295213"

# 雑談関連
CHAT_GENERAL = f"{BASE}/1200096269293596693"
CHAT_RHYTHM = f"{BASE}/1421749620010123355"
CHAT_MESHI = f"{BASE}/1529673772997808211"
CHAT_BASEBALL = f"{BASE}/1502993527108534272"

# 攻略・クラス相談
CONSULT = f"{BASE}/1401026533434458264"
TODAY_JUDGE = f"{BASE}/1458372480795414569"
CLASS_ELF = f"{BASE}/1384171023766781982"
CLASS_ROYAL = f"{BASE}/1384171025151037631"
CLASS_WITCH = f"{BASE}/1384171181090930811"
CLASS_DRAGON = f"{BASE}/1384171197004255325"
CLASS_NIGHT = f"{BASE}/1384171212716113930"
CLASS_BISHOP = f"{BASE}/1384171225345036338"
CLASS_NEMESIS = f"{BASE}/1384171237919428659"
CLASS_2PICK = f"{BASE}/1521414016491065454"

# お知らせ・ルール
RULES = f"{BASE}/1401016513061720084"
ANNOUNCE_OPS = f"{BASE}/1195015362526326904"
ROLE_LIST = f"{BASE}/1401220412368617647"
ROLE_PROOF = f"{BASE}/1402984558819020800"
IDEA_BOX = f"{BASE}/1401393478348443788"

# 大会・イベント
ANNOUNCE_TOUR = f"{BASE}/1401015993395970259"
ANNOUNCE_EVENT = f"{BASE}/1503580110282821802"
TOURNAMENT = f"{BASE}/1513219971281584169"

# その他
DISBOARD = f"{BASE}/1401437902394753125"
BENRI = f"{BASE}/1534017790363697383"          # 便利機能（追加）
DECK_BOSHU = f"{BASE}/1510289539745316874"     # デッキ募集（追加）

# ==============================
# BOT設定
# ==============================
intents = discord.Intents.default()
intents.message_content = True
intents.voice_states = True          # VC入退室の検知に必要
intents.members = True               # ロール保持者の取得に必要（追加）
bot = commands.Bot(command_prefix="!", intents=intents)

# VC入室時の注意事項機能を登録
setup_vc_notice(bot)


# ==============================
# メインメニューView
# ==============================
class MainMenuView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="🎙️ VC・通話", style=discord.ButtonStyle.primary, custom_id="guide_vc")
    async def vc_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(embed=vc_embed(), view=BackView(), ephemeral=True)

    @discord.ui.button(label="💬 雑談チャンネル", style=discord.ButtonStyle.primary, custom_id="guide_chat")
    async def chat_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(embed=chat_embed(), view=BackView(), ephemeral=True)

    @discord.ui.button(label="⚔️ 攻略・クラス相談", style=discord.ButtonStyle.primary, custom_id="guide_class")
    async def class_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(embed=class_embed(), view=BackView(), ephemeral=True)

    @discord.ui.button(label="📢 お知らせ・ルール", style=discord.ButtonStyle.secondary, custom_id="guide_announce")
    async def announce_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(embed=announce_embed(), view=BackView(), ephemeral=True)

    @discord.ui.button(label="🏆 大会・イベント", style=discord.ButtonStyle.secondary, custom_id="guide_tournament")
    async def tournament_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(embed=tournament_embed(), view=BackView(), ephemeral=True)

    @discord.ui.button(label="⚙️ コマンド・その他", style=discord.ButtonStyle.secondary, custom_id="guide_command")
    async def command_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.send_message(embed=command_embed(), view=BackView(), ephemeral=True)


# ==============================
# 戻るボタンView
# ==============================
class BackView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="◀️ メニューに戻る", style=discord.ButtonStyle.danger, custom_id="guide_back")
    async def back_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await interaction.response.edit_message(embed=main_embed(), view=MainMenuView())


# ==============================
# Embed定義
# ==============================
def main_embed():
    e = discord.Embed(
        title="📘 しろくさDiscord ガイド",
        description="知りたい内容のボタンを押してください！\n\n**初めての方はまず** [サーバールール]({}) **を読んでね🙏".format(RULES),
        color=0x2E6DA4
    )
    e.set_footer(text="しろくさDiscord | ガイドBOT")
    return e


def vc_embed():
    e = discord.Embed(title="🎙️ VC・通話チャンネルの使い方", color=0x27AE60)
    e.add_field(name="🔵 通常VC（シャドバ通話）",
        value=f"[メインVC]({VC_MAIN}) → まずはここに入ってね！\n[サブVC]({VC_SUB}) → メインが埋まっていたらこちらへ\n\n誰でも参加・開催OK！",
        inline=False)
    e.add_field(name="📣 通話を開始する前に",
        value=f"初めて通話を開催する場合は\n[VC告知チャンネル]({VC_ANNOUNCE}) に\n**いつ・どんな通話をするか**を投稿してね！",
        inline=False)
    e.add_field(name="🎉 イベントVC",
        value=f"[イベントVC]({VC_EVENT})\nサーバーイベントや大会開催時に使用します",
        inline=False)
    e.add_field(name="📺 配信のルール",
        value="**1つのVCにつき配信は1人まで**です。\n2人目が配信したい場合は、先に配信している人の\n**許可を取ってから**お願いします。\n許可が取れれば2配信までOKです！",
        inline=False)
    e.add_field(name="📝 入室時の注意事項について",
        value="VCに入ると、チャット欄に注意事項が自動で表示されます。\n（5分で自動的に消えます）\n\nルール違反を見かけたら `/通報` からいつでも運営に報告できます。",
        inline=False)
    e.set_footer(text="◀️ 戻るボタンでメニューに戻れます")
    return e


def chat_embed():
    e = discord.Embed(title="💬 雑談チャンネルの使い方", color=0xF39C12)
    e.add_field(name="🎮 シャドバ雑談",
        value=f"[シャドバ雑談]({CHAT_GENERAL})\n攻略以外のシャドバ話はここで！",
        inline=False)
    e.add_field(name="🎵 リズムゲーム雑談",
        value=f"[リズムゲーム雑談]({CHAT_RHYTHM})\nリズムゲーム関連の話題はこちら",
        inline=False)
    e.add_field(name="🍜 飯テロ",
        value=f"[飯テロチャンネル]({CHAT_MESHI})\n食べ物画像を投稿すると\nしろ子から厳しめの診断が届きます😂\n\n累計の査定に応じて称号が自動で付きます。\n`/飯テロ称号一覧` `/飯テロカルテ` で確認できます\n※あくまで面白診断です！",
        inline=False)
    e.add_field(name="⚾ スポーツ雑談",
        value=f"[野球・スポーツ雑談]({CHAT_BASEBALL})\nスポーツ全般の話題はこちら",
        inline=False)
    e.set_footer(text="◀️ 戻るボタンでメニューに戻れます")
    return e


def class_embed():
    e = discord.Embed(title="⚔️ 攻略・クラス相談チャンネル", color=0xE74C3C)
    e.add_field(name="🤔 どのクラスか決まっていない場合",
        value=f"[シャドバ相談所]({CONSULT})\n「どのデッキを使えばいいか」など\n何でも相談OK！",
        inline=False)
    e.add_field(name="📅 あのターン、どうすれば？",
        value=f"[今日の一判断]({TODAY_JUDGE})\n「このターンどう動けばよかった？」\nを話し合うチャンネル",
        inline=False)
    e.add_field(name="🔧 デッキを見てほしいとき",
        value=f"[デッキ募集]({DECK_BOSHU})\n使っている構築を投稿して意見をもらえます\nネタデッキも歓迎！\n\n意見が不要な場合は【共有のみ】タグを付けてね",
        inline=False)
    e.add_field(name="🃏 クラス別相談チャンネル",
        value=(
            f"[エルフ]({CLASS_ELF}) ／ [ロイヤル]({CLASS_ROYAL}) ／ [ウィッチ]({CLASS_WITCH}) ／ [ドラゴン]({CLASS_DRAGON})\n"
            f"[ナイトメア]({CLASS_NIGHT}) ／ [ビショップ]({CLASS_BISHOP}) ／ [ネメシス]({CLASS_NEMESIS}) ／ [2pick]({CLASS_2PICK})\n\n"
            "デッキ相談・攻略・勝てない場合の対処法など\nクラスごとのチャンネルで相談してね！"
        ),
        inline=False)
    e.set_footer(text="◀️ 戻るボタンでメニューに戻れます")
    return e


def announce_embed():
    e = discord.Embed(title="📢 お知らせ・ルール関連", color=0x8E44AD)
    e.add_field(name="📋 サーバールール",
        value=f"[サーバールール]({RULES})\nまず最初に確認してください！",
        inline=False)
    e.add_field(name="📣 運営アナウンス",
        value=f"[運営アナウンス]({ANNOUNCE_OPS})\nサーバーの変更・しろくさの活動告知など",
        inline=False)
    e.add_field(name="🎖️ ロール・経験値",
        value=f"[ロール＆経験値一覧]({ROLE_LIST})\nロール付与の条件を確認できます\n\n[実績ロール証明所]({ROLE_PROOF})\nロール付与の申請はこちら",
        inline=False)
    e.add_field(name="💡 意見・要望",
        value=f"[意見箱]({IDEA_BOX})\n改善点やお困りごとはこちらへ",
        inline=False)
    e.set_footer(text="◀️ 戻るボタンでメニューに戻れます")
    return e


def tournament_embed():
    e = discord.Embed(title="🏆 大会・イベント関連", color=0xF0B429)
    e.add_field(name="📣 大会アナウンス",
        value=f"[大会アナウンス]({ANNOUNCE_TOUR})\n大会の告知はこちらで確認！",
        inline=False)
    e.add_field(name="🎉 イベントアナウンス",
        value=f"[イベントアナウンス]({ANNOUNCE_EVENT})\nサーバーイベントの告知はこちら",
        inline=False)
    e.add_field(name="📂 大会一覧",
        value=f"[大会一覧カテゴリ]({TOURNAMENT})\n開催前・開催中・直近終了した大会を確認できます",
        inline=False)
    e.add_field(name="🎙️ イベントVC",
        value=f"[イベントVC]({VC_EVENT})\n大会・イベント開催時の通話はこちら",
        inline=False)
    e.set_footer(text="◀️ 戻るボタンでメニューに戻れます")
    return e


def command_embed():
    e = discord.Embed(title="⚙️ コマンド・その他", color=0x7F8C8D)
    e.add_field(name="🟩 便利機能まとめ",
        value=f"[便利機能チャンネル]({BENRI})\nサーバーで使えるコマンドや\n通知機能をまとめています",
        inline=False)
    e.add_field(name="📈 ディスボード上げ",
        value=f"[ディスボード上げチャンネル]({DISBOARD})\n`/bump` と入力するとサーバーの\n優先順位を上げられます！\n※2時間に1回更新可能",
        inline=False)
    e.add_field(name="🍜 飯テロの称号確認",
        value="`/飯テロ称号一覧` … 全称号と次の称号まであと何年か\n`/飯テロカルテ` … 自分の累計と現在の称号",
        inline=False)
    e.add_field(name="🚨 ルール違反の報告",
        value="`/通報` … 運営に報告できます\n内容は他の人には見えません",
        inline=False)
    e.add_field(name="📘 このガイドを表示する",
        value="`/ガイド` と入力するといつでもこのメニューが表示されます",
        inline=False)
    e.set_footer(text="◀️ 戻るボタンでメニューに戻れます")
    return e


# ==============================
# スラッシュコマンド
# ==============================
@bot.tree.command(name="ガイド", description="チャンネルの使い方ガイドを表示します")
async def guide(interaction: discord.Interaction):
    await interaction.response.send_message(embed=main_embed(), view=MainMenuView(), ephemeral=True)


@bot.tree.command(name="vc注意事項プレビュー", description="VC入室時に表示される注意事項を自分だけに表示して確認します")
@app_commands.default_permissions(administrator=True)
async def preview_notice(interaction: discord.Interaction):
    await interaction.response.send_message(embed=build_notice_embed(), ephemeral=True)


@bot.tree.command(name="vc注意事項を出す", description="今いるボイスチャンネルに注意事項を手動で表示します")
@app_commands.default_permissions(administrator=True)
async def post_notice(interaction: discord.Interaction):
    if interaction.user.voice is None or interaction.user.voice.channel is None:
        await interaction.response.send_message(
            "先にボイスチャンネルに入ってから実行してください。", ephemeral=True
        )
        return

    channel = interaction.user.voice.channel
    view = build_notice_view()
    if view is not None:
        await channel.send(embed=build_notice_embed(), view=view)
    else:
        await channel.send(embed=build_notice_embed())
    await interaction.response.send_message(
        f"{channel.name} に表示しました。", ephemeral=True
    )


@bot.tree.command(name="通報", description="ルール違反を運営に報告します（内容は他の人に見えません）")
async def report(interaction: discord.Interaction):
    mode = _effective_mode()
    if mode == "form":
        await interaction.response.send_modal(ReportModal())
    elif mode == "ticket":
        await interaction.response.send_message(
            "ルール違反の報告は意見箱で受け付けています。\n"
            f"{REPORT_TICKET_URL} からチケットを作成してください。\n"
            "チケット内のやり取りは、運営とあなたにしか見えません。",
            ephemeral=True,
        )
    else:
        await interaction.response.send_message(
            "現在この機能は準備中です。お手数ですが、しろくさへ直接ご連絡ください。",
            ephemeral=True,
        )


# ==============================
# 起動処理
# ==============================
@bot.event
async def on_ready():
    guild_obj = discord.Object(id=GUILD_ID)

    # 管理者用Cogを登録（重複登録を防ぐ）
    try:
        if bot.get_cog("JohoListCog") is None:
            await bot.add_cog(JohoListCog(bot), guild=guild_obj)
        if bot.get_cog("KouryakuMiningCog") is None:
            await bot.add_cog(KouryakuMiningCog(bot), guild=guild_obj)
        if bot.get_cog("KokenRankCog") is None:
            await bot.add_cog(KokenRankCog(bot), guild=guild_obj)
        if bot.get_cog("SensekiCog") is None:
            await bot.add_cog(SensekiCog(bot), guild=guild_obj)
        if bot.get_cog("PremierSeriesNotifier") is None:
            await bot.add_cog(PremierSeriesNotifier(bot), guild=guild_obj)
    except Exception as e:
        print(f"⚠️ Cog登録に失敗しました: {e}")

    # グローバルコマンドをギルドへコピーして即時同期
    try:
        bot.tree.copy_global_to(guild=guild_obj)
        synced = await bot.tree.sync(guild=guild_obj)
        print(f"✅ コマンド同期完了: {len(synced)} 件")
        for cmd in synced:
            print(f"   - /{cmd.name}")
    except Exception as e:
        print(f"⚠️ コマンド同期に失敗しました: {e}")

    bot.add_view(MainMenuView())
    bot.add_view(BackView())
    bot.add_view(ReportView())

    print(f"✅ BOT起動完了: {bot.user}")


# ==============================
# トークン（環境変数から読み込む）
# ==============================
token = os.environ.get("DISCORD_TOKEN")
if not token:
    print("❌ DISCORD_TOKENが設定されていません")
else:
    bot.run(token)
