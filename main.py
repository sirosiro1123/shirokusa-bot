import discord
from discord.ext import commands
import os
from dotenv import load_dotenv

# 環境変数を読み込む
load_dotenv()

# Intents を設定
intents = discord.Intents.default()
intents.message_content = True
intents.voice_states = True
intents.members = True

# BOT を初期化
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    """BOT が起動したときに実行"""
    print(f"✅ BOT起動: {bot.user}")
    
    try:
        synced = await bot.tree.sync()
        print(f"✅ スラッシュコマンド同期: {len(synced)} 個")
    except Exception as e:
        print(f"❌ スラッシュコマンド同期エラー: {e}")

# Cog を読み込む
async def load_cogs():
    """Cog ファイルを読み込む"""
    cog_files = [
        'vc_notice',
        'joho_list', 
        'kouryaku_mining',
        'premier_notify'
    ]
    
    for cog in cog_files:
        try:
            await bot.load_extension(f'{cog}')
            print(f"✅ Cog 読み込み: {cog}")
        except Exception as e:
            print(f"❌ Cog 読み込みエラー ({cog}): {e}")

# メイン処理
async def main():
    """BOT を起動"""
    await load_cogs()
    
    token = os.getenv('DISCORD_TOKEN')
    if not token:
        print("❌ DISCORD_TOKEN が設定されていません")
        return
    
    await bot.start(token)

# 実行
if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
