import discord
from discord.ext import tasks, commands
from discord import app_commands
import json
from datetime import datetime, timedelta
import asyncio
import os

class PremierSeriesNotifier(commands.Cog):
    """Shadowverse Premier Series の試合イベント自動作成・リマインド"""
    
    def __init__(self, bot):
        self.bot = bot
        self.guild_id = 1194515135071539210
        self.channel_id = 1540616850293923991
        self.role_id = 1540617363777388585
        self.schedule_file = 'data/premier_schedule.json'
        self.created_events_file = 'data/created_events.json'
        self.official_url = 'https://ps.shadowverse-wb.com/26-27/schedule-results/'
        
        # デバッグ出力
        print(f"🔍 ワーキングディレクトリ: {os.getcwd()}")
        print(f"🔍 スケジュールファイルパス: {os.path.abspath(self.schedule_file)}")
        print(f"🔍 ファイル存在確認: {os.path.exists(self.schedule_file)}")
        
        self.created_events = self._load_created_events()
        self.monthly_create.start()
        self.daily_remind.start()
    
    def cog_unload(self):
        self.monthly_create.cancel()
        self.daily_remind.cancel()
    
    def _load_created_events(self):
        """作成済みイベント一覧を読み込む"""
        try:
            with open(self.created_events_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"⚠️ {self.created_events_file} は作成されていません（初回時の正常な状態）")
            return {}
        except json.JSONDecodeError:
            return {}
    
    def _save_created_events(self):
        """作成済みイベント一覧を保存"""
        try:
            os.makedirs('data', exist_ok=True)
            with open(self.created_events_file, 'w', encoding='utf-8') as f:
                json.dump(self.created_events, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ イベント記録保存エラー: {e}")
    
    async def load_schedule(self):
        """JSON ファイルから試合スケジュールを読み込む"""
        try:
            with open(self.schedule_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"✅ スケジュール読み込み成功: {len(data['matches'])} 試合")
                return data
        except FileNotFoundError:
            print(f"❌ {self.schedule_file} が見つかりません")
            print(f"🔍 存在するファイル: {os.listdir('data') if os.path.exists('data') else 'data フォルダなし'}")
            return None
        except json.JSONDecodeError as e:
            print(f"❌ JSON フォーマットエラー: {e}")
            return None
    
    async def get_next_30days_matches(self):
        """今日から30日間の試合を取得"""
        schedule = await self.load_schedule()
        if not schedule:
            return None
        
        today = datetime.now()
        thirty_days_later = today + timedelta(days=30)
        
        upcoming_matches = []
        for round_data in schedule['matches']:
            match_date = datetime.strptime(round_data['date'], '%Y-%m-%d')
            if today <= match_date <= thirty_days_later:
                upcoming_matches.append(round_data)
        
        return upcoming_matches if upcoming_matches else None
    
    async def get_next_match(self):
        """直近の試合を取得"""
        schedule = await self.load_schedule()
        if not schedule:
            return None
        
        today = datetime.now()
        for round_data in schedule['matches']:
            match_date = datetime.strptime(round_data['date'], '%Y-%m-%d')
            if match_date >= today:
                return round_data
        return None
    
    @app_commands.command(
        name="プレミアイベント作成",
        description="プレミアシリーズの1ヶ月分イベントを作成・更新します（手動実行）"
    )
    async def manual_create_premier_events(self, interaction: discord.Interaction):
        """手動でイベント一括作成"""
        await interaction.response.defer()
        
        # スケジュール読み込みテスト
        schedule = await self.load_schedule()
        
        if not schedule:
            await interaction.followup.send(
                "❌ スケジュールファイルが読み込めません\n"
                f"📁 ワーキングディレクトリ: {os.getcwd()}\n"
                f"🔗 ファイルパス: {os.path.abspath(self.schedule_file)}",
                ephemeral=True
            )
            return
        
        await interaction.followup.send(
            f"✅ スケジュール読み込み成功\n"
            f"📊 試合数: {len(schedule['matches'])}\n"
            f"🔗 ファイルパス: {os.path.abspath(self.schedule_file)}",
            ephemeral=True
        )
    
    @tasks.loop(hours=24)
    async def monthly_create(self):
        """毎月1日・15日 09:00 にイベント一括作成"""
        pass
    
    @monthly_create.before_loop
    async def before_monthly_create(self):
        """初回実行まで待機"""
        await self.bot.wait_until_ready()
        now = datetime.now()
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if now >= target:
            target = target + timedelta(days=1)
        wait_seconds = (target - now).total_seconds()
        await asyncio.sleep(wait_seconds)
    
    @tasks.loop(hours=24)
    async def daily_remind(self):
        """毎日09:00 にリマインド送信"""
        pass
    
    @daily_remind.before_loop
    async def before_daily_remind(self):
        """初回実行まで待機"""
        await self.bot.wait_until_ready()
        now = datetime.now()
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if now >= target:
            target = target + timedelta(days=1)
        wait_seconds = (target - now).total_seconds()
        await asyncio.sleep(wait_seconds)
