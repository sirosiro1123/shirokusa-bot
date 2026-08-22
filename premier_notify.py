import discord
from discord.ext import tasks, commands
from discord import app_commands
import json
from datetime import datetime, timedelta
import asyncio

class PremierSeriesNotifier(commands.Cog):
    """Shadowverse Premier Series の試合日程を自動通知（大会前日・当日のみ）"""
    
    def __init__(self, bot):
        self.bot = bot
        self.channel_id = 1540616850293923991  # 📺-プレミアシリーズ
        self.role_id = 1540617363777388585      # プレミアシリーズ観戦者
        self.schedule_file = 'data/premier_schedule.json'
        self.official_url = 'https://ps.shadowverse-wb.com/26-27/schedule-results/'
        self.daily_check.start()
    
    def cog_unload(self):
        self.daily_check.cancel()
    
    async def load_schedule(self):
        """JSON ファイルから試合スケジュールを読み込む"""
        try:
            with open(self.schedule_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"⚠️ Warning: {self.schedule_file} not found")
            return None
        except json.JSONDecodeError:
            print(f"❌ Error: {self.schedule_file} is not valid JSON")
            return None
    
    async def get_upcoming_matches(self):
        """本日と明日の試合を取得"""
        schedule = await self.load_schedule()
        if not schedule:
            return None
        
        today = datetime.now().strftime('%Y-%m-%d')
        tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
        
        upcoming_matches = []
        
        for round_data in schedule['matches']:
            match_date = round_data['date']
            if match_date == today:
                round_data['match_type'] = '当日'
                upcoming_matches.append(round_data)
            elif match_date == tomorrow:
                round_data['match_type'] = '前日'
                upcoming_matches.append(round_data)
        
        return upcoming_matches if upcoming_matches else None
    
    async def get_next_match(self):
        """直近の試合（本日以降の最初の試合）を取得"""
        schedule = await self.load_schedule()
        if not schedule:
            return None
        
        today = datetime.now()
        
        for round_data in schedule['matches']:
            match_date = datetime.strptime(round_data['date'], '%Y-%m-%d')
            if match_date >= today:
                return round_data
        
        return None
    
    async def send_match_notification(self):
        """本日と明日の試合告知を送信"""
        matches = await self.get_upcoming_matches()
        if not matches:
            print("📺 プレミアシリーズ: 本日・明日に試合がありません")
            return
        
        channel = self.bot.get_channel(self.channel_id)
        if not channel:
            print(f"❌ Channel {self.channel_id} not found")
            return
        
        # 各試合を処理
        for round_data in matches:
            round_num = round_data['round']
            title = round_data['title']
            time_str = round_data.get('time', 'TBD')
            match_type = round_data.get('match_type', '未定')
            
            # Embed を作成
            embed = discord.Embed(
                title=f"📺 プレミアシリーズ {title} 【{match_type}】",
                description=f"⏰ **{time_str}** に開始予定\n🔗 [公式ページ]({self.official_url})",
                color=0x1e90ff  # Dodger Blue
            )
            
            # マッチ情報を追加
            match_count = 1
            while f'match_{match_count}' in round_data:
                match = round_data[f'match_{match_count}']
                team_a = match['team_a']
                team_b = match['team_b']
                broadcast = match.get('broadcast', 'TBD')
                
                match_info = f"**{team_a}** vs **{team_b}**"
                if broadcast != 'TBD':
                    match_info += f"\n🔗 [配信を見る]({broadcast})"
                else:
                    match_info += f"\n📡 配信URL未定"
                
                embed.add_field(
                    name=f"ROUND {match_count}",
                    value=match_info,
                    inline=False
                )
                match_count += 1
            
            embed.set_footer(text="Shadowverse Premier Series 26-27")
            
            # ロールを取得してメンション
            guild = channel.guild
            role = guild.get_role(self.role_id)
            mention_text = f"{role.mention}" if role else ""
            
            # 送信
            try:
                await channel.send(mention_text, embed=embed)
                print(f"✅ プレミアシリーズ通知送信: Round {round_num} ({title}) 【{match_type}】")
            except Exception as e:
                print(f"❌ 通知送信エラー: {e}")
    
    # ============================================================
    # スラッシュコマンド
    # ============================================================
    @app_commands.command(
        name="プレミア日程",
        description="直近のプレミアシリーズ試合日程を表示します"
    )
    async def premier_schedule(self, interaction: discord.Interaction):
        """直近の試合情報を表示"""
        await interaction.response.defer()
        
        next_match = await self.get_next_match()
        if not next_match:
            await interaction.followup.send(
                "❌ 試合スケジュールが見つかりません。\n"
                f"🔗 [公式ページで確認]({self.official_url})",
                ephemeral=True
            )
            return
        
        round_num = next_match['round']
        title = next_match['title']
        date_str = next_match['date']
        time_str = next_match.get('time', 'TBD')
        
        # Embed を作成
        embed = discord.Embed(
            title=f"📺 プレミアシリーズ {title}",
            description=f"📅 **{date_str}** ⏰ **{time_str}**\n🔗 [公式ページ]({self.official_url})",
            color=0x1e90ff
        )
        
        # マッチ情報を追加
        match_count = 1
        while f'match_{match_count}' in next_match:
            match = next_match[f'match_{match_count}']
            team_a = match['team_a']
            team_b = match['team_b']
            broadcast = match.get('broadcast', 'TBD')
            
            match_info = f"**{team_a}** vs **{team_b}**"
            if broadcast != 'TBD':
                match_info += f"\n🔗 [配信を見る]({broadcast})"
            else:
                match_info += f"\n📡 配信URL未定"
            
            embed.add_field(
                name=f"ROUND {match_count}",
                value=match_info,
                inline=False
            )
            match_count += 1
        
        embed.set_footer(text="Shadowverse Premier Series 26-27")
        
        await interaction.followup.send(embed=embed)
    
    @tasks.loop(hours=24)
    async def daily_check(self):
        """毎日09:00に実行"""
        await self.send_match_notification()
    
    @daily_check.before_loop
    async def before_daily_check(self):
        """初回実行まで待機"""
        await self.bot.wait_until_ready()
        
        # 次の09:00まで待機
        now = datetime.now()
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        
        # 既に09:00を過ぎていれば、翌日の09:00に設定
        if now >= target:
            target = target + timedelta(days=1)
        
        wait_seconds = (target - now).total_seconds()
        print(f"🔔 プレミアシリーズ通知チェック: {wait_seconds}秒後に開始")
        await asyncio.sleep(wait_seconds)
