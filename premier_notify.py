import discord
from discord.ext import tasks, commands
from discord import app_commands
import json
from datetime import datetime, timedelta
import asyncio

class PremierSeriesNotifier(commands.Cog):
    """Shadowverse Premier Series の試合イベント自動作成・リマインド"""
    
    def __init__(self, bot):
        self.bot = bot
        self.guild_id = 1194515135071539210   # サーバーID
        self.channel_id = 1540616850293923991  # 📺-プレミアシリーズ
        self.role_id = 1540617363777388585      # プレミアシリーズ観戦者
        self.schedule_file = 'data/premier_schedule.json'
        self.created_events_file = 'data/created_events.json'
        self.official_url = 'https://ps.shadowverse-wb.com/26-27/schedule-results/'
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
            return {}
        except json.JSONDecodeError:
            return {}
    
    def _save_created_events(self):
        """作成済みイベント一覧を保存"""
        try:
            with open(self.created_events_file, 'w', encoding='utf-8') as f:
                json.dump(self.created_events, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ イベント記録保存エラー: {e}")
    
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
    
    async def get_upcoming_matches(self):
        """本日と明日の試合を取得（リマインド用）"""
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
    
    async def create_event_for_match(self, round_data):
        """試合情報からDiscordイベントを作成"""
        guild = self.bot.get_guild(self.guild_id)
        if not guild:
            print(f"❌ Guild {self.guild_id} not found")
            return None
        
        round_num = round_data['round']
        title = round_data['title']
        date_str = round_data['date']
        time_str = round_data.get('time', '09:00')
        
        # イベント作成済みチェック
        event_key = f"{round_num}_{date_str}"
        if event_key in self.created_events:
            print(f"⚠️ イベント既作成スキップ: {title}")
            return None
        
        try:
            # 開始時刻を計算
            start_time = datetime.strptime(f"{date_str} {time_str}", '%Y-%m-%d %H:%M')
            
            # マッチ情報を説明に追加
            description = ""
            match_count = 1
            while f'match_{match_count}' in round_data:
                match = round_data[f'match_{match_count}']
                team_a = match['team_a']
                team_b = match['team_b']
                description += f"ROUND {match_count}: {team_a} vs {team_b}\n"
                match_count += 1
            
            description += f"\n🔗 [公式ページで配信を確認]({self.official_url})"
            
            # イベント作成
            event = await guild.create_scheduled_event(
                name=f"📺 プレミアシリーズ {title}",
                description=description,
                start_time=start_time,
                end_time=start_time + timedelta(hours=4),  # 4時間後を終了時刻
                entity_type=discord.ScheduledEventEntityType.external,
                location=self.official_url
            )
            
            self.created_events[f"{round_num}_{date_str}"] = event.id
            self._save_created_events()
            print(f"✅ イベント作成: {title} (ID: {event.id})")
            return event
        except Exception as e:
            print(f"❌ イベント作成エラー: {e}")
            return None
    
    async def send_monthly_events(self):
        """月初・月中：1ヶ月分のイベント一括作成"""
        matches = await self.get_next_30days_matches()
        if not matches:
            print("📺 プレミアシリーズ: 今後30日間に試合がありません")
            return 0, 0
        
        created_count = 0
        skipped_count = 0
        
        # 各試合を処理
        for round_data in matches:
            round_num = round_data['round']
            title = round_data['title']
            date_str = round_data['date']
            
            event_key = f"{round_num}_{date_str}"
            
            # イベント作成済みチェック
            if event_key in self.created_events:
                skipped_count += 1
                continue
            
            # イベント作成
            result = await self.create_event_for_match(round_data)
            if result:
                created_count += 1
        
        return created_count, skipped_count
    
    async def send_match_reminders(self):
        """毎日のリマインド：本日・明日に試合がある場合"""
        matches = await self.get_upcoming_matches()
        if not matches:
            return
        
        channel = self.bot.get_channel(self.channel_id)
        if not channel:
            print(f"❌ Channel {self.channel_id} not found")
            return
        
        # 各試合を処理
        for round_data in matches:
            title = round_data['title']
            date_str = round_data['date']
            time_str = round_data.get('time', 'TBD')
            match_type = round_data.get('match_type', '未定')
            
            # Embed を作成
            if match_type == '前日':
                embed_title = f"📺 プレミアシリーズ {title} 【明日開始】"
                embed_desc = f"⏰ 明日 **{time_str}** に開始予定\n🔗 [公式ページで配信を確認]({self.official_url})"
            else:
                embed_title = f"📺 プレミアシリーズ {title} 【本日開始】"
                embed_desc = f"⏰ 本日 **{time_str}** に開始予定\n🔗 [公式ページで配信を確認]({self.official_url})"
            
            embed = discord.Embed(
                title=embed_title,
                description=embed_desc,
                color=0x1e90ff  # Dodger Blue
            )
            
            # マッチ情報を追加
            match_count = 1
            while f'match_{match_count}' in round_data:
                match = round_data[f'match_{match_count}']
                team_a = match['team_a']
                team_b = match['team_b']
                
                match_info = f"**{team_a}** vs **{team_b}**"
                
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
                print(f"✅ プレミアシリーズリマインド送信: {title} 【{match_type}】")
            except Exception as e:
                print(f"❌ リマインド送信エラー: {e}")
    
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
            description=f"📅 **{date_str}** ⏰ **{time_str}**\n🔗 [公式ページで配信を確認]({self.official_url})",
            color=0x1e90ff
        )
        
        # マッチ情報を追加
        match_count = 1
        while f'match_{match_count}' in next_match:
            match = next_match[f'match_{match_count}']
            team_a = match['team_a']
            team_b = match['team_b']
            
            match_info = f"**{team_a}** vs **{team_b}**"
            
            embed.add_field(
                name=f"ROUND {match_count}",
                value=match_info,
                inline=False
            )
            match_count += 1
        
        embed.set_footer(text="Shadowverse Premier Series 26-27")
        
        await interaction.followup.send(embed=embed)
    
    @app_commands.command(
        name="プレミアイベント作成",
        description="プレミアシリーズの1ヶ月分イベントを作成・更新します（手動実行）"
    )
    async def manual_create_premier_events(self, interaction: discord.Interaction):
        """手動でイベント一括作成"""
        await interaction.response.defer()
        
        channel = self.bot.get_channel(self.channel_id)
        if not channel:
            await interaction.followup.send(
                "❌ チャンネルが見つかりません",
                ephemeral=True
            )
            return
        
        # イベント作成実行
        created_count, skipped_count = await self.send_monthly_events()
        
        # 結果報告
        if created_count == 0 and skipped_count == 0:
            embed = discord.Embed(
                title="❌ プレミアシリーズ イベント作成",
                description="今後30日間に試合がありません",
                color=0xFF6B6B
            )
        else:
            embed = discord.Embed(
                title="✅ プレミアシリーズ イベント作成完了",
                description=f"✅ 新規作成: {created_count} 件\n⏭️ スキップ（既作成）: {skipped_count} 件",
                color=0x1e90ff
            )
        
        embed.set_footer(text="Shadowverse Premier Series 26-27")
        
        await interaction.followup.send(embed=embed)
        
        # チャンネルにも報告
        try:
            await channel.send(embed=embed)
        except Exception as e:
            print(f"⚠️ チャンネル報告エラー: {e}")
    
    @tasks.loop(hours=24)
    async def monthly_create(self):
        """毎月1日・15日 09:00 にイベント一括作成"""
        today = datetime.now()
        if today.day in [1, 15]:  # 1日または15日
            created_count, skipped_count = await self.send_monthly_events()
            
            channel = self.bot.get_channel(self.channel_id)
            if channel and (created_count > 0 or skipped_count > 0):
                try:
                    embed = discord.Embed(
                        title="📅 プレミアシリーズ イベント一括作成完了",
                        description=f"✅ 新規作成: {created_count} 件\n⏭️ スキップ（既作成）: {skipped_count} 件",
                        color=0x1e90ff
                    )
                    embed.set_footer(text="Shadowverse Premier Series 26-27")
                    await channel.send(embed=embed)
                except Exception as e:
                    print(f"❌ 完了報告エラー: {e}")
    
    @monthly_create.before_loop
    async def before_monthly_create(self):
        """初回実行まで待機"""
        await self.bot.wait_until_ready()
        
        now = datetime.now()
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        
        if now >= target:
            target = target + timedelta(days=1)
        
        wait_seconds = (target - now).total_seconds()
        print(f"🔔 プレミアシリーズ月次作成: {wait_seconds}秒後に開始")
        await asyncio.sleep(wait_seconds)
    
    @tasks.loop(hours=24)
    async def daily_remind(self):
        """毎日09:00 にリマインド送信"""
        await self.send_match_reminders()
    
    @daily_remind.before_loop
    async def before_daily_remind(self):
        """初回実行まで待機"""
        await self.bot.wait_until_ready()
        
        now = datetime.now()
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        
        if now >= target:
            target = target + timedelta(days=1)
        
        wait_seconds = (target - now).total_seconds()
        print(f"🔔 プレミアシリーズ日次リマインド: {wait_seconds}秒後に開始")
        await asyncio.sleep(wait_seconds)
