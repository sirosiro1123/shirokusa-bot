"""
senseki.py
シャドバ戦績ツール（フェーズ1：記録できる状態）

詳細仕様は `戦績ツール_実装仕様書.md` を参照。
このファイルはフェーズ1のみを実装しています。

コマンド構成（Discord上のトップレベルは3つだけ）

  /戦績
    記録        対戦結果を記録（ボタン形式）
    取消        直前の1件を取り消し
    確認        今のデッキの直近30戦の勝率（無料）と通算勝率（30戦超・メンバー限定）
    履歴        直近10戦の対戦履歴（自分の分のみ）
    設定        フォーマット・クラス・デッキ・ランクを登録（全項目必須）
    ランク更新   ランク帯とグレードだけ変更（デッキは維持）
    弱点対面     デッキ単位（30戦から解放）：①勝率が低い対面（無料）／②全体平均より見劣りする対面（メンバー限定）
    デッキ推移   今のフォーマットで、使ってきたデッキの変遷と成績
    先後対面     今のデッキで、対面ごとに先攻後攻どちらが得意か
    相手デッキ対面 今のデッキで、相手デッキ別の勝率（十分なデータがあれば先攻後攻も）

  /デッキ
    登録        よく使うデッキを登録（複数可）
    切替        登録済みから選んで切替（ランクは維持）
    検証開始／検証終了／検証集計  個人用デッキバリアント（下記）の管理者フォールバック。
                通常はパネルの「⚙️ 設定を変える」→「🧪 デッキ検証」から使う

  /戦績管理（管理者のみ表示）
    パネル設置    記録・切替・ランク更新・取消・成績確認・弱点対面・デッキ検証の各ボタンを常設
    掲示板設置    全員の使用デッキ＆ランクの一覧を常設（設定変更時に自動更新）
    全体集計      その瞬間のDBから全体・先後別・クラス別・対面別を表示
    環境分析画像  クラス使用率・対面マトリクス・先攻後攻をPNG画像で表示（SENSEKI_ADMIN_USER_ID本人のみ）
    環境レポート設置／投稿／停止  同じ画像＋AI分析を2日に1回12時にサーバーへ自動投稿する
    データ抽出    生データ＋集計のExcelを出力（毎週月曜9:00にDMへ自動送信）
    シート同期    Googleスプレッドシートへ即時同期
    シート診断    スプレッドシート連携の設定状況
    診断          DBの保存先と永続化状況（データが消える事故の確認用）
    権限確認      指定チャンネルでBOTが持っている権限を確認
    分析利用状況  確認・弱点対面などの分析コマンドの利用回数・利用者数
    デッキ名統合  表記ゆれのあるデッキ名を1つに統合（過去の記録も一括書き換え）
    無料トライアル開始／無料トライアル状況  初回参加者向け無料キャンペーンの開始とON後の状況確認
    記事登録／記事一覧／記事削除  対面ごとのおすすめ記事URLの管理
    質問ヒント設置／投稿／停止  毎週金曜20時に「質問のヒント」を自動投稿（AIなし）。ボタンで個人用ヒントとクラス別質問
    質問先設定    クラスごとの質問の投稿先チャンネルを設定
    記録管理      修正・削除の履歴確認／削除済み記録の確認／IDを指定した復元

【データの保存先について】
  Railwayの永続ボリュームのマウント先はプロジェクトによって異なる
  （/data のことも /app/data のこともある）。間違った場所に書くと
  書き込み自体は成功するのに再デプロイで消えるため、気づきにくい。
  そのため _resolve_db_path() が
    1. 環境変数 SENSEKI_DB_PATH（最優先）
    2. すでにDBファイルが存在する場所
    3. マウント済み（＝永続）のディレクトリ
  の順に探す。起動のたびに保存先と永続化状況をログへ出し、
  永続でなければ警告を出す。/戦績管理 診断 でも同じ情報を確認できる。

【有料化スイッチについて（/弱点対面 の公開範囲）】
  現状は誰でも使える。有料化のタイミングで Railway の環境変数に
    SENSEKI_MEMBERS_ONLY=true
  を追加して再デプロイするだけで、コードを触らずメンバー限定に切り替わる。
  メンバー判定は MEMBER_ROLE_IDS（カンマ区切りの環境変数、飯テロBOTと同じ方式）。
  スイッチをオンにする前に MEMBER_ROLE_IDS が未設定だと、誰も条件を満たせず
  実質「使用不可」になってしまうので、オンにするときは MEMBER_ROLE_IDS が
  正しく設定されていることを先に確認すること。

フェーズ2以降で追加するもの（このファイルには未実装）
  - 環境IDは ENV_SCHEDULE の開始日時で自動切替（/環境切替 コマンドは未実装）
  - デッキタイプのオートコンプリート・表記揺れ統合（/デッキ名整理）
  - /戦績比較（メンバー限定・全体勝率との比較）
  - ランク帯別の集計（データは記録済みなので後から出せる）

main.py への統合方法は末尾のコメントを参照。

【デプロイ後に必ず確認すること】
Railwayのログに出る「戦績ツール ストレージ診断」を見て、
「永続ボリュームか: はい」になっていることを確認してください。
「いいえ」の場合は再デプロイでデータが消えます。
"""

import asyncio
import json
import io
import os
import sqlite3
from datetime import datetime, time as dtime, timedelta, timezone

from typing import Union

import discord

try:  # AI分析（環境レポート用）。未導入でもBOTは止めない
    from anthropic import AsyncAnthropic
    ANTHROPIC_IMPORT_ERROR = None
except Exception as _e:
    AsyncAnthropic = None
    ANTHROPIC_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"
from discord import app_commands
from discord.ext import commands, tasks

try:
    from openpyxl import Workbook
    OPENPYXL_IMPORT_ERROR = None
except Exception as _e:  # ライブラリ未導入時に起動を止めない
    Workbook = None
    OPENPYXL_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"

try:
    import senseki_sheets
    SHEETS_MODULE_ERROR = None
except Exception as _e:  # モジュール未配置でも起動を止めない
    senseki_sheets = None
    SHEETS_MODULE_ERROR = f"{type(_e).__name__}: {_e}"

try:
    import matplotlib
    matplotlib.use("Agg")  # Railwayに画面がないためGUIバックエンドは使わない
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    import matplotlib.colors as mcolors
    import numpy as np  # 対面マトリクスのヒートマップ生成に使う（matplotlibの依存なので追加コストなし）
    MATPLOTLIB_IMPORT_ERROR = None
except Exception as _e:  # ライブラリ未導入時に起動を止めない（グラフなしでテキストのみ動く）
    plt = None
    MATPLOTLIB_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"

# ---- 弱点対面グラフ用の日本語フォント ----
# assets/NotoSansJP-Regular.otf を同梱（Noto Sans CJK JPのJapaneseサブセット・OFLライセンス）。
# matplotlibは標準で日本語グリフを持たないため、明示的にフォントファイルを読み込む。
# フォントが無い/読み込み失敗した場合は _JP_FONT_NAME が None のままとなり、
# グラフ生成自体をスキップしてテキストのみ送る（デグレードするが起動は止めない）。
_JP_FONT_NAME = None
if plt is not None:
    # 同梱フォントに太字が無いため、太字指定のたびに「通常の太さで代用」の警告がログに出る。
    # 画像は問題なく作られるので、この警告だけ出さないようにする
    import logging as _logging
    _logging.getLogger("matplotlib.font_manager").setLevel(_logging.ERROR)
    try:
        _FONT_PATH = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "assets", "NotoSansJP-Regular.otf"
        )
        if os.path.exists(_FONT_PATH):
            fm.fontManager.addfont(_FONT_PATH)
            _JP_FONT_NAME = fm.FontProperties(fname=_FONT_PATH).get_name()
            matplotlib.rcParams["font.family"] = _JP_FONT_NAME
            matplotlib.rcParams["axes.unicode_minus"] = False
    except Exception:
        _JP_FONT_NAME = None

# ==============================
# 設定
# ==============================

GUILD_ID = 1194515135071539210

# ---- DBの置き場所 ----
# Railwayの永続ボリュームはプロジェクトによってマウント先が違う。
# /data のこともあれば /app/data のこともあり、間違えると再デプロイのたびに
# データが消える（書き込みは成功するので気づきにくい）。
# そのため「明示指定 → 既存DBのある場所 → マウント済みの場所」の順で探す。

DB_FILENAME = "senseki.db"

# 探索候補。左が優先。
DB_CANDIDATE_DIRS = ["/data", "/app/data"]


def _is_persistent(path: str) -> bool:
    """そのディレクトリが永続ボリュームとしてマウントされているか"""
    try:
        return os.path.ismount(path)
    except Exception:
        return False


def _resolve_db_path() -> str:
    # 1. 環境変数で明示指定されていればそれに従う（最優先）
    explicit = os.environ.get("SENSEKI_DB_PATH")
    if explicit:
        return explicit

    # 2. すでにDBファイルが存在する場所があればそこを使う（データを見失わない）
    for d in DB_CANDIDATE_DIRS:
        candidate = os.path.join(d, DB_FILENAME)
        if os.path.exists(candidate):
            return candidate

    # 3. マウント済み（＝永続化される）ディレクトリを優先して新規作成
    for d in DB_CANDIDATE_DIRS:
        if _is_persistent(d):
            return os.path.join(d, DB_FILENAME)

    # 4. どれも該当しなければ既定値。永続化されない可能性があるため起動時に警告する
    return os.path.join(DB_CANDIDATE_DIRS[0], DB_FILENAME)


DB_PATH = _resolve_db_path()


def _list_real_mounts() -> list:
    """/proc/mounts から、データ保存に使えそうなマウントを拾う（診断用）

    os.path.ismount() だけだと環境によって判定を外すことがあるため、
    実際のマウント一覧も併せて見せて、人間が最終判断できるようにする。
    """
    results = []
    try:
        with open("/proc/mounts", encoding="utf-8", errors="replace") as f:
            for line in f:
                parts = line.split()
                if len(parts) < 3:
                    continue
                device, mount_point, fstype = parts[0], parts[1], parts[2]
                # 仮想FSや明らかにシステム用のものは除外
                if fstype in ("proc", "sysfs", "devpts", "tmpfs", "cgroup", "cgroup2",
                              "mqueue", "devtmpfs", "securityfs", "overlay"):
                    continue
                if mount_point in ("/", "/etc/hosts", "/etc/hostname", "/etc/resolv.conf"):
                    continue
                results.append(f"{mount_point}  ({fstype} / {device})")
    except Exception as e:
        results.append(f"（マウント情報を読めませんでした: {type(e).__name__}）")
    return results


def describe_storage() -> str:
    """DBの保存先と永続化状況を人間が読める形で返す（診断用）"""
    lines = [f"DBパス: {DB_PATH}"]
    directory = os.path.dirname(DB_PATH)
    mounted = _is_persistent(directory)
    lines.append(f"保存先ディレクトリ: {directory}")
    lines.append(f"永続ボリュームか: {'はい' if mounted else '❌ いいえ（再デプロイで消えます）'}")

    if os.path.exists(DB_PATH):
        size = os.path.getsize(DB_PATH)
        lines.append(f"DBファイル: あり（{size:,} バイト）")
    else:
        lines.append("DBファイル: まだありません（初回起動時は正常）")

    lines.append("")
    lines.append("候補ディレクトリの状況:")
    for d in DB_CANDIDATE_DIRS:
        exists = os.path.isdir(d)
        mark = "マウント済み" if _is_persistent(d) else ("存在するが非マウント" if exists else "存在しない")
        has_db = "／DBあり" if os.path.exists(os.path.join(d, DB_FILENAME)) else ""
        lines.append(f"  {d}: {mark}{has_db}")

    lines.append("")
    mounts = _list_real_mounts()
    if mounts:
        lines.append("実際にマウントされている場所:")
        for m in mounts:
            lines.append(f"  {m}")
    else:
        lines.append("実際にマウントされている場所: なし（＝ボリューム未接続）")
    return "\n".join(lines)


JST = timezone(timedelta(hours=9))

# クラスの選択肢（仕様書 2-4）
CLASS_CHOICES = ["エルフ", "ロイヤル", "ウィッチ", "ドラゴン", "ナイトメア", "ビショップ", "ネメシス"]

# フォーマットの選択肢（仕様書 2-5）
FORMAT_LABELS = {"rotation": "ローテーション", "unlimited": "アンリミテッド"}

# ランク帯の選択肢
# ゲーム内では各帯が0〜3に細分化されているが（D0〜D3など）、
# 集計の粒度としては細かすぎるため帯単位でまとめている。
# マスターの上に「グランドマスター」があり、ここに到達すると
# クラス別レーティング（CR）による細分化に切り替わる。
RANK_TIER_CHOICES = [
    "ビギナー",
    "D帯",
    "C帯",
    "B帯",
    "A帯",
    "AA帯",
    "マスター",
    "グランドマスター",
]

# グランドマスター帯のみで使うCRグレード
# EPIC 1650〜1749 / ULTIMATE 1750〜1849 / LEGEND 1850以上 /
# BEYOND 1850以上かつCRランキング100位以内
# CRはクラス別に算出されるため、/戦績設定 のクラスと対応する値を入れる。
CR_GRADE_CHOICES = ["EPIC", "ULTIMATE", "LEGEND", "BEYOND"]

# 「グレードなし」を表す選択肢の値
# グランドマスター昇格直後のCR初期値は1100〜1600で、EPICの下限1650に届かない。
# つまりグレードなしは例外ではなく通常の状態なので、明示的に選べるようにする。
# 未入力での代用はしない（入力し忘れと区別できなくなるため）。
CR_GRADE_NONE_VALUE = "__none__"
CR_GRADE_NONE_LABEL = "グレードなし（CR1650未満）"

# CRグレードを入力できるランク帯
GRAND_MASTER_TIER = "グランドマスター"

# 環境ID の一覧（古い順）。各環境は「開始日時」になった瞬間に自動で切り替わる。
# 新弾のたびに、ここへ1行追加して再デプロイする（開始日時の前に反映しておくこと）。
# 切り替わると BOT が1分以内に検知し、前の環境のデータをExcelで管理者にDMする。
# 過去のデータは消さない（環境IDで分けて保存しているだけ）。
ENV_SCHEDULE = [
    ("2026-08-beyond-2", datetime(2026, 8, 27, 0, 0, tzinfo=JST)),
    # 新弾（2026-10-28）のメンテナンス開始時刻で切り替え
    ("2026-10-beyond-3", datetime(2026, 10, 28, 10, 0, tzinfo=JST)),
]


def current_env_id(now: datetime = None) -> str:
    """今の環境ID。ENV_SCHEDULE の開始日時を過ぎた最後の環境を返す"""
    now = now or datetime.now(JST)
    env = ENV_SCHEDULE[0][0]
    for env_id, start in ENV_SCHEDULE:
        if now >= start:
            env = env_id
    return env


def previous_env_id(now: datetime = None):
    """1つ前の環境ID（最初の環境なら None）"""
    ids = [e for e, _ in ENV_SCHEDULE]
    i = ids.index(current_env_id(now))
    return ids[i - 1] if i > 0 else None


# 環境の切り替わりを検知するために、最後に確認した環境IDを bot_state に保存する
ENV_STATE_KEY = "last_env_id"

# 集計データの定期抽出先（管理者DM）
# しろくさのユーザーID（koken_rank.py の ADOPT_JUDGE_USER_ID と同じ値）
SENSEKI_ADMIN_USER_ID = 346475730458378240

# 定期抽出のタイミング（JST）
WEEKLY_EXPORT_WEEKDAY = 0   # 0=月曜
WEEKLY_EXPORT_HOUR = 9
WEEKLY_EXPORT_MINUTE = 0

# 環境レポート（サーバーへの自動投稿）のタイミング（JST）
# 投稿先は /戦績管理 環境レポート設置 を実行したチャンネル（未設置なら投稿しない）
# 毎日 ENV_REPORT_HOUR:ENV_REPORT_MINUTE に確認し、前回の投稿から
# ENV_REPORT_INTERVAL_DAYS 日以上たっていれば投稿する（手動投稿もここに数える）
ENV_REPORT_INTERVAL_DAYS = 2
ENV_REPORT_HOUR = 12
ENV_REPORT_MINUTE = 0
# 現環境の記録がこれ未満の週は投稿しない（少なすぎると偏った数字が出るため）
ENV_REPORT_MIN_MATCHES = 30
ENV_REPORT_CHANNEL_KEY = "env_report_channel_id"
ENV_REPORT_PANEL_KEY = "env_report_panel_channel_id"
# レポートの感想・考察を書いてもらうチャンネル（任意。設置コマンドで指定）
ENV_REPORT_DISCUSS_KEY = "env_report_discuss_channel_id"
# 前回投稿時点の「環境ID|試合数|記録者数」。前回からの増加分を出すのに使う
ENV_REPORT_LAST_KEY = "env_report_last"
# 前回投稿した日（JSTの日付 YYYY-MM-DD）
ENV_REPORT_LAST_DATE_KEY = "env_report_last_date"

# ---- 環境レポートのAI分析 ----
# APIキーの環境変数名。攻略情報抽出BOT（kouryaku_mining.py）と同じキーを使う想定
ENV_REPORT_AI_KEY_ENV = "ANTHROPIC_API_KEY"
ENV_REPORT_AI_MODEL = os.environ.get("SENSEKI_AI_MODEL", "claude-haiku-5-5")
# AIに渡す相手デッキの件数（当たった回数の多い順）
ENV_REPORT_AI_TOP_DECKS = 10
# 「直近」とみなす日数。直近と、それより前の記録を比べて環境の変化を見る
ENV_REPORT_AI_RECENT_DAYS = 7
# 環境レポートの投稿ごとに、感想・考察用のスレッドを自動で立てるか
ENV_REPORT_CREATE_THREAD = False
# 個人のAI分析（パネルの「🤖 AIで分析」）
# 押すたびにAPI費用がかかるため1人1日1回まで。モデルは環境レポートとは別に指定できる
AI_PERSONAL_MODEL = os.environ.get("SENSEKI_AI_PERSONAL_MODEL", "claude-haiku-5-5")
AI_PERSONAL_MIN_MATCHES = 20      # 今のデッキでこれ未満の記録なら分析しない
AI_PERSONAL_MIN_PAIR = 5          # 対面ごとの数字はこれ以上の試合数だけ渡す
AI_PERSONAL_RECENT = 20           # 「直近」とみなす試合数
AI_PERSONAL_STATE_PREFIX = "ai_personal_last_date:"

ENV_REPORT_AI_DISCLAIMER = (
    "-# 🤖 AI分析は開発中の機能です。集計の読み違いや、実際の環境と異なる内容を含む場合があります。"
    "参考程度にご覧ください。"
)

# Googleスプレッドシートへの日次同期のタイミング（JST）
# 環境変数が未設定なら同期はスキップされる（senseki_sheets.py 参照）
SHEETS_SYNC_HOUR = 9
SHEETS_SYNC_MINUTE = 30

# 統計として扱える最低試合数（仕様書 8-3）
# フェーズ3の /戦績全体 ではこの値を下回る対面は数字を出さない
MIN_SAMPLE_FOR_STATS = 30

# クラス対クラスの対面マトリクス（スプレッドシート・環境分析画像）専用のしきい値。
# 個人の弱点対面（30戦）と違い、こちらは母数がサーバー全体の試合を
# 7クラス×7クラス=49マスに分けるため、30戦だと当面ほぼ全マスが空欄になってしまう。
# コミュニティ全体の「雰囲気」を見るための参考表という位置づけなので、
# 個人向けより緩めのしきい値にしている。
MATCHUP_MATRIX_MIN_SAMPLE = 15


# 個人の弱点対面表示の最低試合数
# 30戦は運用初期には現実的でないため、個人向けは緩めにする。
# その代わり画面上に「参考値」であることを明記する。
# /戦績 確認 で無料で見せる「直近」の試合数（シャドバのリプレイ保存数と合わせる）
RECENT_FREE_MATCHES = 30

# 相手デッキ単位の対面：これだけ戦えば「勝率」を表示（弱点対面の対面別しきい値と同じ）
MIN_SAMPLE_OPPDECK = 3
# 相手デッキ単位の対面：先攻・後攻それぞれこれだけ戦えば内訳も追加表示
MIN_SAMPLE_OPPDECK_FS = 5

MIN_SAMPLE_PERSONAL = 3
# 弱点対面ランキング（デッキ単位）：この試合数未満は「参考値」と明示して表示する
# （以前はここに達するまで非表示にしていたが、実際のデータ量では
#   ほぼ全員が使えない状態になったため、参考値扱いで出す方式に変更）
DECK_RANK_MIN_MATCHES = 30
# ②全体平均との比較（有料）：全体側の母数がこれ未満の対面は比較対象から除外
GLOBAL_MATCHUP_MIN_SAMPLE = 10

# メンバー限定機能の判定に使うロールID（カンマ区切り、環境変数）
# 飯テロBOTの MEMBER_ROLE_IDS と同じ方式
def _parse_role_ids(raw: str) -> set:
    ids = set()
    for part in (raw or "").split(","):
        part = part.strip()
        if part.isdigit():
            ids.add(int(part))
    return ids

# ②（全体平均との比較）を、有料化前に環境限定で無料開放する環境IDの一覧。
# ここに入っている環境の間だけ全員無料。新しい環境も無料開放を続けるなら、
# ENV_SCHEDULE に追加した環境IDをここにも足すこと。
PREMIUM_FREE_ENV_IDS = {"2026-08-beyond-2", "2026-10-beyond-3"}


def premium_free_this_env() -> bool:
    return current_env_id() in PREMIUM_FREE_ENV_IDS

# 初回参加者向けの無料キャンペーン：初めて `/戦績 設定` した日から何日間、
# メンバーでなくても②が使えるか（サブ垢は別アカウント＝別トライアルとして扱う）
TRIAL_DAYS = 30

# キャンペーン自体の開始日時（bot_state に保存・管理者コマンドでON）。
# これより「前」に初回登録していた人（＝既存ユーザー）はキャンペーン対象外。
# こうしないと、ONにした瞬間に既存ユーザー全員が「初めての参加者」扱いになってしまう。
TRIAL_CAMPAIGN_START_KEY = "trial_campaign_start_at"

MEMBER_ROLE_IDS = _parse_role_ids(os.environ.get("MEMBER_ROLE_IDS", ""))


def is_member(user: discord.Member) -> bool:
    if not MEMBER_ROLE_IDS:
        return False
    if not isinstance(user, discord.Member):
        return False
    return any(r.id in MEMBER_ROLE_IDS for r in user.roles)


def _parse_iso(value: str):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=JST)
    return dt


def trial_days_left(started_at: str) -> int:
    """started_at（ISO文字列）から見て、無料キャンペーンの残り日数（切り上げ・0未満は0）"""
    start = _parse_iso(started_at)
    if start is None:
        return 0
    elapsed = datetime.now(JST) - start
    remaining = TRIAL_DAYS - elapsed.days
    return max(0, remaining)


def premium_access(user: discord.Member, started_at: str, campaign_start_at: str):
    """②（有料部分）にアクセスできるか判定する。
    戻り値: (使えるか, 理由 "member"|"env_free"|"trial"|None, トライアル残日数 or None)

    トライアルが有効なのは、
      1) キャンペーンが開始済み（campaign_start_at がある）
      2) その人の初回登録（started_at）が、キャンペーン開始日以降
         （＝キャンペーン開始前からの既存ユーザーは対象外）
      3) 初回登録からTRIAL_DAYS以内
    の全てを満たす場合のみ。
    """
    if is_member(user):
        return True, "member", None
    if premium_free_this_env():
        return True, "env_free", None

    campaign_start = _parse_iso(campaign_start_at)
    start = _parse_iso(started_at)
    if campaign_start is None or start is None:
        return False, None, 0
    if start < campaign_start:
        return False, None, 0  # キャンペーン開始前からの既存ユーザーは対象外

    days_left = trial_days_left(started_at)
    if days_left > 0:
        return True, "trial", days_left
    return False, None, 0


# メンバー限定機能の「有料化スイッチ」
# 運用開始時は誰でも使える。有料化のタイミングでRailwayの環境変数に
#   SENSEKI_MEMBERS_ONLY=true
# を追加して再デプロイするだけで、コードを触らずメンバー限定に切り替わる。
# MEMBER_ROLE_IDS 側は事前に設定しておいても、このスイッチがオフの間は無効。
def _parse_bool(raw: str) -> bool:
    return (raw or "").strip().lower() in ("1", "true", "yes", "on")

SENSEKI_MEMBERS_ONLY = _parse_bool(os.environ.get("SENSEKI_MEMBERS_ONLY", ""))
# 環境レポートに「デッキ相性表」の画像を載せるか。記録が増えるまでは載せない（既定）。
# 戻すときはRailwayの変数に SENSEKI_REPORT_DECK_MATRIX=1 を入れる
REPORT_SHOW_DECK_MATRIX = _parse_bool(os.environ.get("SENSEKI_REPORT_DECK_MATRIX", ""))

# 記録・取消のあと、何秒待ってからシートへ反映するか
# 0にすると毎試合すぐ書きに行くが、連戦時にGoogle側のAPI制限に触れる。
# 待つ間に複数の記録が入れば1回の同期にまとめられる（まとめ書き）。
SHEETS_DEBOUNCE_SECONDS = 60

# データが変わったことを同期タスクに知らせるフラグ
# insert_match / delete_last_match から立てる
_sheets_dirty = False


def mark_dirty():
    """戦績データが変わったことを記録する"""
    global _sheets_dirty
    _sheets_dirty = True


def take_dirty() -> bool:
    """フラグを読んで下ろす"""
    global _sheets_dirty
    was = _sheets_dirty
    _sheets_dirty = False
    return was


# ==============================
# DB初期化・アクセス
# ==============================

def _connect() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _init_db_sync():
    conn = _connect()
    try:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS matches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                recorded_at TEXT NOT NULL,
                env_id TEXT NOT NULL,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                my_deck TEXT,
                rank_tier TEXT,
                cr_grade TEXT,
                opp_class TEXT NOT NULL,
                opp_deck TEXT,
                is_first INTEGER NOT NULL,
                is_win INTEGER NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS user_settings (
                user_id TEXT PRIMARY KEY,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                my_deck TEXT,
                rank_tier TEXT,
                cr_grade TEXT,
                updated_at TEXT NOT NULL,
                started_at TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS deck_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                deck_name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(user_id, format, my_class, deck_name)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS deck_names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                class_name TEXT NOT NULL,
                deck_name TEXT NOT NULL,
                use_count INTEGER NOT NULL DEFAULT 0,
                is_official INTEGER NOT NULL DEFAULT 0,
                env_id TEXT NOT NULL
            )
        """)
        # 既に旧スキーマでテーブルが作られている場合に備えた追加（初回は何もしない）
        for table in ("matches", "user_settings"):
            cols = {r["name"] for r in cur.execute(f"PRAGMA table_info({table})")}
            if "cr_grade" not in cols:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN cr_grade TEXT")

        user_cols = {r["name"] for r in cur.execute("PRAGMA table_info(user_settings)")}
        if "started_at" not in user_cols:
            cur.execute("ALTER TABLE user_settings ADD COLUMN started_at TEXT")
            # 既存ユーザーは無料キャンペーン導入前からの利用者。
            # updated_at（最終更新日時）だと直近でデッキを切替えただけの人が
            # 「今日始めた」扱いになってしまうため、その人の一番古い記録日時
            # （＝実際に使い始めた日）で埋める。記録が無ければ updated_at で代用。
            cur.execute("""
                UPDATE user_settings SET started_at = COALESCE(
                    (SELECT MIN(recorded_at) FROM matches WHERE matches.user_id = user_settings.user_id),
                    updated_at
                )
                WHERE started_at IS NULL
            """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS personal_deck_variants (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                match_id INTEGER NOT NULL,
                user_id TEXT NOT NULL,
                variant_label TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS user_variant_state (
                user_id TEXT NOT NULL,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                my_deck TEXT NOT NULL,
                variant_label TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (user_id, format, my_class, my_deck)
            )
        """)
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_personal_variant_user "
            "ON personal_deck_variants(user_id, variant_label)"
        )
        # 1試合につきタグは1つまで（同じ試合に二重で紐付けようとしたら弾く）
        cur.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_personal_variant_match "
            "ON personal_deck_variants(match_id)"
        )

        # 削除した記録の退避先（matches と同じ列＋削除日時）。集計・シート・レポートは matches しか見ないので、
        # ここに移した記録は自動的に集計から外れ、復元すると元の id のまま matches に戻る
        cur.execute("""
            CREATE TABLE IF NOT EXISTS matches_deleted (
                id INTEGER PRIMARY KEY,
                user_id TEXT NOT NULL,
                recorded_at TEXT NOT NULL,
                env_id TEXT NOT NULL,
                format TEXT NOT NULL,
                my_class TEXT NOT NULL,
                my_deck TEXT,
                rank_tier TEXT,
                cr_grade TEXT,
                opp_class TEXT NOT NULL,
                opp_deck TEXT,
                is_first INTEGER NOT NULL,
                is_win INTEGER NOT NULL,
                deleted_at TEXT NOT NULL,
                variant_label TEXT
            )
        """)
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_matches_deleted_user ON matches_deleted(user_id, deleted_at)"
        )
        # 修正・削除・復元の履歴（管理者がいじり方を確認するため）
        cur.execute("""
            CREATE TABLE IF NOT EXISTS match_edit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                match_id INTEGER NOT NULL,
                user_id TEXT NOT NULL,
                at TEXT NOT NULL,
                action TEXT NOT NULL,
                source TEXT NOT NULL,
                before_json TEXT,
                after_json TEXT
            )
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS idx_match_edit_log_user ON match_edit_log(user_id, at)")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS guide_links (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                my_class TEXT NOT NULL DEFAULT '',
                opp_class TEXT NOT NULL,
                url TEXT NOT NULL,
                note TEXT,
                updated_at TEXT NOT NULL,
                UNIQUE(my_class, opp_class)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bot_state (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS command_usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                command TEXT NOT NULL,
                used_at TEXT NOT NULL
            )
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS idx_command_usage_command ON command_usage(command)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_command_usage_used_at ON command_usage(used_at)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_matches_user_env ON matches(user_id, env_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_matches_env_format ON matches(env_id, format)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_matches_env_classes ON matches(env_id, my_class, opp_class)")
        conn.commit()
    finally:
        conn.close()


async def init_db():
    await asyncio.to_thread(_init_db_sync)


def _get_user_settings_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM user_settings WHERE user_id = ?", (user_id,)
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


async def get_user_settings(user_id: str):
    return await asyncio.to_thread(_get_user_settings_sync, user_id)


def _upsert_user_settings_sync(user_id: str, format_: str, my_class: str, my_deck, rank_tier, cr_grade):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO user_settings (user_id, format, my_class, my_deck, rank_tier, cr_grade, updated_at, started_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                format = excluded.format,
                my_class = excluded.my_class,
                my_deck = excluded.my_deck,
                rank_tier = excluded.rank_tier,
                cr_grade = excluded.cr_grade,
                updated_at = excluded.updated_at
        """, (user_id, format_, my_class, my_deck, rank_tier, cr_grade, now, now))
        conn.commit()
    finally:
        conn.close()


async def upsert_user_settings(user_id: str, format_: str, my_class: str, my_deck, rank_tier, cr_grade):
    await asyncio.to_thread(
        _upsert_user_settings_sync, user_id, format_, my_class, my_deck, rank_tier, cr_grade
    )


def _register_deck_name(conn, class_name: str, deck_name: str):
    """デッキ名プールに登録・使用件数を増やす（オートコンプリート候補用）"""
    if not deck_name:
        return
    row = conn.execute(
        "SELECT id, use_count FROM deck_names "
        "WHERE class_name = ? AND deck_name = ? AND env_id = ?",
        (class_name, deck_name, current_env_id()),
    ).fetchone()
    if row is None:
        conn.execute(
            "INSERT INTO deck_names (class_name, deck_name, use_count, is_official, env_id) "
            "VALUES (?, ?, 1, 0, ?)",
            (class_name, deck_name, current_env_id()),
        )
    else:
        conn.execute(
            "UPDATE deck_names SET use_count = use_count + 1 WHERE id = ?", (row["id"],)
        )


def _insert_match_sync(user_id: str, settings: dict, opp_class: str, opp_deck,
                       is_first: bool, is_win: bool):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        cur = conn.execute("""
            INSERT INTO matches (
                user_id, recorded_at, env_id, format, my_class, my_deck, rank_tier, cr_grade,
                opp_class, opp_deck, is_first, is_win
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            user_id, now, current_env_id(), settings["format"], settings["my_class"],
            settings.get("my_deck"), settings.get("rank_tier"), settings.get("cr_grade"),
            opp_class, opp_deck, 1 if is_first else 0, 1 if is_win else 0,
        ))
        match_id = cur.lastrowid
        # 自分のデッキ・相手のデッキの両方を候補プールに貯める
        _register_deck_name(conn, settings["my_class"], settings.get("my_deck"))
        _register_deck_name(conn, opp_class, opp_deck)
        conn.commit()
        return match_id
    finally:
        conn.close()


async def insert_match(user_id: str, settings: dict, opp_class: str, opp_deck,
                       is_first: bool, is_win: bool):
    match_id = await asyncio.to_thread(
        _insert_match_sync, user_id, settings, opp_class, opp_deck, is_first, is_win
    )
    mark_dirty()
    return match_id


def _get_deck_names_sync(class_name: str, limit: int = 20):
    """そのクラスで使われているデッキ名を使用件数の多い順に返す"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT deck_name FROM deck_names
            WHERE class_name = ? AND env_id = ?
            ORDER BY is_official DESC, use_count DESC, deck_name
            LIMIT ?
        """, (class_name, current_env_id(), limit)).fetchall()
        return [r["deck_name"] for r in rows]
    finally:
        conn.close()


async def get_deck_names(class_name: str, limit: int = 20):
    return await asyncio.to_thread(_get_deck_names_sync, class_name, limit)


def _merge_deck_name_sync(class_name: str, old_name: str, new_name: str, mark_official: bool):
    """表記ゆれのあるデッキ名を1つに統合する（自分側・相手側の記録、登録デッキ、候補プールを一括更新）"""
    conn = _connect()
    try:
        cur = conn.execute(
            "UPDATE matches SET my_deck = ? WHERE my_class = ? AND my_deck = ?",
            (new_name, class_name, old_name),
        )
        updated_self = cur.rowcount

        cur = conn.execute(
            "UPDATE matches SET opp_deck = ? WHERE opp_class = ? AND opp_deck = ?",
            (new_name, class_name, old_name),
        )
        updated_opp = cur.rowcount

        # 現在の設定（今まさに使っているデッキ）：ここを更新しないと、
        # デッキ切替を手動でやり直すまで確認・弱点対面などに反映されない
        cur = conn.execute(
            "UPDATE user_settings SET my_deck = ? WHERE my_class = ? AND my_deck = ?",
            (new_name, class_name, old_name),
        )
        updated_current = cur.rowcount

        # 登録デッキ（ショートカット）：同じユーザー・フォーマットに新名が既にあれば旧レコードは削除、なければ改名
        templates = conn.execute(
            "SELECT id, user_id, format FROM deck_templates WHERE my_class = ? AND deck_name = ?",
            (class_name, old_name),
        ).fetchall()
        for t in templates:
            exists = conn.execute(
                "SELECT id FROM deck_templates WHERE user_id = ? AND format = ? AND my_class = ? AND deck_name = ?",
                (t["user_id"], t["format"], class_name, new_name),
            ).fetchone()
            if exists:
                conn.execute("DELETE FROM deck_templates WHERE id = ?", (t["id"],))
            else:
                conn.execute("UPDATE deck_templates SET deck_name = ? WHERE id = ?", (new_name, t["id"]))

        # 候補プール（現環境）：件数を合算して旧レコードは削除
        old_row = conn.execute(
            "SELECT id, use_count FROM deck_names WHERE class_name = ? AND deck_name = ? AND env_id = ?",
            (class_name, old_name, current_env_id()),
        ).fetchone()
        new_row = conn.execute(
            "SELECT id, use_count FROM deck_names WHERE class_name = ? AND deck_name = ? AND env_id = ?",
            (class_name, new_name, current_env_id()),
        ).fetchone()
        if old_row and new_row:
            conn.execute(
                "UPDATE deck_names SET use_count = use_count + ? WHERE id = ?",
                (old_row["use_count"], new_row["id"]),
            )
            conn.execute("DELETE FROM deck_names WHERE id = ?", (old_row["id"],))
        elif old_row and not new_row:
            conn.execute("UPDATE deck_names SET deck_name = ? WHERE id = ?", (new_name, old_row["id"]))

        if mark_official:
            conn.execute(
                "UPDATE deck_names SET is_official = 1 WHERE class_name = ? AND deck_name = ? AND env_id = ?",
                (class_name, new_name, current_env_id()),
            )

        conn.commit()
        return {
            "matches_self": updated_self,
            "matches_opp": updated_opp,
            "templates": len(templates),
            "current_settings": updated_current,
        }
    finally:
        conn.close()


async def merge_deck_name(class_name: str, old_name: str, new_name: str, mark_official: bool):
    return await asyncio.to_thread(_merge_deck_name_sync, class_name, old_name, new_name, mark_official)


# ---- デッキテンプレート ----

def _add_template_sync(user_id: str, format_: str, my_class: str, deck_name: str) -> bool:
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        try:
            conn.execute(
                "INSERT INTO deck_templates (user_id, format, my_class, deck_name, created_at) "
                "VALUES (?, ?, ?, ?, ?)",
                (user_id, format_, my_class, deck_name, now),
            )
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False  # 同じ組み合わせが登録済み
    finally:
        conn.close()


async def add_template(user_id: str, format_: str, my_class: str, deck_name: str) -> bool:
    return await asyncio.to_thread(_add_template_sync, user_id, format_, my_class, deck_name)


def _list_templates_sync(user_id: str):
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM deck_templates WHERE user_id = ? ORDER BY id", (user_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def list_templates(user_id: str):
    return await asyncio.to_thread(_list_templates_sync, user_id)


def _delete_template_sync(user_id: str, template_id: int) -> bool:
    conn = _connect()
    try:
        cur = conn.execute(
            "DELETE FROM deck_templates WHERE id = ? AND user_id = ?", (template_id, user_id)
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


async def delete_template(user_id: str, template_id: int) -> bool:
    return await asyncio.to_thread(_delete_template_sync, user_id, template_id)


def _apply_template_sync(user_id: str, template_id: int):
    """テンプレートを現在の設定に反映する。ランク帯・グレードは維持する"""
    conn = _connect()
    try:
        t = conn.execute(
            "SELECT * FROM deck_templates WHERE id = ? AND user_id = ?", (template_id, user_id)
        ).fetchone()
        if t is None:
            return None
        cur = conn.execute("SELECT * FROM user_settings WHERE user_id = ?", (user_id,)).fetchone()
        rank_tier = cur["rank_tier"] if cur else None
        cr_grade = cur["cr_grade"] if cur else None
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO user_settings (user_id, format, my_class, my_deck, rank_tier, cr_grade, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                format = excluded.format,
                my_class = excluded.my_class,
                my_deck = excluded.my_deck,
                updated_at = excluded.updated_at
        """, (user_id, t["format"], t["my_class"], t["deck_name"], rank_tier, cr_grade, now))
        conn.commit()
        return dict(t)
    finally:
        conn.close()


async def apply_template(user_id: str, template_id: int):
    return await asyncio.to_thread(_apply_template_sync, user_id, template_id)


def _update_rank_sync(user_id: str, rank_tier: str, cr_grade) -> bool:
    """ランク帯とグレードだけを更新する。デッキ情報は触らない"""
    conn = _connect()
    try:
        row = conn.execute("SELECT user_id FROM user_settings WHERE user_id = ?", (user_id,)).fetchone()
        if row is None:
            return False  # 先に /戦績設定 が必要
        conn.execute(
            "UPDATE user_settings SET rank_tier = ?, cr_grade = ?, updated_at = ? WHERE user_id = ?",
            (rank_tier, cr_grade, datetime.now(JST).isoformat(), user_id),
        )
        conn.commit()
        return True
    finally:
        conn.close()


async def update_rank(user_id: str, rank_tier: str, cr_grade) -> bool:
    return await asyncio.to_thread(_update_rank_sync, user_id, rank_tier, cr_grade)


def _update_format_sync(user_id: str, format_: str) -> bool:
    """フォーマットだけを更新する。クラス・デッキ・ランクは触らない"""
    conn = _connect()
    try:
        row = conn.execute("SELECT user_id FROM user_settings WHERE user_id = ?", (user_id,)).fetchone()
        if row is None:
            return False
        conn.execute(
            "UPDATE user_settings SET format = ?, updated_at = ? WHERE user_id = ?",
            (format_, datetime.now(JST).isoformat(), user_id),
        )
        conn.commit()
        return True
    finally:
        conn.close()


async def update_format(user_id: str, format_: str) -> bool:
    return await asyncio.to_thread(_update_format_sync, user_id, format_)


# ---- 記録の修正・削除・復元（本人用。管理者は履歴を確認できる） ----
EDIT_DAILY_LIMIT = 20          # 1人が1日に「修正＋削除」できる件数（直後の取り消しは数えない）
RESTORE_DAYS = 7               # 削除した記録を本人が戻せる期間
MATCH_COLS = ("user_id", "recorded_at", "env_id", "format", "my_class", "my_deck", "rank_tier",
              "cr_grade", "opp_class", "opp_deck", "is_first", "is_win")


def _log_edit(conn, match_id: int, user_id: str, action: str, source: str, before=None, after=None):
    conn.execute(
        "INSERT INTO match_edit_log (match_id, user_id, at, action, source, before_json, after_json) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (match_id, user_id, datetime.now(JST).isoformat(), action, source,
         json.dumps(before, ensure_ascii=False) if before is not None else None,
         json.dumps(after, ensure_ascii=False) if after is not None else None),
    )


def _soft_delete_core(conn, user_id: str, match_id: int, action: str, source: str):
    """本人の記録を matches から退避テーブルへ移す。戻り値は従来の取消と同じ形（id/opp_class/is_win/recorded_at）"""
    row = conn.execute("SELECT * FROM matches WHERE id = ? AND user_id = ?", (match_id, user_id)).fetchone()
    if row is None:
        return None
    var = conn.execute("SELECT variant_label FROM personal_deck_variants WHERE match_id = ?", (match_id,)).fetchone()
    cols = ", ".join(("id",) + MATCH_COLS)
    conn.execute(
        f"INSERT OR REPLACE INTO matches_deleted ({cols}, deleted_at, variant_label) "
        f"VALUES ({', '.join('?' * (len(MATCH_COLS) + 1))}, ?, ?)",
        [row["id"]] + [row[c] for c in MATCH_COLS] + [datetime.now(JST).isoformat(), var["variant_label"] if var else None],
    )
    conn.execute("DELETE FROM matches WHERE id = ?", (match_id,))
    conn.execute("DELETE FROM personal_deck_variants WHERE match_id = ?", (match_id,))
    _log_edit(conn, match_id, user_id, action, source, before={k: row[k] for k in MATCH_COLS})
    conn.commit()
    return {"id": row["id"], "opp_class": row["opp_class"], "is_win": row["is_win"], "recorded_at": row["recorded_at"]}


def _edits_today(conn, user_id: str) -> int:
    today = datetime.now(JST).date().isoformat()
    return conn.execute(
        "SELECT COUNT(*) FROM match_edit_log WHERE user_id = ? AND at LIKE ? AND action IN ('edit', 'delete')",
        (user_id, today + "%"),
    ).fetchone()[0]


def _limit_message() -> str:
    return f"1日に修正・削除できるのは{EDIT_DAILY_LIMIT}件までです。日付が変わるとまた使えます。"


def _edit_match_sync(user_id: str, match_id: int, fields: dict, source: str):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM matches WHERE id = ? AND user_id = ?", (match_id, user_id)).fetchone()
        if row is None:
            return False, "記録が見つかりません。", None
        if row["env_id"] != current_env_id():
            return False, "前の環境の記録は、修正・削除できません。", None
        keys = ("opp_class", "opp_deck", "is_first", "is_win")
        before = {k: row[k] for k in keys}
        after = {"opp_class": fields["opp_class"], "opp_deck": fields.get("opp_deck") or None,
                 "is_first": 1 if fields["is_first"] else 0, "is_win": 1 if fields["is_win"] else 0}
        if before == after:
            return False, "変更がありません。", None
        if _edits_today(conn, user_id) >= EDIT_DAILY_LIMIT:
            return False, _limit_message(), None
        conn.execute(
            "UPDATE matches SET opp_class = ?, opp_deck = ?, is_first = ?, is_win = ? WHERE id = ? AND user_id = ?",
            (after["opp_class"], after["opp_deck"], after["is_first"], after["is_win"], match_id, user_id),
        )
        _register_deck_name(conn, after["opp_class"], after["opp_deck"])
        _log_edit(conn, match_id, user_id, "edit", source, before, after)
        conn.commit()
        return True, "修正しました。", after
    finally:
        conn.close()


async def edit_match(user_id: str, match_id: int, fields: dict, source: str = "discord"):
    res = await asyncio.to_thread(_edit_match_sync, user_id, match_id, fields, source)
    if res[0]:
        mark_dirty()
    return res


def _delete_match_checked_sync(user_id: str, match_id: int, source: str):
    conn = _connect()
    try:
        row = conn.execute("SELECT env_id FROM matches WHERE id = ? AND user_id = ?", (match_id, user_id)).fetchone()
        if row is None:
            return False, "記録が見つかりません。", None
        if row["env_id"] != current_env_id():
            return False, "前の環境の記録は、修正・削除できません。", None
        if _edits_today(conn, user_id) >= EDIT_DAILY_LIMIT:
            return False, _limit_message(), None
        res = _soft_delete_core(conn, user_id, match_id, "delete", source)
        return True, "削除しました。", res
    finally:
        conn.close()


async def delete_match_checked(user_id: str, match_id: int, source: str = "discord"):
    res = await asyncio.to_thread(_delete_match_checked_sync, user_id, match_id, source)
    if res[0]:
        mark_dirty()
    return res


def _restore_match_sync(user_id: str, match_id: int, source: str, admin: bool = False):
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM matches_deleted WHERE id = ?", (match_id,)).fetchone()
        if row is None or (not admin and row["user_id"] != user_id):
            return False, "戻せる記録が見つかりません。", None
        if not admin:
            if row["env_id"] != current_env_id():
                return False, "前の環境の記録は戻せません。", None
            if datetime.now(JST) - datetime.fromisoformat(row["deleted_at"]) > timedelta(days=RESTORE_DAYS):
                return False, f"削除から{RESTORE_DAYS}日を過ぎた記録は、自分では戻せません。管理者に相談してください。", None
        if conn.execute("SELECT 1 FROM matches WHERE id = ?", (match_id,)).fetchone():
            return False, "この記録はすでに戻されています。", None
        cols = ", ".join(("id",) + MATCH_COLS)
        conn.execute(
            f"INSERT INTO matches ({cols}) VALUES ({', '.join('?' * (len(MATCH_COLS) + 1))})",
            [row["id"]] + [row[c] for c in MATCH_COLS],
        )
        if row["variant_label"]:
            conn.execute(
                "INSERT OR IGNORE INTO personal_deck_variants (match_id, user_id, variant_label, created_at) "
                "VALUES (?, ?, ?, ?)",
                (row["id"], row["user_id"], row["variant_label"], datetime.now(JST).isoformat()),
            )
        conn.execute("DELETE FROM matches_deleted WHERE id = ?", (match_id,))
        _log_edit(conn, match_id, row["user_id"], "admin_restore" if admin else "restore", source,
                  after={k: row[k] for k in MATCH_COLS})
        conn.commit()
        return True, "元に戻しました。", dict(row)
    finally:
        conn.close()


async def restore_match(user_id: str, match_id: int, source: str = "discord", admin: bool = False):
    res = await asyncio.to_thread(_restore_match_sync, user_id, match_id, source, admin)
    if res[0]:
        mark_dirty()
    return res


def _list_user_matches_sync(user_id: str, limit: int, deleted: bool):
    conn = _connect()
    try:
        if deleted:
            since = (datetime.now(JST) - timedelta(days=RESTORE_DAYS)).isoformat()
            rows = conn.execute(
                "SELECT * FROM matches_deleted WHERE user_id = ? AND env_id = ? AND deleted_at >= ? "
                "ORDER BY deleted_at DESC LIMIT ?", (user_id, current_env_id(), since, limit)).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM matches WHERE user_id = ? AND env_id = ? ORDER BY id DESC LIMIT ?",
                (user_id, current_env_id(), limit)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def list_user_matches(user_id: str, limit: int = 10, deleted: bool = False):
    return await asyncio.to_thread(_list_user_matches_sync, user_id, limit, deleted)


async def get_user_match(user_id: str, match_id: int):
    rows = await asyncio.to_thread(_list_user_matches_sync, user_id, 500, False)
    return next((r for r in rows if r["id"] == match_id), None)


def _edit_report_sync(days: int):
    conn = _connect()
    try:
        since = (datetime.now(JST) - timedelta(days=days)).isoformat()
        logs = conn.execute(
            "SELECT * FROM match_edit_log WHERE at >= ? ORDER BY id", (since,)).fetchall()
        users: dict = {}
        for r in logs:
            u = users.setdefault(r["user_id"], {"edit": 0, "delete": 0, "undo": 0, "restore": 0,
                                                "del_lose": 0, "to_win": 0, "web": 0})
            a = r["action"]
            if a in ("restore", "admin_restore"):
                u["restore"] += 1
            elif a in u:
                u[a] += 1
            if r["source"] == "web":
                u["web"] += 1
            try:
                b = json.loads(r["before_json"]) if r["before_json"] else {}
                af = json.loads(r["after_json"]) if r["after_json"] else {}
            except ValueError:
                continue
            if a == "delete" and b.get("is_win") == 0:
                u["del_lose"] += 1
            if a == "edit" and b.get("is_win") == 0 and af.get("is_win") == 1:
                u["to_win"] += 1
        return users
    finally:
        conn.close()


async def edit_report(days: int = 7):
    return await asyncio.to_thread(_edit_report_sync, days)


def _delete_last_match_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT id FROM matches WHERE user_id = ? ORDER BY id DESC LIMIT 1", (user_id,),
        ).fetchone()
        if row is None:
            return None
        return _soft_delete_core(conn, user_id, row["id"], "undo", "discord")
    finally:
        conn.close()


async def delete_last_match(user_id: str):
    result = await asyncio.to_thread(_delete_last_match_sync, user_id)
    if result is not None:
        mark_dirty()
    return result


def _delete_match_by_id_sync(user_id: str, match_id: int):
    """記録完了画面の「取り消す」用。本人の記録だけを消す（実際は退避テーブルへ移す）"""
    conn = _connect()
    try:
        return _soft_delete_core(conn, user_id, match_id, "undo", "discord")
    finally:
        conn.close()


async def delete_match_by_id(user_id: str, match_id: int):
    result = await asyncio.to_thread(_delete_match_by_id_sync, user_id, match_id)
    if result is not None:
        mark_dirty()
    return result


# ---- 個人用デッキバリアント（検証タグ） ----
# 生データ（matches.my_deck）は一切変更しない。共有集計（環境デッキ分布・対面マトリクス等）は
# personal_deck_variants / user_variant_state のどちらのテーブルも参照しないので、
# ここで何をしてもデッキ名の表記ゆれ対策や公開集計には影響しない。

def _get_active_variant_sync(user_id: str, format_: str, my_class: str, my_deck: str):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT variant_label FROM user_variant_state "
            "WHERE user_id = ? AND format = ? AND my_class = ? AND my_deck = ?",
            (user_id, format_, my_class, my_deck),
        ).fetchone()
        return row["variant_label"] if row else None
    finally:
        conn.close()


async def get_active_variant(user_id: str, format_: str, my_class: str, my_deck: str):
    if not my_deck:
        return None
    return await asyncio.to_thread(_get_active_variant_sync, user_id, format_, my_class, my_deck)


def _set_active_variant_sync(user_id: str, format_: str, my_class: str, my_deck: str, variant_label: str):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO user_variant_state (user_id, format, my_class, my_deck, variant_label, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id, format, my_class, my_deck) DO UPDATE SET
                variant_label = excluded.variant_label,
                updated_at = excluded.updated_at
        """, (user_id, format_, my_class, my_deck, variant_label, now))
        conn.commit()
    finally:
        conn.close()


async def set_active_variant(user_id: str, format_: str, my_class: str, my_deck: str, variant_label: str):
    await asyncio.to_thread(_set_active_variant_sync, user_id, format_, my_class, my_deck, variant_label)


def _clear_active_variant_sync(user_id: str, format_: str, my_class: str, my_deck: str) -> bool:
    conn = _connect()
    try:
        cur = conn.execute(
            "DELETE FROM user_variant_state "
            "WHERE user_id = ? AND format = ? AND my_class = ? AND my_deck = ?",
            (user_id, format_, my_class, my_deck),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


async def clear_active_variant(user_id: str, format_: str, my_class: str, my_deck: str) -> bool:
    return await asyncio.to_thread(_clear_active_variant_sync, user_id, format_, my_class, my_deck)


def _insert_personal_variant_sync(match_id: int, user_id: str, variant_label: str):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute(
            "INSERT INTO personal_deck_variants (match_id, user_id, variant_label, created_at) "
            "VALUES (?, ?, ?, ?)",
            (match_id, user_id, variant_label, now),
        )
        conn.commit()
    finally:
        conn.close()


async def insert_personal_variant(match_id: int, user_id: str, variant_label: str):
    await asyncio.to_thread(_insert_personal_variant_sync, match_id, user_id, variant_label)


def _get_variant_summary_sync(user_id: str, format_: str, my_class: str, my_deck: str):
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT v.variant_label AS variant_label,
                   COUNT(*) AS total,
                   SUM(m.is_win) AS wins
            FROM personal_deck_variants v
            JOIN matches m ON m.id = v.match_id
            WHERE v.user_id = ? AND m.format = ? AND m.my_class = ? AND m.my_deck = ?
            GROUP BY v.variant_label
            ORDER BY total DESC
        """, (user_id, format_, my_class, my_deck)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_variant_summary(user_id: str, format_: str, my_class: str, my_deck: str):
    return await asyncio.to_thread(_get_variant_summary_sync, user_id, format_, my_class, my_deck)


def _get_summary_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute("""
            SELECT
                COUNT(*) AS total,
                SUM(is_win) AS wins
            FROM matches
            WHERE user_id = ? AND env_id = ?
        """, (user_id, current_env_id())).fetchone()
        total = row["total"] or 0
        wins = row["wins"] or 0
        return {"total": total, "wins": wins, "losses": total - wins}
    finally:
        conn.close()


async def get_summary(user_id: str):
    return await asyncio.to_thread(_get_summary_sync, user_id)


def _fetch_all_matches_sync(env_only: bool, env_id: str = None):
    conn = _connect()
    try:
        if env_only:
            rows = conn.execute(
                "SELECT * FROM matches WHERE env_id = ? ORDER BY id", (env_id or current_env_id(),)
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM matches ORDER BY id").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def fetch_all_matches(env_only: bool = True, env_id: str = None):
    return await asyncio.to_thread(_fetch_all_matches_sync, env_only, env_id)


def _fetch_all_user_settings_sync():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM user_settings").fetchall()
        return {r["user_id"]: dict(r) for r in rows}
    finally:
        conn.close()


async def fetch_all_user_settings():
    return await asyncio.to_thread(_fetch_all_user_settings_sync)


# ---- BOTの小さな永続状態（掲示板のチャンネルID・メッセージIDなど） ----

def _get_state_sync(key: str):
    conn = _connect()
    try:
        row = conn.execute("SELECT value FROM bot_state WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else None
    finally:
        conn.close()


def _get_record_counts_sync():
    conn = _connect()
    try:
        m = conn.execute("SELECT COUNT(*) AS c FROM matches").fetchone()["c"]
        u = conn.execute("SELECT COUNT(*) AS c FROM user_settings").fetchone()["c"]
        return {"matches": m, "users": u}
    finally:
        conn.close()


async def get_record_counts():
    return await asyncio.to_thread(_get_record_counts_sync)


async def get_state(key: str):
    return await asyncio.to_thread(_get_state_sync, key)


def _set_state_sync(key: str, value: str):
    conn = _connect()
    try:
        conn.execute(
            "INSERT INTO bot_state (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
        conn.commit()
    finally:
        conn.close()


async def set_state(key: str, value: str):
    await asyncio.to_thread(_set_state_sync, key, value)


# ---- 攻略記事リンク ----
# my_class は空文字列("")で「自分のクラスを問わない汎用記事」を表す

def _upsert_guide_link_sync(my_class: str, opp_class: str, url: str, note):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute("""
            INSERT INTO guide_links (my_class, opp_class, url, note, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(my_class, opp_class) DO UPDATE SET
                url = excluded.url, note = excluded.note, updated_at = excluded.updated_at
        """, (my_class, opp_class, url, note, now))
        conn.commit()
    finally:
        conn.close()


async def upsert_guide_link(my_class: str, opp_class: str, url: str, note=None):
    await asyncio.to_thread(_upsert_guide_link_sync, my_class or "", opp_class, url, note)


def _delete_guide_link_sync(my_class: str, opp_class: str) -> bool:
    conn = _connect()
    try:
        cur = conn.execute(
            "DELETE FROM guide_links WHERE my_class = ? AND opp_class = ?", (my_class, opp_class)
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


async def delete_guide_link(my_class: str, opp_class: str) -> bool:
    return await asyncio.to_thread(_delete_guide_link_sync, my_class or "", opp_class)


def _list_guide_links_sync():
    conn = _connect()
    try:
        rows = conn.execute("SELECT * FROM guide_links ORDER BY opp_class, my_class").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def list_guide_links():
    return await asyncio.to_thread(_list_guide_links_sync)


def _find_guide_link_sync(my_class: str, opp_class: str):
    conn = _connect()
    try:
        # 1. クラス指定の記事を優先、2. なければ汎用（my_class=""）にフォールバック
        row = conn.execute(
            "SELECT * FROM guide_links WHERE my_class = ? AND opp_class = ?",
            (my_class, opp_class),
        ).fetchone()
        if row is None:
            row = conn.execute(
                "SELECT * FROM guide_links WHERE my_class = '' AND opp_class = ?",
                (opp_class,),
            ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


async def find_guide_link(my_class: str, opp_class: str):
    return await asyncio.to_thread(_find_guide_link_sync, my_class, opp_class)


# ---- 個人の対面別集計（弱点対面の抽出用） ----



def _get_deck_total_sync(user_id: str, format_: str, my_deck) -> int:
    """指定デッキ・フォーマットでの、そのユーザーの通算試合数（現環境）"""
    conn = _connect()
    try:
        row = conn.execute("""
            SELECT COUNT(*) AS total FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ?
        """, (user_id, current_env_id(), format_, my_deck)).fetchone()
        return row["total"] or 0
    finally:
        conn.close()


async def get_deck_total(user_id: str, format_: str, my_deck) -> int:
    return await asyncio.to_thread(_get_deck_total_sync, user_id, format_, my_deck)


def _get_deck_summary_sync(user_id: str, format_: str, my_deck):
    """指定デッキ・フォーマットでの、そのユーザーの通算勝敗（現環境）"""
    conn = _connect()
    try:
        row = conn.execute("""
            SELECT COUNT(*) AS total, SUM(is_win) AS wins FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ?
        """, (user_id, current_env_id(), format_, my_deck)).fetchone()
        return {"total": row["total"] or 0, "wins": row["wins"] or 0}
    finally:
        conn.close()


async def get_deck_summary(user_id: str, format_: str, my_deck):
    return await asyncio.to_thread(_get_deck_summary_sync, user_id, format_, my_deck)


def _get_deck_recent_summary_sync(user_id: str, format_: str, my_deck, limit: int):
    """指定デッキ・フォーマットでの、直近limit戦の勝敗（新しい順に数える）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT is_win FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ?
            ORDER BY recorded_at DESC LIMIT ?
        """, (user_id, current_env_id(), format_, my_deck, limit)).fetchall()
        total = len(rows)
        wins = sum(r["is_win"] for r in rows)
        return {"total": total, "wins": wins}
    finally:
        conn.close()


async def get_deck_recent_summary(user_id: str, format_: str, my_deck, limit: int = RECENT_FREE_MATCHES):
    return await asyncio.to_thread(_get_deck_recent_summary_sync, user_id, format_, my_deck, limit)


def _get_recent_matches_sync(user_id: str, limit: int):
    """本人の直近の対戦履歴（デッキ・フォーマット問わず、現環境・新しい順）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT format, my_deck, opp_class, opp_deck, is_first, is_win, recorded_at
            FROM matches WHERE user_id = ? AND env_id = ?
            ORDER BY recorded_at DESC LIMIT ?
        """, (user_id, current_env_id(), limit)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_recent_matches(user_id: str, limit: int = 10):
    return await asyncio.to_thread(_get_recent_matches_sync, user_id, limit)


def _get_deck_matchups_sync(user_id: str, format_: str, my_deck):
    """指定デッキ・フォーマットでの、自分の対面別成績（相手クラスごと・現環境）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ?
            GROUP BY opp_class
        """, (user_id, current_env_id(), format_, my_deck)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_deck_matchups(user_id: str, format_: str, my_deck):
    return await asyncio.to_thread(_get_deck_matchups_sync, user_id, format_, my_deck)


def _get_global_deck_matchups_sync(format_: str, my_deck):
    """全ユーザーでの、同じデッキ・フォーマットにおける対面別の全体成績（現環境）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches
            WHERE env_id = ? AND format = ? AND my_deck = ?
            GROUP BY opp_class
        """, (current_env_id(), format_, my_deck)).fetchall()
        return {r["opp_class"]: dict(r) for r in rows}
    finally:
        conn.close()


async def get_global_deck_matchups(format_: str, my_deck):
    return await asyncio.to_thread(_get_global_deck_matchups_sync, format_, my_deck)


def _get_deck_history_sync(user_id: str, format_: str):
    """デッキを使い始めた順に、区間ごとの試合数・勝率を返す（現環境・新しい区間が先頭）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT my_deck, is_win, recorded_at
            FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ?
            ORDER BY recorded_at ASC
        """, (user_id, current_env_id(), format_)).fetchall()
    finally:
        conn.close()

    segments = []
    current = None
    for r in rows:
        deck = r["my_deck"] or "デッキ未設定"
        if current is None or current["deck"] != deck:
            current = {
                "deck": deck, "total": 0, "wins": 0,
                "start": r["recorded_at"], "end": r["recorded_at"],
            }
            segments.append(current)
        current["total"] += 1
        current["wins"] += r["is_win"]
        current["end"] = r["recorded_at"]
    segments.reverse()
    return segments


async def get_deck_history(user_id: str, format_: str):
    return await asyncio.to_thread(_get_deck_history_sync, user_id, format_)


def _get_first_second_matchups_sync(user_id: str, format_: str, my_deck):
    """指定デッキ・フォーマットでの、対面×先攻後攻別の成績（現環境）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT opp_class, is_first, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ?
            GROUP BY opp_class, is_first
        """, (user_id, current_env_id(), format_, my_deck)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_first_second_matchups(user_id: str, format_: str, my_deck):
    return await asyncio.to_thread(_get_first_second_matchups_sync, user_id, format_, my_deck)


def _get_deck_opp_deck_fs_sync(user_id: str, format_: str, my_deck):
    """指定デッキ・フォーマットでの、相手デッキ×先攻後攻別の成績（現環境）"""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT opp_class, opp_deck, is_first, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ? AND opp_deck IS NOT NULL
            GROUP BY opp_class, opp_deck, is_first
        """, (user_id, current_env_id(), format_, my_deck)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_deck_opp_deck_fs(user_id: str, format_: str, my_deck):
    return await asyncio.to_thread(_get_deck_opp_deck_fs_sync, user_id, format_, my_deck)


def _log_command_usage_sync(user_id: str, command: str):
    conn = _connect()
    try:
        now = datetime.now(JST).isoformat()
        conn.execute(
            "INSERT INTO command_usage (user_id, command, used_at) VALUES (?, ?, ?)",
            (user_id, command, now),
        )
        conn.commit()
    finally:
        conn.close()


async def log_command_usage(user_id: str, command: str):
    await asyncio.to_thread(_log_command_usage_sync, user_id, command)


def _get_activity_stats_sync():
    """週ごとの記録者数と、検証タグの利用状況（管理者向け）"""
    conn = _connect()
    try:
        now = datetime.now(JST)
        w1 = (now - timedelta(days=7)).isoformat()
        w2 = (now - timedelta(days=14)).isoformat()
        this_week = conn.execute(
            "SELECT COUNT(DISTINCT user_id) AS u, COUNT(*) AS m FROM matches WHERE recorded_at >= ?", (w1,)
        ).fetchone()
        last_week = conn.execute(
            "SELECT COUNT(DISTINCT user_id) AS u, COUNT(*) AS m FROM matches "
            "WHERE recorded_at >= ? AND recorded_at < ?", (w2, w1)
        ).fetchone()
        variants = conn.execute(
            "SELECT COUNT(DISTINCT user_id) AS u, COUNT(*) AS m FROM personal_deck_variants"
        ).fetchone()
        active_tags = conn.execute("SELECT COUNT(DISTINCT user_id) AS u FROM user_variant_state").fetchone()
        return {
            "this_week_users": this_week["u"], "this_week_matches": this_week["m"],
            "last_week_users": last_week["u"], "last_week_matches": last_week["m"],
            "variant_users": variants["u"], "variant_matches": variants["m"],
            "active_tag_users": active_tags["u"],
        }
    finally:
        conn.close()


async def get_activity_stats():
    return await asyncio.to_thread(_get_activity_stats_sync)


def _get_personal_ai_data_sync(user_id: str, format_: str, my_class: str, my_deck: str):
    """個人のAI分析用：本人の今のデッキの記録と、同じデッキを使ったほかの人の記録（現環境）"""
    conn = _connect()
    try:
        cols = "user_id, recorded_at, opp_class, opp_deck, is_first, is_win"
        mine = conn.execute(
            f"SELECT {cols} FROM matches WHERE env_id = ? AND user_id = ? AND format = ? "
            "AND my_class = ? AND my_deck = ? ORDER BY id",
            (current_env_id(), user_id, format_, my_class, my_deck),
        ).fetchall()
        everyone = conn.execute(
            f"SELECT {cols} FROM matches WHERE env_id = ? AND user_id != ? AND format = ? "
            "AND my_class = ? AND my_deck = ?",
            (current_env_id(), user_id, format_, my_class, my_deck),
        ).fetchall()
        return [dict(r) for r in mine], [dict(r) for r in everyone]
    finally:
        conn.close()


async def get_personal_ai_data(user_id: str, format_: str, my_class: str, my_deck: str):
    return await asyncio.to_thread(_get_personal_ai_data_sync, user_id, format_, my_class, my_deck)


def build_personal_ai_dataset(settings: dict, mine: list, everyone: list) -> str:
    """個人のAI分析に渡す集計テキスト。比較相手は「同じデッキを使ったほかの人」（本人を含めると差が薄まるため）"""
    def stat(rs):
        n = len(rs)
        return n, (sum(1 for r in rs if r["is_win"]) / n * 100 if n else 0.0)

    n, rate = stat(mine)
    gn, grate = stat(everyone)
    lines = [
        f"デッキ：{settings['my_deck']}（{settings['my_class']}・{FORMAT_LABELS.get(settings['format'], settings['format'])}）",
        f"ランク帯：{settings.get('rank_tier') or '不明'}",
        f"本人：{n}戦・勝率{rate:.1f}%　／　同じデッキのほかの人：{gn}戦・勝率{grate:.1f}%",
        "",
        "■ 先攻後攻（本人／同じデッキのほかの人）",
    ]
    for flag, label in ((1, "先攻"), (0, "後攻")):
        a = [r for r in mine if r["is_first"] == flag]
        b = [r for r in everyone if r["is_first"] == flag]
        an, ar = stat(a)
        bn, br = stat(b)
        lines.append(f"- {label}：本人{an}戦・{ar:.1f}%／ほかの人{bn}戦・{br:.1f}%")

    for title, key in (
        ("相手クラス", lambda r: r["opp_class"]),
        ("相手デッキ", lambda r: f"{r['opp_deck']}（{r['opp_class']}）" if r.get("opp_deck") else None),
    ):
        lines.append("")
        lines.append(f"■ {title}別（本人が{AI_PERSONAL_MIN_PAIR}戦以上のもの。本人の勝率／同じデッキのほかの人の勝率）")
        groups: dict = {}
        for r in mine:
            k = key(r)
            if k is not None:
                groups.setdefault(k, []).append(r)
        shown = 0
        for k, rs in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            if len(rs) < AI_PERSONAL_MIN_PAIR:
                continue
            an, ar = stat(rs)
            bn, br = stat([r for r in everyone if key(r) == k])
            lines.append(f"- {k}：本人{an}戦・{ar:.1f}%／ほかの人{bn}戦・{br:.1f}%（差 {ar - br:+.1f}）")
            shown += 1
        if shown == 0:
            lines.append("- （該当なし：対面ごとの試合数がまだ足りない）")

    if n >= AI_PERSONAL_RECENT * 2:
        rn, rr = stat(mine[-AI_PERSONAL_RECENT:])
        bn, br = stat(mine[:-AI_PERSONAL_RECENT])
        lines.append("")
        lines.append(f"■ 調子：直近{rn}戦の勝率{rr:.1f}%／それより前{bn}戦の勝率{br:.1f}%")
    return "\n".join(lines)


# ---- AI分析に添える画像：「もう少し勝率を上げたい対面」 ----
IMPROVE_MAX_ROWS = 5        # 画像に出す対面の数
IMPROVE_MIN_GAMES = AI_PERSONAL_MIN_PAIR   # 対面ごとにこの試合数以上のものだけ
IMPROVE_MIN_GAP = 5.0       # 自分の平均より、これ以上（ポイント）低い対面だけ
IMPROVE_SMALL_SAMPLE = 15   # これ未満の試合数には「試合数が少ない」の印をつける


def compute_improve_rows(mine: list, everyone: list):
    """
    「当たる回数が多く、自分の平均より勝率が低い対面」を、伸びしろの大きい順に返す。
    伸びしろ＝その対面の勝率を自分の平均まで上げたときに増える勝ち数（試合数×差）。
    相手デッキ名の記録が少ないときは、相手クラス別にする。
    戻り値：{"rows": [...], "level": "deck"|"class", "base": 自分の平均勝率, "n": 試合数} か None
    """
    n = len(mine)
    if n == 0:
        return None
    base = sum(1 for r in mine if r["is_win"]) / n * 100

    def build(keyfunc, labelfunc):
        groups: dict = {}
        for r in mine:
            k = keyfunc(r)
            if k is not None:
                groups.setdefault(k, []).append(r)
        others: dict = {}
        for r in everyone:
            k = keyfunc(r)
            if k is not None:
                others.setdefault(k, []).append(r)
        rows = []
        for k, rs in groups.items():
            g = len(rs)
            if g < IMPROVE_MIN_GAMES:
                continue
            rate = sum(1 for r in rs if r["is_win"]) / g * 100
            gap = base - rate
            if gap < IMPROVE_MIN_GAP:
                continue
            o = others.get(k, [])
            orate = sum(1 for r in o if r["is_win"]) / len(o) * 100 if len(o) >= IMPROVE_MIN_GAMES else None
            rows.append({"label": labelfunc(k), "games": g, "wins": sum(1 for r in rs if r["is_win"]),
                         "rate": rate, "gap": gap, "potential": g * gap / 100, "others": orate})
        rows.sort(key=lambda x: -x["potential"])
        return rows[:IMPROVE_MAX_ROWS]

    rows = build(lambda r: (r["opp_class"], r["opp_deck"]) if r.get("opp_deck") else None,
                 lambda k: f"{k[1]}（{k[0]}）")
    level = "deck"
    if len(rows) < 2:
        class_rows = build(lambda r: r["opp_class"], lambda k: k)
        if len(class_rows) > len(rows):
            rows, level = class_rows, "class"
    if not rows:
        return None
    return {"rows": rows, "level": level, "base": base, "n": n}


def render_improve_chart(deck_label: str, result: dict):
    """勝率を上げたい対面の横棒グラフ。フォント未導入などの場合は None"""
    if plt is None or _JP_FONT_NAME is None or not result or not result["rows"]:
        return None
    try:
        rows = list(reversed(result["rows"]))      # barhは先頭が下に積まれるので、優先度が高いものを上に
        base = result["base"]
        height = max(3.0, 0.95 * len(rows) + 1.8)
        fig, ax = plt.subplots(figsize=(8.2, height), dpi=150)
        y = list(range(len(rows)))
        colors = ["#d9534f" if r["gap"] >= 12 else "#f0ad4e" for r in rows]
        ax.barh(y, [r["rate"] for r in rows], color=colors, edgecolor="#333333", linewidth=0.5, height=0.62, zorder=2)
        ax.axvline(base, color="#2b6cb0", linestyle="--", linewidth=1.6, zorder=3)
        ax.set_xlim(0, 140)
        ax.set_ylim(-0.6, len(rows) - 0.2 + 0.45)
        ax.text(base, len(rows) - 0.2 + 0.4, f"あなたの平均 {base:.0f}%", color="#2b6cb0", fontsize=10,
                ha="center", va="top", weight="bold", zorder=5,
                bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.95))
        ax.set_xticks([0, 20, 40, 60, 80, 100])
        ax.set_yticks(y)
        ax.set_yticklabels([("★ " if i == len(rows) - 1 else "") + r["label"] for i, r in enumerate(rows)], fontsize=12)
        ax.set_xlabel("あなたの勝率（%）", fontsize=11)
        box = dict(boxstyle="round,pad=0.12", fc="white", ec="none", alpha=0.92)
        for yi, r in zip(y, rows):
            x0 = max(r["rate"], 0) + 2
            ax.text(x0, yi + 0.2, f"{r['rate']:.0f}%（{r['wins']}勝{r['games'] - r['wins']}敗）",
                    va="center", ha="left", fontsize=11, color="#222222", weight="bold", zorder=4, bbox=box)
            ax.text(x0, yi - 0.02, f"平均より{r['gap']:.0f}pt低い／あと約{r['potential']:.1f}勝分",
                    va="center", ha="left", fontsize=8.8, color="#444444", zorder=4, bbox=box)
            extra = []
            if r["others"] is not None:
                extra.append(f"同じデッキのほかの人 {r['others']:.0f}%")
            if r["games"] < IMPROVE_SMALL_SAMPLE:
                extra.append("※試合数が少ない")
            if extra:
                ax.text(x0, yi - 0.22, "　".join(extra), va="center", ha="left", fontsize=8.4, color="#666666",
                        zorder=4, bbox=box)
        scope = "相手デッキ別" if result["level"] == "deck" else "相手クラス別（デッキ名の記録が少ないため）"
        ax.set_title(f"{deck_label}　もう少し勝率を上げたい対面", fontsize=14, pad=26)
        ax.text(0.0, 1.04, f"{scope}・当たる回数が多く、平均より勝率が低い順（★が一番の伸びしろ）",
                transform=ax.transAxes, fontsize=9.5, color="#666666", ha="left", va="bottom")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="x", color="#e5e5e5", zorder=0)
        fig.tight_layout()
        buf = io.BytesIO()
        fig.savefig(buf, format="png")
        plt.close(fig)
        buf.seek(0)
        return buf
    except Exception as e:
        print(f"⚠️ 勝率を上げたい対面の画像作成に失敗: {type(e).__name__}: {e}")
        return None


# AIの呼び出しに使うモデルが返答できなかったときの予備（Haiku）
AI_FALLBACK_MODEL = "claude-haiku-5-5"
# 返答の上限。モデルによっては内部で考える分にも使うため、短くしすぎると本文が空になる
AI_MAX_TOKENS = 3000


async def ask_ai(api_key: str, model: str, system: str, user_text: str):
    """
    AIに1回質問する。指定モデルで本文が返らなければ予備モデルでもう一度試す。
    戻り値：(本文 or None, 使ったモデル, 失敗の理由 or None)
    """
    client = AsyncAnthropic(api_key=api_key, timeout=90)
    reasons = []
    for m in dict.fromkeys([model, AI_FALLBACK_MODEL]):  # 同じモデルなら1回だけ
        try:
            msg = await client.messages.create(
                model=m, max_tokens=AI_MAX_TOKENS, system=system,
                messages=[{"role": "user", "content": user_text}],
            )
            text = "".join(
                b.text for b in msg.content if getattr(b, "type", "") == "text"
            ).strip()
            if text:
                return text, m, None
            kinds = ",".join(getattr(b, "type", "?") for b in msg.content) or "なし"
            reasons.append(f"{m}: 本文なし（stop_reason={getattr(msg, 'stop_reason', '?')} / 中身={kinds}）")
        except Exception as e:
            reasons.append(f"{m}: {type(e).__name__}: {e}"[:250])
    return None, model, " ／ ".join(reasons)


def _get_command_usage_stats_sync():
    conn = _connect()
    try:
        by_command = conn.execute("""
            SELECT command, COUNT(*) AS total, COUNT(DISTINCT user_id) AS users
            FROM command_usage GROUP BY command ORDER BY total DESC
        """).fetchall()
        since_30 = (datetime.now(JST) - timedelta(days=30)).isoformat()
        recent = conn.execute("""
            SELECT COUNT(*) AS total, COUNT(DISTINCT user_id) AS users
            FROM command_usage WHERE used_at >= ?
        """, (since_30,)).fetchone()
        analyzers = conn.execute(
            "SELECT COUNT(DISTINCT user_id) AS n FROM command_usage"
        ).fetchone()["n"] or 0
        recorders = conn.execute(
            "SELECT COUNT(DISTINCT user_id) AS n FROM matches WHERE env_id = ?", (current_env_id(),)
        ).fetchone()["n"] or 0
        return {
            "by_command": [dict(r) for r in by_command],
            "recent_30_total": recent["total"] or 0,
            "recent_30_users": recent["users"] or 0,
            "analyzers": analyzers,
            "recorders": recorders,
        }
    finally:
        conn.close()


async def get_command_usage_stats():
    return await asyncio.to_thread(_get_command_usage_stats_sync)


def _get_global_summary_sync():
    """現環境の全体集計。クラス別・先後別まで一度に出す"""
    conn = _connect()
    try:
        total_row = conn.execute("""
            SELECT COUNT(*) AS total, SUM(is_win) AS wins,
                   COUNT(DISTINCT user_id) AS users
            FROM matches WHERE env_id = ?
        """, (current_env_id(),)).fetchone()

        by_first = conn.execute("""
            SELECT is_first, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ? GROUP BY is_first
        """, (current_env_id(),)).fetchall()

        by_my_class = conn.execute("""
            SELECT my_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ?
            GROUP BY my_class ORDER BY total DESC
        """, (current_env_id(),)).fetchall()

        by_opp_class = conn.execute("""
            SELECT opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ?
            GROUP BY opp_class ORDER BY total DESC
        """, (current_env_id(),)).fetchall()

        return {
            "total": total_row["total"] or 0,
            "wins": total_row["wins"] or 0,
            "users": total_row["users"] or 0,
            "by_first": [dict(r) for r in by_first],
            "by_my_class": [dict(r) for r in by_my_class],
            "by_opp_class": [dict(r) for r in by_opp_class],
        }
    finally:
        conn.close()


async def get_global_summary():
    return await asyncio.to_thread(_get_global_summary_sync)


def build_ai_dataset(matches: list[dict], env_id: str, now: datetime):
    """
    環境レポートのAI分析に渡す集計テキストを作る（ローテーションのみ）。
    使用率は相手側、勝ちやすさは記録者平均との差、直近と以前の比較を含める。
    記録が無ければ None。
    """
    rows = [m for m in matches if m.get("format") == "rotation"]
    n = len(rows)
    if n == 0:
        return None

    def wins(rs):
        return sum(1 for r in rs if r["is_win"])

    def pct(a, b):
        return a / b * 100 if b else 0.0

    avg = pct(wins(rows), n)
    users: dict[str, int] = {}
    for r in rows:
        users[r["user_id"]] = users.get(r["user_id"], 0) + 1
    top4 = sum(sorted(users.values(), reverse=True)[:4])

    lines = [
        f"環境ID：{env_id}（ローテーションのみ）",
        f"総試合数：{n}戦 / 記録者：{len(users)}名 / 記録者全体の平均勝率：{avg:.1f}%",
        f"記録の偏り：上位4名の記録が全体の{pct(top4, n):.0f}%",
    ]
    ranks: dict[str, int] = {}
    for r in rows:
        k = r.get("rank_tier") or "不明"
        ranks[k] = ranks.get(k, 0) + 1
    lines.append("ランク帯の内訳：" + "、".join(
        f"{k}{pct(v, n):.0f}%" for k, v in sorted(ranks.items(), key=lambda kv: -kv[1])
    ))
    lines.append("")

    # 相手クラス：使用率と、記録者から見た勝率（平均との差）
    lines.append("■ 相手クラスの使用率（＝環境の使用率）と、記録者から見た勝率")
    by_cls: dict[str, list] = {}
    for r in rows:
        by_cls.setdefault(r["opp_class"], []).append(r)
    for c, rs in sorted(by_cls.items(), key=lambda kv: -len(kv[1])):
        w = pct(wins(rs), len(rs))
        lines.append(f"- {c}：{len(rs)}戦（{pct(len(rs), n):.1f}%）・このクラスと当たったときの記録者の勝率{w:.1f}%（平均との差 {w - avg:+.1f}）")
    lines.append("")

    # 相手デッキ
    with_deck = [r for r in rows if r.get("opp_deck")]
    lines.append(f"■ 相手デッキの使用率（相手デッキの入力率{pct(len(with_deck), n):.0f}%。入力のあった試合のみで計算）")
    by_deck: dict[tuple, list] = {}
    for r in with_deck:
        by_deck.setdefault((r["opp_class"], r["opp_deck"]), []).append(r)
    for (c, d), rs in sorted(by_deck.items(), key=lambda kv: -len(kv[1]))[:ENV_REPORT_AI_TOP_DECKS]:
        w = pct(wins(rs), len(rs))
        lines.append(
            f"- {d}（{c}）：{len(rs)}戦（{pct(len(rs), len(with_deck)):.1f}%）・このデッキと当たったときの記録者の勝率{w:.1f}%（平均との差 {w - avg:+.1f}）"
        )
    lines.append("")

    # 先攻後攻
    lines.append("■ 先攻後攻")
    for flag, label in ((1, "先攻"), (0, "後攻")):
        rs = [r for r in rows if r["is_first"] == flag]
        if rs:
            lines.append(f"- {label}：{len(rs)}戦・記録者の勝率{pct(wins(rs), len(rs)):.1f}%")
    lines.append("")

    # 直近と以前の比較（相手クラス・相手デッキの使用率の変化）
    def recorded(r):
        try:
            return datetime.fromisoformat(r["recorded_at"])
        except (TypeError, ValueError):
            return None

    # 直近の試合数が少なすぎるときは、期間を倍に広げて比べる
    for days in (ENV_REPORT_AI_RECENT_DAYS, ENV_REPORT_AI_RECENT_DAYS * 2):
        cutoff = now - timedelta(days=days)
        recent = [r for r in rows if (recorded(r) or now) >= cutoff]
        before = [r for r in rows if (recorded(r) or now) < cutoff]
        if len(recent) >= 30:
            break
    lines.append(f"■ 直近{days}日（{len(recent)}戦）と、それより前（{len(before)}戦）の比較")
    if len(recent) < 30 or len(before) < 30:
        lines.append("- 比較できるだけの試合数がまだない")
    else:
        def shares(rs, key):
            out: dict = {}
            for r in rs:
                k = key(r)
                if k is not None:
                    out[k] = out.get(k, 0) + 1
            total = sum(out.values()) or 1
            return {k: v / total * 100 for k, v in out.items()}, out

        for title, key in (
            ("相手クラス", lambda r: r["opp_class"]),
            ("相手デッキ", lambda r: f"{r['opp_deck']}（{r['opp_class']}）" if r.get("opp_deck") else None),
        ):
            rec_s, rec_c = shares(recent, key)
            bef_s, _ = shares(before, key)
            diffs = [
                (k, rec_s.get(k, 0.0) - bef_s.get(k, 0.0), rec_c.get(k, 0))
                for k in set(rec_s) | set(bef_s)
            ]
            diffs.sort(key=lambda x: -abs(x[1]))
            lines.append(f"{title}（変化の大きい順）：")
            for k, d, c in diffs[:5]:
                lines.append(
                    f"- {k}：以前{bef_s.get(k, 0.0):.1f}% → 直近{rec_s.get(k, 0.0):.1f}%（{d:+.1f}・直近{c}戦）"
                )
    return "\n".join(lines)




def _get_class_matchup_matrix_sync(env_id: str):
    """
    現環境の、クラス対クラスの対面集計（試合数・勝数）。環境分析（Sheets・Discord画像）用。

    記録は「自分のクラス」側からしか行われないため、片方向のGROUP BYだけだと
    例えば「ロイヤル使用者がエルフと対面した記録」はあっても「エルフ使用者が
    ロイヤルと対面した記録」がまだ無い、という非対称なマトリクスになりやすい
    （実際、サーバー全体でもどちらかの組み合わせしか記録がない対面が多かった）。

    そこで、相手側の記録も鏡写しにして合算する：
    「my_class=X, opp_class=Y」の記録だけでなく、「my_class=Y, opp_class=X」の
    記録も（勝敗を反転させて）Xの対Y成績としてカウントする。1試合が両クラスの
    集計に効くので、対面ごとの母数が実質倍になり、X対YとY対Xの勝率は
    必ず合計100%になる（同じ試合プールを両側から見ているだけなので）。
    """
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT class_a AS my_class, class_b AS opp_class,
                   SUM(total) AS total, SUM(wins) AS wins
            FROM (
                SELECT my_class AS class_a, opp_class AS class_b,
                       COUNT(*) AS total, SUM(is_win) AS wins
                FROM matches WHERE env_id = ?
                GROUP BY my_class, opp_class
                UNION ALL
                SELECT opp_class AS class_a, my_class AS class_b,
                       COUNT(*) AS total, SUM(1 - is_win) AS wins
                FROM matches WHERE env_id = ?
                GROUP BY my_class, opp_class
            ) t
            GROUP BY class_a, class_b
        """, (env_id, env_id)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_class_matchup_matrix(env_id: str = None):
    return await asyncio.to_thread(_get_class_matchup_matrix_sync, env_id or current_env_id())


def combine_symmetric_class_stats(by_my_class: list, by_opp_class: list) -> list:
    """
    クラス別の試合数・勝数を、自分側の記録と相手側の記録の両方から合算する。

    by_my_class は「自分がそのクラスを使ったときの成績」、by_opp_class は
    「そのクラスを相手にしたときの、自分の勝敗」（＝相手クラスの成績を裏返したもの）。
    これらを合算すると、自分がそのクラスを使った試合と、相手がそのクラスを
    使った試合の両方が「そのクラスの成績」としてカウントされる。

    例：ロイヤルを使う人が少なくても、ロイヤルと対面した人は多ければ、
    後者の記録（を裏返したもの）からロイヤルの勝率が見えるようになる。
    """
    totals: dict = {}
    wins: dict = {}
    for r in by_my_class:
        c = r["my_class"]
        totals[c] = totals.get(c, 0) + r["total"]
        wins[c] = wins.get(c, 0) + (r["wins"] or 0)
    for r in by_opp_class:
        c = r["opp_class"]
        total = r["total"]
        opp_wins = total - (r["wins"] or 0)  # 自分が負けた試合数＝相手クラスの勝数
        totals[c] = totals.get(c, 0) + total
        wins[c] = wins.get(c, 0) + opp_wins
    return [
        {"my_class": c, "total": totals[c], "wins": wins[c]}
        for c in sorted(totals, key=lambda k: -totals[k])
    ]


# ==============================
# /戦績 の記録画面（1画面方式）
# ==============================
# 1画面に「①相手クラス」「②相手デッキ（任意）」「③先攻/後攻」「④勝敗」を並べる。
# 試合開始時に①③を選んでおき、試合中に分かれば②、試合後に④を押すと記録される。
# 選んだ内容はユーザーごとにメモリに一時保存し、受付時間切れのあとでも
# パネルから開き直せば続きから再開できる（BOTが再起動すると消える）。

# 一時保存した選択内容を復元する期限（分）
FLOW_DRAFT_TTL_MINUTES = 60

# user_id -> {"opp_class", "opp_deck", "is_first", "saved_at", "settings_key"}
_FLOW_DRAFTS: dict[str, dict] = {}


def _settings_key(settings: dict) -> tuple:
    """デッキを切り替えたあとに古い選択内容を復元しないための目印"""
    return (settings.get("format"), settings.get("my_class"), settings.get("my_deck"))


def _save_flow_draft(view: "SensekiFlowView"):
    if view.opp_class is None and view.opp_deck is None and view.is_first is None:
        _FLOW_DRAFTS.pop(view.user_id, None)
        return
    _FLOW_DRAFTS[view.user_id] = {
        "opp_class": view.opp_class,
        "opp_deck": view.opp_deck,
        "is_first": view.is_first,
        "saved_at": datetime.now(JST),
        "settings_key": _settings_key(view.settings),
    }


def _load_flow_draft(user_id: str, settings: dict):
    d = _FLOW_DRAFTS.get(user_id)
    if d is None:
        return None
    expired = datetime.now(JST) - d["saved_at"] > timedelta(minutes=FLOW_DRAFT_TTL_MINUTES)
    if expired or d["settings_key"] != _settings_key(settings):
        _FLOW_DRAFTS.pop(user_id, None)
        return None
    return d


def _clear_flow_draft(user_id: str):
    _FLOW_DRAFTS.pop(user_id, None)


class OpponentClassSelect(discord.ui.Select):
    def __init__(self, selected=None):
        options = [
            discord.SelectOption(label=c, value=c, default=(c == selected))
            for c in CLASS_CHOICES
        ]
        super().__init__(
            placeholder="① 相手クラスを選択", options=options,
            min_values=1, max_values=1, row=0,
        )

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        new_class = self.values[0]
        if new_class != view.opp_class:
            # クラスを選び直したら、前のクラスのデッキ名は持ち越さない
            view.opp_deck = None
        view.opp_class = new_class
        await view.refresh(interaction)


class OpponentDeckSelect(discord.ui.Select):
    """既知のデッキ名から選ぶ。候補になければモーダルで新規入力。未選択でも記録できる"""

    SKIP_VALUE = "__skip__"
    NEW_VALUE = "__new__"

    def __init__(self, known_decks: list[str], selected, enabled: bool):
        options = [
            discord.SelectOption(label="分からない / 入力しない", value=self.SKIP_VALUE),
            discord.SelectOption(label="新しく入力する…", value=self.NEW_VALUE),
        ]
        names = [n[:100] for n in known_decks]
        # モーダルで入力した新しいデッキ名は候補にまだ無いので、先頭に足して選択状態にする
        if selected and selected[:100] not in names:
            names.insert(0, selected[:100])
        # Discordのセレクトは25件が上限。既定の2件を除いた分だけ載せる
        for name in names[:23]:
            options.append(discord.SelectOption(
                label=name, value=name, default=(selected is not None and name == selected[:100]),
            ))
        placeholder = (
            "② 相手デッキ（分かれば・任意）" if enabled else "② 相手デッキ（先に①を選んでください）"
        )
        super().__init__(
            placeholder=placeholder, options=options,
            min_values=1, max_values=1, row=1, disabled=not enabled,
        )

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        value = self.values[0]

        if value == self.NEW_VALUE:
            await interaction.response.send_modal(OpponentDeckModal(view))
            return

        view.opp_deck = None if value == self.SKIP_VALUE else value
        await view.refresh(interaction)


class OpponentDeckModal(discord.ui.Modal, title="相手のデッキ名を入力"):
    deck_name = discord.ui.TextInput(
        label="デッキ名",
        placeholder="例：進化エルフ",
        required=True,
        max_length=50,
    )

    def __init__(self, flow_view: "SensekiFlowView"):
        super().__init__()
        self.flow_view = flow_view

    async def on_submit(self, interaction: discord.Interaction):
        view = self.flow_view
        view.opp_deck = str(self.deck_name.value).strip() or None
        await view.refresh(interaction)


class FirstButton(discord.ui.Button):
    def __init__(self, selected: bool):
        super().__init__(
            label="先攻",
            style=discord.ButtonStyle.primary if selected else discord.ButtonStyle.secondary,
            row=2,
        )

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.is_first = True
        await view.refresh(interaction)


class SecondButton(discord.ui.Button):
    def __init__(self, selected: bool):
        super().__init__(
            label="後攻",
            style=discord.ButtonStyle.primary if selected else discord.ButtonStyle.secondary,
            row=2,
        )

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.is_first = False
        await view.refresh(interaction)


# ---- 記録画面からのランク変更（ワンタップ） ----
# ランクを「1段ずつの階段」として並べる。ランクアップ／ダウンはこの上を1つ動くだけ。
# グランドマスターは、グレードなし→EPIC→ULTIMATE→LEGEND→BEYOND の順。
# ※BEYONDは「LEGEND相当（CR1850以上）かつCRランキング100位以内」なので、
#   LEGENDからの昇段として扱う（実際に条件を満たすかは本人が判断する）。
RANK_LADDER = (
    [(t, None) for t in RANK_TIER_CHOICES if t != GRAND_MASTER_TIER]
    + [(GRAND_MASTER_TIER, None)]
    + [(GRAND_MASTER_TIER, g) for g in CR_GRADE_CHOICES]
)


def rank_step_label(tier: str, grade) -> str:
    """ランクの表示名。ボタンと完了メッセージで使う"""
    if tier == GRAND_MASTER_TIER:
        return f"GM・{grade}" if grade else "GM・グレードなし"
    return tier


def rank_step_index(tier, grade):
    """今のランクが階段の何段目か。未設定や不明なら None"""
    try:
        return RANK_LADDER.index((tier, grade or None))
    except ValueError:
        return None


class FlowRankStepButton(discord.ui.Button):
    """ワンタップでランクを1段上げる／下げる。ボタンに移動先を表示するので押し間違えにくい"""

    def __init__(self, direction: int, target):
        tier, grade = target
        arrow = "⬆" if direction > 0 else "⬇"
        super().__init__(
            label=f"{arrow} {rank_step_label(tier, grade)}",
            style=discord.ButtonStyle.primary, row=4,
        )
        self.direction = direction
        self.target = target

    async def callback(self, interaction: discord.Interaction):
        view: "SensekiFlowView" = self.view
        tier, grade = self.target
        ok = await update_rank(view.user_id, tier, grade)
        if not ok:
            await view.refresh(interaction, notice="⚠️ 先に初期設定をしてください。")
            return
        view.settings = await get_user_settings(view.user_id) or view.settings
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()
        await view.refresh(interaction, notice=f"✅ ランクを **{rank_step_label(tier, grade)}** に更新しました。")


class FlowRankNowButton(discord.ui.Button):
    """4段目の左端に「今のランク」を表示するだけのボタン（押せない）。ランク欄だと気づきやすくする"""

    def __init__(self, tier, grade):
        text = rank_step_label(tier, grade) if tier else "未設定"
        super().__init__(
            label=f"今のランク：{text}", emoji="📈",
            style=discord.ButtonStyle.secondary, disabled=True, row=4,
        )


class FlowRankPickButton(discord.ui.Button):
    """階段で動かせない・大きく変えたいとき用。記録画面のままランクを選び直す"""

    def __init__(self):
        super().__init__(label="ランクを選ぶ", emoji="📈", style=discord.ButtonStyle.secondary, row=4)

    async def callback(self, interaction: discord.Interaction):
        view: "SensekiFlowView" = self.view
        view.rank_mode = "tier"
        await view.refresh(interaction)


class FlowRankTierSelect(discord.ui.Select):
    def __init__(self):
        super().__init__(
            placeholder="ランク帯を選択",
            options=[discord.SelectOption(label=r, value=r) for r in RANK_TIER_CHOICES],
            row=0,
        )

    async def callback(self, interaction: discord.Interaction):
        view: "SensekiFlowView" = self.view
        tier = self.values[0]
        if tier == GRAND_MASTER_TIER:
            view.rank_mode = "grade"
            await view.refresh(interaction)
            return
        await view.apply_rank(interaction, tier, None)


class FlowRankGradeSelect(discord.ui.Select):
    def __init__(self):
        super().__init__(
            placeholder="グレードを選択",
            options=[discord.SelectOption(label=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
                    + [discord.SelectOption(label=g, value=g) for g in CR_GRADE_CHOICES],
            row=0,
        )

    async def callback(self, interaction: discord.Interaction):
        view: "SensekiFlowView" = self.view
        grade = None if self.values[0] == CR_GRADE_NONE_VALUE else self.values[0]
        await view.apply_rank(interaction, GRAND_MASTER_TIER, grade)


class FlowRankBackButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="記録画面に戻る", style=discord.ButtonStyle.secondary, row=1)

    async def callback(self, interaction: discord.Interaction):
        view: "SensekiFlowView" = self.view
        view.rank_mode = None
        await view.refresh(interaction)


class WinButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="勝ち", style=discord.ButtonStyle.success, row=3)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        await view.finalize(interaction, is_win=True)


class LoseButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="負け", style=discord.ButtonStyle.danger, row=3)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        await view.finalize(interaction, is_win=False)


class RepeatButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="同じ設定でもう1試合", style=discord.ButtonStyle.primary, row=0)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.opp_class = None
        view.opp_deck = None
        view.is_first = None
        await view.refresh(interaction)


class UndoButton(discord.ui.Button):
    """記録完了画面から、今記録した1件だけを取り消す"""

    def __init__(self, match_id: int):
        super().__init__(label="今の記録を取り消す", style=discord.ButtonStyle.secondary, row=0)
        self.match_id = match_id

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        deleted = await delete_match_by_id(view.user_id, self.match_id)
        if deleted is None:
            notice = "⚠️ この記録は既に取り消されているため、何もしていません。"
        else:
            notice = "🗑️ 今の記録を取り消しました。選び直して、もう一度勝敗を押してください。"
        # 取り消した試合の選択内容は残したまま入力画面に戻す（勝敗の押し間違いをすぐ直せる）
        await view.refresh(interaction, notice=notice)


class FlowSettingsButton(discord.ui.Button):
    """記録画面から設定メニューを開く。記録時に設定を読み直すので、変えたらそのまま記録できる"""

    def __init__(self):
        super().__init__(label="設定を変える", emoji="⚙️", style=discord.ButtonStyle.secondary, row=2)

    async def callback(self, interaction: discord.Interaction):
        await send_settings_menu(interaction)


class SensekiFlowView(discord.ui.View):
    def __init__(self, user_id: str, settings: dict, origin_interaction: discord.Interaction):
        # 13分。Discordのフォローアップ編集が可能な15分に収まる範囲で設定。
        # 長い試合で時間切れになっても、選んだ内容は一時保存から復元できる。
        super().__init__(timeout=780)
        self.user_id = user_id
        self.settings = settings
        self.origin_interaction = origin_interaction
        self.opp_class = None
        self.opp_deck = None
        self.is_first = None
        self.restored = False
        self.rank_mode = None  # None＝通常の記録画面 / "tier"・"grade"＝ランク選択中
        self._deck_cache: dict[str, list[str]] = {}

    @classmethod
    async def create(cls, user_id: str, settings: dict, origin_interaction: discord.Interaction):
        view = cls(user_id, settings, origin_interaction)
        draft = _load_flow_draft(user_id, settings)
        if draft is not None:
            view.opp_class = draft["opp_class"]
            view.opp_deck = draft["opp_deck"]
            view.is_first = draft["is_first"]
            view.restored = True
        await view.build()
        return view

    def describe_settings(self) -> str:
        """今の自分の設定。全ステップで出し続けて、古いまま記録するのを防ぐ"""
        s = self.settings
        fmt = FORMAT_LABELS.get(s["format"], s["format"])
        deck = s.get("my_deck") or "デッキ未設定"
        rank = s.get("rank_tier") or "ランク未設定"
        if s.get("cr_grade"):
            rank += f"・{s['cr_grade']}"
        elif s.get("rank_tier") == GRAND_MASTER_TIER:
            rank += "・グレードなし"
        return (
            f"🧑 あなた：**{deck}**（{s['my_class']} / {fmt}）\n"
            f"📈 現在のランク：**{rank}**　-# 変わったら下の青いボタンで更新"
        )

    def describe_progress(self) -> str:
        """ここまでに選んだ内容を1行で"""
        if self.opp_class is None and self.is_first is None:
            return ""
        label = f"🆚 相手：**{self.opp_class or '未選択'}**"
        if self.opp_deck:
            label += f"（{self.opp_deck}）"
        if self.is_first is not None:
            label += f" / **{'先攻' if self.is_first else '後攻'}**"
        return label

    def screen(self, body: str) -> str:
        """設定ヘッダー＋進捗＋操作案内をまとめた画面テキスト"""
        parts = [self.describe_settings()]
        progress = self.describe_progress()
        if progress:
            parts.append(progress)
        parts.append("")
        parts.append(body)
        return "\n".join(parts)

    def guide_text(self, notice: str = None) -> str:
        lines = []
        if notice:
            lines.append(notice)
        if self.rank_mode == "tier":
            return "\n".join(lines + ["新しいランク帯を選んでください。"])
        if self.rank_mode == "grade":
            return "\n".join(lines + [f"グレードを選んでください（付いていなければ「{CR_GRADE_NONE_LABEL}」）。"])
        if self.opp_class is not None and self.is_first is not None:
            lines.append("準備OKです。試合が終わったら **④ 勝ち／負け** を押してください。")
            lines.append("-# 相手のデッキが分かったら、勝敗を押す前に②で選べます（任意）。")
        else:
            lines.append(
                "試合開始時に **①相手クラス** と **③先攻／後攻** を選び、"
                "試合が終わったら **④勝ち／負け** を押してください。"
            )
            lines.append("-# ②相手デッキは、分かったときだけ選べばOKです。")
        return "\n".join(lines)

    async def _known_decks(self) -> list[str]:
        if self.opp_class is None:
            return []
        if self.opp_class not in self._deck_cache:
            self._deck_cache[self.opp_class] = await get_deck_names(self.opp_class)
        return self._deck_cache[self.opp_class]

    async def build(self):
        """選択状態に合わせて入力画面の部品を並べ直す"""
        self.clear_items()
        if self.rank_mode == "tier":
            self.add_item(FlowRankTierSelect())
            self.add_item(FlowRankBackButton())
            return
        if self.rank_mode == "grade":
            self.add_item(FlowRankGradeSelect())
            self.add_item(FlowRankBackButton())
            return
        self.add_item(OpponentClassSelect(self.opp_class))
        self.add_item(OpponentDeckSelect(
            await self._known_decks(), self.opp_deck, enabled=self.opp_class is not None,
        ))
        self.add_item(FirstButton(self.is_first is True))
        self.add_item(SecondButton(self.is_first is False))
        self.add_item(WinButton())
        self.add_item(LoseButton())
        self.add_item(FlowSettingsButton())
        # ランク変更（4段目）：今のランク表示＋ワンタップで1段上下＋選び直し
        self.add_item(FlowRankNowButton(self.settings.get("rank_tier"), self.settings.get("cr_grade")))
        idx = rank_step_index(self.settings.get("rank_tier"), self.settings.get("cr_grade"))
        if idx is not None:
            if idx > 0:
                self.add_item(FlowRankStepButton(-1, RANK_LADDER[idx - 1]))
            if idx < len(RANK_LADDER) - 1:
                self.add_item(FlowRankStepButton(+1, RANK_LADDER[idx + 1]))
        self.add_item(FlowRankPickButton())

    async def apply_rank(self, interaction: discord.Interaction, tier: str, grade):
        """ランク選択の結果を保存して、記録画面に戻る"""
        ok = await update_rank(self.user_id, tier, grade)
        if not ok:
            self.rank_mode = None
            await self.refresh(interaction, notice="⚠️ 先に初期設定をしてください。")
            return
        self.settings = await get_user_settings(self.user_id) or self.settings
        self.rank_mode = None
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()
        await self.refresh(
            interaction, notice=f"✅ ランクを **{rank_step_label(tier, grade)}** に更新しました。"
        )

    async def refresh(self, interaction: discord.Interaction, notice: str = None):
        await self.build()
        _save_flow_draft(self)
        await interaction.response.edit_message(
            content=self.screen(self.guide_text(notice)), view=self,
        )

    async def finalize(self, interaction: discord.Interaction, is_win: bool):
        missing = []
        if self.opp_class is None:
            missing.append("①相手クラス")
        if self.is_first is None:
            missing.append("③先攻／後攻")
        if missing:
            await self.refresh(
                interaction,
                notice=f"⚠️ まだ記録していません。{'と'.join(missing)}を選んでから勝敗を押してください。",
            )
            return

        # 記録画面を開いたあとに「⚙️ 設定を変える」でデッキやランクを変えた場合に備え、
        # 記録の直前に最新の設定を読み直す（古い設定のまま記録されるのを防ぐ）
        latest = await get_user_settings(self.user_id)
        if latest is not None:
            self.settings = latest
        match_id = await insert_match(
            self.user_id, self.settings, self.opp_class, self.opp_deck,
            self.is_first, is_win,
        )
        variant_label = await get_active_variant(
            self.user_id, self.settings["format"], self.settings["my_class"],
            self.settings.get("my_deck"),
        )
        if variant_label:
            await insert_personal_variant(match_id, self.user_id, variant_label)

        _clear_flow_draft(self.user_id)
        result_label = "勝ち" if is_win else "負け"
        self.clear_items()
        self.add_item(RepeatButton())
        self.add_item(UndoButton(match_id))
        tag_note = f"\n🏷 検証タグ「{variant_label}」を付けて記録しました。" if variant_label else ""
        await interaction.response.edit_message(
            content=self.screen(
                f"✅ **{result_label}** で記録しました。{tag_note}\n"
                f"続けて記録する場合は「同じ設定でもう1試合」、押し間違えた場合は「今の記録を取り消す」を押してください。\n"
                f"-# 📈 ランクやグレードが変わったときは、次の記録画面の一番下の青いボタンで更新できます。"
            ),
            view=self,
        )

    async def on_timeout(self):
        # 元のインタラクションのフォローアップ編集は15分まで有効。
        # timeout=780(13分)なので基本は成功するが、念のためHTTPExceptionは握りつぶす。
        if self.user_id in _FLOW_DRAFTS:
            text = (
                "⏱ 入力の受付時間が終了しました。\n"
                "パネルの「⚔️ 戦績を記録する」から開き直すと、選んでいた内容から再開できます。"
            )
        else:
            text = (
                "⏱ 入力の受付時間が終了しました。\n"
                "お手数ですが、パネルの「⚔️ 戦績を記録する」からもう一度お願いします。"
            )
        try:
            await self.origin_interaction.edit_original_response(content=text, view=None)
        except discord.HTTPException:
            pass


def _resolve_grade(rank_value: str, raw_grade):
    """
    ランク帯とグレードの整合を取る。
    戻り値: (保存するグレード, 補足メッセージ, エラーメッセージ)

    グランドマスターはグレード指定を必須にするが、「グレードなし」も
    正規の選択肢として認める（昇格直後はCRが1650に届かないため）。
    """
    if rank_value == GRAND_MASTER_TIER:
        if raw_grade is None:
            return None, None, (
                f"{GRAND_MASTER_TIER}を選んだ場合は `グレード` も指定してください。\n"
                f"CRが1650未満でグレードが付いていない場合は"
                f"「{CR_GRADE_NONE_LABEL}」を選んでください。"
            )
        if raw_grade == CR_GRADE_NONE_VALUE:
            return None, None, None
        return raw_grade, None, None

    # グランドマスター以外にCRは存在しない
    if raw_grade is not None and raw_grade != CR_GRADE_NONE_VALUE:
        return None, (
            f"※グレードは{GRAND_MASTER_TIER}帯でのみ記録されるため、今回は保存していません。"
        ), None
    return None, None, None


# ==============================
# /デッキ切替 のUI
# ==============================

def _template_label(t: dict) -> str:
    fmt = "ロテ" if t["format"] == "rotation" else "アンリミ"
    return f"{t['deck_name']}（{t['my_class']} / {fmt}）"


class TemplateSelect(discord.ui.Select):
    NEW_VALUE = "__new__"

    def __init__(self, templates: list[dict], mode: str):
        self.mode = mode  # "switch" or "delete"
        options = []
        if mode == "switch":
            # 候補にないデッキを、この場で登録してそのまま使えるようにする
            options.append(
                discord.SelectOption(label="➕ 新しいデッキを登録する…", value=self.NEW_VALUE)
            )
        limit = 24 if mode == "switch" else 25
        options.extend(
            discord.SelectOption(label=_template_label(t)[:100], value=str(t["id"]))
            for t in templates[:limit]
        )
        placeholder = "使うデッキを選択" if mode == "switch" else "削除するデッキを選択"
        super().__init__(placeholder=placeholder, options=options, min_values=1, max_values=1)

    async def callback(self, interaction: discord.Interaction):
        user_id = str(interaction.user.id)

        if self.values[0] == self.NEW_VALUE:
            view = SetupFlowView(user_id, mode="newdeck")
            await interaction.response.edit_message(
                content=view.header("フォーマットを選んでください。"), view=view
            )
            return

        template_id = int(self.values[0])

        if self.mode == "delete":
            ok = await delete_template(user_id, template_id)
            msg = "🗑️ 削除しました。" if ok else "削除できませんでした。"
            await interaction.response.edit_message(content=msg, view=None)
            return

        t = await apply_template(user_id, template_id)
        if t is None:
            await interaction.response.edit_message(
                content="そのデッキが見つかりませんでした。", view=None
            )
            return
        settings = await get_user_settings(user_id)
        rank = settings.get("rank_tier") or "未設定"
        if settings.get("cr_grade"):
            rank += f"（{settings['cr_grade']}）"
        await interaction.response.edit_message(
            content=(
                f"✅ デッキを切り替えました\n"
                f"{_template_label(t)}\n"
                f"ランク帯：{rank}（変更していません）"
            ),
            view=None,
        )
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()


class TemplateView(discord.ui.View):
    def __init__(self, templates: list[dict], mode: str = "switch"):
        super().__init__(timeout=180)
        self.add_item(TemplateSelect(templates, mode))


# ==============================
# 初回設定フロー / 新規デッキ登録フロー（ボタンから開く）
# ==============================

class SetupFormatSelect(discord.ui.Select):
    def __init__(self):
        options = [
            discord.SelectOption(label="ローテーション", value="rotation"),
            discord.SelectOption(label="アンリミテッド", value="unlimited"),
        ]
        super().__init__(placeholder="フォーマットを選択", options=options)

    async def callback(self, interaction: discord.Interaction):
        view: SetupFlowView = self.view
        view.format = self.values[0]
        view.show_class()
        await interaction.response.edit_message(content=view.header("クラスを選んでください。"), view=view)


class SetupClassSelect(discord.ui.Select):
    def __init__(self):
        options = [discord.SelectOption(label=c, value=c) for c in CLASS_CHOICES]
        super().__init__(placeholder="使用クラスを選択", options=options)

    async def callback(self, interaction: discord.Interaction):
        view: SetupFlowView = self.view
        view.my_class = self.values[0]
        known = await get_deck_names(view.my_class)
        view.show_deck(known)
        await interaction.response.edit_message(
            content=view.header("デッキを選んでください。候補になければ「新しく入力する…」から入力できます。"),
            view=view,
        )


class SetupDeckSelect(discord.ui.Select):
    NEW_VALUE = "__new__"

    def __init__(self, known_decks: list):
        options = [discord.SelectOption(label="新しく入力する…", value=self.NEW_VALUE)]
        for name in known_decks[:24]:
            options.append(discord.SelectOption(label=name[:100], value=name[:100]))
        super().__init__(placeholder="デッキを選択", options=options)

    async def callback(self, interaction: discord.Interaction):
        view: SetupFlowView = self.view
        if self.values[0] == self.NEW_VALUE:
            await interaction.response.send_modal(SetupDeckModal(view))
            return
        view.deck = self.values[0]
        await view.after_deck(interaction)


class SetupDeckModal(discord.ui.Modal, title="デッキ名を入力"):
    deck_name = discord.ui.TextInput(
        label="デッキ名", placeholder="例：ハイランダーネメシス", required=True, max_length=50
    )

    def __init__(self, flow_view):
        super().__init__()
        self.flow_view = flow_view

    async def on_submit(self, interaction: discord.Interaction):
        view = self.flow_view
        name = str(self.deck_name.value).strip()
        if not name:
            await interaction.response.edit_message(
                content="デッキ名が空でした。最初からやり直してください。", view=None
            )
            return
        view.deck = name
        await view.after_deck(interaction)


class SetupRankSelect(discord.ui.Select):
    def __init__(self):
        options = [discord.SelectOption(label=r, value=r) for r in RANK_TIER_CHOICES]
        super().__init__(placeholder="ランク帯を選択", options=options)

    async def callback(self, interaction: discord.Interaction):
        view: SetupFlowView = self.view
        view.rank = self.values[0]
        if view.rank == GRAND_MASTER_TIER:
            view.show_grade()
            await interaction.response.edit_message(
                content=view.header(
                    f"グレードを選んでください（付いていなければ「{CR_GRADE_NONE_LABEL}」でOK）。"
                ),
                view=view,
            )
            return
        view.grade = None
        await view.finalize(interaction)


class SetupGradeSelect(discord.ui.Select):
    def __init__(self):
        options = (
            [discord.SelectOption(label=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
            + [discord.SelectOption(label=g, value=g) for g in CR_GRADE_CHOICES]
        )
        super().__init__(placeholder="グレードを選択", options=options)

    async def callback(self, interaction: discord.Interaction):
        view: SetupFlowView = self.view
        grade_value, _note, _err = _resolve_grade(view.rank, self.values[0])
        view.grade = grade_value
        await view.finalize(interaction)


class SetupFlowView(discord.ui.View):
    """
    mode="setup"   … 初回設定（フォーマット→クラス→デッキ→ランク→グレード）
    mode="newdeck" … デッキの新規登録（フォーマット→クラス→デッキ）。ランクは触らない
    """

    def __init__(self, user_id: str, mode: str = "setup"):
        super().__init__(timeout=300)
        self.user_id = user_id
        self.mode = mode
        self.format = None
        self.my_class = None
        self.deck = None
        self.rank = None
        self.grade = None
        self.show_format()

    def header(self, body: str) -> str:
        title = "🆕 **初回設定**" if self.mode == "setup" else "➕ **デッキの新規登録**"
        chosen = []
        if self.format:
            chosen.append(FORMAT_LABELS.get(self.format, self.format))
        if self.my_class:
            chosen.append(self.my_class)
        if self.deck:
            chosen.append(self.deck)
        line = " / ".join(chosen)
        return f"{title}\n{line}\n\n{body}" if line else f"{title}\n\n{body}"

    def show_format(self):
        self.clear_items()
        self.add_item(SetupFormatSelect())

    def show_class(self):
        self.clear_items()
        self.add_item(SetupClassSelect())

    def show_deck(self, known_decks: list):
        self.clear_items()
        self.add_item(SetupDeckSelect(known_decks))

    def show_rank(self):
        self.clear_items()
        self.add_item(SetupRankSelect())

    def show_grade(self):
        self.clear_items()
        self.add_item(SetupGradeSelect())

    async def after_deck(self, interaction: discord.Interaction):
        """デッキ名が決まったあとの分岐"""
        if self.mode == "newdeck":
            await self.finalize(interaction)
            return
        self.show_rank()
        await interaction.response.edit_message(
            content=self.header("ランク帯を選んでください。"), view=self
        )

    async def finalize(self, interaction: discord.Interaction):
        # どちらのモードでもテンプレートに登録しておく（次から選ぶだけで済む）
        await add_template(self.user_id, self.format, self.my_class, self.deck)

        if self.mode == "setup":
            await upsert_user_settings(
                self.user_id, self.format, self.my_class, self.deck, self.rank, self.grade
            )
            rank_label = self.rank
            if self.grade:
                rank_label += f"（{self.grade}）"
            elif self.rank == GRAND_MASTER_TIER:
                rank_label += "・グレードなし"
            body = (
                "✅ 設定が完了しました\n"
                f"{FORMAT_LABELS.get(self.format, self.format)} / {self.my_class} / "
                f"**{self.deck}** / ランク：**{rank_label}**\n\n"
                "パネルの「⚔️ 戦績を記録する」から記録を始められます。"
            )
        else:
            # 既存のランク・グレードは維持したまま、使用デッキだけ差し替える
            current = await get_user_settings(self.user_id)
            rank = current.get("rank_tier") if current else None
            grade = current.get("cr_grade") if current else None
            await upsert_user_settings(
                self.user_id, self.format, self.my_class, self.deck, rank, grade
            )
            body = (
                "✅ 登録して、使用デッキに設定しました\n"
                f"{FORMAT_LABELS.get(self.format, self.format)} / {self.my_class} / **{self.deck}**\n"
                "ランク帯は変更していません。"
            )

        await interaction.response.edit_message(content=body, view=None)
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()


# ==============================
# 常設パネル（/戦績パネル設置 で1回置く。ボタンは押した人にだけ反応する）
# ==============================

PANEL_SETUP_CUSTOM_ID = "senseki_panel_setup_v1"
PANEL_CHECK_CUSTOM_ID = "senseki_panel_check_v1"
PANEL_WEAK_CUSTOM_ID = "senseki_panel_weak_v1"
PANEL_FORMAT_CUSTOM_ID = "senseki_panel_format_v1"
PANEL_RECORD_CUSTOM_ID = "senseki_panel_record_v1"
PANEL_DECK_SWITCH_CUSTOM_ID = "senseki_panel_deck_switch_v1"
PANEL_RANK_UPDATE_CUSTOM_ID = "senseki_panel_rank_update_v1"
PANEL_UNDO_CUSTOM_ID = "senseki_panel_undo_v1"


class SensekiPanelButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="戦績を記録する",
            style=discord.ButtonStyle.primary,
            emoji="⚔️",
            custom_id=PANEL_RECORD_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        cog = interaction.client.get_cog("SensekiCog")
        if cog is None:
            await interaction.response.send_message(
                "内部エラー：SensekiCogが見つかりません。", ephemeral=True
            )
            return
        # ここでのレスポンスは押した本人にだけ見える（ephemeral）。
        # パネル本体（誰でも見えるメッセージ）は書き換えない。
        await cog._start_record_flow(interaction)


class PanelDeckSwitchButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="デッキ切替",
            style=discord.ButtonStyle.secondary,
            emoji="🔄",
            custom_id=PANEL_DECK_SWITCH_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        user_id = str(interaction.user.id)
        settings = await get_user_settings(user_id)
        if settings is None:
            await interaction.response.send_message(
                "先に `/戦績 設定` で初期設定をしてください。", ephemeral=True
            )
            return
        templates = await list_templates(user_id)
        if not templates:
            # 1つも登録がなければ、そのまま新規登録フローに入る
            view = SetupFlowView(user_id, mode="newdeck")
            await interaction.response.send_message(
                view.header("登録済みのデッキがありません。フォーマットを選んでください。"),
                view=view,
                ephemeral=True,
            )
            return
        await interaction.response.send_message(
            "使うデッキを選んでください。候補にない場合は「➕ 新しいデッキを登録する…」を選んでください。",
            view=TemplateView(templates, "switch"),
            ephemeral=True,
        )


class PanelRankTierSelect(discord.ui.Select):
    def __init__(self):
        options = [discord.SelectOption(label=r, value=r) for r in RANK_TIER_CHOICES]
        super().__init__(placeholder="ランク帯を選択", options=options, min_values=1, max_values=1)

    async def callback(self, interaction: discord.Interaction):
        rank_value = self.values[0]

        if rank_value == GRAND_MASTER_TIER:
            view = discord.ui.View(timeout=180)
            view.add_item(PanelGradeSelect(rank_value))
            await interaction.response.edit_message(
                content=(
                    f"ランク帯：**{rank_value}**\n"
                    f"グレードを選択してください（付いていなければ"
                    f"「{CR_GRADE_NONE_LABEL}」でOK）。"
                ),
                view=view,
            )
            return

        ok = await update_rank(str(interaction.user.id), rank_value, None)
        if not ok:
            await interaction.response.edit_message(
                content="先に `/戦績 設定` で初期設定をしてください。", view=None
            )
            return
        await interaction.response.edit_message(
            content=f"✅ ランク帯を更新しました：{rank_value}", view=None
        )
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()


class PanelGradeSelect(discord.ui.Select):
    def __init__(self, rank_value: str):
        self.rank_value = rank_value
        options = (
            [discord.SelectOption(label=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
            + [discord.SelectOption(label=g, value=g) for g in CR_GRADE_CHOICES]
        )
        super().__init__(placeholder="グレードを選択", options=options, min_values=1, max_values=1)

    async def callback(self, interaction: discord.Interaction):
        raw_grade = self.values[0]
        grade_value, _note, _error = _resolve_grade(self.rank_value, raw_grade)

        ok = await update_rank(str(interaction.user.id), self.rank_value, grade_value)
        if not ok:
            await interaction.response.edit_message(
                content="先に `/戦績 設定` で初期設定をしてください。", view=None
            )
            return
        label = f"{self.rank_value}（{grade_value}）" if grade_value else f"{self.rank_value}・グレードなし"
        await interaction.response.edit_message(
            content=f"✅ ランク帯を更新しました：{label}", view=None
        )
        cog = interaction.client.get_cog("SensekiCog")
        if cog is not None:
            await cog._update_board()


class PanelRankUpdateButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="ランク更新",
            style=discord.ButtonStyle.secondary,
            emoji="📈",
            custom_id=PANEL_RANK_UPDATE_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None:
            await interaction.response.send_message(
                "先に `/戦績 設定` で初期設定をしてください。", ephemeral=True
            )
            return
        view = discord.ui.View(timeout=180)
        view.add_item(PanelRankTierSelect())
        await interaction.response.send_message(
            "ランク帯を選んでください。", view=view, ephemeral=True
        )


class PanelUndoButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="直前の記録を取消",
            style=discord.ButtonStyle.danger,
            emoji="↩️",
            custom_id=PANEL_UNDO_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        deleted = await delete_last_match(str(interaction.user.id))
        if deleted is None:
            await interaction.response.send_message(
                "取り消せる記録が見つかりませんでした。", ephemeral=True
            )
            return
        result_label = "勝ち" if deleted["is_win"] else "負け"
        await interaction.response.send_message(
            f"🗑️ 直前の記録を取り消しました\n相手クラス：{deleted['opp_class']} / {result_label}",
            ephemeral=True,
        )


class PanelSetupButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="初回設定",
            style=discord.ButtonStyle.success,
            emoji="🆕",
            custom_id=PANEL_SETUP_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        user_id = str(interaction.user.id)
        settings = await get_user_settings(user_id)
        if settings is not None:
            # 共有パネルなのでボタン自体は隠せない。代わりに現状を返す
            deck = settings.get("my_deck") or "デッキ未設定"
            rank = settings.get("rank_tier") or "ランク未設定"
            if settings.get("cr_grade"):
                rank += f"・{settings['cr_grade']}"
            elif settings.get("rank_tier") == GRAND_MASTER_TIER:
                rank += "・グレードなし"
            view = discord.ui.View(timeout=180)
            view.add_item(PanelRedoSetupButton())
            await interaction.response.send_message(
                f"✅ すでに設定済みです\n"
                f"**{deck}**（{settings['my_class']} / "
                f"{FORMAT_LABELS.get(settings['format'], settings['format'])}） / "
                f"ランク：**{rank}**\n\n"
                "デッキやランクを変えるだけなら「⚙️ 設定を変える」が早いです。",
                view=view,
                ephemeral=True,
            )
            return
        await interaction.response.send_message(
            SetupFlowView(user_id).header("フォーマットを選んでください。"),
            view=SetupFlowView(user_id),
            ephemeral=True,
        )


class PanelRedoSetupButton(discord.ui.Button):
    """設定済みの人が、あえて最初からやり直すためのボタン"""

    def __init__(self):
        super().__init__(label="最初から設定し直す", style=discord.ButtonStyle.secondary)

    async def callback(self, interaction: discord.Interaction):
        user_id = str(interaction.user.id)
        view = SetupFlowView(user_id)
        await interaction.response.edit_message(
            content=view.header("フォーマットを選んでください。"), view=view
        )


class PanelFormatSelect(discord.ui.Select):
    def __init__(self):
        options = [
            discord.SelectOption(label="ローテーション", value="rotation"),
            discord.SelectOption(label="アンリミテッド", value="unlimited"),
        ]
        super().__init__(placeholder="フォーマットを選択", options=options)

    async def callback(self, interaction: discord.Interaction):
        value = self.values[0]
        ok = await update_format(str(interaction.user.id), value)
        if not ok:
            await interaction.response.edit_message(
                content="先に「🆕 初回設定」を済ませてください。", view=None
            )
            return
        await interaction.response.edit_message(
            content=f"✅ フォーマットを **{FORMAT_LABELS[value]}** に変更しました。", view=None
        )


class PanelFormatButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="フォーマット切替",
            style=discord.ButtonStyle.secondary,
            emoji="🔀",
            custom_id=PANEL_FORMAT_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None:
            await interaction.response.send_message(
                "先に「🆕 初回設定」を済ませてください。", ephemeral=True
            )
            return
        now = FORMAT_LABELS.get(settings["format"], settings["format"])
        view = discord.ui.View(timeout=180)
        view.add_item(PanelFormatSelect())
        await interaction.response.send_message(
            f"現在：**{now}**\n切り替え先を選んでください。", view=view, ephemeral=True
        )



class PanelWeakButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="弱点対面",
            style=discord.ButtonStyle.secondary,
            emoji="📉",
            custom_id=PANEL_WEAK_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        cog = interaction.client.get_cog("SensekiCog")
        if cog is None:
            await interaction.response.send_message("内部エラーです。", ephemeral=True)
            return
        await cog._show_weak_matchups(interaction)


PANEL_VARIANT_CUSTOM_ID = "senseki_panel_variant_v1"
PANEL_SETTINGS_CUSTOM_ID = "senseki_panel_settings_v1"
PANEL_AI_CUSTOM_ID = "senseki_panel_ai_v1"


class VariantTagModal(discord.ui.Modal, title="検証タグを設定"):
    tag_name = discord.ui.TextInput(
        label="タグ名",
        placeholder="例：アンテマリア採用型、フルスペル型 など",
        required=True,
        max_length=50,
    )

    def __init__(self, settings: dict):
        super().__init__()
        self.settings = settings

    async def on_submit(self, interaction: discord.Interaction):
        label = str(self.tag_name.value).strip()
        if not label:
            await interaction.response.send_message("タグ名を入力してください。", ephemeral=True)
            return
        await set_active_variant(
            str(interaction.user.id), self.settings["format"], self.settings["my_class"],
            self.settings["my_deck"], label,
        )
        await interaction.response.send_message(
            f"🏷 検証タグ「{label}」を開始しました。\n"
            f"これ以降、**{self.settings['my_deck']}**で記録する対戦に自動で付与されます"
            f"（この記録はあなたにしか見えません）。\n"
            f"やめる場合は「⚙️ 設定を変える」→「🧪 デッキ検証」→「タグを解除」を押してください。",
            ephemeral=True,
        )


class VariantSetButton(discord.ui.Button):
    def __init__(self, settings: dict):
        self.settings = settings
        super().__init__(label="タグを設定・変更", style=discord.ButtonStyle.primary)

    async def callback(self, interaction: discord.Interaction):
        await interaction.response.send_modal(VariantTagModal(self.settings))


class VariantClearButton(discord.ui.Button):
    def __init__(self, settings: dict):
        self.settings = settings
        super().__init__(label="タグを解除", style=discord.ButtonStyle.secondary)

    async def callback(self, interaction: discord.Interaction):
        ok = await clear_active_variant(
            str(interaction.user.id), self.settings["format"], self.settings["my_class"],
            self.settings["my_deck"],
        )
        if ok:
            msg = "🏷 検証タグを解除しました。以降の記録は通常通りタグなしで残ります。"
        else:
            msg = "現在、有効な検証タグはありません。"
        await interaction.response.edit_message(content=msg, view=None)


class VariantSummaryButton(discord.ui.Button):
    def __init__(self, settings: dict):
        self.settings = settings
        super().__init__(label="集計を見る", style=discord.ButtonStyle.secondary)

    async def callback(self, interaction: discord.Interaction):
        rows = await get_variant_summary(
            str(interaction.user.id), self.settings["format"], self.settings["my_class"],
            self.settings["my_deck"],
        )
        if not rows:
            await interaction.response.edit_message(
                content=(
                    f"**{self.settings['my_deck']}**にはまだ検証タグ付きの記録がありません。\n"
                    f"「タグを設定・変更」からタグを付けてから対戦を記録してください。"
                ),
                view=None,
            )
            return
        active = await get_active_variant(
            str(interaction.user.id), self.settings["format"], self.settings["my_class"],
            self.settings["my_deck"],
        )
        lines = [f"🧪 **{self.settings['my_deck']}** 個人検証タグ別勝率（あなただけに表示）"]
        for r in rows:
            total = r["total"]
            wins = r["wins"] or 0
            rate = wins / total * 100 if total else 0
            mark = " ▶️稼働中" if active and r["variant_label"] == active else ""
            lines.append(
                f"・**{r['variant_label']}**：勝率 {rate:.1f}%"
                f"（{wins}勝{total - wins}敗・{total}戦）{mark}"
            )
        lines.append("-# この集計は生データ（デッキ名）には反映されません。あなただけが見られます。")
        await interaction.response.edit_message(content="\n".join(lines)[:1900], view=None)


class PanelVariantButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="デッキ検証",
            style=discord.ButtonStyle.secondary,
            emoji="🧪",
            custom_id=PANEL_VARIANT_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず「🆕 初回設定」か「⚙️ 設定を変える」→「🔄 デッキ切替」でデッキを登録してください。",
                ephemeral=True,
            )
            return
        active = await get_active_variant(
            str(interaction.user.id), settings["format"], settings["my_class"], settings["my_deck"]
        )
        active_note = f"\n現在のタグ：**{active}**" if active else "\n現在、タグは設定されていません。"
        view = discord.ui.View(timeout=180)
        view.add_item(VariantSetButton(settings))
        view.add_item(VariantClearButton(settings))
        view.add_item(VariantSummaryButton(settings))
        await interaction.response.send_message(
            f"🧪 **{settings['my_deck']}** の個人検証タグ管理{active_note}\n"
            f"（カード1枚単位で構築を変えた場合などに、勝率だけ自分用に分けて記録できます。"
            f"公開集計には一切反映されません）",
            view=view,
            ephemeral=True,
        )


class _SubmenuButton(discord.ui.Button):
    """設定メニュー内のボタン。パネルの既存ボタンと同じ処理を呼ぶ（custom_idは持たない）"""

    def __init__(self, label: str, emoji: str, handler):
        super().__init__(label=label, emoji=emoji, style=discord.ButtonStyle.secondary)
        self.handler = handler

    async def callback(self, interaction: discord.Interaction):
        await self.handler(self, interaction)


async def send_settings_menu(interaction: discord.Interaction):
    """「⚙️ 設定を変える」：デッキ・ランク・フォーマット・検証タグの入口をまとめて出す"""
    settings = await get_user_settings(str(interaction.user.id))
    if settings is None:
        await interaction.response.send_message("先に「🆕 初回設定」を済ませてください。", ephemeral=True)
        return
    view = discord.ui.View(timeout=180)
    view.add_item(_SubmenuButton("デッキ切替", "🔄", PanelDeckSwitchButton.callback))
    view.add_item(_SubmenuButton("ランク更新", "📈", PanelRankUpdateButton.callback))
    view.add_item(_SubmenuButton("フォーマット切替", "🔀", PanelFormatButton.callback))
    view.add_item(_SubmenuButton("デッキ検証", "🧪", PanelVariantButton.callback))
    deck = settings.get("my_deck") or "デッキ未設定"
    rank = settings.get("rank_tier") or "ランク未設定"
    if settings.get("cr_grade"):
        rank += f"・{settings['cr_grade']}"
    await interaction.response.send_message(
        f"⚙️ 今の設定：**{deck}**（{settings['my_class']} / "
        f"{FORMAT_LABELS.get(settings['format'], settings['format'])}） / ランク：**{rank}**\n"
        "変えたい項目を選んでください。\n"
        "-# 🧪 デッキ検証：今のデッキに自分だけの検証タグを付けて、型ごとの勝率を比べられます。",
        view=view,
        ephemeral=True,
    )


class PanelSettingsButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="設定を変える", style=discord.ButtonStyle.secondary, emoji="⚙️",
            custom_id=PANEL_SETTINGS_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        await send_settings_menu(interaction)


class _FollowupInteraction:
    """
    すでに返信済みのインタラクションで、もう1通メッセージを送るための薄いラッパー。
    response.send_message を followup.send に置き換え、それ以外は元のインタラクションに任せる。
    「📊 成績・弱点」で、成績と弱点対面を続けて表示するために使う。
    """

    class _Response:
        def __init__(self, interaction):
            self._i = interaction

        async def send_message(self, content=None, **kwargs):
            kwargs.setdefault("ephemeral", True)
            await self._i.followup.send(content, **kwargs)


    def __init__(self, interaction):
        self._interaction = interaction
        self.response = self._Response(interaction)

    def __getattr__(self, name):
        return getattr(self._interaction, name)


class PanelCheckButton(discord.ui.Button):
    """成績確認と弱点対面をまとめた入口"""

    def __init__(self):
        super().__init__(
            label="成績・弱点",
            style=discord.ButtonStyle.secondary,
            emoji="📊",
            custom_id=PANEL_CHECK_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        cog = interaction.client.get_cog("SensekiCog")
        if cog is None:
            await interaction.response.send_message("内部エラーです。", ephemeral=True)
            return
        await cog._show_check(interaction)
        # 記録があるときだけ、続けて弱点対面を出す（記録が無い人に同じ案内を2回出さない）
        settings = await get_user_settings(str(interaction.user.id))
        if settings and settings.get("my_deck"):
            total = await get_deck_total(str(interaction.user.id), settings["format"], settings["my_deck"])
            if total > 0:
                await cog._show_weak_matchups(_FollowupInteraction(interaction))


class PanelAIButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="AIで分析", style=discord.ButtonStyle.secondary, emoji="🤖",
            custom_id=PANEL_AI_CUSTOM_ID,
        )

    async def callback(self, interaction: discord.Interaction):
        cog = interaction.client.get_cog("SensekiCog")
        if cog is None:
            await interaction.response.send_message("内部エラーです。", ephemeral=True)
            return
        await cog._show_personal_ai(interaction)


# ---- 記録の確認・修正・削除・復元（Discord） ----
PANEL_RECORDS_CUSTOM_ID = "senseki_panel_records_v1"
NO_DECK_VALUE = "__none__"


def _match_label(m: dict) -> str:
    wl = "勝" if m["is_win"] else "負"
    deck = f"({m['opp_deck']})" if m.get("opp_deck") else ""
    ts = (m.get("recorded_at") or "")[5:16].replace("T", " ").replace("-", "/")
    return f"{wl} vs {m['opp_class']}{deck} {'先' if m['is_first'] else '後'} {ts}"[:100]


def _match_summary(cls: str, deck, is_first: bool, is_win: bool) -> str:
    d = f"（{deck}）" if deck else ""
    return f"{'勝ち' if is_win else '負け'}／{cls}{d}／{'先攻' if is_first else '後攻'}"


def _web_url_button():
    base = os.environ.get("WEB_BASE_URL", "").rstrip("/")
    if not base:
        return None
    return discord.ui.Button(label="Webで見る", style=discord.ButtonStyle.link, url=base, emoji="🌐", row=2)


async def build_record_list(user_id: str):
    rows = await list_user_matches(user_id, 10)
    lines = ["📋 **あなたの直近の記録**（今の環境・新しい順・10件まで）"]
    if rows:
        lines.append("修正・削除したい記録を、下から選んでください。")
    else:
        lines.append("まだ今の環境の記録がありません。")
    lines.append(f"-# 修正と削除は1日{EDIT_DAILY_LIMIT}件まで。削除しても{RESTORE_DAYS}日以内なら自分で戻せます。前の環境の記録は変更できません。")
    return "\n".join(lines), RecordListView(user_id, rows)


class RecordPickSelect(discord.ui.Select):
    def __init__(self, rows):
        super().__init__(
            placeholder="修正・削除する記録を選ぶ",
            options=[discord.SelectOption(label=_match_label(m), value=str(m["id"])) for m in rows[:25]],
            row=0,
        )

    async def callback(self, interaction: discord.Interaction):
        uid = str(interaction.user.id)
        m = await get_user_match(uid, int(self.values[0]))
        if m is None:
            await interaction.response.send_message("記録が見つかりません。", ephemeral=True)
            return
        view = RecordEditView(uid, m)
        await view.refresh_deck_options()
        await interaction.response.edit_message(content=view.content(), view=view)


class RecordListView(discord.ui.View):
    def __init__(self, user_id: str, rows):
        super().__init__(timeout=300)
        self.user_id = user_id
        if rows:
            self.add_item(RecordPickSelect(rows))
        self.add_item(RecordRestoreOpenButton())
        link = _web_url_button()
        if link is not None:
            self.add_item(link)


class RecordRestoreOpenButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="削除した記録を戻す", style=discord.ButtonStyle.secondary, emoji="♻️", row=1)

    async def callback(self, interaction: discord.Interaction):
        uid = str(interaction.user.id)
        rows = await list_user_matches(uid, 10, deleted=True)
        if not rows:
            await interaction.response.send_message(
                f"戻せる記録はありません（削除から{RESTORE_DAYS}日以内・今の環境の記録だけ）。", ephemeral=True)
            return
        view = discord.ui.View(timeout=300)
        view.add_item(RecordRestoreSelect(rows))
        view.add_item(RecordBackButton())
        await interaction.response.edit_message(
            content="♻️ **削除した記録**（新しい順）\n戻したい記録を選ぶと、すぐに元に戻ります。", view=view)


class RecordRestoreSelect(discord.ui.Select):
    def __init__(self, rows):
        super().__init__(
            placeholder="戻す記録を選ぶ",
            options=[discord.SelectOption(label=_match_label(m), value=str(m["id"])) for m in rows[:25]],
            row=0,
        )

    async def callback(self, interaction: discord.Interaction):
        ok, msg, _ = await restore_match(str(interaction.user.id), int(self.values[0]), "discord")
        view = discord.ui.View(timeout=120)
        view.add_item(RecordBackButton())
        await interaction.response.edit_message(content=("✅ " if ok else "⚠️ ") + msg, view=view)


class RecordBackButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="一覧に戻る", style=discord.ButtonStyle.secondary, row=4)

    async def callback(self, interaction: discord.Interaction):
        content, view = await build_record_list(str(interaction.user.id))
        await interaction.response.edit_message(content=content, view=view)


class RecordClassSelect(discord.ui.Select):
    def __init__(self, current: str):
        super().__init__(
            placeholder="相手のクラス",
            options=[discord.SelectOption(label=c, value=c, default=(c == current)) for c in CLASS_CHOICES],
            row=0,
        )

    async def callback(self, interaction: discord.Interaction):
        v: RecordEditView = self.view
        if self.values[0] != v.cls:
            v.cls = self.values[0]
            v.deck = None
            await v.refresh_deck_options()
        await interaction.response.edit_message(content=v.content(), view=v)


class RecordDeckSelect(discord.ui.Select):
    def __init__(self, names, current):
        opts = [discord.SelectOption(label="（デッキなし／不明）", value=NO_DECK_VALUE, default=not current)]
        for n in names[:24]:
            opts.append(discord.SelectOption(label=n[:100], value=n, default=(n == current)))
        super().__init__(placeholder="相手のデッキ", options=opts, row=1)

    async def callback(self, interaction: discord.Interaction):
        v: RecordEditView = self.view
        v.deck = None if self.values[0] == NO_DECK_VALUE else self.values[0]
        await v.refresh_deck_options()
        await interaction.response.edit_message(content=v.content(), view=v)


class RecordToggleButton(discord.ui.Button):
    def __init__(self, key: str, value: bool, label: str):
        super().__init__(label=label, row=2)
        self.key, self.value = key, value

    async def callback(self, interaction: discord.Interaction):
        v: RecordEditView = self.view
        setattr(v, self.key, self.value)
        v.sync_buttons()
        await interaction.response.edit_message(content=v.content(), view=v)


class RecordSaveButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="保存", style=discord.ButtonStyle.success, emoji="💾", row=3)

    async def callback(self, interaction: discord.Interaction):
        v: RecordEditView = self.view
        ok, msg, _ = await edit_match(v.user_id, v.match["id"], {
            "opp_class": v.cls, "opp_deck": v.deck, "is_first": v.is_first, "is_win": v.is_win}, "discord")
        back = discord.ui.View(timeout=120)
        back.add_item(RecordBackButton())
        await interaction.response.edit_message(
            content=("✅ " if ok else "⚠️ ") + msg + (f"\n{_match_summary(v.cls, v.deck, v.is_first, v.is_win)}" if ok else ""),
            view=back)


class RecordDeleteButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="削除", style=discord.ButtonStyle.danger, emoji="🗑️", row=3)

    async def callback(self, interaction: discord.Interaction):
        v: RecordEditView = self.view
        m = v.match
        view = RecordDeleteConfirmView(v.user_id, m["id"])
        await interaction.response.edit_message(
            content=("🗑️ **この記録を削除しますか？**\n"
                     f"{_match_summary(m['opp_class'], m.get('opp_deck'), bool(m['is_first']), bool(m['is_win']))}\n"
                     f"-# 削除すると集計から外れます。{RESTORE_DAYS}日以内なら「♻️ 削除した記録を戻す」で戻せます。"),
            view=view)


class RecordDeleteConfirmView(discord.ui.View):
    def __init__(self, user_id: str, match_id: int):
        super().__init__(timeout=120)
        self.user_id, self.match_id = user_id, match_id

    @discord.ui.button(label="削除する", style=discord.ButtonStyle.danger)
    async def do_delete(self, interaction: discord.Interaction, button: discord.ui.Button):
        ok, msg, _ = await delete_match_checked(self.user_id, self.match_id, "discord")
        back = discord.ui.View(timeout=120)
        back.add_item(RecordBackButton())
        await interaction.response.edit_message(content=("✅ " if ok else "⚠️ ") + msg, view=back)

    @discord.ui.button(label="やめる", style=discord.ButtonStyle.secondary)
    async def cancel(self, interaction: discord.Interaction, button: discord.ui.Button):
        content, view = await build_record_list(self.user_id)
        await interaction.response.edit_message(content=content, view=view)


class RecordEditView(discord.ui.View):
    def __init__(self, user_id: str, match: dict):
        super().__init__(timeout=300)
        self.user_id, self.match = user_id, match
        self.cls = match["opp_class"]
        self.deck = match.get("opp_deck") or None
        self.is_first = bool(match["is_first"])
        self.is_win = bool(match["is_win"])
        self.add_item(RecordClassSelect(self.cls))
        self.add_item(RecordDeckSelect([], self.deck))
        self.btn_first = RecordToggleButton("is_first", True, "先攻")
        self.btn_second = RecordToggleButton("is_first", False, "後攻")
        self.btn_win = RecordToggleButton("is_win", True, "勝ち")
        self.btn_lose = RecordToggleButton("is_win", False, "負け")
        for b in (self.btn_first, self.btn_second, self.btn_win, self.btn_lose):
            self.add_item(b)
        self.save_btn = RecordSaveButton()
        self.add_item(self.save_btn)
        self.add_item(RecordDeleteButton())
        back = RecordBackButton()
        back.row = 3
        self.add_item(back)
        self.sync_buttons()

    async def refresh_deck_options(self):
        names = await get_deck_names(self.cls, 24)
        if self.deck and self.deck not in names:
            names = [self.deck] + names[:23]
        for item in list(self.children):
            if isinstance(item, RecordDeckSelect):
                self.remove_item(item)
        sel = RecordDeckSelect(names, self.deck)
        self.add_item(sel)
        # 並びを保つ（クラス→デッキ→ボタン）。add_itemは末尾なので行指定で見た目を固定している
        self.sync_buttons()

    def changed(self) -> bool:
        m = self.match
        return (self.cls, self.deck, self.is_first, self.is_win) != (
            m["opp_class"], m.get("opp_deck") or None, bool(m["is_first"]), bool(m["is_win"]))

    def sync_buttons(self):
        on, off = discord.ButtonStyle.primary, discord.ButtonStyle.secondary
        self.btn_first.style = on if self.is_first else off
        self.btn_second.style = off if self.is_first else on
        self.btn_win.style = on if self.is_win else off
        self.btn_lose.style = off if self.is_win else on
        self.save_btn.disabled = not self.changed()

    def content(self) -> str:
        m = self.match
        before = _match_summary(m["opp_class"], m.get("opp_deck"), bool(m["is_first"]), bool(m["is_win"]))
        after = _match_summary(self.cls, self.deck, self.is_first, self.is_win)
        lines = ["✏️ **記録の修正**", f"変更前：{before}", f"変更後：{after}"]
        lines.append("内容を選び直して「💾 保存」を押してください。" if self.changed() else "-# 変更すると「保存」が押せるようになります。")
        return "\n".join(lines)


class PanelRecordsButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="記録の確認・修正", style=discord.ButtonStyle.secondary, emoji="📋",
                         custom_id=PANEL_RECORDS_CUSTOM_ID)

    async def callback(self, interaction: discord.Interaction):
        content, view = await build_record_list(str(interaction.user.id))
        await interaction.response.send_message(content, view=view, ephemeral=True)


class SensekiPanelView(discord.ui.View):
    """timeout=None＋固定custom_idで永続化。BOT再起動後もボタンが反応し続ける"""

    def __init__(self):
        super().__init__(timeout=None)
        # 1段目：記録と確認
        self.add_item(SensekiPanelButton())
        self.add_item(PanelUndoButton())
        self.add_item(PanelRecordsButton())
        self.add_item(PanelCheckButton())
        self.add_item(PanelAIButton())
        # 2段目：設定
        self.add_item(PanelSetupButton())
        self.add_item(PanelSettingsButton())
        # 質問・相談（ヒント、クラス別チャンネル・シャドバ相談所への質問）
        self.add_item(PanelHintButton())
        self.add_item(PanelAskButton())
        self.add_item(PanelVCButton())


class LegacyPanelView(discord.ui.View):
    """
    以前のパネルにだけあったボタン（弱点対面・デッキ切替・ランク更新・フォーマット切替・デッキ検証）。
    新しいパネルを設置し直すまでの間、古いパネルのボタンも動くように登録だけしておく。
    古いパネルを削除したら、このクラスと cog_load の登録行は消してよい。
    """

    def __init__(self):
        super().__init__(timeout=None)
        self.add_item(PanelWeakButton())
        self.add_item(PanelDeckSwitchButton())
        self.add_item(PanelRankUpdateButton())
        self.add_item(PanelFormatButton())
        self.add_item(PanelVariantButton())


# ==============================
# Excel（生データ＋集計）の生成
# ==============================

def _resolve_display_name(guild, user_id: str) -> str:
    if guild is not None:
        member = guild.get_member(int(user_id))
        if member is not None:
            return member.display_name
    return f"不明なユーザー（{user_id}）"


def build_workbook(matches: list[dict], guild) -> bytes:
    """生データシートと集計シートを持つxlsxをバイト列で返す"""
    wb = Workbook()

    # ---- 生データ ----
    ws_raw = wb.active
    ws_raw.title = "生データ"
    ws_raw.append([
        "記録日時", "ユーザー", "ユーザーID", "環境ID", "フォーマット",
        "自分のクラス", "自分のデッキ", "ランク帯", "グレード",
        "相手クラス", "相手デッキ", "先攻/後攻", "勝敗",
    ])
    for m in matches:
        ws_raw.append([
            m["recorded_at"],
            _resolve_display_name(guild, m["user_id"]),
            m["user_id"],
            m["env_id"],
            FORMAT_LABELS.get(m["format"], m["format"]),
            m["my_class"],
            m.get("my_deck") or "",
            m.get("rank_tier") or "",
            m.get("cr_grade") or "",
            m["opp_class"],
            m.get("opp_deck") or "",
            "先攻" if m["is_first"] else "後攻",
            "勝ち" if m["is_win"] else "負け",
        ])

    # ---- 集計（ユーザー別） ----
    ws_agg = wb.create_sheet("集計")
    ws_agg.append(["ユーザー", "ユーザーID", "試合数", "勝数", "敗数", "勝率(%)"])

    per_user: dict[str, dict] = {}
    for m in matches:
        uid = m["user_id"]
        row = per_user.setdefault(uid, {"total": 0, "wins": 0})
        row["total"] += 1
        row["wins"] += m["is_win"]

    for uid, row in sorted(per_user.items(), key=lambda kv: -kv[1]["total"]):
        total = row["total"]
        wins = row["wins"]
        losses = total - wins
        win_rate = round(wins / total * 100, 1) if total else 0.0
        ws_agg.append([
            _resolve_display_name(guild, uid), uid, total, wins, losses, win_rate,
        ])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================
# 環境レポートの画像（デッキ単位・直感的に読める形）
# ==============================
# ・よく当たるデッキ TOP10（当たる割合・勝ちやすさ・直近の増減）
# ・環境マップ（横：当たる割合／縦：勝ちにくさ）
# ・デッキ相性表（上位8デッキ同士。30戦以上のマスだけ有利/不利を判定）
# いずれもローテーションのみ。使用率は相手側、勝ちやすさは記録者平均との差で見る。

REPORT_TOP_DECKS = 10           # ランキング・環境マップに載せる相手デッキ数
REPORT_MATRIX_DECKS = 8         # 相性表に載せるデッキ数
REPORT_MATRIX_MIN_SHOW = 10     # これ未満の組み合わせは「−」
REPORT_MATRIX_MIN_JUDGE = 30    # これ以上の組み合わせだけ有利/不利を判定
REPORT_FAVOR = 55.0             # 有利とみなす勝率
REPORT_UNFAVOR = 45.0           # 不利とみなす勝率

CLASS_COLORS = {
    "エルフ": "#5DAA4F", "ロイヤル": "#D4A72C", "ウィッチ": "#6F5BD9",
    "ドラゴン": "#D9653B", "ナイトメア": "#A23B6B", "ビショップ": "#9AA3B5",
    "ネメシス": "#3B8FD9",
}


def _parse_recorded(r, default):
    try:
        return datetime.fromisoformat(r["recorded_at"])
    except (TypeError, ValueError, KeyError):
        return default


def compute_report_data(matches: list[dict], now: datetime):
    """環境レポートの画像とAI分析に使う集計（ローテーションのみ）。記録が無ければ None"""
    rows = [m for m in matches if m.get("format") == "rotation"]
    if not rows:
        return None
    n = len(rows)
    avg = sum(1 for r in rows if r["is_win"]) / n * 100
    with_deck = [r for r in rows if r.get("opp_deck")]

    # 直近の期間（試合数が少なければ倍に広げる）
    for days in (ENV_REPORT_AI_RECENT_DAYS, ENV_REPORT_AI_RECENT_DAYS * 2):
        cutoff = now - timedelta(days=days)
        recent = [r for r in with_deck if _parse_recorded(r, now) >= cutoff]
        before = [r for r in with_deck if _parse_recorded(r, now) < cutoff]
        if len(recent) >= 30:
            break

    def key(r):
        return (r["opp_class"], r["opp_deck"])

    total_c, win_c, rec_c, bef_c = {}, {}, {}, {}
    for r in with_deck:
        k = key(r)
        total_c[k] = total_c.get(k, 0) + 1
        win_c[k] = win_c.get(k, 0) + (1 if r["is_win"] else 0)
    for r in recent:
        rec_c[key(r)] = rec_c.get(key(r), 0) + 1
    for r in before:
        bef_c[key(r)] = bef_c.get(key(r), 0) + 1

    items = []
    for k, cnt in sorted(total_c.items(), key=lambda kv: -kv[1])[:REPORT_TOP_DECKS]:
        win_rate = win_c[k] / cnt * 100
        trend = None
        if len(recent) >= 30 and len(before) >= 30:
            trend = rec_c.get(k, 0) / len(recent) * 100 - bef_c.get(k, 0) / len(before) * 100
        items.append({
            "cls": k[0], "deck": k[1], "n": cnt,
            "share": cnt / len(with_deck) * 100,
            "diff": win_rate - avg, "trend": trend, "recent_n": rec_c.get(k, 0),
        })

    # デッキ相性（使った側・当たった側の記録を合算）
    pair_t, pair_w = {}, {}
    for r in rows:
        if not r.get("opp_deck") or not r.get("my_deck"):
            continue
        a = (r["my_class"], r["my_deck"])
        b = (r["opp_class"], r["opp_deck"])
        w = 1 if r["is_win"] else 0
        pair_t[(a, b)] = pair_t.get((a, b), 0) + 1
        pair_w[(a, b)] = pair_w.get((a, b), 0) + w
        pair_t[(b, a)] = pair_t.get((b, a), 0) + 1
        pair_w[(b, a)] = pair_w.get((b, a), 0) + (1 - w)
    top = [(it["cls"], it["deck"]) for it in items[:REPORT_MATRIX_DECKS]]
    matrix = []
    for a in top:
        row = []
        for b in top:
            t = pair_t.get((a, b), 0)
            row.append({"n": t, "rate": (pair_w.get((a, b), 0) / t * 100) if t else None})
        matrix.append(row)

    firsts = {}
    for flag in (1, 0):
        rs = [r for r in rows if r["is_first"] == flag]
        if rs:
            firsts[flag] = (len(rs), sum(1 for r in rs if r["is_win"]) / len(rs) * 100)

    return {
        "n": n, "n_deck": len(with_deck), "avg": avg, "items": items,
        "recent_days": days, "recent_n": len(recent),
        "matrix_decks": top, "matrix": matrix, "firsts": firsts,
    }


def _report_badge(diff: float):
    if diff >= 5:
        return "◎ 勝てている", "#2E8B57"
    if diff >= -2:
        return "○ ふつう", "#7A7A7A"
    if diff >= -6:
        return "△ やや苦戦", "#D08A1E"
    return "× 苦戦", "#C2402F"


def _fig_to_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf


def render_deck_ranking(data: dict, date_label: str):
    """よく当たるデッキ TOP10 の表"""
    items = data["items"]
    if plt is None or _JP_FONT_NAME is None or not items:
        return None
    try:
        height = 2.4 + 0.5 * len(items)
        fig = plt.figure(figsize=(8, height), dpi=150)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis("off")
        ax.set_xlim(0, 100)
        total_h = 14 + 7 * len(items) + 12
        ax.set_ylim(0, total_h)
        top_y = total_h - 4
        ax.text(4, top_y, "今の環境　よく当たるデッキ TOP" + str(len(items)), fontsize=17, weight="bold", va="center")
        ax.text(4, top_y - 4.5, f"ローテーション・相手デッキ{data['n_deck']}戦から集計（{date_label}時点）",
                fontsize=10, color="#666", va="center")
        hdr = top_y - 9
        for x, t in ((4, "順位・デッキ"), (52, "当たる割合"), (77, "勝ちやすさ"), (90, "直近")):
            ax.text(x, hdr, t, fontsize=9.5, color="#888", va="center")
        y = hdr - 6
        mx = max(it["share"] for it in items)
        for i, it in enumerate(items, 1):
            if i % 2:
                ax.add_patch(plt.Rectangle((2, y - 3.2), 96, 6.4, color="#F6F6F4", zorder=0))
            color = CLASS_COLORS.get(it["cls"], "#999999")
            ax.text(4, y, str(i), fontsize=13, weight="bold", va="center")
            ax.add_patch(plt.Rectangle((8.5, y - 2), 1.3, 4, color=color))
            t1 = ax.text(11.5, y, it["deck"], fontsize=12.5, va="center")
            ax.annotate(it["cls"], xycoords=t1, xy=(1, 0.5), xytext=(6, 0), textcoords="offset points",
                        fontsize=9, color="#888", va="center")
            w = it["share"] / mx * 16
            ax.add_patch(plt.Rectangle((52, y - 1.4), w, 2.8, color=color, alpha=0.85))
            ax.text(52 + w + 1, y, f"{it['share']:.1f}%", fontsize=10, va="center")
            label, c = _report_badge(it["diff"])
            ax.text(77, y, label, fontsize=10.5, color=c, va="center", weight="bold")
            if it["trend"] is None or it["recent_n"] < 10:
                a, ac = "ー", "#999999"
            elif it["trend"] >= 3:
                a, ac = "▲", "#C2402F"
            elif it["trend"] <= -3:
                a, ac = "▼", "#3B6FD9"
            else:
                a, ac = "→", "#777777"
            ax.text(92, y, a, fontsize=12, color=ac, va="center", ha="center")
            y -= 7
        ax.text(4, 7, f"勝ちやすさ：記録者全体の平均勝率（{data['avg']:.0f}%）との差　"
                      "◎+5以上／○−2〜+5／△−2〜−6／×−6以下", fontsize=8.5, color="#666")
        ax.text(4, 3, f"直近：{data['recent_days']}日間で当たる割合が ▲増えた／▼減った（3ポイント以上）。"
                      "記録が少ないものは「ー」", fontsize=8.5, color="#666")
        return _fig_to_buf(fig)
    except Exception as e:
        print(f"⚠️ デッキランキング画像の生成に失敗: {type(e).__name__}: {e}")
        return None


def render_env_map(data: dict):
    """環境マップ（横：当たる割合／縦：勝ちにくさ）"""
    items = data["items"]
    if plt is None or _JP_FONT_NAME is None or len(items) < 3:
        return None
    try:
        fig, ax = plt.subplots(figsize=(8, 6.6), dpi=150)
        xs = [it["share"] for it in items]
        ys = [-it["diff"] for it in items]
        ax.axvline(sum(xs) / len(xs), color="#bbbbbb", lw=1, ls="--")
        ax.axhline(0, color="#bbbbbb", lw=1, ls="--")
        max_n = max(it["n"] for it in items)
        for it in items:
            size = 120 + 700 * it["n"] / max_n
            ax.scatter(it["share"], -it["diff"], s=size, color=CLASS_COLORS.get(it["cls"], "#999999"),
                       alpha=0.75, edgecolor="white", lw=1.5, zorder=3)
            ax.annotate(it["deck"], (it["share"], -it["diff"]), xytext=(0, 9 + (size ** 0.5) / 2),
                        textcoords="offset points", ha="center", fontsize=10.5)
        lim = max(abs(v) for v in ys) + 4
        ax.set_ylim(-lim, lim)
        ax.set_xlim(min(xs) - 2, max(xs) + 2.5)
        x_l, x_r = min(xs) - 1.6, max(xs) + 2.2
        ax.text(x_r, lim - 1, "よく当たる＆苦戦\n＝最優先で対策", ha="right", va="top", fontsize=11, color="#C2402F", weight="bold")
        ax.text(x_l, lim - 1, "たまに当たる＆苦戦\n＝見かけたら注意", ha="left", va="top", fontsize=10, color="#D08A1E")
        ax.text(x_r, -lim + 1, "よく当たる＆勝てている", ha="right", va="bottom", fontsize=10, color="#2E8B57")
        ax.text(x_l, -lim + 1, "たまに当たる＆勝てている", ha="left", va="bottom", fontsize=10, color="#888888")
        ax.set_xlabel("当たる割合（%）→ 多い", fontsize=11)
        ax.text(-0.03, 0.98, "勝ちにくい ↑", transform=ax.transAxes, rotation=90, ha="right", va="top", fontsize=11)
        ax.text(-0.03, 0.02, "↓ 勝ちやすい", transform=ax.transAxes, rotation=90, ha="right", va="bottom", fontsize=11)
        ax.set_yticks([])
        ax.set_title("環境マップ：どのデッキを対策すべきか", fontsize=15, pad=12)
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        fig.text(0.5, 0.005, f"縦軸は記録者全体の平均勝率（{data['avg']:.0f}%）との差。円の大きさは試合数",
                 ha="center", fontsize=8.5, color="#666")
        fig.tight_layout(rect=(0, 0.02, 1, 1))
        return _fig_to_buf(fig)
    except Exception as e:
        print(f"⚠️ 環境マップ画像の生成に失敗: {type(e).__name__}: {e}")
        return None


def render_deck_matrix(data: dict):
    """デッキ相性表（左のデッキから見た勝率）"""
    decks = data["matrix_decks"]
    if plt is None or _JP_FONT_NAME is None or len(decks) < 3:
        return None
    try:
        k = len(decks)
        judged = [[float("nan")] * k for _ in range(k)]
        ref = [[float("nan")] * k for _ in range(k)]
        for i in range(k):
            for j in range(k):
                c = data["matrix"][i][j]
                if i == j or c["rate"] is None:
                    continue
                if c["n"] >= REPORT_MATRIX_MIN_JUDGE:
                    judged[i][j] = c["rate"]
                elif c["n"] >= REPORT_MATRIX_MIN_SHOW:
                    ref[i][j] = c["rate"]

        import numpy as _np
        cmap = matplotlib.colormaps["RdYlGn"].copy()
        cmap.set_bad("#EEEEEE")
        fig, ax = plt.subplots(figsize=(1.1 * k + 1.2, 1.0 * k + 1.6), dpi=150)
        ax.imshow(_np.ma.masked_invalid(_np.array(judged)), cmap=cmap, vmin=25, vmax=75)
        ax.imshow(_np.ma.masked_invalid(_np.array(ref)), cmap=cmap, vmin=25, vmax=75, alpha=0.35)
        labels = [f"{d}\n{c}" for c, d in decks]
        ax.set_xticks(range(k))
        ax.set_yticks(range(k))
        ax.set_xticklabels(labels, fontsize=9.5)
        ax.set_yticklabels(labels, fontsize=10)
        ax.xaxis.tick_top()
        for i in range(k):
            for j in range(k):
                c = data["matrix"][i][j]
                if i == j:
                    ax.text(j, i, "同型", ha="center", va="center", fontsize=9, color="#999999")
                elif c["n"] >= REPORT_MATRIX_MIN_JUDGE:
                    v = c["rate"]
                    mark = "有利" if v >= REPORT_FAVOR else "不利" if v <= REPORT_UNFAVOR else "互角"
                    ax.text(j, i, f"{mark}\n{v:.0f}%({c['n']})", ha="center", va="center", fontsize=9.5,
                            weight="bold" if mark != "互角" else "normal")
                elif c["n"] >= REPORT_MATRIX_MIN_SHOW:
                    ax.text(j, i, f"参考\n{c['rate']:.0f}%({c['n']})", ha="center", va="center", fontsize=9, color="#444444")
                else:
                    ax.text(j, i, f"−\n({c['n']})", ha="center", va="center", fontsize=8.5, color="#AAAAAA")
        ax.set_title("デッキ相性表（左のデッキから見た勝率）", fontsize=15, pad=50)
        fig.text(0.5, 0.035,
                 f"よく当たるデッキ上位{k}つ。有利{REPORT_FAVOR:.0f}%以上／不利{REPORT_UNFAVOR:.0f}%以下"
                 f"（{REPORT_MATRIX_MIN_JUDGE}戦以上のみ判定）",
                 ha="center", fontsize=8.5, color="#666")
        fig.text(0.5, 0.012,
                 f"{REPORT_MATRIX_MIN_SHOW}〜{REPORT_MATRIX_MIN_JUDGE - 1}戦は薄い色の「参考」、"
                 f"{REPORT_MATRIX_MIN_SHOW}戦未満は「−」。使った側・当たった側の記録を合算",
                 ha="center", fontsize=8.5, color="#666")
        fig.tight_layout(rect=(0, 0.06, 1, 1))
        return _fig_to_buf(fig)
    except Exception as e:
        print(f"⚠️ デッキ相性表の生成に失敗: {type(e).__name__}: {e}")
        return None


# ==============================
# Cog本体
# ==============================

# ==============================
# 質問のヒント（毎週金曜に自動投稿＋ボタン）／クラス別チャンネルへの質問
# ==============================
# AIは使わない。決まった基準で「聞いてみるとよさそうな相手」を出し、助言として表示するだけ。
HINT_POST_HOUR = 20                 # 自動投稿の時刻（JST）
HINT_POST_WEEKDAY = 4               # 曜日（月=0 … 金=4）
HINT_CHANNEL_KEY = "hint_post_channel_id"
HINT_LAST_DATE_KEY = "hint_post_last_date"
QUESTION_CHANNEL_KEY_PREFIX = "question_channel_"   # + クラス名
HINT_ENV_MIN_GAMES = 10             # 環境全体のヒントに使う最低試合数（30戦未満は「参考」表示）
HINT_ENV_MIN_USERS = 3              # 環境全体のヒントに使う最低記録者数（個人の特定を防ぐ）
HINT_ENV_GAP = -2.0                 # 平均よりこれ以上低いものを「手ごわい」候補にする
HINT_PERSONAL_MIN_TOTAL = 20        # 個人用ヒントを出す最低試合数
HINT_PERSONAL_MIN_GAMES = 5         # 個人用：相手ごとの最低試合数
HINT_PERSONAL_GAP = 15.0            # 個人用：自分の平均よりこれ以上低い相手を苦手候補にする
HINT_FIRST_MIN_GAMES = 10           # 先攻・後攻の差を見る最低試合数（それぞれ）
HINT_FIRST_GAP = 15.0
# 環境レポートの直後に、質問のヒントも一緒に投稿する（既定ON）。
# OFFにしたいときはRailwayの変数 SENSEKI_HINT_WITH_REPORT=0 。OFFのときは毎週金曜の単独投稿に切り替わる
HINT_WITH_REPORT = _parse_bool(os.environ.get("SENSEKI_HINT_WITH_REPORT", "1"))
QUESTION_COOLDOWN_SECONDS = 600     # 質問の連続投稿の間隔
QUESTION_DAILY_LIMIT = 3            # 1人が1日に投稿できる質問数
_QUESTION_LOG: dict[str, list] = {}
CONSULT_KEY = "相談所"   # 「シャドバ相談所」を表すキー（クラス名と区別するため）
# クラス別チャンネルの初期値（/戦績管理 質問先設定 で上書きできる）
CLASS_CHANNEL_DEFAULTS = {
    "エルフ": "1384171023766781982",
    "ロイヤル": "1384171025151037631",
    "ウィッチ": "1384171181090930811",
    "ドラゴン": "1384171197004255325",
    "ナイトメア": "1384171212716113930",
    "ビショップ": "1384171225345036338",
    "ネメシス": "1384171237919428659",
    # デッキ選びの相談（使うデッキが決まっていない人向け）
    CONSULT_KEY: "1401026533434458264",
}


async def class_channel_id(cls: str):
    """クラス別チャンネルのID。管理者の設定があればそれを、なければ初期値を使う"""
    return (await get_state(QUESTION_CHANNEL_KEY_PREFIX + cls)) or CLASS_CHANNEL_DEFAULTS.get(cls)


async def class_channel_map() -> dict:
    out = {c: await class_channel_id(c) for c in CLASS_CHOICES}
    out[CONSULT_KEY] = await class_channel_id(CONSULT_KEY)
    return out


def _consult_note(channels: dict) -> str:
    ch = _chan(channels, CONSULT_KEY)
    return f"🧭 使うデッキが決まっていない・迷っている方は、シャドバ相談所 {ch} へ" if ch else ""


def _chan(channels: dict, cls: str) -> str:
    cid = (channels or {}).get(cls)
    return f"<#{cid}>" if cid else ""


def _rotation_rows(matches: list[dict]) -> list[dict]:
    return [m for m in matches if m.get("format") == "rotation"]


def _env_share_map(rows: list[dict]) -> dict:
    """相手デッキごとの「当たる割合(%)」。相手デッキの入力があった試合だけで計算"""
    with_deck = [r for r in rows if r.get("opp_deck")]
    total = len(with_deck) or 1
    counts: dict = {}
    for r in with_deck:
        k = (r["opp_class"], r["opp_deck"])
        counts[k] = counts.get(k, 0) + 1
    return {k: v / total * 100 for k, v in counts.items()}


def compute_env_hints(matches: list[dict], limit: int = 3):
    """環境全体のヒント：よく当たるのに記録者が勝ちにくい相手。戻り値：(候補のリスト, 記録者全体の平均勝率)"""
    rows = _rotation_rows(matches)
    if not rows:
        return [], 0.0
    avg = sum(1 for r in rows if r["is_win"]) / len(rows) * 100
    shares = _env_share_map(rows)
    groups: dict = {}
    for r in rows:
        if r.get("opp_deck"):
            groups.setdefault((r["opp_class"], r["opp_deck"]), []).append(r)
    out = []
    for (c, d), rs in groups.items():
        n = len(rs)
        users = len({r["user_id"] for r in rs})
        if n < HINT_ENV_MIN_GAMES or users < HINT_ENV_MIN_USERS:
            continue
        rate = sum(1 for r in rs if r["is_win"]) / n * 100
        gap = rate - avg
        if gap > HINT_ENV_GAP:
            continue
        share = shares.get((c, d), 0.0)
        out.append({"class": c, "deck": d, "games": n, "share": share, "rate": rate, "gap": gap,
                    "score": share * (-gap)})
    out.sort(key=lambda x: -x["score"])
    return out[:limit], avg


def compute_personal_hints(matches: list[dict], user_id: str) -> dict:
    """個人用のヒント。自分の記録と、環境全体の「当たる割合」だけを使う（他人の記録は出さない）"""
    all_rows = _rotation_rows(matches)
    mine = [r for r in all_rows if r["user_id"] == user_id]
    res = {"total": len(mine), "enough": len(mine) >= HINT_PERSONAL_MIN_TOTAL,
           "avg": None, "weak": [], "first": None}
    if not mine or not res["enough"]:
        return res
    avg = sum(1 for r in mine if r["is_win"]) / len(mine) * 100
    res["avg"] = avg
    shares = _env_share_map(all_rows)
    groups: dict = {}
    for r in mine:
        if r.get("opp_deck"):
            groups.setdefault((r["opp_class"], r["opp_deck"]), []).append(r)
    weak = []
    for (c, d), rs in groups.items():
        n = len(rs)
        if n < HINT_PERSONAL_MIN_GAMES:
            continue
        rate = sum(1 for r in rs if r["is_win"]) / n * 100
        if rate <= avg - HINT_PERSONAL_GAP:
            weak.append({"class": c, "deck": d, "games": n, "rate": rate, "share": shares.get((c, d), 0.0)})
    weak.sort(key=lambda x: (-x["share"], x["rate"]))
    res["weak"] = weak[:3]
    first = [r for r in mine if r["is_first"] == 1]
    second = [r for r in mine if r["is_first"] == 0]
    if len(first) >= HINT_FIRST_MIN_GAMES and len(second) >= HINT_FIRST_MIN_GAMES:
        fr = sum(1 for r in first if r["is_win"]) / len(first) * 100
        sr = sum(1 for r in second if r["is_win"]) / len(second) * 100
        if abs(fr - sr) >= HINT_FIRST_GAP:
            res["first"] = {"first_rate": fr, "first_n": len(first), "second_rate": sr, "second_n": len(second)}
    return res


def _hint_label(gap: float, games: int) -> str:
    label = "手ごわい相手" if gap <= -6 else "やや手ごわい相手"
    return label + ("（参考）" if games < MIN_SAMPLE_FOR_STATS else "")


def build_env_hint_post(hints: list, avg: float, channels: dict = None) -> str:
    lines = [
        "🎯 **今週の質問のヒント**",
        "",
        f"最近の記録から、「聞いてみるとよさそうな相手」です（記録者全体の平均勝率は約{avg:.0f}%）。",
    ]
    for i, h in enumerate(hints, 1):
        lines.append(
            f"{i}. **{h['deck']}**（{h['class']}）：当たる割合 {h['share']:.0f}%／"
            f"記録者の勝率 {h['rate']:.0f}%（{h['games']}戦）… {_hint_label(h['gap'], h['games'])}"
        )
        ch = _chan(channels, h["class"])
        if ch:
            lines.append(f"　→ 質問は {ch} へ")
    lines += [
        "",
        "質問するときは、**自分のデッキ・ランク帯・どの場面で困ったか**を書くと、答えが付きやすくなります。",
        _consult_note(channels),
        "-# 勝率は「記録者がそのデッキと当たったときの勝率」です。数字は参考で、記録が少ないものは偏りがあります。",
        "",
        "▼ 自分の記録からのヒント／クラスを選んで質問",
    ]
    return "\n".join(lines)


def build_personal_hint_text(res: dict, channels: dict = None, my_class: str = None) -> str:
    all_ch = "　".join(f"{c} {_chan(channels, c)}" for c in CLASS_CHOICES if _chan(channels, c))
    ch_note = f"\n\n📍 質問できるクラス別チャンネル\n{all_ch}" if all_ch else ""
    consult = _consult_note(channels)
    if consult:
        ch_note += f"\n{consult}"
    if res["total"] == 0:
        return ("🎯 現環境の記録がまだありません。\n"
                "#戦歴入力 から記録がたまると、あなた用のヒントが出せるようになります。" + ch_note)
    if not res["enough"]:
        need = HINT_PERSONAL_MIN_TOTAL - res["total"]
        return (f"🎯 現環境の記録は{res['total']}戦です。\n"
                f"あと**{need}戦**記録すると、あなた用のヒントが出せるようになります。\n"
                "-# それまでは、投稿されている「今週の質問のヒント」を参考にしてください。" + ch_note)
    lines = ["🎯 **あなた用の質問のヒント**（現環境・ローテーション）", ""]
    if res["weak"]:
        lines.append("【聞いてみるとよさそうな相手】")
        for i, w in enumerate(res["weak"], 1):
            lines.append(
                f"{i}. **{w['deck']}**（{w['class']}）：あなたの勝率 {w['rate']:.0f}%（{w['games']}戦）"
                f"／環境で当たる割合 {w['share']:.0f}%"
            )
            ch = _chan(channels, w["class"])
            if ch:
                lines.append(f"　→ {w['class']}のチャンネル {ch} で聞いてみるのがおすすめです")
        lines.append("　例：「このデッキと当たったときの、立ち回りで意識することは？」")
    else:
        lines.append("【聞いてみるとよさそうな相手】\n特に苦手がはっきりしている相手は、今のところありません。")
    if res["first"]:
        f = res["first"]
        lines += ["", "【先攻・後攻の差】",
                  f"先攻 {f['first_rate']:.0f}%（{f['first_n']}戦）／後攻 {f['second_rate']:.0f}%（{f['second_n']}戦）"]
        low = "後攻" if f["second_rate"] < f["first_rate"] else "先攻"
        lines.append(f"　→ {low}のときの立ち回りを聞いてみてもよさそうです")
        mine_ch = _chan(channels, my_class) if my_class else ""
        if mine_ch:
            lines.append(f"　　自分のクラスのチャンネル {mine_ch} がおすすめです")
    consult = _consult_note(channels)
    if consult:
        lines += ["", consult]
    lines += ["", "-# あくまで目安です。試合数が少ないうちは偏りが出ます。質問の書き方は「今週の質問のヒント」を参考にしてください。"]
    return "\n".join(lines)


async def submit_question(client, user, cls: str, title: str, body: str, attach: bool):
    """
    クラス別のチャンネル（またはシャドバ相談所）に、質問を投稿してスレッドを作る。
    DiscordのボタンからもWebページからも共通で使う。戻り値：(成功したか, メッセージ, 投稿のURL)
    user は discord.User / discord.Member。
    """
    import time as _time
    user_id = str(user.id)
    now = _time.time()
    log = [t for t in _QUESTION_LOG.get(user_id, []) if now - t < 86400]
    if log and now - log[-1] < QUESTION_COOLDOWN_SECONDS:
        wait = int(QUESTION_COOLDOWN_SECONDS - (now - log[-1]))
        return False, f"連続の投稿は少し間を空けてください（あと約{wait}秒）。", None
    if len(log) >= QUESTION_DAILY_LIMIT:
        return False, f"質問は1日{QUESTION_DAILY_LIMIT}件までです。また明日お願いします。", None
    title = (title or "").strip()
    body = (body or "").strip()
    low = (title + " " + body).lower()
    if "discord.gg/" in low or "discord.com/invite" in low:
        return False, "招待リンクは投稿できません。", None
    if not title or not body:
        return False, "題名と内容を入力してください。", None

    channel_id = await class_channel_id(cls)
    if not channel_id:
        return False, f"「{cls}」の質問先がまだ設定されていません。運営にお知らせください。", None
    channel = client.get_channel(int(channel_id))
    if channel is None:
        try:
            channel = await client.fetch_channel(int(channel_id))
        except discord.HTTPException:
            return False, "質問先のチャンネルが見つかりません。運営にお知らせください。", None
    if isinstance(user, discord.Member):
        perms = channel.permissions_for(user)
        if not (perms.view_channel and perms.send_messages):
            return False, f"{channel.mention} に投稿できる権限がありません。", None

    embed = None
    if attach:
        settings = await get_user_settings(user_id)
        matches = await fetch_all_matches(env_only=True)
        embed = build_question_embed(user_id, cls, _rotation_rows(matches), settings)

    safe_title = discord.utils.escape_mentions(title)[:80]
    safe_body = discord.utils.escape_mentions(body)[:800]
    content = (f"❓ **{safe_title}**\n{safe_body}\n"
               f"-# 質問者：{user.mention}（戦績ツールから投稿）")
    allowed = discord.AllowedMentions(users=[user], roles=False, everyone=False)
    kwargs = {"allowed_mentions": allowed, "view": QuestionVCView()}
    if embed is not None:
        kwargs["embed"] = embed
    try:
        if isinstance(channel, discord.ForumChannel):
            # フォーラムチャンネルの場合は、投稿（スレッド）として作る
            created = await channel.create_thread(name=safe_title[:90], content=content, **kwargs)
            jump_url = created.thread.jump_url if hasattr(created, "thread") else created.jump_url
            post_id = getattr(getattr(created, "message", None), "id", None) or getattr(getattr(created, "thread", None), "id", None)
        else:
            message = await channel.send(content, **kwargs)
            jump_url = message.jump_url
            post_id = message.id
            try:
                await message.create_thread(name=safe_title[:90], auto_archive_duration=1440)
            except discord.HTTPException:
                pass  # スレッドを作れなくても、質問自体は投稿済み
    except discord.HTTPException as e:
        return False, f"投稿に失敗しました（{type(e).__name__}）。運営にお知らせください。", None
    if post_id:
        await set_state(f"{QUESTION_ASKER_KEY_PREFIX}{post_id}", user_id)   # VC予約ボタンで質問者を確認するため
    log.append(now)
    _QUESTION_LOG[user_id] = log
    return True, f"{channel.mention} に質問を投稿しました。", jump_url


# ---- 質問の投稿に付けるデータ（相手クラス別の勝率）とVC予約 ----
QUESTION_ASKER_KEY_PREFIX = "question_asker_"     # + 投稿のID → 質問者のユーザーID
QUESTION_REF_GAMES = 10          # これ未満の試合数は「参考値」と表示する
QUESTION_OTHERS_MIN = 5          # 「他の記録者」の平均を出す最低試合数
QUESTION_GAP_MARK = 6.0          # 他の記録者との差がこれ以上なら🔴/🟢
VC_SLOT_KEYS = ["vc1", "vc2", "vc3"]              # 質問先設定 の「相談VC①〜③」
VC_NOTICE_KEY = "vcnotice"                        # 質問先設定 の「VC予約のお知らせ先」
VC_NOTICE_DEFAULT = "1467189822350295213"         # VC予約を知らせるチャンネルの初期値（予定を見る場所）
VC_RESERVE_DAILY_LIMIT = 2       # 1人が1日に予約できる回数
VC_RESERVE_MINUTES = 120         # 1回の予約の長さ（1人2時間確保）
VC_DEFAULT_IDS = ["1208804697801236564", "1407176497084629122"]   # 全クラス共通の相談VC（シャドバビヨンドVC／VC2）。先着順
VC_AUTO = "auto"                 # 「空いている方におまかせ」
VC_RESERVE_DAYS = 4              # 今日から何日先まで予約できるか
VC_START_HOURS = list(range(12, 24))               # 予約できる開始の「時」（12時〜23時）
VC_START_MINUTES = [0, 15, 30, 45]                 # 予約できる開始の「分」（15分刻み）
_VC_RESERVE_LOG: dict[str, list] = {}


def _win_rate(rows: list):
    return (sum(1 for r in rows if r["is_win"]) / len(rows) * 100) if rows else None


def build_question_embed(user_id: str, cls: str, rows: list, settings: dict = None):
    """質問者の戦績を、他の記録者の平均と並べて見せる枠。読む人が状況を判断するための材料"""
    mine = [r for r in rows if r["user_id"] == user_id]
    others = [r for r in rows if r["user_id"] != user_id]
    lines = []
    if settings:
        deck = settings.get("my_deck") or "デッキ未設定"
        rank = settings.get("rank_tier") or "ランク未設定"
        if settings.get("cr_grade"):
            rank += f"・{settings['cr_grade']}"
        lines.append(f"**使用** {deck}（{settings.get('my_class', '')}）／{rank}")
    if not mine:
        lines.append("現環境の記録はまだありません。")
    else:
        total_wr = _win_rate(mine)
        first = [r for r in mine if r.get("is_first")]
        second = [r for r in mine if not r.get("is_first")]
        part = f"**全体** {len(mine)}戦 {total_wr:.0f}%"
        if first:
            part += f"　先攻 {_win_rate(first):.0f}%（{len(first)}戦）"
        if second:
            part += f"　後攻 {_win_rate(second):.0f}%（{len(second)}戦）"
        lines.append(part)
        lines.append("")
        lines.append("**相手クラス別の勝率**（自分 ／ 他の記録者の平均 ／ 差）")
        by_cls: dict = {}
        for r in mine:
            by_cls.setdefault(r["opp_class"], []).append(r)
        order = sorted(by_cls, key=lambda c: (c != cls, -len(by_cls[c])))[:7]
        for c in order:
            my_rows = by_cls[c]
            my_wr = _win_rate(my_rows)
            oth_rows = [r for r in others if r["opp_class"] == c]
            oth_wr = _win_rate(oth_rows) if len(oth_rows) >= QUESTION_OTHERS_MIN else None
            mark = "⚪"
            diff_txt = "他の記録者のデータが少ないため比較なし"
            if oth_wr is not None:
                diff = my_wr - oth_wr
                mark = "🔴" if diff <= -QUESTION_GAP_MARK else ("🟢" if diff >= QUESTION_GAP_MARK else "⚪")
                diff_txt = f"他 {oth_wr:.0f}%（{len(oth_rows)}戦）→ **{diff:+.0f}pt**"
            ref = " ⚠参考値" if len(my_rows) < QUESTION_REF_GAMES else ""
            target = " ◀今回の相手" if c == cls else ""
            lines.append(f"{mark} {c}{target}　**{my_wr:.0f}%**（{len(my_rows)}戦）{ref}　{diff_txt}")
        recent = sorted(mine, key=lambda r: r.get("recorded_at") or "")[-10:]
        lines.append("")
        lines.append("**直近の試合**　" + "".join("🟢" if r["is_win"] else "🔴" for r in recent))
    embed = discord.Embed(title="📎 質問者のデータ（現環境・ローテーション）",
                          description="\n".join(lines)[:4000], colour=discord.Colour.blurple())
    embed.set_footer(text=f"{QUESTION_REF_GAMES}戦未満は参考値です。「他の記録者」は質問者以外の記録者の平均です。")
    return embed


async def consult_vc_ids() -> list:
    """相談に使うVC。管理者が設定していればそれを、なければ既定の2つ（全クラス共通）を使う"""
    out = []
    for k in VC_SLOT_KEYS:
        v = await get_state(QUESTION_CHANNEL_KEY_PREFIX + k)
        if v:
            out.append(str(v))
    return out or list(VC_DEFAULT_IDS)


_WEEKDAY_JA = "月火水木金土日"


class VCReserveView(discord.ui.View):
    """質問者が、日・時刻・VCを選んで相談の予約を入れる画面（押した人にだけ表示）"""

    def __init__(self, asker_id: str, vc_options: list, question_message, title: str):
        super().__init__(timeout=300)
        self.asker_id = asker_id
        self.question_message = question_message
        self.title = title
        self.day = None
        self.time = None
        self.vc = None
        today = datetime.now(JST).date()
        day_opts = []
        for i in range(VC_RESERVE_DAYS):
            d = today + timedelta(days=i)
            tag = "今日" if i == 0 else ("明日" if i == 1 else "")
            day_opts.append(discord.SelectOption(
                label=f"{d.month}/{d.day}（{_WEEKDAY_JA[d.weekday()]}）{tag}".strip(), value=d.isoformat()))
        self.hour = None
        self.minute = None
        hour_opts = [discord.SelectOption(label=f"{h}時", value=f"{h:02d}") for h in VC_START_HOURS]
        minute_opts = [discord.SelectOption(label=f"{m:02d}分", value=f"{m:02d}") for m in VC_START_MINUTES]
        self.day_sel = discord.ui.Select(placeholder="① 日にちを選ぶ", options=day_opts, row=0)
        self.hour_sel = discord.ui.Select(placeholder="② 開始の「時」を選ぶ", options=hour_opts, row=1)
        self.minute_sel = discord.ui.Select(placeholder="③ 開始の「分」を選ぶ（15分刻み）", options=minute_opts, row=2)
        vc_opts = [discord.SelectOption(label=name[:100], value=cid) for cid, name in vc_options]
        if len(vc_opts) > 1:
            vc_opts.insert(0, discord.SelectOption(label="空いている方（おまかせ）", value=VC_AUTO,
                                                   description="先に空いているVCを取ります"))
        self.vc_sel = discord.ui.Select(placeholder="④ VCを選ぶ", row=3, options=vc_opts)
        self.day_sel.callback = self._on_day
        self.hour_sel.callback = self._on_hour
        self.minute_sel.callback = self._on_minute
        self.vc_sel.callback = self._on_vc
        for item in (self.day_sel, self.hour_sel, self.minute_sel, self.vc_sel):
            self.add_item(item)
        self.go = discord.ui.Button(label="この日時で予約する", emoji="🎤", style=discord.ButtonStyle.success, row=4)
        self.go.callback = self._on_go
        self.add_item(self.go)

    async def _on_day(self, interaction):
        self.day = self.day_sel.values[0]
        await interaction.response.defer()

    def _update_time(self):
        if self.hour and self.minute:
            self.time = f"{self.hour}:{self.minute}"

    async def _on_hour(self, interaction):
        self.hour = self.hour_sel.values[0]
        self._update_time()
        await interaction.response.defer()

    async def _on_minute(self, interaction):
        self.minute = self.minute_sel.values[0]
        self._update_time()
        await interaction.response.defer()

    async def _on_vc(self, interaction):
        self.vc = self.vc_sel.values[0]
        await interaction.response.defer()

    async def _on_go(self, interaction):
        if not (self.day and self.time and self.vc):
            await interaction.response.send_message("日にち・時・分・VCの4つを選んでください。", ephemeral=True)
            return
        await interaction.response.send_modal(VCCommentModal(self))


class VCCommentModal(discord.ui.Modal):
    """予約の確定前に、任意のひとことを入れる画面。空のままでも予約できる"""

    def __init__(self, parent: "VCReserveView"):
        super().__init__(title="VCの予約")
        self.parent_view = parent
        self.theme = None
        if parent.question_message is None:      # 質問に紐づかない、直接のVC予約
            self.theme = discord.ui.TextInput(
                label="テーマ（何をしますか）", required=True, max_length=60,
                placeholder="例：ロイヤルの対策を相談したい／ランクマの練習")
            self.add_item(self.theme)
        self.comment = discord.ui.TextInput(
            label="ひとこと（任意）", required=False, max_length=100,
            placeholder="例：音声だけでもOKです／画面共有で見てもらいたいです",
        )
        self.add_item(self.comment)

    async def on_submit(self, interaction: discord.Interaction):
        p = self.parent_view
        title = p.title
        if self.theme is not None:
            title = discord.utils.escape_mentions(str(self.theme.value or "").strip()) or "VC"
        ok, text = await create_vc_reservation(interaction, p.asker_id, p.day, p.time, p.vc,
                                               p.question_message, title, str(self.comment.value or ""))
        if ok:
            for item in p.children:
                item.disabled = True
            await interaction.response.edit_message(content=text, view=p)
        else:
            await interaction.response.send_message(text, ephemeral=True)


async def create_vc_reservation(interaction, asker_id: str, day: str, hhmm: str, vc_id: str,
                                question_message, title: str, comment: str = ""):
    """Discordの「イベント」としてVCの予約を作り、質問のスレッドに知らせる。戻り値：(成功したか, 表示する文)"""
    import time as _time
    guild = interaction.guild
    if guild is None:
        return False, "サーバー内で使ってください。"
    now_ts = _time.time()
    log = [t for t in _VC_RESERVE_LOG.get(asker_id, []) if now_ts - t < 86400]
    if len(log) >= VC_RESERVE_DAILY_LIMIT:
        return False, f"VCの予約は、1日{VC_RESERVE_DAILY_LIMIT}回までです。"
    h, m = (int(x) for x in hhmm.split(":"))
    d = datetime.fromisoformat(day).date()
    start = datetime(d.year, d.month, d.day, h, m, tzinfo=JST)
    end = start + timedelta(minutes=VC_RESERVE_MINUTES)
    if start < datetime.now(JST) + timedelta(minutes=5):
        return False, "5分以内、または過ぎた時刻は予約できません。もう少し先の時刻を選んでください。"
    try:
        events = await guild.fetch_scheduled_events()
    except discord.HTTPException:
        events = []

    def busy_at(channel_id: int):
        """その時間に、このVCで重なる予約があれば、その開始時刻を返す"""
        for ev in events:
            if getattr(ev, "channel_id", None) != channel_id:
                continue
            ev_end = ev.end_time or (ev.start_time + timedelta(minutes=VC_RESERVE_MINUTES))
            if ev.start_time < end and ev_end > start:
                return ev.start_time.astimezone(JST)
        return None

    candidates = []
    for cid in (await consult_vc_ids() if vc_id == VC_AUTO else [vc_id]):
        ch_ = guild.get_channel(int(cid))
        if ch_ is not None:
            candidates.append(ch_)
    if not candidates:
        return False, "選んだVCが見つかりません。運営にお知らせください。"
    vc = next((c for c in candidates if busy_at(c.id) is None), None)
    if vc is None:
        if vc_id == VC_AUTO:
            return False, "その時間は、どのVCも予約で埋まっています。別の時刻を選んでください。"
        # 選んだVCが埋まっているとき、もう片方が空いていれば教える
        others = [g for g in [guild.get_channel(int(c)) for c in await consult_vc_ids()]
                  if g is not None and g.id != candidates[0].id and busy_at(g.id) is None]
        hint = f"{others[0].name} は空いています。" if others else "どちらのVCも埋まっています。"
        return False, f"{candidates[0].name} は、その時間は予約で埋まっています。{hint}別の時刻かVCを選んでください。"
    day_txt = f"{d.month}/{d.day}（{_WEEKDAY_JA[d.weekday()]}） {h}:{m:02d}〜"
    comment = discord.utils.escape_mentions((comment or "").strip())[:100]
    comment_line = f"\nひとこと：{comment}" if comment else ""
    event = None
    try:
        event = await guild.create_scheduled_event(
            name=(f"🎤 質問の相談：{title}" if question_message is not None else f"🎤 VC：{title}")[:100],
            start_time=start, end_time=end, channel=vc,
            description=((f"質問者：{interaction.user.display_name}\nスレッドの質問について、VCで相談します。参加したい方は「興味あり」を押してください。"
                          if question_message is not None else
                          f"予約者：{interaction.user.display_name}\n参加したい方は「興味あり」を押してください。")
                         + (f"\nひとこと：{comment}" if comment else ""))[:1000],
            entity_type=discord.EntityType.voice, privacy_level=discord.PrivacyLevel.guild_only,
        )
    except discord.Forbidden:
        event = None   # イベントを作る権限がないときは、スレッドへのお知らせだけにする
    except discord.HTTPException as e:
        return False, f"予約を作れませんでした（{type(e).__name__}）。運営にお知らせください。"
    msg = (f"🎤 **VC相談の予約**\n日時：{day_txt}（JST）\nVC：{vc.mention}\n質問者：<@{asker_id}>{comment_line}\n"
           "参加できる方は、イベントの「興味あり」を押すか、このスレッドに一言ください。")
    if event is not None:
        msg += f"\n{event.url}"
    target = None
    if question_message is not None:     # 質問への予約のときだけ、スレッドにも知らせる
        target = interaction.channel if isinstance(interaction.channel, discord.Thread) else getattr(question_message, "thread", None)
        allowed = discord.AllowedMentions(users=[discord.Object(id=int(asker_id))], roles=False, everyone=False)
        try:
            if target is not None:
                await target.send(msg, allowed_mentions=allowed)
            else:
                await question_message.reply(msg, allowed_mentions=allowed)
        except discord.HTTPException:
            pass
    # 予定を見る場所（イベント機能は知らない人が多いため）にも同じ予定を出す
    thread_url = getattr(target, "jump_url", None) or getattr(question_message, "jump_url", None)
    notice_id = (await get_state(QUESTION_CHANNEL_KEY_PREFIX + VC_NOTICE_KEY)) or VC_NOTICE_DEFAULT
    try:
        notice_ch = guild.get_channel(int(notice_id)) or await interaction.client.fetch_channel(int(notice_id))
        notice = (f"🗓️ **VC相談の予定**\n日時：{day_txt}（JST）\nVC：{vc.mention}\n"
                  f"テーマ：{title}\n質問者：<@{asker_id}>{comment_line}") if question_message is not None else (
                  f"🗓️ **VCの予定**\n日時：{day_txt}（JST）\nVC：{vc.mention}\n"
                  f"テーマ：{title}\n予約者：<@{asker_id}>{comment_line}")
        await notice_ch.send(notice, allowed_mentions=discord.AllowedMentions.none())
    except Exception as e:
        print(f"⚠️ VC予約のお知らせを予定チャンネルに出せませんでした: {type(e).__name__}: {e}")
    log.append(now_ts)
    _VC_RESERVE_LOG[asker_id] = log
    note = "" if event is not None else "\n※イベントを作る権限がないため、お知らせだけにしました。"
    return True, f"✅ 予約しました：{day_txt}　{vc.mention}{note}"


class QuestionVCView(discord.ui.View):
    """質問の投稿に付く「VC予約」ボタン（永続ビュー）"""

    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="VCで相談する（予約）", emoji="🎤", style=discord.ButtonStyle.success,
                       custom_id="senseki_question_vc_v1")
    async def reserve(self, interaction: discord.Interaction, button: discord.ui.Button):
        msg = interaction.message
        asker = await get_state(f"{QUESTION_ASKER_KEY_PREFIX}{msg.id}")
        if not asker:
            await interaction.response.send_message("この質問の投稿者の情報が見つかりません。", ephemeral=True)
            return
        if str(interaction.user.id) != str(asker):
            await interaction.response.send_message(
                "予約できるのは質問した本人です。相談に参加したい場合は、本人が予約すると出るイベントの「興味あり」を押すか、"
                "スレッドに書き込んでください。", ephemeral=True)
            return
        vc_ids = await consult_vc_ids()
        options = []
        for cid in vc_ids:
            ch = interaction.guild.get_channel(int(cid)) if interaction.guild else None
            if ch is not None:
                options.append((cid, ch.name))
        if not options:
            await interaction.response.send_message("相談用のVCがまだ設定されていません。運営にお知らせください。",
                                                    ephemeral=True)
            return
        first_line = (msg.content or "").split("\n")[0].replace("❓", "").replace("*", "").strip()
        view = VCReserveView(str(asker), options, msg, first_line or "質問")
        await interaction.response.send_message(
            "🎤 **VCで相談する日時を選んでください**（今日から4日先まで・1人2時間確保・先着順）\n"
            "予約すると、Discordの「イベント」が作られ、スレッドにお知らせが出ます。", view=view, ephemeral=True)



async def start_direct_vc(interaction: discord.Interaction):
    """質問に紐づかないVC予約の入口（パネル・ヒントのボタンから）。押した人にだけ画面を出す"""
    vc_ids = await consult_vc_ids()
    options = []
    for cid in vc_ids:
        ch = interaction.guild.get_channel(int(cid)) if interaction.guild else None
        if ch is not None:
            options.append((cid, ch.name))
    if not options:
        await interaction.response.send_message("相談用のVCが見つかりません。運営にお知らせください。", ephemeral=True)
        return
    view = VCReserveView(str(interaction.user.id), options, None, "")
    await interaction.response.send_message(
        "🎤 **VCを予約します**（今日から4日先まで・1人2時間確保・先着順）\n"
        "日時とVCを選ぶと、Discordの「イベント」になり、予定のチャンネルにもお知らせが出ます。", view=view, ephemeral=True)


async def post_question(interaction: discord.Interaction, cls: str, title: str, body: str, attach: bool):
    """Discordの入力画面（モーダル）から呼ばれる。結果は押した人にだけ表示する"""
    ok, message, url = await submit_question(interaction.client, interaction.user, cls, title, body, attach)
    text = ("✅ " + message + (f"\n{url}" if url else "")) if ok else ("⏳ " if "間を空けて" in message else "") + message
    await interaction.response.send_message(text, ephemeral=True)


class QuestionModal(discord.ui.Modal):
    def __init__(self, cls: str, attach: bool):
        super().__init__(title=("デッキ選びの相談" if cls == CONSULT_KEY else f"{cls}への質問")[:45])
        self.cls = cls
        self.attach = attach
        consult = cls == CONSULT_KEY
        self.q_title = discord.ui.TextInput(
            label="題名", max_length=80,
            placeholder="例：ロイヤルで使うデッキを探しています" if consult else "例：秘術ウィッチへの対策を教えてください",
        )
        self.q_body = discord.ui.TextInput(
            label="内容（使いたいクラス・ランク帯・好み）" if consult else "内容（自分のデッキ・ランク帯・困った場面）",
            style=discord.TextStyle.paragraph, max_length=800,
            placeholder=("例：マスター帯です。ロイヤルで、扱いやすいデッキを知りたいです。" if consult
                         else "例：GMでロイヤルを使っています。後攻のとき、…で困っています。"),
        )
        self.add_item(self.q_title)
        self.add_item(self.q_body)

    async def on_submit(self, interaction: discord.Interaction):
        await post_question(interaction, self.cls, str(self.q_title.value), str(self.q_body.value), self.attach)


class QuestionWriteButton(discord.ui.Button):
    def __init__(self, cls: str, attach: bool):
        super().__init__(
            label="戦績を添付して書く" if attach else "質問を書く",
            emoji="📎" if attach else "✍",
            style=discord.ButtonStyle.primary if attach else discord.ButtonStyle.secondary,
        )
        self.cls = cls
        self.attach = attach

    async def callback(self, interaction: discord.Interaction):
        await interaction.response.send_modal(QuestionModal(self.cls, self.attach))


class QuestionClassSelect(discord.ui.Select):
    def __init__(self):
        super().__init__(
            placeholder="質問先を選択（クラス別／デッキ選びの相談）",
            options=[discord.SelectOption(label=c, value=c) for c in CLASS_CHOICES]
                    + [discord.SelectOption(label="🧭 使うデッキの相談（シャドバ相談所）", value=CONSULT_KEY)],
        )

    async def callback(self, interaction: discord.Interaction):
        cls = self.values[0]
        target = "シャドバ相談所" if cls == CONSULT_KEY else f"**{cls}**のチャンネル"
        if not await class_channel_id(cls):
            await interaction.response.edit_message(
                content=f"「{cls}」の質問先がまだ設定されていません。運営にお知らせください。", view=None)
            return
        view = discord.ui.View(timeout=300)
        view.add_item(QuestionWriteButton(cls, attach=True))
        view.add_item(QuestionWriteButton(cls, attach=False))
        await interaction.response.edit_message(
            content=(f"{target}に投稿します。\n"
                     "📎 戦績を添付すると、使用デッキ・ランク・相手クラス別の勝率（他の記録者の平均との差つき）が質問に添えられます。答えが付きやすくなります。"),
            view=view,
        )


PANEL_HINT_CUSTOM_ID = "senseki_panel_hint_v1"
PANEL_ASK_CUSTOM_ID = "senseki_panel_ask_v1"
PANEL_VC_CUSTOM_ID = "senseki_panel_vc_v1"


async def show_personal_hint(interaction: discord.Interaction):
    """自分用の質問のヒントを、押した人にだけ表示する"""
    matches = await fetch_all_matches(env_only=True)
    res = compute_personal_hints(matches, str(interaction.user.id))
    settings = await get_user_settings(str(interaction.user.id))
    my_class = settings.get("my_class") if settings else None
    await interaction.response.send_message(
        build_personal_hint_text(res, await class_channel_map(), my_class), ephemeral=True)


async def start_question_flow(interaction: discord.Interaction):
    """質問先（クラス別チャンネル／シャドバ相談所）の選択画面を、押した人にだけ表示する"""
    view = discord.ui.View(timeout=300)
    view.add_item(QuestionClassSelect())
    await interaction.response.send_message("質問先を選んでください。", view=view, ephemeral=True)


class PanelHintButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="質問のヒント", style=discord.ButtonStyle.secondary, emoji="🎯",
                         custom_id=PANEL_HINT_CUSTOM_ID)

    async def callback(self, interaction: discord.Interaction):
        await show_personal_hint(interaction)


class PanelAskButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="質問・相談する", style=discord.ButtonStyle.secondary, emoji="❓",
                         custom_id=PANEL_ASK_CUSTOM_ID)

    async def callback(self, interaction: discord.Interaction):
        await start_question_flow(interaction)


class PanelVCButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="VC予約", style=discord.ButtonStyle.success, emoji="🎤",
                         custom_id=PANEL_VC_CUSTOM_ID)

    async def callback(self, interaction: discord.Interaction):
        await start_direct_vc(interaction)


class HintPostView(discord.ui.View):
    """毎週金曜の「質問のヒント」投稿に付くボタン（永続ビュー）"""

    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="自分用のヒントを見る", emoji="🎯", style=discord.ButtonStyle.primary,
                       custom_id="senseki_hint_personal")
    async def personal(self, interaction: discord.Interaction, button: discord.ui.Button):
        await show_personal_hint(interaction)

    @discord.ui.button(label="クラスを選んで質問する", emoji="❓", style=discord.ButtonStyle.secondary,
                       custom_id="senseki_hint_ask")
    async def ask(self, interaction: discord.Interaction, button: discord.ui.Button):
        await start_question_flow(interaction)

    @discord.ui.button(label="VCを予約する", emoji="🎤", style=discord.ButtonStyle.success,
                       custom_id="senseki_hint_vc")
    async def reserve_vc(self, interaction: discord.Interaction, button: discord.ui.Button):
        await start_direct_vc(interaction)


class SensekiCog(commands.Cog):
    # コマンドはグループにまとめている。Discord上は
    #   /戦績 記録 ／ /戦績 設定 …
    # のように表示され、トップレベルのコマンド数は3つで済む。
    # 一般利用はすべてパネルのボタンで完結するため、コマンドは管理者専用にして
    # 一般メンバーのコマンド一覧に出さない（パネルが使えない時の予備手段として残す）
    senseki = app_commands.Group(
        name="戦績",
        description="戦績の記録・確認・設定（管理者専用。通常はパネルを使用）",
        default_permissions=discord.Permissions(administrator=True),
    )
    deck = app_commands.Group(
        name="デッキ",
        description="使用デッキの登録・切替（管理者専用。通常はパネルを使用）",
        default_permissions=discord.Permissions(administrator=True),
    )
    admin = app_commands.Group(
        name="戦績管理",
        description="戦績ツールの管理機能（管理者専用）",
        default_permissions=discord.Permissions(administrator=True),
    )

    def __init__(self, bot: commands.Bot):
        self.bot = bot

    async def cog_load(self):
        await init_db()

        # データが消える事故を早期に発見するため、起動のたびに保存先を明示する
        print("---- 戦績ツール ストレージ診断 ----")
        print(describe_storage())
        try:
            counts = await get_record_counts()
            print(f"記録件数: {counts['matches']}件 / 設定済みユーザー: {counts['users']}名")
        except Exception as e:
            print(f"件数取得に失敗: {type(e).__name__}: {e}")
        if not _is_persistent(os.path.dirname(DB_PATH)):
            print("⚠️ 保存先が永続ボリュームではありません。再デプロイでデータが消えます。")
            print("⚠️ Railway の Volumes でマウント先を確認し、必要なら")
            print("⚠️ 環境変数 SENSEKI_DB_PATH に正しいパスを設定してください。")
        print("-----------------------------------")
        if not self.weekly_export.is_running():
            self.weekly_export.start()
        if not self.env_watch.is_running():
            self.env_watch.start()
        if not self.env_report_loop.is_running():
            self.env_report_loop.start()
        if not self.daily_sheets_sync.is_running():
            self.daily_sheets_sync.start()
        if not self.debounced_sheets_sync.is_running():
            self.debounced_sheets_sync.start()
        if not self.hint_post_loop.is_running():
            self.hint_post_loop.start()
        # 永続ビューの登録。custom_idが一致すれば、BOT再起動後も
        # 過去に送信済みのパネルのボタンが反応し続ける
        self.bot.add_view(SensekiPanelView())
        self.bot.add_view(LegacyPanelView())
        self.bot.add_view(HintPostView())
        self.bot.add_view(QuestionVCView())
        # Webページ（senseki_web.py があり WEB_SESSION_SECRET が設定されているときだけ起動）
        try:
            from senseki_web import start_web
            if getattr(self, "_web_runner", None) is None:
                self._web_runner = await start_web(self.bot)
        except ImportError:
            pass
        except Exception as e:
            print(f"⚠️ Webページの起動に失敗しました（BOT本体は通常どおり動きます）: {type(e).__name__}: {e}")
        # BOT再起動をまたいでも掲示板が最新になるよう一度更新しておく
        try:
            await self._update_board()
        except Exception as e:
            print(f"⚠️ 起動時の戦績掲示板更新に失敗しました: {type(e).__name__}: {e}")

    def cog_unload(self):
        self.weekly_export.cancel()
        self.env_watch.cancel()
        self.env_report_loop.cancel()
        self.daily_sheets_sync.cancel()
        self.debounced_sheets_sync.cancel()
        self.hint_post_loop.cancel()
        runner = getattr(self, "_web_runner", None)
        if runner is not None:
            self._web_runner = None
            asyncio.create_task(runner.cleanup())

    @staticmethod
    def _missing_perm_message(interaction: discord.Interaction, what: str) -> str:
        """チャンネル権限が足りないときに、何をどう直せばよいかを具体的に返す。"""
        ch = interaction.channel
        me = interaction.guild.me if interaction.guild else None
        lines = [
            f"❌ このチャンネルに{what}を設置できませんでした。",
            "BOTに **チャンネルの権限が不足** しています。",
            "",
            f"**対象チャンネル**：{ch.mention if ch else '不明'}",
        ]
        if me is not None and ch is not None:
            try:
                p = ch.permissions_for(me)
                checks = [
                    ("チャンネルを見る", p.view_channel),
                    ("メッセージを送信", p.send_messages),
                    ("メッセージの管理（ピン留め用）", p.manage_messages),
                    ("埋め込みリンク", p.embed_links),
                ]
                lines.append("")
                lines.append("**現在の権限**")
                for name, ok in checks:
                    lines.append(f"{'✅' if ok else '❌'} {name}")
            except Exception:
                pass
        lines += [
            "",
            "**直し方**",
            "チャンネルの編集 → 権限 → BOT（またはBOTのロール）を追加し、",
            "上の❌が付いた項目を許可にしてください。",
        ]
        return "\n".join(lines)

    @admin.command(name="パネル設置", description="このチャンネルに戦績記録パネルを常設します")
    async def senseki_panel_setup(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            message = await interaction.channel.send(
                content=(
                    "⚔️ **戦績記録パネル**\n"
                    "ボタンはすべて押した人にだけ画面が表示されます。\n"
                    "**はじめての方は「🆕 初回設定」から**\n"
                    "\n"
                    "**記録・確認**\n"
                    "・⚔️ 戦績を記録する　・↩️ 直前の記録を取消\n"
                    "・📋 記録の確認・修正（過去の記録を見て、修正・削除。削除しても7日以内なら戻せます）\n"
                    "・📊 成績・弱点（通算・直近の成績と、勝率が低い対面）\n"
                    "・🤖 AIで分析（今のデッキの記録をAIが読み解く。1日1回）\n"
                    "\n"
                    "**設定**\n"
                    "・🆕 初回設定\n"
                    "・⚙️ 設定を変える（デッキ切替・ランク更新・フォーマット切替・デッキ検証）\n"
                    "\n"
                    "**質問・相談**\n"
                    "・🎯 質問のヒント（あなたの記録から、聞いてみるとよさそうな相手とおすすめのチャンネルを表示）\n"
                    "・❓ 質問・相談する（クラス別チャンネル／使うデッキの相談はシャドバ相談所へ投稿。戦績の添付もできます）\n"
                    "・🎤 VC予約（シャドバビヨンドVC／VC2を、先着順・1人2時間で予約。予定のチャンネルにもお知らせが出ます）\n"
                    "\n"
                    "**⚠️ 記録の前に確認してください**\n"
                    "記録画面の上部に、今のあなたのデッキとランクが表示されます。\n"
                    "**違っていたら記録画面の「⚙️ 設定を変える」で直してから記録してください。**\n"
                    "古い設定のまま記録すると、そのデータは集計に使えなくなります。\n"
                    "間違えたときは「↩️ 直前の記録を取消」で消せます。\n"
                    "\n"
                    "**📋 データの取り扱いについて**\n"
                    "記録された戦績は、全体の統計として集計されます。\n"
                    "・個人の成績が名前つきで公開されることはありません\n"
                    "・「このデッキの全体勝率は○%」といった形で使われます\n"
                    "・集計結果は攻略記事や動画の材料になります\n"
                    "・掲示板に出るのは使用デッキとランク帯のみで、勝敗や勝率は出ません\n"
                    "-# 初回は `/戦績 設定` で先にデッキ・ランクを登録してください。"
                ),
                view=SensekiPanelView(),
            )
        except discord.Forbidden:
            await interaction.followup.send(
                self._missing_perm_message(interaction, "パネル"), ephemeral=True
            )
            return
        except Exception as e:
            await interaction.followup.send(
                f"❌ パネルの設置に失敗しました: {type(e).__name__}: {e}", ephemeral=True
            )
            return

        pinned = True
        try:
            await message.pin()
        except discord.Forbidden:
            pinned = False
        except discord.HTTPException:
            pinned = False

        note = "" if pinned else "\n-# ピン留めの権限がないため、手動でピン留めしてください。"
        await interaction.followup.send(f"✅ パネルを設置しました。{note}", ephemeral=True)

    # ---- 現状掲示板（全プレイヤーのデッキ＆ランク一覧） ----

    BOARD_CHANNEL_KEY = "board_channel_id"
    BOARD_MESSAGE_KEY = "board_message_id"

    async def _build_board_embed(self) -> discord.Embed:
        all_settings = await fetch_all_user_settings()
        resolve = self._name_resolver()

        by_class: dict[str, list[str]] = {}
        for uid, s in all_settings.items():
            deck = s.get("my_deck") or "デッキ未設定"
            rank = s.get("rank_tier") or "ランク未設定"
            if s.get("cr_grade"):
                rank += f"・{s['cr_grade']}"
            elif s.get("rank_tier") == GRAND_MASTER_TIER:
                rank += "・グレードなし"
            line = f"・{resolve(uid)}：{deck}（{rank}）"
            by_class.setdefault(s["my_class"], []).append(line)

        embed = discord.Embed(
            title="📋 現在の使用デッキ＆ランク",
            description="`/戦績 設定` `/デッキ 切替` `/戦績 ランク更新` のいずれかを実行すると自動で更新されます。",
            color=0x2E6DA4,
        )
        if not by_class:
            embed.add_field(name="―", value="まだ誰も設定していません。", inline=False)
        else:
            for cls in CLASS_CHOICES:
                if cls in by_class:
                    value = "\n".join(by_class[cls])[:1024]
                    embed.add_field(name=f"【{cls}】", value=value, inline=False)
        embed.timestamp = datetime.now(JST)
        return embed

    async def _update_board(self):
        """設定変更のたびに呼ぶ。掲示板が未設置なら何もしない"""
        channel_id = await get_state(self.BOARD_CHANNEL_KEY)
        message_id = await get_state(self.BOARD_MESSAGE_KEY)
        if not channel_id or not message_id:
            return

        channel = self.bot.get_channel(int(channel_id))
        if channel is None:
            print(f"⚠️ 戦績掲示板のチャンネルが見つかりません（ID: {channel_id}）")
            return

        try:
            message = await channel.fetch_message(int(message_id))
            embed = await self._build_board_embed()
            await message.edit(embed=embed)
        except discord.NotFound:
            print("⚠️ 戦績掲示板のメッセージが見つかりません。`/戦績管理 掲示板設置` をやり直してください。")
        except discord.Forbidden:
            print("⚠️ 戦績掲示板を編集する権限がありません。")
        except Exception as e:
            print(f"⚠️ 戦績掲示板の更新に失敗しました: {type(e).__name__}: {e}")

    @admin.command(name="掲示板設置", description="全員の使用デッキ＆ランクを常時表示する掲示板を作ります")
    async def senseki_board_setup(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True, thinking=True)
        embed = await self._build_board_embed()
        try:
            message = await interaction.channel.send(embed=embed)
        except discord.Forbidden:
            await interaction.followup.send(
                self._missing_perm_message(interaction, "掲示板"), ephemeral=True
            )
            return
        except Exception as e:
            await interaction.followup.send(
                f"❌ 掲示板の設置に失敗しました: {type(e).__name__}: {e}", ephemeral=True
            )
            return

        pinned = True
        try:
            await message.pin()
        except discord.Forbidden:
            pinned = False
        except discord.HTTPException:
            pinned = False

        await set_state(self.BOARD_CHANNEL_KEY, str(interaction.channel.id))
        await set_state(self.BOARD_MESSAGE_KEY, str(message.id))
        note = "" if pinned else "\n-# ピン留めの権限がないため、手動でピン留めしてください。"
        await interaction.followup.send(f"✅ 掲示板を設置しました。{note}", ephemeral=True)

    # ---- Googleスプレッドシート同期 ----

    def _name_resolver(self):
        guild = self.bot.get_guild(GUILD_ID)

        def resolve(user_id: str) -> str:
            if guild is not None:
                member = guild.get_member(int(user_id))
                if member is not None:
                    return member.display_name
            return f"不明なユーザー（{user_id}）"

        return resolve

    async def _run_sheets_sync(self) -> dict:
        matches = await fetch_all_matches(env_only=False)
        global_summary = await get_global_summary()
        matchup_rows = await get_class_matchup_matrix()
        # クラス別集計は「自分が使った試合」と「相手として対面した試合」の両方を
        # 合算した対称版を使う（詳しくは combine_symmetric_class_stats のdocstring参照）
        class_summary = combine_symmetric_class_stats(
            global_summary["by_my_class"], global_summary["by_opp_class"]
        )
        return await asyncio.to_thread(
            senseki_sheets.sync_to_sheets,
            matches, FORMAT_LABELS, self._name_resolver(),
            class_summary, global_summary["by_first"],
            matchup_rows, CLASS_CHOICES, MATCHUP_MATRIX_MIN_SAMPLE, current_env_id(),
        )

    @tasks.loop(time=dtime(hour=SHEETS_SYNC_HOUR, minute=SHEETS_SYNC_MINUTE, tzinfo=JST))
    async def daily_sheets_sync(self):
        if senseki_sheets is None or not senseki_sheets.is_configured():
            return  # 未設定なら静かにスキップ（毎日ログを汚さない）
        try:
            result = await self._run_sheets_sync()
            print(f"✅ 戦績スプレッドシート同期完了: {result['raw']}件 / {result['users']}名")
            for w in result.get("warnings", []):
                print(f"⚠️ 戦績スプレッドシート同期の一部が失敗: {w}")
        except Exception as e:
            print(f"⚠️ 戦績スプレッドシート同期に失敗しました: {type(e).__name__}: {e}")

    @daily_sheets_sync.before_loop
    async def _before_daily_sheets_sync(self):
        await self.bot.wait_until_ready()

    @tasks.loop(seconds=SHEETS_DEBOUNCE_SECONDS)
    async def debounced_sheets_sync(self):
        """記録・取消があった時だけシートへ反映する（まとめ書き）"""
        if senseki_sheets is None or not senseki_sheets.is_configured():
            return
        if not take_dirty():
            return  # 変更がなければAPIを呼ばない
        try:
            result = await self._run_sheets_sync()
            for w in result.get("warnings", []):
                print(f"⚠️ 戦績スプレッドシートの自動同期の一部が失敗: {w}")
        except Exception as e:
            print(f"⚠️ 戦績スプレッドシートの自動同期に失敗しました: {type(e).__name__}: {e}")
            mark_dirty()  # 次の周回で再試行する

    @debounced_sheets_sync.before_loop
    async def _before_debounced_sheets_sync(self):
        await self.bot.wait_until_ready()

    @admin.command(name="シート同期", description="戦績データをGoogleスプレッドシートへ即時同期します")
    async def senseki_sheets_sync(self, interaction: discord.Interaction):
        if senseki_sheets is None:
            await interaction.response.send_message(
                f"senseki_sheets.py を読み込めませんでした。\n```\n{SHEETS_MODULE_ERROR}\n```",
                ephemeral=True,
            )
            return

        if not senseki_sheets.is_configured():
            await interaction.response.send_message(
                "スプレッドシート連携の設定が未完了です。\n"
                f"```\n{senseki_sheets.describe_config()}\n```\n"
                "Railway の Variables と requirements.txt を確認してください。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            result = await self._run_sheets_sync()
            await interaction.followup.send(
                f"✅ 同期しました（{result['raw']}件 / {result['users']}名）\n{result['url']}"
                + (
                    "\n\n⚠️ 一部の追加処理が失敗しました（データ自体は同期済みです）：\n"
                    + "\n".join(f"- {w}" for w in result["warnings"])
                    if result.get("warnings") else ""
                ),
                ephemeral=True,
            )
        except Exception as e:
            print(f"⚠️ 戦績シート同期でエラー: {type(e).__name__}: {e}")
            hint = ""
            if "403" in str(e) or "PERMISSION" in str(e).upper():
                hint = (
                    "\n\nスプレッドシートがサービスアカウントに共有されていない可能性があります。"
                    "`/戦績管理 シート診断` でメールアドレスを確認し、そのアドレスに"
                    "「編集者」権限で共有してください。"
                )
            await interaction.followup.send(
                f"同期に失敗しました。\n```\n{type(e).__name__}: {e}\n```{hint}",
                ephemeral=True,
            )

    @admin.command(name="権限確認", description="このチャンネルでBOTが持っている権限を確認します")
    @app_commands.describe(channel="確認したいチャンネル（省略すると今いるチャンネル）")
    async def perm_check(
        self,
        interaction: discord.Interaction,
        channel: discord.TextChannel | None = None,
    ):
        target = channel or interaction.channel
        me = interaction.guild.me if interaction.guild else None
        if me is None or target is None:
            await interaction.response.send_message(
                "チャンネル情報を取得できませんでした。", ephemeral=True
            )
            return

        p = target.permissions_for(me)
        items = [
            ("チャンネルを見る", p.view_channel, "これが❌だとBOTからチャンネルが存在しない扱いになります（50001）"),
            ("メッセージを送信", p.send_messages, "パネル・掲示板・通知の投稿に必要"),
            ("メッセージの管理", p.manage_messages, "パネルの自動ピン留めに必要"),
            ("埋め込みリンク", p.embed_links, "Embed（掲示板・通知）の表示に必要"),
            ("メッセージ履歴を読む", p.read_message_history, "掲示板の更新に必要"),
        ]
        lines = [f"**{target.mention} でのBOTの権限**", ""]
        ng = []
        for name, ok, why in items:
            lines.append(f"{'✅' if ok else '❌'} {name}")
            if not ok:
                ng.append(f"・**{name}** — {why}")

        if ng:
            lines += [
                "",
                "**不足している権限**",
                *ng,
                "",
                "チャンネルの編集 → 権限 → BOT（またはBOTのロール）を追加して許可してください。",
            ]
        else:
            lines += ["", "問題ありません。"]

        await interaction.response.send_message("\n".join(lines), ephemeral=True)

    @admin.command(name="分析利用状況", description="分析系コマンド（確認・弱点対面など）の利用状況を確認します")
    async def usage_stats_cmd(self, interaction: discord.Interaction):
        stats = await get_command_usage_stats()
        lines = ["📊 **分析コマンドの利用状況**"]
        if not stats["by_command"]:
            lines.append("まだ利用ログがありません。")
            await interaction.response.send_message("\n".join(lines), ephemeral=True)
            return

        for row in stats["by_command"]:
            lines.append(f"・{row['command']}：{row['total']}回（{row['users']}名）")

        lines.append(f"\n直近30日：{stats['recent_30_total']}回（{stats['recent_30_users']}名）")

        recorders = stats["recorders"]
        analyzers = stats["analyzers"]
        rate = (analyzers / recorders * 100) if recorders else 0
        lines.append(
            f"\n戦績を記録したことがある人：{recorders}名\n"
            f"分析コマンドを使ったことがある人：{analyzers}名（記録者の{rate:.0f}%）"
        )
        a = await get_activity_stats()
        lines.append(
            f"\n**記録した人数（目標100名）**\n"
            f"・直近7日：{a['this_week_users']}名（{a['this_week_matches']}戦）\n"
            f"・その前の7日：{a['last_week_users']}名（{a['last_week_matches']}戦）\n"
            f"\n**🧪 デッキ検証**\n"
            f"・タグ付きで記録したことがある人：{a['variant_users']}名（{a['variant_matches']}戦）\n"
            f"・今タグを設定している人：{a['active_tag_users']}名"
        )
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    @admin.command(name="無料トライアル開始", description="初回参加者向け無料トライアルキャンペーンを開始します（このコマンド実行以降に初回登録した人だけが対象）")
    @app_commands.describe(
        上書き="既に開始済みでも、今の日時で開始日を上書きする場合はオン（通常は使いません）",
    )
    async def start_trial_campaign_cmd(
        self, interaction: discord.Interaction, 上書き: bool = False
    ):
        existing = await get_state(TRIAL_CAMPAIGN_START_KEY)
        if existing and not 上書き:
            await interaction.response.send_message(
                f"⚠️ 無料トライアルキャンペーンは既に **{existing[:19].replace('T', ' ')}** から開始済みです。\n"
                f"開始日を今の日時に上書きしたい場合は「上書き」をオンにして再実行してください。",
                ephemeral=True,
            )
            return

        now = datetime.now(JST).isoformat()
        await set_state(TRIAL_CAMPAIGN_START_KEY, now)
        await interaction.response.send_message(
            f"✅ 無料トライアルキャンペーンを開始しました（開始日時：{now[:19].replace('T', ' ')}）\n\n"
            f"これ以降に `/戦績 設定` を初めて実行した人だけが対象になります。\n"
            f"それ以前から使っている既存ユーザーは対象外です（今の環境は"
            + ("引き続き全員無料のままです。" if premium_free_this_env() else "既に有料化されています。")
            + f"）\nトライアル期間：{TRIAL_DAYS}日間",
            ephemeral=True,
        )

    @admin.command(name="無料トライアル状況", description="無料トライアルキャンペーンの開始日と、現在トライアル中の人数を確認します")
    async def trial_campaign_status_cmd(self, interaction: discord.Interaction):
        started = await get_state(TRIAL_CAMPAIGN_START_KEY)
        if not started:
            await interaction.response.send_message(
                "まだ無料トライアルキャンペーンは開始していません。"
                "`/戦績管理 無料トライアル開始` で開始できます。",
                ephemeral=True,
            )
            return

        rows = await fetch_all_user_settings()
        in_trial = 0
        for r in rows.values():
            s = r.get("started_at")
            if not s:
                continue
            start = _parse_iso(s)
            campaign_start = _parse_iso(started)
            if start and campaign_start and start >= campaign_start and trial_days_left(s) > 0:
                in_trial += 1

        await interaction.response.send_message(
            f"📅 無料トライアルキャンペーン開始日：{started[:19].replace('T', ' ')}\n"
            f"現在トライアル中の人数：{in_trial}名\n"
            f"（現在の環境は" + ("全員無料のため、トライアルの有無に関わらず②が使えます。" if premium_free_this_env() else "有料化済みです。") + "）",
            ephemeral=True,
        )

    @admin.command(name="デッキ名統合", description="表記ゆれのあるデッキ名を1つに統合します（過去の記録も書き換えます）")
    @app_commands.describe(
        クラス="対象クラス",
        統合前デッキ名="書き換え前のデッキ名（表記ゆれの方。完全一致で指定）",
        統合後デッキ名="正式名として残すデッキ名",
        公式デッキにする="今後の候補一覧で優先表示する場合はオン（デフォルトON）",
    )
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    async def merge_deck_names_cmd(
        self,
        interaction: discord.Interaction,
        クラス: app_commands.Choice[str],
        統合前デッキ名: str,
        統合後デッキ名: str,
        公式デッキにする: bool = True,
    ):
        await interaction.response.defer(ephemeral=True, thinking=True)
        old_name = 統合前デッキ名.strip()
        new_name = 統合後デッキ名.strip()
        if not old_name or not new_name:
            await interaction.followup.send("デッキ名が空です。", ephemeral=True)
            return
        if old_name == new_name:
            await interaction.followup.send("統合前後で同じ名前です。", ephemeral=True)
            return

        result = await merge_deck_name(クラス.value, old_name, new_name, 公式デッキにする)
        mark_dirty()

        note = "\n・今後の候補一覧で優先表示されるようにしました。" if 公式デッキにする else ""
        await interaction.followup.send(
            f"✅ **{クラス.name}**：「{old_name}」→「{new_name}」に統合しました\n"
            f"・自分側のデッキとして記録された試合：{result['matches_self']}件\n"
            f"・相手側のデッキとして記録された試合：{result['matches_opp']}件\n"
            f"・登録デッキ（ショートカット）：{result['templates']}件\n"
            f"・今まさにこのデッキを使用中だった人：{result['current_settings']}名（切替不要で反映済み）"
            + note,
            ephemeral=True,
        )

    @admin.command(name="記録管理", description="記録の修正・削除の履歴確認と、削除された記録の復元（管理者のみ）")
    @app_commands.describe(
        操作="履歴＝ユーザー別の修正・削除状況／削除済み＝指定した人が削除した記録とID／復元＝IDを指定して戻す",
        ユーザー="「削除済み」で、確認する人",
        記録id="「復元」で、戻す記録のID（「削除済み」で確認できます）",
        日数="「履歴」でさかのぼる日数（デフォルト7日）",
    )
    @app_commands.choices(操作=[
        app_commands.Choice(name="履歴（ユーザー別の修正・削除状況）", value="history"),
        app_commands.Choice(name="削除済み（指定した人の削除した記録）", value="deleted"),
        app_commands.Choice(name="復元（IDを指定して戻す）", value="restore"),
    ])
    async def record_admin_cmd(
        self, interaction: discord.Interaction, 操作: app_commands.Choice[str],
        ユーザー: discord.Member = None, 記録id: int = None,
        日数: app_commands.Range[int, 1, 60] = 7,
    ):
        await interaction.response.defer(ephemeral=True, thinking=True)
        none = discord.AllowedMentions.none()
        if 操作.value == "history":
            report = await edit_report(日数)
            if not report:
                await interaction.followup.send(f"直近{日数}日に、修正・削除・取消はありませんでした。", ephemeral=True)
                return
            rows = sorted(report.items(), key=lambda kv: -(kv[1]["edit"] + kv[1]["delete"]))
            lines = [f"🔎 **記録の修正・削除（直近{日数}日）**",
                     "-# ⚠️は、削除が10件以上／負けの削除が5件以上／負→勝への修正が3件以上のどれかです。"]
            for uid, u in rows[:20]:
                warn = u["delete"] >= 10 or u["del_lose"] >= 5 or u["to_win"] >= 3
                lines.append(
                    f"{'⚠️ ' if warn else ''}<@{uid}>：修正{u['edit']}（負→勝{u['to_win']}）／削除{u['delete']}（うち負け{u['del_lose']}）"
                    f"／取消{u['undo']}／復元{u['restore']}" + ("／Web経由あり" if u["web"] else ""))
            if len(rows) > 20:
                lines.append(f"…ほか{len(rows) - 20}名")
            await interaction.followup.send("\n".join(lines)[:1900], ephemeral=True, allowed_mentions=none)
        elif 操作.value == "deleted":
            if ユーザー is None:
                await interaction.followup.send("「ユーザー」を指定してください。", ephemeral=True)
                return

            def _q():
                conn = _connect()
                try:
                    return [dict(r) for r in conn.execute(
                        "SELECT * FROM matches_deleted WHERE user_id = ? ORDER BY deleted_at DESC LIMIT 15",
                        (str(ユーザー.id),)).fetchall()]
                finally:
                    conn.close()

            rows = await asyncio.to_thread(_q)
            if not rows:
                await interaction.followup.send("削除された記録はありません。", ephemeral=True)
                return
            lines = [f"🗑️ **{ユーザー.display_name} さんの削除済み記録**（新しい順）"]
            for r in rows:
                lines.append(f"・ID **{r['id']}**｜{_match_label(r)}｜削除 {r['deleted_at'][5:16].replace('T', ' ')}")
            lines.append("-# 戻すには、操作「復元」にIDを入れてください。")
            await interaction.followup.send("\n".join(lines)[:1900], ephemeral=True, allowed_mentions=none)
        else:
            if 記録id is None:
                await interaction.followup.send("「記録id」を指定してください。", ephemeral=True)
                return
            ok, msg, _ = await restore_match("", 記録id, "admin", admin=True)
            await interaction.followup.send(("✅ " if ok else "⚠️ ") + msg, ephemeral=True)

    @admin.command(name="診断", description="データの保存先と永続化状況を確認します")
    async def storage_diag(self, interaction: discord.Interaction):
        counts = await get_record_counts()
        persistent = _is_persistent(os.path.dirname(DB_PATH))
        body = describe_storage()
        body += f"\n\n記録件数: {counts['matches']}件\n設定済みユーザー: {counts['users']}名"
        if MATPLOTLIB_IMPORT_ERROR:
            body += f"\n\nグラフ機能: 無効（matplotlib未導入）\n  {MATPLOTLIB_IMPORT_ERROR}"
        elif _JP_FONT_NAME is None:
            body += "\n\nグラフ機能: 無効（日本語フォント未検出。assets/NotoSansJP-Regular.otf を配置してください）"
        else:
            body += f"\n\nグラフ機能: 有効（フォント: {_JP_FONT_NAME}）"
        warn = ""
        if not persistent:
            warn = (
                "\n\n⚠️ **保存先が永続ボリュームではありません。**\n"
                "このままだと再デプロイのたびにデータが消えます。\n"
                "Railway の Volumes でマウント先を確認し、そのパスを環境変数 "
                "`SENSEKI_DB_PATH` に設定してください（例：`/app/data/senseki.db`）。"
            )
        await interaction.response.send_message(
            f"```\n{body}\n```{warn}", ephemeral=True
        )

    @admin.command(name="シート診断", description="スプレッドシート連携の設定状況を確認します")
    async def senseki_sheets_diag(self, interaction: discord.Interaction):
        if senseki_sheets is None:
            await interaction.response.send_message(
                f"senseki_sheets.py を読み込めませんでした。\n```\n{SHEETS_MODULE_ERROR}\n```",
                ephemeral=True,
            )
            return
        config_text = senseki_sheets.describe_config()

        # 何が足りないかを判定して、次にやることだけを出す
        todo = []
        if "未読込" in config_text:
            todo.append(
                "**requirements.txt に `gspread` と `google-auth` を追加**して再デプロイしてください。"
            )
        if "GOOGLE_SERVICE_ACCOUNT_JSON: 未設定" in config_text:
            todo.append("Railway の Variables に `GOOGLE_SERVICE_ACCOUNT_JSON` を設定してください。")
        if "SENSEKI_SPREADSHEET_ID: 未設定" in config_text:
            todo.append("Railway の Variables に `SENSEKI_SPREADSHEET_ID` を設定してください。")

        if todo:
            footer = "**次にやること**\n" + "\n".join(f"{i+1}. {t}" for i, t in enumerate(todo))
        else:
            footer = (
                "✅ 設定はすべて揃っています。`/戦績管理 シート同期` を実行してください。\n"
                "-# 同期時に権限エラーが出る場合は、上記のサービスアカウントのアドレスに"
                "スプレッドシートを「編集者」で共有できているか確認してください。"
            )

        await interaction.response.send_message(
            "```\n" + config_text + "\n```\n" + footer,
            ephemeral=True,
        )

    async def _build_export(self, env_only: bool, env_id: str = None) -> tuple[bytes, str, int]:
        matches = await fetch_all_matches(env_only=env_only, env_id=env_id)
        guild = self.bot.get_guild(GUILD_ID)
        data = await asyncio.to_thread(build_workbook, matches, guild)
        scope = (env_id or "現環境") if env_only else "全期間"
        filename = f"senseki_{scope}_{datetime.now(JST).strftime('%Y%m%d_%H%M')}.xlsx"
        return data, filename, len(matches)

    @tasks.loop(time=dtime(hour=WEEKLY_EXPORT_HOUR, minute=WEEKLY_EXPORT_MINUTE, tzinfo=JST))
    async def weekly_export(self):
        if datetime.now(JST).weekday() != WEEKLY_EXPORT_WEEKDAY:
            return
        if SENSEKI_ADMIN_USER_ID == 0:
            print("⚠️ SENSEKI_ADMIN_USER_ID が未設定のため、戦績データの週次送信をスキップしました")
            return
        if Workbook is None:
            print(f"⚠️ openpyxl が読み込めないため、戦績データの週次送信をスキップしました: {OPENPYXL_IMPORT_ERROR}")
            return
        try:
            data, filename, count = await self._build_export(env_only=True)
            admin = self.bot.get_user(SENSEKI_ADMIN_USER_ID) or await self.bot.fetch_user(SENSEKI_ADMIN_USER_ID)
            await admin.send(
                content=f"📊 戦績データ週次抽出（現環境・{count}件）",
                file=discord.File(io.BytesIO(data), filename=filename),
            )
        except Exception as e:
            print(f"⚠️ 戦績データの週次送信に失敗しました: {type(e).__name__}: {e}")

    @weekly_export.before_loop
    async def _before_weekly_export(self):
        await self.bot.wait_until_ready()

    @tasks.loop(minutes=1)
    async def env_watch(self):
        """ENV_SCHEDULE の切替日時を過ぎたら、前の環境のデータを管理者にExcelで送る"""
        now_env = current_env_id()
        last_env = await get_state(ENV_STATE_KEY)
        if last_env is None:
            # 初回起動：今の環境を覚えるだけ
            await set_state(ENV_STATE_KEY, now_env)
            return
        if last_env == now_env:
            return

        # 先に状態を更新して、送信に失敗しても毎分DMが飛び続けないようにする
        await set_state(ENV_STATE_KEY, now_env)
        print(f"🔄 戦績ツールの環境が切り替わりました：{last_env} → {now_env}")
        mark_dirty()
        try:
            await self._update_board()
        except Exception as e:
            print(f"⚠️ 環境切替時の掲示板更新に失敗しました: {type(e).__name__}: {e}")

        if SENSEKI_ADMIN_USER_ID == 0 or Workbook is None:
            print("⚠️ 管理者IDまたはopenpyxlが無いため、前の環境のExcel送付をスキップしました")
            return
        try:
            admin = self.bot.get_user(SENSEKI_ADMIN_USER_ID) or await self.bot.fetch_user(SENSEKI_ADMIN_USER_ID)
            data, filename, count = await self._build_export(env_only=True, env_id=last_env)
            await admin.send(
                content=(
                    f"🔄 戦績ツールの環境を切り替えました（{last_env} → {now_env}）\n"
                    f"📊 前の環境のデータ（{count}件）を保存用にお送りします。\n"
                    f"-# データは消していません。あとから `/戦績管理 データ抽出 前の環境:True` でも出力できます。"
                ),
                file=discord.File(io.BytesIO(data), filename=filename),
            )
        except Exception as e:
            print(f"⚠️ 前の環境のExcel送付に失敗しました: {type(e).__name__}: {e}")

    @env_watch.before_loop
    async def _before_env_watch(self):
        await self.bot.wait_until_ready()

    # ---- /戦績データ抽出 ----
    @admin.command(name="データ抽出", description="戦績データを生データ＋集計のExcelで出力します")
    @app_commands.describe(
        全期間="オンにすると環境をまたいだ全データを出力します（既定は現環境のみ）",
        前の環境="オンにすると1つ前の環境（前の弾）のデータだけを出力します",
    )
    async def senseki_export(self, interaction: discord.Interaction,
                             全期間: bool = False, 前の環境: bool = False):
        if Workbook is None:
            await interaction.response.send_message(
                f"openpyxlが読み込めませんでした。\n```\n{OPENPYXL_IMPORT_ERROR}\n```\n"
                "requirements.txt に `openpyxl` を追加してください。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            target_env = None
            if 前の環境 and not 全期間:
                target_env = previous_env_id()
                if target_env is None:
                    await interaction.followup.send("1つ前の環境がありません。", ephemeral=True)
                    return
            data, filename, count = await self._build_export(env_only=not 全期間, env_id=target_env)
            if count == 0:
                await interaction.followup.send("記録がまだありません。", ephemeral=True)
                return

            scope_label = "全期間" if 全期間 else (f"前の環境：{target_env}" if target_env else "現環境")
            summary = f"📊 戦績データ抽出（{scope_label}・{count}件）"
            try:
                await interaction.user.send(
                    content=summary,
                    file=discord.File(io.BytesIO(data), filename=filename),
                )
                await interaction.followup.send("DMにExcelを送付しました。", ephemeral=True)
            except discord.Forbidden:
                await interaction.followup.send(
                    content=f"DMを送信できなかったため、こちらに添付します。\n\n{summary}",
                    file=discord.File(io.BytesIO(data), filename=filename),
                    ephemeral=True,
                )
        except Exception as e:
            print(f"⚠️ 戦績データ抽出で予期しないエラー: {type(e).__name__}: {e}")
            await interaction.followup.send(
                f"処理中にエラーが発生しました。\n```\n{type(e).__name__}: {e}\n```",
                ephemeral=True,
            )

    # ---- /戦績全体 ----
    @admin.command(name="全体集計", description="サーバー全体の集計を今この瞬間のデータで表示します")
    async def senseki_global(self, interaction: discord.Interaction):
        g = await get_global_summary()
        total = g["total"]
        if total == 0:
            await interaction.response.send_message(
                "まだ記録がありません。", ephemeral=True
            )
            return

        def rate(wins, n):
            return f"{wins / n * 100:.1f}%" if n else "-"

        lines = [
            "📊 **全体集計（現環境）**",
            f"{total}戦 / {g['wins']}勝{total - g['wins']}敗 / 勝率 {rate(g['wins'], total)}",
            f"記録者：{g['users']}名",
        ]

        first_map = {r["is_first"]: r for r in g["by_first"]}
        parts = []
        for key, label in ((1, "先攻"), (0, "後攻")):
            r = first_map.get(key)
            if r:
                parts.append(f"{label} {rate(r['wins'], r['total'])}（{r['total']}戦）")
        if parts:
            lines.append("\n**先後別**\n" + " / ".join(parts))

        lines.append("\n**使用クラス別**")
        for r in g["by_my_class"]:
            lines.append(f"{r['my_class']}：{rate(r['wins'], r['total'])}（{r['total']}戦）")

        lines.append("\n**対面クラス別**")
        for r in g["by_opp_class"]:
            lines.append(f"vs {r['opp_class']}：{rate(r['wins'], r['total'])}（{r['total']}戦）")

        lines.append(
            f"\n※母数が少ない項目は参考値です（{MIN_SAMPLE_FOR_STATS}戦未満は特に）。"
        )
        await interaction.response.send_message("\n".join(lines), ephemeral=True)

    @admin.command(
        name="環境分析画像",
        description="現環境のクラス使用率・対面マトリクス・先攻後攻をPNG画像で表示します（管理者本人のみ）",
    )
    async def senseki_global_image(self, interaction: discord.Interaction):
        await self._show_global_analysis(interaction)

    # ---- /デッキ登録 ----
    @deck.command(name="登録", description="よく使うデッキを登録しておきます（複数登録できます）")
    @app_commands.describe(
        フォーマット="このデッキで遊ぶフォーマット",
        クラス="このデッキのクラス",
        デッキ名="デッキ名（例：進化ネメシス）",
    )
    @app_commands.choices(フォーマット=[
        app_commands.Choice(name="ローテーション", value="rotation"),
        app_commands.Choice(name="アンリミテッド", value="unlimited"),
    ])
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    async def deck_register(
        self,
        interaction: discord.Interaction,
        フォーマット: app_commands.Choice[str],
        クラス: app_commands.Choice[str],
        デッキ名: str,
    ):
        name = (デッキ名 or "").strip()
        if not name:
            await interaction.response.send_message("デッキ名を入力してください。", ephemeral=True)
            return

        ok = await add_template(
            str(interaction.user.id), フォーマット.value, クラス.value, name
        )
        if not ok:
            await interaction.response.send_message(
                "同じ組み合わせのデッキがすでに登録されています。", ephemeral=True
            )
            return
        await interaction.response.send_message(
            f"✅ 登録しました\n{name}（{クラス.name} / {フォーマット.name}）\n\n"
            "`/デッキ 切替` で選ぶだけで使うデッキを変更できます。",
            ephemeral=True,
        )

    # ---- /デッキ切替 ----
    @deck.command(name="切替", description="登録済みのデッキから、今使うデッキを選びます")
    @app_commands.describe(削除="オンにすると、選んだデッキを登録から削除します")
    async def deck_switch(self, interaction: discord.Interaction, 削除: bool = False):
        templates = await list_templates(str(interaction.user.id))
        if not templates:
            await interaction.response.send_message(
                "登録済みのデッキがありません。`/デッキ 登録` で登録してください。",
                ephemeral=True,
            )
            return

        settings = await get_user_settings(str(interaction.user.id))
        if settings is None and not 削除:
            await interaction.response.send_message(
                "先に `/戦績 設定` でランク帯まで登録してください。", ephemeral=True
            )
            return

        mode = "delete" if 削除 else "switch"
        header = "削除するデッキを選んでください。" if 削除 else "使うデッキを選んでください。"
        await interaction.response.send_message(
            header, view=TemplateView(templates, mode), ephemeral=True
        )

    @deck.command(name="検証開始", description="今のデッキに個人用の検証タグを付けます（自分だけに見える集計用）")
    @app_commands.describe(タグ名="例：アンテマリア採用、フルスペル型 など")
    async def start_variant(self, interaction: discord.Interaction, タグ名: str):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return
        label = タグ名.strip()[:50]
        if not label:
            await interaction.response.send_message("タグ名を入力してください。", ephemeral=True)
            return
        await set_active_variant(
            str(interaction.user.id), settings["format"], settings["my_class"], settings["my_deck"], label
        )
        await interaction.response.send_message(
            f"🏷 検証タグ「{label}」を開始しました。", ephemeral=True
        )

    @deck.command(name="検証終了", description="今のデッキの検証タグを解除します")
    async def stop_variant(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return
        ok = await clear_active_variant(
            str(interaction.user.id), settings["format"], settings["my_class"], settings["my_deck"]
        )
        msg = "🏷 検証タグを解除しました。" if ok else "現在、有効な検証タグはありません。"
        await interaction.response.send_message(msg, ephemeral=True)

    @deck.command(name="検証集計", description="今のデッキの検証タグごとの勝率を確認します（自分だけに表示）")
    async def variant_summary_cmd(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return
        rows = await get_variant_summary(
            str(interaction.user.id), settings["format"], settings["my_class"], settings["my_deck"]
        )
        if not rows:
            await interaction.response.send_message(
                f"**{settings['my_deck']}**にはまだ検証タグ付きの記録がありません。", ephemeral=True
            )
            return
        lines = [f"🧪 **{settings['my_deck']}** 個人検証タグ別勝率（自分だけに表示）"]
        for r in rows:
            total = r["total"]; wins = r["wins"] or 0
            rate = wins / total * 100 if total else 0
            lines.append(f"・**{r['variant_label']}**：勝率 {rate:.1f}%（{wins}勝{total-wins}敗・{total}戦）")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    # ---- /ランク更新 ----
    @senseki.command(name="ランク更新", description="ランク帯とグレードだけを変更します（デッキ設定はそのまま）")
    @app_commands.describe(
        ランク帯="現在のランク帯",
        グレード="グランドマスターの方のみ必須。グレードが付いていない場合は「グレードなし」を選択",
    )
    @app_commands.choices(ランク帯=[app_commands.Choice(name=r, value=r) for r in RANK_TIER_CHOICES])
    @app_commands.choices(グレード=(
        [app_commands.Choice(name=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
        + [app_commands.Choice(name=g, value=g) for g in CR_GRADE_CHOICES]
    ))
    async def rank_update(
        self,
        interaction: discord.Interaction,
        ランク帯: app_commands.Choice[str],
        グレード: app_commands.Choice[str] = None,
    ):
        rank_value = ランク帯.value
        raw_grade = グレード.value if グレード else None

        grade_value, note, error = _resolve_grade(rank_value, raw_grade)
        if error:
            await interaction.response.send_message(error, ephemeral=True)
            return

        ok = await update_rank(str(interaction.user.id), rank_value, grade_value)
        if not ok:
            await interaction.response.send_message(
                "先に `/戦績 設定` で初期設定をしてください。", ephemeral=True
            )
            return

        label = f"{rank_value}（{grade_value}）" if grade_value else rank_value
        lines = [f"✅ ランク帯を更新しました：{label}"]
        if note:
            lines.append(note)
        await interaction.response.send_message("\n".join(lines), ephemeral=True)
        await self._update_board()

    # ---- /戦績設定 ----
    @senseki.command(name="設定", description="フォーマット・自分のクラス・デッキタイプ・ランク帯を登録します")
    @app_commands.describe(
        フォーマット="使用するフォーマット",
        クラス="自分の使用クラス",
        デッキタイプ="デッキ名（例：進化ネメシス）",
        ランク帯="現在のランク帯",
        グレード="グランドマスターの方のみ必須。CRのグレードを選択してください",
    )
    @app_commands.choices(フォーマット=[
        app_commands.Choice(name="ローテーション", value="rotation"),
        app_commands.Choice(name="アンリミテッド", value="unlimited"),
    ])
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(ランク帯=[app_commands.Choice(name=r, value=r) for r in RANK_TIER_CHOICES])
    @app_commands.choices(グレード=(
        [app_commands.Choice(name=CR_GRADE_NONE_LABEL, value=CR_GRADE_NONE_VALUE)]
        + [app_commands.Choice(name=g, value=g) for g in CR_GRADE_CHOICES]
    ))
    async def senseki_settings(
        self,
        interaction: discord.Interaction,
        フォーマット: app_commands.Choice[str],
        クラス: app_commands.Choice[str],
        デッキタイプ: str,
        ランク帯: app_commands.Choice[str],
        グレード: app_commands.Choice[str] = None,
    ):
        rank_value = ランク帯.value
        raw_grade = グレード.value if グレード else None
        deck_value = (デッキタイプ or "").strip()

        # 空白だけの入力を弾く
        if not deck_value:
            await interaction.response.send_message(
                "デッキタイプを入力してください。", ephemeral=True
            )
            return

        grade_value, note, error = _resolve_grade(rank_value, raw_grade)
        if error:
            await interaction.response.send_message(error, ephemeral=True)
            return

        await upsert_user_settings(
            str(interaction.user.id), フォーマット.value, クラス.value,
            deck_value, rank_value, grade_value,
        )
        # ここで指定したデッキはテンプレートにも入れておく（次回から選ぶだけで済む）
        await add_template(
            str(interaction.user.id), フォーマット.value, クラス.value, deck_value
        )
        rank_label = f"{rank_value}（{grade_value}）" if grade_value else rank_value
        lines = [
            "✅ 設定を保存しました",
            f"フォーマット：{フォーマット.name}",
            f"クラス：{クラス.name}",
            f"デッキタイプ：{deck_value}",
            f"ランク帯：{rank_label}",
        ]
        if note:
            lines.append(note)
        lines.append("\n`/戦績 記録` で記録できます。相手クラス・先後・勝敗だけ入力すればOKです。")
        lines.append("ランクが変わったら `/戦績 ランク更新`、デッキを変えるときは `/デッキ 切替` が早いです。")
        await interaction.response.send_message("\n".join(lines), ephemeral=True)
        await self._update_board()

    async def _start_record_flow(self, interaction: discord.Interaction):
        """記録の入り口。/戦績 とパネルのボタンの両方から呼ぶ"""
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None:
            await interaction.response.send_message(
                "先に `/戦績 設定` でフォーマット・クラスを登録してください。",
                ephemeral=True,
            )
            return

        view = await SensekiFlowView.create(str(interaction.user.id), settings, interaction)
        notice = "↩️ 前回選んでいた内容から再開します。" if view.restored else None
        await interaction.response.send_message(
            view.screen(
                view.guide_text(notice) + "\n"
                "-# 上の内容が違う場合は `/デッキ 切替` `/戦績 ランク更新` で変更してください。"
            ),
            view=view,
            ephemeral=True,
        )

    # ---- /戦績 ----
    @senseki.command(name="記録", description="対戦結果を記録します")
    async def senseki_record(self, interaction: discord.Interaction):
        await self._start_record_flow(interaction)

    # ---- /戦績取消 ----
    @senseki.command(name="取消", description="直前に記録した1件を取り消します")
    async def senseki_undo(self, interaction: discord.Interaction):
        deleted = await delete_last_match(str(interaction.user.id))
        if deleted is None:
            await interaction.response.send_message(
                "取り消せる記録が見つかりませんでした。", ephemeral=True
            )
            return
        result_label = "勝ち" if deleted["is_win"] else "負け"
        await interaction.response.send_message(
            f"🗑️ 直前の記録を取り消しました\n相手クラス：{deleted['opp_class']} / {result_label}",
            ephemeral=True,
        )

    # ---- /弱点対面 のグラフ（①のみ・無料枠） ----
    def _render_weak_matchups_chart_sync(self, my_deck: str, format_label: str, rows: list,
                                          deck_total: int = None, min_matches: int = None):
        """①勝率が低い対面を横棒グラフのPNGにする。フォント未導入などの場合はNoneを返す"""
        if plt is None or _JP_FONT_NAME is None or not rows:
            return None
        try:
            # 勝率の高い順に並べる → barhは先頭が下から積まれるので、
            # 一番勝率が低い（＝弱点として一番目立たせたい）対面が一番上に来る
            ordered = sorted(rows, key=lambda r: r["win_rate"], reverse=True)
            labels = [r["opp_class"] for r in ordered]
            values = [r["win_rate"] for r in ordered]

            norm = mcolors.Normalize(vmin=0, vmax=100)
            cmap = matplotlib.colormaps["RdYlGn"]
            colors = [cmap(norm(v)) for v in values]

            height = max(2.2, 0.55 * len(ordered) + 1.2)
            fig, ax = plt.subplots(figsize=(7.8, height), dpi=150)

            y = range(len(ordered))
            ax.barh(y, values, color=colors, edgecolor="#333333", linewidth=0.5, height=0.65)
            ax.set_yticks(list(y))
            ax.set_yticklabels(labels, fontsize=12)
            ax.set_xlim(0, 118)  # 100%のバーでも右側にラベル分の余白を確保
            ax.set_xticks([0, 20, 40, 60, 80, 100])
            ax.axvline(50, color="#888888", linestyle="--", linewidth=1)
            ax.set_xlabel("勝率（%）", fontsize=11)

            for yi, r in zip(y, ordered):
                losses = r["total"] - r["wins"]
                text = f"{r['win_rate']:.1f}%（{r['wins']}勝{losses}敗・{r['total']}戦）"
                ax.text(r["win_rate"] + 2, yi, text, va="center", ha="left", fontsize=10, color="#222222")

            ax.set_title(f"{my_deck}の弱点対面（{format_label}）", fontsize=14, pad=14)
            if deck_total is not None and min_matches is not None and deck_total < min_matches:
                ax.text(
                    0.5, 1.14,
                    f"⚠ 参考値：まだ{deck_total}戦です（{min_matches}戦を超えると精度が安定します）",
                    transform=ax.transAxes, ha="center", va="bottom", fontsize=10.5, color="#b5651d",
                )

            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            fig.tight_layout()

            buf = io.BytesIO()
            fig.savefig(buf, format="png")
            plt.close(fig)
            buf.seek(0)
            return buf
        except Exception:
            return None

    async def _render_weak_matchups_chart(self, my_deck: str, format_label: str, rows: list,
                                           deck_total: int = None, min_matches: int = None):
        return await asyncio.to_thread(
            self._render_weak_matchups_chart_sync, my_deck, format_label, rows, deck_total, min_matches
        )

    # ---- 環境分析（サーバー全体・現環境）のグラフ ----
    # スプレッドシートの「クラス別集計」「対面マトリクス」「先攻後攻集計」と同じ集計を、
    # PNG画像としてDiscord上でいつでも見られるようにするためのもの。
    # 個人の弱点対面と違い母数が母集団全体なので誰でも使ってよい機能という位置づけ。


    def _render_matchup_heatmap_sync(self, matchup_rows: list, classes: list,
                                      min_sample: int, env_id: str):
        """クラス対クラスの対面マトリクスをヒートマップのPNGにする。
        母数がしきい値未満のマスはグレーにして、件数だけ小さく添える
        （空欄にしてしまうと"データがない"のか"信頼できないだけ"なのか区別できないため）。
        """
        if plt is None or _JP_FONT_NAME is None or not matchup_rows:
            return None
        try:
            lookup = {(r["my_class"], r["opp_class"]): r for r in matchup_rows}
            n = len(classes)
            rate_grid = [[None] * n for _ in range(n)]
            total_grid = [[0] * n for _ in range(n)]
            for i, mc in enumerate(classes):
                for j, oc in enumerate(classes):
                    r = lookup.get((mc, oc))
                    total = r["total"] if r else 0
                    wins = r["wins"] if r else 0
                    total_grid[i][j] = total
                    rate_grid[i][j] = (wins / total * 100) if total >= min_sample else None

            display_grid = np.full((n, n), np.nan)
            for i in range(n):
                for j in range(n):
                    if rate_grid[i][j] is not None:
                        display_grid[i, j] = rate_grid[i][j]

            cmap = matplotlib.colormaps["RdYlGn"].copy()
            cmap.set_bad(color="#d9d9d9")  # 母数不足のマス

            fig, ax = plt.subplots(figsize=(8.2, 7.4), dpi=150)
            im = ax.imshow(display_grid, cmap=cmap, vmin=0, vmax=100, aspect="equal")

            ax.set_xticks(range(n))
            ax.set_xticklabels(classes, fontsize=11, rotation=30, ha="right")
            ax.set_yticks(range(n))
            ax.set_yticklabels(classes, fontsize=11)
            ax.set_xlabel("相手のクラス", fontsize=11)
            ax.set_ylabel("自分のクラス", fontsize=11)

            for i in range(n):
                for j in range(n):
                    total = total_grid[i][j]
                    rate = rate_grid[i][j]
                    if rate is not None:
                        text = f"{rate:.0f}%\n({total})"
                        color = "#1a1a1a" if 25 <= rate <= 75 else "white"
                    elif total > 0:
                        text = f"({total})"
                        color = "#555555"
                    else:
                        text = "-"
                        color = "#999999"
                    ax.text(j, i, text, ha="center", va="center", fontsize=9, color=color)

            ax.set_title(f"対面マトリクス（現環境：{env_id}）", fontsize=14, pad=14)
            fig.colorbar(im, ax=ax, label="勝率（%）", fraction=0.046, pad=0.04)
            fig.text(
                0.5, 0.005,
                f"グレーのマスは対戦数が{min_sample}戦未満（参考値扱い）。カッコ内は対戦数。",
                ha="center", fontsize=9, color="#666666",
            )
            fig.tight_layout(rect=[0, 0.02, 1, 1])

            buf = io.BytesIO()
            fig.savefig(buf, format="png")
            plt.close(fig)
            buf.seek(0)
            return buf
        except Exception:
            return None


    async def _show_global_analysis(self, interaction: discord.Interaction):
        """環境分析コマンドの本体。SENSEKI_ADMIN_USER_ID 本人のみ実行できる"""
        if interaction.user.id != SENSEKI_ADMIN_USER_ID:
            await interaction.response.send_message(
                "このコマンドは管理者専用です。", ephemeral=True
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        g, files, _ = await self._build_env_report_files()
        if g["total"] == 0:
            await interaction.followup.send("まだ記録がありません。", ephemeral=True)
            return

        await log_command_usage(str(interaction.user.id), "環境分析")

        if not files:
            await interaction.followup.send(
                "画像の生成に失敗しました（グラフ機能が無効になっている可能性があります。"
                "`/戦績管理 診断` を確認してください）。",
                ephemeral=True,
            )
            return

        await interaction.followup.send(
            f"📊 **現環境（{current_env_id()}）の全体分析**（{g['total']}戦・記録者{g['users']}名）",
            files=files,
            ephemeral=True,
        )

    async def _build_env_report_files(self):
        """
        現環境の集計と、レポート用の画像（discord.File）を返す。
        戻り値：(全体集計 g, 画像のリスト, デッキ単位の集計 data or None)
        画像は「よく当たるデッキTOP10」「環境マップ」「デッキ相性表」「クラス対面マトリクス」の順。
        """
        g = await get_global_summary()
        if g["total"] == 0:
            return g, [], None
        matches = await fetch_all_matches(env_only=True)
        now = datetime.now(JST)
        data = compute_report_data(matches, now)
        matchup_rows = await get_class_matchup_matrix()
        date_label = now.strftime("%m/%d")

        jobs = []
        names = []
        if data is not None:
            jobs += [
                asyncio.to_thread(render_deck_ranking, data, date_label),
                asyncio.to_thread(render_env_map, data),
            ]
            names += ["よく当たるデッキ.png", "環境マップ.png"]
            if REPORT_SHOW_DECK_MATRIX:
                jobs.append(asyncio.to_thread(render_deck_matrix, data))
                names.append("デッキ相性表.png")
        jobs.append(asyncio.to_thread(
            self._render_matchup_heatmap_sync, matchup_rows, CLASS_CHOICES,
            MATCHUP_MATRIX_MIN_SAMPLE, current_env_id(),
        ))
        names.append("クラス対面マトリクス.png")
        images = await asyncio.gather(*jobs)
        files = [discord.File(buf, filename=name) for buf, name in zip(images, names) if buf is not None]
        return g, files, data

    # ---- 環境レポート（2日に1回、サーバーへ自動投稿） ----

    async def _post_env_report(self, force: bool = False):
        """
        環境レポートを設置済みチャンネルに投稿する。
        戻り値：(投稿したか, 理由メッセージ)
        force=True のときは最低試合数のチェックを飛ばす（管理者の手動投稿用）
        """
        channel_id = await get_state(ENV_REPORT_CHANNEL_KEY)
        if not channel_id:
            return False, "投稿先が未設置です。`/戦績管理 環境レポート設置` を実行してください。"
        channel = self.bot.get_channel(int(channel_id))
        if channel is None:
            try:
                channel = await self.bot.fetch_channel(int(channel_id))
            except discord.HTTPException:
                return False, "投稿先のチャンネルが見つかりません。設置し直してください。"

        g, files, data = await self._build_env_report_files()
        if g["total"] == 0:
            return False, "現環境の記録がまだありません。"
        if not force and g["total"] < ENV_REPORT_MIN_MATCHES:
            return False, f"現環境の記録が{g['total']}戦で、投稿の目安（{ENV_REPORT_MIN_MATCHES}戦）に届いていません。"
        if not files:
            return False, "画像の生成に失敗しました。`/戦績管理 診断` を確認してください。"

        env_id = current_env_id()
        growth = ""
        last = await get_state(ENV_REPORT_LAST_KEY)
        if last:
            try:
                last_env, last_total, last_users = last.split("|")
                if last_env == env_id:
                    d_total = g["total"] - int(last_total)
                    d_users = g["users"] - int(last_users)
                    growth = f"（前回から +{max(d_total, 0)}戦・記録者 +{max(d_users, 0)}名）"
            except ValueError:
                pass

        panel_id = await get_state(ENV_REPORT_PANEL_KEY)
        discuss_id = await get_state(ENV_REPORT_DISCUSS_KEY)
        if ENV_REPORT_CREATE_THREAD:
            discuss_line = "💬 感想や考察は、下のスレッドへどうぞ！\n"
        elif discuss_id and discuss_id != str(channel_id):
            discuss_line = f"💬 感想や考察は <#{discuss_id}> へどうぞ！\n"
        elif discuss_id:
            # レポートを考察チャンネルそのものに投稿している場合は「ここへどうぞ」にする
            discuss_line = "💬 感想や考察は、このチャンネルでどうぞ！\n"
        else:
            discuss_line = ""
        where = f"<#{panel_id}> の" if panel_id else "戦績パネルの"
        first_line = ""
        if data and 1 in data["firsts"] and 0 in data["firsts"]:
            (n1, r1), (n0, r0) = data["firsts"][1], data["firsts"][0]
            first_line = f"先攻の勝率 **{r1:.1f}%**（{n1}戦）／後攻の勝率 **{r0:.1f}%**（{n0}戦）\n"

        ai_text = await self._build_ai_commentary(g, data)
        ai_block = ""
        if ai_text:
            ai_block = f"### 🤖 AIの読み解き\n{ai_text}\n{ENV_REPORT_AI_DISCLAIMER}\n\n"

        content = (
            f"## 📊 しろくさサーバー調べ：最新の環境レポート\n"
            f"現環境（{env_id}）の記録 **{g['total']}戦**・記録者 **{g['users']}名**{growth}\n"
            f"\n"
            f"みなさんが記録した試合から、よく当たるデッキ・対策すべきデッキ・デッキ同士の相性を集計しました。\n"
            f"{first_line}"
            f"記録が増えるほど、このレポートの精度が上がります🙌\n"
            f"\n"
            f"{ai_block}"
            f"⚔️ あなたの1戦も次回のレポートに反映されます → {where}「⚔️ 戦績を記録する」から\n"
            f"{discuss_line}"
            f"-# 集計はローテーションのみ。記録者の多いランク帯・デッキに偏る点はご了承ください。"
        )
        try:
            message = await channel.send(content=content, files=files)
        except discord.Forbidden:
            return False, "投稿先のチャンネルにメッセージ・ファイルを送る権限がありません。"
        thread_note = ""
        if ENV_REPORT_CREATE_THREAD:
            try:
                await message.create_thread(
                    name=f"{datetime.now(JST).strftime('%m/%d')} 環境レポートの感想・考察",
                    auto_archive_duration=4320,  # 3日（次のレポートまで残る）
                )
            except discord.HTTPException as e:
                thread_note = f"（スレッドは作れませんでした：BOTに「公開スレッドの作成」権限があるか確認してください / {type(e).__name__}）"
        await set_state(ENV_REPORT_LAST_KEY, f"{env_id}|{g['total']}|{g['users']}")
        await set_state(ENV_REPORT_LAST_DATE_KEY, datetime.now(JST).date().isoformat())
        if HINT_WITH_REPORT:
            # レポートの直後に「質問のヒント」を続けて出す（ヒントに使える相手がいなければ出さない）
            try:
                hint_matches = await fetch_all_matches(env_only=True)
                hints, hint_avg = compute_env_hints(hint_matches)
                if hints:
                    await channel.send(
                        build_env_hint_post(hints, hint_avg, await class_channel_map()), view=HintPostView())
            except Exception as e:
                print(f"⚠️ 質問のヒントの投稿に失敗しました: {type(e).__name__}: {e}")
        ai_note = ""
        if not ai_text:
            reason = getattr(self, "_last_ai_error", None) or "AIの返答が空でした"
            ai_note = f"\n⚠️ AI分析は付けられませんでした（モデル：`{ENV_REPORT_AI_MODEL}`）\n```\n{reason}\n```"
        return True, f"{channel.mention} に投稿しました（{g['total']}戦・記録者{g['users']}名）。{ai_note}{thread_note}"

    async def _build_ai_commentary(self, g: dict, report_data: dict = None):
        """集計値だけを渡して、AIに短い読み解きを書かせる。失敗したら None"""
        self._last_ai_error = None
        api_key = os.environ.get(ENV_REPORT_AI_KEY_ENV)
        if AsyncAnthropic is None or not api_key:
            self._last_ai_error = (
                f"anthropic={ANTHROPIC_IMPORT_ERROR or 'OK'} / APIキー={'あり' if api_key else 'なし'}"
            )
            print(f"ℹ️ AI分析をスキップ：{self._last_ai_error}")
            return None

        matches = await fetch_all_matches(env_only=True)
        data_text = build_ai_dataset(matches, current_env_id(), datetime.now(JST))
        if data_text is None:
            self._last_ai_error = "ローテーションの記録がないため、AIに渡す集計を作れませんでした"
            return None
        if report_data:
            pairs = []
            decks = report_data["matrix_decks"]
            for i, a in enumerate(decks):
                for j, b in enumerate(decks):
                    c = report_data["matrix"][i][j]
                    if i < j and c["n"] >= REPORT_MATRIX_MIN_JUDGE:
                        pairs.append(f"- {a[1]}（{a[0]}） 対 {b[1]}（{b[0]}）：{c['n']}戦・左から見た勝率{c['rate']:.1f}%")
            data_text += (
                f"\n\n■ デッキ同士の相性（{REPORT_MATRIX_MIN_JUDGE}戦以上の組み合わせのみ。"
                f"使った側・当たった側の記録を合算しているので、平均との差ではなく50%を基準に読む）\n"
                + ("\n".join(pairs) if pairs else "- （該当なし：まだ試合数が足りない）")
            )

        system = (
            "あなたはシャドウバース ワールズビヨンドのDiscordコミュニティで、"
            "メンバーが記録したランクマッチの対戦データを読み解く担当です。\n"
            "【データの読み方】\n"
            "- 環境の使用率は「相手側」の数字で判断する（相手はマッチングでランダムに当たるため）。"
            "記録者自身のデッキの数は、記録者の好みに偏るので使用率として扱わない。\n"
            "- データの「勝率」はすべて記録者自身の勝率。相手クラス・相手デッキの行の勝率は"
            "「そのクラス／デッキと当たったとき、記録者が勝った割合」であり、そのデッキ自体の勝率ではない。"
            "勝率が高い＝記録者にとって勝ちやすい相手（そのデッキは弱め）、低い＝手ごわい相手（そのデッキは強め）。"
            "この向きを取り違えない。\n"
            "- 記録者は上位ランク帯が多く、全体の勝率が50%より高い。"
            "勝ちやすさ・勝ちにくさは、必ず「記録者全体の平均勝率との差」で判断する。\n"
            "- 「直近」と「それより前」の比較があれば、使用率が増えた・減ったものを優先して書く。"
            "直近の試合数が100戦未満なら「兆し」「まだ断定できない」と表現する。\n"
            "【厳守するルール】\n"
            "1. 渡された数字だけを根拠にする。カード名・カードの効果・デッキの中身・強い理由には一切触れない"
            "（データに含まれていないため。推測で書くと誤情報になる）。デッキ名は渡されたものをそのまま使う。\n"
            "2. 30戦未満の数字には触れない。\n"
            "3. 記録が一部の記録者に偏っている場合は、そのことを1回だけ短く添える。\n"
            "4. 「最強」「環境トップ」「必ず勝てる」などの断定や煽りは使わない。\n"
            "5. 日本語・です・ます調で、箇条書き3〜4個、全体で300字以内。見出しや前置き、締めの挨拶は書かない。\n"
            "6. 「次の記録で確かめたいこと」などの今後の予告・宣言は書かない。\n"
            "【書き方（読みやすさ）】\n"
            "- 各項目は「結論を先に、数字は後から1つだけ添える」。"
            "例：「秘術ウィッチと当たったときの勝率は約49%で、記録者全体の平均（約60%）より低く、手ごわい相手です」"
            "「ランプ（ドラゴン）との対戦は約70%勝てていて、平均を大きく上回る、勝ちやすい相手です」。\n"
            "- 勝率を書くときは必ず「〇〇と当たったとき」「〇〇との対戦では」と、誰から見た数字かを主語に入れる。"
            "「ランプは勝率70%」のような、デッキ自体の勝率に読める書き方は禁止。\n"
            "- 「平均との差」のポイント数値（-11.6、+9.7など）や、括弧の中に数字を並べる書き方はしない。"
            "勝率は「約49%」のように整数で、試合数は「39戦で」のように文中で1回だけ必要なときに添える。\n"
            "- 相手デッキ・相手クラスについては、勝率が平均より低ければ「手ごわい相手」「勝ちにくい相手」、"
            "高ければ「勝ちやすい相手」と言葉で結論を書く。先攻・後攻は記録者自身の勝率なので「先攻だと約65%勝てている」と書く。"
            "「負けやすい傾向」などの曖昧な言い方は避ける。\n"
            "- 一つの項目に数字は多くても2つまで。難しい統計用語（「平均との差」「乖離」など）は使わない。"
        )
        text, used_model, error = await ask_ai(
            api_key, ENV_REPORT_AI_MODEL, system, f"次の集計を読み解いてください。\n\n{data_text}"
        )
        if not text:
            self._last_ai_error = error
            print(f"⚠️ 環境レポートのAI分析に失敗しました: {error}")
            return None
        if error is None and used_model != ENV_REPORT_AI_MODEL:
            print(f"ℹ️ {ENV_REPORT_AI_MODEL} で返答がなかったため {used_model} で分析しました")
        print(f"✅ 環境レポートのAI分析に成功しました（モデル：{used_model}）")
        # Discordの2000字制限に余裕を残す
        return text[:900]

    @tasks.loop(time=dtime(hour=ENV_REPORT_HOUR, minute=ENV_REPORT_MINUTE, tzinfo=JST))
    async def env_report_loop(self):
        last_date = await get_state(ENV_REPORT_LAST_DATE_KEY)
        if last_date:
            try:
                days = (datetime.now(JST).date() - datetime.fromisoformat(last_date).date()).days
                if days < ENV_REPORT_INTERVAL_DAYS:
                    return
            except ValueError:
                pass
        try:
            posted, reason = await self._post_env_report()
            if not posted:
                print(f"ℹ️ 環境レポートを投稿しませんでした: {reason}")
        except Exception as e:
            print(f"⚠️ 環境レポートの投稿に失敗しました: {type(e).__name__}: {e}")

    @env_report_loop.before_loop
    async def _before_env_report_loop(self):
        await self.bot.wait_until_ready()

    @admin.command(name="環境レポート設置", description="2日に1回12時に、環境レポートをこのチャンネルへ自動投稿します")
    @app_commands.describe(
        パネルのチャンネル="レポート内で「ここから記録」と案内する戦績パネルのチャンネル（任意）",
        考察チャンネル="レポート内で「感想・考察はこちら」と案内するチャンネル（任意）",
    )
    async def env_report_setup(self, interaction: discord.Interaction,
                               パネルのチャンネル: discord.TextChannel = None,
                               考察チャンネル: discord.TextChannel = None):
        await set_state(ENV_REPORT_CHANNEL_KEY, str(interaction.channel_id))
        if パネルのチャンネル is not None:
            await set_state(ENV_REPORT_PANEL_KEY, str(パネルのチャンネル.id))
        if 考察チャンネル is not None:
            await set_state(ENV_REPORT_DISCUSS_KEY, str(考察チャンネル.id))
        await interaction.response.send_message(
            f"✅ 環境レポートの投稿先を {interaction.channel.mention} に設定しました。\n"
            f"{ENV_REPORT_INTERVAL_DAYS}日に1回、{ENV_REPORT_HOUR}:{ENV_REPORT_MINUTE:02d} に自動投稿します"
            f"（現環境の記録が{ENV_REPORT_MIN_MATCHES}戦未満のときは投稿しません）。\n"
            f"-# 今すぐ投稿するには `/戦績管理 環境レポート投稿` を使ってください。",
            ephemeral=True,
        )

    @admin.command(name="環境レポート投稿", description="環境レポートを今すぐ投稿先のチャンネルへ投稿します")
    async def env_report_post_now(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            posted, reason = await self._post_env_report(force=True)
        except Exception as e:
            posted, reason = False, f"エラーが発生しました：{type(e).__name__}: {e}"
        await interaction.followup.send(("✅ " if posted else "❌ ") + reason, ephemeral=True)

    @admin.command(name="環境レポート停止", description="環境レポートの自動投稿を止めます")
    async def env_report_stop(self, interaction: discord.Interaction):
        await set_state(ENV_REPORT_CHANNEL_KEY, "")
        await interaction.response.send_message(
            "⏸ 環境レポートの自動投稿を止めました。再開するには `/戦績管理 環境レポート設置` を実行してください。",
            ephemeral=True,
        )

    # ---- 質問のヒント（毎週金曜の自動投稿） ----

    async def _post_hint(self, force: bool = False):
        """質問のヒントを設置済みチャンネルに投稿する。戻り値：(投稿したか, 理由)"""
        channel_id = await get_state(HINT_CHANNEL_KEY)
        if not channel_id:
            return False, "投稿先が未設置です。`/戦績管理 質問ヒント設置` を実行してください。"
        channel = self.bot.get_channel(int(channel_id))
        if channel is None:
            try:
                channel = await self.bot.fetch_channel(int(channel_id))
            except discord.HTTPException:
                return False, "投稿先のチャンネルが見つかりません。設置し直してください。"
        matches = await fetch_all_matches(env_only=True)
        hints, avg = compute_env_hints(matches)
        if not hints:
            return False, (f"ヒントに使える相手がまだありません（{HINT_ENV_MIN_GAMES}戦以上・"
                           f"記録者{HINT_ENV_MIN_USERS}名以上で、平均より勝ちにくい相手がいる場合に投稿します）。")
        await channel.send(build_env_hint_post(hints, avg, await class_channel_map()), view=HintPostView())
        await set_state(HINT_LAST_DATE_KEY, datetime.now(JST).date().isoformat())
        return True, f"{channel.mention} に投稿しました。"

    @tasks.loop(time=dtime(hour=HINT_POST_HOUR, minute=0, tzinfo=JST))
    async def hint_post_loop(self):
        if HINT_WITH_REPORT:
            return  # 環境レポートと一緒に出す設定のときは、単独の定期投稿はしない
        now = datetime.now(JST)
        if now.weekday() != HINT_POST_WEEKDAY:
            return
        if await get_state(HINT_LAST_DATE_KEY) == now.date().isoformat():
            return
        try:
            posted, reason = await self._post_hint()
            if not posted:
                print(f"ℹ️ 質問のヒントを投稿しませんでした: {reason}")
        except Exception as e:
            print(f"⚠️ 質問のヒントの投稿に失敗しました: {type(e).__name__}: {e}")

    @hint_post_loop.before_loop
    async def _before_hint_post_loop(self):
        await self.bot.wait_until_ready()

    @admin.command(name="質問ヒント設置", description="「質問のヒント」の単独投稿先をこのチャンネルに設定します（通常は環境レポートと一緒に出ます）")
    async def hint_setup(self, interaction: discord.Interaction):
        await set_state(HINT_CHANNEL_KEY, str(interaction.channel_id))
        await interaction.response.send_message(
            f"✅ 質問のヒントの投稿先を {interaction.channel.mention} に設定しました。\n"
            + ("環境レポートの直後に、一緒に投稿されます（この設定は手動投稿用です）。\n" if HINT_WITH_REPORT
               else f"毎週金曜 {HINT_POST_HOUR}:00 に自動投稿します（ヒントに使える相手がいない週は投稿しません）。\n") +
            "-# クラス別の質問先は初期設定済みです（変える場合は `/戦績管理 質問先設定`）。今すぐ投稿するには `/戦績管理 質問ヒント投稿`。",
            ephemeral=True,
        )

    @admin.command(name="質問ヒント投稿", description="質問のヒントを今すぐ投稿先のチャンネルへ投稿します")
    async def hint_post_now(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            posted, reason = await self._post_hint(force=True)
        except Exception as e:
            posted, reason = False, f"エラーが発生しました：{type(e).__name__}: {e}"
        await interaction.followup.send(("✅ " if posted else "❌ ") + reason, ephemeral=True)

    @admin.command(name="質問ヒント停止", description="質問のヒントの自動投稿を止めます")
    async def hint_stop(self, interaction: discord.Interaction):
        await set_state(HINT_CHANNEL_KEY, "")
        await interaction.response.send_message("⏸ 質問のヒントの自動投稿を止めました。", ephemeral=True)

    @admin.command(name="質問先設定", description="クラス別の質問の投稿先チャンネルを設定します")
    @app_commands.describe(クラス="対象クラス（デッキ選びの相談は「シャドバ相談所」、VCの予約に使うVCは「相談VC①〜③」）",
                           チャンネル="質問を投稿するチャンネル（相談VCのときはボイスチャンネル）")
    @app_commands.choices(クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES]
                          + [app_commands.Choice(name="シャドバ相談所", value=CONSULT_KEY)]
                          + [app_commands.Choice(name=f"相談VC{n}", value=k)
                             for n, k in zip("①②③", VC_SLOT_KEYS)]
                          + [app_commands.Choice(name="VC予約のお知らせ先", value=VC_NOTICE_KEY)])
    async def question_channel_setup(self, interaction: discord.Interaction,
                                     クラス: app_commands.Choice[str],
                                     チャンネル: Union[discord.TextChannel, discord.ForumChannel, discord.VoiceChannel]):
        is_vc_slot = クラス.value in VC_SLOT_KEYS
        if is_vc_slot != isinstance(チャンネル, discord.VoiceChannel):
            await interaction.response.send_message(
                "相談VCにはボイスチャンネルを、クラス・シャドバ相談所にはテキスト（またはフォーラム）チャンネルを選んでください。",
                ephemeral=True)
            return
        await set_state(QUESTION_CHANNEL_KEY_PREFIX + クラス.value, str(チャンネル.id))
        if クラス.value == VC_NOTICE_KEY:
            await interaction.response.send_message(
                f"✅ VC予約のお知らせを {チャンネル.mention} に出すように設定しました。", ephemeral=True)
            return
        if is_vc_slot:
            await interaction.response.send_message(
                f"✅ **{クラス.name}** を {チャンネル.mention} に設定しました。質問の「🎤 VCで相談する（予約）」で選べます。\n"
                "-# BOTに「イベントの作成」権限があると、Discordのイベントとして予約されます（なくてもスレッドへのお知らせは出ます）。",
                ephemeral=True)
            return
        await interaction.response.send_message(
            f"✅ **{クラス.name}** の質問先を {チャンネル.mention} に設定しました。\n"
            "-# BOTに、そのチャンネルの「メッセージを送信」「公開スレッドを作成」の権限が必要です。",
            ephemeral=True,
        )

    # ---- 個人のAI分析（パネルの「🤖 AIで分析」） ----
    async def _show_personal_ai(self, interaction: discord.Interaction):
        user_id = str(interaction.user.id)
        settings = await get_user_settings(user_id)
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず「🆕 初回設定」か「⚙️ 設定を変える」でデッキを登録してください。", ephemeral=True
            )
            return

        today = datetime.now(JST).date().isoformat()
        state_key = AI_PERSONAL_STATE_PREFIX + user_id
        if await get_state(state_key) == today:
            await interaction.response.send_message(
                "🤖 AI分析は1日1回までです。また明日お試しください。", ephemeral=True
            )
            return

        mine, everyone = await get_personal_ai_data(
            user_id, settings["format"], settings["my_class"], settings["my_deck"]
        )
        if len(mine) < AI_PERSONAL_MIN_MATCHES:
            await interaction.response.send_message(
                f"🤖 **{settings['my_deck']}**の記録が{len(mine)}戦です。"
                f"AI分析は{AI_PERSONAL_MIN_MATCHES}戦から使えます（今回は回数に数えません）。",
                ephemeral=True,
            )
            return

        api_key = os.environ.get(ENV_REPORT_AI_KEY_ENV)
        if AsyncAnthropic is None or not api_key:
            await interaction.response.send_message("🤖 AI分析は現在準備中です。", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        data_text = build_personal_ai_dataset(settings, mine, everyone)
        system = (
            "あなたはシャドウバース ワールズビヨンドのDiscordコミュニティで、"
            "メンバー本人の対戦記録を読み解く担当です。本人に向けて書く。\n"
            "【データの読み方】\n"
            "- 「ほかの人」は同じデッキを使った本人以外のメンバーの記録。本人の得意・苦手は、ほかの人との差で判断する。"
            "ほかの人の試合数が少ない項目は、比較の信頼度が低いと添える。\n"
            "- 対面ごとの数字は試合数が少ない。15戦未満のものに触れるときは「まだ試合数が少ない」と必ず添える。\n"
            "【厳守するルール】\n"
            "1. 渡された数字だけを根拠にする。カード名・カードの効果・デッキの中身・プレイングの具体的な指示には一切触れない"
            "（データに含まれていないため。推測で書くと誤情報になる）。デッキ名は渡されたものをそのまま使う。\n"
            "2. 本人を責める・煽る書き方はしない。「下手」「弱い」などの言葉は使わない。\n"
            "3. 日本語・です・ます調で、箇条書き4〜5個、全体で400字以内。見出しや前置き、締めの挨拶は書かない。\n"
            "4. 最後の箇条書きは「次に意識して記録・確認するとよいこと」を1つ書く"
            "（例：この対面の試合数を増やして確かめる、先攻後攻の差を見る など。プレイングの指示はしない）。"
        )
        text, _used, error = await ask_ai(
            api_key, AI_PERSONAL_MODEL, system, f"次の記録を読み解いてください。\n\n{data_text}"
        )
        if error:
            print(f"⚠️ 個人のAI分析に失敗しました: {error}")
        if not text:
            await interaction.followup.send(
                "🤖 AI分析に失敗しました。時間をおいてもう一度お試しください（今回は回数に数えません）。",
                ephemeral=True,
            )
            return

        await set_state(state_key, today)
        await log_command_usage(user_id, "AI個人分析")
        improve = compute_improve_rows(mine, everyone)
        chart = None
        if improve is not None:
            chart = await asyncio.to_thread(render_improve_chart, settings["my_deck"], improve)
        body = (
            f"🤖 **{settings['my_deck']}のAI分析**（あなたの記録{len(mine)}戦・同じデッキのほかの人{len(everyone)}戦と比較）\n"
            f"{text[:1400]}\n"
        )
        if chart is not None:
            body += "📉 **もう少し勝率を上げたい対面**を、下の画像にまとめました。\n"
        elif improve is None:
            body += f"📉 平均より勝率が低く、{IMPROVE_MIN_GAMES}戦以上当たっている対面は、今のところありません。\n"
        body += ENV_REPORT_AI_DISCLAIMER
        if chart is not None:
            await interaction.followup.send(
                body, file=discord.File(chart, filename="improve_matchups.png"), ephemeral=True
            )
        else:
            await interaction.followup.send(body, ephemeral=True)

    # ---- /弱点対面 ----
    # SENSEKI_MEMBERS_ONLY が true になるまで①②とも誰でも使える（コマンド自体の公開範囲）。
    # ②（全体平均との比較）は premium_access() で個別に判定：
    #   会員 / 今の環境が無料開放中(PREMIUM_FREE_ENV_IDS) / 初回登録から TRIAL_DAYS 日以内
    #   のいずれかを満たせば表示。無料開放する環境は PREMIUM_FREE_ENV_IDS で管理。
    async def _show_weak_matchups(self, interaction: discord.Interaction):
        """/戦績 弱点対面 とパネルのボタンの両方から呼ぶ"""
        if SENSEKI_MEMBERS_ONLY and not is_member(interaction.user):
            await interaction.response.send_message(
                "この機能はメンバー限定です。メンバーシップについては固定メッセージをご確認ください。",
                ephemeral=True,
            )
            return

        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return

        format_ = settings["format"]
        format_label = "ローテーション" if format_ == "rotation" else "アンリミテッド"
        my_class = settings["my_class"]
        my_deck = settings["my_deck"]

        deck_total = await get_deck_total(str(interaction.user.id), format_, my_deck)
        provisional = deck_total < DECK_RANK_MIN_MATCHES

        rows = await get_deck_matchups(str(interaction.user.id), format_, my_deck)
        candidates = [dict(r) for r in rows if r["total"] >= MIN_SAMPLE_PERSONAL]
        if not candidates:
            await interaction.response.send_message(
                f"**{my_deck}**（{format_label}）はまだ判定できるだけのデータがありません"
                f"（同じ対面が{MIN_SAMPLE_PERSONAL}戦以上たまると表示されます）。",
                ephemeral=True,
            )
            return

        for r in candidates:
            r["win_rate"] = r["wins"] / r["total"] * 100

        lines = [f"📉 **{my_deck}の弱点対面**（{format_label}・現環境）"]
        if provisional:
            lines.append(
                f"-# ⚠️ 参考値：このデッキはまだ{deck_total}戦です"
                f"（{DECK_RANK_MIN_MATCHES}戦を超えると精度が安定します）"
            )

        # ① 単純に勝率が低い対面（無料）
        raw_sorted = sorted(candidates, key=lambda r: r["win_rate"])
        raw_worst = raw_sorted[:3]
        lines.append("\n**① 勝率が低い対面**")
        for r in raw_worst:
            lines.append(
                f"・**{my_class}** vs **{r['opp_class']}**："
                f"勝率 {r['win_rate']:.1f}%（{r['wins']}勝{r['total']-r['wins']}敗・{r['total']}戦）"
            )
            link = await find_guide_link(my_class, r["opp_class"])
            if link:
                lines.append(f"　📖 {link['url']}" + (f"　{link['note']}" if link.get("note") else ""))

        # ② 全体平均より見劣りする対面（メンバー限定・環境限定無料／初回30日間無料キャンペーンあり）
        lines.append("\n**② 全体平均より見劣りする対面**")
        unlocked, reason, days_left = premium_access(
            interaction.user, settings.get("started_at"), await get_state(TRIAL_CAMPAIGN_START_KEY)
        )
        if not unlocked:
            lines.append(
                "-# この項目はメンバー限定です。自分の勝率は良くても、"
                "他プレイヤーの平均と比べると見劣りする対面が分かります。"
            )
        else:
            global_map = await get_global_deck_matchups(format_, my_deck)
            deviations = []
            for r in candidates:
                g = global_map.get(r["opp_class"])
                if not g or g["total"] < GLOBAL_MATCHUP_MIN_SAMPLE:
                    continue
                global_rate = g["wins"] / g["total"] * 100
                deviations.append({**r, "global_rate": global_rate, "diff": r["win_rate"] - global_rate})

            if not deviations:
                lines.append("-# 比較できるだけの全体データがまだ足りません。")
            else:
                deviations.sort(key=lambda r: r["diff"])
                for r in deviations[:3]:
                    lines.append(
                        f"・**{my_class}** vs **{r['opp_class']}**："
                        f"あなた {r['win_rate']:.1f}%（全体平均 {r['global_rate']:.1f}% / "
                        f"乖離 {r['diff']:+.1f}%）"
                    )
                    link = await find_guide_link(my_class, r["opp_class"])
                    if link:
                        lines.append(f"　📖 {link['url']}" + (f"　{link['note']}" if link.get("note") else ""))
                await log_command_usage(str(interaction.user.id), "弱点対面_全体比較")

            if reason == "trial" and days_left <= 7:
                lines.append(f"\n🎁 無料キャンペーン期間中（あと{days_left}日で終了）")

        lines.append(
            f"\n-# 対面ごとに{MIN_SAMPLE_PERSONAL}戦、全体比較は全体側{GLOBAL_MATCHUP_MIN_SAMPLE}戦以上のものだけ表示しています。"
        )

        # グラフは①（無料枠）の全対面を対象にする。フォント未導入時などはNoneが返り、
        # そのままテキストのみで送信される（グラフなしでも壊れない）。
        chart_buf = await self._render_weak_matchups_chart(
            my_deck, format_label, candidates, deck_total=deck_total, min_matches=DECK_RANK_MIN_MATCHES
        )
        file_ = discord.File(chart_buf, filename="weak_matchups.png") if chart_buf else None

        await log_command_usage(str(interaction.user.id), "弱点対面")
        await interaction.response.send_message(
            "\n".join(lines)[:1900], file=file_, ephemeral=True
        )

    @senseki.command(name="弱点対面", description="対面別の勝率が低い相手と、おすすめの攻略記事を表示します")
    async def weak_matchups(self, interaction: discord.Interaction):
        await self._show_weak_matchups(interaction)

    # ---- /戦績 デッキ推移 ----
    @senseki.command(name="デッキ推移", description="今のフォーマットで、これまで使ってきたデッキの変遷と成績を表示します")
    async def deck_history_cmd(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None:
            await interaction.response.send_message(
                "まず `/戦績 設定` で登録してください。", ephemeral=True
            )
            return
        format_ = settings["format"]
        format_label = "ローテーション" if format_ == "rotation" else "アンリミテッド"

        segments = await get_deck_history(str(interaction.user.id), format_)
        if not segments:
            await interaction.response.send_message(
                f"{format_label}での記録がまだありません。", ephemeral=True
            )
            return

        lines = [f"📈 **デッキ推移**（{format_label}・新しい順）"]
        for seg in segments[:10]:
            win_rate = seg["wins"] / seg["total"] * 100 if seg["total"] else 0
            start = seg["start"][:10]
            end = seg["end"][:10]
            period = start if start == end else f"{start}〜{end}"
            lines.append(
                f"・**{seg['deck']}**：{seg['wins']}勝{seg['total']-seg['wins']}敗"
                f"（{seg['total']}戦・勝率{win_rate:.1f}%）　{period}"
            )
        if len(segments) > 10:
            lines.append(f"\n-# 直近10区間のみ表示しています（全{len(segments)}区間）。")

        await log_command_usage(str(interaction.user.id), "デッキ推移")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    # ---- /戦績 先後対面 ----
    @senseki.command(name="先後対面", description="今のデッキで、対面ごとに先攻後攻どちらが得意かを表示します")
    async def first_second_cmd(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return
        format_ = settings["format"]
        format_label = "ローテーション" if format_ == "rotation" else "アンリミテッド"
        my_deck = settings["my_deck"]

        rows = await get_first_second_matchups(str(interaction.user.id), format_, my_deck)
        by_opp = {}
        for r in rows:
            d = by_opp.setdefault(r["opp_class"], {})
            d["first" if r["is_first"] else "second"] = r

        results = []
        for opp_class, d in by_opp.items():
            first = d.get("first")
            second = d.get("second")
            if not first or not second:
                continue
            if first["total"] < MIN_SAMPLE_PERSONAL or second["total"] < MIN_SAMPLE_PERSONAL:
                continue
            first_rate = first["wins"] / first["total"] * 100
            second_rate = second["wins"] / second["total"] * 100
            results.append({
                "opp_class": opp_class, "first_rate": first_rate, "second_rate": second_rate,
                "first_total": first["total"], "second_total": second["total"],
                "diff": first_rate - second_rate,
            })

        if not results:
            await interaction.response.send_message(
                f"**{my_deck}**（{format_label}）は、先攻・後攻それぞれ{MIN_SAMPLE_PERSONAL}戦以上ある"
                f"対面がまだありません。", ephemeral=True,
            )
            return

        results.sort(key=lambda r: abs(r["diff"]), reverse=True)
        lines = [f"🔀 **{my_deck}の先攻・後攻別対面**（{format_label}・現環境）"]
        for r in results[:5]:
            favor = "先攻" if r["diff"] > 0 else "後攻"
            lines.append(
                f"・**{r['opp_class']}** 戦：先攻 {r['first_rate']:.1f}%"
                f"（{r['first_total']}戦） / 後攻 {r['second_rate']:.1f}%（{r['second_total']}戦）"
                f"　→ {favor}が得意（差 {abs(r['diff']):.1f}%）"
            )
        lines.append(f"\n-# 先攻・後攻それぞれ{MIN_SAMPLE_PERSONAL}戦以上ある対面のみ表示しています。")

        await log_command_usage(str(interaction.user.id), "先後対面")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    # ---- /戦績 相手デッキ対面 ----
    @senseki.command(name="相手デッキ対面", description="今のデッキで、相手デッキ別の勝率を表示します（データが十分あれば先攻後攻も）")
    async def opp_deck_matchup_cmd(self, interaction: discord.Interaction):
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return
        format_ = settings["format"]
        format_label = "ローテーション" if format_ == "rotation" else "アンリミテッド"
        my_deck = settings["my_deck"]

        rows = await get_deck_opp_deck_fs(str(interaction.user.id), format_, my_deck)
        by_key = {}
        for r in rows:
            key = (r["opp_class"], r["opp_deck"])
            d = by_key.setdefault(key, {
                "total": 0, "wins": 0,
                "first_total": 0, "first_wins": 0,
                "second_total": 0, "second_wins": 0,
            })
            d["total"] += r["total"]
            d["wins"] += r["wins"]
            if r["is_first"]:
                d["first_total"] += r["total"]
                d["first_wins"] += r["wins"]
            else:
                d["second_total"] += r["total"]
                d["second_wins"] += r["wins"]

        entries = [
            {"opp_class": k[0], "opp_deck": k[1], **v}
            for k, v in by_key.items() if v["total"] >= MIN_SAMPLE_OPPDECK
        ]
        if not entries:
            await interaction.response.send_message(
                f"**{my_deck}**（{format_label}）は、相手デッキが分かる対面が"
                f"まだ{MIN_SAMPLE_OPPDECK}戦以上たまっていません。\n"
                f"-# 対戦相手のデッキ名を入力して記録すると、ここに表示されるようになります。",
                ephemeral=True,
            )
            return

        entries.sort(key=lambda r: r["total"], reverse=True)
        lines = [f"🎯 **{my_deck}の相手デッキ別対面**（{format_label}・現環境）"]
        for r in entries[:8]:
            rate = r["wins"] / r["total"] * 100
            lines.append(
                f"・**{r['opp_class']}**（{r['opp_deck']}）：勝率 {rate:.1f}%"
                f"（{r['wins']}勝{r['total']-r['wins']}敗・{r['total']}戦）"
            )
            if r["first_total"] >= MIN_SAMPLE_OPPDECK_FS and r["second_total"] >= MIN_SAMPLE_OPPDECK_FS:
                first_rate = r["first_wins"] / r["first_total"] * 100
                second_rate = r["second_wins"] / r["second_total"] * 100
                lines.append(
                    f"　先攻 {first_rate:.1f}%（{r['first_total']}戦） / "
                    f"後攻 {second_rate:.1f}%（{r['second_total']}戦）"
                )
            else:
                lines.append(f"　-# 先攻後攻は各{MIN_SAMPLE_OPPDECK_FS}戦以上貯まると内訳が出ます")

        if len(entries) > 8:
            lines.append(f"\n-# 対戦数の多い順に8件のみ表示しています（全{len(entries)}件）。")
        lines.append(f"-# 相手デッキは{MIN_SAMPLE_OPPDECK}戦以上あるものだけ表示しています。")

        await log_command_usage(str(interaction.user.id), "相手デッキ対面")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    # ---- 攻略記事の管理（管理者専用） ----
    @admin.command(name="記事登録", description="対面ごとのおすすめ攻略記事を登録します")
    @app_commands.describe(
        相手クラス="この相手クラスへの対策記事",
        url="記事のURL",
        自分のクラス="任意：自分のクラスを指定すると、その組み合わせでのみ表示されます（省略時は相手クラス共通の記事）",
        メモ="任意：一言コメント",
    )
    @app_commands.choices(相手クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(自分のクラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    async def register_guide_link(
        self,
        interaction: discord.Interaction,
        相手クラス: app_commands.Choice[str],
        url: str,
        自分のクラス: app_commands.Choice[str] = None,
        メモ: str = None,
    ):
        my_class = 自分のクラス.value if 自分のクラス else ""
        await upsert_guide_link(my_class, 相手クラス.value, url, メモ)
        scope = f"{自分のクラス.name} vs {相手クラス.name}" if 自分のクラス else f"vs {相手クラス.name}（汎用）"
        await interaction.response.send_message(f"✅ 登録しました：{scope}\n{url}", ephemeral=True)

    @admin.command(name="記事一覧", description="登録済みの攻略記事を一覧表示します")
    async def list_guide_links_cmd(self, interaction: discord.Interaction):
        links = await list_guide_links()
        if not links:
            await interaction.response.send_message("まだ登録がありません。", ephemeral=True)
            return
        lines = ["📚 **登録済み攻略記事**"]
        for l in links:
            scope = f"{l['my_class']} vs {l['opp_class']}" if l["my_class"] else f"vs {l['opp_class']}（汎用）"
            lines.append(f"・{scope}：{l['url']}")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    @admin.command(name="記事削除", description="登録済みの攻略記事を削除します")
    @app_commands.describe(
        相手クラス="削除する記事の相手クラス",
        自分のクラス="任意：クラス指定ありで登録した記事を消す場合のみ指定",
    )
    @app_commands.choices(相手クラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    @app_commands.choices(自分のクラス=[app_commands.Choice(name=c, value=c) for c in CLASS_CHOICES])
    async def delete_guide_link_cmd(
        self,
        interaction: discord.Interaction,
        相手クラス: app_commands.Choice[str],
        自分のクラス: app_commands.Choice[str] = None,
    ):
        my_class = 自分のクラス.value if 自分のクラス else ""
        ok = await delete_guide_link(my_class, 相手クラス.value)
        msg = "🗑️ 削除しました。" if ok else "該当する記事が見つかりませんでした。"
        await interaction.response.send_message(msg, ephemeral=True)

    # ---- /戦績確認 ----
    @senseki.command(name="確認", description="今使っているデッキ・ランクと、現環境での自分の勝率を確認します")
    async def senseki_check(self, interaction: discord.Interaction):
        await self._show_check(interaction)

    async def _show_check(self, interaction: discord.Interaction):
        """/戦績 確認 とパネルのボタンの両方から呼ぶ"""
        settings = await get_user_settings(str(interaction.user.id))
        if settings is None or not settings.get("my_deck"):
            await interaction.response.send_message(
                "まず `/戦績 設定` で使用デッキを登録してください。", ephemeral=True
            )
            return

        format_ = settings["format"]
        format_label = "ローテーション" if format_ == "rotation" else "アンリミテッド"
        my_deck = settings["my_deck"]
        rank = settings.get("rank_tier") or "ランク未設定"
        if settings.get("cr_grade"):
            rank += f"・{settings['cr_grade']}"
        elif settings.get("rank_tier") == GRAND_MASTER_TIER:
            rank += "・グレードなし"

        lines = [
            f"🧑 現在の設定：**{my_deck}**（{settings['my_class']} / {format_label}） / ランク：**{rank}**",
            "",
        ]

        full = await get_deck_summary(str(interaction.user.id), format_, my_deck)
        if full["total"] == 0:
            lines.append("このデッキではまだ記録がありません。パネルの「⚔️ 戦績を記録する」から記録してみてください。")
            await interaction.response.send_message("\n".join(lines), ephemeral=True)
            return

        recent = await get_deck_recent_summary(str(interaction.user.id), format_, my_deck, RECENT_FREE_MATCHES)
        recent_rate = recent["wins"] / recent["total"] * 100 if recent["total"] else 0
        lines.append(
            f"📊 直近{recent['total']}戦の勝率：{recent_rate:.1f}%（{recent['wins']}勝{recent['total']-recent['wins']}敗）"
        )

        if full["total"] > RECENT_FREE_MATCHES:
            unlocked, reason, days_left = premium_access(
                interaction.user, settings.get("started_at"), await get_state(TRIAL_CAMPAIGN_START_KEY)
            )
            if unlocked:
                full_rate = full["wins"] / full["total"] * 100
                lines.append(
                    f"通算{full['total']}戦：勝率 {full_rate:.1f}%（{full['wins']}勝{full['total']-full['wins']}敗）"
                )
                if reason == "trial" and days_left <= 7:
                    lines.append(f"🎁 無料キャンペーン期間中（あと{days_left}日で終了）")
            else:
                lines.append(f"通算{full['total']}戦（勝率は🔒メンバーシップ限定）")


        overall = await get_summary(str(interaction.user.id))
        if overall["total"] > full["total"]:
            o_rate = overall["wins"] / overall["total"] * 100
            lines.append(
                f"-# 全デッキ通算：{overall['total']}戦"
                f"（{overall['wins']}勝{overall['losses']}敗・勝率{o_rate:.1f}%）"
            )

        await log_command_usage(str(interaction.user.id), "確認")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)

    # ---- /戦績 履歴 ----
    @senseki.command(name="履歴", description="直近の対戦履歴を表示します（自分の分のみ）")
    async def recent_history_cmd(self, interaction: discord.Interaction):
        rows = await get_recent_matches(str(interaction.user.id), 10)
        if not rows:
            await interaction.response.send_message(
                "まだ記録がありません。パネルの「⚔️ 戦績を記録する」から記録してみてください。",
                ephemeral=True,
            )
            return

        lines = ["🕘 **直近の対戦履歴**（新しい順・最大10戦）"]
        for r in rows:
            result = "🟢勝ち" if r["is_win"] else "🔴負け"
            first_second = "先攻" if r["is_first"] else "後攻"
            fmt_label = "ローテ" if r["format"] == "rotation" else "アンリミ"
            deck = r["my_deck"] or "デッキ未設定"
            opp = r["opp_class"] + (f"（{r['opp_deck']}）" if r["opp_deck"] else "")
            date = (r["recorded_at"] or "")[:16].replace("T", " ")
            lines.append(f"・{date}　[{fmt_label}] {deck} vs {opp}　{first_second}　{result}")

        await log_command_usage(str(interaction.user.id), "履歴")
        await interaction.response.send_message("\n".join(lines)[:1900], ephemeral=True)


async def setup(bot: commands.Bot):
    await bot.add_cog(SensekiCog(bot), guild=discord.Object(id=GUILD_ID))


# ==============================
# main.py への統合方法
# ==============================
#
#   from senseki import SensekiCog
#
#   # on_ready 内、他のCog登録と同じ場所に追加：
#   if bot.get_cog("SensekiCog") is None:
#       await bot.add_cog(SensekiCog(bot), guild=guild_obj)
#
# 必要な環境変数（任意）:
#   SENSEKI_DB_PATH … 既定値 /data/senseki.db。ボリュームのマウントパスが
#                      違う場合のみ設定する
#
# 必要なライブラリ（requirements.txt に追加）:
#   openpyxl
#   gspread        … スプレッドシート連携を使う場合のみ
#   google-auth    … 同上
#
# デプロイ前提条件:
#   Railway の shirokusa-bot サービスに永続ボリューム（/data）が必要。
#   未設定の場合は Volumes タブから追加してから反映すること。
#
# SENSEKI_ADMIN_USER_ID は設定済み（週次抽出のDM送付先）。
