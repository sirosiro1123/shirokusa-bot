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
                通常はパネルの「🧪 デッキ検証」ボタンから使う

  /戦績管理（管理者のみ表示）
    パネル設置    記録・切替・ランク更新・取消・成績確認・弱点対面・デッキ検証の各ボタンを常設
    掲示板設置    全員の使用デッキ＆ランクの一覧を常設（設定変更時に自動更新）
    全体集計      その瞬間のDBから全体・先後別・クラス別・対面別を表示
    環境分析画像  クラス使用率・対面マトリクス・先攻後攻をPNG画像で表示（SENSEKI_ADMIN_USER_ID本人のみ）
    データ抽出    生データ＋集計のExcelを出力（毎週月曜9:00にDMへ自動送信）
    シート同期    Googleスプレッドシートへ即時同期
    シート診断    スプレッドシート連携の設定状況
    診断          DBの保存先と永続化状況（データが消える事故の確認用）
    権限確認      指定チャンネルでBOTが持っている権限を確認
    分析利用状況  確認・弱点対面などの分析コマンドの利用回数・利用者数
    デッキ名統合  表記ゆれのあるデッキ名を1つに統合（過去の記録も一括書き換え）
    無料トライアル開始／無料トライアル状況  初回参加者向け無料キャンペーンの開始とON後の状況確認
    記事登録／記事一覧／記事削除  対面ごとのおすすめ記事URLの管理

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
  - 環境ID管理と /環境切替（現状は CURRENT_ENV_ID を固定値で使用）
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
import io
import os
import sqlite3
from datetime import datetime, time as dtime, timedelta, timezone

import discord
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

# 環境ID（暫定固定値）
# フェーズ2で /環境切替 と data/environment.json による管理に置き換える。
# それまでは記録される全レコードがこの環境IDになる。
# 新弾リリース（2026-08-27）に合わせて設定。新弾切り替わり後に変更する場合は
# ここを書き換えてから再デプロイすること。
CURRENT_ENV_ID = "2026-08-beyond-2"

# 集計データの定期抽出先（管理者DM）
# しろくさのユーザーID（koken_rank.py の ADOPT_JUDGE_USER_ID と同じ値）
SENSEKI_ADMIN_USER_ID = 346475730458378240

# 定期抽出のタイミング（JST）
WEEKLY_EXPORT_WEEKDAY = 0   # 0=月曜
WEEKLY_EXPORT_HOUR = 9
WEEKLY_EXPORT_MINUTE = 0

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

# 相手デッキの選択画面を出し始める、候補デッキ名の最低件数
# 0にすると候補ゼロでも画面が出る（「分からない」しか押せない状態になる）。
# 1にしておくと、そのクラスのデッキ名が1つでも登録された時点で選択画面が出る。
# 候補は各ユーザーが初回設定で登録した自分のデッキから自動的に貯まる。
MIN_DECK_CANDIDATES = 1

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

# ②（全体平均との比較）を、有料化前に環境限定で無料開放するスイッチ。
# 今回の環境（2026-08-beyond-2）は無料開放。次の環境に切り替える際は
# 必ず False に戻すこと（CURRENT_ENV_ID の書き換えと同時に見直す）。
PREMIUM_FREE_THIS_ENV = True

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
    if PREMIUM_FREE_THIS_ENV:
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
        (class_name, deck_name, CURRENT_ENV_ID),
    ).fetchone()
    if row is None:
        conn.execute(
            "INSERT INTO deck_names (class_name, deck_name, use_count, is_official, env_id) "
            "VALUES (?, ?, 1, 0, ?)",
            (class_name, deck_name, CURRENT_ENV_ID),
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
            user_id, now, CURRENT_ENV_ID, settings["format"], settings["my_class"],
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
        """, (class_name, CURRENT_ENV_ID, limit)).fetchall()
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
            (class_name, old_name, CURRENT_ENV_ID),
        ).fetchone()
        new_row = conn.execute(
            "SELECT id, use_count FROM deck_names WHERE class_name = ? AND deck_name = ? AND env_id = ?",
            (class_name, new_name, CURRENT_ENV_ID),
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
                (class_name, new_name, CURRENT_ENV_ID),
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


def _delete_last_match_sync(user_id: str):
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT id, opp_class, is_win, recorded_at FROM matches "
            "WHERE user_id = ? ORDER BY id DESC LIMIT 1",
            (user_id,),
        ).fetchone()
        if row is None:
            return None
        conn.execute("DELETE FROM matches WHERE id = ?", (row["id"],))
        conn.execute("DELETE FROM personal_deck_variants WHERE match_id = ?", (row["id"],))
        conn.commit()
        return dict(row)
    finally:
        conn.close()


async def delete_last_match(user_id: str):
    result = await asyncio.to_thread(_delete_last_match_sync, user_id)
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
        """, (user_id, CURRENT_ENV_ID)).fetchone()
        total = row["total"] or 0
        wins = row["wins"] or 0
        return {"total": total, "wins": wins, "losses": total - wins}
    finally:
        conn.close()


async def get_summary(user_id: str):
    return await asyncio.to_thread(_get_summary_sync, user_id)


def _fetch_all_matches_sync(env_only: bool):
    conn = _connect()
    try:
        if env_only:
            rows = conn.execute(
                "SELECT * FROM matches WHERE env_id = ? ORDER BY id", (CURRENT_ENV_ID,)
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM matches ORDER BY id").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def fetch_all_matches(env_only: bool = True):
    return await asyncio.to_thread(_fetch_all_matches_sync, env_only)


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

def _get_personal_matchups_sync(user_id: str):
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT my_class, opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE user_id = ? AND env_id = ?
            GROUP BY my_class, opp_class
        """, (user_id, CURRENT_ENV_ID)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


async def get_personal_matchups(user_id: str):
    return await asyncio.to_thread(_get_personal_matchups_sync, user_id)


def _get_deck_total_sync(user_id: str, format_: str, my_deck) -> int:
    """指定デッキ・フォーマットでの、そのユーザーの通算試合数（現環境）"""
    conn = _connect()
    try:
        row = conn.execute("""
            SELECT COUNT(*) AS total FROM matches
            WHERE user_id = ? AND env_id = ? AND format = ? AND my_deck = ?
        """, (user_id, CURRENT_ENV_ID, format_, my_deck)).fetchone()
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
        """, (user_id, CURRENT_ENV_ID, format_, my_deck)).fetchone()
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
        """, (user_id, CURRENT_ENV_ID, format_, my_deck, limit)).fetchall()
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
        """, (user_id, CURRENT_ENV_ID, limit)).fetchall()
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
        """, (user_id, CURRENT_ENV_ID, format_, my_deck)).fetchall()
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
        """, (CURRENT_ENV_ID, format_, my_deck)).fetchall()
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
        """, (user_id, CURRENT_ENV_ID, format_)).fetchall()
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
        """, (user_id, CURRENT_ENV_ID, format_, my_deck)).fetchall()
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
        """, (user_id, CURRENT_ENV_ID, format_, my_deck)).fetchall()
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
            "SELECT COUNT(DISTINCT user_id) AS n FROM matches WHERE env_id = ?", (CURRENT_ENV_ID,)
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
        """, (CURRENT_ENV_ID,)).fetchone()

        by_first = conn.execute("""
            SELECT is_first, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ? GROUP BY is_first
        """, (CURRENT_ENV_ID,)).fetchall()

        by_my_class = conn.execute("""
            SELECT my_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ?
            GROUP BY my_class ORDER BY total DESC
        """, (CURRENT_ENV_ID,)).fetchall()

        by_opp_class = conn.execute("""
            SELECT opp_class, COUNT(*) AS total, SUM(is_win) AS wins
            FROM matches WHERE env_id = ?
            GROUP BY opp_class ORDER BY total DESC
        """, (CURRENT_ENV_ID,)).fetchall()

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
    return await asyncio.to_thread(_get_class_matchup_matrix_sync, env_id or CURRENT_ENV_ID)


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
# /戦績 のボタン形式フロー
# ==============================

class OpponentClassSelect(discord.ui.Select):
    def __init__(self):
        options = [discord.SelectOption(label=c, value=c) for c in CLASS_CHOICES]
        super().__init__(placeholder="相手クラスを選択", options=options, min_values=1, max_values=1)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.opp_class = self.values[0]

        # そのクラスで既に使われているデッキ名を候補に出す。
        # 候補が貯まるまでは相手デッキの画面自体を出さない。
        # 運用初期に「分からない」を押させ続けるのは手間なだけで、
        # データも増えないため。候補は各自が登録した自分のデッキから自動的に貯まる。
        known = await get_deck_names(view.opp_class)
        if len(known) < MIN_DECK_CANDIDATES:
            view.opp_deck = None
            view.show_first_second()
            await interaction.response.edit_message(
                content=view.screen("先攻／後攻を選んでください。"),
                view=view,
            )
            return

        view.show_opponent_deck(known)
        await interaction.response.edit_message(
            content=view.screen(
                "相手のデッキが分かれば選んでください。分からなければ「分からない」でOKです。"
            ),
            view=view,
        )


class OpponentDeckSelect(discord.ui.Select):
    """既知のデッキ名から選ぶ。候補になければモーダルで新規入力"""

    SKIP_VALUE = "__skip__"
    NEW_VALUE = "__new__"

    def __init__(self, known_decks: list[str]):
        options = [
            discord.SelectOption(label="分からない / 入力しない", value=self.SKIP_VALUE),
            discord.SelectOption(label="新しく入力する…", value=self.NEW_VALUE),
        ]
        # Discordのセレクトは25件が上限。既定の2件を除いた分だけ載せる
        for name in known_decks[:23]:
            options.append(discord.SelectOption(label=name[:100], value=name[:100]))
        super().__init__(
            placeholder="相手のデッキ（任意）", options=options, min_values=1, max_values=1
        )

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        value = self.values[0]

        if value == self.NEW_VALUE:
            await interaction.response.send_modal(OpponentDeckModal(view))
            return

        view.opp_deck = None if value == self.SKIP_VALUE else value
        view.show_first_second()
        await interaction.response.edit_message(
            content=view.screen("先攻／後攻を選んでください。"),
            view=view,
        )


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
        view.show_first_second()
        await interaction.response.edit_message(
            content=view.screen("先攻／後攻を選んでください。"),
            view=view,
        )


class FirstButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="先攻", style=discord.ButtonStyle.primary)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.is_first = True
        view.show_win_lose()
        await interaction.response.edit_message(
            content=view.screen("結果を選んでください。"),
            view=view,
        )


class SecondButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="後攻", style=discord.ButtonStyle.secondary)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.is_first = False
        view.show_win_lose()
        await interaction.response.edit_message(
            content=view.screen("結果を選んでください。"),
            view=view,
        )


class WinButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="勝ち", style=discord.ButtonStyle.success)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        await view.finalize(interaction, is_win=True)


class LoseButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="負け", style=discord.ButtonStyle.danger)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        await view.finalize(interaction, is_win=False)


class RepeatButton(discord.ui.Button):
    def __init__(self):
        super().__init__(label="同じ設定でもう1試合", style=discord.ButtonStyle.primary)

    async def callback(self, interaction: discord.Interaction):
        view: SensekiFlowView = self.view
        view.reset()
        await interaction.response.edit_message(
            content=view.screen("相手クラスを選択してください。"),
            view=view,
        )


class SensekiFlowView(discord.ui.View):
    def __init__(self, user_id: str, settings: dict, origin_interaction: discord.Interaction):
        # 13分。実際の対戦(5〜15分程度)をまたいで「同じ設定でもう1試合」を押せるように、
        # 以前の5分から延長。Discordのフォローアップ編集が可能な15分に収まる範囲で設定。
        super().__init__(timeout=780)
        self.user_id = user_id
        self.settings = settings
        self.origin_interaction = origin_interaction
        self.opp_class = None
        self.opp_deck = None
        self.is_first = None
        self.clear_items()
        self.add_item(OpponentClassSelect())

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
        return f"🧑 あなた：**{deck}**（{s['my_class']} / {fmt}） / ランク：**{rank}**"

    def screen(self, body: str) -> str:
        """設定ヘッダー＋進捗＋操作案内をまとめた画面テキスト"""
        parts = [self.describe_settings()]
        progress = self.describe_progress()
        if progress:
            parts.append(progress)
        parts.append("")
        parts.append(body)
        return "\n".join(parts)

    def describe_progress(self) -> str:
        """ここまでに選んだ内容を1行で"""
        if self.opp_class is None:
            return ""
        label = f"🆚 相手：**{self.opp_class}**"
        if self.opp_deck:
            label += f"（{self.opp_deck}）"
        if self.is_first is not None:
            label += f" / **{'先攻' if self.is_first else '後攻'}**"
        return label

    def show_opponent_deck(self, known_decks: list[str]):
        self.clear_items()
        self.add_item(OpponentDeckSelect(known_decks))

    def show_first_second(self):
        self.clear_items()
        self.add_item(FirstButton())
        self.add_item(SecondButton())

    def show_win_lose(self):
        self.clear_items()
        self.add_item(WinButton())
        self.add_item(LoseButton())

    def reset(self):
        self.opp_class = None
        self.opp_deck = None
        self.is_first = None
        self.clear_items()
        self.add_item(OpponentClassSelect())

    async def finalize(self, interaction: discord.Interaction, is_win: bool):
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

        result_label = "勝ち" if is_win else "負け"
        self.clear_items()
        self.add_item(RepeatButton())
        tag_note = f"\n🏷 検証タグ「{variant_label}」を付けて記録しました。" if variant_label else ""
        await interaction.response.edit_message(
            content=self.screen(
                f"✅ **{result_label}** で記録しました。{tag_note}\n"
                f"続けて記録する場合は下のボタンを押してください。"
            ),
            view=self,
        )

    async def on_timeout(self):
        # 元のインタラクションのフォローアップ編集は15分まで有効。
        # timeout=780(13分)なので基本は成功するが、念のためHTTPExceptionは握りつぶす。
        try:
            await self.origin_interaction.edit_original_response(
                content=(
                    "⏱ 入力の受付時間が終了しました。\n"
                    "お手数ですが、パネルの「⚔️ 戦績を記録する」からもう一度お願いします。"
                ),
                view=None,
            )
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
                "デッキやランクを変えるだけなら「🔄 デッキ切替」「📈 ランク更新」が早いです。",
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


class PanelCheckButton(discord.ui.Button):
    def __init__(self):
        super().__init__(
            label="成績確認",
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
            f"やめる場合はもう一度「🧪 デッキ検証」→「タグを解除」を押してください。",
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
                "まず「🆕 初回設定」か「🔄 デッキ切替」でデッキを登録してください。",
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


class SensekiPanelView(discord.ui.View):
    """timeout=None＋固定custom_idで永続化。BOT再起動後もボタンが反応し続ける"""

    def __init__(self):
        super().__init__(timeout=None)
        # 1段目：記録まわり（よく使う順）
        self.add_item(SensekiPanelButton())
        self.add_item(PanelUndoButton())
        self.add_item(PanelCheckButton())
        self.add_item(PanelWeakButton())
        # 2段目：設定まわり
        self.add_item(PanelSetupButton())
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
# Cog本体
# ==============================

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
        if not self.daily_sheets_sync.is_running():
            self.daily_sheets_sync.start()
        if not self.debounced_sheets_sync.is_running():
            self.debounced_sheets_sync.start()
        # 永続ビューの登録。custom_idが一致すれば、BOT再起動後も
        # 過去に送信済みのパネルのボタンが反応し続ける
        self.bot.add_view(SensekiPanelView())
        # BOT再起動をまたいでも掲示板が最新になるよう一度更新しておく
        try:
            await self._update_board()
        except Exception as e:
            print(f"⚠️ 起動時の戦績掲示板更新に失敗しました: {type(e).__name__}: {e}")

    def cog_unload(self):
        self.weekly_export.cancel()
        self.daily_sheets_sync.cancel()
        self.debounced_sheets_sync.cancel()

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
                    "・📊 成績確認　・📉 弱点対面\n"
                    "\n"
                    "**設定**\n"
                    "・🆕 初回設定　・🔄 デッキ切替（新規登録もここから）\n"
                    "・📈 ランク更新　・🔀 フォーマット切替\n"
                    "\n"
                    "**⚠️ 記録の前に確認してください**\n"
                    "記録画面の上部に、今のあなたのデッキとランクが表示されます。\n"
                    "**違っていたら「🔄 デッキ切替」「📈 ランク更新」で直してから記録してください。**\n"
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
        guild = self.bot.get_guild(GUILD_ID)
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
            matchup_rows, CLASS_CHOICES, MATCHUP_MATRIX_MIN_SAMPLE, CURRENT_ENV_ID,
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
            + ("引き続き全員無料のままです。" if PREMIUM_FREE_THIS_ENV else "既に有料化されています。")
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
            f"（現在の環境は" + ("全員無料のため、トライアルの有無に関わらず②が使えます。" if PREMIUM_FREE_THIS_ENV else "有料化済みです。") + "）",
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

    async def _build_export(self, env_only: bool) -> tuple[bytes, str, int]:
        matches = await fetch_all_matches(env_only=env_only)
        guild = self.bot.get_guild(GUILD_ID)
        data = await asyncio.to_thread(build_workbook, matches, guild)
        scope = "現環境" if env_only else "全期間"
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

    # ---- /戦績データ抽出 ----
    @admin.command(name="データ抽出", description="戦績データを生データ＋集計のExcelで出力します")
    @app_commands.describe(全期間="オンにすると環境をまたいだ全データを出力します（既定は現環境のみ）")
    async def senseki_export(self, interaction: discord.Interaction, 全期間: bool = False):
        if Workbook is None:
            await interaction.response.send_message(
                f"openpyxlが読み込めませんでした。\n```\n{OPENPYXL_IMPORT_ERROR}\n```\n"
                "requirements.txt に `openpyxl` を追加してください。",
                ephemeral=True,
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        try:
            data, filename, count = await self._build_export(env_only=not 全期間)
            if count == 0:
                await interaction.followup.send("記録がまだありません。", ephemeral=True)
                return

            summary = f"📊 戦績データ抽出（{'全期間' if 全期間 else '現環境'}・{count}件）"
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

        view = SensekiFlowView(str(interaction.user.id), settings, interaction)
        await interaction.response.send_message(
            view.screen(
                "相手クラスを選択してください。\n"
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

    def _render_class_summary_image_sync(self, class_summary: list, env_id: str):
        """クラス使用率（円）＋クラス別勝率（横棒）を1枚のPNGにする"""
        if plt is None or _JP_FONT_NAME is None or not class_summary:
            return None
        try:
            rows = sorted(class_summary, key=lambda r: r["total"], reverse=True)
            labels = [r["my_class"] for r in rows]
            totals = [r["total"] for r in rows]
            win_rates = [
                (r["wins"] or 0) / r["total"] * 100 if r["total"] else 0.0 for r in rows
            ]

            norm = mcolors.Normalize(vmin=0, vmax=100)
            cmap = matplotlib.colormaps["RdYlGn"]

            fig, (ax_pie, ax_bar) = plt.subplots(
                2, 1, figsize=(7.4, 9.6), dpi=150,
                gridspec_kw={"height_ratios": [1, 1.1]},
            )

            ax_pie.pie(totals, labels=labels, autopct="%1.1f%%", startangle=90,
                       textprops={"fontsize": 11})
            ax_pie.set_title(f"クラス使用率（現環境：{env_id}）", fontsize=14, pad=12)

            # 勝率バーは勝率の低い順（悪い順）に並べ、悪目立ちする対面を上に来させる
            bar_order = sorted(range(len(rows)), key=lambda i: win_rates[i], reverse=True)
            bar_labels = [labels[i] for i in bar_order]
            bar_values = [win_rates[i] for i in bar_order]
            bar_totals = [totals[i] for i in bar_order]
            colors = [cmap(norm(v)) for v in bar_values]

            y = range(len(bar_labels))
            ax_bar.barh(y, bar_values, color=colors, edgecolor="#333333", linewidth=0.5, height=0.65)
            ax_bar.set_yticks(list(y))
            ax_bar.set_yticklabels(bar_labels, fontsize=12)
            ax_bar.set_xlim(0, 118)
            ax_bar.set_xticks([0, 20, 40, 60, 80, 100])
            ax_bar.axvline(50, color="#888888", linestyle="--", linewidth=1)
            ax_bar.set_xlabel("勝率（%）", fontsize=11)
            for yi, (v, t) in enumerate(zip(bar_values, bar_totals)):
                ax_bar.text(v + 2, yi, f"{v:.1f}%（{t}戦）", va="center", ha="left", fontsize=10)
            ax_bar.set_title(f"クラス別勝率（現環境：{env_id}）", fontsize=14, pad=12)
            ax_bar.spines["top"].set_visible(False)
            ax_bar.spines["right"].set_visible(False)

            fig.tight_layout()
            buf = io.BytesIO()
            fig.savefig(buf, format="png")
            plt.close(fig)
            buf.seek(0)
            return buf
        except Exception:
            return None

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

    def _render_first_second_image_sync(self, first_second_summary: list, env_id: str):
        """先攻/後攻の勝率を横棒2本のPNGにする"""
        if plt is None or _JP_FONT_NAME is None or not first_second_summary:
            return None
        try:
            by_first = {r["is_first"]: r for r in first_second_summary}
            order = [(1, "先攻"), (0, "後攻")]
            labels, values, totals = [], [], []
            for key, label in order:
                r = by_first.get(key)
                if r is None:
                    continue
                labels.append(label)
                totals.append(r["total"])
                values.append((r["wins"] or 0) / r["total"] * 100 if r["total"] else 0.0)

            norm = mcolors.Normalize(vmin=0, vmax=100)
            cmap = matplotlib.colormaps["RdYlGn"]
            colors = [cmap(norm(v)) for v in values]

            fig, ax = plt.subplots(figsize=(5.6, 3.6), dpi=150)
            y = range(len(labels))
            ax.barh(y, values, color=colors, edgecolor="#333333", linewidth=0.5, height=0.55)
            ax.set_yticks(list(y))
            ax.set_yticklabels(labels, fontsize=12)
            ax.set_xlim(0, 118)
            ax.set_xticks([0, 20, 40, 60, 80, 100])
            ax.axvline(50, color="#888888", linestyle="--", linewidth=1)
            ax.set_xlabel("勝率（%）", fontsize=11)
            for yi, (v, t) in enumerate(zip(values, totals)):
                ax.text(v + 2, yi, f"{v:.1f}%（{t}戦）", va="center", ha="left", fontsize=10)
            ax.set_title(f"先攻/後攻の勝率（現環境：{env_id}）", fontsize=13, pad=12)
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

    async def _show_global_analysis(self, interaction: discord.Interaction):
        """環境分析コマンドの本体。SENSEKI_ADMIN_USER_ID 本人のみ実行できる"""
        if interaction.user.id != SENSEKI_ADMIN_USER_ID:
            await interaction.response.send_message(
                "このコマンドは管理者専用です。", ephemeral=True
            )
            return

        await interaction.response.defer(ephemeral=True, thinking=True)
        g = await get_global_summary()
        if g["total"] == 0:
            await interaction.followup.send("まだ記録がありません。", ephemeral=True)
            return
        matchup_rows = await get_class_matchup_matrix()
        # クラス別集計はSheets同期側と同じく、自分が使った試合＋相手として対面した試合の対称合算
        class_summary = combine_symmetric_class_stats(g["by_my_class"], g["by_opp_class"])

        images = await asyncio.gather(
            asyncio.to_thread(self._render_class_summary_image_sync, class_summary, CURRENT_ENV_ID),
            asyncio.to_thread(
                self._render_matchup_heatmap_sync, matchup_rows, CLASS_CHOICES,
                MATCHUP_MATRIX_MIN_SAMPLE, CURRENT_ENV_ID,
            ),
            asyncio.to_thread(self._render_first_second_image_sync, g["by_first"], CURRENT_ENV_ID),
        )
        files = [
            discord.File(buf, filename=name)
            for buf, name in zip(images, ["クラス別集計.png", "対面マトリクス.png", "先攻後攻集計.png"])
            if buf is not None
        ]

        await log_command_usage(str(interaction.user.id), "環境分析")

        if not files:
            await interaction.followup.send(
                "画像の生成に失敗しました（グラフ機能が無効になっている可能性があります。"
                "`/戦績管理 診断` を確認してください）。",
                ephemeral=True,
            )
            return

        await interaction.followup.send(
            f"📊 **現環境（{CURRENT_ENV_ID}）の全体分析**（{g['total']}戦・記録者{g['users']}名）",
            files=files,
            ephemeral=True,
        )

    # ---- /弱点対面 ----
    # SENSEKI_MEMBERS_ONLY が true になるまで①②とも誰でも使える（コマンド自体の公開範囲）。
    # ②（全体平均との比較）は premium_access() で個別に判定：
    #   会員 / 今の環境が無料開放中(PREMIUM_FREE_THIS_ENV) / 初回登録から TRIAL_DAYS 日以内
    #   のいずれかを満たせば表示。次の環境ではPREMIUM_FREE_THIS_ENVをFalseに戻すこと。
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

        lines.append("\n対面別の弱点はパネルの「📉 弱点対面」から確認できます。")

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
