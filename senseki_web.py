"""
戦績ツールのWebページ（スマホ向けの入力・成績・環境）。

BOT（senseki.py）と同じプロセスの中で動く小さなWebサーバー。記録は同じデータベースに保存され、
戦歴シートへの同期もDiscordからの記録と同じ流れで反映される。

【ログインなし（ゲスト）】
  ・環境レポート、環境全体の質問のヒントを見る
  ・戦績を入力する（この端末のブラウザの中にだけ保存。サーバーには送らない）
  ・端末内の記録の勝率・先攻後攻別・一覧・修正・取消
【ログインあり（Discordでログイン＋サーバー参加）】
  ・記録がみんなの集計・戦歴シートに反映される
  ・自分の全記録の一覧・修正・取消、ランクの更新
  ・クラス対面別の勝率、自分用の質問のヒント、質問・相談（Discordへ投稿）

【Railwayの変数】
  WEB_SESSION_SECRET     必須。ログイン状態の署名用のランダムな長い文字列（これが無いとWebは起動しない）
  WEB_BASE_URL           このページの公開URL（例 https://xxxx.up.railway.app）。末尾の / は不要
  DISCORD_CLIENT_ID      Discord開発者ページ（OAuth2）のクライアントID
  DISCORD_CLIENT_SECRET  同・クライアントシークレット
  WEB_INVITE_URL         Webからサーバーへ誘導する招待リンク（任意。Web専用のリンクを作ると流入元が分かる）
  Discord開発者ページの OAuth2 → Redirects に  {WEB_BASE_URL}/callback  を登録する

【安全面】
  ・ログインは Discord の OAuth2（identify のみ）。パスワードは扱わない
  ・セッションは署名つきCookie（HttpOnly / SameSite=Lax）。書き込みAPIは専用ヘッダー必須
  ・本人の記録しか読み書きできない。修正・取消は履歴を web_match_edits に残す
  ・連続入力・1日の入力数・IPごとのアクセス数に上限あり
"""
import asyncio
import base64
import hashlib
import hmac
import json
import os
import secrets
import time
from datetime import datetime, timedelta
from urllib.parse import urlencode

import discord
from aiohttp import ClientSession, ClientTimeout, web

import senseki as S

# ---------------------------------------------------------------- 設定
WEB_BASE_URL = os.environ.get("WEB_BASE_URL", "").rstrip("/")
DISCORD_CLIENT_ID = os.environ.get("DISCORD_CLIENT_ID", "")
DISCORD_CLIENT_SECRET = os.environ.get("DISCORD_CLIENT_SECRET", "")
WEB_SESSION_SECRET = os.environ.get("WEB_SESSION_SECRET", "")
WEB_INVITE_URL = os.environ.get("WEB_INVITE_URL", "")
WEB_PORT = int(os.environ.get("PORT", "8080"))

SESSION_COOKIE = "sw_session"
STATE_COOKIE = "sw_state"
SESSION_DAYS = 30
RECORD_MIN_INTERVAL = 1.5      # 秒。連続入力の間隔
RECORD_DAILY_LIMIT = 300       # 1人が1日に入力できる数
IMPORT_MAX = 30                # ゲストの記録を引き継げる最大数（1人1回だけ）
IP_LIMIT_PER_MIN = 240         # IPごとの1分あたりのAPI呼び出し上限
ENV_CACHE_SECONDS = 60

LOGIN_ENABLED = bool(DISCORD_CLIENT_ID and DISCORD_CLIENT_SECRET and WEB_BASE_URL and WEB_SESSION_SECRET)

_last_record: dict[str, float] = {}
_ip_hits: dict[str, list] = {}
_member_cache: dict[int, tuple] = {}
_env_cache: dict = {"t": 0.0, "data": None}


# ---------------------------------------------------------------- セッション
def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode().rstrip("=")


def _b64d(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def _sign(payload: str) -> str:
    return hmac.new(WEB_SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()


def make_session(uid: str, name: str, avatar: str, days: int = SESSION_DAYS) -> str:
    body = _b64e(json.dumps(
        {"u": str(uid), "n": name, "a": avatar, "exp": int(time.time()) + days * 86400},
        ensure_ascii=False).encode())
    return body + "." + _sign(body)


def read_session(token):
    if not token or "." not in token or not WEB_SESSION_SECRET:
        return None
    body, sig = token.rsplit(".", 1)
    if not hmac.compare_digest(sig, _sign(body)):
        return None
    try:
        data = json.loads(_b64d(body))
    except Exception:
        return None
    if data.get("exp", 0) < time.time():
        return None
    return data


def _secure_cookie() -> bool:
    return WEB_BASE_URL.startswith("https://")


def _client_ip(request) -> str:
    fwd = request.headers.get("X-Forwarded-For", "")
    return (fwd.split(",")[0].strip() if fwd else "") or (request.remote or "?")


def _json(data, status=200):
    return web.json_response(data, status=status, dumps=lambda o: json.dumps(o, ensure_ascii=False))


def _err(msg, status=400):
    return _json({"error": msg}, status)


# ---------------------------------------------------------------- ミドルウェア
@web.middleware
async def security_middleware(request: web.Request, handler):
    if request.path.startswith("/api/"):
        ip = _client_ip(request)
        now = time.time()
        hits = [t for t in _ip_hits.get(ip, []) if now - t < 60]
        if len(hits) >= IP_LIMIT_PER_MIN:
            return _err("アクセスが多すぎます。少し待ってからお試しください。", 429)
        hits.append(now)
        _ip_hits[ip] = hits
        if len(_ip_hits) > 5000:
            _ip_hits.clear()
        if request.method not in ("GET", "HEAD") and request.headers.get("X-SW") != "1":
            return _err("不正なリクエストです。", 403)
    resp = await handler(request)
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "same-origin"
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; img-src 'self' https://cdn.discordapp.com data:; "
        "style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; "
        "connect-src 'self'; frame-ancestors 'none'; form-action 'self' https://discord.com"
    )
    if request.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


# ---------------------------------------------------------------- 共通のデータ処理
def _label(gap: float, n: int) -> str:
    if n < S.HINT_ENV_MIN_GAMES:
        return "データ少なめ"
    if gap >= 5:
        return "勝ちやすい"
    if gap > -2:
        return "ふつう"
    if gap > -6:
        return "やや手ごわい"
    return "手ごわい"


async def _member(bot, uid: int):
    """サーバーのメンバーか確認する（2分だけ結果を覚えて、Discordへの問い合わせを減らす）"""
    now = time.time()
    cached = _member_cache.get(uid)
    if cached and now - cached[0] < 120:
        return cached[1]
    guild = bot.get_guild(S.GUILD_ID)
    m = None
    if guild is not None:
        m = guild.get_member(uid)
        if m is None:
            try:
                m = await guild.fetch_member(uid)
            except discord.HTTPException:
                m = None
    _member_cache[uid] = (now, m)
    if len(_member_cache) > 3000:
        _member_cache.clear()
    return m


def _rank_info(settings: dict) -> dict:
    tier, grade = settings.get("rank_tier"), settings.get("cr_grade")
    idx = S.rank_step_index(tier, grade)
    ladder = S.RANK_LADDER

    def item(i):
        t, g = ladder[i]
        return {"tier": t, "grade": g, "label": S.rank_step_label(t, g)}

    return {
        "label": S.rank_step_label(tier, grade) if tier else "未設定",
        "prev": item(idx - 1) if idx is not None and idx > 0 else None,
        "next": item(idx + 1) if idx is not None and idx < len(ladder) - 1 else None,
    }


def _fetch_user_matches_sync(uid: str, env_id: str, limit=None):
    conn = S._connect()
    try:
        sql = ("SELECT id, recorded_at, format, my_class, my_deck, opp_class, opp_deck, is_first, is_win "
               "FROM matches WHERE user_id = ? AND env_id = ? ORDER BY id DESC")
        args = [uid, env_id]
        if limit:
            sql += " LIMIT ?"
            args.append(limit)
        return [dict(r) for r in conn.execute(sql, args).fetchall()]
    finally:
        conn.close()


def _count_today_sync(uid: str) -> int:
    conn = S._connect()
    try:
        today = datetime.now(S.JST).date().isoformat()
        return conn.execute(
            "SELECT COUNT(*) FROM matches WHERE user_id = ? AND recorded_at LIKE ?", (uid, today + "%")
        ).fetchone()[0]
    finally:
        conn.close()


def _ensure_edit_log(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS web_match_edits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            match_id INTEGER NOT NULL,
            user_id TEXT NOT NULL,
            edited_at TEXT NOT NULL,
            action TEXT NOT NULL,
            before_json TEXT,
            after_json TEXT
        )
    """)


def _update_match_sync(uid: str, match_id: int, fields: dict):
    conn = S._connect()
    try:
        _ensure_edit_log(conn)
        row = conn.execute("SELECT * FROM matches WHERE id = ? AND user_id = ?", (match_id, uid)).fetchone()
        if row is None:
            return False
        before = dict(row)
        conn.execute(
            "UPDATE matches SET opp_class = ?, opp_deck = ?, is_first = ?, is_win = ? WHERE id = ? AND user_id = ?",
            (fields["opp_class"], fields["opp_deck"], fields["is_first"], fields["is_win"], match_id, uid),
        )
        conn.execute(
            "INSERT INTO web_match_edits (match_id, user_id, edited_at, action, before_json, after_json) "
            "VALUES (?, ?, ?, 'update', ?, ?)",
            (match_id, uid, datetime.now(S.JST).isoformat(),
             json.dumps({k: before[k] for k in ("opp_class", "opp_deck", "is_first", "is_win")}, ensure_ascii=False),
             json.dumps(fields, ensure_ascii=False)),
        )
        conn.commit()
        return True
    finally:
        conn.close()


def _log_delete_sync(uid: str, match_id: int):
    conn = S._connect()
    try:
        _ensure_edit_log(conn)
        row = conn.execute("SELECT * FROM matches WHERE id = ? AND user_id = ?", (match_id, uid)).fetchone()
        if row is None:
            return False
        conn.execute(
            "INSERT INTO web_match_edits (match_id, user_id, edited_at, action, before_json, after_json) "
            "VALUES (?, ?, ?, 'delete', ?, NULL)",
            (match_id, uid, datetime.now(S.JST).isoformat(), json.dumps(dict(row), ensure_ascii=False)),
        )
        conn.commit()
        return True
    finally:
        conn.close()


def _compute_user_stats(rows: list) -> dict:
    total = len(rows)
    wins = sum(1 for r in rows if r["is_win"])

    def part(rs):
        n = len(rs)
        return {"n": n, "rate": (sum(1 for r in rs if r["is_win"]) / n * 100) if n else None}

    by_class: dict = {}
    for r in rows:
        by_class.setdefault(r["opp_class"], []).append(r)
    classes = [{"cls": c, "n": len(rs), "rate": sum(1 for r in rs if r["is_win"]) / len(rs) * 100}
               for c, rs in by_class.items()]
    classes.sort(key=lambda x: -x["n"])
    by_deck: dict = {}
    for r in rows:
        if r.get("opp_deck"):
            by_deck.setdefault((r["opp_class"], r["opp_deck"]), []).append(r)
    decks = [{"cls": c, "deck": d, "n": len(rs), "rate": sum(1 for r in rs if r["is_win"]) / len(rs) * 100}
             for (c, d), rs in by_deck.items() if len(rs) >= 3]
    decks.sort(key=lambda x: -x["n"])
    return {
        "total": total, "wins": wins, "rate": (wins / total * 100) if total else None,
        "first": part([r for r in rows if r["is_first"]]),
        "second": part([r for r in rows if not r["is_first"]]),
        "classes": classes, "decks": decks[:8],
    }


async def _env_data():
    now = time.time()
    if _env_cache["data"] is not None and now - _env_cache["t"] < ENV_CACHE_SECONDS:
        return _env_cache["data"]
    matches = await S.fetch_all_matches(env_only=True)
    rows = S._rotation_rows(matches)
    total = len(rows)
    users = len({r["user_id"] for r in rows})
    avg = (sum(1 for r in rows if r["is_win"]) / total * 100) if total else 0.0
    shares = S._env_share_map(rows)
    groups: dict = {}
    for r in rows:
        if r.get("opp_deck"):
            groups.setdefault((r["opp_class"], r["opp_deck"]), []).append(r)
    decks = []
    for (c, d), rs in groups.items():
        n = len(rs)
        rate = sum(1 for r in rs if r["is_win"]) / n * 100
        decks.append({"cls": c, "deck": d, "n": n, "share": shares.get((c, d), 0.0), "rate": rate,
                      "label": _label(rate - avg, n)})
    decks.sort(key=lambda x: -x["share"])
    firsts = [r for r in rows if r["is_first"]]
    seconds = [r for r in rows if not r["is_first"]]
    hints, _ = S.compute_env_hints(matches)
    data = {
        "env_id": S.current_env_id(), "total": total, "users": users, "avg": avg,
        "first": {"n": len(firsts), "rate": (sum(1 for r in firsts if r["is_win"]) / len(firsts) * 100) if firsts else None},
        "second": {"n": len(seconds), "rate": (sum(1 for r in seconds if r["is_win"]) / len(seconds) * 100) if seconds else None},
        "decks": decks[:10],
        "hints": [{"cls": h["class"], "deck": h["deck"], "share": h["share"], "rate": h["rate"], "n": h["games"],
                   "label": _label(h["gap"], h["games"])} for h in hints],
    }
    _env_cache["t"], _env_cache["data"] = now, data
    return data


def _channel_url(cid) -> str:
    return f"https://discord.com/channels/{S.GUILD_ID}/{cid}" if cid else ""


# ---------------------------------------------------------------- ハンドラ
def build_app(bot) -> web.Application:
    app = web.Application(middlewares=[security_middleware])
    routes = web.RouteTableDef()

    async def current_user(request):
        """ログイン中ならセッションの内容を返す（なければ None）"""
        return read_session(request.cookies.get(SESSION_COOKIE))

    async def require_member(request):
        sess = await current_user(request)
        if not sess:
            return None, None, _err("ログインしてください。", 401)
        m = await _member(bot, int(sess["u"]))
        if m is None:
            return sess, None, _err("しろくさサーバーに参加すると使えます。", 403)
        return sess, m, None

    @routes.get("/")
    async def index(request):
        return web.Response(text=INDEX_HTML, content_type="text/html")

    @routes.get("/health")
    async def health(request):
        return web.Response(text="ok")

    # ---- ログイン（Discord OAuth2）
    @routes.get("/login")
    async def login(request):
        if not LOGIN_ENABLED:
            return web.Response(text="ログインはまだ設定されていません。", status=503)
        state = secrets.token_urlsafe(24)
        url = "https://discord.com/oauth2/authorize?" + urlencode({
            "client_id": DISCORD_CLIENT_ID, "redirect_uri": WEB_BASE_URL + "/callback",
            "response_type": "code", "scope": "identify", "state": state, "prompt": "none",
        })
        resp = web.HTTPFound(url)
        resp.set_cookie(STATE_COOKIE, state, max_age=600, httponly=True, secure=_secure_cookie(), samesite="Lax")
        raise resp

    @routes.get("/callback")
    async def callback(request):
        if not LOGIN_ENABLED:
            return web.Response(text="ログインはまだ設定されていません。", status=503)
        code = request.query.get("code")
        state = request.query.get("state")
        if not code or not state or state != request.cookies.get(STATE_COOKIE):
            return web.Response(text="ログインに失敗しました。もう一度お試しください。", status=400)
        try:
            async with ClientSession(timeout=ClientTimeout(total=15)) as http:
                async with http.post("https://discord.com/api/oauth2/token", data={
                    "client_id": DISCORD_CLIENT_ID, "client_secret": DISCORD_CLIENT_SECRET,
                    "grant_type": "authorization_code", "code": code,
                    "redirect_uri": WEB_BASE_URL + "/callback",
                }) as r:
                    tok = await r.json()
                if "access_token" not in tok:
                    return web.Response(text="ログインに失敗しました（認証エラー）。", status=400)
                async with http.get("https://discord.com/api/users/@me",
                                    headers={"Authorization": "Bearer " + tok["access_token"]}) as r:
                    me = await r.json()
        except Exception as e:
            print(f"⚠️ Webログインに失敗: {type(e).__name__}: {e}")
            return web.Response(text="ログインに失敗しました。時間をおいてお試しください。", status=502)
        if "id" not in me:
            return web.Response(text="ログインに失敗しました。", status=400)
        name = me.get("global_name") or me.get("username") or "ユーザー"
        avatar = (f"https://cdn.discordapp.com/avatars/{me['id']}/{me['avatar']}.png?size=64"
                  if me.get("avatar") else "")
        resp = web.HTTPFound("/")
        resp.set_cookie(SESSION_COOKIE, make_session(me["id"], name, avatar), max_age=SESSION_DAYS * 86400,
                        httponly=True, secure=_secure_cookie(), samesite="Lax")
        resp.del_cookie(STATE_COOKIE)
        raise resp

    @routes.get("/logout")
    async def logout(request):
        resp = web.HTTPFound("/")
        resp.del_cookie(SESSION_COOKIE)
        raise resp

    # ---- 状態
    @routes.get("/api/me")
    async def api_me(request):
        sess = await current_user(request)
        out = {"login_enabled": LOGIN_ENABLED, "invite_url": WEB_INVITE_URL, "classes": S.CLASS_CHOICES,
               "consult_key": S.CONSULT_KEY, "logged_in": False, "member": False, "settings": None}
        if not sess:
            return _json(out)
        out.update({"logged_in": True, "name": sess.get("n"), "avatar": sess.get("a")})
        m = await _member(bot, int(sess["u"]))
        if m is None:
            return _json(out)
        out["member"] = True
        settings = await S.get_user_settings(sess["u"])
        if settings:
            out["settings"] = {
                "my_class": settings["my_class"], "my_deck": settings.get("my_deck"),
                "format": S.FORMAT_LABELS.get(settings["format"], settings["format"]),
                "rank": _rank_info(settings),
            }
            out["ladder"] = [{"tier": t, "grade": g, "label": S.rank_step_label(t, g)} for t, g in S.RANK_LADDER]
        return _json(out)

    @routes.get("/api/decks")
    async def api_decks(request):
        cls = request.query.get("cls", "")
        if cls not in S.CLASS_CHOICES:
            return _err("クラスが正しくありません。")
        return _json({"decks": await S.get_deck_names(cls, 30)})

    @routes.get("/api/env")
    async def api_env(request):
        data = dict(await _env_data())
        cmap = await S.class_channel_map()
        data["hints"] = [dict(h, url=_channel_url(cmap.get(h["cls"]))) for h in data["hints"]]
        data["consult_url"] = _channel_url(cmap.get(S.CONSULT_KEY))
        return _json(data)

    # ---- 記録
    @routes.post("/api/record")
    async def api_record(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        try:
            body = await request.json()
        except Exception:
            return _err("入力が正しくありません。")
        uid = sess["u"]
        settings = await S.get_user_settings(uid)
        if not settings:
            return _err("先にDiscordの「🆕 初回設定」をしてください。", 409)
        opp_class = body.get("opp_class")
        if opp_class not in S.CLASS_CHOICES:
            return _err("相手のクラスを選んでください。")
        opp_deck = (body.get("opp_deck") or "").strip() or None
        if opp_deck and opp_deck not in await S.get_deck_names(opp_class, 50):
            return _err("相手デッキが候補にありません。")
        if not isinstance(body.get("is_first"), bool) or not isinstance(body.get("is_win"), bool):
            return _err("先攻・後攻と勝敗を選んでください。")
        now = time.time()
        if now - _last_record.get(uid, 0) < RECORD_MIN_INTERVAL:
            return _err("少し間を空けてください。", 429)
        if await asyncio.to_thread(_count_today_sync, uid) >= RECORD_DAILY_LIMIT:
            return _err("今日の入力数の上限に達しました。", 429)
        _last_record[uid] = now
        match_id = await S.insert_match(uid, settings, opp_class, opp_deck, body["is_first"], body["is_win"])
        variant = await S.get_active_variant(uid, settings["format"], settings["my_class"], settings.get("my_deck"))
        if variant:
            await S.insert_personal_variant(match_id, uid, variant)
        return _json({"id": match_id})

    @routes.delete(r"/api/match/{mid:\d+}")
    async def api_match_delete(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        mid = int(request.match_info["mid"])
        if not await asyncio.to_thread(_log_delete_sync, sess["u"], mid):
            return _err("記録が見つかりません。", 404)
        await S.delete_match_by_id(sess["u"], mid)
        return _json({"ok": True})

    @routes.patch(r"/api/match/{mid:\d+}")
    async def api_match_update(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        try:
            body = await request.json()
        except Exception:
            return _err("入力が正しくありません。")
        opp_class = body.get("opp_class")
        if opp_class not in S.CLASS_CHOICES:
            return _err("相手のクラスを選んでください。")
        opp_deck = (body.get("opp_deck") or "").strip() or None
        if opp_deck and opp_deck not in await S.get_deck_names(opp_class, 50):
            return _err("相手デッキが候補にありません。")
        if not isinstance(body.get("is_first"), bool) or not isinstance(body.get("is_win"), bool):
            return _err("先攻・後攻と勝敗を選んでください。")
        fields = {"opp_class": opp_class, "opp_deck": opp_deck,
                  "is_first": 1 if body["is_first"] else 0, "is_win": 1 if body["is_win"] else 0}
        if not await asyncio.to_thread(_update_match_sync, sess["u"], int(request.match_info["mid"]), fields):
            return _err("記録が見つかりません。", 404)
        S.mark_dirty()
        return _json({"ok": True})

    @routes.post("/api/rank")
    async def api_rank(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        try:
            body = await request.json()
        except Exception:
            return _err("入力が正しくありません。")
        pair = (body.get("tier"), body.get("grade") or None)
        if pair not in S.RANK_LADDER:
            return _err("ランクが正しくありません。")
        if not await S.update_rank(sess["u"], pair[0], pair[1]):
            return _err("先にDiscordの「🆕 初回設定」をしてください。", 409)
        cog = bot.get_cog("SensekiCog")
        if cog is not None:
            try:
                await cog._update_board()
            except Exception as e:
                print(f"⚠️ 掲示板の更新に失敗: {type(e).__name__}: {e}")
        settings = await S.get_user_settings(sess["u"])
        return _json({"rank": _rank_info(settings)})

    # ---- 成績・ヒント・質問
    @routes.get("/api/stats")
    async def api_stats(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        rows = await asyncio.to_thread(_fetch_user_matches_sync, sess["u"], S.current_env_id())
        stats = _compute_user_stats(rows)
        stats["recent"] = [
            {"id": r["id"], "t": (r["recorded_at"] or "")[:16].replace("T", " "), "cls": r["opp_class"],
             "deck": r["opp_deck"] or "", "first": bool(r["is_first"]), "win": bool(r["is_win"])}
            for r in rows[:50]
        ]
        return _json(stats)

    @routes.get("/api/hint")
    async def api_hint(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        matches = await S.fetch_all_matches(env_only=True)
        res = S.compute_personal_hints(matches, sess["u"])
        cmap = await S.class_channel_map()
        settings = await S.get_user_settings(sess["u"])
        out = {"total": res["total"], "enough": res["enough"], "need": max(S.HINT_PERSONAL_MIN_TOTAL - res["total"], 0),
               "weak": [dict(w, url=_channel_url(cmap.get(w["class"]))) for w in res["weak"]],
               "first": res["first"], "consult_url": _channel_url(cmap.get(S.CONSULT_KEY))}
        if res["first"] and settings:
            out["first"] = dict(res["first"], url=_channel_url(cmap.get(settings["my_class"])))
        return _json(out)

    @routes.post("/api/question")
    async def api_question(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        try:
            body = await request.json()
        except Exception:
            return _err("入力が正しくありません。")
        cls = body.get("cls")
        if cls not in S.CLASS_CHOICES and cls != S.CONSULT_KEY:
            return _err("質問先を選んでください。")
        ok, message, url = await S.submit_question(
            bot, m, cls, str(body.get("title", ""))[:80], str(body.get("body", ""))[:800], bool(body.get("attach")))
        if not ok:
            return _err(message, 429 if "間を空けて" in message or "1日" in message else 400)
        return _json({"message": message, "url": url})

    # ---- ゲストの記録の引き継ぎ（1人1回・最大30件）
    @routes.post("/api/import")
    async def api_import(request):
        sess, m, err = await require_member(request)
        if err:
            return err
        uid = sess["u"]
        settings = await S.get_user_settings(uid)
        if not settings:
            return _err("先にDiscordの「🆕 初回設定」をしてください。", 409)
        if await S.get_state("web_import_" + uid):
            return _json({"imported": 0, "already": True})
        try:
            body = await request.json()
        except Exception:
            return _err("入力が正しくありません。")
        recs = body.get("records") or []
        if not isinstance(recs, list):
            return _err("入力が正しくありません。")
        names_cache: dict = {}
        count = 0
        for r in reversed(recs[:IMPORT_MAX]):    # 古い順に入れる
            if not isinstance(r, dict) or r.get("c") not in S.CLASS_CHOICES:
                continue
            c = r["c"]
            if c not in names_cache:
                names_cache[c] = await S.get_deck_names(c, 50)
            deck = r.get("d") if r.get("d") in names_cache[c] else None
            first = r.get("f")
            if first not in (0, 1, True, False):
                continue
            await S.insert_match(uid, settings, c, deck, bool(first), bool(r.get("w")))
            count += 1
        await S.set_state("web_import_" + uid, "1")
        return _json({"imported": count})

    app.add_routes(routes)
    return app


async def start_web(bot):
    """BOTと同じプロセスでWebサーバーを起動する。WEB_SESSION_SECRET が無ければ起動しない"""
    if not WEB_SESSION_SECRET:
        print("ℹ️ WEB_SESSION_SECRET が未設定のため、戦績ツールのWebページは起動しません")
        return None
    app = build_app(bot)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", WEB_PORT)
    await site.start()
    print(f"✅ 戦績ツールのWebページを起動しました（ポート{WEB_PORT} / ログイン{'有効' if LOGIN_ENABLED else '無効'}）")
    return runner


# ---------------------------------------------------------------- 画面（1ファイルで完結）
INDEX_HTML = r'''<!doctype html>
<html lang="ja"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#1c1f2a">
<meta name="apple-mobile-web-app-capable" content="yes">
<title>しろくさ 戦績ツール</title>
<style>
:root{--bg:#1c1f2a;--card:#262a38;--line:#2f3447;--tx:#e6e8f0;--mute:#9aa1b8;--pri:#5865f2;--win:#2e9e5b;--lose:#d6454c;--warn:#f0c75e}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;background:var(--bg);color:var(--tx);font-family:system-ui,-apple-system,"Hiragino Sans","Noto Sans JP",sans-serif;font-size:15px;line-height:1.5;padding-bottom:84px}
.wrap{max-width:560px;margin:0 auto;padding:14px 16px}
h1{font-size:17px;margin:2px 0 12px;display:flex;align-items:center;gap:8px;justify-content:space-between}
h1 img{width:24px;height:24px;border-radius:50%;vertical-align:middle;margin-right:6px}
.card{background:var(--card);border-radius:14px;padding:13px 14px;margin-bottom:12px}
.lbl{font-size:12px;color:var(--mute);margin-bottom:7px}
.banner{border-radius:10px;padding:9px 12px;font-size:13px;margin-bottom:12px}
.b-warn{background:#3b3524;color:var(--warn)}.b-ok{background:#23382b;color:#7fd69e}.b-info{background:#2a3050;color:#b9c2ff}
.chips{display:grid;grid-template-columns:repeat(4,1fr);gap:7px}
.chip{background:#363b4e;border:0;color:var(--tx);border-radius:10px;padding:12px 0;font-size:14px;cursor:pointer}
.chip.on{background:var(--pri);color:#fff;font-weight:700}
select,textarea,input[type=text]{width:100%;background:#2f3447;color:var(--tx);border:0;border-radius:10px;padding:12px;font-size:15px;font-family:inherit}
textarea{min-height:96px;resize:vertical}
.seg{display:flex;background:#363b4e;border-radius:10px;overflow:hidden}
.seg button{flex:1;border:0;background:none;color:var(--tx);padding:13px 0;font-size:15px;cursor:pointer}
.seg button.on{background:var(--pri);color:#fff;font-weight:700}
.row{display:flex;gap:10px}.row>*{flex:1}
.btn{display:block;width:100%;border:0;border-radius:12px;padding:15px;font-size:16px;font-weight:700;color:#fff;text-align:center;text-decoration:none;cursor:pointer;background:#4e5058}
.btn.pri{background:var(--pri)}.btn.win{background:var(--win);padding:20px}.btn.lose{background:var(--lose);padding:20px}.btn.sec{background:#363b4e}
.btn.sm{padding:9px 12px;font-size:13px;display:inline-block;width:auto;border-radius:8px}
.btn[disabled]{opacity:.5}
.rank{display:flex;gap:8px;align-items:center}
.rank .now{flex:1.4;background:#2f3447;border-radius:10px;padding:11px 8px;text-align:center;font-size:13px;color:#c5cae0}
.rank button{flex:1;border:0;border-radius:10px;padding:11px 4px;background:var(--pri);color:#fff;font-weight:700;font-size:13px}
.lock{border:1px dashed #4a5070;border-radius:14px;padding:13px 14px;margin-bottom:12px;color:#b4bad0;font-size:14px}
.bar{height:9px;border-radius:5px;background:var(--pri);margin:4px 0}
table{width:100%;border-collapse:collapse;font-size:14px}
td,th{padding:8px 4px;border-bottom:1px solid var(--line);text-align:left}th{color:var(--mute);font-weight:400;font-size:12px}
td.r,th.r{text-align:right}
.hi{color:#7fd69e}.lo{color:#f08a8a}
.tabs{position:fixed;left:0;right:0;bottom:0;background:#14161d;border-top:1px solid var(--line);display:flex;padding-bottom:env(safe-area-inset-bottom)}
.tabs button{flex:1;border:0;background:none;color:var(--mute);padding:12px 0 13px;font-size:12px}
.tabs button.on{color:#fff;font-weight:700}.tabs span{display:block;font-size:19px}
.toast{position:fixed;left:12px;right:12px;bottom:92px;background:#fff;color:#14161d;border-radius:12px;padding:12px 14px;font-size:14px;display:flex;gap:10px;align-items:center;justify-content:space-between;z-index:30;max-width:536px;margin:0 auto}
.toast button{border:0;background:#14161d;color:#fff;border-radius:8px;padding:7px 12px}
.sheet{position:fixed;inset:0;background:#000b;z-index:40;display:flex;align-items:flex-end}
.sheet>div{background:var(--bg);width:100%;max-width:560px;margin:0 auto;border-radius:18px 18px 0 0;padding:16px;max-height:92vh;overflow:auto}
small,.mute{color:var(--mute);font-size:12px}
a{color:#aab6ff}
.tag{display:inline-block;background:#363b4e;border-radius:6px;padding:2px 8px;font-size:11px;margin-left:6px}
.item{display:flex;align-items:center;gap:8px;padding:9px 0;border-bottom:1px solid var(--line);font-size:14px}
.item .grow{flex:1}.item .mini{border:0;background:#363b4e;color:var(--tx);border-radius:7px;padding:6px 10px;font-size:12px}
.w{color:#7fd69e;font-weight:700}.l{color:#f08a8a;font-weight:700}
</style></head><body>
<div class="wrap" id="app">読み込み中…</div>
<div class="tabs" id="tabs"></div>
<div id="layer"></div>
<script>
"use strict";
const $=(s)=>document.querySelector(s);
const esc=(s)=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
let ME={}, tab=(function(){try{return localStorage.getItem("sw_tab")||"input"}catch(e){return "input"}})();
let sel={cls:null,deck:"",first:null}, decks={}, last=null, ENV=null, ST=null, HINT=null, toastTimer=null;
const GK="sw_guest_v1";
function gload(){try{return JSON.parse(localStorage.getItem(GK))||{r:[],seq:0}}catch(e){return {r:[],seq:0}}}
function gsave(g){try{localStorage.setItem(GK,JSON.stringify(g))}catch(e){}}
async function api(path,method,body){
  try{
    const r=await fetch(path,{method:method||"GET",credentials:"same-origin",headers:{"X-SW":"1","Content-Type":"application/json"},body:body?JSON.stringify(body):undefined});
    let data={};try{data=await r.json()}catch(e){}
    return {ok:r.ok,status:r.status,data};
  }catch(e){return {ok:false,status:0,data:{error:"通信に失敗しました。電波の良い場所でもう一度お試しください。"}}}
}
function toast(msg,undo){
  clearTimeout(toastTimer);
  $("#layer").innerHTML=`<div class="toast"><span>${esc(msg)}</span>${undo?'<button data-act="undo">取り消す</button>':""}</div>`;
  toastTimer=setTimeout(()=>{const l=$("#layer");if(l&&l.firstChild&&l.firstChild.className==="toast")l.innerHTML=""},undo?7000:3500);
}
const isServer=()=>ME.logged_in&&ME.member&&ME.settings;
const pct=(v)=>v==null?"—":Math.round(v)+"%";

/* ---------- 初期化 ---------- */
async function init(){
  const r=await api("/api/me");ME=r.data||{};
  if(isServer()){await tryImport()}
  render();
}
async function tryImport(){
  const g=gload();if(!g.r.length)return;
  const r=await api("/api/import","POST",{records:g.r});
  if(r.ok){
    if(r.data.imported>0)toast("この端末の記録"+r.data.imported+"件を、みんなの集計に反映しました");
    if(r.data.imported>0||r.data.already)gsave({r:[],seq:g.seq});
  }
}
/* ---------- 描画 ---------- */
function renderTabs(){
  const t=[["input","⚔","入力"],["stats","📊","成績"],["env","🌍","環境"]];
  $("#tabs").innerHTML=t.map(x=>`<button class="${tab===x[0]?"on":""}" data-tab="${x[0]}"><span>${x[1]}</span>${x[2]}</button>`).join("");
}
function header(){
  let right="";
  if(ME.logged_in){right=`<span class="mute">${ME.avatar?`<img src="${esc(ME.avatar)}" alt="">`:""}${esc(ME.name)} <a href="/logout">ログアウト</a></span>`}
  return `<h1><span>⚔ しろくさ 戦績ツール</span>${right}</h1>`;
}
function banner(){
  if(!ME.logged_in)return `<div class="banner b-warn">ゲスト（記録はこの端末にだけ保存されます）</div>`;
  if(!ME.member)return `<div class="banner b-info">ログインしました。<b>しろくさサーバーに参加</b>すると、記録がみんなの集計に反映されます。${ME.invite_url?`<br><a href="${esc(ME.invite_url)}">サーバーに参加する</a>`:""}</div>`;
  if(!ME.settings)return `<div class="banner b-info">Discordの <b>🆕 初回設定</b> をすると、記録がみんなの集計に反映されます（それまではこの端末に保存）。</div>`;
  return `<div class="banner b-ok">✅ ${esc(ME.name)} としてログイン中（集計に反映されます）</div>`;
}
function loginBtn(){return ME.login_enabled&&!ME.logged_in?`<a class="btn pri" href="/login">Discordでログイン</a>`:""}
function render(){
  renderTabs();
  const app=$("#app");
  if(tab==="input")app.innerHTML=header()+viewInput();
  else if(tab==="stats"){app.innerHTML=header()+`<div class="card">読み込み中…</div>`;loadStats()}
  else{app.innerHTML=header()+`<div class="card">読み込み中…</div>`;loadEnv()}
}
function viewInput(){
  const chips=(ME.classes||[]).map(c=>`<button class="chip ${sel.cls===c?"on":""}" data-act="cls" data-v="${esc(c)}">${esc(c)}</button>`).join("");
  const dl=sel.cls?(decks[sel.cls]||[]):[];
  const opts=`<option value="">（分かるときだけ選択）</option>`+dl.map(d=>`<option ${sel.deck===d?"selected":""}>${esc(d)}</option>`).join("");
  let rank="";
  if(isServer()){
    const rk=ME.settings.rank;
    rank=`<div class="card"><div class="lbl">📈 ランク（変わったらここで更新）</div><div class="rank">
      <div class="now">今：${esc(rk.label)}</div>
      ${rk.prev?`<button data-act="rank" data-i="prev">⬇ ${esc(rk.prev.label)}</button>`:""}
      ${rk.next?`<button data-act="rank" data-i="next">⬆ ${esc(rk.next.label)}</button>`:""}
    </div><div style="margin-top:8px"><select data-act="rankpick"><option value="">ランクを選び直す…</option>${(ME.ladder||[]).map((l,i)=>`<option value="${i}">${esc(l.label)}</option>`).join("")}</select></div></div>`;
  }
  const hint=(!ME.logged_in&&ME.login_enabled)?`<p class="mute" style="text-align:center">ログインなしでも使えます。ログインすると、みんなの集計に反映され、クラス対面別の勝率なども見られます。</p>${loginBtn()}`:"";
  return banner()+`
  <div class="card"><div class="lbl">① 相手クラス</div><div class="chips">${chips}</div>
    <div class="lbl" style="margin-top:12px">② 相手デッキ（任意）</div><select data-act="deck" ${sel.cls?"":"disabled"}>${opts}</select>
    <div class="lbl" style="margin-top:12px">③ 先攻 / 後攻</div><div class="seg"><button class="${sel.first===1?"on":""}" data-act="first" data-v="1">先攻</button><button class="${sel.first===0?"on":""}" data-act="first" data-v="0">後攻</button></div>
    <div class="lbl" style="margin-top:12px">④ 結果</div><div class="row"><button class="btn win" data-act="win" data-v="1">勝ち</button><button class="btn lose" data-act="win" data-v="0">負け</button></div>
  </div>${rank}${hint}`;
}
async function loadDecks(cls){
  if(decks[cls])return;
  const r=await api("/api/decks?cls="+encodeURIComponent(cls));
  decks[cls]=(r.data&&r.data.decks)||[];
}
async function doRecord(win){
  if(sel.cls===null||sel.first===null){toast("相手クラスと、先攻・後攻を選んでください");return}
  if(isServer()){
    const r=await api("/api/record","POST",{opp_class:sel.cls,opp_deck:sel.deck||null,is_first:sel.first===1,is_win:win});
    if(!r.ok){toast(r.data.error||"記録に失敗しました");return}
    last={server:true,id:r.data.id};
  }else{
    const g=gload();g.seq++;
    g.r.unshift({id:g.seq,t:Date.now(),c:sel.cls,d:sel.deck||"",f:sel.first,w:win?1:0});
    gsave(g);last={server:false,id:g.seq};
  }
  sel={cls:null,deck:"",first:null};
  render();toast((win?"勝ち":"負け")+"で記録しました",true);
}
async function doUndo(){
  if(!last)return;
  if(last.server){const r=await api("/api/match/"+last.id,"DELETE");if(!r.ok){toast(r.data.error||"取り消せませんでした");return}}
  else{const g=gload();g.r=g.r.filter(x=>x.id!==last.id);gsave(g)}
  last=null;toast("取り消しました");if(tab!=="input")render();
}
/* ---------- 成績 ---------- */
function localStats(){
  const rs=gload().r;const t=rs.length,w=rs.filter(x=>x.w).length;
  const part=(a)=>({n:a.length,rate:a.length?a.filter(x=>x.w).length/a.length*100:null});
  return {total:t,wins:w,rate:t?w/t*100:null,first:part(rs.filter(x=>x.f===1)),second:part(rs.filter(x=>x.f===0)),
    recent:rs.slice(0,50).map(x=>({id:x.id,t:new Date(x.t).toLocaleString("ja-JP",{month:"numeric",day:"numeric",hour:"2-digit",minute:"2-digit"}),cls:x.c,deck:x.d,first:x.f===1,win:!!x.w}))};
}
async function loadStats(){
  if(isServer()){const r=await api("/api/stats");ST=r.ok?r.data:null;const h=await api("/api/hint");HINT=h.ok?h.data:null}
  else{ST=localStats();HINT=null}
  if(tab!=="stats")return;
  $("#app").innerHTML=header()+viewStats();
}
function summary(s){
  if(!s||!s.total)return `<div class="card"><div class="lbl">${isServer()?"現環境の成績":"この端末の成績"}</div><div>まだ記録がありません。「入力」から1戦記録してみましょう。</div></div>`;
  return `<div class="card"><div class="lbl">${isServer()?"現環境の成績（あなたの全記録）":"この端末の成績"}</div>
  <div style="font-size:20px;font-weight:700">${s.total}戦 ${s.wins}勝${s.total-s.wins}敗　勝率 ${pct(s.rate)}</div>
  <div class="row" style="margin-top:10px"><div><span class="mute">先攻</span> ${pct(s.first.rate)}<small>（${s.first.n}戦）</small><div class="bar" style="width:${s.first.rate||0}%"></div></div>
  <div><span class="mute">後攻</span> ${pct(s.second.rate)}<small>（${s.second.n}戦）</small><div class="bar" style="width:${s.second.rate||0}%"></div></div></div></div>`;
}
function recent(s){
  if(!s||!s.recent||!s.recent.length)return "";
  return `<div class="card"><div class="lbl">${isServer()?"全記録（新しい順・50件まで）":"直近の記録"}</div>${s.recent.map(x=>`
   <div class="item"><span class="${x.win?"w":"l"}">${x.win?"勝":"負"}</span><span class="grow">vs ${esc(x.cls)}${x.deck?`（${esc(x.deck)}）`:""}<small> ${x.first?"先":"後"}・${esc(x.t)}</small></span>
   <button class="mini" data-act="edit" data-id="${x.id}">修正</button><button class="mini" data-act="del" data-id="${x.id}">取消</button></div>`).join("")}</div>`;
}
function viewStats(){
  let h=banner()+summary(ST);
  if(isServer()&&ST&&ST.classes&&ST.classes.length){
    h+=`<div class="card"><div class="lbl">クラス対面別の勝率（相手クラスごと）</div><table><tr><th>相手</th><th class="r">戦</th><th class="r">勝率</th></tr>${ST.classes.map(c=>`<tr><td>${esc(c.cls)}</td><td class="r">${c.n}</td><td class="r ${c.rate>=60?"hi":c.rate<45?"lo":""}">${pct(c.rate)}</td></tr>`).join("")}</table>
    <small>試合数が少ないうちは偏ります。目安としてご覧ください。</small></div>`;
  }else if(!isServer()){
    h+=`<div class="lock">🔒 <b>クラス対面別の勝率・質問のヒント・質問する</b><br><span class="mute">Discordにログイン（サーバー参加）すると使えます。記録もみんなの集計に反映されます。</span><div style="margin-top:10px">${loginBtn()}${ME.invite_url?`<a class="btn sec" style="margin-top:8px" href="${esc(ME.invite_url)}">Discordサーバーに参加する</a>`:""}</div></div>`;
  }
  if(isServer())h+=viewHint()+viewAsk();
  h+=recent(ST);
  return h;
}
function viewHint(){
  if(!HINT)return "";
  if(!HINT.total)return `<div class="card"><div class="lbl">🎯 質問のヒント</div>まだ現環境の記録がありません。記録がたまると、あなた用のヒントが出ます。</div>`;
  if(!HINT.enough)return `<div class="card"><div class="lbl">🎯 質問のヒント</div>あと<b>${HINT.need}戦</b>記録すると、あなた用のヒントが出ます。</div>`;
  let h=`<div class="card"><div class="lbl">🎯 質問のヒント（あなたの記録から）</div>`;
  if(HINT.weak.length){h+=HINT.weak.map(w=>`<div class="item"><span class="grow"><b>${esc(w.deck)}</b>（${esc(w.class)}）<small> あなたの勝率${pct(w.rate)}（${w.games}戦）／環境で${Math.round(w.share)}%</small></span>${w.url?`<a class="mini btn sm sec" style="text-decoration:none" href="${esc(w.url)}" target="_blank" rel="noopener">${esc(w.class)}のチャンネル</a>`:""}</div>`).join("")}
  else h+=`<div>特に苦手がはっきりしている相手は、今のところありません。</div>`;
  if(HINT.first)h+=`<div style="margin-top:8px">先攻${pct(HINT.first.first_rate)}（${HINT.first.first_n}戦）／後攻${pct(HINT.first.second_rate)}（${HINT.first.second_n}戦）<small> 差が大きいので、${HINT.first.second_rate<HINT.first.first_rate?"後攻":"先攻"}の立ち回りを聞いてみても</small></div>`;
  if(HINT.consult_url)h+=`<div style="margin-top:10px">🧭 使うデッキに迷ったら <a href="${esc(HINT.consult_url)}" target="_blank" rel="noopener">シャドバ相談所</a></div>`;
  return h+`<small>あくまで目安です。試合数が少ないうちは偏りが出ます。</small></div>`;
}
function viewAsk(){
  const o=(ME.classes||[]).map(c=>`<option value="${esc(c)}">${esc(c)}のチャンネル</option>`).join("")+`<option value="${esc(ME.consult_key)}">🧭 使うデッキの相談（シャドバ相談所）</option>`;
  return `<div class="card"><div class="lbl">❓ 質問・相談する（Discordに投稿されます）</div>
  <select id="q-cls">${o}</select><div style="height:8px"></div>
  <input type="text" id="q-title" maxlength="80" placeholder="題名（例：秘術ウィッチへの対策を教えてください）"><div style="height:8px"></div>
  <textarea id="q-body" maxlength="800" placeholder="内容（自分のデッキ・ランク帯・困った場面）"></textarea>
  <label style="display:block;margin:8px 0;font-size:14px"><input type="checkbox" id="q-attach" checked> 📎 自分の戦績を添付する（使用デッキ・ランク・勝率）</label>
  <button class="btn pri" data-act="ask">質問を投稿する</button></div>`;
}
/* ---------- 環境 ---------- */
async function loadEnv(){
  if(!ENV){const r=await api("/api/env");ENV=r.ok?r.data:null}
  if(tab!=="env")return;
  const e=ENV;if(!e){$("#app").innerHTML=header()+`<div class="card">読み込みに失敗しました。</div>`;return}
  const cls={"勝ちやすい":"hi","手ごわい":"lo","やや手ごわい":"lo"};
  let h=`<div class="card"><div class="lbl">📊 今の環境（${esc(e.env_id)}・ローテーション）</div>
   <div style="font-size:17px;font-weight:700">${e.total}戦・記録者${e.users}名</div>
   <div class="mute" style="margin-top:4px">先攻の勝率 ${pct(e.first.rate)}（${e.first.n}戦）／後攻の勝率 ${pct(e.second.rate)}（${e.second.n}戦）</div></div>
   <div class="card"><div class="lbl">よく当たるデッキ TOP10</div>${e.decks.map(d=>`<div style="margin-bottom:6px"><div style="display:flex;justify-content:space-between"><span><b>${esc(d.deck)}</b><small>（${esc(d.cls)}）</small></span><span>${d.share.toFixed(1)}% <span class="${cls[d.label]||""}" style="font-size:12px">${esc(d.label)}</span></span></div><div class="bar" style="width:${Math.min(d.share*5,100)}%"></div></div>`).join("")}
   <small>「勝ちやすい／手ごわい」は、記録者がそのデッキと当たったときの勝率を、記録者全体の平均（約${Math.round(e.avg)}%）と比べたものです。デッキそのものの勝率ではありません。記録者はDiscord内の方で、上位ランク帯に偏りがあります。</small></div>`;
  if(e.hints&&e.hints.length){
    h+=`<div class="card"><div class="lbl">🎯 質問のヒント（手ごわい相手）</div>${e.hints.map(x=>`<div class="item"><span class="grow"><b>${esc(x.deck)}</b>（${esc(x.cls)}）<small> 当たる割合${Math.round(x.share)}%／記録者の勝率${pct(x.rate)}（${x.n}戦）</small></span>${x.url?`<a class="btn sm sec" style="text-decoration:none" href="${esc(x.url)}" target="_blank" rel="noopener">質問する</a>`:""}</div>`).join("")}
    ${e.consult_url?`<div style="margin-top:10px">🧭 使うデッキに迷ったら <a href="${esc(e.consult_url)}" target="_blank" rel="noopener">シャドバ相談所</a></div>`:""}</div>`;
  }
  if(ME.invite_url&&!ME.member)h+=`<a class="btn sec" href="${esc(ME.invite_url)}">Discordサーバーに参加する</a>`;
  $("#app").innerHTML=header()+h;
}
/* ---------- 修正シート ---------- */
async function openEdit(id){
  let rec;
  if(isServer())rec=(ST.recent||[]).find(x=>x.id==id);
  else rec=localStats().recent.find(x=>x.id==id);
  if(!rec)return;
  await loadDecks(rec.cls);
  let st={id,cls:rec.cls,deck:rec.deck,first:rec.first,win:rec.win};
  const draw=()=>{
    const dl=decks[st.cls]||[];
    $("#layer").innerHTML=`<div class="sheet"><div>
     <div class="lbl">記録を修正</div>
     <div class="chips">${(ME.classes||[]).map(c=>`<button class="chip ${st.cls===c?"on":""}" data-e="cls" data-v="${esc(c)}">${esc(c)}</button>`).join("")}</div>
     <div style="height:10px"></div><select data-e="deck"><option value="">（デッキなし）</option>${dl.map(d=>`<option ${st.deck===d?"selected":""}>${esc(d)}</option>`).join("")}</select>
     <div style="height:10px"></div><div class="seg"><button class="${st.first?"on":""}" data-e="first" data-v="1">先攻</button><button class="${!st.first?"on":""}" data-e="first" data-v="0">後攻</button></div>
     <div style="height:10px"></div><div class="seg"><button class="${st.win?"on":""}" data-e="win" data-v="1">勝ち</button><button class="${!st.win?"on":""}" data-e="win" data-v="0">負け</button></div>
     <div style="height:14px"></div><div class="row"><button class="btn sec" data-e="close">閉じる</button><button class="btn pri" data-e="save">保存する</button></div></div></div>`;
  };
  draw();
  const layer=$("#layer");
  layer.onclick=async(ev)=>{
    const b=ev.target.closest("[data-e]");if(!b)return;const k=b.dataset.e;
    if(k==="cls"){st.cls=b.dataset.v;st.deck="";await loadDecks(st.cls);draw()}
    else if(k==="first"){st.first=b.dataset.v==="1";draw()}
    else if(k==="win"){st.win=b.dataset.v==="1";draw()}
    else if(k==="close"){layer.innerHTML="";layer.onclick=null}
    else if(k==="save"){
      if(isServer()){const r=await api("/api/match/"+st.id,"PATCH",{opp_class:st.cls,opp_deck:st.deck||null,is_first:st.first,is_win:st.win});
        if(!r.ok){toast(r.data.error||"保存できませんでした");return}}
      else{const g=gload();const x=g.r.find(y=>y.id==st.id);if(x){x.c=st.cls;x.d=st.deck;x.f=st.first?1:0;x.w=st.win?1:0;gsave(g)}}
      layer.innerHTML="";layer.onclick=null;toast("修正しました");render();
    }
  };
  layer.addEventListener("change",(ev)=>{if(ev.target.dataset.e==="deck")st.deck=ev.target.value});
}
/* ---------- イベント ---------- */
document.addEventListener("click",async(ev)=>{
  const t=ev.target.closest("[data-tab]");
  if(t){tab=t.dataset.tab;try{localStorage.setItem("sw_tab",tab)}catch(e){}render();return}
  const b=ev.target.closest("[data-act]");if(!b)return;
  const a=b.dataset.act;
  if(a==="cls"){sel.cls=b.dataset.v;sel.deck="";await loadDecks(sel.cls);$("#app").innerHTML=header()+viewInput()}
  else if(a==="first"){sel.first=+b.dataset.v;$("#app").innerHTML=header()+viewInput()}
  else if(a==="win"){await doRecord(b.dataset.v==="1")}
  else if(a==="undo"){await doUndo()}
  else if(a==="rank"){
    const rk=ME.settings.rank[b.dataset.i];if(!rk)return;
    const r=await api("/api/rank","POST",{tier:rk.tier,grade:rk.grade});
    if(r.ok){ME.settings.rank=r.data.rank;toast("ランクを "+rk.label+" に更新しました");render()}else toast(r.data.error||"更新できませんでした");
  }
  else if(a==="edit"){await openEdit(b.dataset.id)}
  else if(a==="del"){
    if(!confirm("この記録を取り消しますか？"))return;
    if(isServer()){const r=await api("/api/match/"+b.dataset.id,"DELETE");if(!r.ok){toast(r.data.error||"取り消せませんでした");return}}
    else{const g=gload();g.r=g.r.filter(x=>String(x.id)!==b.dataset.id);gsave(g)}
    toast("取り消しました");loadStats();
  }
  else if(a==="ask"){
    b.disabled=true;
    const r=await api("/api/question","POST",{cls:$("#q-cls").value,title:$("#q-title").value,body:$("#q-body").value,attach:$("#q-attach").checked});
    b.disabled=false;
    if(r.ok){toast("✅ Discordに投稿しました");$("#q-title").value="";$("#q-body").value=""}else toast(r.data.error||"投稿できませんでした");
  }
});
document.addEventListener("change",async(ev)=>{
  const a=ev.target.dataset&&ev.target.dataset.act;
  if(a==="deck"){sel.deck=ev.target.value}
  else if(a==="rankpick"&&ev.target.value!==""){
    const l=ME.ladder[+ev.target.value];
    const r=await api("/api/rank","POST",{tier:l.tier,grade:l.grade});
    if(r.ok){ME.settings.rank=r.data.rank;toast("ランクを "+l.label+" に更新しました");render()}else toast(r.data.error||"更新できませんでした");
  }
});
init();
</script></body></html>
'''
