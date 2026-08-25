import secrets
from datetime import datetime, timedelta

from flask import Flask, jsonify, redirect, render_template, request, send_file, session, url_for

import auth
import config
import graph_client
from export_docx import build_docx, build_txt

app = Flask(__name__)
app.secret_key = config.FLASK_SECRET_KEY
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=90)

MAX_DATE_RANGE_DAYS = 20


def _sid():
    if "sid" not in session:
        session["sid"] = secrets.token_urlsafe(16)
    session.permanent = True
    return session["sid"]

# Simple in-memory cache so the export step doesn't need to re-fetch
# everything (and re-download every image) after the user picks checkboxes.
# Keyed by (sid, chat_id) -> {message_id: message} so one browser session's
# cached messages are never visible to another. Entries accumulate across
# the plain/paged/refresh loads for a chat rather than being replaced.
_message_cache = {}

# Pagination cursor for "load more / infinite scroll" per (sid, chat_id).
# Holds the Graph @odata.nextLink to continue from, or None once history is
# exhausted.
_next_cursor = {}


def _serialize(messages):
    return [
        {
            "id": m["id"],
            "sender": m["sender"],
            "created": m["created"],
            "text": m["text"],
            "image_data_uris": [img["data_uri"] for img in m["images"]],
            "attachments": m["attachments"],
        }
        for m in messages
    ]


def _cache_messages(key, messages):
    _message_cache.setdefault(key, {}).update({m["id"]: m for m in messages})


@app.route("/")
def index():
    sid = _sid()
    user = auth.current_user(sid)
    return render_template(
        "index.html",
        default_topic=config.DEFAULT_CHAT_TOPIC,
        username=user["username"] if user else None,
    )


@app.route("/login")
def login():
    _sid()
    flow = auth.build_auth_code_flow()
    session["flow"] = flow
    return redirect(flow["auth_uri"])


@app.route(config.REDIRECT_PATH)
def auth_callback():
    sid = _sid()
    flow = session.pop("flow", {})
    result = auth.acquire_token_by_flow(sid, flow, request.args)
    if "error" in result:
        return f"Login failed: {result.get('error_description', result['error'])}", 400
    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    auth.logout(_sid())
    session.clear()
    return redirect(url_for("index"))


@app.route("/api/chats")
def api_chats():
    token = auth.get_access_token(_sid())
    if not token:
        return jsonify({"error": "login_required"}), 401
    chats = graph_client.list_chats(token)
    result = []
    for c in chats:
        if c.get("chatType") not in ("group", "oneOnOne", "meeting"):
            continue
        unread = graph_client.chat_unread_info(token, c)
        result.append(
            {
                "id": c["id"],
                "topic": graph_client.chat_display_name(token, c),
                "chat_type": c.get("chatType"),
                "unread": unread["unread"],
                "unread_count": unread["count"],
            }
        )
    return jsonify(result)


@app.route("/api/messages")
def api_messages():
    chat_id = request.args.get("chat_id")
    if not chat_id:
        return jsonify({"error": "chat_id is required"}), 400

    start_date = request.args.get("start")
    end_date = request.args.get("end")
    keyword = request.args.get("keyword") or None

    sid = _sid()
    token = auth.get_access_token(sid)
    if not token:
        return jsonify({"error": "login_required"}), 401
    key = (sid, chat_id)

    if start_date or end_date:
        if start_date and end_date:
            days = (datetime.fromisoformat(end_date) - datetime.fromisoformat(start_date)).days
            if days > MAX_DATE_RANGE_DAYS:
                return jsonify({"error": f"Date range cannot exceed {MAX_DATE_RANGE_DAYS} days"}), 400

        start = f"{start_date}T00:00:00Z" if start_date else None
        end = f"{end_date}T23:59:59Z" if end_date else None
        messages = graph_client.get_messages(token, chat_id, start=start, end=end, keyword=keyword)
        _next_cursor.pop(key, None)
        has_more = False
    else:
        limit = int(request.args.get("limit", 20))
        messages, next_url = graph_client.get_messages_page(token, chat_id, top=limit)
        if keyword:
            messages = [m for m in messages if keyword.lower() in m["text"].lower()]
            has_more = False
        else:
            _next_cursor[key] = next_url
            has_more = bool(next_url)

    _cache_messages(key, messages)
    return jsonify({"messages": _serialize(messages), "has_more": has_more})


@app.route("/api/messages/more")
def api_messages_more():
    chat_id = request.args.get("chat_id")
    if not chat_id:
        return jsonify({"error": "chat_id is required"}), 400

    sid = _sid()
    key = (sid, chat_id)
    next_url = _next_cursor.get(key)
    if not next_url:
        return jsonify({"messages": [], "has_more": False})

    token = auth.get_access_token(sid)
    if not token:
        return jsonify({"error": "login_required"}), 401
    messages, new_next_url = graph_client.get_messages_page(token, chat_id, next_url=next_url)
    _next_cursor[key] = new_next_url

    _cache_messages(key, messages)
    return jsonify({"messages": _serialize(messages), "has_more": bool(new_next_url)})


@app.route("/api/messages/refresh")
def api_messages_refresh():
    chat_id = request.args.get("chat_id")
    if not chat_id:
        return jsonify({"error": "chat_id is required"}), 400

    sid = _sid()
    key = (sid, chat_id)
    since = request.args.get("since") or None
    token = auth.get_access_token(sid)
    if not token:
        return jsonify({"error": "login_required"}), 401

    if since:
        messages = graph_client.get_new_messages(token, chat_id, since)
    else:
        messages, next_url = graph_client.get_messages_page(token, chat_id, top=20)
        _next_cursor[key] = next_url

    _cache_messages(key, messages)
    return jsonify({"messages": _serialize(messages)})


@app.route("/api/export", methods=["POST"])
def api_export():
    payload = request.get_json()
    chat_id = payload.get("chat_id")
    chat_topic = payload.get("chat_topic", "Teams Export")
    message_ids = payload.get("message_ids", [])
    fmt = payload.get("format", "docx")

    cached = _message_cache.get((_sid(), chat_id), {})
    selected = [cached[mid] for mid in message_ids if mid in cached]
    selected.sort(key=lambda m: m["created"])

    if fmt == "txt":
        buf = build_txt(selected, chat_topic)
        return send_file(
            buf, as_attachment=True, download_name="teams_export.txt", mimetype="text/plain"
        )

    buf = build_docx(selected, chat_topic)
    return send_file(
        buf,
        as_attachment=True,
        download_name="teams_export.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    print("Starting server at http://127.0.0.1:5000")
    app.run(port=5000, debug=False)
