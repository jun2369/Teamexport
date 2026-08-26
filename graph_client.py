import base64
import re
from datetime import datetime

import requests
from bs4 import BeautifulSoup

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

HOSTED_CONTENT_RE = re.compile(
    r'https://graph\.microsoft\.com/v1\.0/chats/[^"]+/hostedContents/[^"/]+/\$value'
)


def _get(token, url):
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.json()


def list_chats(token):
    chats = []
    url = f"{GRAPH_BASE}/me/chats?$expand=members,lastMessagePreview&$top=50"
    while url:
        data = _get(token, url)
        chats.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return chats


def find_chat_by_topic(token, topic):
    for chat in list_chats(token):
        if (chat.get("topic") or "").strip() == topic:
            return chat
    return None


_my_id_cache = {}


def _get_my_id(token):
    if token not in _my_id_cache:
        _my_id_cache[token] = _get(token, f"{GRAPH_BASE}/me")["id"]
    return _my_id_cache[token]


def chat_display_name(token, chat):
    topic = (chat.get("topic") or "").strip()
    if topic:
        return topic

    my_id = _get_my_id(token)
    names = [
        m["displayName"]
        for m in chat.get("members", []) or []
        if m.get("userId") != my_id and m.get("displayName")
    ]
    if not names:
        return "(untitled chat)"
    if len(names) > 3:
        return ", ".join(names[:3]) + f" +{len(names) - 3} more"
    return ", ".join(names)


def _parse_dt(value):
    if not value:
        return None
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def _count_messages_since(token, chat_id, since, page_size=25):
    url = f"{GRAPH_BASE}/chats/{chat_id}/messages?$top={page_size}&$orderby=createdDateTime desc"
    data = _get(token, url)
    since_dt = _parse_dt(since)
    count = 0
    for raw in data.get("value", []):
        if raw.get("messageType") != "message":
            continue
        created_dt = _parse_dt(raw.get("createdDateTime"))
        if since_dt and created_dt and created_dt <= since_dt:
            break
        count += 1
    return f"{page_size}+" if count >= page_size else count


def chat_unread_info(token, chat):
    """Best-effort unread status/count based on the chat's viewpoint (last read
    time) vs its most recent message, regardless of mute state."""
    preview = chat.get("lastMessagePreview") or {}
    last_created = preview.get("createdDateTime")
    if not last_created:
        return {"unread": False, "count": None}

    last_read = (chat.get("viewpoint") or {}).get("lastMessageReadDateTime")
    if not last_read:
        # Graph frequently omits viewpoint.lastMessageReadDateTime (a known
        # API gap, especially for group chats). Without it we can't tell
        # whether the chat is actually unread, so don't guess "unread".
        return {"unread": False, "count": None}
    if _parse_dt(last_created) <= _parse_dt(last_read):
        return {"unread": False, "count": None}

    count = _count_messages_since(token, chat["id"], last_read)
    return {"unread": True, "count": count}


def _fetch_hosted_image(token, url):
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    content_type = resp.headers.get("Content-Type", "image/png")
    return resp.content, content_type


def _file_ext(name):
    return name.rsplit(".", 1)[-1].upper() if "." in name else "FILE"


def _html_to_text(html):
    soup = BeautifulSoup(html, "html.parser")
    for br in soup.find_all("br"):
        br.replace_with("\n")
    return soup.get_text().strip()


def _extract_message(token, raw):
    body = raw.get("body", {}) or {}
    html = body.get("content", "") or ""

    images = []
    for match in HOSTED_CONTENT_RE.findall(html):
        try:
            img_bytes, content_type = _fetch_hosted_image(token, match)
        except requests.HTTPError:
            continue
        b64 = base64.b64encode(img_bytes).decode("ascii")
        images.append(
            {
                "bytes_b64": b64,
                "content_type": content_type,
                "data_uri": f"data:{content_type};base64,{b64}",
            }
        )

    text = _html_to_text(html)

    attachments = [
        {"name": a.get("name") or "attachment", "url": a["contentUrl"]}
        for a in (raw.get("attachments") or [])
        if a.get("contentType") == "reference" and a.get("contentUrl")
    ]

    if attachments:
        file_notes = [
            f"[Sent a file: {att['name']} ({_file_ext(att['name'])})]" for att in attachments
        ]
        text = "\n".join(file_notes) if not text else text + "\n" + "\n".join(file_notes)

    sender = "Unknown"
    from_field = raw.get("from") or {}
    if from_field.get("user"):
        sender = from_field["user"].get("displayName", "Unknown")
    elif from_field.get("application"):
        sender = from_field["application"].get("displayName", "Unknown")

    return {
        "id": raw["id"],
        "sender": sender,
        "created": raw.get("createdDateTime", ""),
        "text": text,
        "images": images,
        "attachments": attachments,
    }


def get_messages(token, chat_id, limit=20, start=None, end=None, keyword=None):
    """Fetch messages newest-first, stopping once enough are collected.

    With no start/end, stops after `limit` messages. With a start/end range,
    keeps paging (up to a hard cap) until messages older than `start` are hit,
    since date filtering isn't supported server-side by this Graph endpoint.
    """
    messages = []
    url = f"{GRAPH_BASE}/chats/{chat_id}/messages?$top=50&$orderby=createdDateTime desc"
    ranged = bool(start or end)
    hard_cap = 1000 if ranged else limit

    while url and len(messages) < hard_cap:
        data = _get(token, url)
        stop = False
        for raw in data.get("value", []):
            if raw.get("messageType") != "message":
                continue
            html = (raw.get("body", {}) or {}).get("content", "") or ""
            if not html and not raw.get("attachments"):
                continue

            created = raw.get("createdDateTime", "")
            if start and created < start:
                stop = True
                break
            if end and created > end:
                continue

            if keyword and keyword.lower() not in _html_to_text(html).lower():
                continue

            messages.append(_extract_message(token, raw))
            if not ranged and len(messages) >= limit:
                stop = True
                break
        if stop:
            break
        url = data.get("@odata.nextLink")

    messages.sort(key=lambda m: m["created"])
    return messages


def get_messages_page(token, chat_id, top=20, next_url=None):
    """Fetch one batch of messages (newest-first page), for infinite scroll.

    Pass the `next_url` returned by a previous call to continue further back
    in history. Returns (messages_oldest_first, next_url_or_None).
    """
    url = next_url or (
        f"{GRAPH_BASE}/chats/{chat_id}/messages?$top={top}&$orderby=createdDateTime desc"
    )
    data = _get(token, url)

    messages = []
    for raw in data.get("value", []):
        if raw.get("messageType") != "message":
            continue
        if not (raw.get("body", {}) or {}).get("content"):
            continue
        messages.append(_extract_message(token, raw))

    messages.sort(key=lambda m: m["created"])
    return messages, data.get("@odata.nextLink")


def get_new_messages(token, chat_id, since, cap=100):
    """Fetch messages newer than `since` (ISO datetime), oldest-first."""
    messages = []
    url = f"{GRAPH_BASE}/chats/{chat_id}/messages?$top=50&$orderby=createdDateTime desc"

    while url and len(messages) < cap:
        data = _get(token, url)
        stop = False
        for raw in data.get("value", []):
            if raw.get("messageType") != "message":
                continue
            if not (raw.get("body", {}) or {}).get("content") and not raw.get("attachments"):
                continue
            created = raw.get("createdDateTime", "")
            if created <= since:
                stop = True
                break
            messages.append(_extract_message(token, raw))
        if stop:
            break
        url = data.get("@odata.nextLink")

    messages.sort(key=lambda m: m["created"])
    return messages
