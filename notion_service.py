"""
Notion API client for the Projection Model's Notion Research integration.

Stdlib-only (urllib), same pattern excel_bridge.py already uses for Alpha
Vantage — no pip dependency added for this. Every public function returns
(ok, data_or_error) instead of raising, so callers in excel_bridge.py never
need a try/except around a Notion call; a failure just becomes a JSON error
response to the dashboard.

A Notion research page has two independently-updatable zones, so a
background notes autosave and an explicit "Sync to Notion" (numbers) can
never clobber each other:

    [Heading2] Notes                   <- notesHeadingId: stable, never replaced
    [...blocks...]                     <- notesBlockIds: deleted + re-appended
    [Divider]
    [Heading2] Projection Data — synced <timestamp>
    [...blocks...]                     <- dataBlockIds: deleted + re-appended

Notes content is parsed from lightweight markdown (typed in a plain
<textarea> — see markdown_to_blocks) so a research note can use real Notion
headings/bullets/bold instead of landing as one flat paragraph. Every
block ID above is handed back to the caller (excel_bridge.py) to persist
alongside the saved projection that owns them; this module holds no state
of its own.
"""

import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request

# Overridable so tests can point this at a local stub instead of the real
# Notion API — never changed in normal operation.
API_BASE = os.environ.get("NOTION_API_BASE", "https://api.notion.com/v1").rstrip("/")
NOTION_VERSION = "2022-06-28"


def _request(token, method, path, body=None):
    url = API_BASE + path
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
        "User-Agent": "ProjectionModel/1.0",
    })
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            raw = resp.read()
            return True, (json.loads(raw) if raw else {})
    except urllib.error.HTTPError as exc:
        raw = exc.read()
        message = str(exc)
        try:
            payload = json.loads(raw)
            message = payload.get("message") or message
        except (json.JSONDecodeError, ValueError):
            pass
        # Notion's own codes for "the token no longer works" / "not shared
        # with the integration" / "the page/db was deleted" — surfaced as a
        # status so excel_bridge.py can turn this into one clear message
        # instead of a stack trace, whichever of the three it turns out to be.
        return False, {"status": exc.code, "message": message}
    except urllib.error.URLError as exc:
        return False, {"status": 0, "message": str(getattr(exc, "reason", exc))}


def check_token(token):
    """Cheapest possible call that proves the token still works."""
    ok, data = _request(token, "GET", "/users/me")
    if not ok:
        return False, data
    return True, {"name": data.get("name") or data.get("bot", {}).get("owner", {}).get("type", "workspace")}


def list_databases(token):
    """Databases this integration has been explicitly shared with in Notion."""
    ok, data = _request(token, "POST", "/search", {
        "filter": {"property": "object", "value": "database"},
        "page_size": 50,
    })
    if not ok:
        return False, data
    out = []
    for db in data.get("results", []):
        title = "".join(t.get("plain_text", "") for t in db.get("title", [])) or "Untitled database"
        out.append({"id": db["id"], "name": title})
    out.sort(key=lambda d: d["name"].lower())
    return True, out


def _title_property_name(schema):
    for name, prop in (schema.get("properties") or {}).items():
        if prop.get("type") == "title":
            return name
    return "Name"  # every Notion database has exactly one; this is just a safe fallback


def get_database_title_property(token, database_id):
    ok, data = _request(token, "GET", f"/databases/{database_id}")
    if not ok:
        return False, data
    return True, _title_property_name(data)


# ── block builders ───────────────────────────────────────────────────────────
def _rich_text(content):
    return [{"type": "text", "text": {"content": content[:2000]}}]


def heading_block(text):
    return {"object": "block", "type": "heading_2", "heading_2": {"rich_text": _rich_text(text)}}


def paragraph_block(text):
    return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": _rich_text(text or "—")}}


def divider_block():
    return {"object": "block", "type": "divider", "divider": {}}


# ── notes: lightweight markdown -> Notion blocks ─────────────────────────────
# The dashboard's notes field is a plain <textarea> (no rich-text editor to
# ship), so this is what actually makes a typed note "look organized" once
# it lands in Notion — line-start markers become real block types, and a
# small set of inline markers become rich-text annotations within any of them.
_INLINE_RE = re.compile(
    r"\*\*(?P<bold>.+?)\*\*"
    r"|`(?P<code>.+?)`"
    r"|\*(?P<italic>.+?)\*"
    r"|\[(?P<link_text>.+?)\]\((?P<link_url>[^\s()]+)\)"
)


def _parse_inline(text):
    """**bold**, `code`, *italic*, [text](url) -> Notion rich_text spans."""
    spans, pos = [], 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            spans.append({"type": "text", "text": {"content": text[pos:m.start()]}})
        if m.group("bold") is not None:
            spans.append({"type": "text", "text": {"content": m.group("bold")}, "annotations": {"bold": True}})
        elif m.group("code") is not None:
            spans.append({"type": "text", "text": {"content": m.group("code")}, "annotations": {"code": True}})
        elif m.group("italic") is not None:
            spans.append({"type": "text", "text": {"content": m.group("italic")}, "annotations": {"italic": True}})
        else:
            spans.append({"type": "text", "text": {"content": m.group("link_text"), "link": {"url": m.group("link_url")}}})
        pos = m.end()
    if pos < len(text):
        spans.append({"type": "text", "text": {"content": text[pos:]}})
    if not spans:
        spans = [{"type": "text", "text": {"content": ""}}]
    for s in spans:
        s["text"]["content"] = s["text"]["content"][:2000]
    return spans


_HEADING_RE = re.compile(r"^(#{1,3})\s+(.*)$")
_TODO_RE = re.compile(r"^[-*]\s+\[( |x|X)\]\s+(.*)$")
_BULLET_RE = re.compile(r"^[-*]\s+(.*)$")
_NUMBERED_RE = re.compile(r"^\d+[.)]\s+(.*)$")
_QUOTE_RE = re.compile(r"^>\s?(.*)$")


def markdown_to_blocks(text):
    """Line-start markers -> block type: '# '/'## '/'### ' headings, '- '/'* '
    bullets, '1. ' numbered items, '- [ ] '/'- [x] ' to-dos, '> ' quotes,
    '---' dividers. Blank-line-separated runs of plain text become separate
    paragraphs. Never returns an empty list — Notion pages need >=1 child."""
    if not text or not text.strip():
        return [paragraph_block("")]

    blocks = []
    paragraph_buf = []

    def flush_paragraph():
        if not paragraph_buf:
            return
        content = " ".join(paragraph_buf).strip()
        if content:
            blocks.append({"object": "block", "type": "paragraph", "paragraph": {"rich_text": _parse_inline(content)}})
        paragraph_buf.clear()

    for raw_line in text.replace("\r\n", "\n").split("\n"):
        stripped = raw_line.strip()

        if not stripped:
            flush_paragraph()
            continue

        m = _HEADING_RE.match(stripped)
        if m:
            flush_paragraph()
            key = f"heading_{len(m.group(1))}"
            blocks.append({"object": "block", "type": key, key: {"rich_text": _parse_inline(m.group(2))}})
            continue

        if stripped in ("---", "***"):
            flush_paragraph()
            blocks.append(divider_block())
            continue

        m = _TODO_RE.match(stripped)
        if m:
            flush_paragraph()
            blocks.append({"object": "block", "type": "to_do",
                            "to_do": {"rich_text": _parse_inline(m.group(2)), "checked": m.group(1).lower() == "x"}})
            continue

        m = _BULLET_RE.match(stripped)
        if m:
            flush_paragraph()
            blocks.append({"object": "block", "type": "bulleted_list_item",
                            "bulleted_list_item": {"rich_text": _parse_inline(m.group(1))}})
            continue

        m = _NUMBERED_RE.match(stripped)
        if m:
            flush_paragraph()
            blocks.append({"object": "block", "type": "numbered_list_item",
                            "numbered_list_item": {"rich_text": _parse_inline(m.group(1))}})
            continue

        m = _QUOTE_RE.match(stripped)
        if m:
            flush_paragraph()
            blocks.append({"object": "block", "type": "quote", "quote": {"rich_text": _parse_inline(m.group(1))}})
            continue

        paragraph_buf.append(stripped)

    flush_paragraph()
    return blocks or [paragraph_block("")]


# ── page lookup / creation ───────────────────────────────────────────────────
def find_page_by_ticker(token, database_id, title_prop, ticker):
    ok, data = _request(token, "POST", f"/databases/{database_id}/query", {
        "filter": {"property": title_prop, "title": {"equals": ticker}},
        "page_size": 1,
    })
    if not ok:
        return False, data
    results = data.get("results") or []
    return True, (results[0] if results else None)


def get_page(token, page_id):
    return _request(token, "GET", f"/pages/{page_id}")


def _list_children(token, block_id):
    ok, data = _request(token, "GET", f"/blocks/{block_id}/children?page_size=100")
    if not ok:
        return False, data
    return True, data.get("results") or []


def append_notes_zone(token, page_id, notes_text):
    """For a page this integration didn't create (found by ticker but never
    tracked) — adds a Notes heading + parsed content at the end and hands
    back the heading id (stable) and content block ids (replaced on every
    later notes autosave) so future edits can target this page directly."""
    ok, result = _request(token, "PATCH", f"/blocks/{page_id}/children", {
        "children": [heading_block("Notes"), *markdown_to_blocks(notes_text)],
    })
    if not ok:
        return False, result
    kids = result.get("results", [])
    return True, {
        "notesHeadingId": kids[0]["id"] if kids else None,
        "notesBlockIds": [b["id"] for b in kids[1:]] if len(kids) > 1 else [],
    }


def create_stock_page(token, database_id, title_prop, ticker, notes_text, data_blocks, synced_label):
    """First-ever sync for this ticker. Returns page id/url plus the tracked
    block ids, read back from the created page's children."""
    notes_content = markdown_to_blocks(notes_text)
    children = [
        heading_block("Notes"),
        *notes_content,
        divider_block(),
        heading_block(f"Projection Data — {synced_label}"),
        *data_blocks,
    ]
    ok, page = _request(token, "POST", "/pages", {
        "parent": {"database_id": database_id},
        "properties": {title_prop: {"title": _rich_text(ticker)}},
        "children": children,
    })
    if not ok:
        return False, page

    ok, kids = _list_children(token, page["id"])
    if not ok:
        # Page exists but we couldn't read back block ids for future targeted
        # updates — not fatal, the next full sync just rebuilds from scratch.
        kids = []
    n = len(notes_content)
    notes_heading_id = kids[0]["id"] if kids else None
    notes_block_ids = [b["id"] for b in kids[1:1 + n]] if len(kids) > 1 else []
    # dataBlockIds includes the "Projection Data — <timestamp>" heading
    # itself (index 1+n+1, right after the divider) so replace_data_section()
    # below deletes and recreates that heading too on every sync — otherwise
    # the old heading (with a stale timestamp) would never be tracked and
    # would sit there forever, one more of them piling up on every re-sync.
    data_block_ids = [b["id"] for b in kids[1 + n + 1:]] if len(kids) > 1 + n + 1 else []
    return True, {
        "pageId": page["id"],
        "url": page.get("url") or f"https://notion.so/{page['id'].replace('-', '')}",
        "notesHeadingId": notes_heading_id,
        "notesBlockIds": notes_block_ids,
        "dataBlockIds": data_block_ids,
    }


def replace_data_section(token, page_id, old_block_ids, data_blocks, synced_label):
    """Wholesale-replace the Projection Data blocks, leaving the Notes zone
    (heading, content, divider) and everything above them untouched."""
    for block_id in old_block_ids or []:
        # Best-effort — a block the user already deleted by hand is fine to skip.
        _request(token, "DELETE", f"/blocks/{block_id}")

    ok, result = _request(token, "PATCH", f"/blocks/{page_id}/children", {
        "children": [heading_block(f"Projection Data — {synced_label}"), *data_blocks],
    })
    if not ok:
        return False, result
    # Track the heading's id too (not just the data blocks under it) so the
    # *next* sync deletes this timestamp along with everything else instead
    # of leaving it behind as an orphan.
    new_ids = [b["id"] for b in result.get("results", [])]
    return True, new_ids


def replace_notes_zone(token, page_id, notes_heading_id, old_block_ids, notes_text):
    """Wholesale-replace the Notes content (everything parsed from the typed
    markdown), leaving the "Notes" heading itself and the Projection Data
    section untouched — same delete-then-append pattern as the data section,
    just re-inserted right after the heading via `after` instead of at the
    very end of the page."""
    for block_id in old_block_ids or []:
        _request(token, "DELETE", f"/blocks/{block_id}")

    body = {"children": markdown_to_blocks(notes_text)}
    if notes_heading_id:
        body["after"] = notes_heading_id
    ok, result = _request(token, "PATCH", f"/blocks/{page_id}/children", body)
    if not ok:
        return False, result
    return True, [b["id"] for b in result.get("results", [])]


def error_message(err):
    """A stable, user-facing message regardless of which Notion failure mode
    this was — token revoked, database unshared, or the page/block gone."""
    if not isinstance(err, dict):
        return str(err)
    status = err.get("status")
    detail = err.get("message") or "Unknown error."
    if status in (401, 403):
        return ("Notion rejected the connection — the integration token may have been "
                "revoked, or this database isn't shared with it anymore. " + detail)
    if status == 404:
        return "That Notion page or database no longer exists. " + detail
    if status == 0:
        return "Couldn't reach Notion — check the network and try again."
    return detail
