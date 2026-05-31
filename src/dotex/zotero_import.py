"""Zotero missing-item import helpers.

The local Connector path follows Zotero's connector server import flow
(`server_connector.js`, `/connector/import`, `Zotero.Translate.Import()`), while the
Web API path uses normal API writes with `Zotero-API-Key` headers. This module
never writes to `zotero.sqlite`; Zotero's Local API source documents that local
write access is not supported, so local collection creation is intentionally a
user action in Zotero Desktop.
"""

from __future__ import annotations

import getpass
import json
import sys
import time
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, Iterable
from urllib import error, request
from urllib.parse import urlencode

from dotex.resolve_zotero import BibliographyEntry, normalize_doi, normalize_url

ZOTERO_CONNECTOR_URL = "http://127.0.0.1:23119"
CONNECTOR_API_VERSION = "2"
ZOTERO_API_BASE_URL = "https://api.zotero.org"
ZOTERO_API_VERSION = "3"
ZOTERO_SCHEMA_VERSION = "35"


@dataclass(frozen=True)
class MissingZoteroClassification:
    matched: list[object] = field(default_factory=list)
    unmatched: list[object] = field(default_factory=list)
    insufficient_metadata: list[object] = field(default_factory=list)
    duplicate_candidates: list[object] = field(default_factory=list)

    @property
    def has_missing(self) -> bool:
        return bool(self.unmatched or self.insufficient_metadata or self.duplicate_candidates)


@dataclass
class MissingZoteroHandlingResult:
    mode: str
    imported_count: int = 0
    still_unmatched_count: int = 0
    message: str = ""
    should_reresolve: bool = False


@dataclass
class ZoteroImportSession:
    selected_mode: str | None = None
    confirmed_write: bool = False
    api_key: str | None = None
    user_id: str | None = None
    username: str | None = None
    local_connector_available: bool | None = None


def classify_resolution_records(records: Iterable[object]) -> MissingZoteroClassification:
    matched: list[object] = []
    unmatched: list[object] = []
    insufficient: list[object] = []
    duplicates: list[object] = []
    for record in records:
        if getattr(record, "matched", False):
            matched.append(record)
            continue
        if getattr(record, "duplicate_candidates", None):
            duplicates.append(record)
            continue
        if not has_importable_metadata(record):
            insufficient.append(record)
            continue
        unmatched.append(record)
    return MissingZoteroClassification(matched, unmatched, insufficient, duplicates)


def has_importable_metadata(record: object) -> bool:
    return bool(
        normalize_doi(getattr(record, "source_key", ""))
        or normalize_url(getattr(record, "source_key", ""))
        or getattr(record, "parsed_title", None)
    )


def prompt_for_missing_zotero_mode(
    classification: MissingZoteroClassification,
    *,
    input_func: Callable[[str], str] = input,
    output_func: Callable[[str], None] = print,
) -> str:
    output_func("发现 bibliography 中存在未导入 Zotero 的文献。dotex 不会直接写入 zotero.sqlite。")
    output_func(f"未匹配且可导入: {len(classification.unmatched)}")
    output_func(f"元数据不足: {len(classification.insufficient_metadata)}")
    output_func(f"可能重复: {len(classification.duplicate_candidates)}")
    output_func("1) local  - 通过 Zotero 桌面端 Connector 导入；请先在 Zotero 中选中目标 collection（建议 Dotex/yy-mm-dd）。")
    output_func("2) web    - 通过 Zotero Web API 授权后导入，并创建/复用 Dotex/yy-mm-dd collection。")
    output_func("3) ignore - 不导入，继续生成未导入清单；手动导入后重新转换可获得最稳定的 Zotero 原生字段。")
    if not sys.stdin.isatty():
        output_func("非交互式终端，默认选择 ignore；不会静默导入。")
        return "ignore"
    while True:
        choice = input_func("请选择 [local/web/ignore]: ").strip().casefold()
        if choice in {"1", "local", "l"}:
            return "local"
        if choice in {"2", "web", "w"}:
            return "web"
        if choice in {"3", "ignore", "i"}:
            return "ignore"
        output_func("请输入 local、web 或 ignore。")


def handle_missing_zotero_items(
    classification: MissingZoteroClassification,
    entries: list[BibliographyEntry],
    *,
    session: ZoteroImportSession | None = None,
    input_func: Callable[[str], str] = input,
    output_func: Callable[[str], None] = print,
) -> MissingZoteroHandlingResult:
    if not classification.has_missing:
        return MissingZoteroHandlingResult(mode="none")
    session = session or DEFAULT_IMPORT_SESSION
    if session.selected_mode is None:
        session.selected_mode = prompt_for_missing_zotero_mode(classification, input_func=input_func, output_func=output_func)
    if session.selected_mode == "ignore":
        return MissingZoteroHandlingResult(
            mode="ignore",
            still_unmatched_count=len(classification.unmatched),
            message="用户选择忽略未导入文献；已保留未导入清单，手动导入 Zotero 后请重新转换。",
        )
    if session.selected_mode == "local":
        return import_missing_via_local_connector(classification.unmatched, entries, output_func=output_func, session=session)
    if session.selected_mode == "web":
        return import_missing_via_web_api(classification.unmatched, entries, output_func=output_func, session=session, input_func=input_func)
    return MissingZoteroHandlingResult(mode=session.selected_mode, still_unmatched_count=len(classification.unmatched))


def connector_is_available(base_url: str = ZOTERO_CONNECTOR_URL, timeout: float = 1.0) -> bool:
    try:
        with request.urlopen(f"{base_url}/connector/ping", timeout=timeout) as response:
            return 200 <= response.status < 300
    except Exception:
        return False


def import_missing_via_local_connector(
    missing_records: list[object],
    entries: list[BibliographyEntry],
    *,
    output_func: Callable[[str], None] = print,
    session: ZoteroImportSession | None = None,
    base_url: str = ZOTERO_CONNECTOR_URL,
) -> MissingZoteroHandlingResult:
    session = session or DEFAULT_IMPORT_SESSION
    target = date.today().strftime("Dotex/%y-%m-%d")
    output_func(f"本地 Zotero Connector 导入会写入 Zotero 当前选中的 library/collection；请在 Zotero 中手动创建并选中 {target}。")
    available = connector_is_available(base_url)
    session.local_connector_available = available
    if not available:
        return MissingZoteroHandlingResult(
            mode="local",
            still_unmatched_count=len(missing_records),
            message="Zotero Connector server 不可访问；请启动 Zotero 桌面端并确认 Connector server 已启用后重试。",
        )
    bibtex = build_bibtex_payload(missing_records, entries)
    if not bibtex.strip():
        return MissingZoteroHandlingResult(mode="local", still_unmatched_count=len(missing_records), message="没有足够元数据可导入。")
    post_connector_import(bibtex, base_url=base_url)
    return MissingZoteroHandlingResult(mode="local", imported_count=len(missing_records), should_reresolve=True)


def post_connector_import(bibtex_payload: str, *, base_url: str = ZOTERO_CONNECTOR_URL, session_id: str | None = None) -> bytes:
    session = session_id or f"dotex-{int(time.time())}"
    url = f"{base_url}/connector/import?{urlencode({'session': session})}"
    req = request.Request(
        url,
        data=bibtex_payload.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/x-bibtex; charset=utf-8",
            "X-Zotero-Connector-API-Version": CONNECTOR_API_VERSION,
        },
    )
    with request.urlopen(req, timeout=30) as response:
        return response.read()


def import_missing_via_web_api(
    missing_records: list[object],
    entries: list[BibliographyEntry],
    *,
    output_func: Callable[[str], None] = print,
    input_func: Callable[[str], str] = input,
    session: ZoteroImportSession | None = None,
    client: "ZoteroWebAPIClient | None" = None,
) -> MissingZoteroHandlingResult:
    session = session or DEFAULT_IMPORT_SESSION
    if session.api_key is None:
        output_func("需要 Zotero Web API 授权；API key 不会写入普通日志。")
        if sys.stdin.isatty():
            session.api_key = getpass.getpass("请输入 Zotero API key（输入不会显示）: ").strip()
        else:
            session.api_key = input_func("请输入 Zotero API key: ").strip()
    if not session.api_key:
        return MissingZoteroHandlingResult(mode="web", still_unmatched_count=len(missing_records), message="未提供 API key。")
    client = client or ZoteroWebAPIClient(session.api_key)
    user = client.current_user()
    session.user_id = str(user.get("userID") or user.get("id") or "")
    session.username = str(user.get("username") or "")
    target_collection = ensure_dotex_collection(client)
    unique_records = dedupe_records(missing_records)
    items = [record_to_zotero_api_item(record, target_collection) for record in unique_records]
    if not items:
        return MissingZoteroHandlingResult(mode="web", still_unmatched_count=len(missing_records), message="没有足够元数据可导入。")
    output_func(f"准备通过 Web API 导入 {len(items)} 条文献到 Dotex/{date.today().strftime('%y-%m-%d')}。")
    if not session.confirmed_write and sys.stdin.isatty():
        if input_func("确认写入 Zotero Web API? [yes/no]: ").strip().casefold() not in {"yes", "y"}:
            return MissingZoteroHandlingResult(mode="web", still_unmatched_count=len(missing_records), message="用户取消 Web API 导入。")
        session.confirmed_write = True
    client.create_items(items)
    output_func("Web API 导入完成；Zotero 桌面端可能需要同步后本地数据库才能解析到真实 item identity。")
    return MissingZoteroHandlingResult(mode="web", imported_count=len(items), should_reresolve=True)


class ZoteroWebAPIClient:
    def __init__(self, api_key: str, *, base_url: str = ZOTERO_API_BASE_URL) -> None:
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")

    def _request(self, method: str, path: str, payload: object | None = None) -> object:
        data = None if payload is None else json.dumps(payload).encode("utf-8")
        req = request.Request(
            f"{self.base_url}{path}",
            data=data,
            method=method,
            headers={
                "Content-Type": "application/json",
                "Zotero-API-Key": self.api_key,
                "Zotero-API-Version": ZOTERO_API_VERSION,
                "Zotero-Schema-Version": ZOTERO_SCHEMA_VERSION,
            },
        )
        try:
            with request.urlopen(req, timeout=30) as response:
                raw = response.read()
        except error.HTTPError as exc:
            raw = exc.read()
            raise RuntimeError(f"Zotero Web API request failed with HTTP {exc.code}: {redact_api_key(raw.decode('utf-8', errors='replace'), self.api_key)}") from exc
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def current_user(self) -> dict:
        payload = self._request("GET", "/keys/current")
        return payload if isinstance(payload, dict) else {}

    def create_login_session(self) -> dict:
        payload = self._request("POST", "/keys/sessions", {})
        return payload if isinstance(payload, dict) else {}

    def check_login_session(self, token: str) -> dict:
        payload = self._request("GET", f"/keys/sessions/{token}")
        return payload if isinstance(payload, dict) else {}

    def list_collections(self) -> list[dict]:
        payload = self._request("GET", "/users/me/collections?format=json")
        return payload if isinstance(payload, list) else []

    def create_collection(self, name: str, parent_key: str | None = None) -> dict:
        collection: dict[str, object] = {"name": name}
        if parent_key:
            collection["parentCollection"] = parent_key
        payload = self._request("POST", "/users/me/collections", [collection])
        if isinstance(payload, dict):
            successful = payload.get("successful")
            if isinstance(successful, dict) and successful:
                return next(iter(successful.values()))
        return collection

    def create_items(self, items: list[dict]) -> object:
        return self._request("POST", "/users/me/items", items)


def ensure_dotex_collection(client: ZoteroWebAPIClient) -> str:
    collections = client.list_collections()
    dotex = find_collection(collections, "Dotex", parent_key=False)
    if dotex is None:
        dotex = client.create_collection("Dotex")
    dotex_key = str(dotex.get("key") or dotex.get("collectionKey") or "")
    today_name = date.today().strftime("%y-%m-%d")
    child = find_collection(client.list_collections(), today_name, parent_key=dotex_key)
    if child is None:
        child = client.create_collection(today_name, dotex_key)
    return str(child.get("key") or child.get("collectionKey") or "")


def find_collection(collections: list[dict], name: str, parent_key: str | bool | None = None) -> dict | None:
    for collection in collections:
        data = collection.get("data", collection)
        if data.get("name") != name:
            continue
        parent = data.get("parentCollection")
        if parent_key is False and parent:
            continue
        if isinstance(parent_key, str) and parent != parent_key:
            continue
        return data
    return None


def dedupe_records(records: list[object]) -> list[object]:
    seen: set[str] = set()
    unique: list[object] = []
    for record in records:
        key = dedupe_key(record)
        if key in seen:
            continue
        seen.add(key)
        unique.append(record)
    return unique


def dedupe_key(record: object) -> str:
    return normalize_doi(getattr(record, "source_key", "")) or normalize_url(getattr(record, "source_key", "")) or str(getattr(record, "parsed_title", "")).strip().casefold() or str(getattr(record, "source_key", ""))


def build_bibtex_payload(records: list[object], entries: list[BibliographyEntry]) -> str:
    entry_by_key = {entry.source_key: entry for entry in entries}
    blocks: list[str] = []
    for index, record in enumerate(dedupe_records(records), start=1):
        source_key = str(getattr(record, "source_key", ""))
        entry = entry_by_key.get(source_key)
        title = getattr(record, "parsed_title", None) or (entry.parsed_title if entry else None) or (entry.formatted_reference if entry else source_key)
        doi = normalize_doi(source_key)
        url = normalize_url(source_key)
        key = f"dotex{index}"
        fields = {"title": title}
        if doi:
            fields["doi"] = doi
        if url:
            fields["url"] = url
        blocks.append("@article{" + key + ",\n" + ",\n".join(f"  {name} = {{{escape_bibtex(value)}}}" for name, value in fields.items() if value) + "\n}")
    return "\n\n".join(blocks) + ("\n" if blocks else "")


def record_to_zotero_api_item(record: object, collection_key: str) -> dict:
    source_key = str(getattr(record, "source_key", ""))
    item: dict[str, object] = {
        "itemType": "journalArticle" if normalize_doi(source_key) else "webpage",
        "title": getattr(record, "parsed_title", None) or getattr(record, "formatted_reference", None) or source_key,
        "collections": [collection_key] if collection_key else [],
    }
    doi = normalize_doi(source_key)
    url = normalize_url(source_key)
    if doi:
        item["DOI"] = doi
    if url:
        item["url"] = url
    return item


def escape_bibtex(value: object) -> str:
    return str(value).replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")


def redact_api_key(text: str, api_key: str | None) -> str:
    return text.replace(api_key, "[REDACTED]") if api_key else text


DEFAULT_IMPORT_SESSION = ZoteroImportSession()
