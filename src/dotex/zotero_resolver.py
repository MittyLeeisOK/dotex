from __future__ import annotations

import json
import re
import shutil
import sqlite3
import tempfile
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from collections.abc import Iterator
from pathlib import Path
from urllib.parse import unquote, urlparse


@dataclass
class BibliographyEntry:
    source_key: str
    formatted_reference: str
    parsed_title: str | None


@dataclass
class ZoteroItem:
    item_id: int
    item_key: str
    item_type: str
    uri: str | None
    fields: dict[str, str]
    creators: list[dict[str, str]]

    def to_csl_json(self) -> dict:
        csl_item: dict[str, object] = {
            "id": self.item_key,
            "type": map_zotero_type_to_csl(self.item_type),
            "title": self.fields.get("title", ""),
        }

        if self.creators:
            csl_item["author"] = self.creators

        if self.fields.get("publicationTitle"):
            csl_item["container-title"] = self.fields["publicationTitle"]
        if self.fields.get("volume"):
            csl_item["volume"] = self.fields["volume"]
        if self.fields.get("issue"):
            csl_item["issue"] = self.fields["issue"]
        if self.fields.get("pages"):
            csl_item["page"] = self.fields["pages"]
        if self.fields.get("publisher"):
            csl_item["publisher"] = self.fields["publisher"]
        if self.fields.get("place"):
            csl_item["publisher-place"] = self.fields["place"]
        if self.fields.get("url"):
            csl_item["URL"] = self.fields["url"]
        if self.fields.get("DOI"):
            csl_item["DOI"] = self.fields["DOI"]

        year_match = re.search(r"(19|20)\d{2}", self.fields.get("date", ""))
        if year_match:
            csl_item["issued"] = {"date-parts": [[int(year_match.group(0))]]}

        return csl_item


@dataclass
class ResolutionRecord:
    source_key: str
    formatted_reference: str
    parsed_title: str | None
    matched: bool
    matched_by: str | None
    zotero_item_id: int | None
    zotero_item_key: str | None
    zotero_item_type: str | None
    zotero_title: str | None
    zotero_url: str | None
    zotero_doi: str | None
    zotero_uri: str | None


@dataclass
class ZoteroResolutionReport:
    bibliography_path: str
    zotero_database: str
    total_entries: int
    matched_entries: int
    unmatched_entries: int
    records: list[ResolutionRecord]

    def to_json(self) -> str:
        payload = asdict(self)
        return json.dumps(payload, ensure_ascii=False, indent=2)


def resolve_bibliography_against_zotero(
    bibliography_path: Path,
    zotero_database: Path,
) -> tuple[ZoteroResolutionReport, list[dict]]:
    entries = parse_bibliography_entries(bibliography_path)
    zotero_items = load_zotero_items(zotero_database)

    doi_index = {
        normalize_doi(item.fields["DOI"]): item
        for item in zotero_items
        if item.fields.get("DOI") and normalize_doi(item.fields["DOI"])
    }
    url_index = {
        normalize_url(item.fields["url"]): item
        for item in zotero_items
        if item.fields.get("url") and normalize_url(item.fields["url"])
    }
    title_index: dict[str, list[ZoteroItem]] = {}
    for item in zotero_items:
        title = normalize_title(item.fields.get("title", ""))
        if not title:
            continue
        title_index.setdefault(title, []).append(item)

    records: list[ResolutionRecord] = []
    matched_items: dict[str, ZoteroItem] = {}

    for entry in entries:
        matched_by: str | None = None
        matched_item: ZoteroItem | None = None

        entry_doi = normalize_doi(entry.source_key)
        entry_url = normalize_url(entry.source_key)
        if entry_doi and entry_doi in doi_index:
            matched_item = doi_index[entry_doi]
            matched_by = "doi"
        elif entry_url and entry_url in url_index:
            matched_item = url_index[entry_url]
            matched_by = "url"
        elif entry.parsed_title:
            candidates = title_index.get(normalize_title(entry.parsed_title), [])
            if len(candidates) == 1:
                matched_item = candidates[0]
                matched_by = "title"

        if matched_item is not None:
            matched_items[matched_item.item_key] = matched_item

        records.append(
            ResolutionRecord(
                source_key=entry.source_key,
                formatted_reference=entry.formatted_reference,
                parsed_title=entry.parsed_title,
                matched=matched_item is not None,
                matched_by=matched_by,
                zotero_item_id=matched_item.item_id if matched_item else None,
                zotero_item_key=matched_item.item_key if matched_item else None,
                zotero_item_type=matched_item.item_type if matched_item else None,
                zotero_title=matched_item.fields.get("title") if matched_item else None,
                zotero_url=matched_item.fields.get("url") if matched_item else None,
                zotero_doi=matched_item.fields.get("DOI") if matched_item else None,
                zotero_uri=matched_item.uri if matched_item else None,
            )
        )

    report = ZoteroResolutionReport(
        bibliography_path=str(bibliography_path),
        zotero_database=str(zotero_database),
        total_entries=len(entries),
        matched_entries=sum(1 for record in records if record.matched),
        unmatched_entries=sum(1 for record in records if not record.matched),
        records=records,
    )
    csl_items = [item.to_csl_json() for item in matched_items.values()]
    return report, csl_items


def parse_bibliography_entries(bibliography_path: Path) -> list[BibliographyEntry]:
    text = bibliography_path.read_text(encoding="utf-8")
    entries: list[BibliographyEntry] = []
    token = "\\bibentry"
    index = 0
    while True:
        start = text.find(token, index)
        if start == -1:
            break
        cursor = start + len(token)
        try:
            cursor = skip_whitespace(text, cursor)
            source_key, cursor = read_braced(text, cursor)
            cursor = skip_whitespace(text, cursor)
            formatted_reference, cursor = read_braced(text, cursor)
        except ValueError:
            index = start + len(token)
            continue
        normalized_reference = convert_inline_tex_to_plain(formatted_reference)
        entries.append(
            BibliographyEntry(
                source_key=source_key.strip(),
                formatted_reference=normalized_reference,
                parsed_title=extract_reference_title(normalized_reference),
            )
        )
        index = cursor
    if entries:
        return entries

    refs_display = load_refs_display_map(bibliography_path)
    for key, entry_text in parse_bibtex_entry_blocks(text):
        title = clean_bibtex_field_text(extract_bibtex_field(entry_text, "title") or "")
        doi = clean_bibtex_field_text(extract_bibtex_field(entry_text, "doi") or "")
        url = clean_bibtex_field_text(extract_bibtex_field(entry_text, "url") or "")
        author = clean_bibtex_field_text(extract_bibtex_field(entry_text, "author") or "")
        year = clean_bibtex_field_text(extract_bibtex_field(entry_text, "year") or "")
        source_key = doi or url or key
        formatted_reference = refs_display.get(key) or synthesize_bibtex_display(author, year, title, key)
        entries.append(
            BibliographyEntry(
                source_key=source_key.strip(),
                formatted_reference=formatted_reference.strip(),
                parsed_title=title or None,
            )
        )
    return entries


def load_refs_display_map(bibliography_path: Path) -> dict[str, str]:
    refs_display_path = bibliography_path.with_name("refs_display.json")
    if not refs_display_path.exists():
        return {}
    try:
        payload = json.loads(refs_display_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return {str(key): str(value) for key, value in payload.items()}


def parse_bibtex_entry_blocks(text: str) -> list[tuple[str, str]]:
    entries: list[tuple[str, str]] = []
    for match in re.finditer(r"@(\w+)\{([^,\s{}'\"]+),", text):
        key = match.group(2).strip()
        entry_start = match.end()
        next_entry = text.find("@", entry_start)
        entry_text = text[entry_start : next_entry if next_entry != -1 else len(text)]
        entries.append((key, entry_text))
    return entries


def extract_bibtex_field(entry_text: str, field_name: str) -> str | None:
    match = re.search(rf"\b{re.escape(field_name)}\s*=\s*", entry_text, re.IGNORECASE)
    if match is None:
        return None
    cursor = skip_whitespace(entry_text, match.end())
    if cursor >= len(entry_text):
        return None
    if entry_text[cursor] == "{":
        value, _ = read_braced(entry_text, cursor)
        return value
    if entry_text[cursor] == '"':
        cursor += 1
        value_start = cursor
        while cursor < len(entry_text):
            if entry_text[cursor] == '"' and not is_escaped(entry_text, cursor):
                return entry_text[value_start:cursor]
            cursor += 1
        return entry_text[value_start:]
    value_start = cursor
    while cursor < len(entry_text) and entry_text[cursor] not in ",\n\r":
        cursor += 1
    return entry_text[value_start:cursor]


def clean_bibtex_field_text(value: str) -> str:
    current = value.strip()
    current = current.replace("\n", " ")
    current = current.replace("\r", " ")
    current = re.sub(r"\\&", "&", current)
    current = re.sub(r"[{}]", "", current)
    current = re.sub(r"\s+", " ", current)
    return current.strip()


def synthesize_bibtex_display(author: str, year: str, title: str, key: str) -> str:
    normalized_author = author.replace(" and ", "; ").strip()
    if normalized_author and year:
        return f"{normalized_author} {year}".strip()
    if title and year:
        return f"{title} {year}".strip()
    return key


def load_zotero_items(zotero_database: Path) -> list[ZoteroItem]:
    connection = sqlite3.connect(f"file:{zotero_database.expanduser()}?mode=ro", uri=True)
    try:
        cursor = connection.cursor()
        user_id = load_zotero_user_id(cursor)
        field_rows = cursor.execute(
            """
            select items.itemID, items.key, itemTypes.typeName, fields.fieldName, itemDataValues.value
            from items
            join itemTypes on itemTypes.itemTypeID = items.itemTypeID
            join itemData on itemData.itemID = items.itemID
            join fields on fields.fieldID = itemData.fieldID
            join itemDataValues on itemDataValues.valueID = itemData.valueID
            where fields.fieldName in (
                'title', 'date', 'url', 'DOI', 'publicationTitle', 'pages', 'volume', 'issue',
                'publisher', 'place', 'bookTitle', 'proceedingsTitle', 'websiteTitle'
            )
            """
        ).fetchall()
        creator_rows = cursor.execute(
            """
            select items.itemID, creators.firstName, creators.lastName, creators.fieldMode, itemCreators.orderIndex
            from items
            join itemCreators on itemCreators.itemID = items.itemID
            join creators on creators.creatorID = itemCreators.creatorID
            order by items.itemID, itemCreators.orderIndex
            """
        ).fetchall()
    finally:
        connection.close()

    items: dict[int, ZoteroItem] = {}
    for item_id, item_key, item_type, field_name, field_value in field_rows:
        item = items.setdefault(
            item_id,
            ZoteroItem(
                item_id=item_id,
                item_key=item_key,
                item_type=item_type,
                uri=build_zotero_item_uri(item_key, user_id),
                fields={},
                creators=[],
            ),
        )
        item.fields[field_name] = field_value

    for item_id, first_name, last_name, field_mode, _order_index in creator_rows:
        item = items.get(item_id)
        if item is None:
            continue
        if field_mode == 1 or not first_name:
            item.creators.append({"literal": last_name})
        else:
            item.creators.append({"family": last_name, "given": first_name})

    return list(items.values())


def load_zotero_user_id(cursor: sqlite3.Cursor) -> str | None:
    row = cursor.execute(
        """
        select value from settings
        where setting='account' and key='userID'
        limit 1
        """
    ).fetchone()
    if row is None or row[0] in {None, ""}:
        return None
    return str(row[0])


def build_zotero_item_uri(item_key: str, user_id: str | None) -> str | None:
    if user_id:
        return f"http://zotero.org/users/{user_id}/items/{item_key}"
    if item_key:
        return f"http://zotero.org/users/local/items/{item_key}"
    return None


@contextmanager
def copied_zotero_database(zotero_database: Path) -> Iterator[Path]:
    source = zotero_database.expanduser()
    with tempfile.TemporaryDirectory(prefix="dotex-zotero-") as temp_dir:
        target = Path(temp_dir) / source.name
        shutil.copy2(source, target)
        for suffix in ("-wal", "-shm"):
            companion = source.with_name(source.name + suffix)
            if companion.exists():
                shutil.copy2(companion, target.with_name(target.name + suffix))
        yield target


def map_zotero_type_to_csl(item_type: str) -> str:
    mapping = {
        "blogPost": "post-weblog",
        "book": "book",
        "bookSection": "chapter",
        "conferencePaper": "paper-conference",
        "journalArticle": "article-journal",
        "magazineArticle": "article-magazine",
        "newspaperArticle": "article-newspaper",
        "report": "report",
        "thesis": "thesis",
        "webpage": "webpage",
    }
    return mapping.get(item_type, "article")


def extract_reference_title(reference_text: str) -> str | None:
    for pattern in [r"‘(.+?)’(?=,|\.|$)", r'“(.+?)”(?=,|\.|$)', r'"(.+?)"(?=,|\.|$)', r"'(.+?)'(?=,|\.|$)"]:
        match = re.search(pattern, reference_text)
        if match:
            return match.group(1).strip()
    return None


def normalize_doi(value: str | None) -> str | None:
    if not value:
        return None
    normalized = unquote(value.replace("\\%", "%")).strip()
    normalized = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", normalized, flags=re.I)
    normalized = re.sub(r"^doi:\s*", "", normalized, flags=re.I)
    normalized = normalized.strip().strip("/")
    return normalized.lower() or None


def normalize_url(value: str | None) -> str | None:
    if not value:
        return None
    decoded = unquote(value.replace("\\%", "%")).strip()
    if normalize_doi(decoded):
        decoded = re.sub(r"^https?://(?:dx\.)?doi\.org/", "https://doi.org/", decoded, flags=re.I)
    parsed = urlparse(decoded)
    if not parsed.scheme and not parsed.netloc:
        return None
    netloc = parsed.netloc.lower()
    path = parsed.path.rstrip("/")
    query = f"?{parsed.query}" if parsed.query else ""
    return f"{parsed.scheme.lower()}://{netloc}{path}{query}"


def normalize_title(value: str | None) -> str:
    if not value:
        return ""
    normalized = value.casefold()
    normalized = re.sub(r"[^\w\s]", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def convert_inline_tex_to_plain(text: str) -> str:
    current = text
    previous = None
    while current != previous:
        previous = current
        current = replace_command_one_arg(current, "emph", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "textbf", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "nolinkurl", lambda value: value)
        current = replace_command_one_arg(current, "url", lambda value: value)
    replacements = {
        "\\%": "%",
        "\\_": "_",
        "\\&": "&",
        "\\#": "#",
        "~": " ",
    }
    for old, new in replacements.items():
        current = current.replace(old, new)
    current = re.sub(r"\s+", " ", current).strip()
    return current


def replace_command_one_arg(text: str, command: str, replacer) -> str:
    token = f"\\{command}"
    chunks: list[str] = []
    index = 0
    while index < len(text):
        if text.startswith(token, index) and (index + len(token) == len(text) or not text[index + len(token)].isalpha()):
            cursor = skip_whitespace(text, index + len(token))
            if cursor >= len(text) or text[cursor] != "{":
                chunks.append(text[index])
                index += 1
                continue
            try:
                value, cursor = read_braced(text, cursor)
            except ValueError:
                chunks.append(text[index])
                index += 1
                continue
            chunks.append(replacer(value))
            index = cursor
            continue
        chunks.append(text[index])
        index += 1
    return "".join(chunks)


def skip_whitespace(text: str, index: int) -> int:
    while index < len(text) and text[index].isspace():
        index += 1
    return index


def read_braced(text: str, start: int) -> tuple[str, int]:
    if start >= len(text) or text[start] != "{":
        raise ValueError("expected opening brace")

    depth = 0
    parts: list[str] = []
    index = start
    while index < len(text):
        char = text[index]
        if char == "{" and not is_escaped(text, index):
            depth += 1
            if depth > 1:
                parts.append(char)
        elif char == "}" and not is_escaped(text, index):
            depth -= 1
            if depth == 0:
                return "".join(parts), index + 1
            parts.append(char)
        else:
            parts.append(char)
        index += 1
    raise ValueError("unclosed brace")


def is_escaped(text: str, index: int) -> bool:
    slash_count = 0
    probe = index - 1
    while probe >= 0 and text[probe] == "\\":
        slash_count += 1
        probe -= 1
    return slash_count % 2 == 1
