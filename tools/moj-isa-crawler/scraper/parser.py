from __future__ import annotations

import json
import re
from dataclasses import dataclass
from urllib.parse import urldefrag, urljoin, urlparse

from bs4 import BeautifulSoup, Tag


CONTENT_SELECTORS = [
    "#contentsArea",
    "#contents_area",
    "#contents",
    "#main",
    "#main_area",
    "main",
    "article",
    "body",
]
BREADCRUMB_SELECTORS = [
    ".breadcrumb",
    "#breadcrumb",
    ".topicPath",
    "#topicPath",
    ".topicpath",
    ".pankuzu",
    "#pankuzu",
    "nav[aria-label='breadcrumb']",
]
HEADING_TAGS = ["h1", "h2", "h3", "h4", "h5", "h6"]
RESOURCE_EXTENSIONS = {
    ".pdf",
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".svg",
    ".webp",
    ".ico",
    ".css",
    ".js",
    ".zip",
    ".lzh",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".csv",
    ".mp3",
    ".mp4",
    ".wmv",
}


@dataclass(frozen=True)
class ExtractedLink:
    text: str
    url: str
    category: str


def normalize_space(text: object) -> str:
    value = str(text or "").replace("\ufeff", "").replace("\u200b", "")
    value = value.replace("\xa0", " ").replace("　", " ")
    return re.sub(r"[ \t\r\f\v]+", " ", value).strip()


def normalize_block_text(text: object) -> str:
    lines = [normalize_space(line) for line in str(text or "").splitlines()]
    return "\n".join(line for line in lines if line)


def parse_html(html: str) -> BeautifulSoup:
    return BeautifulSoup(html, "html.parser")


def content_root(soup: BeautifulSoup | Tag) -> Tag:
    for selector in CONTENT_SELECTORS:
        found = soup.select_one(selector)
        if found is not None:
            return found
    if isinstance(soup, Tag):
        return soup
    raise ValueError("No content root found")


def clean_for_text(root: Tag) -> Tag:
    copied = BeautifulSoup(str(root), "html.parser")
    working = copied.find(root.name) or copied
    for node in working.select("script, style, noscript, template"):
        node.decompose()
    return working


def page_title(soup: BeautifulSoup) -> str:
    root = content_root(soup)
    h1 = root.find("h1")
    if h1 is not None:
        return normalize_space(h1.get_text(" ", strip=True))
    if soup.title is not None:
        return normalize_space(soup.title.get_text(" ", strip=True))
    return ""


def h1_text(soup: BeautifulSoup) -> str:
    root = content_root(soup)
    h1 = root.find("h1")
    return normalize_space(h1.get_text(" ", strip=True)) if h1 else ""


def breadcrumb_text(soup: BeautifulSoup) -> str:
    for selector in BREADCRUMB_SELECTORS:
        found = soup.select_one(selector)
        if found is not None:
            return normalize_space(found.get_text(" > ", strip=True))
    return ""


def headings_json(soup: BeautifulSoup) -> str:
    root = content_root(soup)
    headings = []
    for node in root.find_all(HEADING_TAGS):
        text = normalize_space(node.get_text(" ", strip=True))
        if text:
            headings.append({"level": int(node.name[1]), "text": text})
    return json.dumps(headings, ensure_ascii=False)


def visible_body_text(soup: BeautifulSoup) -> tuple[str, int]:
    root = clean_for_text(content_root(soup))
    text = normalize_block_text(root.get_text("\n", strip=True))
    return text, len(text)


def tables_json(soup: BeautifulSoup) -> tuple[str, int]:
    root = content_root(soup)
    tables = []
    for table in root.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = [normalize_space(cell.get_text(" ", strip=True)) for cell in tr.find_all(["th", "td"])]
            if any(cells):
                rows.append(cells)
        if rows:
            tables.append(rows)
    return json.dumps(tables, ensure_ascii=False), len(tables)


def is_pdf_url(url: str) -> bool:
    path = urlparse(url).path.lower()
    return ".pdf" in path


def has_resource_extension(url: str) -> bool:
    path = urlparse(url).path.lower()
    return any(path.endswith(ext) for ext in RESOURCE_EXTENSIONS)


def normalize_url(base_url: str, href: str) -> str:
    absolute = urljoin(base_url, href.strip())
    absolute, _fragment = urldefrag(absolute)
    return absolute


def is_internal(url: str, start_netloc: str, same_netloc_only: bool) -> bool:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        return False
    if same_netloc_only and parsed.netloc != start_netloc:
        return False
    return True


def is_allowed_by_prefix(url: str, allowed_prefixes: list[str]) -> bool:
    return any(url.startswith(prefix) for prefix in allowed_prefixes)


def classify_url(url: str, *, start_netloc: str, allowed_prefixes: list[str], same_netloc_only: bool) -> str:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        return "skipped_scheme"
    if is_pdf_url(url):
        return "pdf"
    if not is_internal(url, start_netloc, same_netloc_only):
        return "external"
    if not is_allowed_by_prefix(url, allowed_prefixes):
        return "internal_out_of_scope"
    if has_resource_extension(url):
        return "internal_resource"
    return "internal_page"


def extract_links(
    soup: BeautifulSoup,
    base_url: str,
    *,
    start_netloc: str,
    allowed_prefixes: list[str],
    same_netloc_only: bool,
) -> list[ExtractedLink]:
    root = content_root(soup)
    links: list[ExtractedLink] = []
    seen: set[tuple[str, str]] = set()
    for a in root.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href or href.startswith("#"):
            continue
        url = normalize_url(base_url, href)
        text = normalize_space(a.get_text(" ", strip=True))
        category = classify_url(url, start_netloc=start_netloc, allowed_prefixes=allowed_prefixes, same_netloc_only=same_netloc_only)
        key = (url, text)
        if key in seen:
            continue
        seen.add(key)
        links.append(ExtractedLink(text=text, url=url, category=category))
    return links


def safe_filename_from_url(url: str) -> str:
    path = urlparse(url).path
    name = path.rsplit("/", 1)[-1] or "download.pdf"
    name = re.sub(r"[^0-9A-Za-z._-]+", "_", name)
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name
