from __future__ import annotations

import dataclasses
import hashlib
import json
import os
import time
from pathlib import Path
from typing import Callable

import requests


DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/123.0 Safari/537.36"
)


@dataclasses.dataclass
class FetchResult:
    url: str
    final_url: str
    status_code: int
    content_type: str
    content_length: str
    last_modified: str
    content: bytes

    def text(self) -> str:
        # moj.go.jp pages are UTF-8. Fall back gently only for old pages that
        # omit charset in the HTTP header.
        return self.content.decode("utf-8", errors="replace")


@dataclasses.dataclass
class DownloadResult:
    url: str
    final_url: str
    status_code: int
    content_type: str
    content_length: str
    last_modified: str
    sha256: str
    saved_path: Path
    bytes_written: int
    response_headers_json: str
    first_bytes_hex: str
    post_write_size: int
    post_write_readback_ok: bool


class FetchError(RuntimeError):
    pass


class Fetcher:
    def __init__(self, *, timeout: int, sleep: float, retries: int, user_agent: str = DEFAULT_USER_AGENT) -> None:
        self.timeout = timeout
        self.sleep = sleep
        self.retries = retries
        self.session = requests.Session()
        self.headers = {
            "User-Agent": user_agent,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        }

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        last_error: Exception | None = None
        for attempt in range(self.retries + 1):
            try:
                response = self.session.request(method, url, headers=self.headers, timeout=self.timeout, **kwargs)
                response.raise_for_status()
                return response
            except requests.RequestException as exc:
                last_error = exc
                if attempt >= self.retries:
                    break
                time.sleep(min(2.0, 0.5 * (attempt + 1)))
            finally:
                if self.sleep > 0:
                    time.sleep(self.sleep)
        raise FetchError(f"{method} {url} failed: {last_error}")

    def get(self, url: str) -> FetchResult:
        response = self._request("GET", url)
        return FetchResult(
            url=url,
            final_url=response.url,
            status_code=response.status_code,
            content_type=response.headers.get("Content-Type", ""),
            content_length=response.headers.get("Content-Length", ""),
            last_modified=response.headers.get("Last-Modified", ""),
            content=response.content,
        )

    def head(self, url: str) -> FetchResult:
        response = self._request("HEAD", url, allow_redirects=True)
        return FetchResult(
            url=url,
            final_url=response.url,
            status_code=response.status_code,
            content_type=response.headers.get("Content-Type", ""),
            content_length=response.headers.get("Content-Length", ""),
            last_modified=response.headers.get("Last-Modified", ""),
            content=b"",
        )

    def download(
        self,
        url: str,
        target: Path,
        *,
        progress_callback: Callable[[int], None] | None = None,
        event_callback: Callable[[str, dict[str, object]], None] | None = None,
    ) -> DownloadResult:
        target.parent.mkdir(parents=True, exist_ok=True)
        response = self._request("GET", url, stream=True)
        response_headers = dict(response.headers)
        response_headers_json = json.dumps(response_headers, ensure_ascii=False, sort_keys=True)
        if event_callback is not None:
            event_callback(
                "response",
                {
                    "status_code": response.status_code,
                    "final_url": response.url,
                    "content_type": response.headers.get("Content-Type", ""),
                    "content_length": response.headers.get("Content-Length", ""),
                    "headers": response_headers_json,
                },
            )
        digest = hashlib.sha256()
        bytes_written = 0
        first_bytes = b""
        part_target = target.with_name(f"{target.name}.part")
        try:
            target.unlink(missing_ok=True)
            part_target.unlink(missing_ok=True)
            with part_target.open("wb") as fh:
                for chunk in response.iter_content(chunk_size=1024 * 128):
                    if not chunk:
                        continue
                    if len(first_bytes) < 5:
                        first_bytes = (first_bytes + chunk)[:5]
                    fh.write(chunk)
                    digest.update(chunk)
                    bytes_written += len(chunk)
                    if progress_callback is not None:
                        progress_callback(bytes_written)
                fh.flush()
                os.fsync(fh.fileno())
            part_size = part_target.stat().st_size
            with part_target.open("rb") as verify_fh:
                part_readback = verify_fh.read(5)
            if event_callback is not None:
                event_callback(
                    "part_verify",
                    {
                        "part_path": str(part_target),
                        "bytes_written": bytes_written,
                        "part_size": part_size,
                        "first_bytes_hex": first_bytes.hex(),
                        "readback_hex": part_readback.hex(),
                    },
                )
            part_target.replace(target)
            post_write_size = target.stat().st_size
            with target.open("rb") as verify_fh:
                final_readback = verify_fh.read(5)
            post_write_readback_ok = final_readback == first_bytes[:5]
            if event_callback is not None:
                event_callback(
                    "final_verify",
                    {
                        "saved_path": str(target),
                        "post_write_size": post_write_size,
                        "readback_hex": final_readback.hex(),
                        "readback_ok": post_write_readback_ok,
                    },
                )
        except OSError as exc:
            target.unlink(missing_ok=True)
            part_target.unlink(missing_ok=True)
            errno_value = getattr(exc, "errno", None)
            winerror_value = getattr(exc, "winerror", None)
            raise FetchError(
                f"filesystem write failed: target={target} part={part_target} "
                f"errno={errno_value} winerror={winerror_value} message={exc}"
            ) from exc
        if bytes_written == 0:
            target.unlink(missing_ok=True)
            raise FetchError(f"GET {url} downloaded 0 bytes")
        content_type = response.headers.get("Content-Type", "")
        if first_bytes and not first_bytes.startswith(b"%PDF-") and "pdf" not in content_type.lower():
            target.unlink(missing_ok=True)
            raise FetchError(f"GET {url} did not return a PDF: Content-Type={content_type!r}, first_bytes={first_bytes!r}")
        if post_write_size != bytes_written:
            target.unlink(missing_ok=True)
            raise FetchError(f"GET {url} write size mismatch: bytes_written={bytes_written}, post_write_size={post_write_size}")
        if not post_write_readback_ok:
            target.unlink(missing_ok=True)
            raise FetchError(f"GET {url} readback mismatch after save: first_bytes={first_bytes!r}")
        return DownloadResult(
            url=url,
            final_url=response.url,
            status_code=response.status_code,
            content_type=content_type,
            content_length=response.headers.get("Content-Length", ""),
            last_modified=response.headers.get("Last-Modified", ""),
            sha256=digest.hexdigest(),
            saved_path=target,
            bytes_written=bytes_written,
            response_headers_json=response_headers_json,
            first_bytes_hex=first_bytes.hex(),
            post_write_size=post_write_size,
            post_write_readback_ok=post_write_readback_ok,
        )
