"""Read SharePoint files in memory using an Edge/Chrome session + REST API.

Delta blocks Office365 SDK OAuth, but after you sign in through Edge the
browser session cookies work with SharePoint's REST API. One Edge login can
fetch multiple files from the same folder without saving to disk (unless
cache_to_disk is enabled).

Usage:
    python scripts/sync_sharepoint_files.py
    python scripts/list_sharepoint_folder.py
    # or set excel.source: sharepoint and run python main.py
"""

from __future__ import annotations

import io
import json
import logging
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import requests

logger = logging.getLogger(__name__)

DEFAULT_SERVER_FOLDER = (
    "/sites/DL002488/MPR  Research/GSE MPR Documents/6 - TESTING"
)

# Populated by sync_sharepoint_files(); keyed by project-relative dest path.
_SHAREPOINT_CACHE: dict[str, bytes] = {}


@dataclass
class SharePointFileSpec:
    name: str
    dest: str
    optional: bool = False


@dataclass
class SharePointSyncResult:
    files: dict[str, bytes] = field(default_factory=dict)
    errors: dict[str, str] = field(default_factory=dict)

    def get(self, dest: str) -> bytes | None:
        return self.files.get(dest)


def _browser_path(sp_cfg: dict, browser: str) -> str | None:
    if browser == "edge":
        for candidate in (
            sp_cfg.get("edge_path"),
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ):
            if candidate and Path(candidate).exists():
                return candidate
        return None
    candidate = sp_cfg.get("chrome_path")
    return candidate if candidate and Path(candidate).exists() else None


def _build_requests_session(driver) -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "application/json;odata=verbose",
        }
    )
    for cookie in driver.get_cookies():
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain"),
            path=cookie.get("path", "/"),
        )
    return session


def begin_sharepoint_browser(config: dict) -> requests.Session:
    """Open Edge/Chrome once for this run; reuse for sync + upload. Call end_sharepoint_browser() when done."""
    existing = config.get("_sharepoint_session")
    if existing is not None and config.get("_sharepoint_driver") is not None:
        logger.debug("Reusing open SharePoint browser session")
        return existing

    sp_cfg = config.get("sharepoint", {})
    folder_page_url = sp_cfg.get("folder_page_url")
    if not folder_page_url:
        raise ValueError(
            "sharepoint.folder_page_url is required. "
            "Paste the full Edge URL when the 6 - TESTING folder is open."
        )

    from sharepoint_selenium import _create_driver

    browser = sp_cfg.get("browser", "edge")
    browser_path = _browser_path(sp_cfg, browser)
    login_wait = int(sp_cfg.get("login_wait_seconds", 180))
    browser_label = "Edge" if browser.lower() == "edge" else "Chrome"

    driver = _create_driver(browser, Path.cwd(), browser_path, headless=False)
    logger.info("Opening SharePoint (%s) — one browser window for this run", browser_label)
    driver.get(folder_page_url)
    print(
        f"\n>>> {browser_label}: sign in to Delta / SharePoint if prompted.\n"
        f">>> Waiting {login_wait}s so the folder view loads.\n"
        f">>> This same window will be reused for upload (no extra tabs).\n"
    )
    time.sleep(login_wait)

    session = _build_requests_session(driver)
    config["_sharepoint_session"] = session
    config["_sharepoint_driver"] = driver
    return session


def end_sharepoint_browser(config: dict) -> None:
    """Close the Edge/Chrome window opened by begin_sharepoint_browser()."""
    driver = config.pop("_sharepoint_driver", None)
    if driver is not None:
        try:
            driver.quit()
        except Exception as exc:
            logger.debug("Browser close: %s", exc)


def browser_session_from_sharepoint(
    folder_page_url: str,
    *,
    browser: str = "edge",
    browser_path: str | None = None,
    login_wait_seconds: int = 180,
) -> requests.Session:
    """One-shot auth: open browser, return session, close browser (standalone scripts only)."""
    from sharepoint_selenium import _create_driver

    driver = _create_driver(browser, Path.cwd(), browser_path, headless=False)
    browser_label = "Edge" if browser.lower() == "edge" else "Chrome"

    try:
        logger.info("Opening SharePoint for live read (%s)", browser_label)
        driver.get(folder_page_url)
        print(
            f"\n>>> {browser_label}: sign in to Delta / SharePoint if prompted.\n"
            f">>> Waiting {login_wait_seconds}s so the folder view loads.\n"
        )
        time.sleep(login_wait_seconds)
        return _build_requests_session(driver)
    finally:
        driver.quit()


def server_path_for_file(sp_cfg: dict, file_name: str) -> str:
    folder = sp_cfg.get("server_folder_path", DEFAULT_SERVER_FOLDER).rstrip("/")
    return f"{folder}/{file_name}"


def _file_server_path(sp_cfg: dict, config: dict) -> str:
    file_name = sp_cfg.get("file_name") or Path(config["excel"]["file_path"]).name
    return server_path_for_file(sp_cfg, file_name)


def default_file_specs(config: dict, base_dir: Path) -> list[SharePointFileSpec]:
    """Build file list from sharepoint.files or sensible defaults."""
    sp_cfg = config.get("sharepoint", {})
    raw_files = sp_cfg.get("files")
    if raw_files:
        specs: list[SharePointFileSpec] = []
        for entry in raw_files:
            if isinstance(entry, str):
                specs.append(SharePointFileSpec(name=entry, dest=_guess_dest(entry, config)))
            else:
                name = entry["name"]
                dest = entry.get("dest") or _guess_dest(name, config)
                specs.append(
                    SharePointFileSpec(
                        name=name,
                        dest=dest,
                        optional=bool(entry.get("optional", False)),
                    )
                )
        return specs

    specs = [
        SharePointFileSpec(
            name=sp_cfg.get("file_name") or Path(config["excel"]["file_path"]).name,
            dest=config["excel"]["file_path"],
        ),
        SharePointFileSpec(
            name=Path(config["powerpoint"]["template_path"]).name,
            dest=config["powerpoint"]["template_path"],
            optional=True,
        ),
    ]
    return specs


def _guess_dest(file_name: str, config: dict) -> str:
    lower = file_name.lower()
    if lower.endswith((".pptx", ".ppt", ".potx")):
        return f"templates/{file_name}"
    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        return f"data/{file_name}"
    if lower.endswith(".pdf"):
        return f"data/{file_name}"
    return f"data/{file_name}"


def download_file_bytes(
    session: requests.Session,
    site_url: str,
    server_relative_path: str,
    *,
    min_bytes: int = 100,
) -> bytes:
    """Download a SharePoint file into memory using an authenticated session."""
    site_url = site_url.rstrip("/")
    path_literal = server_relative_path.replace("'", "''")

    urls = [
        f"{site_url}/_api/web/GetFileByServerRelativeUrl('{path_literal}')/$value",
        f"{site_url}/_api/web/GetFileByServerRelativePath(decodedurl='{server_relative_path}')/$value",
    ]

    last_error: Exception | None = None
    for api_url in urls:
        try:
            response = session.get(api_url, timeout=120)
            if response.status_code in (401, 403):
                raise PermissionError(
                    f"SharePoint denied access (HTTP {response.status_code}). "
                    "Sign in again in Edge or increase login_wait_seconds."
                )
            response.raise_for_status()
            content = response.content
            if _looks_like_login_page(content):
                raise PermissionError(
                    "SharePoint returned a login page instead of the file. "
                    "Increase login_wait_seconds and complete MFA before the wait ends."
                )
            if len(content) < min_bytes:
                raise ValueError(
                    f"Download too small ({len(content)} bytes) for {server_relative_path!r}."
                )
            logger.info("Fetched %s bytes from %s", len(content), server_relative_path)
            return content
        except Exception as exc:
            last_error = exc
            logger.debug("SharePoint API attempt failed: %s", exc)

    raise RuntimeError(f"Could not download {server_relative_path!r}: {last_error}")


def list_folder_files(
    session: requests.Session,
    site_url: str,
    server_folder_path: str,
) -> list[dict[str, Any]]:
    """Return file metadata from a SharePoint folder via REST."""
    site_url = site_url.rstrip("/")
    folder_literal = server_folder_path.replace("'", "''")
    api_url = (
        f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{folder_literal}')/Files"
        "?$select=Name,Length,ServerRelativeUrl,TimeLastModified"
    )
    response = session.get(api_url, timeout=60)
    response.raise_for_status()
    payload = response.json()
    rows = payload.get("d", {}).get("results") or payload.get("value") or []
    return [
        {
            "name": row.get("Name") or row.get("name"),
            "size": row.get("Length") or row.get("length"),
            "server_relative_url": row.get("ServerRelativeUrl") or row.get("serverRelativeUrl"),
        }
        for row in rows
    ]


def _looks_like_login_page(content: bytes) -> bool:
    head = content[:500].lower()
    return b"<!doctype html" in head or (b"<html" in head and b"sign in" in head)


def _authenticate(config: dict) -> tuple[requests.Session, dict]:
    sp_cfg = config.get("sharepoint", {})
    session = config.get("_sharepoint_session")
    if session is not None:
        return session, sp_cfg

    folder_page_url = sp_cfg.get("folder_page_url")
    if not folder_page_url:
        raise ValueError(
            "sharepoint.folder_page_url is required. "
            "Paste the full Edge URL when the 6 - TESTING folder is open."
        )
    browser = sp_cfg.get("browser", "edge")
    login_wait = int(sp_cfg.get("login_wait_seconds", 180))
    session = browser_session_from_sharepoint(
        folder_page_url,
        browser=browser,
        browser_path=_browser_path(sp_cfg, browser),
        login_wait_seconds=login_wait,
    )
    config["_sharepoint_session"] = session
    return session, sp_cfg


def sync_sharepoint_files(config: dict, base_dir: Path | None = None) -> SharePointSyncResult:
    """One Edge login; fetch all configured files (and optionally every file in the folder)."""
    global _SHAREPOINT_CACHE

    base_dir = base_dir or Path.cwd()
    session, sp_cfg = _authenticate(config)
    site_url = sp_cfg.get("site_url", "https://deltaairlines.sharepoint.com/sites/DL002488")
    folder_path = sp_cfg.get("server_folder_path", DEFAULT_SERVER_FOLDER)

    result = SharePointSyncResult()
    specs = default_file_specs(config, base_dir)

    if sp_cfg.get("fetch_all_in_folder"):
        listed = list_folder_files(session, site_url, folder_path)
        known_names = {spec.name for spec in specs}
        for row in listed:
            name = row.get("name")
            if not name or name in known_names:
                continue
            specs.append(
                SharePointFileSpec(
                    name=name,
                    dest=_guess_dest(name, config),
                    optional=True,
                )
            )

    for spec in specs:
        server_path = server_path_for_file(sp_cfg, spec.name)
        try:
            data = download_file_bytes(session, site_url, server_path)
            result.files[spec.dest] = data
            _SHAREPOINT_CACHE[spec.dest] = data
            if sp_cfg.get("cache_to_disk"):
                out = base_dir / spec.dest
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_bytes(data)
                logger.info("Cached %s to %s", spec.name, out)
            else:
                logger.info("Loaded %s into memory (%s bytes)", spec.name, len(data))
        except Exception as exc:
            msg = str(exc)
            result.errors[spec.name] = msg
            if spec.optional:
                logger.warning("Optional SharePoint file skipped (%s): %s", spec.name, msg)
            else:
                logger.error("Required SharePoint file failed (%s): %s", spec.name, msg)
                raise

    config["_sharepoint_session"] = session
    config["_sharepoint_site_url"] = site_url
    config["_sharepoint_folder_path"] = folder_path
    return result


def get_cached_file(dest: str) -> bytes | None:
    """Return bytes for a project-relative dest path if already synced this run."""
    return _SHAREPOINT_CACHE.get(dest)


def attach_cache_to_config(config: dict, result: SharePointSyncResult) -> None:
    config["_sharepoint_files"] = result.files


def fetch_workbook_bytes(config: dict, base_dir: Path | None = None) -> bytes:
    """Return the MPR Excel workbook bytes (uses cache when sync already ran)."""
    excel_dest = config["excel"]["file_path"]
    cached = get_cached_file(excel_dest)
    if cached:
        return cached

    if config.get("_sharepoint_files", {}).get(excel_dest):
        return config["_sharepoint_files"][excel_dest]

    result = sync_sharepoint_files(config, base_dir=base_dir)
    attach_cache_to_config(config, result)
    if excel_dest not in result.files:
        raise FileNotFoundError(f"Excel file not loaded from SharePoint: {excel_dest}")
    return result.files[excel_dest]


def fetch_workbook_buffer(config: dict, base_dir: Path | None = None) -> io.BytesIO:
    return io.BytesIO(fetch_workbook_bytes(config, base_dir=base_dir))


def get_sharepoint_session(config: dict) -> tuple[requests.Session, dict]:
    """Return the SharePoint session for this run (never opens a new browser)."""
    sp_cfg = config.get("sharepoint", {})
    session = config.get("_sharepoint_session")
    if session is not None:
        return session, sp_cfg
    session, sp_cfg = _authenticate(config)
    return session, sp_cfg


def _request_digest(session: requests.Session, site_url: str) -> str:
    """Fetch SharePoint form digest required for POST/PUT uploads."""
    response = session.post(
        f"{site_url.rstrip('/')}/_api/contextinfo",
        headers={"Accept": "application/json;odata=verbose"},
        timeout=60,
    )
    response.raise_for_status()
    payload = response.json()
    info = payload.get("d", {}).get("GetContextWebInformation") or payload.get("GetContextWebInformation")
    if not info or not info.get("FormDigestValue"):
        raise RuntimeError("SharePoint did not return a request digest token.")
    return str(info["FormDigestValue"])


def _upload_headers(session: requests.Session, site_url: str, *, method: str = "POST") -> dict[str, str]:
    digest = _request_digest(session, site_url)
    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/octet-stream",
        "X-RequestDigest": digest,
    }
    if method == "PUT":
        headers["IF-MATCH"] = "*"
        headers["X-HTTP-Method"] = "PUT"
    return headers


def _delete_headers(session: requests.Session, site_url: str) -> dict[str, str]:
    digest = _request_digest(session, site_url)
    return {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
    }


def _delete_sharepoint_file(session: requests.Session, site_url: str, server_path: str) -> bool:
    """Delete an existing SharePoint file so a fresh upload can use the same name."""
    path_literal = server_path.replace("'", "''")
    headers = _delete_headers(session, site_url)
    urls = [
        f"{site_url.rstrip('/')}/_api/web/GetFileByServerRelativeUrl('{path_literal}')",
        f"{site_url.rstrip('/')}/_api/web/GetFileByServerRelativePath(decodedurl='{server_path}')",
    ]
    for api_url in urls:
        try:
            response = session.post(api_url, headers=headers, timeout=120)
            if response.status_code in (200, 204, 404):
                logger.info("Removed previous SharePoint file: %s", server_path)
                return True
        except Exception as exc:
            logger.debug("SharePoint delete attempt failed (%s): %s", api_url, exc)
    return False


def _upload_is_locked(exc: Exception) -> bool:
    message = str(exc).lower()
    return "423" in message or "locked" in message


def _parse_upload_response(response: requests.Response, folder: str, file_name: str) -> str:
    if response.status_code in (401, 403):
        body = response.text[:500]
        raise PermissionError(
            f"SharePoint upload denied (HTTP {response.status_code}). "
            f"Response: {body}"
        )
    if response.status_code == 423:
        raise PermissionError(
            f"SharePoint file is locked (HTTP 423): {folder}/{file_name}. "
            "Close the file in Edge/SharePoint if it is open, then re-run."
        )
    response.raise_for_status()
    try:
        payload = response.json()
    except Exception:
        return f"{folder}/{file_name}"
    server_url = payload.get("d", {}).get("ServerRelativeUrl") or payload.get("ServerRelativeUrl")
    return server_url or f"{folder}/{file_name}"


def _upload_via_add(
    session: requests.Session,
    site_url: str,
    folder: str,
    file_name: str,
    data: bytes,
) -> str:
    folder_literal = folder.replace("'", "''")
    file_literal = file_name.replace("'", "''")
    headers = _upload_headers(session, site_url, method="POST")

    urls = [
        (
            f"{site_url.rstrip('/')}/_api/web/GetFolderByServerRelativeUrl('{folder_literal}')"
            f"/Files/add(url='{file_literal}',overwrite=true)"
        ),
        (
            f"{site_url.rstrip('/')}/_api/web/GetFolderByServerRelativePath("
            f"decodedurl='{folder}')/Files/add(url='{file_literal}',overwrite=true)"
        ),
    ]
    last_error: Exception | None = None
    for api_url in urls:
        try:
            response = session.post(api_url, data=data, headers=headers, timeout=300)
            return _parse_upload_response(response, folder, file_name)
        except Exception as exc:
            last_error = exc
            logger.debug("Upload add attempt failed (%s): %s", api_url, exc)
    raise RuntimeError(f"SharePoint add upload failed: {last_error}")


def _upload_via_put(
    session: requests.Session,
    site_url: str,
    folder: str,
    file_name: str,
    data: bytes,
) -> str:
    server_path = f"{folder.rstrip('/')}/{file_name}"
    path_literal = server_path.replace("'", "''")
    headers = _upload_headers(session, site_url, method="PUT")
    urls = [
        f"{site_url.rstrip('/')}/_api/web/GetFileByServerRelativeUrl('{path_literal}')/$value",
        (
            f"{site_url.rstrip('/')}/_api/web/GetFileByServerRelativePath("
            f"decodedurl='{server_path}')/$value"
        ),
    ]
    last_error: Exception | None = None
    for api_url in urls:
        try:
            response = session.put(api_url, data=data, headers=headers, timeout=300)
            return _parse_upload_response(response, folder, file_name)
        except Exception as exc:
            last_error = exc
            logger.debug("Upload put attempt failed (%s): %s", api_url, exc)
    raise RuntimeError(f"SharePoint overwrite upload failed: {last_error}")


def upload_file_to_sharepoint(
    config: dict,
    file_name: str,
    data: bytes,
    *,
    folder_path: str | None = None,
    local_path: Path | None = None,
) -> str:
    """Upload bytes to SharePoint using the existing session; reuse open browser for UI upload."""
    sp_cfg = config.get("sharepoint", {})
    site_url = sp_cfg.get("site_url", "https://deltaairlines.sharepoint.com/sites/DL002488").rstrip("/")
    folder = folder_path or sp_cfg.get("server_folder_path", DEFAULT_SERVER_FOLDER)
    upload_mode = str(sp_cfg.get("upload_method", "auto")).lower()

    errors: list[str] = []
    session, _ = get_sharepoint_session(config)

    if upload_mode in ("auto", "rest"):
        server_path = f"{folder.rstrip('/')}/{file_name}"
        try:
            try:
                server_path = _upload_via_add(session, site_url, folder, file_name, data)
                logger.info("Uploaded %s to SharePoint (%s bytes)", file_name, len(data))
                return server_path
            except Exception as add_exc:
                logger.info("Add upload failed, trying overwrite: %s", add_exc)
                try:
                    server_path = _upload_via_put(session, site_url, folder, file_name, data)
                    logger.info("Uploaded %s to SharePoint (%s bytes)", file_name, len(data))
                    return server_path
                except Exception as put_exc:
                    if _upload_is_locked(add_exc) or _upload_is_locked(put_exc):
                        logger.info("SharePoint file locked; deleting previous copy then re-uploading")
                        _delete_sharepoint_file(session, site_url, server_path)
                        server_path = _upload_via_add(session, site_url, folder, file_name, data)
                        logger.info("Uploaded %s to SharePoint after delete (%s bytes)", file_name, len(data))
                        return server_path
                    raise put_exc from add_exc
        except Exception as exc:
            errors.append(f"REST: {exc}")
            logger.warning("REST upload failed: %s", exc)

    driver = config.get("_sharepoint_driver")
    if upload_mode in ("auto", "selenium") and local_path is not None:
        from sharepoint_selenium import upload_file_via_selenium

        upload_file_via_selenium(config, local_path, driver=driver)
        logger.info("Uploaded %s to SharePoint via Edge UI (%s bytes)", file_name, len(data))
        return f"{folder}/{file_name}"

    raise PermissionError(
        "Could not upload report to SharePoint. "
        f"Attempts: {'; '.join(errors)}. "
        "Try sharepoint.upload_method: selenium in config.yaml, "
        "or drag the file from output\\ into the SharePoint folder in Edge."
    )
