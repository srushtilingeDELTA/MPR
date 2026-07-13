"""Access Excel data stored in Delta's SharePoint from Python.

Two ways to reach the file are supported, so this works nicely from a desktop:

* ``--source local`` (recommended on a work desktop) - read the file straight
  from a local copy. If the ``GSE MPR Documents`` library is synced through the
  OneDrive/SharePoint client, it shows up as an ordinary folder and needs no
  authentication at all. Point ``--local-path`` at the synced folder or the
  file, or let the tool auto-detect a synced copy.
* ``--source sharepoint`` - authenticate to SharePoint Online, navigate to the
  file, download it, and load it into pandas.

Default target (overridable via CLI flags / function arguments):

    Site   : https://deltaairlines.sharepoint.com/sites/DL002488
    Library: GSE MPR Documents
    Folder : 6 - TESTING
    File   : MPR Actuals and Goals_v2   (.xlsx assumed if no extension given)

Authentication (only needed for ``--source sharepoint``)
--------------------------------------------------------
Delta's tenant almost certainly enforces MFA / conditional access, so plain
username+password auth usually fails. Pick whichever method your account allows:

* ``interactive`` (default) - opens a browser window; you log in normally
  (handles MFA). On a desktop this is the easiest option.
* ``device`` - prints a code + URL to enter on another device; good for
  headless/remote shells.
* ``app`` - app-only auth with an Azure AD app registration
  (``client_id`` + ``client_secret``). No user prompt; needs admin-granted
  ``Sites.Read.All`` (or similar) application permission.
* ``userpass`` - username + password (only works without MFA).

For ``interactive``/``device`` the SDK needs a public-client ``client_id``. This
module falls back to a well-known Microsoft first-party client id so it works
out of the box on most desktops; if your tenant's conditional access blocks it,
register an app and set ``SP_CLIENT_ID`` (a public client with the
``https://<tenant>.sharepoint.com/.default`` delegated permission is enough).

Credentials are read from environment variables so nothing is hard-coded:

    SP_TENANT                          (all methods, e.g. "deltaairlines.onmicrosoft.com")
    SP_CLIENT_ID                       (interactive/device/app; optional for interactive/device)
    SP_CLIENT_SECRET                   (app)
    SP_USERNAME, SP_PASSWORD           (userpass)
    SP_LOCAL_PATH                      (local; a synced folder or the file itself)

Examples
--------
    # Desktop with the library synced via OneDrive - no login needed:
    python sharepoint_excel.py --source local

    # Point at an explicit local copy and read one sheet:
    python sharepoint_excel.py --source local --local-path "~/Delta Air Lines/GSE MPR Documents - 6 - TESTING" --sheet Actuals

    # Browser login against SharePoint Online, then print a sheet summary:
    python sharepoint_excel.py --source sharepoint

    # Headless shell (device code) + save a local copy:
    python sharepoint_excel.py --source sharepoint --auth device --save-to mpr.xlsx
"""

from __future__ import annotations

import argparse
import io
import os
import sys
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Dict, List, Optional, Union

import pandas as pd

# --- Defaults for the file the user asked about -----------------------------

DEFAULT_SITE_URL = "https://deltaairlines.sharepoint.com/sites/DL002488"
DEFAULT_LIBRARY = "GSE MPR Documents"
DEFAULT_FOLDER = "6 - TESTING"
DEFAULT_FILE = "MPR Actuals and Goals_v2"
DEFAULT_TENANT = "deltaairlines.onmicrosoft.com"

# Microsoft first-party "Office" public client id. Broadly pre-consented for
# delegated Microsoft 365 access, so interactive/device login works without
# registering a custom app. Override with SP_CLIENT_ID if conditional access
# blocks it in your tenant.
DEFAULT_PUBLIC_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"

# File extensions we treat as Excel workbooks when scanning a folder.
EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".xls")


@dataclass
class AuthConfig:
    """How to authenticate to SharePoint. Values default to env vars."""

    method: str = "interactive"  # interactive | device | app | userpass
    tenant: str = os.environ.get("SP_TENANT", DEFAULT_TENANT)
    username: Optional[str] = os.environ.get("SP_USERNAME")
    password: Optional[str] = os.environ.get("SP_PASSWORD")
    client_id: Optional[str] = os.environ.get("SP_CLIENT_ID")
    client_secret: Optional[str] = os.environ.get("SP_CLIENT_SECRET")


def build_context(site_url: str, auth: AuthConfig):
    """Return an authenticated ``ClientContext`` for ``site_url``.

    Imports of the SharePoint SDK are done lazily so the pure-pandas helpers in
    this module (and the test-suite) do not require the SDK to be installed.
    """
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext

    method = auth.method.lower()

    if method in ("interactive", "device"):
        client_id = auth.client_id or DEFAULT_PUBLIC_CLIENT_ID
        if method == "interactive":
            return ClientContext(site_url).with_interactive(auth.tenant, client_id)
        return ClientContext(site_url).with_device_flow(auth.tenant, client_id)

    if method == "app":
        if not (auth.client_id and auth.client_secret):
            raise ValueError(
                "app auth requires SP_CLIENT_ID and SP_CLIENT_SECRET "
                "(or pass client_id/client_secret)."
            )
        credentials = ClientCredential(auth.client_id, auth.client_secret)
        return ClientContext(site_url).with_credentials(credentials)

    if method == "userpass":
        if not (auth.username and auth.password):
            raise ValueError(
                "userpass auth requires SP_USERNAME and SP_PASSWORD "
                "(or pass username/password)."
            )
        credentials = UserCredential(auth.username, auth.password)
        return ClientContext(site_url).with_credentials(credentials)

    raise ValueError(
        f"Unknown auth method {auth.method!r}; "
        "expected one of: interactive, device, app, userpass."
    )


def get_library_root_url(ctx, library_title: str) -> str:
    """Return the server-relative URL of a document library's root folder.

    Using the library *title* is more robust than guessing its URL slug, since
    display names (e.g. "GSE MPR Documents") often differ from the URL path.
    """
    library = ctx.web.lists.get_by_title(library_title)
    root = library.root_folder
    ctx.load(root)
    ctx.execute_query()
    return root.serverRelativeUrl


def _stem(name: str) -> str:
    return PurePosixPath(name).stem.casefold()


def match_name(available: List[str], file_name: str) -> Optional[str]:
    """Pick the entry in ``available`` matching ``file_name``.

    Prefers an exact (case-insensitive) match, then falls back to matching the
    stem so "MPR Actuals and Goals_v2" resolves to
    "MPR Actuals and Goals_v2.xlsx". Returns ``None`` if nothing matches.
    """
    target = file_name.casefold()
    for name in available:
        if name.casefold() == target:
            return name

    target_stem = _stem(file_name)
    for name in available:
        if _stem(name) == target_stem:
            return name
    return None


def find_file_url(ctx, folder_url: str, file_name: str) -> str:
    """Find a file within ``folder_url`` and return its server-relative URL.

    Raises ``FileNotFoundError`` with the available names if nothing matches.
    """
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    available = [f.properties.get("Name", "") for f in files]
    match = match_name(available, file_name)
    if match is not None:
        return str(PurePosixPath(folder_url) / match)

    raise FileNotFoundError(
        f"No file matching {file_name!r} in {folder_url!r}. "
        f"Found: {available!r}"
    )


def download_file(ctx, file_server_relative_url: str) -> bytes:
    """Download a file from SharePoint and return its raw bytes."""
    buffer = io.BytesIO()
    (
        ctx.web.get_file_by_server_relative_url(file_server_relative_url)
        .download(buffer)
        .execute_query()
    )
    return buffer.getvalue()


def read_excel_bytes(
    data: bytes,
    sheet_name: Union[str, int, None] = None,
) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """Parse Excel bytes into a DataFrame (one sheet) or a dict of DataFrames.

    ``sheet_name=None`` returns every sheet keyed by name (pandas default when
    reading all sheets).
    """
    effective = 0 if sheet_name is None else sheet_name
    if sheet_name is None:
        # Read all sheets.
        return pd.read_excel(io.BytesIO(data), sheet_name=None, engine="openpyxl")
    return pd.read_excel(io.BytesIO(data), sheet_name=effective, engine="openpyxl")


def load_excel_from_sharepoint(
    site_url: str = DEFAULT_SITE_URL,
    library: str = DEFAULT_LIBRARY,
    folder: str = DEFAULT_FOLDER,
    file_name: str = DEFAULT_FILE,
    sheet_name: Union[str, int, None] = None,
    auth: Optional[AuthConfig] = None,
    save_to: Optional[str] = None,
    server_folder_path: Optional[str] = None,
) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """End-to-end: authenticate, locate the file, download it, load into pandas.

    Returns the parsed workbook. If ``save_to`` is given, the raw file is also
    written to that local path.

    If ``server_folder_path`` is set (e.g. from config), it is used directly
    instead of resolving the library root by title.
    """
    auth = auth or AuthConfig()
    ctx = build_context(site_url, auth)

    if server_folder_path:
        folder_url = server_folder_path
    else:
        library_root = get_library_root_url(ctx, library)
        folder_url = str(PurePosixPath(library_root) / folder) if folder else library_root
    file_url = find_file_url(ctx, folder_url, file_name)

    data = download_file(ctx, file_url)

    if save_to:
        with open(save_to, "wb") as fh:
            fh.write(data)

    return read_excel_bytes(data, sheet_name=sheet_name)


# --- local (OneDrive-synced / downloaded) file access -----------------------


def find_local_file(local_path: Union[str, Path], file_name: str = DEFAULT_FILE) -> Path:
    """Resolve ``file_name`` under a local path.

    ``local_path`` may be the file itself or a directory. Directories are
    searched recursively (so a synced library root will find the file inside
    ``6 - TESTING`` automatically). Matching prefers exact name, then stem, and
    finally any Excel workbook whose stem matches. Raises ``FileNotFoundError``
    if nothing suitable is found.
    """
    path = Path(local_path).expanduser()

    if path.is_file():
        return path

    if not path.exists():
        raise FileNotFoundError(f"Local path does not exist: {path}")

    candidates = [p for p in path.rglob("*") if p.is_file()]
    names = [p.name for p in candidates]
    by_name = {p.name: p for p in candidates}

    match = match_name(names, file_name)
    if match is not None:
        return by_name[match]

    # Fall back to any Excel workbook whose stem matches the requested name.
    target_stem = _stem(file_name)
    for p in candidates:
        if p.suffix.lower() in EXCEL_SUFFIXES and _stem(p.name) == target_stem:
            return p

    excel_files = [n for n in names if Path(n).suffix.lower() in EXCEL_SUFFIXES]
    raise FileNotFoundError(
        f"No file matching {file_name!r} under {path}. "
        f"Excel files found: {excel_files!r}"
    )


def default_local_search_roots(library: str = DEFAULT_LIBRARY) -> List[Path]:
    """Best-effort list of local folders where a synced library might live."""
    roots: List[Path] = []

    env_path = os.environ.get("SP_LOCAL_PATH")
    if env_path:
        roots.append(Path(env_path).expanduser())

    home = Path.home()
    # OneDrive/SharePoint sync roots and any home-level folder that references
    # the library by name (e.g. "GSE MPR Documents - 6 - TESTING").
    for entry in sorted(home.glob("*")):
        if not entry.is_dir():
            continue
        name = entry.name.casefold()
        if "onedrive" in name or library.casefold() in name or "sharepoint" in name:
            roots.append(entry)

    return roots


def autodetect_local_file(
    file_name: str = DEFAULT_FILE,
    library: str = DEFAULT_LIBRARY,
) -> Path:
    """Search common sync locations for the file. Raises if none is found."""
    searched = default_local_search_roots(library)
    for root in searched:
        try:
            return find_local_file(root, file_name)
        except FileNotFoundError:
            continue

    raise FileNotFoundError(
        f"Could not auto-detect {file_name!r} locally. "
        f"Searched: {[str(p) for p in searched]!r}. "
        "Pass --local-path pointing at the synced folder or the file, "
        "or set SP_LOCAL_PATH."
    )


def load_excel_local(
    local_path: Optional[Union[str, Path]] = None,
    file_name: str = DEFAULT_FILE,
    library: str = DEFAULT_LIBRARY,
    sheet_name: Union[str, int, None] = None,
    save_to: Optional[str] = None,
) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """Load the Excel file from a local (OneDrive-synced or downloaded) copy.

    If ``local_path`` is omitted, common sync locations are auto-detected.
    """
    if local_path:
        resolved = find_local_file(local_path, file_name)
    else:
        resolved = autodetect_local_file(file_name, library)

    data = resolved.read_bytes()

    if save_to:
        with open(save_to, "wb") as fh:
            fh.write(data)

    return read_excel_bytes(data, sheet_name=sheet_name)


def _summarize(result: Union[pd.DataFrame, Dict[str, pd.DataFrame]]) -> None:
    if isinstance(result, dict):
        print(f"Loaded {len(result)} sheet(s):")
        for name, df in result.items():
            print(f"  - {name}: {df.shape[0]} rows x {df.shape[1]} cols")
            print(df.head().to_string(index=False))
            print()
    else:
        print(f"Loaded sheet: {result.shape[0]} rows x {result.shape[1]} cols")
        print(result.head().to_string(index=False))


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read an Excel file from Delta SharePoint (online or a local synced copy)."
    )
    parser.add_argument(
        "--source",
        default="sharepoint",
        choices=["sharepoint", "local"],
        help="Where to read the file from. 'local' reads a OneDrive-synced or "
        "downloaded copy (no login). Default: sharepoint.",
    )
    parser.add_argument(
        "--local-path",
        default=None,
        help="For --source local: the synced folder or the file itself. "
        "Defaults to SP_LOCAL_PATH, else common sync locations are auto-detected.",
    )
    parser.add_argument("--site-url", default=DEFAULT_SITE_URL)
    parser.add_argument("--library", default=DEFAULT_LIBRARY)
    parser.add_argument("--folder", default=DEFAULT_FOLDER)
    parser.add_argument("--file", dest="file_name", default=DEFAULT_FILE)
    parser.add_argument(
        "--sheet",
        default=None,
        help="Sheet name or index to load. Omit to load all sheets.",
    )
    parser.add_argument(
        "--auth",
        default="interactive",
        choices=["interactive", "device", "app", "userpass"],
        help="Authentication method (default: interactive browser login).",
    )
    parser.add_argument("--tenant", default=None, help="Azure AD tenant, e.g. contoso.onmicrosoft.com")
    parser.add_argument(
        "--client-id",
        default=None,
        help="Azure AD app (public client) id. Required for interactive/device auth. "
        "Defaults to the SP_CLIENT_ID env var.",
    )
    parser.add_argument("--save-to", default=None, help="Also write the raw file to this local path.")
    return parser


def main(argv: Optional[list] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    auth = AuthConfig(method=args.auth)
    if args.tenant:
        auth.tenant = args.tenant
    if args.client_id:
        auth.client_id = args.client_id

    sheet: Union[str, int, None] = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    try:
        if args.source == "local":
            result = load_excel_local(
                local_path=args.local_path,
                file_name=args.file_name,
                library=args.library,
                sheet_name=sheet,
                save_to=args.save_to,
            )
        else:
            result = load_excel_from_sharepoint(
                site_url=args.site_url,
                library=args.library,
                folder=args.folder,
                file_name=args.file_name,
                sheet_name=sheet,
                auth=auth,
                save_to=args.save_to,
            )
    except Exception as exc:  # noqa: BLE001 - surface a clean message to the CLI user
        print(f"Failed to load file ({args.source}): {exc}", file=sys.stderr)
        return 1

    _summarize(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
