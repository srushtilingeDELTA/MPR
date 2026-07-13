"""Download MPR Excel from SharePoint using Selenium (Delta SSO / MFA).

Use Edge or Chrome with manual login when the Office365 SDK is blocked by
conditional access. Opens the SharePoint folder page from config, waits for
you to sign in, then finds the Excel file and downloads it.

Usage:
    python scripts/download_from_sharepoint.py --method selenium
    python scripts/download_from_sharepoint.py --method selenium --browser edge
"""

from __future__ import annotations

import logging
import time
from pathlib import Path

logger = logging.getLogger(__name__)

DEFAULT_SITE = "https://deltaairlines.sharepoint.com/sites/DL002488"
DEFAULT_LIBRARY = "GSE MPR Documents"
DEFAULT_FOLDER = "6 - TESTING"
DEFAULT_FILE = "MPR Actuals and Goals_v2.xlsx"
DEFAULT_SERVER_FOLDER = (
    "/sites/DL002488/MPR  Research/GSE MPR Documents/6 - TESTING"
)


def _installed_edge_version(browser_path: str | None) -> str | None:
    """Return installed Edge version string, e.g. 149.0.4022.62."""
    import re
    import subprocess

    candidates: list[str] = []
    if browser_path:
        candidates.append(browser_path)
    candidates.extend(
        [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
    )
    for path in candidates:
        if not Path(path).exists():
            continue
        try:
            result = subprocess.run(
                [path, "--version"],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            text = (result.stdout or result.stderr or "").strip()
            match = re.search(r"(\d+\.\d+\.\d+\.\d+)", text)
            if match:
                return match.group(1)
        except Exception:
            continue
    return None


def _create_edge_driver(options, browser_path: str | None):
    """Create Edge WebDriver with a driver version that matches the browser."""
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service

    version = _installed_edge_version(browser_path)
    if version:
        logger.info("Detected Microsoft Edge %s", version)

    errors: list[str] = []

    try:
        return webdriver.Edge(options=options)
    except Exception as exc:
        errors.append(f"Selenium Manager: {exc}")

    if version:
        try:
            from webdriver_manager.microsoft import EdgeChromiumDriverManager

            service = Service(
                EdgeChromiumDriverManager(driver_version=version).install()
            )
            return webdriver.Edge(service=service, options=options)
        except Exception as exc:
            errors.append(f"webdriver-manager (version={version}): {exc}")

    try:
        from webdriver_manager.microsoft import EdgeChromiumDriverManager

        service = Service(EdgeChromiumDriverManager().install())
        return webdriver.Edge(service=service, options=options)
    except Exception as exc:
        errors.append(f"webdriver-manager (latest): {exc}")

    raise RuntimeError(
        "Could not start Microsoft Edge WebDriver. "
        "Update Edge and Selenium, or set sharepoint.browser to chrome. "
        f"Attempts: {'; '.join(errors)}"
    )


def _create_driver(
    browser: str,
    download_dir: Path,
    browser_path: str | None,
    headless: bool,
):
    """Return a Selenium WebDriver for Edge or Chrome with downloads enabled."""
    from selenium import webdriver

    prefs = {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }

    browser = browser.lower()
    if browser == "edge":
        from selenium.webdriver.edge.options import Options

        options = Options()
        if browser_path:
            options.binary_location = browser_path
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_experimental_option("prefs", prefs)
        if not headless:
            options.add_experimental_option("detach", True)
        return _create_edge_driver(options, browser_path)

    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    options = Options()
    if browser_path:
        options.binary_location = browser_path
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )


def download_via_selenium(
    *,
    site_url: str = DEFAULT_SITE,
    library: str = DEFAULT_LIBRARY,
    folder: str = DEFAULT_FOLDER,
    file_name: str = DEFAULT_FILE,
    save_to: Path,
    folder_page_url: str | None = None,
    server_folder_path: str | None = None,
    browser: str = "edge",
    browser_path: str | None = None,
    login_wait_seconds: int = 120,
    headless: bool = False,
) -> Path:
    """Open SharePoint in Edge/Chrome, wait for manual login, download the Excel file."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait

    save_to = Path(save_to)
    save_to.parent.mkdir(parents=True, exist_ok=True)

    driver = _create_driver(browser, save_to.parent, browser_path, headless)

    try:
        if folder_page_url:
            target_url = folder_page_url
        elif server_folder_path:
            encoded = server_folder_path.replace(" ", "%20")
            target_url = f"{site_url.rstrip('/')}/{encoded.lstrip('/')}"
        else:
            folder_path = folder.replace(" ", "%20")
            library_path = library.replace(" ", "%20")
            target_url = f"{site_url.rstrip('/')}/{library_path}/{folder_path}"

        logger.info("Opening SharePoint (%s): %s", browser, target_url)
        driver.get(target_url)

        browser_label = "Edge" if browser.lower() == "edge" else "Chrome"
        print(
            f"\n>>> {browser_label} opened SharePoint.\n"
            f">>> Log in with Delta SSO + MFA if prompted.\n"
            f">>> You have {login_wait_seconds} seconds — open the Excel file if needed.\n"
            f">>> Looking for: {file_name}\n"
        )
        time.sleep(login_wait_seconds)

        wait = WebDriverWait(driver, 30)
        file_stem = Path(file_name).stem
        selectors = [
            (By.PARTIAL_LINK_TEXT, file_stem),
            (By.XPATH, f"//*[contains(@aria-label, '{file_stem}')]"),
            (By.XPATH, f"//*[contains(@data-automationid, 'FieldRenderer-name')]//*[contains(text(), '{file_stem}')]"),
            (By.XPATH, f"//*[contains(text(), '{file_stem}')]"),
        ]

        clicked = False
        for by, value in selectors:
            try:
                element = wait.until(EC.element_to_be_clickable((by, value)))
                element.click()
                clicked = True
                logger.info("Opened file element via %s", value)
                break
            except Exception:
                continue

        if not clicked:
            print(
                f"\n>>> Could not auto-click {file_name!r}.\n"
                f">>> In the {browser_label} window: click the file, then click Download.\n"
                f">>> Waiting 90 more seconds for download to finish...\n"
            )
            time.sleep(90)

        time.sleep(3)

        for label in ("Download", "download"):
            try:
                btn = driver.find_element(By.PARTIAL_LINK_TEXT, label)
                btn.click()
                logger.info("Clicked Download")
                break
            except Exception:
                try:
                    btn = driver.find_element(
                        By.XPATH,
                        f"//button[contains(@aria-label, '{label}')]",
                    )
                    btn.click()
                    logger.info("Clicked Download button")
                    break
                except Exception:
                    continue

        deadline = time.time() + 120
        downloaded: Path | None = None
        while time.time() < deadline:
            for candidate in save_to.parent.glob("*.xlsx"):
                if file_stem.lower() in candidate.stem.lower():
                    if candidate.stat().st_size > 0:
                        downloaded = candidate
                        break
            if downloaded:
                break
            time.sleep(1)

        if not downloaded:
            raise RuntimeError(
                "Download did not complete. In Edge, open the file and choose Download, "
                f"then save manually to: {save_to}"
            )

        if downloaded.resolve() != save_to.resolve():
            downloaded.replace(save_to)

        logger.info("Saved workbook to %s", save_to)
        return save_to

    finally:
        driver.quit()


def upload_file_via_selenium(config: dict, local_path: Path, driver=None) -> None:
    """Upload a local file to SharePoint using Edge/Chrome Upload UI.

    Pass driver= to reuse the window already opened for sync (no second tab).
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait

    sp_cfg = config.get("sharepoint", {})
    folder_page_url = sp_cfg.get("folder_page_url")
    if not folder_page_url:
        raise ValueError("sharepoint.folder_page_url is required for Selenium upload.")

    browser = sp_cfg.get("browser", "edge")
    browser_path = sp_cfg.get("edge_path") if browser == "edge" else sp_cfg.get("chrome_path")
    login_wait = int(sp_cfg.get("upload_login_wait_seconds", sp_cfg.get("login_wait_seconds", 60)))
    upload_wait = int(sp_cfg.get("upload_complete_wait_seconds", 90))

    local_path = Path(local_path).resolve()
    if not local_path.exists():
        raise FileNotFoundError(f"File to upload not found: {local_path}")

    own_driver = driver is None
    browser_label = "Edge" if browser.lower() == "edge" else "Chrome"

    if own_driver:
        driver = _create_driver(browser, local_path.parent, browser_path, headless=False)

    try:
        if own_driver:
            logger.info("Opening SharePoint folder for upload (%s)", browser_label)
            driver.get(folder_page_url)
            print(
                f"\n>>> {browser_label}: confirm the 6 - TESTING folder is open.\n"
                f">>> Waiting {login_wait}s before upload...\n"
            )
            time.sleep(login_wait)
        else:
            logger.info("Reusing open %s window for upload", browser_label)
            print(f"\n>>> Using the same {browser_label} window for upload (no new tab).\n")
            try:
                driver.get(folder_page_url)
            except Exception:
                pass
            time.sleep(5)

        upload_clicked = False
        for by, value in (
            (By.XPATH, "//button[contains(@aria-label,'Upload')]"),
            (By.XPATH, "//button[contains(.,'Upload')]"),
            (By.CSS_SELECTOR, "button[name='Upload']"),
            (By.XPATH, "//*[contains(@data-automationid,'upload')]"),
        ):
            try:
                btn = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((by, value)))
                btn.click()
                upload_clicked = True
                logger.info("Clicked Upload via %s", value)
                break
            except Exception:
                continue

        if not upload_clicked:
            print(
                "\n>>> Could not auto-click Upload.\n"
                ">>> In Edge: click Upload -> Files, then select the report file manually.\n"
                f">>> File: {local_path}\n"
            )
            time.sleep(upload_wait)
            return

        time.sleep(2)
        file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
        if not file_inputs:
            raise RuntimeError("SharePoint upload file input not found after clicking Upload.")

        file_inputs[0].send_keys(str(local_path))
        print(f"\n>>> Uploading {local_path.name} ... waiting {upload_wait}s for completion.\n")
        time.sleep(upload_wait)
        logger.info("Selenium upload finished for %s", local_path.name)
        print(f"\n>>> Upload sent. Check the SharePoint folder in this Edge window.\n")
    finally:
        if own_driver:
            driver.quit()
