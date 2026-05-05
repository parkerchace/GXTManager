"""
GXTManager — Vertiv GXT UPS Management Tool
  Mode 1 — Battery Report:  login → Battery → scrape status/test/type → Communications → MAC → CSV
  Mode 2 — Firmware Upgrade: login → File Transfer → read version → optionally push firmware → CSV
"""

import csv
import os
import re
import subprocess
import urllib.parse
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import queue
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchWindowException
from webdriver_manager.firefox import GeckoDriverManager

LOG_QUEUE = queue.Queue()

# geckodriver resolved once; shared across all parallel workers
_GECKO_PATH: str | None = None

def _get_geckodriver() -> str:
    global _GECKO_PATH
    if _GECKO_PATH is None:
        path = GeckoDriverManager().install()
        real = os.path.realpath(path)
        try:
            os.chmod(real, 0o755)
        except Exception:
            pass
        # macOS Gatekeeper blocks unsigned binaries even after quarantine removal;
        # ad-hoc signing satisfies the check without a developer certificate.
        subprocess.run(["codesign", "-s", "-", "--force", real], capture_output=True)
        _GECKO_PATH = real
    return _GECKO_PATH

BATTERY_LINK_TIMEOUT = 90
TABLE_LOAD_TIMEOUT   = 20
FIRMWARE_XFER_TIMEOUT = 300   # 5 min for a firmware upload

BATTERY_FIELDS = {"UPS Battery Status", "Battery Test Result", "Battery Cabinet Type"}
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def log(msg: str) -> None:
    LOG_QUEUE.put(msg)

def _short_error(exc: Exception) -> str:
    first = str(exc).split("\n")[0].strip()
    return first or type(exc).__name__

def _clean_ip(raw: str) -> str:
    ip = raw.strip()
    ip = re.sub(r'^https?://', '', ip)
    ip = ip.split('/')[0]
    ip = ip.split('?')[0]
    ip = ip.strip()
    # If there's embedded text (e.g. "HOS-BR-UPS 10.70.96.15"), pull out the IPv4 address
    if ' ' in ip:
        m = re.search(r'\b(\d{1,3}(?:\.\d{1,3}){3})\b', ip)
        if m:
            ip = m.group(1)
    return ip

def _model_gen(model: str) -> str:
    """Return 'GXT5', 'GXT4', or '' based on model string."""
    m = model.upper()
    if "GXT5" in m:
        return "GXT5"
    if "GXT4" in m:
        return "GXT4"
    return ""


# ---------------------------------------------------------------------------
# Selenium helpers (shared by both modes)
# ---------------------------------------------------------------------------

def get_all_frames(driver) -> list:
    frames = []
    for tag in ("frame", "iframe"):
        try:
            frames += driver.find_elements(By.TAG_NAME, tag)
        except Exception:
            pass
    return frames


def find_element_anywhere(driver, by, value, timeout=30, label="", require_visible=True):
    deadline = time.time() + timeout
    tag = label or f"{by}={value!r}"

    def _search():
        try:
            el = driver.find_element(by, value)
            if not require_visible or el.is_displayed():
                return el
        except Exception:
            pass
        return None

    while time.time() < deadline:
        # Auto re-auth: if the device challenged us, re-login and re-navigate
        if _is_auth_page(driver):
            creds = getattr(driver, "_au_creds", None)
            if creds:
                loc, ip, user, pw = creds
                log(f"[{loc} | {ip}] Auth challenge detected — re-logging in ...")
                _login(driver, loc, ip, user, pw)
                nav_fn = getattr(driver, "_au_nav_fn", None)
                if nav_fn:
                    try:
                        nav_fn()
                    except Exception:
                        pass
                deadline = time.time() + timeout  # reset search window after re-auth
                time.sleep(1)
                continue

        driver.switch_to.default_content()
        el = _search()
        if el:
            return el
        for frame in get_all_frames(driver):
            try:
                driver.switch_to.frame(frame)
                el = _search()
                if el:
                    return el
            except Exception:
                pass
            driver.switch_to.default_content()
        time.sleep(1)

    raise TimeoutException(f"Element not found after {timeout}s: {tag}")


def js_click(driver, el) -> None:
    driver.execute_script("arguments[0].click();", el)


def scrape_detail_table(driver, timeout=20) -> tuple[str, list[dict]]:
    deadline = time.time() + timeout
    while time.time() < deadline:
        driver.switch_to.default_content()
        contexts = [None] + get_all_frames(driver)
        for ctx in contexts:
            if ctx is not None:
                try:
                    driver.switch_to.frame(ctx)
                except Exception:
                    driver.switch_to.default_content()
                    continue
            tables = driver.find_elements(By.CSS_SELECTOR, "table.detailTable")
            if tables:
                page_updated = ""
                for sel in ("span.lastUpdated","td.lastUpdated","div.lastUpdated","span.updated"):
                    try:
                        page_updated = driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                        if page_updated:
                            break
                    except Exception:
                        pass
                if not page_updated:
                    try:
                        for line in driver.find_element(By.TAG_NAME,"body").text.splitlines():
                            if "Updated:" in line:
                                page_updated = line.strip()
                                break
                    except Exception:
                        pass
                rows = []
                for table in tables:
                    for tr in table.find_elements(By.TAG_NAME, "tr"):
                        try:
                            lbl = tr.find_element(By.CSS_SELECTOR,"td[id^='label']").text.strip()
                            val = tr.find_element(By.CSS_SELECTOR,"td[id^='val']").text.strip()
                            uom = tr.find_element(By.CSS_SELECTOR,"td[id^='uom']").text.strip()
                            if lbl:
                                rows.append({"label": lbl, "value": val, "unit": uom})
                        except Exception:
                            continue
                if rows:
                    return page_updated, rows
            driver.switch_to.default_content()
        time.sleep(1)
    raise TimeoutException("detail table not found or remained empty")


def _make_driver() -> webdriver.Firefox:
    opts = Options()
    opts.headless = False
    opts.set_preference("acceptInsecureCerts", True)
    return webdriver.Firefox(service=Service(_get_geckodriver()), options=opts)


def _login(driver, location, ip, username, password):
    ip         = _clean_ip(ip)
    driver._au_creds = (location, ip, username, password)   # stored for auto-reauth
    login_path = "/web/initialize.htm?mode=newAuth"
    u = urllib.parse.quote(username, safe="")
    p = urllib.parse.quote(password, safe="")

    driver.set_page_load_timeout(30)
    for scheme in ("http", "https"):
        url      = f"{scheme}://{ip}{login_path}"
        auth_url = f"{scheme}://{u}:{p}@{ip}{login_path}"
        log(f"[{location} | {ip}] Navigating to {url} ...")
        try:
            driver.get(auth_url)
        except TimeoutException:
            log(f"[{location} | {ip}] Page load timed out — proceeding")
        except WebDriverException as exc:
            emsg = str(exc)
            if "connectionFailure" in emsg or "neterror" in emsg or "Reached error page" in emsg:
                log(f"[{location} | {ip}] {scheme.upper()} unreachable — trying HTTPS ...")
                continue
            raise
        time.sleep(2)
        cur = driver.current_url
        if "about:neterror" in cur or "connectionFailure" in cur:
            log(f"[{location} | {ip}] {scheme.upper()} unreachable — trying HTTPS ...")
            continue
        break   # page loaded on this scheme

    wait = WebDriverWait(driver, 15)
    try:
        user_field = wait.until(EC.element_to_be_clickable((By.ID, "username")))
        log(f"[{location} | {ip}] Filling login form ...")
        driver.execute_script("arguments[0].scrollIntoView(true);", user_field)
        time.sleep(0.3)
        user_field.clear()
        user_field.send_keys(username)
        pass_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
        driver.execute_script("arguments[0].scrollIntoView(true);", pass_field)
        time.sleep(0.3)
        pass_field.clear()
        pass_field.send_keys(password)
        try:
            submit = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR,"input[type='submit'],button[type='submit']")
            ))
            driver.execute_script("arguments[0].scrollIntoView(true);", submit)
            time.sleep(0.3)
            submit.click()
        except Exception:
            from selenium.webdriver.common.keys import Keys
            pass_field.send_keys(Keys.RETURN)
        log(f"[{location} | {ip}] Login submitted.")
    except TimeoutException:
        log(f"[{location} | {ip}] No login form — proceeding.")

    time.sleep(5)


def _is_auth_page(driver) -> bool:
    """Return True if the current page is a login/401 challenge."""
    try:
        cur = driver.current_url
        if "401" in cur or "unauthorized" in cur.lower() or "newAuth" in cur:
            return True
        driver.switch_to.default_content()
        for ctx in [None] + get_all_frames(driver):
            if ctx is not None:
                try: driver.switch_to.frame(ctx)
                except Exception:
                    driver.switch_to.default_content(); continue
            try:
                body = driver.find_element(By.TAG_NAME, "body").text.lower()
                if "401" in body or "unauthorized" in body or "authentication required" in body:
                    driver.switch_to.default_content()
                    return True
                if driver.find_elements(By.ID, "username") and driver.find_elements(By.ID, "password"):
                    driver.switch_to.default_content()
                    return True
            except Exception:
                pass
            driver.switch_to.default_content()
    except Exception:
        pass
    return False


def _read_model(driver, location, ip) -> str:
    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            el = find_element_anywhere(driver, By.CSS_SELECTOR, "#tab0 span",
                                       timeout=5, label="model", require_visible=False)
            model = el.text.strip()
            if model and model.lower() != "device0":
                log(f"[{location} | {ip}] Model: {model}")
                return model
        except Exception:
            pass
        time.sleep(2)
    log(f"[{location} | {ip}] Model not resolved — page may have loaded Communications-only.")
    return ""


# ---------------------------------------------------------------------------
# Mode 1 — Battery Report
# ---------------------------------------------------------------------------

def process_battery_ip(location: str, ip: str, username: str, password: str) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", status="unknown",
                  page_updated="", scraped_at="", battery_rows=[],
                  ethernet_mac="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)

        # debug frame dump
        driver.switch_to.default_content()
        frames = get_all_frames(driver)
        log(f"[{location} | {ip}] {len(frames)} frame(s) after login")
        for i, f in enumerate(frames):
            try:
                name = f.get_attribute("name") or f.get_attribute("id") or f.get_attribute("src") or "?"
                log(f"[{location} | {ip}]   frame[{i}] → {name}")
            except Exception:
                pass

        result["model"] = _read_model(driver, location, ip)

        # Battery
        battery_rows, page_updated = [], ""
        for attempt in range(1, 5):
            log(f"[{location} | {ip}] Battery attempt {attempt} ...")
            try:
                driver.switch_to.default_content()
                bat = find_element_anywhere(driver, By.ID, "report163860",
                                            timeout=BATTERY_LINK_TIMEOUT,
                                            label="Battery link", require_visible=False)
                js_click(driver, bat)
                time.sleep(2)
                page_updated, all_rows = scrape_detail_table(driver, timeout=TABLE_LOAD_TIMEOUT)
                battery_rows = [r for r in all_rows if r["label"] in BATTERY_FIELDS]
                if battery_rows and any(r["value"] for r in battery_rows):
                    break
                raise TimeoutException("Battery rows found but values still empty")
            except TimeoutException:
                log(f"[{location} | {ip}] Battery attempt {attempt} failed — waiting ...")
                driver.switch_to.default_content()
                time.sleep(8)

        if not battery_rows or all(not r["value"] for r in battery_rows):
            raise TimeoutException("Battery table empty after all retry attempts")
        log(f"[{location} | {ip}] {len(battery_rows)} battery fields captured.")

        # Communications → Active Networking → MAC
        log(f"[{location} | {ip}] Clicking Communications tab ...")
        driver.switch_to.default_content()
        js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=30,
                                               label="Communications tab", require_visible=False))
        time.sleep(2)

        log(f"[{location} | {ip}] Expanding Support ...")
        js_click(driver, find_element_anywhere(driver, By.ID, "164190Plus", timeout=30,
                                               label="Support expand", require_visible=False))
        time.sleep(2)

        log(f"[{location} | {ip}] Clicking Active Networking ...")
        js_click(driver, find_element_anywhere(driver, By.ID, "report164330", timeout=30,
                                               label="Active Networking", require_visible=False))
        time.sleep(2)

        log(f"[{location} | {ip}] Reading Ethernet MAC ...")
        mac_el = find_element_anywhere(driver, By.ID, "val6156_0", timeout=30,
                                       label="Ethernet MAC", require_visible=False)
        ethernet_mac = mac_el.text.strip()
        log(f"[{location} | {ip}] MAC: {ethernet_mac}")

        result.update(status="success", page_updated=page_updated,
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      battery_rows=battery_rows, ethernet_mac=ethernet_mac)
        time.sleep(2)

    except NoSuchWindowException:
        result.update(status="error", error="Browser window closed unexpectedly",
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] ERROR: window closed")
    except TimeoutException as exc:
        result.update(status="timeout", error=_short_error(exc),
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] TIMEOUT: {_short_error(exc)}")
    except WebDriverException as exc:
        result.update(status="error", error=_short_error(exc),
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] WebDriver error: {_short_error(exc)}")
    except Exception as exc:
        result.update(status="error", error=_short_error(exc),
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] Error: {_short_error(exc)}")
    finally:
        if driver:
            try: driver.quit()
            except Exception: pass
    return result


def _build_battery_csv(results: list[dict]) -> str:
    lu: dict[str, str] = {}
    for r in results:
        for row in r["battery_rows"]:
            if row["label"] not in lu:
                lu[row["label"]] = row["unit"]
    metric_cols = [(f"{lbl} ({unit})" if unit else lbl, lbl) for lbl, unit in lu.items()]
    fixed = ["Location","IP","Model","Ethernet MAC","Page Updated","Scraped At","Status","Error"]
    path = os.path.join(SCRIPT_DIR, f"battery_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed+[c for c,_ in metric_cols], extrasaction="ignore")
        w.writeheader()
        for r in results:
            row = dict(Location=r["location"], IP=r["ip"], Model=r["model"],
                       **{"Ethernet MAC": r["ethernet_mac"]},
                       **{"Page Updated": r["page_updated"]},
                       **{"Scraped At": r["scraped_at"]},
                       Status=r["status"], Error=r["error"])
            vm = {br["label"]: br["value"] for br in r["battery_rows"]}
            for col, lbl in metric_cols:
                row[col] = vm.get(lbl, "")
            w.writerow(row)
    return path


def run_battery_scraper(targets, username, password, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_battery_ip(loc, ip, username, password)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
                s += f"  — model={r['model']}, MAC={r['ethernet_mac']}" if r["status"]=="success" else f"  — {r['error']}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_battery_csv(results)
    log(f"\nCSV saved: {path}")
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
        s += f"  — model={r['model']}, MAC={r['ethernet_mac']}" if r["status"]=="success" else f"  — {r['error']}"
        log(s)
    log("Done.")


# ---------------------------------------------------------------------------
# Mode 2 — Firmware Upgrade
# ---------------------------------------------------------------------------

def _wait_for_device_online(driver, location: str, ip: str, timeout: int = 480) -> bool:
    """Poll until the login form is visible. Keeps retrying through page load timeouts and
    reboot holding pages — expected behaviour for several minutes after a firmware upgrade."""
    log(f"[{location} | {ip}] Waiting for login page to be ready (up to {timeout}s) ...")
    deadline = time.time() + timeout
    login_path = "/web/initialize.htm?mode=newAuth"
    attempt = 0
    while time.time() < deadline:
        attempt += 1
        for scheme in ("http", "https"):
            try:
                # Stop any pending load before navigating
                try: driver.execute_script("window.stop();")
                except Exception: pass

                driver.set_page_load_timeout(20)
                driver.get(f"{scheme}://{ip}{login_path}")
                time.sleep(3)

                cur = driver.current_url
                if "neterror" in cur or "connectionFailure" in cur:
                    continue

                body = ""
                try: body = driver.find_element(By.TAG_NAME, "body").text
                except Exception: pass

                # Still on reboot holding page — keep waiting
                if "attempting to reconnect" in body.lower() or "web card has been rebooted" in body.lower():
                    log(f"[{location} | {ip}] Still rebooting (attempt {attempt}) ...")
                    break

                # Wait for login form fields to appear
                try:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "username")))
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
                    log(f"[{location} | {ip}] Login page ready.")
                    return True
                except Exception:
                    log(f"[{location} | {ip}] Page responded but login form not yet visible (attempt {attempt}) ...")
                    break

            except TimeoutException:
                log(f"[{location} | {ip}] Page load timed out (attempt {attempt}) — retrying ...")
            except Exception:
                pass   # connection refused / neterror — keep polling

        time.sleep(12)

    log(f"[{location} | {ip}] Login page did not appear within {timeout}s.")
    return False


def _click_run_alternate(driver, location: str, ip: str) -> bool:
    """Click Enable then Run Alternate and confirm the dialog. Returns True if clicked."""
    try:
        enable_btn = find_element_anywhere(driver, By.ID, "enableComms", timeout=20,
                                           label="Enable", require_visible=False)
        driver.execute_script("arguments[0].scrollIntoView(true);", enable_btn)
        time.sleep(0.5)
        enable_btn.click()
        log(f"[{location} | {ip}] Enable clicked — waiting for Run Alternate to activate ...")

        run_alt = None
        deadline_btn = time.time() + 20
        while time.time() < deadline_btn:
            try:
                btn = find_element_anywhere(driver, By.ID, "commBtn263", timeout=5,
                                            label="Run Alternate", require_visible=False)
                disabled = btn.get_attribute("disabled")
                if disabled is None or disabled.lower() in ("false", ""):
                    run_alt = btn
                    break
            except Exception:
                pass
            time.sleep(1)

        if not run_alt:
            log(f"[{location} | {ip}] Run Alternate button never became enabled.")
            return False

        driver.execute_script("arguments[0].scrollIntoView(true);", run_alt)
        time.sleep(0.5)
        run_alt.click()

        try:
            alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert.accept()
            log(f"[{location} | {ip}] Run Alternate confirmed — device will reboot.")
        except Exception:
            log(f"[{location} | {ip}] Run Alternate clicked (no confirmation dialog appeared).")
        return True
    except Exception as exc:
        log(f"[{location} | {ip}] _click_run_alternate failed: {_short_error(exc)}")
        return False


def _nav_to_firmware_page(driver, location: str, ip: str,
                          username: str = "", password: str = "") -> None:
    """Navigate from post-login home to the Firmware Update detail page.
    Re-authenticates automatically if the device challenges mid-navigation."""
    def _reauth_if_needed():
        if username and _is_auth_page(driver):
            log(f"[{location} | {ip}] Auth challenge detected — re-logging in ...")
            _login(driver, location, ip, username, password)

    log(f"[{location} | {ip}] Navigating to Firmware Update page ...")
    _reauth_if_needed()

    # Some devices load only the Communications tab without the full UPS nav structure.
    # Try clicking it; if it isn't present we're likely already in the right context.
    try:
        js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=10,
                                               label="Communications tab", require_visible=False))
        time.sleep(2)
        _reauth_if_needed()
    except TimeoutException:
        log(f"[{location} | {ip}] Communications tab (#tab4) not found — proceeding without it ...")

    try:
        js_click(driver, find_element_anywhere(driver, By.ID, "report164190", timeout=15,
                                               label="Support", require_visible=False))
        time.sleep(2)
        _reauth_if_needed()
    except TimeoutException:
        log(f"[{location} | {ip}] Support link not found — looking for Firmware Update link directly ...")

    fw_link = find_element_anywhere(
        driver,
        By.XPATH, "//a[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'firmware')]",
        timeout=30, label="Firmware Update link", require_visible=False)
    js_click(driver, fw_link)
    time.sleep(2)


def _read_fw_version_from_page(driver, location: str, ip: str) -> str:
    """Read 'Current Firmware Version' from the Firmware Update detailTable."""
    try:
        _, rows = scrape_detail_table(driver, timeout=20)
        for r in rows:
            if "current firmware version" in r["label"].lower():
                log(f"[{location} | {ip}] Firmware version: {r['value']!r}")
                return r["value"]
        # If row not found, log all labels seen to help debug
        labels = [r["label"] for r in rows]
        log(f"[{location} | {ip}] 'Current Firmware Version' not found; labels seen: {labels}")
    except Exception as exc:
        log(f"[{location} | {ip}] Could not read firmware version: {_short_error(exc)}")
    return ""


def process_firmware_ip(location: str, ip: str, username: str, password: str,
                        upgrade_mode: str,       # "check_only" | "check_and_upgrade" | "force_upgrade"
                        target_version: str,
                        gtx4_file: str, gtx5_file: str) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", gen="",
                  current_version="", target_version=target_version,
                  upgrade_mode=upgrade_mode, upgrade_applied=False,
                  upload_status="", verified_version="", scraped_at="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)
        result["model"] = _read_model(driver, location, ip)
        result["gen"]   = _model_gen(result["model"])

        _nav_to_firmware_page(driver, location, ip, username, password)
        # After reaching the firmware page, store a nav callback so auto-reauth
        # in find_element_anywhere knows how to get back here after re-login
        driver._au_nav_fn = lambda: _nav_to_firmware_page(driver, location, ip, username, password)

        # Read current firmware version from the Firmware Update page
        log(f"[{location} | {ip}] Reading firmware version ...")
        current = _read_fw_version_from_page(driver, location, ip)
        result["current_version"] = current
        log(f"[{location} | {ip}] Current firmware: {current!r}")

        # Decide whether to upgrade
        do_upgrade = False
        if upgrade_mode == "force_upgrade":
            do_upgrade = True
        elif upgrade_mode == "check_and_upgrade":
            tv = " ".join(target_version.split())
            do_upgrade = (current != tv)
            if not do_upgrade:
                log(f"[{location} | {ip}] Already at target version — skipping upgrade.")

        if do_upgrade:
            gen = result["gen"]
            fw_file = gtx5_file if gen == "GXT5" else gtx4_file if gen == "GXT4" else ""
            if not fw_file:
                # Model not resolved (Communications-only page load) — use whichever file was provided
                fw_file = gtx4_file or gtx5_file
                if fw_file:
                    log(f"[{location} | {ip}] Model unknown — using fallback firmware file: {os.path.basename(fw_file)}")
                else:
                    raise ValueError(f"No firmware file selected for model gen '{gen}' ({result['model']})")
            if not os.path.isfile(fw_file):
                raise FileNotFoundError(f"Firmware file not found: {fw_file}")

            upload_status = ""
            recovery_ok = True
            used_recovery = False
            for attempt in range(1, 3):   # up to 2 attempts
                if attempt == 2 and not recovery_ok:
                    log(f"[{location} | {ip}] Recovery failed — not retrying upload.")
                    break

                if attempt == 2:
                    # ── Recovery ────────────────────────────────────────────────────
                    # Step 1: wait for device, sign in, navigate to firmware page
                    log(f"[{location} | {ip}] [RECOVERY 1/4] Waiting for device ...")
                    if not _wait_for_device_online(driver, location, ip, timeout=90):
                        log(f"[{location} | {ip}] Device unreachable — giving up.")
                        recovery_ok = False; break
                    _login(driver, location, ip, username, password)
                    log(f"[{location} | {ip}] [RECOVERY 1/4] Signed in.")
                    try:
                        _nav_to_firmware_page(driver, location, ip, username, password)
                    except Exception as exc:
                        log(f"[{location} | {ip}] [RECOVERY 1/4] Navigation failed: {_short_error(exc)} — giving up.")
                        recovery_ok = False; break

                    # Step 2: check if the upload actually landed despite the redirect/error
                    log(f"[{location} | {ip}] [RECOVERY 2/4] Checking current firmware version ...")
                    version_now = _read_fw_version_from_page(driver, location, ip)
                    tv_check = " ".join(target_version.split())
                    if version_now and version_now == tv_check:
                        log(f"[{location} | {ip}] [RECOVERY 2/4] Upload succeeded — version confirmed: {version_now}")
                        upload_status = "firmware update successful"
                        break  # falls through to post-loop verification

                    log(f"[{location} | {ip}] [RECOVERY 2/4] Version still {version_now!r} — device needs recovery.")

                    # Step 3: Run Alternate to clear device state, then reboot
                    log(f"[{location} | {ip}] [RECOVERY 3/4] Running Run Alternate to clear device state ...")
                    if not _click_run_alternate(driver, location, ip):
                        log(f"[{location} | {ip}] [RECOVERY 3/4] Run Alternate failed — giving up.")
                        recovery_ok = False; break

                    # Step 4: wait for reboot → sign back in → navigate for retry upload
                    log(f"[{location} | {ip}] [RECOVERY 4/4] Waiting for device to reboot ...")
                    if not _wait_for_device_online(driver, location, ip, timeout=480):
                        log(f"[{location} | {ip}] [RECOVERY 4/4] Device did not come back — giving up.")
                        recovery_ok = False; break
                    log(f"[{location} | {ip}] [RECOVERY 4/4] Signing back in ...")
                    _login(driver, location, ip, username, password)
                    try:
                        _nav_to_firmware_page(driver, location, ip, username, password)
                        log(f"[{location} | {ip}] [RECOVERY 4/4] Ready — retrying upload ...")
                    except Exception as exc:
                        log(f"[{location} | {ip}] [RECOVERY 4/4] Navigation failed: {_short_error(exc)} — giving up.")
                        recovery_ok = False; break
                    used_recovery = True
                    # ── End recovery ────────────────────────────────────────────────

                log(f"[{location} | {ip}] Clicking Enable (attempt {attempt}) ...")
                js_click(driver, find_element_anywhere(driver, By.ID, "enableComms", timeout=20,
                                                       label="Enable button", require_visible=False))
                time.sleep(2)

                log(f"[{location} | {ip}] Clicking Web ...")
                js_click(driver, find_element_anywhere(driver, By.ID, "webFwUpdateBtn", timeout=20,
                                                       label="Web button", require_visible=False))
                time.sleep(3)

                log(f"[{location} | {ip}] Uploading {os.path.basename(fw_file)} (attempt {attempt}) ...")
                file_input = find_element_anywhere(
                    driver,
                    By.CSS_SELECTOR, 'input[id="Firmware File Upload"]',
                    timeout=20, label="firmware file input", require_visible=False)
                file_input.send_keys(os.path.abspath(fw_file))
                time.sleep(1)

                try:
                    submit = find_element_anywhere(
                        driver,
                        By.CSS_SELECTOR, "input[type='submit'], button[type='submit'], input[type='button'][value*='Upload'], input[type='button'][value*='Update']",
                        timeout=10, label="upload submit", require_visible=False)
                    js_click(driver, submit)
                except TimeoutException:
                    driver.execute_script("var f = document.querySelector('form'); if(f) f.submit();")

                log(f"[{location} | {ip}] Upload submitted — waiting for result ...")
                upload_status = _wait_for_transfer(driver, location, ip)
                log(f"[{location} | {ip}] Transfer result: {upload_status}")

                _upload_failed = (
                    "device error" in upload_status.lower()
                    or "session expired" in upload_status.lower()
                    or upload_status == "timed out waiting for transfer completion"
                )
                if not _upload_failed:
                    break   # success — don't retry

            result["upload_status"]   = upload_status
            result["upgrade_applied"] = "error" not in upload_status.lower() and "timed out" not in upload_status

            if result["upgrade_applied"]:
                # After a recovery upload the new firmware lands in the alternate slot and the
                # device reboots back to the OLD version. Run Alternate one more time to activate it.
                if used_recovery:
                    log(f"[{location} | {ip}] Recovery upload complete — activating new firmware via Run Alternate ...")
                    if _wait_for_device_online(driver, location, ip, timeout=600):
                        _login(driver, location, ip, username, password)
                        try:
                            _nav_to_firmware_page(driver, location, ip, username, password)
                            if _click_run_alternate(driver, location, ip):
                                log(f"[{location} | {ip}] Run Alternate clicked — waiting for final reboot ...")
                            else:
                                log(f"[{location} | {ip}] Could not click Run Alternate for activation.")
                        except Exception as exc:
                            log(f"[{location} | {ip}] Activation step failed: {_short_error(exc)}")

                # Wait for device, sign in, and verify version
                log(f"[{location} | {ip}] Waiting for device to come back up after upgrade ...")
                if _wait_for_device_online(driver, location, ip, timeout=600):
                    _login(driver, location, ip, username, password)
                    try:
                        _nav_to_firmware_page(driver, location, ip, username, password)
                        verified = _read_fw_version_from_page(driver, location, ip)
                        result["verified_version"] = verified
                        tv = " ".join(target_version.split())
                        if verified and verified == tv:
                            log(f"[{location} | {ip}] Version confirmed: {verified} — upgrade successful.")
                        elif verified:
                            log(f"[{location} | {ip}] Version mismatch: got {verified}, expected {tv}")
                        else:
                            log(f"[{location} | {ip}] Could not read version after upgrade.")
                    except Exception as exc:
                        log(f"[{location} | {ip}] Version check failed: {_short_error(exc)}")
        else:
            result["upload_status"] = "skipped"

        result["scraped_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    except NoSuchWindowException:
        result.update(error="Browser window closed unexpectedly",
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] ERROR: window closed")
    except (TimeoutException, WebDriverException, ValueError, FileNotFoundError, Exception) as exc:
        result.update(error=_short_error(exc),
                      scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] ERROR: {_short_error(exc)}")
        if driver:
            try: driver.quit()
            except Exception: pass
    # Browser window intentionally left open after firmware run so you can verify the result
    return result


def _wait_for_transfer(driver, location: str, ip: str, timeout: int = FIRMWARE_XFER_TIMEOUT) -> str:
    """Poll after submitting firmware upload. Handles success page countdown and Go Home button."""
    deadline = time.time() + timeout
    original_url = driver.current_url
    while time.time() < deadline:
        time.sleep(5)
        try:
            # Session expired mid-upload — device redirected browser to login page
            if _is_auth_page(driver):
                log(f"[{location} | {ip}] Redirected to login during transfer — session expired.")
                return "session expired during transfer"

            url_changed = driver.current_url != original_url
            driver.switch_to.default_content()
            body_text = ""
            for ctx in [None] + get_all_frames(driver):
                if ctx is not None:
                    try: driver.switch_to.frame(ctx)
                    except Exception:
                        driver.switch_to.default_content(); continue
                try:
                    body_text += driver.find_element(By.TAG_NAME, "body").text
                except Exception:
                    pass
                driver.switch_to.default_content()

            body_lower = body_text.lower()

            # Device-side error (503 / write failure)
            if "503" in body_text or "error writing" in body_lower or "service unavailable" in body_lower:
                for line in body_text.splitlines():
                    if line.strip():
                        return f"device error — {line.strip()[:200]}"
                return "device error — 503 Service Unavailable"

            # Success page — "FIRMWARE UPDATE SUCCESSFUL ... Restarting... N seconds"
            if "firmware update successful" in body_lower:
                log(f"[{location} | {ip}] Firmware update successful — waiting for restart countdown ...")
                # Wait for GoHomeB button to become enabled (countdown reaches 0)
                home_deadline = time.time() + 300
                while time.time() < home_deadline:
                    time.sleep(5)
                    try:
                        driver.switch_to.default_content()
                        btn = find_element_anywhere(driver, By.ID, "GoHomeB", timeout=5,
                                                    label="Go Home button", require_visible=False)
                        disabled = btn.get_attribute("disabled")
                        if disabled is None or disabled.lower() in ("false", ""):
                            log(f"[{location} | {ip}] Clicking 'Go Home' ...")
                            btn.click()
                            break
                        # Log remaining time from page if visible
                        try:
                            remaining = [l for l in body_text.splitlines() if "restarting" in l.lower() or "second" in l.lower()]
                            if remaining:
                                log(f"[{location} | {ip}] {remaining[-1].strip()}")
                        except Exception:
                            pass
                    except Exception:
                        pass
                return "firmware update successful"

            if any(w in body_lower for w in ("failed", "invalid", "rejected")):
                return f"error — {body_text[:200]}"

            if url_changed:
                original_url = driver.current_url

        except Exception:
            return "browser closed during transfer"
    return "timed out waiting for transfer completion"


def _build_firmware_csv(results: list[dict]) -> str:
    fixed = ["Location","IP","Model","Gen","Current Version","Target Version",
             "Upgrade Mode","Upgrade Applied","Upload Status","Verified Version","Scraped At","Error"]
    path = os.path.join(SCRIPT_DIR, f"firmware_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed, extrasaction="ignore")
        w.writeheader()
        for r in results:
            if r is None:
                continue
            w.writerow({
                "Location":       r["location"],
                "IP":             r["ip"],
                "Model":          r["model"],
                "Gen":            r["gen"],
                "Current Version":r["current_version"],
                "Target Version": r["target_version"],
                "Upgrade Mode":   r["upgrade_mode"],
                "Upgrade Applied":r["upgrade_applied"],
                "Upload Status":   r["upload_status"],
                "Verified Version":r.get("verified_version",""),
                "Scraped At":      r["scraped_at"],
                "Error":          r["error"],
            })
    return path


def run_firmware_scraper(targets, username, password, upgrade_mode,
                         target_version, gtx4_file, gtx5_file, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_firmware_ip(loc, ip, username, password,
                                        upgrade_mode, target_version, gtx4_file, gtx5_file)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['model']}  v={r['current_version']}  upload={r['upload_status'] or 'n/a'}  err={r['error'] or 'none'}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_firmware_csv(results)
    log(f"\nCSV saved: {path}")
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        log(f"[{r['location']} | {r['ip']}] {r['model']}  current={r['current_version']}  applied={r['upgrade_applied']}  status={r['upload_status'] or r['error'] or 'ok'}")
    log("Done.")


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GXTManager")
        self.resizable(True, True)
        self._build_ui()
        self._poll_log()

    # ── Layout ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}
        self.columnconfigure(0, weight=1)

        # ── Credentials ──
        creds = ttk.LabelFrame(self, text="Credentials  (never stored)")
        creds.grid(row=0, column=0, sticky="ew", padx=8, pady=(10,4))
        ttk.Label(creds, text="Username:").grid(row=0, column=0, sticky="e", **pad)
        self.username_var = tk.StringVar()
        ue = ttk.Entry(creds, textvariable=self.username_var, width=30)
        ue.grid(row=0, column=1, sticky="w", **pad)
        self._bind_paste(ue)
        ttk.Label(creds, text="Password:").grid(row=1, column=0, sticky="e", **pad)
        self.password_var = tk.StringVar()
        pe = ttk.Entry(creds, textvariable=self.password_var, show="*", width=30)
        pe.grid(row=1, column=1, sticky="w", **pad)
        self._bind_paste(pe)

        # ── Mode notebook ──
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=1, column=0, sticky="ew", padx=8, pady=4)

        # Tab 1 — Battery (no extra options)
        bat_tab = ttk.Frame(self.notebook)
        self.notebook.add(bat_tab, text="Battery Report")
        ttk.Label(bat_tab, text="Logs UPS Battery Status, Battery Test Result, Battery Cabinet Type, and Ethernet MAC.",
                  foreground="gray").pack(padx=8, pady=6)

        # Tab 2 — Firmware
        fw_tab = ttk.Frame(self.notebook)
        self.notebook.add(fw_tab, text="Firmware Upgrade")
        self._build_firmware_tab(fw_tab, pad)

        # ── Targets ──
        tgt = ttk.LabelFrame(self, text="Targets — paste two columns from Excel:  Location  [tab]  IP Address")
        tgt.grid(row=2, column=0, sticky="nsew", **pad)
        self.rowconfigure(2, weight=1)
        tgt.columnconfigure(0, weight=1)
        tgt.rowconfigure(1, weight=1)
        hdr = ttk.Frame(tgt)
        hdr.grid(row=0, column=0, sticky="ew", padx=8, pady=(4,0))
        ttk.Label(hdr, text="Location", foreground="gray", font=("Courier",10,"bold"), width=22).pack(side="left")
        ttk.Label(hdr, text="IP Address", foreground="gray", font=("Courier",10,"bold")).pack(side="left")
        self.target_text = scrolledtext.ScrolledText(tgt, width=52, height=10, font=("Courier",11))
        self.target_text.grid(row=1, column=0, sticky="nsew", **pad)
        self._bind_paste(self.target_text)

        # ── Controls ──
        ctrl = ttk.Frame(self)
        ctrl.grid(row=3, column=0, sticky="ew", padx=8, pady=(4,4))
        self.start_btn = ttk.Button(ctrl, text="Start", command=self._start)
        self.start_btn.pack(side="left", padx=4)
        ttk.Button(ctrl, text="Clear Log", command=self._clear_log).pack(side="left", padx=4)
        ttk.Label(ctrl, text="Parallel:").pack(side="left", padx=(12,2))
        self.parallel_var = tk.IntVar(value=3)
        ttk.Spinbox(ctrl, from_=1, to=10, width=3, textvariable=self.parallel_var).pack(side="left")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(ctrl, textvariable=self.status_var, foreground="gray").pack(side="right", padx=8)

        # ── Log ──
        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.grid(row=4, column=0, sticky="nsew", padx=8, pady=(0,8))
        self.rowconfigure(4, weight=2)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_box = scrolledtext.ScrolledText(
            log_frame, width=70, height=14, state="disabled",
            font=("Courier",10), background="#1e1e1e", foreground="#d4d4d4")
        self.log_box.grid(row=0, column=0, sticky="nsew", **pad)

    def _build_firmware_tab(self, parent, pad):
        # Firmware files
        files_frame = ttk.LabelFrame(parent, text="Firmware Files")
        files_frame.grid(row=0, column=0, sticky="ew", **pad)
        files_frame.columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="GTX-5 file:").grid(row=0, column=0, sticky="e", **pad)
        self.gtx5_var = tk.StringVar()
        gtx5_entry = ttk.Entry(files_frame, textvariable=self.gtx5_var, width=45)
        gtx5_entry.grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(files_frame, text="Browse…",
                   command=lambda: self._browse_file(self.gtx5_var)).grid(row=0, column=2, **pad)

        ttk.Label(files_frame, text="GTX-4 file:").grid(row=1, column=0, sticky="e", **pad)
        self.gtx4_var = tk.StringVar()
        gtx4_entry = ttk.Entry(files_frame, textvariable=self.gtx4_var, width=45)
        gtx4_entry.grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(files_frame, text="Browse…",
                   command=lambda: self._browse_file(self.gtx4_var)).grid(row=1, column=2, **pad)

        # Upgrade mode
        mode_frame = ttk.LabelFrame(parent, text="Upgrade Mode")
        mode_frame.grid(row=1, column=0, sticky="ew", **pad)
        self.fw_mode_var = tk.StringVar(value="check_only")
        modes = [
            ("Check version only",                "check_only"),
            ("Check and upgrade if outdated",      "check_and_upgrade"),
            ("Force upgrade (skip version check)", "force_upgrade"),
        ]
        for i, (label, val) in enumerate(modes):
            ttk.Radiobutton(mode_frame, text=label, variable=self.fw_mode_var,
                            value=val, command=self._on_fw_mode).grid(
                row=i, column=0, sticky="w", padx=12, pady=2)

        # Target version (shown only for check_and_upgrade)
        self.tv_frame = ttk.Frame(mode_frame)
        self.tv_frame.grid(row=len(modes), column=0, sticky="w", padx=12, pady=(0,4))
        ttk.Label(self.tv_frame, text="Target version string:").pack(side="left")
        self.target_ver_var = tk.StringVar()
        ttk.Entry(self.tv_frame, textvariable=self.target_ver_var, width=30).pack(side="left", padx=4)
        self.tv_frame.grid_remove()   # hidden until check_and_upgrade selected

    def _on_fw_mode(self):
        if self.fw_mode_var.get() == "check_and_upgrade":
            self.tv_frame.grid()
        else:
            self.tv_frame.grid_remove()

    def _browse_file(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            title="Select firmware file",
            filetypes=[("All files", "*")]
        )
        if path:
            var.set(path)

    # ── Paste fix ───────────────────────────────────────────────────────────

    def _bind_paste(self, widget):
        pending = [False]
        def on_paste(event=None):
            if pending[0]: return "break"
            pending[0] = True
            self.after(10, do_paste)
            return "break"
        def do_paste():
            pending[0] = False
            try:
                text = self.clipboard_get()
            except tk.TclError:
                return
            if isinstance(widget, tk.Text):
                try: widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError: pass
                widget.insert(tk.INSERT, text)
            else:
                try:
                    if widget.selection_present():
                        widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError: pass
                widget.insert(tk.INSERT, text)
        widget.bind("<<Paste>>", on_paste)
        widget.bind("<Command-v>", on_paste)

    # ── Targets parser ───────────────────────────────────────────────────────

    def _parse_targets(self):
        raw = self.target_text.get("1.0", "end")
        targets = []
        for line in raw.splitlines():
            line = line.strip()
            if not line: continue
            parts = line.split("\t")
            if len(parts) >= 2:
                location, ip = parts[0].strip(), _clean_ip(parts[1])
            else:
                raw_val = parts[0].strip()
                ip = _clean_ip(raw_val)
                # If text precedes the IP (e.g. "HOS-BR-UPS 10.70.96.15"), use it as location
                m = re.search(r'\b(\d{1,3}(?:\.\d{1,3}){3})\b', raw_val)
                location = raw_val[:m.start()].strip() if (m and m.start() > 0) else ""
            if ip:
                targets.append((location, ip))
        return targets

    # ── Start ────────────────────────────────────────────────────────────────

    def _start(self):
        username = self.username_var.get().strip()
        password = self.password_var.get()
        targets  = self._parse_targets()
        parallel = max(1, min(10, self.parallel_var.get()))
        mode_tab = self.notebook.index(self.notebook.select())  # 0=battery, 1=firmware

        if not username:
            messagebox.showwarning("Missing input", "Please enter a username."); return
        if not password:
            messagebox.showwarning("Missing input", "Please enter a password."); return
        if not targets:
            messagebox.showwarning("Missing input", "Please paste at least one row of targets."); return

        if mode_tab == 1:  # Firmware
            fw_mode    = self.fw_mode_var.get()
            target_ver = self.target_ver_var.get().strip()
            gtx5_file  = self.gtx5_var.get().strip()
            gtx4_file  = self.gtx4_var.get().strip()
            if fw_mode == "check_and_upgrade" and not target_ver:
                messagebox.showwarning("Missing input", "Enter a target version string."); return
            if fw_mode in ("check_and_upgrade","force_upgrade") and not gtx5_file and not gtx4_file:
                messagebox.showwarning("Missing input", "Select at least one firmware file."); return

            self.start_btn.config(state="disabled")
            self.status_var.set(f"Firmware — {len(targets)} device(s) ...")
            log(f"Starting firmware run ({fw_mode}) on {len(targets)} device(s) ...")
            def worker():
                run_firmware_scraper(targets, username, password,
                                     fw_mode, target_ver, gtx4_file, gtx5_file, parallel)
                LOG_QUEUE.put("__DONE__")
        else:  # Battery
            self.start_btn.config(state="disabled")
            self.status_var.set(f"Battery report — {len(targets)} device(s) ...")
            log(f"Starting battery scrape of {len(targets)} device(s) ...")
            def worker():
                run_battery_scraper(targets, username, password, parallel)
                LOG_QUEUE.put("__DONE__")

        threading.Thread(target=worker, daemon=True).start()

    def _clear_log(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

    def _poll_log(self):
        try:
            while True:
                msg = LOG_QUEUE.get_nowait()
                if msg == "__DONE__":
                    self.start_btn.config(state="normal")
                    self.status_var.set("Finished — CSV saved next to this script")
                else:
                    self.log_box.config(state="normal")
                    self.log_box.insert("end", msg + "\n")
                    self.log_box.see("end")
                    self.log_box.config(state="disabled")
        except queue.Empty:
            pass
        self.after(200, self._poll_log)


if __name__ == "__main__":
    app = App()
    app.mainloop()
