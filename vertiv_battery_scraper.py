"""
GXTManager — Vertiv GXT UPS Management Tool
  Mode 1 — Battery Report:  login → Battery → scrape status/test/type → Communications → MAC → CSV
  Mode 2 — Firmware Upgrade: login → File Transfer → read version → optionally push firmware → CSV
"""

import csv
import json
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
from selenium.webdriver.support.ui import WebDriverWait, Select as SeleniumSelect
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchWindowException
from selenium.webdriver.common.action_chains import ActionChains
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

BATTERY_LINK_TIMEOUT  = 120   # time to wait for the Battery nav link to appear
TABLE_LOAD_TIMEOUT    = 90    # time to wait for the detail table to fully populate
FIRMWARE_XFER_TIMEOUT = 300   # 5 min for a firmware upload
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _config_path() -> str:
    """Return the path to the config file saved on the user's Desktop (works on Mac and Windows)."""
    return os.path.join(os.path.expanduser("~"), "Desktop", "GXTManager_config.json")


def _open_file(path: str) -> None:
    """Open a file with the system default app (works on macOS, Windows, and Linux)."""
    try:
        if os.sys.platform == "darwin":
            subprocess.run(["open", path])
        elif os.sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.run(["xdg-open", path])
    except Exception:
        pass


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


def scrape_battery_page(driver, timeout=90) -> tuple[str, list[dict]]:
    """Scrape battery status and event data from the Vertiv battery detail page.

    The battery page has two useful tables inside #detailPanelArea:
      - #statusTable  - metrics like charge %, voltage, temperature (label/val/uom columns)
      - #eventTable   - alarm events like Replace Battery, Battery Low (label/evtStatus_ columns)

    Both are captured so the CSV has the full picture.
    """
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

            # The battery page wraps everything in #detailPanelArea; bail if it isn't here yet
            try:
                panel = driver.find_element(By.CSS_SELECTOR, "#detailPanelArea")
            except Exception:
                driver.switch_to.default_content()
                continue

            page_updated = ""
            for sel in ("span.lastUpdated", "td.lastUpdated", "div.lastUpdated", "span.updated"):
                try:
                    page_updated = driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                    if page_updated:
                        break
                except Exception:
                    pass
            if not page_updated:
                try:
                    for line in driver.find_element(By.TAG_NAME, "body").text.splitlines():
                        if "Updated:" in line:
                            page_updated = line.strip()
                            break
                except Exception:
                    pass

            rows = []

            # Pull every metric row out of #statusTable (val column holds the live reading)
            try:
                status_tbl = panel.find_element(By.CSS_SELECTOR, "#statusTable table")
                for tr in status_tbl.find_elements(By.TAG_NAME, "tr"):
                    try:
                        lbl = tr.find_element(By.CSS_SELECTOR, "td[id^='label']").text.strip()
                        val = tr.find_element(By.CSS_SELECTOR, "td[id^='val']").text.strip()
                        uom = ""
                        try:
                            uom = tr.find_element(By.CSS_SELECTOR, "td[id^='uom']").text.strip()
                        except Exception:
                            pass
                        if lbl:
                            rows.append({"label": lbl, "value": val, "unit": uom})
                    except Exception:
                        continue
            except Exception:
                pass

            # Pull alarm events from #eventTable (status column uses evtStatus_ not val_)
            try:
                event_tbl = panel.find_element(By.CSS_SELECTOR, "#eventTable table")
                for tr in event_tbl.find_elements(By.TAG_NAME, "tr"):
                    try:
                        lbl = tr.find_element(By.CSS_SELECTOR, "td[id^='label']").text.strip()
                        val = tr.find_element(By.CSS_SELECTOR, "td[id^='evtStatus_']").text.strip()
                        if lbl and val:
                            rows.append({"label": lbl, "value": val, "unit": ""})
                    except Exception:
                        continue
            except Exception:
                pass

            driver.switch_to.default_content()
            if rows and any(r["value"] for r in rows):
                return page_updated, rows

        time.sleep(1)
    raise TimeoutException("Battery detail panel not found or remained empty after timeout")


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

    driver.set_page_load_timeout(60)  # some NICs are very slow to serve even the login page
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
        time.sleep(5)  # let the page finish rendering before we check the URL
        cur = driver.current_url
        if "about:neterror" in cur or "connectionFailure" in cur:
            log(f"[{location} | {ip}] {scheme.upper()} unreachable — trying HTTPS ...")
            continue
        break   # page loaded on this scheme

    wait = WebDriverWait(driver, 45)  # slow NICs can take a long time to render the login form
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

    time.sleep(10)  # wait for the post-login redirect and dashboard to finish loading


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
    deadline = time.time() + 30  # some dashboards are slow to populate the model name
    while time.time() < deadline:
        try:
            el = find_element_anywhere(driver, By.CSS_SELECTOR, "#tab0 span",
                                       timeout=8, label="model", require_visible=False)
            model = el.text.strip()
            if model and model.lower() != "device0":
                log(f"[{location} | {ip}] Model: {model}")
                return model
        except Exception:
            pass
        time.sleep(3)
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

        # Battery - try up to 6 times to give slow NICs a fair chance
        battery_rows, page_updated = [], ""
        for attempt in range(1, 7):
            log(f"[{location} | {ip}] Battery attempt {attempt} ...")
            try:
                driver.switch_to.default_content()
                bat = find_element_anywhere(driver, By.ID, "report163860",
                                            timeout=BATTERY_LINK_TIMEOUT,
                                            label="Battery link", require_visible=False)
                js_click(driver, bat)
                time.sleep(6)  # give the page a moment to start loading before we poll
                page_updated, battery_rows = scrape_battery_page(driver, timeout=TABLE_LOAD_TIMEOUT)
                if battery_rows:
                    break
            except TimeoutException:
                log(f"[{location} | {ip}] Battery attempt {attempt} timed out - waiting 15s before retry ...")
                driver.switch_to.default_content()
                time.sleep(15)

        if not battery_rows:
            raise TimeoutException("Battery table empty after all retry attempts")
        log(f"[{location} | {ip}] {len(battery_rows)} battery fields captured.")

        # Communications → Active Networking → MAC
        log(f"[{location} | {ip}] Clicking Communications tab ...")
        driver.switch_to.default_content()
        js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=60,
                                               label="Communications tab", require_visible=False))
        time.sleep(5)  # wait for the Communications section to fully load

        log(f"[{location} | {ip}] Expanding Support ...")
        js_click(driver, find_element_anywhere(driver, By.ID, "164190Plus", timeout=60,
                                               label="Support expand", require_visible=False))
        time.sleep(5)  # wait for the Support submenu to expand

        log(f"[{location} | {ip}] Clicking Active Networking ...")
        js_click(driver, find_element_anywhere(driver, By.ID, "report164330", timeout=60,
                                               label="Active Networking", require_visible=False))
        time.sleep(5)  # wait for the Active Networking page to load

        log(f"[{location} | {ip}] Reading Ethernet MAC ...")
        mac_el = find_element_anywhere(driver, By.ID, "val6156_0", timeout=60,
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
    _open_file(path)
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


def _wait_for_reboot_page(driver, location: str, ip: str, timeout: int = 480) -> bool:
    """After clicking Run Alternate, stay on the reboot-holding page and wait for the
    login form to appear — the Vertiv page transitions on its own, do NOT navigate away."""
    log(f"[{location} | {ip}] Staying on reboot page — waiting for login to appear (up to {timeout}s) ...")
    deadline = time.time() + timeout
    attempt = 0
    while time.time() < deadline:
        attempt += 1
        try:
            driver.switch_to.default_content()
            body = ""
            try:
                body = driver.find_element(By.TAG_NAME, "body").text.lower()
            except Exception:
                pass

            if driver.find_elements(By.ID, "username") and driver.find_elements(By.ID, "password"):
                log(f"[{location} | {ip}] Login page ready.")
                return True

            if "attempting to reconnect" in body or "web card has been rebooted" in body:
                log(f"[{location} | {ip}] Still rebooting (attempt {attempt}) ...")
            elif body:
                log(f"[{location} | {ip}] Waiting for login (attempt {attempt}) ...")
        except Exception as e:
            log(f"[{location} | {ip}] Waiting ... ({_short_error(e)})")

        time.sleep(10)

    log(f"[{location} | {ip}] Login page did not appear within {timeout}s.")
    return False


def _real_click(driver, el) -> None:
    """Real click with ActionChains fallback if element is not directly interactable."""
    try:
        el.click()
    except Exception:
        ActionChains(driver).move_to_element(el).click().perform()


def _click_run_alternate(driver, location: str, ip: str) -> bool:
    """Click Enable then Run Alternate and confirm the dialog. Returns True if clicked."""
    try:
        enable_btn = find_element_anywhere(driver, By.ID, "enableComms", timeout=20,
                                           label="Enable", require_visible=False)
        driver.execute_script("arguments[0].scrollIntoView(true);", enable_btn)
        time.sleep(0.5)
        _real_click(driver, enable_btn)
        log(f"[{location} | {ip}] Enable clicked — waiting for Run Alternate to activate ...")
        time.sleep(2)  # give the UI time to react before polling

        run_alt = None
        deadline_btn = time.time() + 30
        while time.time() < deadline_btn:
            try:
                btn = find_element_anywhere(driver, By.ID, "commBtn263", timeout=5,
                                            label="Run Alternate", require_visible=False)
                disabled = btn.get_attribute("disabled")
                # Only None means the attribute is absent = button is enabled
                if disabled is None:
                    run_alt = btn
                    break
                log(f"[{location} | {ip}] Run Alternate still disabled (attr={disabled!r}) — waiting ...")
            except Exception as e:
                log(f"[{location} | {ip}] Waiting for Run Alternate button ... ({_short_error(e)})")
            time.sleep(2)

        if not run_alt:
            log(f"[{location} | {ip}] Run Alternate button never became enabled.")
            return False

        driver.execute_script("arguments[0].scrollIntoView(true);", run_alt)
        time.sleep(0.5)
        log(f"[{location} | {ip}] Clicking Run Alternate ...")
        _real_click(driver, run_alt)

        # Native browser confirm dialog
        try:
            alert = WebDriverWait(driver, 15).until(EC.alert_is_present())
            alert.accept()
            log(f"[{location} | {ip}] Run Alternate confirmed — device will reboot.")
        except TimeoutException:
            # Fall back to HTML modal OK button if no native alert appeared
            log(f"[{location} | {ip}] No native alert — checking for HTML OK dialog ...")
            try:
                ok_btn = find_element_anywhere(
                    driver,
                    By.XPATH,
                    "//button[normalize-space(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']"
                    " | //input[@type='button' and normalize-space(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']",
                    timeout=5, label="OK button", require_visible=True)
                _real_click(driver, ok_btn)
                log(f"[{location} | {ip}] HTML OK dialog confirmed — device will reboot.")
            except Exception:
                log(f"[{location} | {ip}] Run Alternate clicked — no confirmation dialog detected.")
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

    # Use the known Firmware Update link ID; fall back to text search if missing
    try:
        fw_link = find_element_anywhere(driver, By.ID, "report164380", timeout=15,
                                        label="Firmware Update link", require_visible=False)
    except TimeoutException:
        fw_link = find_element_anywhere(
            driver,
            By.XPATH, "//a[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'firmware')]",
            timeout=20, label="Firmware Update link (fallback)", require_visible=False)
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

                    # Step 4: stay on reboot page → wait for login → sign back in → nav for retry
                    log(f"[{location} | {ip}] [RECOVERY 4/4] Waiting for device to reboot ...")
                    if not _wait_for_reboot_page(driver, location, ip, timeout=480):
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
                en = find_element_anywhere(driver, By.ID, "enableComms", timeout=20,
                                           label="Enable button", require_visible=False)
                driver.execute_script("arguments[0].scrollIntoView(true);", en)
                time.sleep(0.3)
                _real_click(driver, en)
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
                device_ready = False

                if used_recovery:
                    # Recovery upload succeeded — device rebooted to OLD firmware.
                    # Wait for it to come back, then Run Alternate to activate the new firmware.
                    log(f"[{location} | {ip}] Recovery upload complete — waiting for device, then activating via Run Alternate ...")
                    if _wait_for_device_online(driver, location, ip, timeout=600):
                        _login(driver, location, ip, username, password)
                        try:
                            _nav_to_firmware_page(driver, location, ip, username, password)
                            if _click_run_alternate(driver, location, ip):
                                log(f"[{location} | {ip}] Run Alternate clicked — staying on reboot page ...")
                                # Stay on the reboot-holding page; it transitions to login on its own
                                device_ready = _wait_for_reboot_page(driver, location, ip, timeout=600)
                            else:
                                log(f"[{location} | {ip}] Could not click Run Alternate for activation.")
                        except Exception as exc:
                            log(f"[{location} | {ip}] Activation step failed: {_short_error(exc)}")
                else:
                    # Normal upload — device rebooted after "Go Home"; navigate to login URL
                    log(f"[{location} | {ip}] Waiting for device to come back up after upgrade ...")
                    device_ready = _wait_for_device_online(driver, location, ip, timeout=600)

                if device_ready:
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
    finally:
        if driver:
            try: driver.quit()
            except Exception: pass
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
    _open_file(path)
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        log(f"[{r['location']} | {r['ip']}] {r['model']}  current={r['current_version']}  applied={r['upgrade_applied']}  status={r['upload_status'] or r['error'] or 'ok'}")
    log("Done.")


# ---------------------------------------------------------------------------
# Mode 3 — SNMPv3 Configuration
# ---------------------------------------------------------------------------

def _nav_to_snmpv3_page(driver, location: str, ip: str,
                         username: str = "", password: str = "") -> None:
    def _reauth_if_needed():
        if username and _is_auth_page(driver):
            log(f"[{location} | {ip}] Auth challenge — re-logging in ...")
            _login(driver, location, ip, username, password)

    log(f"[{location} | {ip}] Navigating to SNMPv3 configuration ...")
    _reauth_if_needed()

    try:
        js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=20,
                                               label="Communications tab", require_visible=False))
        time.sleep(2)
        _reauth_if_needed()
    except TimeoutException:
        log(f"[{location} | {ip}] Communications tab not found — proceeding ...")

    js_click(driver, find_element_anywhere(driver, By.ID, "report164220", timeout=20,
                                           label="Protocols", require_visible=False))
    time.sleep(2)
    _reauth_if_needed()

    js_click(driver, find_element_anywhere(driver, By.ID, "report164210", timeout=15,
                                           label="SNMP", require_visible=False))
    time.sleep(2)

    js_click(driver, find_element_anywhere(driver, By.ID, "report16408164210", timeout=15,
                                           label="SNMPv3 User list", require_visible=False))
    time.sleep(2)

    user_link = find_element_anywhere(driver, By.CSS_SELECTOR, 'a[title="SNMPv3 User [1]"]',
                                      timeout=15, label="SNMPv3 User [1]", require_visible=False)
    js_click(driver, user_link)
    time.sleep(2)


def _configure_snmpv3(driver, location: str, ip: str, cfg: dict) -> None:
    log(f"[{location} | {ip}] Clicking Edit ...")
    edit_btn = find_element_anywhere(driver, By.ID, "editButton", timeout=20,
                                     label="Edit button", require_visible=False)
    js_click(driver, edit_btn)
    time.sleep(2)

    log(f"[{location} | {ip}] Enabling SNMPv3 user ...")
    chk = find_element_anywhere(driver, By.ID, "chkbx7385", timeout=15,
                                label="SNMPv3 Enable checkbox", require_visible=False)
    if not chk.is_selected():
        js_click(driver, chk)
        time.sleep(0.5)

    log(f"[{location} | {ip}] Setting SNMPv3 username ...")
    un_field = find_element_anywhere(driver, By.ID, "str7386", timeout=10,
                                     label="SNMPv3 username", require_visible=False)
    un_field.clear()
    un_field.send_keys(cfg["snmp_username"])

    access_map = {"Read Only": "0", "Read/Write": "1", "Traps Only": "2"}
    SeleniumSelect(find_element_anywhere(driver, By.ID, "enum7387", timeout=10,
                                         label="Access Type", require_visible=False)
                   ).select_by_value(access_map.get(cfg["access_type"], "0"))

    auth_map = {"None": "0", "MD5": "1", "SHA": "2"}
    SeleniumSelect(find_element_anywhere(driver, By.ID, "enum7388", timeout=10,
                                         label="Auth Protocol", require_visible=False)
                   ).select_by_value(auth_map.get(cfg["auth_protocol"], "0"))

    if cfg.get("auth_secret"):
        auth_sec = find_element_anywhere(driver, By.ID, "str7389", timeout=10,
                                         label="Auth Secret", require_visible=False)
        auth_sec.clear()
        auth_sec.send_keys(cfg["auth_secret"])

    priv_map = {"None": "0", "DES": "1", "AES": "2"}
    SeleniumSelect(find_element_anywhere(driver, By.ID, "enum7390", timeout=10,
                                         label="Privacy Protocol", require_visible=False)
                   ).select_by_value(priv_map.get(cfg["privacy_protocol"], "0"))

    if cfg.get("privacy_secret"):
        priv_sec = find_element_anywhere(driver, By.ID, "str7391", timeout=10,
                                          label="Privacy Secret", require_visible=False)
        priv_sec.clear()
        priv_sec.send_keys(cfg["privacy_secret"])

    if cfg.get("trap_targets"):
        trap_tgt = find_element_anywhere(driver, By.ID, "str7392", timeout=10,
                                          label="Trap Targets", require_visible=False)
        trap_tgt.clear()
        trap_tgt.send_keys(cfg["trap_targets"])

    if cfg.get("trap_port"):
        trap_port = find_element_anywhere(driver, By.ID, "num7393", timeout=10,
                                           label="Trap Port", require_visible=False)
        trap_port.clear()
        trap_port.send_keys(cfg["trap_port"])

    log(f"[{location} | {ip}] Saving SNMPv3 settings ...")
    save_btn = find_element_anywhere(driver, By.ID, "submitButton", timeout=10,
                                     label="Save button", require_visible=False)
    js_click(driver, save_btn)
    time.sleep(3)
    log(f"[{location} | {ip}] SNMPv3 settings saved.")

    # Disable SNMPv1/v2c
    log(f"[{location} | {ip}] Navigating back to Protocols → SNMP to disable SNMPv1/v2c ...")
    js_click(driver, find_element_anywhere(driver, By.ID, "report164220", timeout=20,
                                           label="Protocols", require_visible=False))
    time.sleep(2)
    js_click(driver, find_element_anywhere(driver, By.ID, "report164210", timeout=15,
                                           label="SNMP", require_visible=False))
    time.sleep(2)

    log(f"[{location} | {ip}] Clicking Edit for SNMP settings ...")
    js_click(driver, find_element_anywhere(driver, By.ID, "editButton", timeout=15,
                                           label="SNMP Edit button", require_visible=False))
    time.sleep(2)

    log(f"[{location} | {ip}] Unchecking SNMPv1/v2c Enable ...")
    v1v2_chk = find_element_anywhere(driver, By.ID, "chkbx7400", timeout=15,
                                     label="SNMPv1/v2c Enable checkbox", require_visible=False)
    if v1v2_chk.is_selected():
        js_click(driver, v1v2_chk)
        time.sleep(0.5)
        log(f"[{location} | {ip}] SNMPv1/v2c unchecked.")
    else:
        log(f"[{location} | {ip}] SNMPv1/v2c was already disabled.")

    log(f"[{location} | {ip}] Saving SNMP settings ...")
    js_click(driver, find_element_anywhere(driver, By.ID, "submitButton", timeout=10,
                                           label="SNMP Save button", require_visible=False))
    time.sleep(3)
    log(f"[{location} | {ip}] SNMPv1/v2c disabled.")


def process_snmpv3_ip(location: str, ip: str, username: str, password: str, cfg: dict) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", status="unknown",
                  restart_required=False, restarted=False, scraped_at="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)
        result["model"] = _read_model(driver, location, ip)

        _nav_to_snmpv3_page(driver, location, ip, username, password)
        driver._au_nav_fn = lambda: _nav_to_snmpv3_page(driver, location, ip, username, password)

        _configure_snmpv3(driver, location, ip, cfg)
        result.update(status="success", scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        # After configuring SNMPv3, check whether the NIC flagged a restart requirement
        log(f"[{location} | {ip}] Checking NIC events after SNMPv3 config ...")
        try:
            driver._au_nav_fn = None
            try:
                js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=20,
                                                       label="Communications tab", require_visible=False))
                time.sleep(2)
            except TimeoutException:
                pass
            js_click(driver, find_element_anywhere(driver, By.ID, "report164180", timeout=20,
                                                   label="Status", require_visible=False))
            time.sleep(3)
            _, nic_rows = scrape_battery_page(driver, timeout=60)
            restart_required = any(
                r["label"] == "System Restart Required" and r["value"].strip().lower() == "active"
                for r in nic_rows
            )
            result["restart_required"] = restart_required
            if restart_required:
                log(f"[{location} | {ip}] System Restart Required is ACTIVE — restarting NIC ...")
                js_click(driver, find_element_anywhere(driver, By.ID, "report164190", timeout=20,
                                                       label="Support", require_visible=False))
                time.sleep(2)
                restarted = _do_nic_restart(driver, location, ip)
                result["restarted"] = restarted
                if restarted:
                    log(f"[{location} | {ip}] NIC restart initiated successfully.")
                else:
                    log(f"[{location} | {ip}] NIC restart could not be completed.")
            else:
                log(f"[{location} | {ip}] No NIC restart required.")
        except Exception as exc:
            log(f"[{location} | {ip}] NIC event check after SNMPv3 failed: {_short_error(exc)}")

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


def _build_snmpv3_csv(results: list[dict]) -> str:
    fixed = ["Location", "IP", "Model", "Status", "Restart Required", "NIC Restarted", "Scraped At", "Error"]
    path = os.path.join(SCRIPT_DIR, f"snmpv3_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed, extrasaction="ignore")
        w.writeheader()
        for r in results:
            if r is None:
                continue
            w.writerow({"Location": r["location"], "IP": r["ip"], "Model": r["model"],
                        "Status": r["status"],
                        "Restart Required": "YES" if r.get("restart_required") else "no",
                        "NIC Restarted":    "YES" if r.get("restarted") else "no",
                        "Scraped At": r["scraped_at"], "Error": r["error"]})
    return path


def run_snmpv3_config(targets, username, password, cfg, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_snmpv3_ip(loc, ip, username, password, cfg)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
                if r["error"]:
                    s += f"  — {r['error']}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_snmpv3_csv(results)
    log(f"\nCSV saved: {path}")
    _open_file(path)
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
        if r["error"]:
            s += f"  — {r['error']}"
        log(s)
    log("Done.")


# ---------------------------------------------------------------------------
# Mode 4 — Silence Alarm
# ---------------------------------------------------------------------------

def process_silence_alarm_ip(location: str, ip: str, username: str, password: str) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", status="unknown",
                  scraped_at="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)
        result["model"] = _read_model(driver, location, ip)

        # Wait for the sidebar to load (either link appearing means it's ready),
        # then prefer System Configuration if it's there -- some models have both,
        # and System Configuration is the one with the Silence Alarm command.
        log(f"[{location} | {ip}] Waiting for sidebar nav links to load ...")
        find_element_anywhere(
            driver,
            By.XPATH, "//*[@id='report163910' or @id='report263940']",
            timeout=60, label="System nav link", require_visible=False)
        # Now that the sidebar is loaded, check whether System Configuration is present
        try:
            nav_link = driver.find_element(By.ID, "report263940")
            log(f"[{location} | {ip}] System Configuration found — using that.")
        except Exception:
            nav_link = driver.find_element(By.ID, "report163910")
            log(f"[{location} | {ip}] System Configuration not present — using System.")
        nav_label = nav_link.get_attribute("id")
        log(f"[{location} | {ip}] Clicking {nav_label} ...")
        driver.execute_script("arguments[0].scrollIntoView(true);", nav_link)
        time.sleep(0.5)
        _real_click(driver, nav_link)
        log(f"[{location} | {ip}] Nav clicked — waiting for page ...")
        time.sleep(4)

        log(f"[{location} | {ip}] Clicking Enable ...")
        en = find_element_anywhere(driver, By.ID, "enableComms", timeout=20,
                                   label="Enable", require_visible=False)
        driver.execute_script("arguments[0].scrollIntoView(true);", en)
        time.sleep(0.5)
        _real_click(driver, en)
        log(f"[{location} | {ip}] Enable clicked — polling for Silence Alarm button ...")
        time.sleep(4)

        silence_btn = None
        deadline_btn = time.time() + 45
        while time.time() < deadline_btn:
            try:
                btn = find_element_anywhere(driver, By.ID, "commBtn6257", timeout=5,
                                            label="Silence Alarm", require_visible=False)
                disabled = btn.get_attribute("disabled")
                if disabled is None:
                    silence_btn = btn
                    log(f"[{location} | {ip}] Silence Alarm button is active.")
                    break
                log(f"[{location} | {ip}] Silence Alarm button disabled (attr={disabled!r}) — waiting ...")
            except Exception as e:
                log(f"[{location} | {ip}] Silence Alarm button not found yet: {_short_error(e)}")
            time.sleep(3)

        if not silence_btn:
            raise TimeoutException("Silence Alarm button never became enabled")

        driver.execute_script("arguments[0].scrollIntoView(true);", silence_btn)
        time.sleep(1)
        log(f"[{location} | {ip}] Clicking Silence Alarm ...")
        _real_click(driver, silence_btn)
        log(f"[{location} | {ip}] Silence Alarm clicked — waiting for confirmation dialog ...")
        time.sleep(2)

        try:
            alert = WebDriverWait(driver, 20).until(EC.alert_is_present())
            log(f"[{location} | {ip}] Dialog: {alert.text!r} — clicking OK ...")
            alert.accept()
            log(f"[{location} | {ip}] Silence Alarm confirmed.")
        except TimeoutException:
            log(f"[{location} | {ip}] No native alert after 20s — checking for HTML OK button ...")
            ok_btn = find_element_anywhere(
                driver,
                By.XPATH,
                "//button[normalize-space(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']"
                " | //input[@type='button' and normalize-space(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']",
                timeout=10, label="OK button", require_visible=True)
            _real_click(driver, ok_btn)
            log(f"[{location} | {ip}] HTML OK confirmed.")

        time.sleep(5)
        result.update(status="success", scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] Done.")

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


def _build_silence_csv(results: list[dict]) -> str:
    fixed = ["Location", "IP", "Model", "Status", "Scraped At", "Error"]
    path = os.path.join(SCRIPT_DIR, f"silence_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed, extrasaction="ignore")
        w.writeheader()
        for r in results:
            if r is None:
                continue
            w.writerow({"Location": r["location"], "IP": r["ip"], "Model": r["model"],
                        "Status": r["status"], "Scraped At": r["scraped_at"], "Error": r["error"]})
    return path


def run_silence_alarm(targets, username, password, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_silence_alarm_ip(loc, ip, username, password)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
                if r["error"]:
                    s += f"  — {r['error']}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_silence_csv(results)
    log(f"\nCSV saved: {path}")
    _open_file(path)
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
        if r["error"]:
            s += f"  — {r['error']}"
        log(s)
    log("Done.")


# ---------------------------------------------------------------------------
# Shared NIC restart helper (used by Restart NIC tab and post-SNMPv3 flow)
# ---------------------------------------------------------------------------

def _do_nic_restart(driver, location: str, ip: str) -> bool:
    """Click Enable then Restart on the Support page and confirm the dialog.
    Returns True if the restart was successfully initiated."""
    try:
        en = find_element_anywhere(driver, By.ID, "enableComms", timeout=20,
                                   label="Enable", require_visible=False)
        driver.execute_script("arguments[0].scrollIntoView(true);", en)
        time.sleep(0.5)
        _real_click(driver, en)
        log(f"[{location} | {ip}] Enable clicked — waiting for Restart button ...")
        time.sleep(3)

        restart_btn = None
        deadline = time.time() + 30
        while time.time() < deadline:
            try:
                btn = find_element_anywhere(driver, By.ID, "commBtn139", timeout=5,
                                            label="Restart", require_visible=False)
                if btn.get_attribute("disabled") is None:
                    restart_btn = btn
                    break
                log(f"[{location} | {ip}] Restart button still disabled — waiting ...")
            except Exception as e:
                log(f"[{location} | {ip}] Restart button not found yet: {_short_error(e)}")
            time.sleep(2)

        if not restart_btn:
            log(f"[{location} | {ip}] Restart button never became enabled.")
            return False

        driver.execute_script("arguments[0].scrollIntoView(true);", restart_btn)
        time.sleep(0.5)
        log(f"[{location} | {ip}] Clicking Restart ...")
        _real_click(driver, restart_btn)

        try:
            alert = WebDriverWait(driver, 15).until(EC.alert_is_present())
            log(f"[{location} | {ip}] Dialog: {alert.text!r} — clicking OK ...")
            alert.accept()
            log(f"[{location} | {ip}] NIC restart confirmed.")
        except TimeoutException:
            try:
                ok_btn = find_element_anywhere(
                    driver,
                    By.XPATH,
                    "//button[normalize-space(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']"
                    " | //input[@type='button' and normalize-space(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']",
                    timeout=5, label="OK button", require_visible=True)
                _real_click(driver, ok_btn)
                log(f"[{location} | {ip}] NIC restart confirmed via HTML dialog.")
            except Exception:
                log(f"[{location} | {ip}] Restart clicked — no confirmation dialog detected.")
        return True
    except Exception as exc:
        log(f"[{location} | {ip}] NIC restart failed: {_short_error(exc)}")
        return False


# ---------------------------------------------------------------------------
# Mode 5 — Turn Output ON
# ---------------------------------------------------------------------------

def process_output_ip(location: str, ip: str, username: str, password: str) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", status="unknown",
                  scraped_at="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)
        result["model"] = _read_model(driver, location, ip)

        log(f"[{location} | {ip}] Clicking Output ...")
        out_link = find_element_anywhere(driver, By.ID, "report163870", timeout=60,
                                         label="Output", require_visible=False)
        driver.execute_script("arguments[0].scrollIntoView(true);", out_link)
        time.sleep(0.5)
        _real_click(driver, out_link)
        log(f"[{location} | {ip}] Output clicked — waiting for page ...")
        time.sleep(4)

        log(f"[{location} | {ip}] Clicking Enable ...")
        en = find_element_anywhere(driver, By.ID, "enableOutput", timeout=20,
                                   label="Enable Output", require_visible=False)
        driver.execute_script("arguments[0].scrollIntoView(true);", en)
        time.sleep(0.5)
        _real_click(driver, en)
        log(f"[{location} | {ip}] Enable clicked — polling for Turn Output ON button ...")
        time.sleep(4)

        output_btn = None
        deadline_btn = time.time() + 45
        while time.time() < deadline_btn:
            try:
                btn = find_element_anywhere(driver, By.ID, "commBtn5816", timeout=5,
                                            label="Turn Output ON", require_visible=False)
                if btn.get_attribute("disabled") is None:
                    output_btn = btn
                    log(f"[{location} | {ip}] Turn Output ON button is active.")
                    break
                log(f"[{location} | {ip}] Turn Output ON button still disabled — waiting ...")
            except Exception as e:
                log(f"[{location} | {ip}] Turn Output ON button not found yet: {_short_error(e)}")
            time.sleep(3)

        if not output_btn:
            raise TimeoutException("Turn Output ON button never became enabled")

        driver.execute_script("arguments[0].scrollIntoView(true);", output_btn)
        time.sleep(1)
        log(f"[{location} | {ip}] Clicking Turn Output ON ...")
        _real_click(driver, output_btn)
        time.sleep(2)

        try:
            alert = WebDriverWait(driver, 20).until(EC.alert_is_present())
            log(f"[{location} | {ip}] Dialog: {alert.text!r} — clicking OK ...")
            alert.accept()
            log(f"[{location} | {ip}] Turn Output ON confirmed.")
        except TimeoutException:
            log(f"[{location} | {ip}] No native alert — checking for HTML OK button ...")
            ok_btn = find_element_anywhere(
                driver,
                By.XPATH,
                "//button[normalize-space(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']"
                " | //input[@type='button' and normalize-space(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))='OK']",
                timeout=10, label="OK button", require_visible=True)
            _real_click(driver, ok_btn)
            log(f"[{location} | {ip}] HTML OK confirmed.")

        time.sleep(5)
        result.update(status="success", scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        log(f"[{location} | {ip}] Done.")

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


def _build_output_csv(results: list[dict]) -> str:
    fixed = ["Location", "IP", "Model", "Status", "Scraped At", "Error"]
    path = os.path.join(SCRIPT_DIR, f"output_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed, extrasaction="ignore")
        w.writeheader()
        for r in results:
            if r is None:
                continue
            w.writerow({"Location": r["location"], "IP": r["ip"], "Model": r["model"],
                        "Status": r["status"], "Scraped At": r["scraped_at"], "Error": r["error"]})
    return path


def run_output(targets, username, password, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_output_ip(loc, ip, username, password)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
                if r["error"]:
                    s += f"  - {r['error']}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_output_csv(results)
    log(f"\nCSV saved: {path}")
    _open_file(path)
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
        if r["error"]:
            s += f"  - {r['error']}"
        log(s)
    log("Done.")


# ---------------------------------------------------------------------------
# Mode 6 — Restart NIC
# ---------------------------------------------------------------------------

def process_restart_nic_ip(location: str, ip: str, username: str, password: str) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", status="unknown",
                  scraped_at="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)
        result["model"] = _read_model(driver, location, ip)

        log(f"[{location} | {ip}] Navigating to Communications > Support ...")
        try:
            js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=20,
                                                   label="Communications tab", require_visible=False))
            time.sleep(2)
        except TimeoutException:
            log(f"[{location} | {ip}] Communications tab not found — proceeding ...")

        js_click(driver, find_element_anywhere(driver, By.ID, "report164190", timeout=20,
                                               label="Support", require_visible=False))
        time.sleep(3)

        restarted = _do_nic_restart(driver, location, ip)
        if restarted:
            result.update(status="success", scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            log(f"[{location} | {ip}] NIC restart initiated.")
        else:
            result.update(status="error", error="Restart button never became enabled",
                          scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

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


def _build_restart_nic_csv(results: list[dict]) -> str:
    fixed = ["Location", "IP", "Model", "Status", "Scraped At", "Error"]
    path = os.path.join(SCRIPT_DIR, f"restart_nic_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed, extrasaction="ignore")
        w.writeheader()
        for r in results:
            if r is None:
                continue
            w.writerow({"Location": r["location"], "IP": r["ip"], "Model": r["model"],
                        "Status": r["status"], "Scraped At": r["scraped_at"], "Error": r["error"]})
    return path


def run_restart_nic(targets, username, password, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_restart_nic_ip(loc, ip, username, password)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
                if r["error"]:
                    s += f"  - {r['error']}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_restart_nic_csv(results)
    log(f"\nCSV saved: {path}")
    _open_file(path)
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
        if r["error"]:
            s += f"  - {r['error']}"
        log(s)
    log("Done.")


# ---------------------------------------------------------------------------
# Mode 7 — Check NIC Events
# ---------------------------------------------------------------------------

def process_nic_events_ip(location: str, ip: str, username: str, password: str) -> dict:
    ip = _clean_ip(ip)
    result = dict(location=location, ip=ip, model="", status="unknown",
                  restart_required=False, event_rows=[], scraped_at="", error="")
    driver = None
    try:
        driver = _make_driver()
        _login(driver, location, ip, username, password)
        result["model"] = _read_model(driver, location, ip)

        log(f"[{location} | {ip}] Navigating to Communications > Status ...")
        try:
            js_click(driver, find_element_anywhere(driver, By.ID, "tab4", timeout=20,
                                                   label="Communications tab", require_visible=False))
            time.sleep(2)
        except TimeoutException:
            log(f"[{location} | {ip}] Communications tab not found — proceeding ...")

        js_click(driver, find_element_anywhere(driver, By.ID, "report164180", timeout=20,
                                               label="Status", require_visible=False))
        time.sleep(3)

        log(f"[{location} | {ip}] Scraping NIC status and events ...")
        _, rows = scrape_battery_page(driver, timeout=60)

        restart_required = any(
            r["label"] == "System Restart Required" and r["value"].strip().lower() == "active"
            for r in rows
        )
        result.update(status="success", restart_required=restart_required,
                      event_rows=rows, scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        if restart_required:
            log(f"[{location} | {ip}] System Restart Required is ACTIVE.")
        else:
            log(f"[{location} | {ip}] No restart required.")
        log(f"[{location} | {ip}] {len(rows)} NIC status fields captured.")

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


def _build_nic_events_csv(results: list[dict]) -> str:
    lu: dict[str, str] = {}
    for r in results:
        for row in r.get("event_rows", []):
            if row["label"] not in lu:
                lu[row["label"]] = row["unit"]
    event_cols = [(f"{lbl} ({unit})" if unit else lbl, lbl) for lbl, unit in lu.items()]
    fixed = ["Location", "IP", "Model", "Restart Required", "Status", "Scraped At", "Error"]
    path = os.path.join(SCRIPT_DIR, f"nic_events_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fixed + [c for c, _ in event_cols], extrasaction="ignore")
        w.writeheader()
        for r in results:
            if r is None:
                continue
            row = {"Location": r["location"], "IP": r["ip"], "Model": r["model"],
                   "Restart Required": "YES" if r.get("restart_required") else "no",
                   "Status": r["status"], "Scraped At": r["scraped_at"], "Error": r["error"]}
            vm = {er["label"]: er["value"] for er in r.get("event_rows", [])}
            for col, lbl in event_cols:
                row[col] = vm.get(lbl, "")
            w.writerow(row)
    return path


def run_nic_events(targets, username, password, max_parallel=3):
    results = [None] * len(targets)
    def _run(idx, loc, ip):
        time.sleep(idx * 1.5)
        return idx, process_nic_events_ip(loc, ip, username, password)
    with ThreadPoolExecutor(max_workers=max_parallel) as pool:
        futures = {pool.submit(_run, i, loc, ip): i for i,(loc,ip) in enumerate(targets)}
        for fut in as_completed(futures):
            try:
                idx, r = fut.result()
                results[idx] = r
                s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
                if r.get("restart_required"):
                    s += "  - RESTART REQUIRED"
                if r["error"]:
                    s += f"  - {r['error']}"
                log(f"Finished: {s}")
            except Exception as exc:
                log(f"Worker error: {_short_error(exc)}")
    path = _build_nic_events_csv(results)
    log(f"\nCSV saved: {path}")
    _open_file(path)
    log("\n=== SUMMARY ===")
    for r in (r for r in results if r):
        s = f"[{r['location']} | {r['ip']}] {r['status'].upper()}"
        if r.get("restart_required"):
            s += "  - RESTART REQUIRED"
        if r["error"]:
            s += f"  - {r['error']}"
        log(s)
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
        creds = ttk.LabelFrame(self, text="Credentials")
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
        cfg_btns = ttk.Frame(creds)
        cfg_btns.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 6))
        ttk.Button(cfg_btns, text="Save Config", command=self._save_config).pack(side="left", padx=(0, 6))
        ttk.Button(cfg_btns, text="Load Config", command=self._load_config).pack(side="left")
        ttk.Label(cfg_btns, text="Saves/loads all credentials and SNMPv3 settings to your Desktop.",
                  foreground="gray").pack(side="left", padx=(10, 0))

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

        # Tab 3 — SNMPv3
        snmp_tab = ttk.Frame(self.notebook)
        self.notebook.add(snmp_tab, text="SNMPv3 Config")
        self._build_snmpv3_tab(snmp_tab, pad)

        # Tab 4 — Silence Alarm
        silence_tab = ttk.Frame(self.notebook)
        self.notebook.add(silence_tab, text="Silence Alarm")
        ttk.Label(silence_tab,
                  text="Navigates to System or System Configuration, enables commands, and silences the alarm.",
                  foreground="gray").pack(padx=8, pady=6)

        # Tab 5 — Turn Output ON
        output_tab = ttk.Frame(self.notebook)
        self.notebook.add(output_tab, text="Output")
        ttk.Label(output_tab,
                  text="Turns UPS output back on after a power outage. Clicks Output, Enable, Turn Output ON.",
                  foreground="gray").pack(padx=8, pady=6)

        # Tab 6 — Restart NIC
        restart_tab = ttk.Frame(self.notebook)
        self.notebook.add(restart_tab, text="Restart NIC")
        ttk.Label(restart_tab,
                  text="Navigates to Communications > Support, enables commands, and restarts the NIC card.",
                  foreground="gray").pack(padx=8, pady=6)

        # Tab 7 — Check NIC Events
        nic_events_tab = ttk.Frame(self.notebook)
        self.notebook.add(nic_events_tab, text="NIC Events")
        ttk.Label(nic_events_tab,
                  text="Navigates to Communications > Status and logs all NIC events. Flags System Restart Required in the CSV.",
                  foreground="gray").pack(padx=8, pady=6)

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

    def _build_snmpv3_tab(self, parent, pad):
        cfg_frame = ttk.LabelFrame(parent, text="SNMPv3 User 1 Settings  (never stored)")
        cfg_frame.grid(row=0, column=0, sticky="ew", **pad)
        cfg_frame.columnconfigure(1, weight=1)

        ttk.Label(cfg_frame, text="SNMPv3 Username:").grid(row=0, column=0, sticky="e", **pad)
        self.snmpv3_user_var = tk.StringVar()
        un_e = ttk.Entry(cfg_frame, textvariable=self.snmpv3_user_var, width=30)
        un_e.grid(row=0, column=1, sticky="w", **pad)
        self._bind_paste(un_e)

        ttk.Label(cfg_frame, text="Access Type:").grid(row=1, column=0, sticky="e", **pad)
        self.snmpv3_access_var = tk.StringVar(value="Read Only")
        ttk.Combobox(cfg_frame, textvariable=self.snmpv3_access_var, width=20,
                     state="readonly", values=["Read Only", "Read/Write", "Traps Only"]
                     ).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(cfg_frame, text="Auth Protocol:").grid(row=2, column=0, sticky="e", **pad)
        self.snmpv3_auth_proto_var = tk.StringVar(value="None")
        ttk.Combobox(cfg_frame, textvariable=self.snmpv3_auth_proto_var, width=20,
                     state="readonly", values=["None", "MD5", "SHA"]
                     ).grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(cfg_frame, text="Auth Secret:").grid(row=3, column=0, sticky="e", **pad)
        self.snmpv3_auth_secret_var = tk.StringVar()
        auth_e = ttk.Entry(cfg_frame, textvariable=self.snmpv3_auth_secret_var, show="*", width=30)
        auth_e.grid(row=3, column=1, sticky="w", **pad)
        self._bind_paste(auth_e)

        ttk.Label(cfg_frame, text="Privacy Protocol:").grid(row=4, column=0, sticky="e", **pad)
        self.snmpv3_priv_proto_var = tk.StringVar(value="None")
        ttk.Combobox(cfg_frame, textvariable=self.snmpv3_priv_proto_var, width=20,
                     state="readonly", values=["None", "DES", "AES"]
                     ).grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(cfg_frame, text="Privacy Secret:").grid(row=5, column=0, sticky="e", **pad)
        self.snmpv3_priv_secret_var = tk.StringVar()
        priv_e = ttk.Entry(cfg_frame, textvariable=self.snmpv3_priv_secret_var, show="*", width=30)
        priv_e.grid(row=5, column=1, sticky="w", **pad)
        self._bind_paste(priv_e)

        ttk.Label(cfg_frame, text="Trap Targets:").grid(row=6, column=0, sticky="e", **pad)
        self.snmpv3_trap_targets_var = tk.StringVar()
        trap_e = ttk.Entry(cfg_frame, textvariable=self.snmpv3_trap_targets_var, width=30)
        trap_e.grid(row=6, column=1, sticky="w", **pad)
        self._bind_paste(trap_e)

        ttk.Label(cfg_frame, text="Trap Port:").grid(row=7, column=0, sticky="e", **pad)
        self.snmpv3_trap_port_var = tk.StringVar()
        port_e = ttk.Entry(cfg_frame, textvariable=self.snmpv3_trap_port_var, width=10)
        port_e.grid(row=7, column=1, sticky="w", **pad)
        self._bind_paste(port_e)

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

    # ── Config save / load ───────────────────────────────────────────────────

    def _save_config(self):
        path = filedialog.asksaveasfilename(
            title="Save config",
            initialfile="GXTManager_config.json",
            defaultextension=".json",
            filetypes=[("JSON config", "*.json"), ("All files", "*")],
        )
        if not path:
            return
        cfg = {
            "username":               self.username_var.get(),
            "password":               self.password_var.get(),
            "snmpv3_username":        self.snmpv3_user_var.get(),
            "snmpv3_access_type":     self.snmpv3_access_var.get(),
            "snmpv3_auth_protocol":   self.snmpv3_auth_proto_var.get(),
            "snmpv3_auth_secret":     self.snmpv3_auth_secret_var.get(),
            "snmpv3_privacy_protocol": self.snmpv3_priv_proto_var.get(),
            "snmpv3_privacy_secret":  self.snmpv3_priv_secret_var.get(),
            "snmpv3_trap_targets":    self.snmpv3_trap_targets_var.get(),
            "snmpv3_trap_port":       self.snmpv3_trap_port_var.get(),
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    def _load_config(self):
        path = filedialog.askopenfilename(
            title="Load config",
            defaultextension=".json",
            filetypes=[("JSON config", "*.json"), ("All files", "*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.username_var.set(cfg.get("username", ""))
            self.password_var.set(cfg.get("password", ""))
            self.snmpv3_user_var.set(cfg.get("snmpv3_username", ""))
            self.snmpv3_access_var.set(cfg.get("snmpv3_access_type", "Read Only"))
            self.snmpv3_auth_proto_var.set(cfg.get("snmpv3_auth_protocol", "None"))
            self.snmpv3_auth_secret_var.set(cfg.get("snmpv3_auth_secret", ""))
            self.snmpv3_priv_proto_var.set(cfg.get("snmpv3_privacy_protocol", "None"))
            self.snmpv3_priv_secret_var.set(cfg.get("snmpv3_privacy_secret", ""))
            self.snmpv3_trap_targets_var.set(cfg.get("snmpv3_trap_targets", ""))
            self.snmpv3_trap_port_var.set(cfg.get("snmpv3_trap_port", ""))
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc))

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
        elif mode_tab == 2:  # SNMPv3
            snmpv3_username = self.snmpv3_user_var.get().strip()
            if not snmpv3_username:
                messagebox.showwarning("Missing input", "Please enter an SNMPv3 username."); return
            cfg = {
                "snmp_username":    snmpv3_username,
                "access_type":      self.snmpv3_access_var.get(),
                "auth_protocol":    self.snmpv3_auth_proto_var.get(),
                "auth_secret":      self.snmpv3_auth_secret_var.get(),
                "privacy_protocol": self.snmpv3_priv_proto_var.get(),
                "privacy_secret":   self.snmpv3_priv_secret_var.get(),
                "trap_targets":     self.snmpv3_trap_targets_var.get().strip(),
                "trap_port":        self.snmpv3_trap_port_var.get().strip(),
            }
            self.start_btn.config(state="disabled")
            self.status_var.set(f"SNMPv3 Config — {len(targets)} device(s) ...")
            log(f"Starting SNMPv3 configuration on {len(targets)} device(s) ...")
            def worker():
                run_snmpv3_config(targets, username, password, cfg, parallel)
                LOG_QUEUE.put("__DONE__")
        elif mode_tab == 3:  # Silence Alarm
            self.start_btn.config(state="disabled")
            self.status_var.set(f"Silence Alarm — {len(targets)} device(s) ...")
            log(f"Starting Silence Alarm on {len(targets)} device(s) ...")
            def worker():
                run_silence_alarm(targets, username, password, parallel)
                LOG_QUEUE.put("__DONE__")
        elif mode_tab == 4:  # Output
            self.start_btn.config(state="disabled")
            self.status_var.set(f"Turn Output ON — {len(targets)} device(s) ...")
            log(f"Starting Turn Output ON on {len(targets)} device(s) ...")
            def worker():
                run_output(targets, username, password, parallel)
                LOG_QUEUE.put("__DONE__")
        elif mode_tab == 5:  # Restart NIC
            self.start_btn.config(state="disabled")
            self.status_var.set(f"Restart NIC — {len(targets)} device(s) ...")
            log(f"Starting NIC restart on {len(targets)} device(s) ...")
            def worker():
                run_restart_nic(targets, username, password, parallel)
                LOG_QUEUE.put("__DONE__")
        elif mode_tab == 6:  # NIC Events
            self.start_btn.config(state="disabled")
            self.status_var.set(f"NIC Events — {len(targets)} device(s) ...")
            log(f"Starting NIC event check on {len(targets)} device(s) ...")
            def worker():
                run_nic_events(targets, username, password, parallel)
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
