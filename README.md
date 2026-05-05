# GXTManager

A macOS GUI tool for managing Vertiv GXT-4 and GXT-5 UPS units over the network. Run battery health reports and push comm card firmware upgrades across your whole UPS fleet — no command line, no Vertiv cloud subscription required.

> **Note on platform:** This script is built and tested on macOS. It can be adapted to run on Windows with some Python knowledge or LLM assistance (primarily swapping out the launcher script and geckodriver setup).

---

## No cloud management system needed

Vertiv offers cloud-based NIC management platforms, but you don't need to buy into any of that to use this tool. All you need is:

1. Each UPS device's **IP address** on your network
2. A **label** for it — just a name you give the device so results are easy to read (e.g. the room number, closet name, or whatever makes sense to you)
3. The **web interface credentials** for the devices (the username and password you'd use to log in through a browser)

The script logs into each device's built-in web interface directly, the same way you would manually in a browser.

---

## What it does

**Battery Report**
- Logs into each device, pulls UPS Battery Status, Battery Test Result, Battery Cabinet Type, and Ethernet MAC
- Exports everything to a timestamped CSV you can open in Excel

**Firmware Upgrade**
- Checks the current comm card firmware version on each device
- Uploads a new `.fl` firmware file to devices that need it
- Handles errors automatically — if a device returns a 503 or drops the session mid-upload, the tool recovers and retries without any manual intervention
- Confirms the installed version after each upgrade
- Exports a per-device result CSV

---

## Requirements

- macOS (tested on macOS 14+)
- Python 3.9+
- Firefox

Install Python dependencies:

```bash
pip3 install -r requirements.txt
```

`geckodriver` is downloaded and signed automatically on first run.

---

## How to run

Double-click **Run Vertiv Scraper.command** in Finder. If macOS blocks it, right-click → Open.

Or from Terminal:

```bash
python3 vertiv_battery_scraper.py
```

---

## How to use it

### Step 1 — Enter credentials

Type your UPS web interface username and password into the Credentials fields at the top. These are never saved to disk.

### Step 2 — Paste your device list

The target box accepts two columns copied straight from a spreadsheet — **Location** (your label for the UPS) and **IP Address**, separated by a tab:

```
BLDG-100-CLOSET-A    10.0.1.15
SERVER-ROOM-UPS      10.0.1.22
4TH-FLOOR-SW         10.0.1.30
```

Copy those two columns from Excel or any spreadsheet and paste directly into the box. You can also paste just a plain list of IPs (one per line) if you don't need location labels.

### Step 3 — Choose a tab and run

**Battery tab** → click **Start** to pull battery status from all devices. Results save to `battery_report_<timestamp>.csv` next to the script.

**Firmware tab** → select your `.fl` firmware file(s), choose an upgrade mode, then click **Start**. Results save to `firmware_report_<timestamp>.csv`.

---

## Firmware upgrade behavior

| Scenario | What the tool does |
|---|---|
| Normal upload | Submits file → waits for "FIRMWARE UPDATE SUCCESSFUL" → clicks Go Home → waits for reboot → verifies version |
| 503 / upload error | Waits for device to recover → signs in → clicks Run Alternate → waits for reboot → uploads again → activates new firmware → verifies version |
| Session drops mid-upload | Detects redirect to login page → checks if upload already landed → proceeds or recovers as needed |
| Auth challenge at any step | Re-authenticates automatically and retries |

---

## Output columns

### Battery report

`Location, IP, Model, Ethernet MAC, Page Updated, Scraped At, Status, Error, UPS Battery Status, Battery Test Result, Battery Cabinet Type`

### Firmware report

`Location, IP, Model, Ethernet MAC, Scraped At, Status, Error, Current Version, Upgrade Applied, Verified Version`

---

## Supported devices

- Vertiv Liebert GXT4
- Vertiv Liebert GXT5
