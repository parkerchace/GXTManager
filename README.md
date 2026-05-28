# GXTManager

A GUI tool for managing Vertiv GXT-4 and GXT-5 UPS units over the network. Run battery health reports, push firmware upgrades, configure SNMPv3, silence alarms, restore output after outages, and manage NIC cards across your whole UPS fleet. No command line experience needed, no Vertiv cloud subscription required.

Works on **macOS and Windows**.

---

## What you need before running this

1. **Python 3.9 or newer** - download from [python.org](https://www.python.org/downloads/)
   - On Windows: during installation, check the box that says **"Add Python to PATH"** before clicking Install Now
   - On macOS: the installer from python.org is recommended (the built-in macOS Python is outdated)
2. **Firefox** - download from [mozilla.org](https://www.mozilla.org/firefox/)
   - The script controls Firefox to log into each UPS web interface. It must be installed but you do not need to do anything special with it
3. **The files from this repo** - either download the ZIP from GitHub or clone it somewhere easy to find, like your Desktop

Everything else (geckodriver, Python packages) is handled automatically on first run.

---

## How to run it

### On macOS

1. Open the folder in Finder
2. Double-click **Run Vertiv Scraper.command**
3. If macOS says it can't be opened, right-click it and choose **Open**, then click **Open** again in the dialog
4. A terminal window will appear, install the required packages on first run, then open the app

### On Windows

1. Open the folder in File Explorer
2. Double-click **Run Vertiv Scraper.bat**
3. A command prompt window will appear, install the required packages on first run, then open the app
4. If Windows Defender SmartScreen appears, click **More info** then **Run anyway**

> The first launch takes about 30 seconds longer than usual while it sets up a virtual environment and downloads the required packages. Every launch after that is fast.

---

## No cloud management system needed

Vertiv offers cloud-based NIC management platforms, but you don't need any of that to use this tool. All you need is:

1. Each UPS device's **IP address** on your network
2. A **label** for it - just a name you give the device so results are easy to read (e.g. the room number, closet name, or whatever makes sense to you)
3. The **web interface credentials** for the devices (the username and password you would use to log in through a browser)

The script logs into each device's built-in web interface directly, the same way you would manually in a browser.

---

## What it does

**Battery Report**
- Logs into each device and scrapes every field from the Battery status page: charge percentage, time remaining, state of health, voltage, temperature, test result, last replaced date, and more
- Also captures battery alarm events: Replace Battery, Battery Low, Battery Test Failed, etc.
- Exports everything to a timestamped CSV that opens automatically when the run finishes

**Firmware Upgrade**
- Checks the current comm card firmware version on each device
- Uploads a new `.fl` firmware file to devices that need it
- Handles errors automatically. If a device returns a 503 or drops the session mid-upload, the tool recovers and retries without any manual intervention
- Confirms the installed version after each upgrade
- Exports a per-device result CSV

**SNMPv3 Config**
- Configures SNMPv3 User 1 on each device (username, access type, auth protocol, auth secret, privacy protocol, privacy secret, trap targets, and trap port)
- Disables SNMPv1/v2c after applying SNMPv3 settings
- Automatically checks NIC events after configuration. If a System Restart Required flag is active, the NIC is restarted automatically
- CSV includes Restart Required and NIC Restarted columns so you can see at a glance which devices needed it

**Silence Alarm**
- Navigates to the right page for the device model (some use System, others use System Configuration), enables commands, and clicks Silence Alarm
- Confirms the dialog and logs the result

**Output**
- Turns UPS output back on after a power outage
- Navigates to Output, enables commands, and clicks Turn Output ON
- Useful when you get outage alerts and need to remotely restore output across multiple devices

**Restart NIC**
- Navigates to Communications > Support, enables commands, and restarts the NIC card
- Good for applying changes that require a reboot or clearing a System Restart Required flag

**NIC Events**
- Navigates to Communications > Status and scrapes all NIC status and event fields
- Flags any device with System Restart Required: Active as YES in the CSV so they are easy to spot
- Useful as a follow-up check after firmware upgrades or SNMPv3 changes

---

## How to use it

### Step 1 - Enter credentials

Type your UPS web interface username and password into the Credentials fields at the top. These are never saved to disk.

### Step 2 - Paste your device list

The target box accepts two columns copied straight from a spreadsheet: **Location** (your label for the UPS) and **IP Address**, separated by a tab:

```
BLDG-100-CLOSET-A    10.0.1.15
SERVER-ROOM-UPS      10.0.1.22
4TH-FLOOR-SW         10.0.1.30
```

Copy those two columns from Excel or any spreadsheet and paste directly into the box. You can also paste just a plain list of IPs (one per line) if you don't need location labels.

### Step 3 - Choose a tab and run

| Tab | What it does |
|---|---|
| **Battery Report** | Pulls all battery metrics and alarm events from each device |
| **Firmware Upgrade** | Checks and optionally upgrades comm card firmware |
| **SNMPv3 Config** | Configures SNMPv3 User 1 and disables SNMPv1/v2c |
| **Silence Alarm** | Silences the audible alarm on each device |
| **Output** | Turns UPS output back on after a power outage |
| **Restart NIC** | Restarts the NIC card on each device |
| **NIC Events** | Checks NIC status and flags devices that need a restart |

Click **Start** to run against all pasted targets. The Parallel spinner controls how many devices are worked simultaneously (default 3). Each completed run saves a timestamped CSV next to the script and opens it automatically.

---

## Firmware upgrade behavior

| Scenario | What the tool does |
|---|---|
| Normal upload | Submits file, waits for "FIRMWARE UPDATE SUCCESSFUL", clicks Go Home, waits for reboot, signs in, verifies version |
| 503 / upload error | Waits for device, signs in, navigates to Firmware Update, clicks Enable, clicks Run Alternate, accepts confirmation dialog, stays on reboot page until login appears, signs in, uploads again, activates new firmware via Run Alternate, waits for reboot, verifies version |
| Session drops mid-upload | Detects redirect to login page immediately, checks if upload already landed (version match), skips recovery if it did, otherwise runs full recovery |
| Auth challenge at any step | Re-authenticates automatically and retries the current step |
| Communications-only page load | Skips missing nav elements and proceeds directly to Firmware Update link |

---

## Output columns

### Battery report

All fields from the Battery status page are captured dynamically, so the exact columns depend on what each device reports. Common ones include:

`Location, IP, Model, Ethernet MAC, Page Updated, Scraped At, Status, Error, UPS Battery Status, Battery Charge Status, Battery Test Result, Battery Cabinet Type, Battery Time Remaining, Battery Percentage Charge, Battery Current, DC Bus Voltage, Battery Temperature, Battery State of Health, Battery last replaced time, Battery Self Test, Replace Battery, Battery Low, Battery Test Failed, ...`

### Firmware report

`Location, IP, Model, Gen, Current Version, Target Version, Upgrade Mode, Upgrade Applied, Upload Status, Verified Version, Scraped At, Error`

### SNMPv3 report

`Location, IP, Model, Status, Restart Required, NIC Restarted, Scraped At, Error`

### Silence Alarm / Output / Restart NIC reports

`Location, IP, Model, Status, Scraped At, Error`

### NIC Events report

`Location, IP, Model, Restart Required, Status, Scraped At, Error, System Status, System Restart Required, ...` (plus a column for every event the NIC reports)

---

## Supported devices

- Vertiv Liebert GXT4
- Vertiv Liebert GXT5
