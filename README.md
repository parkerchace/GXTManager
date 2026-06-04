# GXTManager

A tool for managing Vertiv GXT-4 and GXT-5 UPS units over the network. Run battery health reports, push firmware upgrades, configure SNMPv3, silence alarms, restore output after outages, and manage NIC cards across your whole UPS fleet.

**No command line experience needed. No Vertiv cloud subscription required. Works on Mac and Windows.**

---

## Setup (do this once)

You only need to do this the first time. It takes about 5 minutes.

### Step 1 - Install Python

Go to [python.org/downloads](https://www.python.org/downloads/) and download the latest version.

**On Windows - this part is important:** When the installer opens, look for a checkbox at the bottom of the first screen that says **"Add Python to PATH"**. Check that box before you click Install Now. If you miss it, uninstall Python and run the installer again.

**On Mac:** Use the installer from python.org. The version built into macOS is outdated and will not work.

### Step 2 - Install Firefox

Go to [mozilla.org/firefox](https://www.mozilla.org/firefox/) and install it if you don't already have it. The tool uses Firefox to log into each UPS's web interface automatically. You don't need to configure anything in Firefox itself.

### Step 3 - Download the tool

Download this repo as a ZIP from GitHub (click the green **Code** button, then **Download ZIP**). Unzip it somewhere easy to find, like your Desktop.

That's it for setup. Everything else (drivers, Python packages) downloads itself the first time you launch.

---

## How to launch it

### On Mac

1. Open the folder in Finder
2. Double-click **Run Vertiv Scraper.command**
3. If Mac says it can't be opened: right-click the file, choose **Open**, then click **Open** again in the popup that appears
4. A black terminal window will appear. The first launch downloads some packages automatically, which takes about 30 seconds. The app window will open when it's ready.

### On Windows

1. Open the folder in File Explorer
2. Double-click **Run Vertiv Scraper.bat**
3. A black command prompt window will appear. The first launch downloads some packages automatically, which takes about 30 seconds. The app window will open when it's ready.
4. If a blue "Windows protected your PC" screen appears, click **More info**, then **Run anyway**. This happens because the file isn't signed by a publisher Windows recognizes. It's safe.

> After the first launch, every launch after that opens in just a few seconds.

---

## What you need to use it

You don't need any Vertiv cloud account or special software. You just need:

- The **IP address** of each UPS on your network
- A **label** for each one - just a name that makes it easy to identify in reports (room number, closet name, etc.)
- The **username and password** for the UPS web interfaces (the same login you'd use if you opened the device's IP in a browser)

---

## How to use it

### Step 1 - Enter your credentials

Type your UPS web interface username and password into the Credentials fields at the top of the app. These are never saved unless you click **Save Config**.

### Step 2 - Paste your device list

In the **Targets** box, paste a list of your devices. The easiest way is to copy two columns from a spreadsheet - Location and IP Address - and paste them directly:

```
BLDG-100-CLOSET-A    10.0.1.15
SERVER-ROOM-UPS      10.0.1.22
4TH-FLOOR-SW         10.0.1.30
```

You can also paste just a plain list of IP addresses (one per line) if you don't need location labels.

### Step 3 - Pick a tab and click Start

| Tab | What it does |
|---|---|
| **Battery Report** | Pulls battery charge, health, temperature, time remaining, and alarm events from each device |
| **Firmware Upgrade** | Checks the firmware version on each device and upgrades it if needed |
| **SNMPv3 Config** | Configures SNMPv3 settings and disables the older SNMPv1/v2c |
| **Silence Alarm** | Silences the audible alarm on each device |
| **Output** | Turns UPS output back on after a power outage |
| **Restart NIC** | Restarts the network card on each device |
| **NIC Events** | Checks NIC status and flags devices that need a restart |

Click **Start** to run against every device in your list. The **Parallel** spinner controls how many devices are worked at the same time (default is 3 - you can increase this if you have a lot of devices and want it to run faster).

When a run finishes, the results are saved as a CSV file next to the tool and opened automatically.

### Saving and loading your settings

If you run this tool regularly, use **Save Config** to save your credentials and SNMPv3 settings to a file so you don't have to retype them every time. Use **Load Config** to load them back. Keep that file somewhere safe - it contains passwords.

---

## Firmware upgrade notes

The firmware tab handles a few tricky situations automatically:

| Situation | What happens |
|---|---|
| Normal upload | Uploads file, waits for success confirmation, verifies the new version |
| Device returns an error (503) | Waits for the device to recover, then retries the upload automatically |
| Connection drops mid-upload | Checks if the firmware already landed, skips re-upload if it did |
| Device asks to log in again | Re-authenticates automatically and picks up where it left off |

---

## Output columns

### Battery report

Columns vary by device model, but common ones include:

`Location, IP, Model, Status, UPS Battery Status, Battery Charge Status, Battery Time Remaining, Battery Percentage Charge, Battery Temperature, Battery State of Health, Battery last replaced time, Battery Test Result, Replace Battery, Battery Low, Battery Test Failed, Scraped At, Error`

### Firmware report

`Location, IP, Model, Current Version, Target Version, Upgrade Applied, Verified Version, Scraped At, Error`

### SNMPv3 report

`Location, IP, Model, Status, Restart Required, NIC Restarted, Scraped At, Error`

### Silence Alarm / Output / Restart NIC reports

`Location, IP, Model, Status, Scraped At, Error`

### NIC Events report

`Location, IP, Model, Restart Required, Status, Scraped At, Error` (plus a column for every event the NIC reports)

---

## Supported devices

- Vertiv Liebert GXT4
- Vertiv Liebert GXT5
