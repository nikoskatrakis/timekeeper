# Timekeeper

A free, open-source, ad-free time tracking app for macOS. Free forever. No ads.

Timekeeper is inspired by the Pomodoro technique. You can manipulate the time recorded in the database where its saved (you'll need to download an app to open it, if you dont have one) and the app will export times in Excel. Manual recording is also feasible. Finally there are keyboard shortcuts to start/pause/stop the countdown, and a menu with various options. 

---

## Features

- **Countdown timer** — ticks down for each task interval and break, shown as a circular progress arc that shifts colour (green → orange → red)
- **Task management** — create, rename, and switch between tasks at any time, even mid-interval
- **Automatic time recording** — every interval is logged with date, start time, end time, and duration
- **Manual time entry** — add time to any task for periods when Timekeeper was not running
- **Daily goal** — set a target number of intervals per day and track your progress
- **Excel export** — all time entries are written to `timekeeper.xlsx` and can be opened directly from the app
- **File/folder links** — attach a file or folder to any task for quick access; links appear as clickable hyperlinks in the spreadsheet
- **Menu bar icon** — lives in the macOS menu bar so it is always one click away
- **Keyboard shortcuts** — control the app without leaving your current window
- **API** — set task names and adjust interval/break durations programmatically

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl + B | Start / Pause interval |
| Ctrl + S | Stop interval |
| Ctrl + M | Open menu |

---

## Installation

Below is a set of instructions to install it under your machine, in the /Documents/Timekeeper area. You can obviously install it elsewhere, but I have not tested that, I'm not a programmer to know what needs changing. 

### Requirements

- macOS
- Python 3.9 or later
- A virtual environment (recommended)

### 1. Clone the repository

```bash
cd ~/Documents
git clone https://github.com/nikosandthepython/timekeeper.git
cd timekeeper
```

### 2. Create and activate a virtual environment under Documents/timekeeper

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install PySide6 openpyxl pyobjc-framework-Cocoa pyinstaller
```

### 4. Build the app

```bash
bash build_app.sh
```

This builds `timekeeper.app` and copies it to `/Applications`. You can then launch it from the Dock or Finder like any other macOS app.
> You must allow it to save to your documnets folder. it will create a Timekeeper folder under Documents where it'll save its database and the associated excel spreadsheet. 
> If macOS blocks the app on first launch: System Settings → Privacy & Security → Open Anyway

---

## Running without building

You can run Timekeeper directly from the terminal without building an app bundle:

```bash
python3 timekeeper_app.py
```

---

## Data & Privacy

All task and time data is stored locally in a SQLite database (`timekeeper.db`). **Nothing is sent to the internet.** Everything stays on your computer.

---

## Licence

Timekeeper is open-source software released under the [MIT Licence](https://opensource.org/licenses/MIT).

This means you are free to use, copy, modify, merge, publish, distribute, and sell copies of the software — for personal or commercial use — as long as the original licence notice is included. No restrictions. No fees. No ads.
