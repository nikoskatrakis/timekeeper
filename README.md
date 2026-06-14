# Timekeeper

A free, open-source time tracking app for macOS. Free forever. No ads. Ever.

---

## What it does

Timekeeper helps you work efficiently across multiple tasks by keeping a countdown timer always visible in your Mac menu bar and automatically recording how long you spend on each task.

---

## How it works

### The menu bar item
Timekeeper lives in the native macOS menu bar at the top of the screen. It shows the current task name and a live countdown clock (e.g. `⏱  Write report  27:14`). During a break it shows `☕  Break  04:32`. Clicking the menu bar item toggles the main panel open or closed.

### The main panel
When open, the panel shows:
- The name of the current task
- A **digital countdown clock** ticking down to zero
- A **circular progress arc** that shifts colour as time runs out (green → orange → red for tasks; yellow for breaks)
- Controls for managing tasks and timers (see Features below)

Clicking anywhere outside the panel hides it. The timer keeps running in the menu bar. Clicking the menu bar item brings the panel back.

---

## Features

### Timer
- Countdown timer for **task intervals** (default 30 minutes, adjustable)
- Countdown timer for **break intervals** (default 5 minutes, adjustable)
- **Strict alternation** — a break must follow every task interval and vice versa; you cannot run two task or two break intervals consecutively
- **Pre-loaded next timer** — when an interval ends or is stopped, the next timer appears stationary (yellow for break, green for task) so you can start it when you're ready
- **Pause / Resume** — pause any running timer and resume it later
- **Stop** — stop the current timer early; the next timer is pre-loaded automatically
- **Sounds** — a fanfare plays when a task interval ends; a beep plays when a break ends
- Circular progress arc with colour shift: green → orange → red as time runs out

### Task management
- **New Task** — create a task at any time, even mid-interval; the timer keeps running and time is split automatically
- **Switch Task** — switch to an existing task mid-interval; the timer continues uninterrupted and elapsed time is recorded against the previous task
- **Rename a task** — change the name of the current task at any time
- **Add time manually** — add time retroactively to any task for occasions when Timekeeper was not running
- **No time recorded during breaks** — switching tasks during a break has no effect on time recording

### File and folder linking
- Link any task to a file or folder on your computer via the 📎 button
- If a link already exists, you are prompted before it is overwritten
- Links are stored in the Excel timesheet as clickable hyperlinks

### Excel timesheet
Timekeeper automatically maintains an Excel file (`timekeeper.xlsx`) with two tabs. Click **Open Linked Excel Timesheet** in the panel to open it.

**Tab 1 — Time Log**
| Date | Task Name | Duration (min) | Start | End | Manual Entry |
|------|-----------|---------------|-------|-----|--------------|
| 2026-06-14 | Write report | 27.5 | 09:00:00 | 09:27:30 | No |
| 2026-06-14 | Review emails | 30.0 | 09:32:00 | 10:02:00 | No |

**Tab 2 — Task Links**
| Task Name | Linked File / Folder | Date Created |
|-----------|----------------------|--------------|
| Write report | /Users/niko/Documents/report.docx | 2026-06-14 |

The spreadsheet updates automatically whenever a task period is saved — no manual export needed.

### Data persistence
- All task and time data is stored locally in a SQLite database (`timekeeper.db`)
- Nothing is sent to the internet; everything stays on your computer
- The Excel file is regenerated automatically from the database

### On quit
When you click **Quit Timekeeper**, any time accumulated in the current task interval is saved automatically before the app exits.

---

## How to run

1. Open `tk_v2.ipynb` in VS Code
2. Click **Run All** (▶▶)
3. The panel appears and the menu bar item is ready

To stop the app, click **Quit Timekeeper** in the panel. The Jupyter kernel stays alive.

---

## Technical details

- **Language:** Python (Jupyter Notebook — `tk_v2.ipynb`)
- **Platform:** macOS (primary)
- **Key libraries:**
  - `PySide6` — Qt GUI framework
  - `pyobjc-framework-Cocoa` — native macOS menu bar integration via NSStatusBar
  - `openpyxl` — Excel spreadsheet creation and updating
  - `sqlite3` — internal storage of task and time data (standard library)
- **Architecture:** SOLID principles throughout — each class has a single responsibility
- **Open source:** Published on GitHub under an open source licence

---

## What it is not

- It is not a subscription service
- It has no ads
- It does not send any data anywhere
- It does not require an internet connection

---

## Planned future improvements

- Standalone `.app` packaging (no VS Code required to launch)
- Database reset / fresh start feature
- Cross-platform support (Windows, Linux)
- Sync across devices
