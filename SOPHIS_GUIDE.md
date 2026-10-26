# Sophis Strikes Automation – Step-by-Step Guide

This guide explains what the automation does, what you need to run it, and how to run it safely, one step at a time.
You do not need to know anything about the project, but you do need to know the Sophis Risque screens used below.

---

## 1. What does it do?

Entering the strikes of an option contract in Sophis is repetitive: for each code you select the contract, set the
right *Term*, open the **Strikes <NAME>** window and type a *Minimum* and a *Maximum* for every maturity date.

The script does this for you, starting from an Excel file:

| Stage | What happens | Who does it |
|-------|--------------|-------------|
| **1. Extraction** | The Excel is read. The script prints a summary and tells you **which code to look for first** in Sophis. | Script |
| **2. Selection and opening** | You find the code in the *Listed market* grid and select its row. The script then reads the row's **Term**, sets the same Term in the dropdown above the right-hand table, and opens the **Strikes** window with the first button next to *Available contracts*. | You (selection) + Script |
| **3. Filling** | **3 seconds after** it clicks the icon, the script starts filling the **Strikes** window (Maturity / Minimum / Maximum / Ric): it finds each date and types its Minimum and Maximum, confirming each value with Enter. It does not check first whether the window is open. | Script (automatic) |

> **The script never presses OK.** When it finishes, the **Strikes** window is left open so that you can check every
> value and press **OK yourself**.

### How it works (so you know what to expect)
Sophis runs through Citrix, so the script cannot read Sophis directly. It works like a person would: it **takes
screenshots, reads the text on them (OCR), and moves the mouse / presses keys** on your screen. Because of that:

* Sophis must be **visible on the screen** (not minimized, not covered by another window, including the terminal).
* **Do not touch the mouse or keyboard while the script is typing.**
* Moving the mouse to a **corner of the screen** aborts the script immediately (safety feature, "fail-safe").

---

## 2. Requirements

### 2.1 Computer and software
* The computer from which you **see Sophis** (the one that shows the Citrix window), running **Windows**.
  The script must run on that machine, not inside Citrix.
* **Python 3.10 or newer** (check with `python --version`). Download it from <https://www.python.org/downloads/>
  and tick **"Add Python to PATH"** during installation.
* **Git** (optional; you can also download the repository as a ZIP file from GitHub).
* Internet access for the first installation only.

### 2.2 Screen
* Use the **same screen setup every time** (same monitor, same resolution and Windows display scaling).
  If you change the scaling or the resolution, run the calibration again (section 4).
* The Citrix/Sophis window can be anywhere on the screen and can be moved between runs – the script finds it by
  its table headers – but it must be **fully visible**.

### 2.3 Sophis screens
Two Sophis screens are used:
1. The contract list ("**Listed market**" grid with columns *Mnemo, Extended Mnemo, Description, RIC, Term, …*),
   with the **Term dropdown** and the **Available contracts** label above the right-hand table.
2. The **Strikes <NAME>** window with the grid *Maturity | Minimum | Maximum | Ric* (opened by the first button next
   to *Available contracts*).

### 2.4 The Excel file
The Excel (`.xlsx`, `.xls` or `.csv`; an `.xls` that is really an HTML or XML table is also read) must always have this layout (a header row is fine, it is skipped):

| Column | Content | Example | Used? |
|--------|---------|---------|-------|
| A | Bloomberg equity | `SNDK US Equity` | no |
| B | Call code | `SNDK US 01/15/27 C45` | no |
| C | Put code | `SNDK US 01/15/27 P45` | no |
| **D** | **Date** (month/day/year) | `01/15/27` | **yes** |
| **E** | **Value** | `45.00` | **yes** |
| **F** | **Code** | `SNDK` | **yes** |
| G – L | other values / letters | `A`, `100.00`, … | no |

How the values are chosen:
* Rows are grouped by **code (column F)** and, inside each code, by **date (column D)**.
* For each date, the **first** value in column E (reading top to bottom) is the **Minimum** and the **last** value
  is the **Maximum**. It is the *order of the rows*, not the size of the numbers, that counts.
* If a file has several codes, each code is processed separately (one Strikes window per code).
* If a date has only one row, that value is used for both Minimum and Maximum (a warning is printed).

---

## 3. One-time installation

Open **PowerShell** or **Command Prompt** and run:

```
git clone https://github.com/sebashills3-gif/Sebas-1
cd Sebas-1
git checkout claude/zealous-mccarthy-2tjacr
pip install -r requirements_sophis.txt
```

*(No Git? Download the repository as a ZIP from GitHub, unzip it, open the terminal in that folder and run only the
last command. Make sure you use the branch `claude/zealous-mccarthy-2tjacr`.)*

This installs the libraries for mouse/keyboard control (`pyautogui`), image handling (`pillow`, `opencv-python`),
text recognition (`rapidocr-onnxruntime`) and Excel reading (`openpyxl`, `xlrd`).
The first time the script starts it may take a few seconds to load the text-recognition engine.

Optional – only if you want to start step 2 with a keyboard shortcut instead of the console (see 5.3):
```
pip install keyboard
```

Quick check that everything is installed:
```
python sophis_auto.py --help
python sophis_excel.py --help
```
Both must print a help text without errors.

> **Updating later:** `git pull` inside the folder gets the latest version of the script.

---

## 4. One-time calibration (and when to repeat it)

The script has to learn where things are on your screen. This is done **once**; the result is saved in
`sophis_config.json` and the folder `sophis_anchors/` (they are not shared; each computer has its own).

**Repeat the calibration when:** you change the screen resolution or Windows scaling, switch monitors, change the
Sophis layout (columns, grid size), or the script keeps failing to find the grid.

### Steps
1. Open Sophis and show the *Listed market* screen with the grid and the right-hand table fully visible.
   Make sure the **Term column is wide enough** to show the whole text (e.g. "Short term index" instead of
   "Short ter…"). Keep **at least a few rows** visible in the grid.
2. Open a terminal in the project folder and run:
   ```
   python sophis_auto.py calibrate
   ```
3. The script asks for **14 mouse positions**, one at a time. For each one it shows an instruction, then counts
   down (5 seconds by default). **Put the mouse at the described place before the countdown ends and do not move
   it.** The position is recorded automatically when the countdown ends.
   (Use `--wait 8` to get more time per point.)
   Two of the points are *pairs of corners* that define a **header strip** (the script remembers how it looks, so it
   can find the window later even if you move it):
   * the grey header of the grid: top-left corner of "Mnemo" and bottom-right corner at the end of "Description";
   * the header of the right-hand table: top-left of "Date" and bottom-right after "Name".

   The other points are the text of the first / second / last visible row, the "Extended Mnemo" column, the borders of
   the RIC and Term columns, the dropdown (text and arrow), and the **first small button next to "Available
   contracts"**.
4. At the end the script validates the calibration and prints, for both header strips, a *match* close to `1.00`
   and a *shift* of `(+0,+0)`. If it prints an error (for example that the last row must be the last *fully*
   visible row), read the message and run `calibrate` again.
5. To fix only some points without repeating everything:
   ```
   python sophis_auto.py calibrate --only row1 row2 rowN
   ```

> The Citrix window is allowed to move after calibration. The mouse positions are stored relative to the headers
> and are recomputed on every run.

---

## 5. Running it

### 5.1 Before you start (checklist)
- [ ] The Excel file is closed or saved, and follows the layout of section 2.4.
- [ ] Sophis is open on the *Listed market* screen; the grid and the right-hand table are fully visible.
- [ ] The Term column is wide enough to read the Term completely.
- [ ] No other window covers Sophis. The **terminal is allowed to be open, but not on top of the Sophis window**
      once the script starts clicking (put it on another monitor, or in a corner outside the areas Sophis uses).
- [ ] Any **previous Strikes window is closed** (OK/Cancel pressed by you).
- [ ] You are not going to use the mouse or keyboard while the script is working.

### 5.2 Optional – preview the Excel first
This only reads the file and prints what the script understood; it does not touch Sophis:
```
python sophis_excel.py strikes.xlsx
python sophis_excel.py                 # no file given: a window opens to choose it
```
Check that each code, each date and each Minimum/Maximum is what you expect. Use `--code SNDK` to see one code and
`--sheet NAME` to choose a sheet (for `.xls`, the first sheet is used by default).

### 5.3 Run the three stages
```
python sophis_auto.py run
```
A **file selector window opens**: choose the Excel (`.xlsx`, `.xls` or `.csv`). (If your computer cannot show the
window, the console asks you to paste the path instead. You can also give the path directly:
`python sophis_auto.py run strikes.xlsx`.) Useful options: `--code SNDK` to process only one code,
`--sheet NAME`, `--click-wait 8`, `--open-delay 5`, `--slowdown 1.5` for a slower, more careful run; see section 7.

**Stage 1 – Extraction.** The console prints something like:
```
[stage 1] Excel read
SNDK: 2 date(s)
   15 Jan 2027   Minimum       45   Maximum      100   (3 row(s))
   19 Feb 2027   Minimum       30   Maximum       70   (2 row(s))

Search FIRST in Sophis: SNDK
```
If the file has more codes, it also lists the order in which they will be processed. Read the warnings, if any.

**Stage 2 – Your selection, then automatic opening.**
1. In Sophis, find the code printed after *Search FIRST in Sophis* in the *Listed market* grid and **select its row
   with one single click**. (Click once, on the row, and do not double-click. Do not type in the cells.)
2. Come back to the console and **press ENTER**. You can instead type a number of seconds (for example `8`) if you
   need more time to switch back to Sophis. Also available: `s` = skip this code, `q` = quit.
3. A short countdown runs: **switch to Sophis, leave it in front and do not touch anything**.
   (If you prefer to avoid switching windows, start the script with `--hotkey f9` after `pip install keyboard`, and
   press **F9** instead of ENTER while Sophis is in front.)
4. The script now, by itself:
   * finds the selected row and reads its **Mnemo** and **Term**,
   * sets the **dropdown** above the right-hand table to that Term and checks that it matches,
   * clicks the first button next to **Available contracts** to open the **Strikes** window.

**Stage 3 – Automatic filling.** Exactly **3 seconds after the click on the icon** the script starts by itself
(it does not wait to "see" the window, so make sure the Strikes window really opens in that time; if Sophis is slow,
use `--open-delay 5`). For every date of the code it:
1. finds the row with that date in the Maturity column (scrolling the list if needed),
2. double-clicks the **Minimum** cell, types the value and presses **Enter**,
3. moves to the **Maximum** cell of the same row, types the value and presses **Enter**.

Do not touch anything until the console says the code is finished. A summary is printed:
```
SNDK: 2 date(s) written, 0 skipped
The update window was LEFT OPEN: check the values and press OK in Sophis yourself.
```

### 5.4 After each code
1. **Check the values** in the Strikes window against the Excel summary.
2. Press **OK** (or Cancel) in Sophis yourself.
3. If there are more codes, the console shows the next one: go back to Stage 2 for that code
   ("find it, select the row with one click, press ENTER").
4. At the end the script prints a final summary of all codes (written / skipped dates).

### 5.5 What gets skipped (and why)
A date is **skipped, not guessed**, and reported in the summary, when:
* it is **not in the Maturity list** of the Strikes window;
* it appears **more than once** in the list (including rows with a suffix such as `-1`), because the script cannot
  know which row you mean. Fill those by hand, or fill them with the advanced command in section 6.3.
* its row cannot be read.

Everything else in the batch continues normally.

---

## 6. Other useful commands

### 6.1 Read-only checks (nothing is typed)
```
python sophis_auto.py diagnose            # shows how the selected row of the grid is read
python sophis_auto.py window              # reads the open Strikes window (dates, columns)
python sophis_auto.py window --date "19 Oct 2026" --show   # moves the pointer over that date's Minimum/Maximum cells
```
Images are saved in the folder `sophis_debug/`.

### 6.2 Stage 2 only (select a row, set the Term, open the window)
```
python sophis_auto.py open
```

### 6.3 Stage 3 only (the Strikes window is already open)
Dry run – locates the cells but clicks and types nothing:
```
python sophis_auto.py fill --excel strikes.xlsx --code SNDK --dry-run
```
Real run:
```
python sophis_auto.py fill --excel strikes.xlsx --code SNDK
python sophis_auto.py fill --entry "15 Jan 2027" 45 100 --entry "20 Nov 2026#1" 30 70
```
`DATE#0` is the first row with a repeated date and `DATE#1` the second one.
The dates are processed **one at a time** (find the date, type its Minimum and Maximum, then the next date). With `fill` (unlike `run`) a missing or repeated date **stops** the run at that date; the dates before it have already been written. `--dry-run` is the only mode that scans all the dates first.

---

## 7. Options that help when something is slow or fails

| Option | Effect |
|--------|--------|
| `--slowdown 1.5` (or `2`) | Waits longer after each action. Use it if Citrix is slow and clicks arrive late. |
| `--click-wait 8` | More seconds to switch back to Sophis after pressing ENTER (stage 2). |
| `--open-delay 5` | Seconds between the click on the *Available contracts* icon and the start of the filling (default 3). Raise it if the Strikes window opens slowly. |
| `--no-widen` | Does not try to enlarge the Term column automatically. |
| `--strict-verify` | Stops if a typed value reads back differently (default: warn and continue, because the on-screen text is sometimes misread). |
| `--no-verify` | Does not read the cells back after typing. |
| `--no-trace` | Does not save a screenshot at each step (by default they go to `sophis_debug/`). |
| `--hotkey f9` | Starts stage 2 with a global key instead of ENTER (needs `pip install keyboard`). |

---

## 8. Troubleshooting

| What you see | What to do |
|--------------|------------|
| `Calibration …` / the grid or the right table is not found | The Sophis window is covered, minimized, or the layout/scaling changed. Make it fully visible; if nothing changed, run `calibrate` again. |
| It cannot tell which row is selected | Select the row again with **one** click, wait a second, press ENTER and do not move the mouse. Check that the row is visibly highlighted. |
| `Could not read the Term of row …` | The Term column is too narrow. Widen it by dragging the right border of its header, then retry. |
| The dropdown is not set to the right Term | Check that the Term text is fully visible; retry with a bigger `--slowdown`. Send the screenshot from `sophis_debug/` if it persists. |
| `could not read the update window` / `Could not find the Maturity … header on screen` | The filling started before the Strikes window was open (Sophis was slow) or the icon was not clicked. Make sure the window is open and answer `r` (retry only repeats the filling), or run again with a bigger `--open-delay`. |
| `could not be opened as an old Excel file` | The `.xls` is damaged or protected. Open it in Excel and save it as `.xlsx`. |
| `The minimum cell … did not open for editing … NOTHING was typed` | Another cell/editor may be open, or the window lost focus. Click the title bar of the Strikes window once and answer `r`. |
| `WARNING: the … cell was read back as … but … was typed` | Usually a misread of the screen, not a typing error. Compare the cell with the Excel value by eye. |
| `SKIPPED <date>: … not in the list / appears N times` | See section 5.5. Fill that date by hand, or use `fill --entry "DATE#n" MIN MAX`. |
| The script clicked something unexpected | Move the mouse to a **corner of the screen** to abort immediately. Press **Cancel** in Sophis if values were already typed that you do not want, and run again after checking the setup. |
| Anything else | Copy the whole console output and the newest files in `sophis_debug/` (screenshots) and send them to the maintainer. |

When the script stops with an error it asks `r` (retry), `s` (skip this code) or `q` (quit). Retrying after the
Strikes window is already open only repeats the filling stage; it does not open the window a second time. Retrying
re-types all dates of that code, which is harmless but worth checking afterwards.

---

## 9. Safe-use rules (summary)

1. **You** select the row in Sophis; the script does not search for it.
2. Never touch the mouse or keyboard while the script is typing.
3. The script **never presses OK** – always check the values and confirm them yourself.
4. Abort at any time by moving the mouse to a corner of the screen (or `Ctrl+C` in the terminal).
5. Check the console summary after every code, especially any `SKIPPED` or `WARNING` lines.
6. After changing screen resolution, scaling or the Sophis layout, run `calibrate` again.

---

## 10. Files in the project

| File | Purpose |
|------|---------|
| `sophis_auto.py` | Main program (commands: `calibrate`, `run`, `open`, `fill`, `window`, `diagnose`). |
| `sophis_excel.py` | Reads the Excel/CSV, shows the file selector, and extracts the codes, dates and Minimum/Maximum values. |
| `requirements_sophis.txt` | Libraries to install. |
| `test_sophis_auto.py`, `test_sophis_excel.py` | Automatic tests (developers only; run with `pytest`). |
| `sophis_config.json`, `sophis_anchors/` | Your calibration (created by `calibrate`; specific to your computer). |
| `sophis_debug/` | Screenshots and logs saved by the script while it works; useful when reporting a problem. |
