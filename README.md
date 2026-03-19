# Attendance Aggregator

A Python script that merges **UBLinked CSV attendance exports** into a single, always up-to-date Excel workbook. One command after every meeting.

Built for **UB SHPE** by Nicholas Medina (Club Secretary 2025-26), but works for any club using EngageSUNY / CampusGroups.

https://github.com/nickmedina5991/Club-Attendance-Aggregator

> **tl;dr** Put the script and your CSV files in the same folder. Run `python attendance_aggregator.py`. It will scan the folder and create or update `attendance_summary.xlsx` automatically.

---

## Features

- Parses the UBLinked multi-line CSV export format automatically
- Creates or updates `attendance_summary.xlsx` with every member ever seen
- Adds a **date column per event** with a checkmark for attended, blank for absent
- Sorts members by **attendance count** (highest to lowest)
- Keeps an **Import Log** sheet so you know which files have been processed
- Skips already-imported files automatically to prevent double-counting
- Never deletes data, only adds or increments

---

## Requirements

- Python 3.10+
- pandas
- openpyxl

Install dependencies with:

```bash
pip install pandas openpyxl
```

> On some systems you may need `pip3` instead of `pip`. Run `python --version` to confirm you're on 3.10 or above.

---

## Folder Setup

Keep everything in one dedicated folder. The script always reads from and writes to `attendance_summary.xlsx` in the same directory it is run from.

```
SHPE Attendance/
    attendance_aggregator.py      <- the script
    README.md
    attendance_summary.xlsx       <- created automatically on first run
    GBM_01_23_2026.csv            <- already processed
    GBM_01_30_2026.csv            <- already processed
    GBM_02_06_2026.csv            <- new export, drop it here then run
```

---

## Usage

### After each meeting

1. **Export attendance from UBLinked**
   Go to your event, open the *Attendance* tab, and click *Export*. Save the CSV into your folder and give it a corresponding file name.

2. **Open a terminal in that folder**
   - **Windows:**  Open File Explorer and navigate to the desired folder. Click the address bar at the top, type `cmd` (for Command Prompt) or `powershell` (for PowerShell), and press `Enter`.
   - **Mac:** Right-click the folder in Finder and select *New Terminal at Folder*

3. **Run the script**:

```bash
python attendance_aggregator.py
```

With no arguments the script scans the folder for all CSV files, skips any already imported, and updates `attendance_summary.xlsx`.

4. Open `attendance_summary.xlsx`. The new date column is added and counts are updated.

---

## Command Reference

```bash
# No arguments -- scans every CSV in the current folder (recommended)
python attendance_aggregator.py

# Single specific CSV
python attendance_aggregator.py GBM_02_20_2026.csv

# Multiple specific CSVs
python attendance_aggregator.py GBM_02_20_2026.csv Workshop_02_09_2026.csv
```

---

## Duplicate Import Protection

Each time a CSV is successfully processed, its filename is recorded in the **Import Log** sheet. On every subsequent run, the script reads that log and filters out any files whose names already appear in it before doing any processing.

If you accidentally pass in a file that was already imported, you will see:

```
Skipping 1 already-imported file(s). Check Import Log for details.

  Nothing new to import. Exiting.
```

No changes are made to `attendance_summary.xlsx` in this case.

> **Important:** Do not rename CSV files after scanning them. The duplicate check is based on the original filename, so renaming a file will cause it to be treated as a new import.

---

## Output Workbook

### `Summary` sheet

| Column | Description |
|---|---|
| `First Name` | Member's first name |
| `Last Name` | Member's last name |
| `Campus Email` | Campus email, used as the unique ID to match members across files |
| `Attendance Count` | Total events attended. Rows sorted highest to lowest. |
| `1/23/2026`, `2/6/2026`, ... | One column per meeting. Checkmark = attended, blank = absent. The totals row shows how many people attended each event. |

Columns are frozen after `Attendance Count` so you can scroll right through many dates while names and attendance counts stay visible.

### `Import Log` sheet

A running record of every CSV processed, including the event name, meeting date, and the timestamp it was imported. This log is also what the script uses to detect and skip duplicate imports.

---

## Expected CSV Format

The script is built for the standard **UBLinked "Event Attendance By Event"** export. The file structure looks like this:

```
Event Attendance By Event

SHPE GBM
Start Date,1/23/2026
End Date,1/23/2026
First Name,Last Name,Campus Email,...
"Nicholas","Palma","nicpalma@buffalo.edu",... 
"Brandon","Smith","bsmith24@buffalo.edu",... 
...
```

All columns beyond `First Name`, `Last Name`, and `Campus Email` are ignored.

---

## Troubleshooting

**"No CSV files found"**
Make sure your terminal is inside the SHPE Attendance folder. Run `pwd` (Mac/Linux) or `cd` (Windows) to check your current location.

**Skipped file warning:**
```
Skipping 'file.csv' - could not find 'First Name' header row.
```
The CSV is not a standard attendance export. Make sure you exported from the *Track Attendance* tab of the event.

**Member appears twice / duplicate row**
Members are matched by campus email. If someone has two different emails they will appear as two rows. You can manually merge them in the Excel file.

**Rebuilding from scratch**
Delete `attendance_summary.xlsx` and re-run the script with all your CSVs. It will rebuild the full history from the files you pass in. Since the Import Log is also deleted, all files will be treated as new imports.

---

## How It Works

```
UBLinked CSV exports
        |
        v
  load_import_log()  checks filenames already recorded in the Import Log
        |
        v
  parse_csv()        strips multi-line header, extracts event name and date
        |
        v
  load_existing()    reads current attendance_summary.xlsx (if it exists)
        |
        v
  merge()            increments counts, adds checkmarks for new dates, adds new members
        |
        v
  write_excel()      writes sorted Summary sheet and Import Log
        |
        v
  attendance_summary.xlsx
```
