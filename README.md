# attendance_management

Python-based attendance management desktop application targeting **Windows 11**.

## Features

| Feature | Details |
|---|---|
| **Excel storage** | One file per year (`Attendance_Sheet_YYYY.xlsx`), one sheet per month (e.g. `Jan`) |
| **Columns** | Date / Start time / End time / Work time — sorted by date ascending |
| **Break deduction** | Work time > 6 h → 1-hour break automatically deducted |
| **Monthly total** | When all calendar days for a month are recorded, a total row is appended |
| **Start / End buttons** | Timestamps written to Excel immediately and displayed in the window |
| **Auto end-time** | On startup, if yesterday's end time is missing the app queries Windows Event Log for the first hibernate / lock / shutdown event that lasted ≥ 3 hours |
| **Persistent folder** | Selected folder saved to `~/.attendance_config.json`; shown as `Empty` on first use |

## Requirements

- Python 3.10 or later
- Windows 11 (the Windows Event Log feature requires Windows; all other features work cross-platform)

## Installation

```powershell
# 1. Clone / download the repository
# 2. Install Python dependencies
pip install -r requirements.txt
```

## Usage

```powershell
python main.py
```

1. Click **フォルダ選択** to choose where the Excel files will be saved.
2. Click **始業** when you start work — the time is written to Excel immediately.
3. Click **終業** when you finish work — the time and calculated work time are written to Excel.
4. Click **Excelを開く** to open the current month's file directly.

## File layout

```
<selected folder>/
└── Attendance_Sheet_2025.xlsx
    ├── Jan   (columns: 日付 | 始業時間 | 終業時間 | 労働時間)
    ├── Feb
    └── ...
```
