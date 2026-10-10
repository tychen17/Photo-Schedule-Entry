# Photo Schedule Entry

Batch-convert photos with travel-info filenames into a structured Excel itinerary.

If you name your photos with dates and locations directly in the filename, this tool scans an entire folder and parses them into a spreadsheet — no manual transcription needed.

## Features

- **GUI workflow**: Select folder → scan & preview → export Excel. No command line required.
- **Inline editing**: Double-click any cell to edit directly in the table; press Enter to confirm, Esc to cancel.
- **Color coding**: Yellow = needs review, Green = manually edited, Red = missing date.
- **Auto event classification**: Keyword-based routing into four columns: Exhibit / Event / Performance / Sports.
- **Filter by issues**: One-click filtering to show only rows that need attention.
- **Chinese–English parsing**: Domestic addresses auto-fill province and city; foreign addresses split by state/region and city dictionaries.
- **Issue report**: Exports a plain-text checklist alongside the Excel file, listing each pending file and the reason.

## Usage

### Option 1: Run the exe (Windows)

Download `PhotoScheduleEntry.exe` from the [Releases](../../releases) page and double-click. No Python required.

### Option 2: Run from source

```bash
# 1. Install dependencies (only one third-party library)
pip install -r requirements.txt

# 2. Launch
python photo_schedule_gui.py
```

Requires Python 3.8+. tkinter is included with standard Python distributions.

### Steps

1. Click "Browse…" and select your photo folder.
2. Adjust scan options as needed, then click **Start Scan**.
3. Review yellow rows in the preview table; double-click cells to correct.
4. When satisfied, click **Export Excel**.

Output files are written **inside the selected photo folder**:

| File | Contents |
|---|---|
| `日程记录_YYYYMMDD_HHMMSS.xlsx` | Main spreadsheet, 9 columns |
| `待人工确认_YYYYMMDD_HHMMSS.txt` | Pending items with reasons |

## Filename Convention

The parser reads filenames in this order:

```
{date}{country}{province/state}{city}{venue}{event}
```

- **Date** (required): Formats like `2021年1月11日`. Missing dates are marked red.
- **Country**: Falls back to China if unrecognized (outputs `中国`).
- **Province / City**: Split using built-in administrative division and city dictionaries; no separator required.
- **Venue / Event**: Separated by venue-type suffixes (museum, gallery, theater, stadium, etc.) and event keywords.

### Supported Formats

`.jpg` `.jpeg` `.png` `.gif` `.bmp` `.tif` `.tiff` `.webp` `.avif` `.heic` `.heif`
Plus common RAW: `.dng` `.arw` `.cr2` `.cr3` `.nef` `.nrw` `.orf` `.pef` `.raf` `.raw` `.rw2` `.rwl` `.sr2` `.srf` `.srw` `.x3f`

## Output Columns

```
Date | Country | Province/State | City | Venue | Exhibit | Event | Performance | Sports
```

Events are placed only in their matching column; the rest remain empty.

## Custom Dictionaries

The built-in dictionaries cover only **major cities** (US metropolitan areas, state capitals, and world cities). To add smaller cities you visit frequently — without modifying the source:

Place a `_local_places.py` file next to `photo_schedule.py`:

```python
EXTRA_CITIES_EN = ["Exampletown", "Sample City"]
EXTRA_CITIES_CN = ["示例市"]
EXTRA_CITY_TO_EN = {"示例市": "Exampletown"}
```

All three variables are optional. The program loads them automatically on startup. This file is excluded by `.gitignore` — **it contains places you personally visit; do not commit it to a public repository**.

> Why not ship a comprehensive dictionary? A gazetteer that "covers someone's complete travel history" would leak that person's itinerary by itself. Keeping personal additions local is an intentional design trade-off.

## Parsing Limitations

Filenames are free text. The program marks the following cases in yellow for human review rather than guessing:

- Fields with no separators and locations not in the dictionary
- Words that could be either a venue or an event name
- Missing dates, or a filename containing only a city name with no specific venue

Yellow is not failure — the program chooses to ask rather than silently fill in wrong data.

## Building an exe

```bash
pip install pyinstaller
pyinstaller 照片日程录入.spec
```

Output: `dist/PhotoScheduleEntry.exe`.

> On Windows the exe may trigger SmartScreen (no code signing). Choose "Run anyway".

## Project Layout

| File | Description |
|---|---|
| `photo_schedule.py` | Parsing engine: filename splitting, event classification, dictionaries, Excel generation |
| `photo_schedule_gui.py` | GUI: table preview, inline editing, status stats, logging |
| `照片日程录入.spec` | PyInstaller build configuration |

The GUI and engine are decoupled — call the engine directly for batch scripting:

```python
from photo_schedule import scan_folder, parse_filename

for path in scan_folder("/path/to/photos"):
    print(parse_filename(path))
```

## Privacy

- All parsing happens **entirely offline** — no network access, no data upload.
- The program reads only filenames, not photo contents.
- Generated spreadsheets and checklists are written only to your chosen folder.