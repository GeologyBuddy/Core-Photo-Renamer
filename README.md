# GB•CoreNamer 2026

> Bulk-rename drill core photo photos using interval data — hole by hole, fast and accurately.

<img width="501" height="407" alt="image" src="https://github.com/user-attachments/assets/04e69135-ac8b-439b-8bdc-8356139763d1" />


---


## What It Does

GB•CoreNamer reads an interval spreadsheet (Excel or CSV) and a folder of core photos, then renames every image to a standardised filename that encodes the **hole ID**, **box range**, and **depth interval** — automatically, in bulk, with a live preview before anything is touched. 

**Output filename format:**
```
{HoleID}_{Bx001-004}_{0000.0m-0016.0m}_{Dry|Wet}.jpg
```
Example:
```
DDH001_Bx001-004_0000.0m-0016.0m_Dry.jpg
```

---

## Features

- **Flexible interval file support** — loads `.xlsx`, `.xls`, and `.csv` files
- **Column mapping dialog** — map any column name to the required fields at load time; no need to rename your spreadsheet headers
- **Live rename preview** — see old → new filename pairs in a scrollable table before committing
- **Dry / Wet photo type detection** — automatically detected from the folder name
- **Threaded renaming** — progress bar updates in real time; the UI stays responsive during large batches
- **One-click undo** — reverses the last rename operation completely
- **Collision-safe** — appends a counter suffix if a target filename already exists
- **Partial group support** — handles the last group of boxes even when fewer than four rows remain

---

## Requirements

| Dependency | Version |
|---|---|
| Python | 3.8 or later |
| pandas | any recent |
| Pillow | any recent |

Install dependencies:
```bash
pip install pandas pillow openpyxl
```

> `openpyxl` is required for `.xlsx` support via pandas.

---

## Running the App

```bash
python GB-CoreNamer.py
```

No arguments needed — the GUI handles everything.

---

## Usage

1. **Load your interval file** — click Browse under *Interval File Selection* and select your `.xlsx` or `.csv` file.
2. **Map your columns** — a dialog will appear asking you to match your column names to: `Hole ID`, `Box Number`, `From (m)`, `To (m)`.
3. **Select your image folder** — click Browse under *Image Folder Selection*. Name the folder with `dry` or `wet` (case-insensitive) and the photo type will be set automatically.
4. **Preview** — the table fills with old → new filename pairs. Check before proceeding.
5. **Rename** — click **Rename Files**. The progress bar tracks completion.
6. **Undo if needed** — click **Undo Rename** to restore all original filenames.

---

## Interval File Format

Your spreadsheet must contain at minimum these four columns (column names can be anything — you map them in step 2):

| Hole ID | Box Number | From (m) | To (m) |
|---|---|---|---|
| DDH001 | 1 | 0.0 | 4.0 |
| DDH001 | 2 | 4.0 | 8.0 |
| DDH001 | 3 | 8.0 | 12.0 |
| DDH001 | 4 | 12.0 | 16.0 |

Photos are grouped in sets of four rows. Each group of four produces one renamed file. If the total number of rows is not a multiple of four, the final partial group is still processed.

---

## Folder Naming Convention

The app detects photo type from the image folder name:

| Folder name contains | Photo type tag |
|---|---|
| `dry` (any case) | `Dry` |
| `wet` (any case) | `Wet` |
| anything else | *(empty)* |

---

## Image Formats

Only the JPG format is supported. The app does not convert or modify image data — only the filename is changed.

---

## License

GB•CoreNamer is licensed under the GNU General Public License.
Created and developed by **Mehmet Duyan** — ©Mehmet Duyan, 2026.

---

## Links

- 🌐 [geologybuddy.com](https://geologybuddy.com)
- 🐙 [github.com/GeologyBuddy](https://github.com/GeologyBuddy)
- ▶️ [youtube.com/@GeologyBuddy](https://www.youtube.com/@GeologyBuddy)
- 📸 [instagram.com/geologybuddy](https://www.instagram.com/geologybuddy/)
