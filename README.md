# NJN Comparator

**Created by Arturo Rebolledo**

Compares PPAP NJN Excel documents revision by revision and highlights every change directly inside the document — same format as the original. Add 2 or more revision files and it compares every pair, color-codes every cell/row that changed, and collapses repetitive part-revision renames into a single flagged change instead of dozens of duplicate diffs.

---

## Option A — Download & Run (no Python needed)

Go to the [**Releases**](../../releases/latest) page and download:

| Platform | File | How to open |
|----------|------|-------------|
| **Windows** | `NJN_Comparator.exe` | Double-click |
| **macOS** | `NJN_Comparator_macOS.zip` | Unzip → **right-click** the `.app` → **Open** *(first time only, to bypass Gatekeeper)* |

---

## Option B — Run from source (Python)

### First time setup

**macOS** — double-click `install_mac.sh`
> If macOS blocks it: right-click → Open → Open

**Windows** — double-click `install_windows.bat`

Both scripts automatically install Python and all packages if missing, then launch the app.

### After setup (daily use)

| Platform | Launch |
|----------|--------|
| macOS | double-click `install_mac.sh` |
| Windows | double-click `install_windows.bat` |

Or from terminal:
```bash
python app.py
```

---

## How it works

1. Click **+ Add Files** and select 2 or more NJN `.xlsx` files
2. Choose an output folder (defaults to Desktop)
3. Click **▶ COMPARE NOW**
4. One Excel file is generated per pair (every file is compared against every other file you added)

### Color legend in the output files

| Color | Label | Meaning |
|-------|-------|---------|
| 🟡 Lemon yellow cell | CHANGED | Cell value was modified — hover the cell to see the previous value |
| 🟢 Mint green row | ADDED | Row is new in this revision |
| 🔴 Strawberry pink row | DELETED | Row was removed — shown in its original position |
| 🔵 Sky-blue banner | SYSTEMIC | A whole column changed across most rows (e.g. a part-revision letter that appears throughout) |
| 🟣 Lavender banner/cell | RENAMED | The same value was substituted for another one in many places (e.g. a part-number revision bump like `656475500F` → `656475500G`) — every occurrence is still highlighted, but it's reported once instead of once per row |
| ⚪ Light gray | UNCHANGED | No difference |

A legend explaining each color is also added at the bottom of every generated report.

---

## Security & privacy

NJN Comparator works **100% offline** — it never touches the network, so it's safe to run on confidential PPAP/BOM documents.

- **No network code at all.** Neither `app.py` nor `comparator.py` imports `requests`, `urllib`, `socket`, or any other networking module — there is nothing in the source that *could* phone home, even by accident.
- **No telemetry, no analytics, no auto-update check.** The app doesn't call out anywhere, ever.
- **Local-only file I/O.** It only reads the `.xlsx` files you pick and writes the comparison report next to them (or to the output folder you choose), using `openpyxl`. Your documents never leave your machine.
- **Verify it yourself** — the entire source is in this repo (`app.py` + `comparator.py`); a quick read or a `grep -i "requests\|urllib\|socket\|http"` over both files will show zero matches.

The only outbound connection in this whole *project* is GitHub's own infrastructure building the `.exe`/`.app` releases via Actions when a new version is tagged — the app itself, once downloaded, runs fully local.

---

## Requirements (if running from source)

- Python 3.9+
- `customtkinter`
- `openpyxl`

```bash
pip install customtkinter openpyxl
```

---

## Build executables locally

```bash
pip install pyinstaller customtkinter openpyxl

# macOS
pyinstaller --windowed --collect-all customtkinter --name "NJN_Comparator" app.py

# Windows
pyinstaller --onefile --windowed --collect-all customtkinter --name "NJN_Comparator" app.py
```
