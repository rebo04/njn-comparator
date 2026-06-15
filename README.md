# NJN Comparator

Compares PPAP NJN Excel documents revision by revision and highlights every change directly inside the document — same format as the original, with orange/green/red cells.

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
4. One Excel file is generated per comparison pair

### Color legend in the output files

| Color | Meaning |
|-------|---------|
| 🟠 Orange cell | Value changed — hover to see the previous value in the comment |
| 🟢 Green row | Row is new in this revision |
| 🔴 Red row | Row was removed (shown at the bottom of the sheet) |
| White | No change |

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
