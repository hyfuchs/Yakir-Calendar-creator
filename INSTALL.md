# Install / Move to Another Machine — Yakir Calendar Creator

The project is plain Python 3 + a few pip packages. It runs on **WSL, Windows, or Mac**.
Copy the whole project folder, install Python + the packages, and run the same commands.

## What to copy

Copy the project directory — everything lives inside it now:
- `Yakir-Calendar-creator/` — all `.py` scripts, `yakir_logo.png`, the `.md` docs, `requirements.txt`,
  and the data CSVs (`modiin_calendar.csv`, `modiin_weeks.csv`)
- the Modiin candles PDF(s) (input for regenerating the CSVs)

The old master `.xlsx` (in the parent folder) is **not** needed to generate the calendar
(validation only) — copy it only if you want the `--source xls` fallback.

---

## Requirements (all platforms)

- **Python 3.9 or newer**
- Python packages (from `requirements.txt`): `pymupdf`, `ephem`, `pyluach`, `python-docx`, `openpyxl`
- **Hebrew fonts** (for correct rendering): present by default on Windows and Mac.
- **For PDF export** (optional): Microsoft Word **or** LibreOffice. The `.docx` is always produced; PDF is a convenience.

---

## Windows (native, no WSL)

1. Install Python from https://python.org — **check "Add Python to PATH"** during install.
2. Open **PowerShell** or **Command Prompt** in the project folder:
   ```powershell
   cd "C:\path\to\Yakir-Calendar-creator"
   python -m pip install -r requirements.txt
   ```
3. Run with `python` (not `python3`):
   ```powershell
   python parse_modiin_pdf.py --pdf "path\to\Modiin-<year>-Candels.pdf" --out "..\modiin_calendar.csv"
   python modiin_weeks.py
   python yakir_calendar.py --theme purple --pdf --hyrmo 578701 578702 ... --out yakir_5787.docx
   ```
- **PDF**: Word is auto-detected on Windows (the script drives it). Nothing extra needed if Word is installed.
- **Fonts**: Arial/Calibri + Hebrew are built in.

---

## Mac

1. Install Python 3 (if not present):
   ```bash
   brew install python3        # or download from python.org
   ```
2. In Terminal, in the project folder:
   ```bash
   cd "/path/to/Yakir-Calendar-creator"
   python3 -m pip install -r requirements.txt
   ```
3. Run exactly as on WSL (`python3 ...`, same commands as HOW_TO_RUN.md).
- **PDF**: install LibreOffice (`brew install --cask libreoffice`) — the script uses its
  `soffice` headless converter. (The Windows/Word COM path does not apply on Mac.)
  Alternatively, open the `.docx` and export PDF from Word/Pages manually.
- **Fonts**: Arial + Hebrew are built in. If Hebrew looks wrong, the script falls back to
  DejaVu Sans (the code auto-detects Windows fonts; on Mac it uses the default).
  > Note: `yakir_calendar.py` picks fonts by checking for `/mnt/c/Windows/Fonts` (a WSL
  > path). On native Mac/Windows it will use the `DejaVu Sans` fallback name unless you
  > adjust the `HAS_WIN`/font constants near the top of the file to `"Arial"`/`"Calibri"`.

---

## WSL (current setup — for reference)

```bash
cd "/mnt/c/yfuchs/OneDrive/Documents/Yakir-Calendar-experimental/Yakir-Calendar-creator"
pip install -r requirements.txt        # add --break-system-packages if pip refuses
python3 modiin_weeks.py
python3 yakir_calendar.py --theme purple --pdf --hyrmo 578701 ... --out yakir_5787.docx
```
- PDF works via the installed Windows Word (the script converts the WSL path and drives Word over PowerShell).
- Fonts: reads the Windows font folder at `/mnt/c/Windows/Fonts`.

---

## Quick sanity check after install

```bash
python3 -c "import fitz, ephem, pyluach, docx, openpyxl; print('all deps OK')"
python3 yakir_calendar.py --hyrmo 578601 --out /tmp/test.docx   # one month, no PDF
```
If that produces a `.docx`, the install is good.

---

## One portability note (fonts)

The font auto-detection in `yakir_calendar.py` keys off the WSL path `/mnt/c/Windows/Fonts`.
On **native Windows or Mac** that path won't exist, so it uses the `DejaVu Sans` fallback.
For best appearance on a native laptop, set the font constants near the top of
`yakir_calendar.py`:
```python
F_ARIAL   = "Arial"
F_CALIBRI = "Calibri"
```
(Both are present on Windows and Mac.) This is the one small edit worth making when moving
off WSL. Everything else runs unchanged.
