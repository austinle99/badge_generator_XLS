# Conference Badge Generator

Create print-ready conference badges from a Word template and an attendee spreadsheet. The script reads attendee data, merges it into a Word layout, exports each badge to PDF, and finally renders PNGs ready for printing.

## Features
- Automatic column normalisation for Vietnamese and English headers (`Họ và tên`, `Chức vụ`, etc.).
- Auto-detects gender prefixes when the column is missing.
- Bundled Poppler support for PDF → PNG conversion (no system installation required).
- Optional Microsoft Word COM fallback when direct conversion fails.
- PowerShell build script to package the project for distribution.

## Requirements
- Windows 10/11 with Microsoft Word installed (required by `docx2pdf`/COM fallback).
- Python 3.9+.
- Poppler binaries (already included when you run the build script or keep the `poppler/` folder).
- The following Python packages:
  - `pandas`, `docxtpl`, `docx2pdf`, `pdf2image`, `openpyxl`, `pywin32`

## Initial Setup
1. Open PowerShell in the project root.
2. Install dependencies (use `--user` if you do not have admin rights):
   ```powershell
   python -m pip install --upgrade pip
   python -m pip install pandas docxtpl docx2pdf pdf2image openpyxl pywin32
   ```
   _Tip:_ keep the same Python version across machines for consistent wheel compatibility.
3. If you obtained Poppler separately, set the `POPPLER_PATH` environment variable to its `bin` directory, or place the Poppler folder inside this project (`poppler/.../Library/bin`). The script automatically detects both locations.

## Preparing the Word Template
1. Open `badge_template.docx` in Microsoft Word.
2. Insert your background artwork and set it to **Behind Text** so you can type over it.
3. Add text boxes (or plain paragraphs) where you want each field to appear and type the placeholders exactly:
   - `{{prefix}}`
   - `{{name}}`
   - `{{position}}`
4. Format each placeholder with your desired font, colour, and alignment. Ensure the font supports Vietnamese diacritics.
5. Save the file and close Word before running the script.

## Updating the Attendee Spreadsheet
Edit `danh_sach.xlsx` and fill in attendee data. The script recognises a variety of common header names:
- Name columns: `Họ và tên`, `Họ tên`, `Tên`, `Name`
- Position columns: `Chức vụ`, `Chức danh`, `Position`, `Title`
- Prefix column (optional): `Danh xưng`, `Prefix`, `Mr/Mrs`

If the prefix column is absent, the script guesses `Mr.`/`Mrs.` based on common Vietnamese female name components.

## Running the Generator
```powershell
$env:PYTHONUTF8 = '1'            # ensures Unicode output
python badge_generator.py
```

Generated PNG files are placed in `badges/` with zero-padded indexing (e.g., `001_Nguyen_Van_An.png`). Temporary DOCX/PDF files are removed automatically after each badge is created.

## Packaging the Project
Use the provided PowerShell build script to assemble a distributable archive that includes this script, the template, sample data, and (optionally) Poppler binaries.

```powershell
.\build.ps1
```

The script creates a `dist/` directory containing:
- `conference_badge_generator.zip` — the distributable package.
- `package/` — staging folder used to build the archive (safe to delete after distribution).

### Customising the Build
- Pass a custom output directory or archive name:
  ```powershell
  .\build.ps1 -OutputDirectory out -ArchiveName badge_tools.zip
  ```
- To exclude Poppler (for machines where it is already installed system-wide), run the script with `-IncludePoppler:$false`.

## Suggested Distribution Checklist
1. Run the badges script locally to confirm the template and data look correct.
2. Execute `build.ps1` to produce `conference_badge_generator.zip`.
3. Share the ZIP along with setup instructions (this README) with the target machine.
4. On the recipient machine:
   - Extract the ZIP.
   - Install Python + dependencies.
   - Ensure Microsoft Word is installed/sign-in ready.
   - Run `python badge_generator.py` from the extracted folder.

## Troubleshooting
- **Poppler not found**: ensure `poppler/pdftoppm` exists in the project or set `POPPLER_PATH` to its `bin` directory. The script reports when it falls back to the system PATH.
- **PDF conversion fails**: confirm Microsoft Word is installed and activated. `docx2pdf` relies on it on Windows.
- **PNG conversion fails after adding Poppler**: verify the `bin` directory (containing `pdftoppm.exe`) is accessible. Run `pdftoppm -v` in PowerShell to test.
- **Template exports blank badges**: make sure the placeholders are typed as plain text (`{{name}}` etc.), not embedded in images or smart quotes.

---
Questions or improvements? Open an issue or tweak the script—it's designed to be easy to extend (e.g., additional fields, different output formats).
