# Counseling Appointment Data Extractor — SARS PDF

A Python script that reads counseling appointment records exported from the **SARS Short Name History Report** (Detail view) and outputs a structured, analysis-ready Excel file.

This script works for any counseling program that uses SARS — including EOPS, CARE, NEXT UP, Veterans, Promise, CalWORKs, and others. The only thing that changes between programs is the list of reason codes, which is explained below.

---

## Background

Counseling departments that use SARS generate appointment history reports that can be exported in two formats: **Excel** or **PDF**.

### Why not just use the Excel export?

The Excel version is easier to request but comes with several problems that make it unreliable for analysis:

- **No counselor name in the data** — the counselor's name appears in the report header but is not included as a column in the exported rows, so there is no way to filter or group appointments by counselor
- **Rows get skipped** — when a reason code comment is too long, SARS splits it across multiple rows in the export, and Excel misreads those rows, dropping data
- **Reason codes are not parsed** — all reason codes are lumped together in a single cell with no separation, making it impossible to filter or analyze individual codes
- **Summary tables corrupt the data** — the attendance summary block at the bottom of each counselor section gets mixed into the data rows, producing garbage records
- **No clean structure for dashboards** — the raw Excel export is not in a format that can be directly loaded into Power BI or pivot tables without significant manual cleanup

### Why PDF instead?

Requesting the report as a PDF requires an extra extraction step, but the PDF preserves the full layout — including the counselor name in the page header — and gives complete, unaltered data. This script handles the extraction automatically, producing a clean structured Excel file that is ready for analysis and dashboards.

This project was built to support institutional research workflows at a community college and is demonstrated here using fully anonymized mock data (cartoon character names, fictional student IDs, and celebrity counselor names). No real student data is included in this repository.

---

## What the Script Does

1. **Reads the PDF** page by page using `pdfplumber`
2. **Extracts the counselor name** from the page header on each page (`Short Name: [NAME]`)
3. **Identifies and skips** column header rows and the summary block at the bottom of each counselor section
4. **Parses each appointment row** into structured fields: Source, Student ID, Student Name, Date, Time, Duration, Reason Code(s)/Comments, and Attendance Status
5. **Adds calculated columns**: Term code, Term label, Intersession flag, and Modality (In-Person vs Online)
6. **Parses reason codes** from the comments field — each recognized code gets its own column (`Reason_1` through `Reason_20`)
7. **Saves everything** to an Excel file in the same folder as the PDF

---

## Adapting This Script to Any Counseling Program

Every counseling program that uses SARS has its own set of reason codes — the short labels counselors select when logging an appointment. These codes are configured locally at each campus and can vary by program and by year.

**Before running this script on your PDF:**

1. Request the current reason code list from your SARS administrator, or pull it directly from the SARS system under your program's configuration
2. Open `SARS_extraction.py` and find the `REASON_CODES` list at the top of the file
3. Replace or update the list to match your program's codes exactly as they appear in SARS
4. Save the script and run it

The script will only detect and capture the codes that appear in `REASON_CODES`. Any code not in the list will be silently ignored, so keeping this list current is important — especially at the start of a new academic year when codes may be added or changed.

### Example: updating for a different program

```python
# EOPS / CARE / NEXT UP example
REASON_CODES = [
    'AA', 'ADMIN', 'CA', 'CANCEL', 'CARE CON', 'CC', 'CE', 'CERT',
    'CNCOMPL', 'CRS RETAKE', 'CSU TRAN', 'DISQ', 'ECOMPL', 'ED PLAN',
    'ESARSONLINE', 'WORKSHOP', 'GENCOUN', 'HOLD', 'NEXT UP', 'NOTES',
    'OTHER PERS', 'OUT', 'PETITION', 'PRICE', 'SEP', 'SICK', 'SPRBRK',
    'TRANS', 'UC TRAN', 'VACATION', 'VIRTUAL INFO', 'WRKSHOP', 'WALKIN',
    'WALKIN & WRAPUP'
]

# Veterans example
REASON_CODES = [
    'CANCEL', 'VA-CLASS ASSISTANCE', 'VA-EDUCATIONAL PLAN', 'VA-FILES',
    'VA-PHONE', 'VA-TEMPPLANS', 'VETERAN', 'VSC-ACADEMIC ADV',
    'VSC-COMPUTER USE', 'VSC-GUEST PRESENTER', 'VSC-SOCIAL EVENTS',
    'VSC-STUDY', 'VSC-WORKSHOPS',
]
```

Note that some programs have a small number of reason codes (Veterans uses 13) while others have many more (EOPS/CARE/NEXT UP uses over 30). Programs with fewer codes will naturally produce more records with `Total_Reasons = 0` — this is expected and means the appointment had free-text comments but no recognized code. If this happens frequently, review your PDF and check whether counselors are using codes that are not in your current list.

---

## Output Columns

| Column | Description |
|---|---|
| Source | Where the appointment was booked (e.g. THE GRID or Drop-in) |
| Appointment Type (Source) | Appointment or Drop-in |
| Attendance Status | Attended / Not Attended / Not Marked |
| INTERSESSION | Intersession label if the date falls in January intersession, otherwise blank |
| Counselor | Counselor short name extracted from the page header |
| Student ID | Student identifier |
| Student Name | Student name as it appears in SARS |
| Date | Appointment date (YYYY-MM-DD) |
| Term | Term code (e.g. 22247) |
| Term Recode | Human-readable term (e.g. Fall 2024) |
| Time | Appointment time |
| Duration | Duration in minutes (blank for Drop-ins) |
| Total_Reasons | Number of reason codes found in the comments |
| Modality | In-Person or Online (inferred from keywords in comments) |
| Reason Code(s) / Comments | Full comment text from SARS |
| Reason_1 … Reason_20 | Individual parsed reason codes |

---

## How Reason Code Parsing Works

Counselors in SARS can select multiple reason codes to describe a single appointment. In the raw export, all selected codes appear together in one cell — for example: `"CC, NEXT UP - follow up meeting - THE GRID"`. This makes it impossible to filter or count individual codes in Power BI or pivot tables.

This script solves that by scanning each comment string against the known reason code list and storing each match in its own dedicated column:

```
Reason_1 = 'CC'
Reason_2 = 'NEXT UP'
Total_Reasons = 2
```

### Why Reason_1 through Reason_20?

Counselors sometimes select multiple reason codes to describe a single meeting. By splitting each code into its own column, every individual reason is captured and accessible. Up to 20 reason codes per appointment are supported, which covers any realistic combination a counselor would select.

Having each reason code in its own column means you can **unpivot** the `Reason_1` through `Reason_20` columns in Power BI to get one row per reason code per appointment. This lets you accurately count how many times each reason code was used, filter appointments by reason, and build breakdowns by counselor, term, or modality — none of which is possible when all codes are crammed into a single cell.

---

## Known Manual Step — Long Rows That Span Pages

When a counselor's appointment comment is very long, the last row of a table on one page may overflow and continue onto the next page without a closing bottom border. Because `pdfplumber` uses table borders to detect where rows begin and end, an unclosed row is invisible to the script — it cannot be read as a valid row.

**How to fix this before running the script:**

1. Open the PDF in Adobe Acrobat or any PDF editor
2. Find the page where the long row is cut off — the bottom border of that row will be missing
3. Copy the bottom border line from any other closed row in the table and paste it onto the open row to close it
4. Save the PDF
5. Run the script normally

Copying an existing row border is the fastest approach — it takes only a few seconds and ensures the line is the correct weight and style without having to draw anything from scratch.

**What the script does automatically for wrapped text within a page:**

When `pdfplumber` splits a long comment across two rows within the same page, the second fragment has no student ID, no date, and no time — only text in the comments position. The script detects this and automatically appends the fragment to the previous record's `Reason Code(s) / Comments` field, keeping the full comment intact.

---

## Requirements

```
pandas
pdfplumber
openpyxl
```

Install with:

```bash
pip install pandas pdfplumber openpyxl
```

---

## How to Run

1. Clone or download this repository
2. Place your SARS PDF in the same folder as the script
3. Open `SARS_extraction.py` and update these two lines in `main()`:

```python
pdf_directory = r"C:\path\to\your\folder"   # folder containing the PDF
pdf_filename  = "your_sars_report.pdf"       # exact filename
```

4. Update the `REASON_CODES` list at the top of the script to match your program's codes
5. Double-click the script, or run from the command line:

```bash
python SARS_extraction.py
```

6. The output Excel file (`SARS_Cleaned_Data.xlsx`) will appear in the same folder as the PDF

---

## Files in This Repository

| File | Description |
|---|---|
| `SARS_extraction.py` | Main extraction script |
| `Mock_SARS_Counseling_Data.pdf` | Sample mock PDF for testing |
| `README.md` | This file |

---

## Mock Data Note

The sample PDF included in this repository uses:
- **Cartoon characters** as student names (Mickey Mouse, Bugs Bunny, SpongeBob, etc.)
- **Fictional letter-based IDs** (GGG-001, GGG-002, etc.)
- **Celebrity names** as counselors (Oprah W, Serena W, Michelle O, etc.)
- **Real SARS reason code formats** so the parsing logic can be tested end to end

No real student records, counselor names, or institutional data are included.
