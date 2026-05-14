"""
Counseling Appointment Data Extractor — SARS PDF
=================================================
Reads a SARS-generated PDF (Short Name History Report, Detail view),
extracts appointment records table by table, and outputs a structured
Excel file with parsed reason codes, term codes, modality, and
intersession flags.

Works for any counseling program that uses SARS — update the REASON_CODES
list at the top of this file to match your program's codes before running.

Dependencies:
    pip install pandas pdfplumber openpyxl

Usage:
    1. Update REASON_CODES below to match your program's reason codes.
    2. Set pdf_directory and pdf_filename in main() to point to your PDF.
    3. Double-click the script or run: python SARS_extraction.py
    4. The output Excel file is saved in the same folder as the PDF.
"""

import pandas as pd
import pdfplumber
import re
import os
from datetime import datetime


# ── Reason code list ───────────────────────────────────────────────────────
#
# Update this list to match the reason codes used by your counseling program.
# Every counseling program that uses SARS has its own set of codes configured
# locally — and they can vary by campus and by year.
#
# Before running this script, request the current reason code list from your
# SARS administrator, or pull it from the SARS system under your program's
# configuration. Then replace the codes below with your program's exact codes.
#
# The script will only detect and capture the codes listed here. Any code
# that appears in your PDF but is not in this list will be silently ignored.
# Review your PDF each academic year in case codes have been added or changed.
#
REASON_CODES = [
    'AA', 'ADMIN', 'CA', 'CANCEL', 'CARE CON', 'CC', 'CE', 'CERT',
    'CNCOMPL', 'CRS RETAKE', 'CSU TRAN', 'DISQ', 'ECOMPL', 'ED PLAN',
    'ESARSONLINE', 'WORKSHOP', 'GENCOUN', 'HOLD', 'NEXT UP', 'NOTES',
    'OTHER PERS', 'OUT', 'PETITION', 'PRICE', 'SEP', 'SICK', 'SPRBRK',
    'TRANS', 'UC TRAN', 'VACATION', 'VIRTUAL INFO', 'WRKSHOP', 'WALKIN',
    'WALKIN & WRAPUP'
]


# ── Text cleaning ──────────────────────────────────────────────────────────

def clean_text(text):
    """Normalize whitespace and strip leading/trailing spaces."""
    if not text:
        return ""
    text = str(text).strip()
    text = re.sub(r'\s+', ' ', text)
    return text


# ── Counselor name extraction ──────────────────────────────────────────────

def extract_counselor_name_from_page(page_text):
    """
    Extract the counselor short name from the page header.
    Looks for the pattern: Short Name: [NAME]
    Handles multi-word names such as 'MICHELLE O' or 'ALEM'.
    """
    if not page_text:
        return ""

    pattern = r'Short\s*Name:\s*([A-Z][A-Z\s]+?)(?:\n|$|\d)'
    match = re.search(pattern, page_text, re.IGNORECASE)
    if match:
        name = match.group(1).strip()
        name = re.sub(r'\s+', ' ', name)
        return name

    return ""


# ── Row classification ─────────────────────────────────────────────────────

def is_summary_row(row):
    """
    Return True if the row belongs to the summary block at the bottom
    of each counselor section (Attendance Status, Attended, Total, etc.).
    """
    if not row or len(row) < 2:
        return False

    student_id_text = str(row[1] if len(row) > 1 else "").strip().upper()
    source_text     = str(row[0] if len(row) > 0 else "").strip().upper()

    summary_indicators = [
        'ATTENDANCE STATUS', 'ATTENDED', 'NOT ATTENDED',
        'NOT MARKED', 'DROP-INS', '#', 'TOTAL', 'HOURS', 'MINUTES'
    ]

    if any(ind in student_id_text for ind in summary_indicators):
        return True

    if 'TOTAL' in source_text:
        return True

    row_text = " ".join(str(cell).strip().upper() for cell in row if cell)
    if any(ind in row_text for ind in ['# HOURS MINUTES', 'DESCRIPTION COUNT', 'ADDITIONAL CONTACTS']):
        return True

    return False


def is_header_row(row):
    """Return True if the row contains column header labels."""
    if not row:
        return False

    row_text = " ".join(str(cell).strip().upper() for cell in row if cell)
    header_terms = [
        'SOURCE', 'STUDENT ID', 'STUDENT NAME', 'DATE',
        'TIME', 'DURATION', 'REASON CODE', 'ATTENDANCE STATUS'
    ]
    return sum(1 for term in header_terms if term in row_text) >= 3


# ── PDF extraction ─────────────────────────────────────────────────────────

def extract_pdf_table_data(pdf_path):
    """
    Open the SARS PDF and extract all appointment records.

    Two-pass approach:
      Pass 1 — read the counselor name from the page header on every page.
      Pass 2 — extract table rows, skipping headers and summary blocks.

    The counselor name is carried forward across continuation pages so that
    every record is correctly attributed even when the header only appears
    on the first page of a section.
    """
    print(f"Reading PDF: {os.path.basename(pdf_path)}")

    all_data            = []
    summary_rows_skipped = 0
    header_rows_skipped  = 0
    page_counselors      = {}

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)

        # Pass 1: extract counselor names
        print("\nExtracting counselor names from each page...")
        for page_num, page in enumerate(pdf.pages, 1):
            page_text = page.extract_text()
            if page_text:
                counselor_name = extract_counselor_name_from_page(page_text)
                if counselor_name:
                    page_counselors[page_num] = counselor_name
                    print(f"  Page {page_num}: Found counselor '{counselor_name}'")
                else:
                    page_counselors[page_num] = ""
                    print(f"  Page {page_num}: No counselor name found")

        print("\n" + "="*60)
        print("Processing tables...")
        print("="*60)

        # Pass 2: extract table data
        last_known_counselor = ""

        for page_num, page in enumerate(pdf.pages, 1):
            print(f"\nProcessing page {page_num}/{total_pages}")

            # Carry the counselor name forward across continuation pages
            if page_counselors.get(page_num, ""):
                last_known_counselor = page_counselors[page_num]
            current_counselor = last_known_counselor

            tables = page.extract_tables({
                "vertical_strategy":   "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance":      3,
                "join_tolerance":      3,
                "edge_min_length":     3,
                "min_words_vertical":  1,
                "min_words_horizontal": 1
            })

            if not tables:
                print(f"  No tables found on page {page_num}")
                continue

            print(f"  Found {len(tables)} table(s) on page {page_num}")

            for table_num, table in enumerate(tables):
                if not table or len(table) == 0:
                    continue

                print(f"  Processing table {table_num + 1} with {len(table)} rows...")
                data_rows_found     = 0
                summary_block_started = False  # stop appending once summary begins

                for row in table:
                    if not row or not any(cell for cell in row if cell and str(cell).strip()):
                        continue

                    row_cleaned = [clean_text(cell) for cell in row]

                    if is_summary_row(row_cleaned):
                        print(f"    Skipping summary row: {row_cleaned[:2]}...")
                        summary_rows_skipped += 1
                        summary_block_started = True
                        continue

                    if is_header_row(row_cleaned):
                        print(f"    Skipping header row")
                        header_rows_skipped += 1
                        continue

                    row_text = " ".join(row_cleaned)

                    # Identify valid data rows by student ID format and date
                    has_student_id = bool(re.search(r'[A-Z]{3}-\d{3}', row_text))
                    has_date       = bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}', row_text))

                    if has_student_id and has_date:
                        record = parse_table_row(row_cleaned, page_num)
                        if record:
                            record['Counselor'] = current_counselor
                            all_data.append(record)
                            data_rows_found += 1
                    else:
                        # Append wrapped reason text only before the summary block
                        if (not summary_block_started
                                and all_data
                                and is_continuation_row(row_cleaned, all_data[-1])):
                            append_to_previous_record(row_cleaned, all_data[-1])
                            print(f"    Appended continuation text to previous record")

                print(f"    Extracted {data_rows_found} data rows from table {table_num + 1}")

    print(f"\n{'='*60}")
    print(f"EXTRACTION SUMMARY:")
    print(f"  Total pages processed:          {total_pages}")
    print(f"  Summary rows skipped:           {summary_rows_skipped}")
    print(f"  Header rows skipped:            {header_rows_skipped}")
    print(f"  Appointment records extracted:  {len(all_data)}")

    counselor_counts = {}
    for record in all_data:
        counselor = record.get('Counselor', 'Unknown')
        counselor_counts[counselor] = counselor_counts.get(counselor, 0) + 1

    print(f"\nCounselor distribution:")
    for counselor, count in counselor_counts.items():
        print(f"  {counselor or 'Unknown'}: {count}")
    print(f"{'='*60}")

    return all_data


# ── Continuation row handling ──────────────────────────────────────────────

def is_continuation_row(row, previous_record):
    """
    Return True when a row contains wrapped text from the previous record's
    Reason Code / Comments field rather than a new appointment record.
    Continuation rows have no student ID, date, or time.
    """
    if not row or not previous_record:
        return False

    row_text = " ".join(str(cell).strip() for cell in row if cell)

    has_student_id = bool(re.search(r'[A-Z]{3}-\d{3}', row_text))
    has_date       = bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}', row_text))
    has_time       = bool(re.search(r'\d{1,2}:\d{2}\s*[APMapm]{2}', row_text))

    if has_student_id or has_date or has_time:
        return False

    has_text       = bool(re.search(r'[A-Za-z]', row_text)) and len(row_text.strip()) > 3
    prev_has_reason = previous_record.get('Reason', '') != ''

    return has_text and prev_has_reason


def append_to_previous_record(row, previous_record):
    """Append continuation text to the Reason field of the previous record."""
    if not row or not previous_record:
        return

    for cell in row:
        if cell and str(cell).strip() and re.search(r'[A-Za-z]', str(cell)):
            continuation_text = str(cell).strip()
            current_reason = previous_record.get('Reason', '')
            previous_record['Reason'] = (
                current_reason + " " + continuation_text if current_reason
                else continuation_text
            )
            break


# ── Row parsing ────────────────────────────────────────────────────────────

def parse_table_row(row, page_num):
    """
    Parse one table row into a structured dictionary.
    Fields are identified by content patterns rather than fixed column indexes
    because pdfplumber may merge or split cells depending on the PDF layout.
    """
    try:
        source = student_id = student_name = ""
        date = time = duration = attendance = ""

        row_text       = " ".join(row)
        row_text_upper = row_text.upper()

        # Source
        if 'THE GRID' in row_text_upper:
            source = 'THE GRID'
        elif 'DROP-IN' in row_text_upper:
            source = 'Drop-in'

        # Student ID
        id_match = re.search(r'([A-Z]{3}-\d{3})', row_text)
        if id_match:
            student_id = id_match.group(1)

        # Date
        date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', row_text)
        if date_match:
            date = date_match.group(1)

        # Time
        time_match = re.search(r'(\d{1,2}:\d{2}\s*[APMapm]{2})', row_text, re.IGNORECASE)
        if time_match:
            time = time_match.group(1)

        # Attendance status — checked in priority order (Not Attended before Attended)
        attendance_terms = [
            ('Not Attended', 'NOT ATTENDED'),
            ('Attended',     'ATTENDED'),
            ('Not Marked',   'NOT MARKED'),
            ('Drop-in',      'DROP-IN'),
        ]
        for display_term, search_term in attendance_terms:
            if search_term in row_text_upper:
                for cell in row:
                    if cell and search_term.lower() in str(cell).lower():
                        cell_text = str(cell).strip()
                        if 'Not Attended' in cell_text or 'NOT ATTENDED' in cell_text:
                            attendance = 'Not Attended'
                        elif 'Not Marked'   in cell_text or 'NOT MARKED'   in cell_text:
                            attendance = 'Not Marked'
                        elif 'Attended'     in cell_text or 'ATTENDED'     in cell_text:
                            attendance = 'Attended'
                        elif 'Drop-in'      in cell_text or 'DROP-IN'      in cell_text:
                            attendance = 'Drop-in'
                        break
                if not attendance:
                    attendance = display_term
                break

        # Duration — digits immediately after the time field
        if time_match:
            after_time = row_text[row_text.find(time_match.group(1)) + len(time_match.group(1)):]
            dur_match  = re.search(r'(-|\d+)\s', after_time)
            if dur_match:
                duration = dur_match.group(1)

        # Student name — text between the ID and the date
        if id_match and date_match:
            id_end   = row_text.find(id_match.group(1)) + len(id_match.group(1))
            date_pos = row_text.find(date_match.group(1))
            if id_end < date_pos:
                name_text   = row_text[id_end:date_pos].strip()
                name_text   = re.sub(r'^\W+|\W+$', '', name_text)
                name_text   = re.sub(r'[\-\|]', '', name_text).strip()
                student_name = name_text

        # Reason — whatever remains after removing all identified fields
        reason_text = row_text
        for value in [source, student_id, student_name, date, time, attendance]:
            if value:
                pos = reason_text.lower().find(value.lower())
                if pos != -1:
                    reason_text = reason_text[:pos] + reason_text[pos + len(value):]
        if duration and duration != '-':
            reason_text = reason_text.replace(duration, '')

        reason = clean_text(reason_text)
        reason = re.sub(r'^[\s\-|]+|[\s\-|]+$', '', reason)
        reason = re.sub(r'\s+', ' ', reason)

        record = {
            'Page': page_num, 'Source': source,
            'Student ID': student_id, 'Student Name': student_name,
            'Date': date, 'Time': time, 'Duration': duration,
            'Reason': reason, 'Attendance': attendance
        }

        for key in record:
            if isinstance(record[key], str):
                record[key] = clean_text(record[key])

        return record if (student_id and date) else None

    except Exception as e:
        print(f"Error parsing row: {e}\nRow data: {row}")
        return None


# ── Date / time / duration helpers ────────────────────────────────────────

def parse_date_time(date_str, time_str):
    """Convert MM/DD/YYYY and H:MM AM/PM strings to standardized formats."""
    if not date_str or not time_str:
        return "", ""
    date_str = clean_text(date_str)
    time_str = clean_text(time_str)
    try:
        parts = date_str.split('/')
        if len(parts) == 3:
            month, day, year = parts[0].zfill(2), parts[1].zfill(2), parts[2]
            if len(year) == 2:
                year = f"20{year}"
            date_formatted = f"{year}-{month}-{day}"
        else:
            return "", ""

        m = re.search(r'(\d{1,2}):(\d{2})\s*(AM|PM)', time_str, re.IGNORECASE)
        time_formatted = f"{int(m.group(1))}:{m.group(2)} {m.group(3).upper()}" if m else time_str

        return date_formatted, time_formatted
    except Exception as e:
        print(f"Warning parsing date/time: {date_str} {time_str} - {e}")
        return "", ""


def clean_duration(duration):
    """Return duration as a string of minutes, defaulting to '0' for blanks or dashes."""
    if not duration:
        return "0"
    duration = clean_text(duration)
    if duration == '-' or duration.lower() == 'drop-in':
        return "0"
    return duration


def determine_appointment_type(source):
    """Map source value to appointment type label."""
    return 'Drop-in' if source.upper() in ('DROP-IN',) else 'Appointment'


def determine_modality(reason_text):
    """Infer In-Person vs Online from keywords in the reason/comment field."""
    if not reason_text:
        return "In-Person"
    lower = str(reason_text).lower()
    if any(kw in lower for kw in ('zoom', 'online', 'virtual')):
        return "Online"
    return "In-Person"


# ── Term and intersession helpers ──────────────────────────────────────────

def get_intersession(date_str):
    """
    Return an intersession label if the date falls within a January intersession
    period, otherwise return an empty string.
    Update the date ranges below to match your institution's academic calendar.
    """
    if not date_str:
        return ""
    try:
        dt    = datetime.strptime(date_str, "%Y-%m-%d")
        year  = dt.year
        month = dt.month
        day   = dt.day

        intersession_windows = {
            2020: (2, 29),
            2021: (4, 30),
            2022: (3, 29),
            2023: (3, 28),
            2024: (2, 27),
            2025: (6, 31),
        }

        if month == 1 and year in intersession_windows:
            start, end = intersession_windows[year]
            if start <= day <= end:
                return f"Intersession {year}"

        return ""
    except:
        return ""


def get_term_codes(date_str):
    """
    Convert a YYYY-MM-DD date string to a term code and term label.
    Term digit convention: 3 = Spring, 5 = Summer, 7 = Fall.
    Update the month boundaries below to match your institution's calendar.
    """
    if not date_str:
        return "", ""
    try:
        dt    = datetime.strptime(date_str, "%Y-%m-%d")
        year  = dt.year
        month = dt.month
        day   = dt.day

        if 1 <= month <= 5:
            term_digit, term_name = 3, "Spring"
        elif month in (6, 7):
            term_digit, term_name = 5, "Summer"
        elif month == 8:
            term_digit, term_name = (5, "Summer") if day <= 14 else (7, "Fall")
        elif 9 <= month <= 12:
            term_digit, term_name = 7, "Fall"
        else:
            term_digit, term_name = 0, "Unknown"

        term_code   = f"2{str(year)[-2:]}{term_digit}"
        term_recode = f"{term_name} {year}"
        return term_code, term_recode
    except:
        return "", ""


# ── Reason code parsing ────────────────────────────────────────────────────

def parse_reasons_into_columns(reason_series):
    """
    Scan each reason string for known reason codes and store each match
    in its own column (Reason_1 through Reason_20).

    Matching uses three patterns per code:
      1. Code surrounded by whitespace
      2. Code followed by a dash separator (common SARS format: 'CC - text')
      3. Code at the end of the string
    """
    results = []

    for reason in reason_series:
        reason_str    = str(reason).upper() if pd.notna(reason) else ""
        found_reasons = []

        for code in REASON_CODES:
            code_upper = code.upper()
            patterns = [
                r'(?:^|\s)' + re.escape(code_upper) + r'(?:\s|$)',
                r'(?:^|\s)' + re.escape(code_upper) + r'\s*-',
                r'(?:^|\s)' + re.escape(code_upper) + r'$',
            ]
            for pattern in patterns:
                if re.search(pattern, reason_str):
                    found_reasons.append(code)
                    break

        row = {'Total_Reasons': len(found_reasons)}
        for i, val in enumerate(found_reasons[:20], 1):
            row[f'Reason_{i}'] = val
        for i in range(len(found_reasons) + 1, 21):
            row[f'Reason_{i}'] = ""
        results.append(row)

    return pd.DataFrame(results)


# ── DataFrame assembly ─────────────────────────────────────────────────────

def create_final_dataframe(data):
    """Assemble the final output DataFrame from the list of raw record dicts."""
    if not data:
        return pd.DataFrame()

    processed_rows = []
    for record in data:
        date_fmt, time_fmt = parse_date_time(record.get('Date', ''), record.get('Time', ''))
        processed_rows.append({
            'Source':                    record.get('Source', ''),
            'Appointment Type (Source)': determine_appointment_type(record.get('Source', '')),
            'Attendance Status':         record.get('Attendance', ''),
            'Counselor':                 record.get('Counselor', ''),
            'Student ID':                record.get('Student ID', ''),
            'Student Name':              record.get('Student Name', ''),
            'Date':                      date_fmt,
            'Time':                      time_fmt,
            'Duration':                  clean_duration(record.get('Duration', '')),
            'Reason Code(s) / Comments': record.get('Reason', ''),
            'Page':                      record.get('Page', ''),
        })

    df = pd.DataFrame(processed_rows)

    df['INTERSESSION'] = df['Date'].apply(get_intersession)
    df['Term'], df['Term Recode'] = zip(*df['Date'].apply(get_term_codes))
    df['Modality'] = df['Reason Code(s) / Comments'].apply(determine_modality)

    reason_df = parse_reasons_into_columns(df['Reason Code(s) / Comments'])
    for col in reason_df.columns:
        df[col] = reason_df[col]

    final_columns = [
        'Source', 'Appointment Type (Source)', 'Attendance Status', 'INTERSESSION',
        'Counselor', 'Student ID', 'Student Name', 'Date', 'Term', 'Term Recode',
        'Time', 'Duration', 'Total_Reasons', 'Modality', 'Reason Code(s) / Comments',
        *[f'Reason_{i}' for i in range(1, 21)]
    ]

    for col in final_columns:
        if col not in df.columns:
            df[col] = ""

    return df[[c for c in final_columns if c in df.columns]]


# ── Entry point ────────────────────────────────────────────────────────────

def main():
    print("="*80)
    print("COUNSELING APPOINTMENT DATA EXTRACTOR — SARS PDF")
    print("="*80)

    # Point these two variables to your PDF file.
    # The output Excel will be saved in the same folder.
    pdf_directory = r"C:\path\to\your\folder"   # <-- update this path
    pdf_filename  = "your_sars_report.pdf"       # <-- update this filename
    pdf_path      = os.path.join(pdf_directory, pdf_filename)

    if not os.path.exists(pdf_path):
        print(f"\nERROR: File not found:\n  {pdf_path}")
        print("\nPDF files in that directory:")
        try:
            for f in os.listdir(pdf_directory):
                if f.lower().endswith('.pdf'):
                    print(f"  - {f}")
        except Exception as e:
            print(f"  Could not list directory: {e}")
        input("\nPress Enter to exit...")
        return

    print("\nExtracting data from PDF...")
    raw_data = extract_pdf_table_data(pdf_path)

    if not raw_data:
        print("No data extracted.")
        input("\nPress Enter to exit...")
        return

    print(f"\nSuccessfully extracted {len(raw_data):,} appointment records")
    print("\nProcessing data...")
    df = create_final_dataframe(raw_data)

    if df.empty:
        print("No data to save.")
        input("\nPress Enter to exit...")
        return

    output_file = os.path.join(pdf_directory, "SARS_Cleaned_Data.xlsx")
    try:
        df.to_excel(output_file, index=False)
        print(f"\nSUCCESS — file saved to:\n  {output_file}")
        print(f"Total records: {len(df):,}")

        print("\n" + "="*80)
        print("DATA SUMMARY")
        print("="*80)

        for label, series in [
            ("Source",            df['Source']),
            ("Counselor",         df['Counselor']),
            ("Attendance Status", df['Attendance Status']),
            ("INTERSESSION",      df['INTERSESSION']),
            ("Modality",          df['Modality']),
        ]:
            print(f"\n{label}:")
            for val, cnt in series.value_counts().items():
                print(f"  '{val}': {cnt:,}")

        if 'Total_Reasons' in df.columns:
            print("\nRecords by reason count:")
            print(df['Total_Reasons'].value_counts().sort_index().to_string())

            reason_counts = {}
            for i in range(1, 21):
                col = f'Reason_{i}'
                if col in df.columns:
                    for val, cnt in df[col].value_counts().items():
                        if val:
                            reason_counts[val] = reason_counts.get(val, 0) + cnt
            print("\nTop 10 reason codes:")
            for code, cnt in sorted(reason_counts.items(), key=lambda x: x[1], reverse=True)[:10]:
                print(f"  {code}: {cnt:,}")

        print("\n" + "="*80)

    except Exception as e:
        print(f"\nError saving Excel file: {e}")

    input("\nProcessing complete. Press Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + "="*80)
        print("UNEXPECTED ERROR:")
        print(f"  {type(e).__name__}: {e}")
        print("\nCommon fixes:")
        print("  - Install missing libraries:")
        print("      py -m pip install pandas pdfplumber openpyxl")
        print("  - Check that pdf_directory and pdf_filename are correct in main()")
        print("="*80)
        input("\nPress Enter to close...")
