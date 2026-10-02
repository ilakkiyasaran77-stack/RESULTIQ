"""
services/reporter.py
====================
Module 6 — Pure Python report engine for ResultIQ.

No Flask or SQLAlchemy imports.  Can be unit-tested in isolation.

Public API
----------
generate_report(source_filepath, worksheets_json, sheet_name, output_filepath)
    -> dict  (summary statistics)

The function:
  1. Reads the specified sheet from the source Excel file.
  2. Identifies MARKS columns from the pre-classified worksheets JSON.
  3. Coerces each mark to numeric safely (absent/invalid → 0 with a note).
  4. Computes per-student: Total, Average, Percentage, Result (PASS/FAIL).
  5. Writes a professionally styled .xlsx to output_filepath.
  6. Returns a summary stats dict.
"""

from __future__ import annotations

import re
import pandas as pd
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ #
# Constants
# ------------------------------------------------------------------ #

# Absent / non-numeric tokens (mirrors analyser.py — keep in sync)
ABSENT_TOKENS = frozenset({
    'aa', 'ab', 'absent', 'na', 'n/a', 'n.a', 'n.a.', '-', '--', '---',
    'nil', 'null', 'none', 'not appeared', 'withheld', 'w/h', 'wh',
    'medical', 'detained', 'ex', 'exempted',
})

# Pass/Fail thresholds
PASS_OVERALL_PERCENT = 40.0   # minimum overall percentage to pass
PASS_SUBJECT_MIN    = 50.0   # minimum per-subject percentage to pass

# Assumed maximum marks per subject when not explicitly known
DEFAULT_MAX_PER_SUBJECT = 100


# ------------------------------------------------------------------ #
# Helper — safe numeric coercion
# ------------------------------------------------------------------ #

def _to_numeric(value) -> float | None:
    """
    Convert a cell value to float.

    Returns
    -------
    float  if the value is numeric.
    None   if the value is NaN, blank, or an absent token.
    Raises ValueError for truly unexpected non-numeric text.
    """
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.lower() in ABSENT_TOKENS:
        return None
    # Strip commas (e.g. "1,000")
    s = s.replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


# ------------------------------------------------------------------ #
# Core computation
# ------------------------------------------------------------------ #

def _compute_results(df: pd.DataFrame, marks_cols: list[str]) -> pd.DataFrame:
    """
    Add result columns to *df* (mutates a copy).

    Columns added
    -------------
    <col>_numeric  — coerced float (NaN = absent/invalid)  [internal, dropped]
    Total Marks    — sum of valid marks (absent → 0)
    Subjects Appeared — count of subjects with a valid mark
    Average        — mean of valid marks (0 if none)
    Percentage     — (Total / max_possible) × 100
    Result         — PASS / FAIL
    """
    df = df.copy()
    n_subjects = len(marks_cols)
    max_possible = n_subjects * DEFAULT_MAX_PER_SUBJECT

    numeric_cols = []
    for col in marks_cols:
        num_col = f'__num_{col}'
        df[num_col] = df[col].apply(_to_numeric)
        numeric_cols.append(num_col)

    # Total: treat absent/missing as 0
    df['Total Marks'] = df[numeric_cols].fillna(0).sum(axis=1)

    # Subjects appeared = count of non-None entries
    df['Subjects Appeared'] = df[numeric_cols].notna().sum(axis=1)

    # Average over appeared subjects only
    df['Average'] = df.apply(
        lambda row: (
            row['Total Marks'] / row['Subjects Appeared']
            if row['Subjects Appeared'] > 0 else 0.0
        ),
        axis=1
    ).round(2)

    # Percentage based on full possible marks (not just appeared subjects)
    df['Percentage'] = (
        (df['Total Marks'] / max_possible * 100) if max_possible > 0 else 0.0
    ).round(2)

    # PASS/FAIL logic
    # A student passes if:
    #   1. Overall percentage >= PASS_OVERALL_PERCENT
    #   2. Each appeared subject has >= PASS_SUBJECT_MIN %
    def _result(row) -> str:
        if row['Percentage'] < PASS_OVERALL_PERCENT:
            return 'FAIL'
        # Check per-subject threshold
        for num_col in numeric_cols:
            val = row[num_col]
            if val is not None and not pd.isna(val):
                subject_pct = (val / DEFAULT_MAX_PER_SUBJECT) * 100
                if subject_pct < PASS_SUBJECT_MIN:
                    return 'FAIL'
        return 'PASS'

    df['Result'] = df.apply(_result, axis=1)

    # Drop internal numeric columns
    df.drop(columns=numeric_cols, inplace=True)

    return df


# ------------------------------------------------------------------ #
# Excel styling helpers
# ------------------------------------------------------------------ #

_HEADER_FILL   = PatternFill('solid', fgColor='1E2A4A')   # dark navy
_HEADER_FONT   = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
_BODY_FONT     = Font(name='Calibri', size=11)
_BOLD_FONT     = Font(name='Calibri', bold=True, size=11)
_CENTER        = Alignment(horizontal='center', vertical='center', wrap_text=True)
_LEFT          = Alignment(horizontal='left',   vertical='center', wrap_text=True)

def _thin_border():
    thin = Side(style='thin', color='D1D5DB')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _thick_border():
    """Thick border matching the reference output workbook analysis block."""
    thick = Side(style='thick')
    return Border(left=thick, right=thick, top=thick, bottom=thick)


FOOTER_REGEX = re.compile(
    r'(faculty|incharge|in-charge|h\.?o\.?d|head of|signature|staff|principal|dean|'
    r'prepared by|verified by|advisor|coordinator|controller|result analysis|total|average|passed|appeared)',
    re.IGNORECASE
)


def _is_valid_mark_or_absent(value) -> bool:
    """Return True if value is numeric, NaN/None/blank, or a known absent token."""
    if value is None or pd.isna(value):
        return True
    s = str(value).strip()
    if not s or s.lower() in ABSENT_TOKENS:
        return True
    try:
        float(s.replace(',', ''))
        return True
    except ValueError:
        return False


def _is_footer_row(row_values: list) -> bool:
    """Return True if row text matches signature, footer, or metadata patterns."""
    non_empty = [str(v).strip() for v in row_values if v is not None and str(v).strip() != '']
    if not non_empty:
        return False
    row_text = ' '.join(non_empty)
    return bool(FOOTER_REGEX.search(row_text))


def _append_per_column_analysis(ws, per_column_stats):
    """Write a horizontal Result Analysis block below the student data.

    Key rules (derived from reference file inspection):
    - Analysis starts IMMEDIATELY after the last student row (no blank gap).
    - If the target rows are already blank (like the 15 empty rows in the
      reference file before the footer), we write into them to preserve footer position.
    - If the target rows contain data (e.g., footer immediately after students),
      we use ws.insert_rows() to push that data downward safely.
    - Borders on analysis cells are THICK, matching the student data rows.
    - Borders are applied across ALL columns (A through max_col).
    - n_data_rows = last_data_row - first_data_row + 1 (= total student count).
    """
    if not per_column_stats:
        return

    # -- 1. Locate header row and map marks column names -> ws col indices --
    header_row = 1
    col_map = {}
    name_col_idx = None
    identity_col_indices = []

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        found = False
        for c_idx, val in enumerate(row, start=1):
            if isinstance(val, str):
                stripped = val.strip()
                if stripped in per_column_stats:
                    col_map[stripped] = c_idx
                    found = True
                if stripped.lower() in ('name of the student', 'name', 'student name'):
                    name_col_idx = c_idx
                if stripped.lower() in ('sl. no', 'sl no', 's.no', 'roll no', 'regiter number', 'register number', 'reg no'):
                    identity_col_indices.append(c_idx)
        if found:
            header_row = row_idx
            break

    if not col_map:
        return

    marks_col_indices = list(col_map.values())
    if name_col_idx is None:
        first_marks_col = min(marks_col_indices)
        name_col_idx = max(1, first_marks_col - 1)

    # -- 2. Dynamically find student data rows (first_data_row .. last_data_row) --
    first_data_row = header_row + 1
    last_data_row = header_row

    for r in range(first_data_row, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in row_vals if v is not None and str(v).strip() != '']
        if not non_empty:
            continue

        # Check if row matches footer/signature patterns
        if _is_footer_row(row_vals):
            break

        # Check if marks columns have invalid non-mark text (e.g., 'HOD/IT')
        has_invalid_marks_text = any(
            not _is_valid_mark_or_absent(ws.cell(row=r, column=c).value)
            for c in marks_col_indices
        )
        if has_invalid_marks_text:
            break

        # Check if row has student data (identity or marks)
        has_data = any(
            ws.cell(row=r, column=c).value is not None and str(ws.cell(row=r, column=c).value).strip() != ''
            for c in (marks_col_indices + [name_col_idx] + identity_col_indices)
        )
        if has_data:
            last_data_row = r

    if last_data_row < first_data_row:
        # Fallback if no student data found
        last_data_row = first_data_row

    n_data_rows = last_data_row - first_data_row + 1
    start_row = last_data_row + 1

    # -- 3. Row labels --
    LABEL_COL = name_col_idx
    row_labels = [
        'appeared',
        'passed ',
        '% passed',
        'below 50%',
        '50%-60%',
        '61%-70%',
        '71%-80%',
        '81%-90%',
        '91%-100%',
    ]
    num_rows = len(row_labels)

    # -- Smart Insertion Logic --
    # Check if we have enough completely blank rows below start_row
    # If any cell in the next `num_rows` rows has a value, we must insert rows.
    needs_insert = False
    for r in range(start_row, start_row + num_rows):
        if r > ws.max_row:
            break
        # Check if row r has any data
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value is not None and str(ws.cell(row=r, column=c).value).strip() != '':
                needs_insert = True
                break
        if needs_insert:
            break

    if needs_insert:
        ws.insert_rows(start_row, amount=num_rows)

    border = _thick_border()
    max_col = ws.max_column

    # -- 4. Write labels and borders --
    for i, label in enumerate(row_labels):
        current_row = start_row + i
        for c in range(1, max_col + 1):
            cell = ws.cell(row=current_row, column=c)
            cell.border = border
            cell.font = _BODY_FONT
            cell.alignment = _CENTER
        lc = ws.cell(row=current_row, column=LABEL_COL, value=label)
        lc.border = border
        lc.font = _BODY_FONT
        lc.alignment = _LEFT

    # -- 5. Write Excel formulas --
    for col_name in per_column_stats:
        if col_name not in col_map:
            continue
        c_idx = col_map[col_name]
        col_letter = get_column_letter(c_idx)
        data_start = first_data_row
        rng = f"{col_letter}{data_start}:{col_letter}{last_data_row}"

        appeared_row = start_row
        passed_row   = start_row + 1

        col_values = [
            f'={n_data_rows}-COUNTIF({rng},"aa")',
            f'=COUNTIF({rng},">=50")',
            f'={col_letter}{passed_row}/{col_letter}{appeared_row}*100',
            f'=COUNTIF({rng},"<50")',
            f'=COUNTIFS({rng},">=50",{rng},"<=60")',
            f'=COUNTIFS({rng},">=61",{rng},"<=70")',
            f'=COUNTIFS({rng},">=71",{rng},"<=80")',
            f'=COUNTIFS({rng},">=81",{rng},"<=90")',
            f'=COUNTIFS({rng},">=91",{rng},"<=100")',
        ]

        for i, value in enumerate(col_values):
            cell = ws.cell(row=start_row + i, column=c_idx, value=value)
            cell.border = border
            cell.font = _BODY_FONT
            cell.alignment = _CENTER



# ------------------------------------------------------------------ #
# Summary statistics
# ------------------------------------------------------------------ #

def _compute_results_per_column(df: pd.DataFrame, marks_cols: list[str]) -> dict:
    """Compute statistics for each marks column individually."""
    stats_by_col = {}
    for col in marks_cols:
        num_col = df[col].apply(_to_numeric)
        
        appeared_mask = num_col.notna()
        appeared = int(appeared_mask.sum())
        
        passes = 0
        fails = 0
        bins = {
            'below 50%': 0, '50%-60%': 0, '61%-70%': 0, '71%-80%': 0, '81%-90%': 0, '91%-100%': 0
        }
        
        for val in num_col:
            if pd.isna(val):
                continue
            pct = (val / DEFAULT_MAX_PER_SUBJECT) * 100
            if pct >= PASS_SUBJECT_MIN:
                passes += 1
            else:
                fails += 1
                
            if pct < 50:
                bins['below 50%'] += 1
            elif pct <= 60:
                bins['50%-60%'] += 1
            elif pct <= 70:
                bins['61%-70%'] += 1
            elif pct <= 80:
                bins['71%-80%'] += 1
            elif pct <= 90:
                bins['81%-90%'] += 1
            else:
                bins['91%-100%'] += 1
                
        pass_rate = round((passes / appeared * 100), 2) if appeared > 0 else 0.0
        
        stats_by_col[col] = {
            'appeared': appeared,
            'pass_count': passes,
            'fail_count': fails,
            'pass_rate': pass_rate,
            'bins': bins
        }
    return stats_by_col


def _build_stats(df: pd.DataFrame, marks_cols: list[str]) -> dict:
    """Return a summary statistics dict from the computed results DataFrame."""
    total   = len(df)
    # Appeared = students who have at least one valid mark
    appeared = int((df['Subjects Appeared'] > 0).sum())
    passes  = int((df['Result'] == 'PASS').sum())
    fails   = int((df['Result'] == 'FAIL').sum())
    
    # Calculate percentage based on appeared population or total?
    # Usually pass % is based on Appeared.
    pass_rate = round((passes / appeared * 100), 2) if appeared > 0 else 0.0
    
    avg_pct = round(float(df['Percentage'].mean()), 2) if total > 0 else 0.0
    highest = round(float(df['Percentage'].max()), 2)  if total > 0 else 0.0
    lowest  = round(float(df['Percentage'].min()), 2)  if total > 0 else 0.0
    
    bins = {
        'below 50%': 0, '50%-60%': 0, '61%-70%': 0, '71%-80%': 0, '81%-90%': 0, '91%-100%': 0
    }
    
    for pct in df['Percentage']:
        if pd.isna(pct):
            continue
        if pct < 50:
            bins['below 50%'] += 1
        elif pct <= 60:
            bins['50%-60%'] += 1
        elif pct <= 70:
            bins['61%-70%'] += 1
        elif pct <= 80:
            bins['71%-80%'] += 1
        elif pct <= 90:
            bins['81%-90%'] += 1
        else:
            bins['91%-100%'] += 1

    per_column_stats = _compute_results_per_column(df, marks_cols)

    return {
        'total_students':  total,
        'appeared':        appeared,
        'pass_count':      passes,
        'fail_count':      fails,
        'pass_rate':       pass_rate,
        'avg_percentage':  avg_pct,
        'highest_pct':     highest,
        'lowest_pct':      lowest,
        'marks_columns':   marks_cols,
        'num_subjects':    len(marks_cols),
        'bins':            bins,
        'per_column_stats': per_column_stats,
    }



def _filter_student_dataframe(df: pd.DataFrame, marks_cols: list[str]) -> pd.DataFrame:
    """Filter raw DataFrame to keep only valid student data rows."""
    valid_row_indices = []
    for idx, row in df.iterrows():
        # Check if entire row is null/empty
        row_non_null = row.dropna()
        if row_non_null.empty:
            continue

        # Check if row is a footer/signature row
        row_vals = [v for v in row.values if pd.notna(v) and str(v).strip() != '']
        if not row_vals or _is_footer_row(row_vals):
            break

        # Check if any marks column has invalid non-mark text
        has_invalid_mark = False
        for col in marks_cols:
            if col in row and not _is_valid_mark_or_absent(row[col]):
                has_invalid_mark = True
                break
        if has_invalid_mark:
            break

        # Check if at least one marks cell or identity cell is non-null
        has_any_data = any(pd.notna(row[col]) and str(row[col]).strip() != '' for col in marks_cols if col in row)
        if not has_any_data:
            if not any(pd.notna(v) and str(v).strip() != '' for v in row.values):
                continue

        valid_row_indices.append(idx)

    if not valid_row_indices:
        return df.iloc[0:0].copy()
    return df.loc[valid_row_indices].copy()


# ------------------------------------------------------------------ #
# Public API
# ------------------------------------------------------------------ #

def generate_report(
    source_filepath:  str,
    worksheets_json:  list[dict],
    sheet_name:       str,
    output_filepath:  str,
) -> dict:
    # ---- 1. Find marks columns from pre-classified JSON ----
    target_sheet = None
    for s in worksheets_json:
        if s.get('name') == sheet_name:
            target_sheet = s
            break

    if target_sheet is None:
        raise ValueError(f"Sheet '{sheet_name}' not found in analysis data.")

    marks_cols = [
        col['name']
        for col in target_sheet.get('columns', [])
        if col.get('classification') == 'marks'
    ]

    if not marks_cols:
        raise ValueError(
            f"No columns classified as MARKS were found in sheet '{sheet_name}'. "
            "Please run the analysis step (Analyse Columns) before generating a report."
        )

    # ---- 2. Read source data using the detected header row ----
    ext = source_filepath.rsplit('.', 1)[-1].lower()
    engine = 'openpyxl' if ext == 'xlsx' else 'xlrd'

    header_row_idx = target_sheet.get('header_row_index', None)
    if header_row_idx is None:
        from services.analyser import _detect_header_row
        header_row_idx = _detect_header_row(source_filepath, sheet_name)

    try:
        df = pd.read_excel(source_filepath, sheet_name=sheet_name,
                           header=header_row_idx, engine=engine)
    except Exception as exc:
        raise ValueError(f"Cannot read sheet '{sheet_name}': {exc}") from exc

    if df.empty:
        raise ValueError(f"Sheet '{sheet_name}' contains no data.")

    df_cols = [str(c).strip() for c in df.columns]
    df.columns = df_cols

    valid_marks = [c for c in marks_cols if c in df_cols]
    if not valid_marks:
        raise ValueError(
            "MARKS columns from analysis were not found in the source file. "
            "Please re-analyse the file before generating a report."
        )

    # ---- 3. Filter DF to true student rows & compute result stats ----
    df_students = _filter_student_dataframe(df, valid_marks)
    if df_students.empty:
        df_students = df.copy()

    df_result = _compute_results(df_students, valid_marks)
    stats = _build_stats(df_result, valid_marks)

    # ---- 4. Clone workbook and append analysis ----
    try:
        wb = openpyxl.load_workbook(source_filepath)
        ws = wb[sheet_name]

        _append_per_column_analysis(ws, stats['per_column_stats'])

        # ---- Force Excel to recalculate all formulas on open ----
        # openpyxl never evaluates formulas; cached values remain None.
        # Setting calcMode="auto" and fullCalcOnLoad=True writes a <calcPr>
        # element that instructs Excel to perform a full recalculation the
        # moment the workbook is opened, so every formula cell shows its
        # computed value without the user pressing F9.
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True

        wb.save(output_filepath)
    except Exception as exc:
        raise ValueError(f"Failed to clone and modify workbook: {exc}") from exc
    finally:
        if 'wb' in locals():
            wb.close()

    return stats

