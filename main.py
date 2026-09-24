import pandas as pd
import numpy as np
from io import BytesIO

# ============================================================================
# CONFIGURATION -- adjust these to control comparison behavior
# ============================================================================

# If True, leading/trailing whitespace in DATA VALUES (not column names) is
# stripped before comparison, meaning "foo" and "foo " would be treated as
# the SAME value. If False (default), whitespace differences in data values
# are preserved as real differences, matching the same "be exact" philosophy
# requested for column names.
# Comparison is deliberately insensitive to incidental outer whitespace; it
# does not alter internal whitespace ("New York" remains different from
# "NewYork").
STRIP_WHITESPACE_IN_DATA_VALUES = True


# Sentinel used internally to represent "missing" (NaN/None/NaT) so that
# groupby treats all missing values in a column as equal to each other,
# without ever colliding with a real data value. Uses characters extremely
# unlikely to appear in real spreadsheet data.
_MISSING_SENTINEL = "<<<__MISSING__>>>"

# Columns whose values look like clock times (e.g. "07:00" text, or
# datetime.time objects from Excel) will be normalized to a fixed-format
# string "HH:MM:SS" before comparison, so "07:00" and time(7, 0, 0) compare
# equal. Detection is automatic (see _looks_like_time_series below), but you
# can force specific columns here if auto-detection misses them:
FORCE_TIME_COLUMNS = []  # e.g. ["Start Time", "Doors Closed"]

# Columns whose values are dates/datetimes will be converted to "MM-DD-YYYY"
# text (time-of-day dropped entirely) before comparison AND before the
# source file is saved/archived. Detection is automatic (see
# _looks_like_date_series below), but you can force specific columns here if
# auto-detection misses them:
FORCE_DATE_COLUMNS = []  # e.g. ["Flight Date", "Booking Date"]


# ============================================================================
# RESULT HIGHLIGHTING CONFIGURATION (NEW -- output formatting only)
# ============================================================================
# These control the cell fill colors applied to the GENERATED RESULT
# workbook after it is written. Nothing here affects comparison logic in any
# way -- the comparison runs exactly as before, and this is purely a
# post-processing pass over the finished output file so the differences that
# were already identified are visually obvious instead of having to be
# hunted for by eye.
#
# All colors are deliberately LIGHT/pastel so black cell text stays readable
# and printing/screenshotting the report still looks clean. Values are plain
# 6-digit RRGGBB hex strings (openpyxl format, no leading '#').
HIGHLIGHT_RESULTS = True                 # set False to write a plain, unhighlighted workbook

HIGHLIGHT_CHANGED_VALUE_COLOR = "FFF2CC"  # light amber  -- a value that differs between A and B
HIGHLIGHT_EXTRA_IN_A_COLOR    = "FCE4E4"  # light red    -- rows/columns present (or excess) in A only
HIGHLIGHT_EXTRA_IN_B_COLOR    = "DDEBF7"  # light blue   -- rows/columns present (or excess) in B only
HIGHLIGHT_SUMMARY_COLOR       = "FFF2CC"  # light amber  -- non-zero difference metrics on Summary

# Autofill changes source data before comparison; they are not a separate
# comparison result and therefore must not receive their own report colour.
HIGHLIGHT_AUTOFILLED_VALUES = False
HIGHLIGHT_AUTOFILLED_COLOR = "E2F0D9"


# ============================================================================
# STEP 0: LOAD
# ============================================================================

def expand_powerbi_merged_cells(path, sheet_name=0):
    """Expand merged cells in one worksheet and save the workbook.

    Excel stores a value only in the top-left cell of a merged range.  The
    range is unmerged first because openpyxl's other cells in a merged range
    are read-only, then that top-left value is written to every former member
    of the range.  This deliberately affects only ranges Excel explicitly
    marks as merged; ordinary blank cells are left untouched.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[sheet_name] if isinstance(sheet_name, int) else wb[sheet_name]

    # Copy because unmerge_cells changes ws.merged_cells.ranges while iterating.
    for merged_range in list(ws.merged_cells.ranges):
        min_col, min_row, max_col, max_row = merged_range.bounds
        value = ws.cell(row=min_row, column=min_col).value
        ws.unmerge_cells(str(merged_range))
        for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                min_col=min_col, max_col=max_col):
            for cell in row:
                cell.value = value

    wb.save(path)


def load_sheet(path, sheet_name=0, expand_powerbi_merges=False):
    """
    Load a single sheet from an .xlsx file, preserving column headers
    EXACTLY as written (no auto-stripping of whitespace), so that a header
    difference like "Doors Closed" vs " Doors Closed" is detectable rather
    than silently normalized away.

    Parameters
    ----------
    path : str
        Path to the .xlsx file.
    sheet_name : str or int
        Sheet name or index to read (default: first sheet).

    Also detects which cells are PERCENTAGE-FORMATTED in the source
    workbook. Excel stores a percentage as a plain fraction -- e.g. 0.63%
    is stored internally as the number 0.0063 -- with "%" applied purely
    as a display format (cell.number_format, e.g. "0.00%") that a plain
    value read never exposes. Without tracking this separately, a
    percent-formatted numeric cell and a sheet that stores the same
    percentage as literal text ("0.63%") would normalize to two
    completely different-looking values and always appear to mismatch.
    See convert_percent_columns_to_text() for how this mask is used.

    Returns
    -------
    (pandas.DataFrame, pandas.DataFrame)
        The raw sheet data (headers untouched, values read as openpyxl
        naturally infers them -- dtype normalization happens later,
        explicitly, in compare_data()), and a same-shape, same-column-
        names boolean DataFrame marking which cells were percentage-
        formatted in the source workbook.
    """
    # Read cell-by-cell via openpyxl (rather than pd.read_excel) so each
    # cell's number_format can be captured alongside its value -- pandas'
    # own Excel reader discards that formatting information entirely.
    import openpyxl

    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[sheet_name] if isinstance(sheet_name, int) else wb[sheet_name]

    if expand_powerbi_merges:
        # Expand merges in this in-memory workbook. The comparison
        # representation needs every member of an Excel merged range to
        # have its logical value AND its number_format -- otherwise a
        # non-anchor cell that's now holding a copied percent value could
        # still show "General" format and be missed by the percent-mask
        # detection below.
        for merged_range in list(ws.merged_cells.ranges):
            min_col, min_row, max_col, max_row = merged_range.bounds
            top_left = ws.cell(row=min_row, column=min_col)
            value, number_format = top_left.value, top_left.number_format
            ws.unmerge_cells(str(merged_range))
            for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                    min_col=min_col, max_col=max_col):
                for cell in row:
                    cell.value = value
                    cell.number_format = number_format

    # Header row: used exactly as found, no whitespace-stripping or other
    # transformation, so a header difference like "Doors Closed" vs
    # " Doors Closed" stays detectable rather than silently normalized away.
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    data_rows, percent_rows = [], []
    last_row = ws.max_row or 1
    if last_row >= 2:
        for row in ws.iter_rows(min_row=2, max_row=last_row):
            values, percents = [], []
            for cell in row:
                values.append(cell.value)
                percents.append(cell.number_format is not None and "%" in cell.number_format)
            data_rows.append(values)
            percent_rows.append(percents)

    df = pd.DataFrame(data_rows, columns=header)
    percent_mask = pd.DataFrame(percent_rows, columns=header)
    return df, percent_mask


# ============================================================================
# MERGED-CELL FORWARD-FILL (for exports where merged cells leave blanks)
# ============================================================================
# Some exports (notably Tableau, when the source had merged cells) only
# populate a value on the FIRST row of a repeated block and leave every row
# below it blank -- rather than repeating the value on every row, the way a
# flat/PowerBI-style export usually does. If left as-is, this makes every
# row after the first in each block look like it "differs" from the other
# sheet, even though the data is identical -- it's just structured
# differently.
#
# List the column names here that you KNOW have this merged-cell pattern in
# ONE of your two sheets (forward-filling is harmless on the sheet that's
# already fully populated -- there's nothing blank to fill). Left empty by
# default so this is opt-in, not a silent guess: blindly forward-filling
# every column would incorrectly "fix" a column that's legitimately blank.
FORWARD_FILL_COLUMNS = []  # e.g. ["Entry Date", "StaffID", "Full Name"]


# MODIFIED: FORWARD FILL
# NEW: AUTO-FILL TRACKING
def _ffill_one_column_by_position(df, col_pos):
    """
    Forward-fills blank cells in the column at integer position `col_pos`,
    on the assumption that the blanks come from merged cells in the
    original spreadsheet (a value only present on the first row of a
    repeated block). Mutates `df` in place (caller is expected to have
    already copied it) using positional access throughout, so it is safe
    even when multiple columns share the same (possibly blank/None)
    header -- label-based access like df[col_name] silently returns/sets
    EVERY column with that label at once, which is not what "autofill
    this one column" means.

    Returns the set of `(column_name, original_dataframe_index)` pairs for
    precisely the cells that changed from blank to a prior value.
    """
    col_name = df.columns[col_pos]
    values = df.iloc[:, col_pos].copy()
    # ffill only recognizes NaN, while exports often represent blank merged
    # cells as empty/whitespace strings. Convert those to missing first.
    # pandas leaves leading missing values missing, which is the required
    # first-row behavior.
    blank = values.isna() | values.map(lambda v: isinstance(v, str) and not v.strip())
    filled = values.mask(blank).ffill()
    # A leading blank remains missing after ffill and was not auto-filled.
    # Only record cells for which ffill supplied a usable previous value.
    actually_filled = blank & filled.notna()
    autofilled_cells = {(col_name, index) for index in df.index[actually_filled]}
    df.isetitem(col_pos, filled)
    return autofilled_cells


def forward_fill_merged_cells(df, columns):
    """
    Forward-fills blank cells in the given columns (by NAME), on the
    assumption that the blanks come from merged cells in the original
    spreadsheet. Returns a NEW dataframe -- does not mutate the input.
    Also returns a set of ``(column_name, original_dataframe_index)``
    pairs for precisely the cells that changed from blank to a prior
    value.

    NOTE: this name-based entry point is kept for FORWARD_FILL_COLUMNS
    (a short, hand-typed config list). If a name in `columns` happens to
    match more than one column (e.g. duplicate/blank headers), EVERY
    matching column is filled -- that is the documented behavior for this
    name-based path. The interactive prompt (prompt_for_autofill) does
    NOT use this function for that reason; it fills by exact column
    position instead, see forward_fill_merged_cells_by_position below.
    """
    df = df.copy()
    autofilled_cells = set()
    for col in columns:
        if col in df.columns:
            positions = [i for i, name in enumerate(df.columns) if name == col]
            for col_pos in positions:
                autofilled_cells.update(_ffill_one_column_by_position(df, col_pos))
    return df, autofilled_cells


def forward_fill_merged_cells_by_position(df, positions):
    """
    Forward-fills blank cells in the columns at the given integer
    positions (0-based). Returns a NEW dataframe -- does not mutate the
    input -- plus the set of `(column_name, original_dataframe_index)`
    pairs that were actually filled.

    Unlike forward_fill_merged_cells (name-based), this only ever touches
    the exact column(s) the caller pointed at, even if other columns
    happen to share the same (or a blank/None) header label.
    """
    df = df.copy()
    autofilled_cells = set()
    for col_pos in positions:
        autofilled_cells.update(_ffill_one_column_by_position(df, col_pos))
    return df, autofilled_cells


# MODIFIED: AUTO-FILL PROMPT (preserves tracking information)
def prompt_for_autofill(df, sheet_label):
    """Interactively forward-fill only columns explicitly selected by user."""
    while True:
        print(f"\nDo you want to autofill any columns in {sheet_label}?")
        print("1. Yes\n2. No")
        choice = input("Select 1 or 2: ").strip()
        if choice == "2":
            return df, set()
        if choice == "1":
            break
        print("Please enter 1 (Yes) or 2 (No).")

    columns = list(df.columns)
    print(f"Available {sheet_label} columns:")
    for number, column in enumerate(columns, start=1):
        print(f"  {number}. {column}")
    while True:
        selected = input("Select one or more column numbers (comma-separated): ").strip()
        try:
            indices = [int(item.strip()) for item in selected.split(",") if item.strip()]
            if not indices or any(index < 1 or index > len(columns) for index in indices):
                raise ValueError
            # Positional, not name-based: some exports have duplicate or
            # blank (None) headers, and selecting by name would silently
            # fill EVERY column sharing that label instead of just the one
            # the user picked (df[col] returns/sets all matching columns
            # at once when a label isn't unique).
            positions = list(dict.fromkeys(index - 1 for index in indices))
            return forward_fill_merged_cells_by_position(df, positions)
        except ValueError:
            print("Enter one or more valid column numbers, separated by commas.")


# ============================================================================
# DATE COLUMN DETECTION + CONVERSION (MM-DD-YYYY, time dropped entirely)
# ============================================================================
# This section handles MULTIPLE real-world date formats -- numeric
# (little/middle/big-endian, any separator), spelled-out month names, ISO
# 8601 with time/zone, ISO week dates (with an explicit weekday digit),
# ordinal/Julian dates, and RFC 2822 -- rather than assuming one fixed
# layout. It is used both to normalize dates for comparison and to actually
# rewrite the source DataFrame (and, in the folder workflow, the source
# file on disk) so the converted MM-DD-YYYY values are what gets archived.
#
# Deliberately UNSUPPORTED (returns the original value unparsed rather than
# guessing, since none of these can be resolved from the string alone):
#   - 2-digit-year numeric dates ("05-06-07" -- which part is the year, and
#     which century, is genuinely ambiguous)
#   - Bare integers as Unix epoch timestamps (indistinguishable from a
#     numeric ID/invoice/phone number by content alone)
#   - ISO week dates with no weekday digit ("2026-W39" -- which day of that
#     7-day week is meant is not present in the string)

import re
import datetime as _dt
from email.utils import parsedate_to_datetime

_MONTH_WORD_RE = re.compile(
    r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)", re.IGNORECASE
)

# Used only when a date column has ZERO unambiguous day/month evidence in
# its own data (see _infer_dayfirst_for_series below). Change this if your
# organization's default convention for such columns is day-first instead.
DEFAULT_DAYFIRST_WHEN_AMBIGUOUS = False  # False = assume MM/DD (US-style)


def _is_nan(v):
    try:
        return pd.isna(v)
    except (TypeError, ValueError):
        return False


def _parse_flexible_date(v, dayfirst=None):
    """
    Parse a single date-like value using structural rules, falling back to
    an explicit `dayfirst` decision ONLY for the genuinely ambiguous
    4-digit-year numeric case (e.g. "4/1/2024"). Returns a
    pandas.Timestamp, or pd.NaT if it can't be confidently parsed.

    Priority order (most-specific / least-ambiguous first):
      1. Already a real date/datetime object -> pass through untouched.
      2. ISO week date WITH explicit weekday digit (YYYY-Www-D).
      3. Ordinal/Julian date (YYYY/DDD or DDD/YYYY).
      4. RFC 2822 (stdlib email parser -- exact per spec, no guessing).
      5. Contains a spelled-out month name -> unambiguous, hand to pandas
         directly (covers "September 23, 2026", "23 September 2026",
         "Sep 23, 2026", "Wednesday, September 23, 2026", "23-SEP-2026").
      6. ISO 8601 with time/zone component (YYYY-MM-DDTHH:MM:SSZ etc).
      7. Plain numeric date with exactly one 4-digit part -> that part IS
         the year; if it's first, order is fixed (Y-M-D); if it's last,
         the other two need `dayfirst` to disambiguate (covers
         DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, MM-DD-YYYY, DD.MM.YYYY,
         YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD).
      Anything else (2-digit-year numeric dates, bare integers, weekday-
      less ISO weeks, unparseable text) -> pd.NaT, left unparsed.
    """
    if v is None or (isinstance(v, float) and _is_nan(v)):
        return pd.NaT
    if isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)):
        return pd.Timestamp(v)
    if not isinstance(v, str):
        return pd.NaT

    s = v.strip()
    if not s:
        return pd.NaT

    # -- ISO week date, weekday digit REQUIRED: 2026-W39-3 --
    m = re.match(r"^(\d{4})-W(\d{2})-(\d)$", s)
    if m:
        year, week, weekday = int(m[1]), int(m[2]), int(m[3])
        try:
            return pd.Timestamp(_dt.date.fromisocalendar(year, week, weekday))
        except ValueError:
            return pd.NaT
        # NOTE: "YYYY-Www" with no trailing "-D" intentionally falls through
        # to the numeric branch below, where it fails the 3-part check and
        # correctly returns pd.NaT rather than guessing a weekday.

    # -- Ordinal/Julian: YYYY/DDD or DDD/YYYY --
    m = re.match(r"^(\d{4})[/-](\d{1,3})$", s) or re.match(r"^(\d{1,3})[/-](\d{4})$", s)
    if m:
        a, b = m[1], m[2]
        year, doy = (int(a), int(b)) if len(a) == 4 else (int(b), int(a))
        try:
            return pd.Timestamp(_dt.date(year, 1, 1) + _dt.timedelta(days=doy - 1))
        except ValueError:
            return pd.NaT

    # -- RFC 2822 (weekday + comma + numeric offset/named zone at the end) --
    if "," in s and re.search(r"[+-]\d{4}$|GMT$|UTC$", s):
        try:
            return pd.Timestamp(parsedate_to_datetime(s))
        except (TypeError, ValueError):
            pass

    # -- Spelled-out month name -> structurally unambiguous --
    if _MONTH_WORD_RE.search(s):
        return pd.to_datetime(s, errors="coerce")  # dayfirst irrelevant here

    # -- ISO 8601 with time component --
    if "T" in s or re.search(r"\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}", s):
        return pd.to_datetime(s.replace("Z", "+00:00"), errors="coerce")

    # -- Plain numeric date: needs exactly one 4-digit part to be safe --
    parts = re.split(r"[-/.]", s)
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        lengths = [len(p) for p in parts]
        if lengths.count(4) == 1:
            nums = [int(p) for p in parts]
            y_idx = lengths.index(4)
            year = nums[y_idx]
            rest = [n for i, n in enumerate(nums) if i != y_idx]
            if y_idx == 0:
                month, day = rest  # Y-M-D, order fixed by position
            else:
                month, day = (rest[1], rest[0]) if dayfirst else (rest[0], rest[1])
            try:
                return pd.Timestamp(_dt.date(year, month, day))
            except ValueError:
                return pd.NaT
        # all parts <=2 digits (2-digit year) -- unsupported, falls through

    return pd.NaT


def _infer_dayfirst_for_series(series):
    """
    Scan a column's date-like strings for UNAMBIGUOUS evidence of day/month
    order: a numeric date with exactly one 4-digit part, where the other
    two parts include a value >12 (which can only be a day, never a
    month). Returns True (day-first), False (month-first), or None (no
    evidence found -- caller falls back to DEFAULT_DAYFIRST_WHEN_AMBIGUOUS).

    Raises ValueError if the column contains unambiguous evidence for BOTH
    conventions -- that means this one column genuinely mixes DD/MM and
    MM/DD rows and cannot be safely auto-normalized without fixing the
    source data.
    """
    saw_dayfirst = False
    saw_monthfirst = False
    for v in series.dropna():
        if not isinstance(v, str):
            continue  # real date/datetime objects are never ambiguous
        s = v.strip()
        if _MONTH_WORD_RE.search(s) or "T" in s or "W" in s:
            continue  # not a plain D/M/Y numeric date -- no evidence here
        parts = re.split(r"[-/.]", s)
        if len(parts) != 3 or not all(p.isdigit() for p in parts):
            continue
        lengths = [len(p) for p in parts]
        if lengths.count(4) != 1:
            continue  # 2-digit-year case -- skipped, no evidence taken from it
        nums = [int(p) for p in parts]
        y_idx = lengths.index(4)
        if y_idx == 0:
            continue  # Y-M-D -- order is fixed, not evidence for D/M ambiguity
        rest = [n for i, n in enumerate(nums) if i != y_idx]
        first, second = rest
        if first > 12 and second <= 12:
            saw_dayfirst = True
        elif second > 12 and first <= 12:
            saw_monthfirst = True

    if saw_dayfirst and saw_monthfirst:
        raise ValueError(
            "This column mixes DD/MM and MM/DD dates -- cannot safely "
            "auto-detect a single convention. Fix the source data or "
            "split the column before comparing."
        )
    if saw_dayfirst:
        return True
    if saw_monthfirst:
        return False
    return None


def _looks_like_date_series(series):
    """
    Heuristic: does this column hold date/datetime values?
      1. Native date/datetime objects (Timestamp/datetime/date/
         np.datetime64) -- unambiguous, always counted.
      2. datetime.time values (clock times, no date component) are
         excluded -- those belong to the separate time-column pipeline,
         not this one.
      3. Text values that _parse_flexible_date can confidently parse under
         EITHER day-first or month-first assumption (the actual convention
         is resolved separately by _infer_dayfirst_for_series once the
         column is confirmed to be date-like).
    ALL sampled non-null, non-time values must match for the column to be
    treated as a date column, so a column that's mostly free text is left
    alone even if one value happens to be parseable.
    """
    sample = series.dropna().head(20)
    if len(sample) == 0:
        return False
    hits = 0
    checked = 0
    for v in sample:
        if isinstance(v, _dt.time):
            continue  # belongs to the separate time-column pipeline
        checked += 1
        if isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date, np.datetime64)):
            hits += 1
        elif not pd.isna(_parse_flexible_date(v, dayfirst=False)) or \
                not pd.isna(_parse_flexible_date(v, dayfirst=True)):
            hits += 1
    return checked > 0 and hits == checked


def _to_mmddyyyy(v, dayfirst=False):
    """
    Convert a single date/datetime-like value to 'MM-DD-YYYY' text, with
    the time-of-day component dropped entirely. Non-date, unparseable, and
    missing values pass through UNCHANGED (see the module-level docstring
    for exactly which formats are deliberately left unparsed).
    """
    if isinstance(v, _dt.time):
        return v  # not a date value -- leave untouched
    parsed = _parse_flexible_date(v, dayfirst=dayfirst)
    return parsed.strftime("%m-%d-%Y") if not pd.isna(parsed) else v


def convert_date_columns_to_mmddyyyy(df, force_columns=None):
    """
    Detects date-related columns in `df` (dynamically via
    _looks_like_date_series, plus any explicit overrides passed in
    `force_columns` or listed in FORCE_DATE_COLUMNS), infers each column's
    own day/month convention from unambiguous values in its own data (see
    _infer_dayfirst_for_series), and converts values to 'MM-DD-YYYY' text,
    dropping the time-of-day portion completely. Returns a NEW dataframe --
    does not mutate the input.

    Each column's convention is detected independently, so Sheet A and
    Sheet B can each be resolved correctly even if one is DD/MM and the
    other is MM/DD -- callers run this once per sheet.

    Columns that merely contain numbers (flight numbers, IDs, pax counts,
    etc.) are left completely untouched, since detection requires either a
    real date/datetime object or a text value that _parse_flexible_date can
    confidently resolve.
    """
    df = df.copy()
    force = set(force_columns or []) | set(FORCE_DATE_COLUMNS)
    for col in df.columns:
        if col in force or _looks_like_date_series(df[col]):
            dayfirst = _infer_dayfirst_for_series(df[col])
            if dayfirst is None:
                dayfirst = DEFAULT_DAYFIRST_WHEN_AMBIGUOUS
                print(f"[WARN] '{col}': no unambiguous day/month evidence in this "
                      f"column -- defaulting to {'DD/MM' if dayfirst else 'MM/DD'}. "
                      f"Double-check this column if it matters.")
            else:
                print(f"[INFO] '{col}': detected {'DD/MM' if dayfirst else 'MM/DD'} "
                      f"from unambiguous values in the data.")
            df[col] = df[col].map(lambda v: _to_mmddyyyy(v, dayfirst))
    return df


# ============================================================================
# PERCENTAGE COLUMN DETECTION + CONVERSION (canonical "0.63%" text)
# ============================================================================
# Excel stores a percentage-formatted cell as a raw FRACTION (0.63% is
# stored as 0.0063), with "%" applied purely as a display format that a
# plain value read never exposes. A sheet that instead stores the same
# percentage as literal TEXT ("0.63%") has no such hidden fraction -- the
# string already IS the display value. Left alone, these two
# representations of the identical real-world percentage would compare as
# totally different values. This section converts BOTH into the same
# canonical percent-text form so they compare equal.

# Number of decimal places used for the canonical percent text, and for
# re-rounding percent values already stored as text. Matches Excel's
# common "0.00%" display format; change this if your sheets consistently
# use a different precision (e.g. 1 for "0.00%" -> "0.0%", 0 for whole
# percents like "63%").
PERCENT_DECIMAL_PLACES = 2

_PERCENT_TEXT_RE = re.compile(r"^([+-]?\d+(?:\.\d+)?)\s*%$")


def _excel_round(value, decimal_places):
    """
    Round `value` to `decimal_places` using round-half-AWAY-FROM-ZERO --
    the convention Excel itself uses for on-screen display. Python's
    built-in round() uses round-half-TO-EVEN ("banker's rounding"), which
    disagrees with Excel on exact halfway values: round(0.625, 2) is 0.62
    in Python but Excel displays 0.63%. Uses Decimal(str(value)) rather
    than Decimal(value) to avoid binary floating-point representation
    artifacts (e.g. 0.615 sometimes being stored as slightly less than
    0.615 under the hood).
    """
    from decimal import Decimal, ROUND_HALF_UP
    quantum = Decimal(1).scaleb(-decimal_places)  # e.g. Decimal('0.01') for 2 places
    return float(Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP))


def _percent_value_to_text(v, decimal_places=PERCENT_DECIMAL_PLACES):
    """
    Convert a raw Excel percent-formatted FRACTION (e.g. 0.00625, meaning
    0.625%) into canonical percent text (e.g. "0.63%"), rounding to
    `decimal_places` using Excel's own round-half-away-from-zero
    convention (see _excel_round). Non-numeric and missing values pass
    through unchanged.
    """
    if v is None or (isinstance(v, float) and _is_nan(v)):
        return v
    if isinstance(v, bool):
        return v  # guard: bool is a subclass of int in Python
    if isinstance(v, (int, float)):
        return f"{_excel_round(v * 100, decimal_places):.{decimal_places}f}%"
    return v


def _normalize_percent_text(v, decimal_places=PERCENT_DECIMAL_PLACES):
    """
    Round an ALREADY-percent-suffixed text value (e.g. "0.630%", "0.63 %")
    to the same canonical decimal precision used by _percent_value_to_text,
    so a sheet that stores percentages as literal text still compares
    equal to a sheet that stores them as percent-formatted numbers, even
    if the two differ only in trailing-zero style. Values that don't match
    the "<number>%" text pattern pass through unchanged.
    """
    if not isinstance(v, str):
        return v
    m = _PERCENT_TEXT_RE.match(v.strip())
    if not m:
        return v
    number = float(m.group(1))
    return f"{_excel_round(number, decimal_places):.{decimal_places}f}%"


def convert_percent_columns_to_text(df, percent_mask=None):
    """
    Converts every cell flagged in `percent_mask` (a same-shape boolean
    DataFrame produced by load_sheet()'s percent-format detection) from
    its raw Excel fraction (e.g. 0.00625) into canonical percent text
    (e.g. "0.63%"). Also re-rounds any column that ALREADY holds literal
    percent-suffixed text to the same decimal precision, so text-based and
    numeric-format percentages -- whether from the same sheet or two
    different sheets -- compare equal instead of looking like two
    different values for the same real quantity. Returns a NEW dataframe
    -- does not mutate the input.

    If `percent_mask` is None (the CSV loading path, which carries no
    cell-format metadata at all), only the text re-rounding step runs --
    a bare numeric percentage in a CSV (e.g. "0.0063" with no "%" and no
    formatting information) is indistinguishable from an ordinary decimal
    number and is intentionally left untouched.
    """
    df = df.copy()
    for col in df.columns:
        if percent_mask is not None and col in percent_mask.columns:
            mask = percent_mask[col].fillna(False).to_numpy()
            if mask.any():
                # Rebuild the column as a plain Python list rather than a
                # masked in-place assignment: a column that loaded as pure
                # float64 (no text mixed in) locks pandas to that dtype,
                # and assigning percent-text strings into a subset of it
                # raises a LossySetitemError. Replacing the whole column
                # at once lets pandas re-infer dtype (-> object) safely.
                values = df[col].tolist()
                df[col] = [
                    _percent_value_to_text(v) if flagged else v
                    for v, flagged in zip(values, mask)
                ]
        # Re-round any value that's ALREADY percent-suffixed text -- covers
        # CSV-sourced percentages and the just-converted cells above alike.
        df[col] = df[col].map(_normalize_percent_text)
    return df


# ============================================================================
# WRITE MODIFIED DATA BACK TO THE SOURCE FILE (so archives reflect changes)
# ============================================================================

def save_dataframe_to_source(df, path, sheet_name=0):
    """
    Writes `df` back to the ORIGINAL source file at `path`, overwriting its
    data in place. Used so that any preprocessing applied in-memory
    (autofill, date conversion, etc.) is reflected in the actual file
    BEFORE that file gets moved to the archive folder.

    - .csv / .tsv: the file is simply overwritten with the modified
      dataframe.
    - .xlsx / .xls: the existing workbook is loaded so other sheets and the
      overall workbook structure are preserved; only the target sheet's
      cell contents are cleared and rewritten from the dataframe.
    """
    from pathlib import Path as _Path
    path = _Path(path)

    if path.suffix.lower() in (".csv", ".tsv"):
        sep = "\t" if path.suffix.lower() == ".tsv" else ","
        df.to_csv(path, index=False, sep=sep)
        return

    import openpyxl
    wb = openpyxl.load_workbook(path)
    if isinstance(sheet_name, int):
        ws = wb.worksheets[sheet_name]
    else:
        ws = wb[sheet_name]

    # A source sheet may still contain merged ranges because comparison-time
    # expansion is deliberately in-memory.  Remove those definitions before
    # replacing its rows; otherwise openpyxl retains stale merged-cell
    # objects that conflict with the rewritten flat table.
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    # Clear existing rows on the target sheet only, then rewrite header +
    # data from the (possibly modified) dataframe. Other sheets in the
    # workbook are untouched.
    if ws.max_row and ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))

    wb.save(path)


# ============================================================================
# STEP 1: SCHEMA CHECK
# ============================================================================

def compare_schema(df_a, df_b):
    """
    Compare headers after trimming only leading/trailing whitespace.  The
    returned common names preserve Sheet A's original display names.

    Returns
    -------
    dict with keys:
        'only_in_a'   : list of column names present only in Sheet A
        'only_in_b'   : list of column names present only in Sheet B
        'common'      : list of column names present in both (preserves
                         Sheet A's column order, since order is not
                         guaranteed to match and A is our reference)
        'count_a'     : total column count in Sheet A
        'count_b'     : total column count in Sheet B
    """
    cols_a = list(df_a.columns)
    cols_b = list(df_b.columns)

    normalize_header = lambda c: c.strip() if isinstance(c, str) else c
    normalized_a = [normalize_header(c) for c in cols_a]
    normalized_b = [normalize_header(c) for c in cols_b]
    # Ambiguous headers cannot be safely matched after trimming.
    if len(set(normalized_a)) != len(normalized_a) or len(set(normalized_b)) != len(normalized_b):
        raise ValueError("A sheet has duplicate column names after trimming leading/trailing whitespace.")
    set_a, set_b = set(normalized_a), set(normalized_b)
    only_in_a = [c for c, normalized in zip(cols_a, normalized_a) if normalized not in set_b]
    only_in_b = [c for c, normalized in zip(cols_b, normalized_b) if normalized not in set_a]
    common = [c for c, normalized in zip(cols_a, normalized_a) if normalized in set_b]
    b_name_by_normalized = dict(zip(normalized_b, cols_b))
    common_b = [b_name_by_normalized[normalize_header(c)] for c in common]

    return {
        "only_in_a": only_in_a,
        "only_in_b": only_in_b,
        "common": common,
        "common_b": common_b,
        "count_a": len(cols_a),
        "count_b": len(cols_b),
    }


def align_common_headers(df_b, schema):
    """Give B's whitespace-equivalent common headers A's display names.

    This alignment is confined to the in-memory comparison dataframe. Source
    headers remain untouched when files are loaded and archived.
    """
    rename_map = dict(zip(schema["common_b"], schema["common"]))
    return df_b.rename(columns=rename_map)


# ============================================================================
# STEP 2: RECORD COUNT CHECK
# ============================================================================

def compare_row_counts(df_a, df_b):
    """
    Compare total row counts between the two sheets.

    Returns
    -------
    dict with keys: 'rows_a', 'rows_b', 'difference' (rows_a - rows_b)
    """
    rows_a = len(df_a)
    rows_b = len(df_b)
    return {
        "rows_a": rows_a,
        "rows_b": rows_b,
        "difference": rows_a - rows_b,
    }


# ============================================================================
# STEP 3: DATA NORMALIZATION HELPERS (used inside compare_data)
# ============================================================================

def _looks_like_time_series(series):
    """
    Heuristic: does this column look like it holds clock-time values, either
    as datetime.time objects (openpyxl's native representation for Excel
    time-formatted cells) or as "HH:MM" / "HH:MM:SS" text?
    Used only to decide whether to apply time normalization automatically.
    """
    import datetime
    import re

    time_text_re = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")
    sample = series.dropna()
    if len(sample) == 0:
        return False
    # Look at up to 20 non-null values as a representative sample
    sample = sample.head(20)
    hits = 0
    for v in sample:
        if isinstance(v, datetime.time):
            hits += 1
        elif isinstance(v, str) and time_text_re.match(v.strip()):
            hits += 1
    return hits == len(sample)  # ALL sampled values must look like times


def _normalize_time_value(v):
    """Convert a time-like value (datetime.time or 'HH:MM'/'HH:MM:SS' text)
    to a canonical 'HH:MM:SS' string. Non-time or missing values pass through
    unchanged so they can be caught by the general missing-value handling."""
    import datetime

    if v is None or (isinstance(v, float) and pd.isna(v)):
        return v
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, str):
        s = v.strip()
        parts = s.split(":")
        if len(parts) == 2:
            s = s + ":00"
        return s
    return v


def _normalize_date_value(v, dayfirst=False):
    """
    Return MM-DD-YYYY for real datetimes and date-like text (time-of-day
    dropped entirely), using the same multi-format parser as
    convert_date_columns_to_mmddyyyy() (see _parse_flexible_date). Values it
    can't confidently parse pass through unchanged.

    NOTE: by the time this runs inside normalize_dataframe(), date columns
    have typically already been converted to 'MM-DD-YYYY' strings by
    convert_date_columns_to_mmddyyyy() upstream (see run_comparison /
    run_folder_comparison), so `dayfirst` here is just a safety-net default
    for any date-like value that slipped through unconverted -- it does not
    re-decide the convention for columns already normalized upstream.
    """
    if isinstance(v, _dt.time):
        return v
    parsed = _parse_flexible_date(v, dayfirst=dayfirst)
    return parsed.strftime("%m-%d-%Y") if not pd.isna(parsed) else v


def _normalize_numeric_like(v):
    """
    Normalize values that are numeric in nature but may be stored
    inconsistently (e.g. "007" text vs 7 numeric, or 7.0 vs 7).
    Strategy: if a value can be parsed as a float AND represents an integer
    value, normalize to a plain int string ("7"). Otherwise normalize floats
    to a fixed representation. Non-numeric strings pass through unchanged.
    This is intentionally conservative: it only touches values that clearly
    parse as numbers, so real text data is never altered.
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        f = float(v)
        if f.is_integer():
            return str(int(f))
        return repr(f)
    if isinstance(v, str):
        s = v.strip() if STRIP_WHITESPACE_IN_DATA_VALUES else v
        # Only attempt numeric parsing on strings that look purely numeric
        # (optionally signed, optional decimal point) to avoid mangling
        # things like phone numbers or IDs that merely contain digits.
        candidate = s.strip()
        try:
            f = float(candidate)
        except (ValueError, TypeError):
            return s
        # Guard against things like "1e10" style scientific text being
        # unintentionally reformatted in a surprising way -- only normalize
        # simple decimal-looking strings.
        if not candidate.replace(".", "", 1).replace("-", "", 1).isdigit():
            return s
        if f.is_integer():
            return str(int(f))
        return repr(f)
    return v


# Text values that should be treated as equivalent to a genuinely blank/
# NULL cell during comparison (case-insensitive, whitespace-insensitive).
# Some exports write one of these literal strings instead of leaving the
# cell truly empty, and to a reviewer they mean the same thing: "no value
# here". Extend this list if your source data uses other placeholder text.
NULL_LIKE_TEXT_VALUES = {"null", "n/a", "na", "none", "nan", "nat", "#n/a"}


def _normalize_value(v, strip_whitespace):
    """
    General-purpose value normalization applied to every cell before
    comparison:
      - Missing values (None/NaN/NaT) become an empty string "", so a
        genuinely blank/NULL cell on one side compares equal to an empty
        string on the other side (they are treated as the same value).
      - Text placeholders for "no value" (see NULL_LIKE_TEXT_VALUES --
        e.g. the literal text "NULL", "N/A", "None") are ALSO normalized
        to "", so a cell that holds that text on one side still matches a
        truly blank/NULL/"" cell on the other side.
      - Strings optionally have leading/trailing whitespace stripped,
        controlled by STRIP_WHITESPACE_IN_DATA_VALUES.
      - EVERYTHING ELSE is cast to str(). This is a deliberate safety net:
        it guarantees every normalized column ends up as plain text, so two
        sheets can NEVER fail to merge/compare due to a dtype mismatch
        (e.g. one sheet's column read in as real datetime objects, the
        other as text -- see _normalize_date_value above for the case that
        actually happens most often). Since numeric-like and time-like
        values are already converted to their canonical string form
        earlier in the pipeline, this final str() cast does not change
        their meaning -- it only guards against any other type slipping
        through unconverted.
    """
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass  # pd.isna can choke on some object types; safe to ignore

    if isinstance(v, str):
        s = v.strip() if strip_whitespace else v
        if s.strip().lower() in NULL_LIKE_TEXT_VALUES or s.strip() == "":
            return ""
        return s

    return str(v)  # safety net -- see docstring above


def normalize_dataframe(df, columns, apply_numeric_normalization=True,
                        apply_previous_value_fill=False):
    """
    Apply the configured normalization steps to `columns` of `df`, returning
    a NEW dataframe (does not mutate the input). This is the single place
    where "007" vs 7, time text vs datetime.time, date objects vs date
    text, and NaN vs None vs "" discrepancies are resolved -- deliberately
    visible/inspectable rather than happening implicitly inside groupby.

    Parameters
    ----------
    df : pandas.DataFrame
    columns : list of str
        Which columns to normalize (typically the common columns).
    apply_numeric_normalization : bool
        Whether to attempt numeric-string normalization (e.g. "007" -> "7").
        Exposed as a parameter so you can disable it and inspect raw
        mismatches if you suspect it's over-normalizing your data.

    Returns
    -------
    pandas.DataFrame
        Copy of df restricted to `columns`, with normalized values.
    """
    out = df[columns].copy()

    for col in columns:
        if apply_previous_value_fill:
            # This is comparison-time Tableau expansion.  It does not alter
            # the source dataframe or fill a leading blank from a future row.
            values = out[col]
            blank = values.isna() | values.map(lambda v: isinstance(v, str) and not v.strip())
            out[col] = values.mask(blank).ffill()
        is_time_col = col in FORCE_TIME_COLUMNS or _looks_like_time_series(out[col])

        def _norm(v, is_time_col=is_time_col):
            v = _normalize_date_value(v)
            if is_time_col:
                v = _normalize_time_value(v)
            if apply_numeric_normalization:
                v = _normalize_numeric_like(v)
            return _normalize_value(v, STRIP_WHITESPACE_IN_DATA_VALUES)

        out[col] = out[col].map(_norm)
        # Force the whole column to a uniform pandas string dtype as a
        # second safety net -- guarantees pd.merge() on this column can
        # never hit a "trying to merge on str and datetime64" (or similar)
        # dtype error, no matter what odd type a source cell held.
        out[col] = out[col].astype(str)

    return out


# ============================================================================
# STEP 3: DATA COMPARISON (multiset / bag comparison)
# ============================================================================

def compare_data(df_a, df_b, common_columns, tableau_previous_value_fill=False,
                 autofilled_cells_a=None, autofilled_cells_b=None,
                 display_df_a=None, display_df_b=None):
    """
    Multiset comparison of df_a and df_b restricted to `common_columns`.

    Equivalent SQL concept:
        WITH counts_a AS (
            SELECT <common_columns>, COUNT(*) AS cnt
            FROM sheet_a GROUP BY <common_columns>
        ),
        counts_b AS (
            SELECT <common_columns>, COUNT(*) AS cnt
            FROM sheet_b GROUP BY <common_columns>
        )
        SELECT ... FROM counts_a FULL OUTER JOIN counts_b
        USING (<common_columns>)
        WHERE cnt_a <> cnt_b

    Returns
    -------
    dict with keys:
        'extra_in_a' : DataFrame of common columns + CountInA + CountInB,
                        one row per distinct value-combination, for
                        combinations where A has strictly more occurrences
                        than B. CountInA/CountInB show the raw counts (not
                        just the difference) so you can see the full picture.
        'extra_in_b' : same, reversed.
    """
    # --- normalize both sides identically before grouping ---
    # NOTE: tableau_previous_value_fill defaults to False. This is a
    # separate, comparison-time-only "treat blank as same as row above"
    # step that used to run unconditionally on EVERY common column,
    # regardless of which columns the user actually chose in the
    # interactive autofill prompt (or FORWARD_FILL_COLUMNS). That silently
    # overrode the user's column selection at comparison time, and could
    # also replace a genuinely-missing NULL cell with a copied prior value
    # instead of letting it normalize to "" -- causing false mismatches
    # against a legitimately blank/"" cell on the other side. The real,
    # user-controlled autofill already happened earlier during
    # preprocessing (_preprocess_and_save -> prompt_for_autofill /
    # FORWARD_FILL_COLUMNS) and is already baked into df_a/df_b by this
    # point. Only pass True here if you explicitly want this ADDITIONAL,
    # comparison-only "same as row above" behavior applied to every
    # common column on top of that.
    norm_a = normalize_dataframe(df_a, common_columns,
                                 apply_previous_value_fill=tableau_previous_value_fill)
    norm_b = normalize_dataframe(df_b, common_columns,
                                 apply_previous_value_fill=tableau_previous_value_fill)

    # --- STEP: group by the full row-tuple and count occurrences ---
    # groupby(...).size() is the pandas equivalent of GROUP BY + COUNT(*).
    # dropna=False ensures groups containing our MISSING_SENTINEL (which is
    # a real string, not NaN, by this point) are never silently dropped --
    # included defensively in case any raw NaN slips through normalization.
    counts_a = (
        norm_a.groupby(common_columns, dropna=False)
        .size()
        .reset_index(name="CountInA")
    )
    counts_b = (
        norm_b.groupby(common_columns, dropna=False)
        .size()
        .reset_index(name="CountInB")
    )

    # --- STEP: full outer join on the row-tuple ---
    # A tuple present only in A gets CountInB = NaN -> filled with 0, and
    # vice versa. This is what lets us detect "entirely new/removed rows"
    # as well as "same row, different multiplicity".
    merged = pd.merge(counts_a, counts_b, on=common_columns, how="outer")
    merged["CountInA"] = merged["CountInA"].fillna(0).astype(int)
    merged["CountInB"] = merged["CountInB"].fillna(0).astype(int)

    # --- STEP: split into extra-in-A / extra-in-B ---
    # Only tuples with UNEQUAL counts are "different" -- equal counts (even
    # if >1, i.e. duplicates) are considered fully matched and excluded.
    extra_in_a = merged[merged["CountInA"] > merged["CountInB"]].copy()
    extra_in_b = merged[merged["CountInB"] > merged["CountInA"]].copy()

    # Sort for readability (largest discrepancy first)
    extra_in_a["_diff"] = extra_in_a["CountInA"] - extra_in_a["CountInB"]
    extra_in_b["_diff"] = extra_in_b["CountInB"] - extra_in_b["CountInA"]
    extra_in_a = extra_in_a.sort_values("_diff", ascending=False).drop(columns="_diff")
    extra_in_b = extra_in_b.sort_values("_diff", ascending=False).drop(columns="_diff")

    # --- STEP: restore original (un-normalized) display values ---
    # The counts/grouping used normalized values for correctness, but for
    # human review it's more useful to show the actual original data. We do
    # this by taking one representative original row per matched normalized
    # tuple from the sheet that has extra copies.
    extra_in_a = _attach_after_fill_values(
    extra_in_a, common_columns, norm_a, autofilled_cells_a
    )
    extra_in_b = _attach_after_fill_values(
        extra_in_b, common_columns, norm_b, autofilled_cells_b
    )

    return {"extra_in_a": extra_in_a, "extra_in_b": extra_in_b}


def _attach_after_fill_values(result_df, common_columns, normalized_df, autofilled_cells=None):
    """
    result_df currently holds NORMALIZED values for common_columns plus
    CountInA/CountInB. Replace them with one representative row of the
    AFTER-FILL (post-normalization) values from normalized_df -- i.e. the
    exact values that were actually compared -- instead of going back to
    the pre-fill original.
    """
    if result_df.empty:
        result_df = result_df.reset_index(drop=True)
        result_df.attrs["autofilled_columns_by_row"] = {}
        return result_df

    normalized_df = normalized_df.copy()
    normalized_df["_orig_index"] = normalized_df.index

    # First matching normalized row per normalized tuple (duplicates are
    # identical after normalization, so any one instance is representative).
    first_match = normalized_df.drop_duplicates(subset=common_columns, keep="first")

    lookup = pd.merge(
        result_df[common_columns + ["CountInA", "CountInB"]],
        first_match[common_columns + ["_orig_index"]],
        on=common_columns,
        how="left",
    )

    if lookup["_orig_index"].isna().any():
        missing = lookup[lookup["_orig_index"].isna()]
        raise ValueError(
            f"_attach_after_fill_values: no matching row found for {len(missing)} group(s)."
        )

    # Pull the AFTER-FILL row directly -- no trip back to the original df.
    after_fill_rows = normalized_df.loc[lookup["_orig_index"], common_columns].reset_index(drop=True)
    after_fill_rows["CountInA"] = lookup["CountInA"].values
    after_fill_rows["CountInB"] = lookup["CountInB"].values

    tracked = autofilled_cells or set()
    after_fill_rows.attrs["autofilled_columns_by_row"] = {
        output_row: {column for column, source_row in tracked if source_row == original_row}
        for output_row, original_row in enumerate(lookup["_orig_index"])
    }

    return after_fill_rows


# ============================================================================
# STEP 4 (OPTIONAL): LIKELY-EDITED ROW PAIRING (fuzzy match)
# ============================================================================

# Maximum number of common columns that may differ for two rows (one from
# ExtraInA, one from ExtraInB) to be considered "the same record, edited"
# rather than two unrelated extra rows. Tune this: lower = stricter (fewer,
# more confident pairs); higher = looser (more pairs, more false positives).
MAX_DIFFERING_COLUMNS_FOR_MATCH = 3


def _values_equal(v1, v2):
    """
    Compares two ORIGINAL (non-normalized) cell values for the purposes of
    deciding whether a column "differs" when building the human-readable
    pairing report. Unlike a plain `!=`, this treats two missing values
    (NaN/None) as EQUAL to each other -- because by definition NaN != NaN in
    Python/pandas, a naive comparison would flag two blank cells as a
    "difference" even though there's nothing there to review.
    """
    def _is_missing(v):
        if v is None:
            return True
        try:
            return bool(pd.isna(v))
        except (TypeError, ValueError):
            return False

    if _is_missing(v1) and _is_missing(v2):
        return True
    return v1 == v2


# Safety cap: the fuzzy-matcher below is O(n*m) -- comparing every leftover
# row in A against every leftover row in B. That's fine for small numbers of
# mismatched rows, but becomes computationally infeasible fast: 1,000 x
# 1,000 = 1,000,000 comparisons is fine (~seconds); 30,000 x 30,000 = 900
# million is NOT (can run for hours). If the number of rows to pair exceeds
# this cap, Step 4 is skipped automatically with a clear message rather than
# hanging silently.
MAX_ROWS_FOR_FUZZY_MATCH = 1200


def find_likely_edited_pairs(extra_in_a, extra_in_b, common_columns,
                              max_differing_columns=MAX_DIFFERING_COLUMNS_FOR_MATCH):
    """
    Attempts to pair up rows from extra_in_a and extra_in_b that are likely
    the SAME underlying record with a small number of edited values, rather
    than two completely unrelated rows.

    Approach (greedy nearest-match, not a full optimal assignment):
    For each row in extra_in_a (in order), compare it against every not-yet-
    matched row in extra_in_b, counting how many common columns differ.
    Pick the candidate with the FEWEST differing columns. If that minimum is
    <= max_differing_columns, treat them as a matched pair, remove both from
    the pool, and record which columns differed and how.

    This is intentionally simple (O(n*m) comparisons) and greedy, not a true
    optimal matching -- fine for typical small numbers of extra/mismatched
    rows. If either side exceeds MAX_ROWS_FOR_FUZZY_MATCH rows, this step is
    skipped entirely (see the 'skipped' key in the return value) rather than
    attempting a comparison that could take hours.

    Parameters
    ----------
    extra_in_a, extra_in_b : pandas.DataFrame
        Output of compare_data()['extra_in_a'] / ['extra_in_b']. Each row
        here represents ONE excess occurrence (duplicates already expanded
        out by count during pairing, see below).
    common_columns : list of str
    max_differing_columns : int
        Pairing threshold -- see module-level default above.

    Returns
    -------
    dict with keys:
        'pairs'          : DataFrame, one row per (matched A-row, B-row,
                            differing column) triple -- i.e. "long" format,
                            showing Column / ValueInA / ValueInB per diff,
                            grouped under a PairID so you can see which diffs
                            belong to the same matched pair.
        'pairs_full'     : DataFrame, one row per pair with the FULL A and B
                            record side by side.
        'unmatched_a'    : DataFrame, rows from extra_in_a that could not be
                            paired within the threshold (genuinely new/
                            removed rows, not edits).
        'unmatched_b'    : same, for extra_in_b.
        'skipped'        : bool -- True if this step was skipped because
                            extra_in_a/extra_in_b were too large (see
                            MAX_ROWS_FOR_FUZZY_MATCH). When True, 'pairs' and
                            'pairs_full' are empty and 'unmatched_a'/
                            'unmatched_b' are just extra_in_a/extra_in_b
                            unchanged (nothing was attempted).
    """
    # ---- Safety guard: bail out early if this would be computationally
    # infeasible, rather than hanging for hours. ----
    if len(extra_in_a) > MAX_ROWS_FOR_FUZZY_MATCH or len(extra_in_b) > MAX_ROWS_FOR_FUZZY_MATCH:
        print(
            f"  [Step 4 SKIPPED] {len(extra_in_a)} extra rows in A and "
            f"{len(extra_in_b)} extra rows in B -- exceeds MAX_ROWS_FOR_FUZZY_MATCH "
            f"({MAX_ROWS_FOR_FUZZY_MATCH}). Row-by-row fuzzy pairing at this scale "
            f"would take an infeasible amount of time (n x m comparisons), so it's "
            f"being skipped. This usually means something upstream isn't matching as "
            f"expected -- see ExtraInA/ExtraInB in the output to investigate before "
            f"re-running with fewer mismatched rows."
        )
        empty_pairs = pd.DataFrame(columns=["PairID", "Column", "ValueInA", "ValueInB"])
        full_columns = ["PairID", "DifferingColumns"]
        for col in common_columns:
            full_columns += [f"{col} [A]", f"{col} [B]"]
        empty_pairs_full = pd.DataFrame(columns=full_columns)
        return {
            "pairs": empty_pairs,
            "pairs_full": empty_pairs_full,
            "unmatched_a": extra_in_a,
            "unmatched_b": extra_in_b,
            "skipped": True,
            "pair_autofilled": {},
            "unmatched_autofilled_a": {},
            "unmatched_autofilled_b": {},
        }

    # Expand each row out by its "excess count" (CountInA - CountInB, or
    # vice versa) so that if a row has 2 excess copies, it's available to be
    # matched twice. Each expanded instance gets a stable row id.
    #
    # Alongside the RAW columns (for human-readable display), we also carry
    # a NORMALIZED copy of each common column (prefixed "__norm__"), built
    # via the same normalize_dataframe() pipeline used for the multiset
    # comparison in compare_data(). Diff detection below compares the
    # NORMALIZED values, not the raw ones -- otherwise a date stored as text
    # in one sheet and as a real datetime object in the other (or "77960" as
    # a string vs 77960 as an int) would always look "different" even
    # though they mean the same thing, falsely inflating DifferingColumns.
    def _expand(df, norm_df, count_col_self, count_col_other):
        rows = []
        autofilled_by_row = df.attrs.get("autofilled_columns_by_row", {})
        for idx, row in df.iterrows():
            excess = int(row[count_col_self]) - int(row[count_col_other])
            norm_row = norm_df.loc[idx]
            for _ in range(max(excess, 0)):
                combined = row.copy()
                # Internal-only marker. It is removed before any dataframe
                # is written to the report workbook.
                combined["__autofilled_columns__"] = autofilled_by_row.get(idx, set())
                for col in common_columns:
                    combined[f"__norm__{col}"] = norm_row[col]
                rows.append(combined)
        norm_cols = [f"__norm__{c}" for c in common_columns]
        if not rows:
            return pd.DataFrame(columns=list(df.columns) + norm_cols)
        return pd.DataFrame(rows).reset_index(drop=True)

    norm_a = normalize_dataframe(extra_in_a, common_columns)
    norm_b = normalize_dataframe(extra_in_b, common_columns)
    pool_a = _expand(extra_in_a, norm_a, "CountInA", "CountInB")
    pool_b = _expand(extra_in_b, norm_b, "CountInB", "CountInA")

    unmatched_b_idx = set(pool_b.index)
    pair_records = []
    full_pair_records = []
    pair_autofilled = {}
    unmatched_a_idx = []
    pair_id = 0

    for i, row_a in pool_a.iterrows():
        best_j = None
        best_diff_cols = None
        best_diff_count = None

        for j in unmatched_b_idx:
            row_b = pool_b.loc[j]
            # Compare NORMALIZED values to decide what actually differs --
            # see the _expand docstring above for why raw comparison would
            # be misleading here.
            diff_cols = [
                c for c in common_columns
                if row_a[f"__norm__{c}"] != row_b[f"__norm__{c}"]
            ]
            if best_diff_count is None or len(diff_cols) < best_diff_count:
                best_diff_count = len(diff_cols)
                best_diff_cols = diff_cols
                best_j = j
                if best_diff_count == 0:
                    break  # can't do better than an exact match on remaining cols

        if best_j is not None and best_diff_count <= max_differing_columns:
            unmatched_b_idx.discard(best_j)
            row_b = pool_b.loc[best_j]
            pair_id += 1
            # NEW: AUTO-FILL TRACKING -- side/column provenance for the two
            # mismatched report formats.  This is never added as a result
            # column, so it cannot affect comparison output.
            pair_autofilled[pair_id] = {
                "A": set(row_a.get("__autofilled_columns__", set())),
                "B": set(row_b.get("__autofilled_columns__", set())),
            }
            for col in best_diff_cols:
                pair_records.append({
                    "PairID": pair_id,
                    "Column": col,
                    "ValueInA": row_a[col],
                    "ValueInB": row_b[col],
                })
            if not best_diff_cols:
                # Identical on all common columns -- shouldn't normally
                # happen (would've matched in compare_data), but guard
                # anyway so a pair is still visible in the output.
                pair_records.append({
                    "PairID": pair_id, "Column": "(no differences)",
                    "ValueInA": "", "ValueInB": "",
                })

            # ---- Build the FULL wide record for this pair: every common
            # column from both sheets, side by side, plus a summary of
            # which columns actually differ. This is what makes the pair
            # easy to look up and hand to a dev team -- you get the whole
            # row from Sheet A and the whole row from Sheet B, not just the
            # 1-3 fields that changed.
            full_row = {
                "PairID": pair_id,
                "DifferingColumns": ", ".join(best_diff_cols) if best_diff_cols else "(identical)",
            }
            for col in common_columns:
                full_row[f"{col} [A]"] = row_a[col]
                full_row[f"{col} [B]"] = row_b[col]
            full_pair_records.append(full_row)
        else:
            unmatched_a_idx.append(i)

    pairs_df = pd.DataFrame(pair_records, columns=["PairID", "Column", "ValueInA", "ValueInB"])

    # Column order for the wide/full table: PairID, DifferingColumns, then
    # each common column's A/B values grouped together (not all A's then
    # all B's) so related values sit next to each other when scanning.
    full_columns = ["PairID", "DifferingColumns"]
    for col in common_columns:
        full_columns += [f"{col} [A]", f"{col} [B]"]
    pairs_full_df = pd.DataFrame(full_pair_records, columns=full_columns)

    # Strip the internal __norm__ helper columns before returning -- they
    # were only needed for diff detection above, not for display.
    # display_cols must be computed PER SIDE, not shared -- pool_a and
    # pool_b can have different columns whenever Sheet A/B schemas differ
    # (e.g. extra columns added to only one sheet).
    display_cols_a = [c for c in pool_a.columns if not c.startswith("__")]
    display_cols_b = [c for c in pool_b.columns if not c.startswith("__")]
    unmatched_a_df = (
        pool_a.loc[unmatched_a_idx, display_cols_a].reset_index(drop=True)
        if unmatched_a_idx else pool_a.loc[:, display_cols_a].iloc[0:0]
    )
    unmatched_b_df = (
        pool_b.loc[list(unmatched_b_idx), display_cols_b].reset_index(drop=True)
        if unmatched_b_idx else pool_b.loc[:, display_cols_b].iloc[0:0]
    )
    unmatched_autofilled_a = {
        output_row: set(pool_a.loc[pool_row, "__autofilled_columns__"])
        for output_row, pool_row in enumerate(unmatched_a_idx)
    }
    unmatched_autofilled_b = {
        output_row: set(pool_b.loc[pool_row, "__autofilled_columns__"])
        for output_row, pool_row in enumerate(list(unmatched_b_idx))
    }
    return {
        "pairs": pairs_df,             # long format: one row per differing field
        "pairs_full": pairs_full_df,   # wide format: full A + B record per pair
        "unmatched_a": unmatched_a_df,
        "unmatched_b": unmatched_b_df,
        "skipped": False,
        "pair_autofilled": pair_autofilled,
        "unmatched_autofilled_a": unmatched_autofilled_a,
        "unmatched_autofilled_b": unmatched_autofilled_b,
    }


# ============================================================================
# RESULT HIGHLIGHTING (NEW -- runs AFTER the result workbook is written)
# ============================================================================
# Everything below is pure output formatting. It re-opens the finished
# result .xlsx with openpyxl and paints a light background fill on the cells
# that represent the differences ALREADY identified by the comparison above.
# No comparison logic is re-run, no values are recalculated, and no data is
# added, removed, or reordered -- only cell fill colors are set. If anything
# in here fails for any reason, the failure is caught and reported, and the
# (already complete and correct) workbook is simply left unhighlighted.
#
# Colour key, applied consistently across every sheet:
#   light amber (HIGHLIGHT_CHANGED_VALUE_COLOR) -> a value that DIFFERS
#   light red   (HIGHLIGHT_EXTRA_IN_A_COLOR)    -> present/excess in A only
#   light blue  (HIGHLIGHT_EXTRA_IN_B_COLOR)    -> present/excess in B only

def _solid_fill(hex_color):
    """Build a solid openpyxl PatternFill from a 6-digit RRGGBB hex string."""
    from openpyxl.styles import PatternFill
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def _header_index_map(ws):
    """
    Map {header text -> 1-based column index} by reading row 1 of a
    worksheet. Reading the headers back off the sheet (rather than assuming
    a fixed column order) keeps this robust if the column layout of any
    result sheet ever changes.
    """
    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value)] = cell.column
    return headers


def _highlight_entire_data_rows(ws, fill):
    """Fill every populated data cell (row 2 downward) on a worksheet."""
    if ws.max_row is None or ws.max_row < 2:
        return 0
    filled_rows = 0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                            min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.fill = fill
        filled_rows += 1
    return filled_rows


def _highlight_column_by_name(ws, header_name, fill):
    """Fill every data cell under one named column, if that column exists."""
    headers = _header_index_map(ws)
    col_idx = headers.get(header_name)
    if not col_idx or ws.max_row is None or ws.max_row < 2:
        return 0
    count = 0
    for row_idx in range(2, ws.max_row + 1):
        ws.cell(row=row_idx, column=col_idx).fill = fill
        count += 1
    return count


def _highlight_mismatched_records(ws, fill, autofilled_by_pair=None, autofill_fill=None):
    """
    MismatchedRecords (wide format): for each matched pair, read that row's
    'DifferingColumns' value and highlight ONLY the '<col> [A]' and
    '<col> [B]' cells for the columns actually named there. Columns that
    matched between the two sheets stay unhighlighted, so what changed
    jumps out immediately even on a very wide record.
    """
    headers = _header_index_map(ws)
    diff_idx = headers.get("DifferingColumns")
    if not diff_idx or ws.max_row is None or ws.max_row < 2:
        return 0

    highlighted = 0
    for row_idx in range(2, ws.max_row + 1):
        pair_id = ws.cell(row=row_idx, column=headers.get("PairID", 1)).value
        raw = ws.cell(row=row_idx, column=diff_idx).value
        if raw not in (None, "", "(identical)"):
            for name in [part.strip() for part in str(raw).split(",") if part.strip()]:
                for suffix in ("[A]", "[B]"):
                    col_idx = headers.get(f"{name} {suffix}")
                    if col_idx:
                        ws.cell(row=row_idx, column=col_idx).fill = fill
                        highlighted += 1
            # Also tint the DifferingColumns cell itself so the row is easy to
            # spot when scrolling horizontally through a wide pair record.
            ws.cell(row=row_idx, column=diff_idx).fill = fill
        # NEW: AUTO-FILL HIGHLIGHTING. Apply last, making green the explicit
        # precedence when a value is both changed and auto-filled.
        for side, suffix in (("A", "[A]"), ("B", "[B]")):
            for name in (autofilled_by_pair or {}).get(pair_id, {}).get(side, set()):
                col_idx = headers.get(f"{name} {suffix}")
                if col_idx and autofill_fill:
                    ws.cell(row=row_idx, column=col_idx).fill = autofill_fill
    return highlighted


def _highlight_mismatched_records_short(ws, fill_a, fill_b, fill_changed,
                                        autofilled_by_pair=None, autofill_fill=None):
    """
    MismatchedRecordsShort (long format): every row here IS a difference,
    so highlight the Column name (amber), the A value (light red) and the
    B value (light blue) -- side-by-side colours make it obvious at a glance
    which sheet each value came from.
    """
    headers = _header_index_map(ws)
    if ws.max_row is None or ws.max_row < 2:
        return 0
    targets = [
        ("Column", fill_changed),
        ("ValueInA", fill_a),
        ("ValueInB", fill_b),
    ]
    count = 0
    for row_idx in range(2, ws.max_row + 1):
        for header_name, fill in targets:
            col_idx = headers.get(header_name)
            if col_idx:
                ws.cell(row=row_idx, column=col_idx).fill = fill
                count += 1
        # NEW: AUTO-FILL HIGHLIGHTING -- green takes precedence over the
        # normal A/B difference colour for the actual supplied value cell.
        pair_id = ws.cell(row=row_idx, column=headers.get("PairID", 1)).value
        column_name = ws.cell(row=row_idx, column=headers.get("Column", 1)).value
        pair_info = (autofilled_by_pair or {}).get(pair_id, {})
        for side, header_name in (("A", "ValueInA"), ("B", "ValueInB")):
            if column_name in pair_info.get(side, set()):
                col_idx = headers.get(header_name)
                if col_idx and autofill_fill:
                    ws.cell(row=row_idx, column=col_idx).fill = autofill_fill
    return count


# NEW: AUTO-FILL HIGHLIGHTING
def _highlight_autofilled_rows(ws, autofilled_by_row, fill):
    """Highlight only tracked source-value cells in a direct row report."""
    if not autofilled_by_row or ws.max_row is None or ws.max_row < 2:
        return 0
    headers = _header_index_map(ws)
    count = 0
    for output_row, columns in autofilled_by_row.items():
        for column in columns:
            col_idx = headers.get(str(column))
            if col_idx:
                ws.cell(row=output_row + 2, column=col_idx).fill = fill
                count += 1
    return count


# NEW: AUTO-FILL HIGHLIGHTING
def _add_summary_legend(ws, fill_changed, fill_a, fill_b, fill_autofilled=None):
    """Place a colour legend to the right of metrics without changing them."""
    start_col = max(ws.max_column + 2, 4)
    ws.cell(row=1, column=start_col, value="Color Legend")
    entries = [
        ("Changed Value", fill_changed),
        ("Extra in A", fill_a),
        ("Extra in B", fill_b),
    ]
    if fill_autofilled:
        entries.append(("Auto-Filled Value", fill_autofilled))
    for row, (label, fill) in enumerate(entries, start=2):
        cell = ws.cell(row=row, column=start_col, value=label)
        cell.fill = fill


def _highlight_summary(ws, fill):
    """
    Summary: highlight the Value cell of any metric that represents a
    detected difference and is NON-ZERO. Metrics that are simply
    descriptive (file names, total column/row counts) and differences that
    came out as zero are deliberately left plain, so the highlighted cells
    are exactly the ones worth investigating.
    """
    headers = _header_index_map(ws)
    metric_idx = headers.get("Metric")
    value_idx = headers.get("Value")
    if not metric_idx or not value_idx or ws.max_row is None or ws.max_row < 2:
        return 0

    difference_metrics = {
        "Columns only in A",
        "Columns only in B",
        "Row count difference (A-B)",
        "Distinct row-combos extra in A",
        "Distinct row-combos extra in B",
        "Likely-edited pairs found",
        "Still-unmatched extra rows in A",
        "Still-unmatched extra rows in B",
    }

    count = 0
    for row_idx in range(2, ws.max_row + 1):
        metric = ws.cell(row=row_idx, column=metric_idx).value
        if metric not in difference_metrics:
            continue
        value = ws.cell(row=row_idx, column=value_idx).value
        try:
            is_nonzero = float(value) != 0
        except (TypeError, ValueError):
            is_nonzero = False
        if is_nonzero:
            ws.cell(row=row_idx, column=metric_idx).fill = fill
            ws.cell(row=row_idx, column=value_idx).fill = fill
            count += 1
    return count


def highlight_result_workbook(output_path, autofill_display=None):
    """
    Post-process the generated result workbook at `output_path`, applying a
    light background fill to every cell that represents an identified
    difference. Called at the very end of run_comparison() and
    run_folder_comparison(), immediately after the workbook is written.

    Sheet-by-sheet behaviour:
      Summary                -> non-zero difference metrics (amber)
      MissingInB             -> each column name found only in A (light red)
      MissingInA             -> each column name found only in B (light blue)
      ExtraInA               -> whole row (light red)
      ExtraInB               -> whole row (light blue)
      MismatchedRecords      -> only the '[A]'/'[B]' cells whose column is
                                listed in that row's DifferingColumns (amber)
      MismatchedRecordsShort -> Column (amber), ValueInA (red), ValueInB (blue)
      UnmatchedInA           -> whole row (light red)
      UnmatchedInB           -> whole row (light blue)

    Purely cosmetic and fully optional: set HIGHLIGHT_RESULTS = False at the
    top of the file to skip it, and any unexpected error here is caught and
    printed without disturbing the workbook that was already written.
    """
    if not HIGHLIGHT_RESULTS:
        return

    try:
        import openpyxl

        fill_changed = _solid_fill(HIGHLIGHT_CHANGED_VALUE_COLOR)
        fill_a = _solid_fill(HIGHLIGHT_EXTRA_IN_A_COLOR)
        fill_b = _solid_fill(HIGHLIGHT_EXTRA_IN_B_COLOR)
        fill_summary = _solid_fill(HIGHLIGHT_SUMMARY_COLOR)
        fill_autofilled = _solid_fill(HIGHLIGHT_AUTOFILLED_COLOR) if HIGHLIGHT_AUTOFILLED_VALUES else None

        wb = openpyxl.load_workbook(output_path)

        if "Summary" in wb.sheetnames:
            _highlight_summary(wb["Summary"], fill_summary)
            _add_summary_legend(wb["Summary"], fill_changed, fill_a, fill_b, fill_autofilled)

        # Schema differences: the single column of names on each sheet IS
        # the difference, so the whole column gets tinted.
        if "MissingInB" in wb.sheetnames:
            _highlight_column_by_name(wb["MissingInB"], "ColumnOnlyInA", fill_a)
        if "MissingInA" in wb.sheetnames:
            _highlight_column_by_name(wb["MissingInA"], "ColumnOnlyInB", fill_b)

        # Extra / unmatched rows: the ENTIRE row is the finding (this whole
        # record is excess on one side), so the full row is tinted rather
        # than any individual cell.
        for sheet_name, fill in (
            ("ExtraInA", fill_a),
            ("ExtraInB", fill_b),
            ("UnmatchedInA", fill_a),
            ("UnmatchedInB", fill_b),
        ):
            if sheet_name in wb.sheetnames:
                _highlight_entire_data_rows(wb[sheet_name], fill)

        # Direct row reports retain their original displayed columns, so
        # tracked source-row provenance identifies exact cells to repaint.
        if HIGHLIGHT_AUTOFILLED_VALUES:
            for sheet_name in ("ExtraInA", "ExtraInB", "UnmatchedInA", "UnmatchedInB"):
                if sheet_name in wb.sheetnames:
                    _highlight_autofilled_rows(
                        wb[sheet_name], (autofill_display or {}).get(sheet_name, {}), fill_autofilled
                    )

        # Likely-edited pairs: only the fields that actually changed.
        if "MismatchedRecords" in wb.sheetnames:
            _highlight_mismatched_records(
                wb["MismatchedRecords"], fill_changed,
                (autofill_display or {}).get("MismatchedRecords", {}),
                fill_autofilled if HIGHLIGHT_AUTOFILLED_VALUES else None,
            )
        if "MismatchedRecordsShort" in wb.sheetnames:
            _highlight_mismatched_records_short(
                wb["MismatchedRecordsShort"], fill_a, fill_b, fill_changed,
                (autofill_display or {}).get("MismatchedRecordsShort", {}),
                fill_autofilled if HIGHLIGHT_AUTOFILLED_VALUES else None,
            )

        wb.save(output_path)
        print(
            "Highlighting applied to result workbook "
            f"(changed={HIGHLIGHT_CHANGED_VALUE_COLOR}, "
            f"extra-in-A={HIGHLIGHT_EXTRA_IN_A_COLOR}, "
            f"extra-in-B={HIGHLIGHT_EXTRA_IN_B_COLOR})."
        )
    except Exception as exc:  # never let cosmetics break a finished report
        print(f"[WARNING] Could not apply highlighting to {output_path}: {exc}")
        print("The result workbook itself was written successfully and is complete.")


# NEW: AUTO-FILL TRACKING
def _autofill_display_info(extra_in_a, extra_in_b, match_result):
    """Collect non-tabular provenance used only by the formatting pass."""
    return {
        "ExtraInA": extra_in_a.attrs.get("autofilled_columns_by_row", {}),
        "ExtraInB": extra_in_b.attrs.get("autofilled_columns_by_row", {}),
        "UnmatchedInA": match_result.get("unmatched_autofilled_a", {}),
        "UnmatchedInB": match_result.get("unmatched_autofilled_b", {}),
        "MismatchedRecords": match_result.get("pair_autofilled", {}),
        "MismatchedRecordsShort": match_result.get("pair_autofilled", {}),
    }


# ============================================================================
# ORCHESTRATION
# ============================================================================

def run_comparison(path_a, path_b, sheet_a=0, sheet_b=0, output_path="comparison_output.xlsx",
                   interactive_autofill=True):
    """
    Runs the full comparison pipeline in order (schema -> row counts ->
    data), prints a console summary, and writes results to an output .xlsx
    workbook with one sheet per result set.
    """
    print("=" * 70)
    print("LOADING SHEETS")
    print("=" * 70)
    # Merged ranges are a property of the actual source worksheet, not its
    # SheetA/SheetB role. Expand them before reading either input.
    # Keep untouched display copies separate from the working comparison
    # copies.  The latter expand merged cells; the former are used only when
    # writing human-readable difference rows.
    display_df_a, _ = load_sheet(path_a, sheet_a)
    display_df_b, _ = load_sheet(path_b, sheet_b)
    df_a, percent_mask_a = load_sheet(path_a, sheet_a, expand_powerbi_merges=True)
    df_b, percent_mask_b = load_sheet(path_b, sheet_b, expand_powerbi_merges=True)
    autofilled_cells_a = set()
    autofilled_cells_b = set()
    if interactive_autofill:
        df_a, autofilled_cells_a = prompt_for_autofill(df_a, "SheetA")
        df_b, autofilled_cells_b = prompt_for_autofill(df_b, "SheetB")
    if FORWARD_FILL_COLUMNS:
        print(f"Forward-filling merged-cell columns: {FORWARD_FILL_COLUMNS}")
        df_a, configured_cells_a = forward_fill_merged_cells(df_a, FORWARD_FILL_COLUMNS)
        df_b, configured_cells_b = forward_fill_merged_cells(df_b, FORWARD_FILL_COLUMNS)
        autofilled_cells_a.update(configured_cells_a)
        autofilled_cells_b.update(configured_cells_b)
    # Convert date/datetime columns to MM-DD-YYYY (time dropped) before the
    # rest of the pipeline runs, so schema/row-count/data comparison all see
    # the normalized date representation. Each sheet's own day/month
    # convention is inferred independently -- see convert_date_columns_to_mmddyyyy.
    df_a = convert_date_columns_to_mmddyyyy(df_a)
    df_b = convert_date_columns_to_mmddyyyy(df_b)
    # Convert percent-formatted numeric cells (and already-text percentages)
    # to the same canonical "0.63%" text form on both sides -- see the
    # PERCENTAGE COLUMN DETECTION + CONVERSION section for why this is
    # needed (Excel stores a percent-formatted cell as a raw fraction,
    # invisible to a plain value read).
    df_a = convert_percent_columns_to_text(df_a, percent_mask_a)
    df_b = convert_percent_columns_to_text(df_b, percent_mask_b)
    print(f"Sheet A: {path_a!r} (sheet={sheet_a!r}) -> {df_a.shape[0]} rows, {df_a.shape[1]} cols")
    print(f"Sheet B: {path_b!r} (sheet={sheet_b!r}) -> {df_b.shape[0]} rows, {df_b.shape[1]} cols")

    # ---- 1. SCHEMA CHECK ----
    print("\n" + "=" * 70)
    print("STEP 1: SCHEMA CHECK")
    print("=" * 70)
    schema = compare_schema(df_a, df_b)
    df_b_comparison = align_common_headers(df_b, schema)
    display_df_b = align_common_headers(display_df_b, schema)
    # Result sheets use Sheet A's display names for whitespace-equivalent
    # common headers, so translate only the in-memory tracking labels too.
    b_to_result_header = dict(zip(schema["common_b"], schema["common"]))
    autofilled_cells_b = {
        (b_to_result_header.get(column, column), row)
        for column, row in autofilled_cells_b
    }
    print(f"Total columns in A: {schema['count_a']}")
    print(f"Total columns in B: {schema['count_b']}")
    print(f"Columns only in A ({len(schema['only_in_a'])}): {schema['only_in_a']}")
    print(f"Columns only in B ({len(schema['only_in_b'])}): {schema['only_in_b']}")
    print(f"Common columns ({len(schema['common'])}): {schema['common']}")

    if not schema["common"]:
        print("\nNo common columns -- cannot perform data comparison. Stopping.")
        return

    # ---- 2. RECORD COUNT CHECK ----
    print("\n" + "=" * 70)
    print("STEP 2: RECORD COUNT CHECK")
    print("=" * 70)
    row_counts = compare_row_counts(df_a, df_b)
    print(f"Rows in A: {row_counts['rows_a']}")
    print(f"Rows in B: {row_counts['rows_b']}")
    print(f"Difference (A - B): {row_counts['difference']}")

    # ---- 3. DATA COMPARISON ----
    print("\n" + "=" * 70)
    print("STEP 3: DATA COMPARISON (multiset / GROUP BY + COUNT(*) approach)")
    print("=" * 70)
    print(f"Comparing on {len(schema['common'])} common columns.")
    print(f"STRIP_WHITESPACE_IN_DATA_VALUES = {STRIP_WHITESPACE_IN_DATA_VALUES}")
    data_result = compare_data(df_a, df_b_comparison, schema["common"],
                               autofilled_cells_a=autofilled_cells_a,
                               autofilled_cells_b=autofilled_cells_b,
                               display_df_a=display_df_a,
                               display_df_b=display_df_b)
    extra_in_a = data_result["extra_in_a"]
    extra_in_b = data_result["extra_in_b"]
    print(f"Distinct row-combinations with EXTRA occurrences in A: {len(extra_in_a)}")
    print(f"Distinct row-combinations with EXTRA occurrences in B: {len(extra_in_b)}")

    # ---- STEP 4: LIKELY-EDITED ROW PAIRING (fuzzy match) ----
    print("\n" + "=" * 70)
    print("STEP 4: LIKELY-EDITED ROW PAIRING (fuzzy match)")
    print("=" * 70)
    print(f"MAX_DIFFERING_COLUMNS_FOR_MATCH = {MAX_DIFFERING_COLUMNS_FOR_MATCH}")
    match_result = find_likely_edited_pairs(extra_in_a, extra_in_b, schema["common"])
    pairs_df = match_result["pairs"]
    pairs_full_df = match_result["pairs_full"]
    unmatched_a = match_result["unmatched_a"]
    unmatched_b = match_result["unmatched_b"]
    num_pairs = pairs_df["PairID"].nunique() if not pairs_df.empty else 0
    print(f"Likely-edited pairs found: {num_pairs}")
    print(f"Still-unmatched extra rows in A (genuinely new/removed, not edits): {len(unmatched_a)}")
    print(f"Still-unmatched extra rows in B: {len(unmatched_b)}")

    # ---- WRITE OUTPUT WORKBOOK ----
    print("\n" + "=" * 70)
    print(f"WRITING OUTPUT -> {output_path}")
    print("=" * 70)

    missing_in_a_df = pd.DataFrame({"ColumnOnlyInB": schema["only_in_b"]})
    missing_in_b_df = pd.DataFrame({"ColumnOnlyInA": schema["only_in_a"]})
    summary_df = pd.DataFrame(
        {
            "Metric": [
                "Columns in A", "Columns in B",
                "Columns only in A", "Columns only in B",
                "Rows in A", "Rows in B", "Row count difference (A-B)",
                "Distinct row-combos extra in A", "Distinct row-combos extra in B",
                "Likely-edited pairs found", "Still-unmatched extra rows in A", "Still-unmatched extra rows in B",
            ],
            "Value": [
                schema["count_a"], schema["count_b"],
                len(schema["only_in_a"]), len(schema["only_in_b"]),
                row_counts["rows_a"], row_counts["rows_b"], row_counts["difference"],
                len(extra_in_a), len(extra_in_b),
                num_pairs, len(unmatched_a), len(unmatched_b),
            ],
        }
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        missing_in_b_df.to_excel(writer, sheet_name="MissingInB", index=False)  # cols in A, absent from B
        missing_in_a_df.to_excel(writer, sheet_name="MissingInA", index=False)  # cols in B, absent from A
        extra_in_a.to_excel(writer, sheet_name="ExtraInA", index=False)
        extra_in_b.to_excel(writer, sheet_name="ExtraInB", index=False)
        pairs_full_df.to_excel(writer, sheet_name="MismatchedRecords", index=False)      # full A+B record per pair
        pairs_df.to_excel(writer, sheet_name="MismatchedRecordsShort", index=False)      # just the differing fields
        unmatched_a.to_excel(writer, sheet_name="UnmatchedInA", index=False)
        unmatched_b.to_excel(writer, sheet_name="UnmatchedInB", index=False)

    # ---- HIGHLIGHT THE IDENTIFIED DIFFERENCES (cosmetic post-processing) ----
    highlight_result_workbook(output_path, _autofill_display_info(extra_in_a, extra_in_b, match_result))

    print("Done.")


# ============================================================================
# ENTRY POINT
# ============================================================================

# ============================================================================
# FOLDER-BASED WORKFLOW (auto-discover, compare, archive, timestamped output)
# ============================================================================
# Assumption (matches how you described the workflow):
#   - PATH_A (Sheet A) = the file you drop into the TABLEAU folder
#   - PATH_B (Sheet B) = the file you drop into the POWERBI folder
# If that's backwards for you, just swap the two folder paths below.

import shutil
from pathlib import Path
from datetime import datetime

# ---- EDIT THESE FOLDER PATHS ONCE, THEN NEVER TOUCH THEM AGAIN ----
# TABLEAU_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\NeedToTestTBL")
# POWERBI_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\NeedToTestPBI")
# ARCHIVE_TABLEAU_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\Tableau_Archive")
# ARCHIVE_POWERBI_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\Powerbi_Archive")
# RESULT_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\Comparision_Tbl_PBI")

BASE_FOLDER = Path(__file__).resolve().parent

TABLEAU_FOLDER = BASE_FOLDER / "NeedToTestTBL"
POWERBI_FOLDER = BASE_FOLDER / "NeedToTestPBI"

ARCHIVE_TABLEAU_FOLDER = BASE_FOLDER / "Tableau_Archive"
ARCHIVE_POWERBI_FOLDER = BASE_FOLDER / "Powerbi_Archive"

RESULT_FOLDER = BASE_FOLDER / "Comparision_Tbl_PBI"

# File extensions to look for when auto-discovering the file in each folder.
VALID_EXTENSIONS = (".xlsx", ".xls", ".csv")


def _find_single_file(folder):
    """
    Look inside `folder` for exactly one data file (.xlsx/.xls/.csv).
    Raises a clear error if the folder has zero or more than one candidate,
    since auto-discovery only works safely when there's no ambiguity about
    which file to pick up.
    """
    folder = Path(folder)
    if not folder.exists():
        raise FileNotFoundError(f"Folder does not exist: {folder}")

    candidates = [
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in VALID_EXTENSIONS
        and not f.name.startswith("~$")  # ignore Excel's temp lock files
    ]

    if len(candidates) == 0:
        raise FileNotFoundError(f"No data file found in {folder}. Drop exactly one file there.")
    if len(candidates) > 1:
        names = ", ".join(f.name for f in candidates)
        raise RuntimeError(
            f"Found {len(candidates)} files in {folder}, expected exactly 1: {names}\n"
            f"Remove the extra file(s) so it's unambiguous which one to compare."
        )
    return candidates[0]


def _load_any(path, sheet_name=0, expand_powerbi_merges=False):
    """
    Reads .xlsx/.xls via load_sheet(), or .csv directly, based on
    extension. Always returns (dataframe, percent_format_mask) -- the mask
    is None for CSV, since CSV carries no cell-formatting metadata at all
    (see convert_percent_columns_to_text).
    """
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=object), None
    return load_sheet(path, sheet_name, expand_powerbi_merges=expand_powerbi_merges)


def _preprocess_and_save(path, sheet_index, sheet_label, interactive_autofill,
                         expand_powerbi_merges=True):
    """
    Loads a source file, applies (in order) user-selected autofill,
    configured merged-cell forward-fill, and date-column conversion to
    MM-DD-YYYY, then writes the result BACK to the same source file on
    disk so the modifications are preserved before archiving.

    Returns the modified in-memory dataframe (already reflecting exactly
    what was written to disk) so the caller doesn't need to re-read the
    file.
    """
    df, percent_mask = _load_any(path, sheet_index, expand_powerbi_merges=expand_powerbi_merges)
    autofilled_cells = set()
    if interactive_autofill:
        df, autofilled_cells = prompt_for_autofill(df, sheet_label)
    if FORWARD_FILL_COLUMNS:
        print(f"Forward-filling merged-cell columns in {sheet_label}: {FORWARD_FILL_COLUMNS}")
        df, configured_cells = forward_fill_merged_cells(df, FORWARD_FILL_COLUMNS)
        autofilled_cells.update(configured_cells)
    df = convert_date_columns_to_mmddyyyy(df)
    df = convert_percent_columns_to_text(df, percent_mask)
    save_dataframe_to_source(df, path, sheet_index)
    return df, autofilled_cells


def run_folder_comparison(sheet_a_index=0, sheet_b_index=0,
                           max_differing_columns=MAX_DIFFERING_COLUMNS_FOR_MATCH,
                           interactive_autofill=True):
    """
    Full folder-based workflow:
      1. Auto-find the one file in TABLEAU_FOLDER (Sheet A) and POWERBI_FOLDER (Sheet B).
      2. For each file: load -> autofill -> merged-cell fill -> convert date
         columns to MM-DD-YYYY -> SAVE the modified data back to that same
         source file on disk (so the file itself now reflects every change
         made during preprocessing).
      3. Run the full comparison (schema, row counts, data, likely-edited
         pairing) using the modified data.
      4. Write results to RESULT_FOLDER as:
             <TableauFileName>_<YYYYMMDD>_<HHMMSS>.xlsx
         and then highlight every identified difference in that workbook.
      5. Move the (now-modified) source files into their Archive folders as:
             <filename>_TBL_<YYYYMMDD>.xlsx   (from Tableau folder)
             <filename>_PBI_<YYYYMMDD>.xlsx   (from Powerbi folder)
         so the source folders are empty again and ready for the next run,
         and the archive contains the actual processed input, not the
         untouched original.
    """
    RESULT_FOLDER.mkdir(parents=True, exist_ok=True)
    ARCHIVE_TABLEAU_FOLDER.mkdir(parents=True, exist_ok=True)
    ARCHIVE_POWERBI_FOLDER.mkdir(parents=True, exist_ok=True)

    # ---- 1. Discover files ----
    path_a = _find_single_file(TABLEAU_FOLDER)   # Sheet A
    path_b = _find_single_file(POWERBI_FOLDER)   # Sheet B

    # Safety check: if TABLEAU_FOLDER and POWERBI_FOLDER accidentally point
    # to the same folder (or somehow resolve to the same file), stop here
    # with a clear message instead of comparing a file to itself and then
    # crashing later when trying to archive the same file twice.
    if path_a.resolve() == path_b.resolve():
        raise RuntimeError(
            f"TABLEAU_FOLDER and POWERBI_FOLDER both resolved to the SAME file:\n"
            f"  {path_a}\n"
            f"Check that TABLEAU_FOLDER and POWERBI_FOLDER point to two DIFFERENT "
            f"folders, each containing only its own file."
        )

    print(f"Found Tableau file (Sheet A): {path_a}")
    print(f"Found Powerbi file (Sheet B): {path_b}")

    # Preserve the source representation for result-sheet display.  Working
    # copies below may expand merges, forward-fill report blanks, and convert
    # dates before grouping, but those transformations must not replace the
    # values a reviewer sees in ExtraInA/ExtraInB and mismatch reports.
    display_df_a, _ = _load_any(path_a, sheet_a_index)
    display_df_b, _ = _load_any(path_b, sheet_b_index)

    now = datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")

    # ---- 2. Per-file preprocessing, with the modified data written back to
    # the source file BEFORE comparison/archiving (Sheet A fully processed
    # and saved, then Sheet B fully processed and saved) ----
    print("\n" + "=" * 70)
    print("PREPROCESSING SHEET A (autofill -> date conversion -> save to source file)")
    print("=" * 70)
    df_a, autofilled_cells_a = _preprocess_and_save(
        path_a, sheet_a_index, "SheetA", interactive_autofill,
        expand_powerbi_merges=False,
    )

    print("\n" + "=" * 70)
    print("PREPROCESSING SHEET B (autofill -> date conversion -> save to source file)")
    print("=" * 70)
    df_b, autofilled_cells_b = _preprocess_and_save(
        path_b, sheet_b_index, "SheetB", interactive_autofill,
        expand_powerbi_merges=True,
    )

    schema = compare_schema(df_a, df_b)
    df_b_comparison = align_common_headers(df_b, schema)
    display_df_b = align_common_headers(display_df_b, schema)
    b_to_result_header = dict(zip(schema["common_b"], schema["common"]))
    autofilled_cells_b = {
        (b_to_result_header.get(column, column), row)
        for column, row in autofilled_cells_b
    }
    row_counts = compare_row_counts(df_a, df_b)

    print("\n" + "=" * 70)
    print("SCHEMA CHECK")
    print("=" * 70)
    print(f"Columns only in A ({len(schema['only_in_a'])}): {schema['only_in_a']}")
    print(f"Columns only in B ({len(schema['only_in_b'])}): {schema['only_in_b']}")

    print("\n" + "=" * 70)
    print("ROW COUNT CHECK")
    print("=" * 70)
    print(f"Rows in A: {row_counts['rows_a']}  Rows in B: {row_counts['rows_b']}  "
          f"Diff: {row_counts['difference']}")

    if not schema["common"]:
        print("No common columns -- cannot perform data comparison. Stopping (files NOT archived).")
        return

    print("\n" + "=" * 70)
    print("DATA COMPARISON")
    print("=" * 70)
    data_result = compare_data(df_a, df_b_comparison, schema["common"],
                               autofilled_cells_a=autofilled_cells_a,
                               autofilled_cells_b=autofilled_cells_b,
                               display_df_a=display_df_a,
                               display_df_b=display_df_b)
    extra_in_a = data_result["extra_in_a"]
    extra_in_b = data_result["extra_in_b"]
    print(f"Extra in A: {len(extra_in_a)}   Extra in B: {len(extra_in_b)}")

    match_result = find_likely_edited_pairs(extra_in_a, extra_in_b, schema["common"], max_differing_columns)
    pairs_df = match_result["pairs"]
    pairs_full_df = match_result["pairs_full"]
    unmatched_a = match_result["unmatched_a"]
    unmatched_b = match_result["unmatched_b"]
    num_pairs = pairs_df["PairID"].nunique() if not pairs_df.empty else 0
    print(f"Likely-edited pairs: {num_pairs}   Unmatched A: {len(unmatched_a)}   Unmatched B: {len(unmatched_b)}")

    # ---- 3. Write timestamped output into Result folder ----
    output_filename = f"{path_a.stem}_{date_str}_{time_str}.xlsx"
    output_path = RESULT_FOLDER / output_filename

    missing_in_a_df = pd.DataFrame({"ColumnOnlyInB": schema["only_in_b"]})
    missing_in_b_df = pd.DataFrame({"ColumnOnlyInA": schema["only_in_a"]})
    summary_df = pd.DataFrame({
        "Metric": [
            "Tableau file (A)", "Powerbi file (B)",
            "Columns in A", "Columns in B", "Columns only in A", "Columns only in B",
            "Rows in A", "Rows in B", "Row count difference (A-B)",
            "Distinct row-combos extra in A", "Distinct row-combos extra in B",
            "Likely-edited pairs found", "Still-unmatched extra rows in A", "Still-unmatched extra rows in B",
        ],
        "Value": [
            path_a.name, path_b.name,
            schema["count_a"], schema["count_b"], len(schema["only_in_a"]), len(schema["only_in_b"]),
            row_counts["rows_a"], row_counts["rows_b"], row_counts["difference"],
            len(extra_in_a), len(extra_in_b),
            num_pairs, len(unmatched_a), len(unmatched_b),
        ],
    })

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        missing_in_b_df.to_excel(writer, sheet_name="MissingInB", index=False)
        missing_in_a_df.to_excel(writer, sheet_name="MissingInA", index=False)
        extra_in_a.to_excel(writer, sheet_name="ExtraInA", index=False)
        extra_in_b.to_excel(writer, sheet_name="ExtraInB", index=False)
        pairs_full_df.to_excel(writer, sheet_name="MismatchedRecords", index=False)      # full A+B record per pair
        pairs_df.to_excel(writer, sheet_name="MismatchedRecordsShort", index=False)      # just the differing fields
        unmatched_a.to_excel(writer, sheet_name="UnmatchedInA", index=False)
        unmatched_b.to_excel(writer, sheet_name="UnmatchedInB", index=False)

    # ---- HIGHLIGHT THE IDENTIFIED DIFFERENCES (cosmetic post-processing) ----
    highlight_result_workbook(output_path, _autofill_display_info(extra_in_a, extra_in_b, match_result))

    print(f"\nResult written to: {output_path}")

    # ---- 4. Archive the (already-modified-on-disk) source files ----
    archived_a = ARCHIVE_TABLEAU_FOLDER / f"{path_a.stem}_TBL_{date_str}{path_a.suffix}"
    archived_b = ARCHIVE_POWERBI_FOLDER / f"{path_b.stem}_PBI_{date_str}{path_b.suffix}"

    # Avoid silently overwriting an existing archive file from an earlier
    # run today -- append the time too if a same-day archive already exists.
    if archived_a.exists():
        archived_a = ARCHIVE_TABLEAU_FOLDER / f"{path_a.stem}_TBL_{date_str}_{time_str}{path_a.suffix}"
    if archived_b.exists():
        archived_b = ARCHIVE_POWERBI_FOLDER / f"{path_b.stem}_PBI_{date_str}_{time_str}{path_b.suffix}"

    shutil.move(str(path_a), str(archived_a))
    shutil.move(str(path_b), str(archived_b))
    print(f"Archived Tableau file (modified) -> {archived_a}")
    print(f"Archived Powerbi file (modified) -> {archived_b}")
    print("\nDone. Source folders are now empty and ready for the next run.")


if __name__ == "__main__":
    run_folder_comparison()