

import pandas as pd
import numpy as np

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


# ============================================================================
# STEP 0: LOAD
# ============================================================================

def load_sheet(path, sheet_name=0):
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

    Returns
    -------
    pandas.DataFrame
        The raw sheet data, with headers untouched and all values read as
        openpyxl/pandas naturally infers them (dtype normalization happens
        later, explicitly, in compare_data()).
    """
    # engine="openpyxl" reads .xlsx reliably; dtype=object keeps pandas from
    # silently coercing types on read (e.g. turning "007" into 7), so we can
    # inspect and normalize types ourselves in a controlled, visible step.
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl", dtype=object)

    # pandas.read_excel does NOT strip header whitespace by default, but we
    # assert/document that explicitly here so future pandas versions or
    # accidental changes don't silently break this guarantee.
    # (No transformation is applied to df.columns -- this is intentional.)
    return df


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


def forward_fill_merged_cells(df, columns):
    """
    Forward-fills blank cells in the given columns, on the assumption that
    the blanks come from merged cells in the original spreadsheet (a value
    only present on the first row of a repeated block). Returns a NEW
    dataframe -- does not mutate the input.
    """
    df = df.copy()
    for col in columns:
        if col in df.columns:
            # ffill only recognizes NaN, while exports often represent blank
            # merged cells as empty/whitespace strings.  Convert those to
            # missing first.  pandas leaves leading missing values missing,
            # which is the required first-row behavior.
            values = df[col].copy()
            blank = values.isna() | values.map(lambda v: isinstance(v, str) and not v.strip())
            df[col] = values.mask(blank).ffill()
    return df


def prompt_for_autofill(df, sheet_label):
    """Interactively forward-fill only columns explicitly selected by user."""
    while True:
        print(f"\nDo you want to autofill any columns in {sheet_label}?")
        print("1. Yes\n2. No")
        choice = input("Select 1 or 2: ").strip()
        if choice == "2":
            return df
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
            selected_columns = list(dict.fromkeys(columns[index - 1] for index in indices))
            return forward_fill_merged_cells(df, selected_columns)
        except ValueError:
            print("Enter one or more valid column numbers, separated by commas.")


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


def _normalize_date_value(v):
    """Return YYYY-MM-DD for real datetimes and ISO-style datetime text."""
    import datetime
    import re

    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date, np.datetime64)):
        parsed = pd.to_datetime(v, errors="coerce")
        return parsed.strftime("%Y-%m-%d") if not pd.isna(parsed) else v
    # Restrict text parsing to unambiguous ISO dates, so ordinary identifiers
    # and free text are never accidentally converted to dates.
    if isinstance(v, str) and re.match(r"^\s*\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?\s*$", v):
        parsed = pd.to_datetime(v.strip(), errors="coerce")
        return parsed.strftime("%Y-%m-%d") if not pd.isna(parsed) else v
    return v


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


def _normalize_value(v, strip_whitespace):
    """
    General-purpose value normalization applied to every cell before
    comparison:
      - Missing values (None/NaN/NaT) become the MISSING sentinel, so they
        group together with each other, distinct from "" or the text "None".
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
        return _MISSING_SENTINEL
    if isinstance(v, float) and pd.isna(v):
        return _MISSING_SENTINEL
    try:
        if pd.isna(v):
            return _MISSING_SENTINEL
    except (TypeError, ValueError):
        pass  # pd.isna can choke on some object types; safe to ignore

    if isinstance(v, str):
        return v.strip() if strip_whitespace else v

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

def compare_data(df_a, df_b, common_columns, tableau_previous_value_fill=True):
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
    norm_a = normalize_dataframe(df_a, common_columns,
                                 apply_previous_value_fill=tableau_previous_value_fill)
    norm_b = normalize_dataframe(df_b, common_columns)

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
    extra_in_a = _attach_original_values(extra_in_a, df_a, common_columns, norm_a)
    extra_in_b = _attach_original_values(extra_in_b, df_b, common_columns, norm_b)

    return {"extra_in_a": extra_in_a, "extra_in_b": extra_in_b}


def _attach_original_values(result_df, original_df, common_columns, normalized_df):
    """
    result_df currently holds NORMALIZED values for common_columns (since it
    came from grouping on normalized_df) plus CountInA/CountInB. Replace the
    normalized values with one representative row of ORIGINAL (pre-
    normalization) values from original_df, so the output is readable real
    data rather than sentinel-substituted/type-coerced strings.
    """
    if result_df.empty:
        return result_df.reset_index(drop=True)

    # Attach a temporary key to normalized_df so we can look up, per
    # normalized tuple, the index of one matching original row.
    normalized_df = normalized_df.copy()
    normalized_df["_orig_index"] = normalized_df.index

    # First matching original row per normalized tuple (any one instance is
    # representative -- they're duplicates of each other after all).
    first_match = normalized_df.drop_duplicates(subset=common_columns, keep="first")

    lookup = pd.merge(
        result_df[common_columns + ["CountInA", "CountInB"]],
        first_match[common_columns + ["_orig_index"]],
        on=common_columns,
        how="left",
    )

    orig_rows = original_df.loc[lookup["_orig_index"]].reset_index(drop=True)
    orig_rows["CountInA"] = lookup["CountInA"].values
    orig_rows["CountInB"] = lookup["CountInB"].values
    return orig_rows


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
        for idx, row in df.iterrows():
            excess = int(row[count_col_self]) - int(row[count_col_other])
            norm_row = norm_df.loc[idx]
            for _ in range(max(excess, 0)):
                combined = row.copy()
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
    display_cols_a = [c for c in pool_a.columns if not c.startswith("__norm__")]
    display_cols_b = [c for c in pool_b.columns if not c.startswith("__norm__")]
    unmatched_a_df = (
        pool_a.loc[unmatched_a_idx, display_cols_a].reset_index(drop=True)
        if unmatched_a_idx else pool_a.loc[:, display_cols_a].iloc[0:0]
    )
    unmatched_b_df = (
        pool_b.loc[list(unmatched_b_idx), display_cols_b].reset_index(drop=True)
        if unmatched_b_idx else pool_b.loc[:, display_cols_b].iloc[0:0]
    )
    return {
        "pairs": pairs_df,             # long format: one row per differing field
        "pairs_full": pairs_full_df,   # wide format: full A + B record per pair
        "unmatched_a": unmatched_a_df,
        "unmatched_b": unmatched_b_df,
        "skipped": False,
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
    df_a = load_sheet(path_a, sheet_a)
    df_b = load_sheet(path_b, sheet_b)
    if interactive_autofill:
        df_a = prompt_for_autofill(df_a, "SheetA")
        df_b = prompt_for_autofill(df_b, "SheetB")
    if FORWARD_FILL_COLUMNS:
        print(f"Forward-filling merged-cell columns: {FORWARD_FILL_COLUMNS}")
        df_a = forward_fill_merged_cells(df_a, FORWARD_FILL_COLUMNS)
        df_b = forward_fill_merged_cells(df_b, FORWARD_FILL_COLUMNS)
    print(f"Sheet A: {path_a!r} (sheet={sheet_a!r}) -> {df_a.shape[0]} rows, {df_a.shape[1]} cols")
    print(f"Sheet B: {path_b!r} (sheet={sheet_b!r}) -> {df_b.shape[0]} rows, {df_b.shape[1]} cols")

    # ---- 1. SCHEMA CHECK ----
    print("\n" + "=" * 70)
    print("STEP 1: SCHEMA CHECK")
    print("=" * 70)
    schema = compare_schema(df_a, df_b)
    df_b_comparison = align_common_headers(df_b, schema)
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
    data_result = compare_data(df_a, df_b_comparison, schema["common"])
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
TABLEAU_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\NeedToTestTBL")

POWERBI_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\NeedToTestPBI")

ARCHIVE_TABLEAU_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\Tableau_Archive")

ARCHIVE_POWERBI_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\Powerbi_Archive")

RESULT_FOLDER = Path(r"D:\Blick_Tickets\Microsft_fabric\Data_Validations_Tableau_Powerbi\Comparision_Tbl_PBI")

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


def _load_any(path, sheet_name=0):
    """Reads .xlsx/.xls via load_sheet(), or .csv directly, based on extension."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=object)
    return load_sheet(path, sheet_name)


def run_folder_comparison(sheet_a_index=0, sheet_b_index=0,
                           max_differing_columns=MAX_DIFFERING_COLUMNS_FOR_MATCH,
                           interactive_autofill=True):
    """
    Full folder-based workflow:
      1. Auto-find the one file in TABLEAU_FOLDER (Sheet A) and POWERBI_FOLDER (Sheet B).
      2. Run the full comparison (schema, row counts, data, likely-edited pairing).
      3. Write results to RESULT_FOLDER as:
             <TableauFileName>_<YYYYMMDD>_<HHMMSS>.xlsx
      4. Move the source files into their Archive folders as:
             <filename>_TBL_<YYYYMMDD>.xlsx   (from Tableau folder)
             <filename>_PBI_<YYYYMMDD>.xlsx   (from Powerbi folder)
         so the source folders are empty again and ready for the next run.
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

    now = datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")

    # ---- 2. Load + run the existing comparison pipeline ----
    df_a = _load_any(path_a, sheet_a_index)
    df_b = _load_any(path_b, sheet_b_index)
    if interactive_autofill:
        df_a = prompt_for_autofill(df_a, "SheetA")
        df_b = prompt_for_autofill(df_b, "SheetB")
    if FORWARD_FILL_COLUMNS:
        print(f"Forward-filling merged-cell columns: {FORWARD_FILL_COLUMNS}")
        df_a = forward_fill_merged_cells(df_a, FORWARD_FILL_COLUMNS)
        df_b = forward_fill_merged_cells(df_b, FORWARD_FILL_COLUMNS)

    schema = compare_schema(df_a, df_b)
    df_b_comparison = align_common_headers(df_b, schema)
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
    data_result = compare_data(df_a, df_b_comparison, schema["common"])
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

    print(f"\nResult written to: {output_path}")

    # ---- 4. Archive the source files ----
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
    print(f"Archived Tableau file -> {archived_a}")
    print(f"Archived Powerbi file -> {archived_b}")
    print("\nDone. Source folders are now empty and ready for the next run.")


if __name__ == "__main__":
    run_folder_comparison()
