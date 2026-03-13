"""
outlook_email_parser.py
-----------------------
Reads emails from a specific Microsoft Outlook folder, filters by received date
(exactly X days ago), parses HTML tables from each matching email, and returns
all parsed rows as a single consolidated pandas DataFrame.

Requirements:
    pip install pywin32 pandas lxml html5lib beautifulsoup4

Compatibility: Python 3.9+
Platform:      Windows only (uses win32com.client / Outlook COM interface)
"""

from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s – %(message)s",
)
log = logging.getLogger("outlook_email_parser")


# ---------------------------------------------------------------------------
# 1. Outlook namespace
# ---------------------------------------------------------------------------

def get_outlook_namespace():
    """
    Connect to the running Outlook instance via COM and return the MAPI
    namespace object.

    Raises
    ------
    ImportError
        If pywin32 is not installed.
    RuntimeError
        If Outlook is not running or cannot be connected to.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise ImportError(
            "pywin32 is required. Install it with:  pip install pywin32"
        ) from exc

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace
    except Exception as exc:
        raise RuntimeError(
            f"Could not connect to Outlook. Make sure Outlook is open. Error: {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# 2. Folder navigation
# ---------------------------------------------------------------------------

def get_folder(
    namespace,
    folder_path: Tuple[str, ...],
    mailbox_name: Optional[str] = None,
):
    """
    Traverse a nested folder path inside Outlook and return the target folder.

    Parameters
    ----------
    namespace:
        MAPI namespace returned by ``get_outlook_namespace()``.
    folder_path:
        Ordered tuple of folder names, e.g. ``("Inbox", "MCBox", "Trades", "EMail")``.
        The first element is resolved against the root of ``mailbox_name`` (or the
        default store if ``mailbox_name`` is None).
    mailbox_name:
        Display name of the mailbox / shared mailbox, e.g. ``"shared@company.com"``.
        When ``None`` the default store (the user's own mailbox) is used.

    Returns
    -------
    Outlook Folder COM object.

    Raises
    ------
    ValueError
        If ``folder_path`` is empty.
    KeyError
        If any folder in the path cannot be found.
    """
    if not folder_path:
        raise ValueError("folder_path must contain at least one folder name.")

    # ---- resolve root store ------------------------------------------------
    if mailbox_name:
        root = None
        for store in namespace.Stores:
            if store.DisplayName.strip().lower() == mailbox_name.strip().lower():
                root = store.GetRootFolder()
                break
        if root is None:
            available = [s.DisplayName for s in namespace.Stores]
            raise KeyError(
                f"Mailbox '{mailbox_name}' not found. "
                f"Available stores: {available}"
            )
    else:
        root = namespace.DefaultStore.GetRootFolder()

    # ---- walk the path -----------------------------------------------------
    current_folder = root
    for segment in folder_path:
        matched = None
        try:
            sub_folders = current_folder.Folders
        except Exception as exc:
            raise KeyError(
                f"Cannot list sub-folders of '{current_folder.Name}': {exc}"
            ) from exc

        for i in range(1, sub_folders.Count + 1):
            try:
                sf = sub_folders.Item(i)
                if sf.Name.strip().lower() == segment.strip().lower():
                    matched = sf
                    break
            except Exception:
                continue

        if matched is None:
            available = []
            for i in range(1, sub_folders.Count + 1):
                try:
                    available.append(sub_folders.Item(i).Name)
                except Exception:
                    pass
            raise KeyError(
                f"Sub-folder '{segment}' not found under '{current_folder.Name}'. "
                f"Available: {available}"
            )
        current_folder = matched

    log.info("Resolved folder: %s", " > ".join(folder_path))
    return current_folder


# ---------------------------------------------------------------------------
# 3. Fetch messages from a specific day
# ---------------------------------------------------------------------------

def fetch_messages_from_day(
    folder,
    days_ago: int,
    reference_date: Optional[datetime] = None,
) -> List:
    """
    Return all MailItem objects whose ``ReceivedTime`` falls on the calendar day
    that is exactly ``days_ago`` days before ``reference_date`` (default: today).

    Filtering is done in Python to avoid Restrict() locale / timezone pitfalls.

    Parameters
    ----------
    folder:
        Outlook Folder COM object.
    days_ago:
        How many days back to look (0 = today, 1 = yesterday, …).
    reference_date:
        The anchor date. Defaults to ``datetime.now()`` (local time).

    Returns
    -------
    List of Outlook MailItem COM objects.
    """
    if reference_date is None:
        reference_date = datetime.now()

    target_date = (reference_date - timedelta(days=days_ago)).date()
    log.info("Filtering emails received on: %s", target_date)

    mail_items = []
    items = folder.Items
    total = items.Count
    log.info("Total items in folder: %d", total)

    for i in range(1, total + 1):
        try:
            item = items.Item(i)
        except Exception as exc:
            log.warning("Could not access item %d: %s", i, exc)
            continue

        # Only process genuine mail items (Class == 43)
        try:
            if item.Class != 43:
                continue
        except Exception:
            continue

        # Safely read ReceivedTime
        try:
            received = item.ReceivedTime
            if received is None:
                continue
            # COM dates come back as Python datetime objects via pywin32
            if hasattr(received, "date"):
                received_date = received.date()
            else:
                # Fallback: parse string representation
                received_date = datetime.strptime(
                    str(received)[:10], "%Y-%m-%d"
                ).date()
        except Exception as exc:
            log.debug("Skipping item %d – bad ReceivedTime: %s", i, exc)
            continue

        if received_date == target_date:
            mail_items.append(item)

    log.info("Found %d matching mail item(s) for %s.", len(mail_items), target_date)
    return mail_items


# ---------------------------------------------------------------------------
# 4. HTML table extraction
# ---------------------------------------------------------------------------

def try_read_html_tables(html_body: str) -> List[pd.DataFrame]:
    """
    Attempt to parse all HTML tables from ``html_body`` using pandas + lxml/html5lib.

    Returns a (possibly empty) list of raw DataFrames – one per parseable table.
    Malformed or empty tables are silently skipped.

    Parameters
    ----------
    html_body:
        Raw HTML string from ``item.HTMLBody``.

    Returns
    -------
    List[pd.DataFrame]
    """
    if not html_body or not html_body.strip():
        return []

    try:
        tables = pd.read_html(html_body, flavor="lxml")
    except Exception:
        try:
            tables = pd.read_html(html_body, flavor="html5lib")
        except Exception as exc:
            log.debug("Could not parse HTML tables: %s", exc)
            return []

    valid: List[pd.DataFrame] = []
    for df in tables:
        if df is None or df.empty:
            continue
        if df.shape[0] < 1 or df.shape[1] < 2:
            # Need at least 1 data row and 2 columns (attribute + one value)
            continue
        valid.append(df)

    return valid


# ---------------------------------------------------------------------------
# 5. Table normalisation helpers
# ---------------------------------------------------------------------------

_WHITESPACE_RE = re.compile(r"\s+")


def _clean_column_name(raw: str) -> str:
    """Lowercase, strip, replace inner whitespace with underscores."""
    s = str(raw).strip()
    s = _WHITESPACE_RE.sub("_", s)
    return s.lower()


def normalize_table(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Normalise a raw DataFrame produced by ``try_read_html_tables``.

    Assumes layout:
        column 0  → attribute / field names
        column 1+ → one security per column

    Steps:
    1. Drop fully-NaN rows and columns.
    2. Fill remaining NaN with empty string.
    3. Use the first column as the attribute index.
    4. Return None if the table is unusable after cleaning.

    Parameters
    ----------
    df:
        Raw DataFrame from ``pd.read_html``.

    Returns
    -------
    Normalised DataFrame or None.
    """
    # Drop columns that are entirely NaN
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    if df.empty or df.shape[1] < 2:
        return None

    # Reset column names to simple integers for predictable indexing
    df = df.reset_index(drop=True)
    df.columns = list(range(df.shape[1]))

    # Fill NaN → empty string
    df = df.fillna("")

    # Convert all values to str
    df = df.astype(str)

    # Trim whitespace everywhere
    df = df.applymap(lambda x: x.strip())  # type: ignore[attr-defined]

    return df


# ---------------------------------------------------------------------------
# 6. Convert one normalised table → security rows
# ---------------------------------------------------------------------------

def table_to_security_rows(
    df: pd.DataFrame,
    table_index: int,
    email_meta: Dict,
) -> List[Dict]:
    """
    Convert a normalised table into a list of flat row dicts, one per security column.

    Parameters
    ----------
    df:
        Normalised DataFrame (output of ``normalize_table``).
    table_index:
        Zero-based position of this table within the email.
    email_meta:
        Dict with keys: ``entry_id``, ``subject``, ``received_time``.

    Returns
    -------
    List of dicts, each representing one security.
    """
    rows: List[Dict] = []

    # Column 0 = attribute names; columns 1..N = securities
    attr_col = df[0].tolist()
    num_security_cols = df.shape[1] - 1

    # Normalise attribute names → column headers for the output row
    norm_attrs = [_clean_column_name(a) if a else f"field_{i}"
                  for i, a in enumerate(attr_col)]

    # Detect optional column headers (first row of df might be a header row for
    # the security columns, e.g. ISIN / ticker).  Heuristic: if the first data
    # row has no attribute name it is likely a header row.
    security_headers: List[str] = [""] * num_security_cols
    data_start_row = 0

    if attr_col and not attr_col[0].strip():
        # First row looks like headers for security columns
        for col_idx in range(1, df.shape[1]):
            security_headers[col_idx - 1] = str(df.iloc[0, col_idx]).strip()
        data_start_row = 1

    for sec_idx in range(num_security_cols):
        col_idx = sec_idx + 1  # column index in df

        record: Dict = {
            **email_meta,
            "table_index": table_index,
            "security_index": sec_idx,
            "security_header": security_headers[sec_idx],
        }

        for row_idx in range(data_start_row, len(attr_col)):
            attr_name = norm_attrs[row_idx] or f"field_{row_idx}"
            try:
                value = df.iloc[row_idx, col_idx]
            except IndexError:
                value = ""
            record[attr_name] = value

        rows.append(record)

    return rows


# ---------------------------------------------------------------------------
# 7. Main orchestration function
# ---------------------------------------------------------------------------

def parse_outlook_folder_to_dataframe(
    folder_path: Tuple[str, ...],
    days_ago: int,
    mailbox_name: Optional[str] = None,
    reference_date: Optional[datetime] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    High-level entry point.  Connects to Outlook, resolves the folder, fetches
    emails from ``days_ago`` days back, parses HTML tables, and consolidates
    everything into one output DataFrame.

    Parameters
    ----------
    folder_path:
        Tuple of nested folder names, e.g. ``("Inbox", "MCBox", "Trades", "EMail")``.
    days_ago:
        How many calendar days back to look (0 = today, 1 = yesterday, …).
    mailbox_name:
        Optional shared/delegate mailbox display name.
    reference_date:
        Override the anchor date (defaults to ``datetime.now()``).

    Returns
    -------
    (results_df, summary_df)

    results_df:
        One row per security per table per email.  Columns include email metadata
        (entry_id, subject, received_time, table_index, security_index,
        security_header) plus all normalised attribute columns.

    summary_df:
        One row per processed email showing: subject, received_time, entry_id,
        tables_found, tables_parsed, total_security_rows.
    """
    namespace = get_outlook_namespace()
    folder = get_folder(namespace, folder_path, mailbox_name=mailbox_name)
    messages = fetch_messages_from_day(folder, days_ago, reference_date=reference_date)

    all_rows: List[Dict] = []
    summary_rows: List[Dict] = []

    for msg in messages:
        # ---- extract metadata ----------------------------------------------
        try:
            entry_id = msg.EntryID
        except Exception:
            entry_id = ""

        try:
            subject = msg.Subject or ""
        except Exception:
            subject = ""

        try:
            received_time = msg.ReceivedTime
        except Exception:
            received_time = None

        log.info("Processing: '%s' (%s)", subject, received_time)

        email_meta: Dict = {
            "entry_id": entry_id,
            "subject": subject,
            "received_time": received_time,
        }

        # ---- read HTML body ------------------------------------------------
        try:
            html_body = msg.HTMLBody or ""
        except Exception:
            html_body = ""

        if not html_body.strip():
            log.debug("  No HTML body found, skipping.")
            summary_rows.append(
                {**email_meta, "tables_found": 0, "tables_parsed": 0, "total_security_rows": 0}
            )
            continue

        # ---- parse tables --------------------------------------------------
        raw_tables = try_read_html_tables(html_body)
        tables_found = len(raw_tables)
        tables_parsed = 0
        security_rows_count = 0

        for tbl_idx, raw_df in enumerate(raw_tables):
            norm_df = normalize_table(raw_df)
            if norm_df is None:
                log.debug("  Table %d: could not normalise, skipping.", tbl_idx)
                continue

            rows = table_to_security_rows(norm_df, tbl_idx, email_meta)
            if rows:
                all_rows.extend(rows)
                tables_parsed += 1
                security_rows_count += len(rows)
                log.debug(
                    "  Table %d: %d security row(s) extracted.", tbl_idx, len(rows)
                )

        summary_rows.append(
            {
                **email_meta,
                "tables_found": tables_found,
                "tables_parsed": tables_parsed,
                "total_security_rows": security_rows_count,
            }
        )

    # ---- build output DataFrames -------------------------------------------
    if all_rows:
        results_df = pd.DataFrame(all_rows)
    else:
        results_df = pd.DataFrame(
            columns=[
                "entry_id", "subject", "received_time",
                "table_index", "security_index", "security_header",
            ]
        )

    summary_df = pd.DataFrame(summary_rows) if summary_rows else pd.DataFrame(
        columns=[
            "entry_id", "subject", "received_time",
            "tables_found", "tables_parsed", "total_security_rows",
        ]
    )

    log.info(
        "Done. %d email(s) processed, %d total output row(s).",
        len(summary_rows),
        len(all_rows),
    )
    return results_df, summary_df


# ---------------------------------------------------------------------------
# 8. CLI / quick-run entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # -----------------------------------------------------------------------
    # CONFIG – edit these values before running
    # -----------------------------------------------------------------------
    CONFIG: Dict = {
        # Nested folder path inside Outlook
        "folder_path": ("Inbox", "MCBox", "Trades", "EMail"),

        # How many days back (0 = today, 1 = yesterday, …)
        "days_ago": 1,

        # Set to None to use your default/primary mailbox
        # Set to a string like "shared@company.com" for a shared/delegate mailbox
        "mailbox_name": None,
    }
    # -----------------------------------------------------------------------

    results, summary = parse_outlook_folder_to_dataframe(
        folder_path=CONFIG["folder_path"],
        days_ago=CONFIG["days_ago"],
        mailbox_name=CONFIG["mailbox_name"],
    )

    print("\n========== SUMMARY ==========")
    print(summary.to_string(index=False))

    print("\n========== RESULTS (first 10 rows) ==========")
    print(results.head(10).to_string(index=False))

    # Optionally save to CSV
    # results.to_csv("parsed_securities.csv", index=False)
    # summary.to_csv("email_summary.csv", index=False)
