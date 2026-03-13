from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time, timedelta
from io import StringIO
from typing import Any

import pandas as pd

try:
    import win32com.client  # type: ignore[import-untyped]
except ImportError as exc:  # pragma: no cover - import failure is environment-specific
    raise ImportError(
        "pywin32 is required. Install it with: pip install pywin32"
    ) from exc


@dataclass
class OutlookTableParserConfig:
    mailbox: str | None = None
    folder_path: tuple[str, ...] = ("Inbox",)
    days_ago: int = 1
    outlook_profile: str | None = None
    include_empty_tables: bool = False


def get_outlook_namespace(profile_name: str | None = None) -> Any:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if profile_name:
        namespace.Logon(Profile=profile_name)
    return namespace


def get_folder(namespace: Any, mailbox: str | None, folder_path: tuple[str, ...]) -> Any:
    if not folder_path:
        raise ValueError("folder_path must contain at least one folder name")

    if mailbox:
        folder = namespace.Folders.Item(mailbox)
        parts_to_walk = folder_path
    else:
        folder = namespace.GetDefaultFolder(6)
        if folder.Name != folder_path[0]:
            if folder_path[0].lower() == "inbox":
                parts_to_walk = folder_path[1:]
            else:
                folder = namespace.Folders.Item(folder_path[0])
                parts_to_walk = folder_path[1:]
        else:
            parts_to_walk = folder_path[1:]

    for part in parts_to_walk:
        if part.lower() == "inbox" and getattr(folder, "Name", "").lower() == "inbox":
            continue
        folder = folder.Folders.Item(part)

    return folder


def get_day_bounds(days_ago: int) -> tuple[datetime, datetime]:
    if days_ago < 0:
        raise ValueError("days_ago must be >= 0")

    today = datetime.now().date()
    target_day = today - timedelta(days=days_ago)
    start = datetime.combine(target_day, time.min)
    end = start + timedelta(days=1)
    return start, end


def format_outlook_datetime(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")


def fetch_messages_from_day(folder: Any, days_ago: int) -> list[Any]:
    start, end = get_day_bounds(days_ago)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    restriction = (
        f"[ReceivedTime] >= '{format_outlook_datetime(start)}' AND "
        f"[ReceivedTime] < '{format_outlook_datetime(end)}'"
    )
    restricted = items.Restrict(restriction)
    return [item for item in restricted if getattr(item, "Class", None) == 43]


def try_read_html_tables(html: str) -> list[pd.DataFrame]:
    if not html or "<table" not in html.lower():
        return []

    try:
        tables = pd.read_html(StringIO(html))
    except ValueError:
        return []

    cleaned: list[pd.DataFrame] = []
    for table in tables:
        normalized = normalize_table(table)
        if normalized is not None:
            cleaned.append(normalized)
    return cleaned


def normalize_table(table: pd.DataFrame) -> pd.DataFrame | None:
    if table.empty or table.shape[1] < 2:
        return None

    working = table.copy()
    working = working.dropna(axis=0, how="all").dropna(axis=1, how="all")
    if working.empty or working.shape[1] < 2:
        return None

    working.columns = [str(col).strip() for col in working.columns]
    working = working.applymap(clean_cell_value)

    # Some email tables repeat header labels as the first row.
    first_row = [str(value).strip().lower() for value in working.iloc[0].tolist()]
    if working.shape[0] > 1 and first_row == [str(col).strip().lower() for col in working.columns]:
        working = working.iloc[1:].reset_index(drop=True)
        if working.empty or working.shape[1] < 2:
            return None

    first_col = working.iloc[:, 0].astype(str).str.strip()
    non_empty_ratio = (first_col != "").mean()
    if non_empty_ratio < 0.5:
        return None

    return working.reset_index(drop=True)


def clean_cell_value(value: Any) -> Any:
    if pd.isna(value):
        return None

    text = str(value).replace("\xa0", " ").strip()
    return text if text else None


def table_to_security_rows(
    table: pd.DataFrame,
    message_subject: str,
    received_time: datetime,
    entry_id: str,
    table_index: int,
) -> list[dict[str, Any]]:
    attribute_series = table.iloc[:, 0].astype(str).str.strip()
    records: list[dict[str, Any]] = []

    for security_position in range(1, table.shape[1]):
        security_values = table.iloc[:, security_position]
        row_data = {
            "message_entry_id": entry_id,
            "message_subject": message_subject,
            "received_time": received_time,
            "table_index": table_index,
            "security_index": security_position,
            "security_column_header": infer_security_header(table.columns[security_position]),
        }

        has_value = False
        for attribute_name, attribute_value in zip(attribute_series, security_values):
            if not attribute_name or attribute_name.lower() == "nan":
                continue

            normalized_name = make_column_name(attribute_name)
            row_data[normalized_name] = attribute_value
            has_value = has_value or attribute_value not in (None, "")

        if has_value:
            records.append(row_data)

    return records


def infer_security_header(value: Any) -> str | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text or text.lower().startswith("unnamed:"):
        return None
    return text


def make_column_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() else "_" for ch in value.strip().lower())
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.strip("_") or "attribute"


def parse_outlook_folder_to_dataframe(
    config: OutlookTableParserConfig,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    namespace = get_outlook_namespace(config.outlook_profile)
    folder = get_folder(namespace, config.mailbox, config.folder_path)
    messages = fetch_messages_from_day(folder, config.days_ago)

    parsed_rows: list[dict[str, Any]] = []
    message_summaries: list[dict[str, Any]] = []

    for message in messages:
        subject = getattr(message, "Subject", "") or ""
        received_time = getattr(message, "ReceivedTime", None)
        entry_id = getattr(message, "EntryID", "")
        html_body = getattr(message, "HTMLBody", "") or ""

        tables = try_read_html_tables(html_body)
        message_summaries.append(
            {
                "entry_id": entry_id,
                "subject": subject,
                "received_time": received_time,
                "table_count": len(tables),
            }
        )

        if not tables and config.include_empty_tables:
            parsed_rows.append(
                {
                    "message_entry_id": entry_id,
                    "message_subject": subject,
                    "received_time": received_time,
                    "table_index": None,
                    "security_index": None,
                    "security_column_header": None,
                }
            )
            continue

        for table_index, table in enumerate(tables, start=1):
            parsed_rows.extend(
                table_to_security_rows(
                    table=table,
                    message_subject=subject,
                    received_time=received_time,
                    entry_id=entry_id,
                    table_index=table_index,
                )
            )

    return pd.DataFrame(parsed_rows), message_summaries


if __name__ == "__main__":
    config = OutlookTableParserConfig(
        mailbox=None,
        folder_path=("Inbox", "YourSubfolder"),
        days_ago=1,
    )

    df, messages = parse_outlook_folder_to_dataframe(config)

    print("Messages checked:")
    print(pd.DataFrame(messages))
    print()
    print("Parsed output:")
    print(df)
