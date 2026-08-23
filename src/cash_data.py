"""Cash-count data handling, kept free of tkinter so it can be tested.

The second-offering subtraction lives here. The BC-40 reports cumulative
totals within a session, so the 2차 figures are (second report - first report)
per denomination. Getting that wrong is silent: the numbers still look
plausible, they are just the wrong numbers. Hence the guards below.
"""

import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Sequence

import pandas as pd

# Row order of the denomination block in the report template (A7:A13).
DENOMINATIONS: Sequence[int] = (1, 2, 5, 10, 20, 50, 100)

SNAPSHOT_PREFIX = ".offering_1차_"


def pdf_to_dataframe(pdf_path: str) -> pd.DataFrame:
    """Extract the single-page BC-40 table into a DataFrame."""
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        table = pdf.pages[0].extract_table()
    if not table:
        raise ValueError(f"No table found in {pdf_path}")
    if len(table) < 5:
        raise ValueError(f"Unexpected table shape in {pdf_path}: {len(table)} rows")
    return pd.DataFrame(table[4:], columns=table[3])


def _to_int(value) -> Optional[int]:
    """Parse a BC-40 cell to int, tolerating '1,920', '$50', spaces and None."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    text = re.sub(r"[^\d\-]", "", str(value))
    if not text or text == "-":
        return None
    try:
        return int(text)
    except ValueError:
        return None


def clean_denomination_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise the raw PDF table to integer DENO/QTY/AMT rows.

    Drops the TOTAL row by testing whether DENO parses as a number, rather
    than by chopping the last row - which silently ate a real denomination
    whenever the machine omitted the total.
    """
    required = {"DENO", "QTY", "AMT"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"PDF table is missing column(s): {sorted(missing)}")

    out = df.copy()
    for column in ("DENO", "QTY", "AMT"):
        out[column] = out[column].map(_to_int)
    out = out[out["DENO"].notna()]
    out = out[out[["QTY", "AMT"]].notna().all(axis=1)]
    out = out.astype({"DENO": int, "QTY": int, "AMT": int})
    return out.sort_values("DENO").reset_index(drop=True)[["DENO", "QTY", "AMT"]]


def align_to_denominations(df: pd.DataFrame,
                           denominations: Sequence[int] = DENOMINATIONS) -> pd.DataFrame:
    """Force one row per template denomination, in template order.

    Without this, a report that omits zero-count denominations shifts every
    subsequent row up by one and the quantities land beside the wrong labels.
    """
    indexed = df.set_index("DENO")
    rows = []
    for denomination in denominations:
        if denomination in indexed.index:
            row = indexed.loc[denomination]
            rows.append({"DENO": denomination, "QTY": int(row["QTY"]), "AMT": int(row["AMT"])})
        else:
            rows.append({"DENO": denomination, "QTY": 0, "AMT": 0})
    return pd.DataFrame(rows)


@dataclass
class SubtractionResult:
    frame: pd.DataFrame
    warnings: List[str] = field(default_factory=list)
    negative_denominations: List[int] = field(default_factory=list)

    @property
    def looks_wrong(self) -> bool:
        return bool(self.negative_denominations)


def subtract_offering(current: pd.DataFrame, previous: pd.DataFrame) -> SubtractionResult:
    """Derive the 2차 figures as current - previous, per denomination.

    A negative result means the assumption broke - almost always because the
    machine was cleared between offerings, so the second report is standalone
    rather than cumulative. That is surfaced instead of being written out.
    """
    # Callers pass frames that have already been through
    # clean_denomination_frame(); aligning both guarantees the two sides line
    # up denomination-for-denomination before subtracting.
    current = align_to_denominations(current)
    previous = align_to_denominations(previous)

    merged = current.merge(previous, on="DENO", how="left", suffixes=("", "_prev"))
    merged[["QTY_prev", "AMT_prev"]] = merged[["QTY_prev", "AMT_prev"]].fillna(0).astype(int)
    merged["QTY"] = merged["QTY"] - merged["QTY_prev"]
    merged["AMT"] = merged["AMT"] - merged["AMT_prev"]

    result = merged[["DENO", "QTY", "AMT"]].reset_index(drop=True)
    negatives = sorted(int(d) for d in result.loc[result["QTY"] < 0, "DENO"])

    warnings = []
    if negatives:
        warnings.append(
            "2차 계산 결과가 음수입니다 (해당 권종: "
            + ", ".join(f"${d}" for d in negatives)
            + "). 기계가 1차와 2차 사이에 초기화된 것으로 보입니다."
        )
    elif int(result["QTY"].sum()) == 0:
        warnings.append("2차 헌금 수량이 0입니다. 2차 PDF가 새로 생성되지 않았을 수 있습니다.")

    return SubtractionResult(frame=result, warnings=warnings, negative_denominations=negatives)


# --------------------------------------------------------------------------
# Snapshot persistence
# --------------------------------------------------------------------------

def snapshot_path(output_folder: str, run_date: str, mass_time: str) -> str:
    """Where the 1차 counts are parked between the two offerings."""
    safe = re.sub(r"[^\w가-힣]", "", mass_time or "")
    return os.path.join(output_folder, f"{SNAPSHOT_PREFIX}{run_date}_{safe}.json")


def save_offering_snapshot(path: str, df: pd.DataFrame, pdf_file: str = "") -> None:
    """Persist the 1차 counts so a restart cannot lose them.

    The original code held these in a module global only, so quitting the app
    between offerings meant the 2차 subtraction silently did not happen and the
    cumulative totals were written as if they were the second offering.
    """
    payload = {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "source_pdf": pdf_file,
        "rows": df[["DENO", "QTY", "AMT"]].astype(int).to_dict(orient="records"),
    }
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def load_offering_snapshot(path: str):
    """Return (DataFrame, metadata) or (None, {}) when there is no snapshot."""
    if not path or not os.path.exists(path):
        return None, {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        frame = pd.DataFrame(payload["rows"])[["DENO", "QTY", "AMT"]].astype(int)
    except (ValueError, KeyError, TypeError, OSError):
        return None, {}
    return frame, {k: v for k, v in payload.items() if k != "rows"}


def write_offering_block(sheet, frame: pd.DataFrame, qty_col: int, start_row: int) -> int:
    """Write QTY/AMT down two adjacent columns, one row per denomination.

    Rows are positional, so the frame must already be aligned with
    align_to_denominations() or the values land beside the wrong labels.
    Returns the number of rows written.
    """
    for offset, row in enumerate(frame.itertuples(index=False)):
        sheet.cell(row=start_row + offset, column=qty_col, value=int(row.QTY))
        sheet.cell(row=start_row + offset, column=qty_col + 1, value=int(row.AMT))
    return len(frame)
