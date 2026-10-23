"""Parse the uploaded team/project spreadsheet into per-team deck data.

Expected spreadsheet shape: "long" format, one row per team per project.

Required columns (case/space-insensitive):
    team_name     - who the deck is for
    status        - "in_progress" or "completed"
    project_name  - project name

Optional columns, used when present:
    owner, start_date
    in_progress only: target_date, percent_complete, notes
    completed only:   completed_date, outcome

Any unrecognized columns are ignored.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field

import pandas as pd

from . import formatting as fmt

REQUIRED_COLUMNS = {"team_name", "status", "project_name"}


class SpreadsheetError(ValueError):
    """Raised when the uploaded spreadsheet can't be parsed or is missing required columns."""


@dataclass
class TeamDeckData:
    team_name: str
    in_progress_rows: list[dict] = field(default_factory=list)
    completed_rows: list[dict] = field(default_factory=list)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip().lower().replace(" ", "_").replace("-", "_") for c in df.columns
    ]
    return df


def _read_excel_data_sheet(buf: io.BytesIO) -> pd.DataFrame:
    """Multi-sheet workbooks (e.g. our template, which has an Instructions tab
    before the data tab) shouldn't silently read the wrong sheet."""
    xls = pd.ExcelFile(buf)
    sheet_names = xls.sheet_names
    if len(sheet_names) == 1:
        return pd.read_excel(xls, sheet_name=sheet_names[0])

    for name in sheet_names:
        if name.strip().lower() == "projects":
            return pd.read_excel(xls, sheet_name=name)

    for name in sheet_names:
        candidate = _normalize_columns(pd.read_excel(xls, sheet_name=name))
        if REQUIRED_COLUMNS.issubset(set(candidate.columns)):
            return candidate

    return pd.read_excel(xls, sheet_name=sheet_names[0])


def load_spreadsheet(filename: str, content: bytes) -> pd.DataFrame:
    buf = io.BytesIO(content)
    lower = filename.lower()
    try:
        if lower.endswith(".csv"):
            df = pd.read_csv(buf)
        elif lower.endswith((".xlsx", ".xlsm", ".xls")):
            df = _read_excel_data_sheet(buf)
        else:
            raise SpreadsheetError(
                f"Unsupported file type for '{filename}'. Please upload a .xlsx or .csv file."
            )
    except SpreadsheetError:
        raise
    except Exception as exc:  # pandas raises many different error types
        raise SpreadsheetError(f"Could not read '{filename}': {exc}") from exc

    df = _normalize_columns(df)
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise SpreadsheetError(
            "Spreadsheet is missing required column(s): "
            + ", ".join(sorted(missing))
            + ". Found columns: "
            + ", ".join(df.columns)
        )
    return df


def build_team_decks(df: pd.DataFrame) -> tuple[dict[str, TeamDeckData], list[str]]:
    """Returns (teams_by_name, warnings). Rows with bad/missing required
    fields are skipped and reported as warnings rather than aborting the run."""
    teams: dict[str, TeamDeckData] = {}
    warnings: list[str] = []

    for idx, row in df.iterrows():
        team_name = fmt.format_text(row.get("team_name"))
        if not team_name:
            warnings.append(f"Row {idx + 2}: missing team_name, skipped.")
            continue
        status = fmt.format_text(row.get("status")).lower().replace(" ", "_")
        project_name = fmt.format_text(row.get("project_name"))
        if status not in ("in_progress", "completed"):
            warnings.append(
                f"Row {idx + 2} ({team_name} / {project_name or 'unnamed project'}): "
                f"status must be 'in_progress' or 'completed', got '{status or '(blank)'}'. Skipped."
            )
            continue

        deck = teams.setdefault(team_name, TeamDeckData(team_name))

        common = {
            "project_name": project_name,
            "owner": fmt.format_text(row.get("owner")),
            "start_date": fmt.format_date(row.get("start_date")),
        }

        if status == "in_progress":
            common.update(
                {
                    "target_date": fmt.format_date(row.get("target_date")),
                    "percent_complete": fmt.format_percent(row.get("percent_complete")),
                    "notes": fmt.format_text(row.get("notes")),
                }
            )
            deck.in_progress_rows.append(common)
        else:
            common.update(
                {
                    "completed_date": fmt.format_date(row.get("completed_date")),
                    "outcome": fmt.format_text(row.get("outcome")),
                }
            )
            deck.completed_rows.append(common)

    if not teams:
        raise SpreadsheetError(
            "No usable rows found.\n" + "\n".join(warnings)
            if warnings
            else "No usable rows found in the spreadsheet."
        )

    return teams, warnings
