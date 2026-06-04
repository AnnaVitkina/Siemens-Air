"""
Extract Air Rate Card and Airport Zones from a DHL Global Air Fact Sheet workbook.

Prompts for the source file and sheet tabs, then writes a consolidated workbook
to the processing folder with Rate Card and Zones tabs.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from project_paths import INPUT_DIR, PROCESSING_DIR, ensure_project_dirs

DEFAULT_SHEETS = {
    "rate_card": "air rate card",
    "zones": "Airport zones GAT",
}

OUTPUT_SHEET_NAMES = {
    "rate_card": "Rate Card",
    "zones": "Zones",
}

RATE_CARD_HEADER_MARKERS = (
    "country_contract_owner",
    "origin_country",
    "forwarder_name",
)

RATE_CARD_COLUMN_RENAMES = {
    "origin_airport_code": "Origin Zone",
    "destination_airport_code": "Destination Zone",
    "siemens_division": "business unit code",
    "siemens_business": "siemens division",
    "siemens_buisness": "siemens division",
}


@dataclass
class ExtractionResult:
    rate_card: pd.DataFrame
    zones: pd.DataFrame
    source_file: Path
    output_file: Path
    rate_card_header_row: int
    zones_header_row: int


def list_input_files() -> list[Path]:
    ensure_project_dirs()
    return sorted(INPUT_DIR.glob("*.xlsx")) + sorted(INPUT_DIR.glob("*.xls"))


def list_workbook_sheets(file_path: Path) -> list[str]:
    return pd.ExcelFile(file_path).sheet_names


def resolve_default_sheet(sheet_names: list[str], default_name: str) -> str | None:
    if default_name in sheet_names:
        return default_name

    lowered = {name.lower(): name for name in sheet_names}
    if default_name.lower() in lowered:
        return lowered[default_name.lower()]

    default_lower = default_name.lower()
    for name in sheet_names:
        if default_lower in name.lower():
            return name

    return None


def prompt_choice(prompt: str, options: list[str], default: str | None = None) -> str:
    print(f"\n{prompt}")
    for index, option in enumerate(options, start=1):
        marker = " (default)" if default and option == default else ""
        print(f"  {index}. {option}{marker}")

    while True:
        raw = input("Enter number or exact name: ").strip()
        if not raw and default:
            return default
        if raw.isdigit():
            choice_index = int(raw) - 1
            if 0 <= choice_index < len(options):
                return options[choice_index]
        if raw in options:
            return raw
        print("Invalid choice. Try again.")


def prompt_sheet_name(
    sheet_names: list[str],
    label: str,
    default_name: str,
) -> str:
    suggested = resolve_default_sheet(sheet_names, default_name)

    print(f"\n{label}")
    print("Available tabs:")
    for index, name in enumerate(sheet_names, start=1):
        marker = "  <-- auto-selected" if suggested and name == suggested else ""
        print(f"  {index}. {name}{marker}")

    if suggested:
        print(
            f'\nAuto-selected: "{suggested}" '
            f'(tab name contains "{default_name}")'
        )
        prompt_text = "Press Enter to confirm, or enter tab number: "
    else:
        print(f'\nNo tab matched "{default_name}".')
        prompt_text = "Enter tab number: "

    while True:
        raw = input(prompt_text).strip()
        if not raw:
            if suggested:
                return suggested
            print("Please enter a tab number.")
            continue
        if raw.isdigit():
            choice_index = int(raw) - 1
            if 0 <= choice_index < len(sheet_names):
                return sheet_names[choice_index]
        print("Invalid choice. Enter a valid tab number.")


def _row_contains_marker(row: pd.Series, markers: tuple[str, ...]) -> bool:
    for value in row:
        if pd.isna(value):
            continue
        text = str(value).strip().lower()
        if any(marker in text for marker in markers):
            return True
    return False


def rename_rate_card_columns(columns: pd.Index | list[str]) -> list[str]:
    renamed: list[str] = []
    for column in columns:
        name = str(column).strip()
        mapped = RATE_CARD_COLUMN_RENAMES.get(name.lower())
        if mapped is not None:
            renamed.append(mapped)
        else:
            renamed.append(name.replace("_", " "))
    return renamed


def find_rate_card_header_row(df_raw: pd.DataFrame) -> int:
    for row_index in range(len(df_raw)):
        if _row_contains_marker(df_raw.iloc[row_index], RATE_CARD_HEADER_MARKERS):
            return row_index
    raise ValueError(
        "Could not find the Rate Card header row "
        f"(expected one of: {', '.join(RATE_CARD_HEADER_MARKERS)})."
    )


def extract_rate_card_df(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    header_row = find_rate_card_header_row(df_raw)
    headers = df_raw.iloc[header_row].tolist()
    column_names: list[str] = []
    seen: dict[str, int] = {}

    for column_index, header in enumerate(headers):
        if pd.isna(header) or str(header).strip() == "":
            name = f"Column_{column_index + 1}"
        else:
            name = str(header).strip()

        if name not in seen:
            seen[name] = 1
            column_names.append(name)
            continue

        seen[name] += 1
        column_names.append(f"{name}_{seen[name]}")

    rate_card = df_raw.iloc[header_row + 1 :].copy()
    rate_card.columns = rename_rate_card_columns(column_names)
    rate_card = rate_card.dropna(how="all").reset_index(drop=True)
    return rate_card, header_row


def find_zones_header_row(df_raw: pd.DataFrame) -> int:
    for row_index in range(min(10, len(df_raw))):
        first_cell = df_raw.iloc[row_index, 0]
        if pd.isna(first_cell):
            continue
        text = str(first_cell).strip().lower()
        if "airport" in text and "code" in text:
            return row_index
    return 0


def extract_zones_df(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    header_row = find_zones_header_row(df_raw)
    header_values = df_raw.iloc[header_row].tolist()

    column_names: list[str] = []
    for column_index, header in enumerate(header_values):
        if pd.isna(header) or str(header).strip() == "":
            column_names.append(f"Column_{column_index + 1}")
        else:
            column_names.append(str(header).strip())

    zones = df_raw.iloc[header_row + 1 :].copy()
    zones.columns = column_names
    zones = zones.dropna(how="all").reset_index(drop=True)

    primary_columns = [
        name
        for name in zones.columns
        if zones[name].notna().any()
    ]
    if primary_columns:
        zones = zones[primary_columns]

    airport_col = zones.columns[0]
    zone_col = zones.columns[1] if len(zones.columns) > 1 else None

    zones = zones[zones[airport_col].notna()].copy()
    if zone_col is not None:
        zones = zones[zones[zone_col].notna()].copy()

    zones = zones.reset_index(drop=True)
    return zones, header_row


def extract_air_workbook(
    file_path: Path,
    sheet_mapping: dict[str, str],
) -> ExtractionResult:
    rate_card_raw = pd.read_excel(
        file_path,
        sheet_name=sheet_mapping["rate_card"],
        header=None,
    )
    zones_raw = pd.read_excel(
        file_path,
        sheet_name=sheet_mapping["zones"],
        header=None,
    )

    rate_card_df, rate_card_header_row = extract_rate_card_df(rate_card_raw)
    zones_df, zones_header_row = extract_zones_df(zones_raw)

    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    output_file = PROCESSING_DIR / f"{file_path.stem}_extracted.xlsx"

    metadata_df = pd.DataFrame(
        [
            {"Field": "Source file", "Value": file_path.name},
            {"Field": "Rate Card tab", "Value": sheet_mapping["rate_card"]},
            {"Field": "Zones tab", "Value": sheet_mapping["zones"]},
            {"Field": "Rate Card header row (0-based)", "Value": rate_card_header_row},
            {"Field": "Zones header row (0-based)", "Value": zones_header_row},
            {"Field": "Rate Card rows", "Value": len(rate_card_df)},
            {"Field": "Zones rows", "Value": len(zones_df)},
        ]
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        rate_card_df.to_excel(
            writer,
            sheet_name=OUTPUT_SHEET_NAMES["rate_card"],
            index=False,
        )
        zones_df.to_excel(
            writer,
            sheet_name=OUTPUT_SHEET_NAMES["zones"],
            index=False,
        )
        metadata_df.to_excel(writer, sheet_name="Metadata", index=False)

    return ExtractionResult(
        rate_card=rate_card_df,
        zones=zones_df,
        source_file=file_path,
        output_file=output_file,
        rate_card_header_row=rate_card_header_row,
        zones_header_row=zones_header_row,
    )


def run_interactive_extraction() -> ExtractionResult:
    input_files = list_input_files()
    if not input_files:
        raise FileNotFoundError(f"No Excel files found in {INPUT_DIR}")

    print("Files available in the input folder:")
    selected_file = prompt_choice(
        "Which file should be processed?",
        [path.name for path in input_files],
        default=input_files[0].name if len(input_files) == 1 else None,
    )
    file_path = INPUT_DIR / selected_file

    sheet_names = list_workbook_sheets(file_path)
    sheet_mapping = {
        "rate_card": prompt_sheet_name(
            sheet_names,
            "Rate Card (look for tab containing: air rate card)",
            DEFAULT_SHEETS["rate_card"],
        ),
        "zones": prompt_sheet_name(
            sheet_names,
            "Zones (look for tab containing: Airport zones GAT)",
            DEFAULT_SHEETS["zones"],
        ),
    }

    result = extract_air_workbook(file_path, sheet_mapping)

    print(f"\nSaved extracted workbook to: {result.output_file}")
    print(f"Rate Card rows: {len(result.rate_card)}")
    print(f"Zones rows: {len(result.zones)}")
    print(f"Rate Card header row: {result.rate_card_header_row}")
    print(f"Zones header row: {result.zones_header_row}")

    return result


if __name__ == "__main__":
    run_interactive_extraction()
