"""
Extract Air Rate Card and Airport Zones from a DHL Global Air Fact Sheet workbook.

Prompts for the source file and sheet tabs, then writes a consolidated workbook
to the processing folder with Rate Card and Zones tabs.
"""

from __future__ import annotations

import re
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
    "contract owner",
    "origin country",
    "forwarder",
)

RATE_CARD_COLUMN_RENAMES = {
    "origin_airport_code": "Origin Zone",
    "destination_airport_code": "Destination Zone",
    "origin airport code": "Origin Zone",
    "destination airport code": "Destination Zone",
    "origin airport code zone": "Origin Zone",
    "destination airport code zone": "Destination Zone",
    "siemens_division": "business unit code",
    "siemens_business": "siemens division",
    "siemens_buisness": "siemens division",
    "airline surcharge per kg": "fix fuel surcharge",
    "applicable for": "appli for DGR or Special",
    "volume weight ratio": "volume weight ratio",
    "volume weight ratio e g 0 0 0 1 6 6 1 7 7 1 8 8": "volume weight ratio",
    "siemens division in case of special agreement": "siemens division",
    "service codes": "service code extended",
    "2000 01 01 00 00 00": "ValidFrom",
    "2000 12 31 00 00 00": "ValidUntil",
    "1999 12 12 00 00 00": "date of update",
}

# Hardcoded mapping for dual-header templates:
# first-row business labels -> second-row technical labels.
RATE_CARD_UPPER_TO_LOWER_HEADER_RENAMES = {
    "country of contract owner iso code": "country_contract_owner",
    "published scm pi lo internal": "published_P_PC_LO_int",
    "forwarder name": "forwarder_name",
    "forwarder short code": "forwarder_short_code",
    "origin country": "origin_country",
    "origin country iso code": "origin_country_code",
    "origin city": "origin_city",
    "origin airport code zone": "origin_airport_code",
    "destination country": "destination_country",
    "destination country iso code": "destination_country_code",
    "destination city": "destination_city",
    "destination airport code zone": "destination_airport_code",
    "service level": "service_level",
    "service code": "service_code",
    "special service category e g dgr sp airline service odc uld cool": "appli_for_DGR_or_Special",
    "transit time in hours": "transit_time_in_hours",
    "remarks": "Remarks",
    "currency": "Currency",
    "minimum": "Minimum",
    "flat": "Flat",
    "45 kg": "C_45_kg",
    "100 kg": "C_100_kg",
    "300 kg": "C_300_kg",
    "500 kg": "C_500_kg",
    "1000 kg": "C_1000_kg",
    "3000 kg": "C_3000_kg",
    "99999 kg": "C_99999_kg",
    "fix fuel surcharge sfi rules in general req": "fix_fuel_surcharge",
    "fuel surcharge as per outlay": "fuel_surcharge_outlay",
    "fix security surcharge sfi rules in general req": "fix_security_surcharge",
    "security surcharge as per outlay": "security_surch_outlay",
    "dgr fee yes as per outlay amount fix": "DGR_fee",
    "rate applicable for main deck dims details adjacent": "Airdeck",
    "maximum lenght x": "max_length_cm",
    "dimensions width x": "max_width_cm",
    "in cm height": "max_height_cm",
    "airline": "airline",
    "frequency": "frequency",
    "volume weight ratio e g 0 0 0 1 6 6 1 7 7 1 8 8": "volume_weight_ratio",
    "currency outbound": "Currency_Outbound",
    "flat opc outbound processing charge": "Flat_OPC",
    "currency inbound": "Currency_Inbound",
    "flat ipc inbound processing charge": "Flat_IPC",
    "origin airport zone cluster": "origin_apt_zone_cluster",
    "dest airport zone cluster": "dest_apt_zone_cluster",
    "ratescope": "Rate_scope",
    "status": "Status",
    "id": "ID",
    "version": "Version",
    "siemens business in case of special agreement": "Siemens_Business",
    "business unit code in case of special agreement": "siemens_division",
    "siemens division (in case of special agreement)": "siemens_division",
    "responsible person of siemens email address": "responsible_person",
    "additional information s": "additional_information",
    "service code extended": "service_code_extended",
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
    normalized_markers = {_normalize_header_key(marker) for marker in markers}
    for value in row:
        if pd.isna(value):
            continue
        text = _normalize_header_key(str(value))
        text_tokens = set(text.split())
        if any(
            marker in text
            or set(marker.split()).issubset(text_tokens)
            for marker in normalized_markers
        ):
            return True
    return False


def _normalize_header_key(text: str) -> str:
    normalized = str(text).strip().lower()
    normalized = normalized.replace("\n", " ").replace("_", " ").replace("/", " ")
    normalized = re.sub(r"[^a-z0-9 ]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def rename_rate_card_columns(columns: pd.Index | list[str]) -> list[str]:
    renamed: list[str] = []
    for column in columns:
        name = str(column).strip()
        upper_alias = RATE_CARD_UPPER_TO_LOWER_HEADER_RENAMES.get(_normalize_header_key(name))
        if upper_alias is not None:
            name = upper_alias
        mapped = RATE_CARD_COLUMN_RENAMES.get(_normalize_header_key(name))
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


def extract_air_workbook_rate_card_only(
    file_path: Path,
    rate_card_sheet: str,
) -> ExtractionResult:
    rate_card_raw = pd.read_excel(
        file_path,
        sheet_name=rate_card_sheet,
        header=None,
    )
    rate_card_df, rate_card_header_row = extract_rate_card_df(rate_card_raw)
    zones_df = pd.DataFrame()

    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    output_file = PROCESSING_DIR / f"{file_path.stem}_extracted.xlsx"

    metadata_df = pd.DataFrame(
        [
            {"Field": "Source file", "Value": file_path.name},
            {"Field": "Rate Card tab", "Value": rate_card_sheet},
            {"Field": "Zones tab", "Value": "N/A (LATAM mode)"},
            {"Field": "Rate Card header row (0-based)", "Value": rate_card_header_row},
            {"Field": "Zones header row (0-based)", "Value": -1},
            {"Field": "Rate Card rows", "Value": len(rate_card_df)},
            {"Field": "Zones rows", "Value": 0},
        ]
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        rate_card_df.to_excel(
            writer,
            sheet_name=OUTPUT_SHEET_NAMES["rate_card"],
            index=False,
        )
        metadata_df.to_excel(writer, sheet_name="Metadata", index=False)

    return ExtractionResult(
        rate_card=rate_card_df,
        zones=zones_df,
        source_file=file_path,
        output_file=output_file,
        rate_card_header_row=rate_card_header_row,
        zones_header_row=-1,
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
