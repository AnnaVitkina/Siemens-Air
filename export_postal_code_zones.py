"""
Export postal code zones from extracted Rate Card and Zones data to a txt file.

Reads from the processing folder and writes a tab-separated zones file to output.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from project_paths import OUTPUT_DIR, PROCESSING_DIR

ORIGIN_ZONE_COLUMN = "Origin Zone"
DESTINATION_ZONE_COLUMN = "Destination Zone"
SERVICE_CODE_COLUMN = "service code extended"
ZONES_AIRPORT_COLUMN = "Airport code (5-digit)"
ZONES_REGION_COLUMN = "Zone GAT2020"

OUTPUT_COLUMNS = ["Name", "Country", "Postal Code", "Excluded"]
ALL_COUNTRY_POSTAL_CODES = ", ".join(chr(code) for code in range(ord("A"), ord("Z") + 1))

AIRPORT_ZONE_PATTERN = re.compile(r"^[A-Z]{2}[A-Z0-9]{3}$")
ALL_COUNTRY_ZONE_PATTERN = re.compile(r"^ALL-([A-Z]{2})$")
US_REGION_ZONE_PATTERN = re.compile(r"^US-.+")


@dataclass
class AllCountryExclusion:
    all_zone: str
    airport_zone: str
    exclusion_name: str
    country: str
    excluded_postal_code: str


@dataclass
class ZonesExportResult:
    output_path: Path
    row_count: int
    airport_zone_count: int
    all_country_zone_count: int
    us_region_zone_count: int
    other_zone_count: int
    exclusion_zone_count: int = 0
    lane_replacement_count: int = 0


def _normalize_zone_name(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _normalize_service_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _build_exclusion_zone_name(all_zone: str, airport_zone: str) -> str:
    return f"{all_zone} (excl. {airport_zone})"


def _country_airport_zones(
    zone_names: list[str],
    country_code: str,
    all_zone: str,
) -> list[str]:
    return sorted(
        zone_name
        for zone_name in zone_names
        if AIRPORT_ZONE_PATTERN.match(zone_name)
        and zone_name.startswith(country_code)
        and zone_name != all_zone
    )


def _build_region_postal_lookup(
    zones_df: pd.DataFrame,
) -> dict[str, frozenset[str]]:
    lookup: dict[str, set[str]] = {}
    if ZONES_AIRPORT_COLUMN not in zones_df.columns or ZONES_REGION_COLUMN not in zones_df.columns:
        return {}

    for _, row in zones_df.iterrows():
        region = _normalize_zone_name(row[ZONES_REGION_COLUMN])
        airport = _normalize_zone_name(row[ZONES_AIRPORT_COLUMN])
        if not region or not airport or len(airport) <= 2:
            continue
        lookup.setdefault(region, set()).add(airport[2:])
    return {region: frozenset(postal_codes) for region, postal_codes in lookup.items()}


def _expand_zone_to_postal_codes(
    zone_name: str,
    region_postal_lookup: dict[str, frozenset[str]],
) -> frozenset[str]:
    zone_name = _normalize_zone_name(zone_name)
    if not zone_name:
        return frozenset()
    if ALL_COUNTRY_ZONE_PATTERN.match(zone_name):
        return frozenset()
    if US_REGION_ZONE_PATTERN.match(zone_name):
        return region_postal_lookup.get(zone_name, frozenset())
    if AIRPORT_ZONE_PATTERN.match(zone_name):
        return frozenset({zone_name[2:]})
    if len(zone_name) >= 3 and zone_name[:2].isalpha():
        return frozenset({zone_name[2:]})
    return frozenset({zone_name})


def _zones_overlap(
    zone_a: str,
    zone_b: str,
    region_postal_lookup: dict[str, frozenset[str]],
) -> bool:
    postal_a = _expand_zone_to_postal_codes(zone_a, region_postal_lookup)
    postal_b = _expand_zone_to_postal_codes(zone_b, region_postal_lookup)
    if not postal_a or not postal_b:
        return _normalize_zone_name(zone_a) == _normalize_zone_name(zone_b)
    return bool(postal_a & postal_b)


def detect_all_country_exclusions(
    rate_card: pd.DataFrame,
    zone_names: list[str],
    zones_df: pd.DataFrame,
) -> list[AllCountryExclusion]:
    exclusions: list[AllCountryExclusion] = []
    seen: set[tuple[str, str]] = set()
    region_postal_lookup = _build_region_postal_lookup(zones_df)

    all_country_zones = [
        zone_name
        for zone_name in zone_names
        if ALL_COUNTRY_ZONE_PATTERN.match(zone_name)
    ]

    for all_zone in all_country_zones:
        country_match = ALL_COUNTRY_ZONE_PATTERN.match(all_zone)
        if country_match is None:
            continue
        country_code = country_match.group(1)

        for airport_zone in _country_airport_zones(zone_names, country_code, all_zone):
            origin_conflicts = _find_origin_side_conflicts(
                rate_card,
                all_zone,
                airport_zone,
                region_postal_lookup,
            )
            destination_conflicts = _find_destination_side_conflicts(
                rate_card,
                all_zone,
                airport_zone,
                region_postal_lookup,
            )
            if not origin_conflicts and not destination_conflicts:
                continue

            key = (all_zone, airport_zone)
            if key in seen:
                continue
            seen.add(key)

            exclusions.append(
                AllCountryExclusion(
                    all_zone=all_zone,
                    airport_zone=airport_zone,
                    exclusion_name=_build_exclusion_zone_name(all_zone, airport_zone),
                    country=country_code,
                    excluded_postal_code=airport_zone[2:],
                )
            )

    return exclusions


def _find_origin_side_conflicts(
    rate_card: pd.DataFrame,
    all_zone: str,
    airport_zone: str,
    region_postal_lookup: dict[str, frozenset[str]],
) -> set[tuple[str, str]]:
    conflicts: set[tuple[str, str]] = set()
    service_codes = {
        _normalize_service_code(value)
        for value in rate_card[SERVICE_CODE_COLUMN].dropna()
    }

    for service_code in service_codes:
        service_mask = (
            rate_card[SERVICE_CODE_COLUMN].map(_normalize_service_code) == service_code
        )
        group = rate_card.loc[service_mask]

        all_rows = group[
            group[ORIGIN_ZONE_COLUMN].map(_normalize_zone_name) == all_zone
        ]
        airport_rows = group[
            group[ORIGIN_ZONE_COLUMN].map(_normalize_zone_name) == airport_zone
        ]
        if all_rows.empty or airport_rows.empty:
            continue

        for _, row_all in all_rows.iterrows():
            destination_all = _normalize_zone_name(row_all[DESTINATION_ZONE_COLUMN])
            for _, row_airport in airport_rows.iterrows():
                destination_airport = _normalize_zone_name(
                    row_airport[DESTINATION_ZONE_COLUMN]
                )
                if _zones_overlap(
                    destination_all,
                    destination_airport,
                    region_postal_lookup,
                ):
                    conflicts.add((destination_all, service_code))
                    break

    return conflicts


def _find_destination_side_conflicts(
    rate_card: pd.DataFrame,
    all_zone: str,
    airport_zone: str,
    region_postal_lookup: dict[str, frozenset[str]],
) -> set[tuple[str, str]]:
    conflicts: set[tuple[str, str]] = set()
    service_codes = {
        _normalize_service_code(value)
        for value in rate_card[SERVICE_CODE_COLUMN].dropna()
    }

    for service_code in service_codes:
        service_mask = (
            rate_card[SERVICE_CODE_COLUMN].map(_normalize_service_code) == service_code
        )
        group = rate_card.loc[service_mask]

        all_rows = group[
            group[DESTINATION_ZONE_COLUMN].map(_normalize_zone_name) == all_zone
        ]
        airport_rows = group[
            group[DESTINATION_ZONE_COLUMN].map(_normalize_zone_name) == airport_zone
        ]
        if all_rows.empty or airport_rows.empty:
            continue

        for _, row_all in all_rows.iterrows():
            origin_all = _normalize_zone_name(row_all[ORIGIN_ZONE_COLUMN])
            for _, row_airport in airport_rows.iterrows():
                origin_airport = _normalize_zone_name(row_airport[ORIGIN_ZONE_COLUMN])
                if _zones_overlap(origin_all, origin_airport, region_postal_lookup):
                    conflicts.add((origin_all, service_code))
                    break

    return conflicts


def apply_all_country_zone_exclusions(
    rate_card: pd.DataFrame,
    zones_df: pd.DataFrame,
) -> tuple[pd.DataFrame, list[AllCountryExclusion], int]:
    updated = rate_card.copy()
    zone_names = collect_unique_zone_names(updated)
    region_postal_lookup = _build_region_postal_lookup(zones_df)
    exclusions = detect_all_country_exclusions(updated, zone_names, zones_df)
    replacement_count = 0

    for exclusion in exclusions:
        origin_conflicts = _find_origin_side_conflicts(
            rate_card,
            exclusion.all_zone,
            exclusion.airport_zone,
            region_postal_lookup,
        )
        for destination_zone, service_code in origin_conflicts:
            mask = (
                updated[ORIGIN_ZONE_COLUMN].map(_normalize_zone_name) == exclusion.all_zone
            ) & (
                updated[DESTINATION_ZONE_COLUMN].map(_normalize_zone_name)
                == destination_zone
            ) & (
                updated[SERVICE_CODE_COLUMN].map(_normalize_service_code) == service_code
            )
            replacement_count += int(mask.sum())
            updated.loc[mask, ORIGIN_ZONE_COLUMN] = exclusion.exclusion_name

        destination_conflicts = _find_destination_side_conflicts(
            rate_card,
            exclusion.all_zone,
            exclusion.airport_zone,
            region_postal_lookup,
        )
        for origin_zone, service_code in destination_conflicts:
            mask = (
                updated[DESTINATION_ZONE_COLUMN].map(_normalize_zone_name)
                == exclusion.all_zone
            ) & (
                updated[ORIGIN_ZONE_COLUMN].map(_normalize_zone_name) == origin_zone
            ) & (
                updated[SERVICE_CODE_COLUMN].map(_normalize_service_code) == service_code
            )
            replacement_count += int(mask.sum())
            updated.loc[mask, DESTINATION_ZONE_COLUMN] = exclusion.exclusion_name

    return updated, exclusions, replacement_count


def _build_all_country_exclusion_row(exclusion: AllCountryExclusion) -> dict[str, str]:
    return {
        "Name": exclusion.exclusion_name,
        "Country": exclusion.country,
        "Postal Code": ALL_COUNTRY_POSTAL_CODES,
        "Excluded": exclusion.excluded_postal_code,
    }


def collect_unique_zone_names(rate_card: pd.DataFrame) -> list[str]:
    zone_names: set[str] = set()
    for column in (ORIGIN_ZONE_COLUMN, DESTINATION_ZONE_COLUMN):
        if column not in rate_card.columns:
            continue
        for value in rate_card[column].dropna():
            name = _normalize_zone_name(value)
            if name:
                zone_names.add(name)
    return sorted(zone_names)


def _build_airport_zone_row(zone_name: str) -> dict[str, str]:
    return {
        "Name": zone_name,
        "Country": zone_name[:2],
        "Postal Code": zone_name[-3:],
        "Excluded": "",
    }


def _build_all_country_zone_row(zone_name: str) -> dict[str, str]:
    match = ALL_COUNTRY_ZONE_PATTERN.match(zone_name)
    if match is None:
        raise ValueError(f"Invalid ALL country zone: {zone_name}")
    return {
        "Name": zone_name,
        "Country": match.group(1),
        "Postal Code": ALL_COUNTRY_POSTAL_CODES,
        "Excluded": "",
    }


def _build_us_region_zone_row(zone_name: str, zones_df: pd.DataFrame) -> dict[str, str]:
    if ZONES_AIRPORT_COLUMN not in zones_df.columns:
        raise ValueError(f"Missing column: {ZONES_AIRPORT_COLUMN}")
    if ZONES_REGION_COLUMN not in zones_df.columns:
        raise ValueError(f"Missing column: {ZONES_REGION_COLUMN}")

    matching_airports = zones_df.loc[
        zones_df[ZONES_REGION_COLUMN].astype(str).str.strip() == zone_name,
        ZONES_AIRPORT_COLUMN,
    ]
    postal_codes = sorted(
        {
            _normalize_zone_name(airport_code)[2:]
            for airport_code in matching_airports.dropna()
            if len(_normalize_zone_name(airport_code)) > 2
        }
    )

    return {
        "Name": zone_name,
        "Country": zone_name[:2],
        "Postal Code": ", ".join(postal_codes),
        "Excluded": "",
    }


def _build_generic_zone_row(zone_name: str) -> dict[str, str]:
    if len(zone_name) >= 2 and zone_name[:2].isalpha():
        return {
            "Name": zone_name,
            "Country": zone_name[:2],
            "Postal Code": zone_name[2:],
            "Excluded": "",
        }
    return {
        "Name": zone_name,
        "Country": "",
        "Postal Code": zone_name,
        "Excluded": "",
    }


def build_postal_code_zones(
    rate_card: pd.DataFrame,
    zones_df: pd.DataFrame,
    exclusions: list[AllCountryExclusion] | None = None,
) -> tuple[pd.DataFrame, dict[str, int]]:
    if exclusions is None:
        exclusions = detect_all_country_exclusions(
            rate_card,
            collect_unique_zone_names(rate_card),
            zones_df,
        )

    rows: list[dict[str, str]] = []
    counts = {
        "airport": 0,
        "all_country": 0,
        "us_region": 0,
        "other": 0,
        "exclusion": 0,
    }
    exclusion_names = {exclusion.exclusion_name for exclusion in exclusions}

    for zone_name in collect_unique_zone_names(rate_card):
        if zone_name in exclusion_names:
            continue
        if ALL_COUNTRY_ZONE_PATTERN.match(zone_name):
            row = _build_all_country_zone_row(zone_name)
            counts["all_country"] += 1
        elif US_REGION_ZONE_PATTERN.match(zone_name):
            row = _build_us_region_zone_row(zone_name, zones_df)
            counts["us_region"] += 1
        elif AIRPORT_ZONE_PATTERN.match(zone_name):
            row = _build_airport_zone_row(zone_name)
            counts["airport"] += 1
        else:
            row = _build_generic_zone_row(zone_name)
            counts["other"] += 1
        rows.append(row)

    for exclusion in exclusions:
        rows.append(_build_all_country_exclusion_row(exclusion))
        counts["exclusion"] += 1

    processed = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    if not processed.empty:
        processed = processed.sort_values("Name", kind="stable").reset_index(drop=True)
    return processed, counts


def write_postal_code_zones_txt(
    rate_card: pd.DataFrame,
    zones_df: pd.DataFrame,
    output_path: Path,
    exclusions: list[AllCountryExclusion] | None = None,
    lane_replacement_count: int = 0,
) -> ZonesExportResult:
    processed, counts = build_postal_code_zones(rate_card, zones_df, exclusions=exclusions)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8", newline="") as handle:
        handle.write("\t".join(OUTPUT_COLUMNS) + "\n")
        for _, row in processed.iterrows():
            handle.write("\t".join(str(row[column]) for column in OUTPUT_COLUMNS) + "\n")

    return ZonesExportResult(
        output_path=output_path,
        row_count=len(processed),
        airport_zone_count=counts["airport"],
        all_country_zone_count=counts["all_country"],
        us_region_zone_count=counts["us_region"],
        other_zone_count=counts["other"],
        exclusion_zone_count=counts["exclusion"],
        lane_replacement_count=lane_replacement_count,
    )


def export_zones_from_rate_card(
    rate_card: pd.DataFrame,
    zones_df: pd.DataFrame,
    output_file: Path,
) -> tuple[ZonesExportResult, pd.DataFrame]:
    updated_rate_card, exclusions, replacement_count = apply_all_country_zone_exclusions(
        rate_card,
        zones_df,
    )
    result = write_postal_code_zones_txt(
        updated_rate_card,
        zones_df,
        output_file,
        exclusions=exclusions,
        lane_replacement_count=replacement_count,
    )
    return result, updated_rate_card


def load_extracted_rate_card(processing_file: Path) -> pd.DataFrame:
    return pd.read_excel(processing_file, sheet_name="Rate Card", header=0)


def load_extracted_zones(processing_file: Path) -> pd.DataFrame:
    return pd.read_excel(processing_file, sheet_name="Zones", header=0)


def export_zones_from_processing_file(
    processing_file: Path,
    output_file: Path | None = None,
) -> ZonesExportResult:
    rate_card = load_extracted_rate_card(processing_file)
    zones_df = load_extracted_zones(processing_file)
    if output_file is None:
        output_file = OUTPUT_DIR / f"{processing_file.stem.replace('_extracted', '')}_zones.txt"
    result, _ = export_zones_from_rate_card(rate_card, zones_df, output_file)
    return result


def list_processing_files() -> list[Path]:
    if not PROCESSING_DIR.exists():
        return []
    return sorted(PROCESSING_DIR.glob("*_extracted.xlsx"))


def run_interactive_zones_export() -> ZonesExportResult:
    processing_files = list_processing_files()
    if not processing_files:
        raise FileNotFoundError(f"No extracted files found in {PROCESSING_DIR}")

    print("Extracted files in processing folder:")
    for index, file_path in enumerate(processing_files, start=1):
        print(f"  {index}. {file_path.name}")

    while True:
        raw = input("Enter file number to export zones from: ").strip()
        if raw.isdigit():
            choice_index = int(raw) - 1
            if 0 <= choice_index < len(processing_files):
                selected_file = processing_files[choice_index]
                break
        print("Invalid choice. Try again.")

    result = export_zones_from_processing_file(selected_file)
    print(f"\nSaved postal code zones file to: {result.output_path}")
    print(f"Total rows: {result.row_count}")
    print(f"Airport zones: {result.airport_zone_count}")
    print(f"ALL country zones: {result.all_country_zone_count}")
    print(f"US region zones: {result.us_region_zone_count}")
    print(f"Other zones: {result.other_zone_count}")
    print(f"ALL exclusion zones: {result.exclusion_zone_count}")
    print(f"Lane zone replacements: {result.lane_replacement_count}")
    return result


if __name__ == "__main__":
    run_interactive_zones_export()
