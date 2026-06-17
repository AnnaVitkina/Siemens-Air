"""
End-to-end Siemens Air rate card pipeline.

Steps:
  1. Choose input Excel file from input/
  2. Confirm sheet tabs (Rate Card, Zones)
  3. Extract to processing/{name}_extracted.xlsx
  4. Build matrix workbook in output/
  5. Export postal code zones txt in output/

Google Colab:
    from google.colab import drive
    drive.mount("/content/drive")
    exec(open("/content/Siemens-Air/run_pipeline.py").read())
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import pandas as pd

# Must run before any project imports — exec(open(...)) does not set sys.path or __file__.
_COLAB_PROJECT_DIRS = (
    Path("/content/Siemens-Air"),
    Path("/content/Siemens-air"),
)


def _resolve_project_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        pass
    for path in _COLAB_PROJECT_DIRS:
        if path.is_dir():
            return path
    return Path.cwd()


PROJECT_DIR = _resolve_project_dir()
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

LATAM_RATE_CARD_DEFAULT_SHEET = "rate card"

from project_paths import BASE_DIR, INPUT_DIR, OUTPUT_DIR, PROCESSING_DIR, ensure_project_dirs  # noqa: E402

from build_air_rate_card_matrix import (  # noqa: E402
    MatrixBuildResult,
    apply_lane_country_filters,
    build_matrix_from_rate_card,
    prompt_lane_country_filters,
)
from export_postal_code_zones import (  # noqa: E402
    ZonesExportResult,
    export_zones_from_rate_card,
)
from extract_air_data import (  # noqa: E402
    DEFAULT_SHEETS,
    ExtractionResult,
    extract_air_workbook,
    extract_air_workbook_rate_card_only,
    list_input_files,
    list_workbook_sheets,
    prompt_choice,
    prompt_sheet_name,
)


@dataclass
class PipelineResult:
    source_file: Path
    extraction: ExtractionResult
    matrix: MatrixBuildResult
    zones: ZonesExportResult | None


def prompt_shipper_mode() -> Literal["standard", "latam"]:
    choice = prompt_choice(
        "Which shipper is this file for?",
        ["Siemens Healthineers/Divisions", "Siemens LATAM"],
        default="Siemens Healthineers/Divisions",
    )
    if choice == "Siemens LATAM":
        return "latam"
    return "standard"


def prompt_siemens_business_scope() -> Literal["all_di", "all_di_shs", "both"]:
    choice = prompt_choice(
        "Which Siemens_Business scope should be used?",
        [
            "ALL and DI",
            "ALL, DI, and SHS",
            "Both",
        ],
        default="ALL and DI",
    )
    if choice == "ALL, DI, and SHS":
        return "all_di_shs"
    if choice == "Both":
        return "both"
    return "all_di"


def filter_rate_card_by_siemens_business(
    rate_card: pd.DataFrame,
    allowed_values: set[str],
) -> pd.DataFrame:
    column_name = "siemens division"
    if column_name not in rate_card.columns:
        raise ValueError(f"Missing column: {column_name}")

    normalized = (
        rate_card[column_name]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .replace({"ALL": "ALL"})
    )
    keep_mask = normalized.isin({value.upper() for value in allowed_values})
    return rate_card[keep_mask].reset_index(drop=True)


def run_interactive_pipeline() -> PipelineResult:
    ensure_project_dirs()

    print(f"Code folder:    {BASE_DIR}")
    print(f"  input/       -> {INPUT_DIR}")
    print(f"  processing/  -> {PROCESSING_DIR}")
    print(f"  output/      -> {OUTPUT_DIR}")
    if INPUT_DIR != BASE_DIR / "input":
        print("  (data folders on Google Drive)")
    print()

    input_files = list_input_files()
    if not input_files:
        raise FileNotFoundError(
            f"No Excel files found in {INPUT_DIR}. "
            "Upload a .xlsx file to the input folder and run again."
        )

    print("Files available in the input folder:")
    shipper_mode = prompt_shipper_mode()
    selected_name = prompt_choice(
        "Which file should be processed?",
        [path.name for path in input_files],
        default=input_files[0].name if len(input_files) == 1 else None,
    )
    file_path = INPUT_DIR / selected_name

    print(f"\n--- Sheet selection for: {file_path.name} ---")
    sheet_names = list_workbook_sheets(file_path)
    if shipper_mode == "latam":
        rate_card_tab = prompt_sheet_name(
            sheet_names,
            "Rate Card (look for tab containing: rate card)",
            LATAM_RATE_CARD_DEFAULT_SHEET,
        )
    else:
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

    print("\n--- Step 1/3: Extracting source workbook ---")
    if shipper_mode == "latam":
        extraction = extract_air_workbook_rate_card_only(file_path, rate_card_tab)
    else:
        extraction = extract_air_workbook(file_path, sheet_mapping)
    print(f"Saved extracted workbook: {extraction.output_file}")
    print(f"  Rate Card rows: {len(extraction.rate_card)}")
    if shipper_mode == "latam":
        print("  Zones rows: 0 (LATAM mode)")
    else:
        print(f"  Zones rows: {len(extraction.zones)}")

    business_scope: Literal["all_di", "all_di_shs", "both"] = "all_di"
    if shipper_mode == "standard":
        business_scope = prompt_siemens_business_scope()

    de_only, latam_only, drop_empty_or_zero_cost_columns = prompt_lane_country_filters()
    prefiltered_rate_card = apply_lane_country_filters(
        extraction.rate_card,
        de_only=de_only,
        latam_only=latam_only,
    )
    if len(prefiltered_rate_card) == 0:
        raise ValueError("No lanes remain after selected lane filters.")
    if de_only or latam_only:
        print(
            "Applied lane filters before downstream processing: "
            f"{len(extraction.rate_card)} -> {len(prefiltered_rate_card)} rows"
        )

    if shipper_mode == "standard" and business_scope != "both":
        allowed = {"ALL", "DI"} if business_scope == "all_di" else {"ALL", "DI", "SHS"}
        before_business_count = len(prefiltered_rate_card)
        prefiltered_rate_card = filter_rate_card_by_siemens_business(
            prefiltered_rate_card,
            allowed,
        )
        if len(prefiltered_rate_card) == 0:
            raise ValueError("No lanes remain after Siemens_Business filtering.")
        print(
            "Applied Siemens_Business filter before downstream processing: "
            f"{before_business_count} -> {len(prefiltered_rate_card)} rows"
        )

    output_stem = file_path.stem
    zones_output = OUTPUT_DIR / f"{output_stem}_zones.txt"
    matrix_output = OUTPUT_DIR / f"{output_stem}_matrix.xlsx"
    zones: ZonesExportResult | None = None
    rate_card = prefiltered_rate_card
    rate_card_all_di: pd.DataFrame | None = None
    rate_card_all_di_shs: pd.DataFrame | None = None

    if shipper_mode == "latam":
        print("\n--- Step 2/3: Skipping zones export (LATAM mode) ---")
    elif business_scope == "both":
        print(
            "\n--- Step 2/3: Applying exclusions per Siemens_Business scope "
            "(independent for each tab) ---"
        )
        rate_card_all_di_input = filter_rate_card_by_siemens_business(rate_card, {"ALL", "DI"})
        rate_card_all_di_shs_input = filter_rate_card_by_siemens_business(
            rate_card, {"ALL", "DI", "SHS"}
        )
        if len(rate_card_all_di_input) == 0:
            raise ValueError("No lanes remain for scope ALL and DI.")
        if len(rate_card_all_di_shs_input) == 0:
            raise ValueError("No lanes remain for scope ALL, DI, and SHS.")

        zones_all_di_output = OUTPUT_DIR / f"{output_stem}_ALL_DI_zones.txt"
        zones_all_di_shs_output = OUTPUT_DIR / f"{output_stem}_ALL_DI_SHS_zones.txt"

        zones_all_di, rate_card_all_di = export_zones_from_rate_card(
            rate_card_all_di_input,
            extraction.zones,
            zones_all_di_output,
        )
        zones_all_di_shs, rate_card_all_di_shs = export_zones_from_rate_card(
            rate_card_all_di_shs_input,
            extraction.zones,
            zones_all_di_shs_output,
        )
        zones = zones_all_di

        print(f"Saved zones file (ALL_DI): {zones_all_di.output_path}")
        print(
            "  ALL_DI replacements: "
            f"{zones_all_di.lane_replacement_count} | rows: {len(rate_card_all_di)}"
        )
        print(f"Saved zones file (ALL_DI_SHS): {zones_all_di_shs.output_path}")
        print(
            "  ALL_DI_SHS replacements: "
            f"{zones_all_di_shs.lane_replacement_count} | rows: {len(rate_card_all_di_shs)}"
        )
    else:
        print("\n--- Step 2/3: Applying ALL-country exclusions and exporting zones ---")
        zones, rate_card = export_zones_from_rate_card(
            rate_card,
            extraction.zones,
            zones_output,
        )
        print(f"Saved zones file: {zones.output_path}")
        print(f"  Total rows: {zones.row_count}")
        print(f"  Airport zones: {zones.airport_zone_count}")
        print(f"  ALL country zones: {zones.all_country_zone_count}")
        print(f"  US region zones: {zones.us_region_zone_count}")
        print(f"  Other zones: {zones.other_zone_count}")
        print(f"  ALL exclusion zones: {zones.exclusion_zone_count}")
        print(f"  Lane zone replacements: {zones.lane_replacement_count}")

    print("\n--- Step 3/3: Building matrix rate card ---")
    if shipper_mode == "standard" and business_scope == "both":
        if rate_card_all_di is None or rate_card_all_di_shs is None:
            raise ValueError("Internal error: missing scoped rate cards for both-tab build.")
        matrix = build_matrix_from_rate_card(
            rate_card_all_di,
            matrix_output,
            de_only=False,
            latam_only=False,
            drop_empty_or_zero_cost_columns=drop_empty_or_zero_cost_columns,
            ignore_volume_weight_ratio=False,
            include_dgr_fee_shipment_column=True,
            sheet_name="ALL_DI",
            append_sheet_if_exists=False,
        )
        matrix_all_di_shs = build_matrix_from_rate_card(
            rate_card_all_di_shs,
            matrix_output,
            de_only=False,
            latam_only=False,
            drop_empty_or_zero_cost_columns=drop_empty_or_zero_cost_columns,
            ignore_volume_weight_ratio=False,
            include_dgr_fee_shipment_column=True,
            sheet_name="ALL_DI_SHS",
            append_sheet_if_exists=True,
        )
        print(f"Saved matrix workbook: {matrix.matrix_path}")
        print(
            f"  Tab ALL_DI rows after filters: {matrix.row_count} | "
            f"Cost blocks: {matrix.cost_block_count}"
        )
        print(
            f"  Tab ALL_DI_SHS rows after filters: {matrix_all_di_shs.row_count} | "
            f"Cost blocks: {matrix_all_di_shs.cost_block_count}"
        )
        print(f"  Shipment columns: {matrix.shipment_column_count}")
    else:
        matrix = build_matrix_from_rate_card(
            rate_card,
            matrix_output,
            de_only=False,
            latam_only=False,
            drop_empty_or_zero_cost_columns=drop_empty_or_zero_cost_columns,
            ignore_volume_weight_ratio=(shipper_mode == "latam"),
            include_dgr_fee_shipment_column=(shipper_mode == "standard"),
        )
        print(f"Saved matrix workbook: {matrix.matrix_path}")
        print(f"  Data rows: {matrix.row_count}")
        print(f"  Shipment columns: {matrix.shipment_column_count}")
        print(f"  Cost blocks: {matrix.cost_block_count}")

    print("\n--- Pipeline complete ---")
    print(f"Source file:        {file_path}")
    print(f"Extracted workbook: {extraction.output_file}")
    print(f"Matrix workbook:    {matrix.matrix_path}")
    if shipper_mode == "standard" and business_scope == "both":
        print(f"Zones file:         {output_stem}_ALL_DI_zones.txt and {output_stem}_ALL_DI_SHS_zones.txt")
    elif zones is not None:
        print(f"Zones file:         {zones.output_path}")
    else:
        print("Zones file:         skipped (LATAM mode)")

    return PipelineResult(
        source_file=file_path,
        extraction=extraction,
        matrix=matrix,
        zones=zones,
    )


if __name__ == "__main__":
    pipeline_result = run_interactive_pipeline()
