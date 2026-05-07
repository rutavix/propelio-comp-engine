"""
Excel report writer.

Takes the subject Property and the scored top comps from comp_engine
and writes a 4-sheet workbook (Subject Property, Top Comps, Neighborhood
Summary, ARV Analysis) to disk. If the target file is locked because
it's open in Excel, falls back to a timestamped sibling filename so the
run still produces output. No scoring or scraping happens here.
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd
from openpyxl.styles import Font

import config
from comp_engine import arv_summary, clean_neighborhood, neighborhood_summary
from scraper import Property


logger = logging.getLogger(__name__)


def _fmt_usd(val: Any) -> str:
    """Format a number as ``$1,234,567``; empty string when missing."""
    if val is None or val == "":
        return ""
    try:
        n = int(round(float(val)))
        return f"${n:,}"
    except (TypeError, ValueError):
        return str(val)


def _fmt_num(val: Any) -> str:
    """Whole-number comma formatting (sqft, year, count)."""
    if val is None or val == "":
        return ""
    try:
        n = int(round(float(val)))
        return f"{n:,}"
    except (TypeError, ValueError):
        return str(val)


def generate_excel(
    subject_property: Property,
    comps: List[Dict],
    output_path: Union[str, Path] = config.OUTPUT_FILE,
    *,
    valuation: Optional[Dict[str, Optional[float]]] = None,
) -> Path:
    """Write a 4-sheet workbook (subject, comps, neighborhoods, ARV)."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if valuation is None and isinstance(subject_property.extra, dict):
        extra_val = subject_property.extra.get("valuation")
        if isinstance(extra_val, dict):
            valuation = extra_val  # type: ignore[assignment]

    subject_df = _subject_frame(subject_property)
    comps_df = _comps_frame(comps)
    neighborhood_df = _neighborhood_frame(comps)
    arv_df = _arv_frame(comps, valuation)

    candidate_paths = [output_path, _timestamped_alternative(output_path)]

    last_perm_error: Exception | None = None
    for attempt, path in enumerate(candidate_paths, start=1):
        try:
            _write_workbook(
                path,
                subject_df,
                comps_df,
                neighborhood_df,
                arv_df,
                add_top_comps_missing_sqft_note=bool(
                    comps
                    and any(c.get("sqft_missing_not_filtered") for c in comps)
                ),
            )
        except PermissionError as exc:
            last_perm_error = exc
            if attempt < len(candidate_paths):
                logger.warning(
                    "PermissionError writing %s (file likely open in Excel); "
                    "retrying with timestamped filename.", path,
                )
                continue
            logger.error(
                "PermissionError on both %s and the timestamped fallback %s",
                output_path, path,
            )
            raise OSError(
                f"Could not write Excel report to {output_path} or "
                f"timestamped fallback {path}: {exc}"
            ) from exc
        except OSError as exc:
            logger.exception("Failed to write Excel report to %s", path)
            raise OSError(
                f"Could not write Excel report to {path}: {exc}"
            ) from exc

        if path != output_path:
            logger.warning(
                "Wrote Excel report to fallback path %s (original %s was "
                "locked: %s)", path, output_path, last_perm_error,
            )
        else:
            logger.info("Wrote Excel report to %s", path)
        return path.resolve()

    raise OSError(
        f"Could not write Excel report; exhausted candidate paths "
        f"({[str(p) for p in candidate_paths]})"
    )


_TOP_COMPS_MISSING_SQFT_NOTE = (
    "* Living area missing from API — sqft comparability unverified "
    "for this comp (included, not filtered)"
)


def _write_workbook(
    path: Path,
    subject_df: pd.DataFrame,
    comps_df: pd.DataFrame,
    neighborhood_df: pd.DataFrame,
    arv_df: pd.DataFrame,
    *,
    add_top_comps_missing_sqft_note: bool = False,
) -> None:
    """Materialize the 4-sheet workbook at ``path``."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        subject_df.to_excel(writer, sheet_name="Subject Property", index=False)
        comps_df.to_excel(writer, sheet_name="Top Comps", index=False)
        neighborhood_df.to_excel(
            writer, sheet_name="Neighborhood Summary", index=False
        )
        arv_df.to_excel(writer, sheet_name="ARV Analysis", index=False)
        if add_top_comps_missing_sqft_note and len(comps_df) > 0:
            _append_top_comps_missing_sqft_note(
                writer.sheets["Top Comps"], len(comps_df)
            )
        for sheet_name in writer.sheets:
            _autosize(writer.sheets[sheet_name])


def _append_top_comps_missing_sqft_note(worksheet, n_comp_rows: int) -> None:
    """Blank row after data, then italic gray note in A/B (not a data row)."""
    # Row 1 = headers; data rows 2 .. n_comp_rows + 1; blank n_comp_rows + 2;
    # note row n_comp_rows + 3
    note_row = n_comp_rows + 3
    footnote_font = Font(italic=True, color="FF9E9E9E")

    cell_a = worksheet.cell(row=note_row, column=1, value="Notes:")
    cell_b = worksheet.cell(row=note_row, column=2, value=_TOP_COMPS_MISSING_SQFT_NOTE)
    cell_a.font = footnote_font
    cell_b.font = footnote_font


def _timestamped_alternative(path: Path) -> Path:
    """Build ``<stem>_YYYYMMDD_HHMMSS<suffix>`` next to ``path``."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{stamp}{path.suffix}")


def _subject_frame(subject: Property) -> pd.DataFrame:
    """Subject sheet: core Property fields + parcel valuation / enrichment."""
    extra: Dict[str, Any] = (
        dict(subject.extra) if isinstance(subject.extra, dict) else {}
    )
    pe: Dict[str, Any] = (
        extra.get("parcel_enrichment")
        if isinstance(extra.get("parcel_enrichment"), dict) else {}
    )
    val: Dict[str, Any] = (
        extra.get("valuation") if isinstance(extra.get("valuation"), dict) else {}
    )

    rows: List[Dict[str, str]] = []

    def row(field: str, value: Any, notes: str = "") -> None:
        rows.append({"Field": field, "Value": value if value != "" else "", "Notes": notes})

    row("Address", subject.address or "")
    nb_raw = subject.neighborhood or pe.get("subdivision") or ""
    row("Neighborhood (Raw)", nb_raw)
    row(
        "Neighborhood (Clean)",
        clean_neighborhood(nb_raw),
        "Trim, spacing, and title case (no subdivision aliasing)",
    )
    row("Lot Size (sqft)", _fmt_num(subject.lot_size))

    liv = subject.sqft if subject.sqft is not None else pe.get("sqft")
    row("Living Area (sqft)", _fmt_num(liv), "CMA or parcel summary/buildingDetail")

    yb = subject.year_built if subject.year_built is not None else pe.get("year_built")
    row("Year Built", _fmt_num(yb), "CMA or parcel summary")

    row("Status", subject.status or pe.get("derived_status") or "")

    est = pe.get("estimated_value") if pe.get("estimated_value") is not None else val.get("estimate")
    low = pe.get("estimate_low") if pe.get("estimate_low") is not None else val.get("low")
    high = pe.get("estimate_high") if pe.get("estimate_high") is not None else val.get("high")
    row("Estimated Value", _fmt_usd(est), "valuationDetail.estimate.value")
    row("Estimate Low", _fmt_usd(low), "")
    row("Estimate High", _fmt_usd(high), "")

    dsp = pe.get("display_price")
    row("Display / Primary Price", _fmt_usd(dsp), "lastSalePrice or estimate.value")

    row("Last Sale Price", _fmt_usd(pe.get("last_sale_price")), "parcel lastSalePrice")
    row("Last Sale Date", str(pe.get("last_sale_date") or ""), "parcel lastSaleDate")

    row("Source", subject.source or "")

    return pd.DataFrame(rows)


def _comps_frame(comps: List[Dict]) -> pd.DataFrame:
    """Selected top comps only — currency formatted."""
    if not comps:
        return pd.DataFrame(columns=[
            "Address", "Price", "Lot Size (sqft)", "Living Area (sqft)",
            "Year Built", "Status",
            "Neighborhood (Raw)", "Neighborhood (Clean)",
            "Source",
            "Lot Confidence", "Nbhd Mult", "Nbhd Match Type", "Nbhd Match",
            "Neighborhood Match Reason", "Confidence Formula",
            "Boundary Flag", "Distance (mi)", "Subdivision Penalty Applied",
            "Final Confidence",
        ])

    rows = []
    for item in comps:
        prop: Property = item["property"]
        raw_nb = item.get("neighborhood_raw", prop.neighborhood or "")
        clean_nb = item.get(
            "neighborhood_clean",
            clean_neighborhood(prop.neighborhood or ""),
        )
        rows.append({
            "Address": prop.address,
            "Price": _fmt_usd(prop.price),
            "Lot Size (sqft)": _fmt_num(prop.lot_size),
            "Living Area (sqft)": (
                "N/A (sqft missing — not filtered)"
                if item.get("sqft_missing_not_filtered")
                else _fmt_num(prop.sqft)
            ),
            "Year Built": _fmt_num(prop.year_built),
            "Status": prop.status or "",
            "Neighborhood (Raw)": raw_nb if raw_nb else "",
            "Neighborhood (Clean)": clean_nb or "Unknown",
            "Source": prop.source,
            "Lot Confidence": item.get("lot_confidence"),
            "Nbhd Mult": item.get("neighborhood_multiplier"),
            "Nbhd Match Type": item.get("nbhd_match_type", ""),
            "Nbhd Match": item.get("neighborhood_match"),
            "Neighborhood Match Reason": item.get("nbhd_reason") or "",
            "Confidence Formula": item.get("conf_formula") or "",
            "Boundary Flag": item.get("boundary_flag") or "",
            "Distance (mi)": (
                f"{float(item['distance_miles']):.4f}"
                if item.get("distance_miles") is not None
                else ""
            ),
            "Subdivision Penalty Applied": item.get(
                "subdivision_penalty_applied", "",
            ),
            "Final Confidence": (
                f"{float(item['confidence']):.2f} *"
                if item.get("sqft_missing_not_filtered")
                else f"{float(item['confidence']):.2f}"
            ),
        })
    return pd.DataFrame(rows)


def _arv_frame(
    comps: List[Dict],
    valuation: Optional[Dict[str, Optional[float]]],
) -> pd.DataFrame:
    """Build the "ARV Analysis" sheet (``arv`` = avg of sold comps in ARV pool)."""
    summary = arv_summary(comps, valuation)
    note = summary.get("arv_note") or ""
    headline = summary.get("arv_headline") or ""
    arv_val = summary.get("arv")
    missing_sqft = summary.get("arv_missing_sqft_addresses") or []
    missing_sqft_text = (
        ", ".join(str(x) for x in missing_sqft)
        if missing_sqft
        else "None - all ARV comps have living area data"
    )
    rows: List[Dict[str, Any]] = [
        {
            "Field": "ARV basis (headline)",
            "Value": headline,
            "Notes": "Plain-language summary; same pool as numeric ARV below",
        },
        {
            "Field": "Subject Estimated Value",
            "Value": _fmt_usd(summary.get("subject_estimate")),
            "Notes": "Propelio valuationDetail.estimate.value",
        },
        {
            "Field": "Subject Estimate Low",
            "Value": _fmt_usd(summary.get("subject_estimate_low")),
            "Notes": "",
        },
        {
            "Field": "Subject Estimate High",
            "Value": _fmt_usd(summary.get("subject_estimate_high")),
            "Notes": "",
        },
        {
            "Field": "Top Sold Comps Average (ARV)",
            "Value": _fmt_usd(arv_val),
            "Notes": (
                "Same value as ARV: mean of sold comps in the ARV pool "
                "(exact/similar first; see arv_source). Not a mix of list/sale."
            ),
        },
        {
            "Field": "ARV Low",
            "Value": _fmt_usd(summary.get("arv_low")),
            "Notes": "Min sold price used for ARV (or priced fallback)",
        },
        {
            "Field": "ARV High",
            "Value": _fmt_usd(summary.get("arv_high")),
            "Notes": "Max sold price used for ARV (or priced fallback)",
        },
        {
            "Field": "ARV Source",
            "Value": summary.get("arv_source") or "",
            "Notes": note or "Derivation detail",
        },
        {
            "Field": "Sold Comps (status)",
            "Value": _fmt_num(summary.get("sold_comp_count")),
            "Notes": "Sold / closed comps counted in ARV pool",
        },
        {
            "Field": "Missing Sqft Comps in ARV Pool",
            "Value": missing_sqft_text,
            "Notes": "Comps kept because living area was missing (not filterable)",
        },
    ]
    return pd.DataFrame(rows)


def _neighborhood_frame(comps: List[Dict]) -> pd.DataFrame:
    summary = neighborhood_summary(comps)
    if not summary:
        return pd.DataFrame(columns=[
            "Neighborhood", "Comp Count", "Sold Count", "Avg Price",
            "Avg Sold Price", "Avg Lot Size (sqft)",
        ])

    rows: List[Dict[str, Any]] = []
    for neighborhood, stats in summary.items():
        rows.append({
            "Neighborhood": neighborhood,
            "Comp Count": stats["comp_count"],
            "Sold Count": stats.get("sold_count", 0),
            "Avg Price": _fmt_usd(stats["avg_price"]),
            "Avg Sold Price": _fmt_usd(stats.get("avg_sold_price")),
            "Avg Lot Size (sqft)": _fmt_num(stats.get("avg_lot_size")),
        })

    all_prices = [
        item["property"].price
        for item in comps
        if item.get("property") and item["property"].price is not None
    ]
    sold_n = sum(1 for item in comps if item.get("is_sold"))

    total_avg = None
    if all_prices:
        total_avg = round(sum(all_prices) / len(all_prices))

    sold_prices_total = [
        float(item["property"].price)
        for item in comps
        if item.get("is_sold") and item.get("property")
        and item["property"].price is not None
    ]
    total_avg_sold = None
    if sold_prices_total:
        total_avg_sold = round(
            sum(sold_prices_total) / len(sold_prices_total),
        )

    rows.append({
        "Neighborhood": "TOTAL COMPS",
        "Comp Count": len(comps),
        "Sold Count": sold_n,
        "Avg Price": _fmt_usd(total_avg),
        "Avg Sold Price": _fmt_usd(total_avg_sold),
        "Avg Lot Size (sqft)": "",
    })
    return pd.DataFrame(rows)


def _autosize(worksheet) -> None:
    """Roughly fit column widths to the longest cell value."""
    for column_cells in worksheet.columns:
        try:
            length = max(
                (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                default=10,
            )
            letter = column_cells[0].column_letter
            worksheet.column_dimensions[letter].width = min(length + 2, 60)
        except (AttributeError, ValueError):
            continue
