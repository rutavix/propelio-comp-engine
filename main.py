"""
Command-line entry point.

This is the glue: it parses the user's CLI flags (subject address,
radius, lot-size override, output path, etc.), kicks off the Propelio
scraper, hands the results to the comp engine for scoring, prints a
short text summary to the console, and writes the Excel report. All the
real work lives in scraper.py, comp_engine.py, and output.py — this
file just orchestrates them.

Usage::

    python main.py "123 Main St, Austin, TX"
    python main.py "123 Main St" --radius 0.75 --output ./output/report.xlsx
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import sys
from pathlib import Path
from typing import List, Optional

import config
from comp_engine import (
    arv_summary,
    filter_comps,
    neighborhood_summary,
    pool_status_counts,
)
from output import generate_excel
from scraper import (
    Property,
    PropelioScraperError,
    scrape_comps,
)


logger = logging.getLogger(__name__)


def _fmt_console_money(val: Optional[float]) -> str:
    if val is None:
        return "n/a"
    try:
        return f"${float(val):,.0f}"
    except (TypeError, ValueError):
        return "n/a"


def _parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Pull real estate comps from Propelio for a subject address.",
    )
    parser.add_argument("address", help="Subject property address.")
    parser.add_argument(
        "--radius",
        type=float,
        default=config.RADIUS_MILES,
        help=f"Search radius in miles (default: {config.RADIUS_MILES}).",
    )
    parser.add_argument(
        "--lot-size",
        type=int,
        default=None,
        metavar="SQFT",
        help=(
            "Subject lot size in square feet. Overrides whatever the "
            "scraper recovers from Propelio's CMA / lead-details endpoints. "
            "Use this when the API doesn't return a lot size for the "
            "subject property and you know it from elsewhere."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=config.OUTPUT_FILE,
        help=f"Excel output path (default: {config.OUTPUT_FILE}).",
    )
    parser.add_argument(
        "--username",
        default=config.PROPELIO_USERNAME,
        help="Propelio username (overrides config.py).",
    )
    parser.add_argument(
        "--password",
        default=config.PROPELIO_PASSWORD,
        help="Propelio password (overrides config.py).",
    )
    parser.add_argument(
        "--new-builds",
        action="store_true",
        help=(
            "Restrict comps to new builds (year_built >= "
            f"{config.NEW_BUILDS_MIN_YEAR}). Default behavior is the "
            "opposite: drop properties built within "
            f"{config.NEW_BUILD_YEARS} years of the current year."
        ),
    )
    parser.add_argument(
        "--expand-radius",
        action="store_true",
        dest="expand_radius",
        help=(
            "Prefer exact subdivision comps first; if fewer than three exact "
            "matches, fill from nearby subdivisions with a 0.6x confidence "
            "penalty (testing)."
        ),
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose (DEBUG) logging.",
    )
    return parser.parse_args(argv)


def _build_subject(
    address: str,
    scraped_subject: Property,
    *,
    lot_size_override: Optional[int] = None,
) -> Property:
    """Build the subject Property the comp engine will score against.

    Uses the address the user typed (so display matches their input)
    but pulls ``lot_size`` / ``sqft`` / ``year_built`` / ``neighborhood``
    from whatever the scraper recovered from the CMA payload.

    ``lot_size_override`` (CLI flag ``--lot-size``) wins over the
    scraper-resolved value when supplied — useful when Propelio doesn't
    return a lot size for the subject property and the user knows it
    from elsewhere.
    """
    final_lot_size = (
        float(lot_size_override) if lot_size_override is not None
        else scraped_subject.lot_size
    )
    extra = (
        dict(scraped_subject.extra)
        if isinstance(scraped_subject.extra, dict) else {}
    )
    pe = extra.get("parcel_enrichment") if isinstance(extra.get("parcel_enrichment"), dict) else {}

    merged_sqft = scraped_subject.sqft if scraped_subject.sqft is not None else pe.get("sqft")
    merged_year = (
        scraped_subject.year_built
        if scraped_subject.year_built is not None else pe.get("year_built")
    )
    merged_price = (
        pe.get("display_price")
        if pe.get("display_price") is not None
        else scraped_subject.price
    )
    merged_status = pe.get("derived_status") or scraped_subject.status
    merged_nb = scraped_subject.neighborhood or pe.get("subdivision") or ""

    if pe.get("lat") is not None and extra.get("lat") is None:
        extra["lat"] = pe["lat"]
    if pe.get("lon") is not None and extra.get("lon") is None:
        extra["lon"] = pe["lon"]

    return Property(
        address=address,
        lot_size=final_lot_size,
        sqft=merged_sqft,
        year_built=merged_year,
        neighborhood=merged_nb if merged_nb else None,
        price=merged_price,
        status=merged_status,
        source="user_input",
        extra=extra,
    )


def _print_summary(
    subject: Property,
    comps: List[dict],
    *,
    pool_size: int,
    status_counts: dict,
) -> None:
    print()
    print("=" * 72)
    print(f"Subject: {subject.address}")
    if subject.neighborhood:
        print(f"  Neighborhood: {subject.neighborhood}")
    if subject.lot_size:
        print(f"  Lot size: {subject.lot_size:,.0f} sqft")
    if subject.sqft:
        print(f"  Living area: {subject.sqft:,.0f} sqft")
    if subject.year_built:
        print(f"  Year built: {subject.year_built}")
    print("=" * 72)

    print(
        f"Scored pool: {pool_size} comp(s) "
        f"(sold={status_counts.get('sold', 0)}, "
        f"pending={status_counts.get('pending', 0)}, "
        f"for_sale={status_counts.get('for_sale', 0)}, "
        f"other={status_counts.get('other', 0)}, "
        f"unknown={status_counts.get('unknown', 0)})"
    )

    if not comps:
        print("No qualifying comps found.")
        return

    print(f"Top {len(comps)} comps (sorted by final confidence; "
          "guarantees at least one sold comp when available):")
    for i, item in enumerate(comps, start=1):
        prop = item["property"]
        price = f"${prop.price:,.0f}" if prop.price else "n/a"
        lot = f"{prop.lot_size:,.0f}" if prop.lot_size else "n/a"
        raw_nb = item.get("neighborhood_raw", prop.neighborhood or "")
        clean_nb = item.get("neighborhood_clean") or prop.neighborhood or ""
        mtype = item.get("neighborhood_match", "n/a")
        nb_m = item.get("neighborhood_multiplier", 1.0)
        reason = item.get("selection_reason", "top-N by final confidence")
        print(
            f"  {i}. {prop.address}\n"
            f"     price={price}  |  lot={lot} sqft  |  status={prop.status or 'n/a'}  "
            f"|  conf={item['confidence']:.2f}\n"
            f"     nbhd raw={raw_nb!r}  |  clean={clean_nb!r}  |  "
            f"match={mtype} (nbhd mult {nb_m})  |  {reason}"
        )

    print()
    print(
        f"Neighborhoods present in top {len(comps)}:"
    )
    for neighborhood, stats in neighborhood_summary(comps).items():
        avg_price = f"${stats['avg_price']:,.0f}" if stats["avg_price"] else "n/a"
        print(
            f"  - {neighborhood}: {stats['comp_count']} comps, "
            f"avg price {avg_price}"
        )

    val = (
        subject.extra.get("valuation")
        if isinstance(subject.extra, dict) else None
    )
    arv = arv_summary(comps, val if isinstance(val, dict) else None)
    print()
    print("ARV breakdown:")
    print(f"  ARV pool: {arv.get('arv_pool_kind') or 'n/a'}")
    if arv.get("arv_pool_note"):
        print(f"  Pool note: {arv['arv_pool_note']}")
    addrs = arv.get("arv_comp_addresses") or []
    if addrs:
        print(f"  ARV pool comps ({len(addrs)}): {addrs}")
    sold_used = arv.get("arv_sold_addresses") or []
    if sold_used:
        print(f"  Sold comps used for ARV ({len(sold_used)}): {sold_used}")
    print(
        "  Top sold comps average (ARV): "
        f"{_fmt_console_money(arv.get('top_sold_comps_average') or arv.get('arv'))}"
    )
    print(f"  ARV: {_fmt_console_money(arv.get('arv'))}")
    lo, hi = arv.get("arv_low"), arv.get("arv_high")
    print(
        f"  Low / High: {_fmt_console_money(lo)} — {_fmt_console_money(hi)}"
    )
    print(f"  Source: {arv.get('arv_source') or 'n/a'}")
    print(f"  Headline: {arv.get('arv_headline') or 'n/a'}")
    print(f"  Sold comps in ARV pool (status): {arv.get('sold_comp_count', 0)}")
    print()


async def _run(args: argparse.Namespace) -> int:
    """Async pipeline: scrape -> filter -> write Excel."""
    try:
        scraped_subject, properties = await scrape_comps(
            args.address,
            radius=args.radius,
            username=args.username,
            password=args.password,
        )
    except PropelioScraperError as exc:
        logger.error("Scraping failed: %s", exc)
        return 2
    except Exception:
        logger.exception("Unexpected error during scraping")
        return 3

    subject = _build_subject(
        args.address, scraped_subject, lot_size_override=args.lot_size,
    )
    lot_source = (
        "CLI override --lot-size" if args.lot_size is not None
        else "Propelio CMA / lead-details"
    )
    logger.info(
        "Subject: address=%r lot_size=%s (%s) sqft=%s year_built=%s",
        subject.address, subject.lot_size, lot_source,
        subject.sqft, subject.year_built,
    )

    if not properties:
        logger.warning("No comp candidates returned from Propelio.")
        print("No comp candidates returned for the given address.")
        return 1

    if subject.lot_size is None:
        # Only fires when the scraper couldn't resolve a lot size AND
        # the user didn't supply --lot-size. With either source set,
        # subject.lot_size is non-None and this branch is skipped.
        logger.warning(
            "Subject lot size could not be determined from CMA payload "
            "and no --lot-size override was supplied; filter_comps will "
            "return no comps. Check the 'CMA payload keys' / 'CMA params' "
            "log lines to see what shape Propelio actually returned, then "
            "either extend _extract_subject in scraper.py or rerun with "
            "--lot-size <SQFT>."
        )

    comps, scored = filter_comps(
        subject,
        properties,
        prefer_new_builds=args.new_builds,
        radius_miles=args.radius,
        expand_neighborhood=args.expand_radius,
    )
    counts = pool_status_counts(scored)
    logger.info(
        "Scored pool: size=%d sold=%d pending=%d for_sale=%d other=%d unknown=%d "
        "(prefer_new_builds=%s)",
        len(scored),
        counts.get("sold", 0), counts.get("pending", 0),
        counts.get("for_sale", 0), counts.get("other", 0),
        counts.get("unknown", 0),
        args.new_builds,
    )

    valuation = (
        subject.extra.get("valuation")
        if isinstance(subject.extra, dict) else None
    )
    try:
        output_path = generate_excel(
            subject, comps, args.output, valuation=valuation,
        )
    except OSError as exc:
        logger.error("Failed to write Excel report: %s", exc)
        return 4

    _print_summary(
        subject, comps,
        pool_size=len(scored),
        status_counts=counts,
    )
    print(f"Report written to: {output_path}")
    return 0


def main(argv: List[str] | None = None) -> int:
    args = _parse_args(argv if argv is not None else sys.argv[1:])

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    if sys.platform.startswith("win"):
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    return asyncio.run(_run(args))


if __name__ == "__main__":
    raise SystemExit(main())
