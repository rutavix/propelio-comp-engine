"""
Comp scoring and filtering.

Given a subject Property and a list of candidate Property records (from
scraper.py), this module decides which ones are usable comps and ranks
them. It applies the lot-size / living-area / new-build filters, scores
each comp by similarity (lot fit, neighborhood match, distance), picks
the top N for the Excel report, and computes the ARV summary numbers.
No HTTP, no I/O, no Excel — pure logic.
"""

from __future__ import annotations

import logging
import math
import re
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Tuple

import config
from scraper import Property


logger = logging.getLogger(__name__)


# --- Neighborhood helpers --------------------------------------------------

SIMILARITY_THRESHOLD = 0.9


def clean_neighborhood(name: Optional[str]) -> str:
    """Trim, collapse whitespace, and title-case for display/grouping.

    Does **not** merge different subdivisions into one label — only cosmetic
    normalization.
    """
    if not name:
        return ""
    collapsed = " ".join(str(name).strip().split())
    return collapsed.title()


def _neighborhood_key(name: Optional[str]) -> str:
    """Lowercase, single-spaced key for equality checks."""
    if not name:
        return ""
    return " ".join(str(name).strip().split()).lower()


def _neighborhood_similarity(a: str, b: str) -> float:
    """Ratio in [0, 1] using built-in difflib (no extra dependencies)."""
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return float(SequenceMatcher(None, a, b).ratio())


def _nbhd_match_type_label(match: str) -> str:
    """Excel-facing neighborhood tier."""
    if match in ("exact", "similar", "same_city", "different_area", "unknown"):
        return match
    return "other"


# Listing statuses treated as "sold" for scoring pool / top-N guarantees
_SOLD_STATUS_TOKENS = frozenset({"sold", "closed", "recently_sold"})


def _normalize_status(status: Optional[str]) -> str:
    """Lower-cased, whitespace-collapsed status string (``""`` if unset)."""
    if not status:
        return ""
    return " ".join(str(status).strip().lower().split())


def _is_sold(status: Optional[str]) -> bool:
    return _normalize_status(status) in _SOLD_STATUS_TOKENS


_NB_MATCH_RANK = {
    "exact": 4,
    "similar": 3,
    "same_city": 2,
    "different_area": 1,
    "unknown": 0,
}

NEW_BUILD_CONFIDENCE_BOOST = 0.15
NEW_BUILD_BOOST_YEAR = 2018


def calculate_lot_confidence(
    subject_lot: Optional[float],
    comp_lot: Optional[float],
) -> float:
    """Confidence score for how well ``comp_lot`` matches ``subject_lot``."""
    if not subject_lot or not comp_lot or subject_lot <= 0 or comp_lot <= 0:
        return 0.0

    deviation = abs(comp_lot - subject_lot) / subject_lot

    if deviation <= 0.25:
        return 1.0
    if deviation >= 0.33:
        return 0.0

    span = 0.33 - 0.25
    progress = (deviation - 0.25) / span
    return round(0.9 - progress * (0.9 - 0.5), 4)


def _is_new_build(
    year_built: Optional[int],
    reference_year: Optional[int] = None,
    horizon: int = config.NEW_BUILD_YEARS,
) -> bool:
    """Treat anything built within ``horizon`` years as a new build."""
    if year_built is None:
        return False
    ref = reference_year or datetime.utcnow().year
    return (ref - year_built) <= horizon


def _within_lot_tolerance(
    subject_lot: Optional[float],
    comp_lot: Optional[float],
    tolerance: float = config.LOT_SIZE_TOLERANCE,
) -> bool:
    if not subject_lot or not comp_lot or subject_lot <= 0:
        return False
    return abs(comp_lot - subject_lot) / subject_lot <= tolerance


def _within_living_area_tolerance(
    subject_sqft: Optional[float],
    comp_sqft: Optional[float],
    tolerance: float = config.LIVING_AREA_TOLERANCE,
) -> tuple[bool, Optional[float]]:
    """Check if comp living area is within tolerance of subject living area.
    
    Returns (passes_filter, variance_pct) where variance_pct is None if filter passes
    or the actual percentage if it fails.
    """
    if not subject_sqft or subject_sqft <= 0:
        return True, None  # No subject sqft, can't filter
    if not comp_sqft or comp_sqft <= 0:
        return True, None  # No comp sqft, can't filter
    
    variance = abs(comp_sqft - subject_sqft) / subject_sqft
    if variance <= tolerance:
        return True, None
    else:
        return False, variance


def _city_from_address(address: Optional[str]) -> Optional[str]:
    """Best-effort US-style city: segment after first comma."""
    if not address:
        return None
    parts = [p.strip() for p in address.split(",")]
    if len(parts) >= 2 and parts[1]:
        return parts[1].lower()
    return None


def _coords_from_property(prop: Property) -> Optional[Tuple[float, float]]:
    """Return (lat, lon) from ``Property.extra`` or ``parcel_enrichment``."""
    ex = prop.extra if isinstance(prop.extra, dict) else {}
    lat, lon = ex.get("lat"), ex.get("lon")
    if lat is None or lon is None:
        pe = ex.get("parcel_enrichment")
        if isinstance(pe, dict):
            if lat is None:
                lat = pe.get("lat")
            if lon is None:
                lon = pe.get("lon")
    if lat is None or lon is None:
        return None
    try:
        return float(lat), float(lon)
    except (TypeError, ValueError):
        return None


def _haversine_miles(
    lat1: float, lon1: float, lat2: float, lon2: float,
) -> float:
    """Great-circle distance in miles (WGS84 sphere approx)."""
    r = 3959.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = (
        math.sin(dphi / 2) ** 2
        + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    )
    return 2 * r * math.asin(min(1.0, math.sqrt(a)))


def _street_line_key(address: Optional[str]) -> str:
    """First line of address without leading street number, lowercased."""
    if not address:
        return ""
    first = address.split(",")[0].strip()
    first = re.sub(r"^\d+[A-Za-z]?\s+", "", first, count=1)
    return " ".join(first.lower().split())


def _major_street_boundary(
    subject_property: Property,
    comp: Property,
) -> tuple[bool, str, float]:
    """If comp is close but on a different street, treat as major-street boundary."""
    sc = _coords_from_property(subject_property)
    cc = _coords_from_property(comp)
    if not sc or not cc:
        return False, "", 1.0
    dist_mi = _haversine_miles(sc[0], sc[1], cc[0], cc[1])
    dist_ft = dist_mi * 5280
    if dist_ft < config.MAJOR_STREET_MIN_DISTANCE_FT:
        return False, "", 1.0
    if dist_ft > config.MAJOR_STREET_MAX_DISTANCE_FT:
        return False, "", 1.0
    sk = _street_line_key(subject_property.address)
    ck = _street_line_key(comp.address)
    if not sk or not ck or sk == ck:
        return False, "", 1.0
    return True, "Major-Street Boundary", float(config.MAJOR_STREET_BOUNDARY_MULT)


def _neighborhood_match_and_multiplier(
    subject_property: Property,
    comp: Property,
) -> tuple[str, float, str]:
    """Classify subdivision match: exact / similar / same_city / different_area / unknown."""
    sub_clean = clean_neighborhood(subject_property.neighborhood or "")
    comp_clean = clean_neighborhood(comp.neighborhood or "")
    sub_key = _neighborhood_key(subject_property.neighborhood or "")
    comp_key = _neighborhood_key(comp.neighborhood or "")
    sub_city = _city_from_address(subject_property.address)
    comp_city = _city_from_address(comp.address)

    if not sub_key and not comp_key:
        return ("exact", 1.0, "Exact subdivision match")

    if not sub_key or not comp_key:
        return (
            "unknown",
            0.8,
            "Missing subdivision name — reduced confidence",
        )

    if sub_key == comp_key:
        return ("exact", 1.0, "Exact subdivision match")

    sim = _neighborhood_similarity(sub_key, comp_key)
    if sim >= SIMILARITY_THRESHOLD:
        return ("similar", 0.9, "Normalized match (case/spacing)")

    if sub_city and comp_city and sub_city == comp_city:
        return ("same_city", 0.8, "Different subdivision, same city")

    return (
        "different_area",
        0.6,
        "Different subdivision and city / area",
    )


def score_pool(
    subject_property: Property,
    properties_list: Iterable[Property],
    confidence_threshold: float = config.CONFIDENCE_THRESHOLD,
    *,
    prefer_new_builds: bool = False,
    new_builds_min_year: int = config.NEW_BUILDS_MIN_YEAR,
    max_distance_miles: Optional[float] = None,
) -> List[Dict]:
    """Score every comp candidate; return the qualifying pool."""
    subject_lot = subject_property.lot_size
    if not subject_lot:
        logger.warning("Subject property has no lot size; cannot score comps.")
        return []

    subject_sqft = subject_property.sqft
    sub_coords = _coords_from_property(subject_property)

    scored: List[Dict] = []
    for prop in properties_list:
        if prop is subject_property or prop.address == subject_property.address:
            continue
        if not _within_lot_tolerance(subject_lot, prop.lot_size):
            continue
        
        # Add living area filter BEFORE scoring
        sqft_missing_not_filtered = False
        if subject_sqft and subject_sqft > 0 and (prop.sqft is None or prop.sqft <= 0):
            sqft_missing_not_filtered = True
            logger.info(
                "[SQFT FILTER] Kept %s - living area missing, cannot filter "
                "(included with flag)",
                prop.address,
            )
        passes_filter, variance_pct = _within_living_area_tolerance(subject_sqft, prop.sqft)
        if not passes_filter:
            logger.info(
                "[SQFT FILTER] Excluded %s - %s sqft vs subject %s sqft (%.0f%% variance)",
                prop.address, prop.sqft, subject_sqft, variance_pct * 100
            )
            continue
            
        if prefer_new_builds:
            if prop.year_built is None or prop.year_built < new_builds_min_year:
                continue
        elif _is_new_build(prop.year_built):
            continue

        cc = _coords_from_property(prop)
        dist_mi: Optional[float] = None
        if sub_coords and cc:
            dist_mi = _haversine_miles(
                sub_coords[0], sub_coords[1], cc[0], cc[1],
            )

        if max_distance_miles is not None and sub_coords and cc:
            if dist_mi is not None and dist_mi > max_distance_miles:
                continue

        lot_confidence = calculate_lot_confidence(subject_lot, prop.lot_size)
        nb_match, nb_mult, nb_reason = _neighborhood_match_and_multiplier(
            subject_property, prop,
        )
        cross_subdivision_penalty = nb_match in (
            "same_city", "different_area", "unknown",
        )
        cross_mult = (
            float(config.CROSS_SUBDIVISION_CONF_MULT)
            if cross_subdivision_penalty
            else 1.0
        )

        prod_pre_boundary = float(lot_confidence * nb_mult * cross_mult)

        boundary_mult = 1.0
        boundary_flag = "no_coordinates"

        # Boundary penalty logic based on match type
        if dist_mi is None:
            boundary_flag = "no_coordinates"
        elif nb_match in ("exact", "similar"):
            # Exact/similar should never get an extra boundary penalty.
            boundary_flag = "no_boundary_penalty"
            boundary_mult = 1.0
        elif nb_match == "same_city":
            if dist_mi > float(config.EXPAND_NEIGHBORHOOD_DISTANCE_FLAG_MI):
                boundary_flag = "major_street_possible"
                boundary_mult = 0.9
            else:
                boundary_flag = "within_0.3mi"
                boundary_mult = 1.0
        elif nb_match == "unknown":
            boundary_flag = "unknown_subdivision"
            boundary_mult = 0.8
        else:
            # Keep existing adjacent-street heuristic for other cross-area matches.
            if dist_mi > float(config.EXPAND_NEIGHBORHOOD_DISTANCE_FLAG_MI):
                boundary_flag = "major_street_possible"
            else:
                boundary_flag = "no_boundary_penalty"
            hit, _bh_label, heuristic_mult = _major_street_boundary(
                subject_property, prop,
            )
            boundary_mult = float(heuristic_mult) if hit else 1.0

        prod_after_boundary = prod_pre_boundary * boundary_mult

        confidence = prod_after_boundary
        nb_boost_applied = False
        if prefer_new_builds and prop.year_built is not None:
            if prop.year_built >= NEW_BUILD_BOOST_YEAR:
                before_boost = confidence
                confidence = min(1.0, confidence + NEW_BUILD_CONFIDENCE_BOOST)
                nb_boost_applied = confidence > before_boost - 1e-9

        # Two decimal places for investor-facing scores (avoids float dust / long tails)
        confidence = round(confidence, 2)

        if confidence < confidence_threshold:
            continue

        dist_part = (
            f"{dist_mi:.4f}" if dist_mi is not None else "n/a"
        )
        b_eff_numeric = float(boundary_mult)
        
        # Build confidence formula (math only, no text description)
        conf_formula = (
            f"{lot_confidence:.4f} lot_conf × {nb_mult:.4f} nbhd_mult × "
            f"{cross_mult:.4f} cross_subdivision_mult × "
            f"{b_eff_numeric:.4f} boundary_adjacent_mult"
        )
        if nb_boost_applied:
            conf_formula += (
                f"; +{NEW_BUILD_CONFIDENCE_BOOST:.2f} new_build_boost "
                f"(cap 1.0)"
            )
        conf_formula += f" = {confidence:.4f} final"
        conf_formula += (
            f" | distance_mi={dist_part}; boundary_flag={boundary_flag}"
        )

        # DEBUG log with all multipliers
        logger.debug(
            f"[CONF] {prop.address}: lot={lot_confidence:.4f} × nbhd={nb_mult:.4f} × "
            f"cross_subdiv={cross_mult:.4f} × boundary={b_eff_numeric:.4f} × "
            f"expand=1.0000 = {confidence:.4f}"
        )

        if (
            nb_mult < 1.0 or cross_mult < 1.0 or b_eff_numeric < 1.0
            or boundary_flag == "major_street_possible"
        ):
            logger.info(
                "Neighborhood adjustment: %s | status=%s | conf_formula=%s | "
                "match=%r | nbhd_reason=%r",
                prop.address,
                prop.status or "n/a",
                conf_formula,
                nb_match,
                nb_reason,
            )

        raw_nb = prop.neighborhood or ""
        clean_nb = clean_neighborhood(raw_nb)
        disp_clean = clean_nb if clean_nb else "Unknown"
        scored.append({
            "property": prop,
            "confidence": confidence,
            "neighborhood": disp_clean,
            "neighborhood_raw": raw_nb,
            "neighborhood_clean": disp_clean,
            "neighborhood_display": disp_clean,
            "neighborhood_key": disp_clean,
            "lot_confidence": lot_confidence,
            "neighborhood_match": nb_match,
            "neighborhood_multiplier": nb_mult,
            "neighborhood_reason": nb_reason,
            "nbhd_reason": nb_reason,
            "conf_formula": conf_formula,
            "nbhd_match_type": _nbhd_match_type_label(nb_match),
            "cross_subdivision_penalty": cross_subdivision_penalty,
            "cross_subdivision_multiplier": cross_mult,
            "subdivision_penalty_applied": "Yes" if cross_subdivision_penalty else "No",
            "boundary_flag": boundary_flag,
            "boundary_multiplier": b_eff_numeric,
            "distance_miles": dist_mi,
            "sqft_missing_not_filtered": sqft_missing_not_filtered,
            "confidence_formula": conf_formula,
            "status_normalized": _normalize_status(prop.status),
            "is_sold": _is_sold(prop.status),
        })

    return scored


def pool_status_counts(scored: List[Dict]) -> Dict[str, int]:
    buckets = Counter({
        "sold": 0, "pending": 0, "for_sale": 0, "other": 0, "unknown": 0,
    })
    for item in scored:
        status = item.get("status_normalized") or ""
        if not status:
            buckets["unknown"] += 1
        elif status in _SOLD_STATUS_TOKENS:
            buckets["sold"] += 1
        elif status == "pending":
            buckets["pending"] += 1
        elif status in {"for_sale", "for sale", "active", "listed"}:
            buckets["for_sale"] += 1
        else:
            buckets["other"] += 1
    return dict(buckets)


def _build_expansion_merged_pool(
    scored: List[Dict],
    confidence_threshold: float,
) -> List[Dict]:
    """Exact and similar comps unchanged; only same_city/different_area/unknown get ``EXPAND_NEIGHBORHOOD_CONF_MULT``."""
    mult = float(config.EXPAND_NEIGHBORHOOD_CONF_MULT)
    merged: List[Dict] = []
    for item in scored:
        # Don't apply expand penalty to exact or similar matches
        if item.get("neighborhood_match") in ("exact", "similar"):
            merged.append({**item, "expanded_neighborhood_fill": False})
            continue
        
        # Apply expand penalty only to same_city, different_area, unknown
        new_conf = round(float(item["confidence"]) * mult, 2)
        if new_conf < confidence_threshold:
            continue
        prior_f = item.get("conf_formula") or item.get("confidence_formula") or ""
        expand_suffix = (
            f" × {mult:.2f} expand_nearby_subdivision_mult => {new_conf:.2f} final"
        )
        merged_formula = (prior_f + expand_suffix) if prior_f else expand_suffix.strip()
        prop_y: Property = item["property"]
        
        # DEBUG log with expand multiplier
        logger.debug(
            f"[CONF] {prop_y.address}: lot={float(item.get('lot_confidence') or 0.0):.4f} × "
            f"nbhd={float(item.get('neighborhood_multiplier') or 1.0):.4f} × "
            f"cross_subdiv={float(item.get('cross_subdivision_multiplier') or 1.0):.4f} × "
            f"boundary={float(item.get('boundary_multiplier') or 1.0):.4f} × "
            f"expand={mult:.4f} = {new_conf:.4f}"
        )
        
        merged.append({
            **item,
            "confidence": new_conf,
            "expanded_neighborhood_fill": True,
            "neighborhood_expansion_penalty_mult": mult,
            "conf_formula": merged_formula,
            "confidence_formula": merged_formula,
        })
    return merged


def _sort_key(item: Dict, *, prefer_new_builds: bool) -> tuple:
    """Prefer higher confidence → stronger nbhd match → lot fit → newer year."""
    prop = item["property"]
    yb = prop.year_built if prop.year_built is not None else -1
    return (
        -float(item.get("confidence") or 0.0),
        -_NB_MATCH_RANK.get(item.get("neighborhood_match"), 0),
        -float(item.get("lot_confidence") or 0.0),
        -yb if prefer_new_builds else 0,
    )


def select_top_comps(
    scored: List[Dict],
    top_n: int = config.TOP_COMP_COUNT,
    *,
    prefer_new_builds: bool = False,
) -> List[Dict]:
    """Pick top ``top_n`` comps; prefer ≥2 sold in the result when the pool allows."""
    if not scored:
        return []

    ordered = sorted(
        scored,
        key=lambda x: _sort_key(x, prefer_new_builds=prefer_new_builds),
    )
    selected = list(ordered[:top_n])

    pool_sold_outside = [x for x in ordered[top_n:] if x.get("is_sold")]
    sold_in_selected = lambda: sum(1 for x in selected if x.get("is_sold"))

    # Prefer at least **two** sold comps in the final slice when available
    if len([x for x in ordered if x.get("is_sold")]) >= 2:
        need = 2 - sold_in_selected()
        k = 0
        while need > 0 and k < len(pool_sold_outside):
            non_sold_idx = [i for i, x in enumerate(selected) if not x.get("is_sold")]
            if not non_sold_idx:
                break
            drop_i = non_sold_idx[-1]
            selected[drop_i] = pool_sold_outside[k]
            k += 1
            need -= 1

    # Guarantee ≥1 sold when pool has any sold (legacy behavior)
    elif sold_in_selected() == 0 and pool_sold_outside:
        non_sold_idx = [i for i, x in enumerate(selected) if not x.get("is_sold")]
        if non_sold_idx:
            drop_i = non_sold_idx[-1]
            selected[drop_i] = pool_sold_outside[0]

    for item in selected:
        if item.get("sqft_missing_not_filtered"):
            prop = item.get("property")
            addr = prop.address if prop else "?"
            logger.warning(
                "[ARV WARNING] %s selected in top comps but living area is unknown "
                "— sqft comparability unverified",
                addr,
            )

    annotated: List[Dict] = []
    for item in selected:
        annotated.append({**item, "selection_reason": "ranked by confidence & rules"})

    statuses_in_top = Counter(
        i.get("status_normalized") or "unknown" for i in annotated
    )
    logger.info(
        "Selected %d/%d comps (status mix in top-N: %s; pool size %d)",
        len(annotated), top_n,
        ", ".join(f"{s}={c}" for s, c in sorted(statuses_in_top.items())),
        len(scored),
    )
    return annotated


def filter_comps(
    subject_property: Property,
    properties_list: Iterable[Property],
    confidence_threshold: float = config.CONFIDENCE_THRESHOLD,
    top_n: int = config.TOP_COMP_COUNT,
    *,
    prefer_new_builds: bool = False,
    new_builds_min_year: int = config.NEW_BUILDS_MIN_YEAR,
    radius_miles: Optional[float] = None,
    expand_neighborhood: bool = False,
) -> Tuple[List[Dict], List[Dict]]:
    """Score comps with controlled radius expansion when the pool is thin.

    When ``expand_neighborhood`` is True, prefer up to ``top_n`` comps with
    **exact** subdivision match; if fewer than ``top_n`` exist, merge in
    other matches with a 0.6x confidence penalty (see config).

    Returns ``(top_comps, full_scored_pool)``.
    """
    req = config.RADIUS_MILES if radius_miles is None else float(radius_miles)
    steps = sorted({s for s in config.RADIUS_EXPANSION_STEPS if s <= req + 1e-9})
    if not steps or steps[-1] < req - 1e-9:
        steps = sorted(set(steps) | {req})

    scored: List[Dict] = []
    eff_radius = req
    for eff_radius in steps:
        scored = score_pool(
            subject_property,
            properties_list,
            confidence_threshold,
            prefer_new_builds=prefer_new_builds,
            new_builds_min_year=new_builds_min_year,
            max_distance_miles=eff_radius,
        )
        if len(scored) >= config.MIN_POOL_BEFORE_RADIUS_EXPAND:
            break

    if not scored:
        logger.info("No comps passed filters for %s", subject_property.address)
        return [], []

    logger.info(
        "Effective comp search radius: %.2f mi (requested max %.2f mi; "
        "pool size %d)",
        eff_radius,
        req,
        len(scored),
    )

    counts = pool_status_counts(scored)
    logger.info(
        "Scored pool status mix: sold=%d pending=%d for_sale=%d other=%d unknown=%d "
        "(pool size %d)",
        counts["sold"], counts["pending"], counts["for_sale"],
        counts["other"], counts["unknown"], len(scored),
    )

    exact_pool = [x for x in scored if x.get("neighborhood_match") == "exact"]
    if expand_neighborhood:
        if len(exact_pool) >= top_n:
            selected = select_top_comps(
                exact_pool,
                top_n=top_n,
                prefer_new_builds=prefer_new_builds,
            )
            logger.info(
                "Expand-radius: using %d top comp(s) from exact subdivision only "
                "(%d exact in pool).",
                len(selected),
                len(exact_pool),
            )
        else:
            logger.warning(
                "Expanding to nearby neighborhoods - only %d exact matches found",
                len(exact_pool),
            )
            merged = _build_expansion_merged_pool(
                scored, confidence_threshold,
            )
            selected = select_top_comps(
                merged,
                top_n=top_n,
                prefer_new_builds=prefer_new_builds,
            )
    else:
        selected = select_top_comps(
            scored, top_n=top_n, prefer_new_builds=prefer_new_builds,
        )
    return selected, scored


def neighborhood_summary(comps: List[Dict]) -> Dict[str, Dict[str, Any]]:
    """Aggregate by **neighborhood_clean** (trim / spacing / case only)."""
    merge_order: List[str] = []
    grouped: Dict[str, List[Property]] = defaultdict(list)

    for item in comps:
        prop: Property = item["property"]
        ck = item.get("neighborhood_clean") or clean_neighborhood(
            prop.neighborhood or "",
        )
        if not ck:
            ck = "Unknown"
        if ck not in merge_order:
            merge_order.append(ck)
        grouped[ck].append(item["property"])

    summary: Dict[str, Dict[str, Any]] = {}
    for ck in merge_order:
        props = grouped[ck]
        sold_props = [p for p in props if _is_sold(p.status)]
        sold_prices = [p.price for p in sold_props if p.price is not None]
        summary[ck] = {
            "comp_count": len(props),
            "sold_count": len(sold_props),
            "avg_price": _avg(p.price for p in props),
            "avg_sold_price": _avg(sold_prices) if sold_prices else None,
            "avg_lot_size": _avg(p.lot_size for p in props),
        }
    return summary


def _avg(values: Iterable[Optional[float]]) -> Optional[float]:
    nums = [v for v in values if v is not None]
    if not nums:
        return None
    return round(sum(nums) / len(nums), 2)


def _arv_select_pool(comps: List[Dict]) -> tuple[List[Dict], str, str]:
    """Prefer exact+similar; else same_city; else all. Returns pool, kind, note."""
    primary = [
        c for c in comps
        if c.get("neighborhood_match") in ("exact", "similar")
    ]
    if primary:
        return primary, "exact_similar", ""

    same_city = [c for c in comps if c.get("neighborhood_match") == "same_city"]
    if same_city:
        return (
            same_city,
            "same_city_fallback",
            "No exact or similar subdivision match in top comps; "
            "ARV uses same-city comps only.",
        )

    return (
        list(comps),
        "all_fallback",
        "No same-city subdivision matches; ARV uses all top comps.",
    )


def arv_summary(
    comps: List[Dict],
    valuation: Optional[Dict[str, Optional[float]]] = None,
) -> Dict[str, Any]:
    """Single-source ARV: average of **sold** comps in the ARV pool (top-comp order)."""

    out: Dict[str, Any] = {
        "subject_estimate": None,
        "subject_estimate_low": None,
        "subject_estimate_high": None,
        "top3_avg_price": None,
        "top_sold_comps_average": None,
        "arv": None,
        "arv_low": None,
        "arv_high": None,
        "arv_source": None,
        "arv_note": "",
        "arv_headline": "",
        "sold_comp_count": 0,
        "arv_sold_priced_count": 0,
        "arv_pool_kind": "",
        "arv_pool_note": "",
        "arv_comp_addresses": [],
        "arv_sold_addresses": [],
        "arv_missing_sqft_addresses": [],
    }
    if isinstance(valuation, dict):
        out["subject_estimate"] = valuation.get("estimate")
        out["subject_estimate_low"] = valuation.get("low")
        out["subject_estimate_high"] = valuation.get("high")

    if not comps:
        return out

    pool, pool_kind, pool_note = _arv_select_pool(comps)
    out["arv_pool_kind"] = pool_kind
    out["arv_pool_note"] = pool_note
    out["arv_comp_addresses"] = [
        item["property"].address for item in pool if item.get("property")
    ]
    out["arv_missing_sqft_addresses"] = [
        item["property"].address
        for item in pool
        if item.get("property")
        and (item["property"].sqft is None or float(item["property"].sqft) <= 0)
    ]

    sold_rows_ordered = [
        item for item in pool
        if item.get("property")
        and _is_sold(item["property"].status)
        and item["property"].price is not None
    ]
    sold_prices = [float(item["property"].price) for item in sold_rows_ordered]
    out["sold_comp_count"] = len(sold_rows_ordered)
    out["arv_sold_priced_count"] = len(sold_prices)
    out["arv_sold_addresses"] = [
        item["property"].address for item in sold_rows_ordered
    ]

    top_prices_all = [
        float(item["property"].price)
        for item in pool
        if item.get("property") and item["property"].price is not None
    ]

    def _headline(base: str) -> str:
        if pool_note:
            return f"{base} — {pool_note}"
        return base

    def _assign_arv_from_sold() -> None:
        """Set arv / low / high / averages from ``sold_prices`` (single source)."""
        if not sold_prices:
            return
        avg = round(sum(sold_prices) / len(sold_prices))
        arv_val = float(avg)
        out["arv"] = arv_val
        out["arv_low"] = float(min(sold_prices))
        out["arv_high"] = float(max(sold_prices))
        # Same numeric path for Excel + console (no mixed-status average).
        out["top_sold_comps_average"] = arv_val
        out["top3_avg_price"] = arv_val

    _assign_arv_from_sold()

    if len(sold_prices) >= 2:
        n = len(sold_prices)
        out["arv_headline"] = _headline(f"ARV based on {n} sold comps")
        out["arv_source"] = f"avg of {n} sold comps (pool={pool_kind})"
        out["arv_note"] = out["arv_headline"]
        return out

    if len(sold_prices) == 1:
        out["arv_headline"] = _headline(
            "ARV based on 1 sold comp (limited confidence)",
        )
        out["arv_source"] = f"single sold comp (pool={pool_kind})"
        out["arv_note"] = out["arv_headline"]
        return out

    # No sold comps: ARV falls back to priced pool; keep averages aligned.
    out["top_sold_comps_average"] = None
    out["top3_avg_price"] = None
    if top_prices_all:
        fb = float(round(sum(top_prices_all) / len(top_prices_all)))
        out["arv"] = fb
        out["arv_low"] = float(min(top_prices_all))
        out["arv_high"] = float(max(top_prices_all))
        out["top3_avg_price"] = fb
        out["arv_source"] = f"fallback - no sold comps (pool={pool_kind})"
        out["arv_note"] = _headline(
            "Mean of priced comps in ARV pool (no sold closings)",
        )
        out["arv_headline"] = out["arv_note"]
    else:
        out["arv_source"] = f"fallback - no priced comps (pool={pool_kind})"
        out["arv_note"] = _headline("No priced comps in ARV pool")
        out["arv_headline"] = out["arv_note"]

    return out
