"""
Propelio API client.

This module talks to Propelio's private REST API and nothing else: it
logs in with the credentials in config.py, looks an address up to a
Propelio lead id (parcel suggest -> parcel detail -> legacy/leads),
fetches the CMA (comps) payload for that lead, and converts each raw
record into a flat ``Property`` dataclass for the rest of the project
to consume. All scoring, filtering, and Excel rendering live elsewhere.
"""

from __future__ import annotations

import asyncio
import json
import logging
import re
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import requests

import config


logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

PROPELIO_API_BASE = "https://api.propelio.com"
LOGIN_URL = f"{PROPELIO_API_BASE}/login"
CMA_URL_TEMPLATE = f"{PROPELIO_API_BASE}/legacy/cma/{{lead_id}}"
PARCEL_SUGGEST_BASE = f"{PROPELIO_API_BASE}/parcels/v1/parcels/suggest"
PARCEL_DETAIL_TEMPLATE = f"{PROPELIO_API_BASE}/parcels/v1/parcels/{{parcel_id}}"
LEGACY_LEADS_WITHADDRESS_URL = f"{PROPELIO_API_BASE}/legacy/leads/withaddress"


def _coerce_positive_float(val: Any) -> Optional[float]:
    if val is None:
        return None
    try:
        f = float(val)
    except (TypeError, ValueError):
        return None
    if f > 0:
        return f
    return None


def _parcel_subject_lot_sqft(parcel: Dict[str, Any]) -> Optional[float]:
    """Subject lot size from ``GET /parcels/v1/parcels/{uuid}`` (sq ft when present)."""
    land = parcel.get("landDetail")
    if isinstance(land, dict):
        v = _coerce_positive_float(land.get("lotArea"))
        if v is not None:
            return v
    summary = parcel.get("summary")
    if isinstance(summary, dict):
        v = _coerce_positive_float(summary.get("lotSqft"))
        if v is not None:
            return v
    return None


def _parse_optional_datetime(val: Any):
    """Best-effort parse to timezone-aware UTC ``datetime``."""
    from datetime import datetime, timezone

    if val is None:
        return None
    if isinstance(val, datetime):
        return val if val.tzinfo else val.replace(tzinfo=timezone.utc)
    if isinstance(val, (int, float)):
        ts = float(val)
        if ts > 1e12:
            ts /= 1000.0
        return datetime.fromtimestamp(ts, tz=timezone.utc)
    s = str(val).strip()
    if not s:
        return None
    if "T" in s:
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
        except ValueError:
            pass
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").replace(tzinfo=timezone.utc)
    except ValueError:
        return None


def _parcel_subject_enrichment(
    parcel: Dict[str, Any],
    valuation: Dict[str, Optional[float]],
) -> Dict[str, Any]:
    """Subject-facing fields from parcel API (valuationDetail + summary + sale).

    Used by ``output`` Subject sheet and passed via ``Property.extra``.
    """
    from datetime import datetime, timezone

    summary = parcel.get("summary") if isinstance(parcel.get("summary"), dict) else {}
    bd = (
        parcel.get("buildingDetail")
        if isinstance(parcel.get("buildingDetail"), dict) else {}
    )

    last_sale_price = _coerce_positive_float(
        parcel.get("lastSalePrice") or parcel.get("last_sale_price")
        or summary.get("lastSalePrice") or summary.get("last_sale_price"),
    )
    last_sale_raw = (
        parcel.get("lastSaleDate") or parcel.get("last_sale_date")
        or summary.get("lastSaleDate") or summary.get("last_sale_date")
    )
    last_sale_dt = _parse_optional_datetime(last_sale_raw)

    est = valuation.get("estimate") if valuation else None
    display_price = last_sale_price if last_sale_price is not None else est

    sqft = _coerce_positive_float(summary.get("sqft")) or _coerce_positive_float(
        bd.get("livingArea") or bd.get("living_area"),
    )

    yb = summary.get("yearBuilt") or summary.get("year_built")
    year_built: Optional[int] = None
    if yb is not None:
        try:
            year_built = int(float(yb))
        except (TypeError, ValueError):
            year_built = None

    derived_status = "active/unknown"
    if last_sale_dt is not None:
        delta = datetime.now(timezone.utc) - last_sale_dt
        if delta.days <= 183:
            derived_status = "recently_sold"
        else:
            derived_status = "active/unknown"

    addr = parcel.get("address") if isinstance(parcel.get("address"), dict) else {}
    loc = addr.get("location") if isinstance(addr.get("location"), dict) else {}
    plat = loc.get("latitude") or addr.get("lat") or addr.get("latitude")
    plon = (
        loc.get("longitude") or addr.get("lon") or addr.get("longitude")
        or addr.get("lng")
    )
    plat_f: Optional[float] = None
    plon_f: Optional[float] = None
    if plat is not None:
        try:
            plat_f = float(plat)
        except (TypeError, ValueError):
            plat_f = None
    if plon is not None:
        try:
            plon_f = float(plon)
        except (TypeError, ValueError):
            plon_f = None

    subdivision = (
        (addr.get("subdivision") or addr.get("subDivision") or "").strip() or None
    )

    out_pe: Dict[str, Any] = {
        "estimated_value": est,
        "estimate_low": valuation.get("low") if valuation else None,
        "estimate_high": valuation.get("high") if valuation else None,
        "last_sale_price": last_sale_price,
        "last_sale_date": last_sale_raw,
        "display_price": display_price,
        "sqft": sqft,
        "year_built": year_built,
        "derived_status": derived_status,
        "subdivision": subdivision,
        "lat": plat_f,
        "lon": plon_f,
    }
    return out_pe


def _parcel_valuation_estimate(
    parcel: Dict[str, Any],
) -> Dict[str, Optional[float]]:
    """Extract subject valuation (``estimate`` / ``low`` / ``high``) from parcel.

    Reads ``valuationDetail.estimate.{value,low,high}`` first, falling
    back to ``valuationDetail.{low,high,estimateLow,estimateHigh}`` for
    range fields. Always returns a dict with all three keys (values may
    be ``None``) so consumers can render unconditionally.
    """
    out: Dict[str, Optional[float]] = {
        "estimate": None, "low": None, "high": None,
    }
    vd = parcel.get("valuationDetail")
    if not isinstance(vd, dict):
        return out

    est = vd.get("estimate")
    if isinstance(est, dict):
        out["estimate"] = _coerce_positive_float(est.get("value"))
        out["low"] = _coerce_positive_float(est.get("low"))
        out["high"] = _coerce_positive_float(est.get("high"))
    elif est is not None:
        out["estimate"] = _coerce_positive_float(est)

    if out["low"] is None:
        out["low"] = _coerce_positive_float(
            vd.get("low") or vd.get("estimateLow")
        )
    if out["high"] is None:
        out["high"] = _coerce_positive_float(
            vd.get("high") or vd.get("estimateHigh")
        )
    return out


def _withaddress_payload_from_parcel(parcel: Dict[str, Any]) -> Dict[str, Any]:
    """JSON body for ``POST /legacy/leads/withaddress``.

    No ``confirmationKey`` here — the first POST is expected to return
    HTTP 409 with ``error.confirmationKey`` set; the caller adds that key
    to this payload and POSTs again to actually create / fetch the lead.
    """
    addr = parcel.get("address")
    if not isinstance(addr, dict):
        raise PropelioScraperError(
            "Parcel detail response missing a dict 'address' for withaddress."
        )
    line1 = (addr.get("line1") or addr.get("street") or "").strip()
    city = (addr.get("city") or "").strip()
    state = (addr.get("state") or "").strip()
    postal = (
        addr.get("postalCode") or addr.get("postal_code") or addr.get("zip") or ""
    )
    postal = str(postal).strip()
    loc = addr.get("location")
    if not isinstance(loc, dict):
        raise PropelioScraperError(
            "Parcel address missing 'location' dict (lat/lon required for withaddress)."
        )
    lat = loc.get("latitude")
    lon = loc.get("longitude")
    if lat is None or lon is None:
        raise PropelioScraperError(
            "Parcel address.location missing latitude and/or longitude."
        )
    if not line1 or not city or not state or not postal:
        raise PropelioScraperError(
            "Parcel address missing line1, city, state, or postalCode for withaddress."
        )
    return {
        "line1": line1,
        "city": city,
        "state": state,
        "zip": postal,
        "lat": lat,
        "lon": lon,
        "useIncludedConsumables": True,
    }


def _extract_confirmation_key_from_withaddress_json(
    payload: Dict[str, Any],
) -> Optional[str]:
    """Pull ``confirmationKey`` out of a withaddress JSON object (esp. 409s)."""
    for source in (
        payload.get("error"),
        payload.get("result"),
        payload.get("data"),
        payload,
    ):
        if not isinstance(source, dict):
            continue
        for key in ("confirmationKey", "confirmation_key"):
            val = source.get(key)
            if isinstance(val, str) and val.strip():
                return val.strip()
    return None


def _extract_lead_id_from_withaddress_json(payload: Dict[str, Any]) -> Optional[str]:
    """Best-effort ``result.id`` from a withaddress JSON object (incl. 409 shapes)."""

    def _id_from_dict(node: Dict[str, Any]) -> Optional[str]:
        lid = node.get("id")
        if lid is None or lid == "":
            return None
        return str(lid).strip()

    result = payload.get("result")
    if isinstance(result, dict):
        got = _id_from_dict(result)
        if got:
            return got

    err = payload.get("error")
    if isinstance(err, dict):
        for key in ("result", "data", "body", "response", "lead", "payload"):
            node = err.get(key)
            if isinstance(node, dict):
                got = _id_from_dict(node)
                if got:
                    return got
                inner = node.get("result")
                if isinstance(inner, dict):
                    got = _id_from_dict(inner)
                    if got:
                        return got
        got = _id_from_dict(err)
        if got:
            return got

    errors = payload.get("errors")
    if isinstance(errors, list):
        for item in errors:
            if not isinstance(item, dict):
                continue
            r = item.get("result")
            if isinstance(r, dict):
                got = _id_from_dict(r)
                if got:
                    return got

    data = payload.get("data")
    if isinstance(data, dict):
        r2 = data.get("result")
        if isinstance(r2, dict):
            got = _id_from_dict(r2)
            if got:
                return got

    return None


# ---------------------------------------------------------------------------
# Property dataclass / errors
# ---------------------------------------------------------------------------

@dataclass
class Property:
    """Normalized representation of a property record."""

    address: str
    price: Optional[float] = None
    lot_size: Optional[float] = None       # square feet
    sqft: Optional[float] = None           # living area
    year_built: Optional[int] = None
    status: Optional[str] = None           # sold | for_sale | pending
    neighborhood: Optional[str] = None
    source: str = "propelio"
    extra: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data.pop("extra", None)
        return data


class PropelioScraperError(Exception):
    """Raised when the API client cannot complete an operation."""


# ---------------------------------------------------------------------------
# Authenticated client
# ---------------------------------------------------------------------------

class PropelioClient:
    """Authenticated ``requests.Session`` against Propelio's private API."""

    def __init__(
        self,
        username: str,
        password: str,
        *,
        timeout: int = config.HTTP_TIMEOUT_SECONDS,
        base_url: str = PROPELIO_API_BASE,
        proxies: Optional[Dict[str, str]] = None,
    ) -> None:
        print(f"[DEBUG] PROXY_URL loaded: {config.PROXY_URL}")
        if not username or not password:
            raise PropelioScraperError(
                "Propelio credentials are not configured. Set "
                "PROPELIO_USERNAME and PROPELIO_PASSWORD in config.py."
            )

        self.username = username
        self.password = password
        self.timeout = timeout
        self.base_url = base_url.rstrip("/")

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": config.HTTP_USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "en-US,en;q=0.9",
            "Content-Type": "application/json",
            "Origin": "https://genesis.propelio.com",
            "Referer": "https://genesis.propelio.com/",
        })

        # Route every request through the configured proxy. Pass
        # ``proxies={}`` to disable (useful for tests).
        proxy_map = config.PROPELIO_PROXIES if proxies is None else proxies
        if proxy_map:
            self.session.proxies.update(proxy_map)
            logger.info(
                "PropelioClient routing through proxy %s",
                _redact_proxy(proxy_map.get("https") or proxy_map.get("http")),
            )

        self._token: Optional[str] = None
        self._logged_in = False

    # -- Login ---------------------------------------------------------------

    def login(self) -> None:
        """POST credentials to ``/login`` and capture the resulting auth.

        On success, either a bearer token is added to the session's
        ``Authorization`` header or the session's cookie jar carries
        the auth implicitly (Propelio may use either; we accept both).
        """
        if self._logged_in:
            return

        login_url = f"{self.base_url}/login"
        payload = {"email": self.username, "password": self.password}

        try:
            response = self.session.post(
                login_url, json=payload, timeout=self.timeout
            )
        except requests.RequestException as exc:
            raise PropelioScraperError(
                f"Login network error against {login_url}: {exc}"
            ) from exc

        if not response.ok:
            raise PropelioScraperError(
                f"Login returned HTTP {response.status_code} "
                f"({response.reason}): {response.text[:500]}"
            )

        token = self._extract_token(response)
        if token:
            self._token = token
            self.session.headers["Authorization"] = f"Bearer {token}"
            logger.info("Logged in via bearer token (length=%d)", len(token))
        elif self.session.cookies:
            logger.info(
                "Logged in via session cookies (%d cookie(s) set)",
                len(self.session.cookies),
            )
        else:
            raise PropelioScraperError(
                "Login succeeded (HTTP 200) but neither a token nor a "
                "session cookie was returned. Body snippet: "
                f"{response.text[:500]}"
            )

        self._logged_in = True

    @staticmethod
    def _extract_token(response: requests.Response) -> Optional[str]:
        """Try the common JSON-token shapes Propelio might use."""
        try:
            body = response.json()
        except ValueError:
            return None
        if not isinstance(body, dict):
            return None

        # VERIFY: adjust these keys once you see the real login response.
        for key in ("token", "access_token", "sessionToken", "jwt"):
            value = body.get(key)
            if isinstance(value, str) and value:
                return value

        nested = body.get("data")
        if isinstance(nested, dict):
            for key in ("token", "access_token", "sessionToken", "jwt"):
                value = nested.get(key)
                if isinstance(value, str) and value:
                    return value

        return None

    # -- REST helpers --------------------------------------------------------

    def _get_json(self, url: str, *, params: Optional[Dict[str, Any]] = None) -> Any:
        """GET ``url`` (auto-login first) and return the parsed JSON body.

        Raises ``PropelioScraperError`` on transport failures, non-2xx
        responses, or non-JSON bodies.
        """
        self.login()
        try:
            response = self.session.get(url, params=params, timeout=self.timeout)
        except requests.RequestException as exc:
            raise PropelioScraperError(
                f"GET network error against {url}: {exc}"
            ) from exc

        if not response.ok:
            raise PropelioScraperError(
                f"GET {url} returned HTTP {response.status_code} "
                f"({response.reason}): {response.text[:500]}"
            )

        try:
            return response.json()
        except ValueError as exc:
            raise PropelioScraperError(
                f"GET {url} returned non-JSON: {response.text[:500]}"
            ) from exc

    def _parcel_suggest_try(self, tier: str, address: str) -> Optional[str]:
        """GET ``/parcels/v1/parcels/suggest/{tier}?query=...``; first ``items[].id``."""
        url = f"{PARCEL_SUGGEST_BASE}/{tier}"
        self.login()
        try:
            response = self.session.get(
                url, params={"query": address}, timeout=self.timeout,
            )
        except requests.RequestException as exc:
            logger.warning("Parcel suggest %s GET network error: %s", tier, exc)
            return None

        snippet = repr(response.text[:300])
        logger.info(
            "Parcel suggest %s -> %s [HTTP %s] body[:300]=%s",
            tier, response.url, response.status_code, snippet,
        )

        if not response.ok:
            logger.warning(
                "Parcel suggest %s non-OK HTTP %s", tier, response.status_code,
            )
            return None

        try:
            payload = response.json()
        except ValueError:
            logger.warning("Parcel suggest %s returned non-JSON body", tier)
            return None

        if not isinstance(payload, dict):
            return None
        items = payload.get("items")
        if not isinstance(items, list) or not items:
            return None
        first = items[0]
        if not isinstance(first, dict):
            return None
        pid = first.get("id")
        if not pid:
            return None
        logger.info(
            "Parcel suggest %s: using first match id=%r score=%r label=%r",
            tier, pid, first.get("score"), first.get("address"),
        )
        return str(pid)

    def _suggest_parcel_uuid(self, address: str) -> str:
        """Walk ``exact`` → ``close`` → ``fuzzy`` until one returns an id."""
        for tier in ("exact", "close", "fuzzy"):
            pid = self._parcel_suggest_try(tier, address)
            if pid:
                return pid
        raise PropelioScraperError(
            f"No parcel match for {address!r} (suggest exact / close / fuzzy "
            f"all returned no items)."
        )

    def _legacy_cma_probe(self, cma_id: str) -> bool:
        """Lightweight GET to see whether ``/legacy/cma/{cma_id}`` responds OK."""
        url = CMA_URL_TEMPLATE.format(lead_id=cma_id)
        self.login()
        try:
            response = self.session.get(url, timeout=self.timeout)
        except requests.RequestException as exc:
            logger.warning("CMA probe GET failed for %s: %s", url, exc)
            return False
        snippet = repr(response.text[:120])
        logger.info(
            "CMA probe GET %s [HTTP %s] body[:120]=%s",
            url, response.status_code, snippet,
        )
        return bool(response.ok)

    def _get_parcel_detail(self, parcel_uuid: str) -> Dict[str, Any]:
        """GET ``/parcels/v1/parcels/{parcel_uuid}`` (normalized address + lot)."""
        url = PARCEL_DETAIL_TEMPLATE.format(parcel_id=parcel_uuid)
        raw = self._get_json(url)
        return self._unwrap_list_envelope(raw, url, "Parcel detail")

    def _post_withaddress(
        self, body: Dict[str, Any], *, attempt_label: str,
    ) -> Tuple[int, Dict[str, Any], str]:
        """POST ``/legacy/leads/withaddress`` once, returning ``(status, json, raw)``.

        Logs the **full** response text and parsed JSON. Raises
        :class:`PropelioScraperError` on transport / non-(200|409) /
        non-JSON / non-object responses.
        """
        self.login()
        try:
            response = self.session.post(
                LEGACY_LEADS_WITHADDRESS_URL,
                json=body,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise PropelioScraperError(
                f"withaddress {attempt_label} POST network error: {exc}"
            ) from exc

        logger.info(
            "withaddress %s -> %s [HTTP %s] full body: %s",
            attempt_label,
            LEGACY_LEADS_WITHADDRESS_URL,
            response.status_code,
            response.text,
        )

        if response.status_code not in (200, 409):
            raise PropelioScraperError(
                f"withaddress {attempt_label} returned HTTP "
                f"{response.status_code}: {response.text[:500]}"
            )

        try:
            payload = response.json()
        except ValueError as exc:
            raise PropelioScraperError(
                f"withaddress {attempt_label} returned non-JSON: "
                f"{response.text[:500]}"
            ) from exc

        if not isinstance(payload, dict):
            raise PropelioScraperError(
                f"withaddress {attempt_label} JSON was not an object: "
                f"{str(payload)[:300]}"
            )

        logger.info(
            "withaddress %s parsed JSON: %s",
            attempt_label,
            json.dumps(payload, default=str),
        )
        return response.status_code, payload, response.text

    def _lead_id_from_withaddress(self, body: Dict[str, Any]) -> str:
        """Two-step ``/legacy/leads/withaddress`` flow with ``confirmationKey``.

        Step 1
            POST ``body`` (no ``confirmationKey``). Expected HTTP 409 with
            ``error.confirmationKey`` set. If the server happens to return
            200 with a usable ``result.id`` already, short-circuits and
            returns that lead id.

        Step 2
            POST the same ``body`` with ``confirmationKey`` added. Expected
            HTTP 200 with ``{"result": {"id": ...}}``. A second 409 (or any
            response without a usable ``result.id``) is logged and raised.
        """
        status1, payload1, raw1 = self._post_withaddress(
            body, attempt_label="step1",
        )

        if status1 == 200:
            lead_id = _extract_lead_id_from_withaddress_json(payload1)
            if lead_id:
                logger.info(
                    "withaddress step1 returned lead_id=%s without "
                    "needing a confirmationKey",
                    lead_id,
                )
                return lead_id
            raise PropelioScraperError(
                "withaddress step1 returned HTTP 200 but no extractable "
                f"lead id; keys={list(payload1.keys())}; raw={raw1!r}"
            )

        confirmation_key = _extract_confirmation_key_from_withaddress_json(
            payload1
        )
        if not confirmation_key:
            raise PropelioScraperError(
                "withaddress step1 returned HTTP 409 but no "
                "'confirmationKey' field was found in the response; "
                f"keys={list(payload1.keys())}; raw={raw1!r}"
            )
        logger.info(
            "withaddress step1 409 -> retrying with confirmationKey=%r",
            confirmation_key,
        )

        body2 = dict(body)
        body2["confirmationKey"] = confirmation_key

        status2, payload2, raw2 = self._post_withaddress(
            body2, attempt_label="step2",
        )

        if status2 == 409:
            raise PropelioScraperError(
                "withaddress step2 still returned HTTP 409 after sending "
                f"confirmationKey={confirmation_key!r}; raw={raw2!r}"
            )

        lead_id = _extract_lead_id_from_withaddress_json(payload2)
        if lead_id:
            return lead_id

        raise PropelioScraperError(
            "withaddress step2 returned no extractable lead id "
            f"(HTTP {status2}); keys={list(payload2.keys())}; raw={raw2!r}"
        )

    def find_lead_id(
        self, address: str,
    ) -> Tuple[str, Optional[float], Dict[str, Any]]:
        """Resolve ``address`` to ``(lead_id, subject_lot_sqft, parcel_bundle)``.

        ``parcel_bundle`` is ``{"valuation": {...}, "enrichment": {...}}``
        (see :func:`_parcel_subject_enrichment`).

        Steps:

            1. ``GET …/suggest/exact`` (then ``close``, ``fuzzy``) → parcel UUID.
            2. ``GET …/parcels/v1/parcels/{uuid}`` → address + lat/lon +
               lot fields + ``valuationDetail.estimate``.
            3. ``POST /legacy/leads/withaddress`` twice: the first call
               returns HTTP 409 with ``error.confirmationKey``; we POST
               again with that key included to get HTTP 200 +
               ``result.id`` (the lead id).
            4. Subject lot: ``landDetail.lotArea`` or ``summary.lotSqft``
               when present.
            5. Subject valuation: ``valuationDetail.estimate.value`` plus
               optional ``low`` / ``high`` (always returned as a dict
               with all three keys; values may be ``None``).

        Logs a probe ``GET /legacy/cma/{lead_id}``. No hardcoded lead-id shortcut.
        """
        parcel_uuid = self._suggest_parcel_uuid(address)
        parcel = self._get_parcel_detail(parcel_uuid)
        subject_lot = _parcel_subject_lot_sqft(parcel)
        if subject_lot is not None:
            logger.info("Parcel detail subject lot (sq ft): %s", subject_lot)
        valuation = _parcel_valuation_estimate(parcel)
        if any(v is not None for v in valuation.values()):
            logger.info(
                "Parcel valuation: estimate=%s low=%s high=%s",
                valuation["estimate"], valuation["low"], valuation["high"],
            )

        body = _withaddress_payload_from_parcel(parcel)
        lead_id = self._lead_id_from_withaddress(body)

        self._legacy_cma_probe(lead_id)
        logger.info(
            "Resolved address %r -> lead_id %r (parcel uuid=%r, subject_lot=%s)",
            address, lead_id, parcel_uuid, subject_lot,
        )
        enrichment = _parcel_subject_enrichment(parcel, valuation)
        parcel_bundle = {"valuation": valuation, "enrichment": enrichment}
        return lead_id, subject_lot, parcel_bundle

    # -- CMA / lead helpers --------------------------------------------------

    def get_cma(self, lead_id: str) -> Dict[str, Any]:
        """Fetch the CMA payload for ``lead_id`` from ``/legacy/cma/{lead_id}``.

        Propelio wraps the CMA in a single-element list, e.g.::

            [{"id": "1528869", "data": {"sales": [...]}, ...}]

        We unwrap the first element here so callers (and
        ``_parse_cma_response``) can treat the result as a plain dict.
        A bare-dict response is also accepted in case the API ever
        changes the envelope.
        """
        url = CMA_URL_TEMPLATE.format(lead_id=lead_id)
        raw = self._get_json(url)
        first = self._unwrap_list_envelope(raw, url, "CMA response")
        wrapper_id = first.get("id") if isinstance(raw, list) else None
        if isinstance(raw, list):
            logger.info(
                "Unwrapped CMA list response (len=%d, first.id=%s) for lead_id=%s",
                len(raw), wrapper_id, lead_id,
            )
        return first

    def get_lead_details(self, lead_id: str) -> Dict[str, Any]:
        """Fetch the full lead record from ``/legacy/leads/{lead_id}``.

        Used by :func:`_extract_subject` as a fallback when the CMA
        payload doesn't carry the subject's lot size. The response is
        unwrapped from any list envelope (mirroring ``get_cma``) so the
        caller gets a flat dict it can introspect.
        """
        url = f"{self.base_url}/legacy/leads/{lead_id}"
        raw = self._get_json(url)
        details = self._unwrap_list_envelope(raw, url, "Lead details")
        logger.info(
            "Lead details for lead_id=%s: top keys=%s",
            lead_id, list(details.keys()),
        )
        return details

    @staticmethod
    def _unwrap_list_envelope(raw: Any, url: str, what: str) -> Dict[str, Any]:
        """Coerce a list-or-dict response to a flat dict, raising on edges.

        Used by :meth:`get_cma` and :meth:`get_lead_details` because
        Propelio sometimes wraps single-record responses in
        ``[{...}]``.
        """
        if isinstance(raw, list):
            if not raw:
                raise PropelioScraperError(
                    f"{what} from {url} was an empty list"
                )
            first = raw[0]
            if not isinstance(first, dict):
                raise PropelioScraperError(
                    f"{what} list at {url} contained non-dict element: "
                    f"{str(first)[:300]}"
                )
            return first
        if isinstance(raw, dict):
            return raw
        raise PropelioScraperError(
            f"{what} from {url} was neither dict nor list: "
            f"{str(raw)[:300]}"
        )


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------

def login_propelio(
    username: Optional[str] = None,
    password: Optional[str] = None,
    *,
    timeout: int = config.HTTP_TIMEOUT_SECONDS,
) -> PropelioClient:
    """Log into Propelio's private API and return an authenticated client.

    Raises ``PropelioScraperError`` if credentials are missing or login
    fails for any reason.
    """
    user = username or config.PROPELIO_USERNAME
    pw = password or config.PROPELIO_PASSWORD
    client = PropelioClient(user, pw, timeout=timeout)
    client.login()
    return client


def search_properties(
    address: str,
    radius: float = config.RADIUS_MILES,
    *,
    client: Optional[PropelioClient] = None,
    username: Optional[str] = None,
    password: Optional[str] = None,
) -> Tuple[Property, List[Property]]:
    """Return ``(subject, comps)`` for ``address``.

    Flow: resolve the address to a ``lead_id`` (parcel suggest → parcel
    detail → ``/legacy/leads/withaddress``), then fetch ``/legacy/cma/{id}``
    and parse the result.

    ``subject`` is always a :class:`Property` so callers don't need to
    null-check; if the CMA payload doesn't carry recoverable subject
    metadata, a synthetic Property with only ``address`` populated is
    returned and the comp engine will skip filtering (logged at WARNING).

    ``radius`` is *not* sent to the server — Propelio's CMA endpoint
    returns its own pre-selected comp set; ``comp_engine.filter_comps``
    enforces the radius threshold post-hoc. The kwarg is kept in the
    signature for compatibility with existing callers (``main.py``) and
    is logged for visibility.

    If ``client`` is omitted, a fresh authenticated client is built from
    the supplied credentials (or ``config.PROPELIO_USERNAME`` /
    ``PROPELIO_PASSWORD``).
    """
    if not address:
        raise ValueError("address must be a non-empty string")

    if client is None:
        client = login_propelio(username, password)

    logger.info(
        "Looking up Propelio lead for %r (filter radius=%.2f mi)",
        address, radius,
    )

    try:
        lead_id, subject_lot_sqft, parcel_bundle = client.find_lead_id(address)
        payload = client.get_cma(lead_id)
    except PropelioScraperError:
        raise
    except Exception as exc:
        logger.exception("Unexpected error while querying Propelio")
        raise PropelioScraperError(f"Search failed: {exc}") from exc

    subject, comps = _parse_cma_response(
        payload,
        searched_address=address,
        client=client,
        lead_id=lead_id,
        subject_lot_sqft=subject_lot_sqft,
    )
    if subject is None:
        subject = Property(address=address, source="user_input")
    valuation = parcel_bundle.get("valuation") if isinstance(parcel_bundle, dict) else None
    enrichment = parcel_bundle.get("enrichment") if isinstance(parcel_bundle, dict) else None
    if isinstance(valuation, dict) and any(v is not None for v in valuation.values()):
        subject.extra["valuation"] = valuation
    if isinstance(enrichment, dict):
        subject.extra["parcel_enrichment"] = enrichment

    logger.info(
        "Returned subject + %d comp(s) for address=%r (lead_id=%s)",
        len(comps), address, lead_id,
    )
    return subject, comps


async def scrape_comps(
    address: str,
    radius: float = config.RADIUS_MILES,
    username: Optional[str] = None,
    password: Optional[str] = None,
) -> Tuple[Property, List[Property]]:
    """Async-compatible wrapper used by ``main.py``.

    Returns ``(subject, comps)``. Internally runs the blocking
    ``requests`` calls in a worker thread via :func:`asyncio.to_thread`,
    so the event loop stays responsive even though the underlying
    client is synchronous.
    """
    return await asyncio.to_thread(
        search_properties,
        address,
        radius,
        client=None,
        username=username,
        password=password,
    )


# ---------------------------------------------------------------------------
# Response parsing
# ---------------------------------------------------------------------------

def _parse_cma_response(
    payload: Dict[str, Any],
    searched_address: str = "",
    *,
    client: Optional["PropelioClient"] = None,
    lead_id: Optional[str] = None,
    subject_lot_sqft: Optional[float] = None,
) -> Tuple[Optional[Property], List[Property]]:
    """Normalize a ``/legacy/cma/{lead_id}`` record into ``(subject, comps)``.

    Expected shape (after :meth:`PropelioClient.get_cma` has already
    unwrapped the outer ``[...]`` list)::

        {"id": "1528869", "params": {...}, "data": {"sales": [...]}, ...}

    A flat ``payload["sales"]`` is also accepted as a fallback for the
    comps list. The ``subject`` Property is recovered (best-effort) so
    the comp engine has a lot size to score against; see
    :func:`_extract_subject` for the strategies tried. When ``client``
    and ``lead_id`` are provided, ``_extract_subject`` will fall back
    to ``GET /legacy/leads/{lead_id}`` if the CMA payload alone doesn't
    carry a usable lot size. ``subject_lot_sqft`` comes from the parcel
    detail endpoint when available and fills a missing subject lot size.
    """
    if isinstance(payload, dict):
        logger.info("CMA payload keys: %s", list(payload.keys()))
        logger.info("CMA params: %s", payload.get("params"))

    container: Any = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(container, dict):
        container = payload

    sales = container.get("sales") if isinstance(container, dict) else None
    if not isinstance(sales, list):
        top_keys = (
            list(payload.keys())
            if isinstance(payload, dict)
            else type(payload).__name__
        )
        data_keys = (
            list(container.keys())
            if isinstance(container, dict) and container is not payload
            else None
        )
        logger.info(
            "No 'sales' array in CMA response (top keys=%s, data keys=%s)",
            top_keys, data_keys,
        )
        sales = []

    comps: List[Property] = []
    for raw in sales:
        if not isinstance(raw, dict):
            continue
        try:
            comps.append(_parse_property(raw))
        except (TypeError, ValueError) as exc:
            logger.debug("Skipping malformed sale %r: %s", raw, exc)
            continue
    logger.info("Parsed %d comp record(s) from CMA sales array", len(comps))

    subject = _extract_subject(
        payload,
        searched_address,
        sales,
        client=client,
        lead_id=lead_id,
        subject_lot_sqft=subject_lot_sqft,
    )
    if subject:
        logger.info(
            "Resolved subject: address=%r lot_size=%s sqft=%s year_built=%s",
            subject.address, subject.lot_size, subject.sqft, subject.year_built,
        )
    else:
        logger.warning(
            "Could not resolve subject property metadata from CMA payload; "
            "filter_comps will receive an empty subject and skip filtering."
        )

    return subject, comps


def _extract_subject(
    payload: Dict[str, Any],
    searched_address: str,
    sales: List[Any],
    *,
    client: Optional["PropelioClient"] = None,
    lead_id: Optional[str] = None,
    subject_lot_sqft: Optional[float] = None,
) -> Optional[Property]:
    """Heuristically recover the subject Property from a CMA payload.

    Strategies, tried in order:

        1. ``payload["data"]["subject"]`` / ``payload["subject"]`` —
           an explicit subject sub-object.
        2. ``payload["data"]["property"]`` / ``payload["property"]`` —
           alternate name some Propelio responses use.
        3. Top-level ``lot_size`` / ``lotSize`` / ``sqft`` / etc. on
           ``payload["data"]`` or ``payload`` itself (CMAs sometimes
           include the subject's stats at the envelope level).
        4. First ``data.sales[i]`` whose ``address.line1`` matches
           BOTH the street number AND the street name from
           ``searched_address``.
        5. ``GET /legacy/leads/{lead_id}`` — full property record
           fallback. Fires when ``client`` and ``lead_id`` are
           supplied AND the prior strategies didn't produce a
           ``lot_size``.

    When ``subject_lot_sqft`` is set (from ``GET …/parcels/{uuid}``), it
    supplies or patches ``lot_size`` before the lead-details fallback.

    Returns ``None`` if nothing usable is found across all strategies.
    """
    subject = _subject_from_cma_payload(payload, searched_address, sales)

    if subject_lot_sqft is not None and subject_lot_sqft > 0:
        if subject is None:
            logger.info(
                "Subject lot_size from parcel detail API: %s", subject_lot_sqft,
            )
            subject = Property(
                address=searched_address or "Unknown",
                lot_size=float(subject_lot_sqft),
                source="propelio_parcel",
            )
        elif subject.lot_size is None:
            logger.info(
                "Patched subject lot_size from parcel detail API: %s",
                subject_lot_sqft,
            )
            subject = Property(
                address=subject.address,
                price=subject.price,
                lot_size=float(subject_lot_sqft),
                sqft=subject.sqft,
                year_built=subject.year_built,
                status=subject.status,
                neighborhood=subject.neighborhood,
                source=subject.source,
                extra=dict(subject.extra),
            )

    needs_fallback = (
        client is not None
        and lead_id
        and (subject is None or subject.lot_size is None)
    )
    if needs_fallback:
        subject = _augment_subject_from_lead_details(
            subject, searched_address, client, str(lead_id),
        )

    return subject


def _subject_from_cma_payload(
    payload: Dict[str, Any],
    searched_address: str,
    sales: List[Any],
) -> Optional[Property]:
    """Strategies 1–4 from :func:`_extract_subject` (CMA payload only)."""
    if not isinstance(payload, dict):
        return None
    data = payload.get("data") if isinstance(payload.get("data"), dict) else {}

    # Strategy 1 + 2: explicit subject node under data or top-level.
    for source in (data, payload):
        if not isinstance(source, dict):
            continue
        for key in ("subject", "subjectProperty", "subject_property", "property"):
            node = source.get(key)
            if isinstance(node, dict) and _looks_like_property_dict(node):
                try:
                    parsed = _parse_property(node)
                except (TypeError, ValueError):
                    continue
                logger.info(
                    "Subject extracted via explicit %r node", key
                )
                return _override_address(parsed, searched_address)

    # Strategy 3: synthesize from top-level fields.
    for source_name, source in (("data", data), ("payload", payload)):
        if not isinstance(source, dict):
            continue
        lot_size = _lot_size_from_dict(source)
        if lot_size:
            logger.info(
                "Subject extracted from top-level %s.lot_size=%s",
                source_name, lot_size,
            )
            return Property(
                address=searched_address or "Unknown",
                lot_size=lot_size,
                sqft=_sqft_from_dict(source),
                year_built=_year_built_from_dict(source),
                neighborhood=_neighborhood_from_dict(source),
                source="propelio_subject",
            )

    # Strategy 4: strict street-number + street-name match against sales.
    # Number-only matches caused false positives — e.g. "3761 Dunhaven"
    # incorrectly matched a "3761 Seguin Dr" comp. Both components now
    # have to line up.
    needle_number, needle_name = _parse_address_components(searched_address)
    if needle_number and needle_name:
        for raw in sales:
            if not isinstance(raw, dict):
                continue
            sale_addr = raw.get("address")
            line1 = ""
            if isinstance(sale_addr, dict):
                line1 = sale_addr.get("line1") or sale_addr.get("street") or ""
            elif isinstance(sale_addr, str):
                line1 = sale_addr
            if not _line1_matches(line1, needle_number, needle_name):
                continue
            try:
                parsed = _parse_property(raw)
            except (TypeError, ValueError):
                continue
            logger.info(
                "Subject matched in sales via number=%r + name=%r -> %r",
                needle_number, needle_name, parsed.address,
            )
            return _override_address(parsed, searched_address)
        logger.info(
            "Strategy 4: no sale matched both number=%r and name=%r "
            "across %d sales; deferring to lead-details fallback.",
            needle_number, needle_name, len(sales),
        )
    elif searched_address:
        logger.info(
            "Strategy 4: could not parse street number + name from %r; "
            "skipping sales-array matching.", searched_address,
        )

    return None


def _augment_subject_from_lead_details(
    subject: Optional[Property],
    searched_address: str,
    client: "PropelioClient",
    lead_id: str,
) -> Optional[Property]:
    """Strategy 5: fetch ``/legacy/leads/{lead_id}`` and patch missing fields.

    If ``subject`` is ``None``, builds a fresh ``Property`` from
    whatever the lead record provides. If ``subject`` is non-``None``,
    patches in any missing ``lot_size`` / ``sqft`` / ``year_built`` /
    ``neighborhood`` without overwriting fields already populated.

    Falls back to returning ``subject`` unchanged on:
      - ``PropelioScraperError`` from the lead-details request, or
      - a lead record that has no recognizable lot size.
    """
    try:
        details = client.get_lead_details(lead_id)
    except PropelioScraperError as exc:
        logger.warning(
            "Lead-details fallback failed for lead_id=%s: %s", lead_id, exc,
        )
        return subject

    details_lot = _lot_size_from_dict(details)
    if not details_lot:
        logger.warning(
            "Lead details for lead_id=%s contained no recognizable lot size "
            "(top keys=%s); leaving subject as-is.",
            lead_id, list(details.keys()),
        )
        return subject

    details_sqft = _sqft_from_dict(details)
    details_year = _year_built_from_dict(details)
    details_neighborhood = _neighborhood_from_dict(details)

    if subject is None:
        logger.info(
            "Subject built from /legacy/leads/%s (lot_size=%s, sqft=%s, "
            "year_built=%s, neighborhood=%s)",
            lead_id, details_lot, details_sqft, details_year,
            details_neighborhood,
        )
        return Property(
            address=searched_address or "Unknown",
            lot_size=details_lot,
            sqft=details_sqft,
            year_built=details_year,
            neighborhood=details_neighborhood,
            source="propelio_subject",
        )

    logger.info(
        "Patched subject from /legacy/leads/%s (lot_size %s -> %s)",
        lead_id, subject.lot_size, subject.lot_size or details_lot,
    )
    return Property(
        address=subject.address,
        price=subject.price,
        lot_size=subject.lot_size or details_lot,
        sqft=subject.sqft or details_sqft,
        year_built=subject.year_built or details_year,
        status=subject.status,
        neighborhood=subject.neighborhood or details_neighborhood,
        source=subject.source,
        extra=dict(subject.extra),
    )


def _looks_like_property_dict(node: Dict[str, Any]) -> bool:
    """True if ``node`` carries enough fields to plausibly be a property."""
    if not isinstance(node, dict):
        return False
    interesting = (
        "address", "lot_size", "lotSize", "lot_sqft", "lotSqft",
        "sqft", "livingArea", "year_built", "yearBuilt",
    )
    return any(node.get(k) is not None for k in interesting)


def _street_number(address: str) -> Optional[str]:
    """Return the leading numeric token of ``address``, e.g. ``"3761"``."""
    if not address:
        return None
    first = address.strip().split(maxsplit=1)[0] if address.strip() else ""
    if first.isdigit():
        return first
    return None


# Common US street-type tokens we strip when computing the street name body
# so "3761 Dunhaven Lane" matches Propelio's "3761 Dunhaven Ln" and similar.
_STREET_TYPE_TOKENS = frozenset({
    "st", "street",
    "ave", "av", "avenue",
    "rd", "road",
    "dr", "drive",
    "ln", "lane",
    "blvd", "boulevard",
    "ct", "court",
    "pl", "place",
    "way",
    "cir", "circle",
    "ter", "terrace",
    "pkwy", "parkway",
    "hwy", "highway",
    "trl", "trail",
    "loop",
    "row",
    "sq", "square",
    "xing", "crossing",
})


def _parse_address_components(address: str) -> Tuple[Optional[str], Optional[str]]:
    """Extract ``(street_number, street_name)`` from a freeform address.

    The street name is the part between the street number and the
    trailing street-type token (``Lane`` / ``Dr`` / ``Rd`` / etc.),
    lowercased and joined with single spaces. Anything after the first
    comma (city / state / zip) is dropped before parsing.

    Examples::

        "3761 Dunhaven Lane, Dallas, TX 75220" -> ("3761", "dunhaven")
        "1234 Old Mill Rd"                     -> ("1234", "old mill")
        "3761 N Dunhaven Ln, Dallas, TX"       -> ("3761", "n dunhaven")
        "Foo Bar"                              -> (None, None)
    """
    if not address:
        return None, None

    head = address.strip().split(",", maxsplit=1)[0]
    tokens = head.split()
    if len(tokens) < 2 or not tokens[0].isdigit():
        return None, None

    street_number = tokens[0]
    rest = tokens[1:]
    if rest and rest[-1].rstrip(".").lower() in _STREET_TYPE_TOKENS:
        rest = rest[:-1]

    if not rest:
        return street_number, None

    street_name = " ".join(rest).lower()
    return street_number, street_name


def _line1_matches(line1: Any, needle_number: str, needle_name: str) -> bool:
    """True if ``line1`` starts with ``needle_number`` AND contains ``needle_name``.

    Number boundary is enforced (so ``"3761"`` doesn't match ``"37615"``)
    by requiring whitespace or punctuation immediately after the digits.
    Both fields are matched case-insensitively.
    """
    if not line1:
        return False
    text = str(line1).strip().lower()
    if not text or not text.startswith(needle_number):
        return False
    after = text[len(needle_number):len(needle_number) + 1]
    if after and not (after.isspace() or after in ",.-"):
        return False
    return needle_name in text


def _override_address(prop: Property, searched_address: str) -> Property:
    """Return a copy of ``prop`` with ``address`` replaced by the user's input.

    Preserves all other fields. We prefer the user's address spelling
    over Propelio's normalization for display consistency.
    """
    if not searched_address:
        return prop
    return Property(
        address=searched_address,
        price=prop.price,
        lot_size=prop.lot_size,
        sqft=prop.sqft,
        year_built=prop.year_built,
        status=prop.status,
        neighborhood=prop.neighborhood,
        source=prop.source,
        extra=dict(prop.extra),
    )


def _parse_property(raw: Dict[str, Any]) -> Property:
    """Build a ``Property`` from one ``data.sales[i]`` record.

    Tolerates several Propelio field-name spellings: ``salePrice`` /
    ``sale_price`` / ``soldPrice``, ``livingArea`` / ``sqft``,
    ``lotSize`` / ``lot_size`` / ``lotSqft``, ``yearBuilt`` /
    ``year_built``, ``status`` / ``saleStatus`` / ``sale_status``.
    Stores beds/baths/MLS/lat/lon and any unmapped keys on ``extra``.
    """
    addr_obj = raw.get("address") if isinstance(raw.get("address"), dict) else {}

    address = _format_address(addr_obj) or str(raw.get("address") or "Unknown")

    price = _to_float(
        raw.get("price")
        or raw.get("sale_price") or raw.get("salePrice")
        or raw.get("sold_price") or raw.get("soldPrice")
        or raw.get("close_price") or raw.get("closePrice")
        or raw.get("list_price") or raw.get("listPrice")
    )
    sqft = _sqft_from_dict(raw)
    lot_size = _lot_size_from_dict(raw)
    year_built = _year_built_from_dict(raw)
    status = _normalize_status(
        raw.get("status")
        or raw.get("sale_status") or raw.get("saleStatus")
        or raw.get("listing_status") or raw.get("listingStatus")
        or raw.get("mls_status") or raw.get("mlsStatus")
        or raw.get("property_status") or raw.get("propertyStatus")
        or raw.get("transaction_status") or raw.get("transactionStatus")
        or raw.get("close_status") or raw.get("closeStatus")
    )
    neighborhood = (
        addr_obj.get("subdivision")
        or addr_obj.get("neighborhood")
        or raw.get("subdivision")
        or raw.get("neighborhood")
    )

    extra: Dict[str, Any] = {}
    cp = _to_float(raw.get("close_price") or raw.get("closePrice"))
    if cp is not None:
        extra["close_price"] = cp
    elif price is not None and status == "sold":
        extra["close_price"] = price

    for key in ("beds", "baths", "mls", "source"):
        value = raw.get(key)
        if value is not None and value != "":
            extra[key] = value
    if addr_obj:
        lat = addr_obj.get("lat") or addr_obj.get("latitude")
        lon = addr_obj.get("lon") or addr_obj.get("lng") or addr_obj.get("longitude")
        if lat is not None:
            extra["lat"] = lat
        if lon is not None:
            extra["lon"] = lon

    return Property(
        address=address,
        price=price,
        lot_size=lot_size,
        sqft=sqft,
        year_built=year_built,
        status=status,
        neighborhood=neighborhood,
        source="propelio",
        extra=extra,
    )


def _format_address(addr_obj: Dict[str, Any]) -> Optional[str]:
    """Render a Propelio address sub-object as ``"line1, city, ST zip"``."""
    if not addr_obj:
        return None
    line1 = (addr_obj.get("line1") or addr_obj.get("street") or "").strip()
    city = (addr_obj.get("city") or "").strip()
    state = (addr_obj.get("state") or "").strip()
    zip_code = (
        addr_obj.get("zip")
        or addr_obj.get("zipcode")
        or addr_obj.get("postal_code")
        or ""
    )
    zip_code = str(zip_code).strip()

    state_zip = " ".join(part for part in (state, zip_code) if part)
    parts = [p for p in (line1, city, state_zip) if p]
    return ", ".join(parts) if parts else None


def _normalize_status(value: Any) -> Optional[str]:
    """Map MLS-style strings to ``sold`` / ``pending`` / ``for_sale``.

    Order matters: check **inactive / withdrawn** before **active** so
    ``inactive`` is not misclassified as ``for_sale``. ``sold`` /
    ``pending`` / ``for_sale`` are never used to *drop* comps in
    ``comp_engine`` — every normalized value is kept on
    :attr:`Property.status` for reporting.
    """
    if value is None:
        return None
    cleaned = str(value).strip().lower()
    if not cleaned:
        return None

    inactive_markers = (
        "inactive", "withdrawn", "cancelled", "canceled", "expired",
        "terminated", "off market", "not available",
    )
    if any(m in cleaned for m in inactive_markers):
        return None

    pending_markers = (
        "pending", "pend", "contingent", "under contract", "option",
        "backup", "hold",
    )
    if any(m in cleaned for m in pending_markers):
        return "pending"

    if re.search(r"\b(sold|closed|settled|recorded)\b", cleaned):
        return "sold"
    if re.search(r"\b(sls|sld)\b", cleaned):
        return "sold"

    # ``for_sale`` bucket: active listings and anything still on-market.
    active_markers = (
        "active", "for sale", "for lease", "listing", "listed",
        "coming soon", "new construction",
    )
    if any(m in cleaned for m in active_markers):
        return "for_sale"

    # Fallback: preserve unknown vendor codes verbatim (lowercased).
    return cleaned


def _to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        cleaned = str(value).replace("$", "").replace(",", "").strip()
        return float(cleaned.split()[0]) if cleaned else None
    except (ValueError, IndexError):
        return None


def _to_int(value: Any) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    try:
        return int(str(value).replace(",", "").strip())
    except (ValueError, AttributeError):
        return None


def _candidate_sources(d: Any) -> List[Dict[str, Any]]:
    """Return the dicts to inspect for property fields.

    For lead-details responses the interesting fields can live at the
    top level, under ``data``, or under ``params``; we walk each
    container in that priority order.
    """
    if not isinstance(d, dict):
        return []
    sources: List[Dict[str, Any]] = [d]
    for key in ("data", "params"):
        nested = d.get(key)
        if isinstance(nested, dict):
            sources.append(nested)
    return sources


_SQFT_PER_ACRE = 43_560.0


def _lot_size_from_dict(d: Any) -> Optional[float]:
    """Lot size in square feet, tolerant of name spellings + ``acres``.

    Priority order, applied per container (top-level, then ``d['data']``,
    then ``d['params']``):

      1. ``lot_size_sqft`` — Propelio's CMA sale-row spelling, direct
         square-foot value.
      2. ``lot_size_acres`` — Propelio's CMA sale-row acres value,
         converted via ``× 43,560``.
      3. Other direct-sqft aliases: ``lot_sqft`` / ``lotSqft`` /
         ``lot_size`` / ``lotSize`` / ``lot_square_feet`` /
         ``lotSquareFeet``.
      4. Other acres aliases: ``acres`` / ``acre`` / ``lot_acres`` /
         ``lotAcres`` (also × 43,560).
    """
    for source in _candidate_sources(d):
        direct = _to_float(source.get("lot_size_sqft"))
        if direct:
            return direct

        acres = _to_float(source.get("lot_size_acres"))
        if acres:
            return round(acres * _SQFT_PER_ACRE, 2)

        for key in (
            "lot_sqft", "lotSqft",
            "lot_size", "lotSize",
            "lot_square_feet", "lotSquareFeet",
        ):
            value = _to_float(source.get(key))
            if value:
                return value

        for key in ("acres", "acre", "lot_acres", "lotAcres"):
            value = _to_float(source.get(key))
            if value:
                return round(value * _SQFT_PER_ACRE, 2)
    return None


def _sqft_from_dict(d: Any) -> Optional[float]:
    """Living area in square feet (sqft / square_feet / livingArea / gla)."""
    for source in _candidate_sources(d):
        for key in (
            "sqft", "square_feet", "squareFeet",
            "living_area", "livingArea",
            "gla",
        ):
            value = _to_float(source.get(key))
            if value:
                return value
    return None


def _year_built_from_dict(d: Any) -> Optional[int]:
    """Year built (year_built / yearBuilt / year)."""
    for source in _candidate_sources(d):
        for key in ("year_built", "yearBuilt", "year"):
            value = _to_int(source.get(key))
            if value:
                return value
    return None


def _neighborhood_from_dict(d: Any) -> Optional[str]:
    """Neighborhood / subdivision, checking the address sub-object too."""
    for source in _candidate_sources(d):
        for key in ("neighborhood", "subdivision"):
            value = source.get(key)
            if value:
                return str(value)
        addr = source.get("address")
        if isinstance(addr, dict):
            for key in ("subdivision", "neighborhood"):
                value = addr.get(key)
                if value:
                    return str(value)
    return None


def _redact_proxy(url: Optional[str]) -> str:
    """Return a proxy URL with the password masked for safe logging."""
    if not url:
        return "<none>"
    try:
        scheme, _, rest = url.partition("://")
        if "@" not in rest:
            return url
        creds, _, host = rest.rpartition("@")
        user, _, _ = creds.partition(":")
        return f"{scheme}://{user}:***@{host}"
    except Exception:
        return "<proxy>"


