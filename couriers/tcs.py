"""TCS Express (Pakistan) tracking.

The public page https://www.tcsexpress.com/track/<number> is a React app that
calls its own bridge endpoint, which forwards to TCS's tracking API. We call the
same bridge the website uses, so results match what you see in the browser.
"""

import re
from datetime import datetime

import requests

from .base import Checkpoint, TrackingResult, TrackingError

NAME = "TCS"
BRIDGE_URL = "https://www.tcsexpress.com/apibridge"
TIMEOUT = 30

HEADERS = {
    "Content-Type": "application/json",
    "Origin": "https://www.tcsexpress.com",
    "Referer": "https://www.tcsexpress.com/track/",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}

# Statuses that mean the parcel reached the receiver. "undelivered" / "not delivered"
# / "attempt" are checked first so they never count as delivered.
NOT_DELIVERED_WORDS = ("undeliver", "not deliver", "attempt", "refus", "return", "rts", "hold")
DELIVERED_WORDS = ("delivered", "deliverd", "delivery done", "handed over")

DATE_FORMATS = (
    "%d-%b-%Y %I:%M %p", "%d-%b-%Y %I:%M:%S %p", "%d-%b-%Y %H:%M", "%d-%b-%Y %H:%M:%S",
    "%d/%m/%Y %I:%M %p", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y",
    "%d %b %Y %I:%M %p", "%d %b %Y %H:%M", "%d %b %Y", "%d-%b-%Y", "%d/%m/%Y",
    "%b %d %Y %I:%M %p", "%b %d %Y %H:%M", "%b %d %Y", "%B %d %Y %I:%M %p", "%B %d %Y",
    "%d %B %Y %I:%M %p", "%d %B %Y",
)


def parse_dt(text) -> datetime | None:
    """Best-effort parse of the date strings TCS returns. Returns None if unparseable."""
    if not text:
        return None
    s = str(text).strip()
    if s.endswith("Z"):
        s = s[:-1]
    # drop fractional seconds and weekday names, collapse spaces/commas
    s = re.sub(r"\.\d+$", "", s)
    s = re.sub(r"\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*", "", s, flags=re.I)
    s = re.sub(r"[,]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    # last try: ISO with microseconds
    try:
        return datetime.fromisoformat(str(text).strip().replace("Z", ""))
    except ValueError:
        return None


RETURNED_WORDS = ("returned to sender", "returned to shipper", "return to sender", "return to shipper", "rts")


def _first(items, number: str = ""):
    """First entry, preferring the one whose consignmentno matches `number`
    (a reference-number search can return several shipments)."""
    if isinstance(items, dict):
        return items
    if not isinstance(items, list) or not items:
        return {}
    if number:
        for it in items:
            if it and str(it.get("consignmentno") or "").strip() == number:
                return it
    return items[0] or {}


def _is_returned(status: str) -> bool:
    return any(w in (status or "").lower() for w in RETURNED_WORDS)


def fetch_raw(number: str) -> dict:
    payload = {
        "body": {
            "url": "trackapinew",
            "type": "GET",
            "headers": {},
            "param": f"consignee={number}",
            "payload": {},
        }
    }
    try:
        resp = requests.post(BRIDGE_URL, json=payload, headers=HEADERS, timeout=TIMEOUT)
    except requests.RequestException as exc:
        raise TrackingError(f"Could not reach TCS: {exc}") from exc
    if resp.status_code != 200:
        raise TrackingError(f"TCS returned HTTP {resp.status_code}")
    try:
        data = resp.json()
    except ValueError as exc:
        raise TrackingError("TCS returned a non-JSON response") from exc
    if not data.get("isSuccess"):
        raise TrackingError(f"TCS error: {data.get('responseData')}")
    return data.get("responseData") or {}


def _is_delivered(status: str) -> bool:
    s = (status or "").lower()
    if any(w in s for w in NOT_DELIVERED_WORDS):
        return False
    return any(w in s for w in DELIVERED_WORDS)


def track(number: str) -> TrackingResult:
    number = re.sub(r"\s+", "", str(number))
    data = fetch_raw(number)

    info = _first(data.get("shipmentinfo"), number)
    delivery = _first(data.get("deliveryinfo"), number)
    raw_cps = data.get("checkpoints") or []
    if number and any(str((cp or {}).get("consignmentno") or "") == number for cp in raw_cps):
        raw_cps = [cp for cp in raw_cps if str((cp or {}).get("consignmentno") or "") == number]
    summary_text = str(data.get("shipmentsummary") or "").strip()

    if not info and not raw_cps:
        return TrackingResult(
            courier=NAME, number=number, found=False, state="NOT FOUND",
            summary=summary_text or "No data found for this tracking number.", raw=data,
        )

    checkpoints = []
    for cp in raw_cps:
        raw_time = str(cp.get("datetime") or "").strip()
        checkpoints.append(Checkpoint(
            time=parse_dt(raw_time), raw_time=raw_time,
            status=str(cp.get("status") or "").strip(),
            detail=str(cp.get("recievedby") or cp.get("receivedby") or "").strip(),
        ))
    # newest first; keep original order for items without a parsable time
    if checkpoints and all(c.time for c in checkpoints):
        checkpoints.sort(key=lambda c: c.time, reverse=True)

    booking_date = parse_dt(info.get("bookingdate"))
    delivered_on = parse_dt(delivery.get("datetime")) if delivery else None
    received_by = str((delivery or {}).get("recievedby") or (delivery or {}).get("receivedby") or "").strip()

    latest = checkpoints[0] if checkpoints else None
    latest_status = latest.status if latest else ""
    latest_time = latest.time if latest else None

    # TCS's deliveryinfo block is the final outcome. It carries status "Delivered"
    # (code OK) or "Returned To Sender" (code RS), so check it before the checkpoints.
    delivery_status = str((delivery or {}).get("status") or "").strip()
    delivery_code = str((delivery or {}).get("code") or "").strip().upper()

    returned = (_is_returned(delivery_status) or delivery_code == "RS"
                or (not delivery and _is_returned(latest_status)))
    delivered = (not returned) and (
        _is_delivered(delivery_status) or delivery_code == "OK"
        or (not delivery_status and bool(delivery) and (delivered_on or received_by))
        or (not delivery and (_is_delivered(latest_status) or _is_delivered(summary_text)))
    )
    if (delivered or returned) and not delivered_on:
        hit = next((c for c in checkpoints
                    if (_is_returned(c.status) if returned else _is_delivered(c.status))), None)
        delivered_on = hit.time if hit else latest_time
        if hit and delivered and not received_by:
            received_by = hit.detail

    end = delivered_on or datetime.now()
    days_in_transit = (end.date() - booking_date.date()).days if booking_date else None

    state = "DELIVERED" if delivered else "RETURNED" if returned else "IN TRANSIT"
    if state == "DELIVERED":
        summary = "Delivered"
        if received_by:
            summary += f" to {received_by}"
        if delivered_on:
            summary += f" on {delivered_on.strftime('%d %b %Y %I:%M %p')}"
    elif state == "RETURNED":
        summary = "Returned to sender"
        if delivered_on:
            summary += f" on {delivered_on.strftime('%d %b %Y %I:%M %p')}"
    else:
        summary = latest_status or summary_text or "In transit"
        if latest_time:
            summary += f" ({latest_time.strftime('%d %b %Y %I:%M %p')})"
        elif latest and latest.raw_time:
            summary += f" ({latest.raw_time})"

    return TrackingResult(
        courier=NAME, number=number, found=True, state=state, summary=summary,
        origin=str(info.get("origin") or "").strip(),
        destination=str(info.get("destination") or "").strip(),
        booking_date=booking_date, delivered_on=delivered_on, received_by=received_by,
        latest_status=latest_status, latest_time=latest_time,
        days_in_transit=days_in_transit, checkpoints=checkpoints, raw=data,
    )
