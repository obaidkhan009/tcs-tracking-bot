"""Courier registry. Add a new courier by writing a module with a `track(number)`
function that returns a TrackingResult, then register it in COURIERS below."""

from .base import TrackingResult, TrackingError
from . import tcs

COURIERS = {
    "tcs": tcs,
}

ALIASES = {
    "tcs": "tcs",
    "tcs express": "tcs",
    "tcs courier": "tcs",
    "tcs couriers": "tcs",
}


def get_courier(name: str | None):
    """Return the courier module for a name from the sheet (case-insensitive).
    Defaults to TCS when the cell is empty."""
    key = (name or "tcs").strip().lower()
    key = ALIASES.get(key, key)
    if key not in COURIERS:
        raise TrackingError(
            f"Courier '{name}' is not supported yet. Supported: {', '.join(sorted(COURIERS))}"
        )
    return COURIERS[key]


__all__ = ["TrackingResult", "TrackingError", "get_courier", "COURIERS"]
