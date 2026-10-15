from dataclasses import dataclass, field
from datetime import datetime


class TrackingError(Exception):
    """Raised when a courier lookup fails (network, bad number, unsupported courier)."""


@dataclass
class Checkpoint:
    time: datetime | None
    raw_time: str
    status: str
    detail: str = ""


@dataclass
class TrackingResult:
    courier: str
    number: str
    found: bool
    state: str                      # DELIVERED | IN TRANSIT | NOT FOUND | ERROR
    summary: str = ""               # one-line status in plain words
    origin: str = ""
    destination: str = ""
    booking_date: datetime | None = None
    delivered_on: datetime | None = None
    received_by: str = ""
    latest_status: str = ""
    latest_time: datetime | None = None
    days_in_transit: int | None = None
    checkpoints: list[Checkpoint] = field(default_factory=list)
    raw: dict = field(default_factory=dict)

    def eta_text(self, expected_transit_days: int = 0) -> str:
        """Human text for 'when will it arrive'.

        TCS does not publish an expected delivery date, so by default this only
        states facts (days in transit, latest scan). If expected_transit_days > 0
        an estimate of booking date + that many days is added, clearly labelled."""
        if self.state == "DELIVERED":
            when = self.delivered_on.strftime("%d %b %Y %I:%M %p") if self.delivered_on else "unknown date"
            who = f" (received by {self.received_by})" if self.received_by else ""
            return f"Delivered on {when}{who}."
        if self.state == "RETURNED":
            when = f" on {self.delivered_on.strftime('%d %b %Y %I:%M %p')}" if self.delivered_on else ""
            return f"Returned to sender{when}. Contact the local TCS office."
        if self.state != "IN TRANSIT":
            return ""
        days = f"{self.days_in_transit} day(s)" if self.days_in_transit is not None else "unknown time"
        latest = self.latest_status or "no scan yet"
        if not expected_transit_days:
            return (f"In transit for {days}. Latest scan: {latest}. "
                    f"TCS does not give an expected delivery date.")
        if not self.booking_date:
            return f"In transit. Latest scan: {latest}. No booking date returned, so no estimate is possible."
        from datetime import timedelta
        due = self.booking_date + timedelta(days=expected_transit_days)
        now = datetime.now()
        if due.date() >= now.date():
            days_left = (due.date() - now.date()).days
            when = "today" if days_left == 0 else f"in about {days_left} day{'s' if days_left != 1 else ''}"
            return (f"In transit for {self.days_in_transit} day(s). Estimated arrival {when} "
                    f"({due.strftime('%d %b %Y')}, based on a typical {expected_transit_days}-day TCS window).")
        overdue = (now.date() - due.date()).days
        return (f"In transit for {self.days_in_transit} day(s), {overdue} day(s) past the typical "
                f"{expected_transit_days}-day window. Worth calling TCS (021-111-123-456).")
