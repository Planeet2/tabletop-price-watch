import logging
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, DecimalException, ROUND_HALF_UP
from xml.etree import ElementTree

import requests

ECB_URL = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"
CBR_URL = "https://www.cbr.ru/scripts/XML_daily.asp"
ECB_NS = "{http://www.ecb.int/vocabulary/2002-08-01/eurofxref}"
MOSCOW_TIME = timezone(timedelta(hours=3))
logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class Rate:
    usd_per_unit: Decimal
    effective_date: date | None = None
    source: str = ""

    def to_usd(self, amount: Decimal) -> Decimal:
        return (amount * self.usd_per_unit).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def ecb_rate(body: bytes, url: str) -> Rate:
    root = ElementTree.fromstring(body)
    day = root.find(f".//{ECB_NS}Cube[@time]")
    if day is None:
        raise ValueError("ECB publication date is missing")
    usd = day.find(f"{ECB_NS}Cube[@currency='USD']")
    if usd is None:
        raise ValueError("ECB USD rate is missing")
    return Rate(Decimal(usd.attrib["rate"]), date.fromisoformat(day.attrib["time"]), url)


def cbr_rate(body: bytes, url: str) -> Rate:
    # The feed declares Windows-1251. Transcode for XML parsers limited to UTF encodings.
    xml = body.decode("cp1251").replace('encoding="windows-1251"', 'encoding="utf-8"', 1)
    root = ElementTree.fromstring(xml)
    usd = root.find("./Valute[CharCode='USD']")
    if usd is None:
        raise ValueError("Bank of Russia USD rate is missing")
    rubles = Decimal(usd.findtext("Value", "").replace(",", "."))
    dollars = Decimal(usd.findtext("Nominal", ""))
    return Rate(dollars / rubles, datetime.strptime(root.attrib["Date"], "%d.%m.%Y").date(), url)


def collect_rates(session: requests.Session, observed: datetime) -> dict[str, Rate]:
    rates = {"USD": Rate(Decimal(1))}
    ecb_day = observed.astimezone(timezone.utc).date()
    moscow_day = observed.astimezone(MOSCOW_TIME).date()
    for currency, url, parse, day in (
        ("EUR", ECB_URL, ecb_rate, ecb_day),
        ("RUB", f"{CBR_URL}?date_req={moscow_day:%d/%m/%Y}", cbr_rate, moscow_day),
    ):
        try:
            response = session.get(url, timeout=20, allow_redirects=False)
            if response.status_code != 200:
                raise ValueError(f"HTTP {response.status_code}")
            rate = parse(response.content, url)
            if not rate.usd_per_unit.is_finite() or rate.usd_per_unit <= 0:
                raise ValueError("invalid USD rate")
            if not 0 <= (day - rate.effective_date).days <= 7:
                raise ValueError(f"reference rate date {rate.effective_date} is outside the seven-day window")
            rates[currency] = rate
        except (requests.RequestException, ElementTree.ParseError, DecimalException, KeyError, ValueError) as exc:
            logger.warning("%s reference rate unavailable · %s", currency, exc)
    return rates
