import argparse
import csv
import json
import logging
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.parse import urlsplit
from urllib.robotparser import RobotFileParser

import requests

from .retailers import RETAILERS, parse_offer

USER_AGENT = "TabletopPriceWatch/0.1"
FIELDS = ("Game", "Type", "Retailer", "Price USD", "Availability", "Listing title", "Source ID",
          "Identity check", "Source URL", "Run started UTC", "Detail", "Offer key")
logger = logging.getLogger(__name__)


class SourceError(Exception):
    pass


class StoreClient:
    def __init__(self, session: requests.Session, host: str):
        self.session = session
        self.host = host
        self.next_request = 0.0
        self.delay = 3.0
        self.robots = RobotFileParser()
        self.paused = False

    def get(self, url: str) -> requests.Response:
        if self.paused:
            raise SourceError("source paused after access restriction or rate limit")
        time.sleep(max(0, self.next_request - time.monotonic()))
        self.next_request = time.monotonic() + self.delay
        response = self.session.get(url, timeout=20, allow_redirects=False)
        if response.status_code == 429:
            self.paused = True
            raise SourceError(f"HTTP 429; Retry-After {response.headers.get('Retry-After', 'unspecified')}")
        if response.status_code in {401, 403}:
            self.paused = True
        if response.status_code != 200:
            raise SourceError(f"HTTP {response.status_code}")
        return response

    def load_robots(self) -> None:
        response = self.get(f"https://{self.host}/robots.txt")
        self.robots.parse(response.text.splitlines())
        rate = self.robots.request_rate(USER_AGENT)
        self.delay = max(self.delay, self.robots.crawl_delay(USER_AGENT) or 0,
                         rate.seconds / rate.requests if rate and rate.requests else 0)
        self.next_request = time.monotonic() + self.delay

    def product(self, url: str) -> bytes:
        parts = urlsplit(url)
        if parts.scheme != "https" or parts.netloc != self.host:
            raise SourceError("URL is outside the configured retailer")
        if not self.robots.can_fetch(USER_AGENT, url):
            raise SourceError("robots.txt disallows this URL")
        return self.get(url).content


def normalized(code: str) -> str:
    return "".join(char for char in code.upper() if char.isalnum())


def verify_identity(game: dict, offer) -> str:
    actual = normalized(offer.identity)
    expected = game.get("upc") if offer.method == "UPC" else game.get("mpn")
    if expected and actual:
        target = normalized(expected)
        aliases = {target}
        if offer.method in {"SKU", "MPN"}:
            aliases.update(f"{prefix}{target}" for prefix in ("CSG", "CTG", "ASM"))
            if target.startswith("AEG"):
                aliases.add(target[3:])
            if target.startswith("STM"):
                aliases.add(target[3:])
            if target.startswith("NMG"):
                aliases.update({target[3:], f"NM{target[3:]}"})
        if actual not in aliases:
            raise ValueError(f"expected {offer.method} {expected}, got {offer.identity}")
        return offer.method
    raise ValueError(f"missing independently verified {offer.method}")


def collect(catalog: list[dict], observed: str) -> list[list]:
    rows = {}
    with requests.Session() as session:
        session.headers["User-Agent"] = USER_AGENT
        session.trust_env = False
        for key, (label, host) in RETAILERS.items():
            client = StoreClient(session, host)
            robots_error = ""
            try:
                client.load_robots()
            except (SourceError, requests.RequestException) as exc:
                robots_error = str(exc)
            for game in catalog:
                url = game["offers"].get(key, "")
                title = identity = method = detail = ""
                price = None
                status = "No listing"
                if url:
                    try:
                        if robots_error:
                            raise SourceError(f"robots.txt: {robots_error}")
                        offer = parse_offer(key, client.product(url))
                        title, identity = offer.title, offer.identity
                        method = verify_identity(game, offer)
                        price, status = offer.price, offer.status
                    except (KeyError, TypeError, ValueError) as exc:
                        status, detail = "Review required", str(exc)
                    except (SourceError, requests.RequestException) as exc:
                        status, detail = "Fetch failed", str(exc)
                rows[game["id"], key] = [game["name"], game["type"], label, price, status, title,
                                          identity, method, url, observed, detail, f"{game['name']}|{label}"]
                logger.info("%s · %s · %s%s", label, game["name"], status, f" · {detail}" if detail else "")
    return [rows[game["id"], key] for game in catalog for key in RETAILERS]


def save(path: Path, rows: list[list]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with NamedTemporaryFile(dir=path.parent, mode="w", newline="", encoding="utf-8-sig",
                                prefix=".snapshot-", delete=False) as stream:
            temporary = Path(stream.name)
            writer = csv.writer(stream, lineterminator="\n")
            writer.writerow(FIELDS)
            writer.writerows(
                [f"'{value}" if isinstance(value, str) and value.startswith(("=", "+", "-", "@", "\t", "\r"))
                 else value for value in row] for row in rows
            )
        temporary.replace(path)
    finally:
        if temporary:
            temporary.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Collect US board game offers for the Excel dashboard.")
    parser.add_argument("--catalog", type=Path, default=Path("catalog.json"))
    parser.add_argument("--data-dir", type=Path, default=Path("data"))
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    catalog = json.loads(args.catalog.read_text())
    if not catalog or any(len({game[field] for game in catalog}) != len(catalog) for field in ("id", "name")):
        parser.error("catalog must have at least one game and unique IDs and names")
    start = datetime.now(timezone.utc)
    rows = collect(catalog, start.isoformat(timespec="seconds"))
    history = args.data_dir / "history" / f"{start:%Y%m%dT%H%M%S%fZ}.csv"
    save(history, rows)
    usable = sum(row[4] in {"In stock", "Out of stock", "No new copy", "Preorder", "Backorder"} for row in rows)
    if usable:
        save(args.data_dir / "latest.csv", rows)
    logger.info("Saved %d checked rows (%d verified offers) to %s", len(rows), usable, history)
    return 0 if usable else 1


if __name__ == "__main__":
    raise SystemExit(main())
