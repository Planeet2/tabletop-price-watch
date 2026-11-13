import json
import re
from dataclasses import dataclass
from decimal import Decimal

from bs4 import BeautifulSoup, Tag

RETAILERS = {
    "miniature_market": ("Miniature Market", "www.miniaturemarket.com"),
    "cardhaus": ("Cardhaus", "www.cardhaus.com"),
    "noble_knight": ("Noble Knight", "www.nobleknight.com"),
}
STOCK = {
    "InStock": "In stock", "LimitedAvailability": "In stock", "OutOfStock": "Out of stock",
    "SoldOut": "Out of stock", "Discontinued": "Out of stock", "PreOrder": "Preorder",
    "PreSale": "Preorder", "BackOrder": "Backorder", "OnlineOnly": "Unknown stock",
}


@dataclass(frozen=True)
class Offer:
    title: str
    price: Decimal | None
    status: str
    identity: str
    method: str


def money(value: str) -> Decimal:
    value = value.strip()
    if not re.fullmatch(r"\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?", value):
        raise ValueError(f"invalid price {value!r}")
    return Decimal(value.replace("$", "").replace(",", "")).quantize(Decimal("0.01"))


def product_json(soup: BeautifulSoup) -> dict:
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            value = json.loads(script.get_text())
        except ValueError:
            continue
        if isinstance(value, dict) and value.get("@type") == "Product":
            return value
    raise ValueError("Product structured data is missing")


def content(node: Tag, prop: str) -> str:
    element = node.select_one(f'[itemprop="{prop}"]')
    if element is None:
        raise ValueError(f"missing {prop}")
    return str(element.get("content") or element.get_text(" ", strip=True)).strip()


def stock(value: str) -> str:
    return STOCK.get(value.rsplit("/", 1)[-1], "Unknown stock")


def _miniature_market(soup: BeautifulSoup) -> Offer:
    product = soup.select_one('[itemtype$="/Product"]')
    if product is None:
        raise ValueError("Product markup is missing")
    if content(product, "priceCurrency") != "USD":
        raise ValueError("expected USD")
    title = product.select_one("h1.product-detail-name")
    if title is None:
        raise ValueError("product title is missing")
    return Offer(title.get_text(" ", strip=True), money(content(product, "price")),
                 stock(content(product, "availability")), content(product, "sku"), "SKU")


def _cardhaus(soup: BeautifulSoup) -> Offer:
    product = product_json(soup)
    offer = product["offers"]
    if offer["priceCurrency"] != "USD":
        raise ValueError("expected USD")
    upc = next((term.find_next_sibling().get_text(strip=True) for term in soup.select("dt")
                if term.get_text(strip=True) == "UPC:" and term.find_next_sibling()), "")
    return Offer(product["name"], money(str(offer["price"])), stock(offer["availability"]), upc, "UPC")


def _noble_knight(soup: BeautifulSoup) -> Offer:
    heading = soup.select_one("h1")
    if heading is None:
        raise ValueError("product title is missing")
    title = heading.get_text(" ", strip=True)
    part = next((line.select_one(".value").get_text(" ", strip=True) for line in soup.select(".info-line")
                 if line.select_one(".label") and line.select_one(".value")
                 and line.select_one(".label").get_text(strip=True) == "MFG. Part #"), "")
    prices = []
    for option in soup.select(".condition-col.in_stock"):
        condition = option.select_one(".condition-value a")
        quantity = option.select_one("input.condition-quantity")
        if condition is None or quantity is None:
            continue
        if condition.get_text(" ", strip=True) not in {"SW (MINT/New)", "MINT/New"}:
            continue
        label = quantity.get("aria-label", "")
        match = re.search(r" Price (\$[\d,.]+)$", label)
        if not match:
            raise ValueError(f"USD price is missing for new copy: {label!r}")
        prices.append(money(match[1]))
    if prices:
        product = product_json(soup)
        if product["offers"]["priceCurrency"] != "USD":
            raise ValueError("expected USD")
    return Offer(title, min(prices) if prices else None, "In stock" if prices else "No new copy", part, "MPN")


PARSERS = {"miniature_market": _miniature_market, "cardhaus": _cardhaus, "noble_knight": _noble_knight}


def parse_offer(retailer: str, html: bytes) -> Offer:
    return PARSERS[retailer](BeautifulSoup(html, "html.parser"))
