import json
import re
from dataclasses import dataclass
from decimal import Decimal

from bs4 import BeautifulSoup, Tag

@dataclass(frozen=True)
class Retailer:
    label: str
    host: str
    market: str
    currency: str


RETAILERS = {
    "miniature_market": Retailer("Miniature Market", "www.miniaturemarket.com", "United States", "USD"),
    "cardhaus": Retailer("Cardhaus", "www.cardhaus.com", "United States", "USD"),
    "noble_knight": Retailer("Noble Knight", "www.nobleknight.com", "United States", "USD"),
    "philibert": Retailer("Philibert", "www.philibertnet.com", "France", "EUR"),
    "edinorog": Retailer("Edinorog", "edinorog.shop", "Russia", "RUB"),
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
    price = Decimal(value.replace("$", "").replace(",", "")).quantize(Decimal("0.01"))
    if price <= 0:
        raise ValueError("price must be positive")
    return price


def product_json(soup: BeautifulSoup) -> dict:
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            value = json.loads(script.get_text())
        except ValueError:
            continue
        if isinstance(value, dict):
            for item in value.get("@graph", [value]):
                if isinstance(item, dict) and item.get("@type") == "Product":
                    return item
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


def _philibert(soup: BeautifulSoup) -> Offer:
    product = product_json(soup)
    offer = product["offers"]
    if offer["priceCurrency"] != "EUR":
        raise ValueError("expected EUR")
    status = stock(offer.get("availability", ""))
    if status == "Unknown stock":
        availability = soup.select_one("#availability-after-price-container")
        if availability is not None:
            if availability.select_one("button.btn-add-to-cart[data-cart='1']:not([disabled])"):
                status = "In stock"
            elif "out of stock" in availability.get_text(" ", strip=True).lower():
                status = "Out of stock"
    return Offer(product["name"], money(str(offer["price"])), status, product.get("gtin", ""), "UPC")


def _edinorog(soup: BeautifulSoup) -> Offer:
    product = product_json(soup)
    offers = product["offers"]
    if len(offers) != 1 or offers[0].get("itemCondition") != "https://schema.org/NewCondition":
        raise ValueError("expected one new-copy offer")
    offer = offers[0]
    if offer["priceCurrency"] != "RUB":
        raise ValueError("expected RUB")
    title = product["name"]
    if "на английском языке" not in title.lower():
        raise ValueError("English edition is not confirmed on the listing")
    return Offer(title, money(str(offer["price"])), stock(offer["availability"]), offer.get("gtin", ""), "UPC")


PARSERS = {
    "miniature_market": _miniature_market, "cardhaus": _cardhaus, "noble_knight": _noble_knight,
    "philibert": _philibert, "edinorog": _edinorog,
}


def parse_offer(retailer: str, html: bytes) -> Offer:
    return PARSERS[retailer](BeautifulSoup(html, "html.parser"))
