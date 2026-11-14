# Tabletop Price Watch

Tabletop Price Watch compares prices for matching English-language board games and expansions in the United States, France and Russia. A Python scraper saves dated CSV snapshots with prices in local currency and USD. The Excel workbook compares the lowest in-stock offer from each market.

[Download the Excel dashboard](outputs/Tabletop%20Price%20Watch.xlsx?raw=1) · [Browse the sample observations](examples/offers.csv) · [Inspect the matched catalog](catalog.json)

## What the snapshot shows

The November 14, 2026 UTC snapshot covers **10 editions across five retailers**. It verified 36 listings. Five editions were available in at least two markets, with two available in all three.

| Edition | United States | France | Russia | Spread |
| --- | ---: | ---: | ---: | ---: |
| Finspan | **$41.81** | €44.95 → $49.68 | ₽5,290 → $64.39 | $22.58 |
| Ark Nova | **$61.83** | €74.90 → $82.78 | ₽8,020 → $97.62 | $35.79 |
| Wyrmspan | **$53.64** | €64.95 → $71.78 | Out of stock | $18.14 |

The spread is the difference between the highest and lowest market minima. The percentage expresses that difference relative to the lowest price. The international view uses each shop’s listed price and excludes adjustments for taxes, delivery and import duties.

## Work in Excel

- **Dashboard** compares one edition. The yellow selector updates the chart, local and converted prices, availability and the amount above the cheapest available offer.
- **Markets** compares every edition at once. Each market column shows the lowest verified in-stock price in USD. Blank cells indicate that the snapshot contains no qualifying price for that edition and market. A single qualifying market has a price and no spread.
- **US Basket** estimates local economy delivery. Change the yellow **Basket qty** cells on **Comparison** to compare the three US retailers for your basket.
- **Offers** records each checked URL, identifier, status, local price, currency, conversion rate and official rate date. Only verified in-stock offers with a USD price qualify for the international ranking.

The workbook estimates standard shipping within the contiguous US before sales tax. It uses published thresholds and Noble Knight’s stated below-threshold maximum. Enter a Cardhaus quote in its yellow **Base estimate** cell for below-threshold comparisons. [Miniature Market](https://www.miniaturemarket.com/utility/customer-service) · [Cardhaus](https://www.cardhaus.com/shipping-faq/) · [Noble Knight](https://www.nobleknight.com/ordering-and-shipping)

## Refresh the data

Use Python 3.11 or newer from the repository root.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install .
tabletop-price-watch
```

For PowerShell, use `py -m venv .venv` and `.\.venv\Scripts\Activate.ps1` for the first two steps.

Each run writes `data/history/<timestamp>.csv` and updates `data/latest.csv` when it obtains verified observations. To refresh the shipped workbook with the same catalog, import `latest.csv` in Excel with **Source ID** and **FX date** set to Text. Import the price and rate columns as numbers using the decimal point from the CSV. Copy the 50 data rows across all 18 columns and paste values into **Offers** starting at A2. Excel recalculates the formulas and chart using **Offer key** to find each listing.

For environments with the `@oai/artifact-tool` workspace package available, rebuild the workbook directly from the CSV. This also updates the workbook ranges if you change `catalog.json`.

```bash
node build_workbook.mjs data/latest.csv 'outputs/Tabletop Price Watch.xlsx'
```

## Matching and collection

The watchlist pairs retailer URLs with publisher part numbers and barcodes. Miniature Market supplies product microdata, Cardhaus supplies a structured offer and UPC, and Noble Knight lists multiple conditions. The Noble Knight parser takes the cheapest available listing marked `MINT/New`. Philibert and Edinorog are checked against the exact barcode. The Russian parser also checks that the listing describes the English edition. A mismatch is recorded as **Review required**. Coverage is limited to the curated URL list, which can be edited in `catalog.json`. [Philibert](https://www.philibertnet.com/en/stonemaier-games/154440-finspan-850032180931.html) · [Edinorog](https://edinorog.shop/collection/wyrmspan)

The EUR rate comes from the European Central Bank. For RUB, the scraper divides the Bank of Russia’s quoted USD nominal by its price in rubles. USD prices are rounded to cents. Each rate can be up to seven calendar days old to account for publication schedules. A missing or unusable rate leaves the local price visible and the USD conversion blank. [ECB reference rates](https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml) · [Bank of Russia feed](https://cbr.ru/development/sxml/)

The collector checks `robots.txt` and applicable crawl delays before fetching a store. An access denial or HTTP 429 pauses that source for the run. Failed rows remain visible in the snapshot. The project only collects item identifiers, prices, stock and source links.

Identifiers for the Azul editions were checked against the US distributor’s catalog. [Asmodee USA catalog](https://www.asmodeena.com/AUSA-Active-05012026.pdf)

## Tests

```bash
python -m unittest discover -s tests -v
```
