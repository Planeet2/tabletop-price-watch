# Tabletop Price Watch

Tabletop Price Watch collects public offers for board games and expansions from three US retailers. It saves dated CSV snapshots and compares verified editions in an interactive Excel workbook.

[Download the Excel dashboard](outputs/Tabletop%20Price%20Watch.xlsx?raw=1) · [Browse the sample observations](examples/offers.csv) · [Inspect the matched catalog](catalog.json)

## What the snapshot shows

The November 13, 2026 UTC snapshot checks **10 editions across three stores**. It verifies 28 listings, of which 16 have a new copy in stock. Four editions have at least two in-stock offers, including one tie.

| Example | Snapshot result |
| --- | --- |
| Finspan | Noble Knight listed a new copy at **$41.81**. The other two stores listed **$45.00**, a **$3.19** gap before shipping. |
| Cascadia | Miniature Market and Cardhaus tied at **$31.99**. |
| Three-game basket | Miniature Market **$100.98** and Cardhaus **$101.98** with estimated economy shipping. Noble Knight had two of the three games. |

## Work in Excel

- **Dashboard** compares the selected basket, estimated delivery, in-stock coverage and lowest-price counts. Use the yellow game selector to inspect an edition's cheapest offer and price gap.
- **Comparison** shows all three store prices and statuses. Change the yellow **Basket qty** cells to select items. The price gap is the highest minus the lowest in-stock price. The percentage uses the lowest price as its denominator.
- **Offers** keeps the checked URLs, source IDs and match status for auditing. Missing listings and unavailable new copies stay visible. Only verified in-stock offers can win a comparison.

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

Each run writes `data/history/<timestamp>.csv` and updates `data/latest.csv` when it obtains verified observations. To refresh the shipped workbook with the same catalog, import `latest.csv` in Excel with **Source ID** set to Text. Copy the 30 data rows across all 12 columns and paste values into **Offers** starting at A2. The comparison matches rows by **Offer key** and Excel recalculates the formulas and chart.

For environments with the `@oai/artifact-tool` workspace package available, rebuild the workbook directly from the CSV. This also updates the workbook ranges if you change `catalog.json`.

```bash
node build_workbook.mjs data/latest.csv 'outputs/Tabletop Price Watch.xlsx'
```

## Matching and collection

The watchlist pairs retailer URLs with publisher part numbers and UPCs. The scraper checks each page’s identifier against that list. Miniature Market supplies product microdata, Cardhaus supplies a structured offer and UPC, and Noble Knight lists multiple conditions. The Noble Knight parser takes the cheapest available listing marked `MINT/New` and uses the displayed sale price when present. A mismatch appears as **Review required** and is excluded from the price comparison. The initial ten editions are a curated watchlist.

The collector checks `robots.txt` and applicable crawl delays before fetching a store. An access denial or HTTP 429 pauses that source for the run. Failed rows remain visible in the snapshot. The project only collects item identifiers, prices, stock and source links.

Identifiers for the Azul editions were checked against the US distributor’s catalog. [Asmodee USA catalog](https://www.asmodeena.com/AUSA-Active-05012026.pdf)

## Tests

```bash
python -m unittest discover -s tests -v
```
