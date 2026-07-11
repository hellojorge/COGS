# COGS — Menu Profitability

Cost-of-goods and profitability tracking for Arepas House.

## Contents

- `output/Menu_Profitability.xlsx` — the workbook. Sheets:
  - **Menu Profitability** — every menu item (from the Square catalog) with menu price, ingredient cost, gross profit, and food-cost % (color coded: green < 30%, yellow 30–40%, red > 40%).
  - **Recipes** — the ingredient lines behind each item. Portion quantities are draft estimates; edit them and everything recalculates.
  - **Ingredient Prices** — vendor prices from supplier invoices (Jun 12 – Jul 9, 2026), converted to unit costs. Rows are flagged INVOICE / CHECK / ESTIMATE / PRICE NEEDED.
  - **Invoice Log** — raw line items extracted from every invoice, for reference.
  - **README** — how to use it and what still needs review.
- `data/invoices_extracted.json` — structured line items extracted from the invoice PDFs in Google Drive (`Claude Managed/Supplier Invoices`).
- `scripts/build_workbook.py` — regenerates the workbook (`python3 scripts/build_workbook.py`, requires `openpyxl`).

## Data sources

- Menu items and prices: Square catalog (27 items, pulled 2026-07-11).
- Ingredient prices: US Foods, Mission & Vegetable, Latin Market Mana, Restaurant Depot invoices.
- Prudential/Clean Green towel service, fuel surcharges and delivery fees are treated as overhead, not COGS.
