#!/usr/bin/env python3
"""Build the Arepas House menu profitability workbook.

Data sources:
- Menu items + prices: Square catalog (pulled 2026-07-11)
- Ingredient prices: supplier invoices in Google Drive "Claude Managed/Supplier Invoices"
  (US Foods, Mission & Vegetable, Latin Market Mana, Restaurant Depot), Jun 12 - Jul 9 2026
- Recipe portions: DRAFT estimates from Square item descriptions - to be corrected by owner

Usage: python3 scripts/build_workbook.py  ->  output/Menu_Profitability.xlsx
"""
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT = os.path.join(REPO, "output", "Menu_Profitability.xlsx")
INVOICES = os.path.join(REPO, "data", "invoices_extracted.json")

# ---------------------------------------------------------------------------
# Ingredient price database
# (name, vendor, last_invoice_date, invoice_no, pack_desc, pack_price, yield_qty, unit, status, note)
# status: INVOICE = straight from invoice | CHECK = invoice price but pack/yield assumed
#         ESTIMATE = no recent invoice, market estimate - replace with real price
#         PRICE NEEDED = item seen on invoice but price unreadable
# ---------------------------------------------------------------------------
INGREDIENTS = [
    # --- Latin Market Mana ---
    ("Tequenos (frozen)", "Latin Market Mana", "2026-07-07", "SVFPDYSHPK2NC", "160 units/case", 87.75, 160, "each", "INVOICE", ""),
    ("Cachapa pack (Palmita, w/ queso)", "Latin Market Mana", "2026-07-07", "SVFPDYSHPK2NC", "1 pack", 13.99, 1, "pack", "CHECK", "How many cachapas per pack? Update yield qty."),
    ("Queso de Mano", "Latin Market Mana", "2026-07-07", "SVFPDYSHPK2NC", "26 oz wheel", 11.00, 26, "oz", "INVOICE", ""),
    ("Nata (crema)", "Latin Market Mana", "2026-07-07", "SVFPDYSHPK2NC", "32 lb tub", 116.00, 512, "oz", "INVOICE", ""),
    ("Concentrado de Pina", "Latin Market Mana", "2026-07-07", "SVFPDYSHPK2NC", "1 gal", 15.39, 128, "fl oz", "INVOICE", ""),
    ("Concentrado de Parchita", "Latin Market Mana", "2026-07-04", "0RJ9HAABW6KF8", "1 gal", 15.39, 128, "fl oz", "INVOICE", ""),
    ("Frescolita (12 oz can)", "Latin Market Mana", "2026-06-23", "PSEKS3P317GJR", "24-can case", 23.00, 24, "each", "INVOICE", ""),
    ("Cocosette", "Latin Market Mana", "2026-06-12", "VA375MKQ3HGWG", "18 x 50 g", 14.73, 18, "each", "INVOICE", ""),
    ("Samba (Strawberry)", "Latin Market Mana", "2026-06-12", "VA375MKQ3HGWG", "20 x 32 g", 18.81, 20, "each", "INVOICE", ""),
    ("Cri-Cri", "Latin Market Mana", "2026-06-12", "VA375MKQ3HGWG", "12 x 27 g", 11.93, 12, "each", "INVOICE", ""),
    ("Savoy Chocolate de Leche", "Latin Market Mana", "2026-06-12", "VA375MKQ3HGWG", "12 x 30 g", 13.50, 12, "each", "INVOICE", ""),
    ("Galak White Chocolate", "Latin Market Mana", "2026-06-12", "VA375MKQ3HGWG", "12 x 30 g", 11.93, 12, "each", "INVOICE", ""),
    ("Pirucream (300 g can)", "Latin Market Mana", "2026-06-12", "VA375MKQ3HGWG", "300 g can", 6.97, 1, "each", "INVOICE", ""),
    # --- US Foods ---
    ("Beef rib lifter meat (raw)", "US Foods", "2026-07-09", "3190-0146", "catch weight, $/lb", 6.23, 16, "oz", "INVOICE", "162.2 lb delivered on last invoice"),
    ("Chicken breast jumbo (raw)", "US Foods", "2026-07-09", "3190-0146", "4/10 lb case", 107.34, 640, "oz", "CHECK", "Extended price partly unreadable on scan - verify $107.34/case"),
    ("Cheddar shredded", "US Foods", "2026-07-09", "3190-0146", "4/5 lb case", 139.18, 320, "oz", "CHECK", "Extended price partly unreadable on scan - verify $139.18/case"),
    ("Frying shortening (canola liquid)", "US Foods", "2026-07-09", "3190-0146", "35 lb case", 42.04, 560, "oz", "INVOICE", ""),
    # --- Mission & Vegetable ---
    ("Red bell pepper (morron)", "Mission & Vegetable", "2026-07-09", "41123", "case (~25 lb assumed)", 38.75, 400, "oz", "CHECK", "Case weight not on invoice - assumed 25 lb"),
    ("Yellow onion", "Mission & Vegetable", "2026-07-09", "41123", "sack (~50 lb assumed)", 25.75, 800, "oz", "CHECK", "Sack weight not on invoice - assumed 50 lb"),
    ("Apio (celery)", "Mission & Vegetable", "2026-07-09", "41123", "case (~28 lb assumed)", 27.75, 448, "oz", "CHECK", "Case weight not on invoice - assumed 28 lb"),
    ("Avocado (#48 case)", "Mission & Vegetable", "2026-07-09", "41123", "48-count case", 46.75, 48, "each", "INVOICE", ""),
    ("Habanero", "Mission & Vegetable", "2026-06-25", "40457", "case (~8 lb assumed)", 43.75, 128, "oz", "CHECK", "Case weight not on invoice - assumed 8 lb"),
    # --- Restaurant Depot ---
    ("Mayonnaise", "Restaurant Depot", "2026-07-06", "20544606249422176", "4 x 1 gal", 44.55, 512, "fl oz", "INVOICE", ""),
    ("Black beans (cooked)", "Restaurant Depot", "2026-07-06", "20544606249422176", "25 lb dry bag", 13.33, 880, "oz", "CHECK", "25 lb dry x ~2.2 cooked yield = 55 lb cooked (880 oz)"),
    ("Seasoning salt", "Restaurant Depot", "2026-07-06", "20544606249422176", "10 lb jar (substituted)", 67.44, 160, "oz", "CHECK", "20 lb was substituted w/ 10 lb jar - verify charge"),
    ("Canola salad oil", "Restaurant Depot", "2026-07-06", "20544606249422176", "35 lb", 41.82, 560, "oz", "INVOICE", ""),
    ("Peeled garlic", "Restaurant Depot", "2026-07-06", "20544606249422176", "jar (~5 lb assumed)", 15.95, 80, "oz", "CHECK", "Jar size not on receipt - assumed 5 lb"),
    # --- On invoice but price unreadable ---
    ("Yuca fries frozen", "US Foods (?)", "2026-06-29", "?", "20 x 1 lb", None, 320, "oz", "PRICE NEEDED", "07-03 invoice scan unreadable - ENTER PACK PRICE"),
    ("Sweet plantain IQF", "US Foods (?)", "2026-06-29", "?", "4 x 6 lb", None, 384, "oz", "PRICE NEEDED", "07-03 invoice scan unreadable - ENTER PACK PRICE"),
    # --- Not on recent invoices: market estimates, replace with your real prices ---
    ("Harina P.A.N. (cornmeal)", "-", "-", "-", "50 lb (ESTIMATE)", 45.00, 800, "oz", "ESTIMATE", "No invoice in last 30 days - replace with real price"),
    ("White cheese (mozzarella/llanero)", "-", "-", "-", "per lb (ESTIMATE)", 3.25, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Cotija cheese", "-", "-", "-", "per lb (ESTIMATE)", 4.50, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Green plantain", "-", "-", "-", "each (ESTIMATE)", 0.60, 1, "each", "ESTIMATE", "Replace with real price"),
    ("Lettuce", "-", "-", "-", "head (ESTIMATE)", 2.50, 1, "each", "ESTIMATE", "Replace with real price"),
    ("Tomato", "-", "-", "-", "per lb (ESTIMATE)", 1.80, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Cabbage", "-", "-", "-", "per lb (ESTIMATE)", 0.90, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Carrot", "-", "-", "-", "per lb (ESTIMATE)", 0.80, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Mustard", "-", "-", "-", "1 gal (ESTIMATE)", 12.00, 128, "fl oz", "ESTIMATE", "Replace with real price"),
    ("Lime", "-", "-", "-", "each (ESTIMATE)", 0.25, 1, "each", "ESTIMATE", "Replace with real price"),
    ("Papelon (panela)", "-", "-", "-", "per lb (ESTIMATE)", 2.50, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Sugar", "-", "-", "-", "per lb (ESTIMATE)", 0.60, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Butter", "-", "-", "-", "per lb (ESTIMATE)", 3.50, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Malta Polar (12 oz can)", "-", "-", "-", "each (ESTIMATE)", 1.20, 1, "each", "ESTIMATE", "Replace with real price"),
    ("Vegan mayo", "-", "-", "-", "1 gal (ESTIMATE)", 18.00, 128, "fl oz", "ESTIMATE", "Replace with real price"),
    ("Carre Savoy (each)", "-", "-", "-", "each (ESTIMATE)", 2.20, 1, "each", "ESTIMATE", "Not on recent Mana invoices"),
    ("Ham", "-", "-", "-", "per lb (ESTIMATE)", 3.00, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Potato", "-", "-", "-", "per lb (ESTIMATE)", 0.70, 16, "oz", "ESTIMATE", "Replace with real price"),
    ("Guasacaca sauce (house-made)", "-", "-", "-", "per fl oz (ESTIMATE)", 0.20, 1, "fl oz", "ESTIMATE", "Rough cost/fl oz - can build as sub-recipe later"),
    ("To-go container/wrap", "-", "-", "-", "per order (ESTIMATE)", 0.45, 1, "each", "ESTIMATE", "Replace with real price"),
    ("Drink cup + lid + straw (20 oz)", "-", "-", "-", "each (ESTIMATE)", 0.28, 1, "each", "ESTIMATE", "Replace with real price"),
]

# ---------------------------------------------------------------------------
# Menu (from Square catalog) and DRAFT recipes.
# (menu_item, category, price, [(ingredient, qty, note), ...])
# Qty units always match the ingredient's unit in the Ingredient Prices sheet.
# ALL PORTIONS ARE DRAFT ESTIMATES from item descriptions - owner to correct.
# ---------------------------------------------------------------------------
MENU = [
    ("Yucca Fries (7 units)", "Appetizers", 6.99, [
        ("Yuca fries frozen", 8, "DRAFT: ~8 oz per order"),
        ("Frying shortening (canola liquid)", 1.5, "DRAFT: oil absorbed in frying"),
        ("Guasacaca sauce (house-made)", 1.5, "DRAFT: dipping sauce"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Plantains and Cheese (7 units)", "Appetizers", 6.99, [
        ("Sweet plantain IQF", 8, "DRAFT: ~8 oz per order"),
        ("Queso de Mano", 1.5, "DRAFT: white cheese topping"),
        ("Frying shortening (canola liquid)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Tequeños (5 units)", "Appetizers", 8.99, [
        ("Tequenos (frozen)", 5, "5 units per order"),
        ("Frying shortening (canola liquid)", 1, "DRAFT: air fried - may be less"),
        ("Guasacaca sauce (house-made)", 1.5, "DRAFT: dipping sauce"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Empanadas", "Appetizers", 4.95, [
        ("Harina P.A.N. (cornmeal)", 2.5, "DRAFT: dough per empanada"),
        ("Beef rib lifter meat (raw)", 1.5, "DRAFT: avg filling (beef/chicken/cheese)"),
        ("Frying shortening (canola liquid)", 1.5, "DRAFT"),
        ("To-go container/wrap", 0.5, "DRAFT: shared packaging"),
    ]),
    ("Pastelitos", "Appetizers", 4.95, [
        ("Harina P.A.N. (cornmeal)", 2.5, "DRAFT: dough"),
        ("Ham", 1, "DRAFT: ham & cheese filling"),
        ("White cheese (mozzarella/llanero)", 1, "DRAFT"),
        ("Frying shortening (canola liquid)", 1.5, "DRAFT"),
        ("To-go container/wrap", 0.5, "DRAFT: shared packaging"),
    ]),
    ("Arepa Cheese", "Arepas", 6.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Queso de Mano", 2.5, "DRAFT: thick melted cheese"),
        ("Butter", 0.3, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Arepa Domino", "Arepas", 6.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Black beans (cooked)", 3, "DRAFT"),
        ("Cotija cheese", 1, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Arepa Chicken", "Arepas", 12.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Chicken breast jumbo (raw)", 4.5, "DRAFT: raw wt for ~3.5 oz cooked shredded"),
        ("Sweet plantain IQF", 2, "DRAFT"),
        ("Cheddar shredded", 1, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Arepa Pabellon", "Arepas", 13.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Beef rib lifter meat (raw)", 3.5, "DRAFT: raw wt for shredded beef"),
        ("Black beans (cooked)", 2.5, "DRAFT"),
        ("Sweet plantain IQF", 2, "DRAFT"),
        ("Avocado (#48 case)", 0.25, "DRAFT: 1/4 avocado"),
        ("Queso de Mano", 1, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Arepas Vegetarian", "Arepas", 10.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Black beans (cooked)", 2.5, "DRAFT"),
        ("Avocado (#48 case)", 0.33, "DRAFT: 1/3 avocado"),
        ("Sweet plantain IQF", 2, "DRAFT"),
        ("Tomato", 1, "DRAFT"),
        ("Lettuce", 0.1, "DRAFT: 1/10 head"),
        ("Queso de Mano", 1, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Arepa Vegan", "Arepas", 10.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Black beans (cooked)", 2.5, "DRAFT"),
        ("Avocado (#48 case)", 0.33, "DRAFT"),
        ("Sweet plantain IQF", 2, "DRAFT"),
        ("Tomato", 1, "DRAFT"),
        ("Lettuce", 0.1, "DRAFT"),
        ("Vegan mayo", 1, "DRAFT: vegan guasacaca base"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Arepa Queen", "Arepas", 11.99, [
        ("Harina P.A.N. (cornmeal)", 3, "DRAFT: arepa shell"),
        ("Chicken breast jumbo (raw)", 4, "DRAFT: chicken salad"),
        ("Mayonnaise", 1, "DRAFT"),
        ("Mustard", 0.3, "DRAFT"),
        ("Avocado (#48 case)", 0.4, "DRAFT: in salad + extra"),
        ("Sweet plantain IQF", 2, "DRAFT"),
        ("Queso de Mano", 1, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Cachapas", "Mains", 16.99, [
        ("Cachapa pack (Palmita, w/ queso)", 0.2, "DRAFT: assumes 5 cachapas/pack - VERIFY"),
        ("Queso de Mano", 4, "DRAFT: melted paisa-style cheese"),
        ("Butter", 0.5, "DRAFT"),
        ("Nata (crema)", 1, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Patacon (Original)", "Mains", 19.95, [
        ("Green plantain", 2, "DRAFT: 2 plantains for the 2 slices"),
        ("Chicken breast jumbo (raw)", 3, "DRAFT"),
        ("Beef rib lifter meat (raw)", 3, "DRAFT"),
        ("Cabbage", 2, "DRAFT: coleslaw"),
        ("Carrot", 0.5, "DRAFT: coleslaw"),
        ("Mayonnaise", 1, "DRAFT: slaw + sauces"),
        ("Queso de Mano", 1.5, "DRAFT"),
        ("Avocado (#48 case)", 0.33, "DRAFT"),
        ("Tomato", 1, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1.5, "DRAFT: 3 sauces"),
        ("Frying shortening (canola liquid)", 2, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Patacon (Vegetarian)", "Mains", 15.99, [
        ("Green plantain", 2, "DRAFT"),
        ("Cabbage", 2, "DRAFT: coleslaw"),
        ("Carrot", 0.5, "DRAFT: coleslaw"),
        ("Mayonnaise", 1, "DRAFT"),
        ("Queso de Mano", 1.5, "DRAFT"),
        ("Avocado (#48 case)", 0.33, "DRAFT"),
        ("Tomato", 1.5, "DRAFT"),
        ("Lettuce", 0.15, "DRAFT"),
        ("Guasacaca sauce (house-made)", 1.5, "DRAFT"),
        ("Frying shortening (canola liquid)", 2, "DRAFT"),
        ("To-go container/wrap", 1, ""),
    ]),
    ("Passion Fruit (20 oz)", "Drinks", 4.75, [
        ("Concentrado de Parchita", 4, "DRAFT: fl oz concentrate per 20 oz cup"),
        ("Sugar", 1, "DRAFT"),
        ("Drink cup + lid + straw (20 oz)", 1, ""),
    ]),
    ("Pineapple Juice (20 oz)", "Drinks", 4.75, [
        ("Concentrado de Pina", 4, "DRAFT: fl oz concentrate per 20 oz cup"),
        ("Sugar", 1, "DRAFT"),
        ("Drink cup + lid + straw (20 oz)", 1, ""),
    ]),
    ("Papelon con Limon (20 oz)", "Drinks", 4.75, [
        ("Papelon (panela)", 2, "DRAFT"),
        ("Lime", 1, "DRAFT"),
        ("Drink cup + lid + straw (20 oz)", 1, ""),
    ]),
    ("Malta Polar (12 oz)", "Drinks", 4.25, [
        ("Malta Polar (12 oz can)", 1, "ESTIMATE price - not on recent invoices"),
    ]),
    ("Frescolita (12 oz)", "Drinks", 4.00, [
        ("Frescolita (12 oz can)", 1, ""),
    ]),
    ("Cri-Cri", "Candy", 1.75, [("Cri-Cri", 1, "")]),
    ("Pirucream", "Candy", 15.00, [("Pirucream (300 g can)", 1, "")]),
    ("Cocosette", "Candy", 2.00, [("Cocosette", 1, "")]),
    ("Samba (Strawberry)", "Candy", 2.00, [("Samba (Strawberry)", 1, "")]),
    ("Carré (Savoy Hazelnut Milk Chocolate)", "Candy", 3.75, [("Carre Savoy (each)", 1, "ESTIMATE price")]),
    ("Savoy - Chocolate de Leche", "Candy", 1.75, [("Savoy Chocolate de Leche", 1, "")]),
    ("Galak - White Chocolate", "Candy", 1.75, [("Galak White Chocolate", 1, "")]),
]

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
CHECK_FILL = PatternFill("solid", fgColor="FFF2CC")     # yellow - verify
EST_FILL = PatternFill("solid", fgColor="FCE4D6")       # orange - estimate
NEED_FILL = PatternFill("solid", fgColor="F8CBAD")      # red-orange - price needed
DRAFT_FONT = Font(italic=True, color="808080", size=9)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '"$"#,##0.00'
MONEY4 = '"$"#,##0.0000'
PCT = "0.0%"


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build():
    wb = Workbook()

    # ================= Menu Profitability =================
    ws = wb.active
    ws.title = "Menu Profitability"
    ws["A1"] = "Arepas House - Menu Profitability"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Prices from Square catalog (Jul 2026). Costs roll up from the Recipes sheet - edit portions there and this sheet recalculates."
    ws["A2"].font = DRAFT_FONT
    headers = ["Menu Item", "Category", "Menu Price", "Ingredient Cost", "Gross Profit $", "Food Cost %", "Margin %", "Notes"]
    ws.append([])
    ws.append(headers)
    hdr_row = 4
    style_header(ws, hdr_row, len(headers))
    item_notes = {
        "Yucca Fries (7 units)": "INCOMPLETE - yuca price missing (unreadable invoice)",
        "Plantains and Cheese (7 units)": "INCOMPLETE - plantain price missing (unreadable invoice)",
        "Cachapas": "Verify cachapas per pack on Ingredient Prices sheet",
        "Malta Polar (12 oz)": "Cost is an estimate - no recent invoice",
        "Carré (Savoy Hazelnut Milk Chocolate)": "Cost is an estimate - no recent invoice",
    }
    r = hdr_row
    for name, cat, price, lines in MENU:
        r += 1
        uses_missing = any(i in ("Yuca fries frozen", "Sweet plantain IQF") for i, _, _ in lines)
        ws.cell(row=r, column=1, value=name).border = BORDER
        ws.cell(row=r, column=2, value=cat).border = BORDER
        c = ws.cell(row=r, column=3, value=price); c.number_format = MONEY; c.border = BORDER
        c = ws.cell(row=r, column=4, value=f'=SUMIF(Recipes!$A:$A,$A{r},Recipes!$F:$F)'); c.number_format = MONEY; c.border = BORDER
        c = ws.cell(row=r, column=5, value=f'=C{r}-D{r}'); c.number_format = MONEY; c.border = BORDER
        c = ws.cell(row=r, column=6, value=f'=IF(C{r}=0,"",D{r}/C{r})'); c.number_format = PCT; c.border = BORDER
        c = ws.cell(row=r, column=7, value=f'=IF(C{r}=0,"",E{r}/C{r})'); c.number_format = PCT; c.border = BORDER
        note = item_notes.get(name, "")
        nc = ws.cell(row=r, column=8, value=note)
        nc.border = BORDER
        nc.font = DRAFT_FONT
        if uses_missing:
            for col in range(1, 9):
                ws.cell(row=r, column=col).fill = NEED_FILL
    last = r
    # Food-cost % traffic lights (col F)
    rng = f"F{hdr_row+1}:F{last}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0.3"], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0.3", "0.4"], fill=PatternFill("solid", fgColor="FFEB9C")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.4"], fill=PatternFill("solid", fgColor="FFC7CE")))
    set_widths(ws, [34, 12, 12, 14, 13, 12, 10, 46])
    ws.freeze_panes = f"A{hdr_row+1}"

    # ================= Recipes =================
    ws = wb.create_sheet("Recipes")
    ws["A1"] = "Recipes - DRAFT portions, please correct"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Qty unit must match the ingredient's unit on the Ingredient Prices sheet. "
                "Edit Qty (col C) to your real serving sizes; costs update automatically. Add rows freely - keep the menu item name identical to the Menu Profitability sheet.")
    ws["A2"].font = DRAFT_FONT
    headers = ["Menu Item", "Ingredient", "Qty", "Unit", "Unit Cost", "Line Cost", "Notes"]
    ws.append([])
    ws.append(headers)
    hdr_row = 4
    style_header(ws, hdr_row, len(headers))
    unit_by_ing = {i[0]: i[7] for i in INGREDIENTS}
    r = hdr_row
    for name, _cat, _price, lines in MENU:
        for ing, qty, note in lines:
            r += 1
            ws.cell(row=r, column=1, value=name).border = BORDER
            ws.cell(row=r, column=2, value=ing).border = BORDER
            c = ws.cell(row=r, column=3, value=qty); c.border = BORDER
            ws.cell(row=r, column=4, value=unit_by_ing[ing]).border = BORDER
            c = ws.cell(row=r, column=5, value=f"=IFERROR(VLOOKUP($B{r},'Ingredient Prices'!$A:$I,9,FALSE),0)")
            c.number_format = MONEY4; c.border = BORDER
            c = ws.cell(row=r, column=6, value=f'=IF(ISNUMBER(E{r}),C{r}*E{r},0)'); c.number_format = MONEY; c.border = BORDER
            nc = ws.cell(row=r, column=7, value=note); nc.border = BORDER; nc.font = DRAFT_FONT
    last = r
    dv = DataValidation(type="list", formula1=f"='Ingredient Prices'!$A$5:$A${4+len(INGREDIENTS)}", allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(f"B{hdr_row+1}:B{last+40}")
    set_widths(ws, [34, 34, 8, 8, 11, 11, 48])
    ws.freeze_panes = f"A{hdr_row+1}"

    # ================= Ingredient Prices =================
    ws = wb.create_sheet("Ingredient Prices")
    ws["A1"] = "Ingredient Prices - from supplier invoices (Jun 12 - Jul 9, 2026)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Unit Cost = Pack Price / Yield Qty. Yellow = invoice price but pack size assumed (verify). "
                "Orange = no recent invoice, market ESTIMATE (replace). Dark orange = price unreadable on scan (enter it).")
    ws["A2"].font = DRAFT_FONT
    headers = ["Ingredient", "Vendor", "Last Invoice", "Invoice #", "Pack", "Pack Price", "Yield Qty", "Unit", "Unit Cost", "Status", "Notes"]
    ws.append([])
    ws.append(headers)
    hdr_row = 4
    style_header(ws, hdr_row, len(headers))
    r = hdr_row
    for name, vendor, date, inv, pack, price, yqty, unit, status, note in INGREDIENTS:
        r += 1
        vals = [name, vendor, date, inv, pack, price, yqty, unit]
        for ci, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=ci, value=v)
            cell.border = BORDER
            if ci == 6:
                cell.number_format = MONEY
        c = ws.cell(row=r, column=9, value=f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(G{r}),G{r}>0),F{r}/G{r},"")')
        c.number_format = MONEY4; c.border = BORDER
        ws.cell(row=r, column=10, value=status).border = BORDER
        nc = ws.cell(row=r, column=11, value=note); nc.border = BORDER; nc.font = DRAFT_FONT
        fill = {"CHECK": CHECK_FILL, "ESTIMATE": EST_FILL, "PRICE NEEDED": NEED_FILL}.get(status)
        if fill:
            for col in range(1, 12):
                ws.cell(row=r, column=col).fill = fill
    set_widths(ws, [32, 20, 12, 20, 22, 11, 10, 7, 11, 14, 46])
    ws.freeze_panes = f"A{hdr_row+1}"

    # ================= Invoice Log =================
    ws = wb.create_sheet("Invoice Log")
    ws["A1"] = "Raw invoice line items (reference)"
    ws["A1"].font = TITLE_FONT
    headers = ["Vendor", "Invoice Date", "Invoice #", "Description", "Qty", "Pack/Size", "Unit Price", "Line Total"]
    ws.append([])
    ws.append(headers)
    style_header(ws, 3, len(headers))
    r = 3
    for inv in json.load(open(INVOICES)):
        for li in inv.get("line_items", []):
            r += 1
            row = [inv.get("vendor"), inv.get("date"), str(inv.get("invoice_number") or ""),
                   li.get("description"), li.get("qty") if not isinstance(li.get("qty"), str) else li.get("qty"),
                   li.get("pack_size"), li.get("unit_price") if isinstance(li.get("unit_price"), (int, float)) else li.get("unit_price"),
                   li.get("total") if isinstance(li.get("total"), (int, float)) else li.get("total")]
            for ci, v in enumerate(row, start=1):
                cell = ws.cell(row=r, column=ci, value=v)
                cell.border = BORDER
                if ci in (7, 8) and isinstance(v, (int, float)):
                    cell.number_format = MONEY
    set_widths(ws, [28, 12, 22, 52, 8, 20, 11, 11])
    ws.freeze_panes = "A4"

    # ================= README =================
    ws = wb.create_sheet("README")
    lines = [
        ("Arepas House - Menu Profitability Workbook", TITLE_FONT),
        ("Generated 2026-07-11 by Claude from Square catalog + supplier invoices in Google Drive 'Claude Managed/Supplier Invoices'.", None),
        ("", None),
        ("HOW IT WORKS", Font(bold=True)),
        ("1. 'Ingredient Prices' holds what you pay vendors, converted to a cost per oz / fl oz / each.", None),
        ("2. 'Recipes' lists what goes into each menu item. Line Cost = Qty x Unit Cost (looked up automatically).", None),
        ("3. 'Menu Profitability' sums each item's recipe lines and compares to the Square menu price.", None),
        ("   Food Cost % is color coded: green < 30%, yellow 30-40%, red > 40%.", None),
        ("", None),
        ("WHAT YOU NEED TO REVIEW (in priority order)", Font(bold=True)),
        ("A. Recipe portions (Recipes sheet, col C) are DRAFT estimates from the menu descriptions.", None),
        ("   Correct them to your real serving sizes - everything recalculates.", None),
        ("B. Dark-orange rows on 'Ingredient Prices': yuca fries and sweet plantain prices were unreadable", None),
        ("   on the 07-03 invoice scan. Enter the pack price - until then Yucca Fries and Plantains & Cheese costs are understated.", None),
        ("C. Orange ESTIMATE rows: items not on any invoice in the last 30 days (Harina PAN, produce for slaw,", None),
        ("   packaging, Malta, etc.). Replace with your real prices.", None),
        ("D. Yellow CHECK rows: invoice price is real but the pack size/weight was not printed, so the yield is assumed.", None),
        ("E. Cachapas: the Palmita pack price is $13.99 but I don't know how many cachapas come per pack (yield qty = 1 for now).", None),
        ("", None),
        ("NOTES", Font(bold=True)),
        ("- Meat quantities are raw weight; cooked shrinkage is roughly built into the draft portions.", None),
        ("- Black beans unit cost is per cooked oz (25 lb dry bag x ~2.2 hydration yield).", None),
        ("- Towel service (Prudential/Clean Green), fuel surcharges and delivery fees are overhead, not COGS - they appear only in the Invoice Log.", None),
        ("- US Foods 07-09 scan was partly unreadable: chicken $107.34/case and cheddar $139.18/case are best readings - verify.", None),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font
    ws.column_dimensions["A"].width = 120

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    build()
