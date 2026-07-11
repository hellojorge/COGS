#!/usr/bin/env python3
"""Emit the menu profitability model as a single CSV for Google Sheets conversion.

Side-by-side layout (no fragile cross-section row math):
  cols A-G : menu summary        cols I-O : recipes        cols Q-Z : ingredient prices
Computed columns (D,E,F,M,N,Y) use one ARRAYFORMULA in row 4 so the CSV stays small.
"""
import csv
import io
import sys

sys.path.insert(0, "scripts")
from build_workbook import INGREDIENTS, MENU

unit_by_ing = {i[0]: i[7] for i in INGREDIENTS}

menu_rows = [[name, cat, price, "", "", "", note] for name, cat, price, note in
             [(n, c, p, "") for n, c, p, _l in MENU]]
recipe_rows = []
for name, _cat, _price, lines in MENU:
    for ing, qty, note in lines:
        recipe_rows.append([name, ing, qty, unit_by_ing[ing], "", "", note])
ing_rows = []
for name, vendor, date, _inv, pack, price, yqty, unit, status, note in INGREDIENTS:
    ing_rows.append([name, vendor, date, pack, price if price is not None else "",
                     yqty, unit, status, "", note])

m_end = 3 + len(menu_rows)        # menu data rows 4..m_end
r_end = 3 + len(recipe_rows)      # recipe rows 4..r_end
i_end = 3 + len(ing_rows)         # ingredient rows 4..i_end

menu_rows[0][3] = f'=ARRAYFORMULA(IF(A4:A{m_end}="",,ROUND(SUMIF($I:$I,A4:A{m_end},$N:$N),2)))'
menu_rows[0][4] = f'=ARRAYFORMULA(IF(A4:A{m_end}="",,ROUND(C4:C{m_end}-D4:D{m_end},2)))'
menu_rows[0][5] = f'=ARRAYFORMULA(IF(A4:A{m_end}="",,ROUND(D4:D{m_end}/C4:C{m_end}*100,1)))'
recipe_rows[0][4] = f'=ARRAYFORMULA(IF(J4:J{r_end}="",,IFERROR(VLOOKUP(J4:J{r_end},$Q:$Y,9,FALSE),0)))'
recipe_rows[0][5] = f'=ARRAYFORMULA(IF(J4:J{r_end}="",,ROUND(K4:K{r_end}*M4:M{r_end},4)))'
ing_rows[0][8] = f'=ARRAYFORMULA(IF(Q4:Q{i_end}="",,IFERROR(ROUND(U4:U{i_end}/V4:V{i_end},4),"")))'

nrows = max(len(menu_rows), len(recipe_rows), len(ing_rows))
out = io.StringIO()
w = csv.writer(out, lineterminator="\n")
w.writerow(["AREPAS HOUSE - MENU PROFITABILITY (menu | recipes | ingredient prices)"])
w.writerow(["Edit recipe Qty (col K) and ingredient prices (cols U/V); everything recalculates. Portions are DRAFT - correct them."])
w.writerow(
    ["Menu Item", "Category", "Menu Price", "Ingredient Cost", "Profit $", "Food Cost %", "Notes", ""]
    + ["Menu Item", "Ingredient", "Qty", "Unit", "Unit Cost", "Line Cost", "Notes", ""]
    + ["Ingredient", "Vendor", "Last Invoice", "Pack", "Pack Price", "Yield Qty", "Unit", "Status", "Unit Cost", "Notes"]
)
for i in range(nrows):
    m = menu_rows[i] if i < len(menu_rows) else [""] * 7
    rr = recipe_rows[i] if i < len(recipe_rows) else [""] * 7
    g = ing_rows[i] if i < len(ing_rows) else [""] * 10
    w.writerow(m + [""] + rr + [""] + g)

sys.stdout.write(out.getvalue())
