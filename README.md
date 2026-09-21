# AzureUpsell

Upsell contest leaderboard for the **Azure** group of stores (Checkmate Pizza, a Domino's franchise; stores 3116, 3802, 3882, 5161, 5166, 8604).

`build.py` reads the Domino's *UPSELL BY TM* reports and produces a single self-contained page, `dist/index.html`:

- overall leader and top store
- team member board (upsell rate, with a toggle for upsell orders; 20-order minimum to be ranked)
- store standings
- coupon-by-store grid
- store detail: per-store team member board (10-order minimum) and every coupon code with who used it

Team members appear as first name + last initial. Internet orders are never ranked; they only show as a memo column in store detail.

## Rebuild

```
pip install openpyxl pandas
python3 build.py
```

Put the reports in `data/` (not committed, because they contain full employee names and codes):

- `data/<period workbook>.xlsx` - multi-sheet report, one sheet per coupon code
- `data/daily/YYYY-MM-DD_<coupon>.xlsx` - one single-day report per coupon (11 per day)

Each report's own Begin/End Date header drives the period label; the build stops if two reports cover the same day for a coupon, and if a report's rows don't add up to its own store totals.

To add a new day, drop that day's 11 files into `data/daily/` and run `python3 build.py`.

`template.html` holds the page design (including the Azure banner); `build.py` fills in the data.
