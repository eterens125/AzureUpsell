# AzureUpsell

Upsell contest leaderboard for the **Azure** group of stores (Checkmate Pizza, a Domino's franchise; stores 3116, 3802, 3882, 5161, 5166, 8604).

`build.py` reads the Domino's *UPSELL BY TM* reports and produces a single self-contained page, `dist/index.html`:

- overall leader and top store
- team member board (upsell rate, with a toggle for upsell orders; 20-order minimum to be ranked)
- store standings
- coupon-by-store grid
- store detail: per-store team member board (10-order minimum) and every coupon code with who used it

Team members appear as first name + last initial. Internet orders are never ranked; they only show as a memo column in store detail.

## Periods

The page has one tab per Domino's 4-week period (`PERIOD_DEFS` in `build.py`): Period 10 (Sep 7 - Oct 4, final) and Period 11 (Oct 5 - Nov 2, in progress). Each period has its own boards, standings, heatmap, store detail and progress chart. A report must sit inside one period. To start Period 12, add an entry to `PERIOD_DEFS`.

## Update workflow

```
pip install openpyxl pandas
# drop the new reports in data/daily/, named YYYY-MM-DD_YYYY-MM-DD_<coupon>.xlsx (informational; dates come from each report's header)
python3 build.py --rollup     # merge into data/baseline.json, rebuild dist/index.html
git add -A && git commit -m "Update leaderboard" && git push
```

`data/baseline.json` is committed. It holds only aggregate totals per period with team members as "First L." - no full names and no employee codes - so older raw reports never need to be kept. Raw reports (`data/*.xlsx`, `data/daily/*.xlsx`) are gitignored because they contain full names and employee codes. Reports ending on or before a period's recorded `through` date are skipped, so re-running is safe, and a report that straddles that date is refused so a day can't be counted twice.

Because the baseline has no employee codes, a new report row is merged into an existing "First L." only when they share a store; same-named people at other stores are kept as separate people and listed as a NOTE in the build output.

Each report's own Begin/End Date header drives the dates; the build stops if two reports cover the same day for a coupon, and if a report's rows don't add up to its own store totals.

`python3 build.py` without `--rollup` rebuilds the page without changing the baseline.

`template.html` holds the page design (including the Azure banner); `build.py` fills in the data.
