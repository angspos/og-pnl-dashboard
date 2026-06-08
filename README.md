# Olive Garden P&L Dashboard

An interactive Excel/VBA profit-and-loss dashboard modeled on multi-store regional reporting for Olive Garden (Darden Restaurants). All figures are **dummy data** — nothing here is real Darden financial information.

## What it's for

Corporate P&L reports are built for finance teams, not for the people who actually move the numbers. A server stretched on a Friday night, a line cook watching food cost, a manager covering a close — they don't open a 30-row P&L statement. They want one screen that tells them whether their store is up or down versus last year, in real time.

This dashboard serves two audiences from the same underlying data.

**The full P&L view** preserves the structure of a real P&L statement: revenue at the top, line-item variable costs and labor in the middle, operating income at the bottom. Current period vs. prior period side by side, with GST/OP/WK, actual dollars, and percentages on every line. Useful for anyone learning how a restaurant P&L flows, and for managers reviewing the full picture.

**The vs-last-year view** strips the same data down to what operators actually need on shift: am I up or down, and by how much? Green or red bars next to every line, weighted by magnitude, readable in a glance.

## How it toggles

Each toggle acts on every dashboard sheet at once, so switching stores never resets the view:

- **% vs $ vs Per Guest** — switch how growth-vs-last-year is expressed
- **Month / Quarter / Year** — change the time period across every section
- **PnL / PnL min** — expand the full line-item P&L or collapse it to a summary
- **GSTOP / GSTOP Graph** — toggle the GST/OP block between numbers and chart view

A regional manager can flip Vancouver → Tacoma → Olympia → Federal Way without re-toggling each one.

## Regional comparison

The **Side By Side** sheet pulls every store's headline numbers onto a single comparison view, so a regional manager can spot which store is dragging the region down or pulling it up without opening eight tabs.

## Data

Eight dummy text files in `txtdata/`, one per store. The `Data` module reads them into hidden `<Store> Data` sheets that feed every dashboard. To repopulate: **Tools > Macros > RunAll > Run** (or **Opt+F8** / **Alt+F8** to open the macro menu).

The path is relative — `txtdata/` just needs to live next to the workbook. Move the folder anywhere on disk and it still works.

## Under the hood

Each button calls a one-line wrapper (`GrowthActual`, `PnL`, etc.) that hands off to `ApplyToDashboards <mode>`. That helper:

1. Saves the current `ScreenUpdating` and `Calculation` state
2. Loops every visible worksheet that has the dashboard's signature shapes
3. Applies the per-mode column-width and shape-visibility logic
4. Restores the saved state on exit, even if a sheet is missing a shape

Missing shapes and ranges fail silently — one sheet's drift doesn't abort the whole run. See `Stores.bas` for the full module.

## Layout

```
PnL-Dashboard.xlsm        Main workbook (macro-enabled)
Stores.bas                Standalone copy of the dashboard-toggle VBA module
                          (mirrors what's inside the workbook; here for readability)
txtdata/                  Dummy store-level data, one .txt per store
  Federal Way.txt
  Mill Plain.txt
  Olympia.txt
  Puyallup.txt
  Silverdale.txt
  Tacoma.txt
  Tukwila.txt
  Vancouver Mall.txt
```

## Background

I built the original version of this in 2018 as an Assistant Manager at the Tacoma Olive Garden, self-teaching VBA to make weekly corporate P&L reports readable across eight Pacific Northwest stores. The two-audience design came from running shifts — finance views weren't reaching the people who actually moved the numbers. This repo is a clean, dummy-data version of that work, refactored in 2026 for multi-sheet sync, proper state restoration, and defensive error handling.

## License

MIT — see [LICENSE](LICENSE).
