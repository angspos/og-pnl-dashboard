# Olive Garden P&L Dashboard

An interactive Excel/VBA profit-and-loss dashboard modeled on a multi-store regional report for Olive Garden (Darden Restaurants). All figures in this repo are **dummy data** generated for educational and portfolio purposes. Nothing here is real Darden financial information.

## What it does

Eight stores' worth of weekly, quarterly, and year-to-date sales and P&L data, each on its own dashboard sheet, plus a side-by-side comparison view. The dashboard collapses dense corporate reporting into a single screen by hiding/showing column groups and chart overlays through VBA-driven toggles.

Toggles (each acts on all dashboard sheets at once):

- **Growth %** vs. **Growth $** vs. **Growth Per Guest** — switch how growth-vs-last-year is expressed
- **PnL / PnL min** — expand or collapse the full line-item P&L block
- **GSTOP / GSTOP Graph** — switch GST/OP between number view and chart view

Switch sheets and the view stays consistent — every store reflects the toggle you just clicked.

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

## How the data loads

The `Data` module loops the eight text files into eight hidden `<Store> Data` sheets that feed the dashboards. Run **`RunAll`** from the macro menu (Tools > Macros, or Opt+F8 / Alt+F8) to reload every store at once.

The path is relative — `txtdata/` just needs to live next to the workbook. Move the folder anywhere on disk and it keeps working.

## How the toggles work

Each button calls a one-line wrapper (`GrowthActual`, `PnL`, etc.) that hands off to `ApplyToDashboards <mode>`. That helper:

1. Saves the current `ScreenUpdating` and `Calculation` state
2. Loops every visible worksheet that has the dashboard's signature shapes
3. Applies the per-mode column-width and shape-visibility logic
4. Restores the saved state on exit, even if a sheet is missing a shape

The shape and range setters are defensive — a missing shape on one sheet doesn't abort the whole run. See `Stores.bas` for the full module.

## Running the workbook

1. Open `PnL-Dashboard.xlsm` in Excel (macros must be enabled)
2. To repopulate from the dummy data: **Tools > Macros > RunAll > Run**
3. Use the toggles on any dashboard sheet — every sheet syncs

## Background

I built the original version of this in 2018 as an Assistant Manager at the Tacoma Olive Garden, self-teaching VBA to make corporate weekly P&L reports readable across eight Pacific Northwest stores. This repo is a clean, dummy-data version of that work — refactored in 2026 to apply toggles across all sheets at once instead of one-at-a-time, with proper state restoration and defensive error handling.

## License

MIT — see [LICENSE](LICENSE).
