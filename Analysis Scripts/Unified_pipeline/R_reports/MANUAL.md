# EMT Report — User Manual

This folder has two tools:

| Tool | Input | Use it for | Run with |
|---|---|---|---|
| **EMT report** (§1–9) | PB-analyzed Excel workbooks across batches | Figures, metric superplots and statistics comparing conditions | `render-report.R` |
| **Plate trace report** (§10) | Raw time-force CSVs from one 24-well plate | Looking at every well's raw trace in plate layout: QC before analysis | `render-plates.R` |
| **Both, from the launcher** (§11) | A raw data folder + plate map | One click: Python analysis, then both reports | `Master_Analysis_Launcher.py` → `render-pipeline.R` |

The EMT report turns PB-analyzed EMT Excel workbooks into a single self-contained HTML
report, plus figure-ready PNG panels and statistics CSVs. Everything that is
specific to one dataset (where the files are, what the conditions are called,
colors, which batches to use) lives in **one YAML config file**. The R code
never needs editing for a new dataset.

---

## 1. Quick start (new dataset, 5 steps)

1. **Make a batch CSV** listing every data folder (§3).
2. **Copy the example config:** `cp configs/examples/dataset-example.yml configs/mydataset.yml`
   (dataset configs in `configs/` are personal and git-ignored)
3. **Edit the config:** set `name`, `batch_csv`, `output_dir`, `conditions`, `colors` (§4).
4. **Render:**
   ```bash
   cd R_reports    # from the project folder
   Rscript render-report.R configs/mydataset.yml
   ```
5. **Open** `output/mydataset/mydataset-emt-report.html`. Read the
   *Data summary* section first for warnings (§7).

---

## 2. What's in the folder

| Path | What it is |
|---|---|
| `render-report.R` | Entry point. Renders the report once per config. |
| `explore.R` | Starter script for interactive exploration in an R session (§6). |
| `emt-report.qmd` | The report template (Quarto). Edit only to change report layout. |
| `configs/*.yml` | One config per dataset. **This is what you edit.** Git-ignored. |
| `configs/examples/` | Templates: a dataset config, a batch CSV, and a plate config for the sample data. |
| `R/emt-config.R` | Loads a config, fills defaults, validates it. |
| `R/emt-import.R` | Reads the batch CSV and Excel workbooks. |
| `R/emt-plots.R` | Trace plots, superplots, fatigue curves, PNG saving. |
| `R/emt-stats.R` | Condition-vs-reference t-tests. |
| `R/emt-data.R` | `load_emt_data()` and plot shortcuts for interactive use. |
| `render-plates.R` | Entry point for the plate trace report (§10). |
| `plate-report.qmd` | Plate trace report template. |
| `configs/plates/*.yml` | Optional plate trace configs (§10.3). Git-ignored. |
| `R/plate-data.R` | Plate map, time-force file discovery and reading. |
| `R/plate-plots.R` | The plate-layout grid plot. |
| `../sample_data/` | One example batch (plate map, plus `timeforce/` recordings if you copied them in). |
| `renv.lock`, `renv/`, `.Rprofile` | Pinned R packages (renv). `.Rprofile` activates the project library whenever R starts in this folder. |
| `EMT-Reports.Rproj` | RStudio project. Open it so the working directory and package library are right. |
| `output/<dataset>/` | Everything a render produces (§5). |

### Requirements

- **R 4.5** with the packages pinned in `renv.lock`. The project's `setup.sh` /
  `setup.bat` installs them; by hand, start R in this folder and run
  `renv::restore()`. They go into `renv/library/`, not your personal library.
  After adding a package to the code, run `renv::snapshot()` and commit `renv.lock`.
- **Quarto.** `uv sync` in the project folder installs it into `../.venv/`, and the
  R scripts here use that copy automatically. If it's missing they fall back to
  `QUARTO_PATH`, your PATH, a standalone install, or RStudio's bundled copy.
- **OneDrive files available offline** (or at least synced) for every data folder
  in the batch CSV. Cloud-only placeholders can fail to open.

---

## 3. Input data

### 3.1 Workbook folders

Each folder holds the analyzed workbooks for **one batch × one analysis type**,
one workbook per day of culture:

```
.../Batch-1/analysis-results/FF/
    d10_FF_waveformkinetics.xlsx
    d12_FF_waveformkinetics.xlsx
    d14_FF_waveformkinetics.xlsx
    ...
```

Rules:

- The file name **must start with `d<day>_`** (e.g. `d7_`, `D14_`). That number
  is the day. Files without it are skipped (with a warning in the report).
- If a day has more than one `.xlsx` (e.g. a re-run), the **most recently
  modified** one is used and the report warns you which.
- Other files (`.md`, `.csv`, Excel lock files `~$...`) are ignored.

Required sheets and columns per analysis type:

| Analysis (`analysis` in CSV) | Trace sheet | Metrics sheet |
|---|---|---|
| `ff` (force-frequency) | `Representative Traces` | `Relaxation Summary` (needs a `Hz` column) |
| `75hz` (kinetics) | `Representative Traces` | `Relaxation Summary` |
| `fatigue` | `Condition Avg Fatigue` | `Fatigue Summary` |

- `Representative Traces` columns: `Time (seconds)`, `Normalized Trace`,
  `Standard Deviation`, `Force (μN)`, `Force SD`, `Smoothed Force (μN)`,
  `Smoothed Force SD`, `Condition`
- `Condition Avg Fatigue` columns: `Condition`, `Stim Number`,
  `Mean Normalized Force (% of T1)`, `StdDev Normalized Force (% of T1)`,
  `N (Wells contributing)`
- Metrics sheets: identifier columns `Plate`, `Well`, `Condition` (plus
  `Stim Peak Time`, `Stim Number`, `Hz` where present). **Every other numeric
  column is treated as a metric** and is plotted in the sweep and tested in the
  stats — no list to maintain.
- **One value per well.** Where a sheet has one row per stim (75 Hz: three tetani
  per well), every metric is averaged across the stims of each well, by day and
  Hz. Columns ending in `Well Avg` are ignored, so use the per-stim names
  (`R50 Time`, `TT50P Time`, …), which the report averages itself.

Columns are matched **by name**, not position. A missing sheet or column stops
the render with a message naming the batch, file and what's missing.

### 3.2 The batch CSV

One row per batch × analysis. Columns:

| Column | Required | Meaning |
|---|---|---|
| `batch` | yes | Batch (differentiation) number. Any integers; gaps are fine. |
| `analysis` | yes | `ff`, `75hz` or `fatigue` (case-insensitive). |
| `directory` | yes | Full path to the folder of workbooks. |
| `exclude` | no | `1` = skip this row; `0` or blank = use it. |

Example:

```csv
batch,analysis,directory,exclude
1,ff,/Users/you/Library/CloudStorage/OneDrive-UW/.../Batch-1/analysis-results/FF/,0
1,75hz,/Users/you/Library/CloudStorage/OneDrive-UW/.../Batch-1/analysis-results/75Hz/,0
1,fatigue,/Users/you/Library/CloudStorage/OneDrive-UW/.../Batch-1/analysis-results/Fatigue/,0
2,ff,/Users/you/Library/CloudStorage/OneDrive-UW/.../Batch-2/analysis-results/FF/,0
2,75hz,/Users/you/Library/CloudStorage/OneDrive-UW/.../Batch-2/analysis-results/75Hz/,0
2,fatigue,/Users/you/Library/CloudStorage/OneDrive-UW/.../Batch-2/analysis-results/Fatigue/,1
```

Tips:

- **Get exact paths.** macOS: drag the folder from Finder into Terminal, or
  right-click → hold Option → *Copy "…" as Pathname*. Windows: Shift + right-click
  the folder → *Copy as path*, then drop the surrounding quotes. Forward slashes
  (`C:/Users/you/OneDrive - UW/...`) work in R and avoid escaping problems.
- **Watch the lab folder name.** The local OneDrive path is
  `Regnier_Mack Labs` (underscore). Paths copied from the web/SharePoint may say
  `Regnier+Mack`, which won't exist locally.
- **Paths containing commas** must be wrapped in double quotes.
- You can omit an analysis type entirely (e.g. no fatigue rows); those report
  sections are simply skipped.
- Save the CSV in `configs/`, next to your config (`batch_csv: configs/<name>.csv`;
  paths in configs are relative to this `R_reports/` folder). Batch CSVs hold paths
  on your machine, so `configs/*.csv` is git-ignored. A template is in
  `configs/examples/example-batches.csv`. Naming convention so far:
  `YYYY-MM-DD_<DATASET>_EMT_data_files.csv`.

---

## 4. The config file

Start from `configs/examples/dataset-example.yml` (copy it to `configs/`). Paths are
relative to the `R_reports/` folder.

### 4.1 Settings you must set

```yaml
name: MYDATASET                              # shown in the report; also names the HTML file
batch_csv: 2026-11-17_MYDATASET_EMT_data_files.csv
output_dir: output/mydataset                 # use a new folder per dataset

conditions:                                  # display label -> raw labels in the workbooks
  WT:      [UC2WT]
  MutHet:  [HetMut-Cl12, HetMutCl12]
  MutHom:  [HomMut-Cl7]
```

**`conditions`** is the most important setting:

- Left side = the name you want in plots/tables. Right side = every spelling of
  that condition that appears in the workbooks' `Condition` column (clone names
  often differ between batches, e.g. `T178I-Cl113` vs `T178ICl113`).
- The display label itself also matches automatically (so a batch that already
  says `WT` maps to `WT` without listing it).
- **Order matters:** it sets legend/plot order, and the **first condition is
  the reference** every other condition is compared against in the stats.
- Labels not listed are kept as-is and flagged under *Import warnings* in the
  report — that's your cue to add them.

To find the raw labels, open one workbook from each batch and look at the
`Condition` column, or in R:

```r
unique(readxl::read_excel("path/to/d10_FF_waveformkinetics.xlsx", "Relaxation Summary")$Condition)
```

### 4.2 Settings you'll often change

```yaml
colors:                 # optional; missing ones are filled from a colorblind-safe palette
  WT:     "#000000"
  MutHet: "#009E73"
  MutHom: "#E69F00"

labels:
  condition: Genotype   # legend title for color
  batch: Differentiation  # legend title for point shape

days_max: 7             # only show/test days <= this; omit (or .inf) for all days

trace_batches: [1]      # batches whose average traces are drawn in panels A/B and fatigue curves
metric_batches: all     # batches used in metric panels, sweep and stats; or a list, e.g. [1, 4, 5, 8]
```

- `metric_batches: all` = every non-excluded batch in the CSV. If you give a
  list instead, the report notes which non-excluded batches were left out.
- Use `exclude` in the CSV to drop a batch everywhere; use `metric_batches` to
  drop it from metrics/stats only.

### 4.3 Panels

```yaml
kinetics:
  stim_time_s: 48       # which 75 Hz stim to window for panels A/B
  window_s: 3           # seconds shown after the stim
  baseline_s: 0.01      # baseline = mean force over (stim, stim + baseline_s]

kinetics_metrics:       # panels C, D, E, ... from the 75 Hz metrics sheet
  "Max tetanic-BK (μN)": "Active Force (μN)"   # sheet column name: y-axis label
  "TT50P Time": "Time to 50% Peak (s)"
  "R50 Time": "Time to 50% Relax (s)"
  "TTP90 Time": "Time to 90% Peak (s)"

ff_metric:              # the force-frequency panel (next letter after kinetics)
  name: "FF Normalized (%)"
  label: "Percent of 100Hz Tension (%)"

fatigue_metrics:        # fatigue summary panels
  "Fatigue Resistance Index (%)": "Fatigue Resistance Index (%)"
  "Last Normalized Force (% of T1)": "Final Force (% of first)"
```

Metric names must match the sheet column header **exactly** (including `μ`,
`%` and spaces). Panel letters are assigned automatically, so adding or
removing a kinetics metric re-letters the panels after it.

### 4.4 Statistics

```yaml
stats:
  unit: batch           # batch (recommended) or well
  paired: false         # true = paired t-test on batch means
  p_adjust: BH          # any p.adjust() method: BH, bonferroni, holm, none, ...
```

- Each condition is compared to the reference (first condition) with a t-test,
  separately for every metric × day (× Hz for force-frequency).
- `unit: batch` averages wells within each batch first, so *n* = number of
  batches. **With only 1–2 batches most tests will be `n < 2` or underpowered —
  that's expected, not a bug.**
- `unit: well` treats every well as independent (pseudo-replication); p-values
  will look smaller than they should. Use only for exploration.
- `paired: true` requires `unit: batch` and uses only batches that contain both
  conditions.
- p-values are adjusted within each metric (across days, Hz and comparisons).

### 4.5 Other settings

```yaml
sweep: true             # false skips the all-metrics section (faster render)
seed: 1                 # jitter seed, so point positions are reproducible
panel_sizes:            # saved PNG width, height in inches
  traces:  [10, 2.5]
  metric:  [4, 4]
  ff:      [14, 3.5]
  fatigue: [10, 2.5]
  sweep:   [8, 6]
```

Anything you leave out of the config falls back to the defaults in
`R/emt-config.R`. The only required keys are `batch_csv` and `conditions`.

---

## 5. Running it

### From Terminal

```bash
cd R_reports    # from the project folder
Rscript render-report.R configs/mydataset.yml   # one dataset
Rscript render-report.R                          # every config in configs/
```

### From RStudio

Open `EMT-Reports.Rproj` (so the working directory is this folder and renv is active), then:

```r
source("render-report.R")
render_emt_report("configs/mydataset.yml")
```

The console prints `Report: /full/path/...html` when done. A full render with the
sweep takes a few minutes, mostly reading workbooks over OneDrive and saving
PNGs.

### What you get

```
output/mydataset/
├── mydataset-emt-report.html   # self-contained; email or share as one file
├── panels/                     # figure panels, 300 dpi, no legends — for assembling figures
│   ├── A_normalized_force.png
│   ├── B_active_force.png
│   ├── C_Max_tetanic_BK_uN.png ...
│   ├── G_FF_FF_Normalized_pct.png
│   ├── fatigue_curves.png
│   └── fatigue_<metric>.png
├── all-metrics/                # one superplot per metric per analysis, 150 dpi
│   └── 75hz_<metric>.png, ff_<metric>.png, fatigue_<metric>.png
└── stats/
    ├── stats_75hz.csv
    ├── stats_ff.csv
    └── stats_fatigue.csv
```

Re-rendering overwrites these files. Filenames replace `μ` with `u`, `%` with
`pct` and other symbols with `_`.

Stats CSV columns: `metric, day, [Hz], comparison, n_ref, n_cmp, mean_ref,
mean_cmp, difference (cmp − ref), p_value, p_adj, note, unit`. `note` explains
a missing p-value (`n < 2`, `no variance`).

### Report sections

1. **Data summary** — settings used, batches left out/excluded, QC and import
   warnings, wells per condition × day × batch.
2. **Figure panels** — A/B 75 Hz traces, kinetics metric superplots, force-frequency.
3. **Fatigue** — fatigue curves and summary metrics.
4. **Statistics** — panel-metric table, then searchable tables of every metric.
5. **All-metrics sweep** — every numeric column in every metrics sheet.

**Reading superplots:** faint points = individual wells, bold points = batch
means (shape = batch), x = day (or Hz for force-frequency, faceted by day).

---

## 6. Exploring interactively in R

Use this to poke at the data and try new plots without re-rendering the whole
report. `explore.R` is a worked example; open it and run it section by section.

### Load once

Open `EMT-Reports.Rproj` in RStudio, then:

```r
library(tidyverse)
for (f in list.files("R", pattern = "\\.R$", full.names = TRUE)) source(f)

emt <- load_emt_data("configs/mydataset.yml")
# override batches without editing the config:
# emt <- load_emt_data("configs/mydataset.yml", metric_batches = "all", trace_batches = c(1, 2))
```

This is the slow step (it reads every workbook over OneDrive), so do it once per
session. To skip it after a restart, cache it with
`saveRDS(emt, "output/mydataset/emt-data.rds")` and reload it with `readRDS()`.
Re-run `load_emt_data()` whenever the workbooks change.

`emt` is a list:

| Element | Contents |
|---|---|
| `emt$cfg` | The loaded config (`colors`, `conditions`, `days_max`, `labels`, ...) |
| `emt$m75` | 75 Hz metrics: one row per well × metric |
| `emt$mff` | Force-frequency metrics: one row per well × Hz × metric |
| `emt$mfat` | Fatigue summary metrics: one row per well × metric |
| `emt$traces75` | 75 Hz condition-average traces (trace batches) |
| `emt$fat_tr` | Fatigue condition-average curves (trace batches) |
| `emt$metric_batches`, `emt$trace_batches` | Batch numbers in use |

Metric tables are long format with columns `Plate, Well, Condition, [Hz],
batch, day, formal_date, metric, value`. That makes filtering and plotting
any metric straightforward.

### Plot like the report

```r
unique(emt$m75$metric)    # what you can plot

emt_superplot(emt, emt$m75, "R50 Time", y_label = "Time to 50% Relax (s)")
emt_superplot(emt, emt$mfat, "Fatigue Resistance Index (%)")
emt_ff_superplot(emt, "FF Normalized (%)")

# override any config value for one plot
emt_superplot(emt, emt$m75, "Max tetanic-BK (μN)", days_max = Inf, batches = 2)

kin <- prep_kinetics_traces(emt$traces75, stim_time_s = 48, window_s = 3)
kinetics_trace_plot(kin, "normalized", emt$cfg$colors)
fatigue_curve_plot(emt$fat_tr, emt$cfg$colors, emt$cfg$days_max)
```

All of these return ordinary ggplot objects, so you can add layers:
`p + facet_wrap(~batch)`, `p + scale_y_continuous(limits = c(0, NA))`, etc.

### Your own plots and stats

```r
emt$m75 |>
  filter(metric == "R50 Time", day <= 7) |>
  ggplot(aes(factor(day), value, color = Condition)) +
  geom_boxplot() +
  scale_color_manual(values = emt$cfg$colors)

compare_conditions(filter(emt$m75, day <= 7), by = "day")              # same tests as the report
compare_conditions(filter(emt$mff, day <= 7), by = c("day", "Hz"))
```

### Save

```r
save_panel(p, "R50_custom", size = c(4, 4), dir = "output/mydataset/custom")   # 300 dpi PNG
ggsave("output/mydataset/custom/my_plot.pdf", p, width = 4, height = 4)
```

When a plot becomes a keeper, add it to the config (e.g. a new entry under
`kinetics_metrics`) so the report produces it on every render.

**Tip:** if you change a file in `R/`, re-run the `source()` loop to pick up
the change. You don't need to reload the data.

---

## 7. Checking the results

Always read the **Data summary** before trusting the figures:

| Message | What it means | What to do |
|---|---|---|
| *Condition label(s) not in config `conditions`* | A workbook uses a spelling you haven't mapped; it shows up as an extra condition. | Add it to the right list in `conditions`. |
| *QC: wells with zero or negative active force* | Peak detection probably failed upstream for that file. Normalized metrics for those wells are meaningless. They are **not** removed. | Re-run the upstream analysis for that day, or exclude the batch. |
| *skipping file(s) without a leading d<day>_* | A workbook name doesn't start with `d<number>_`. | Rename it if it should be included. |
| *multiple workbooks for day(s) …; using the most recent* | Old and new versions are in the same folder. | Confirm the right file was used; move old versions out. |
| *Batch(es) … are not in `metric_batches`* | Informational: you chose a subset. | Nothing, if intended. |

Also check the **Wells per condition** table: a 0 or unexpectedly low count
usually means a condition label wasn't mapped or a day is missing.

---

## 8. Troubleshooting

| Error | Fix |
|---|---|
| `Quarto not found` | Run `uv sync` in the project folder, install Quarto, or set `QUARTO_PATH` (§2). |
| `there is no package called ...` | Run `renv::restore()` in this folder (or the project's setup script). |
| `Config file not found` / `batch_csv not found` | Run from the project folder; paths in the config are relative to it. |
| `Batch N: directory does not exist` | Fix the path in the batch CSV (check `Regnier_Mack` vs `Regnier+Mack`, trailing spaces, that OneDrive is synced). |
| `Batch N: no .xlsx files in …` | Wrong folder, or files aren't downloaded from OneDrive. |
| `sheet '…' not found in …` | The workbook came from a different analysis version, or the CSV row has the wrong `analysis` type for that folder. |
| `… is missing column(s): …` | The workbook format changed. If it's a deliberate rename, update the column map at the top of `R/emt-import.R`. |
| `No non-excluded '…' rows` | Every row for that analysis is excluded, or `metric_batches`/`trace_batches` lists batches that aren't in the CSV. |
| A panel is empty | The metric name in `kinetics_metrics` / `fatigue_metrics` / `ff_metric` doesn't exactly match a column header, or `days_max` filters out all days. |
| `Quarto render failed` with an R error | Scroll up in the console; the R error above it names the problem. To debug interactively, open `emt-report.qmd` in RStudio, set `params$config` at the top, and run chunks one by one. |

---

## 9. Checklist for a new dataset

- [ ] All workbook folders synced locally; filenames start with `d<day>_`
- [ ] Batch CSV created with `batch, analysis, directory, exclude`
- [ ] Config copied from `configs/examples/` into `configs/`, and `name`, `batch_csv`, `output_dir` set
- [ ] `conditions` lists every raw label, reference condition first
- [ ] `days_max`, `trace_batches`, `metric_batches` set
- [ ] Rendered; *Import warnings* empty or understood
- [ ] *QC* table reviewed
- [ ] Wells-per-condition counts look right

---

## 10. Plate trace report (raw time-force data)

This report covers the step **upstream** of the EMT report. For one plate, it draws
every well's raw force trace in the same 4 × 6 layout as the plate. That makes it easy
to spot dead wells, drifting baselines, plate-map mistakes or a protocol that
didn't run, before anything is analyzed.

### 10.1 Quick start

Point it at a batch folder:

```bash
cd R_reports    # from the project folder
Rscript render-plates.R "/Users/you/Library/CloudStorage/OneDrive-UW/Regnier_Mack Labs/.../2026-11-05_R672C-EMT-Batch3"
Rscript render-plates.R ../sample_data             # the bundled example
```

From RStudio: `source("render-plates.R")`, then `render_plate_report("<folder or config>")`.

The report is saved to `output/plates/<folder name>/<name>-plate-traces.html`.
PNGs (200 dpi) go to `output/plates/<folder name>/plate-traces/<protocol>/d<day>_<protocol>.png`.
Animated GIFs (one per protocol, stepping through the days) go to
`output/plates/<folder name>/animations/<protocol>_all-days.gif`.
A batch with ~7 days × 7 protocols takes 1–2 minutes.

### 10.2 What it expects

**Where it looks:** point it at any folder. It reads from `<folder>/timeforce/` if
that exists, otherwise from the folder itself, including subfolders at any depth.
All of these work:

```bash
Rscript render-plates.R ".../2026-11-05_R672C-EMT-Batch3"                           # has timeforce/
Rscript render-plates.R ".../2026-01-19_T178I-EMT-cl113cl152_batch-8/timeforce"
Rscript render-plates.R ".../2025-11-17_T178I-EMT-cl113cl152-batch-4/pulse3D-files"
Rscript render-plates.R ".../2025-05-16_DA-EMT_Batch-1/mantarray-recordings/pulseFiles/T178I"
```

**Plate map:** the tool looks for `*Plate Map*.csv` in that folder, then up to 3
folders above it. If one folder holds several maps (Batch 1 has `R672C Plate Map.csv`
and `T178I Plate Map.csv`), it uses the one whose name appears in the data
folder's path (`…/pulseFiles/T178I` → T178I). If it can't tell which to use, it says
so; point at the plate's own folder or use a config. 2025 T178I Batch 3 has no plate
map at all, so it needs a config (§10.3).

The plate is named after the batch folder holding the plate map. Folders named
`timeforce`, `pulse…` or `mantarray…` are skipped. The plate name is added when it's
needed to tell plates apart, e.g. `2025-05-16_DA-EMT_Batch-1 T178I`. The output goes
to `output/plates/<that name>/`.

**Plate map format:** the plate layout as a CSV. The first column holds row letters,
the header holds column numbers, and each cell holds that well's sample/group label.
Blank cells are shown as `(empty)`.

```
,1,2,3,4,5,6
A,UC2WT,UC2WT,R672CHet,R672CHet,R672CHom,R672CHom
B,...
```

**Time-force files:** one file per day × protocol, in either format:

| Format | Used by | What's read |
|---|---|---|
| CSV export | 2026 batches (`timeforce/`) | `Time (s)` + one column per well (`A1`, `B1`, …) |
| Pulse3D "full" workbook, `*__full__*.xlsx` | 2025 batches (`pulse3D-files/`, `pulseFiles/`, `pulse-analysis/`) | sheet `continuous-waveforms`: `Time (seconds)` + `A1 - Active Twitch Force (μN)` … |

Other `.xlsx` files, such as the PB-analyzed `d10_FF_waveformkinetics.xlsx`, are
ignored, so pointing at a whole batch folder won't pick them up.

**File names:** the day and protocol are read from the file name:

| File name | Day | Protocol |
|---|---|---|
| `d7_R672C-Batch-3_FF.csv` | 7 | FF |
| `ML26187408-6__2026_07_24_230936_d7_DAEMTR672CBatch2_FF.csv` | 7 | FF |
| `ML25296405-6_2026_01_26_235812__d10_T178I-Cl113Cl152-batch-8_1Hz.csv` | 10 | 1Hz |
| `ML25101401-2__2025_05_19_212941_d3_DAEMTT178I_75Hz__full__2025-05-23_15-01-46.xlsx` | 3 | 75Hz |

The rules:

1. A Pulse3D `__full__<timestamp>` ending is dropped.
2. Everything up to the first `__` is an instrument/timestamp prefix and is ignored
   when looking for the day.
3. The day is `d<number>_` at the start of what's left or after an underscore.
4. The protocol is the last `_`-separated word.

Protocol names ignore case, `-`, `_` and spaces. So `75HZ` → `75Hz`, and
`ASfatigue` / `ASFatigue` / `AS-Fatigue` all become one protocol. Genuinely different
names, like `Priming-round2`, `afterResting75Hz` or `primingg`, get their own section,
drawn as the whole recording. If a day × protocol has more than one file, the newest
is plotted and the report lists the others.

### 10.3 Config (optional)

Use a config to map plate-map labels to the display names and colors your EMT
report uses, pick days, change plot windows, or point at a plate map elsewhere.
Copy `configs/examples/plate-sample-data.yml` to `configs/plates/` and edit it:

```yaml
name: R672C Batch 3
timeforce_dir: /…/2026-11-05_R672C-EMT-Batch3/timeforce   # required
plate_map: /…/2026-11-05_R672C-EMT-Batch3/R672C Plate Map.csv  # required
output_dir: output/plates/r672c-batch3
days: all                  # or [3, 7]
conditions:                # optional, same format as the EMT config
  WT: [UC2WT]
colors:
  WT: "#000000"
animation: true            # GIF per protocol across days
animation_delay_s: 1       # seconds per day
```

Run it with `Rscript render-plates.R configs/plates/r672c-batch3.yml`. To render every
config in `configs/plates/`, run `Rscript render-plates.R` with no arguments.

### 10.4 How each protocol is drawn

| Protocol | Default | Why |
|---|---|---|
| `02Hz` | Whole protocol, **0–55 s** | Stims at 5, 10 … 50 s |
| `1Hz` | Whole protocol, **0–35 s** | Stims each second from 3 to 32 s |
| `75Hz` | One tetanus: 0.2 s before to 3 s after **t = 48 s** | Same stim the EMT report's panels A/B use |
| `Fatigue` | Whole protocol, **0–60 s** | Stimulation runs from ~3 s to ~55.5 s in every Batch 2/3 file |
| `FF` | Whole protocol, **0–250 s** | Stimulation runs from ~3 s to ~240 s |
| anything else (e.g. `ASFatigue`, `Priming`) | Whole recording | No default; add one in the config if needed |

A snapshot uses one time window for the whole plate, because the plate is
stimulated together, so x = 0 is the same moment in every well. Full-protocol
plots drop anything recorded after `end_s`, which is the recording left running.
If a file stops before `end_s`, the report says so.

To change one, override it in a config. You only need the keys you're changing:

```yaml
protocols:
  1Hz:       {type: snapshot, time_s: 20, pre_s: 0.1, post_s: 0.9}   # one twitch instead
  Fatigue:   {end_s: 45}
  ASFatigue: {type: full, end_s: 40}
  Priming:   {type: snapshot, time_s: 10, pre_s: 0.2, post_s: 2}
```

### 10.5 Reading the plots

- One panel per well, in plate position. The panel title shows the well and its **raw
  plate-map label**, so you can check the map. Line color uses the mapped group name.
- The grey `max N` is the highest force (μN) in that panel's window.
- **Per-day plots:** every well in a day's plot has the same x and y limits, so
  weak or dead wells are obvious. Limits are set from that day's data.
- **All days ▶ tab (animation):** one frame per day, with x and y limits fixed
  across **all days** of that protocol, so changes over time are directly
  comparable. Early, weak days can look nearly flat next to later days; check the
  per-day tab for their shape. Needs at least two days. Turn it off with
  `animation: false`.
- Full-protocol traces are thinned to 4,000 points per well for plotting. Each
  bin's minimum and maximum are kept, so peaks are exact. Set `max_points: 0` to
  plot every sample.
- The force is raw (no baseline subtraction), so passive tension and drift are
  visible.

### 10.6 Interactively

```r
library(tidyverse)
for (f in list.files("R", pattern = "\\.R$", full.names = TRUE)) source(f)

cfg   <- load_plate_config("../sample_data")          # folder or yml
pm    <- read_plate_map(cfg$plate_map, cfg$conditions)
files <- list_timeforce_files(cfg$timeforce_dir)
files                                                  # day, protocol, file

tr <- load_protocol_trace(files$file[files$protocol == "FF"], "FF", cfg)
plate_grid_plot(tr, pm, plate_colors(pm$group, cfg$colors))

# shared limits across several days, like the animation
trs <- lapply(files$file[files$protocol == "FF"], load_protocol_trace, protocol = "FF", cfg = cfg)
lim <- plate_limits(trs)
plate_grid_plot(trs[[1]], pm, plate_colors(pm$group), limits = lim)

# any window, any file: long table of time, well, force
raw <- read_timeforce(files$file[1], t_from = 40, t_to = 60)
```

---

## 11. Running everything from the Master Analysis Launcher

`Python_analyze/Master_Analysis_Launcher.py` runs the Python analyses and then
these reports in one go. In its **Reports (R)** section:

| Option | What it does | Output |
|---|---|---|
| Plate trace report | §10 report on the raw files in the data folder, using the plate map you picked | `<batch>/reports/plates/` |
| Single-batch EMT report | EMT report on this batch's `analyzed/75Hz`, `FF` and `Fatigue` workbooks only | `<batch>/reports/emt/` |
| Add to dataset report | Adds/updates this batch's rows (as batch #N) in the chosen config's batch CSV, then re-renders that dataset's report | the config's `output_dir` |

`<batch>` is the folder above the data folder, next to `analyzed/`. Reports
render even when no files were new, so you can re-run just the reports after
changing a config. Twitch results aren't used by any report.

- **Single-batch settings:** panels come from `configs/single-batch-defaults.yml`,
  or from the dataset config when "Add to dataset report" is ticked. Conditions come
  from that dataset config, or otherwise straight from the plate-map labels, with the
  first label in plate order as the reference. Stats use `unit: well`, because one
  batch can't give batch-level p-values. Treat them as exploratory.
- **Dataset batch CSV:** existing rows for that batch × analysis keep their
  `exclude` value and any extra columns. Only `directory` is updated. Before any
  change, the CSV is copied to `<csv>.bak-<timestamp>`.
- **Generated configs** are saved in `<batch>/reports/_configs/`. To re-render by
  hand: `Rscript render-report.R <batch>/reports/_configs/emt.yml`.
- **Rscript** is found automatically (macOS: `/usr/local/bin`, Homebrew, R.framework;
  Windows: `Program Files\R` and per-user installs). Otherwise set it in the launcher.
  Quarto is the copy `uv sync` installed into `.venv/` (see §2).

The launcher calls `render-pipeline.R`, which also works on its own:

```bash
cd R_reports
Rscript render-pipeline.R --timeforce-dir ../sample_data/timeforce \
  --plate-map "../sample_data/R672C Plate Map.csv" --analyzed-dir ../sample_data/analyzed \
  --reports-dir ../sample_data/reports --name "R672C Batch 3" --plates --emt-batch \
  [--dataset-config configs/mydataset.yml --batch-number 3]
```
