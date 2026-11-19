# EMT Pipeline

Analysis and reporting for EMT (engineered muscle tissue) force recordings from
24-well plates. One click in the **Master Analysis Launcher** runs the Python
analyses on a batch's raw time-force files and then renders HTML reports from the
raw and analyzed data.

```
raw time-force CSVs + plate map
        │
        ├─► Python analyses (75Hz, Twitch, Force-Frequency, Fatigue) ─► analyzed/*.xlsx
        │                                                                   │
        └─► R: plate trace report (raw QC)                                  ├─► R: single-batch EMT report
                                                                            └─► R: multi-batch dataset report (optional)
```

## Folder layout

| Path | What it is |
|---|---|
| `Python_analyze/Master_Analysis_Launcher.py` | **Start here.** GUI that runs the analyses and the reports. |
| `Python_analyze/*_Script.py` | One analysis per protocol (75Hz, Twitch, Force_Frequency, Fatigue). |
| `R_reports/` | Quarto/R reports. `MANUAL.md` there covers them in detail. |
| `R_reports/render-pipeline.R` | The single entry point the launcher calls to render reports. |
| `R_reports/configs/*.yml` | Your multi-batch dataset configs. Personal, so git-ignored. Templates are in `configs/examples/`. |
| `R_reports/configs/single-batch-defaults.yml` | Panel settings for single-batch reports. |
| `R_reports/configs/examples/` | Templates to copy: a dataset config, a batch CSV, a plate config for the sample data. |
| `setup.sh`, `setup.bat` | One-time setup (Python env, Quarto and R packages) for macOS / Windows. |
| `launch_pipeline.sh`, `launch_pipeline.bat`, `EMT Pipeline.app` | Start the launcher in the project's uv env. |
| `pyproject.toml`, `uv.lock`, `.python-version` | Python version and dependencies, managed by uv. |
| `R_reports/renv.lock`, `renv/`, `.Rprofile` | Pinned R package versions, managed by renv. |
| `R_reports/EMT-Reports.Rproj` | RStudio project for working on the reports. |
| `assets/` | App icon (`emt_icon.svg` / `.png`) and `make_app.sh`, which builds the app. |
| `sample_data/` | One batch of example data (R672C Batch 3, days 3–7). The raw `timeforce/` files aren't in git (see Setup). |

## Setup

You need **uv** (manages Python and Quarto) and **R 4.5** (for the reports). Everything
else is installed per project by the setup script, so it won't touch other Python or R
projects on your machine.

1. **Install uv**
   - macOS: `curl -LsSf https://astral.sh/uv/install.sh | sh`
   - Windows (PowerShell): `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
2. **Install R 4.5** from CRAN: [macOS](https://cran.r-project.org/bin/macosx/) or
   [Windows](https://cran.r-project.org/bin/windows/base/). RStudio is optional. R is only
   needed for the reports; the analyses run without it.
3. **Get the code:** `git clone <repo URL>`, or in GitHub Desktop use *File → Clone
   repository*. Put it somewhere you'll keep it; the macOS app remembers the location.
4. **Run the setup script** from the project folder: `./setup.sh` (macOS) or
   double-click `setup.bat` (Windows). It runs `uv sync`, which installs Python 3.12,
   pandas, numpy, openpyxl, tqdm and Quarto into `.venv/`. Then it runs
   `renv::restore()` in `R_reports/`, which installs the pinned R packages into
   `R_reports/renv/library/`. The first run downloads about 1 GB and takes a few minutes.
5. **Sample data (optional, recommended for a first test run):** the raw recordings
   are too big for git. Copy the `timeforce/` folder from <SAMPLE DATA LINK> into
   `sample_data/`.
6. **Your dataset configs (optional):** to use *Add to dataset report*, copy
   `R_reports/configs/examples/dataset-example.yml` into `R_reports/configs/` and edit
   it ([MANUAL §4](R_reports/MANUAL.md)). These configs hold paths on your machine, so
   they stay out of git.

**Updating:** `git pull`. If `uv.lock` or `R_reports/renv.lock` changed, run the setup
script again. (`uv run`, used by the launch scripts, picks up Python changes on its own.)

## Quick start

Start the launcher any of these ways:

- **macOS app:** double-click `EMT Pipeline.app` (setup offers to build it, or run
  `assets/make_app.sh`). Drag it to the Dock for one-click access.
- **Windows:** double-click `launch_pipeline.bat`. For a desktop icon, right-click it →
  *Send to* → *Desktop (create shortcut)*.
- **From a terminal:** `./launch_pipeline.sh`, or on any OS
  `uv run python Python_analyze/Master_Analysis_Launcher.py`.

Then:

1. **Data folder:** the folder holding the raw time-force files, e.g.
   `sample_data/timeforce`.
2. **Plate map file:** the batch's plate map CSV, e.g. `sample_data/R672C Plate Map.csv`.
3. Tick the analyses and reports you want (all are on by default, except the
   dataset report).
4. Click **Run**. Progress, including R's output, streams into the run log. A
   full run on the sample data takes about 3 minutes.

## Input data

**Time-force files** go directly in the data folder, one per day × protocol, named
`d<day>_<anything>_<Protocol>.csv` (or `.xlsx`). The last `_`-separated word
picks the analysis:

| File suffix | Analysis | Output subfolder |
|---|---|---|
| `_75Hz` | 75Hz kinetics | `analyzed/75Hz/` |
| `_1Hz`, `_02Hz` | Twitch | `analyzed/Twitch/` |
| `_FF`, `_FvF` | Force-frequency | `analyzed/FF/` |
| `_Fatigue` | Fatigue | `analyzed/Fatigue/` |

Example: `d4_R672C-Batch-3_FF.csv` is day 4, force-frequency. Files with other
suffixes are skipped and listed in the run summary.

**Plate map:** a CSV in plate layout, where the first column is the row letter,
the header is the column numbers, and each cell is that well's condition label:

```
,1,2,3,4,5,6
A,UC2WT,UC2WT,R672CHet,R672CHet,R672CHom,R672CHom
B,...
```

## What you get

Outputs go next to the data folder, in the batch folder:

```
<batch>/
├── timeforce/                      # your raw data (the selected data folder)
├── analyzed/
│   ├── 75Hz/  Twitch/  FF/  Fatigue/   # d<day>_..._*.xlsx workbooks
│   ├── analysis_run_manifest.csv       # what has been analyzed (for incremental runs)
│   └── analysis_run_summary_<time>.txt # per-run log: files, timings, errors, report paths
└── reports/
    ├── plates/   <name>-plate-traces.html   # raw traces in plate layout, per protocol and day
    │             plate-traces/<protocol>/*.png, animations/*.gif
    ├── emt/      <name>-emt-report.html     # this batch's figures, metrics and stats
    │             panels/, all-metrics/, stats/*.csv
    └── _configs/                            # configs generated for these reports
```

`<name>` is the batch folder name (with `sample_data` that's `sample-data`). Every
HTML report is self-contained, so you can share it as a single file.

## Launcher options

**Performance & incremental processing**
- *Run analysis types in parallel* — runs the analyses in separate processes,
  sharing the worker budget.
- *Process only new/changed files* — skips files whose input, plate map and
  script are unchanged since the last successful run (tracked in the manifest).
  *Force reprocess all* ignores the manifest.

**Reports (R)**
- *Plate trace report* — every well's raw trace in plate layout, for QC before
  trusting the analysis.
- *Single-batch EMT report* — figures, superplots and stats for this batch.
  Conditions come from the plate-map labels, and the first label (in plate order)
  is the reference. Stats treat **wells** as replicates, because one batch can't
  give batch-level p-values. Treat them as exploratory.
- *Add to dataset report* — adds this batch, as batch #N, to the selected dataset
  config's batch CSV (backing it up first as `<csv>.bak-<timestamp>`) and re-renders
  that multi-batch report. Output goes to the config's `output_dir`, e.g.
  `R_reports/output/mydataset/`. When a dataset is chosen, its conditions and colors
  are also used for the single-batch and plate reports.
- *Rscript* — found automatically; set it here if it isn't.
- *Open reports folder when done.*

Reports render even when there are no new files to analyze, so you can re-run
the reports alone after editing a config.

## How metrics are summarized

- The EMT reports read the `Relaxation Summary` (75Hz, FF) and `Fatigue Summary`
  sheets. Every numeric column is a metric.
- **One value per well:** 75Hz has three tetani per well, and every metric is
  averaged across them (by day, and by Hz for FF). The pre-computed `... Well Avg`
  columns are ignored, so configs use the per-stim names such as `R50 Time`.
- Twitch results are not used by any report yet.

## Running reports without the launcher

From `R_reports/`. The R scripts use the Quarto in `../.venv` automatically. On Windows,
if `Rscript` isn't on your PATH, use its full path, e.g.
`& "C:\Program Files\R\R-4.5.3\bin\Rscript.exe" render-plates.R ../sample_data` in PowerShell.

```bash
# everything the launcher does, for one batch
Rscript render-pipeline.R --timeforce-dir ../sample_data/timeforce \
  --plate-map "../sample_data/R672C Plate Map.csv" --analyzed-dir ../sample_data/analyzed \
  --reports-dir ../sample_data/reports --name "R672C Batch 3" --plates --emt-batch \
  [--dataset-config configs/mydataset.yml --batch-number 3]

Rscript render-plates.R ../sample_data                 # plate trace report only
Rscript render-report.R configs/mydataset.yml          # one multi-batch dataset report
Rscript render-report.R ../sample_data/reports/_configs/emt.yml   # re-render a single-batch report
```

See [`R_reports/MANUAL.md`](R_reports/MANUAL.md) for config options, the batch
CSV format, interactive exploration in R, and how to read the reports.

## The macOS app

`EMT Pipeline.app` is built by `assets/make_app.sh`. It is a tiny AppleScript app
that runs `launch_pipeline.sh` from this project folder.

- **First launch:** macOS asks whether *EMT Pipeline* may access files in your
  Desktop folder. Click **Allow**. If you clicked *Don't Allow*, turn it on in
  System Settings → Privacy & Security → Files and Folders.
- **Moved the project folder?** Run `assets/make_app.sh` again. The app stores the
  project's full path. You can put the app itself anywhere (Dock, /Applications).
  After a rebuild, macOS may ask for Desktop access again.
- **Logs:** each launch appends the launcher's output to
  `~/Library/Logs/EMT_Pipeline.log`. If the launcher doesn't come up, the app
  shows the last lines of that log.
- **Icon:** `assets/emt_icon.svg` shows a 24-well plate, a twitch/tetanus force
  trace and the stimulation bolt. To change it, edit the SVG, delete
  `assets/emt_icon.png` and rerun `assets/make_app.sh`; this needs the R packages
  from setup. The running launcher also uses `emt_icon.png` as its Dock and window
  icon.

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` when analyses start | Start the launcher with a launch script or `uv run`, not a system `python`. |
| `there is no package called ...` in the run log | The R packages aren't installed. Run the setup script again (it runs `renv::restore()`). |
| "Rscript not found" | Set the Rscript path in the Reports section (e.g. `/usr/local/bin/Rscript`, or `C:\Program Files\R\R-4.5.3\bin\Rscript.exe`). |
| `Quarto not found` in the log | Run `uv sync` in the project folder (installs Quarto into `.venv/`), or install Quarto from quarto.org, or set `QUARTO_PATH`. |
| renv says the project is out of sync | Someone changed the lockfile. Run the setup script, or `renv::restore()` from `R_reports/`. |
| A report step shows FAILED | The run log shows R's error just above the result line, and the run summary lists it too. |
| An analysis shows 0 files | Check the file suffix (table above) and that files are directly in the data folder, not a subfolder. |
| Old figures still in a report folder | Re-rendering overwrites files but doesn't delete old ones; clear `panels/` or `all-metrics/` if metrics were renamed. |
| Unexpected extra conditions in a report | A plate-map or workbook label isn't mapped in the dataset config's `conditions` (MANUAL §4.1). |

## For maintainers

- **Python packages:** `uv add <pkg>` (or edit `pyproject.toml`, then `uv lock`).
  The analysis libraries are pinned to the minor versions they were validated with. To
  upgrade one, run `uv lock --upgrade-package pandas` and compare the sample-data
  outputs before committing.
- **R packages:** install inside the project (`cd R_reports`, start R, then
  `install.packages("pkg")`). Use it in the code, then run `renv::snapshot()` and
  commit `renv.lock`.
- **Why the R repo is a dated snapshot:** `renv.lock` installs from Posit Package
  Manager's snapshot of CRAN for 2026-11-17. A dated snapshot keeps serving prebuilt
  macOS and Windows binaries of exactly those versions, so labmates never compile
  packages. (CRAN, and PPM's `latest`, only have binaries of each package's current
  version.) To move everything forward, in `R_reports/` run
  `options(repos = c(CRAN = "https://packagemanager.posit.co/cran/<YYYY-MM-DD>"))`,
  then `renv::update()` and `renv::snapshot()`, and put the new date in
  `R_reports/.Rprofile` (which sets the repo for every R session started in that
  folder). Re-render the sample reports before committing.
- Commit `uv.lock` and `R_reports/renv.lock` whenever they change.
