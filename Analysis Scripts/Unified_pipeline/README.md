# EMT Pipeline

Analysis and reporting for EMT (engineered muscle tissue) force recordings from
24-well plates. The **Master Analysis Launcher** works one batch at a time: one
click runs the Python analyses on a batch's raw time-force files, then renders that
batch's plate trace report and EMT report. Batches can also be registered in a
dataset, whose combined reports are rendered separately from the launcher's
*Dataset reports* section.

```
Run batch (once per batch)
  raw time-force files + plate map
        │
        ├─► Python analyses (75Hz, Twitch, Force-Frequency, Fatigue) ─► analyzed/*.xlsx
        │                                                                   │
        └─► R: plate trace report (raw QC)                                  └─► R: single-batch EMT report
        (optional: register the batch's folders in a dataset's batch CSV)

Run dataset reports (whenever you like)
  every batch registered in a dataset ─► R: dataset EMT report + combined plate trace report
```

## Folder layout

| Path | What it is |
|---|---|
| `Python_analyze/Master_Analysis_Launcher.py` | **Start here.** GUI with two sections: run one batch (analyses + its reports), and render a dataset's reports across batches. |
| `Python_analyze/*_Script.py` | One analysis per protocol (75Hz, Twitch, Force_Frequency, Fatigue). |
| `R_reports/` | Quarto/R reports. `MANUAL.md` there covers them in detail. |
| `R_reports/render-pipeline.R` | The single entry point the launcher calls to render reports. |
| `<experiment folder>/<dataset>.yml`, `<dataset>-batches.csv` | A dataset config and its batch CSV, kept with the data (made by the launcher's *New…* / *Register*, or by hand). Older configs in `R_reports/configs/` still work; those are git-ignored. Templates are in `R_reports/configs/examples/`. |
| `R_reports/output/` | Dataset reports of configs kept in `R_reports/configs/`. Git-ignored. |
| `R_reports/configs/single-batch-defaults.yml` | Panel settings for single-batch reports. |
| `R_reports/configs/examples/` | Templates to copy: a dataset config, a batch CSV, plate configs (one batch, or several combined). |
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
6. **Your dataset configs (optional):** to start a dataset, click **New…** next to
   *Register in dataset* in the launcher, or copy
   `R_reports/configs/examples/dataset-example.yml` into your experiment folder and edit
   it ([MANUAL §4](R_reports/MANUAL.md)). Then add it in the launcher with *Open…*.

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

Then, in the **Batch** section:

1. **Data folder:** the folder holding the batch's raw time-force files, e.g.
   `sample_data/timeforce`.
2. **Plate map file:** the batch's plate map CSV, e.g. `sample_data/R672C Plate Map.csv`.
3. Tick the analyses and reports you want (all are on by default). To add the batch
   to a dataset, tick *Register in dataset* and pick the dataset, or click **New…** to
   start one; the batch number is filled in for you.
4. Click **Run batch**. Progress, including R's output, streams into the run log. A
   full run on the sample data takes about 3 minutes.

For reports across batches, pick the dataset in the **Dataset reports** section and
click **Run dataset reports**.

## Input data

**Time-force files** go in the data folder, either directly or in subfolders (for
example one per protocol), one file per day × protocol. The file name gives the day
(`d<day>_`) and, as its last `_`-separated word, the protocol, which picks the
analysis:

| File suffix | Analysis | Output subfolder |
|---|---|---|
| `_75Hz` | 75Hz kinetics | `analyzed/75Hz/` |
| `_1Hz`, `_02Hz` | Twitch | `analyzed/Twitch/` |
| `_FF`, `_FvF` | Force-frequency | `analyzed/FF/` |
| `_Fatigue` | Fatigue | `analyzed/Fatigue/` |

Examples: `d4_R672C-Batch-3_FF.csv` is day 4, force-frequency. Pulse3D exports
work too: in `ML25101402-2__2025_05_20_230523_d4_DAEMTR672C_75Hz__full__2025-05-23_15-02-14.xlsx`
the `__full__…` ending is ignored, so it's day 4, 75Hz.

- Files with other protocols (e.g. ASFatigue, Priming) are skipped and listed in the
  run summary.
- The **file name, not the folder, picks the analysis.** A 75Hz file sitting in an
  `FF/` folder is analyzed as 75Hz, and the run log, run summary and end-of-run
  message warn about it so you can check it.
- If one day × protocol has several files, the newest is used and the others are
  listed in the run summary.

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

For a batch run on several plates (one folder per plate, see *plate* under Launcher
options), `analyzed/` and `reports/` sit next to the plate folders, with one subfolder
per plate:

```
<batch>/
├── dmd/  wt/                        # plate folders: plate map + timeforce files
├── analyzed/
│   ├── dmd/  75Hz/ Twitch/ FF/ Fatigue/   # d<day>_..._DMD.xlsx
│   └── wt/   ...                          # d<day>_..._WT.xlsx
└── reports/
    ├── dmd/plates/  wt/plates/      # each plate's plate trace report
    └── emt/                         # the batch's single EMT report, all plates
```

Dataset reports go to the dataset config's `output_dir`. For a config made with
*New…* that is `reports/<dataset>/` next to the config, e.g.
`SR4370_exp/reports/srexpt/`: the EMT report at the top, the combined plate trace
report in `plates/`, and the generated plate config in `_configs/`.

## Launcher options

**Performance & incremental processing**
- *Run analysis types in parallel* — runs the analyses at the same time,
  splitting the worker budget evenly between them.
- *Total worker budget* — how many worker processes to use in total (defaults to
  the CPU core count). Each analysis processes its files, then writes its
  per-day workbooks, across its share of the workers. An analysis can't use more
  workers than it has files (or days), so small batches won't load every core.
  Each worker holds its own copy of the data it is working on (a full batch on
  18 workers peaked around 19 GB), so lower this if the machine runs short of
  RAM. The plate trace reports also read files and draw plots across this many R
  worker processes (at most 8; see MANUAL §10.8).
- *Process only new/changed files* — skips files whose input, plate map and
  script are unchanged since the last successful run (tracked in the manifest).
  *Force reprocess all* ignores the manifest.

**Batch section: reports**
- *Plate trace report* — every well's raw trace in plate layout, for QC before
  trusting the analysis.
- *Single-batch EMT report* — figures, superplots and stats for this batch, for
  every recorded day. Conditions come from the plate-map labels, and the first label
  (in plate order) is the reference. Stats treat **wells** as replicates, because one
  batch can't give batch-level p-values. Treat them as exploratory.
- *Register in dataset … as batch #N* — adds this batch's analyzed folders and its raw
  data folder + plate map to the dataset config's batch CSV (backing it up first as
  `<csv>.bak-<timestamp>`). It doesn't render the dataset reports. The batch number
  is suggested: this batch's existing number if it's already registered, otherwise the
  next free one. When registering, the dataset's conditions and colors are also used
  for this batch's own reports.
- *New…* — creates a dataset config, `<name>.yml`, in a folder you pick (it suggests
  the experiment folder, the one holding `analyzed/`). Its batch CSV is kept next to it,
  its reports go to `reports/<name>/` beside it, and its conditions and colors are
  filled in from the selected plate map. Check those afterwards (MANUAL §4.1), especially when batches spell a condition
  differently.
- *plate* — only for a batch run on several plates (e.g. genotypes split across
  plates). Give each plate its own folder (plate map + timeforce files), pick it as
  the data folder and run each plate with the same batch # and its own plate name
  (e.g. DMD, WT). Results go to `analyzed/<plate folder>/` next to the plate folders,
  with file names ending in `_<plate>`. Each plate gets its own plate trace report in
  `reports/<plate folder>/`. The EMT report covers every plate registered for the
  batch so far and goes to `<batch folder>/reports/emt/`, where the batch folder
  holds the plate folders. Leave it blank for single-plate batches.

Batch reports render even when there are no new files to analyze, so you can re-run
the reports alone after editing a config.

**Dataset reports section**
- *Dataset* — a dataset config. The list holds configs made with *New…* or added with
  *Open…*, configs found in a data folder's experiment folder when you pick that data
  folder, and older ones in `R_reports/configs/`. The table shows its registered
  batches: their folder, which analyzed data and raw data each has, and anything
  excluded in the batch CSV. *Refresh* re-reads the configs and CSV.
- *Dataset EMT report* — the multi-batch EMT report, with batch-level stats
  (MANUAL §4). Output goes to the config's `output_dir`, e.g.
  `<experiment folder>/reports/mydataset/`.
- *Combined plate trace report* — every registered batch's raw traces in one report,
  one tab per batch within each protocol (MANUAL §10.7). Output goes to
  `<output_dir>/plates/`.
- **Run dataset reports** renders the ticked reports for every batch in the CSV.
  It doesn't run any analyses.

**Shared**
- *Rscript* — found automatically; set it here if it isn't.
- *Open reports folder when done.*
- Only one run at a time: both Run buttons are disabled while either is running.

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
# what "Run batch" renders, for one batch (the analyses run in Python first)
Rscript render-pipeline.R --timeforce-dir ../sample_data/timeforce \
  --plate-map "../sample_data/R672C Plate Map.csv" --analyzed-dir ../sample_data/analyzed \
  --reports-dir ../sample_data/reports --name "R672C Batch 3" --plates --emt-batch \
  [--dataset-config configs/mydataset.yml --register --batch-number 3]

# what "Run dataset reports" does
Rscript render-pipeline.R --dataset-config configs/mydataset.yml --dataset-emt --dataset-plates

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
| An analysis shows 0 files | Check the file names: they need `d<day>_` and the protocol as the last word (table above). Unmatched files are listed in the run summary. |
| "in a folder for a different protocol" warning | A file's name and its folder name disagree (e.g. a 75Hz file in `FF/`). It was analyzed by its name; move or rename it if the name is wrong. |
| Dataset table is empty, or *Combined plate trace report* fails with "No `timeforce` rows" | Register the batches from the Batch section (re-running a batch with only *Register* ticked is enough), or add rows to the batch CSV by hand (MANUAL §3.2). |
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
