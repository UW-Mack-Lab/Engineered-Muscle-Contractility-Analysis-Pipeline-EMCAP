import os
import sys
import csv
import glob
import re
import queue
import shutil
import subprocess
import threading
import traceback
import importlib.util
from dataclasses import dataclass
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

from file_pool import WORKER_CRASHED
from filenames import protocol_from_filename


DATA_EXTENSIONS = {".csv", ".xlsx"}
MAX_PARALLEL_ANALYSES = 4
MANIFEST_FILENAME = "analysis_run_manifest.csv"

# R reports live in a sibling folder of this script's folder (override with EMT_R_REPORTS_DIR)
R_REPORTS_DIR = os.environ.get(
    "EMT_R_REPORTS_DIR",
    os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "R_reports")),
)
R_PIPELINE_SCRIPT = "render-pipeline.R"
R_RESULT_PREFIX = "PIPELINE_RESULT "
SINGLE_BATCH_DEFAULTS = "single-batch-defaults.yml"
ANSI_ESCAPE = re.compile(r"\x1b\[[0-9;]*[A-Za-z]")


@dataclass(frozen=True)
class AnalysisDefinition:
    name: str
    suffix_tokens: set
    script_filename: str
    function_name: str
    output_subfolder: str


ANALYSES = {
    "75Hz": AnalysisDefinition(
        name="75Hz",
        suffix_tokens={"75hz"},
        script_filename="75Hz_Script.py",
        function_name="run_75hz_analysis",
        output_subfolder="75Hz",
    ),
    "Twitch": AnalysisDefinition(
        name="Twitch",
        suffix_tokens={"1hz", "02hz"},
        script_filename="Twitch_Script.py",
        function_name="run_twitch_analysis",
        output_subfolder="Twitch",
    ),
    "ForceFrequency": AnalysisDefinition(
        name="Force Frequency",
        suffix_tokens={"ff", "fvf"},
        script_filename="Force_Frequency_Script.py",
        function_name="run_force_frequency_analysis",
        output_subfolder="FF",
    ),
    "Fatigue": AnalysisDefinition(
        name="Fatigue",
        suffix_tokens={"fatigue"},
        script_filename="Fatigue_Script.py",
        function_name="run_fatigue_analysis",
        output_subfolder="Fatigue",
    ),
}


MANIFEST_COLUMNS = [
    "timestamp",
    "analysis_key",
    "analysis_name",
    "input_path",
    "input_size",
    "input_mtime_ns",
    "plate_map_path",
    "plate_map_mtime_ns",
    "script_path",
    "script_mtime_ns",
    "output_dir",
    "status",
    "error",
]


def _now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _is_temp_or_hidden(name: str) -> bool:
    return name.startswith(".") or name.startswith("~$")


DAY_PATTERN = re.compile(r"(?:^|_)[dD](\d+)_")

# Protocols recorded but not analyzed. They only matter for spotting misfiled
# files (a folder named after one of these holding another protocol's file).
OTHER_PROTOCOL_TOKENS = {"asfatigue", "priming"}

# Output folders the pipeline writes next to (or, if the batch folder itself is
# picked as the data folder, inside) the data folder. Never searched for inputs.
OUTPUT_FOLDER_NAMES = {"analyzed", "reports"}


def _extract_day(filename: str):
    """Recording day from 'd<day>_' at the start or after an underscore, else None."""
    m = DAY_PATTERN.search(filename)
    return int(m.group(1)) if m else None


def _norm_protocol(token: str) -> str:
    """Compare protocol names ignoring case, '-', '_' and spaces (75HZ == 75Hz)."""
    return re.sub(r"[-_ ]", "", token).lower()


def _load_callable(script_path: str, function_name: str):
    if not os.path.isfile(script_path):
        raise FileNotFoundError(f"Script not found: {script_path}")

    module_name = f"analysis_module_{os.path.basename(script_path).replace('.', '_')}_{_now_stamp()}"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load module spec from {script_path}")

    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    if not hasattr(mod, function_name):
        raise AttributeError(f"{function_name} not found in {script_path}")

    fn = getattr(mod, function_name)
    if not callable(fn):
        raise TypeError(f"{function_name} exists but is not callable in {script_path}")

    return fn


def _run_analysis_job(payload: dict):
    """
    Process-safe worker entrypoint.
    Loads target script in child process and executes the run function.
    """
    key = payload["key"]
    definition_name = payload["definition_name"]
    script_path = payload["script_path"]
    function_name = payload["function_name"]
    files = payload["files"]
    plate_map_path = payload["plate_map_path"]
    output_dir = payload["output_dir"]
    file_workers = payload["file_workers"]
    file_tag = payload.get("file_tag", "")

    started = datetime.now()
    try:
        run_fn = _load_callable(script_path, function_name)
        run_fn(
            input_files=files,
            plate_map_path=plate_map_path,
            output_dir=output_dir,
            max_workers=file_workers,
            file_tag=file_tag,
        )
        ended = datetime.now()
        return {
            "key": key,
            "name": definition_name,
            "status": "DONE",
            "duration_seconds": round((ended - started).total_seconds(), 2),
            "error": None,
            "traceback": None,
        }
    except Exception as exc:
        ended = datetime.now()
        return {
            "key": key,
            "name": definition_name,
            "status": "FAILED",
            "duration_seconds": round((ended - started).total_seconds(), 2),
            "error": str(exc),
            "traceback": traceback.format_exc(),
        }


def _file_sig(path: str):
    st = os.stat(path)
    return str(os.path.abspath(path)), int(st.st_size), int(st.st_mtime_ns)


def _dependency_sig(plate_map_path: str, script_path: str):
    plate_st = os.stat(plate_map_path)
    script_st = os.stat(script_path)
    return (
        str(os.path.abspath(plate_map_path)),
        int(plate_st.st_mtime_ns),
        str(os.path.abspath(script_path)),
        int(script_st.st_mtime_ns),
    )


def _signature_key(analysis_key: str, input_sig: tuple, dep_sig: tuple):
    input_path, input_size, input_mtime_ns = input_sig
    plate_map_path, plate_map_mtime_ns, script_path, script_mtime_ns = dep_sig
    return (
        analysis_key,
        input_path,
        str(input_size),
        str(input_mtime_ns),
        plate_map_path,
        str(plate_map_mtime_ns),
        script_path,
        str(script_mtime_ns),
    )


def _manifest_path(analyzed_root: str):
    return os.path.join(analyzed_root, MANIFEST_FILENAME)


def _load_done_manifest_keys(manifest_path: str):
    keys = set()
    if not os.path.isfile(manifest_path):
        return keys

    with open(manifest_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") != "DONE":
                continue
            key = (
                row.get("analysis_key", ""),
                row.get("input_path", ""),
                row.get("input_size", ""),
                row.get("input_mtime_ns", ""),
                row.get("plate_map_path", ""),
                row.get("plate_map_mtime_ns", ""),
                row.get("script_path", ""),
                row.get("script_mtime_ns", ""),
            )
            keys.add(key)
    return keys


def _append_manifest_rows(manifest_path: str, rows):
    write_header = not os.path.isfile(manifest_path)
    with open(manifest_path, "a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS)
        if write_header:
            writer.writeheader()
        for row in rows:
            writer.writerow(row)


def _output_dir_has_files(path: str):
    if not os.path.isdir(path):
        return False
    for name in os.listdir(path):
        full = os.path.join(path, name)
        if os.path.isfile(full) and not _is_temp_or_hidden(name):
            return True
    return False


@dataclass
class Discovery:
    """Result of discover_files(). File names in the lists are relative to the
    data folder; `routed` holds absolute paths."""
    routed: dict          # analysis key -> [path, ...]
    unmatched: list       # data files whose protocol isn't analyzed (ASFatigue, Priming, ...)
    ignored: list         # hidden/temp/non-data files
    ambiguous: list       # (file, [analysis keys]) matching more than one analysis
    duplicates: list      # (older file, newer file used instead) for the same day x protocol
    misfiled: list        # (file, folder protocol, file protocol)


def discover_files(data_dir: str) -> Discovery:
    """
    Time-force files anywhere under data_dir (protocol subfolders are fine), routed
    to an analysis by the protocol at the end of the *file name*, not the folder.
    Reads plain exports (d4_R672C-Batch-3_FF.csv) and Pulse3D names
    (ML…__2025_05_20_230523_d4_DAEMTR672C_75Hz__full__<export time>.xlsx).
    When a day x protocol has several files, the newest is used.
    """
    known_tokens = {_norm_protocol(t) for d in ANALYSES.values() for t in d.suffix_tokens}
    known_tokens |= OTHER_PROTOCOL_TOKENS

    found = {}  # (analysis key, day or path, token) -> [paths]
    unmatched, ignored, ambiguous, misfiled = [], [], [], []

    for root, dirs, files in os.walk(data_dir):
        dirs[:] = sorted(d for d in dirs
                         if not d.startswith(".") and d.lower() not in OUTPUT_FOLDER_NAMES)
        for entry in sorted(files):
            full_path = os.path.join(root, entry)
            rel = os.path.relpath(full_path, data_dir)

            if _is_temp_or_hidden(entry):
                ignored.append(rel)
                continue
            if os.path.splitext(entry)[1].lower() not in DATA_EXTENSIONS:
                ignored.append(rel)
                continue

            token = protocol_from_filename(entry)
            if not token:
                unmatched.append(rel)
                continue

            token_n = _norm_protocol(token)
            matches = [
                key for key, definition in ANALYSES.items()
                if token_n in {_norm_protocol(t) for t in definition.suffix_tokens}
            ]
            if len(matches) == 0:
                unmatched.append(rel)
                continue
            if len(matches) > 1:
                ambiguous.append((rel, matches))
                continue

            # A protocol-named folder holding another protocol's file: analyzed by
            # its file name, but flagged so a misfiled recording doesn't go unnoticed.
            folder = os.path.basename(root) if os.path.abspath(root) != os.path.abspath(data_dir) else ""
            folder_n = _norm_protocol(folder)
            if folder_n in known_tokens and folder_n != token_n:
                misfiled.append((rel, folder, token))

            day = _extract_day(entry)
            group = (matches[0], day if day is not None else full_path, token_n)
            found.setdefault(group, []).append(full_path)

    routed = {key: [] for key in ANALYSES.keys()}
    duplicates = []
    for (key, _, _), paths in sorted(found.items(), key=lambda kv: str(kv[0])):
        paths.sort(key=os.path.getmtime, reverse=True)
        routed[key].append(paths[0])
        for older in paths[1:]:
            duplicates.append((os.path.relpath(older, data_dir), os.path.relpath(paths[0], data_dir)))
    for key in routed:
        routed[key].sort()

    return Discovery(routed, unmatched, ignored, ambiguous, duplicates, misfiled)


ANALYZED_TYPE_DIRS = {d.output_subfolder.lower() for d in ANALYSES.values()}


def _plate_folder(data_dir: str) -> str:
    """The folder a plate's data lives in: the data folder, or its parent when the
    data folder is just 'timeforce' (Expt/dmd/timeforce -> Expt/dmd)."""
    folder = os.path.abspath(data_dir)
    if os.path.basename(folder).lower() == "timeforce":
        folder = os.path.dirname(folder)
    return folder


def output_roots(data_dir: str, plate: str = ""):
    """(analyzed root, reports root) for a data folder.
    No plate:   next to the data folder, Batch/timeforce -> Batch/analyzed, Batch/reports.
    With plate: one batch run on several plates, each in its own folder. Outputs go
                next to the plate folders, under the plate folder's name:
                Expt/dmd (or Expt/dmd/timeforce) -> Expt/analyzed/dmd, Expt/reports/dmd."""
    if plate:
        folder = _plate_folder(data_dir)
        parent, name = os.path.dirname(folder), os.path.basename(folder)
        return os.path.join(parent, "analyzed", name), os.path.join(parent, "reports", name)
    parent = os.path.dirname(os.path.abspath(data_dir))
    return os.path.join(parent, "analyzed"), os.path.join(parent, "reports")


def file_tag_for(plate: str) -> str:
    """The plate name as it is appended to output file names ('' for none)."""
    return re.sub(r'[\\/:*?"<>|\s]+', "-", plate or "").strip("-")


def foreign_inputs(analyzed_root: str, data_dir: str):
    """Inputs in analyzed_root's manifest that came from outside this plate's
    folder: a sign that another plate (or batch) already writes there."""
    manifest = _manifest_path(analyzed_root)
    if not os.path.isfile(manifest):
        return []
    own = _norm_dir(_plate_folder(data_dir))
    out = set()
    with open(manifest, encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            path = _norm_dir(row.get("input_path", ""))
            if path and not path.startswith(own + os.sep):
                out.add(row["input_path"])
    return sorted(out)


def build_output_dirs(data_dir: str, selected_keys, plate: str = ""):
    analyzed_root, _ = output_roots(data_dir, plate)
    os.makedirs(analyzed_root, exist_ok=True)

    output_map = {}
    for key in selected_keys:
        out_sub = ANALYSES[key].output_subfolder
        out_dir = os.path.join(analyzed_root, out_sub)
        os.makedirs(out_dir, exist_ok=True)
        output_map[key] = out_dir

    return analyzed_root, output_map


def write_summary_log(analyzed_root: str, lines):
    log_path = os.path.join(analyzed_root, f"analysis_run_summary_{_now_stamp()}.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return log_path


# R reports -------------------------------------------------------------------

def build_reports_dir(data_dir: str, plate: str = ""):
    """reports/ next to analyzed/ (see output_roots)."""
    _, reports_root = output_roots(data_dir, plate)
    os.makedirs(reports_root, exist_ok=True)
    return reports_root


def _batch_name(data_dir: str) -> str:
    """Name reports after the batch folder: the data folder, or its parent when
    the data folder is just 'timeforce' (matches the R plate-report naming)."""
    folder = os.path.abspath(data_dir)
    name = os.path.basename(folder)
    if name.lower() == "timeforce":
        name = os.path.basename(os.path.dirname(folder)) or name
    return name


def _r_version_key(path):
    """Sort key for ...\\R-4.5.3\\bin\\Rscript.exe paths by version number."""
    m = re.search(r"R-(\d+)\.(\d+)\.(\d+)", path)
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)


def _find_quarto():
    """Quarto installed into this env by `uv sync` (the quarto-cli package).
    Returns None if it isn't there; R then looks for a standalone install."""
    bin_dir = os.path.dirname(sys.executable)
    names = ["quarto.exe", "quarto.cmd"] if sys.platform == "win32" else ["quarto"]
    for name in names:
        candidate = os.path.join(bin_dir, name)
        if os.path.isfile(candidate):
            return candidate
    return None


def _r_subprocess_env(workers=None):
    """Environment for Rscript: points R at the env's Quarto unless QUARTO_PATH
    is already set, and passes the worker budget to the plate reports
    (EMT_R_WORKERS, see R_reports/R/parallel.R)."""
    env = dict(os.environ)
    if workers:
        env["EMT_R_WORKERS"] = str(int(workers))
    if not env.get("QUARTO_PATH"):
        quarto = _find_quarto()
        if quarto:
            env["QUARTO_PATH"] = quarto
    return env


def _find_rscript():
    """Locate Rscript. Apps launched from Finder get a minimal PATH, so also
    check the usual install locations."""
    candidates = [os.environ.get("RSCRIPT_PATH", ""), shutil.which("Rscript") or ""]
    if sys.platform == "win32":
        # Machine-wide and per-user ("just for me") R installs, newest first
        roots = [r"C:\Program Files\R", os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "R")]
        for root in roots:
            candidates += sorted(glob.glob(os.path.join(root, "R-*", "bin", "Rscript.exe")),
                                 key=_r_version_key, reverse=True)
    else:
        candidates += [
            "/usr/local/bin/Rscript",
            "/opt/homebrew/bin/Rscript",
            "/Library/Frameworks/R.framework/Resources/bin/Rscript",
            "/usr/bin/Rscript",
        ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


# Dataset configs live in the experiment folder (any folder), or in
# R_reports/configs/. Ones outside R_reports/configs/ are remembered here, one
# absolute path per line, so the launcher can list them.
CONFIGS_DIR = os.path.join(R_REPORTS_DIR, "configs")
KNOWN_DATASETS_FILE = os.path.join(CONFIGS_DIR, "known-datasets.txt")
BATCH_CSV_LINE = re.compile(r"""^batch_csv:\s*(?:"([^"]*)"|'([^']*)'|([^#\s][^#]*?))\s*(?:#.*)?$""")


def _same_path(a: str, b: str) -> bool:
    return _norm_dir(a) == _norm_dir(b)


def _config_path(value: str) -> str:
    """Absolute path of a dataset config as listed in the launcher
    ('configs/x.yml' is relative to R_reports/; others are '~/...' or absolute)."""
    v = os.path.expanduser(value.strip())
    return os.path.normpath(v if os.path.isabs(v) else os.path.join(R_REPORTS_DIR, v))


def _config_display(path: str) -> str:
    """How a config is listed: configs/x.yml for R_reports/configs/, else ~/... path."""
    path = os.path.normpath(os.path.abspath(path))
    if _same_path(os.path.dirname(path), CONFIGS_DIR):
        return f"configs/{os.path.basename(path)}"
    home = os.path.expanduser("~")
    if _norm_dir(path).startswith(_norm_dir(home) + os.sep):
        return "~" + path[len(home):]
    return path


def _config_base_dir(path: str) -> str:
    """Folder a config's relative paths start from: its own folder, except for
    configs in R_reports/configs/, which are relative to R_reports/ (same rule as
    config_base_dir() in R_reports/R/emt-config.R)."""
    folder = os.path.dirname(os.path.abspath(path))
    return R_REPORTS_DIR if _same_path(folder, CONFIGS_DIR) else folder


def _batch_csv_value(path: str):
    """The raw top-level `batch_csv:` value of a config, or None. A plain line
    read, so no YAML library needed."""
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                m = BATCH_CSV_LINE.match(line.rstrip("\n"))
                if m:
                    return next(g for g in m.groups() if g is not None).strip() or None
    except (OSError, UnicodeDecodeError):
        return None
    return None


def _dataset_batch_csv(config_value: str):
    """Absolute path of a dataset config's batch CSV, or None."""
    path = _config_path(config_value)
    value = _batch_csv_value(path)
    if not value:
        return None
    value = os.path.expanduser(value)
    return os.path.normpath(value if os.path.isabs(value) else os.path.join(_config_base_dir(path), value))


def _read_known_datasets():
    try:
        with open(KNOWN_DATASETS_FILE, encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]
    except OSError:
        return []


def _remember_dataset(path: str):
    """Add a config outside R_reports/configs/ to the known-datasets list."""
    path = os.path.normpath(os.path.abspath(path))
    if _same_path(os.path.dirname(path), CONFIGS_DIR):
        return
    known = _read_known_datasets()
    if any(_same_path(k, path) for k in known):
        return
    os.makedirs(CONFIGS_DIR, exist_ok=True)
    with open(KNOWN_DATASETS_FILE, "w", encoding="utf-8") as f:
        f.write("".join(f"{k}\n" for k in known + [path]))


def _find_dataset_configs(folder: str):
    """Dataset configs (.yml with a `batch_csv:` line) directly in folder."""
    if not folder or not os.path.isdir(folder):
        return []
    out = []
    for name in sorted(os.listdir(folder)):
        p = os.path.join(folder, name)
        if (name.lower().endswith((".yml", ".yaml")) and not _is_temp_or_hidden(name)
                and os.path.isfile(p) and _batch_csv_value(p)):
            out.append(p)
    return out


def _list_dataset_configs():
    """Dataset configs as listed in the launcher: known experiment-folder configs
    that still exist, then R_reports/configs/*.yml."""
    out = [_config_display(p) for p in _read_known_datasets() if os.path.isfile(p)]
    if os.path.isdir(CONFIGS_DIR):
        for name in sorted(os.listdir(CONFIGS_DIR)):
            if name.lower().endswith((".yml", ".yaml")) and name != SINGLE_BATCH_DEFAULTS:
                out.append(f"configs/{name}")
    return list(dict.fromkeys(out))


def _read_batch_rows(csv_path):
    """Rows of a dataset batch CSV as dicts ([] if it doesn't exist yet)."""
    if not csv_path or not os.path.isfile(csv_path):
        return []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _norm_dir(path: str) -> str:
    return os.path.normcase(os.path.normpath(os.path.abspath(path.strip()))) if path and path.strip() else ""


def _batch_no(row):
    try:
        return int(float(row.get("batch", "")))
    except ValueError:
        return None


def _suggest_registration(rows, data_dir: str):
    """(batch number, plate) for this data folder: its existing registration if its
    raw or analyzed folders are already in the CSV, else (highest number + 1, "")."""
    numbers = [n for n in (_batch_no(r) for r in rows) if n is not None]
    if data_dir and os.path.isdir(data_dir):
        folder = _norm_dir(_plate_folder(data_dir))
        plate_root = _norm_dir(output_roots(data_dir, "plate")[0])   # Expt/analyzed/<plate folder>
        batch_root = _norm_dir(output_roots(data_dir)[0])            # Batch/analyzed
        for r in rows:
            d = _norm_dir(r.get("directory", ""))
            # raw folder (with or without timeforce/), or <analyzed root>/<type>
            if d and (_norm_dir(_plate_folder(d)) == folder
                      or os.path.dirname(d) in (plate_root, batch_root)):
                n = _batch_no(r)
                if n is not None:
                    return n, (r.get("plate") or "").strip()
    return (max(numbers) + 1 if numbers else 1), ""


def _data_folder_of(path: str) -> str:
    """The plate/batch folder an analyzed directory belongs to:
    Batch/analyzed(/75Hz) -> Batch; Expt/analyzed/dmd(/75Hz) -> Expt/dmd."""
    p = os.path.normpath(path)
    if os.path.basename(p).lower() in ANALYZED_TYPE_DIRS:
        p = os.path.dirname(p)
    if os.path.basename(os.path.dirname(p)).lower() == "analyzed":
        return os.path.join(os.path.dirname(os.path.dirname(p)), os.path.basename(p))
    if os.path.basename(p).lower() in {"analyzed", "timeforce"}:
        return os.path.dirname(p)
    return p


def _summarize_batches(rows):
    """One summary per batch number x plate for the dataset table:
    (number, plate, folder, {analysis: True/'excluded'})."""
    by_batch = {}
    for r in rows:
        n = _batch_no(r)
        if n is None:
            continue
        by_batch.setdefault((n, (r.get("plate") or "").strip()), []).append(r)
    out = []
    for n, plate in sorted(by_batch):
        group = by_batch[(n, plate)]
        raw = [r["directory"].strip() for r in group
               if (r.get("analysis") or "").strip().lower() == "timeforce" and (r.get("directory") or "").strip()]
        dirs = [os.path.normpath(r["directory"].strip()) for r in group if (r.get("directory") or "").strip()]
        if raw:
            folder = _plate_folder(raw[0])
        else:
            try:
                folder = os.path.commonpath(dirs) if dirs else ""
            except ValueError:          # different drives on Windows
                folder = dirs[0]
            folder = _data_folder_of(folder) if folder else ""
        present = {}
        for r in by_batch[(n, plate)]:
            excluded = (r.get("exclude") or "0").strip() not in ("", "0")
            present[(r.get("analysis") or "").strip().lower()] = "excluded" if excluded else True
        out.append((n, plate, folder, present))
    return out


# New dataset configs -------------------------------------------------------------

DATASET_TEMPLATE = os.path.join("configs", "examples", "dataset-example.yml")
# Same order as plate_colors() in R_reports/R/plate-data.R
OKABE_ITO = ["#000000", "#009E73", "#E69F00", "#D55E00", "#CC79A7", "#0072B2", "#56B4E9", "#F0E442"]


def _slugify(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def _plate_map_labels(path: str):
    """Condition labels in a plate-map CSV, in plate order (row by row)."""
    labels = []
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            for i, row in enumerate(csv.reader(f)):
                if i == 0 or not row or not row[0].strip():
                    continue
                for cell in row[1:]:
                    cell = cell.strip()
                    if cell and cell not in labels:
                        labels.append(cell)
    except OSError:
        pass
    return labels


def _yaml_str(x: str) -> str:
    return '"' + x.replace("\\", "\\\\").replace('"', '\\"') + '"'


def new_dataset_config_text(name: str, slug: str, labels, template_text: str, in_r_configs: bool = False) -> str:
    """A dataset config from the example template: name, batch CSV and output
    folder set; conditions/colors seeded from plate-map labels (identity mapping).
    The batch CSV sits next to the config and the reports go to reports/<slug>/
    beside it; for a config in R_reports/configs/, paths are relative to R_reports/."""
    text = template_text
    body_start = re.search(r"^[^#\s]", text, flags=re.M)
    if in_r_configs:
        paths_note, batch_csv, output_dir = "R_reports/", f"configs/{slug}-batches.csv", f"output/{slug}"
    else:
        paths_note, batch_csv, output_dir = "this file's folder", f"{slug}-batches.csv", f"reports/{slug}"
    header = (f"# {name} dataset config, created by the Master Analysis Launcher on "
              f"{datetime.now():%Y-%m-%d}.\n"
              "# Edit `conditions` (display name -> every raw plate-map spelling) and\n"
              f"# `colors`; see R_reports/MANUAL.md section 4. Paths are relative to {paths_note}.\n\n")
    text = header + (text[body_start.start():] if body_start else text)
    text = re.sub(r"^name:.*$", lambda _m: f"name: {_yaml_str(name)}", text, count=1, flags=re.M)
    text = re.sub(r"^batch_csv:.*$", lambda _m: f"batch_csv: {batch_csv}", text, count=1, flags=re.M)
    text = re.sub(r"^output_dir:.*$", lambda _m: f"output_dir: {output_dir}", text, count=1, flags=re.M)
    # A new dataset's batch numbers may not start at 1
    text = re.sub(r"^trace_batches:.*$", "trace_batches: all", text, count=1, flags=re.M)
    if labels:
        width = max(len(_yaml_str(l)) for l in labels) + 1
        conds = "".join(f"  {(_yaml_str(l) + ':').ljust(width)} [{_yaml_str(l)}]\n" for l in labels)
        cols = "".join(f"  {(_yaml_str(l) + ':').ljust(width)} \"{OKABE_ITO[i % len(OKABE_ITO)]}\"\n"
                       for i, l in enumerate(labels))
        text = re.sub(r"^conditions:\n(?:[ \t]+.*\n)+", lambda _m: "conditions:\n" + conds, text, count=1, flags=re.M)
        text = re.sub(r"^colors:\n(?:[ \t]+.*\n)+", lambda _m: "colors:\n" + cols, text, count=1, flags=re.M)
    return text


def _open_folder(path: str):
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def _run_r_pipeline(args, log_fn, workers=None):
    """
    Run render-pipeline.R from R_REPORTS_DIR, streaming its output to log_fn.
    Returns (results, returncode, error): results are the parsed
    PIPELINE_RESULT lines as dicts {step, status, detail}; error is set only
    when R could not be started at all.
    """
    results = []
    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        proc = subprocess.Popen(
            args,
            cwd=R_REPORTS_DIR,
            env=_r_subprocess_env(workers),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            **kwargs,
        )
    except OSError as exc:
        return results, None, str(exc)

    for raw in proc.stdout:
        line = ANSI_ESCAPE.sub("", raw).rstrip()
        if not line:
            continue
        if line.startswith(R_RESULT_PREFIX):
            parts = line[len(R_RESULT_PREFIX):].strip().split(" ", 2)
            if len(parts) >= 2:
                results.append({
                    "step": parts[0],
                    "status": parts[1],
                    "detail": parts[2] if len(parts) > 2 else "",
                })
                continue
        log_fn(f"  R | {line}")
    proc.wait()
    return results, proc.returncode, None


class MasterAnalysisLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("Master Analysis Launcher")
        self.root.geometry("1100x960")

        self.data_dir_var = tk.StringVar()
        self.plate_map_var = tk.StringVar()

        self.parallel_analyses_var = tk.BooleanVar(value=True)
        self.process_new_only_var = tk.BooleanVar(value=True)
        self.force_reprocess_var = tk.BooleanVar(value=False)

        self.total_worker_budget_var = tk.IntVar(value=max(1, os.cpu_count() or 1))

        self.analysis_vars = {
            "75Hz": tk.BooleanVar(value=True),
            "Twitch": tk.BooleanVar(value=True),
            "ForceFrequency": tk.BooleanVar(value=True),
            "Fatigue": tk.BooleanVar(value=True),
        }

        # Batch reports, and registering the batch in a dataset
        self.plate_report_var = tk.BooleanVar(value=True)
        self.emt_batch_report_var = tk.BooleanVar(value=True)
        self.register_var = tk.BooleanVar(value=False)
        self.register_config_var = tk.StringVar()
        self.batch_number_var = tk.IntVar(value=1)
        self.plate_id_var = tk.StringVar()

        # Dataset reports
        self.dataset_config_var = tk.StringVar()
        self.dataset_emt_var = tk.BooleanVar(value=True)
        self.dataset_plates_var = tk.BooleanVar(value=True)

        self.rscript_var = tk.StringVar(value=_find_rscript() or "")
        self.open_reports_var = tk.BooleanVar(value=False)

        self._log_queue = queue.Queue()
        self._running = False

        self._set_app_icon()
        self._build_ui()
        self.root.after(100, self._drain_log_queue)

    def _set_app_icon(self):
        """Window/Dock icon from assets/emt_icon.png (skipped if missing)."""
        icon_path = os.path.normpath(os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "..", "assets", "emt_icon.png"))
        try:
            self._icon = tk.PhotoImage(file=icon_path)
            self.root.iconphoto(True, self._icon)
        except tk.TclError:
            pass

    def _build_ui(self):
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True)
        dataset_configs = _list_dataset_configs()

        # Batch: analyze one batch and render its own reports ------------------
        batch = ttk.LabelFrame(frame, text="Batch — analyze one batch and render its reports")
        batch.grid(row=0, column=0, sticky="we", padx=10, pady=(10, 6))
        batch.grid_columnconfigure(1, weight=1)

        checks = ttk.Frame(batch)
        checks.grid(row=0, column=0, columnspan=3, sticky="w", padx=4, pady=(4, 0))
        ttk.Label(checks, text="Analyses:").grid(row=0, column=0, padx=6, sticky="w")
        ttk.Checkbutton(checks, text="75Hz", variable=self.analysis_vars["75Hz"]).grid(row=0, column=1, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(checks, text="Twitch (1Hz / 02Hz)", variable=self.analysis_vars["Twitch"]).grid(row=0, column=2, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(checks, text="Force Frequency (FF / FvF)", variable=self.analysis_vars["ForceFrequency"]).grid(row=0, column=3, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(checks, text="Fatigue", variable=self.analysis_vars["Fatigue"]).grid(row=0, column=4, padx=8, pady=4, sticky="w")

        ttk.Label(batch, text="Data folder:").grid(row=1, column=0, sticky="w", padx=10, pady=4)
        ttk.Entry(batch, textvariable=self.data_dir_var, width=88).grid(row=1, column=1, sticky="we", padx=10)
        ttk.Button(batch, text="Browse", command=self._browse_data_dir).grid(row=1, column=2, sticky="e", padx=10)

        ttk.Label(batch, text="Plate map file (.csv):").grid(row=2, column=0, sticky="w", padx=10, pady=4)
        ttk.Entry(batch, textvariable=self.plate_map_var, width=88).grid(row=2, column=1, sticky="we", padx=10)
        ttk.Button(batch, text="Browse", command=self._browse_plate_map).grid(row=2, column=2, sticky="e", padx=10)

        perf = ttk.LabelFrame(batch, text="Performance & incremental processing")
        perf.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=6)
        ttk.Checkbutton(perf, text="Run analysis types in parallel (processes)",
                        variable=self.parallel_analyses_var).grid(row=0, column=0, sticky="w", padx=10, pady=4)
        ttk.Label(perf, text="Total worker budget (shared across analyses/files):").grid(row=0, column=1, sticky="e", padx=10)
        ttk.Spinbox(perf, from_=1, to=64, textvariable=self.total_worker_budget_var, width=6).grid(
            row=0, column=2, sticky="w", padx=(0, 10))
        ttk.Checkbutton(perf, text="Process only new/changed files (manifest-based)",
                        variable=self.process_new_only_var).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=4)
        ttk.Checkbutton(perf, text="Force reprocess all (ignore manifest)",
                        variable=self.force_reprocess_var).grid(row=1, column=2, sticky="w", padx=10, pady=4)

        rep = ttk.Frame(batch)
        rep.grid(row=4, column=0, columnspan=3, sticky="w", padx=4, pady=(2, 0))
        ttk.Label(rep, text="Reports:").grid(row=0, column=0, padx=6, sticky="w")
        ttk.Checkbutton(rep, text="Plate trace report (raw data)",
                        variable=self.plate_report_var).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(rep, text="Single-batch EMT report (analyzed data)",
                        variable=self.emt_batch_report_var).grid(row=0, column=2, sticky="w", padx=8, pady=4)

        reg = ttk.Frame(batch)
        reg.grid(row=5, column=0, columnspan=3, sticky="w", padx=4, pady=(0, 4))
        ttk.Checkbutton(reg, text="Register in dataset", variable=self.register_var).grid(
            row=0, column=0, sticky="w", padx=6, pady=4)
        self.register_combo = ttk.Combobox(reg, textvariable=self.register_config_var,
                                           values=dataset_configs, state="readonly", width=48)
        self.register_combo.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(reg, text="New…", width=5, command=self._new_dataset).grid(row=0, column=2, padx=(0, 6))
        ttk.Label(reg, text="as batch #").grid(row=0, column=3, sticky="w", padx=(6, 0))
        ttk.Spinbox(reg, from_=1, to=999, textvariable=self.batch_number_var, width=5).grid(
            row=0, column=4, sticky="w", padx=(4, 6))
        ttk.Label(reg, text="plate").grid(row=0, column=5, sticky="w", padx=(6, 0))
        ttk.Entry(reg, textvariable=self.plate_id_var, width=10).grid(row=0, column=6, sticky="w", padx=(4, 6))
        ttk.Label(reg, text="(plate: only if this batch was run on several plates, e.g. DMD, WT;\n"
                            "results go to analyzed/<plate folder>/ next to the plate folders, "
                            "file names end in _<plate>)",
                  foreground="grey40").grid(row=0, column=7, sticky="w", padx=6)

        self.run_button = ttk.Button(batch, text="Run batch", command=self.run_pipeline)
        self.run_button.grid(row=6, column=2, sticky="e", padx=10, pady=(2, 8))

        # Dataset: reports across every registered batch -------------------------
        ds = ttk.LabelFrame(frame, text="Dataset reports — every batch registered in a dataset")
        ds.grid(row=1, column=0, sticky="we", padx=10, pady=6)
        ds.grid_columnconfigure(0, weight=1)

        top = ttk.Frame(ds)
        top.grid(row=0, column=0, columnspan=2, sticky="w", padx=4, pady=(4, 2))
        ttk.Label(top, text="Dataset:").grid(row=0, column=0, padx=6, sticky="w")
        self.dataset_combo = ttk.Combobox(top, textvariable=self.dataset_config_var,
                                          values=dataset_configs, state="readonly", width=60)
        self.dataset_combo.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(top, text="Open…", command=self._open_dataset).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(top, text="Refresh", command=self._refresh_datasets).grid(row=0, column=3, padx=6)
        self.dataset_csv_label = ttk.Label(top, text="", foreground="grey40")
        self.dataset_csv_label.grid(row=0, column=4, sticky="w", padx=6)

        cols = ("batch", "plate", "folder", "75hz", "ff", "fatigue", "timeforce")
        self.batch_tree = ttk.Treeview(ds, columns=cols, show="headings", height=4)
        headings = {"batch": "#", "plate": "Plate", "folder": "Folder", "75hz": "75Hz", "ff": "FF",
                    "fatigue": "Fatigue", "timeforce": "Raw (plates)"}
        widths = {"batch": 40, "plate": 60, "folder": 520, "75hz": 70, "ff": 70, "fatigue": 70,
                  "timeforce": 90}
        for c in cols:
            self.batch_tree.heading(c, text=headings[c])
            self.batch_tree.column(c, width=widths[c], stretch=(c == "folder"),
                                   anchor="w" if c == "folder" else "center")
        self.batch_tree.grid(row=1, column=0, sticky="we", padx=10, pady=4)
        tree_scroll = ttk.Scrollbar(ds, orient="vertical", command=self.batch_tree.yview)
        tree_scroll.grid(row=1, column=1, sticky="ns", pady=4)
        self.batch_tree.configure(yscrollcommand=tree_scroll.set)

        bottom = ttk.Frame(ds)
        bottom.grid(row=2, column=0, columnspan=2, sticky="we", padx=4, pady=(0, 8))
        bottom.grid_columnconfigure(2, weight=1)
        ttk.Checkbutton(bottom, text="Dataset EMT report",
                        variable=self.dataset_emt_var).grid(row=0, column=0, sticky="w", padx=6)
        ttk.Checkbutton(bottom, text="Combined plate trace report",
                        variable=self.dataset_plates_var).grid(row=0, column=1, sticky="w", padx=8)
        self.dataset_run_button = ttk.Button(bottom, text="Run dataset reports",
                                             command=self.run_dataset_reports)
        self.dataset_run_button.grid(row=0, column=3, sticky="e", padx=6)

        # Shared settings + log ---------------------------------------------------
        shared = ttk.Frame(frame)
        shared.grid(row=2, column=0, sticky="we", padx=10, pady=(4, 0))
        shared.grid_columnconfigure(1, weight=1)
        ttk.Label(shared, text="Rscript:").grid(row=0, column=0, sticky="w", padx=(4, 6))
        ttk.Entry(shared, textvariable=self.rscript_var, width=60).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(shared, text="Browse", command=self._browse_rscript).grid(row=0, column=2, padx=6)
        ttk.Checkbutton(shared, text="Open reports folder when done",
                        variable=self.open_reports_var).grid(row=0, column=3, sticky="w", padx=10)
        ttk.Button(shared, text="Clear Log", command=self._clear_log).grid(row=0, column=4, padx=6)

        ttk.Label(frame, text="Run log:", font=("Segoe UI", 10, "bold")).grid(
            row=3, column=0, sticky="w", padx=10, pady=(8, 2))
        self.log_box = ScrolledText(frame, height=14, width=128)
        self.log_box.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(4, weight=1)

        if dataset_configs:
            self.register_config_var.set(dataset_configs[0])
            self.dataset_config_var.set(dataset_configs[0])
        self.data_dir_var.trace_add("write", lambda *_: self._on_data_dir_changed())
        self.register_combo.bind("<<ComboboxSelected>>", lambda _e: self._suggest_batch_number())
        self.dataset_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_batch_table())
        self._refresh_batch_table()

    # Dataset helpers ------------------------------------------------------------

    def _suggest_batch_number(self):
        """Fill batch # and plate: this data folder's existing registration, or the
        next free batch number with no plate."""
        cfg = self.register_config_var.get().strip()
        if not cfg:
            return
        rows = _read_batch_rows(_dataset_batch_csv(cfg))
        number, plate = _suggest_registration(rows, self.data_dir_var.get().strip())
        self.batch_number_var.set(number)
        self.plate_id_var.set(plate)

    def _experiment_folder(self):
        """The folder holding this data folder's analyzed/ (Batch/timeforce -> Batch;
        with a plate name, Expt/dmd -> Expt), or None without a valid data folder."""
        data_dir = self.data_dir_var.get().strip()
        if not data_dir or not os.path.isdir(data_dir):
            return None
        plate = self.plate_id_var.get().strip()
        analyzed_root, _ = output_roots(data_dir, plate)
        return os.path.dirname(os.path.dirname(analyzed_root) if plate else analyzed_root)

    def _on_data_dir_changed(self):
        """Pick up dataset configs kept in this data folder's experiment folder
        (the plate/batch folder or the folder above it), then suggest a batch #."""
        data_dir = self.data_dir_var.get().strip()
        if data_dir and os.path.isdir(data_dir):
            folder = _plate_folder(data_dir)
            found = _find_dataset_configs(folder) + _find_dataset_configs(os.path.dirname(folder))
            for p in found:
                _remember_dataset(p)
            if found:
                self._refresh_datasets(select=_config_display(found[0]) if len(found) == 1 else None)
                return
        self._suggest_batch_number()

    def _select_dataset(self, value: str):
        self.register_config_var.set(value)
        self.dataset_config_var.set(value)

    def _open_dataset(self):
        """Add an existing dataset config (e.g. one kept in an experiment folder)."""
        path = filedialog.askopenfilename(
            title="Select a dataset config",
            initialdir=self._experiment_folder() or None,
            filetypes=[("Dataset config", "*.yml *.yaml"), ("All files", "*.*")])
        if not path:
            return
        if not _batch_csv_value(path):
            messagebox.showerror("Open dataset", f"{path}\nhas no `batch_csv:` line, so it isn't a dataset config.")
            return
        _remember_dataset(path)
        self._refresh_datasets(select=_config_display(path))
        self._log(f"Added dataset config {path}")

    def _new_dataset(self):
        """Create <folder>/<name>.yml from the example template, seeded with the
        current plate map's labels, and select it. The folder defaults to the
        experiment folder (the one holding analyzed/)."""
        name = simpledialog.askstring(
            "New dataset", "Name for the new dataset (e.g. T178I clones):", parent=self.root)
        if not name or not name.strip():
            return
        name = name.strip()
        slug = _slugify(name)
        if not slug:
            messagebox.showerror("New dataset", "Use at least one letter or digit in the name.")
            return
        folder = filedialog.askdirectory(
            title="Folder to keep the dataset config in (e.g. the experiment folder)",
            initialdir=self._experiment_folder() or None, parent=self.root)
        if not folder:
            return
        path = os.path.join(folder, f"{slug}.yml")
        if os.path.exists(path):
            messagebox.showerror("New dataset", f"{path} already exists. Add it with Open… instead.")
            return
        template = os.path.join(R_REPORTS_DIR, DATASET_TEMPLATE)
        try:
            with open(template, encoding="utf-8") as f:
                template_text = f.read()
        except OSError as exc:
            messagebox.showerror("New dataset", f"Could not read the template {template}:\n{exc}")
            return
        plate_map = self.plate_map_var.get().strip()
        labels = _plate_map_labels(plate_map) if os.path.isfile(plate_map) else []
        in_r_configs = _same_path(folder, CONFIGS_DIR)
        with open(path, "w", encoding="utf-8") as f:
            f.write(new_dataset_config_text(name, slug, labels, template_text, in_r_configs))
        self._log(f"Created dataset config {path}")

        _remember_dataset(path)
        self._refresh_datasets(select=_config_display(path))
        self.register_var.set(True)
        if labels:
            detail = (f"Conditions and colors were filled in from this plate map: {', '.join(labels)}.\n\n"
                      "If other batches spell a condition differently, add those spellings to its "
                      "list, and change the colors if you like.")
        else:
            detail = ("No plate map was selected, so `conditions` and `colors` are still the "
                      "template's examples. Replace them before rendering the dataset report.")
        where = ("" if in_r_configs else
                 f"\n\nIts batch CSV ({slug}-batches.csv) is kept next to it and the dataset "
                 f"reports go to {os.path.join(folder, 'reports', slug)}.")
        messagebox.showinfo("New dataset", f"Created {path}.\n\n{detail}{where}")

    def _refresh_datasets(self, select=None):
        configs = _list_dataset_configs()
        self.register_combo.configure(values=configs)
        self.dataset_combo.configure(values=configs)
        if select in configs:
            self._select_dataset(select)
        for var in (self.register_config_var, self.dataset_config_var):
            if var.get() not in configs:
                var.set(configs[0] if configs else "")
        self._refresh_batch_table()
        self._suggest_batch_number()

    def _refresh_batch_table(self):
        self.batch_tree.delete(*self.batch_tree.get_children())
        cfg = self.dataset_config_var.get().strip()
        if not cfg:
            self.dataset_csv_label.configure(text="No dataset configs yet: New… creates one, Open… adds an existing one.")
            return
        csv_path = _dataset_batch_csv(cfg)
        if not csv_path:
            self.dataset_csv_label.configure(text="Config has no batch_csv.")
            return
        if not os.path.isfile(csv_path):
            self.dataset_csv_label.configure(text=f"{os.path.basename(csv_path)} (no batches registered yet)")
            return
        batches = _summarize_batches(_read_batch_rows(csv_path))
        self.dataset_csv_label.configure(text=f"{os.path.basename(csv_path)} · {len(batches)} batch(es)")
        mark = {True: "✓", "excluded": "excluded"}
        for n, plate, folder, present in batches:
            self.batch_tree.insert("", "end", values=(
                n, plate, folder,
                *(mark.get(present.get(a), "–") for a in ("75hz", "ff", "fatigue", "timeforce")),
            ))

    def _browse_data_dir(self):
        path = filedialog.askdirectory(title="Select data directory")
        if path:
            self.data_dir_var.set(path)

    def _browse_plate_map(self):
        path = filedialog.askopenfilename(
            title="Select plate map file",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
        )
        if path:
            self.plate_map_var.set(path)

    def _browse_rscript(self):
        path = filedialog.askopenfilename(title="Select Rscript executable")
        if path:
            self.rscript_var.set(path)

    def _log(self, msg: str):
        """Thread-safe: queue the line; the Tk main loop writes it to the log box."""
        stamp = datetime.now().strftime("%H:%M:%S")
        self._log_queue.put(f"[{stamp}] {msg}")

    def _drain_log_queue(self):
        lines = []
        try:
            while True:
                lines.append(self._log_queue.get_nowait())
        except queue.Empty:
            pass
        if lines:
            self.log_box.insert("end", "\n".join(lines) + "\n")
            self.log_box.see("end")
        self.root.after(100, self._drain_log_queue)

    def _clear_log(self):
        self.log_box.delete("1.0", "end")

    @staticmethod
    def _compute_worker_plan(total_budget: int, job_files: list):
        """
        Returns (analysis_processes, [file_workers for each job]).

        The budget is split evenly (rounded up) across the analyses, and each
        analysis is capped at its file count since extra workers would sit idle.
        File size is a poor guide to run time (Fatigue files are small but slow).
        """
        total_budget = max(1, int(total_budget))
        num_jobs = max(1, len(job_files))
        analysis_processes = min(MAX_PARALLEL_ANALYSES, num_jobs, total_budget)

        share = -(-total_budget // analysis_processes)  # ceiling division
        file_workers = [max(1, min(len(files), share)) for files in job_files]
        return analysis_processes, file_workers

    # Run ----------------------------------------------------------------------

    def run_pipeline(self):
        """Validate inputs on the main thread, then run everything in a worker thread."""
        if self._running:
            return

        data_dir = self.data_dir_var.get().strip()
        plate_map_path = self.plate_map_var.get().strip()
        selected = [k for k, var in self.analysis_vars.items() if var.get()]

        do_plates = self.plate_report_var.get()
        do_emt_batch = self.emt_batch_report_var.get()
        do_register = self.register_var.get()
        register_config = self.register_config_var.get().strip()
        rscript = self.rscript_var.get().strip()
        any_reports = do_plates or do_emt_batch or do_register

        if not data_dir or not os.path.isdir(data_dir):
            messagebox.showerror("Invalid input", "Please select a valid data directory.")
            return
        if not plate_map_path or not os.path.isfile(plate_map_path):
            messagebox.showerror("Invalid input", "Please select a valid plate map CSV file.")
            return
        if not selected and not any_reports:
            messagebox.showerror("Nothing to run", "Select at least one analysis type or report.")
            return

        try:
            total_budget = int(self.total_worker_budget_var.get())
        except (tk.TclError, ValueError):
            total_budget = 0
        if total_budget < 1:
            messagebox.showerror("Invalid worker budget", "Worker budget must be >= 1.")
            return

        batch_number = None
        if any_reports and not self._check_r(rscript):
            return
        if do_register:
            if not register_config:
                messagebox.showerror("No dataset config", "Pick the dataset to register this batch in.")
                return
            try:
                batch_number = int(self.batch_number_var.get())
            except (tk.TclError, ValueError):
                batch_number = 0
            if batch_number < 1:
                messagebox.showerror("Invalid batch number", "Batch number must be >= 1.")
                return

        plate = self.plate_id_var.get().strip()
        if selected:
            analyzed_root, _ = output_roots(data_dir, plate)
            foreign = foreign_inputs(analyzed_root, data_dir)
            if foreign and not messagebox.askyesno(
                "Analyzed folder already in use",
                f"{analyzed_root}\nalready holds results analyzed from another data folder, e.g.\n"
                f"{os.path.dirname(foreign[0])}\n\n"
                "If that is another plate of this batch, choose No and fill in a plate name: "
                "each plate then gets its own analyzed/<plate folder>/.\n\n"
                "Continue anyway? Workbooks for the same day and protocol would be saved with a "
                "timestamp and the reports would use the newest.",
                icon="warning", default="no",
            ):
                return

        settings = {
            "data_dir": data_dir,
            "plate_map_path": plate_map_path,
            "selected": selected,
            "total_budget": total_budget,
            "parallel": self.parallel_analyses_var.get(),
            "process_new_only": self.process_new_only_var.get(),
            "force_reprocess": self.force_reprocess_var.get(),
            "do_plates": do_plates,
            "do_emt_batch": do_emt_batch,
            "register_config": register_config if do_register else None,
            "batch_number": batch_number,
            "plate": plate,   # output layout + file names; registered as the plate when registering
            "rscript": rscript,
            "open_reports": self.open_reports_var.get(),
        }

        self._set_running(True)
        threading.Thread(target=self._pipeline_worker, args=(settings,), daemon=True).start()

    def _check_r(self, rscript):
        """Rscript and render-pipeline.R are both there; shows an error if not."""
        if not rscript or not os.path.isfile(rscript):
            messagebox.showerror(
                "Rscript not found",
                "Reports need R. Set the Rscript path below the Dataset section "
                "(e.g. /usr/local/bin/Rscript), or untick the reports.",
            )
            return False
        if not os.path.isfile(os.path.join(R_REPORTS_DIR, R_PIPELINE_SCRIPT)):
            messagebox.showerror(
                "R reports not found",
                f"Could not find {R_PIPELINE_SCRIPT} in:\n{R_REPORTS_DIR}",
            )
            return False
        return True

    def _set_running(self, running: bool):
        """One job at a time: both Run buttons are disabled while either runs."""
        self._running = running
        state = ["disabled"] if running else ["!disabled"]
        self.run_button.state(state)
        self.dataset_run_button.state(state)

    def run_dataset_reports(self):
        """Render the dataset-wide reports for the dataset picked in the Dataset section."""
        if self._running:
            return
        config = self.dataset_config_var.get().strip()
        do_emt = self.dataset_emt_var.get()
        do_plates = self.dataset_plates_var.get()
        rscript = self.rscript_var.get().strip()
        if not config:
            messagebox.showerror("No dataset config",
                                 "Pick a dataset, or add one with Open… (or New… in the batch section).")
            return
        if not do_emt and not do_plates:
            messagebox.showerror("Nothing to run", "Tick at least one dataset report.")
            return
        if not self._check_r(rscript):
            return
        try:
            workers = max(1, int(self.total_worker_budget_var.get()))
        except (tk.TclError, ValueError):
            workers = None
        settings = {"config": config, "do_emt": do_emt, "do_plates": do_plates, "workers": workers,
                    "rscript": rscript, "open_reports": self.open_reports_var.get()}
        self._set_running(True)
        threading.Thread(target=self._dataset_worker, args=(settings,), daemon=True).start()

    def _dataset_worker(self, s):
        try:
            kind, title, msg = self._execute_dataset(s)
        except Exception as exc:
            self._log(f"ERROR: {exc}")
            for line in traceback.format_exc().splitlines():
                self._log(f"  {line}")
            kind, title, msg = "error", "Dataset run failed", str(exc)
        self.root.after(0, self._finish_run, kind, title, msg)

    def _execute_dataset(self, s):
        args = [s["rscript"], R_PIPELINE_SCRIPT, "--dataset-config", _config_path(s["config"])]
        if s["do_emt"]:
            args.append("--dataset-emt")
        if s["do_plates"]:
            args.append("--dataset-plates")
        self._log(f"Rendering dataset reports for {s['config']} (this can take a while)...")
        started = datetime.now()
        results, returncode, error = _run_r_pipeline(args, self._log, s.get("workers"))
        duration = round((datetime.now() - started).total_seconds(), 1)
        if error:
            self._log(f"ERROR running R: {error}")
            return "error", "Dataset run failed", error
        if not results:
            detail = f"Rscript exited with status {returncode} before any report finished; see the run log."
            self._log(f"ERROR: {detail}")
            return "error", "Dataset run failed", detail

        lines, failed, out_dir = [], 0, None
        for r in results:
            self._log(f"Report {r['step']}: {r['status']} — {r['detail']}")
            lines.append(f"{r['step']}: {r['status']} — {r['detail']}")
            if r["status"] == "OK":
                out_dir = out_dir or os.path.dirname(r["detail"])
            else:
                failed += 1
        self._log(f"Dataset reports finished in {duration} s.")
        if out_dir and s["open_reports"]:
            _open_folder(out_dir)
        msg = "\n".join(lines)
        if failed:
            return "warning", "Dataset reports finished with errors", msg
        return "info", "Dataset reports complete", msg

    def _pipeline_worker(self, s):
        try:
            kind, title, msg = self._execute(s)
        except Exception as exc:
            self._log(f"ERROR: {exc}")
            for line in traceback.format_exc().splitlines():
                self._log(f"  {line}")
            kind, title, msg = "error", "Run failed", str(exc)
        self.root.after(0, self._finish_run, kind, title, msg)

    def _finish_run(self, kind, title, msg):
        self._set_running(False)
        self._refresh_batch_table()
        show = {"info": messagebox.showinfo, "warning": messagebox.showwarning}.get(kind, messagebox.showerror)
        show(title, msg)

    def _execute(self, s):
        data_dir = s["data_dir"]
        plate_map_path = s["plate_map_path"]
        selected = s["selected"]

        self._log("Starting run...")
        self._log(f"Data directory: {data_dir}")
        self._log(f"Plate map: {plate_map_path}")

        analyzed_root, output_map = build_output_dirs(data_dir, selected, s["plate"])
        manifest_path = _manifest_path(analyzed_root)
        if s["plate"]:
            self._log(f"Plate {s['plate']}: results go to {analyzed_root}, "
                      f"file names end in _{file_tag_for(s['plate'])}")

        summary = []
        summary.append("MASTER ANALYSIS RUN SUMMARY")
        summary.append("=" * 80)
        summary.append(f"Timestamp: {datetime.now().isoformat(timespec='seconds')}")
        summary.append(f"Data directory: {data_dir}")
        summary.append(f"Plate map: {plate_map_path}")
        summary.append(f"Plate: {s['plate'] or '(none)'}")
        summary.append(f"Analyzed root: {analyzed_root}")
        summary.append(f"Parallel analyses: {s['parallel']}")
        summary.append(f"Total worker budget: {s['total_budget']}")
        summary.append(f"Process only new/changed: {s['process_new_only']}")
        summary.append(f"Force reprocess: {s['force_reprocess']}")
        summary.append(f"Manifest path: {manifest_path}")
        summary.append("")

        runs_ok, runs_failed = 0, 0
        if selected:
            runs_ok, runs_failed = self._run_analyses(s, analyzed_root, output_map, manifest_path, summary)

        report_results = []
        reports_root = None
        if s["do_plates"] or s["do_emt_batch"] or s["register_config"]:
            reports_root = build_reports_dir(data_dir, s["plate"])
            report_results = self._run_reports(s, analyzed_root, reports_root, summary)

        log_path = write_summary_log(analyzed_root, summary)
        self._log(f"Summary log written: {log_path}")

        if reports_root and s["open_reports"]:
            _open_folder(reports_root)

        reports_failed = [r for r in report_results if r["status"] != "OK"]
        lines = []
        misfiled = s.get("misfiled") or []
        if misfiled:
            lines.append(f"WARNING: {len(misfiled)} file(s) were in a folder for a different "
                         "protocol and were analyzed by their file name; see the run log.")
        if selected and runs_ok + runs_failed == 0:
            lines.append("No new/changed files to analyze.")
        elif selected:
            lines.append(f"Analyses completed: {runs_ok}, failed: {runs_failed}.")
        if report_results:
            lines.append(f"Reports rendered: {len(report_results) - len(reports_failed)}, "
                         f"failed: {len(reports_failed)}.")
            for r in reports_failed:
                lines.append(f"  {r['step']}: {r['detail']}")
        if any(r["step"] == "register" and r["status"] == "OK" for r in report_results):
            plate = f", plate {s['plate']}" if s["plate"] else ""
            lines.append(f"Registered as batch #{s['batch_number']}{plate} in {s['register_config']}. "
                         "Render the dataset reports from the Dataset section.")
        for r in report_results:
            if r["status"] == "OK" and r["step"] in ("plates", "emt-batch"):
                lines.append(f"{'Plate trace' if r['step'] == 'plates' else 'EMT'} report: {r['detail']}")
        lines.append(f"Summary: {log_path}")
        msg = "\n".join(lines)

        if runs_failed or reports_failed:
            return "warning", "Run completed with errors", msg
        if misfiled:
            return "warning", "Run complete, with warnings", msg
        return "info", "Run complete", msg

    def _run_analyses(self, s, analyzed_root, output_map, manifest_path, summary):
        """The Python analyses. Returns (runs_ok, runs_failed); appends to `summary`."""
        data_dir = s["data_dir"]
        plate_map_path = s["plate_map_path"]
        selected = s["selected"]
        total_budget = s["total_budget"]

        found = discover_files(data_dir)
        routed, unmatched, ignored = found.routed, found.unmatched, found.ignored
        s["misfiled"] = found.misfiled

        if found.misfiled:
            summary.append("Files in a folder for a different protocol (analyzed by file name):")
            for fname, folder, token in found.misfiled:
                msg = f'{fname} is in a "{folder}" folder but named {token}; analyzed as {token}'
                summary.append(f"  - {msg}")
                self._log(f"WARNING: {msg}")
            summary.append("")

        if found.duplicates:
            summary.append("Older duplicates for the same day x protocol (skipped; newest used):")
            for older, newer in found.duplicates:
                summary.append(f"  - {older}  (using {newer})")
                self._log(f"Skipped older duplicate: {older} (using {newer})")
            summary.append("")

        if found.ambiguous:
            self._log("Warning: Found files with ambiguous analysis matches; these were skipped.")
            summary.append("Ambiguous file matches (skipped):")
            for fname, matches in found.ambiguous:
                label_matches = [ANALYSES[m].name for m in matches]
                summary.append(f"  - {fname} -> {', '.join(label_matches)}")
                self._log(f"Skipped ambiguous: {fname} -> {', '.join(label_matches)}")
            summary.append("")

        if unmatched:
            self._log(f"{len(unmatched)} file(s) with no analysis (e.g. ASFatigue, Priming); listed in the run summary.")
            summary.append("Unmatched files (skipped):")
            for fname in unmatched:
                summary.append(f"  - {fname}")
            summary.append("")

        if ignored:
            summary.append("Ignored files (hidden/temp/non-data):")
            for fname in ignored:
                summary.append(f"  - {fname}")
            summary.append("")

        script_base_dir = os.path.dirname(os.path.abspath(__file__))
        done_keys = _load_done_manifest_keys(manifest_path)

        jobs = []
        skipped_already_done = 0

        for key in selected:
            definition = ANALYSES[key]
            discovered_files = routed.get(key, [])
            out_dir = output_map[key]

            self._log(f"{definition.name}: discovered {len(discovered_files)} file(s).")
            summary.append(f"{definition.name}:")
            summary.append(f"  Output folder: {out_dir}")
            summary.append(f"  Discovered files: {len(discovered_files)}")

            if not discovered_files:
                summary.append("  Status: SKIPPED (no matching files)")
                summary.append("")
                continue

            script_path = os.path.join(script_base_dir, definition.script_filename)
            dep_sig = _dependency_sig(plate_map_path, script_path)

            files_to_run = []
            skipped_here = 0
            should_filter_manifest = s["process_new_only"] and (not s["force_reprocess"])

            if should_filter_manifest and _output_dir_has_files(out_dir):
                for f in discovered_files:
                    key_sig = _signature_key(key, _file_sig(f), dep_sig)
                    if key_sig in done_keys:
                        skipped_here += 1
                    else:
                        files_to_run.append(f)
            else:
                files_to_run = list(discovered_files)

            skipped_already_done += skipped_here
            summary.append(f"  Skipped already analyzed: {skipped_here}")
            summary.append(f"  Files to process now: {len(files_to_run)}")

            if not files_to_run:
                summary.append("  Status: SKIPPED (all files already analyzed)")
                summary.append("")
                continue

            jobs.append((key, definition, files_to_run, script_path, dep_sig))

        if not jobs:
            summary.append(f"SKIPPED_ALREADY_ANALYZED total: {skipped_already_done}")
            summary.append("")
            self._log("No new/changed files to analyze.")
            return 0, 0

        runs_ok = 0
        runs_failed = 0
        manifest_rows = []

        # Decide orchestration plan.
        if s["parallel"] and len(jobs) > 1:
            proc_workers, job_workers = self._compute_worker_plan(total_budget, [j[2] for j in jobs])
            self._log(
                f"Parallel mode: {proc_workers} analysis process(es); file workers: "
                + ", ".join(f"{j[1].name} {w}" for j, w in zip(jobs, job_workers))
            )

            payloads = []
            for (key, definition, files, script_path, dep_sig), file_workers in zip(jobs, job_workers):
                payloads.append({
                    "key": key,
                    "definition_name": definition.name,
                    "script_path": script_path,
                    "function_name": definition.function_name,
                    "files": files,
                    "plate_map_path": plate_map_path,
                    "output_dir": output_map[key],
                    "file_workers": file_workers,
                    "file_tag": file_tag_for(s["plate"]),
                })
                summary.append(f"  File workers: {file_workers}")

            with ProcessPoolExecutor(max_workers=proc_workers) as executor:
                future_map = {
                    executor.submit(_run_analysis_job, payload): payload
                    for payload in payloads
                }

                for fut in as_completed(future_map):
                    payload = future_map[fut]
                    try:
                        result = fut.result()
                    except BrokenProcessPool:
                        # The analysis process itself died (usually out of memory).
                        result = {
                            "key": payload["key"],
                            "name": payload["definition_name"],
                            "status": "FAILED",
                            "duration_seconds": 0,
                            "error": WORKER_CRASHED,
                            "traceback": traceback.format_exc(),
                        }
                    key = result["key"]
                    name = result["name"]
                    status = result["status"]

                    self._log(f"{name}: {status} ({result['duration_seconds']}s)")
                    summary.append(f"  Duration: {result['duration_seconds']} s")
                    summary.append(f"  Status: {status}")

                    for f in payload["files"]:
                        input_path, input_size, input_mtime_ns = _file_sig(f)
                        plate_path, plate_mtime_ns, script_path, script_mtime_ns = _dependency_sig(
                            payload["plate_map_path"], payload["script_path"]
                        )
                        manifest_rows.append({
                            "timestamp": datetime.now().isoformat(timespec="seconds"),
                            "analysis_key": key,
                            "analysis_name": name,
                            "input_path": input_path,
                            "input_size": input_size,
                            "input_mtime_ns": input_mtime_ns,
                            "plate_map_path": plate_path,
                            "plate_map_mtime_ns": plate_mtime_ns,
                            "script_path": script_path,
                            "script_mtime_ns": script_mtime_ns,
                            "output_dir": payload["output_dir"],
                            "status": status,
                            "error": result["error"] or "",
                        })

                    if status == "DONE":
                        runs_ok += 1
                    else:
                        runs_failed += 1
                        summary.append(f"  Error: {result['error']}")
                        summary.append("  Traceback:")
                        for line in (result["traceback"] or "").splitlines():
                            summary.append(f"    {line}")
                    summary.append("")

        else:
            # Sequential analyses, but each analysis can use full worker budget internally.
            file_workers = max(1, total_budget)
            self._log(f"Sequential analysis mode; {file_workers} file worker(s) per analysis.")

            for key, definition, files, script_path, dep_sig in jobs:
                summary.append(f"  File workers: {file_workers}")
                status = "DONE"
                err_text = ""
                tb_text = ""
                try:
                    run_fn = _load_callable(script_path, definition.function_name)
                    self._log(f"Running {definition.name}...")
                    started = datetime.now()
                    run_fn(
                        input_files=files,
                        plate_map_path=plate_map_path,
                        output_dir=output_map[key],
                        max_workers=file_workers,
                        file_tag=file_tag_for(s["plate"]),
                    )
                    ended = datetime.now()
                    duration = round((ended - started).total_seconds(), 2)
                    self._log(f"Completed {definition.name} ({duration}s).")
                    summary.append(f"  Duration: {duration} s")
                    summary.append("  Status: DONE")
                    runs_ok += 1
                except Exception as exc:
                    status = "FAILED"
                    err_text = str(exc)
                    tb_text = traceback.format_exc()
                    self._log(f"ERROR in {definition.name}: {exc}")
                    summary.append("  Status: FAILED")
                    summary.append(f"  Error: {exc}")
                    summary.append("  Traceback:")
                    for line in tb_text.splitlines():
                        summary.append(f"    {line}")
                    runs_failed += 1

                for f in files:
                    input_path, input_size, input_mtime_ns = _file_sig(f)
                    plate_path, plate_mtime_ns, script_abs, script_mtime_ns = _dependency_sig(
                        plate_map_path, script_path
                    )
                    manifest_rows.append({
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "analysis_key": key,
                        "analysis_name": definition.name,
                        "input_path": input_path,
                        "input_size": input_size,
                        "input_mtime_ns": input_mtime_ns,
                        "plate_map_path": plate_path,
                        "plate_map_mtime_ns": plate_mtime_ns,
                        "script_path": script_abs,
                        "script_mtime_ns": script_mtime_ns,
                        "output_dir": output_map[key],
                        "status": status,
                        "error": err_text if status != "DONE" else "",
                    })

                summary.append("")

        if manifest_rows:
            _append_manifest_rows(manifest_path, manifest_rows)
            self._log(f"Manifest updated: {manifest_path}")

        summary.append("Run totals:")
        summary.append(f"  Analyses completed: {runs_ok}")
        summary.append(f"  Analyses failed: {runs_failed}")
        summary.append(f"  SKIPPED_ALREADY_ANALYZED total: {skipped_already_done}")
        summary.append("")
        return runs_ok, runs_failed

    def _run_reports(self, s, analyzed_root, reports_root, summary):
        """Render the R reports via render-pipeline.R. Returns the parsed step results."""
        data_dir = os.path.abspath(s["data_dir"])
        args = [
            s["rscript"], R_PIPELINE_SCRIPT,
            "--timeforce-dir", data_dir,
            "--plate-map", os.path.abspath(s["plate_map_path"]),
            "--analyzed-dir", analyzed_root,
            "--reports-dir", reports_root,
            "--name", _batch_name(data_dir),
        ]
        if s["do_plates"]:
            args.append("--plates")
        if s["do_emt_batch"]:
            args.append("--emt-batch")
        if s["register_config"]:
            # The dataset's conditions/colors are also used for this batch's reports
            args += ["--dataset-config", _config_path(s["register_config"]), "--register",
                     "--batch-number", str(s["batch_number"])]
            if s["plate"]:
                # Several plates in this batch: the EMT report covers all of them
                args += ["--plate", s["plate"]]

        summary.append("Reports:")
        summary.append(f"  Reports folder: {reports_root}")
        self._log("Rendering R reports (this can take a few minutes)...")
        started = datetime.now()
        results, returncode, error = _run_r_pipeline(args, self._log, s["total_budget"])
        duration = round((datetime.now() - started).total_seconds(), 2)

        if error:
            self._log(f"ERROR running R: {error}")
            summary.append(f"  Status: FAILED ({error})")
            summary.append("")
            return [{"step": "R", "status": "FAILED", "detail": error}]
        if not results:
            detail = f"Rscript exited with status {returncode} before any report finished; see the run log."
            self._log(f"ERROR: {detail}")
            summary.append(f"  Status: FAILED ({detail})")
            summary.append("")
            return [{"step": "R", "status": "FAILED", "detail": detail}]

        for r in results:
            self._log(f"Report {r['step']}: {r['status']} — {r['detail']}")
            summary.append(f"  {r['step']}: {r['status']} — {r['detail']}")
        summary.append(f"  Duration: {duration} s")
        summary.append("")
        return results


if __name__ == "__main__":
    root = tk.Tk()
    app = MasterAnalysisLauncher(root)
    root.mainloop()
