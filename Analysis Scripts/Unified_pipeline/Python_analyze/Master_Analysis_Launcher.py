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
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


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


def _extract_suffix_token(filename: str):
    """
    Extract token between final '_' and extension.
    Example: d4_expt1_FF.csv -> FF
    """
    stem, _ = os.path.splitext(filename)
    if "_" not in stem:
        return None
    return stem.rsplit("_", 1)[1]


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

    started = datetime.now()
    try:
        run_fn = _load_callable(script_path, function_name)
        run_fn(
            input_files=files,
            plate_map_path=plate_map_path,
            output_dir=output_dir,
            max_workers=file_workers,
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


def discover_files(data_dir: str):
    routed = {key: [] for key in ANALYSES.keys()}
    unmatched = []
    ignored = []
    ambiguous = []

    entries = sorted(os.listdir(data_dir))
    for entry in entries:
        full_path = os.path.join(data_dir, entry)

        if not os.path.isfile(full_path):
            continue
        if _is_temp_or_hidden(entry):
            ignored.append(entry)
            continue

        ext = os.path.splitext(entry)[1].lower()
        if ext not in DATA_EXTENSIONS:
            ignored.append(entry)
            continue

        token = _extract_suffix_token(entry)
        if not token:
            unmatched.append(entry)
            continue

        token_l = token.lower()
        matches = [
            key for key, definition in ANALYSES.items()
            if token_l in definition.suffix_tokens
        ]

        if len(matches) == 0:
            unmatched.append(entry)
            continue
        if len(matches) > 1:
            ambiguous.append((entry, matches))
            continue

        routed[matches[0]].append(full_path)

    return routed, unmatched, ignored, ambiguous


def build_output_dirs(data_dir: str, selected_keys):
    parent = os.path.dirname(os.path.abspath(data_dir))
    analyzed_root = os.path.join(parent, "analyzed")
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

def build_reports_dir(data_dir: str):
    """<parent of data folder>/reports, next to analyzed/."""
    parent = os.path.dirname(os.path.abspath(data_dir))
    reports_root = os.path.join(parent, "reports")
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


def _r_subprocess_env():
    """Environment for Rscript: points R at the env's Quarto unless QUARTO_PATH
    is already set."""
    env = dict(os.environ)
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


def _list_dataset_configs():
    """Dataset EMT configs, as paths relative to R_REPORTS_DIR (configs/x.yml)."""
    cfg_dir = os.path.join(R_REPORTS_DIR, "configs")
    if not os.path.isdir(cfg_dir):
        return []
    out = []
    for name in sorted(os.listdir(cfg_dir)):
        if name.lower().endswith((".yml", ".yaml")) and name != SINGLE_BATCH_DEFAULTS:
            out.append(f"configs/{name}")
    return out


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


def _run_r_pipeline(args, log_fn):
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
            env=_r_subprocess_env(),
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
        self.root.geometry("1100x860")

        self.data_dir_var = tk.StringVar()
        self.plate_map_var = tk.StringVar()

        self.parallel_analyses_var = tk.BooleanVar(value=True)
        self.process_new_only_var = tk.BooleanVar(value=True)
        self.force_reprocess_var = tk.BooleanVar(value=False)

        cpu_total = max(1, (os.cpu_count() or 1))
        default_budget = min(12, cpu_total)
        self.total_worker_budget_var = tk.IntVar(value=default_budget)

        self.analysis_vars = {
            "75Hz": tk.BooleanVar(value=True),
            "Twitch": tk.BooleanVar(value=True),
            "ForceFrequency": tk.BooleanVar(value=True),
            "Fatigue": tk.BooleanVar(value=True),
        }

        # R report options
        self.plate_report_var = tk.BooleanVar(value=True)
        self.emt_batch_report_var = tk.BooleanVar(value=True)
        self.dataset_report_var = tk.BooleanVar(value=False)
        self.dataset_config_var = tk.StringVar()
        self.batch_number_var = tk.IntVar(value=1)
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
        padding = {"padx": 10, "pady": 8}

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Select analyses:", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, sticky="w", **padding
        )

        checks = ttk.Frame(frame)
        checks.grid(row=1, column=0, columnspan=4, sticky="w", padx=10)

        ttk.Checkbutton(checks, text="75Hz", variable=self.analysis_vars["75Hz"]).grid(row=0, column=0, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(checks, text="Twitch (1Hz / 02Hz)", variable=self.analysis_vars["Twitch"]).grid(row=0, column=1, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(checks, text="Force Frequency (FF / FvF)", variable=self.analysis_vars["ForceFrequency"]).grid(row=0, column=2, padx=8, pady=4, sticky="w")
        ttk.Checkbutton(checks, text="Fatigue", variable=self.analysis_vars["Fatigue"]).grid(row=0, column=3, padx=8, pady=4, sticky="w")

        ttk.Label(frame, text="Data folder:").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.data_dir_var, width=88).grid(row=2, column=1, columnspan=2, sticky="we", padx=10)
        ttk.Button(frame, text="Browse", command=self._browse_data_dir).grid(row=2, column=3, sticky="e", padx=10)

        ttk.Label(frame, text="Plate map file (.csv):").grid(row=3, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.plate_map_var, width=88).grid(row=3, column=1, columnspan=2, sticky="we", padx=10)
        ttk.Button(frame, text="Browse", command=self._browse_plate_map).grid(row=3, column=3, sticky="e", padx=10)

        perf = ttk.LabelFrame(frame, text="Performance & Incremental Processing")
        perf.grid(row=4, column=0, columnspan=4, sticky="we", padx=10, pady=8)

        ttk.Checkbutton(
            perf,
            text="Run analysis types in parallel (processes)",
            variable=self.parallel_analyses_var,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=6)

        ttk.Label(perf, text="Total worker budget (shared across analyses/files):").grid(
            row=0, column=1, sticky="e", padx=10
        )
        ttk.Spinbox(perf, from_=1, to=64, textvariable=self.total_worker_budget_var, width=6).grid(
            row=0, column=2, sticky="w", padx=(0, 10)
        )

        ttk.Checkbutton(
            perf,
            text="Process only new/changed files (manifest-based)",
            variable=self.process_new_only_var,
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=4)

        ttk.Checkbutton(
            perf,
            text="Force reprocess all (ignore manifest)",
            variable=self.force_reprocess_var,
        ).grid(row=1, column=2, sticky="w", padx=10, pady=4)

        reports = ttk.LabelFrame(frame, text="Reports (R) — rendered after the analyses")
        reports.grid(row=5, column=0, columnspan=4, sticky="we", padx=10, pady=8)

        ttk.Checkbutton(
            reports,
            text="Plate trace report (raw data)",
            variable=self.plate_report_var,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=4)
        ttk.Checkbutton(
            reports,
            text="Single-batch EMT report (analyzed data)",
            variable=self.emt_batch_report_var,
        ).grid(row=0, column=1, columnspan=2, sticky="w", padx=10, pady=4)
        ttk.Checkbutton(
            reports,
            text="Open reports folder when done",
            variable=self.open_reports_var,
        ).grid(row=0, column=3, sticky="w", padx=10, pady=4)

        ttk.Checkbutton(
            reports,
            text="Add to dataset report:",
            variable=self.dataset_report_var,
        ).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        dataset_configs = _list_dataset_configs()
        self.dataset_combo = ttk.Combobox(
            reports, textvariable=self.dataset_config_var, values=dataset_configs,
            state="readonly", width=36,
        )
        self.dataset_combo.grid(row=1, column=1, sticky="w", padx=10)
        if dataset_configs:
            self.dataset_config_var.set(dataset_configs[0])
        batch_frame = ttk.Frame(reports)
        batch_frame.grid(row=1, column=2, sticky="w", padx=10)
        ttk.Label(batch_frame, text="as batch #").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(batch_frame, from_=1, to=999, textvariable=self.batch_number_var, width=5).grid(
            row=0, column=1, sticky="w", padx=(4, 0)
        )

        ttk.Label(reports, text="Rscript:").grid(row=2, column=0, sticky="w", padx=10, pady=(4, 8))
        ttk.Entry(reports, textvariable=self.rscript_var, width=60).grid(
            row=2, column=1, columnspan=2, sticky="we", padx=10, pady=(4, 8)
        )
        ttk.Button(reports, text="Browse", command=self._browse_rscript).grid(
            row=2, column=3, sticky="w", padx=10, pady=(4, 8)
        )
        reports.grid_columnconfigure(2, weight=1)

        btns = ttk.Frame(frame)
        btns.grid(row=6, column=0, columnspan=4, sticky="w", padx=10, pady=10)

        self.run_button = ttk.Button(btns, text="Run", command=self.run_pipeline)
        self.run_button.grid(row=0, column=0, padx=5)
        ttk.Button(btns, text="Clear Log", command=self._clear_log).grid(row=0, column=1, padx=5)

        ttk.Label(frame, text="Run log:", font=("Segoe UI", 10, "bold")).grid(row=7, column=0, sticky="w", **padding)
        self.log_box = ScrolledText(frame, height=24, width=128)
        self.log_box.grid(row=8, column=0, columnspan=4, sticky="nsew", padx=10, pady=(0, 10))

        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_rowconfigure(8, weight=1)

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
    def _compute_worker_plan(total_budget: int, num_jobs: int):
        """
        Returns (analysis_processes, file_workers_per_job).
        Uses budget-aware balancing to avoid CPU oversubscription.
        """
        total_budget = max(1, int(total_budget))
        num_jobs = max(1, int(num_jobs))

        analysis_processes = min(MAX_PARALLEL_ANALYSES, num_jobs, total_budget)
        file_workers = max(1, total_budget // analysis_processes)
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
        do_dataset = self.dataset_report_var.get()
        dataset_config = self.dataset_config_var.get().strip()
        rscript = self.rscript_var.get().strip()
        any_reports = do_plates or do_emt_batch or do_dataset

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
        if any_reports:
            if not rscript or not os.path.isfile(rscript):
                messagebox.showerror(
                    "Rscript not found",
                    "Reports need R. Set the Rscript path in the Reports section "
                    "(e.g. /usr/local/bin/Rscript), or untick the reports.",
                )
                return
            if not os.path.isfile(os.path.join(R_REPORTS_DIR, R_PIPELINE_SCRIPT)):
                messagebox.showerror(
                    "R reports not found",
                    f"Could not find {R_PIPELINE_SCRIPT} in:\n{R_REPORTS_DIR}",
                )
                return
        if do_dataset:
            if not dataset_config:
                messagebox.showerror("No dataset config", "Pick a dataset config for the dataset report.")
                return
            try:
                batch_number = int(self.batch_number_var.get())
            except (tk.TclError, ValueError):
                batch_number = 0
            if batch_number < 1:
                messagebox.showerror("Invalid batch number", "Batch number must be >= 1.")
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
            "dataset_config": dataset_config if do_dataset else None,
            "batch_number": batch_number,
            "rscript": rscript,
            "open_reports": self.open_reports_var.get(),
        }

        self._running = True
        self.run_button.state(["disabled"])
        threading.Thread(target=self._pipeline_worker, args=(settings,), daemon=True).start()

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
        self._running = False
        self.run_button.state(["!disabled"])
        show = {"info": messagebox.showinfo, "warning": messagebox.showwarning}.get(kind, messagebox.showerror)
        show(title, msg)

    def _execute(self, s):
        data_dir = s["data_dir"]
        plate_map_path = s["plate_map_path"]
        selected = s["selected"]

        self._log("Starting run...")
        self._log(f"Data directory: {data_dir}")
        self._log(f"Plate map: {plate_map_path}")

        analyzed_root, output_map = build_output_dirs(data_dir, selected)
        manifest_path = _manifest_path(analyzed_root)

        summary = []
        summary.append("MASTER ANALYSIS RUN SUMMARY")
        summary.append("=" * 80)
        summary.append(f"Timestamp: {datetime.now().isoformat(timespec='seconds')}")
        summary.append(f"Data directory: {data_dir}")
        summary.append(f"Plate map: {plate_map_path}")
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
        if s["do_plates"] or s["do_emt_batch"] or s["dataset_config"]:
            reports_root = build_reports_dir(data_dir)
            report_results = self._run_reports(s, analyzed_root, reports_root, summary)

        log_path = write_summary_log(analyzed_root, summary)
        self._log(f"Summary log written: {log_path}")

        if reports_root and s["open_reports"]:
            _open_folder(reports_root)

        reports_failed = [r for r in report_results if r["status"] != "OK"]
        lines = []
        if selected and runs_ok + runs_failed == 0:
            lines.append("No new/changed files to analyze.")
        elif selected:
            lines.append(f"Analyses completed: {runs_ok}, failed: {runs_failed}.")
        if report_results:
            lines.append(f"Reports rendered: {len(report_results) - len(reports_failed)}, "
                         f"failed: {len(reports_failed)}.")
            for r in reports_failed:
                lines.append(f"  {r['step']}: {r['detail']}")
        if reports_root:
            lines.append(f"Reports folder: {reports_root}")
        lines.append(f"Summary: {log_path}")
        msg = "\n".join(lines)

        if runs_failed or reports_failed:
            return "warning", "Run completed with errors", msg
        return "info", "Run complete", msg

    def _run_analyses(self, s, analyzed_root, output_map, manifest_path, summary):
        """The Python analyses. Returns (runs_ok, runs_failed); appends to `summary`."""
        data_dir = s["data_dir"]
        plate_map_path = s["plate_map_path"]
        selected = s["selected"]
        total_budget = s["total_budget"]

        routed, unmatched, ignored, ambiguous = discover_files(data_dir)

        if ambiguous:
            self._log("Warning: Found files with ambiguous analysis matches; these were skipped.")
            summary.append("Ambiguous file matches (skipped):")
            for fname, matches in ambiguous:
                label_matches = [ANALYSES[m].name for m in matches]
                summary.append(f"  - {fname} -> {', '.join(label_matches)}")
                self._log(f"Skipped ambiguous: {fname} -> {', '.join(label_matches)}")
            summary.append("")

        if unmatched:
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
            proc_workers, file_workers = self._compute_worker_plan(total_budget, len(jobs))
            self._log(
                f"Parallel mode: {proc_workers} analysis process(es); "
                f"{file_workers} file worker(s) per analysis."
            )

            payloads = []
            for key, definition, files, script_path, dep_sig in jobs:
                payloads.append({
                    "key": key,
                    "definition_name": definition.name,
                    "script_path": script_path,
                    "function_name": definition.function_name,
                    "files": files,
                    "plate_map_path": plate_map_path,
                    "output_dir": output_map[key],
                    "file_workers": file_workers,
                })
                summary.append(f"  File workers: {file_workers}")

            with ProcessPoolExecutor(max_workers=proc_workers) as executor:
                future_map = {
                    executor.submit(_run_analysis_job, payload): payload
                    for payload in payloads
                }

                for fut in as_completed(future_map):
                    payload = future_map[fut]
                    result = fut.result()
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
        if s["dataset_config"]:
            args += ["--dataset-config", s["dataset_config"], "--batch-number", str(s["batch_number"])]

        summary.append("Reports:")
        summary.append(f"  Reports folder: {reports_root}")
        self._log("Rendering R reports (this can take a few minutes)...")
        started = datetime.now()
        results, returncode, error = _run_r_pipeline(args, self._log)
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
