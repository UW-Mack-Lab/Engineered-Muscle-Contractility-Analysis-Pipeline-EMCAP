import os
import csv
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


class MasterAnalysisLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("Master Analysis Launcher")
        self.root.geometry("1080x760")

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

        self._build_ui()

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

        btns = ttk.Frame(frame)
        btns.grid(row=5, column=0, columnspan=4, sticky="w", padx=10, pady=10)

        ttk.Button(btns, text="Run", command=self.run_pipeline).grid(row=0, column=0, padx=5)
        ttk.Button(btns, text="Clear Log", command=self._clear_log).grid(row=0, column=1, padx=5)

        ttk.Label(frame, text="Run log:", font=("Segoe UI", 10, "bold")).grid(row=6, column=0, sticky="w", **padding)
        self.log_box = ScrolledText(frame, height=26, width=128)
        self.log_box.grid(row=7, column=0, columnspan=4, sticky="nsew", padx=10, pady=(0, 10))

        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_rowconfigure(7, weight=1)

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

    def _log(self, msg: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {msg}"
        self.log_box.insert("end", line + "\n")
        self.log_box.see("end")
        self.root.update_idletasks()

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

    def run_pipeline(self):
        data_dir = self.data_dir_var.get().strip()
        plate_map_path = self.plate_map_var.get().strip()
        selected = [k for k, var in self.analysis_vars.items() if var.get()]

        if not data_dir or not os.path.isdir(data_dir):
            messagebox.showerror("Invalid input", "Please select a valid data directory.")
            return
        if not plate_map_path or not os.path.isfile(plate_map_path):
            messagebox.showerror("Invalid input", "Please select a valid plate map CSV file.")
            return
        if not selected:
            messagebox.showerror("No analyses selected", "Select at least one analysis type.")
            return

        total_budget = self.total_worker_budget_var.get()
        if total_budget < 1:
            messagebox.showerror("Invalid worker budget", "Worker budget must be >= 1.")
            return

        self._log("Starting run...")
        self._log(f"Data directory: {data_dir}")
        self._log(f"Plate map: {plate_map_path}")

        routed, unmatched, ignored, ambiguous = discover_files(data_dir)
        analyzed_root, output_map = build_output_dirs(data_dir, selected)
        manifest_path = _manifest_path(analyzed_root)

        summary = []
        summary.append("MASTER ANALYSIS RUN SUMMARY")
        summary.append("=" * 80)
        summary.append(f"Timestamp: {datetime.now().isoformat(timespec='seconds')}")
        summary.append(f"Data directory: {data_dir}")
        summary.append(f"Plate map: {plate_map_path}")
        summary.append(f"Analyzed root: {analyzed_root}")
        summary.append(f"Parallel analyses: {self.parallel_analyses_var.get()}")
        summary.append(f"Total worker budget: {total_budget}")
        summary.append(f"Process only new/changed: {self.process_new_only_var.get()}")
        summary.append(f"Force reprocess: {self.force_reprocess_var.get()}")
        summary.append(f"Manifest path: {manifest_path}")
        summary.append("")

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
            should_filter_manifest = self.process_new_only_var.get() and (not self.force_reprocess_var.get())

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
            log_path = write_summary_log(analyzed_root, summary)
            self._log(f"No new/changed files to process. Summary log written: {log_path}")
            messagebox.showinfo("Run complete", f"No new/changed files to process.\nSummary: {log_path}")
            return

        runs_ok = 0
        runs_failed = 0
        manifest_rows = []

        # Decide orchestration plan.
        if self.parallel_analyses_var.get() and len(jobs) > 1:
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

        log_path = write_summary_log(analyzed_root, summary)
        self._log(f"Summary log written: {log_path}")

        if runs_failed > 0:
            messagebox.showwarning(
                "Run completed with errors",
                f"Run finished with {runs_failed} failed analysis type(s).\nSummary: {log_path}",
            )
        else:
            messagebox.showinfo("Run complete", f"All selected analyses finished.\nSummary: {log_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MasterAnalysisLauncher(root)
    root.mainloop()
