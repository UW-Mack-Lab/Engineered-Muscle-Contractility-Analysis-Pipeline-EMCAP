# Importers for PB-analyzed EMT Excel workbooks --------------------------------
# Generalized from importMultipleBatches / importExcelSeries / getFilesAndDays /
# importMetricBatches / importMetrics in pulse+AnalysisPlotting-fcns.R.

suppressPackageStartupMessages({
  library(tidyverse)
  library(readxl)
})

# Which workbook sheet holds which kind of data, per analysis type
emt_sheets <- list(
  ff      = list(traces = "Representative Traces", metrics = "Relaxation Summary"),
  `75hz`  = list(traces = "Representative Traces", metrics = "Relaxation Summary"),
  fatigue = list(traces = "Condition Avg Fatigue", metrics = "Fatigue Summary")
)

# Sheet column name -> internal name. Renaming by name (not position) so a
# reordered sheet doesn't silently swap columns.
trace_columns <- c(
  time        = "Time (seconds)",
  norm_trace  = "Normalized Trace",
  norm_sd     = "Standard Deviation",
  force       = "Force (μN)",
  force_sd    = "Force SD",
  smoothed    = "Smoothed Force (μN)",
  smoothed_sd = "Smoothed Force SD",
  condition   = "Condition"
)

fatigue_trace_columns <- c(
  condition   = "Condition",
  stim_number = "Stim Number",
  norm_force  = "Mean Normalized Force (% of T1)",
  sd_force    = "StdDev Normalized Force (% of T1)",
  n_wells     = "N (Wells contributing)"
)

# Columns that identify a well/stim rather than being a metric
metric_id_columns <- c("Plate", "Well", "Condition", "Stim Peak Time",
                       "Stim Number", "Hz", "file", "day", "batch")

# Batch sheet ----------------------------------------------------------------

#' Read the batch-locations CSV, honoring the `exclude` column.
read_batch_sheet <- function(path, analysis = NULL, batches = NULL) {
  b <- read.csv(path, stringsAsFactors = FALSE, check.names = FALSE)
  needed <- c("batch", "analysis", "directory")
  miss <- setdiff(needed, names(b))
  if (length(miss)) stop("Batch CSV ", path, " is missing column(s): ",
                         paste(miss, collapse = ", "), call. = FALSE)
  if (!"exclude" %in% names(b)) b$exclude <- 0
  b <- b |>
    mutate(analysis = tolower(trimws(analysis)),
           exclude  = coalesce(as.integer(exclude), 0L))
  attr(b, "excluded") <- filter(b, exclude != 0)
  b <- filter(b, exclude == 0)
  if (!is.null(analysis)) b <- filter(b, .data$analysis == tolower(!!analysis))
  if (!is.null(batches) && !identical(batches, "all")) {
    b <- filter(b, batch %in% as.numeric(unlist(batches)))
  }
  b
}

# Files ----------------------------------------------------------------------

#' List analyzed workbooks in a directory, one per day.
#' Day comes from a leading "d<number>_" in the file name. When a day has more
#' than one workbook (e.g. a re-run saved with a timestamp suffix), the most
#' recently modified file is used and a warning names the others.
list_day_files <- function(dir, batch = NA) {
  if (!dir.exists(dir)) {
    stop(sprintf("Batch %s: directory does not exist:\n  %s", batch, dir), call. = FALSE)
  }
  fs <- list.files(dir, pattern = "\\.xlsx$", full.names = TRUE)
  fs <- fs[!startsWith(basename(fs), "~$")]  # Excel lock files
  if (length(fs) == 0) {
    stop(sprintf("Batch %s: no .xlsx files in\n  %s", batch, dir), call. = FALSE)
  }
  out <- tibble(file = fs,
                day = as.integer(str_match(basename(fs), "^[dD](\\d+)_")[, 2]),
                mtime = file.mtime(fs))
  no_day <- filter(out, is.na(day))
  if (nrow(no_day)) {
    warning(sprintf("Batch %s: skipping file(s) without a leading d<day>_: %s",
                    batch, paste(basename(no_day$file), collapse = ", ")), call. = FALSE)
  }
  out <- filter(out, !is.na(day))
  dups <- out |> group_by(day) |> filter(n() > 1) |> ungroup()
  if (nrow(dups)) {
    kept <- dups |> group_by(day) |> slice_max(mtime, n = 1, with_ties = FALSE) |> ungroup()
    warning(sprintf(
      "Batch %s: multiple workbooks for day(s) %s; using the most recent: %s",
      batch, paste(sort(unique(dups$day)), collapse = ", "),
      paste(basename(kept$file), collapse = ", ")), call. = FALSE)
  }
  out |>
    group_by(day) |>
    slice_max(mtime, n = 1, with_ties = FALSE) |>
    ungroup() |>
    arrange(day) |>
    select(file, day)
}

read_sheet_checked <- function(file, sheet, batch) {
  sheets <- excel_sheets(file)
  if (!sheet %in% sheets) {
    stop(sprintf("Batch %s: sheet '%s' not found in %s (sheets: %s)",
                 batch, sheet, basename(file), paste(sheets, collapse = ", ")),
         call. = FALSE)
  }
  suppressMessages(read_excel(file, sheet = sheet))
}

rename_required <- function(df, mapping, file, batch) {
  miss <- setdiff(mapping, names(df))
  if (length(miss)) {
    stop(sprintf("Batch %s: %s is missing column(s): %s",
                 batch, basename(file), paste(miss, collapse = ", ")), call. = FALSE)
  }
  df |> select(all_of(mapping))  # select() with a named vector renames too
}

# Conditions -------------------------------------------------------------------

#' Map raw condition labels (clone names etc.) to display labels and set the
#' plotting order. `conditions` is a named list: display label -> raw labels.
#' Raw labels not in the map are kept as-is, with a warning.
recode_conditions <- function(x, conditions) {
  lookup <- unlist(lapply(names(conditions), function(nm) {
    setNames(rep(nm, length(conditions[[nm]]) + 1), c(nm, conditions[[nm]]))
  }))
  x <- as.character(x)
  unmapped <- setdiff(unique(x[!is.na(x)]), names(lookup))
  if (length(unmapped)) {
    warning("Condition label(s) not in config `conditions`, kept as-is: ",
            paste(unmapped, collapse = ", "), call. = FALSE)
  }
  out <- ifelse(x %in% names(lookup), lookup[x], x)
  factor(out, levels = unique(c(names(conditions), sort(unmapped))))
}

day_label <- function(day) {
  d <- sort(unique(day))
  factor(paste0("Day ", day), levels = paste0("Day ", d))
}

# Importers ------------------------------------------------------------------

#' Per-condition averaged traces ("Representative Traces"; for fatigue the
#' "Condition Avg Fatigue" sheet) for every non-excluded batch.
import_traces <- function(batch_csv, analysis, conditions, batches = "all") {
  analysis <- tolower(analysis)
  b <- read_batch_sheet(batch_csv, analysis, batches)
  if (nrow(b) == 0) stop("No non-excluded '", analysis, "' rows in ", batch_csv,
                         " for the requested batches.", call. = FALSE)
  sheet <- emt_sheets[[analysis]]$traces
  mapping <- if (analysis == "fatigue") fatigue_trace_columns else trace_columns

  pmap(list(b$directory, b$batch), function(dir, batch) {
    list_day_files(dir, batch) |>
      pmap(function(file, day) {
        read_sheet_checked(file, sheet, batch) |>
          rename_required(mapping, file, batch) |>
          mutate(day = day, batch = batch)
      }) |>
      list_rbind()
  }) |>
    list_rbind() |>
    mutate(condition   = recode_conditions(condition, conditions),
           formal_date = day_label(day))
}

stim_columns <- c("Stim Peak Time", "Stim Number")

#' One value per well (x Hz) per metric: metrics measured on each stim (e.g. the
#' three 75 Hz tetani in "Relaxation Summary") are averaged across stims.
#' Pre-computed "... Well Avg" columns are dropped: they repeat the same rounded
#' value on every stim row and only exist for some metrics, so every metric is
#' averaged here the same way instead. Sheets with one row per well (x Hz),
#' like FF and fatigue, pass through unchanged.
well_average <- function(long) {
  long |>
    filter(!str_detect(metric, regex("well avg$", ignore_case = TRUE))) |>
    group_by(across(any_of(setdiff(metric_id_columns, stim_columns))), metric) |>
    summarise(value = mean(value, na.rm = TRUE), .groups = "drop") |>
    mutate(value = if_else(is.nan(value), NA_real_, value))
}

#' Per-well metrics ("Relaxation Summary" / "Fatigue Summary"), long format:
#' one row per well (x Hz) x metric, averaged across stims (see well_average()).
import_metrics <- function(batch_csv, analysis, conditions, batches = "all") {
  analysis <- tolower(analysis)
  b <- read_batch_sheet(batch_csv, analysis, batches)
  if (nrow(b) == 0) stop("No non-excluded '", analysis, "' rows in ", batch_csv,
                         " for the requested batches.", call. = FALSE)
  sheet <- emt_sheets[[analysis]]$metrics

  pmap(list(b$directory, b$batch), function(dir, batch) {
    list_day_files(dir, batch) |>
      pmap(function(file, day) {
        read_sheet_checked(file, sheet, batch) |>
          mutate(file = basename(file), day = day, batch = batch) |>
          pivot_longer(cols = where(is.numeric) & !any_of(metric_id_columns),
                       names_to = "metric", values_to = "value") |>
          select(any_of(metric_id_columns), metric, value) |>
          well_average()
      }) |>
      list_rbind()
  }) |>
    list_rbind() |>
    mutate(Condition   = recode_conditions(Condition, conditions),
           formal_date = day_label(day))
}

#' QC: wells whose baseline-subtracted force is <= 0, which usually means the
#' upstream peak detection failed for that file (and makes any normalized
#' metric, e.g. FF Normalized (%), meaningless for that well).
qc_nonpositive_force <- function(metrics, analysis,
                                 force_metric = c("Max tetanic-BK (μN)", "T1 BK Force (μN)")) {
  if (is.null(metrics)) return(NULL)
  bad <- filter(metrics, metric %in% force_metric, value <= 0)
  if (!"Hz" %in% names(bad)) bad$Hz <- NA_real_
  bad |>
    group_by(batch, day, Condition) |>
    summarise(wells = n_distinct(Plate, Well),
              Hz = paste(sort(unique(na.omit(Hz))), collapse = ", "),
              .groups = "drop") |>
    mutate(analysis = analysis, .before = 1)
}

#' Sample-size table: wells per condition x day x batch.
well_counts <- function(metrics) {
  metrics |>
    distinct(batch, day, Condition, Plate, Well) |>
    count(batch, day, Condition, name = "wells") |>
    pivot_wider(names_from = Condition, values_from = wells, values_fill = 0) |>
    arrange(batch, day)
}
