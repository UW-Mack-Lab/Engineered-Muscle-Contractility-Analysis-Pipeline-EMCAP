# Load a whole dataset in one call, for interactive use -------------------------
# The report does the same steps inline in emt-report.qmd's setup chunk.

#' Read the config and every analysis it points to.
#'
#' Returns a list:
#'   cfg            the loaded config (colors, conditions, labels, ...)
#'   batch_sheet    non-excluded rows of the batch CSV
#'   metric_batches / trace_batches   resolved batch numbers
#'   m75, mff, mfat           per-well metrics, long format (one row per well x metric)
#'   traces75, fat_tr         condition-average traces
#'   sp_args        the arguments every superplot needs, for do.call()
#' Analyses missing from the batch CSV come back as NULL.
#'
#' `metric_batches` / `trace_batches` override the config ("all" or a vector).
load_emt_data <- function(config,
                          metric_batches = NULL, trace_batches = NULL,
                          traces = TRUE) {
  cfg <- load_emt_config(config)
  if (!is.null(metric_batches)) cfg$metric_batches <- metric_batches
  if (!is.null(trace_batches))  cfg$trace_batches  <- trace_batches

  batch_sheet <- read_batch_sheet(cfg$batch_csv)
  mb <- resolve_batches(cfg$metric_batches, batch_sheet$batch)
  tb <- resolve_batches(cfg$trace_batches,  batch_sheet$batch)
  has <- function(a) any(batch_sheet$analysis == a)

  step <- function(label, a, expr) {
    if (!has(a)) return(NULL)
    message("Reading ", label, " ...")
    expr
  }

  out <- list(
    cfg = cfg,
    batch_sheet = batch_sheet,
    metric_batches = mb,
    trace_batches = tb,
    m75  = step("75 Hz metrics",   "75hz",    import_metrics(cfg$batch_csv, "75hz",    cfg$conditions, mb)),
    mff  = step("FF metrics",      "ff",      import_metrics(cfg$batch_csv, "ff",      cfg$conditions, mb)),
    mfat = step("fatigue metrics", "fatigue", import_metrics(cfg$batch_csv, "fatigue", cfg$conditions, mb)),
    traces75 = if (traces) step("75 Hz traces",   "75hz",    import_traces(cfg$batch_csv, "75hz",    cfg$conditions, tb)),
    fat_tr   = if (traces) step("fatigue traces", "fatigue", import_traces(cfg$batch_csv, "fatigue", cfg$conditions, tb)),
    sp_args = list(colors = cfg$colors, days_max = cfg$days_max,
                   batches = mb, labels = cfg$labels, seed = cfg$seed)
  )
  message("Done: ", cfg$name, " (metric batches ", paste(mb, collapse = ", "),
          "; trace batches ", paste(tb, collapse = ", "), ")")
  out
}

#' Superplot one metric with the dataset's colors/labels/batches.
#' `...` overrides any of them, e.g. days_max = Inf, batches = 2, y_label = "...".
emt_superplot <- function(emt, metrics, metric_name, ...) {
  args <- modifyList(emt$sp_args, list(...))
  do.call(metric_superplot, c(list(metrics, metric_name), args))
}

#' Same for force-frequency metrics (x = Hz, faceted by day).
emt_ff_superplot <- function(emt, metric_name, ...) {
  args <- modifyList(emt$sp_args, list(...))
  do.call(ff_superplot, c(list(emt$mff, metric_name), args))
}
