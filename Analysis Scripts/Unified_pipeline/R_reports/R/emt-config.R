# Config loading for the EMT report ------------------------------------------
# Every dataset-specific choice lives in a YAML file under configs/.
# This file fills defaults and validates so the rest of the code can trust it.

emt_config_defaults <- list(
  name            = "EMT dataset",
  output_dir      = "output",
  days_max        = Inf,
  trace_batches   = "all",
  metric_batches  = "all",
  colors          = NULL,
  labels          = list(condition = "Genotype", batch = "Differentiation"),
  kinetics        = list(stim_time_s = 48, window_s = 3, baseline_s = 0.01),
  kinetics_metrics = list(),
  ff_metric       = list(name = "FF Normalized (%)", label = "Percent of 100Hz Tension (%)"),
  fatigue_metrics = list(),
  stats           = list(unit = "batch", paired = FALSE, p_adjust = "BH"),
  sweep           = TRUE,
  seed            = 1,
  panel_sizes     = list(
    traces  = c(10, 2.5),
    metric  = c(4, 4),
    ff      = c(14, 3.5),
    fatigue = c(10, 2.5),
    sweep   = c(8, 6)
  )
)

# Recursive merge: values in `x` override `defaults`, nested lists merged.
merge_defaults <- function(x, defaults) {
  for (nm in names(defaults)) {
    if (is.null(x[[nm]])) {
      x[[nm]] <- defaults[[nm]]
    } else if (is.list(defaults[[nm]]) && is.list(x[[nm]]) && !is.null(names(defaults[[nm]]))) {
      x[[nm]] <- merge_defaults(x[[nm]], defaults[[nm]])
    }
  }
  x
}

load_emt_config <- function(path) {
  if (!file.exists(path)) stop("Config file not found: ", path, call. = FALSE)
  cfg <- merge_defaults(yaml::read_yaml(path), emt_config_defaults)

  if (is.null(cfg$batch_csv)) stop("Config must set `batch_csv`.", call. = FALSE)
  if (!file.exists(cfg$batch_csv)) {
    stop("batch_csv not found: ", cfg$batch_csv, call. = FALSE)
  }
  if (is.null(cfg$conditions) || length(cfg$conditions) == 0) {
    stop("Config must set `conditions` (display label -> list of raw labels).", call. = FALSE)
  }

  # Colors: named by display label; fill any missing ones from a colorblind-safe set
  okabe_ito <- c("#000000", "#009E73", "#E69F00", "#D55E00", "#CC79A7",
                 "#0072B2", "#56B4E9", "#F0E442")
  cond_names <- names(cfg$conditions)
  cols <- unlist(cfg$colors)
  missing_cols <- setdiff(cond_names, names(cols))
  if (length(missing_cols) > 0) {
    spare <- setdiff(okabe_ito, cols)
    cols <- c(cols, setNames(spare[seq_along(missing_cols)], missing_cols))
  }
  cfg$colors <- cols[cond_names]

  cfg$days_max <- as.numeric(cfg$days_max)
  cfg$panel_sizes <- lapply(cfg$panel_sizes, as.numeric)
  cfg
}

# "all" (or NULL) means no filter; otherwise a numeric vector of batch numbers
resolve_batches <- function(x, available) {
  if (is.null(x) || identical(x, "all")) return(sort(unique(available)))
  as.numeric(unlist(x))
}
