# One-shot report rendering for the Master Analysis Launcher -------------------
# Called by Python_analyze/Master_Analysis_Launcher.py, but works on its own too.
# Run from this folder. One batch (the launcher's "Run batch"):
#
#   Rscript render-pipeline.R \
#     --timeforce-dir "/path/batch/timeforce" --plate-map "/path/batch/X Plate Map.csv" \
#     --analyzed-dir "/path/batch/analyzed" --reports-dir "/path/batch/reports" \
#     --name "R672C Batch 3" --plates --emt-batch \
#     [--dataset-config configs/r672c.yml --register --batch-number 3 [--plate A]]
#
# A dataset (the launcher's "Run dataset reports"):
#
#   Rscript render-pipeline.R --dataset-config configs/r672c.yml --dataset-emt --dataset-plates
#
# Steps (each optional, each independent; one failing doesn't stop the others):
#   --plates          plate trace report on the raw time-force files
#                     -> <reports-dir>/plates/
#   --emt-batch       EMT report on this batch's analyzed workbooks only
#                     (stats use unit: well) -> <reports-dir>/emt/
#   --register        add/update this batch's rows (batch = --batch-number, and
#                     plate = --plate when one batch was run on several plates)
#                     in the dataset config's batch CSV: its analyzed
#                     75Hz/FF/Fatigue folders plus a `timeforce` row (raw folder +
#                     plate map). Renders nothing itself. With --plate, --emt-batch
#                     then covers every registered plate of the batch (one report
#                     per batch, in <batch folder>/reports/emt/).
#   --dataset-emt     dataset EMT report -> the config's output_dir
#   --dataset-plates  combined plate trace report of every batch with a
#                     `timeforce` row -> <output_dir>/plates/
# --dataset-config on its own only lends its conditions and colors to the
# --plates / --emt-batch reports, so labels match the dataset report.
#
# The configs and batch CSV generated for the batch steps are kept in
# <reports-dir>/_configs/ (dataset plates: <output_dir>/_configs/plates.yml), so
# you can re-render by hand with render-plates.R / render-report.R. Each step
# prints one line:
#   PIPELINE_RESULT <step> OK <html or csv path>
#   PIPELINE_RESULT <step> FAILED <message>
# and the exit status is non-zero if any step failed.

suppressPackageStartupMessages(library(tidyverse))
source("render-plates.R")      # render_plate_report(); also sources render-report.R
source("R/emt-config.R")       # merge_defaults()
source("R/emt-import.R")       # recode_conditions(), read_batch_sheet()
source("R/plate-data.R")       # read_plate_map(), find_plate_map()

# Analyzed subfolder (as the launcher names it) -> analysis type in the batch CSV
analyzed_subfolders <- c(`75Hz` = "75hz", FF = "ff", Fatigue = "fatigue")
single_batch_defaults <- "configs/single-batch-defaults.yml"
# Combined plate reports embed every batch's PNGs; this keeps the HTML manageable
dataset_plates_dpi <- 150

parse_args <- function(args) {
  flags <- c("plates", "emt-batch", "register", "dataset-emt", "dataset-plates")
  out <- list()
  i <- 1
  while (i <= length(args)) {
    a <- args[[i]]
    if (!startsWith(a, "--")) stop("Unexpected argument: ", a, call. = FALSE)
    key <- substring(a, 3)
    if (key %in% flags) {
      out[[key]] <- TRUE
      i <- i + 1
    } else {
      if (i == length(args)) stop("Missing value for ", a, call. = FALSE)
      out[[key]] <- args[[i + 1]]
      i <- i + 2
    }
  }
  out
}

require_args <- function(opt, keys, step) {
  miss <- keys[vapply(keys, \(k) is.null(opt[[k]]), logical(1))]
  if (length(miss)) stop(step, " needs ", paste0("--", miss, collapse = ", "), call. = FALSE)
}

write_config <- function(x, path) {
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  yaml::write_yaml(x, path, fileEncoding = "UTF-8")
  path
}

#' Analysis types with workbooks in the analyzed folder: named vector
#' analysis type -> directory, e.g. c(`75hz` = ".../analyzed/75Hz/")
find_analyzed <- function(analyzed_dir) {
  dirs <- file.path(analyzed_dir, names(analyzed_subfolders))
  has_xlsx <- vapply(dirs, function(d) {
    dir.exists(d) && any(!startsWith(list.files(d, "\\.xlsx$"), "~$"))
  }, logical(1))
  if (!any(has_xlsx)) {
    stop("No analyzed 75Hz, FF or Fatigue workbooks in ", analyzed_dir, call. = FALSE)
  }
  setNames(paste0(dirs[has_xlsx], "/"), analyzed_subfolders[has_xlsx])
}

#' Plate-map labels as an identity `conditions` list, in plate order.
#' The first label becomes the reference condition for stats.
conditions_from_plate_map <- function(plate_map) {
  labs <- unique(na.omit(read_plate_map(plate_map)$label))
  setNames(lapply(labs, list), labs)
}

# Batch steps ------------------------------------------------------------------

step_plates <- function(opt, dataset_cfg) {
  require_args(opt, c("timeforce-dir", "plate-map"), "--plates")
  cfg <- list(name = opt$name,
              timeforce_dir = opt$`timeforce-dir`,
              plate_map = opt$`plate-map`,
              output_dir = file.path(opt$`reports-dir`, "plates"))
  if (!is.null(dataset_cfg)) {
    cfg$conditions <- dataset_cfg$conditions
    cfg$colors <- dataset_cfg$colors
  }
  path <- write_config(cfg, file.path(opt$`reports-dir`, "_configs", "plate-traces.yml"))
  render_plate_report(path)
}

#' Every registered plate of this batch (--batch-number) in the dataset's batch
#' CSV: the analyzed rows, the plate maps, and where the batch report goes —
#' <batch folder>/reports when all plate folders share a parent (the batch
#' folder, e.g. .../Batch-4/PlateA and .../Batch-4/PlateB), otherwise
#' <dataset output_dir>/batch-<n>.
batch_plates <- function(opt, dataset_cfg) {
  batch_no <- as.numeric(opt$`batch-number`)
  sheet <- read_batch_sheet(dataset_cfg$batch_csv)
  sheet <- sheet[as.numeric(sheet$batch) == batch_no, ]
  rows <- sheet[sheet$analysis %in% analyzed_subfolders, ]
  rows <- rows[order(rows$plate, rows$analysis), ]
  if (nrow(rows) == 0) stop("No analyzed rows for batch ", batch_no, " in ", dataset_cfg$batch_csv, call. = FALSE)

  tf <- sheet[sheet$analysis == "timeforce", ]
  tf <- tf[order(tf$plate), ]
  pm <- if ("plate_map" %in% names(tf)) as.character(tf$plate_map) else rep(NA, nrow(tf))
  plate_maps <- vapply(seq_len(nrow(tf)), function(i) {
    if (!is.na(pm[i]) && nzchar(trimws(pm[i]))) pm[i] else find_plate_map(tf$directory[i])$path
  }, character(1))

  batch_dirs <- unique(dirname(vapply(rows$directory, plate_folder, character(1))))
  if (length(batch_dirs) == 1) {
    reports_dir <- file.path(batch_dirs, "reports")
    name <- basename(batch_dirs)
  } else {
    out <- dataset_cfg$output_dir %||% "output"
    reports_dir <- file.path(out, paste0("batch-", format(batch_no)))
    name <- paste(dataset_cfg$name %||% "Dataset", "Batch", format(batch_no))
    message("Plates of batch ", format(batch_no), " are in different parent folders; ",
            "the batch report goes to ", reports_dir)
  }
  list(rows = rows, plate_maps = plate_maps, reports_dir = reports_dir, name = name,
       plates = unique(rows$plate))
}

step_emt_batch <- function(opt, dataset_cfg) {
  require_args(opt, "analyzed-dir", "--emt-batch")
  # A batch run on several plates: one report for the whole batch, from every
  # plate registered so far (this one was registered just before).
  per_plate <- nzchar(plate_id(opt[["plate"]])) && isTRUE(opt$register) && !is.null(dataset_cfg)
  if (per_plate) {
    bp <- batch_plates(opt, dataset_cfg)
    reports_dir <- bp$reports_dir
    name <- bp$name
    rows <- data.frame(batch = 1, plate = bp$rows$plate, analysis = bp$rows$analysis,
                       directory = bp$rows$directory, exclude = 0)
    plate_maps <- bp$plate_maps
    message("Batch report covers plate(s): ", paste(bp$plates, collapse = ", "))
  } else {
    reports_dir <- opt$`reports-dir`
    name <- opt$name
    found <- find_analyzed(opt$`analyzed-dir`)
    rows <- data.frame(batch = 1, analysis = names(found), directory = unname(found), exclude = 0)
    plate_maps <- opt$`plate-map`
  }
  # Absolute, so the generated config's paths don't depend on where it sits
  cfg_dir <- normalizePath(file.path(reports_dir, "_configs"), mustWork = FALSE)
  dir.create(cfg_dir, recursive = TRUE, showWarnings = FALSE)
  batch_csv <- file.path(cfg_dir, "batch.csv")
  write.csv(rows, batch_csv, row.names = FALSE)

  base <- dataset_cfg %||% yaml::read_yaml(single_batch_defaults)
  if (length(base$conditions) == 0) {
    conds <- lapply(plate_maps, conditions_from_plate_map)
    base$conditions <- do.call(c, conds)[!duplicated(unlist(lapply(conds, names)))]
  }
  base$name <- name
  base$batch_csv <- batch_csv
  base$output_dir <- file.path(dirname(cfg_dir), "emt")
  base$trace_batches <- list(1)
  base$metric_batches <- "all"
  base$days_max <- NULL   # a batch report shows every recorded day of that batch
  base$plates <- NULL
  # One batch: batch-level stats would all be n < 2. Wells as replicates is
  # pseudo-replicated, so treat these p-values as exploratory only.
  base$stats <- list(unit = "well", paired = FALSE, p_adjust = base$stats$p_adjust %||% "BH")
  path <- write_config(base, file.path(cfg_dir, "emt.yml"))
  render_emt_report(path)
}

# Dataset steps ----------------------------------------------------------------

read_dataset_csv <- function(csv) {
  out <- if (file.exists(csv)) {
    read.csv(csv, stringsAsFactors = FALSE, check.names = FALSE, colClasses = "character")
  } else {
    data.frame(batch = character(), analysis = character(),
               directory = character(), exclude = character())
  }
  if (!"exclude" %in% names(out)) out$exclude <- rep("0", nrow(out))
  out
}

#' "" for NULL/NA/blank, else the trimmed plate id
plate_id <- function(x) {
  x <- if (is.null(x)) "" else trimws(as.character(x))
  ifelse(is.na(x), "", x)
}

#' The folder holding a plate's data, from one of its analyzed/<type> folders:
#'   .../Batch-4/PlateA/analyzed/FF/  -> .../Batch-4/PlateA  (analyzed/ inside the plate folder)
#'   .../Expt/analyzed/dmd/FF/        -> .../Expt/dmd        (analyzed/ next to the plate folders)
plate_folder <- function(dir) {
  d <- sub("[/\\\\]+$", "", dir)
  up <- dirname(d)
  if (tolower(basename(up)) == "analyzed") return(dirname(up))
  if (tolower(basename(dirname(up))) == "analyzed") {
    return(file.path(dirname(dirname(up)), basename(up)))
  }
  dirname(d)
}

#' Add or update this batch's rows in the dataset's batch CSV: one per analyzed
#' folder, plus a `timeforce` row (raw folder + plate map) for the combined plate
#' report. Existing rows keep `exclude` and any extra columns. Renders nothing.
step_register <- function(opt, dataset_cfg) {
  require_args(opt, c("batch-number", "analyzed-dir", "timeforce-dir", "plate-map"), "--register")
  batch_no <- suppressWarnings(as.numeric(opt$`batch-number`))
  if (length(batch_no) != 1 || is.na(batch_no)) {
    stop("--batch-number must be a number.", call. = FALSE)
  }
  plate <- plate_id(opt[["plate"]])
  csv <- dataset_cfg$batch_csv
  if (is.null(csv)) stop("Dataset config has no `batch_csv`.", call. = FALSE)

  dirs <- c(find_analyzed(opt$`analyzed-dir`), timeforce = opt$`timeforce-dir`)
  plate_maps <- c(rep("", length(dirs) - 1), opt$`plate-map`)

  old <- read_dataset_csv(csv)
  new <- old
  if (!"plate_map" %in% names(new)) new$plate_map <- rep("", nrow(new))
  if (nzchar(plate) && !"plate" %in% names(new)) new$plate <- rep("", nrow(new))
  row_plate <- if ("plate" %in% names(new)) plate_id(new$plate) else rep("", nrow(new))
  same_batch <- suppressWarnings(as.numeric(new$batch)) == batch_no
  same_batch[is.na(same_batch)] <- FALSE

  # Plates of one batch each need their own analyzed/ folder; two plates whose
  # data folders share a parent would overwrite each other's workbooks.
  clash <- same_batch & row_plate != plate &
    normalizePath(new$directory, mustWork = FALSE) %in% normalizePath(dirs, mustWork = FALSE)
  if (any(clash)) {
    stop("Batch ", format(batch_no), " plate '", row_plate[clash][1], "' already uses ",
         new$directory[clash][1], ". Give each plate its own folder (plate map + timeforce/) ",
         "and analyze it with its plate name filled in, so it gets its own ",
         "analyzed/<plate folder>/.", call. = FALSE)
  }
  if (!nzchar(plate) && any(same_batch & nzchar(row_plate))) {
    stop("Batch ", format(batch_no), " is registered per plate (",
         paste(unique(row_plate[same_batch & nzchar(row_plate)]), collapse = ", "),
         "). Give this plate's name too.", call. = FALSE)
  }

  for (i in seq_along(dirs)) {
    a <- names(dirs)[i]
    hit <- same_batch & tolower(trimws(new$analysis)) == a & row_plate == plate
    if (any(hit)) {
      new$directory[hit] <- dirs[[i]]
      new$plate_map[hit] <- plate_maps[i]
    } else {
      row <- setNames(as.list(rep("", ncol(new))), names(new))
      row[c("batch", "analysis", "directory", "exclude", "plate_map")] <-
        list(format(batch_no), a, dirs[[i]], "0", plate_maps[i])
      if ("plate" %in% names(new)) row$plate <- plate
      new <- rbind(new, as.data.frame(row, check.names = FALSE, stringsAsFactors = FALSE))
      row_plate <- c(row_plate, plate)
      same_batch <- c(same_batch, TRUE)
    }
  }

  if (!identical(new, old)) {
    if (file.exists(csv)) {
      bak <- paste0(csv, ".bak-", format(Sys.time(), "%Y%m%d_%H%M%S"))
      file.copy(csv, bak)
      message("Backed up batch CSV: ", normalizePath(bak))
    }
    dir.create(dirname(csv), recursive = TRUE, showWarnings = FALSE)
    write.csv(new, csv, row.names = FALSE, na = "")
    message("Updated batch ", format(batch_no), if (nzchar(plate)) paste0(" plate ", plate),
            " in ", normalizePath(csv))
  } else {
    message("Batch CSV already up to date: ", normalizePath(csv))
  }
  csv
}

#' Combined plate config for every non-excluded `timeforce` row of the dataset's
#' batch CSV, with the dataset's conditions and colors. Settings in an optional
#' `plates:` block of the dataset config (days, protocols, png_dpi, ...) override
#' the defaults. Returns the path of the written config.
dataset_plates_config <- function(dataset_config) {
  ds <- resolve_config_paths(yaml::read_yaml(dataset_config), dataset_config)
  csv <- ds$batch_csv
  if (is.null(csv) || !file.exists(csv)) {
    stop("Batch CSV not found: ", csv %||% "(no `batch_csv` in config)", call. = FALSE)
  }
  rows <- read_batch_sheet(csv, "timeforce")
  if (nrow(rows) == 0) {
    stop("No `timeforce` rows in ", csv, ". Register the batches from the launcher's ",
         "batch section (or add rows with analysis = timeforce, directory = the raw ",
         "time-force folder).", call. = FALSE)
  }
  rows <- rows[order(as.numeric(rows$batch), rows$plate), ]
  pm_col <- if ("plate_map" %in% names(rows)) as.character(rows$plate_map) else rep(NA, nrow(rows))
  batches <- lapply(seq_len(nrow(rows)), function(i) {
    tf <- sub("[/\\\\]+$", "", rows$directory[i])
    pm <- pm_col[i]
    if (is.na(pm) || !nzchar(trimws(pm))) pm <- find_plate_map(tf)$path
    name <- paste("Batch", format(as.numeric(rows$batch[i])))
    if (nzchar(rows$plate[i])) name <- paste0(name, " · ", rows$plate[i])
    list(name = name, timeforce_dir = tf, plate_map = pm)
  })

  out_dir <- ds$output_dir %||%
    file.path(config_base_dir(dataset_config), "output", gsub("[^A-Za-z0-9]+", "-", tolower(ds$name)))
  cfg <- list(name = ds$name %||% "Dataset",
              output_dir = file.path(out_dir, "plates"),
              conditions = ds$conditions,
              colors = ds$colors,
              png_dpi = dataset_plates_dpi)
  cfg <- utils::modifyList(cfg, ds$plates %||% list())
  cfg$batches <- batches
  write_config(cfg, file.path(out_dir, "_configs", "plates.yml"))
}

step_dataset_plates <- function(opt) {
  render_plate_report(dataset_plates_config(opt$`dataset-config`))
}

step_dataset_emt <- function(opt) {
  render_emt_report(opt$`dataset-config`)
}

# Main -------------------------------------------------------------------------

run_step <- function(step, fn, ...) {
  message("\n=== ", step, " ===")
  res <- tryCatch(list(ok = TRUE, value = fn(...)),
                  error = function(e) list(ok = FALSE, value = conditionMessage(e)))
  line <- if (res$ok) {
    paste("PIPELINE_RESULT", step, "OK", normalizePath(res$value))
  } else {
    paste("PIPELINE_RESULT", step, "FAILED", gsub("\\s*\n\\s*", " ", res$value))
  }
  cat(line, "\n", sep = "")
  res$ok
}

if (sys.nframe() == 0) {
  opt <- parse_args(commandArgs(trailingOnly = TRUE))
  batch_steps <- isTRUE(opt$plates) || isTRUE(opt$`emt-batch`)
  dataset_steps <- isTRUE(opt$register) || isTRUE(opt$`dataset-emt`) || isTRUE(opt$`dataset-plates`)
  if (batch_steps) {
    for (k in c("reports-dir", "name")) {
      if (is.null(opt[[k]])) stop("--", k, " is required for --plates / --emt-batch.", call. = FALSE)
    }
  }
  if (dataset_steps && is.null(opt$`dataset-config`)) {
    stop("--register, --dataset-emt and --dataset-plates need --dataset-config.", call. = FALSE)
  }
  for (k in c("timeforce-dir", "plate-map", "analyzed-dir", "reports-dir")) {
    if (!is.null(opt[[k]]) && file.exists(opt[[k]])) opt[[k]] <- normalizePath(opt[[k]])
  }
  dataset_cfg <- if (!is.null(opt$`dataset-config`)) {
    resolve_config_paths(yaml::read_yaml(opt$`dataset-config`), opt$`dataset-config`)
  }

  ok <- logical()
  # register first: a per-plate batch report reads the batch's plates from the CSV
  if (isTRUE(opt$register))         ok <- c(ok, run_step("register", step_register, opt, dataset_cfg))
  if (isTRUE(opt$plates))           ok <- c(ok, run_step("plates", step_plates, opt, dataset_cfg))
  if (isTRUE(opt$`emt-batch`))      ok <- c(ok, run_step("emt-batch", step_emt_batch, opt, dataset_cfg))
  if (isTRUE(opt$`dataset-emt`))    ok <- c(ok, run_step("dataset-emt", step_dataset_emt, opt))
  if (isTRUE(opt$`dataset-plates`)) ok <- c(ok, run_step("dataset-plates", step_dataset_plates, opt))
  if (!length(ok)) {
    message("Nothing to do: give --plates, --emt-batch, --register, --dataset-emt and/or --dataset-plates.")
  }
  quit(status = if (all(ok)) 0 else 1)
}
