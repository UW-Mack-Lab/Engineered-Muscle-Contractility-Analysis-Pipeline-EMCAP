# One-shot report rendering for the Master Analysis Launcher -------------------
# Called by Python_analyze/Master_Analysis_Launcher.py after the analyses, but
# works on its own too. Run from this folder:
#
#   Rscript render-pipeline.R \
#     --timeforce-dir "/path/batch/timeforce" --plate-map "/path/batch/X Plate Map.csv" \
#     --analyzed-dir "/path/batch/analyzed" --reports-dir "/path/batch/reports" \
#     --name "R672C Batch 3" --plates --emt-batch \
#     [--dataset-config configs/r672c.yml --batch-number 3]
#
# Steps (each optional, each independent; one failing doesn't stop the others):
#   --plates          plate trace report on the raw time-force files
#                     -> <reports-dir>/plates/
#   --emt-batch       EMT report on this batch's analyzed workbooks only
#                     (stats use unit: well) -> <reports-dir>/emt/
#   --dataset-config  add/update this batch's rows (batch = --batch-number) in
#                     that config's batch CSV, then re-render the dataset report
#                     -> wherever that config's output_dir points
#
# The configs and batch CSV generated for the first two steps are kept in
# <reports-dir>/_configs/, so you can re-render by hand with render-plates.R /
# render-report.R. Each step prints one line:
#   PIPELINE_RESULT <step> OK <html path>
#   PIPELINE_RESULT <step> FAILED <message>
# and the exit status is non-zero if any step failed.

suppressPackageStartupMessages(library(tidyverse))
source("render-plates.R")      # render_plate_report(); also sources render-report.R
source("R/emt-config.R")       # merge_defaults()
source("R/emt-import.R")       # recode_conditions()
source("R/plate-data.R")       # read_plate_map()

# Analyzed subfolder (as the launcher names it) -> analysis type in the batch CSV
analyzed_subfolders <- c(`75Hz` = "75hz", FF = "ff", Fatigue = "fatigue")
single_batch_defaults <- "configs/single-batch-defaults.yml"

parse_args <- function(args) {
  flags <- c("plates", "emt-batch")
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

# Steps ----------------------------------------------------------------------

step_plates <- function(opt, dataset_cfg) {
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

step_emt_batch <- function(opt, dataset_cfg) {
  found <- find_analyzed(opt$`analyzed-dir`)
  cfg_dir <- file.path(opt$`reports-dir`, "_configs")
  dir.create(cfg_dir, recursive = TRUE, showWarnings = FALSE)
  batch_csv <- file.path(cfg_dir, "batch.csv")
  write.csv(data.frame(batch = 1, analysis = names(found), directory = unname(found), exclude = 0),
            batch_csv, row.names = FALSE)

  base <- dataset_cfg %||% yaml::read_yaml(single_batch_defaults)
  if (length(base$conditions) == 0) {
    base$conditions <- conditions_from_plate_map(opt$`plate-map`)
  }
  base$name <- opt$name
  base$batch_csv <- batch_csv
  base$output_dir <- file.path(opt$`reports-dir`, "emt")
  base$trace_batches <- list(1)
  base$metric_batches <- "all"
  # One batch: batch-level stats would all be n < 2. Wells as replicates is
  # pseudo-replicated, so treat these p-values as exploratory only.
  base$stats <- list(unit = "well", paired = FALSE, p_adjust = base$stats$p_adjust %||% "BH")
  path <- write_config(base, file.path(cfg_dir, "emt.yml"))
  render_emt_report(path)
}

#' Add or update this batch's rows in the dataset's batch CSV, then render it.
step_dataset <- function(opt, dataset_cfg) {
  batch_no <- suppressWarnings(as.numeric(opt$`batch-number`))
  if (length(batch_no) != 1 || is.na(batch_no)) {
    stop("--batch-number must be a number when --dataset-config is given.", call. = FALSE)
  }
  csv <- dataset_cfg$batch_csv
  if (is.null(csv)) stop("Dataset config has no `batch_csv`.", call. = FALSE)
  found <- find_analyzed(opt$`analyzed-dir`)

  old <- if (file.exists(csv)) {
    read.csv(csv, stringsAsFactors = FALSE, check.names = FALSE, colClasses = "character")
  } else {
    data.frame(batch = character(), analysis = character(),
               directory = character(), exclude = character())
  }
  if (!"exclude" %in% names(old)) old$exclude <- "0"
  new <- old
  for (a in names(found)) {
    hit <- suppressWarnings(as.numeric(new$batch)) == batch_no &
      tolower(trimws(new$analysis)) == a
    hit[is.na(hit)] <- FALSE
    if (any(hit)) {
      new$directory[hit] <- found[[a]]
    } else {
      row <- setNames(as.list(rep("", ncol(new))), names(new))
      row[c("batch", "analysis", "directory", "exclude")] <-
        list(format(batch_no), a, found[[a]], "0")
      new <- rbind(new, as.data.frame(row, check.names = FALSE, stringsAsFactors = FALSE))
    }
  }

  if (!identical(new, old)) {
    if (file.exists(csv)) {
      bak <- paste0(csv, ".bak-", format(Sys.time(), "%Y%m%d_%H%M%S"))
      file.copy(csv, bak)
      message("Backed up batch CSV: ", normalizePath(bak))
    }
    write.csv(new, csv, row.names = FALSE, na = "")
    message("Updated batch ", format(batch_no), " in ", normalizePath(csv))
  } else {
    message("Batch CSV already up to date: ", normalizePath(csv))
  }
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
  for (k in c("reports-dir", "name")) {
    if (is.null(opt[[k]])) stop("--", k, " is required.", call. = FALSE)
  }
  for (k in c("timeforce-dir", "plate-map", "analyzed-dir", "reports-dir")) {
    if (!is.null(opt[[k]]) && file.exists(opt[[k]])) opt[[k]] <- normalizePath(opt[[k]])
  }
  dataset_cfg <- if (!is.null(opt$`dataset-config`)) yaml::read_yaml(opt$`dataset-config`)

  ok <- logical()
  if (isTRUE(opt$plates))      ok <- c(ok, run_step("plates", step_plates, opt, dataset_cfg))
  if (isTRUE(opt$`emt-batch`)) ok <- c(ok, run_step("emt-batch", step_emt_batch, opt, dataset_cfg))
  if (!is.null(dataset_cfg))   ok <- c(ok, run_step("dataset", step_dataset, opt, dataset_cfg))
  if (!length(ok)) message("Nothing to do: give --plates, --emt-batch and/or --dataset-config.")
  quit(status = if (all(ok)) 0 else 1)
}
