# Render plate-report.qmd: raw time-force traces in plate layout ----------------
# Usage (from this folder):
#   Rscript render-plates.R "/path/to/2026-11-05_R672C-EMT-Batch3"   # a batch folder
#   Rscript render-plates.R configs/plates/my-batch.yml              # a config
#   Rscript render-plates.R                                           # every configs/plates/*.yml
# Or in RStudio: source("render-plates.R"), then render_plate_report("<folder or yml>")
#
# A batch folder needs exactly one *Plate Map*.csv; time-force CSVs are read
# from its timeforce/ subfolder (any depth), or from the folder itself.
# The report lands at <output_dir>/<name>-plate-traces.html, with PNGs in
# <output_dir>/plate-traces/<protocol>/.

source("render-report.R")  # for render_qmd(), move_file()

render_plate_report <- function(config) {
  suppressPackageStartupMessages(library(tidyverse))
  source("R/emt-config.R", local = TRUE)  # merge_defaults()
  source("R/emt-import.R", local = TRUE)  # recode_conditions()
  source("R/plate-data.R", local = TRUE)
  cfg <- load_plate_config(config)
  html <- paste0(gsub("[^A-Za-z0-9]+", "-", tolower(cfg$name)), "-plate-traces.html")
  render_qmd("plate-report.qmd", config, html)
  dir.create(cfg$output_dir, recursive = TRUE, showWarnings = FALSE)
  dest <- file.path(cfg$output_dir, html)
  move_file(html, dest)
  message("Report: ", normalizePath(dest))
  invisible(dest)
}

if (sys.nframe() == 0) {
  args <- commandArgs(trailingOnly = TRUE)
  configs <- if (length(args)) args else list.files("configs/plates", "\\.ya?ml$", full.names = TRUE)
  if (!length(configs)) stop("Give a batch folder or config, or add configs to configs/plates/.")
  for (cf in configs) render_plate_report(cf)
}
