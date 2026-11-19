# Render emt-report.qmd once per config -----------------------------------------
# Usage (from this folder):
#   Rscript render-report.R                       # every configs/*.yml
#   Rscript render-report.R configs/mydataset.yml # just one
# Or in RStudio: source("render-report.R"), then render_emt_report("configs/mydataset.yml")
#
# Each report lands at <output_dir>/<name>-emt-report.html alongside the
# panels/, all-metrics/ and stats/ folders.

find_quarto <- function() {
  # QUARTO_PATH is set by the Master Analysis Launcher. Run by hand from this
  # folder, the copy `uv sync` installed into ../.venv (the quarto-cli Python
  # package) is used next; the rest are standalone installs.
  candidates <- c(Sys.getenv("QUARTO_PATH"),
                  file.path("..", ".venv", "bin", "quarto"),
                  file.path("..", ".venv", "Scripts", "quarto.exe"),
                  Sys.which("quarto"),
                  "/Applications/quarto/bin/quarto",
                  "/Applications/RStudio.app/Contents/Resources/app/quarto/bin/quarto",
                  file.path(Sys.getenv("LOCALAPPDATA"), "Programs", "Quarto", "bin", "quarto.exe"),
                  "C:/Program Files/Quarto/bin/quarto.exe",
                  "C:/Program Files/RStudio/resources/app/bin/quarto/bin/quarto.exe")
  hit <- candidates[nzchar(candidates) & file.exists(candidates)]
  if (!length(hit)) {
    stop("Quarto not found. Run `uv sync` in the project folder, install Quarto ",
         "from https://quarto.org, or set QUARTO_PATH.")
  }
  normalizePath(hit[[1]])
}

#' Render a .qmd with `config` as its parameter. Quarto is told to run the
#' chunks with this same R, so they see the same (renv) package library.
render_qmd <- function(qmd, config, html) {
  rscript <- file.path(R.home("bin"), if (.Platform$OS.type == "windows") "Rscript.exe" else "Rscript")
  old <- Sys.getenv("QUARTO_R", unset = NA)
  Sys.setenv(QUARTO_R = rscript)
  on.exit(if (is.na(old)) Sys.unsetenv("QUARTO_R") else Sys.setenv(QUARTO_R = old), add = TRUE)
  status <- system2(find_quarto(),
                    c("render", qmd,
                      "-P", shQuote(paste0("config:", config)),
                      "--output", html))
  if (status != 0) stop("Quarto render failed for ", config)
  invisible(html)
}

# file.rename() can't move across volumes (external drive, some cloud folders),
# so fall back to copy + delete.
move_file <- function(from, to) {
  if (file.exists(to)) unlink(to)
  if (!suppressWarnings(file.rename(from, to))) {
    if (!file.copy(from, to, overwrite = TRUE)) stop("Could not move ", from, " to ", to)
    unlink(from)
  }
  invisible(to)
}

render_emt_report <- function(config) {
  source("R/emt-config.R", local = TRUE)
  cfg <- load_emt_config(config)
  html <- paste0(gsub("[^A-Za-z0-9]+", "-", tolower(cfg$name)), "-emt-report.html")
  render_qmd("emt-report.qmd", config, html)
  dir.create(cfg$output_dir, recursive = TRUE, showWarnings = FALSE)
  dest <- file.path(cfg$output_dir, html)
  move_file(html, dest)
  message("Report: ", normalizePath(dest))
  invisible(dest)
}

if (sys.nframe() == 0) {
  args <- commandArgs(trailingOnly = TRUE)
  configs <- if (length(args)) args else
    setdiff(list.files("configs", "\\.ya?ml$", full.names = TRUE),
            "configs/single-batch-defaults.yml")
  if (!length(configs)) stop("Give a config, or add dataset configs to configs/ (see configs/examples/).")
  for (cf in configs) render_emt_report(cf)
}
