# Parallel map for the plate reports ---------------------------------------------
# Reading time-force files and drawing plate plots are independent per file, so
# they run in worker processes: a socket cluster whose workers load the same
# packages and source this folder's R/ files. Not fork (mclapply): on macOS a
# forked child that draws text crashes in the system font libraries. Pass
# everything FUN needs as arguments (`...`): workers don't see the report's
# variables.

# More workers than this rarely helps (a report has ~5-100 plots) and each one
# holds its own copy of the plotting stack in memory.
max_report_workers <- 8L

#' Worker count: EMT_R_WORKERS (the launcher's worker budget), else one less than
#' the physical cores; capped at max_report_workers. 1 = run serially.
report_workers <- function() {
  n <- suppressWarnings(as.integer(Sys.getenv("EMT_R_WORKERS", "")))
  if (is.na(n) || n < 1) {
    cores <- parallel::detectCores(logical = FALSE)
    n <- if (is.na(cores)) 1L else max(1L, cores - 1L)
  }
  min(n, max_report_workers)
}

.par_state <- new.env()

par_cluster <- function(n) {
  cl <- .par_state$cl
  if (!is.null(cl) && length(cl) >= n) return(cl)
  par_stop()
  cl <- parallel::makePSOCKcluster(n)
  # Same folder and package library (renv) as this session
  parallel::clusterCall(cl, function(wd, libs) {
    setwd(wd)
    .libPaths(libs)
    suppressPackageStartupMessages({
      library(tidyverse)
      library(ggprism)
    })
    for (f in list.files("R", pattern = "\\.R$", full.names = TRUE)) source(f)
    NULL
  }, getwd(), .libPaths())
  .par_state$cl <- cl
  cl
}

#' Stop the worker cluster, if one was started.
par_stop <- function() {
  if (!is.null(.par_state$cl)) {
    try(parallel::stopCluster(.par_state$cl), silent = TRUE)
    .par_state$cl <- NULL
  }
  invisible(NULL)
}

#' lapply(X, FUN, ...) over worker processes, results in the order of X. Items
#' are handed out one at a time, so put the slowest ones first. FUN should catch
#' its own errors (return the condition) so one bad file doesn't stop the rest.
par_lapply <- function(X, FUN, ..., workers = report_workers()) {
  if (workers <= 1 || length(X) <= 1) return(lapply(X, FUN, ...))
  # Started once at full size and reused by later calls (see par_stop())
  parallel::parLapplyLB(par_cluster(workers), X, FUN, ...)
}
