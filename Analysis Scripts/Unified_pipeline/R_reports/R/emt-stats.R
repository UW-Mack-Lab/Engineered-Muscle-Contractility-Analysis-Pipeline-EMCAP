# Condition comparisons for every metric x day (x Hz) --------------------------

suppressPackageStartupMessages(library(tidyverse))

#' Compare each condition against the reference condition with a t-test, for
#' every metric and every combination of the `by` columns.
#'
#' unit = "batch": wells are averaged within each batch first, so n = number
#'   of batches (differentiations). This is the default and the right unit when
#'   batch-to-batch variation is the main source of noise.
#' unit = "well": every well is an observation. Wells in one batch are not
#'   independent, so these p-values are optimistic (pseudo-replication).
#' paired = TRUE (unit = "batch" only): paired t-test on batch means, using
#'   only batches that have both conditions.
#' p-values are adjusted with `p_adjust` within each metric (across days, Hz
#'   and comparisons).
compare_conditions <- function(metrics, by = "day", unit = c("batch", "well"),
                               paired = FALSE, p_adjust = "BH", ref = NULL) {
  unit <- match.arg(unit)
  if (paired && unit != "batch") stop("paired = TRUE requires unit = 'batch'.", call. = FALSE)

  conds <- levels(droplevels(factor(metrics$Condition)))
  ref <- ref %||% conds[1]
  others <- setdiff(conds, ref)
  if (length(others) == 0) return(tibble())

  d <- metrics |> filter(is.finite(value))
  if (unit == "batch") {
    d <- d |>
      group_by(across(all_of(c("metric", by, "batch", "Condition")))) |>
      summarise(value = mean(value), .groups = "drop")
  }

  one_test <- function(df, cmp) {
    a <- filter(df, Condition == ref)
    b <- filter(df, Condition == cmp)
    if (paired) {
      pr <- inner_join(a, b, by = "batch", suffix = c("_ref", "_cmp"))
      x <- pr$value_ref; y <- pr$value_cmp
    } else {
      x <- a$value; y <- b$value
    }
    note <- NA_character_
    p <- NA_real_
    if (length(x) < 2 || length(y) < 2) {
      note <- "n < 2"
    } else if (paired && sd(y - x) == 0 || !paired && sd(x) == 0 && sd(y) == 0) {
      note <- "no variance"
    } else {
      p <- tryCatch(t.test(y, x, paired = paired)$p.value,
                    error = function(e) { note <<- conditionMessage(e); NA_real_ })
    }
    tibble(comparison = paste(cmp, "vs", ref),
           n_ref = length(x), n_cmp = length(y),
           mean_ref = mean(x), mean_cmp = mean(y),
           difference = mean(y) - mean(x),
           p_value = p, note = note)
  }

  d |>
    group_by(across(all_of(c("metric", by)))) |>
    group_modify(function(df, key) map(others, \(cmp) one_test(df, cmp)) |> list_rbind()) |>
    ungroup() |>
    group_by(metric) |>
    mutate(p_adj = p.adjust(p_value, method = p_adjust)) |>
    ungroup() |>
    relocate(p_adj, .after = p_value) |>
    mutate(unit = if (paired) "batch (paired)" else unit)
}
