# Plot functions for the EMT report --------------------------------------------
# Generalized from the figure script and metricSuperPlot / metricSuperPlotFF in
# pulse+AnalysisPlotting-fcns.R. Nothing dataset-specific is hardcoded here:
# days, batches, colors and labels all come in as arguments (from the config).

suppressPackageStartupMessages({
  library(tidyverse)
  library(ggprism)
})

# ggplot's default shapes first (so <= 6 batches look like the old figures),
# then more, so > 6 batches don't silently drop points
batch_shapes <- c(16, 17, 15, 3, 7, 8, 0, 1, 2, 4, 5, 6, 18, 9, 10, 11, 12, 13, 14)

scale_condition <- function(colors) {
  list(scale_color_manual(values = colors, aesthetics = c("colour", "fill")))
}

scale_batch_shape <- function() {
  scale_shape_manual(values = batch_shapes)
}

# Kinetics traces (panels A and B) ---------------------------------------------

#' Window the condition-average 75 Hz traces around one stim, subtract a
#' baseline and normalize to the peak within the window.
prep_kinetics_traces <- function(traces, stim_time_s = 48, window_s = 3,
                                 baseline_s = 0.01, days_max = Inf) {
  traces |>
    filter(day <= days_max,
           time > stim_time_s,
           time < stim_time_s + window_s) |>
    group_by(condition, day, batch) |>
    mutate(
      in_baseline  = time <= stim_time_s + baseline_s + 1e-9,
      # If the baseline window catches no samples, fall back to the first one
      baseline     = if (any(in_baseline)) mean(force[in_baseline]) else force[which.min(time)],
      bk_sub_force = force - baseline,
      new_norm     = bk_sub_force / max(bk_sub_force),
      rel_time     = time - stim_time_s
    ) |>
    ungroup() |>
    select(-in_baseline)
}

#' Panel A (type = "normalized") or panel B (type = "active").
#' One line per condition per batch.
kinetics_trace_plot <- function(kin, type = c("normalized", "active"),
                                colors, window_s = 3, ncol = 5) {
  type <- match.arg(type)
  p <- kin |>
    ggplot(aes(x = rel_time,
               y = if (type == "normalized") new_norm * 100 else bk_sub_force,
               color = condition,
               group = interaction(condition, batch))) +
    geom_line() +
    facet_wrap(~formal_date, ncol = ncol, scales = "free") +
    theme_prism() +
    scale_condition(colors) +
    labs(x = "Time (s)",
         y = if (type == "normalized") "Normalized Force (%)" else "Active Force (μN)")
  if (type == "normalized") {
    p + coord_cartesian(xlim = c(0, window_s), ylim = c(-10, 100))
  } else {
    p + coord_cartesian(xlim = c(0, window_s))
  }
}

# Metric superplots --------------------------------------------------------------

filter_metric <- function(metrics, metric_name, days_max, batches) {
  if (!metric_name %in% metrics$metric) {
    stop("Metric '", metric_name, "' not found. Available: ",
         paste(unique(metrics$metric), collapse = " | "), call. = FALSE)
  }
  metrics |>
    filter(metric == metric_name,
           day <= days_max,
           batch %in% batches)
}

#' Faint points = wells; bold points = batch means; x = day.
metric_superplot <- function(metrics, metric_name, colors, days_max = Inf,
                             batches = unique(metrics$batch), y_label = metric_name,
                             labels = list(condition = "Genotype", batch = "Differentiation"),
                             seed = 1) {
  raw_df <- filter_metric(metrics, metric_name, days_max, batches)

  summary_df <- raw_df |>
    group_by(Condition, day, batch, metric) |>
    summarise(value = mean(value, na.rm = TRUE), .groups = "drop")

  pos <- position_jitterdodge(jitter.width = 0.7, dodge.width = 0.9, seed = seed)

  ggplot(raw_df, aes(x = as.factor(day), y = value, color = Condition,
                     shape = as.factor(batch), group = Condition)) +
    geom_point(alpha = 0.1, position = pos, na.rm = TRUE) +
    geom_point(data = summary_df, size = 3, position = pos, na.rm = TRUE) +
    theme_prism() +
    theme(legend.title = element_text()) +
    scale_condition(colors) +
    scale_batch_shape() +
    labs(x = "Day", y = y_label,
         color = labels$condition, shape = labels$batch)
}

#' Force-frequency superplot: bars = grand mean, faint points = wells,
#' bold points = batch means; x = stimulation frequency, faceted by day.
ff_superplot <- function(metrics, metric_name, colors, days_max = Inf,
                         batches = unique(metrics$batch), y_label = metric_name,
                         labels = list(condition = "Genotype", batch = "Differentiation"),
                         ncol = 5, seed = 1) {
  raw_df <- filter_metric(metrics, metric_name, days_max, batches) |>
    mutate(Hz = factor(Hz, levels = sort(unique(Hz))))

  summary_df <- raw_df |>
    group_by(Condition, day, formal_date, batch, Hz, metric) |>
    summarise(value = mean(value, na.rm = TRUE), .groups = "drop")

  bar_df <- raw_df |>
    group_by(Condition, Hz, formal_date) |>
    summarise(value = mean(value, na.rm = TRUE), .groups = "drop")

  dodge_w <- 0.9
  pos <- position_jitterdodge(jitter.width = 0.7, dodge.width = dodge_w, seed = seed)

  ggplot() +
    geom_col(data = bar_df, aes(x = Hz, y = value, fill = Condition),
             alpha = 0.5, width = 0.8, position = position_dodge(width = dodge_w)) +
    geom_point(data = raw_df,
               aes(x = Hz, y = value, color = Condition,
                   shape = as.factor(batch), group = Condition),
               alpha = 0.1, position = pos, na.rm = TRUE) +
    geom_point(data = summary_df,
               aes(x = Hz, y = value, color = Condition),
               size = 2, position = pos, na.rm = TRUE) +
    facet_wrap(~formal_date, ncol = ncol) +
    theme_prism() +
    theme(legend.title = element_text(),
          axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    scale_condition(colors) +
    scale_batch_shape() +
    labs(x = "Stimulation Frequency (Hz)", y = y_label,
         color = labels$condition, fill = labels$condition, shape = labels$batch)
}

# Fatigue --------------------------------------------------------------------

#' Condition-average normalized force vs stim number, faceted by day.
#' With one batch the ribbon is the sheet's across-well SD; with several
#' batches the line is the mean of the batch averages and the ribbon is the
#' SD across batches.
fatigue_curve_plot <- function(fatigue_traces, colors, days_max = Inf,
                               batches = unique(fatigue_traces$batch), ncol = 5) {
  d <- fatigue_traces |>
    filter(day <= days_max, batch %in% batches)
  n_b <- n_distinct(d$batch)
  d <- d |>
    group_by(condition, day, formal_date, stim_number) |>
    summarise(sd_force   = if (n_b == 1) sd_force[1] else sd(norm_force, na.rm = TRUE),
              norm_force = mean(norm_force, na.rm = TRUE),
              .groups = "drop")

  ggplot(d, aes(x = stim_number, y = norm_force,
                ymin = norm_force - sd_force, ymax = norm_force + sd_force,
                color = condition, fill = condition)) +
    geom_ribbon(alpha = 0.3, color = NA) +
    geom_line() +
    facet_wrap(~formal_date, ncol = ncol) +
    theme_prism() +
    scale_condition(colors) +
    labs(x = "Stimulation Number", y = "Force (% of first)")
}

# Saving ---------------------------------------------------------------------

safe_filename <- function(x) {
  x |>
    str_replace_all("μ", "u") |>
    str_replace_all("%", "pct") |>
    str_replace_all("[^A-Za-z0-9]+", "_") |>
    str_remove_all("^_|_$")
}

save_panel <- function(plot, name, size, dir, dpi = 300) {
  dir.create(dir, recursive = TRUE, showWarnings = FALSE)
  path <- file.path(dir, paste0(safe_filename(name), ".png"))
  ggsave(path, plot = plot, width = size[1], height = size[2], dpi = dpi)
  invisible(path)
}
