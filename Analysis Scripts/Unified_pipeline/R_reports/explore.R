# Interactive exploration of an EMT dataset -------------------------------------
# Open EMT-Reports.Rproj in RStudio (so the working directory is this folder),
# then run this script a section at a time (Cmd+Enter / Cmd+Shift+Enter).
# Nothing here writes files unless you run a ggsave()/write_csv() line.

suppressPackageStartupMessages(library(tidyverse))
for (f in list.files("R", pattern = "\\.R$", full.names = TRUE)) source(f)

# 1. Load ----------------------------------------------------------------------
# Slow part (reads every workbook over OneDrive). Do it once per session.
emt <- load_emt_data("configs/mydataset.yml")
# Override batches without editing the config:
# emt <- load_emt_data("configs/mydataset.yml", metric_batches = "all", trace_batches = c(1, 2))

# Optional: cache to disk so a restarted session doesn't re-read OneDrive.
# Re-run load_emt_data() whenever the workbooks change.
# saveRDS(emt, "output/mydataset/emt-data.rds")
# emt <- readRDS("output/mydataset/emt-data.rds")

cfg  <- emt$cfg
m75  <- emt$m75    # 75 Hz kinetics, one row per well x metric
mff  <- emt$mff    # force-frequency, one row per well x Hz x metric
mfat <- emt$mfat   # fatigue summary, one row per well x metric

# 2. Look around -----------------------------------------------------------------
unique(m75$metric)                       # what you can plot
m75 |> distinct(batch, day) |> arrange(batch, day)
m75 |> filter(day <= cfg$days_max) |> well_counts()

# Summary table: batch means for one metric
m75 |>
  filter(metric == "R50 Time") |>
  group_by(Condition, day, batch) |>
  summarise(mean = mean(value, na.rm = TRUE), wells = n(), .groups = "drop")

# 3. Same plots as the report ------------------------------------------------------
emt_superplot(emt, m75, "R50 Time", y_label = "Time to 50% Relax (s)")
emt_superplot(emt, mfat, "Fatigue Resistance Index (%)")
emt_ff_superplot(emt, "FF Normalized (%)")

# Override anything for one plot: all days, one batch, ...
emt_superplot(emt, m75, "Max tetanic-BK (μN)", days_max = Inf)
emt_superplot(emt, m75, "Max tetanic-BK (μN)", batches = 2)

# Traces
kin <- prep_kinetics_traces(emt$traces75, stim_time_s = 48, window_s = 3,
                            days_max = cfg$days_max)
kinetics_trace_plot(kin, "normalized", cfg$colors)
fatigue_curve_plot(emt$fat_tr, cfg$colors, cfg$days_max)

# 4. Customize: they're ordinary ggplots ------------------------------------------
p <- emt_superplot(emt, m75, "R50 Time")
p + facet_wrap(~batch) + ggtitle("R50 by batch")
p + scale_y_continuous(limits = c(0, NA))

# 5. Your own plots from the long tables -------------------------------------------
# Two metrics against each other, per well
m75 |>
  filter(metric %in% c("Max tetanic-BK (μN)", "R50 Time"), day <= cfg$days_max) |>
  pivot_wider(id_cols = c(batch, day, Plate, Well, Condition),
              names_from = metric, values_from = value, values_fn = mean) |>
  ggplot(aes(`Max tetanic-BK (μN)`, `R50 Time`, color = Condition)) +
  geom_point() +
  facet_wrap(~day) +
  scale_color_manual(values = cfg$colors) +
  ggprism::theme_prism()

# 6. Stats on demand --------------------------------------------------------------
compare_conditions(filter(m75, day <= cfg$days_max), by = "day")
compare_conditions(filter(m75, metric == "R50 Time"), by = "day", unit = "well")

# 7. Save something you like ---------------------------------------------------------
# save_panel(p, "R50_custom", size = c(4, 4), dir = "output/mydataset/custom")
# ggsave("output/mydataset/custom/my_plot.pdf", p, width = 4, height = 4)
