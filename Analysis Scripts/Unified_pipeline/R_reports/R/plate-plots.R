# Plate-layout trace plots ---------------------------------------------------------

suppressPackageStartupMessages({
  library(tidyverse)
  library(ggprism)
})

#' Shared x/y limits for one or more traces (from load_protocol_trace()).
#' x = the plotted window, or the data range when the window is open-ended;
#' y = the force range over every well in every trace.
plate_limits <- function(traces) {
  if (!is.null(traces$data)) traces <- list(traces)
  all_t <- unlist(lapply(traces, \(tr) range(tr$data$time)))
  win <- traces[[1]]$window
  x <- c(if (is.finite(win[1])) win[1] else min(all_t),
         if (is.finite(win[2])) win[2] else max(all_t))
  y <- range(unlist(lapply(traces, \(tr) range(tr$data$force, na.rm = TRUE))))
  list(x = x, y = y)
}

#' One panel per well, arranged like the plate (rows A-D, columns 1-6 for a
#' 24-well plate; the layout comes from the plate map). Line color = group.
#' Every panel shares the same x and y limits: `limits` from plate_limits(),
#' defaulting to this trace alone. Each panel is labelled with the well, its
#' plate-map label, and the maximum force in the window.
plate_grid_plot <- function(trace, plate_map, colors, limits = plate_limits(trace),
                            title = NULL, subtitle = NULL, x_label = "Time (s)") {
  d <- trace$data |>
    inner_join(select(plate_map, well, group), by = "well",
               relationship = "many-to-one") |>
    mutate(well = factor(well, levels = levels(plate_map$well)))

  strip <- plate_map |>
    mutate(strip = paste0(well, "  ", coalesce(label, "(empty)")))
  strip_labels <- setNames(strip$strip, as.character(strip$well))

  peaks <- d |>
    group_by(well) |>
    summarise(peak = max(force, na.rm = TRUE), .groups = "drop") |>
    mutate(text = sprintf("max %.0f", peak))

  n_col <- nlevels(plate_map$col)

  ggplot(d, aes(time, force, color = group)) +
    geom_line(linewidth = 0.3) +
    geom_text(data = peaks, aes(x = Inf, y = Inf, label = text),
              inherit.aes = FALSE, hjust = 1.1, vjust = 1.4, size = 2.6, color = "grey40") +
    facet_wrap(~well, ncol = n_col, drop = FALSE, scales = "fixed",
               labeller = as_labeller(strip_labels)) +
    scale_color_manual(values = colors, drop = FALSE, name = NULL) +
    scale_x_continuous(breaks = scales::breaks_extended(n = 4)) +
    coord_cartesian(xlim = limits$x, ylim = limits$y) +
    theme_prism(base_size = 9) +
    theme(legend.position = "bottom",
          legend.text = element_text(size = 10),
          plot.title = element_text(size = 14),
          strip.text = element_text(size = 8, hjust = 0),
          panel.spacing.x = unit(1.2, "lines"),
          panel.spacing.y = unit(0.6, "lines")) +
    guides(color = guide_legend(override.aes = list(linewidth = 1.5))) +
    labs(x = x_label, y = "Force (μN)", title = title, subtitle = subtitle)
}

#' Write a list of ggplots as an animated GIF, one frame per plot.
make_plate_animation <- function(plots, file, size = c(14, 8), dpi = 100, delay_s = 1) {
  tmp <- tempfile("frames")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE))
  frames <- vapply(seq_along(plots), function(i) {
    f <- file.path(tmp, sprintf("frame%03d.png", i))
    ggsave(f, plots[[i]], width = size[1], height = size[2], dpi = dpi, bg = "white")
    f
  }, character(1))
  dir.create(dirname(file), recursive = TRUE, showWarnings = FALSE)
  magick::image_read(frames) |>
    magick::image_animate(delay = round(100 * delay_s), loop = 0, optimize = TRUE) |>
    magick::image_write(file)
  invisible(file)
}
