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

#' One plate image to draw, for run_plate_job(): a day's PNG (one trace, limits
#' from that trace) or an animation (several traces, shared `limits`, one frame
#' each). `file` is the output path.
plate_png_job <- function(trace, plate_map, title, x_label, file) {
  list(kind = "png", trace = trace, plate_map = plate_map, title = title,
       x_label = x_label, file = file)
}

plate_gif_job <- function(traces, plate_map, limits, titles, subtitle, x_label, file) {
  list(kind = "gif", traces = traces, plate_map = plate_map, limits = limits,
       titles = titles, subtitle = subtitle, x_label = x_label, file = file)
}

#' Draw one job (runs in a worker, see par_lapply()). Returns the file path,
#' or the error so the report can show it instead of stopping.
run_plate_job <- function(job, colors, cfg) {
  tryCatch({
    dir.create(dirname(job$file), recursive = TRUE, showWarnings = FALSE)
    if (job$kind == "png") {
      plt <- plate_grid_plot(job$trace, job$plate_map, colors, plate_limits(job$trace),
                             title = job$title, x_label = job$x_label)
      ggsave(job$file, plot = plt, width = cfg$plot_size[1], height = cfg$plot_size[2],
             dpi = cfg$png_dpi)
    } else {
      frames <- map2(job$traces, job$titles, \(tr, t)
        plate_grid_plot(tr, job$plate_map, colors, job$limits, title = t,
                        subtitle = job$subtitle, x_label = job$x_label))
      make_plate_animation(frames, job$file, cfg$plot_size, cfg$animation_dpi,
                           cfg$animation_delay_s)
    }
    job$file
  }, error = identity)
}

#' Read every file for plotting, in parallel. Returns one entry per file:
#' the trace, or the error. `protocols` gives each file's protocol.
load_protocol_traces <- function(files, protocols, cfg) {
  par_lapply(seq_along(files), \(i, files, protocols, cfg)
    tryCatch(load_protocol_trace(files[i], protocols[i], cfg), error = identity),
    files = files, protocols = protocols, cfg = cfg)
}

#' Draw every job in parallel, animations first (they take longest).
#' Returns the results (path or error) in the order of `jobs`.
run_plate_jobs <- function(jobs, colors, cfg) {
  if (!length(jobs)) return(list())
  kinds <- vapply(jobs, `[[`, "", "kind")
  ord <- order(kinds != "gif")
  out <- vector("list", length(jobs))
  out[ord] <- par_lapply(jobs[ord], run_plate_job, colors = colors, cfg = cfg)
  out
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
