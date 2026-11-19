# Raw time-force data from 24-well plates ---------------------------------------
# Config, plate map and time-force CSV loading for plate-report.qmd.
# Upstream of the EMT report: these are the per-well force traces exported
# before the PB analysis produces the Excel workbooks.

suppressPackageStartupMessages(library(tidyverse))

# How each protocol is drawn when the config doesn't say otherwise.
#   snapshot: one contraction, window [time_s - pre_s, time_s + post_s]
#   full:     the whole protocol, [0, end_s]; anything after end_s is assumed to
#             be the recording left running and is dropped. end_s: null = whole file.
# Defaults come from the stim schedules seen in the R672C batch 2/3 files:
#   02Hz stims at 5, 10 ... 50 s; 1Hz every second 3-32 s; 75Hz tetani at
#   3, 48, 93, 139 s; Fatigue ~3-55.5 s; FF ~3-240 s.
plate_protocol_defaults <- list(
  `02Hz`  = list(type = "full", end_s = 55),
  `1Hz`   = list(type = "full", end_s = 35),
  `75Hz`  = list(type = "snapshot", time_s = 48, pre_s = 0.2, post_s = 3),
  Fatigue = list(type = "full", end_s = 60),
  FF      = list(type = "full", end_s = 250)
)

plate_config_defaults <- list(
  days        = "all",
  protocols   = list(),
  conditions  = NULL,
  colors      = NULL,
  max_points  = 4000,      # per-well point cap for full traces (min/max decimation); 0 = off
  plot_size   = c(14, 8),
  png_dpi     = 200,
  animation   = TRUE,      # one GIF per protocol stepping through the days
  animation_delay_s = 1,   # seconds each day is shown
  animation_dpi = 100
)

plate_map_pattern <- "plate[ _-]?map.*\\.csv$"

#' Find the plate map for a data folder by looking in the folder, then up to
#' `levels` folders above it. Where one folder holds several plate maps
#' (e.g. "R672C Plate Map.csv" + "T178I Plate Map.csv"), the one whose prefix
#' appears in the data folder's path is used (…/pulseFiles/T178I -> T178I).
#' Returns list(path, dir = folder it was found in, plate = prefix or NULL).
find_plate_map <- function(dir, levels = 3) {
  d <- dir
  for (i in 0:levels) {
    hits <- list.files(d, plate_map_pattern, ignore.case = TRUE, full.names = TRUE)
    if (length(hits) == 1) return(list(path = hits, dir = d, plate = NULL))
    if (length(hits) > 1) {
      prefix <- trimws(sub("[ _-]*plate[ _-]?map.*$", "", basename(hits), ignore.case = TRUE))
      below <- substring(dir, nchar(d) + 1)
      match <- nzchar(prefix) & vapply(prefix, \(p) grepl(p, below, fixed = TRUE), logical(1))
      if (sum(match) == 1) return(list(path = hits[match], dir = d, plate = prefix[match]))
      stop("Several plate maps in ", d, ":\n  ", paste(basename(hits), collapse = "\n  "),
           "\nPoint at the plate's own data folder (e.g. …/", prefix[1],
           "), or use a config file and set `plate_map`.", call. = FALSE)
    }
    if (dirname(d) == d) break
    d <- dirname(d)
  }
  stop("No plate map (*Plate Map*.csv) found in ", dir, " or the ", levels,
       " folders above it. Use a config file and set `plate_map`.", call. = FALSE)
}

# Folder names that say what kind of data is inside rather than which batch
generic_folder <- function(x) {
  grepl("^(timeforce|pulse.*|mantarray.*|ma-recordings)$", tolower(x))
}

#' `path` is either a YAML config or a data folder. For a folder, files are
#' read from <folder>/timeforce if it exists, otherwise from the folder itself
#' (recursively). The plate map is found by find_plate_map(). The plate is
#' named after the folder holding the plate map (or the one above it, if that
#' folder is just "timeforce"/"pulse…"), plus the plate-map prefix when that
#' folder holds several plates.
load_plate_config <- function(path) {
  if (dir.exists(path)) {
    folder <- normalizePath(path)
    tf <- file.path(folder, "timeforce")
    pmap <- find_plate_map(folder)
    batch_dir <- pmap$dir
    # Walk past folders that only describe the data ("timeforce", "pulseFiles")
    # or are named after the plate itself (…/pulseFiles/T178I holding
    # "T178I Plate Map.csv"); in the latter case keep the plate in the name.
    plate_of_map <- trimws(sub("[ _-]*plate[ _-]?map.*$", "", basename(pmap$path), ignore.case = TRUE))
    repeat {
      b <- basename(batch_dir)
      if (dirname(batch_dir) == batch_dir) break
      if (nzchar(plate_of_map) && tolower(b) == tolower(plate_of_map)) {
        pmap$plate <- plate_of_map
      } else if (!generic_folder(b)) break
      batch_dir <- dirname(batch_dir)
    }
    name <- paste(c(basename(batch_dir), pmap$plate), collapse = " ")
    cfg <- list(name = name,
                timeforce_dir = if (dir.exists(tf)) tf else folder,
                plate_map = pmap$path,
                output_dir = file.path("output", "plates", gsub("[^A-Za-z0-9._-]+", "_", name)))
  } else if (file.exists(path)) {
    cfg <- yaml::read_yaml(path)
  } else {
    stop("Not a config file or folder: ", path, call. = FALSE)
  }

  cfg <- merge_defaults(cfg, plate_config_defaults)
  cfg$name <- cfg$name %||% "EMT plate"
  cfg$output_dir <- cfg$output_dir %||% file.path("output", "plates",
                                                 gsub("[^A-Za-z0-9]+", "-", tolower(cfg$name)))
  for (k in c("timeforce_dir", "plate_map")) {
    if (is.null(cfg[[k]])) stop("Config must set `", k, "`.", call. = FALSE)
  }
  if (!dir.exists(cfg$timeforce_dir)) stop("timeforce_dir not found: ", cfg$timeforce_dir, call. = FALSE)
  if (!file.exists(cfg$plate_map)) stop("plate_map not found: ", cfg$plate_map, call. = FALSE)

  # Per-protocol settings: config entries override the built-in defaults key by key
  protos <- plate_protocol_defaults
  for (nm in names(cfg$protocols)) {
    protos[[nm]] <- merge_defaults(cfg$protocols[[nm]], protos[[nm]] %||% list(type = "full"))
  }
  cfg$protocols <- protos
  cfg$plot_size <- as.numeric(cfg$plot_size)
  cfg
}

# Settings for a protocol, falling back to "whole file" for unknown ones
protocol_settings <- function(cfg, protocol) {
  cfg$protocols[[protocol]] %||% list(type = "full", end_s = NULL)
}

# Plate map --------------------------------------------------------------------

#' Plate map CSV in plate layout: first column = row letters, header = column
#' numbers, cells = sample/group label. Returns one row per well, in plate
#' order (A1, A2, ... then B1, ...), with `group` mapped through
#' `conditions` and ordered like the config.
read_plate_map <- function(path, conditions = NULL) {
  m <- suppressWarnings(  # "incomplete final line" from Excel-saved CSVs
    read.csv(path, check.names = FALSE, stringsAsFactors = FALSE,
             fileEncoding = "UTF-8-BOM", colClasses = "character"))
  names(m)[1] <- "row"
  m <- m |> filter(nzchar(trimws(row)))
  rows <- trimws(m$row)
  cols <- trimws(names(m)[-1])
  long <- m |>
    pivot_longer(-row, names_to = "col", values_to = "label") |>
    mutate(row = trimws(row), col = trimws(col),
           label = na_if(trimws(label), ""),
           well = paste0(row, col),
           row = factor(row, levels = rows),
           col = factor(col, levels = cols)) |>
    arrange(row, col)

  raw <- coalesce(long$label, "(empty)")
  long$group <- if (is.null(conditions)) {
    factor(raw, levels = unique(raw))
  } else {
    recode_conditions(raw, conditions)
  }
  long |> mutate(well = factor(well, levels = well))
}

plate_colors <- function(groups, colors = NULL) {
  lv <- levels(groups)
  okabe_ito <- c("#000000", "#009E73", "#E69F00", "#D55E00", "#CC79A7",
                 "#0072B2", "#56B4E9", "#F0E442")
  cols <- unlist(colors)
  missing <- setdiff(lv, names(cols))
  spare <- setdiff(okabe_ito, cols)
  if (length(missing) > length(spare)) spare <- c(spare, scales::hue_pal()(length(missing)))
  cols <- c(cols, setNames(spare[seq_along(missing)], missing))
  if ("(empty)" %in% lv) cols["(empty)"] <- "grey70"
  cols[lv]
}

# Files ----------------------------------------------------------------------

#' Every time-force file under `dir` (recursively), with day and protocol
#' parsed from the name. Two formats are read:
#'   - CSV exports: "Time (s)" + one column per well (A1, B1, ...)
#'   - Pulse3D "full" workbooks (*__full__*.xlsx), sheet "continuous-waveforms".
#'     Other .xlsx files (e.g. PB-analyzed d10_FF_waveformkinetics.xlsx) are ignored.
#' Name parsing:
#'   - a Pulse3D "__full__<timestamp>" suffix is dropped first;
#'   - everything up to the first "__" is an instrument/timestamp prefix and is
#'     ignored when looking for the day;
#'   - day = "d<number>_" at the start or after an underscore;
#'   - protocol = the last "_"-separated token.
#'   d7_R672C-Batch-3_FF.csv                                               -> day 7, FF
#'   ML26187408-6__2026_07_24_230936_d7_DAEMTR672CBatch2_FF.csv            -> day 7, FF
#'   ML25296405-6_2026_01_26_235812__d10_T178I-Cl113Cl152-batch-8_1Hz.csv  -> day 10, 1Hz
#'   ML25101401-2__2025_05_19_212941_d3_DAEMTT178I_75Hz__full__2025-05-23_15-01-46.xlsx -> day 3, 75Hz
#' Protocol names are matched ignoring case, "-" and "_" (see canonical_protocol()).
#' Files where day or protocol can't be parsed are returned in attr(, "skipped").
#' When one day x protocol has several files, the newest is kept and the rest
#' are returned in attr(, "duplicates").
list_timeforce_files <- function(dir, days = "all", known = names(plate_protocol_defaults)) {
  fs <- list.files(dir, pattern = "\\.(csv|xlsx)$", recursive = TRUE, full.names = TRUE,
                   ignore.case = TRUE)
  base <- basename(fs)
  is_xlsx <- grepl("\\.xlsx$", base, ignore.case = TRUE)
  keep <- !grepl(plate_map_pattern, base, ignore.case = TRUE) &
    !startsWith(base, "~$") &                       # Excel lock files
    (!is_xlsx | grepl("__full__", base, fixed = TRUE))
  fs <- fs[keep]; base <- base[keep]; is_xlsx <- is_xlsx[keep]

  name <- sub("__full__.*$", "", base)                        # Pulse3D export suffix
  name <- sub("\\.(csv|xlsx)$", "", name, ignore.case = TRUE)
  stem <- sub("^.*?__", "", name, perl = TRUE)                # drop prefix up to the first "__"
  out <- tibble(file = fs,
                format = if_else(is_xlsx, "Pulse3D xlsx", "CSV"),
                day = as.integer(str_match(stem, "(?:^|_)[dD](\\d+)_")[, 2]),
                protocol = str_match(name, "_([^_]+)$")[, 2],
                mtime = file.mtime(fs))
  out$protocol <- canonical_protocol(out$protocol, known)
  skipped <- filter(out, is.na(day) | is.na(protocol))
  out <- filter(out, !is.na(day), !is.na(protocol))
  if (!identical(days, "all") && !is.null(days)) out <- filter(out, day %in% as.numeric(unlist(days)))

  out <- out |> group_by(day, protocol) |> arrange(desc(mtime), .by_group = TRUE)
  dups <- out |> filter(row_number() > 1) |> ungroup()
  out <- out |> slice(1) |> ungroup() |> arrange(day, protocol)
  attr(out, "skipped") <- skipped
  attr(out, "duplicates") <- dups
  out
}

#' Group spellings of the same protocol, ignoring case, "-", "_" and spaces:
#' "75HZ" -> "75Hz", "AS-Fatigue"/"ASfatigue" -> one name. A spelling of a
#' configured protocol (`known`) takes the configured name; otherwise the most
#' common spelling wins.
canonical_protocol <- function(x, known = character()) {
  norm <- function(v) tolower(gsub("[-_ ]", "", v))
  key <- norm(x)
  known_map <- setNames(known, norm(known))
  common <- tapply(x, key, \(v) names(sort(table(v), decreasing = TRUE))[1])
  ifelse(is.na(x), NA_character_,
         ifelse(key %in% names(known_map), known_map[key], common[key]))
}

#' One time-force file (CSV export or Pulse3D full workbook) -> long table
#' (time, well, force), optionally limited to [t_from, t_to].
read_timeforce <- function(file, t_from = -Inf, t_to = Inf) {
  if (grepl("\\.xlsx$", file, ignore.case = TRUE)) {
    sheets <- readxl::excel_sheets(file)
    if (!"continuous-waveforms" %in% sheets) {
      stop("No 'continuous-waveforms' sheet in ", basename(file), call. = FALSE)
    }
    x <- suppressMessages(readxl::read_excel(file, "continuous-waveforms"))
  } else {
    x <- suppressMessages(readr::read_csv(file, show_col_types = FALSE, progress = FALSE))
  }
  tcol <- grep("^time", names(x), ignore.case = TRUE, value = TRUE)[1]
  if (is.na(tcol)) stop("No time column in ", basename(file), call. = FALSE)
  # CSV: "A1"; Pulse3D: "A1 - Active Twitch Force (μN)" (not "A1 Peak Amplitudes")
  wells <- grep("^[A-Z]+[0-9]+( - .*[Ff]orce.*)?$", names(x), value = TRUE)
  if (length(wells) == 0) stop("No well columns (A1, B1, ...) in ", basename(file), call. = FALSE)
  x |>
    select(time = all_of(tcol), all_of(wells)) |>
    rename_with(\(n) sub(" - .*$", "", n), -time) |>
    filter(!is.na(time), time >= t_from, time <= t_to) |>
    pivot_longer(-time, names_to = "well", values_to = "force")
}

#' Keep each bin's min and max so peaks and the noise envelope survive while a
#' 250 s trace drops from ~25,000 to `max_points` points per well.
decimate_minmax <- function(d, max_points) {
  n <- max(table(d$well))
  if (!max_points || n <= max_points) return(d)
  n_bins <- max_points %/% 2
  d |>
    group_by(well) |>
    mutate(bin = cut(time, n_bins, labels = FALSE)) |>
    group_by(well, bin) |>
    slice(unique(c(which.min(force), which.max(force)))) |>
    arrange(time, .by_group = TRUE) |>
    ungroup() |>
    select(-bin)
}

#' Read one file for plotting according to its protocol settings.
#' Returns list(data, window = c(from, to), file_end, settings).
load_protocol_trace <- function(file, protocol, cfg) {
  s <- protocol_settings(cfg, protocol)
  if (identical(s$type, "snapshot")) {
    win <- c(s$time_s - s$pre_s, s$time_s + s$post_s)
  } else {
    win <- c(0, s$end_s %||% Inf)
  }
  d <- read_timeforce(file, win[1], win[2])
  file_end <- if (nrow(d)) max(d$time) else NA_real_
  if (identical(s$type, "snapshot")) {
    # x = time from the stim, so every snapshot reads 0 = stim
    d <- mutate(d, time = time - s$time_s)
    win <- win - s$time_s
  } else {
    d <- decimate_minmax(d, cfg$max_points)
  }
  list(data = d, window = win, file_end = file_end, settings = s)
}
