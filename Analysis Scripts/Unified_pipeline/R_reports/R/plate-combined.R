# Combined plate trace report: several batches in one report ----------------------
# Config and file discovery for plate-combined-report.qmd. A combined config is a
# plate config with a `batches:` list instead of one timeforce_dir/plate_map.
# Every other setting (conditions, colors, protocols, days, plot sizes, ...) is
# shared by all batches; a batch entry can override any of them.

#' TRUE if `path` is a YAML config with a `batches:` list.
is_combined_plate_config <- function(path) {
  is.character(path) && length(path) == 1 && !dir.exists(path) && file.exists(path) &&
    grepl("\\.ya?ml$", path, ignore.case = TRUE) &&
    !is.null(yaml::read_yaml(path)$batches)
}

#' Returns list(name, output_dir, settings = shared plate settings,
#' batches = list of per-batch plate configs, as load_plate_config() makes them).
#' Each `batches:` entry is either a batch folder (same rules as
#' `render-plates.R <folder>`) or a map with name / timeforce_dir / plate_map and
#' optionally any per-batch overrides (e.g. days).
load_combined_plate_config <- function(path) {
  top <- yaml::read_yaml(path)
  if (!length(top$batches)) stop("Combined plate config needs a `batches:` list.", call. = FALSE)
  shared <- top[setdiff(names(top), c("batches", "name", "output_dir"))]

  batches <- lapply(seq_along(top$batches), function(i) {
    b <- top$batches[[i]]
    if (is.character(b)) {
      auto <- load_plate_config(b)
      b <- auto[c("name", "timeforce_dir", "plate_map")]
    }
    b$name <- b$name %||% paste("Batch", i)
    tmp <- tempfile(fileext = ".yml")
    on.exit(unlink(tmp))
    yaml::write_yaml(utils::modifyList(shared, b), tmp)
    load_plate_config(tmp)
  })
  nms <- vapply(batches, `[[`, "", "name")
  if (anyDuplicated(nms)) stop("Batch names must be unique: ", paste(nms, collapse = ", "), call. = FALSE)

  name <- top$name %||% "Combined plates"
  list(name = name,
       output_dir = top$output_dir %||% file.path("output", "plates",
                                                  gsub("[^A-Za-z0-9]+", "-", tolower(name))),
       settings = batches[[1]],   # shared settings (protocols, plot sizes, ...) after defaults
       batches = setNames(batches, nms))
}

#' Time-force files of every batch, with a `batch` column. Protocol spellings are
#' unified across batches too (Batch 1 "ASfatigue" + Batch 2 "ASFatigue" -> one).
#' attr(, "duplicates") / attr(, "skipped") are combined the same way.
list_combined_timeforce_files <- function(ccfg) {
  per <- lapply(names(ccfg$batches), function(nm) {
    b <- ccfg$batches[[nm]]
    f <- list_timeforce_files(b$timeforce_dir, b$days, known = names(b$protocols))
    list(files = mutate(f, batch = nm, .before = 1),
         dups = mutate(attr(f, "duplicates"), batch = nm, .before = 1),
         skipped = mutate(attr(f, "skipped"), batch = nm, .before = 1),
         misfiled = mutate(attr(f, "misfiled"), batch = nm, .before = 1))
  })
  known <- names(ccfg$settings$protocols)
  files <- bind_rows(lapply(per, `[[`, "files"))
  files$protocol <- canonical_protocol(files$protocol, known)
  files$batch <- factor(files$batch, levels = names(ccfg$batches))
  attr(files, "duplicates") <- bind_rows(lapply(per, `[[`, "dups"))
  attr(files, "skipped") <- bind_rows(lapply(per, `[[`, "skipped"))
  attr(files, "misfiled") <- bind_rows(lapply(per, `[[`, "misfiled"))
  files
}
