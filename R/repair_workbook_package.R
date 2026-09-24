#' Repair a saved workbook so Excel opens it without a repair prompt
#'
#' @description Cleans invalid parts that openxlsx (4.2.8.1) leaves behind in a
#'   saved .xlsx, each of which triggers Excel's "problem with some content"
#'   repair prompt:
#'   1. parts with no declared content type (e.g. a stray `[trash]/*.dat` folder)
#'   2. orphaned worksheet relationships pointing at drawing/vmlDrawing files
#'      that were never created
#'   3. duplicate legacy comment VML shapes (duplicate ids, or two shapes
#'      anchored to the same comment cell; the last one is kept)
#'
#' @param path A character string of the saved .xlsx file path, repaired in place
#'
#' @return `path`, invisibly
#' @export
#'
#' @examples
#' \dontrun{
#' openxlsx::saveWorkbook(wb, "tables.xlsx", overwrite = TRUE)
#' repair_workbook_package("tables.xlsx")
#' }
repair_workbook_package <- function(path) {
  path <- normalizePath(path, mustWork = TRUE)  # must be absolute - the function setwd()s into a scratch dir below
  work <- tempfile("xlsxrepair_")
  dir.create(work)
  on.exit(unlink(work, recursive = TRUE, force = TRUE), add = TRUE)
  
  utils::unzip(path, exdir = work)
  
  # (1) strip any part with no declared content type
  ct_file <- file.path(work, "[Content_Types].xml")
  ct_txt  <- paste(readLines(ct_file, warn = FALSE), collapse = "")
  declared_ext <- regmatches(ct_txt, gregexpr('Default Extension="[^"]*"', ct_txt))[[1]]
  declared_ext <- sub('Default Extension="([^"]*)"', "\\1", declared_ext)
  declared_override <- regmatches(ct_txt, gregexpr('PartName="[^"]*"', ct_txt))[[1]]
  declared_override <- sub('PartName="([^"]*)"', "\\1", declared_override)
  
  all_files <- list.files(work, recursive = TRUE, full.names = TRUE)
  for (f in all_files) {
    rel <- sub(work, "", f, fixed = TRUE)
    rel <- paste0("/", gsub("\\\\", "/", rel))
    rel <- sub("^/+", "/", rel)
    ext <- tolower(sub(".*[.]", "", rel))
    if (!(ext %in% declared_ext) && !(rel %in% declared_override)) unlink(f)
  }
  repeat {
    empty_dirs <- list.dirs(work, recursive = TRUE, full.names = TRUE)
    empty_dirs <- empty_dirs[vapply(empty_dirs, function(d) length(list.files(d, all.files = TRUE, no.. = TRUE)) == 0, logical(1))]
    if (length(empty_dirs) == 0) break
    unlink(empty_dirs, recursive = TRUE)
  }
  
  # (2) drop orphaned worksheet relationships (missing target AND unused r:id)
  rels_dir <- file.path(work, "xl/worksheets/_rels")
  if (dir.exists(rels_dir)) {
    rels_files <- list.files(rels_dir, pattern = "[.]rels$", full.names = TRUE)
    for (relf in rels_files) {
      sheet_base <- sub("[.]rels$", "", basename(relf))
      sheetxml   <- file.path(work, "xl/worksheets", sheet_base)
      if (!file.exists(sheetxml)) next
      sheet_txt  <- paste(readLines(sheetxml, warn = FALSE), collapse = "\n")
      
      txt <- paste(readLines(relf, warn = FALSE), collapse = "")
      rel_elems <- regmatches(txt, gregexpr("<Relationship[[:space:]]+[^>]*/>", txt, perl = TRUE))[[1]]
      if (length(rel_elems) == 0) next
      
      keep <- character(0)
      for (el in rel_elems) {
        id     <- sub('.*Id="([^"]*)".*', "\\1", el)
        target <- sub('.*Target="([^"]*)".*', "\\1", el)
        resolved <- normalizePath(file.path(work, "xl/worksheets", target), mustWork = FALSE)
        target_exists <- file.exists(resolved)
        id_used <- grepl(paste0('r:id="', id, '"'), sheet_txt, fixed = TRUE)
        if (!target_exists && !id_used) next
        keep <- c(keep, el)
      }
      
      if (length(keep) != length(rel_elems)) {
        header <- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        writeLines(paste0(header, paste(keep, collapse = ""), "</Relationships>"), relf, useBytes = TRUE)
      }
    }
  }
  
  # (3) de-duplicate legacy comment VML shapes - openxlsx's load/save round-trip
  # duplicates the <v:shapetype>/<v:shape> content of any vmlDrawing*.vml this
  # workbook already carries (Cover Sheet/QA-QC's real cell comments here),
  # leaving multiple elements sharing the same "id" attribute. Duplicate VML
  # ids are invalid and a known trigger for Excel's "problem with some
  # content" repair prompt. Keeping only the first occurrence of each id
  # restores the file to what the source template itself actually defines.
  vml_files <- list.files(file.path(work, "xl/drawings"), pattern = "^vmlDrawing[0-9]+[.]vml$", full.names = TRUE)
  for (vf in vml_files) {
    vdoc <- xml2::read_xml(vf)
    vns  <- xml2::xml_ns(vdoc)
    for (xpath in c(".//v:shapetype", ".//v:shape")) {
      nodes <- xml2::xml_find_all(vdoc, xpath, vns)
      seen <- character(0)
      for (nd in nodes) {
        id <- xml2::xml_attr(nd, "id")
        if (is.na(id)) next
        if (id %in% seen) xml2::xml_remove(nd) else seen <- c(seen, id)
      }
    }
    
    # (3b) the source template itself (independent of the openxlsx round-trip
    # above) separately carries a second, older leftover shape per comment -
    # e.g. an old flat-colour "note" shape alongside a newer theme-coloured
    # one - both anchored to the same cell, likely left behind when Excel
    # migrated its comment-rendering style without cleaning up the prior
    # shape. Two shapes anchored to one real comment is itself invalid and,
    # per testing, enough on its own to trigger the repair prompt even after
    # (3)'s id-dedup. Group remaining shapes by anchor cell (Row/Column) and
    # keep only the last one in document order - the more recently-authored,
    # modern-styled shape - per cell.
    shapes <- xml2::xml_find_all(vdoc, ".//v:shape", vns)
    if (length(shapes) > 0) {
      keys <- vapply(shapes, function(s) {
        row <- xml2::xml_text(xml2::xml_find_first(s, ".//x:Row", vns))
        col <- xml2::xml_text(xml2::xml_find_first(s, ".//x:Column", vns))
        paste(row, col, sep = ",")
      }, character(1))
      for (k in unique(keys[duplicated(keys)])) {
        idx <- which(keys == k)
        for (di in idx[-length(idx)]) xml2::xml_remove(shapes[[di]])
      }
    }
    
    xml2::write_xml(vdoc, vf, options = "no_declaration")
  }
  
  # re-zip in place, [Content_Types].xml first by convention
  old_wd <- getwd()
  setwd(work)
  on.exit(setwd(old_wd), add = TRUE)
  everything <- list.files(".", recursive = TRUE, all.files = TRUE, full.names = FALSE)
  everything <- everything[!grepl("^[.][.]?$", basename(everything))]
  ct <- "[Content_Types].xml"
  rest <- setdiff(everything, ct)
  zip::zip(path, files = c(ct, rest), recurse = FALSE)
  
  invisible(path)
}
