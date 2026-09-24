#' Set a worksheet's tab colour
#'
#' @description openxlsx only exposes tab colour via
#'   `addWorksheet(tabColour = ...)` at sheet-creation time. This sets or
#'   replaces it afterward by editing the sheet's `sheetPr` XML directly,
#'   preserving any other `sheetPr` content already there (e.g. `<outlinePr/>`).
#'
#' @param wb An openxlsx Workbook object
#' @param sheetName A character string of the sheet name
#' @param colour A character string of the colour as `"#RRGGBB"` or `"#AARRGGBB"`
#'
#' @return `wb`, invisibly
#' @export
#'
#' @examples
#' \dontrun{
#' set_tab_colour(wb, "Table 1", "#9BC2E6")
#' }
set_tab_colour <- function(wb, sheetName, colour) {
  idx <- which(names(wb) == sheetName)
  if (length(idx) == 0) stop("Sheet '", sheetName, "' not found in workbook.")

  argb <- toupper(gsub("^#", "", colour))
  if (nchar(argb) == 6) argb <- paste0("FF", argb)  # RGB -> ARGB (opaque) if no alpha given

  cur <- wb$worksheets[[idx]]$sheetPr
  cur <- if (length(cur) == 0 || !nzchar(cur)) "<sheetPr/>" else cur

  node <- xml2::read_xml(cur)
  existing <- xml2::xml_find_first(node, "tabColor")
  if (!inherits(existing, "xml_missing")) xml2::xml_remove(existing)
  xml2::xml_add_child(node, "tabColor", rgb = argb)

  wb$worksheets[[idx]]$sheetPr <- trimws(as.character(node, options = "no_declaration"))
  invisible(wb)
}
