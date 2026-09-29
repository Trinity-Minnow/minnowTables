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
  if (nchar(argb) == 6) argb <- paste0("FF", argb)

  # everything openxlsx has stored for this sheet, merged into one <sheetPr> node
  cur  <- wb$worksheets[[idx]]$sheetPr
  cur  <- cur[nzchar(trimws(cur))]
  node <- xml2::read_xml("<sheetPr/>")
  for (el in cur) {
    x <- xml2::read_xml(el)
    if (xml2::xml_name(x) == "sheetPr") {
      xml2::xml_attrs(node) <- c(xml2::xml_attrs(node), xml2::xml_attrs(x))
      for (ch in xml2::xml_children(x)) xml2::xml_add_child(node, ch)
    } else {
      xml2::xml_add_child(node, x)
    }
  }

  # children: drop any old tabColor, add the new one, keep schema order, drop duplicates
  kids <- xml2::xml_children(node)
  kids <- kids[xml2::xml_name(kids) != "tabColor"]
  child_txt  <- c(sprintf('<tabColor rgb="%s"/>', argb),
                  vapply(kids, function(k) as.character(k, options = "no_declaration"), character(1)))
  child_name <- c("tabColor", xml2::xml_name(kids))
  ord <- order(match(child_name, c("tabColor", "outlinePr", "pageSetUpPr"), nomatch = 99))
  child_txt <- unique(child_txt[ord])

  attrs <- xml2::xml_attrs(node)
  wb$worksheets[[idx]]$sheetPr <- if (length(attrs) == 0) {
    child_txt   # openxlsx wraps these in <sheetPr> when saving, and pageSetup() can still append to it
  } else {
    paste0("<sheetPr ", paste0(names(attrs), '="', attrs, '"', collapse = " "), ">",
           paste(child_txt, collapse = ""), "</sheetPr>")
  }
  invisible(wb)
}
