#' Convert Excel column width units to pixels
#'
#' @description Applies Excel's own column-width-to-pixel formula, so the
#'   real rendered width of a (merged) range can be measured. Used with
#'   [estimate_wrapped_lines()] to size caption and note rows.
#'
#' @param width_units A numeric vector of column widths, in the same units
#'   passed to `openxlsx::setColWidths()`
#' @param mdw Maximum digit width of the workbook's default font, in pixels
#'   (7 for Calibri 11)
#'
#' @return A numeric vector of widths in pixels
#' @export
#'
#' @examples
#' excel_width_to_px(12)
#' sum(excel_width_to_px(c(12, 10, 10)))
excel_width_to_px <- function(width_units, mdw = 7) {
  floor(((256 * width_units + floor(128 / mdw)) / 256) * mdw)
}
