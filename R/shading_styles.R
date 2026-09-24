#' Cell shading styles
#'
#' @description Fill styles reused for shading table cells: blue marks a
#'   positive/increase, green a negative/decrease.
#'
#' @format An openxlsx Style object
#' @name shading_styles
#' @examples
#' \dontrun{
#' openxlsx::addStyle(wb, "Table 1", blue_style, rows = 5, cols = 3, stack = TRUE)
#' }
NULL

#' @rdname shading_styles
#' @export
blue_style  <- openxlsx::createStyle(fgFill = "#9BC2E6")

#' @rdname shading_styles
#' @export
green_style <- openxlsx::createStyle(fgFill = "#C4D79B")
