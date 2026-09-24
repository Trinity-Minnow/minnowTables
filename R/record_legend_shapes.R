#' Record legend swatch positions for later drawing
#'
#' @description openxlsx has no drawing/shape API, so legend swatches can't be
#'   created while building the workbook. Instead each table's swatch
#'   positions/colours are recorded here, and the actual text-box shapes are
#'   drawn into the saved .xlsx afterwards by [add_legend_swatch_shapes()].
#'
#' @details Appends to a list named `legend_shape_specs`, which the caller must
#'   create in the global environment first (`legend_shape_specs <- list()`)
#'   and later pass to [add_legend_swatch_shapes()].
#'
#' @param sheet A character string of the sheet name
#' @param rows A numeric vector of the rows to anchor a swatch on (column 1)
#' @param colours A character vector of fill colours, one per row
#'
#' @return Called for its side effect on `legend_shape_specs`
#' @export
#'
#' @examples
#' \dontrun{
#' legend_shape_specs <- list()
#' record_legend_shapes("Table 1", rows = c(22, 23), colours = c("#9BC2E6", "#C4D79B"))
#' }
record_legend_shapes <- function(sheet, rows, colours) {
  legend_shape_specs[[length(legend_shape_specs) + 1]] <<- list(sheet = sheet, rows = rows, colours = colours)
}
