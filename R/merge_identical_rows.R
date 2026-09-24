#' Merge runs of identical values in one column
#'
#' @description Vertically merges consecutive cells in a column that hold the
#'   same value. Each column is merged independently; use [merge_key_cols()]
#'   when a column's runs must not cross a parent column's boundaries.
#'
#' @param wb An openxlsx Workbook object
#' @param sheet A character string of the sheet name (or its index)
#' @param data A data frame whose rows line up with the sheet's rows
#' @param col_index A number of the column to merge (in both `data` and the sheet)
#' @param row_start A number of the first row to consider
#' @param row_end A number of the last row to consider
#'
#' @return Called for its side effect on `wb`
#' @export
#' @importFrom openxlsx mergeCells
#'
#' @examples
#' \dontrun{
#' merge_identical_rows(wb, "Table 1", df, col_index = 1, row_start = 4, row_end = 20)
#' }
merge_identical_rows <- function(wb, sheet, data, col_index, row_start, row_end) {

  values <- data[row_start:row_end, col_index]
  rle_vals <- rle(as.character(values))

  end_rows <- cumsum(rle_vals$lengths) + row_start - 1
  start_rows <- end_rows - rle_vals$lengths + 1

  for (i in seq_along(rle_vals$values)) {

    if (!is.na(rle_vals$values[i]) && rle_vals$lengths[i] > 1) {

      mergeCells(
        wb,
        sheet = sheet,
        cols = col_index,
        rows = start_rows[i]:end_rows[i]
      )
    }
  }
}
