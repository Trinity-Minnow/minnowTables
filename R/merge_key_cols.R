#' Merge nested key columns vertically
#'
#' @description Vertically merges each of `cols` (in the order given) over the
#'   rows it spans, nested: a later column's run only continues while every
#'   earlier column in `cols` is also unchanged. For example, Year only merges
#'   within one Area's block, Station Group only within one (Area, Year).
#'   Plain `rle()` on a single column would instead merge across a parent
#'   boundary whenever the value happens to repeat there.
#'
#' @param wb An openxlsx Workbook object
#' @param sheetnm A character string of the sheet name
#' @param df A data frame in the same row order it was/will be written to `sheetnm`
#' @param cols A character vector of column names to merge, outermost first
#' @param start_row A number of the sheet row where the first row of `df` sits
#'
#' @return Called for its side effect on `wb`
#' @export
#' @importFrom openxlsx mergeCells
#'
#' @examples
#' \dontrun{
#' merge_key_cols(wb, "Table 1", df, cols = c("Area", "Year"), start_row = 4)
#' }
merge_key_cols = function(wb, sheetnm, df, cols, start_row){
  key = rep("", nrow(df))
  for (col in cols){
    key = paste(key, as.character(df[[col]]), sep="\r")
    r = rle(key)
    row = start_row
    for (len in r$lengths){
      end_row = row + len - 1
      if (len > 1) mergeCells(wb, sheetnm, cols=which(names(df)==col), rows=row:end_row)
      row = row + len
    }
  }
}
