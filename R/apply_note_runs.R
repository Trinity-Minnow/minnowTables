#' Apply recorded rich-text runs to a saved workbook
#'
#' @description Reopens a saved .xlsx with openxlsx2 and rewrites each note cell
#'   recorded by [record_note_runs()] as rich text built by
#'   [richtext_with_highlight()]. Run after `openxlsx::saveWorkbook()`.
#'
#' @param path A character string of the saved .xlsx file path
#' @param specs A list of specs built by [record_note_runs()], usually
#'   `note_run_specs`
#'
#' @return `path`, invisibly
#' @export
#'
#' @examples
#' \dontrun{
#' openxlsx::saveWorkbook(wb, "tables.xlsx", overwrite = TRUE)
#' apply_note_runs("tables.xlsx", note_run_specs)
#' }
apply_note_runs <- function(path, specs) {
  if (length(specs) == 0) return(invisible(path))
  wb2 <- openxlsx2::wb_load(path)
  for (spec in specs) {
    rt <- richtext_with_highlight(spec$text, spec$regex, spec$style, perl = spec$perl)
    wb2$add_data(sheet = spec$sheet, dims = paste0("A", spec$row), x = rt)
  }
  wb2$save(path)
  invisible(path)
}
