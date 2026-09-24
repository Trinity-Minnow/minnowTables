#' Record note cells to patch with rich-text runs
#'
#' @description openxlsx has no rich-text API (`createStyle()` styles a whole
#'   cell), and Unicode bold/subscript look-alikes render badly or don't exist
#'   for every letter. Instead, write every note as one plain-text cell as
#'   usual, record here which cells need styled words, and after
#'   `saveWorkbook()` patch them into rich-text runs within the same cell with
#'   [apply_note_runs()].
#'
#' @details Appends to a list named `note_run_specs`, which the caller must
#'   create in the global environment first (`note_run_specs <- list()`) and
#'   later pass to [apply_note_runs()].
#'
#' @param sheet A character string of the sheet name
#' @param row A number of the row of the note cell (in column A)
#' @param text A character string of the full note text
#' @param regex A character string of the regular expression matching the
#'   text to style
#' @param style A list of formatting arguments passed to `openxlsx2::fmt_txt()`,
#'   e.g. `list(bold = TRUE)` or `list(vert_align = "subscript")`
#' @param perl A logical of whether `regex` is a Perl-compatible regex
#'
#' @return Called for its side effect on `note_run_specs`
#' @export
#'
#' @examples
#' \dontrun{
#' note_run_specs <- list()
#' record_note_runs("Table 1", row = 22, text = note[1], regex = "Bold font",
#'                  style = list(bold = TRUE))
#' }
record_note_runs <- function(sheet, row, text, regex, style, perl = TRUE) {
  note_run_specs[[length(note_run_specs) + 1]] <<- list(sheet = sheet, row = row, text = text, regex = regex, style = style, perl = perl)
}
