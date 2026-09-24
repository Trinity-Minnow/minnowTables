#' Build rich text with styled regex matches
#'
#' @description Builds one cell's worth of alternating plain/styled
#'   `openxlsx2::fmt_txt()` runs from every match of `regex` in `text`.
#'
#' @details Font and size are set explicitly on every run, not just the styled
#'   ones: rich-text runs don't inherit the cell's style, so a run without them
#'   renders in Excel's default Calibri 11.
#'
#' @param text A character string of the full cell text
#' @param regex A character string of the regular expression matching the
#'   text to style
#' @param style A list of formatting arguments passed to `openxlsx2::fmt_txt()`
#'   for matched runs
#' @param font A character string of the font name for all runs
#' @param size A number of the font size for all runs
#' @param perl A logical of whether `regex` is a Perl-compatible regex
#'
#' @return An openxlsx2 `fmt_txt` object
#' @export
#'
#' @examples
#' \dontrun{
#' richtext_with_highlight("Bold font = significant", "Bold font", list(bold = TRUE))
#' }
richtext_with_highlight <- function(text, regex, style, font = "Calibri", size = 9, perl = TRUE) {
  m <- gregexpr(regex, text, perl = perl)[[1]]
  if (m[1] == -1L) return(openxlsx2::fmt_txt(text, font = font, size = size))
  starts <- as.integer(m)
  ends   <- starts + attr(m, "match.length") - 1L
  pieces <- list()
  pos <- 1L
  for (i in seq_along(starts)) {
    if (starts[i] > pos) {
      pieces[[length(pieces) + 1]] <- openxlsx2::fmt_txt(substr(text, pos, starts[i] - 1L), font = font, size = size)
    }
    run_args <- c(list(x = substr(text, starts[i], ends[i]), font = font, size = size), style)
    pieces[[length(pieces) + 1]] <- do.call(openxlsx2::fmt_txt, run_args)
    pos <- ends[i] + 1L
  }
  if (pos <= nchar(text)) {
    pieces[[length(pieces) + 1]] <- openxlsx2::fmt_txt(substr(text, pos, nchar(text)), font = font, size = size)
  }
  Reduce(`+`, pieces)
}
