#' Superscript a footnote's leading marker letter
#'
#' @description Footnote text is marked with a leading `"<letter> "` (e.g.
#'   `"a Years that share..."`). This superscripts just that leading marker so
#'   it matches the markers in the header/caption text, leaving entries with no
#'   marker (e.g. `"Area x Year..."`) untouched.
#'
#' @param txt A character vector of footnote texts
#'
#' @return A character vector the same length as `txt`
#' @export
#'
#' @examples
#' superscript_leading_marker(c("a Years that share a letter", "Area x Year interaction"))
superscript_leading_marker <- function(txt) {
  # per-element loop, not vectorized ifelse() - ifelse() evaluates both branches
  # for every element, which would call sup() on non-marker leading characters
  # (e.g. the "A" in "Area x Year...") and error on an unmapped letter
  vapply(txt, function(t) {
    if (is.na(t) || !grepl("^[a-z] ", t)) return(t)
    paste0(sup(substr(t, 1, 1)), substring(t, 2))
  }, character(1), USE.NAMES = FALSE)
}
