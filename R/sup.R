#' Unicode superscript letters
#'
#' @description openxlsx has no rich-text API, so footnote-marker letters are
#'   superscripted using their dedicated Unicode code points instead. These
#'   render as true superscript glyphs in plain text. There is no superscript
#'   "q".
#'
#' @format A named character vector mapping lowercase letters to superscript
#'   glyphs
#' @export
superscript_letters <- c(a="\u1d43", b="\u1d47", c="\u1d9c", d="\u1d48", e="\u1d49",
                         f="\u1da0", g="\u1d4d", h="\u02b0", i="\u2071", j="\u02b2",
                         k="\u1d4f", l="\u02e1", m="\u1d50", n="\u207f", o="\u1d52",
                         p="\u1d56", r="\u02b3", s="\u02e2", t="\u1d57", u="\u1d58",
                         v="\u1d5b", w="\u02b7", x="\u02e3", y="\u02b8", z="\u1dbb")

#' Convert letters to Unicode superscript
#'
#' @description Looks up footnote-marker letters in [superscript_letters].
#'
#' @param letter A character vector of single lowercase letters
#'
#' @return A character vector of superscript glyphs; errors if a letter has no
#'   mapping
#' @export
#'
#' @examples
#' paste0("Year", sup("a"))
#' sup(c("a", "b"))
sup <- function(letter) {
  out <- superscript_letters[letter]
  if (anyNA(out)) stop("No superscript mapping for letter(s): ", paste(letter[is.na(out)], collapse=", "))
  unname(out)
}
