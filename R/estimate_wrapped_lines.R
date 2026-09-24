#' Estimate how many lines text wraps to in an Excel cell
#'
#' @description Measures each text's real wrapped line count using the actual
#'   font/size and merged width in pixels, greedily word-wrapping the same way
#'   Excel does. A fixed "chars per line" ratio either clips a long note or
#'   leaves blank space at the bottom of the row; this avoids both.
#'
#' @details Windows only: text widths are measured with
#'   `grDevices::windowsFonts()` on an off-screen PNG device.
#'
#' @param texts A character vector of texts to measure
#' @param total_width_px A number of the available width in pixels, e.g.
#'   `sum(excel_width_to_px(col.widths))`
#' @param font A character string of the font name
#' @param size A number of the font size in points
#' @param margin_px A number of pixels subtracted from `total_width_px` for
#'   cell padding
#'
#' @return An integer vector of wrapped line counts, one per element of `texts`
#' @export
#'
#' @examples
#' \dontrun{
#' estimate_wrapped_lines("Note: \"-\" = no data.", sum(excel_width_to_px(c(12, 10))))
#' }
estimate_wrapped_lines <- function(texts, total_width_px, font = "Calibri", size = 9, margin_px = 10) {
  grDevices::windowsFonts(Calibri = grDevices::windowsFont(font))
  tmp <- tempfile(fileext = ".png")
  grDevices::png(tmp, width = 1600, height = 200, pointsize = size, res = 96)
  on.exit({ grDevices::dev.off(); unlink(tmp) })
  graphics::par(family = "Calibri")
  usable_px <- total_width_px - margin_px
  space_w <- graphics::strwidth(" ", units = "inches") * 96
  vapply(texts, function(text) {
    words <- strsplit(text, " ")[[1]]
    lines <- 1L
    cur <- 0
    for (w in words) {
      ww  <- graphics::strwidth(w, units = "inches") * 96
      add <- if (cur == 0) ww else ww + space_w
      if (cur + add > usable_px) { lines <- lines + 1L; cur <- ww } else { cur <- cur + add }
    }
    lines
  }, integer(1))
}
