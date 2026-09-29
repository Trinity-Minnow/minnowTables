#' Write and format a table on a new worksheet (Calibri)
#'
#' @description Adds a worksheet and writes `xx` as a formatted table: Calibri
#'   font, hairline inner borders, bold header rows, thin separator lines,
#'   vertically merged grouping columns, a medium outer border, a bold wrapped
#'   caption in row 1 and note/footnote rows below the table.
#'
#' @param wb An openxlsx Workbook object
#' @param xx A data frame of the table body (written without column names)
#' @param sheetNm A character string of the new sheet name
#' @param header A list of header rows, each a character vector with one entry
#'   per column; element `i` is written to row `header.row[i]`
#' @param dat.row A number of the row the data starts on
#' @param header.row A numeric vector of the header rows
#' @param merged.cells A numeric vector of columns whose runs of identical values
#'   are merged vertically (`NA` to skip merging)
#' @param thin.lines.cols A numeric vector of columns that get a thin right border
#' @param thin.lines.rows A numeric vector of rows that get a thin bottom border
#' @param caption A character string of the table caption, written to row 1
#' @param col.widths An optional numeric vector of column widths (same units
#'   passed to `openxlsx::setColWidths()`). When given, the caption row is
#'   sized to its real wrapped line count with [estimate_wrapped_lines()]
#'   instead of staying clipped to one default-height row
#' @param note A character vector of note/footnote lines written below the table
#' @param shading.rows A number of the leading `note` lines that are
#'   shading-legend lines. Their text starts at column 2 (merged across the remaining columns) so column 1 stays clear for a colour swatch drawn there; all other note lines span the full table width
#'
#' @return Called for its side effect on `wb`
#' @export
#' @importFrom openxlsx addWorksheet writeData createStyle addStyle mergeCells setRowHeights
#' @importFrom dplyr ungroup mutate select all_of group_by across summarise arrange lag if_else
#' @importFrom tibble rownames_to_column
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#' excel.format.tcc(wb, df, "Table 1",
#'   header = list(c("Area", "Year", "Mean", "SD")),
#'   dat.row = 4, header.row = 3, merged.cells = c(1, 2),
#'   caption = "Table 1. Summary statistics")
#' }
excel.format.tcc <- function(wb, xx, sheetNm, header,
                             dat.row = max(header.row) + 1,          # FIX 1
                             header.row = c(3, 4),
                             merged.cells = c(1, 2),
                             thin.lines.cols = 1,
                             thin.lines.rows = max(header.row),      # FIX 5
                             caption,
                             col.widths = NULL,
                             note = paste0("Note: \"-\" = no data."),
                             shading.rows = 0,
                             grid.lines = TRUE) {

  font <- "Calibri"

  xx <- as.data.frame(dplyr::ungroup(xx))
  total_rows <- nrow(xx)
  total_cols <- ncol(xx)

  # FIX 1 - checks
  if (length(header) != length(header.row))
    stop("`header` has ", length(header), " rows but `header.row` has ", length(header.row), " rows.")
  if (dat.row <= max(header.row))
    stop("`dat.row` (", dat.row, ") must be below the last header row (", max(header.row), "), ",
         "otherwise the header overwrites the data.")

  top_row  <- min(header.row)              # FIX 5
  last_row <- dat.row + total_rows - 1     # FIX 4

  addWorksheet(wb, sheetNm, gridLines = grid.lines)

  writeData(wb, sheet = sheetNm, xx, startRow = dat.row, startCol = 1, colNames = FALSE)

  # font and alignment
  addStyle(wb, sheet = sheetNm, style = createStyle(fontName = font, fontSize = 10),
           rows = 1:last_row, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet = sheetNm, style = createStyle(halign = "center", valign = "center"),
           rows = 1:last_row, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)

  # hair inner lines
  addStyle(wb, sheetNm, style = createStyle(border = c("top", "bottom", "left", "right"), borderStyle = "hair"),
           rows = top_row:last_row, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)   # FIX 5

  # header rows
  for (i in seq_along(header)) {
    writeData(wb, sheetNm, t(header[[i]]), startRow = header.row[i], colNames = FALSE)
    addStyle(wb, sheetNm,
             style = createStyle(textDecoration = "bold", border = c("top", "bottom", "left", "right"),
                                 borderStyle = "hair", halign = "center"),
             rows = header.row[i], cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
  }

  # thin column lines (FIX 5 start at top_row; FIX 6 also the left edge of the next column)
  addStyle(wb, sheetNm, style = createStyle(border = "right", borderStyle = "thin"),
           rows = top_row:last_row, cols = thin.lines.cols, gridExpand = TRUE, stack = TRUE)
  next_cols <- thin.lines.cols[thin.lines.cols + 1 <= total_cols] + 1
  if (length(next_cols))
    addStyle(wb, sheetNm, style = createStyle(border = "left", borderStyle = "thin"),
             rows = top_row:last_row, cols = next_cols, gridExpand = TRUE, stack = TRUE)

  # thin row lines (FIX 6 also the top edge of the next row)
  addStyle(wb, sheetNm, style = createStyle(border = "bottom", borderStyle = "thin"),
           rows = thin.lines.rows, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
  next_rows <- thin.lines.rows[thin.lines.rows + 1 <= last_row] + 1
  if (length(next_rows))
    addStyle(wb, sheetNm, style = createStyle(border = "top", borderStyle = "thin"),
             rows = next_rows, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)

  # merged grouping columns (FIX 2 - consecutive runs, nested within the columns to the left)
  if (!all(is.na(merged.cells))) {
    for (i in merged.cells) {
      key <- do.call(paste, c(lapply(xx[seq_len(i)], as.character), sep = "\r"))
      run <- cumsum(c(TRUE, key[-1] != key[-length(key)]))
      ends   <- tapply(seq_along(run), run, max) + dat.row - 1
      starts <- tapply(seq_along(run), run, min) + dat.row - 1
      for (k in seq_along(ends)) {
        if (ends[k] > starts[k]) mergeCells(wb, sheetNm, cols = i, rows = starts[k]:ends[k])
        addStyle(wb, sheetNm, style = createStyle(border = "bottom", borderStyle = "thin"),
                 rows = ends[k], cols = i:total_cols, gridExpand = TRUE, stack = TRUE)
        if (ends[k] < last_row)   # FIX 6
          addStyle(wb, sheetNm, style = createStyle(border = "top", borderStyle = "thin"),
                   rows = ends[k] + 1, cols = i:total_cols, gridExpand = TRUE, stack = TRUE)
      }
    }
  }

  # medium outline (FIX 3 left edge to last_row; FIX 4 last_row from dat.row)
  addStyle(wb, sheetNm, style = createStyle(border = "top", borderStyle = "medium"),
           rows = top_row, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheetNm, style = createStyle(border = "bottom", borderStyle = "medium"),
           rows = last_row, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheetNm, style = createStyle(border = "left", borderStyle = "medium"),
           rows = top_row:last_row, cols = 1, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheetNm, style = createStyle(border = "right", borderStyle = "medium"),
           rows = top_row:last_row, cols = total_cols, gridExpand = TRUE, stack = TRUE)

  # caption
  caption_size <- 11
  caption_style <- createStyle(fontColour = "#000000", textDecoration = "bold", halign = "left", valign = "top",
                               wrapText = TRUE, fontName = font, fontSize = caption_size)
  writeData(wb, sheet = sheetNm, x = caption, startRow = 1, startCol = 1)
  mergeCells(wb, sheet = sheetNm, cols = 1:total_cols, rows = 1)
  addStyle(wb, sheet = sheetNm, style = caption_style, rows = 1, cols = 1, gridExpand = TRUE, stack = TRUE)
  if (!is.null(col.widths)) {
    caption_width_px <- sum(excel_width_to_px(col.widths))
    caption_lines <- estimate_wrapped_lines(caption, caption_width_px, font = font, size = caption_size)
    setRowHeights(wb, sheetNm, rows = 1, heights = pmax(15, caption_lines * 15 + 4))
  }

  # notes (FIX 4 position from last_row; wrapped, row height sized when col.widths is given)
  for (i in seq_along(note)) {
    note_row   <- last_row + 1 + i
    note_text  <- note[i]
    is_shading <- i <= shading.rows
    note_col   <- if (is_shading) 2 else 1
    merge_cols <- if (is_shading) 2:total_cols else 1:total_cols
    writeData(wb, sheet = sheetNm, x = note_text, startRow = note_row, startCol = note_col)
    mergeCells(wb, sheet = sheetNm, cols = merge_cols, rows = note_row)
    addStyle(wb, sheet = sheetNm,
             style = createStyle(fontName = font, fontSize = 9, halign = "left", valign = "top", wrapText = TRUE),
             rows = note_row, cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
    if (!is.null(col.widths)) {
      note_lines <- estimate_wrapped_lines(note_text, sum(excel_width_to_px(col.widths[merge_cols])),
                                           font = font, size = 9)
      setRowHeights(wb, sheetNm, rows = note_row, heights = pmax(13, note_lines * 13 + 4))
    }
  }
}
