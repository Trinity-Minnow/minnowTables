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
excel.format.tcc = function(wb,
                        xx,
                        sheetNm,
                        header, # list of table column headers
                        dat.row=4, # row that data write on
                        header.row=c(3,4), # header row
                        merged.cells=c(1,2), # which cols need to be merged
                        thin.lines.cols= 1, # solid thin lines on columns
                        thin.lines.rows=3, # solid thin lines on rows
                        caption, # table caption
                        col.widths = NULL, # optional column widths (same units passed to setColWidths()) -
                                       # when given, the caption row is measured with estimate_wrapped_lines()
                                       # and sized to its real wrapped line count instead of staying clipped
                                       # to one default-height row
                        note =paste0("Note: \"-\" = no data."),# list of notes that need to be in
                        shading.rows=0 # number of LEADING `note` lines that are shading-legend lines -
                                       # their text starts at column 2 (merged 2:total_cols) so column 1
                                       # stays clear for a colour swatch the caller draws there; all other
                                       # note/footnote lines span the full table width (merged 1:total_cols)
                        ){
  
  addWorksheet(wb, sheetNm)  
  
  
  # Table dimensions
  total_rows <- nrow(xx)
  total_cols <- ncol(xx)
  
  
  # Write table starting at row 2 (row 1 could be caption)
  writeData(wb, sheet = sheetNm, xx, startRow = dat.row, startCol = 1,colNames = FALSE)
  
  # font Arial and size 10 --------------------------------------------------
  
  
  # Create a style for Arial, size 10
  arial_style <- createStyle(fontName = "Calibri", fontSize = 10)
  
  # Apply style to all table cells (headers + data)
  addStyle(
    wb, 
    sheet = sheetNm,
    style = arial_style,
    rows = 1:(total_rows +max(header.row)),  # adjust: +1 for header row, +1 if needed
    cols = 1:total_cols,
    gridExpand = TRUE, stack = TRUE
  )
  
  # centered all cells except 1st column
  
  # Apply style to all table cells (headers + data)
  addStyle(
    wb, 
    sheet = sheetNm,
    style =  createStyle(halign = "center", valign = "center"),
    rows = 1:(total_rows + max(header.row)),  # adjust: +1 for header row, +1 if needed
    cols = 1:total_cols,
    gridExpand = TRUE, stack = TRUE
  )
  
  
  # dotted line and centered for the whole table ----------------------------
  
  # 1️⃣ Dotted lines for the whole table (inside)
  dotted_style <- createStyle(
    border = c("top", "bottom", "left", "right"),
    borderStyle = "hair"
  )
  
  addStyle(wb,sheetNm, style = dotted_style,
           rows = 3:(total_rows+max(header.row)), cols = 1:total_cols,
           gridExpand = TRUE, stack = TRUE)
  
  
  # build  header adn bold them------------------------------------------------------
  for (i in 1:length(header)){
   # browser()
    #i=1
   header_top =  header[i][[1]]
  
  
   
   writeData(wb,sheetNm, t(header_top), startRow =header.row[i], colNames = FALSE)
   
   highlight_style <- createStyle(
    textDecoration = "bold",
    border = c("top", "bottom","left", "right"),
    borderStyle ="hair",
    halign = "center"
  )
  
  addStyle(wb, sheetNm, style =    highlight_style,
           rows = header.row[i], cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)
  
}

  
  

  
  # add solid lines on columns and rows-----------------------------------------------


  addStyle(wb,sheetNm, style = createStyle(border = c("right"),   borderStyle = "thin"),
           rows = 3:(total_rows+ max(header.row)), cols = thin.lines.cols, gridExpand = TRUE, stack = TRUE)
  
  
  addStyle(wb,sheetNm, style = createStyle(border = c("bottom"),   borderStyle = "thin"),
           rows = thin.lines.rows,  1:total_cols, gridExpand = TRUE, stack = TRUE)

  
  
  # merging cells -----------------------------------------------------------
  if (all(is.na(merged.cells))==F){
  for (i in merged.cells) {
    
    colnm <- names(xx)[i]
    
    temp <- xx %>%
      ungroup() %>%
      rownames_to_column("id") %>%
      mutate(id = as.numeric(id)) %>%
      dplyr::select(id, all_of(colnm)) %>%
      group_by(across(all_of(colnm))) %>%
      summarise(
        end = max(id) + dat.row - 1,
        .groups = "drop"
      ) %>%
      arrange(end) %>%
      mutate(
        start = lag(end) + 1 )%>%
      mutate( start = if_else(is.na(start), dat.row, start)  )
    
    for (k in seq_len(nrow(temp))) {
      
      start_row <- temp$start[k]
      end_row   <- temp$end[k]
      
      mergeCells(
        wb, sheetNm,
        cols = i,
        rows = start_row:end_row
      )
      
      addStyle(
        wb, sheetNm,
        style = createStyle(border = "bottom", borderStyle = "thin"),
        rows = end_row,
        cols = i:total_cols,
        gridExpand = TRUE,
        stack = TRUE
      )
    }
  }
  }
  
  
  
  # thick border ------------------------------------------------------------
  
  addStyle(wb, sheetNm, style = createStyle(border = c("top"),   borderStyle = "medium"),
           rows = min(header.row), cols = 1:total_cols, gridExpand = TRUE, stack = TRUE)


  
  addStyle(wb, sheetNm, style = createStyle(border = c("bottom"),   borderStyle = "medium"),
           rows = total_rows+max(header.row), cols = 1:total_cols, gridExpand = TRUE, stack = TRUE) 
  
  addStyle(wb, sheetNm, style = createStyle(border = c("left"),   borderStyle = "medium"),
           rows =min(header.row):(total_rows+min(header.row)), cols = 1, gridExpand = TRUE, stack = TRUE)
  
  addStyle(wb,sheetNm, style = createStyle(border = c("right"),   borderStyle = "medium"),
           rows = min(header.row):(total_rows+max(header.row)), cols = total_cols, gridExpand = TRUE, stack = TRUE)
  
  
  
  
  # caption on bold and size 11, wrapped to fit its merged width -------------

  caption_size <- 11

  # Create a style for the caption
  caption_style <-createStyle(fontColour ="#000000" ,   textDecoration = "bold", halign = "left", valign = "top",# "#135783"
                              wrapText = TRUE, fontName = "Calibri", fontSize = caption_size)


  # currently can do Table XX in minnow col and the rest in back so will get the whole caption in minnow and need to manual change in excel
  # Write the caption
  writeData(wb, sheet = sheetNm,
            x =caption,
            startRow = 1, startCol = 1)

  mergeCells(wb, sheet = sheetNm, cols = 1:(total_cols), rows = 1)

  # Apply the style to the caption cell
  addStyle(
    wb,
    sheet = sheetNm,
    style = caption_style,
    rows = 1,
    cols = 1,
    gridExpand = TRUE, stack = TRUE
  )

  # size the caption row to its real wrapped line count (same measure-then-size
  # approach used for the note/footnote rows below), so a long caption wraps
  # and shows in full instead of staying clipped to one default-height row
  if (!is.null(col.widths)) {
    caption_width_px <- sum(excel_width_to_px(col.widths))
    caption_lines <- estimate_wrapped_lines(caption, caption_width_px, size = caption_size)
    setRowHeights(wb, sheetNm, rows = 1, heights = pmax(15, caption_lines * 15 + 4))
  }


  # add footnote ------------------------------------------------------------
  for (i in 1:length(note)){

  # Add a note below the table (2 rows after the last row of table)
  note_row <- total_rows +  max(header.row) + 1 +i
  note_text <- note[i]

  # the first `shading.rows` lines are shading-legend lines: leave column 1
  # blank (for a colour swatch the caller draws there) and start the text at
  # column 2; every other note/footnote line spans the full table width
  is_shading <- i <= shading.rows
  note_col   <- if (is_shading) 2 else 1
  merge_cols <- if (is_shading) 2:total_cols else 1:total_cols

  # Write the note
  writeData(wb, sheet = sheetNm, x = note_text, startRow = note_row, startCol = note_col)

  # merge note across the available columns
  mergeCells(wb, sheet = sheetNm, cols = merge_cols, rows = note_row)

  # Optional: style the note (italic and smaller font)
  note_style <-createStyle(fontName = "Calibri", fontSize = 9,halign = "left")
  addStyle(wb, sheet = sheetNm, style = note_style, rows = note_row,
           cols = 1:(total_cols), gridExpand = TRUE, stack = TRUE)

  }
  #
  #
  # # add footnote 2------------------------------------------------------------
  # 
  # # Add a note below the table (2 rows after the last row of table)
  # note_row <- total_rows +    header.row +3
  # note_text <-  paste0("a Total density and biomass are reported for all organisms in the sample.")
  # 
  # # Write the note
  # writeData(wb, sheet = sheetNm, x = note_text, startRow = note_row, startCol = 1)
  # 
  # # Optional: merge note across all columns
  # mergeCells(wb, sheet = sheetNm, cols = 1:(total_cols), rows = note_row)
  # 
  # # Optional: style the note (italic and smaller font)
  # note_style <-createStyle(fontName = "Arial", fontSize = 9,halign = "left")
  # addStyle(wb, sheet = sheetNm, style = note_style, rows = note_row, 
  #          cols = 1:(total_cols), gridExpand = TRUE, stack = TRUE)
  
  # shdrows = xx %>%
  #   rownames_to_column("id")%>%
  #   filter(grepl("Exceedance", clean) ) %>%
  #   mutate(id=as.numeric(id)+row-1)%>%
  #   pull(id)
  # 
  # cols =4:total_cols
  # # 
  # 
  # for (x in shdrows){
  #   # x=45
  #   conditionalFormatting(
  #     wb, sheet = tbname,
  #     cols = cols,
  #     rows = x,
  #     type = "expression",
  #     rule = 'AND(TRIM(INDIRECT("RC",FALSE))<>"-", TRIM(INDIRECT("RC",FALSE))<>"0%")',
  #     style = createStyle(bgFill = "#A5A5A5")
  #   )
  # }
  # 
  
} 
