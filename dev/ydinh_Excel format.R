

# real per-note wrapped-line count ------------------------------------------
# a fixed "chars per line" ratio is only ever a rough guess at how much text
# fits per line - it either clips a long note (too generous a ratio) or
# leaves blank space at the bottom of the row (too stingy a ratio,
# overcounting the lines actually needed). This instead measures each note's
# real wrapped line count using the workbook's actual font/size and the
# actual merged width in pixels (via Excel's own column-width-to-pixel
# formula), greedily word-wrapping the same way Excel does. Shared by every
# script that sizes note/footnote rows (originally written for the 1-way/
# 2-way ANOVA table scripts; reused as-is by the appendix summary-stats
# table script since the wrapping math itself is generic).
excel_width_to_px <- function(width_units, mdw = 7) {
  floor(((256 * width_units + floor(128 / mdw)) / 256) * mdw)
}

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
  


excel.format.mn = function(wb,
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
                                           # their text is written to column 2 instead of column 1, leaving
                                           # column 1 clear for a colour swatch the caller draws there
){
  
  addWorksheet(wb, sheetNm)  
  
  
  # Table dimensions
  total_rows <- nrow(xx)
  total_cols <- ncol(xx)
  
  
  # Write table starting at row 2 (row 1 could be caption)
  writeData(wb, sheet = sheetNm, xx, startRow = dat.row, startCol = 1,colNames = FALSE)
  
  # font Arial and size 10 --------------------------------------------------
  
  
  # Create a style for Arial, size 10
  arial_style <- createStyle(fontName = "Arial", fontSize = 10)
  
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
                              wrapText = TRUE, fontName = "Arial", fontSize = caption_size)


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
    # column 2; every other note/footnote line is written to column 1 as before
    is_shading <- i <= shading.rows
    note_col   <- if (is_shading) 2 else 1
    merge_cols <- if (is_shading) 2:total_cols else 1:total_cols

    # Write the note
    writeData(wb, sheet = sheetNm, x = note_text, startRow = note_row, startCol = note_col)

    # merge note across the available columns - starting at note_col, not a
    # fixed column 2, so a non-shading note (written to column 1) actually gets
    # merged with column 1 instead of leaving it as a separate, blank cell
    mergeCells(wb, sheet = sheetNm, cols = merge_cols, rows = note_row)
    
    # Optional: style the note (italic and smaller font)
    note_style <-createStyle(fontName = "Arial", fontSize = 9,halign = "left")
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
# merged ------------------------------------------------------------------
# Area (col 1)
# area_rle <- rle(as.character(sum_fmt$Area))
# r <- data_start
# for (len in area_rle$lengths) {
#   end_row <- r + len - 1
#   if (len > 1) mergeCells(wb_fmt, "sum_stats", cols = 1, rows = r:end_row)
#   if (end_row < last_row)
#     addStyle(wb_fmt, "sum_stats",
#              createStyle(border = "bottom", borderStyle = "thin", borderColour = border_col),
#              rows = end_row, cols = 1:10, stack = TRUE)
#   r <- r + len
# }
# 
# # Endpoint (col 2)
# ep_key <- paste(sum_fmt$Area, sum_fmt$Endpoint)
# ep_rle <- rle(ep_key)
# r <- data_start
# for (len in ep_rle$lengths) {
#   end_row <- r + len - 1
#   if (len > 1) mergeCells(wb_fmt, "sum_stats", cols = 2, rows = r:end_row)
#   if (end_row < last_row)
#     addStyle(wb_fmt, "sum_stats",
#              createStyle(border = "bottom", borderStyle = "thin", borderColour = border_col),
#              rows = end_row, cols = 2:10, stack = TRUE)
#   r <- r + len
# }
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

# merged ------------------------------------------------------------------
# merge_key_cols(): vertically merges each of `cols` (in the order given) over
# the rows it spans, nested - a later column's run only continues while every
# earlier column in `cols` is also unchanged, so e.g. Year only merges within
# one Area's block, Station Group only within one (Area, Year), etc. Plain
# rle() on a single column would instead merge across a parent boundary
# whenever that column's value happens to repeat there (e.g. the same Year
# ending one Area's rows and starting the next). df must already be in the
# same row order it was/will be written to sheetnm at start_row.
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


# set tab colour for excel ------------------------------------------------
# openxlsx only exposes tab colour via addWorksheet(tabColour=...) at sheet-
# creation time; this sets/replaces it afterward by editing the sheet's
# sheetPr XML directly, preserving any other sheetPr content already there
# (e.g. Cover Sheet/QA-QC template sheets carry an <outlinePr/>).
set_tab_colour <- function(wb, sheetName, colour) {
  idx <- which(names(wb) == sheetName)
  if (length(idx) == 0) stop("Sheet '", sheetName, "' not found in workbook.")

  argb <- toupper(gsub("^#", "", colour))
  if (nchar(argb) == 6) argb <- paste0("FF", argb)  # RGB -> ARGB (opaque) if no alpha given

  cur <- wb$worksheets[[idx]]$sheetPr
  cur <- if (length(cur) == 0 || !nzchar(cur)) "<sheetPr/>" else cur

  node <- xml2::read_xml(cur)
  existing <- xml2::xml_find_first(node, "tabColor")
  if (!inherits(existing, "xml_missing")) xml2::xml_remove(existing)
  xml2::xml_add_child(node, "tabColor", rgb = argb)

  wb$worksheets[[idx]]$sheetPr <- trimws(as.character(node, options = "no_declaration"))
  invisible(wb)
}


# floating legend swatches (true Excel text-box shapes) --------------------
# openxlsx has no drawing/shape API, so the grey/blue/green legend swatches
# in column 1 can't be created while building wb in R - instead each table's
# swatch positions/colours are recorded here via record_legend_shapes(), and
# the actual 0.15in x 0.5in text-box shapes are drawn directly into the saved
# .xlsx package afterwards by add_legend_swatch_shapes() (below, called right
# before repair_workbook_package()), the same raw-zip-surgery approach that
# function already uses.
# legend_shape_specs <- list()
record_legend_shapes <- function(sheet, rows, colours) {
  legend_shape_specs[[length(legend_shape_specs) + 1]] <<- list(sheet = sheet, rows = rows, colours = colours)
}

# draw the legend swatch text boxes -----------------------------------------
# Adds one true floating Excel text-box shape (0.15in x 0.5in, solid-filled)
# per row recorded via record_legend_shapes(), anchored at column 1 of the
# given row on each sheet. openxlsx has no shape/text-box API, so this works
# directly on the just-saved .xlsx package: unzip, write a drawingN.xml part
# per sheet holding its shapes, wire it up via [Content_Types].xml + the
# sheet's _rels + a <drawing r:id=".."/> tag on the sheet's own xml, re-zip.
# Must run before repair_workbook_package() so the new drawing parts are
# already in place when that function's orphan-relationship check runs (it
# keeps any worksheet relationship whose r:id is actually referenced in the
# sheet xml, which ours is).
add_legend_swatch_shapes <- function(path, specs, width_in = 0.5, height_in = 0.15) {
  if (length(specs) == 0) return(invisible(path))
  path <- normalizePath(path, mustWork = TRUE)
  work <- tempfile("xlsxshapes_")
  dir.create(work)
  on.exit(unlink(work, recursive = TRUE, force = TRUE), add = TRUE)
  utils::unzip(path, exdir = work)
  
  cx_num <- round(width_in  * 914400)
  cy_num <- round(height_in * 914400)
  cx <- format(cx_num, scientific = FALSE)
  cy <- format(cy_num, scientific = FALSE)
  
  # centre the swatch within column 1's cell instead of anchoring it to the
  # cell's top-left corner. Column 1 is set to a width of 12 (characters) via
  # setColWidths() at the call site, which Excel/openxlsx renders as ~12.71
  # characters -> 89px -> 847725 EMU under the workbook's default Calibri 11
  # font; the note rows carry no per-row height override, so they use the
  # workbook's default row height (15pt -> 190500 EMU). colOff/rowOff below
  # shift the box by half the leftover space on each axis.
  col1_width_emu <- 847725
  row_height_emu <- 190500
  col_off <- format(max(0, round((col1_width_emu - cx_num) / 2)), scientific = FALSE)
  row_off <- format(max(0, round((row_height_emu - cy_num) / 2)), scientific = FALSE)
  
  read_txt  <- function(f) paste(readLines(f, warn = FALSE), collapse = "\n")
  write_txt <- function(txt, f) writeLines(txt, f, useBytes = TRUE)
  
  # sheet name -> r:id (from xl/workbook.xml) -> target path (from its rels)
  wb_txt   <- read_txt(file.path(work, "xl/workbook.xml"))
  sheet_tags <- regmatches(wb_txt, gregexpr("<sheet [^>]*/>", wb_txt))[[1]]
  name_to_rid <- setNames(sub('.*r:id="([^"]*)".*', "\\1", sheet_tags),
                          sub('.*name="([^"]*)".*', "\\1", sheet_tags))
  
  wbrels_txt <- read_txt(file.path(work, "xl/_rels/workbook.xml.rels"))
  rel_tags   <- regmatches(wbrels_txt, gregexpr("<Relationship[^>]*/>", wbrels_txt))[[1]]
  rid_to_target <- setNames(sub('.*Target="([^"]*)".*', "\\1", rel_tags),
                            sub('.*Id="([^"]*)".*', "\\1", rel_tags))
  
  dir.create(file.path(work, "xl/drawings"), showWarnings = FALSE)
  dir.create(file.path(work, "xl/drawings/_rels"), showWarnings = FALSE)
  
  ct_file <- file.path(work, "[Content_Types].xml")
  ct_txt  <- read_txt(ct_file)
  
  # next free "drawingN.xml" name - considers both real files on disk AND names
  # already declared in Content_Types (openxlsx pre-declares a drawing part for
  # sheets it never actually draws on; see the reuse logic below), so a fresh
  # name here can never collide with one of those not-yet-materialised ones
  next_drawing_n <- function() {
    declared <- regmatches(ct_txt, gregexpr("/xl/drawings/drawing[0-9]+[.]xml", ct_txt))[[1]]
    declared_nums <- as.integer(sub(".*drawing([0-9]+)[.]xml", "\\1", declared))
    existing <- list.files(file.path(work, "xl/drawings"), pattern = "^drawing[0-9]+[.]xml$")
    existing_nums <- as.integer(sub("drawing([0-9]+)[.]xml", "\\1", existing))
    suppressWarnings(max(c(declared_nums, existing_nums, 0), na.rm = TRUE)) + 1
  }
  
  for (spec in specs) {
    target <- rid_to_target[[name_to_rid[[spec$sheet]]]]
    if (is.null(target)) next  # sheet not found - skip rather than corrupt the package
    sheet_path <- file.path(work, "xl/worksheets", basename(target))
    rels_dir   <- file.path(work, "xl/worksheets/_rels")
    dir.create(rels_dir, showWarnings = FALSE)
    sheet_rels_path <- file.path(rels_dir, paste0(basename(target), ".rels"))
    
    # a worksheet may carry at most one "drawing"-type relationship. openxlsx
    # (4.2.8.1) leaves one such relationship on every worksheet it writes whose
    # target file it never actually creates (see repair_workbook_package()'s
    # note on this same quirk below) - reuse that existing, still-unmaterialised
    # relationship/target name instead of minting a second drawing relationship,
    # which Excel rejects as invalid and would otherwise corrupt the file.
    reuse_rid <- NA_character_
    reuse_target <- NA_character_
    if (file.exists(sheet_rels_path)) {
      sheet_rels_txt0 <- read_txt(sheet_rels_path)
      rel_tags0  <- regmatches(sheet_rels_txt0, gregexpr("<Relationship[^>]*/>", sheet_rels_txt0))[[1]]
      is_drawing <- grepl('Type="[^"]*/relationships/drawing"', rel_tags0)
      if (any(is_drawing)) {
        tag0 <- rel_tags0[which(is_drawing)[1]]
        reuse_rid    <- sub('.*Id="([^"]*)".*', "\\1", tag0)
        reuse_target <- sub('.*Target="([^"]*)".*', "\\1", tag0)
      }
    }
    reuse_path <- if (!is.na(reuse_target)) normalizePath(file.path(work, "xl/worksheets", reuse_target), mustWork = FALSE) else NA
    reuse_ok   <- !is.na(reuse_rid) && !file.exists(reuse_path)  # only reuse if still never actually created
    
    if (reuse_ok) {
      new_rid      <- reuse_rid
      drawing_path <- reuse_path
      drawing_name <- basename(drawing_path)
      dir.create(dirname(drawing_path), showWarnings = FALSE, recursive = TRUE)
    } else {
      drawing_name <- paste0("drawing", next_drawing_n(), ".xml")
      drawing_path <- file.path(work, "xl/drawings", drawing_name)
    }
    
    shapes_xml <- vapply(seq_along(spec$rows), function(i) {
      row0 <- spec$rows[i] - 1L  # OOXML anchors are 0-indexed
      fill <- gsub("^#", "", spec$colours[i])
      paste0(
        '<xdr:oneCellAnchor>',
        '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>', col_off, '</xdr:colOff>',
        '<xdr:row>', row0, '</xdr:row><xdr:rowOff>', row_off, '</xdr:rowOff></xdr:from>',
        '<xdr:ext cx="', cx, '" cy="', cy, '"/>',
        '<xdr:sp macro="" textlink="">',
        '<xdr:nvSpPr><xdr:cNvPr id="', 1000 + i, '" name="Legend Swatch ', tools::file_path_sans_ext(drawing_name), '-', i, '"/>',
        '<xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>',
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="', cx, '" cy="', cy, '"/></a:xfrm>',
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
        '<a:solidFill><a:srgbClr val="', fill, '"/></a:solidFill>',
        '<a:ln><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln>',
        '</xdr:spPr>',
        '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" rtlCol="0" anchor="ctr"/>',
        '<a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></xdr:txBody>',
        '</xdr:sp>',
        '<xdr:clientData/>',
        '</xdr:oneCellAnchor>'
      )
    }, character(1))
    
    write_txt(paste0(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
      paste(shapes_xml, collapse = ""),
      '</xdr:wsDr>'
    ), drawing_path)
    
    drawing_rels_path <- file.path(dirname(drawing_path), "_rels", paste0(drawing_name, ".rels"))
    dir.create(dirname(drawing_rels_path), showWarnings = FALSE, recursive = TRUE)
    write_txt(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      drawing_rels_path
    )
    
    # register the drawing part's content type - openxlsx may have already
    # declared this exact PartName (see the reuse note above), and duplicate
    # Overrides for one PartName are themselves invalid, so only add if missing
    ct_partname <- paste0('PartName="/xl/drawings/', drawing_name, '"')
    if (!grepl(ct_partname, ct_txt, fixed = TRUE)) {
      ct_txt <- sub("</Types>",
                    paste0('<Override ', ct_partname,
                           ' ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'),
                    ct_txt, fixed = TRUE)
    }
    
    # wire the worksheet -> drawing relationship, unless we're reusing an
    # already-existing one (reuse_ok) above
    if (!reuse_ok) {
      if (file.exists(sheet_rels_path)) {
        sheet_rels_txt <- read_txt(sheet_rels_path)
        existing_ids <- sub('.*Id="([^"]*)".*', "\\1",
                            regmatches(sheet_rels_txt, gregexpr("<Relationship[^>]*/>", sheet_rels_txt))[[1]])
      } else {
        sheet_rels_txt <- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        existing_ids <- character(0)
      }
      n <- 0
      repeat { n <- n + 1; new_rid <- paste0("rIdSwatch", n); if (!(new_rid %in% existing_ids)) break }
      sheet_rels_txt <- sub("</Relationships>",
                            paste0('<Relationship Id="', new_rid,
                                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"',
                                   ' Target="../drawings/', drawing_name, '"/></Relationships>'),
                            sheet_rels_txt, fixed = TRUE)
      write_txt(sheet_rels_txt, sheet_rels_path)
    }
    
    # add <drawing r:id=".."/> to the worksheet xml, right before tableParts/
    # extLst if present (schema order), else right before </worksheet>
    sheet_txt <- read_txt(sheet_path)
    if (!grepl("<drawing ", sheet_txt, fixed = TRUE)) {
      tag <- paste0('<drawing r:id="', new_rid, '"/>')
      if (grepl("<extLst>", sheet_txt, fixed = TRUE)) {
        sheet_txt <- sub("<extLst>", paste0(tag, "<extLst>"), sheet_txt, fixed = TRUE)
      } else if (grepl("<tableParts", sheet_txt, fixed = TRUE)) {
        sheet_txt <- sub("<tableParts", paste0(tag, "<tableParts"), sheet_txt, fixed = TRUE)
      } else {
        sheet_txt <- sub("</worksheet>", paste0(tag, "</worksheet>"), sheet_txt, fixed = TRUE)
      }
      write_txt(sheet_txt, sheet_path)
    }
  }
  
  write_txt(ct_txt, ct_file)
  
  old_wd <- getwd()
  setwd(work)
  on.exit(setwd(old_wd), add = TRUE)
  everything <- list.files(".", recursive = TRUE, all.files = TRUE, full.names = FALSE)
  everything <- everything[!grepl("^[.][.]?$", basename(everything))]
  ct <- "[Content_Types].xml"
  rest <- setdiff(everything, ct)
  zip::zip(path, files = c(ct, rest), recurse = FALSE)
  
  invisible(path)
}


# repair the saved package - openxlsx (4.2.8.1) leaves two kinds of invalid,
# Excel-repair-triggering junk behind after this workbook's save:
#  1. any part with no declared content type (e.g. this template historically
#     carried a stray "[trash]/*.dat" folder from a prior Excel edit)
#  2. orphaned worksheet relationships - a drawing/vmlDrawing relationship
#     entry pointing at a file that was never created, left on every sheet
#     that was added via addWorksheet() to a workbook whose template already
#     has cell comments elsewhere (Cover Sheet/QA-QC here)
# Both are unreferenced by any actual cell content, so removing them is safe;
# see repair_workbook_package() below.
repair_workbook_package <- function(path) {
  path <- normalizePath(path, mustWork = TRUE)  # must be absolute - the function setwd()s into a scratch dir below
  work <- tempfile("xlsxrepair_")
  dir.create(work)
  on.exit(unlink(work, recursive = TRUE, force = TRUE), add = TRUE)
  
  utils::unzip(path, exdir = work)
  
  # (1) strip any part with no declared content type
  ct_file <- file.path(work, "[Content_Types].xml")
  ct_txt  <- paste(readLines(ct_file, warn = FALSE), collapse = "")
  declared_ext <- regmatches(ct_txt, gregexpr('Default Extension="[^"]*"', ct_txt))[[1]]
  declared_ext <- sub('Default Extension="([^"]*)"', "\\1", declared_ext)
  declared_override <- regmatches(ct_txt, gregexpr('PartName="[^"]*"', ct_txt))[[1]]
  declared_override <- sub('PartName="([^"]*)"', "\\1", declared_override)
  
  all_files <- list.files(work, recursive = TRUE, full.names = TRUE)
  for (f in all_files) {
    rel <- sub(work, "", f, fixed = TRUE)
    rel <- paste0("/", gsub("\\\\", "/", rel))
    rel <- sub("^/+", "/", rel)
    ext <- tolower(sub(".*[.]", "", rel))
    if (!(ext %in% declared_ext) && !(rel %in% declared_override)) unlink(f)
  }
  repeat {
    empty_dirs <- list.dirs(work, recursive = TRUE, full.names = TRUE)
    empty_dirs <- empty_dirs[vapply(empty_dirs, function(d) length(list.files(d, all.files = TRUE, no.. = TRUE)) == 0, logical(1))]
    if (length(empty_dirs) == 0) break
    unlink(empty_dirs, recursive = TRUE)
  }
  
  # (2) drop orphaned worksheet relationships (missing target AND unused r:id)
  rels_dir <- file.path(work, "xl/worksheets/_rels")
  if (dir.exists(rels_dir)) {
    rels_files <- list.files(rels_dir, pattern = "[.]rels$", full.names = TRUE)
    for (relf in rels_files) {
      sheet_base <- sub("[.]rels$", "", basename(relf))
      sheetxml   <- file.path(work, "xl/worksheets", sheet_base)
      if (!file.exists(sheetxml)) next
      sheet_txt  <- paste(readLines(sheetxml, warn = FALSE), collapse = "\n")
      
      txt <- paste(readLines(relf, warn = FALSE), collapse = "")
      rel_elems <- regmatches(txt, gregexpr("<Relationship[[:space:]]+[^>]*/>", txt, perl = TRUE))[[1]]
      if (length(rel_elems) == 0) next
      
      keep <- character(0)
      for (el in rel_elems) {
        id     <- sub('.*Id="([^"]*)".*', "\\1", el)
        target <- sub('.*Target="([^"]*)".*', "\\1", el)
        resolved <- normalizePath(file.path(work, "xl/worksheets", target), mustWork = FALSE)
        target_exists <- file.exists(resolved)
        id_used <- grepl(paste0('r:id="', id, '"'), sheet_txt, fixed = TRUE)
        if (!target_exists && !id_used) next
        keep <- c(keep, el)
      }
      
      if (length(keep) != length(rel_elems)) {
        header <- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        writeLines(paste0(header, paste(keep, collapse = ""), "</Relationships>"), relf, useBytes = TRUE)
      }
    }
  }
  
  # (3) de-duplicate legacy comment VML shapes - openxlsx's load/save round-trip
  # duplicates the <v:shapetype>/<v:shape> content of any vmlDrawing*.vml this
  # workbook already carries (Cover Sheet/QA-QC's real cell comments here),
  # leaving multiple elements sharing the same "id" attribute. Duplicate VML
  # ids are invalid and a known trigger for Excel's "problem with some
  # content" repair prompt. Keeping only the first occurrence of each id
  # restores the file to what the source template itself actually defines.
  vml_files <- list.files(file.path(work, "xl/drawings"), pattern = "^vmlDrawing[0-9]+[.]vml$", full.names = TRUE)
  for (vf in vml_files) {
    vdoc <- xml2::read_xml(vf)
    vns  <- xml2::xml_ns(vdoc)
    for (xpath in c(".//v:shapetype", ".//v:shape")) {
      nodes <- xml2::xml_find_all(vdoc, xpath, vns)
      seen <- character(0)
      for (nd in nodes) {
        id <- xml2::xml_attr(nd, "id")
        if (is.na(id)) next
        if (id %in% seen) xml2::xml_remove(nd) else seen <- c(seen, id)
      }
    }
    
    # (3b) the source template itself (independent of the openxlsx round-trip
    # above) separately carries a second, older leftover shape per comment -
    # e.g. an old flat-colour "note" shape alongside a newer theme-coloured
    # one - both anchored to the same cell, likely left behind when Excel
    # migrated its comment-rendering style without cleaning up the prior
    # shape. Two shapes anchored to one real comment is itself invalid and,
    # per testing, enough on its own to trigger the repair prompt even after
    # (3)'s id-dedup. Group remaining shapes by anchor cell (Row/Column) and
    # keep only the last one in document order - the more recently-authored,
    # modern-styled shape - per cell.
    shapes <- xml2::xml_find_all(vdoc, ".//v:shape", vns)
    if (length(shapes) > 0) {
      keys <- vapply(shapes, function(s) {
        row <- xml2::xml_text(xml2::xml_find_first(s, ".//x:Row", vns))
        col <- xml2::xml_text(xml2::xml_find_first(s, ".//x:Column", vns))
        paste(row, col, sep = ",")
      }, character(1))
      for (k in unique(keys[duplicated(keys)])) {
        idx <- which(keys == k)
        for (di in idx[-length(idx)]) xml2::xml_remove(shapes[[di]])
      }
    }
    
    xml2::write_xml(vdoc, vf, options = "no_declaration")
  }
  
  # re-zip in place, [Content_Types].xml first by convention
  old_wd <- getwd()
  setwd(work)
  on.exit(setwd(old_wd), add = TRUE)
  everything <- list.files(".", recursive = TRUE, all.files = TRUE, full.names = FALSE)
  everything <- everything[!grepl("^[.][.]?$", basename(everything))]
  ct <- "[Content_Types].xml"
  rest <- setdiff(everything, ct)
  zip::zip(path, files = c(ct, rest), recurse = FALSE)
  
  invisible(path)
}



# superscript footnote markers - openxlsx has no rich-text/mixed-formatting
# API (createStyle() only styles a whole cell), so footnote markers are
# superscripted via their Unicode code point instead (matches
# 2_Data_Analyses.R / 1b_Anova 1-way table.R's own convention).
superscript_letters <- c(a="ᵃ", b="ᵇ", c="ᶜ")
sup <- function(letter) unname(superscript_letters[letter])

# real per-note wrapped-line count - measures each note's actual wrapped
# line count for the workbook's font/size and merged width in pixels
# (matches 2_Data_Analyses.R / 1b_Anova 1-way table.R's own inline helpers
# of the same name).
excel_width_to_px <- function(width_units, mdw = 7) {
  floor(((256 * width_units + floor(128 / mdw)) / 256) * mdw)
}
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



# clean up the caption table breaks ---------------------------------------


# caption_clean <- gsub("\\s*\n\\s*", " ", caption)   # replace line breaks (and surrounding spaces) with a single space
# caption_clean <- trimws(gsub("\\s+", " ", caption_clean))  # collapse any double spaces, trim ends


# cell shading styles, reused by every shading block below instead of being
# re-created each time - blue = positive/increase, green = negative/decrease
blue_style  <- createStyle(fgFill = "#9BC2E6")
green_style <- createStyle(fgFill = "#C4D79B")

# superscript footnote markers ------------------------------------------------
# openxlsx has no rich-text/mixed-formatting API (createStyle() only styles a
# whole cell), so footnote-marker letters are superscripted using their
# dedicated Unicode code points instead - these render as true superscript
# glyphs in plain text, no special Excel handling required.
superscript_letters <- c(a="ᵃ", b="ᵇ", c="ᶜ", d="ᵈ", e="ᵉ",
                         f="ᶠ", g="ᵍ", h="ʰ", i="ⁱ", j="ʲ",
                         k="ᵏ", l="ˡ", m="ᵐ", n="ⁿ", o="ᵒ",
                         p="ᵖ", r="ʳ", s="ˢ", t="ᵗ", u="ᵘ",
                         v="ᵛ", w="ʷ", x="ˣ", y="ʸ", z="ᶻ")
sup <- function(letter) {
  out <- superscript_letters[letter]
  if (anyNA(out)) stop("No superscript mapping for letter(s): ", paste(letter[is.na(out)], collapse=", "))
  unname(out)
}
# footnotes.xlsx marks its footnote text with a leading "<letter> " (e.g. "a
# Years that share..."); this superscripts just that leading marker so it
# matches the markers embedded in the header/caption text, leaving footnote
# entries with no marker (e.g. "Area x Year...") untouched.
superscript_leading_marker <- function(txt) {
  # per-element loop, not vectorized ifelse() - ifelse() evaluates both branches
  # for every element, which would call sup() on non-marker leading characters
  # (e.g. the "A" in "Area x Year...") and error on an unmapped letter
  vapply(txt, function(t) {
    if (is.na(t) || !grepl("^[a-z] ", t)) return(t)
    paste0(sup(substr(t, 1, 1)), substring(t, 2))
  }, character(1), USE.NAMES = FALSE)
}


# rich-text runs within a note cell ----------------------------------------
# openxlsx (used for everything else in this script) has no rich-text/mixed-
# formatting API - createStyle() only styles a whole cell - so any word that
# needed special formatting inside a note (e.g. "Bold font", or the
# "observed"/"predicted" subscripts in footnote b's MOD formula) used to be
# faked with special Unicode code points instead. That doesn't really work:
#  - Unicode's "Mathematical Bold" glyphs (used for "Bold font") aren't in
#    Calibri, so Excel's worksheet grid substitutes a fallback font whose
#    baseline sits visibly lower than Calibri's - the term visibly "drops"
#    out of line even though it's correct in the formula bar (which doesn't
#    use that same fallback).
#  - Unicode subscript letters don't even exist for every letter needed here
#    ("observed"/"predicted" need b, r, v, d, i, c - none of which have a
#    subscript code point), so there's no way to spell them that way at all.
# Splitting a note into separate cells for each styled word side-steps both
# problems, but changes the note from one cell into several. Instead: keep
# writing every note as one plain-text cell as usual, and once the workbook
# is saved, patch specific cells' content into several *rich-text runs within
# the same cell* - openxlsx2 (unlike openxlsx) supports real per-run
# formatting (bold, subscript, etc.) inside a single cell, matching however
# many matches a regex finds and leaving the rest of the note plain.
# record_note_runs() queues up which (sheet, row, text, regex, style) to
# patch; the actual patch happens once, after saveWorkbook(), via
# apply_note_runs() below.
#note_run_specs <- list()
record_note_runs <- function(sheet, row, text, regex, style, perl = TRUE) {
  note_run_specs[[length(note_run_specs) + 1]] <<- list(sheet = sheet, row = row, text = text, regex = regex, style = style, perl = perl)
}

# builds one cell's worth of alternating plain/styled fmt_txt() runs from
# every match of `regex` in `text` - size/font are given explicitly on every
# run (not just the styled one): a run with no rPr (or one that only sets its
# own style) renders in Excel's default font/size (Calibri 11), not the
# cell's own note_style (Calibri 9); rich-text runs don't inherit from the
# cell's style the way plain cell text does.
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

# real per-note wrapped-line count ------------------------------------------
# a fixed "chars per line" ratio is only ever a rough guess at how much text
# fits per line - it either clips a long note (too generous a ratio) or, as
# was happening here, leaves blank space at the bottom of the row (too
# stingy a ratio, overcounting the lines actually needed). This instead
# measures each note's real wrapped line count using the workbook's actual
# font/size and the actual merged width in pixels (via Excel's own column-
# width-to-pixel formula), greedily word-wrapping the same way Excel does.
excel_width_to_px <- function(width_units, mdw = 7) {
  floor(((256 * width_units + floor(128 / mdw)) / 256) * mdw)
}

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