#' Draw legend swatch text boxes into a saved workbook
#'
#' @description Adds one floating, solid-filled Excel text-box shape per row
#'   recorded by [record_legend_shapes()], centred in column 1 of that row.
#'   openxlsx has no shape API, so this edits the saved .xlsx package directly:
#'   unzip, write a `drawingN.xml` part per sheet, wire it up via
#'   `[Content_Types].xml`, the sheet's `_rels` and a `<drawing r:id=".."/>`
#'   tag on the sheet XML, then re-zip.
#'
#' @details Run after `openxlsx::saveWorkbook()` and before
#'   [repair_workbook_package()], so the new drawing parts are in place when
#'   the repair step's orphan-relationship check runs. Swatch centring assumes
#'   column 1 is 12 characters wide and note rows use the default 15pt height.
#'
#' @param path A character string of the saved .xlsx file path
#' @param specs A list of specs built by [record_legend_shapes()], usually
#'   `legend_shape_specs`
#' @param width_in A number of the swatch width in inches
#' @param height_in A number of the swatch height in inches
#'
#' @return `path`, invisibly
#' @export
#' @importFrom stats setNames
#'
#' @examples
#' \dontrun{
#' openxlsx::saveWorkbook(wb, "tables.xlsx", overwrite = TRUE)
#' add_legend_swatch_shapes("tables.xlsx", legend_shape_specs)
#' repair_workbook_package("tables.xlsx")
#' }
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
