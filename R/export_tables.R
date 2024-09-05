#' Create tibble of sheet titles for contents page
#'
#' @param table_layout Table layout object created by metadata functions
#' @export
make_contents_table <- function(table_layout) {
  contents <- table_layout %>%
    dplyr::select(.data$sheet_name, .data$sheet_title) %>%
    dplyr::rename(Sheet = .data$sheet_name, Description = .data$sheet_title) %>%
    dplyr::add_row(Sheet = "Notes", Description = "List of notes") %>%
    dplyr::slice(order(.data$Sheet != "Notes"))

  contents
}

#' Add contents sheet to give given workbook based on tibble of sheet titles
#'
#' @param wb Workbook object
#' @param contents Contents
#' @param contents_title Character string of title to include on contents page.
#' @export
add_contents_sheet <- function(wb,
                               contents,
                               contents_title) {
  contents_table <-
    contents %>%
    dplyr::slice(-1) %>%
    dplyr::mutate(`Table Number` = openxlsx::makeHyperlinkString(
      sheet = .data$Sheet,
      row = 1,
      col = 1,
      text = .data$Sheet
    ))

  class(contents_table$`Table Number`) <-
    c(class(contents_table$`Table Number`), "formula")

  openxlsx::modifyBaseFont(wb,
                           fontSize = 12,
                           fontColour = "black",
                           fontName = "Roboto"
  )

  openxlsx::addWorksheet(wb, "Contents")

  openxlsx::setColWidths(
    wb,
    "Contents",
    cols = 1:2,
    width = c(27, "auto"),
    ignoreMergedCells = TRUE
  )

  # Make table headings bold
  openxlsx::addStyle(
    wb,
    "Contents",
    rows = 1,
    cols = 1,
    style = openxlsx::createStyle(
      fontSize = 15,
      wrapText = FALSE,
      textDecoration = "bold"
    )
  )

  openxlsx::showGridLines(wb,
                          "Contents",
                          showGridLines = FALSE
  )

  # Format table numbers as hyperlinks
  openxlsx::addStyle(
    wb,
    "Contents",
    rows = 4:(nrow(contents) + 2),
    cols = 1,
    style = openxlsx::createStyle(fontColour = "blue", textDecoration = "underline")
  )

  # Add title
  openxlsx::writeData(wb,
                      "Contents",
                      x = c(contents_title, "Table of Contents")
  )

  # Add contents
  openxlsx::writeDataTable(
    wb,
    "Contents",
    startRow = 3,
    x = contents_table %>% dplyr::select(.data$`Table Number`, .data$Description),
    tableStyle = "TableStyleLight1",
    withFilter = openxlsx::openxlsx_getOp("withFilter", FALSE)
  )

  return(wb)
}


#' Add notes sheet to given workbook
#'
#' @param wb Workbook object
#' @param contents Contents
#' @param notes_list List of notes to include in publication
#' @export
add_notes_sheet <- function(wb,
                            contents,
                            notes_list) {
  openxlsx::addWorksheet(wb, sheetName = "Notes")

  notes_list <- notes_list %>%
    dplyr::rename(`Note number` = .data$note_number,
                  `Note text` = .data$note_text)

  # Format notes table
  openxlsx::setColWidths(
    wb,
    "Notes",
    width = 12,
    cols = 1,
    ignoreMergedCells = TRUE
  )

  openxlsx::setColWidths(
    wb,
    "Notes",
    width = 100,
    cols = 2,
    ignoreMergedCells = TRUE
  )

  openxlsx::addStyle(
    wb,
    "Notes",
    rows = 1,
    cols = 1,
    style = openxlsx::createStyle(
      fontSize = 15,
      wrapText = FALSE,
      textDecoration = "bold"
    )
  )

  openxlsx::addStyle(
    wb,
    "Notes",
    rows = c(1:nrow(notes_list) + 4),
    cols = 2,
    style = openxlsx::createStyle(wrapText = TRUE)
  )

  # Add heading text
  openxlsx::writeData(wb, "Notes", x = "List of notes")
  openxlsx::writeData(wb, "Notes", startRow = 2, x = "This worksheet displays 1 table")
  openxlsx::writeData(wb, "Notes",
                      startRow = 3,
                      x = "The notes within this table are referred to in other worksheets of this workbook."
  )

  # Add notes table
  openxlsx::writeDataTable(
    wb,
    "Notes",
    startRow = 4,
    x = notes_list,
    tableStyle = "TableStyleLight1",
    withFilter = openxlsx::openxlsx_getOp("withFilter", FALSE)
  )

  return(wb)
}

#' Format Columns
#'
#' @param wb Workbook object
#' @param sheet_name Name of sheet
#' @param table Table to be formatted
#' @param column Column to be formatted
#' @param start_row Row to start formatting
#' @param end_row Row to end formatting
#' @export
format_columns <- function(wb,
                           sheet_name,
                           table,
                           column,
                           start_row,
                           end_row) {
  # Format numbers with commas
  if (purrr::is_integer(table[[column]])) {
    openxlsx::addStyle(wb, sheet_name,
                       rows = start_row:end_row,
                       cols = column,
                       style = openxlsx::createStyle(numFmt = "#,##0;-;0", halign = "right"),
                       gridExpand = TRUE
    )



  # Format % columns
  } else if (stringr::str_detect(colnames(table[column]), "[Pp]ercent|[Pp]ercentage|[Pp]roportion")) {
    openxlsx::addStyle(wb, sheet_name,
                       rows = start_row:end_row,
                       cols = column,
                       style = openxlsx::createStyle(numFmt = "0%;-;0%", halign = "right"),
                       gridExpand = TRUE
    )

  }
  # Format £ columns
  if (str_detect(colnames(table[column]), "[Vv]alue")) {
    addStyle(wb, sheet_name,
             rows = start_row:end_row,
             cols = column,
             style = createStyle(numFmt = "£#,##0;-;£0", halign = "right"),
             gridExpand = TRUE
    )

  }
}

#' Format Rows
#'
#' @param wb Workbook object
#' @param sheet_name Name of sheet
#' @param table Table to be formatted
#' @param table_row Row of table to be formatted
#' @param sheet_row Row on sheet to be formatted
#' @param start_col Start column of formatting
#' @param end_col End row of formatting
#' @export
format_rows <- function(wb,
                        sheet_name,
                        table,
                        table_row,
                        sheet_row,
                        start_col,
                        end_col) {
  # Make total rows bold
  if (table[[table_row, 1]] == "Total" |
      stringr::str_starts(table[[table_row, 1]],"Financial")) {
    openxlsx::addStyle(wb, sheet_name,
                       rows = sheet_row,
                       cols = start_col:end_col,
                       style = openxlsx::createStyle(textDecoration = "bold"),
                       gridExpand = TRUE,
                       stack = TRUE
    )
  }
  # Format % rows
  else if (str_starts(table[[table_row, 1]],"Percentage")) {
    addStyle(wb, sheet_name,
             rows = sheet_row,
             cols = 2:end_col,
             style = createStyle(numFmt = "0%;-;0%", halign = "right"),
             gridExpand = TRUE
    )

  }
}

#' Add data tables to existing worksheet
#'
#' @param wb Workbook object
#' @param sheet_name Sheet name
#' @param sheet_title Sheet title
#' @param tables Tables to be included on sheet
#' @param header_rows Number of header rows
#' @export
add_data_tables <- function(wb,
                            sheet_name,
                            sheet_title,
                            tables,
                            header_rows) {
  tables %>%
    dplyr::select(.data$table,
                  .data$start_row,
                  .data$end_row,
                  .data$table_title,
                  .data$table_name) %>%
    purrr::pmap(function(table, start_row, end_row, table_title, table_name) {
      xlsss::add_data_table(wb,
                            sheet_name,
                            sheet_title,
                            table,
                            start_row,
                            end_row,
                            header_rows,
                            table_name)
      if(nrow(tables) > 1){
        openxlsx::writeData(wb, sheet_name,
                            x = table_title,
                            startRow = start_row + header_rows - 1,
                            startCol = 1
        )
        openxlsx::addStyle(wb,
                           sheet_name,
                           rows = start_row + header_rows - 1,
                           cols = 1,
                           style = openxlsx::createStyle(
                             wrapText = FALSE,
                             textDecoration = "bold"
                           )
        )
      }
    })
}

#' Create table sheet with header wording, tables and notes
#'
#' @param wb Workbook object
#' @param sheet_name Sheet name
#' @param sheet_title Sheet title
#' @param sheet_tables Sheet tables
#' @param notes_list List of notes
#' @param note_mapping Notes to included against each table
#' @param n_tables Number of tables
#' @param notes_start Row that notes start on
#' @export
add_data_sheet <- function(wb,
                           sheet_name,
                           sheet_title,
                           sheet_tables,
                           notes_list,
                           note_mapping,
                           n_tables,
                           notes_start) {
  # Define start and end points for tables
  header_rows <- 5

  # end_row <- header_rows + nrow(table[[1]]) + 1

  # Add data table sheet
  openxlsx::addWorksheet(wb, sheetName = sheet_name)

  # Format header section
  openxlsx::showGridLines(wb, sheet_name, showGridLines = FALSE)

  openxlsx::addStyle(wb,
                     sheet_name,
                     rows = 1,
                     cols = 1,
                     style = openxlsx::createStyle(
                       fontSize = 15,
                       wrapText = FALSE,
                       textDecoration = "bold"
                     )
  )

  # Add title and header text
  openxlsx::writeData(wb, sheet_name,
                      x = paste(
                        sheet_title,
                        paste0("[note ", note_mapping, "]",
                               collapse = " "
                        )
                      )
  )

  openxlsx::writeData(wb, sheet_name,
                      x = paste(
                        "This worksheet contains",
                        n_tables,
                        dplyr::if_else(n_tables == 1, "table.", "tables.")
                      ),
                      startRow = 2,
                      startCol = 1
  )

  openxlsx::writeData(wb, sheet_name,
                      x = paste0("Banded rows are used in ",dplyr::if_else(n_tables == 1, "this table", "these tables"),". To remove them, highlight the table, go to the Design tab and uncheck the banded rows box."),
                      startRow = 3,
                      startCol = 1
  )

  openxlsx::writeData(wb, sheet_name,
                      x = paste0(
                        "Notes are located below the",
                        dplyr::if_else(n_tables == 1, " table ", " tables "),
                        "beginning in cell A",
                        notes_start + header_rows,
                        " and in the notes sheet of this document."
                      ),
                      startRow = 4,
                      startCol = 1
  )




  openxlsx::writeData(wb, sheet_name,
                      x = "[c] indicates that a figure has been suppressed for disclosure control purposes.",
                      startRow = 5,
                      startCol = 1
  )



  # Add additional notes to header




  xlsss::add_data_tables(wb, sheet_name, sheet_title, sheet_tables, header_rows)

  # Add notes
  openxlsx::writeData(wb, sheet_name,
                      startRow = notes_start + header_rows,
                      x = dplyr::tibble(note = paste0("[note ", note_mapping, "]"),
                                 note_text = notes_list$note_text[note_mapping]),
                      colNames = FALSE
  )



  return(wb)
}

#' Add Data Table
#'
#' @param wb Workbook object
#' @param sheet_name Sheet name
#' @param sheet_title Sheet title
#' @param table Table to be added
#' @param start_row Sheet row on which table starts
#' @param end_row Sheet row on which table ends
#' @param header_rows Number of header rows
#' @param table_name Name of table
#' @export
add_data_table <- function(wb,
                           sheet_name,
                           sheet_title,
                           table,
                           start_row,
                           end_row,
                           header_rows,
                           table_name) {
  # Format data table
  openxlsx::addStyle(wb,
                     sheet_name,
                     rows = header_rows + start_row:header_rows + end_row - 1,
                     cols = 1,
                     style = openxlsx::createStyle(wrapText = FALSE)
  )


  # Add data table
  openxlsx::writeDataTable(wb, sheet_name,
                           startRow = header_rows + start_row,
                           tableName = stringr::str_replace_all(table_name, " ","_"),
                           x = table,
                           colNames = TRUE,
                           rowNames = FALSE,
                           keepNA = TRUE,
                           na.string = "n/a",
                           tableStyle = "TableStyleLight1",
                           headerStyle = openxlsx::createStyle(wrapText = TRUE),
                           stack = TRUE,
                           withFilter = openxlsx::openxlsx_getOp("withFilter", FALSE)
  )

  # Format table headers
  openxlsx::addStyle(wb, sheet_name,
                     rows = header_rows + start_row,
                     cols = 2:length(table),
                     style = openxlsx::createStyle(halign = "right", wrapText = TRUE)
  )

  # Format columns
  openxlsx::setColWidths(
    wb,
    sheet_name,
    cols = 1,
    width = 22,
    ignoreMergedCells = TRUE
  )

  openxlsx::setColWidths(
    wb,
    sheet_name,
    cols = 2:length(table),
    width = 12,
    ignoreMergedCells = TRUE
  )

  purrr::walk(
    1:length(table),
    ~ xlsss::format_columns(wb, sheet_name, table, .x, header_rows + 1 + start_row, header_rows + end_row - 1)
  )

  purrr::walk(
    1:nrow(table),
    ~ xlsss::format_rows(wb, sheet_name, table, .x, start_row + .x + header_rows, 1, length(table))
  )



  purrr::walk(1:length(table),
              function(y) purrr::walk(1:nrow(table),
                                      function(x) xlsss::negative_to_c(wb, sheet_name, table, y, x, start_row, end_row, header_rows)))


  return(wb)
}

#' Negative to "c"
#'
#' @param wb Workbook object
#' @param sheet_name Sheet name
#' @param table Table to be formatted
#' @param column Column to format
#' @param row Row to format
#' @param start_row Row to start formatting
#' @param end_row Row to end formatting
#' @param header_rows Number of header rows
#' @export
negative_to_c <- function (wb,
                           sheet_name,
                           table,
                           column,
                           row,
                           start_row,
                           end_row,
                           header_rows){
  if (table[[column]][[row]] == -1 & (!is.na(table[[column]][[row]]))){
    openxlsx::writeData(wb,
                        sheet_name,
                        x = "[c]",
                        startCol = column,
                        startRow = row + start_row + header_rows
    )
  }
}

#' Add Data Sheet
#'
#' @param wb Workbook object
#' @param table_layout Table layout generated by metadata function
#' @param notes_list List of notes
#' @export
add_data_sheets <- function(wb, table_layout, notes_list) {
  purrr::walk(
    1:nrow(table_layout),
    ~ xlsss::add_data_sheet(
      wb,
      table_layout[[1]][[.x]],
      table_layout[[2]][[.x]],
      table_layout[[5]][[.x]],
      notes_list,
      table_layout[[4]][[.x]],
      table_layout[[6]][[.x]],
      table_layout[[7]][[.x]]
    )
  )

  return(wb)
}

#' Create excel tables
#'
#' @param metadata metadata object created by metadata functions
#' @param table_data Tibble of tables. Must include columns: name, table and title.
#' `name` column must match the table names specified in the metadata object.
#' `table` column contains the tables to be outputted to excel
#' `title` column is only used where more than one table is included on a sheet
#' and is the subtitle to be printed above the table.
#' @param notes_list List of notes in publication
#' @param contents_title Title of contents page
#' @export
make_output_tables <- function(metadata,
                               table_data,
                               notes_list,
                               contents_title) {

  if(!tibble::is_tibble(table_data)){
    table_data <- table_list_to_tibble(table_data)}

  table_layout <- create_table_layout(metadata, table_data)

  wb <- openxlsx::createWorkbook()

  openxlsx::modifyBaseFont(wb, fontSize = 12, fontColour = "black", fontName = "Roboto")

  contents <- xlsss::make_contents_table(table_layout)

  wb <- xlsss::add_contents_sheet(wb, contents, contents_title)

  wb <- xlsss::add_notes_sheet(wb, contents, notes_list)

  wb <- xlsss::add_data_sheets(wb, table_layout, notes_list)

  wb
}

#' Create and save excel tables
#'
#' @param metadata metadata object created by metadata functions
#' @param table_data Tibble of tables. Must include columns: name, table and title.
#' `name` column must match the table names specified in the metadata object.
#' `table` column contains the tables to be outputted to excel
#' `title` column is only used where more than one table is included on a sheet
#' and is the subtitle to be printed above the table.
#' @param notes_list List of notes in publication
#' @param contents_title Title of contents page
#' @param workbook_filename Filename to export workbook to
#' @export
save_output_tables <- function(metadata,
                               table_data,
                               notes_list,
                               contents_title,
                               workbook_filename) {


  wb <- make_output_tables(metadata,
                           table_data,
                           notes_list,
                           contents_title)

  openxlsx::saveWorkbook(wb, workbook_filename, overwrite = TRUE)

  paste0("Tables have been saved to ",workbook_filename)
}
