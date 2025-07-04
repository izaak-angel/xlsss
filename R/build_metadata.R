#' Initialise Metadata Object
#'
#' @description Creates an empty metadata object which can be added to with
#' `xlsss::add_sheet_to_metadata`
#' @export
create_metadata <- function(){

  dplyr::tibble(
    sheet_name = c(),
    sheet_title = c(),
    table_names = list(),
    table_notes = list()
    )
}

#' Add Sheet to Metadata Object
#'
#' @description Adds a new sheet to a metadata object.
#' @param metadata Metadata object
#' @param sheet_name Name to be displayed on sheet tab
#' @param sheet_title Title to be show in bold in cell A1 on sheet
#' @param table_names Character vector of table names to include on this sheet. Names must match those in the table_list object
#' @param table_notes Numeric vector of notes to be included underneath tables on this sheet
#' @export
add_sheet_to_metadata <- function(metadata,
                                  sheet_name,
                                  sheet_title,
                                  table_names,
                                  table_notes){

  metadata %>%
    dplyr::bind_rows(
      dplyr::tibble(
        sheet_name = sheet_name,
        sheet_title = sheet_title,
        table_names = list(table_names),
        table_notes = list(table_notes)
      )
    )
}

#' Combine Metadata and Tables to Create Layout
#'
#' Takes a metadata object and adds specified table_data to create a table layout to
#' be passed to `xlsss::make_output_tables`
#' @param metadata Metadata object
#' @param table_data Tibble of tables. Must include columns: name, table and title.
#' `name` column must match the table names specified in the metadata object.
#' `table` column contains the tables to be outputted to excel
#' `title` column is only used where more than one table is included on a sheet
#' and is the subtitle to be printed above the table.
#' @export
create_table_layout <- function(metadata, table_data, table_headings = FALSE) {
  metadata %>%
    dplyr::mutate(
      tables = purrr::map(
        .data$table_names,
        ~ generate_table_metadata(.x,
                                  table_data,
                                  table_headings = table_headings)
        )
      ) %>%
    dplyr::mutate(
      n_tables = purrr::map(.data$tables, nrow) %>%
        as.numeric(),
      notes_start =
        purrr::map(
          .data$tables,
          ~ .x %>%
            dplyr::select(.data$n_rows) %>%
            purrr::map(~ sum(.x)) %>%
            as.numeric()
      )
    ) %>%
    dplyr::select(.data$sheet_name,
                  .data$sheet_title,
                  .data$table_names,
                  .data$table_notes,
                  .data$tables,
                  .data$n_tables,
                  .data$notes_start)

}

#' Generate Metadata for a Single Sheet
#'
#' @param table_names Table name
#' @param table_data Tibble of tables
#' @param padding_rows_multi Row gap for sheet with multiple tables
#' @param padding_rows_single Row gap for sheet with single table
#' @export
generate_table_metadata <- function(table_names,
                                    table_data,
                                    padding_rows_multi = 2,
                                    padding_rows_single = 1,
                                    table_headings = FALSE){

  dplyr::tibble(
    table_name = table_names,
    table = table_data %>%
      dplyr::filter(.data$name %in% table_names) %>%
      dplyr::pull(.data$table),
    table_title = table_data %>%
      dplyr::filter(.data$name %in% table_names) %>%
      dplyr::pull(.data$title)
  ) %>%
    dplyr::mutate(
      n_rows = dplyr::case_when(
        nrow(.) > 1 ~ purrr::map(.data$table, ~ .x %>% nrow) %>%
          as.numeric() + 1 + padding_rows_multi,
        nrow(.) == 1 ~ purrr::map(.data$table, ~ .x %>% nrow) %>%
          as.numeric() + 1 + padding_rows_single
      ),
      n_cols = purrr::map(.data$table, length) %>% as.numeric(),
      end_row = cumsum(.data$n_rows),
      notes_start = .data$end_row,
      start_row = dplyr::case_when(
        nrow(.) > 1 | table_headings ~ .data$end_row - .data$n_rows + padding_rows_multi,
        nrow(.) == 1 ~ .data$end_row - .data$n_rows + padding_rows_single
      ),
      start_col = 1,
      end_col = .data$n_cols
    )

}


#' Convert List of Tables to table_data
#'
#' @description Helper function for quickly turning named list of tables into a
#' table_data object for use in the `xlsss::create_table_layout` function.
#' @param table_list List of tables to be converted to table_data
#' @export
table_list_to_tibble <- function(table_list){
  dplyr::tibble(
    name = names(table_list),
    table = table_list,
    title = paste(names(table_list)))
}


