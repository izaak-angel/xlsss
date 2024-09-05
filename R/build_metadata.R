#' Create metadata
#'
#' @export
create_metadata <- function(){

  dplyr::tibble(
    sheet_name = c(),
    sheet_title = c(),
    table_names = list(),
    table_notes = list()
    )
}

#' Instructs metadata to add tables to a given worksheet
#'
#' @param metadata Metadata object
#' @param sheet_name Sheet name
#' @param sheet_title Sheet title
#' @param table_names List of table names
#' @param table_notes List of table notes
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

#' Add table metadata for all tables
#'
#' @param metadata Metadata object
#' @param table_list List of tables
#' @export
add_tables_to_metadata <- function(metadata, table_list) {
  metadata %>%
    dplyr::mutate(
      tables = purrr::map(.data$table_names,
                          ~ generate_table_metadata(.x, table_list))) %>%
    dplyr::mutate(
      n_tables = purrr::map(.data$tables, nrow) %>% as.numeric(),
      notes_start = purrr::map(
        .data$tables,
        ~ .x %>%
          dplyr::select(.data$n_rows) %>%
          purrr::map( ~ sum(.x)) %>%
          as.numeric()
      )
    )

}

#' Generate metadata for a single sheet
#'
#' @param table_names Table name
#' @param table_list List of tables
#' @param padding_rows_multi Row gap for sheet with multiple tables
#' @param padding_rows_single Row gap for sheet with single table
#' @export
generate_table_metadata <- function(table_names,
                                    table_list,
                                    padding_rows_multi = 2,
                                    padding_rows_single = 1){

  dplyr::tibble(
    table_name = table_names,
    table = table_list %>%
      dplyr::filter(.data$name %in% table_names) %>%
      dplyr::pull(.data$table),
    table_title = table_list %>%
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
        nrow(.) > 1 ~ .data$end_row - .data$n_rows + padding_rows_multi,
        nrow(.) == 1 ~ .data$end_row - .data$n_rows + padding_rows_single
      ),
      start_col = 1,
      end_col = .data$n_cols
    )

}
