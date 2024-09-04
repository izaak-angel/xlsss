#' Create metadata
#'
#'
#'

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

add_sheet_to_metadata <- function(metadata,
                                  sheet_name,
                                  sheet_title,
                                  table_names,
                                  table_notes,
                                  tables){

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
#'
#'

add_tables_to_metadata <- function(metadata, table_list) {
  metadata %>%
    dplyr::mutate(
      tables = purrr::map(table_names,
                          ~ generate_table_metadata(.x, table_list))) %>%
    dplyr::mutate(
      n_tables = purrr::map(tables, nrow) %>% as.numeric(),
      notes_start = purrr::map(
        tables,
        ~ .x %>%
          dplyr::select(n_rows) %>%
          purrr::map( ~ sum(.x)) %>%
          as.numeric()
      )
    )

}

#' Generate metadata for a single sheet
generate_table_metadata <- function(table_names,
                                    table_list,
                                    padding_rows_multi = 2,
                                    padding_rows_single = 1){

  dplyr::tibble(
    table_name = table_names,
    table = table_list %>%
      dplyr::filter(name %in% table_names) %>%
      dplyr::pull(table),
    table_title = table_list %>%
      dplyr::filter(name %in% table_names) %>%
      dplyr::pull(title)
  ) %>%
    dplyr::mutate(
      n_rows = dplyr::case_when(
        nrow(.) > 1 ~ purrr::map(table, ~ .x %>% nrow) %>%
          as.numeric() + 1 + padding_rows_multi,
        nrow(.) == 1 ~ purrr::map(table, ~ .x %>% nrow) %>%
          as.numeric() + 1 + padding_rows_single
      ),
      n_cols = purrr::map(table, length) %>% as.numeric(),
      end_row = cumsum(n_rows),
      notes_start = end_row,
      start_row = dplyr::case_when(
        nrow(.) > 1 ~ end_row - n_rows + padding_rows_multi,
        nrow(.) == 1 ~ end_row - n_rows + padding_rows_single
      ),
      start_col = 1,
      end_col = n_cols
    )

}


