#' Create Template Script to Output Tables
#'
#' @param filename File to output template script to
#' @export
create_template_output <- function(filename = "template"){
  cat(
    "### Functions to amend and add ###
# Create metadata object defining each sheet and specifying which tables to include
build_metadata <- function(){
  create_metadata() %>%
  # eg add first sheet with one table
  add_sheet_to_metadata(
    sheet_name = \"Applications by month\",
    sheet_title = \"Application outcomes by month\",
    table_names = \"table_1\",
    table_notes = c(1,2,3)) %>%
  # eg add second sheet with one table
  add_sheet_to_metadata(
    sheet_name = \"Applications by age and LA\",
    sheet_title = \"Application outcomes by age and local authority\",
    table_names = c(\"table_2\",\"table_3\"),
    table_notes = c(1,2,3))
    # ...etc...
}

# Create list of notes
build_notes <- function(){
  tibble::tribble(
  ~note_number, ~note_text,
  \"[note 1]\", \"note text 1\",
  \"[note 2]\", \"note text 2\",
  \"[note 3]\", \"note text 3\"
  )
}

# Create tibble containing tables to output

# name column should correspond to table names in metadata,
# table column contains the table objects themselves
# title column contains title to print above table if more than one table on sheet
# otherwise this column is not used
build_tables <- function(){
  tibble::tribble(
  ~name, ~table, ~title
  \"table_1\", table_1, \"na\",
  \"table_2\", table_2, \"Application outcomes by age\",
  \"table_3\", table_3, \"Application outcomes by local authority\"
  )
}

# Apply bespoke formatting to final workbook if required
tweak_formatting <- function(wb, tweak) {
# Use openxlsx functions to apply whatever formatting adjustments are needed
# Below is simply an example
  setColWidths(wb,
              \"Applications by month\",
              cols = 1,
              width = 22,
              ignoreMergedCells = TRUE)
}


### Add to end of pipeline ###

metadata <- build_metadata()

notes_list <- build_notes()

table_data <- build_tables()


# Create workbook object
save_output_tables(
  metadata = metadata,
  table_data = table_data,
  notes_list = notes_list,
  contents_title = \"Statistical tables as at DATE\",
  use_low = FALSE, # set to true if you are using [low] for 1 and 2
  table_headings = FALSE) # set to true if you would like to include table headings for all tables, even if only one table per page

# Apply final formatting adjustments
wb <- tweak_formatting(wb)

# Output final workbook
openxlsx::saveWorkbook(wb, FILENAME, overwrite = TRUE)


"
, file = paste0(filename,".R"), append = FALSE)
}
