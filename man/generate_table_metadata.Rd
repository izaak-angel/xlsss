% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/build_metadata.R
\name{generate_table_metadata}
\alias{generate_table_metadata}
\title{Generate Metadata for a Single Sheet}
\usage{
generate_table_metadata(
  table_names,
  table_data,
  padding_rows_multi = 2,
  padding_rows_single = 1,
  table_headings = FALSE
)
}
\arguments{
\item{table_names}{Table name}

\item{table_data}{Tibble of tables}

\item{padding_rows_multi}{Row gap for sheet with multiple tables}

\item{padding_rows_single}{Row gap for sheet with single table}

\item{table_headings}{Use subheadings for all tables}
}
\description{
Generate Metadata for a Single Sheet
}
