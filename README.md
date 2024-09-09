
<!-- README.md is generated from README.Rmd. Please edit that file -->

# xlsss

<!-- badges: start -->

[![R-CMD-check](https://github.com/izaak-jephson/xlsss/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/izaak-jephson/xlsss/actions/workflows/R-CMD-check.yaml)
<!-- badges: end -->

The goal of xlsss is to automate production of accessible statistical
tables for Social Security Scotland.

## Installation

You can install the development version of xlsss from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("ScotGovAnalysis/xlsss",
                          upgrade = "never")
```

## Getting started

The package comes with a template R script to incorporate into your
analysis pipeline. After installation, run:

``` r
xlsss::create_template_output("template_filepath.R")
```

to create the template R script at the specified filepath. This should
then be edited to the specifics of your analysis and desired table
layout. Instructions are contained within the script itself.

## Contributing to the package

Contributions to the package are very welcome. If you would like to
contribute, please fork the directory and open a pull request.

Details of how to do so are available
[here](https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/proposing-changes-to-your-work-with-pull-requests/creating-a-pull-request).
