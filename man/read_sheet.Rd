% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/read_sheet.R
\name{read_sheet}
\alias{read_sheet}
\title{Load a single sheet of MS Excel File}
\usage{
read_sheet(filename, sheetname, varname, environment = parent.frame(),
  verbose = TRUE)
}
\arguments{
\item{filename}{The path and the name to the MS Excel File.}

\item{sheetname}{Name of the sheet to be loaded.}

\item{varname}{If not missing, is the name of the new variable filled with
the content of the specified sheet.}

\item{environment}{Environment were the new variables will be located. By
default is parent environment.}

\item{verbose}{If set to \code{TRUE} usefull messages are shown.}
}
\description{
Given a MS Excel File and the name of a sheet it loads the sheet contentent
as variables.
}

