#' Load a single sheet of MS Excel File
#'
#' Given a MS Excel File and the name of a sheet it loads the sheet contentent
#' as variables.
#'
#' @param filename The path and the name to the MS Excel File.
#' @param sheetname Name of the sheet to be loaded.
#' @param environment Environment were the new variables will be located. By
#' default is parent environment.
#' @param varname If not missing, is the name of the new variable filled with
#' the content of the specified sheet.
#' @param verbose If set to \code{TRUE} usefull messages are shown.
#' @export read_sheet
read_sheet <- function(filename, sheetname, varname, environment=parent.frame(), 
                     verbose=TRUE) {
    if(!file.exists(filename)) {
        stop("Given file '", filename, "' does not exists.")
    }
    wb <- XLConnect::loadWorkbook(filename, create=FALSE)
    if(!sheetname %in% XLConnect::getSheets(wb)) {
        stop("Given sheet '", sheetname, "' does not exist in given file.")
    }
    if(verbose) {
        message("Loading sheet '", sheetname, "'.")
    }
    var <- XLConnect::readWorksheet(wb, sheetname)
    if(missing(varname)) {
        assign(sheetname, var, envir=environment)
    } else {
        assign(varname, var, envir=environment)
    }
}