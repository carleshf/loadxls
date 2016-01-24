#' Load a full MS Excel File
#'
#' Given a MS Excel File it loads the full workbook creating as many variables as
#' sheets exists in the workbook.
#'
#' @param filename The path and the name to the MS Excel File.
#' @param environment Environment were the new variables will be located. By
#' default is parent environment.
#' @param verbose If set to \code{TRUE} usefull messages are shown.
#' @export read_all
read_all <- function(filename, environment=parent.frame(), verbose=TRUE) {
    if(!file.exists(filename)) {
        stop("Given file '", filename, "' does not exists.")
    }
    wb <- XLConnect::loadWorkbook(filename, create=FALSE)
    trash <- lapply(XLConnect::getSheets(wb), function(sheet) {
        if(verbose) {
            message("Loading sheet '", sheet, "'.")
        }
        var <- XLConnect::readWorksheet(wb, sheet)
        assign(sheet, var, envir=environment)
    })
    rm(trash)
}