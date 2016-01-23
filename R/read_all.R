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
    wb <- XLConnect::loadWorkbook(filename, create = FALSE)
    if(verbose) {
        message("Sheets: ", paste(XLConnect::getSheets(wb), collapse=" "))
    }

    sh <- XLConnect::getSheets(wb)
    trash <- lapply(sh, function(sheet) {
        if(verbose) {
            message("Loading sheet '", sheet, "'.")
        }
        var <- XLConnect::readWorksheet(wb, sheet)
        assign(sheet, var, envir=environment)
    })
    rm(trash)
}
