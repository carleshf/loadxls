#' Write single object to a MS Excel File
#'
#' Given a single \code{data.frame}, it saves it as sheet in a MS Excel File.
#'
#' @param data \code{data.frame} to be saved as sheet in a MS Excel File.
#' @param sheetname Name to be given to the sheet.
#' @param filename The path and the name to the MS Excel File.
#' @param replace If set to \code{TRUE} it replaces the content of the
#' given sheet with the current value of \code{data}.
#' @param verbose If set to \code{TRUE} usefull messages are shown.
#' @export write_sheet
write_sheet <- function(data, sheetname, filename, replace=FALSE,
                        verbose=TRUE) {
    wb <- XLConnect::loadWorkbook(filename, create=TRUE)
    sheets <- XLConnect::getSheets(wb)
    if(!replace & (sheetname %in% sheets)) {
        stop("Given sheet's name '", sheetname, "' already exists.")
    }
    XLConnect::createSheet(wb, name=sheetname)
    XLConnect::writeWorksheet(wb, data=data, sheet=sheetname, header=TRUE)
    XLConnect::saveWorkbook(wb)
}