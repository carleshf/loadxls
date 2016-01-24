#' Write objects to a MS Excel File
#'
#' Given a collection of R objects (aka \code{data.frame}s), it saves them as
#' different sheets in a MS Excel File.
#'
#' @param ... \code{data.frame}s to be saved as sheets in a MS Excel File.
#' @param filename The path and the name to the MS Excel File.
#' @param verbose If set to \code{TRUE} usefull messages are shown.
#' @export write_all
write_all <- function(..., filename, verbose=TRUE) {
    if(verbose & file.exists(filename)) {
        warning("Given file '", filename, "' already exists.")
    }
    wb <- XLConnect::loadWorkbook(filename, create=TRUE)
    
    elms <- as.character(substitute(list(...)))[-1L]
    trash <- lapply(elms, function(sheetname) {
        if(verbose) {
            message("Saveing object '", sheetname, "'.")
        }
        XLConnect::createSheet(wb, name=sheetname)
        XLConnect::writeWorksheet(wb, data=get(sheetname), 
                                  sheet=sheetname, header=TRUE)
    })
    XLConnect::saveWorkbook(wb)
    rm(trash)
}