#' Write objects to a MS Excel File
#'
#' Given a collection of R objects (aka \code{data.frame}s), it saves them as
#' different sheets in a MS Excel File.
#'
#' @param ... \code{data.frame}s to be saved as sheets in a MS Excel File.
#' @param filename The path and the name to the MS Excel File.
#' @param include.rownames If set to \code{TRUE} adds the \code{data.frame}'s
#' rownames as a new column in the sheet.
#' @param verbose If set to \code{TRUE} useful messages are shown.
#' @export write_all
write_all <- function(..., filename, include.rownames=FALSE, verbose=TRUE) {
    if(verbose & file.exists(filename)) {
        warning("Given file '", filename, "' already exists.")
    }
    wb <- XLConnect::loadWorkbook(filename, create=TRUE)
    
    elms <- as.character(substitute(list(...)))[-1L]
    trash <- lapply(elms, function(sheetname) {
        if(verbose) {
            if(include.rownames) {
                message("Saving object '", sheetname, "' with data.frame's row-names.")
            } else {
                message("Saving object '", sheetname, "'.")
            }
        }
        data <- get(sheetname)
        if(include.rownames) {
            cn <- colnames(data)   
            data$row.names.loadxls <- rownames(data)
            data <- data[ , c("row.names.loadxls", cn)]
            colnames(data) <- c("", cn)
        }
        XLConnect::createSheet(wb, name=sheetname)
        XLConnect::writeWorksheet(wb, data=data, 
                                  sheet=sheetname, header=TRUE)
    })
    XLConnect::saveWorkbook(wb)
    rm(trash)
}