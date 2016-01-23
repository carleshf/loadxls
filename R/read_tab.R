#' Load a single tab of MS Excel File
#'
#' Given a MS Excel File and the name of a tab it loads the tab contentent
#' as variables.
#'
#' @param filename The path and the name to the MS Excel File.
#' @param tabname Name of the tab to be loaded.
#' @param environment Environment were the new variables will be located. By
#' default is parent environment.
#' @param verbose If set to \code{TRUE} usefull messages are shown.
#' @export read_all
read_tab <- function(filename, tabname, environment=parent.frame(), verbose=TRUE) {
  
}