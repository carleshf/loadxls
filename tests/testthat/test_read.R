library(loadxls)

contect("Testing read_all")
filename <- paste0(path.package("loadxls"),"/extdata/hnp2015_es.xlsx")

test_that("read_all for all vars", {
    x <- ls()
    read_all(filename, verbose=FALSE)
    y <- ls()

    expect_equal(sum(
      c("Country", "Country-Series", "Data", "Series", "Series-Time", "Sheet7") 
      %in% y), 6)
    
    rm(x, y, Country, `Country-Series`, Data, Series, `Series-Time`, Sheet7)
})
