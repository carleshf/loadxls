library(loadxls)

context("Testing read_all")

filename <- system.file("extdata", "hnp2015_es.xlsx", package = "loadxls")

test_that("read_all for all vars", {
    read_all(filename, verbose=FALSE)
    y <- ls()

    expect_equal(sum(
      c("Country", "Country-Series", "Data", "Series", "Series-Time", "Sheet7") 
      %in% y), 6)
    
    rm(y, Country, `Country-Series`, Data, Series, `Series-Time`, Sheet7)
})

test_that("read_sheet - non name given", {
    read_sheet(filename, "Country")
    y <- ls()
    expect_true("Country" %in% y)
    rm(Country)
})

test_that("read_sheet - name given", {
    read_sheet(filename, "Country", "test")
    y <- ls()
    expect_true("test" %in% y)
    rm(test)
})