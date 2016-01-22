library(loadxls)

contect("Testing read_all")
filename <- paste0(path.package("loadxls"),"/inst/extdata/hnp2015_es.xlsx")

test_that("read_all for all vars", {
    x <- ls()
    read_all(filename)
    y <- ls()

    expect_equal(str_length("a"), 1)
    expect_equal(str_length("ab"), 2)
    expect_equal(str_length("abc"), 3)
})
