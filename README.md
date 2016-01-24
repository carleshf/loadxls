# loadxls

`loadxls` is an R package to read and write from an to Microsoft Excel Files. It is a wrapper of the excellent but complicated R package `XLConnect`. `loadxls` makes the operations of load the content of a spreadsheet super-easy, also save `data.frame`s to them.

## Installation

Open an R session and load `devtools`. Then type:

```R
install_github("carleshf/loadxls")
```

## Documentation

`loadxls` implements four functions:

 1. `read_all`: This functions load the full content of a MS Excel file into the current workspace.
 2. `read_sheet`: This functions load a specific sheet of a MS Excel file into the current workspace.
 3. `write_all`: This function writes the given objects to a MS Excel file, creating a sheet for each object.
 4. `write_sheet`: This function writes a given object to a specified sheet of a MS Excel file.
