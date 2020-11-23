# Reading tables from spreadsheets
This code creates a dataset of tables in form of a json line file, 
from a directory containing excel sheets in `xls` or `xlsx` format.
This code uses Apache POM library.

## Compilation
use `sbt` to compile the code. In the root directory, where `build.sbt` is, run `sbt compile`.

## Usage
Currently the path to the input directory and the path to the output file is hard-coded.
please edit the scala file `src/main/scala/test/Main.scala`, at line 606.
`createDataset` functions takes two inputs, first is the path to the directory containing the excel files,
and the second is the path to the output file (gzipped json line file).

```
createDataset("/media/majid/data/data_new/fbi_tables_all/",
    "/media/majid/data/Download/fbi_tables_all.jl.gz")
```