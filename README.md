SummerSchoolExcelParser
=======================

parser and collator for the daily evaluation form for SISW summer school

Usage
-----

In the big text field add newline separated paths to xlsx files in the format found in the Admission folder.

The files can have one TEMPLATE sheet which is ignored. All other sheets are taken as being in the correct format.

All data will be squashed in the final output.

Press save and select a path to a file (e.g. `c:\output.xlsx`) to get the output.

If you modify the structure of the excel file by adding or removing columns, please create a new file and then run the tool on both. This is because of a limitation that assumes that one book has the same format, and that the format can vary only between books. Basically the TEMPLATE sheet should have the same columns as the rest of the book.

Building
--------

You need VS2012 and OFFICE14 with the .net interop assembly pack thing installed.

TODO
----

* refactor code because it's ugly
