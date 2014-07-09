SummerSchoolExcelParser
=======================

parser and collator for the daily evaluation form for SISW summer school

Usage
-----

In the big text field add newline separated paths to xlsx files in the format found in the Admission folder.

The files can have one TEMPLATE sheet which is ignored. All other sheets are taken as being in the correct format.

All data will be squashed in the final output.

Press save and select a path to a file (e.g. `c:\output.xlsx`) to get the output.

Building
--------

You need VS2012 and OFFICE14 with the .net interop assembly pack thing installed.

TODO
----

* refactor code because it's ugly
