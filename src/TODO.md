TODO list
=========

1.  * > done: Make library `lib/xlstree.py` as a container of all `xlsx.xml` related tools.
    * > impossible, etree isn't a class: Make `xlstree` as a class, inherited from `lxml.etree`.

1. > done: Move file import, export and deletion of emty cells from `valvcvt.py` to methods of `xlstree`.

1.  * Make `xlsx2htm` &ndash; converter from `xlsx.xml` to `html`.
    * Make `xlstree.export_html()`.
    (Not relevant to `valvcvt`.)

1.  * > done: Fork empty cells deletion to separate tool `xslxtrim`.
    * > done: Move test script `valvcvt.sh` contents to `test_xslxtrim.sh` as well.
    * > done: Delete destination files at very begin of the test.
    * > done: Rename test script `valvcvt.sh` to `test_valvcvt.sh`.

1. > done: `xlstree.append_xlsx_sheet()` doesn't delete appended sheets

1. > done: Move top level `.py` files to the subfolder `src`, folder `lib` &ndash; to `src/lib`

1. Rename `xlstree.write()` to ``xlstree.save()`.

1.  * Make `xlstree.export_csv()`
    * Implement it into `valvcvt.py` &ndash; add third parameter &ndash; output format &ndash; `csv` or `xlsx`.

1. README.md

1. pydoc

1. Make tests for:
    * `xlstree.append_xlsx_sheet()`
    * `xlstree.append_xlsx()`
    * forthcomming `xlstree.export_csv()`
    * rename output of `test_xlstrim.sh`
    * > done: adapt `test_valvcvt.sh` &ndash; redirect to different output

1. Extend `xlstree.append_xlsx()` to concatenate all worksheets of the files.
At the moment only the very first worksheet tables are concatenated.
(Not relevant to the `valvcvt` project, but anyway.)

1. Extend both
    * `xlstree.append_xlsx()` and
    * `xlstree.append_xlsx_sheet()` to move cells together with their styles.
    * Investigate numbering order of the styles.

1. Check correspondence between `<Worksheet>` and `<Table>` tags.

1. > done: this is the `Microsoft Excel 2003 XML` format: Compare `xlsx` and `xslx.xml` formats.

1. Fix `prod/trim.sh` to skip `README.txt`.

1. Clone `lib` from `kppylib`.
