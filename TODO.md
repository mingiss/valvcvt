TODO list
=========

1.  * Make library `lib/xlstree.py` as a container of all `xlsx.xml` related tools.
    * Make `xlstree` as a class, inherited from `lxml.etree`.

1. Move file import, export and deletion of emty cells from `valvcvt.py` to methods of `xlstree`.

1. Make `xlsx2htm` &ndash; converter from `xlsx.xml` to `html`.

1.  * Fork empty cells deletion to separate tool `xslxtrim`.
    * Move test script `valvcvt.sh` contents to `test_xslxtrim.sh` as well.
    * Delete destination files at very begin of the test.
    * Rename test script `valvcvt.sh` to `test_valvcvt.sh`.
