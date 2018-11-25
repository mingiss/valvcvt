TODO list
=========

1.  * > done: Make library `lib/xlstree.py` as a container of all `xlsx.xml` related tools.
    * > impossible, etree isn't a class: Make `xlstree` as a class, inherited from `lxml.etree`.

1. > done: Move file import, export and deletion of emty cells from `valvcvt.py` to methods of `xlstree`.

1. Make `xlsx2htm` &ndash; converter from `xlsx.xml` to `html`.

1.  * > done: Fork empty cells deletion to separate tool `xslxtrim`.
    * > done: Move test script `valvcvt.sh` contents to `test_xslxtrim.sh` as well.
    * > done: Delete destination files at very begin of the test.
    * > done: Rename test script `valvcvt.sh` to `test_valvcvt.sh`.

1. `xlstree.concat_sheets()` doesn't delete appended sheets

1. Move top level `.py` files to the subfolder `src`, folder `lib` &ndash; to `src/lib`

1. Rename `xlstree.write()` to ``xlstree.save()`.
