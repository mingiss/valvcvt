<style>
ul {list-style-type: none;}
</style>

valvrec &ndash; tool for collecting special format xlsx data from several files to single formalized output file
================================================================================================================

# File formats

The tool collects loose formatted data from a set of xlsx data files, exported to `Microsoft Excel 2003 XML` format.
Output &ndash; comma separated `CSV` table, containing simply concatenated all data from the input sheets with data headings spread to beginning of each row.

Input data is presented as collection of xlsx workbook, exported to xml files.
Data chunks are spread over several sheets in files, not bound to fixed coordinates.
Multiple portions could be placed one over another vertically, but could be formed as bigger columns spread horizontally to each other.
Data are recognized as cell groups in format of 3 columns and several rows, having the caption `Type`, `PN [bar]`, `filename [.stp]` over the group.
There are headings of several levels, relevant to some single segment of the data or to the group of segments.
Headings are recognized as cells with the text, coll-spanned over thye data cells horizontally and placed gradually leftwards and upwards according to their level.
There is not set, which categories of headings will be presented and which omitted.
To which category the heading belongs, could be recognized using a dictionary of pattern words, which could be present in one or another category.
To which rows of data headings are relevant, is stated through an empty row-spanned cell beneath the heading cell.
Data file examples could be seen in the folder test:
- `test/cartridge.orig.xml`

Output file should contain all data collected in rows with 3 data cells on the far right end and all headings collected in the left to the data.
Headings should be provided in the order:
- `Kategorie`, `Familie`, `Material`, `Bauform`, `Serie/Verbindungstyp`, `Metrisch/UNC`

Headings of the category `Kategorie` are provided as the input file names, `Familie` -- as sheet names,
`Material` could be provided either in the sheet with the data, or as suffices of sheet names as well.
Rest headings are provided next to the data in the sheets.

# Script files provided

The main python scripts are:
- `src/valvrec.py` &ndash; the final tool for collecting the data
- `src/valvcvt.py` &ndash; the prototype example without recognition of data chunk position and heading category
- `src/xlstrim.py` &ndash; example for cleaning empty cells at the end of input files
- `src/lib/xlstree.py` &ndash; common methods for processing of `Microsoft Excel 2003 XML` files
- `prod/prodrec.sh` &ndash; production stage launch bash script for `valvrec.py`
- `prod/prod.sh` &ndash; production script for `valvcvt.py`
- `prod/trim.sh` &ndash; script for starting the script `xlstrim.py`
- `test/test_valvrec.sh` &ndash; test stage launch bash script for `valvrec.py`
- `test/test_valvcvt.sh` &ndash; test script for `valvcvt.py`
- `test/test_xlstrim.sh` &ndash; test for `xlstrim.py`

# Script details

## `valvrec`

`src/valvrec.py` has two parameters:
- name of a text file with the list of all input `XML` file names
- name of the output `CSV` file

Script requires `Python 3` and package `lxml` to be installed.

`prod/prodrec.sh` starts execution of the data collecting process using `src/valvrec.py`.
Input files should be copied to the folder `prod\orig`, output file produced &ndash; `prod/resrec.csv`.
Input file list is generated to the file `prod/files_orig.txt`.

`test/test_valvrec.sh` launches script `src/valvrec.py` and applies it to the example files `test/*.xml`.
Result file `test/resrec.csv` could be checked for changes using the `git status` or `git diff` commands.


## `valvcvt`

`src/valvcvt.py` has three parameters:
- name of the input file list
- name of the output file
- output format -- either `xml` or `csv`

`prod/prod.sh` collects data from `XML` files in `prod\orig` and writes collected data to the files
`prod/result.csv`, `prod/result.xlsx.xml` and `prod/result.xlsx.fmt.xml`.
`prod/result.xlsx.fmt.xml` is a human readable variant, produced using the tool `xmllint`.

`test/test_valvcvt.sh` launches test of the script `src/valvcvt.py` with an output to
`test/result.csv`, `test/result.xlsx.xml` and `test/result.xlsx.fmt.xml`.


## `xlstrim.py`

`src/xlstrim.py` has two parameters:
- name of the input xml file
- name of the output xml file

Script could be used for compacting of `Microsoft Excel 2003 XML` files.
It deletes empty cells on ends of the rows, as well as empty rows at the end of the sheets.

`prod/trim.sh` strips empty cells of files in the folder `prod/orig` and writes cleaned versions of them to `prod/trimmed`.

`test/test_xlstrim.sh` cleans files `test/*.orig.xml` and writes the results to files `test/*.xml`.
