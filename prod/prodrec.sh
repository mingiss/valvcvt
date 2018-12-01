#! /bin/bash

# script to extract data from special format xlsx.xml files
# source files should be placed to subfolder orig, output is written to the file resrec.csv

ls orig | sed "s/^/orig\//g" > files_orig.txt
ls trimmed | sed "s/^/trimmed\//g" > files_trimmed.txt

valvcvt_path=../src
export PYTHONPATH=$valvcvt_path/lib:$PYTHONPATH

python3 $valvcvt_path/valvrec.py files_orig.txt resrec.csv
