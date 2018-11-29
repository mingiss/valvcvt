#! /bin/bash

# script to concatenate special format xlsx.xml files
# source files should be placed to subfolder orig, output is written to the file result.xlsx.xml

ls orig | sed "s/^/orig\//g" > files_orig.txt
ls trimmed | sed "s/^/trimmed\//g" > files_trimmed.txt

valvcvt_path=../src
export PYTHONPATH=$valvcvt_path/lib:$PYTHONPATH

# python3 $valvcvt_path/valvcvt.py "orig/2-way ball valves flangeable with SAE connections.xlsx.xml" result.xlsx.xml
python3 $valvcvt_path/valvcvt.py files_orig.txt result.xlsx.xml xml
xmllint --format result.xlsx.xml > result.xlsx.fmt.xml

python3 $valvcvt_path/valvcvt.py files_orig.txt result.csv
