#! /bin/bash

valvcvt_path=../src
export PYTHONPATH=$valvcvt_path/lib:$PYTHONPATH

python3 $valvcvt_path/valvrec.py files_orig.txt recres.csv
