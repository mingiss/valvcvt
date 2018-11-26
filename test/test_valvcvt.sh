#! /bin/bash

valvcvt_path=..
export PYTHONPATH=$valvcvt_path/lib:$PYTHONPATH

# head -c 1000 "2-way ball valves flangeable with SAE connections.xlsx.xml" > zzz.txt

# tidy "2-way ball valves flangeable with SAE connections.xlsx.xml" 2-way_ball.orig.tidy.xml
# xmllint --format "2-way ball valves flangeable with SAE connections.xlsx.xml" > 2-way_ball.orig.fmt.xml
# xmllint --format 2-way_ball.orig.xml > 2-way_ball.orig.fmt.xml

# xmllint --format "Cartridge ball valves.xlsx.xml" > cartridge.orig.fmt.xml
# xmllint --format cartridge.orig.xml > cartridge.orig.fmt.xml

# python3 /home/mingis/F/kp/src/xml/valvcvt/src/valvcvt/valvcvt.py files_orig.txt result.xlsx.xml
python3 $valvcvt_path/valvcvt.py files_orig.txt result.xlsx.xml
xmllint --format result.xlsx.xml > result.xlsx.fmt.xml
