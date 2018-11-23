#! /bin/bash

export PYTHONPATH=../lib:$PYTHONPATH

# head -c 1000 "2-way ball valves flangeable with SAE connections.xlsx.xml" > zzz.txt

# tidy "2-way ball valves flangeable with SAE connections.xlsx.xml" 2-way_ball.orig.tidy.xml
# xmllint --format "2-way ball valves flangeable with SAE connections.xlsx.xml" > 2-way_ball.orig.fmt.xml

# python3.4 /home/mingis/F/kp/src/xml/valvcvt/src/valvcvt/valvcvt.py "2-way ball valves flangeable with SAE connections.xlsx.xml" 2-way_ball.xml
python3.4 ../valvcvt.py 2-way_ball.orig.xml 2-way_ball.xml
xmllint --format 2-way_ball.xml > 2-way_ball.fmt.xml

# xmllint --format "Cartridge ball valves.xlsx.xml" > cartridge.orig.fmt.xml

# python3.4 /home/mingis/F/kp/src/xml/valvcvt/src/valvcvt/valvcvt.py "Cartridge ball valves.xlsx.xml" cartridge.xml
python3.4 ../valvcvt.py cartridge.orig.xml cartridge.xml
xmllint --format cartridge.xml > cartridge.fmt.xml

