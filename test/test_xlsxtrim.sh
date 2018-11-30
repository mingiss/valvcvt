#! /bin/bash

valvcvt_path=../src
export PYTHONPATH=$valvcvt_path/lib:$PYTHONPATH

# head -c 1000 "2-way ball valves flangeable with SAE connections.xlsx.xml" > zzz.txt

# tidy "2-way ball valves flangeable with SAE connections.xlsx.xml" 2-way_ball.orig.tidy.xml
# xmllint --format "2-way ball valves flangeable with SAE connections.xlsx.xml" > 2-way_ball.orig.fmt.xml

rm 2-way_ball.xml
rm 2-way_ball.fmt.xml
# python3 /home/mingis/F/kp/src/xml/valvcvt/src/valvcvt/src/xlsxtrim.py "2-way ball valves flangeable with SAE connections.xlsx.xml" 2-way_ball.xml
python3 $valvcvt_path/xlsxtrim.py 2-way_ball.orig.xml 2-way_ball.xml
xmllint --format 2-way_ball.xml > 2-way_ball.fmt.xml

# xmllint --format "Cartridge ball valves.xlsx.xml" > cartridge.orig.fmt.xml

rm cartridge.xml
rm cartridge.fmt.xml
# python3 /home/mingis/F/kp/src/xml/valvcvt/src/valvcvt/src/xlsxtrim.py "Cartridge ball valves.xlsx.xml" cartridge.xml
python3 $valvcvt_path/xlsxtrim.py cartridge.orig.xml cartridge.xml
xmllint --format cartridge.xml > cartridge.fmt.xml

rm 2-way_SAE.xml
rm 2-way_SAE.fmt.xml
python3 $valvcvt_path/xlsxtrim.py 2-way_SAE.orig.xml 2-way_SAE.xml
xmllint --format 2-way_SAE.xml > 2-way_SAE.fmt.xml

rm 2-way_threaded.xml
rm 2-way_threaded.fmt.xml
python3 $valvcvt_path/xlsxtrim.py 2-way_threaded.orig.xml 2-way_threaded.xml
xmllint --format 2-way_threaded.xml > 2-way_threaded.fmt.xml
