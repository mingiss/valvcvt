set valvcvt_path=..\src
set PYTHONPATH=%valvcvt_path%\lib:%PYTHONPATH%

set all_makefile=F:\kp\src\vtex\xml\tex4ht\tools\all_makefile\src_kp\all_makefile
set pythonpath=D:\bin\python33
set path=%pythonpath%;%path%

set jobname=2-way_ball

:: %all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o %jobname%.orig.tidy.xml "2-way ball valves flangeable with SAE connections.xlsx.xml"
:: %all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o %jobname%.orig.tidy.xml %jobname%.orig.xml

:: %all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o cartridge.orig.tidy.xml "Cartridge ball valves.xlsx.xml"
:: %all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o cartridge.orig.tidy.xml cartridge.orig.xml

python %valvcvt_path%\valvcvt.py files_orig.txt result.xlsx.xml
%all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o result.xlsx.tidy.xml result.xlsx.xml
