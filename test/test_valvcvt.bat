set PYTHONPATH=..\lib:%PYTHONPATH%

set all_makefile=F:\kp\src\vtex\xml\tex4ht\tools\all_makefile\src_kp\all_makefile
set pythonpath=D:\bin\python33
set path=%pythonpath%;%path%

set jobname=2-way_ball

:: %all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o %jobname%.orig.tidy.xml "2-way ball valves flangeable with SAE connections.xlsx.xml"

:: python F:\kp\src\xml\valvcvt\src\valvcvt\valvcvt.py "2-way ball valves flangeable with SAE connections.xlsx.xml" 2-way_ball.xml
python ..\valvcvt.py 2-way_ball.orig.xml 2-way_ball.xml
%all_makefile%\bat\tidy.exe -xml -wrap 80 -indent -utf8 -o %jobname%.tidy.xml %jobname%.xml

