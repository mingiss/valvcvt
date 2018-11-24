#! /bin/bash

# script to delete trailing empty cells and rows from xslx.xml files
# source files should be placed to subfolder orig, output is written to subfolder trimmed

# cd orig
# mv "2-way ball valves flangeable with SAE connections.xlsx.xml"                              2-way_ball_valves_flangeable_with_SAE_connections.xlsx.xml
# mv "2-way ball valves for gas with threaded connections.xlsx.xml"                            2-way_ball_valves_for_gas_with_threaded_connections.xlsx.xml
# mv "2-way ball valves for isocyanate with threaded connections.xlsx.xml"                     2-way_ball_valves_for_isocyanate_with_threaded_connections.xlsx.xml
# mv "2-way ball valves for paints and lacquers.xlsx.xml"                                      2-way_ball_valves_for_paints_and_lacquers.xlsx.xml
# mv "2-way ball valves with DIN connections.xlsx.xml"                                         2-way_ball_valves_with_DIN_connections.xlsx.xml
# mv "2-way ball valves with fire-safe approval.xlsx.xml"                                      2-way_ball_valves_with_fire-safe_approval.xlsx.xml
# mv "2-way ball valves with ISO connections.xlsx.xml"                                         2-way_ball_valves_with_ISO_connections.xlsx.xml
# mv "2-way ball valves with SAE connections.xlsx.xml"                                         2-way_ball_valves_with_SAE_connections.xlsx.xml
# mv "2-way ball valves with threaded connections.xlsx.xml"                                    2-way_ball_valves_with_threaded_connections.xlsx.xml
# mv "2-way ball valves with welding ends.xlsx.xml"                                            2-way_ball_valves_with_welding_ends.xlsx.xml
# mv "2-way highest pressure ball valves with threaded connections.xlsx.xml"                   2-way_highest_pressure_ball_valves_with_threaded_connections.xlsx.xml
# mv "2-way low-pressure ball valves with threaded connections.xlsx.xml"                       2-way_low-pressure_ball_valves_with_threaded_connections.xlsx.xml
# mv "3-2-way selector ball valves with threaded connections - SAE connections.xlsx.xml"       3-2-way_selector_ball_valves_with_threaded_connections_-_SAE_connections.xlsx.xml
# mv "3-way and 4-way ball valves with threaded connections.xlsx.xml"                          3-way_and_4-way_ball_valves_with_threaded_connections.xlsx.xml
# mv "Ball valves for manifold mounting.xlsx.xml"                                              Ball_valves_for_manifold_mounting.xlsx.xml
# mv "Cartridge ball valves.xlsx.xml"                                                          Cartridge_ball_valves.xlsx.xml
# cd ..

valvcvt_path=..
export PYTHONPATH=$valvcvt_path/lib:$PYTHONPATH

mkdir trimmed

# file_list=`ls orig`
# file_list=`ls orig | sed "s/^/\"/g" | sed  "s/\$/\"/g"`
# file_list=`ls orig | sed "s/ /\\\\ /g"`
file_list=`ls orig | sed "s/ /_/g"`
for fb in $file_list
do
    fs=`echo $fb | sed "s/_/ /g"`
    echo "Processing $fs"
    python3.4 $valvcvt_path/xlsxtrim.py "orig/$fs" "trimmed/$fs"
done
