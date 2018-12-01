#!/usr/bin/python
# coding=UTF-8

"""
xlsxtrim.py

Trim empty cells out of OpenOffice Calc exported xml file.

Using:
    python xlstrim.py input.xml output.xml
"""

__author__ = "Mindaugas Piešina"
__version__ = "0.0.1"
__email__ = "mpiesina@netscape.net"
__status__ = "Prototype"

import sys
import codecs

from xlstree import XlsTree

def main():
    # ----------------------------------
    if (len(sys.argv) < 3):
        print("Error: Give input and output file names as parameters")
        sys.exit(2)

    in_fname = sys.argv[1]
    out_fname = sys.argv[2]

    tree = XlsTree()

    if (not tree.load(in_fname)):
        print("Error: " + tree.last_error)
        sys.exit(1)

    # ----------------------------------
    tree.trim()

    # ----------------------------------
    if (not tree.write(out_fname)):
        print("Error: " + tree.last_error)
        sys.exit(1)

if __name__ == "__main__":
    main()
