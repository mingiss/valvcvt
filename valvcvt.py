#!/usr/bin/python
# coding=UTF-8

"""
valvcvt.py

Strip empty cells out of OpenOffice Calc exported xml file.

Using:
    python valvcvt.py input.xml output.xml
"""

__author__ = "Mindaugas Piešina"
__version__ = "0.0.1"
__email__ = "mpiesina@netscape.net"
__status__ = "Prototype"

import sys
import codecs

try:
    from lxml import etree
except ImportError:
    print ('no lxml')
    import xml.etree.ElementTree as etree

from xlstree import xlstree
from xlstree import ns_xsl
from xlstree import ns_pref

# ----------------------------------
class valvtree(xlstree):
    '''xlsx.xml tools localized to valvcvt'''

    def del_heads(self):
        for tab in self.dom.xpath('//xmlns:Table', namespaces = ns_xsl):
            prev_spanned = False
            for row in tab.xpath('xmlns:Row', namespaces = ns_xsl):
                row_spanned = False
                for cell in row.xpath('xmlns:Cell', namespaces = ns_xsl):
                    try:
                        span = int(cell.get(ns_pref + 'MergeAcross'))
                        if (span > 1):
                            row_spanned = True
                    except:
                        pass
                if (prev_spanned and (not row_spanned)):
                    # previous row had spanned cells -- this row is a heading -- deleting
                    print('Removed heading: ' + ''.join(row.xpath('.//text()')))
                    row.getparent().remove(row)
                prev_spanned = row_spanned


# ----------------------------------
def main():
    # ----------------------------------
    if (len(sys.argv) < 3):
        print("Error: Give input and output file names as parameters")
        sys.exit(2)

    in_fname = sys.argv[1]
    out_fname = sys.argv[2]

    tree = valvtree()

    if (not tree.load(in_fname)):
        print("Error: " + tree.last_error)
        sys.exit(1)

    # ----------------------------------
    tree.trim()
    tree.del_heads()

    # ----------------------------------
    if (not tree.write(out_fname)):
        print("Error: " + tree.last_error)
        sys.exit(1)

if __name__ == "__main__":
    main()
