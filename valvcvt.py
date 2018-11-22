#!/usr/bin/python
# coding=UTF-8
__author__ = "Mindaugas Piešina"
__version__ = "0.0.1"
__maintainer__ = ""
__email__ = "mpiesina@netscape.net"
__status__ = "Development"

"""
Script to strip empty cells out of OpenOffice Calc exported xml file
using:
    python valvcvt.py input.xml output.xml
"""

import sys
import codecs

ns = 'urn:schemas-microsoft-com:office:spreadsheet'
ns_pref = '{%s}' % ns
ns_xsl = {'xmlns': ns}

try:
    from lxml import etree
except ImportError:
    print ('no lxml')
    import xml.etree.ElementTree as etree

# from lxml import objectify

def main():
    # ----------------------------------
    if (len(sys.argv) < 3):
        print("Error: Give input and output file names as parameters")
        sys.exit(2)

    in_fname = sys.argv[1]
    out_fname = sys.argv[2]

    try:
        with open(in_fname, 'rb') as in_file:
            in_data = in_file.read()
    except OSError as err:
        print("Error: Unable to open file %s (%s)" % (in_fname, err))
        sys.exit(1)

    try:
        new_dom = etree.fromstring(in_data)
    except Exception as err:
        print("Error: Bad input file format (%s)" % err)
        sys.exit(1)

    # objectify.deannotate(new_dom, cleanup_namespaces = True)

    # ----------------------------------
    # delete empty cells at the end of the rows
    for row in new_dom.xpath('//xmlns:Row', namespaces = ns_xsl):
        for cell in reversed(row.xpath('xmlns:Cell', namespaces = ns_xsl)):
            if (cell.tail or cell.text or list(cell)):
                break
            row.remove(cell)

    # delete empty rows at the end of the tables    
    for tab in new_dom.xpath('//xmlns:Table', namespaces = ns_xsl):
        for row in reversed(tab.xpath('xmlns:Row', namespaces = ns_xsl)):
            if (row.tail or row.text or list(row)):
                break
            tab.remove(row)

    # collect dictionary of used styles
    used_styles = {}
    for cell in new_dom.xpath('//xmlns:Cell', namespaces = ns_xsl):
        used_styles[cell.get(ns_pref + 'StyleID')] = True

    #for nod in new_dom.xpath('//*'):
    #    if (nod.get(ns_pref + 'StyleID') and (nod.tag != ns_pref + 'Cell')):
    #        print(nod.tag + ' ' + nod.get(ns_pref + 'StyleID'))

    # delete unused styles
    for sty in new_dom.xpath('//xmlns:Style', namespaces = ns_xsl):
        if (sty.get(ns_pref + 'ID') not in used_styles.keys()):
            sty.getparent().remove(sty)

    # ----------------------------------
    try:
        out_data = etree.tostring(new_dom).decode()
    except Exception as err:
        print("Error: Unable to convert data (%s)" % err)
        sys.exit(1)

    try:
        with open(out_fname, 'w') as out_file:
            out_file.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            out_file.write('<?mso-application progid="Excel.Sheet"?>')
            out_file.write(out_data)
    except Exception as err:
        print("Error: Unable to write file %s (%s)" % (out_fname, err))
        sys.exit(1)

if __name__ == "__main__":
    main()

