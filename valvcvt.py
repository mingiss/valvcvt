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

    # objectify.deannotate(new_dom, cleanup_namespaces=True)

    # ----------------------------------
    for wb in new_dom.xpath('//xmlns:Workbook', namespaces = ns_xsl):
        print(wb.attrib)
#        if (tab.tail == '\n\n'):
#            tab.tail = '\n'
#
#    for row in new_dom.xpath('//row'):
#        if (row.text and not row.text.strip()):
#            row.text = ''
#
#    for entry in new_dom.xpath('//row/entry'):
#        if (entry.tail and not entry.tail.strip()):
#            entry.tail = ''
#        char = entry.get('char')
#        if (char and (ord(char) > 127)):
#            entry.set('char', '?')
#
#    for tag in new_dom.iter():
#        attrs = dict(tag.attrib) 
#        tag.attrib.clear()
#        attr_keys = attrs.keys()
#        for key in sorted(attr_keys):
#            tag.set(key, attrs[key])
    
    # ----------------------------------
    try:
        out_data = etree.tostring(new_dom).decode()
    except Exception as err:
        print("Error: Unable to convert data (%s)" % err)
        sys.exit(1)

    try:
        with open(out_fname, 'w') as out_file:
            out_file.write(out_data)
    except Exception as err:
        print("Error: Unable to write file %s (%s)" % (out_fname, err))
        sys.exit(1)

if __name__ == "__main__":
    main()

