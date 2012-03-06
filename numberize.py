import sys
import re
import xml.etree.ElementTree
import StringIO

assert sys.argv[1]

doc = xml.etree.ElementTree.parse(sys.argv[1])
root = doc.getroot()

f = open("/tmp/de", "w")

_tagre = re.compile(r"(?:\{[^}]*\})?(.*)")
def strip_prefix(tagname):
    return re.match(_tagre, tagname).group(1)

def permissible_label_char(c):
    return c.isupper()

mapping = dict()

def search_and_replace(current, current_number):
    for child in current:
        if strip_prefix(child.tag) == "p":
            current_number = search_and_replace_paragraph(child, current_number)
        else:
            search_and_replace(child, current_number)

_labre = re.compile(r"\(([A-Z]+)[a-z]*\)\t")
def search_and_replace_paragraph(elem, start_number):
    text, links = flatten(elem)
    for match in (re.finditer(_labre, text) or []):
        mapping[match.group(1)] = start_number
        replace_in_linked_string(text, match.start(), match.end()-1, links, str(start_number))
        start_number += 1
        sys.stderr.write("Found (%s)\n" % match.group(1))
    return start_number

def search_and_replace2(current):
    for child in current:
        if strip_prefix(child.tag) == "p":
            search_and_replace_paragraph2(child)
        else:
            search_and_replace2(child)

_labre2 = re.compile(r"\(([A-Z]+)[a-z]*\)[^\t]")
def search_and_replace_paragraph2(elem):
    text, links = flatten(elem)
    f.write(text.encode('utf-8') + "\n\n")
    for match in (re.finditer(_labre2, text) or []):
        if not mapping.has_key(match.group(1)):
            sys.stderr.write("WARNING: Bad reference to (%s)\n" % match.group(1))
        else:
            sys.stderr.write("Replacing (%s) with (%i)\n" % (match.group(1), mapping[match.group(1)]))
            replace_in_linked_string(text, match.start(), match.end(), links, str(mapping[match.group(1)]))

def flatten_(elem, text, links):
    if strip_prefix(elem.tag) == "span":
        if elem.text:
            # Convert <tab> tags to tabs in the text.
            tabs_at = []
            for child in elem:
                if strip_prefix(child.tag) == "tab":
                    tabs_at.append(len(elem.text) - (child.tail and len(child.tail) or 0))

            for i in xrange(len(elem.text)):
                if i in tabs_at:
                    sys.stderr.write("TAB\n")
                    text.write('\t')
                text.write(elem.text[i])
            if len(tabs_at) > 0 and tabs_at[len(tabs_at)-1] == len(elem.text):
                sys.stderr.write("TAB\n")
                text.write('\t')

            links[(links['current_i'], links['current_i']+len(elem.text))] = elem
            links['current_i'] += len(elem.text) + len(tabs_at)
    else:
        for c in elem:
            current_i = flatten_(c, text, links)

def flatten(elem):
    text = StringIO.StringIO()
    links = dict(current_i=0)
    flatten_(elem, text, links)
    return (text.getvalue(), links)

def replace_in_linked_string(string, start, end, links, replacement):
    ks = links.keys()
    ks.sort()
    rks = filter(lambda k: k[0] >= start and k[1] <= end, ks)
    replacement_pos = 0
    for k in rks:
        st = start - k[0]
        en = max(end, k[1])
        new = string[0:st] + replacement[replacement_pos:replacement_pos+en-st] + string[en:]
        replacement_pos += en-st
        links[k].text = new

search_and_replace(root, 1)
search_and_replace2(root)
#xml.etree.ElementTree.dump(root)
