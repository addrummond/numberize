import sys
import os
import re
import StringIO
import xml.etree.ElementTree
import zipfile
import tempfile

for x in [("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"),
          ("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"),
          ("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"),
          ("xlink", "http://www.w3.org/1999/xlink"),
          ("ooo", "http://openoffice.org/2004/office"),
          ("dc", "http://purl.org/dc/elements/1.1/"),
          ("meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"),
          ("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"),
          ("draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"),
          ("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
          ("number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"),
          ("chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"),
          ("svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"),
          ("dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"),
          ("math", "http://www.w3.org/1998/Math/MathML"),
          ("form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"),
          ("script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"),
          ("ooow", "http://openoffice.org/2004/writer"),
          ("oooc", "http://openoffice.org/2004/calc"),
          ("dom", "http://www.w3.org/2001/xml-events")]:
    xml.etree.ElementTree.register_namespace(x[0], x[1])

assert sys.argv[1]

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

_labre = re.compile(r"\(([A-Z]+)\)\t")
def search_and_replace_paragraph(elem, start_number):
    text, links = flatten(elem)
    for match in (re.finditer(_labre, text) or []):
        mapping[match.group(1)] = start_number
        replace_in_linked_string(text, match.start() + 1, match.end() - 2, links, str(start_number))
        start_number += 1
#        sys.stderr.write("Found (%s)\n" % match.group(1))
    return start_number

def search_and_replace2(current):
    for child in current:
        if strip_prefix(child.tag) == "p":
            search_and_replace_paragraph2(child)
        else:
            search_and_replace2(child)

_labre2 = re.compile(r"\(([A-Z]+)([a-z]*)\)[^\t]")
def search_and_replace_paragraph2(elem):
    text, links = flatten(elem)
    for match in (re.finditer(_labre2, text) or []):
        if not mapping.has_key(match.group(1)):
            sys.stderr.write("WARNING: Bad reference to (%s)\n" % match.group(1))
        else:
#            sys.stderr.write("Replacing (%s) with (%i)\n" % (match.group(1), mapping[match.group(1)]))
            replace_in_linked_string(text, match.start() + 1, match.end() - 2 - len(match.group(2)), links, str(mapping[match.group(1)]))

def flatten_(elem, text, links):
    if strip_prefix(elem.tag) == "span":
        if elem.text:
            text.write(elem.text)
            links[(links['current_i'], links['current_i']+len(elem.text))] = dict(type="text", elem=elem)
            links['current_i'] += len(elem.text)

        for child in elem:
            pr = strip_prefix(child.tag)
            if child.tail or pr == "tab": # Regex matching is sensitive to tabs so must include these.
                if pr == "tab":
                    text.write('\t')
                text.write(child.tail or "")
                l = len(child.tail or "") + (pr == "tab" and 1 or 0)
                links[(links['current_i'], links['current_i'] + l)] = dict(type="tail", elem=elem)
                links['current_i'] += l
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
    rks = filter(lambda k: k[0] >= start and k[0] <= end, ks)

    assert rks

    into = getattr(links[rks[0]]['elem'], links[rks[0]]['type']) or ""
    new = into[:start-rks[0][0]] + replacement + into[start-rks[0][0]+(end-start):]
    setattr(links[rks[0]]['elem'], links[rks[0]]['type'], new)

    for k in rks[1:-1]:
        setattr(links[k]['elem'], links[k]['type'], "")

    if len(rks) >= 3:
        into2 = getattr(links[rks[-1]]['elem'], links[rks[-1]]['type']) or ""
        setattr(links[k]['elem'], links[k]['type'], into2[end - rks[-1][0]:])

with zipfile.ZipFile(sys.argv[1], 'r') as odt:
    with odt.open('content.xml', 'r') as content:
        doc = xml.etree.ElementTree.parse(content)
        root = doc.getroot()
        search_and_replace(root, 1)
        search_and_replace2(root)

        with zipfile.ZipFile("out.odt", 'w') as newone:
            for n in odt.namelist():
                if n != "content.xml":
                    with odt.open(n, 'r') as f:
                        newone.writestr(n, f.read())
                else:
                    with tempfile.NamedTemporaryFile(delete=False) as t:
                        t.write("""<?xml version="1.0" encoding="utf-8" standalone="yes"?>""")
                        doc.write(t, 'utf-8')
                        t.close()
                        newone.write(t.name, n)
                        os.remove(t.name)
