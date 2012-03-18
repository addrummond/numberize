import sys
import os
import re
import StringIO
import xml.etree.ElementTree
import zipfile
import tempfile

# From http://www.daniweb.com/software-development/python/code/216865
def int2roman(number):
    numerals = { 1 : "i", 4 : "iv", 5 : "v", 9 : "ix", 10 : "x", 40 : "xl", 
        50 : "l", 90 : "xc", 100 : "c", 400 : "cd", 500 : "d", 900 : "cm", 1000 : "m" }
    result = ""
    for value, numeral in sorted(numerals.items(), reverse=True):
        while number >= value:
            result += numeral
            number -= value
    return result

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

heading_style_to_level = dict() # E.g. P60 -> 3
mapping = dict()
heading_numbers = dict() # E.g. "XX" -> [1,4,2] (= 1.4.2)

STYLEPREF = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
TEXTPREF = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
_hre = re.compile(r"_(\d+)$")
def get_heading_styles(root):
    for elem in root:
        if strip_prefix(elem.tag) == "automatic-styles":
            for c in elem:
                if c.attrib.has_key(STYLEPREF + 'family') and c.attrib[STYLEPREF + 'family'] == 'paragraph':
                    s = c.attrib[STYLEPREF + 'parent-style-name']
                    if s.startswith("Heading_"):
                        m = re.search(_hre, s)
                        if m:
                            heading_style_to_level[c.attrib[STYLEPREF + 'name']] = int(m.group(1))

def search_and_replace_(current, current_number, current_rm_number, current_heading_number):
    for child in current:
        if strip_prefix(child.tag) == "p":
            current_number, current_rm_number = search_and_replace_paragraph(child, current_number, current_rm_number)
        elif strip_prefix(child.tag) == "h" and heading_style_to_level[child.attrib[TEXTPREF + "style-name"]] > 1:
            current_heading_number = search_and_replace_heading(child, current_heading_number)
        else:
            search_and_replace_(child, current_number, current_rm_number, current_heading_number)

def search_and_replace(root, current_number, current_rm_number, current_heading_number):
    get_heading_styles(root)
    search_and_replace_(root, current_number, current_rm_number, current_heading_number)

_labre = re.compile(r"\((#?[A-Z]+)\)\t")
def search_and_replace_paragraph(elem, start_number, start_rm_number):
    text, links = flatten(elem)
    for match in (re.finditer(_labre, text) or []):
        if mapping.has_key(match.group(1).lstrip('#')):
            sys.stderr.write("WARNING: Label (%s) is multiply defined.\n" % match.group(1).lstrip('#'))

        if match.group(1).startswith('#'):
            rm = int2roman(start_rm_number)
            mapping[match.group(1)[1:]] = rm
            replace_in_linked_string(text, match.start(1), match.end(1), links, rm)
            start_rm_number += 1
        else:
            mapping[match.group(1)] = start_number
            replace_in_linked_string(text, match.start(1), match.end(1), links, str(start_number))
            start_number += 1
    return start_number, start_rm_number

def str_heading_number(l):
    return '.'.join(map(str, l))

_headre = re.compile(r"^([A-Z]+)\.(.*)$")
_unnumbered = re.compile(r"^(\*\s*).*$")
def search_and_replace_heading(elem, start_number):
    text, links = flatten(elem)

    level = heading_style_to_level[elem.attrib[TEXTPREF + "style-name"]]

    if level - 1 == len(start_number): # Non-embedded heading
        start_number[-1] += 1
    elif level - 1 < len(start_number):
        for _ in xrange(level - 1, len(start_number)): start_number.pop()
        start_number[-1] += 1
    elif level - 1 > len(start_number):
        start_number.append(1)

    match = re.match(_headre, text)
    if match:
        if heading_mumbers.has_key(match.group(1)):
            sys.stderr.write("WARNING: Heading label '%s' is multiply defined.\n" % match.group(1))
        heading_numbers[match.group(1)] = map(lambda x: x, start_number)
        replace_in_linked_string(text, match.start(1), match.end(1), links, str_heading_number(start_number))
    else:
        match2 = re.match(_unnumbered, text)
        if match2:
            # An unnumbered heading -- just remove the '*' and any whitespace following it.
            replace_in_linked_string(text, 0, match2.end(1), links, "")
        else:
            replace_in_linked_string(text, 0, 1, links, str_heading_number(start_number) + '. ' + text[0])
    return start_number

def search_and_replace2(current):
    for child in current:
        if strip_prefix(child.tag) == "p":
            search_and_replace_paragraph2(child)
        else:
            search_and_replace2(child)

_labre2 = re.compile(r"\(([A-Z]+)(?:[a-z]*)(?:-([A-Z]+)(?:[a-z]*))?\)[^\t]")
_headre2 = re.compile(r"\$([A-Z]+)")
def search_and_replace_paragraph2(elem):
    text, links = flatten(elem)
    for match in (re.finditer(_labre2, text) or []):
        if not mapping.has_key(match.group(1)):
            sys.stderr.write("WARNING: Bad reference to (%s)\n" % match.group(1))
        else:
            replace_in_linked_string(text, match.start(1), match.end(1), links, str(mapping[match.group(1)]))

        if match.group(2):
            if not mapping.has_key(match.group(2)):
                sys.stderr.write("WARNING: Bad reference to (%s)\n" % match.group(2))
            else:
                replace_in_linked_string(text, match.start(2), match.end(2), links, str(mapping[match.group(2)]))

    text, links = flatten(elem)
    for match in (re.finditer(_headre2, text) or []):
        sl = [x for x in links.keys() if x != 'current_i']
        sl.sort()
        if not heading_numbers.has_key(match.group(1)):
            sys.stderr.write("WARNING: Bad reference to $%s\n" % match.group(1))
        else:
            replace_in_linked_string(text, match.start(), match.end(), links, str_heading_number(heading_numbers[match.group(1)]))

def flatten_(elem, text, links):
    if strip_prefix(elem.tag) == "span" and not (len(list(elem)) > 0 and strip_prefix(elem[0].tag) == "note"):
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
                links[(links['current_i'], links['current_i'] + l)] = dict(type="tail", elem=child)
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
    rks = filter(lambda k: k[0] < end and k[1] > start, ks)

    assert rks
#    debug_print_linked_string(string, links, keys=rks)

    into = getattr(links[rks[0]]['elem'], links[rks[0]]['type']) or ""
    new = into[:start-rks[0][0]] + replacement + into[start-rks[0][0]+(end-start):]
    setattr(links[rks[0]]['elem'], links[rks[0]]['type'], new)
#    sys.stderr.write("Replacing 1st with '%s'\n" % new)

    for k in rks[1:-1]:
        setattr(links[k]['elem'], links[k]['type'], "")
#        sys.stderr.write("Setting %i to empty\n" % rks.index(k))

    # This hasn't ever happened, so this code isn't actually tested.
    if len(rks) >= 2:
        into2 = getattr(links[rks[-1]]['elem'], links[rks[-1]]['type']) or ""
        new = into2[end - rks[-1][0]:]
        setattr(links[rks[-1]]['elem'], links[rks[-1]]['type'], new)
#        sys.stderr.write("Setting last to '%s'\n" % new)

def debug_print_linked_string(string, links, keys=None):
    if keys is None: keys = links.keys()
    sys.stderr.write("[[\n")
    lks = links.keys()
    lks.sort()
    for k in lks:
        if k != "current_i" and k in keys:
            sys.stderr.write("%i, %i [%s]: '%s'\n" % (k[0], k[1], links[k]['type'], string[k[0]:k[1]].encode('utf-8')))
    sys.stderr.write("]]\n\n")

with zipfile.ZipFile(sys.argv[1], 'r') as odt:
    with odt.open('content.xml', 'r') as content:
        doc = xml.etree.ElementTree.parse(content)
        root = doc.getroot()
        search_and_replace(root, 1, 1, [0])
        search_and_replace2(root)

        with zipfile.ZipFile(sys.argv[2] or "out.odt", 'w') as newone:
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
