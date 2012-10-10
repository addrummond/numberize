import sys
import os
import re
import StringIO
import xml.etree.ElementTree
import zipfile
import tempfile
import time

now = time.gmtime()

MAX_SUB_EXAMPLE_LETTERS = 3
FIX_FOOTNOTE_NUMBERING = True

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

def permissible_label_char(c):
    return c.isupper()

heading_style_to_level = dict() # E.g. P60 -> 3
mapping = dict()
heading_numbers = dict() # E.g. "XX" -> [1,4,2] (= 1.4.2)
fn_numbers = dict() # E.g. "XX" -> 7

STYLEPREF = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
TEXTPREF = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
OFFICEPREF = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"

T_AUTOMATIC_STYLES = OFFICEPREF + 'automatic-styles'
T_FAMILY = STYLEPREF + 'family'
T_PARENT_STYLE_NAME = STYLEPREF + 'parent-style-name'
T_NAME = STYLEPREF + 'name'
T_P = TEXTPREF + 'p'
T_H = TEXTPREF + 'h'
T_STYLE_NAME = TEXTPREF + 'style-name'
T_NOTE = TEXTPREF + 'note'
T_NOTE_CITATION = TEXTPREF + 'note-citation'
T_TAB = TEXTPREF + 'tab'
T_SPAN = TEXTPREF + 'span'
T_S = TEXTPREF + 's'

_hre = re.compile(r"_(\d+)$")
def get_heading_styles(root):
    for elem in root:
        if elem.tag == T_AUTOMATIC_STYLES:
            for c in elem:
                if c.attrib.has_key(T_FAMILY) and c.attrib[T_FAMILY] == 'paragraph':
                    s = c.attrib[T_PARENT_STYLE_NAME]
                    if s.startswith("Heading_"):
                        m = re.search(_hre, s)
                        if m:
                            heading_style_to_level[c.attrib[T_NAME]] = int(m.group(1))

def search_and_replace_(current, current_number, current_rm_number, current_heading_number, current_fn_number):
    for child in current:
        if child.tag == T_P:
            current_number, current_rm_number, current_fn_number = \
                search_and_replace_paragraph(child, current_number, current_rm_number, current_fn_number)
        elif child.tag == T_H: #and heading_style_to_level[child.attrib[T_STYLE_NAME]] > 1:
            current_heading_number = search_and_replace_heading(child, current_heading_number)
        else:
            search_and_replace_(child, current_number, current_rm_number, current_heading_number, current_fn_number)

def search_and_replace(root, current_number, current_rm_number, current_heading_number, current_fn_number):
    get_heading_styles(root)
    search_and_replace_(root, current_number, current_rm_number, current_heading_number, current_fn_number)

_fnre = re.compile(r"^\s*(!*)([A-Z]+)(\.\s*).*")
def frisk_for_footnotes(elem, current_fn_number):
    if elem.tag == T_NOTE:
        text, links = flatten(elem)
        m = re.match(_fnre, text)
        if m:
            if m.group(1): # It's escaped; delete a '!' and move on.
                replace_in_linked_string(text, m.start(1), m.end(1), links, m.group(1)[1:])
            else:
                if fn_numbers.has_key(m.group(2)):
                    sys.stderr.write("WARNING: Footnote label '%s' is multiply defined.\n" % m.group(2))
                fn_numbers[m.group(2)] = current_fn_number
                replace_in_linked_string(text, m.start(), m.end(3), links, "")

        current_fn_number += 1
    else:
        for child in elem:
            current_fn_number = frisk_for_footnotes(child, current_fn_number)
    return current_fn_number

_labre = re.compile(r"\((!*)(#?[A-Z]+)\)(?:\t|(?:  ))")
def search_and_replace_paragraph(elem, start_number, start_rm_number, current_fn_number):
    current_fn_number = frisk_for_footnotes(elem, current_fn_number)

    text, links = flatten(elem)
    for match in (re.finditer(_labre, text) or []):
        if match.group(1): # It's escaped; delete a '!' and move on.
            replace_in_linked_string(text, match.start(1), match.end(1), links, match.group(1)[1:])
        else:
            if mapping.has_key(match.group(2).lstrip('#')):
                sys.stderr.write("WARNING: Label (%s) is multiply defined.\n" % match.group(2).lstrip('#'))

            if match.group(2).startswith('#'):
                rm = int2roman(start_rm_number)
                mapping[match.group(2)[1:]] = rm
                replace_in_linked_string(text, match.start(2), match.end(2), links, rm)
                start_rm_number += 1
            else:
                mapping[match.group(2)] = start_number
                replace_in_linked_string(text, match.start(2), match.end(2), links, str(start_number))
                start_number += 1
    return start_number, start_rm_number, current_fn_number

def str_heading_number(l):
    return '.'.join(map(str, l))

_headre = re.compile(r"^([A-Z]+)\.(.*)$")
_unnumbered = re.compile(r"^(\*\s*).*$")
def search_and_replace_heading(elem, start_number):
    text, links = flatten(elem)

    level = heading_style_to_level[elem.attrib[TEXTPREF + "style-name"]]

    if level == len(start_number): # Non-embedded heading
        pass
    elif level < len(start_number):
        for _ in xrange(level, len(start_number)): start_number.pop()
    elif level > len(start_number):
        start_number.append(0)

    # Updates start number appropriately if we find we're dealing with an unnumbered heading.
    def update_start_number():
        if level <= len(start_number):
            start_number[-1] += 1        

    match = re.match(_headre, text)
    if match:
        if heading_numbers.has_key(match.group(1)):
            sys.stderr.write("WARNING: Heading label '%s' is multiply defined.\n" % match.group(1))
        update_start_number()
        heading_numbers[match.group(1)] = map(lambda x: x, start_number)
        replace_in_linked_string(text, match.start(1), match.end(1), links, str_heading_number(start_number))
    else:
        match2 = re.match(_unnumbered, text)
        if match2:
            # An unnumbered heading -- just remove the '*' and any whitespace following it.
            replace_in_linked_string(text, 0, match2.end(1), links, "")
        else:
            # It's a numbered heading without a label.
            update_start_number()
            replace_in_linked_string(text, 0, 1, links, str_heading_number(start_number) + '. ' + (text and text[0] or ''))

    return start_number

def search_and_replace2(current):
    for child in current:
        if child.tag == T_P:
            search_and_replace_paragraph2(child)
        else:
            search_and_replace2(child)

_labre2 = r"\((!?)([A-Z]+)(?:(?:[a-z]{0,%i}(?:[-/][a-z]{0,%i})?)|(?:[-/]([A-Z]+)))\)(?!\t|(?:  ))" % ((MAX_SUB_EXAMPLE_LETTERS,)*2)
_headre2 = re.compile(r"(\$+)([A-Z]+)")
_fnre2 = re.compile(r"(\^+)([A-Z]+)")
_replre = re.compile(r"%\[([a-zA-z]+)\]")
def search_and_replace_paragraph2(elem):
    text, links = flatten(elem)
    for match in (re.finditer(_labre2, text) or []):
        if match.group(1): # It's escaped; delete the '!' and move on.
            replace_in_linked_string(text, match.start(1), match.end(1), links, "")
        else:
            if not mapping.has_key(match.group(2)):
                sys.stderr.write("WARNING: Bad reference to (%s)\n" % match.group(2))
            else:
                replace_in_linked_string(text, match.start(2), match.end(2), links, str(mapping[match.group(2)]))

                if match.group(3):
                    if not mapping.has_key(match.group(3)):
                        sys.stderr.write("WARNING: Bad reference to (%s)\n" % match.group(3))
                    else:
                        replace_in_linked_string(text, match.start(3), match.end(3), links, str(mapping[match.group(3)]))

    text, links = flatten(elem)
    for match in (re.finditer(_headre2, text) or []):
        if len(match.group(1)) > 1: # It's escaped; strip a '$' and move on.
            replace_in_linked_string(text, match.start(1), match.start(1)+1, "")
        else:
            sl = [x for x in links.spans.keys()]
            sl.sort()
            if not heading_numbers.has_key(match.group(2)):
                sys.stderr.write("WARNING: Bad reference to heading $%s\n" % match.group(2))
            else:
                replace_in_linked_string(text, match.start(), match.end(), links, str_heading_number(heading_numbers[match.group(2)]))

    text, links = flatten(elem)
    for match in (re.finditer(_fnre2, text) or []):
        if len(match.group(1)) > 1: # It's escaped; strip a '^' and move on.
            replace_in_linked_string(text, match.start(1), match.start(1)+1, "")
        else:
            if not fn_numbers.has_key(match.group(2)):
                sys.stderr.write("WARNING: Bad reference to footnote ^%s\n" % match.group(2))
            else:
                replace_in_linked_string(text, match.start(), match.end(), links, str(fn_numbers[match.group(2)]))

    text, links = flatten(elem)
    for match in (re.finditer(_replre, text) or []):
        w = match.group(1).upper()
        r = ""
        repl = dict(DAY=lambda: time.strftime("%d"), DAYNAME=lambda: time.strftime("%A"),
                    MONTH=lambda: time.strftime("%m"), MONTHNAME=lambda: time.strftime("%B"),
                    YEAR=lambda: time.strftime("%Y"))
        if repl.has_key(w):
            r = repl[w]()
        else:
            sys.stderr.write("WARNING: Unrecognized replacement '%s'\n" % w)
        replace_in_linked_string(text, match.start(), match.end(), links, r)

def number_footnotes(elem, cite_count=1, fn_count=1):
    if elem.tag == T_NOTE_CITATION:
        elem.text = str(cite_count)
        cite_count += 1
    elif elem.tag == T_NOTE:
        elem.attrib[TEXTPREF + 'id'] = "ftn%i" % fn_count
        fn_count += 1

    for c in elem:
        cite_count, fn_count = number_footnotes(c, cite_count, fn_count)
    return cite_count, fn_count

class Links(object):
    def __init__(self):
        self.spans = { }
        self.current_i = 0

def flatten_(elem, text, links):
    if elem.tag == T_SPAN and not (len(list(elem)) > 0 and elem[0].tag == T_NOTE):
        if elem.text:
            text.write(elem.text)
            links.spans[(links.current_i, links.current_i+len(elem.text))] = dict(type="text", elem=elem)
            links.current_i += len(elem.text)

        for child in elem:
            if child.tail or child.tag == T_TAB or child.tag == T_S: # Regex matching is sensitive to tabs so must include these.
                add = 0
                
                if child.tag == T_TAB or child.tag == T_S:
                    add = 1
                    if child.attrib.has_key(TEXTPREF + 'c'):
                        add = child.attrib[TEXTPREF + 'c']
                        try:
                            add = int(add)
                        except:
                            sys.stderr.write("ERROR: Bad value for 'c' attribute")
                            sys.exit(1)
                    text.write((child.tag == T_TAB and '\t' or ' ') * add)

                links.current_i += add

                if child.tail:
                    text.write(child.tail)
                    links.spans[(links.current_i, links.current_i + len(child.tail))] = dict(type="tail", elem=child)
                    links.current_i += len(child.tail)
    else:
        for c in elem:
            current_i = flatten_(c, text, links)

    if elem.tail and elem.tail != '':
        text.write(elem.tail)
        links.spans[(links.current_i, links.current_i+len(elem.tail))] = dict(type="tail", elem=elem)
        links.current_i += len(elem.tail)

def flatten(elem):
    text = StringIO.StringIO()
    links = Links()
    flatten_(elem, text, links)
    return (text.getvalue(), links)

def replace_in_linked_string(string, start, end, links, replacement):
    ks = links.spans.keys()
    ks.sort()
    rks = filter(lambda k: k[0] < end and k[1] > start, ks)

    if len(rks) == 0:
        return
#    debug_print_linked_string(string, links, keys=rks)

    into = getattr(links.spans[rks[0]]['elem'], links.spans[rks[0]]['type']) or ""
    new = into[:start-rks[0][0]] + replacement + into[start-rks[0][0]+(end-start):]
    setattr(links.spans[rks[0]]['elem'], links.spans[rks[0]]['type'], new)
#    sys.stderr.write("Replacing 1st with '%s'\n" % new)

    for k in rks[1:-1]:
        setattr(links.spans[k]['elem'], links.spans[k]['type'], "")
#        sys.stderr.write("Setting %i to empty\n" % rks.index(k))

    if len(rks) >= 2:
        into2 = getattr(links.spans[rks[-1]]['elem'], links.spans[rks[-1]]['type']) or ""
        new = into2[end - rks[-1][0]:]
        setattr(links.spans[rks[-1]]['elem'], links.spans[rks[-1]]['type'], new)
#        sys.stderr.write("Setting last to '%s'\n" % new)

def debug_print_linked_string(string, links, keys=None):
    if keys is None: keys = links.spans.keys()
    sys.stderr.write("[[\n")
    lks = links.spans.keys()
    lks.sort()
    for k in lks:
        if k != "current_i" and k in keys:
            sys.stderr.write("%i, %i [%s]: '%s'\n" % (k[0], k[1], links.spans[k]['type'], string[k[0]:k[1]].encode('utf-8')))
    sys.stderr.write("]]\n\n")

with zipfile.ZipFile(sys.argv[1], 'r') as odt:
    with odt.open('content.xml', 'r') as content:
        doc = xml.etree.ElementTree.parse(content)
        root = doc.getroot()
        search_and_replace(root, 1, 1, [0], 1)
        search_and_replace2(root)
        if FIX_FOOTNOTE_NUMBERING:
            number_footnotes(root, 1)

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
