This is a quick and dirty script for automatically numbering headings
and example sentences in google docs documents. It's designed for
linguistics papers.

Download the document as an ODT file, then run it through the script
as follows:

    python numberize.py infile.odt outfile.odt

The output can be imported back into google docs, or opened using
OpenOffice or a number of other packages.

This should work with the standard OS X Python installation (so long
as the Python version is >= 2.5).

To introduce a numbered example, write a series of upper case letters
enclosed in parantheses, followed by a tab:

    (XYZ)   My example...

To refer back to the example, simply write `(XYZ)` in the text. If you
want the example to be numbered using Roman numerals, introduce the
upper case letters with a `#`:

    (#ABC)  My Roman-numbered example.

Then refer back to it as `(ABC)`. The script handles references to
a/b/c examples in a very basic way. If you have an example of the
following sort:

    (LMN)   a. Blah blah blah...
            b. Blah blah blah...

You can refer either to `(LMN)`, `(LMNa)` or `(LMNb)`. Note that the
script is not aware of the sub-examples (you could just as well write
`(LMNz)`, which would come out as `(2z)` if `(LMN)` is the second
arabic-numbered example in the document).

References to non-existent examples are ignored. This can happen
sometimes if you put acronymys in parantheses (e.g. "The Uniformity of
Theta Assignment Hypothesis (UTAH)").  In this case, so long as there
is no example labeled UTAH, nothing will go wrong.

Headings and sub-headings are automatically numbered in the output. If
you want to refer back to a section, introduce the heading with a
sequence of upper case latters followed by a period:

    ABC. My heading

You can then get the number of this heading using `$ABC`.