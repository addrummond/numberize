This is a quick and dirty script for automatically numbering headings
and example sentences in google docs documents. It's designed for
linguistics papers.

Download the document as an ODT file, then run it through the script
as follows:

    python numberize.py infile.odt outfile.odt

The output can be imported back into google docs, or opened using
OpenOffice or a number of other packages.

This should work with the standard OS X Python installation (so long
as the Python version is ≥ 2.5).

To introduce a numbered example, write a series of upper case letters
enclosed in parantheses, followed by a tab:

    (XYZ)   My example...

To refer back to the example, simply write `(XYZ)` in the text. In
addition to references to individual examples, the script recognizes
ranges of the following form:

    ...blah blah blah examples (AB-XY)...

If you want the example to be numbered using Roman numerals, introduce
the upper case letters with a `#`:

    (#ABC)  My Roman-numbered example.

Then refer back to it as `(ABC)`. The script handles references to
a/b/c subexamples in a very basic way. If you have an example of the
following sort:

    (LMN)   a. Blah blah blah...
            b. Blah blah blah...

You can refer either to `(LMN)`, `(LMNa)` or `(LMNb)`. Note that the
script is not aware of the sub-examples (you could just as well write
`(LMNz)`, which would come out as `(2z)` if `(LMN)` were the second
arabic-numbered example in the document).  You can also write things
like `(ABc-d)`, which comes out as `(20c-d)`, or (ABa/e), which comes
out as `(20a/e)`.

To avoid bracketed words being mistaken for references to examples,
the script assumes that no subexamples are labeled with more than one
letter. (If a brackted word is mistaken for a reference, this is not
usually a problem, unless an example happens to be defined with the
corresponding label.) You can change this default by modifying the
value of `MAX_SUB_EXAMPLE_LETTERS`.

References to non-existent examples are ignored. This can happen
sometimes if you put acronyms in parantheses (e.g. "The Uniformity of
Theta Assignment Hypothesis (UTAH)").  In this case, so long as there
is no example labeled UTAH, nothing will go wrong.

Occasionally, you may want to explicitly specify that something which
looks like a reference to an example is in fact not. (For example, if
you have the word 'I' in parantheses in a gloss, and an example
labeled `(I)`) You can prevent a sequence of capital letters in
parentheses from being interpreted as a reference to an example by
insering a bang: `(!FOO)`. The output `(!FOO)` can be produced using
`(!!FOO)`; `(!)` will come out unaltered. It should rarely if ever be
necessary to use bangs in the manner -- the need arises only if you
happen to have defined an example with the corresponding label.

Headings and sub-headings are automatically numbered in the output. If
you want to refer back to a section, introduce the heading with a
sequence of upper case latters followed by a period:

    ABC. My heading

You can then get the number of this heading using `$ABC`. If you do
not want a heading to be numbered, precede it with a `*`:

    *My heading

The `*` is stripped in the output and the heading is left
unnumbered. To produce the literal output `$ABC`, use `$$ABC`.
