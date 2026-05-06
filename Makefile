DOCNAME = Merkle-tree-HDF5
TEXMFHOME_OVERRIDE = /dev/null

LATEXFLAGS = -interaction=nonstopmode -halt-on-error
LATEX = TEXMFHOME=$(TEXMFHOME_OVERRIDE) pdflatex $(LATEXFLAGS)
BIBTEX = bibtex
PANDOC = pandoc
PANDOCFLAGS = --standalone --bibliography=references.bib --citeproc \
              --toc --wrap=none

AUXFILES = $(DOCNAME).aux $(DOCNAME).bbl $(DOCNAME).blg \
           $(DOCNAME).log $(DOCNAME).out $(DOCNAME).toc

.PHONY: all clean distclean view md

all: $(DOCNAME).pdf

$(DOCNAME).pdf: $(DOCNAME).tex references.bib
	$(LATEX) $(DOCNAME)
	$(BIBTEX) $(DOCNAME)
	$(LATEX) $(DOCNAME)
	$(LATEX) $(DOCNAME)

md: $(DOCNAME).md

$(DOCNAME).md: $(DOCNAME).tex references.bib
	$(PANDOC) $(PANDOCFLAGS) -f latex -t gfm -o $@ $<

clean:
	rm -f $(AUXFILES)

distclean: clean
	rm -f $(DOCNAME).pdf $(DOCNAME).md

view: $(DOCNAME).pdf
	xdg-open $(DOCNAME).pdf &
