# PUBLICATION_HELPERS

Small Python utilities that help with publishing and writing scientific papers.

## MDPI to author-year citations

`mdpi_to_elsevier_citation.py` converts numbered citations such as `[1]`, `[2, 3]`, or `[4–7]` in a Word manuscript into author-year citations based on the manuscript's reference list.

Example:

```bash
python mdpi_to_elsevier_citation.py PAPER_mdpi.docx -o PAPER_apa7.docx
```

The script supports common reference formats including personal authors, institutional authors, standards, webpages, and APA/MDPI-style entries.

### Requirements

```bash
pip install python-docx
```

> Use a different output filename so the original manuscript is preserved.
