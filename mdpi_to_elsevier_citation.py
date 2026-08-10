#```python
#!/usr/bin/env python3

"""
mdpi_to_elsevier_citation.py

Convert numerical citations in a DOCX manuscript to APA 7-style
author-year citations using the manuscript's reference list.

Examples
--------
[1]
    -> (Smith, 2023)

[1, 2]
    -> (Jones et al., 2022; Smith, 2023)

[1-3]
    -> (Jones et al., 2022; Miller & White, 2021; Smith, 2023)

Supported reference styles include:

MDPI:
    Nakajima, K.; Takeda, O.; Miki, T.; ... 2010, ...

APA-like:
    Koinig, G., Aberger, J., Lorber, B., ... & Tischberger-Aldrian, A. (2026).

Non-inverted names:
    Huellen M, Schrade C, Wilhelm U, Zulhan Z (2006) ...

Full name:
    Nils Kroell (2021). ...

Institution:
    Institute for Energy Economics and Financial Analysis. Steel Fact Sheet. 2022.

Undated institution:
    Bundesvereinigung Deutscher Stahlrecycling- und Entsorgungsunternehmen e.V.
    (n.d.). European Steel Scrap Specification.

Undated webpage:
    Leuchtenmüller, M. Hydrogen Reduction Lab. ...
    (accessed on 08 May 2026)

Standard:
    BS EN 10025-2:2019; ...

Requirements
------------
pip install python-docx

Usage
-----
python mdpi_to_elsevier_citation.py input.docx -o output.docx
"""

from __future__ import annotations

import argparse
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from docx import Document


# ======================================================================
# DATA MODEL
# ======================================================================

@dataclass
class Reference:
    number: int
    authors: list[str]
    year: str
    cite_year: str = ""
    raw_text: str = ""


# ======================================================================
# REGULAR EXPRESSIONS
# ======================================================================

# Numerical citations:
#
# [1]
# [1, 3]
# [1; 3]
# [1-3]
# [1–3]
# [1—3, 7]
#
CITATION_RE = re.compile(
    r"""
    \[
        (?P<inside>
            \s*\d+
            (?:\s*[-–—]\s*\d+)?
            (?:
                \s*[,;]\s*
                \d+
                (?:\s*[-–—]\s*\d+)?
            )*
        )
    \]
    """,
    re.X,
)


# Reference numbering:
#
# 1. Smith...
# 1) Smith...
# [1] Smith...
#
REFERENCE_NUMBER_RE = re.compile(
    r"""
    ^\s*
    (?:
        \[(?P<bracket>\d+)\]
        |
        (?P<normal>\d+)[.)]
    )
    \s*
    (?P<text>.*)
    $
    """,
    re.X | re.S,
)


# Parenthesized publication date:
#
# Koinig ... (2026)
# Institution ... (n.d.)
#
PAREN_DATE_RE = re.compile(
    r"""
    ^\s*
    (?P<authors>.+?)
    \s*
    \(
        (?P<year>
            (?:18|19|20)\d{2}[a-z]?
            |
            n\.?\s*d\.?
        )
    \)
    \.?
    \s*
    (?P<rest>.*)
    $
    """,
    re.I | re.X | re.S,
)


# Corporate / institutional author hints
CORPORATE_HINT_RE = re.compile(
    r"""
    \b(
        institute
        |association
        |university
        |agency
        |committee
        |society
        |federation
        |bundesvereinigung
        |ministerium
        |ministry
        |department
        |organisation
        |organization
        |commission
        |council
        |authority
        |administration
        |office
        |board
        |foundation
        |laboratory
        |centre
        |center
        |gmbh
        |ag
        |inc
        |ltd
        |e\.?\s*v\.?
    )\b
    """,
    re.I | re.X,
)


# ======================================================================
# BASIC HELPERS
# ======================================================================

def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_name(text: str) -> str:
    return normalize_whitespace(text).strip(" ,;.")


def normalize_year(year: str) -> str:
    year = year.strip()

    if re.fullmatch(
        r"n\.?\s*d\.?",
        year,
        flags=re.I,
    ):
        return "n.d."

    return year


def text_before_access_date(text: str) -> str:
    """
    Remove access-date material so an access year such as 2026 is not
    mistakenly interpreted as the publication year.
    """

    return re.split(
        r"\b(?:accessed|retrieved|last accessed)\b",
        text,
        maxsplit=1,
        flags=re.I,
    )[0]


def looks_like_web_reference(text: str) -> bool:
    return bool(
        re.search(
            r"https?://|www\.|Available\s+online",
            text,
            flags=re.I,
        )
    )


# ======================================================================
# YEAR EXTRACTION
# ======================================================================

def extract_year(
    text: str,
    allow_undated_web: bool = False,
) -> str | None:

    source = text_before_access_date(text)

    # Explicit n.d.
    if re.search(
        r"\bn\.?\s*d\.?\b",
        source,
        flags=re.I,
    ):
        return "n.d."

    # (2026)
    years = re.findall(
        r"\(((?:18|19|20)\d{2}[a-z]?)\)",
        source,
        flags=re.I,
    )

    if years:
        return years[-1]

    # Typical MDPI:
    #
    # 2010, 44, 5594–5600
    #
    years = re.findall(
        r"\b((?:18|19|20)\d{2}[a-z]?)\s*,",
        source,
        flags=re.I,
    )

    if years:
        return years[-1]

    # General fallback
    years = re.findall(
        r"\b((?:18|19|20)\d{2}[a-z]?)\b",
        source,
        flags=re.I,
    )

    if years:
        return years[-1]

    # A webpage with no publication year
    if (
        allow_undated_web
        and looks_like_web_reference(source)
    ):
        return "n.d."

    return None


# ======================================================================
# STANDARD REFERENCES
# ======================================================================

def parse_standard_reference(
    text: str,
) -> tuple[list[str], str] | None:

    # Examples:
    #
    # BS EN 10025-2:2019;
    # ISO 14001:2015;
    # DIN EN ISO 1234:2020;
    #
    match = re.match(
        r"""
        ^\s*
        (?P<standard>
            (?:
                BS\s+EN
                |
                DIN\s+EN\s+ISO
                |
                DIN\s+EN
                |
                EN
                |
                ISO
                |
                IEC
                |
                ASTM
                |
                DIN
                |
                ÖNORM
                |
                ONR
                |
                CEN
                |
                CENELEC
            )
            \s+
            [A-Z0-9./\-]+
            (?:\s+[A-Z0-9./\-]+)*
        )
        :
        (?P<year>(?:18|19|20)\d{2})
        \s*;
        """,
        text,
        flags=re.I | re.X,
    )

    if match:

        standard = normalize_whitespace(
            match.group("standard")
        )

        return [
            standard
        ], match.group("year")

    return None


# ======================================================================
# MDPI AUTHOR LIST
# ======================================================================

def parse_mdpi_authors(
    text: str,
) -> tuple[list[str], str]:

    """
    Parse:

        Koinig, G.; Aberger, J.; Lorber, B.; Kuhn, N. Title...

    ->

        ["Koinig", "Aberger", "Lorber", "Kuhn"]

    """

    authors: list[str] = []
    remaining = text.strip()

    # Author ending in semicolon:
    #
    # Koinig, G.;
    #
    semicolon_author_re = re.compile(
        r"""
        ^\s*
        (?P<surname>[^,;]+?)
        ,\s*
        (?P<initials>
            (?:
                [A-ZÀ-ÖØ-Þ]
                (?:[-A-ZÀ-ÖØ-Þ])?
                \.?
                \s*
            ){1,6}
        )
        ;
        \s*
        """,
        re.X,
    )

    while True:

        match = semicolon_author_re.match(
            remaining
        )

        if not match:
            break

        authors.append(
            normalize_name(
                match.group("surname")
            )
        )

        remaining = remaining[
            match.end():
        ].lstrip()

    # Last author:
    #
    # Kuhn, N. Title...
    #
    final_author_re = re.compile(
        r"""
        ^\s*
        (?P<surname>[^,;]+?)
        ,\s*
        (?P<initials>
            (?:
                [A-ZÀ-ÖØ-Þ]
                (?:[-A-ZÀ-ÖØ-Þ])?
                \.?
                \s*
            ){1,6}
        )
        \.\s+
        (?P<rest>.+)
        $
        """,
        re.X | re.S,
    )

    match = final_author_re.match(
        remaining
    )

    if match:

        authors.append(
            normalize_name(
                match.group("surname")
            )
        )

        remaining = match.group(
            "rest"
        ).strip()

    return authors, remaining


# ======================================================================
# APA-STYLE AUTHOR LIST
# ======================================================================

def parse_apa_inverted_authors(
    author_text: str,
) -> list[str]:

    """
    Parse APA-like lists such as:

        Koinig, G., Aberger, J., Lorber, B.,
        Kuhn, N., Martinelli, W., Radmann, Y.,
        & Tischberger-Aldrian, A.

    ->

        [
            "Koinig",
            "Aberger",
            "Lorber",
            "Kuhn",
            "Martinelli",
            "Radmann",
            "Tischberger-Aldrian",
        ]
    """

    text = normalize_whitespace(
        author_text
    ).strip(" .")

    # Make separators easier to parse.
    text = re.sub(
        r",\s*&\s*",
        ", ",
        text,
    )

    # Find repeated:
    #
    # Surname, G.
    # Surname, G. H.
    # Surname, G.H.
    #
    pattern = re.compile(
        r"""
        (?:
            ^
            |
            ,\s*
        )
        (?P<surname>
            [^,]+?
        )
        ,\s*
        (?P<initials>
            (?:
                [A-ZÀ-ÖØ-Þ]
                (?:[-A-ZÀ-ÖØ-Þ])?
                \.?
                \s*
            ){1,6}
        )
        (?=
            ,\s*
            |
            $
        )
        """,
        re.X,
    )

    authors = []

    for match in pattern.finditer(text):

        surname = normalize_name(
            match.group("surname")
        )

        if surname:
            authors.append(
                surname
            )

    return authors


# ======================================================================
# "HUELLEN M, SCHRADE C ..." STYLE
# ======================================================================

def parse_surname_initial_list(
    author_text: str,
) -> list[str]:

    """
    Parse:

        Huellen M, Schrade C, Wilhelm U, Zulhan Z

    ->

        ["Huellen", "Schrade", "Wilhelm", "Zulhan"]
    """

    chunks = [
        normalize_whitespace(x)
        for x in author_text.split(",")
    ]

    if len(chunks) < 2:
        return []

    authors = []

    pattern = re.compile(
        r"""
        ^
        (?P<surname>.+?)
        \s+
        (?P<initials>
            (?:
                [A-ZÀ-ÖØ-Þ]
                \.?
                \s*
            ){1,6}
        )
        $
        """,
        re.X,
    )

    for chunk in chunks:

        chunk = chunk.strip()

        match = pattern.match(
            chunk
        )

        if not match:
            return []

        authors.append(
            normalize_name(
                match.group("surname")
            )
        )

    return authors


# ======================================================================
# PARENTHESIZED AUTHOR + DATE
# ======================================================================

def parse_parenthetical_author_part(
    author_text: str,
) -> list[str]:

    """
    Handles:

        Koinig, G., Aberger, J., ... & Tischberger-Aldrian, A.
        Huellen M, Schrade C, Wilhelm U, Zulhan Z
        Nils Kroell
        Association, I. I. M.
        Bundesvereinigung ... e.V.
    """

    text = normalize_whitespace(
        author_text
    ).strip(" .")

    # --------------------------------------------------------------
    # 1. APA inverted personal author list
    # --------------------------------------------------------------

    apa_authors = parse_apa_inverted_authors(
        text
    )

    # Require either multiple parsed authors or an exact single-author
    # inverted pattern.
    if len(apa_authors) >= 2:
        return apa_authors

    single_inverted = re.fullmatch(
        r"""
        (?P<surname>[^,]+?)
        ,\s*
        (?:
            [A-ZÀ-ÖØ-Þ]
            \.?
            \s*
        ){1,6}
        """,
        text,
        flags=re.X,
    )

    if single_inverted:
        return [
            normalize_name(
                single_inverted.group("surname")
            )
        ]

    # --------------------------------------------------------------
    # 2. Huellen M, Schrade C, ...
    # --------------------------------------------------------------

    surname_initial_authors = (
        parse_surname_initial_list(text)
    )

    if surname_initial_authors:
        return surname_initial_authors

    # --------------------------------------------------------------
    # 3. Corporate author
    # --------------------------------------------------------------

    if CORPORATE_HINT_RE.search(text):
        return [text]

    # --------------------------------------------------------------
    # 4. Full personal name
    #
    # Nils Kroell -> Kroell
    # --------------------------------------------------------------

    words = text.split()

    if 2 <= len(words) <= 5:

        # Preserve particles belonging to the family name where possible.
        #
        # Nils de Boer -> de Boer
        #
        particles = {
            "de",
            "del",
            "della",
            "der",
            "van",
            "von",
            "da",
            "dos",
            "di",
        }

        if (
            len(words) >= 3
            and words[-2].casefold() in particles
        ):
            return [
                " ".join(words[-2:])
            ]

        return [
            words[-1]
        ]

    # Conservative fallback.
    return [text]


def parse_parenthetical_reference(
    text: str,
) -> tuple[list[str], str] | None:

    match = PAREN_DATE_RE.match(
        text
    )

    if not match:
        return None

    authors = (
        parse_parenthetical_author_part(
            match.group("authors")
        )
    )

    year = normalize_year(
        match.group("year")
    )

    if not authors:
        return None

    return authors, year


# ======================================================================
# INSTITUTIONAL AUTHOR
# ======================================================================

def parse_institutional_reference(
    text: str,
) -> tuple[list[str], str] | None:

    """
    Example:

        Institute for Energy Economics and Financial Analysis.
        Steel Fact Sheet. 2022.
    """

    match = re.match(
        r"""
        ^\s*
        (?P<organization>.+?)
        \.
        \s+
        (?P<rest>.+)
        $
        """,
        text,
        flags=re.X | re.S,
    )

    if not match:
        return None

    organization = normalize_whitespace(
        match.group("organization")
    )

    remaining = match.group(
        "rest"
    )

    year = extract_year(
        remaining,
        allow_undated_web=True,
    )

    if not year:
        return None

    if len(organization) > 250:
        return None

    return [organization], year


# ======================================================================
# MAIN REFERENCE PARSER
# ======================================================================

def parse_reference_metadata(
    text: str,
) -> tuple[list[str], str | None]:

    text = normalize_whitespace(
        text
    )

    # 1. Standards
    standard = parse_standard_reference(
        text
    )

    if standard:
        return standard

    # 2. APA / Harvard / Springer-like "(2026)"
    parenthetical = (
        parse_parenthetical_reference(
            text
        )
    )

    if parenthetical:
        return parenthetical

    # 3. MDPI author list
    authors, remaining = (
        parse_mdpi_authors(
            text
        )
    )

    if authors:

        year = extract_year(
            remaining,
            allow_undated_web=True,
        )

        if year:
            return authors, year

    # 4. Institutional author
    institutional = (
        parse_institutional_reference(
            text
        )
    )

    if institutional:
        return institutional

    return [], None


# ======================================================================
# YEAR SUFFIXES: 2026a, 2026b
# ======================================================================

def assign_year_suffixes(
    refs: dict[int, Reference],
) -> None:

    """
    Exact same author list + exact same year:

        Smith et al., 2026a
        Smith et al., 2026b

    For undated works:

        Smith, n.d.-a
        Smith, n.d.-b
    """

    groups = defaultdict(list)

    for ref in refs.values():

        key = (
            tuple(
                author.casefold()
                for author in ref.authors
            ),
            ref.year.casefold(),
        )

        groups[key].append(ref)

    for group in groups.values():

        group.sort(
            key=lambda r: r.number
        )

        if len(group) == 1:

            group[0].cite_year = (
                group[0].year
            )

            continue

        for index, ref in enumerate(group):

            suffix = chr(
                ord("a") + index
            )

            if ref.year == "n.d.":

                ref.cite_year = (
                    f"n.d.-{suffix}"
                )

            else:

                ref.cite_year = (
                    f"{ref.year}{suffix}"
                )


# ======================================================================
# APA 7 AUTHOR DISAMBIGUATION
# ======================================================================

def full_author_label(
    authors: list[str],
) -> str:

    if len(authors) == 1:
        return authors[0]

    if len(authors) == 2:
        return (
            f"{authors[0]} "
            f"& {authors[1]}"
        )

    return (
        ", ".join(authors[:-1])
        + f", & {authors[-1]}"
    )


def build_author_labels(
    refs: dict[int, Reference],
) -> dict[int, str]:

    """
    Build APA 7 in-text author labels.

    Normal:
        1 author:  Koinig
        2 authors: Koinig & Aberger
        3+:        Koinig et al.

    APA 7 ambiguity handling:

    If these references both exist in 2026:

        Koinig, Aberger, Smith, ...
        Koinig, Müller, Jones, ...

    plain "Koinig et al., 2026" would be ambiguous.

    The script then expands them only as far as necessary:

        Koinig, Aberger, et al., 2026
        Koinig, Müller, et al., 2026
    """

    labels: dict[int, str] = {}

    # Start with normal APA labels.
    for number, ref in refs.items():

        n = len(ref.authors)

        if n == 1:

            labels[number] = (
                ref.authors[0]
            )

        elif n == 2:

            labels[number] = (
                f"{ref.authors[0]} "
                f"& {ref.authors[1]}"
            )

        else:

            labels[number] = (
                f"{ref.authors[0]} "
                f"et al."
            )

    # Find potentially ambiguous 3+ author groups.
    ambiguity_groups = defaultdict(list)

    for number, ref in refs.items():

        if len(ref.authors) < 3:
            continue

        key = (
            ref.authors[0].casefold(),
            ref.year.casefold(),
        )

        ambiguity_groups[key].append(
            number
        )

    for numbers in ambiguity_groups.values():

        # Remove exact duplicate author lists; those are handled by
        # year suffixes rather than expanded names.
        distinct_author_sets = {
            tuple(
                x.casefold()
                for x in refs[number].authors
            )
            for number in numbers
        }

        if len(distinct_author_sets) <= 1:
            continue

        for number in numbers:

            ref = refs[number]

            others = [
                refs[x]
                for x in numbers
                if x != number
                and tuple(
                    a.casefold()
                    for a in refs[x].authors
                )
                != tuple(
                    a.casefold()
                    for a in ref.authors
                )
            ]

            # Find shortest author prefix that distinguishes this work.
            prefix_length = 2

            while (
                prefix_length
                <= len(ref.authors)
            ):

                prefix = tuple(
                    x.casefold()
                    for x in ref.authors[
                        :prefix_length
                    ]
                )

                ambiguous = False

                for other in others:

                    other_prefix = tuple(
                        x.casefold()
                        for x in other.authors[
                            :prefix_length
                        ]
                    )

                    if prefix == other_prefix:
                        ambiguous = True
                        break

                if not ambiguous:
                    break

                prefix_length += 1

            prefix_length = min(
                prefix_length,
                len(ref.authors),
            )

            selected = ref.authors[
                :prefix_length
            ]

            if prefix_length < len(
                ref.authors
            ):

                labels[number] = (
                    ", ".join(selected)
                    + ", et al."
                )

            else:

                labels[number] = (
                    full_author_label(
                        ref.authors
                    )
                )

    return labels


# ======================================================================
# EXPAND NUMERICAL CITATIONS
# ======================================================================

def expand_numbers(
    citation: str,
) -> list[int]:

    numbers = []

    items = re.split(
        r"\s*[,;]\s*",
        citation.strip(),
    )

    for item in items:

        parts = re.split(
            r"\s*[-–—]\s*",
            item,
        )

        if len(parts) == 1:

            numbers.append(
                int(parts[0])
            )

            continue

        if len(parts) == 2:

            start = int(parts[0])
            end = int(parts[1])

            if start <= end:

                numbers.extend(
                    range(
                        start,
                        end + 1,
                    )
                )

            else:

                numbers.extend(
                    range(
                        start,
                        end - 1,
                        -1,
                    )
                )

    return numbers


# ======================================================================
# APA 7 CITATION CONSTRUCTION
# ======================================================================

def year_sort_key(
    year: str,
):

    if year.startswith("n.d."):
        return (9999, year)

    match = re.match(
        r"(\d{4})",
        year,
    )

    if match:
        return (
            int(match.group(1)),
            year,
        )

    return (
        9998,
        year,
    )


def build_apa_citation(
    numbers: list[int],
    refs: dict[int, Reference],
    author_labels: dict[int, str],
) -> str:

    """
    Construct an APA 7 parenthetical citation.

    Examples:

        (Koinig et al., 2026)

        (Koinig et al., 2024, 2026)

        (Koinig et al., 2026; Smith & Jones, 2024)

    Works are alphabetized by author label.
    """

    # Remove duplicate reference numbers.
    unique_numbers = []

    seen_numbers = set()

    for number in numbers:

        if number not in seen_numbers:

            unique_numbers.append(
                number
            )

            seen_numbers.add(
                number
            )

    # Group ONLY references with the exact same author list.
    #
    # This avoids incorrectly merging two different Koinig et al.
    # author groups merely because both abbreviate to "Koinig et al."
    groups = defaultdict(list)

    for number in unique_numbers:

        ref = refs[number]

        key = tuple(
            author.casefold()
            for author in ref.authors
        )

        groups[key].append(
            number
        )

    citation_parts = []

    for numbers_in_group in groups.values():

        # All members have the exact same author list.
        representative = refs[
            numbers_in_group[0]
        ]

        author = author_labels[
            numbers_in_group[0]
        ]

        years = sorted(
            {
                refs[number].cite_year
                for number in numbers_in_group
            },
            key=year_sort_key,
        )

        citation_parts.append(
            (
                author,
                representative.authors,
                f"{author}, "
                + ", ".join(years),
            )
        )

    # APA: alphabetize works inside one parenthetical citation.
    citation_parts.sort(
        key=lambda item: (
            item[0].casefold(),
            tuple(
                x.casefold()
                for x in item[1]
            ),
        )
    )

    return (
        "("
        + "; ".join(
            item[2]
            for item in citation_parts
        )
        + ")"
    )


def make_replacement(
    match: re.Match,
    refs: dict[int, Reference],
    author_labels: dict[int, str],
) -> str:

    original = match.group(0)

    numbers = expand_numbers(
        match.group("inside")
    )

    if not numbers:
        return original

    missing = [
        number
        for number in numbers
        if number not in refs
    ]

    if missing:

        print(
            f"WARNING: Reference(s) {missing} could not be found. "
            f"Leaving {original} unchanged."
        )

        return original

    return build_apa_citation(
        numbers,
        refs,
        author_labels,
    )


# ======================================================================
# WORD RUN REPLACEMENT
# ======================================================================

def replace_citations_in_paragraph(
    paragraph,
    refs: dict[int, Reference],
    author_labels: dict[int, str],
) -> int:

    """
    Replace numerical citations while preserving Word run formatting
    as much as possible.

    This also handles citations split across multiple Word runs.
    """

    runs = paragraph.runs

    if not runs:
        return 0

    full_text = "".join(
        run.text
        for run in runs
    )

    matches = list(
        CITATION_RE.finditer(
            full_text
        )
    )

    if not matches:
        return 0

    count = 0

    # Right-to-left so previous character positions remain valid.
    for match in reversed(matches):

        replacement = make_replacement(
            match,
            refs,
            author_labels,
        )

        if replacement == match.group(0):
            continue

        # Map current runs to text offsets.
        positions = []

        pos = 0

        for index, run in enumerate(runs):

            start = pos
            end = pos + len(run.text)

            positions.append(
                (
                    index,
                    start,
                    end,
                )
            )

            pos = end

        citation_start = match.start()
        citation_end = match.end()

        start_run_index = None
        end_run_index = None

        start_offset = None
        end_offset = None

        for (
            index,
            run_start,
            run_end,
        ) in positions:

            if (
                run_start
                <= citation_start
                < run_end
            ):

                start_run_index = index

                start_offset = (
                    citation_start
                    - run_start
                )

                break

        for (
            index,
            run_start,
            run_end,
        ) in positions:

            if (
                run_start
                < citation_end
                <= run_end
            ):

                end_run_index = index

                end_offset = (
                    citation_end
                    - run_start
                )

                break

        if (
            start_run_index is None
            or end_run_index is None
            or start_offset is None
            or end_offset is None
        ):

            print(
                "WARNING: Could not map citation "
                f"{match.group(0)} back to Word runs."
            )

            continue

        # Citation inside one run
        if (
            start_run_index
            == end_run_index
        ):

            run = runs[
                start_run_index
            ]

            run.text = (
                run.text[
                    :start_offset
                ]
                + replacement
                + run.text[
                    end_offset:
                ]
            )

        # Citation spans multiple runs
        else:

            first_run = runs[
                start_run_index
            ]

            last_run = runs[
                end_run_index
            ]

            prefix = first_run.text[
                :start_offset
            ]

            suffix = last_run.text[
                end_offset:
            ]

            first_run.text = (
                prefix
                + replacement
            )

            for index in range(
                start_run_index + 1,
                end_run_index,
            ):

                runs[index].text = ""

            last_run.text = suffix

        count += 1

    return count


# ======================================================================
# TABLES
# ======================================================================

def iter_table_paragraphs(
    table,
):

    for row in table.rows:

        for cell in row.cells:

            for paragraph in cell.paragraphs:
                yield paragraph

            for nested_table in cell.tables:

                yield from (
                    iter_table_paragraphs(
                        nested_table
                    )
                )


# ======================================================================
# HEADERS / FOOTERS
# ======================================================================

def iter_header_footer_paragraphs(
    doc,
):

    for section in doc.sections:

        for paragraph in (
            section.header.paragraphs
        ):
            yield paragraph

        for table in (
            section.header.tables
        ):

            yield from (
                iter_table_paragraphs(
                    table
                )
            )

        for paragraph in (
            section.footer.paragraphs
        ):
            yield paragraph

        for table in (
            section.footer.tables
        ):

            yield from (
                iter_table_paragraphs(
                    table
                )
            )


# ======================================================================
# FIND REFERENCES
# ======================================================================

def find_references_heading(
    doc,
) -> int:

    for index, paragraph in enumerate(
        doc.paragraphs
    ):

        text = normalize_whitespace(
            paragraph.text
        )

        if re.fullmatch(
            r"references",
            text,
            flags=re.I,
        ):

            return index

    raise RuntimeError(
        "Could not find a paragraph containing only 'References'."
    )


# ======================================================================
# PARSE DOCX REFERENCE LIST
# ======================================================================

def parse_docx_references(
    doc,
    reference_heading_index: int,
) -> dict[int, Reference]:

    refs_raw: dict[int, str] = {}

    automatic_number = 1
    current_number = None

    paragraphs = doc.paragraphs[
        reference_heading_index + 1:
    ]

    for paragraph in paragraphs:

        raw = normalize_whitespace(
            paragraph.text
        )

        if not raw:
            continue

        match = (
            REFERENCE_NUMBER_RE.match(
                raw
            )
        )

        # Explicit:
        #
        # 1. ...
        #
        if match:

            number = int(
                match.group("bracket")
                or match.group("normal")
            )

            text = normalize_whitespace(
                match.group("text")
            )

            refs_raw[number] = text

            current_number = number

            automatic_number = max(
                automatic_number,
                number + 1,
            )

            continue

        # Word automatic numbering
        style_name = (
            paragraph.style.name
            if paragraph.style
            else ""
        )

        ppr = paragraph._p.pPr

        has_word_numbering = (
            ppr is not None
            and ppr.numPr is not None
        )

        is_list_style = (
            "list"
            in style_name.casefold()
        )

        if (
            has_word_numbering
            or is_list_style
        ):

            number = (
                automatic_number
            )

            automatic_number += 1

            refs_raw[number] = raw

            current_number = number

            continue

        # First unnumbered reference
        if current_number is None:

            number = (
                automatic_number
            )

            automatic_number += 1

            refs_raw[number] = raw

            current_number = number

            continue

        # Determine whether this is a new reference or a continuation.
        possible_authors, possible_year = (
            parse_reference_metadata(
                raw
            )
        )

        if (
            possible_authors
            and possible_year
        ):

            number = (
                automatic_number
            )

            automatic_number += 1

            refs_raw[number] = raw

            current_number = number

        else:

            refs_raw[current_number] = (
                refs_raw[current_number]
                + " "
                + raw
            )

    # --------------------------------------------------------------
    # Parse metadata
    # --------------------------------------------------------------

    refs: dict[int, Reference] = {}

    for (
        number,
        text,
    ) in refs_raw.items():

        authors, year = (
            parse_reference_metadata(
                text
            )
        )

        if (
            not authors
            or not year
        ):

            print()
            print(
                f"WARNING: Could not parse reference {number}:"
            )
            print(
                f"  {text[:350]}"
            )
            print()

            continue

        refs[number] = Reference(
            number=number,
            authors=authors,
            year=year,
            cite_year=year,
            raw_text=text,
        )

    assign_year_suffixes(
        refs
    )

    return refs


# ======================================================================
# DOCUMENT CONVERSION
# ======================================================================

def convert_docx(
    input_file: Path,
    output_file: Path,
) -> None:

    print(
        f"Opening: {input_file}"
    )

    doc = Document(
        str(input_file)
    )

    reference_heading_index = (
        find_references_heading(
            doc
        )
    )

    refs = parse_docx_references(
        doc,
        reference_heading_index,
    )

    if not refs:

        raise RuntimeError(
            "No references could be parsed."
        )

    author_labels = (
        build_author_labels(
            refs
        )
    )

    print()
    print(
        f"Parsed {len(refs)} references:"
    )

    print(
        "-" * 90
    )

    for number in sorted(refs):

        ref = refs[number]

        label = author_labels[
            number
        ]

        print(
            f"{number:>4}: "
            f"{label}, "
            f"{ref.cite_year}"
        )

    print(
        "-" * 90
    )

    print()

    total_replacements = 0

    # --------------------------------------------------------------
    # Main manuscript body only
    # --------------------------------------------------------------

    for paragraph in doc.paragraphs[
        :reference_heading_index
    ]:

        total_replacements += (
            replace_citations_in_paragraph(
                paragraph,
                refs,
                author_labels,
            )
        )

    # --------------------------------------------------------------
    # Tables
    # --------------------------------------------------------------

    for table in doc.tables:

        for paragraph in (
            iter_table_paragraphs(
                table
            )
        ):

            total_replacements += (
                replace_citations_in_paragraph(
                    paragraph,
                    refs,
                    author_labels,
                )
            )

    # --------------------------------------------------------------
    # Headers / footers
    # --------------------------------------------------------------

    for paragraph in (
        iter_header_footer_paragraphs(
            doc
        )
    ):

        total_replacements += (
            replace_citations_in_paragraph(
                paragraph,
                refs,
                author_labels,
            )
        )

    # Save
    doc.save(
        str(output_file)
    )

    print()
    print(
        f"Converted {total_replacements} numerical citation(s)."
    )

    print(
        "Saved converted manuscript as:"
    )

    print(
        f"  {output_file}"
    )


# ======================================================================
# CLI
# ======================================================================

def main() -> None:

    parser = argparse.ArgumentParser(
        description=(
            "Convert numbered citations in a DOCX manuscript "
            "to APA 7 author-year citations."
        )
    )

    parser.add_argument(
        "input",
        help="Input DOCX manuscript",
    )

    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Output DOCX manuscript",
    )

    parser.add_argument(
        "--backup",
        action="store_true",
        help=(
            "Create a .bak copy of the input manuscript."
        ),
    )

    args = parser.parse_args()

    input_path = Path(
        args.input
    )

    output_path = Path(
        args.output
    )

    if not input_path.exists():

        raise FileNotFoundError(
            f"Input file does not exist: {input_path}"
        )

    if (
        input_path.suffix.lower()
        != ".docx"
    ):

        raise ValueError(
            "Input file must be a .docx file."
        )

    if (
        output_path.suffix.lower()
        != ".docx"
    ):

        raise ValueError(
            "Output file must be a .docx file."
        )

    if (
        input_path.resolve()
        == output_path.resolve()
    ):

        raise ValueError(
            "Input and output must have different filenames.\n"
            "Example:\n"
            "  input.docx -o input_apa7.docx"
        )

    if args.backup:

        backup_path = (
            input_path.with_suffix(
                input_path.suffix
                + ".bak"
            )
        )

        shutil.copy2(
            input_path,
            backup_path,
        )

        print(
            f"Backup created: {backup_path}"
        )

    convert_docx(
        input_path,
        output_path,
    )


if __name__ == "__main__":
    main()
#```
