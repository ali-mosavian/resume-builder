"""WordprocessingML snippets shared by the .docx builder and the template generators."""

import re

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.xmlchemy import BaseOxmlElement

CM = 567  # twips per cm


def xml(snippet: str, *prefixes: str) -> BaseOxmlElement:
    """Parse XML written without namespace declarations; they're added to its root (w: unless named)."""
    return parse_xml(re.sub(r"^<([\w:]+)", rf"<\1 {nsdecls(*(prefixes or ('w',)))}", snippet, count=1))


def shading(fill: str) -> BaseOxmlElement:
    return xml(f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>')
