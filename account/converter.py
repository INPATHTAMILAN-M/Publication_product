import re
from typing import Dict, List, Any
from jinja2 import Template
from io import TextIOWrapper
from django.core.exceptions import ObjectDoesNotExist
from .models import LatexDocument

class LatexConverter:
    def __init__(self, template_key, template_content=None):
        self.template_key = template_key
        self.template_text = template_content or self.default_template()

    def convert(self, extracted_data: dict) -> str:
        filled_data = self.prepare_template_data(extracted_data)
        template = Template(self.template_text)
        return template.render(**filled_data)

    def prepare_template_data(self, data: dict) -> dict:
        metadata = data.get("metadata", {})
        body = data.get("body", [])
        references = data.get("references", [])
        tables = data.get("tables", [])

        return {
            "title": metadata.get("title", "Untitled"),
            "short_title": metadata.get("title", "Untitled")[:30],
            "abstract": metadata.get("abstract", ""),
            "keywords": ', '.join(metadata.get("keywords", [])),
            "author_block": self.format_authors(metadata.get("authors", [])),
            "credit_roles": self.format_roles(metadata.get("authors", [])),
            "content": self.format_body(body),
            "bibliography": self.format_references(references),
            "non_technical_summary": "",
            "acknowledgements": "",
            "data_availability": "",
            "competing_interests": "",
            "methods_content": "",
            "doi": "",
            "editor_name": "",
            "received_date": "",
            "accepted_date": "",
            "published_date": "",
            "year": "2025",
            "volume": "1",
            "paper_number": "001",
            "bib_file": "references.bib"
        }

    def format_authors(self, authors):
        if not authors:
            return "Unknown Author"
        author_lines = []
        address_lines = []
        affil_map = {}
        for idx, author in enumerate(authors):
            affil = author.get('affiliation', 'Unknown Institute')
            affil_key = affil_map.setdefault(affil, len(affil_map) + 1)
            author_lines.append(f"\\author[inst{affil_key}]{{{author.get('name', 'Unknown')}}}")
        for affil, num in affil_map.items():
            address_lines.append(f"\\address[inst{num}]{{{affil}}}")
        return '\n'.join(author_lines + address_lines)
    
    def format_roles(self, authors):
        roles = []
        for author in authors:
            if author.get("role"):
                roles.append(f"{author['name']} ({author['role']})")
        return '\\\\'.join(roles)

    def format_body(self, body_sections):
        latex = []
        for section in body_sections:
            if section.get("type") == "section":
                latex.append(f"\\section{{{section.get('heading','')}}}")
                for item in section.get("content", []):
                    latex.append(self.format_content_item(item))
        return '\n\n'.join(latex)

    def format_content_item(self, item):
        if item.get("type") == "paragraph":
            return self.clean_latex_text(item.get("text", ""))
        elif item.get("type") == "figure":
            return self.format_figure(item)
        elif item.get("type") == "list":
            return self.format_list(item)
        elif item.get("type") == "table":
            return self.format_table(item)
        elif item.get("type") == "subsection":
            return f"\\subsection{{{item.get('heading','')}}}\n\n" + \
                   '\n\n'.join([self.format_content_item(subitem) for subitem in item.get("content", [])])
        return ""

    def clean_latex_text(self, text):
        # Escape special LaTeX characters
        text = text.replace('&', '\\&')
        text = text.replace('%', '\\%')
        text = text.replace('$', '\\$')
        text = text.replace('#', '\\#')
        text = text.replace('_', '\\_')
        text = text.replace('{', '\\{')
        text = text.replace('}', '\\}')
        text = text.replace('~', '\\textasciitilde')
        text = text.replace('^', '\\textasciicircum')
        text = text.replace('\\', '\\textbackslash')
        return text

    def format_figure(self, figure):
        caption = self.clean_latex_text(figure.get("caption", ""))
        label = figure.get("label", "").replace(" ", "")
        path = figure.get("content", "").replace("\\", "/")  # Fix path separators
        
        return (
            "\\begin{figure}[H]\n"
            "\\centering\n"
            f"\\includegraphics[width=0.8\\textwidth]{{{path}}}\n"
            f"\\caption{{{caption}}}\n"
            f"\\label{{{label}}}\n"
            "\\end{figure}"
        )

    def format_list(self, item):
        items = '\n'.join([f"  \\item {self.clean_latex_text(i)}" for i in item.get("items",[])])
        return "\\begin{itemize}\n" + items + "\n\\end{itemize}"

    def format_table(self, table):
        headers = [self.clean_latex_text(h) for h in table.get("header", [])]
        rows = [[self.clean_latex_text(cell) for cell in row] for row in table.get("rows", [])]
        col_format = " | ".join(["l"] * len(headers))
        latex_rows = [' & '.join(row) + " \\\\" for row in rows]
        
        return (
            "\\begin{table}[H]\n\\centering\n"
            f"\\caption{{{table.get('label','')}}}\n"
            f"\\label{{tab:{table.get('label','').lower()}}}\n"
            f"\\begin{{tabular}}{{|{col_format}|}}\n\\hline\n"
            + ' & '.join(headers) + " \\\\\n\\hline\n"
            + '\n'.join(latex_rows) + "\n\\hline\n"
            "\\end{tabular}\n\\end{table}"
        )

    def format_references(self, refs):
        if not refs:
            return ""
        # Remove duplicates by citation text
        unique_refs = []
        seen = set()
        for r in refs:
            citation = r['citation']
            if citation not in seen:
                seen.add(citation)
                unique_refs.append(r)
        return '\n'.join([f"\\bibitem{{ref{r['id']}}} {self.clean_latex_text(r['citation'])}" for r in unique_refs])

    def default_template(self):
        return r"""
\documentclass{llncs}

% PACKAGES
\usepackage[utf8]{inputenc}
\usepackage{amsmath,amsfonts,amssymb}
\usepackage{graphicx}
\usepackage{cite}
\usepackage{hyperref}
\usepackage{url}
\usepackage{float}
\usepackage{caption}
\usepackage{booktabs}

% TITLE
\title{ {{ title }} }

% AUTHORS BLOCK
\author{ {{ author_block | safe }} }

\begin{document}

\maketitle

% ABSTRACT
\begin{abstract}
{{ abstract }}
\end{abstract}

% KEYWORDS
{% if keywords %}
\keywords{ {{ keywords }} }
{% endif %}

% MAIN CONTENT
{{ content | safe }}

% REFERENCES
\bibliographystyle{splncs04}
\begin{thebibliography}{99}
{{ bibliography | safe }}
\end{thebibliography}

\end{document}
        """

def text_to_latex(extracted_data: dict) -> str:
    try:
        active_template = LatexDocument.objects.get(is_active=True)
        f = active_template.tex_file.open()
        with TextIOWrapper(f, encoding='utf-8') as text_file:
            template_content = text_file.read()
    except LatexDocument.DoesNotExist:
        template_content = None

    if template_content:
        converter = LatexConverter(
            template_key=active_template.title.lower(),
            template_content=template_content
        )
    else:
        converter = LatexConverter(template_key='default')

    return converter.convert(extracted_data)