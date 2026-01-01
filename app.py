"""
Streamlit app for converting pasted rich text (Word or Google Docs) into Moodle-friendly HTML.

Requirements:
  pip install streamlit beautifulsoup4 streamlit-quill

Run:
  streamlit run app.py
"""

import re
from dataclasses import dataclass
from typing import Iterable

import streamlit as st
from bs4 import BeautifulSoup, Comment, NavigableString
from streamlit_quill import st_quill


@dataclass
class CleanOptions:
    remove_inline_styles: bool = True
    remove_classes: bool = True
    remove_ids: bool = True
    unwrap_spans: bool = True
    remove_empty_tags: bool = True
    remove_comments: bool = True
    remove_office_tags: bool = True
    collapse_whitespace: bool = True
    convert_b_i_to_strong_em: bool = True
    keep_only_moodle_safe_attributes: bool = True


VOID_TAGS = {
    "area",
    "base",
    "br",
    "col",
    "embed",
    "hr",
    "img",
    "input",
    "link",
    "meta",
    "param",
    "source",
    "track",
    "wbr",
}

OFFICE_TAGS = {
    "o:p",
    "v:shapetype",
    "v:shape",
    "v:imagedata",
    "xml",
    "style",
}

# Keep a small allowlist of attributes that Moodle typically tolerates.
# You can extend this list based on your institution’s Moodle config and content needs.
GLOBAL_ALLOWED_ATTRS = set()

ALLOWED_ATTRS_BY_TAG = {
    "a": {"href", "title"},
    "img": {"src", "alt", "title", "width", "height"},
    "th": {"colspan", "rowspan", "scope"},
    "td": {"colspan", "rowspan"},
    "table": {"summary"},
}


def _safe_collapse_whitespace_in_text_nodes(soup: BeautifulSoup) -> None:
    """
    Collapses repeated whitespace in text nodes only.
    This avoids collapsing important whitespace between HTML tags.
    """
    for text_node in soup.find_all(string=True):
        if isinstance(text_node, Comment):
            continue
        new_text = re.sub(r"[ \t\r\n]+", " ", str(text_node))
        if new_text != str(text_node):
            text_node.replace_with(new_text)


def _remove_office_specific_content(soup: BeautifulSoup) -> None:
    """Strip common Microsoft Office-specific tags and attributes."""
    for office_tag in OFFICE_TAGS:
        for node in soup.find_all(office_tag):
            node.decompose()

    for tag in soup.find_all(True):
        for attribute in list(tag.attrs):
            if attribute.lower().startswith("mso") or attribute.lower().startswith("o:"):
                del tag.attrs[attribute]


def _convert_b_i_to_strong_em(soup: BeautifulSoup) -> None:
    """Normalize presentational tags to semantic equivalents."""
    for b in soup.find_all("b"):
        b.name = "strong"
    for i in soup.find_all("i"):
        i.name = "em"


def _strip_disallowed_attributes(soup: BeautifulSoup) -> None:
    """Keep only Moodle-safe attributes (small allowlist)."""
    for tag in soup.find_all(True):
        allowed = set(GLOBAL_ALLOWED_ATTRS)
        allowed |= ALLOWED_ATTRS_BY_TAG.get(tag.name, set())
        for attr in list(tag.attrs):
            if attr not in allowed:
                del tag.attrs[attr]


def _tag_children_are_only_breaks_and_whitespace(tag) -> bool:
    """
    Treat tags like <p><br></p> as empty.
    Quill commonly produces paragraphs that contain only <br>.
    """
    for child in tag.contents:
        if isinstance(child, NavigableString):
            if str(child).strip():
                return False
        else:
            if child.name != "br":
                return False
    return True


def _remove_empty_tags(soup: BeautifulSoup) -> None:
    """
    Remove tags with no meaningful text or children.
    Repeats until stable so newly emptied parents are removed too.
    """
    removed = True
    while removed:
        removed = False
        for tag in list(soup.find_all(True)):
            if tag.name in VOID_TAGS:
                continue

            text = tag.get_text(strip=True)

            # If it has real text, keep it
            if text:
                continue

            # If it has element children other than <br>, keep it
            element_children = [c for c in tag.contents if getattr(c, "name", None)]
            if element_children:
                # If the only element children are <br>, treat as empty
                if _tag_children_are_only_breaks_and_whitespace(tag):
                    tag.decompose()
                    removed = True
                continue

            # No text and no element children
            tag.decompose()
            removed = True


def clean_html(raw_html: str, options: CleanOptions) -> str:
    """Clean HTML produced from rich paste for Moodle use."""
    soup = BeautifulSoup(raw_html or "", "html.parser")

    # If Quill returns wrappers, prefer just the body contents later
    # but still clean the full soup first.
    if options.remove_comments:
        for comment in soup.find_all(string=lambda t: isinstance(t, Comment)):
            comment.extract()

    if options.remove_office_tags:
        _remove_office_specific_content(soup)

    if options.remove_inline_styles:
        for tag in soup.find_all(style=True):
            del tag["style"]

    if options.remove_classes:
        for tag in soup.find_all(class_=True):
            del tag["class"]

    if options.remove_ids:
        for tag in soup.find_all(id=True):
            del tag["id"]

    if options.unwrap_spans:
        for span in soup.find_all("span"):
            span.unwrap()

    if options.convert_b_i_to_strong_em:
        _convert_b_i_to_strong_em(soup)

    if options.keep_only_moodle_safe_attributes:
        _strip_disallowed_attributes(soup)

    if options.collapse_whitespace:
        _safe_collapse_whitespace_in_text_nodes(soup)

    if options.remove_empty_tags:
        _remove_empty_tags(soup)

    # Output only body contents when present, to avoid <html><body> wrappers
    if soup.body:
        cleaned = soup.body.decode_contents()
    else:
        cleaned = soup.decode()

    return cleaned.strip()


def render_sidebar() -> CleanOptions:
    st.sidebar.header("Cleaning options")

    return CleanOptions(
        remove_inline_styles=st.sidebar.checkbox(
            "Remove inline styles", value=True, help="Strips style attributes (best for Moodle consistency)."
        ),
        remove_classes=st.sidebar.checkbox(
            "Remove CSS classes", value=True, help="Strips class attributes."
        ),
        remove_ids=st.sidebar.checkbox(
            "Remove element IDs", value=True, help="Strips id attributes."
        ),
        unwrap_spans=st.sidebar.checkbox(
            "Unwrap span tags", value=True, help="Removes <span> while keeping the text inside."
        ),
        remove_empty_tags=st.sidebar.checkbox(
            "Remove empty tags", value=True, help="Removes empty tags, including <p><br></p>."
        ),
        remove_comments=st.sidebar.checkbox(
            "Remove comments", value=True, help="Removes HTML comments."
        ),
        remove_office_tags=st.sidebar.checkbox(
            "Strip Office-specific markup", value=True, help="Removes common Word-generated tags and attributes."
        ),
        collapse_whitespace=st.sidebar.checkbox(
            "Collapse whitespace", value=True, help="Normalizes extra spaces and line breaks in text."
        ),
        convert_b_i_to_strong_em=st.sidebar.checkbox(
            "Convert b and i to strong and em", value=True, help="Makes emphasis more semantic and consistent."
        ),
        keep_only_moodle_safe_attributes=st.sidebar.checkbox(
            "Keep only Moodle-safe attributes",
            value=True,
            help="Keeps a small allowlist like href and src. Helps Moodle content stay clean and predictable.",
        ),
    )


def render_app() -> None:
    st.set_page_config(page_title="Word to Moodle HTML Cleaner", layout="wide")
    st.title("Word to Moodle HTML Cleaner")
    st.caption("Paste rich text from Word or Google Docs on the left. Copy clean Moodle-friendly HTML on the right.")

    options = render_sidebar()

    left, right = st.columns(2)

    with left:
        st.subheader("Rich text input")
        st.caption("Paste from Word or Google Docs here. Formatting is preserved on paste.")
        raw_html = st_quill(
            placeholder="Paste here...",
            html=True,
            key="rich_paste_editor",
        )
        st.info(
            "Tip: If you see tables or complex layouts paste poorly, that usually needs a DOCX upload converter. "
            "This tool is focused on pasted rich text."
        )

    cleaned_html = clean_html(raw_html or "", options)
    output_html = cleaned_html or "<p>Cleaned HTML will appear here after you paste content.</p>"

    with right:
        st.subheader("Clean HTML for Moodle")
        st.text_area(
            "Copy from here",
            value=output_html,
            height=260,
            help="Click in the box, then select all and copy into Moodle.",
        )
        st.code(output_html, language="html")
        st.download_button(
            "Download cleaned HTML",
            data=cleaned_html if cleaned_html else "",
            file_name="moodle_cleaned.html",
            mime="text/html",
            disabled=not bool(cleaned_html.strip()),
        )

    st.subheader("Preview")
    st.caption("Rendered preview of the cleaned HTML.")
    st.components.v1.html(output_html, height=300, scrolling=True)


if __name__ == "__main__":
    render_app()
