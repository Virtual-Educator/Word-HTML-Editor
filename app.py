"""
Streamlit app for converting pasted rich text into Moodle-friendly HTML.

Requirements:
  pip install streamlit st-tiny-editor beautifulsoup4 html5lib

Run:
  streamlit run app.py
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

import streamlit as st
from bs4 import BeautifulSoup, Comment, NavigableString, Tag
from st_tiny_editor import tiny_editor


@dataclass
class CleanOptions:
    remove_comments: bool = True
    remove_classes: bool = True
    remove_ids: bool = True
    unwrap_spans: bool = True
    remove_office_tags: bool = True
    remove_empty_tags: bool = True

    infer_headings_from_font_size: bool = True
    keep_only_text_align_style: bool = True
    keep_only_moodle_safe_attributes: bool = True

    pretty_print_html: bool = True
    show_compact_copy: bool = True


VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input",
    "link", "meta", "param", "source", "track", "wbr",
}

OFFICE_TAGS = {
    "o:p", "v:shapetype", "v:shape", "v:imagedata", "xml", "style",
}

GLOBAL_ALLOWED_ATTRS = set()
ALLOWED_ATTRS_BY_TAG = {
    "a": {"href", "title", "target", "rel"},
    "img": {"src", "alt", "title", "width", "height"},
    "th": {"colspan", "rowspan", "scope"},
    "td": {"colspan", "rowspan"},
    "table": {"summary"},
}

ALLOWED_TAGS = {
    "p", "br",
    "h1", "h2", "h3", "h4", "h5", "h6",
    "ul", "ol", "li",
    "strong", "em", "b", "i", "u",
    "a",
    "blockquote",
    "pre", "code",
    "hr",
    "table", "thead", "tbody", "tr", "th", "td",
    "img",
    "span",
    "div",
}


def _wrap_fragment(raw_html: str) -> BeautifulSoup:
    wrapped = f"<div id='__root__'>{raw_html or ''}</div>"
    return BeautifulSoup(wrapped, "html5lib")


def _root(soup: BeautifulSoup) -> Tag:
    found = soup.find("div", {"id": "__root__"})
    return found if isinstance(found, Tag) else soup


def _parse_style(style_value: str) -> Dict[str, str]:
    styles: Dict[str, str] = {}
    if not style_value:
        return styles
    for part in style_value.split(";"):
        part = part.strip()
        if not part or ":" not in part:
            continue
        prop, val = part.split(":", 1)
        prop = prop.strip().lower()
        val = val.strip()
        if prop:
            styles[prop] = val
    return styles


def _font_size_to_pt(value: str) -> Optional[float]:
    if not value:
        return None
    m = re.match(r"^\s*([0-9]*\.?[0-9]+)\s*(px|pt)\s*$", value.strip().lower())
    if not m:
        return None
    num = float(m.group(1))
    unit = m.group(2)
    if unit == "pt":
        return num
    return num * 0.75  # px -> pt approx


def _candidate_font_size_pt(tag: Tag) -> Optional[float]:
    style = _parse_style(tag.get("style", ""))
    fs = _font_size_to_pt(style.get("font-size", ""))
    if fs is not None:
        return fs

    first_span = tag.find("span", recursive=False)
    if isinstance(first_span, Tag):
        span_style = _parse_style(first_span.get("style", ""))
        fs = _font_size_to_pt(span_style.get("font-size", ""))
        if fs is not None:
            return fs

    return None


def _infer_headings_from_styles(root: Tag) -> None:
    thresholds: Tuple[Tuple[float, str], ...] = (
        (22.0, "h2"),
        (18.0, "h3"),
        (16.0, "h4"),
    )

    for tag in list(root.find_all(["p", "div"])):
        if tag.find_parent("li") is not None:
            continue

        text = tag.get_text(strip=True)
        if not text:
            continue

        if len(text) > 120:
            continue

        fs_pt = _candidate_font_size_pt(tag)
        if fs_pt is None:
            continue

        for threshold, heading_tag in thresholds:
            if fs_pt >= threshold:
                tag.name = heading_tag
                break


def _remove_office_specific_content(root: Tag) -> None:
    for office_tag in OFFICE_TAGS:
        for node in root.find_all(office_tag):
            node.decompose()

    for tag in root.find_all(True):
        for attribute in list(tag.attrs):
            attr_l = attribute.lower()
            if attr_l.startswith("mso") or attr_l.startswith("o:"):
                del tag.attrs[attribute]


def _strip_disallowed_tags(root: Tag) -> None:
    for tag in list(root.find_all(True)):
        if tag.name not in ALLOWED_TAGS:
            tag.unwrap()


def _strip_disallowed_attributes(root: Tag) -> None:
    for tag in root.find_all(True):
        allowed = set(GLOBAL_ALLOWED_ATTRS)
        allowed |= ALLOWED_ATTRS_BY_TAG.get(tag.name, set())

        for attr in list(tag.attrs):
            if attr == "style":
                continue
            if attr not in allowed:
                del tag.attrs[attr]


def _keep_only_text_align_style(root: Tag) -> None:
    for tag in root.find_all(True):
        if "style" not in tag.attrs:
            continue
        style_map = _parse_style(tag.get("style", ""))
        align = style_map.get("text-align", "").strip().lower()

        if align in {"left", "right", "center", "justify"}:
            tag["style"] = f"text-align: {align}"
        else:
            del tag.attrs["style"]


def _remove_all_styles(root: Tag) -> None:
    for tag in root.find_all(style=True):
        del tag.attrs["style"]


def _unwrap_spans(root: Tag) -> None:
    for span in root.find_all("span"):
        span.unwrap()


def _convert_b_i_to_strong_em(root: Tag) -> None:
    for b in root.find_all("b"):
        b.name = "strong"
    for i in root.find_all("i"):
        i.name = "em"


def _tag_is_effectively_empty(tag: Tag) -> bool:
    if tag.name in VOID_TAGS:
        return False

    if tag.get_text(strip=True):
        return False

    for child in tag.contents:
        if isinstance(child, NavigableString):
            if str(child).strip():
                return False
        else:
            if getattr(child, "name", None) != "br":
                return False

    return True


def _remove_empty_tags(root: Tag) -> None:
    removed = True
    while removed:
        removed = False
        for tag in list(root.find_all(True)):
            if tag.name in VOID_TAGS:
                continue
            if _tag_is_effectively_empty(tag):
                tag.decompose()
                removed = True


def _pretty_html(fragment_html: str) -> str:
    soup = BeautifulSoup(fragment_html or "", "html5lib")
    body = soup.body
    content = body.decode_contents() if body else soup.decode()

    pretty = BeautifulSoup(content, "html.parser").prettify()
    pretty = re.sub(r"\n\s*\n+", "\n\n", pretty).strip()
    return pretty


def _compact_html(fragment_html: str) -> str:
    """
    Compact copy version:
    - Keeps line breaks between block elements
    - Removes indentation and repeated spaces
    - Avoids turning everything into one long line
    """
    html = fragment_html.strip()
    html = re.sub(r"[ \t]+\n", "\n", html)
    html = re.sub(r"\n[ \t]+", "\n", html)
    html = re.sub(r"[ \t]{2,}", " ", html)
    return html.strip()


def clean_html(raw_html: str, options: CleanOptions) -> str:
    soup = _wrap_fragment(raw_html)
    root = _root(soup)

    if options.remove_comments:
        for comment in root.find_all(string=lambda t: isinstance(t, Comment)):
            comment.extract()

    if options.remove_office_tags:
        _remove_office_specific_content(root)

    _strip_disallowed_tags(root)

    if options.infer_headings_from_font_size:
        _infer_headings_from_styles(root)

    if options.keep_only_text_align_style:
        _keep_only_text_align_style(root)
    else:
        _remove_all_styles(root)

    if options.keep_only_moodle_safe_attributes:
        _strip_disallowed_attributes(root)

    if options.remove_classes:
        for tag in root.find_all(class_=True):
            del tag.attrs["class"]

    if options.remove_ids:
        for tag in root.find_all(id=True):
            del tag.attrs["id"]

    _convert_b_i_to_strong_em(root)

    if options.unwrap_spans:
        _unwrap_spans(root)

    if options.remove_empty_tags:
        _remove_empty_tags(root)

    cleaned = root.decode_contents().strip()

    if options.pretty_print_html:
        cleaned = _pretty_html(cleaned)

    return cleaned


def render_sidebar() -> CleanOptions:
    st.sidebar.header("Cleaning options")

    return CleanOptions(
        remove_comments=st.sidebar.checkbox("Remove comments", value=True),
        remove_classes=st.sidebar.checkbox("Remove CSS classes", value=True),
        remove_ids=st.sidebar.checkbox("Remove element IDs", value=True),
        unwrap_spans=st.sidebar.checkbox("Unwrap span tags", value=True),
        remove_office_tags=st.sidebar.checkbox("Strip Office-specific markup", value=True),
        remove_empty_tags=st.sidebar.checkbox("Remove empty tags", value=True),
        infer_headings_from_font_size=st.sidebar.checkbox(
            "Convert large font paragraphs to headings",
            value=True,
        ),
        keep_only_text_align_style=st.sidebar.checkbox(
            "Keep only text-align styling",
            value=True,
        ),
        keep_only_moodle_safe_attributes=st.sidebar.checkbox(
            "Keep only Moodle-safe attributes",
            value=True,
        ),
        pretty_print_html=st.sidebar.checkbox(
            "Pretty print HTML output",
            value=True,
        ),
        show_compact_copy=st.sidebar.checkbox(
            "Show compact copy version",
            value=True,
        ),
    )


def _get_tinymce_api_key() -> str:
    """
    Avoid StreamlitSecretNotFoundError when no secrets.toml exists.
    Returns "" if no key is available.
    """
    try:
        return str(st.secrets.get("TINY_API_KEY", ""))
    except Exception:
        return ""


def render_app() -> None:
    st.set_page_config(page_title="Word to Moodle HTML Cleaner", layout="wide")
    st.title("Word to Moodle HTML Cleaner")
    st.caption("Paste rich text on the left. Copy Moodle-friendly HTML on the right.")

    options = render_sidebar()
    left, right = st.columns(2)

    with left:
        st.subheader("Rich text input")

        api_key = _get_tinymce_api_key()

        raw_html = tiny_editor(
            apiKey=api_key,
            height=520,
            initialValue="",
            menubar=False,
            plugins=["lists", "link", "table", "paste", "code", "autolink"],
            toolbar=(
                "undo redo | blocks | bold italic underline | "
                "alignleft aligncenter alignright alignjustify | "
                "bullist numlist | link table | removeformat | code"
            ),
        )

        st.info(
            "Moodle may sanitize HTML differently in Atto vs TinyMCE depending on filters. "
            "After copying, check the page in Moodle and confirm headings and lists remain correct."
        )

        if not api_key:
            st.caption(
                "TinyMCE API key not found. If you want to add one, create .streamlit/secrets.toml "
                "in your project folder and add: TINY_API_KEY=\"your_key\""
            )

    cleaned_html = clean_html(raw_html or "", options) if raw_html else ""
    output_html = cleaned_html or "<p>Cleaned HTML will appear here after you paste content.</p>"

    with right:
        st.subheader("Clean HTML for Moodle")

        st.text_area(
            "Readable HTML (review/edit here)",
            value=output_html,
            height=240,
        )
        st.code(output_html, language="html")

        if options.show_compact_copy and cleaned_html:
            st.text_area(
                "Compact copy version (less whitespace)",
                value=_compact_html(output_html),
                height=140,
            )

        st.download_button(
            "Download cleaned HTML",
            data=cleaned_html if cleaned_html else "",
            file_name="moodle_cleaned.html",
            mime="text/html",
            disabled=not bool(cleaned_html.strip()),
        )

    st.subheader("Preview")
    st.components.v1.html(output_html, height=320, scrolling=True)


if __name__ == "__main__":
    render_app()
