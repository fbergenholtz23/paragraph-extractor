"""
redline_reader.py — reads a .docx file and prints sections that contain
tracked changes (redlines), showing the original text and the suggested text,
along with any comments attached to them.

Usage:
    python redline_reader.py <file.docx>
"""

import sys
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _tag(local):
    return f"{{{W}}}{local}"


def has_tracked_changes(para_elem):
    """Return True if the paragraph contains any insertion or deletion marks."""
    return (
        para_elem.find(f".//{_tag('ins')}") is not None
        or para_elem.find(f".//{_tag('del')}") is not None
    )


def get_text_variants(para_elem):
    """
    Returns (original, suggested) text for a paragraph.
      original  = text as it was before changes (deletions in, insertions out)
      suggested = text as it would be after changes (insertions in, deletions out)
    """
    original = []
    suggested = []

    def walk(elem, in_ins, in_del):
        tag = elem.tag

        if tag == _tag("ins"):
            for child in elem:
                walk(child, True, in_del)
            return

        if tag == _tag("del"):
            for child in elem:
                walk(child, in_ins, True)
            return

        if tag == _tag("t"):
            text = elem.text or ""
            if not in_ins:
                original.append(text)
            if not in_del:
                suggested.append(text)
            return

        if tag == _tag("delText"):
            original.append(elem.text or "")
            return

        for child in elem:
            walk(child, in_ins, in_del)

    walk(para_elem, False, False)
    return "".join(original), "".join(suggested)


def get_comment_ids(para_elem):
    """Return the set of comment IDs referenced inside this paragraph."""
    ids = set()
    for elem in para_elem.findall(f".//{_tag('commentRangeStart')}"):
        cid = elem.get(_tag("id"))
        if cid is not None:
            ids.add(cid)
    for elem in para_elem.findall(f".//{_tag('commentReference')}"):
        cid = elem.get(_tag("id"))
        if cid is not None:
            ids.add(cid)
    return ids


def load_comments(docx_path):
    """Parse word/comments.xml and return a dict of id -> {author, text}."""
    comments = {}
    with zipfile.ZipFile(docx_path) as z:
        if "word/comments.xml" not in z.namelist():
            return comments
        with z.open("word/comments.xml") as f:
            root = ET.parse(f).getroot()
    for comment in root.findall(_tag("comment")):
        cid = comment.get(_tag("id"))
        author = comment.get(_tag("author"), "Unknown")
        text = "".join(t.text or "" for t in comment.iter(_tag("t")))
        comments[cid] = {"author": author, "text": text}
    return comments


def load_document_paragraphs(docx_path):
    """Return a list of paragraph XML elements from word/document.xml."""
    with zipfile.ZipFile(docx_path) as z:
        with z.open("word/document.xml") as f:
            root = ET.parse(f).getroot()
    return root.findall(f".//{_tag('p')}")


def get_para_style(para_elem):
    pPr = para_elem.find(f"{{{W}}}pPr")
    if pPr is not None:
        pStyle = pPr.find(f"{{{W}}}pStyle")
        if pStyle is not None:
            return pStyle.get(f"{{{W}}}val")
    return None


def is_clause_start(para_elem):
    """
    Return True if this paragraph begins a new numbered clause.
    sapcontractsectionlev1 always starts a new section.
    sapcontractsectionlev2 only starts a new clause when it carries a real
    list number (numId != 0), distinguishing it from unnumbered continuation
    paragraphs that share the same style.
    """
    style = get_para_style(para_elem)
    if style == "sapcontractsectionlev1":
        return True
    if style == "sapcontractsectionlev2":
        pPr = para_elem.find(f"{{{W}}}pPr")
        if pPr is not None:
            numPr = pPr.find(f"{{{W}}}numPr")
            if numPr is not None:
                numIdEl = numPr.find(f"{{{W}}}numId")
                if numIdEl is not None:
                    num_id = numIdEl.get(f"{{{W}}}val")
                    return num_id is not None and num_id != "0"
    return False


def get_plain_text(para_elem):
    """Return the plain text of a paragraph (no change tracking applied)."""
    parts = []
    for t in para_elem.iter(f"{{{W}}}t"):
        parts.append(t.text or "")
    for t in para_elem.iter(f"{{{W}}}delText"):
        parts.append(t.text or "")
    return "".join(parts)


def group_clauses(paragraphs):
    """
    Split the document into logical clauses, where each clause starts at a
    sapcontractsectionlev1/lev2 paragraph. Returns a list of (start, end)
    index tuples (inclusive).
    """
    boundaries = [i for i, p in enumerate(paragraphs) if is_clause_start(p)]
    if not boundaries:
        return [(0, len(paragraphs) - 1)]
    clauses = []
    for j, start in enumerate(boundaries):
        end = boundaries[j + 1] - 1 if j + 1 < len(boundaries) else len(paragraphs) - 1
        clauses.append((start, end))
    return clauses


def analyze(docx_path):
    comments = load_comments(docx_path)
    paragraphs = load_document_paragraphs(docx_path)
    clauses = group_clauses(paragraphs)

    redlined_clauses = [
        (start, end) for start, end in clauses
        if any(has_tracked_changes(paragraphs[i]) for i in range(start, end + 1))
    ]

    if not redlined_clauses:
        print("No redlined sections found.")
        return

    for i, (start, end) in enumerate(redlined_clauses, 1):
        original_lines = []
        suggested_lines = []
        section_comment_ids = set()

        for para in paragraphs[start:end + 1]:
            if has_tracked_changes(para):
                orig, sugg = get_text_variants(para)
                if orig.strip():
                    original_lines.append(orig)
                if sugg.strip():
                    suggested_lines.append(sugg)
                section_comment_ids.update(get_comment_ids(para))
            else:
                text = get_plain_text(para).strip()
                if text:
                    original_lines.append(text)
                    suggested_lines.append(text)

        print(f"{'=' * 60}")
        print(f"  Redlined Section #{i}")
        print(f"{'=' * 60}")

        print("\n  ORIGINAL:")
        for line in original_lines:
            print(f"    {line}")

        print("\n  SUGGESTED:")
        for line in suggested_lines:
            print(f"    {line}")

        if section_comment_ids:
            print("\n  COMMENTS:")
            for cid in sorted(section_comment_ids, key=lambda x: int(x)):
                c = comments.get(cid)
                if c:
                    print(f"    [{c['author']}]: {c['text']}")

        print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python redline_reader.py <file.docx>")
        sys.exit(1)
    analyze(sys.argv[1])
