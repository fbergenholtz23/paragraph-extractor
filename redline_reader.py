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
import re
from collections import defaultdict

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _tag(local):
    return f"{{{W}}}{local}"


def load_numbering(docx_path):
    """
    Parse word/numbering.xml and return:
    - numbering_map: numId -> { ilvl: { 'numFmt': ..., 'start': ..., 'lvlText': ... } }
    - num_id_to_abs_id: numId -> abstractNumId
    """
    numbering_map = {}
    num_id_to_abs_id = {}
    with zipfile.ZipFile(docx_path) as z:
        if "word/numbering.xml" not in z.namelist():
            return numbering_map, num_id_to_abs_id
        with z.open("word/numbering.xml") as f:
            root = ET.parse(f).getroot()

    abstract_nums = {}
    for abstract_num in root.findall(_tag("abstractNum")):
        abs_id = abstract_num.get(_tag("abstractNumId"))
        levels = {}
        for lvl in abstract_num.findall(_tag("lvl")):
            ilvl = int(lvl.get(_tag("ilvl")))
            num_fmt_el = lvl.find(_tag("numFmt"))
            start_el = lvl.find(_tag("start"))
            lvl_text_el = lvl.find(_tag("lvlText"))
            levels[ilvl] = {
                "numFmt": num_fmt_el.get(_tag("val")) if num_fmt_el is not None else "decimal",
                "start": int(start_el.get(_tag("val"))) if start_el is not None else 1,
                "lvlText": lvl_text_el.get(_tag("val")) if lvl_text_el is not None else f"%{ilvl+1}"
            }
        abstract_nums[abs_id] = levels

    for num in root.findall(_tag("num")):
        num_id = num.get(_tag("numId"))
        abs_ref = num.find(_tag("abstractNumId"))
        abs_id = abs_ref.get(_tag("val")) if abs_ref is not None else None
        num_id_to_abs_id[num_id] = abs_id
        
        # Start with levels from abstractNum
        levels = abstract_nums.get(abs_id, {}).copy()
        
        # Apply overrides
        for override in num.findall(_tag("lvlOverride")):
            ilvl_attr = override.get(_tag("ilvl"))
            if ilvl_attr is None:
                continue
            ilvl = int(ilvl_attr)
            lvl_el = override.find(_tag("lvl"))
            if lvl_el is not None:
                if ilvl not in levels:
                    levels[ilvl] = {}
                num_fmt_el = lvl_el.find(_tag("numFmt"))
                start_el = lvl_el.find(_tag("start"))
                lvl_text_el = lvl_el.find(_tag("lvlText"))
                if num_fmt_el is not None:
                    levels[ilvl]["numFmt"] = num_fmt_el.get(_tag("val"))
                if start_el is not None:
                    levels[ilvl]["start"] = int(start_el.get(_tag("val")))
                if lvl_text_el is not None:
                    levels[ilvl]["lvlText"] = lvl_text_el.get(_tag("val"))
            
            start_override = override.find(_tag("startOverride"))
            if start_override is not None:
                if ilvl not in levels:
                    levels[ilvl] = {"numFmt": "decimal", "lvlText": f"%{ilvl+1}"}
                levels[ilvl]["start"] = int(start_override.get(_tag("val")))

        numbering_map[num_id] = levels

    return numbering_map, num_id_to_abs_id


def load_styles(docx_path):
    """Parse word/styles.xml and return a map: styleId -> (numId, ilvl)."""
    style_numbering = {}
    with zipfile.ZipFile(docx_path) as z:
        if "word/styles.xml" not in z.namelist():
            return style_numbering
        with z.open("word/styles.xml") as f:
            root = ET.parse(f).getroot()
    
    for style in root.findall(_tag("style")):
        style_id = style.get(_tag("styleId"))
        pPr = style.find(_tag("pPr"))
        if pPr is not None:
            numPr = pPr.find(_tag("numPr"))
            if numPr is not None:
                num_id_el = numPr.find(_tag("numId"))
                ilvl_el = numPr.find(_tag("ilvl"))
                num_id = num_id_el.get(_tag("val")) if num_id_el is not None else None
                ilvl = int(ilvl_el.get(_tag("val"))) if ilvl_el is not None else 0
                style_numbering[style_id] = (num_id, ilvl)
    return style_numbering


def get_paragraph_numbering(para_elem, style_map):
    """Return (numId, ilvl) for a paragraph, or (None, None)."""
    pPr = para_elem.find(_tag("pPr"))
    if pPr is not None:
        numPr = pPr.find(_tag("numPr"))
        if numPr is not None:
            numId_el = numPr.find(_tag("numId"))
            ilvl_el = numPr.find(_tag("ilvl"))
            num_id = numId_el.get(_tag("val")) if numId_el is not None else None
            ilvl = int(ilvl_el.get(_tag("val"))) if ilvl_el is not None else 0
            return num_id, ilvl
    
    # Check style map if no direct numPr
    style = get_para_style(para_elem)
    if style in style_map:
        return style_map[style]
        
    return None, None


def resolve_label_part(count, num_fmt):
    """Convert counter value to a formatted label string."""
    if num_fmt == "decimal":
        return str(count)
    elif num_fmt == "lowerLetter":
        return chr(ord("a") + count - 1)
    elif num_fmt == "upperLetter":
        return chr(ord("A") + count - 1)
    elif num_fmt in ("lowerRoman", "upperRoman"):
        val = count
        roman_map = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
        result = ""
        for arabic, roman in roman_map:
            while val >= arabic:
                result += roman
                val -= arabic
        if not result:
            return str(count)
        return result if num_fmt == "lowerRoman" else result.upper()
    return str(count)


def build_hierarchical_label(counters, abs_id, num_id, ilvl, numbering_map):
    """Build a label based on the lvlText template from word/numbering.xml."""
    # Use the specific numId's level info for the template
    level_info = numbering_map.get(num_id, {}).get(ilvl, {})
    template = level_info.get("lvlText", f"%{ilvl+1}")
    
    def replace_placeholder(match):
        # Placeholder is like %1, %2, etc. (1-indexed level)
        lvl_idx = int(match.group(1)) - 1
        lvl_info = numbering_map.get(num_id, {}).get(lvl_idx, {})
        num_fmt = lvl_info.get("numFmt", "decimal")
        # Use the counters associated with the abstractNumId
        count = counters[abs_id][lvl_idx]
        return resolve_label_part(count, num_fmt)

    label = re.sub(r"%(\d+)", replace_placeholder, template)
    return label


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
    numbering_map, num_id_to_abs_id = load_numbering(docx_path)
    style_map = load_styles(docx_path)
    paragraphs = load_document_paragraphs(docx_path)
    clauses = group_clauses(paragraphs)

    # Global tracking of counters by abstractNumId
    # counters[absId][ilvl] = current count
    all_counters = defaultdict(lambda: defaultdict(int))
    
    # Pre-calculate labels for ALL paragraphs in the document
    para_info = []
    for para in paragraphs:
        num_id, ilvl = get_paragraph_numbering(para, style_map)
        abs_id = num_id_to_abs_id.get(num_id)
        
        if num_id and num_id != "0" and abs_id:
            levels_info = numbering_map.get(num_id, {})
            # Ensure counters are initialized for this abstractNumId
            for lvl_idx in range(ilvl + 1):
                if all_counters[abs_id][lvl_idx] == 0:
                    all_counters[abs_id][lvl_idx] = levels_info.get(lvl_idx, {}).get("start", 1)
                elif lvl_idx == ilvl:
                    all_counters[abs_id][lvl_idx] += 1
            
            # Reset deeper levels
            for deeper in list(all_counters[abs_id].keys()):
                if deeper > ilvl:
                    all_counters[abs_id][deeper] = 0
            
            label = build_hierarchical_label(all_counters, abs_id, num_id, ilvl, numbering_map)
            indent = "  " if ilvl > 0 else ""
        else:
            label = ""
            indent = ""
        para_info.append((label, indent))

    redlined_clauses = [
        (start, end) for start, end in clauses
        if any(has_tracked_changes(paragraphs[i]) for i in range(start, end + 1))
    ]

    if not redlined_clauses:
        print("No redlined sections found.")
        return

    for i, (start, end) in enumerate(redlined_clauses, 1):
        original_section_lines = []  # full clause before changes
        new_section_lines = []       # full clause after changes
        redline_only_lines = []      # just the redlined paragraphs
        section_comments = {}        # cid -> [para_labels]

        for para_idx in range(start, end + 1):
            para = paragraphs[para_idx]
            label, indent = para_info[para_idx]
            
            # Capture comments for all paragraphs in the section
            para_cids = get_comment_ids(para)
            for cid in para_cids:
                if cid not in section_comments:
                    section_comments[cid] = []
                section_comments[cid].append(label)

            prefix = f"{label} " if label else ""

            if has_tracked_changes(para):
                orig, sugg = get_text_variants(para)
                if orig.strip():
                    original_section_lines.append(f"{indent}{prefix}{orig}")
                if sugg.strip():
                    new_section_lines.append(f"{indent}{prefix}{sugg}")
                    redline_only_lines.append(f"{indent}{prefix}{sugg}")
                elif orig.strip():
                    redline_only_lines.append(f"{indent}{prefix}[DELETED: {orig.strip()}]")
            else:
                text = get_plain_text(para).strip()
                if text:
                    original_section_lines.append(f"{indent}{prefix}{text}")
                    new_section_lines.append(f"{indent}{prefix}{text}")

        # Use the label of the first paragraph in the section for the header
        section_label = para_info[start][0].strip(".")
        print(f"\nSection #{section_label}\n")

        print("Original section:\n")
        for line in original_section_lines:
            print(f"    {line}")

        print("\nOriginal suggestion (redline):\n")
        for line in redline_only_lines:
            print(f"    {line}")

        print("\nNew section:\n")
        for line in new_section_lines:
            print(f"    {line}")

        if section_comments:
            print("\nComments:\n")
            for cid in sorted(section_comments.keys(), key=lambda x: int(x)):
                c = comments.get(cid)
                if c:
                    valid_labels = [l for l in section_comments[cid] if l]
                    if valid_labels:
                        labels_str = ", ".join(valid_labels)
                        print(f"    [{c['author']}] (Para {labels_str}): {c['text']}")
                    else:
                        print(f"    [{c['author']}]: {c['text']}")

        print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python redline_reader.py <file.docx>")
        sys.exit(1)
    analyze(sys.argv[1])
