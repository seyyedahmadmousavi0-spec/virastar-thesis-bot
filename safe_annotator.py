"""Conservative annotation on an independent package, with atomic publication."""
import copy
import hashlib
import unicodedata
import io
import os
import tempfile
import zipfile
from collections import defaultdict
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from doc_guard import ensure_distinct_output, sha256_file, DocumentIntegrityError

def _simple_runs(p):
    # Fields, drawings, hyperlinks, tracked edits and math require a different mapper.
    if any(child.tag not in {qn("w:pPr"), qn("w:r")} for child in p._p):
        return False
    return all(all(c.tag in {qn("w:rPr"), qn("w:t")} for c in r._r) for r in p.runs)

def _split_run_at(p, r, start_idx, end_idx):
    text = r.text
    parent = r._r.getparent()
    index = parent.index(r._r)
    result = []
    for part in (text[:start_idx], text[start_idx:end_idx], text[end_idx:]):
        new = p.add_run(part)
        if r._r.rPr is not None:
            new._r.insert(0, copy.deepcopy(r._r.rPr))
        result.append(new)
    parent.remove(r._r)
    for offset, run in enumerate(result):
        parent.insert(index + offset, run._r)
    result[1].font.highlight_color = WD_COLOR_INDEX.YELLOW
    return tuple(result)

def _find_and_highlight_target(p, target_text, start=-1, end=-1):
    if not target_text or not _simple_runs(p):
        return None
    text = p.text
    if start >= 0:
        if end <= start or text[start:end] != target_text:
            return None
    else:
        if text.count(target_text) != 1:
            return None
        start = text.find(target_text)
        end = start + len(target_text)
    if (start < len(text) and unicodedata.category(text[start]).startswith("M")) or \
       (end < len(text) and unicodedata.category(text[end]).startswith("M")):
        return None
    spans, cursor = [], 0
    for run in p.runs:
        spans.append((cursor, cursor + len(run.text), run))
        cursor += len(run.text)
    targets = []
    for left, right, run in spans:
        if left < end and start < right:
            _, middle, _ = _split_run_at(p, run, max(start, left)-left, min(end, right)-left)
            targets.append(middle)
    return targets or None

def _payload(doc):
    # Text and every non-run-property XML leaf, in document order.
    ignored = {qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")}
    result = []
    def visit(node):
        if node.tag == qn("w:rPr") or node.tag in ignored:
            return
        if node.tag != qn("w:r"):
            if len(node) == 0:
                result.append((node.tag, tuple(sorted(node.attrib.items())), node.text))
            elif node.tag not in {qn("w:t")}:
                result.append(("structure", node.tag, tuple(sorted(node.attrib.items()))))
        for child in node:
            visit(child)
    # Text runs may split; compare concatenated w:t separately.
    text = "".join(doc.element.xpath("//w:t/text()"))
    visit(doc.element)
    result = [r for r in result if r[0] != qn("w:t")]
    return text, result

def annotate_document(doc, issues, out_path, max_comments=250, per_par=4):
    if doc is None or not hasattr(doc, "add_comment"):
        return False
    source = getattr(doc, "_review_source_path", None)
    initial_hash = getattr(doc, "_review_source_hash", None)
    if source:
        ensure_distinct_output(source, out_path)
        if sha256_file(source) != initial_hash:
            raise DocumentIntegrityError("Source changed since analysis")
        original_doc = Document(source)
        if doc.element.xml != original_doc.element.xml:
            raise DocumentIntegrityError("In-memory document changed since analysis")
        working = original_doc
    else:
        snapshot = io.BytesIO()
        doc.save(snapshot)
        snapshot.seek(0)
        working = Document(snapshot)
    before = _payload(working)
    grouped = defaultdict(list)
    paragraphs = {}
    for issue in issues:
        if issue.confidence == "C" or issue.protected_overlap:
            continue
        if issue.xml_path and issue.source_part.lstrip("/") == "word/document.xml" and issue.story in {"body", "table"}:
            key = ("xml", issue.xml_path)
            if key not in paragraphs:
                from lxml import etree
                from docx.text.paragraph import Paragraph
                try:
                    found = etree.XPath(issue.xml_path, namespaces={k: v for k, v in working.element.nsmap.items() if k})(working.element)
                except etree.XPathError:
                    continue
                if len(found) != 1 or found[0].tag != qn("w:p"):
                    continue
                paragraphs[key] = Paragraph(found[0], working)
        elif issue.par_index >= 0 and not issue.source_part:
            key = ("body", issue.par_index)
            if issue.par_index >= len(working.paragraphs):
                continue
            paragraphs[key] = working.paragraphs[issue.par_index]
        else:
            continue
        grouped[key].append(issue)
    count = 0
    for key, items in grouped.items():
        paragraph = paragraphs[key]
        # Resolve every target before modifying XML. No paragraph fallback.
        if not _simple_runs(paragraph):
            continue
        seen = set()
        pending = []
        covered = []
        for issue in items:
            if count + len(pending) >= max_comments or len(pending) >= per_par:
                break
            key = (issue.rule_id, issue.char_start, issue.char_end, issue.target_text)
            if key in seen:
                continue
            seen.add(key)
            left, right = issue.char_start, issue.char_end
            if left < 0 and issue.target_text and paragraph.text.count(issue.target_text) == 1:
                left = paragraph.text.find(issue.target_text)
                right = left + len(issue.target_text)
            if any(left < b and a < right for a, b in covered):
                continue
            runs = _find_and_highlight_target(paragraph, issue.target_text, left, right)
            if runs:
                pending.append((runs, issue))
                covered.append((left, right))
        for runs, issue in pending:
            working.add_comment(runs, text=issue.rule_id + " [" + issue.confidence + "] "
                                + issue.message + ("\nSuggestion: " + issue.suggestion if issue.suggestion else ""),
                                author="Research Editor", initials="RE")
            count += 1
    doc._last_annotation_summary = {"requested": len(issues), "applied": count,
                                    "unapplied": len(issues) - count}
    if _payload(working) != before:
        raise DocumentIntegrityError("Annotation changed document payload")
    buffer = io.BytesIO()
    working.save(buffer)
    # Retain all original package bytes except main document and comment wiring.
    generated = zipfile.ZipFile(buffer)
    export = io.BytesIO()
    if source:
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(export, "w") as output:
            names = set(original.namelist())
            replace = {"word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
            # Existing valid packages may use a custom comments part name.
            for relationship in working.part.rels.values():
                if relationship.reltype.endswith("/comments") and not relationship.is_external:
                    replace.add(str(relationship.target_part.partname).lstrip("/"))
            for item in original.infolist():
                data = generated.read(item.filename) if item.filename in replace and item.filename in generated.namelist() else original.read(item.filename)
                if item.filename in {"word/_rels/document.xml.rels", "[Content_Types].xml"} and item.filename in generated.namelist():
                    from lxml import etree
                    old_tree = etree.fromstring(original.read(item.filename))
                    new_tree = etree.fromstring(data)
                    def entry_key(node):
                        return (node.tag, node.get("Id") or node.get("PartName") or node.get("Extension"))
                    known = {entry_key(node): node for node in old_tree}
                    for node in new_tree:
                        key = entry_key(node)
                        if key not in known:
                            old_tree.append(node)
                        elif dict(known[key].attrib) != dict(node.attrib):
                            raise DocumentIntegrityError("Conflicting package relationship")
                    data = etree.tostring(old_tree, xml_declaration=True, encoding="UTF-8", standalone=True)
                output.writestr(item, data)
            for name in generated.namelist():
                if name not in names:
                    output.writestr(name, generated.read(name))
    else:
        export.write(buffer.getvalue())
    generated.close()
    export.seek(0)
    reopened = Document(export)
    if _payload(reopened) != before:
        raise DocumentIntegrityError("Saved annotation failed content validation")
    destination = os.path.abspath(out_path)
    descriptor, temporary = tempfile.mkstemp(prefix=".review_", suffix=".docx", dir=os.path.dirname(destination))
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(export.getvalue())
        if source:
            ensure_distinct_output(source, destination)
            if sha256_file(source) != initial_hash:
                raise DocumentIntegrityError("Source changed during export")
        os.replace(temporary, destination)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return count > 0
