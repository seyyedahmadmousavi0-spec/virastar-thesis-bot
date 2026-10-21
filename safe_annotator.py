"""Conservative annotation on an independent package, with atomic publication."""
import copy
import hashlib
import unicodedata
import io
import os
import tempfile
import zipfile
from collections import Counter, defaultdict
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

def _comment_text(issue):
    """Keep qualifiers intact; never cut a long sentence into a stronger claim."""
    from rule_registry import rule_spec
    labels = {"A": "اطمینان بالا", "B": "نیازمند بررسی", "C": "مبهم؛ فقط بررسی دستی"}
    reason = " ".join(str(issue.message or "").split())
    suggestion = " ".join(str(issue.suggestion or "").split())
    if not reason or len(reason) > 160:
        reason = rule_spec(issue.rule_id).title + "؛ دلیل کامل در گزارش آمده است."
    if len(suggestion) > 160:
        suggestion = "پیش از هر تغییر، پیشنهاد کامل گزارش را با متن اصلی مقایسه کنید."
    elif not suggestion:
        suggestion = "این مورد را با توجه به بافت متن بازبینی کنید."
    return (f"{issue.rule_id} | {issue.confidence} — {labels[issue.confidence]}\n"
            f"دلیل: {reason}\nپیشنهاد: {suggestion}")


def _exact_target(paragraph, issue):
    """Resolve a unique literal witness without changing a single run."""
    target = issue.target_text
    left, right = issue.char_start, issue.char_end
    if not target:
        return None
    text = paragraph.text
    if left < 0:
        if text.count(target) != 1:
            return None
        left = text.find(target)
        right = left + len(target)
    if not 0 <= left < right <= len(text) or text[left:right] != target:
        return None
    if unicodedata.category(text[left]).startswith("M") or (
            right < len(text) and unicodedata.category(text[right]).startswith("M")):
        return None
    return left, right


def annotate_document(doc, issues, out_path, max_comments=250, per_par=4,
                      include_ambiguous=False):
    """Annotate exact witnesses on a copy, selecting A then B globally.

    The return value retains its historical meaning (at least one new comment).
    ``_last_annotation_summary.output_created`` describes successful publication
    separately, including the useful zero-comment review copy.
    """
    for name, value in (("max_comments", max_comments), ("per_par", per_par)):
        if name == "per_par" and value is None:
            continue
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise ValueError(f"{name} must be a nonnegative integer")
    if not isinstance(include_ambiguous, bool):
        raise ValueError("include_ambiguous must be a boolean")
    issues = list(issues)
    records = [dict(index=index, rule_id=issue.rule_id, confidence=issue.confidence,
                    block_id=issue.block_id, source_part=issue.source_part,
                    xml_path=issue.xml_path, story=issue.story, par_index=issue.par_index,
                    char_start=issue.char_start, char_end=issue.char_end,
                    status="unapplied", reason="location_unavailable")
               for index, issue in enumerate(issues)]
    summary = {"requested": len(issues), "applied": 0, "unapplied": len(issues),
               "max_comments": max_comments, "per_par": per_par,
               "include_ambiguous": include_ambiguous, "output_created": False,
               "reasons": {}, "records": records}
    if doc is None or not hasattr(doc, "add_comment"):
        for record in records:
            record["reason"] = "annotation_unavailable"
        summary["reasons"] = dict(Counter(record["reason"] for record in records))
        if doc is not None:
            doc._last_annotation_summary = summary
        return False
    doc._last_annotation_summary = summary
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
    selected = 0
    covered = defaultdict(list)
    seen = defaultdict(set)
    # Select before adding comment markers: these markers intentionally make
    # _simple_runs reject later editing of a paragraph containing old comments.
    priority = {"A": 0, "B": 1, "C": 2}
    for index, issue in sorted(enumerate(issues), key=lambda pair: priority[pair[1].confidence]):
        record = records[index]
        if issue.confidence == "C" and not include_ambiguous:
            record["reason"] = "ambiguous_excluded"
            continue
        if issue.protected_overlap:
            record["reason"] = "protected_content"
            continue
        if issue.xml_path and issue.source_part.lstrip("/") == "word/document.xml" and issue.story in {"body", "table"}:
            key = ("xml", issue.xml_path)
            if key not in paragraphs:
                from lxml import etree
                from docx.text.paragraph import Paragraph
                try:
                    found = etree.XPath(issue.xml_path, namespaces={k: v for k, v in working.element.nsmap.items() if k})(working.element)
                except etree.XPathError:
                    record["reason"] = "location_unavailable"
                    continue
                if len(found) != 1 or found[0].tag != qn("w:p"):
                    record["reason"] = "location_unavailable"
                    continue
                paragraphs[key] = Paragraph(found[0], working)
        elif issue.par_index >= 0 and not issue.source_part and not issue.xml_path and issue.story in {"", "body"}:
            key = ("body", issue.par_index)
            if issue.par_index >= len(working.paragraphs):
                continue
            paragraphs[key] = working.paragraphs[issue.par_index]
        else:
            record["reason"] = "unsupported_story" if issue.source_part or issue.story else "location_unavailable"
            continue
        paragraph = paragraphs[key]
        if not _simple_runs(paragraph):
            record["reason"] = "unsafe_structure"
            continue
        target = _exact_target(paragraph, issue)
        if target is None:
            record["reason"] = "uncertain_location"
            continue
        left, right = target
        identity = (issue.rule_id, left, right, issue.target_text)
        if identity in seen[key]:
            record["reason"] = "duplicate"
            continue
        seen[key].add(identity)
        if any(left < b and a < right for a, b in covered[key]):
            record["reason"] = "overlapping_target"
            continue
        if selected >= max_comments:
            record["reason"] = "comment_limit"
            continue
        if per_par is not None and len(grouped[key]) >= per_par:
            record["reason"] = "paragraph_limit"
            continue
        grouped[key].append((issue, index, left, right))
        covered[key].append(target)
        selected += 1
    count = 0
    for key, items in grouped.items():
        paragraph = paragraphs[key]
        pending = []
        for issue, index, left, right in items:
            runs = _find_and_highlight_target(paragraph, issue.target_text, left, right)
            if not runs:
                raise DocumentIntegrityError("Selected witness changed during annotation")
            pending.append((runs, issue, index))
        for runs, issue, index in pending:
            working.add_comment(runs, text=_comment_text(issue),
                                author="Research Editor", initials="RE")
            records[index].update(status="applied", reason="annotated")
            count += 1
    summary.update(applied=count, unapplied=len(issues) - count,
                   reasons=dict(Counter(record["reason"] for record in records
                                        if record["status"] != "applied")))
    if _payload(working) != before:
        raise DocumentIntegrityError("Annotation changed document payload")
    buffer = io.BytesIO()
    working.save(buffer)
    # Retain all original package bytes except main document and comment wiring.
    generated = zipfile.ZipFile(buffer)
    export = io.BytesIO()
    if source and count == 0:
        # With no comments, the independent review copy can retain every byte.
        with open(source, "rb") as original_stream:
            export.write(original_stream.read())
    elif source:
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
    summary["output_created"] = True
    return count > 0
