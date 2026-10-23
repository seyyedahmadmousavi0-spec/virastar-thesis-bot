"""Put every finding in comments on existing text in one reviewed source Word.

Certified exact/paragraph anchors are retained. Findings without a safe exact
location receive explicitly qualified contextual comments on existing source
paragraphs. No body text, paragraph, heading or appendix is ever added.
"""
from collections import Counter
import copy
import gc
import hashlib
import io
import json
from pathlib import Path
import posixpath
import zipfile

from lxml import etree
from docx.oxml.ns import qn

from doc_guard import (DocumentIntegrityError, atomic_write,
                       ensure_distinct_output, reject_unsafe_xml, sha256_file)
from word_review_report import (MAX_REVIEW_BYTES, MAX_REVIEW_CHARACTERS,
                                MAX_REVIEW_FINDINGS, WordReviewReportError,
                                _xml_text)
from report_artifacts import ANNOTATION_REASONS, CONFIDENCE_LABELS, STORY_LABELS, location_label


_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
_COMMENT_TAGS = {qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")}


def _tree(data):
    reject_unsafe_xml(data)
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def _serialize(tree):
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _canonical(node):
    # Source paragraphs can be large; retain fixed-size evidence between the
    # generation and reopening stages instead of another full XML copy.
    return hashlib.sha256(etree.tostring(node, method="c14n")).digest()


def _read_package(path):
    # Callers normally already validated the input; repeat relevant package
    # bounds so this public finalization function is safe independently too.
    archive = zipfile.ZipFile(path)
    try:
        entries = archive.infolist()
        names = [item.filename for item in entries]
        if (len(names) > 5000 or len({name.casefold() for name in names}) != len(names)
                or sum(item.file_size for item in entries) > 256 * 1024 * 1024
                or any(item.file_size > 64 * 1024 * 1024 or item.flag_bits & 1 for item in entries)
                or not {"word/document.xml", "[Content_Types].xml"}.issubset(names)):
            raise DocumentIntegrityError("Invalid or oversized reviewed Word package")
        return archive
    except BaseException:
        archive.close()
        raise


def _comments_part(archive):
    relationships_name = "word/_rels/document.xml.rels"
    if relationships_name in archive.namelist():
        relationships = _tree(archive.read(relationships_name))
    else:
        relationships = etree.Element(f"{{{_REL_NS}}}Relationships", nsmap={None: _REL_NS})
    matches = [node for node in relationships if node.get("Type") == _COMMENTS_REL]
    if len(matches) > 1:
        raise DocumentIntegrityError("Multiple Word comments relationships")
    if matches:
        relation = matches[0]
        target = relation.get("Target", "")
        if (relation.get("TargetMode") == "External" or not target or "\\" in target
                or "#" in target or "?" in target):
            raise DocumentIntegrityError("Unsupported Word comments target")
        name = posixpath.normpath(target.lstrip("/") if target.startswith("/")
                                 else posixpath.join("word", target))
        if name.startswith("../") or name not in archive.namelist():
            raise DocumentIntegrityError("Missing Word comments part")
        comments = _tree(archive.read(name))
        if comments.tag != qn("w:comments"):
            raise DocumentIntegrityError("Invalid Word comments part")
        return relationships, name, comments
    return relationships, None, etree.Element(qn("w:comments"), nsmap={"w": qn("w:p").split("}")[0][1:]})


def _comment_map(comments):
    result = {}
    numeric_ids = set()
    for node in comments:
        if node.tag != qn("w:comment"):
            continue
        identity = node.get(qn("w:id"))
        if (identity is None or not identity.isdecimal() or identity in result
                or int(identity) in numeric_ids):
            raise DocumentIntegrityError("Ambiguous Word comment IDs")
        result[identity] = node
        numeric_ids.add(int(identity))
    return result


def _payload_digest(document):
    """Match the annotator's invariant without retaining a second source DOM."""
    text_digest, structure_digest = hashlib.sha256(), hashlib.sha256()
    for node in document.iter(qn("w:t")):
        text_digest.update((node.text or "").encode("utf-8"))

    def visit(node):
        if node.tag == qn("w:rPr") or node.tag in _COMMENT_TAGS:
            return
        record = None
        if node.tag not in {qn("w:r"), qn("w:t")}:
            if not len(node):
                record = (node.tag, tuple(sorted(node.attrib.items())), node.text)
            else:
                record = ("structure", node.tag, tuple(sorted(node.attrib.items())))
        if record is not None:
            structure_digest.update(json.dumps(record, ensure_ascii=False).encode("utf-8"))
            structure_digest.update(b"\x00")
        for child in node:
            visit(child)
    visit(document)
    return text_digest.digest(), structure_digest.digest()


def _comment_paragraph(text):
    paragraph = etree.Element(qn("w:p"))
    properties = etree.SubElement(paragraph, qn("w:pPr"))
    etree.SubElement(properties, qn("w:bidi"))
    run = etree.SubElement(paragraph, qn("w:r"))
    formatting = etree.SubElement(run, qn("w:rPr"))
    etree.SubElement(formatting, qn("w:rFonts"), {qn("w:ascii"): "Arial", qn("w:hAnsi"): "Arial", qn("w:cs"): "Arial"})
    etree.SubElement(formatting, qn("w:sz"), {qn("w:val"): "22"})
    etree.SubElement(formatting, qn("w:szCs"), {qn("w:val"): "22"})
    etree.SubElement(formatting, qn("w:rtl"))
    for index, line in enumerate(_xml_text(text).split("\n")):
        if index:
            etree.SubElement(run, qn("w:br"))
        leaf = etree.SubElement(run, qn("w:t"))
        leaf.set(qn("xml:space"), "preserve")
        leaf.text = line
    return paragraph


def _add_comment_anchor(first_run, last_run, identity):
    first_run.addprevious(etree.Element(qn("w:commentRangeStart"), {qn("w:id"): identity}))
    end = etree.Element(qn("w:commentRangeEnd"), {qn("w:id"): identity})
    last_run.addnext(end)
    reference = etree.Element(qn("w:r"))
    etree.SubElement(reference, qn("w:commentReference"), {qn("w:id"): identity})
    end.addnext(reference)


def _safe_runs(paragraph):
    allowed = {qn("w:pPr"), qn("w:r"), qn("w:commentRangeStart"),
               qn("w:commentRangeEnd"), qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
    if any(child.tag not in allowed for child in paragraph):
        return None
    runs = []
    for run in paragraph.findall(qn("w:r")):
        text = []
        for child in run:
            if child.tag in {qn("w:rPr"), qn("w:commentReference")}:
                continue
            if child.tag == qn("w:t"):
                text.append(child.text or "")
            elif child.tag in {qn("w:tab"), qn("w:cr")}:
                text.append("\t" if child.tag == qn("w:tab") else "\n")
            elif child.tag == qn("w:br") and child.get(qn("w:type"), "textWrapping") == "textWrapping":
                text.append("\n")
            else:
                return None
        if "".join(text):
            runs.append((run, "".join(text)))
    return runs if "".join(text for _, text in runs).strip() else None


def _highlight_runs(runs):
    # Insert highlight in the WordprocessingML rPr schema order. Existing
    # scientific text and all other run properties stay intact.
    following = {qn("w:" + name) for name in (
        "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
        "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange")}
    for run, _ in runs:
        properties = run.find(qn("w:rPr"))
        if properties is None:
            properties = etree.Element(qn("w:rPr"))
            run.insert(0, properties)
        highlight = properties.find(qn("w:highlight"))
        if highlight is None:
            highlight = etree.Element(qn("w:highlight"))
            successor = next((child for child in properties if child.tag in following), None)
            if successor is None:
                properties.append(highlight)
            else:
                successor.addprevious(highlight)
        highlight.set(qn("w:val"), "yellow")


def _context_candidates(document):
    body = document.find(qn("w:body"))
    body_indexes = {paragraph: index for index, paragraph in enumerate(body.findall(qn("w:p")))}
    candidates = []
    excluded = {qn("w:" + tag) for tag in ("txbxContent", "del", "ins", "moveFrom", "moveTo", "sdt")}
    for paragraph in body.iter(qn("w:p")):
        if any(ancestor.tag in excluded for ancestor in paragraph.iterancestors()):
            continue
        runs = _safe_runs(paragraph)
        if runs:
            candidates.append({"paragraph": paragraph, "runs": runs,
                               "text": "".join(text for _, text in runs),
                               "index": body_indexes.get(paragraph, -1),
                               "path": document.getroottree().getpath(paragraph)})
    return candidates


def _resolve_context(issue, candidates, document, by_path, by_index, by_element, heading_candidates):
    unsupported = issue.story not in {"", "body", "table"} or issue.source_part.lstrip("/") not in {"", "word/document.xml"}
    if not unsupported:
        if issue.xml_path:
            candidate = by_path.get(issue.xml_path)
            if candidate is not None:
                return candidate, "paragraph"
            try:
                found = document.xpath(issue.xml_path, namespaces={key: value for key, value in document.nsmap.items() if key})
            except etree.XPathError:
                found = []
            if len(found) == 1:
                candidate = by_element.get(found[0])
                if candidate is not None:
                    return candidate, "paragraph"
        elif issue.par_index >= 0:
            candidate = by_index.get(issue.par_index)
            if candidate is not None:
                return candidate, "paragraph"
    # A related heading is a navigation aid only. Its text must match the
    # reported location unambiguously; never guess a semantic source phrase.
    location = " ".join(str(issue.loc or "").casefold().split())
    headings = []
    for title, item in heading_candidates:
        if title in location:
            headings.append(item)
    candidate = headings[0] if len(headings) == 1 else candidates[0]
    return candidate, "story" if unsupported else "document"


def _context_text(issue, index, scope, reason):
    story_label = STORY_LABELS.get(issue.story, "بخش گزارش‌شدهٔ سند")
    if scope == "story":
        qualifier = f"ارجاع به {story_label}؛ عبارت زردشده محل دقیق ایراد نیست"
    elif scope == "paragraph":
        qualifier = "ارجاع در سطح پاراگراف؛ عبارت زردشده لزوماً خطا نیست"
    else:
        qualifier = "در سطح سند؛ عبارت زردشده لزوماً خطا نیست"
    lines = [f"مورد {index + 1} | {issue.rule_id} | {issue.confidence} — {CONFIDENCE_LABELS[issue.confidence]} [{qualifier}]",
             "محل گزارش‌شده: " + location_label(issue.loc),
             "علت کامنت ارجاعی: " + ANNOTATION_REASONS.get(reason, "محل دقیق برای درج کامنت تأیید نشد") + ".",
             "دلیل: " + str(issue.message or "این مورد را با متن اصلی بررسی کنید."),
             "پیشنهاد: " + str(issue.suggestion or "پیش از هر تغییر، متن اصلی و بافت پژوهش را دستی بررسی کنید.")]
    if issue.confidence == "C":
        lines.append("فقط بررسی دستی؛ این مورد خطای قطعی محسوب نمی‌شود.")
    if issue.protected_overlap:
        lines.append("محتوای علمی حفاظت‌شده است؛ اعداد، فرمول‌ها و متن علمی اصلاح خودکار نشده‌اند.")
    if issue.location_status == "exact" and issue.target_text:
        lines.append("شاهد گزارش‌شده از محل اصلی: «" + issue.target_text + "»")
    elif issue.snippet:
        lines.append("متن پیرامونی ثبت‌شده؛ محل دقیق تأیید نشده: «" + issue.snippet + "»")
    return _xml_text("\n".join(lines))


def _add_context_comments(document, comments, remaining, findings, records, next_id):
    candidates = _context_candidates(document)
    if not candidates:
        raise WordReviewReportError("متن موجود و ایمنی برای اتصال کامنت پیدا نشد؛ فایل ناقص منتشر نشد.")
    by_path = {item["path"]: item for item in candidates}
    by_index = {item["index"]: item for item in candidates if item["index"] >= 0}
    by_element = {item["paragraph"]: item for item in candidates}
    heading_candidates = []
    for item in candidates:
        style = item["paragraph"].find("./" + qn("w:pPr") + "/" + qn("w:pStyle"))
        style_name = (style.get(qn("w:val"), "") if style is not None else "").casefold()
        title = " ".join(item["text"].casefold().split())
        if style_name.startswith(("heading", "title", "عنوان")) and 3 <= len(title) <= 200:
            heading_candidates.append((title, item))
    new_ids, context_records, highlighted = set(), [], set()
    characters = 0
    for index in remaining:
        issue = findings[index]
        anchor, scope = _resolve_context(issue, candidates, document, by_path, by_index, by_element, heading_candidates)
        text = _context_text(issue, index, scope, records[index].get("reason", ""))
        characters += len(text)
        if characters > MAX_REVIEW_CHARACTERS:
            raise WordReviewReportError("حجم کامنت‌ها از ظرفیت Word کامل بیشتر است؛ فایل ناقص منتشر نشد.")
        identity = str(next_id)
        next_id += 1
        new_ids.add(identity)
        if anchor["path"] not in highlighted:
            _highlight_runs(anchor["runs"])
            highlighted.add(anchor["path"])
        _add_comment_anchor(anchor["runs"][0][0], anchor["runs"][-1][0], identity)
        comment = etree.SubElement(comments, qn("w:comment"),
                                   {qn("w:id"): identity, qn("w:author"): "ویراستار پژوهش", qn("w:initials"): "VP"})
        comment.append(_comment_paragraph(text))
        context_records.append({"index": index, "anchor_par_index": anchor["index"],
                                "anchor_xml_path": anchor["path"], "scope": scope,
                                "source_story": issue.story, "reason": records[index].get("reason", "")})
    return new_ids, context_records


def _validate_new_anchors(document, identities):
    starts = Counter(node.get(qn("w:id")) for node in document.iter(qn("w:commentRangeStart")))
    ends = Counter(node.get(qn("w:id")) for node in document.iter(qn("w:commentRangeEnd")))
    references = Counter(node.get(qn("w:id")) for node in document.iter(qn("w:commentReference")))
    active = set()
    yellow = set()
    for node in document.iter():
        if node.tag == qn("w:commentRangeStart"):
            active.add(node.get(qn("w:id")))
        elif node.tag == qn("w:commentRangeEnd"):
            active.discard(node.get(qn("w:id")))
        elif node.tag == qn("w:r"):
            highlight = node.find("./" + qn("w:rPr") + "/" + qn("w:highlight"))
            if highlight is not None and highlight.get(qn("w:val")) == "yellow" and any(
                    leaf.text for leaf in node.iter(qn("w:t"))):
                yellow.update(active)
    if any(starts[identity] != 1 or ends[identity] != 1 or references[identity] != 1
           or identity not in yellow for identity in identities):
        raise DocumentIntegrityError("A reviewed finding lost its independent comment or yellow witness")


def _validated_records(issues, stats):
    summary = stats.get("annotation_summary")
    if not isinstance(summary, dict) or summary.get("output_created") is not True:
        raise DocumentIntegrityError("Source annotation was not completed")
    records = summary.get("records")
    if (summary.get("requested") != len(issues) or not isinstance(records, list)
            or len(records) != len(issues)):
        raise DocumentIntegrityError("Incomplete annotation audit")
    mapped = {}
    for record in records:
        index = record.get("index")
        if not isinstance(index, int) or isinstance(index, bool) or index in mapped or not 0 <= index < len(issues):
            raise DocumentIntegrityError("Invalid annotation finding index")
        issue = issues[index]
        if (any(record.get(key) != getattr(issue, key) for key in
                ("rule_id", "confidence", "block_id", "source_part", "xml_path", "story", "par_index", "char_start", "char_end"))
                or record.get("status") not in {"applied", "unapplied"}
                or (record.get("status") == "applied" and record.get("anchor_scope") not in {"exact", "paragraph"})):
            raise DocumentIntegrityError("Annotation audit does not match findings")
        mapped[index] = record
    applied = sum(record["status"] == "applied" for record in records)
    exact = sum(record["status"] == "applied" and record["anchor_scope"] == "exact" for record in records)
    paragraph = applied - exact
    if (summary.get("applied") != applied or summary.get("unapplied") != len(issues) - applied
            or summary.get("exact_applied") != exact or summary.get("paragraph_applied") != paragraph
            or summary.get("highlighted") != applied):
        raise DocumentIntegrityError("Inconsistent source annotation counts")
    return mapped, applied, exact, paragraph


def finalize_word_review(source_path, annotated_path, issues, stats, rules=None):
    """Atomically comment every finding on existing source Word text only.

    ``annotated_path`` must be the conservative annotated copy with the matching
    ``stats['annotation_summary']``. It is replaced only after all findings have
    a verified independent comment and yellow witness. The input file is never
    overwritten. This function returns counts; the caller stores the result in
    ``stats['unified_word_review']`` only after successful publication.
    """
    source_path, annotated_path = Path(source_path), Path(annotated_path)
    ensure_distinct_output(source_path, annotated_path)
    initial_hash = sha256_file(source_path)
    reviewed_hash = sha256_file(annotated_path)
    findings = list(issues)
    if len(findings) > MAX_REVIEW_FINDINGS:
        raise WordReviewReportError("تعداد یافته‌ها از ظرفیت Word کامل بیشتر است؛ فایل ناقص منتشر نشد.")
    records, applied, exact, paragraph_count = _validated_records(findings, stats)
    remaining = [index for index, record in records.items() if record["status"] != "applied"]
    remaining.sort()
    try:
        with _read_package(source_path) as source, _read_package(annotated_path) as reviewed:
            source_payload = _payload_digest(_tree(source.read("word/document.xml")))
            document = _tree(reviewed.read("word/document.xml"))
            if _payload_digest(document) != source_payload:
                raise DocumentIntegrityError("Source review changed scientific document payload")
            body = document.find(qn("w:body"))
            if body is None:
                raise DocumentIntegrityError("Reviewed Word has no document body")
            original_count = len(body)
            paragraph_nodes = len(list(document.iter(qn("w:p"))))
            text_nodes = len(list(document.iter(qn("w:t"))))
            relationships, comments_name, comments = _comments_part(reviewed)
            _, _, source_comments = _comments_part(source)
            original_comments = _comment_map(comments)
            old_comments = _comment_map(source_comments)
            source_new_ids = set(original_comments) - set(old_comments)
            if (not set(old_comments).issubset(original_comments) or len(source_new_ids) != applied
                    or len(original_comments) != len(old_comments) + applied):
                raise DocumentIntegrityError("Source review comment count does not match audit")
            # The source annotator has already checked scientific payload. Here
            # validate the actual newly added source comments and their yellow
            # anchors before treating them as complete findings.
            _validate_new_anchors(document, source_new_ids)
            if not remaining:
                if sha256_file(source_path) != initial_hash or sha256_file(annotated_path) != reviewed_hash:
                    raise DocumentIntegrityError("Word changed during finalization")
                return {"output_created": True, "requested": len(findings), "source_comments": applied,
                        "context_comments": 0, "appendix_comments": 0, "comments": applied, "highlighted": applied,
                        "exact_comments": exact, "paragraph_comments": paragraph_count, "context_records": []}

            comment_snapshots = {identity: _canonical(node) for identity, node in original_comments.items()}
            known_ids = set(original_comments)
            for node in document.iter():
                if node.tag in _COMMENT_TAGS:
                    identity = node.get(qn("w:id"))
                    if identity is not None:
                        if not identity.isdecimal():
                            raise DocumentIntegrityError("Invalid existing Word anchor ID")
                        known_ids.add(identity)
            next_id = max((int(identity) for identity in known_ids), default=-1) + 1
            del node
            context_ids, context_records = _add_context_comments(
                document, comments, remaining, findings, records, next_id)
            if (_payload_digest(document) != source_payload
                    or len(body) != original_count
                    or len(list(document.iter(qn("w:p")))) != paragraph_nodes
                    or len(list(document.iter(qn("w:t")))) != text_nodes):
                raise DocumentIntegrityError("Context comments changed source text or paragraphs")

            replacements = {"word/document.xml": _serialize(document)}
            if comments_name is None:
                names = set(reviewed.namelist())
                comments_name = "word/review-comments.xml"
                suffix = 1
                while comments_name in names:
                    comments_name = f"word/review-comments-{suffix}.xml"
                    suffix += 1
                existing_rel_ids = {node.get("Id") for node in relationships}
                number = 1
                while f"rId{number}" in existing_rel_ids:
                    number += 1
                etree.SubElement(relationships, f"{{{_REL_NS}}}Relationship", {
                    "Id": f"rId{number}", "Type": _COMMENTS_REL,
                    "Target": posixpath.relpath(comments_name, "word")})
                replacements["word/_rels/document.xml.rels"] = _serialize(relationships)
                types = _tree(reviewed.read("[Content_Types].xml"))
                etree.SubElement(types, f"{{{_CT_NS}}}Override", {
                    "PartName": "/" + comments_name, "ContentType": _COMMENTS_TYPE})
                replacements["[Content_Types].xml"] = _serialize(types)
            replacements[comments_name] = _serialize(comments)
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w") as output:
                for item in reviewed.infolist():
                    output.writestr(copy.copy(item), replacements.pop(item.filename, None) or reviewed.read(item.filename))
                for name, data in replacements.items():
                    output.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
            payload = buffer.getvalue()
            if len(payload) > MAX_REVIEW_BYTES:
                raise WordReviewReportError("حجم فایل از ظرفیت Word کامل بیشتر است؛ فایل ناقص منتشر نشد.")
            original_comment_ids = set(original_comments)
            old_comment_count = len(old_comments)
            had_original_comments = bool(original_comments)
            # Child nodes keep their entire owning XML tree alive. Release all
            # of them before reopening the saved package on the free host.
            del document, body, comments, source_comments, original_comments, old_comments
            del buffer
            gc.collect()
            # Reopen the actual completed ZIP, rather than trusting the in-memory
            # construction. Original paragraphs/section settings and opaque parts
            # remain intact, while each new ID has all three Word references.
            with zipfile.ZipFile(io.BytesIO(payload)) as reopened:
                final_document = _tree(reopened.read("word/document.xml"))
                final_body = final_document.find(qn("w:body"))
                if (_payload_digest(final_document) != source_payload
                        or len(final_body) != original_count
                        or len(list(final_document.iter(qn("w:p")))) != paragraph_nodes
                        or len(list(final_document.iter(qn("w:t")))) != text_nodes):
                    raise DocumentIntegrityError("Context comments changed source document payload")
                final_comments = _comment_map(_tree(reopened.read(comments_name)))
                if (set(final_comments) != original_comment_ids | context_ids
                        or len(final_comments) != old_comment_count + len(findings)
                        or any(_canonical(final_comments[identity]) != snapshot
                               for identity, snapshot in comment_snapshots.items())):
                    raise DocumentIntegrityError("Unified Word lost or changed comments")
                _validate_new_anchors(final_document, source_new_ids | context_ids)
                changed = {"word/document.xml", comments_name}
                if "word/_rels/document.xml.rels" not in reviewed.namelist() or not had_original_comments:
                    # Creation of comments wiring is the only other permitted
                    # change. Existing wiring is left byte-identical otherwise.
                    changed.update({"word/_rels/document.xml.rels", "[Content_Types].xml"})
                for name in reviewed.namelist():
                    if name not in changed and reopened.read(name) != reviewed.read(name):
                        raise DocumentIntegrityError("Unified Word changed an unrelated package part")
            if sha256_file(source_path) != initial_hash or sha256_file(annotated_path) != reviewed_hash:
                raise DocumentIntegrityError("Word changed during finalization")
            # Windows cannot replace an archive while its read handle is open.
            # All validation/reads are complete; the context's final close is
            # idempotent and still protects exception paths above.
            reviewed.close()
            atomic_write(annotated_path, payload, source_path=source_path)
            return {"output_created": True, "requested": len(findings), "source_comments": applied,
                    "context_comments": len(remaining), "appendix_comments": 0,
                    "comments": len(findings), "highlighted": len(findings),
                    "exact_comments": exact, "paragraph_comments": paragraph_count,
                    "context_records": context_records}
    finally:
        if sha256_file(source_path) != initial_hash:
            raise DocumentIntegrityError("Source changed during Word finalization")
