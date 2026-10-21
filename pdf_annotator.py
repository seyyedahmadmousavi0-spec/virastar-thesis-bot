"""Experimental PDF notes: exact extraction coordinates, independent output only.

No text search, OCR, normalization or editing of content streams is performed.
Persian note rendering in third-party readers still needs manual validation, so
callers must keep this feature opt-in and retain the complete page-aware report.
"""
from collections import Counter
import hashlib
import math
import os
import re
import tempfile
import unicodedata

from doc_guard import (DocumentIntegrityError, ensure_distinct_output,
                       sha256_file, validate_document)
from pdf_handler import pymupdf


def _record(index, issue, reason):
    return {"index": index, "rule_id": str(getattr(issue, "rule_id", "")),
            "confidence": getattr(issue, "confidence", ""),
            "page_index": getattr(issue, "page_index", -1),
            "status": "unapplied", "reason": reason}


def _finish(records, created=False):
    applied = sum(record["status"] == "applied" for record in records)
    return {"requested": len(records), "applied": applied,
            "unapplied": len(records) - applied, "output_created": created,
            "reasons": dict(Counter(record["reason"] for record in records
                                    if record["status"] != "applied")),
            "records": records, "experimental": True}


def _streams(document, xrefs=None):
    """Preserve every stream's encoded bytes AND decoding dictionary exactly.

    This proves preservation of decoded content without eagerly decompressing
    arbitrary image/font streams (a small input may contain enormous streams).
    Form, image and font streams are included, not only page /Contents.
    """
    selected = range(1, document.xref_length()) if xrefs is None else xrefs
    return {xref: (hashlib.sha256(document.xref_stream_raw(xref)).hexdigest(),
                   document.xref_object(xref))
            for xref in selected if document.xref_is_stream(xref)}


def _payload(document):
    return [(page.get_text("text", sort=False), tuple(page.mediabox),
             tuple(page.cropbox), page.rotation, tuple(page.get_contents()))
            for page in document]


def _existing_notes(document):
    return {annotation.xref: document.xref_object(annotation.xref)
            for page in document for annotation in (page.annots() or ())}


def _page_blocks(page):
    """Reconstruct the extractor's exact blocks, one codepoint per raw entry.

    Raw text and blocks must agree byte-for-byte in Unicode, including newlines.
    The filtered block numbering is the one used by pdf_handler.py. Any mismatch
    leaves that block unannotatable instead of guessing a different occurrence.
    """
    extracted = [item for item in page.get_text("blocks", sort=False)
                 if len(item) >= 7 and item[6] == 0 and item[4].strip()]
    raw = [block for block in page.get_text("rawdict", sort=False)["blocks"]
           if block.get("type") == 0 and block.get("lines")]
    result = {}
    # There is no safe positional pairing after a missing or merged raw block.
    if len(extracted) != len(raw):
        return result
    for number, (block, raw_block) in enumerate(zip(extracted, raw), 1):
        characters, positions = [], []
        for line in raw_block["lines"]:
            for span in line["spans"]:
                for char in span["chars"]:
                    value = char["c"]
                    characters.append(value)
                    # A single glyph mapping to several codepoints cannot be
                    # partially selected with justified character boundaries.
                    positions.extend([(line["dir"], span, char, len(value))] * len(value))
            characters.append("\n")
            positions.append(None)
        text = "".join(characters)
        if text != block[4] or any(abs(a - b) > 0.01 for a, b in zip(block[:4], raw_block["bbox"])):
            continue
        result[number] = (text, positions)
    return result


def _target_quads(page, issue, blocks):
    """Return exact raw-character quads or a reason; never search for a target."""
    page_index = getattr(issue, "page_index", -1)
    match = re.fullmatch(r"pdf:p([1-9][0-9]*):b([1-9][0-9]*)", getattr(issue, "block_id", ""))
    if not match or int(match[1]) != page_index + 1:
        return None, "missing_exact_location"
    if getattr(issue, "source_part", "") not in {"", f"pdf/page/{page_index + 1}"}:
        return None, "source_mismatch"
    item = blocks.get(int(match[2]))
    if item is None:
        return None, "extraction_mapping_mismatch"
    text, positions = item
    start, end = getattr(issue, "char_start", -1), getattr(issue, "char_end", -1)
    target = getattr(issue, "target_text", "")
    if (type(start) is not int or type(end) is not int or not isinstance(target, str)
            or not target or not 0 <= start < end <= len(text)
            or text[start:end] != target):
        return None, "evidence_mismatch"
    if (unicodedata.category(text[start]).startswith("M")
            or (end < len(text) and unicodedata.category(text[end]).startswith("M"))):
        return None, "grapheme_boundary"
    bounds = page.rect * page.derotation_matrix
    quads, used = [], set()
    for char, position in zip(text[start:end], positions[start:end]):
        if position is None:
            # Newlines have no page glyph, but do not join separate lines.
            continue
        direction, span, raw_char, codepoints = position
        if codepoints != 1:
            return None, "ambiguous_glyph"
        quad = pymupdf.recover_char_quad(direction, span, raw_char)
        coordinates = tuple(value for point in quad for value in point)
        if not all(math.isfinite(value) for value in coordinates):
            return None, "invalid_geometry"
        if quad.rect.is_empty or quad.rect.get_area() <= 0:
            if char.isspace() or unicodedata.category(char) in {"Cf", "Mn", "Me"}:
                continue
            return None, "invalid_geometry"
        if not quad.is_convex or not (bounds + (-1, -1, 1, 1)).contains(quad.rect):
            return None, "invalid_geometry"
        key = tuple(round(value, 5) for value in coordinates)
        if key not in used:
            used.add(key)
            quads.append(quad)
    return (quads, "") if quads else (None, "no_visible_target")


def _note_text(issue):
    def compact(value):
        return " ".join(str(value or "").split())
    reason, suggestion = compact(issue.message), compact(issue.suggestion)
    if not reason or len(reason) > 180:
        reason = "این مورد نیازمند بازبینی است؛ دلیل کامل در گزارش آمده است."
    if len(suggestion) > 180:
        suggestion = "پیش از هر تغییر، پیشنهاد کامل گزارش را با متن اصلی مقایسه کنید."
    return (f"{compact(issue.rule_id)} [{issue.confidence}]\n"
            f"دلیل: {reason}\n"
            f"پیشنهاد: {suggestion or 'این مورد را در متن اصلی بررسی کنید.'}")


def annotate_pdf(source_path, issues, out_path, max_comments=250):
    """Create a verified PDF copy containing highlight notes for exact A/B spans.

    Returns an accounting record for every supplied issue. Skipped findings must
    remain in the ordinary report. Source aliases raise DocumentIntegrityError.
    Parser/geometry/save failures produce no output; document text is never logged.
    """
    if type(max_comments) is not int or not 0 <= max_comments <= 5000:
        raise ValueError("max_comments must be an integer between 0 and 5000")
    ensure_distinct_output(source_path, out_path)
    validate_document(source_path, "pdf")
    initial_hash = sha256_file(source_path)
    issues = list(issues)
    records = [_record(index, issue, "not_processed") for index, issue in enumerate(issues)]
    temporary = None
    try:
        if pymupdf is None:
            for record in records:
                record["reason"] = "dependency_unavailable"
            return _finish(records)
        with pymupdf.open(source_path) as document:
            unsafe = ("encrypted" if document.needs_pass or document.is_encrypted else
                      "repaired_pdf" if document.is_repaired else
                      "signed_pdf" if document.get_sigflags() > 0 else
                      "page_budget" if len(document) > 1000 else "")
            if unsafe:
                for record in records:
                    record["reason"] = unsafe
                return _finish(records)
            before, streams, old_notes = _payload(document), _streams(document), _existing_notes(document)
            if sum(len(page[0]) for page in before) > 4_000_000:
                for record in records:
                    record["reason"] = "character_budget"
                return _finish(records)
            cache, applied, covered, expected_vertices = {}, 0, [], {}
            for index in sorted(range(len(issues)), key=lambda n: ({"A": 0, "B": 1}.get(getattr(issues[n], "confidence", ""), 2), n)):
                issue, record = issues[index], records[index]
                if getattr(issue, "confidence", "") not in {"A", "B"}:
                    record["reason"] = "confidence_filtered"
                    continue
                if getattr(issue, "protected_overlap", False):
                    record["reason"] = "protected_content"
                    continue
                if applied >= max_comments:
                    record["reason"] = "comment_limit"
                    continue
                page_index = getattr(issue, "page_index", -1)
                if type(page_index) is not int or not 0 <= page_index < len(document):
                    record["reason"] = "missing_exact_location"
                    continue
                page = document[page_index]
                if page_index not in cache:
                    cache[page_index] = _page_blocks(page)
                quads, reason = _target_quads(page, issue, cache[page_index])
                if reason:
                    record["reason"] = reason
                    continue
                location = (page_index, issue.block_id, issue.char_start, issue.char_end)
                if location in covered:
                    record["reason"] = "duplicate_target"
                    continue
                if any(location[:2] == previous[:2] and location[2] < previous[3]
                       and previous[2] < location[3] for previous in covered):
                    record["reason"] = "overlapping_target"
                    continue
                annotation = page.add_highlight_annot(quads)
                annotation.set_info(title="ویراستار پژوهش", content=_note_text(issue),
                                    subject=f"{issue.rule_id} [{issue.confidence}]")
                annotation.set_colors(stroke=(1, 0.88, 0.2) if issue.confidence == "A" else (0.6, 0.82, 1))
                annotation.update()
                covered.append(location)
                record.update(status="applied", reason="", annotation_xref=annotation.xref)
                expected_vertices[annotation.xref] = [tuple(point) for quad in quads for point in quad]
                applied += 1
            if not applied:
                return _finish(records)
            destination = os.path.abspath(out_path)
            descriptor, temporary = tempfile.mkstemp(prefix=".review_pdf_", suffix=".pdf", dir=os.path.dirname(destination))
            os.close(descriptor)
            # No cleanup, compression, content rewriting, renumbering or repair.
            document.save(temporary, garbage=0, clean=0, deflate=False, no_new_id=True)
        with pymupdf.open(temporary) as reopened:
            if (_payload(reopened) != before or _streams(reopened, streams) != streams
                    or any(reopened.xref_object(xref) != value for xref, value in old_notes.items())):
                raise DocumentIntegrityError("PDF annotation changed original content")
            for record in records:
                if record["status"] != "applied":
                    continue
                page = reopened[record["page_index"]]
                annotation = page.load_annot(record["annotation_xref"])
                if annotation is None or annotation.info.get("content") != _note_text(issues[record["index"]]):
                    raise DocumentIntegrityError("Saved PDF note failed validation")
                expected = expected_vertices[record["annotation_xref"]]
                vertices = annotation.vertices or []
                if (annotation.type[0] != 8 or len(vertices) != len(expected)
                        or any(abs(left - right) > 0.001 for a, b in zip(expected, vertices)
                               for left, right in zip(a, b))):
                    raise DocumentIntegrityError("Saved PDF coordinates failed validation")
        ensure_distinct_output(source_path, destination)
        if sha256_file(source_path) != initial_hash:
            raise DocumentIntegrityError("Source changed during PDF annotation")
        os.replace(temporary, destination)
        temporary = None
        return _finish(records, True)
    except DocumentIntegrityError:
        raise
    except Exception:
        # No partial copy is published, including when a later annotation fails.
        for record in records:
            if record["status"] == "applied" or record["reason"] == "not_processed":
                record.update(status="unapplied", reason="annotation_failed")
                record.pop("annotation_xref", None)
        return _finish(records)
    finally:
        if temporary is not None and os.path.exists(temporary):
            os.unlink(temporary)
        if sha256_file(source_path) != initial_hash:
            raise DocumentIntegrityError("Source changed during PDF annotation")
