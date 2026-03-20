import os
import json
import uuid
import re
import datetime
import logging
import anthropic

logger = logging.getLogger("dokureview.reviewer")
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH


SEVERITY_COLORS = {
    "red": "FFB3B3",
    "orange": "FFD9A0",
    "green": "C6EFCE",
}

SEVERITY_LABELS = {
    "red": "Kritisch",
    "orange": "Überarbeitungsbedarf",
    "green": "In Ordnung",
}

MODEL = "claude-sonnet-4-5"
MAX_TOKENS = 8096


def highlight_run(run, hex_color):
    """Sets background color of a run via shading."""
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    rPr.append(shd)


def add_comment(doc, paragraph, run, comment_text, author="Document Reviewer"):
    """Adds a Word comment to a run via direct XML manipulation."""
    # Get or create comments part
    comments_part = _get_or_create_comments_part(doc)
    comments_element = comments_part.element

    # Generate a unique comment ID
    existing_ids = [
        int(c.get(qn("w:id"), 0))
        for c in comments_element.findall(qn("w:comment"))
    ]
    comment_id = str(max(existing_ids, default=-1) + 1)

    # Build w:comment element
    comment_elem = OxmlElement("w:comment")
    comment_elem.set(qn("w:id"), comment_id)
    comment_elem.set(qn("w:author"), author)
    comment_elem.set(
        qn("w:date"), datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    )
    comment_elem.set(qn("w:initials"), "DR")

    comment_para = OxmlElement("w:p")
    comment_run = OxmlElement("w:r")
    comment_rpr = OxmlElement("w:rPr")
    comment_run.append(comment_rpr)
    comment_t = OxmlElement("w:t")
    comment_t.text = comment_text
    comment_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    comment_run.append(comment_t)
    comment_para.append(comment_run)
    comment_elem.append(comment_para)
    comments_element.append(comment_elem)

    # Insert commentRangeStart before the run
    range_start = OxmlElement("w:commentRangeStart")
    range_start.set(qn("w:id"), comment_id)
    run._r.addprevious(range_start)

    # Insert commentRangeEnd after the run
    range_end = OxmlElement("w:commentRangeEnd")
    range_end.set(qn("w:id"), comment_id)
    run._r.addnext(range_end)

    # Insert commentReference after rangeEnd
    ref_run = OxmlElement("w:r")
    comment_ref = OxmlElement("w:commentReference")
    comment_ref.set(qn("w:id"), comment_id)
    ref_run.append(comment_ref)
    range_end.addnext(ref_run)


def _get_or_create_comments_part(doc):
    """Gets or creates the comments part of the document."""
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

    part = doc.part

    # Check if comments part already exists
    try:
        for rel in part.rels.values():
            if rel.reltype == REL_TYPE:
                return rel.target_part
    except Exception:
        pass

    # Create new comments part
    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:mv="urn:schemas-microsoft-com:mac:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="mv mo w14 wp14"></w:comments>'
    )

    comments_part = Part(
        PackURI("/word/comments.xml"),
        CT_COMMENTS,
        comments_xml.encode("utf-8"),
        part.package,
    )
    part.relate_to(comments_part, REL_TYPE)
    return comments_part


def call_claude(text, role):
    """Calls Claude API and returns parsed JSON findings."""
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    prompt = f"""Analysiere das folgende Dokument sorgfältig.

Gib deine Analyse als strukturiertes JSON zurück. Keine weiteren Erklärungen, nur valides JSON.

Format:
{{
  "summary": "Kurze Gesamtbewertung (2-4 Sätze)",
  "findings": [
    {{
      "id": 1,
      "severity": "red" | "orange" | "green",
      "quote": "Exakte Textstelle aus dem Dokument (so kurz wie möglich, max. 200 Zeichen)",
      "comment": "Erläuterung des Befundes und Handlungsempfehlung"
    }}
  ]
}}

Severity-Bedeutung:
- red: Kritisches Problem, sofortiger Handlungsbedarf
- orange: Problematisch, sollte überarbeitet werden
- green: Positiv hervorzuheben oder nur geringes Risiko

Dokument:
---
{text}
---"""

    logger.info("Claude-Aufruf: Modell=%s, Rolle='%s', Textlänge=%d Zeichen", MODEL, role.get("name"), len(text))

    for attempt in range(2):
        if attempt > 0:
            logger.warning("Claude-Retry (Versuch %d/2)", attempt + 1)

        response = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=role["system_prompt"],
            messages=[{"role": "user", "content": prompt}],
        )

        raw = response.content[0].text.strip()
        logger.debug("Claude-Antwort (%d Zeichen): %s...", len(raw), raw[:100])

        # Extract JSON: try fenced code block first, then bare JSON object
        json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", raw)
        if json_match:
            raw = json_match.group(1)
        else:
            # Handle truncated responses where closing ``` is missing
            bare_match = re.search(r"(\{[\s\S]*\})", raw)
            if bare_match:
                raw = bare_match.group(1)
                logger.warning("Kein geschlossener Code-Block gefunden, verwende rohen JSON-Block")

        try:
            result = json.loads(raw)
            if "summary" in result and "findings" in result:
                logger.info("Claude-Antwort erfolgreich geparst: %d Findings", len(result["findings"]))
                return result
        except (json.JSONDecodeError, KeyError) as e:
            logger.warning("JSON-Parse-Fehler (Versuch %d): %s", attempt + 1, e)
            if attempt == 1:
                raise ValueError(f"Claude hat kein valides JSON zurückgegeben: {raw[:200]}")
            continue

    raise ValueError("Claude-Analyse fehlgeschlagen")


def find_quote_in_paragraphs(paragraphs, quote):
    """
    Fuzzy search for quote in paragraphs.
    Returns list of (para_index, start_char, end_char) tuples.
    """
    # Normalize quote
    norm_quote = re.sub(r"\s+", " ", quote.strip()).lower()
    if not norm_quote:
        return []

    results = []
    for i, para in enumerate(paragraphs):
        text = para.text
        norm_text = re.sub(r"\s+", " ", text.strip()).lower()
        idx = norm_text.find(norm_quote)
        if idx != -1:
            results.append((i, idx, idx + len(norm_quote)))
            break  # Use first match

    return results


def split_run_at_positions(para, start, end):
    """
    Splits runs in paragraph to isolate characters from start to end.
    Returns the run covering [start, end).
    """
    # Build character map: list of (run_index, char_in_run)
    runs = list(para.runs)
    char_map = []
    for ri, run in enumerate(runs):
        for ci in range(len(run.text)):
            char_map.append((ri, ci))

    if not char_map or start >= len(char_map) or end > len(char_map):
        return None

    # Find which runs we need to split
    start_run_idx, start_char_idx = char_map[start]
    end_run_idx, end_char_idx = char_map[end - 1]

    # Split end run first (to not mess up indices)
    if end_char_idx < len(runs[end_run_idx].text) - 1:
        _split_run(para, end_run_idx, end_char_idx + 1)
        runs = list(para.runs)
        # Rebuild char_map after split
        char_map = []
        for ri, run in enumerate(runs):
            for ci in range(len(run.text)):
                char_map.append((ri, ci))
        start_run_idx, start_char_idx = char_map[start]
        end_run_idx, end_char_idx = char_map[end - 1]

    # Split start run
    if start_char_idx > 0:
        _split_run(para, start_run_idx, start_char_idx)
        runs = list(para.runs)
        char_map = []
        for ri, run in enumerate(runs):
            for ci in range(len(run.text)):
                char_map.append((ri, ci))
        start_run_idx, _ = char_map[start]
        end_run_idx, _ = char_map[end - 1]

    # If the quote spans multiple runs, merge them into one
    runs = list(para.runs)
    if start_run_idx == end_run_idx:
        return runs[start_run_idx]

    # Merge runs from start_run_idx to end_run_idx
    first_run = runs[start_run_idx]
    merged_text = "".join(r.text for r in runs[start_run_idx: end_run_idx + 1])
    first_run.text = merged_text
    for r in runs[start_run_idx + 1: end_run_idx + 1]:
        r._r.getparent().remove(r._r)

    return first_run


def _split_run(para, run_idx, split_pos):
    """Splits a run at split_pos, inserting a new run after it."""
    runs = list(para.runs)
    run = runs[run_idx]
    original_text = run.text
    before = original_text[:split_pos]
    after = original_text[split_pos:]

    run.text = before

    # Create new run with same formatting
    new_r = OxmlElement("w:r")
    # Copy rPr if exists
    rPr = run._r.find(qn("w:rPr"))
    if rPr is not None:
        import copy
        new_r.append(copy.deepcopy(rPr))
    new_t = OxmlElement("w:t")
    new_t.text = after
    new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    new_r.append(new_t)
    run._r.addnext(new_r)


def apply_finding_to_paragraph(doc, out_para, finding, quote):
    """Highlights and comments the quote in the output paragraph."""
    color = SEVERITY_COLORS.get(finding.get("severity", "green"), "C6EFCE")
    comment_text = f"[{SEVERITY_LABELS.get(finding.get('severity','green'), '')}] {finding.get('comment', '')}"

    actual_text = out_para.text

    # Build regex that allows any whitespace between words (handles collapsed spaces)
    words = re.sub(r"\s+", " ", quote.strip()).split()
    if not words:
        return False
    pattern = r"\s+".join(re.escape(w) for w in words)
    m = re.search(pattern, actual_text, re.IGNORECASE)
    if not m:
        return False

    start, end = m.start(), m.end()

    target_run = split_run_at_positions(out_para, start, end)
    if target_run is None:
        return False

    highlight_run(target_run, color)
    try:
        add_comment(doc, out_para, target_run, comment_text)
    except Exception as e:
        logger.warning("Kommentar konnte nicht eingefügt werden: %s", e)

    return True


def copy_paragraph(out_doc, src_para):
    """
    Copies a paragraph with formatting into out_doc, safely.
    Skips relationship-based elements (hyperlinks, drawings, objects) that would
    produce invalid r:id references in the new document and corrupt it.
    """
    import copy as _copy

    out_para = out_doc.add_paragraph()
    # Clear the default empty paragraph element that add_paragraph creates
    for child in list(out_para._p):
        out_para._p.remove(child)

    for child in src_para._p:
        tag = child.tag
        if tag == qn("w:pPr"):
            # Paragraph properties: alignment, indentation, spacing, style, numbering.
            # These contain no r:id relationship references — safe to deep-copy.
            out_para._p.append(_copy.deepcopy(child))
        elif tag == qn("w:r"):
            # Regular run: rPr (bold/italic/font/size/color) has no r:id refs — safe.
            out_para._p.append(_copy.deepcopy(child))
        elif tag == qn("w:hyperlink"):
            # Hyperlink carries an r:id ref we must NOT copy. Extract inner runs only.
            for inner_run in child.findall(qn("w:r")):
                out_para._p.append(_copy.deepcopy(inner_run))
        # Skip: w:drawing, w:object, w:pict (images) and w:bookmarkStart/End,
        # w:proofErr, w:ins, w:del — all either carry r:id refs or are irrelevant.

    return out_para


def add_summary_section(out_doc, summary):
    """Adds the summary block at the beginning of the document."""
    heading = out_doc.add_heading("Review-Zusammenfassung", level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1A, 0x3A, 0x5C)

    para = out_doc.add_paragraph(summary)
    # Light gray background
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F2F2F2")
    pPr.append(shd)

    out_doc.add_paragraph()  # spacer


def add_findings_appendix(out_doc, findings):
    """Appends a findings summary table at the end of the document."""
    out_doc.add_page_break()
    heading = out_doc.add_heading("Anhang: Review-Findings", level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1A, 0x3A, 0x5C)

    table = out_doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"

    # Header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Schweregrad"
    hdr_cells[1].text = "Textstelle"
    hdr_cells[2].text = "Kommentar"

    for cell in hdr_cells:
        for run in cell.paragraphs[0].runs:
            run.bold = True

    # Data rows
    for finding in findings:
        row_cells = table.add_row().cells
        severity = finding.get("severity", "green")
        row_cells[0].text = SEVERITY_LABELS.get(severity, severity)
        row_cells[1].text = finding.get("quote", "")[:200]
        row_cells[2].text = finding.get("comment", "")

        # Color the severity cell
        color = SEVERITY_COLORS.get(severity, "C6EFCE")
        tc = row_cells[0]._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)


def process_document(text, role, output_dir, source_doc=None):
    """
    Main entry point: takes document text and role, returns path to output DOCX.
    If source_doc (python-docx Document) is provided, paragraph formatting is preserved.
    """
    logger.info("Starte Dokumentenverarbeitung: %d Zeichen, Rolle='%s'", len(text), role.get("name"))

    # Call Claude
    analysis = call_claude(text, role)
    summary = analysis.get("summary", "")
    findings = analysis.get("findings", [])
    logger.info("Analyse erhalten: %d Findings", len(findings))

    # Build output DOCX
    out_doc = Document()

    # Add summary at the top
    add_summary_section(out_doc, summary)

    # Divider
    divider = out_doc.add_heading("Dokument", level=2)
    divider.runs[0].font.color.rgb = RGBColor(0x1A, 0x3A, 0x5C)

    unmatched_findings = []
    output_paragraphs = []

    if source_doc is not None:
        # Copy paragraphs with full formatting from the original DOCX
        for src_para in source_doc.paragraphs:
            out_para = copy_paragraph(out_doc, src_para)
            if src_para.text.strip():
                output_paragraphs.append((src_para.text, out_para))
    else:
        # Plain text input: create simple paragraphs
        for para_text in [p.strip() for p in text.split("\n") if p.strip()]:
            out_para = out_doc.add_paragraph(para_text)
            output_paragraphs.append((para_text, out_para))

    # Apply findings
    for finding in findings:
        quote = finding.get("quote", "").strip()
        if not quote:
            continue

        matched = False
        for _para_text, out_para in output_paragraphs:
            if apply_finding_to_paragraph(out_doc, out_para, finding, quote):
                matched = True
                break

        if not matched:
            logger.warning("Quote nicht im Dokument gefunden (Finding #%s): '%s...'",
                           finding.get("id", "?"), quote[:60])
            unmatched_findings.append(finding)

    # Append unmatched findings as footnotes at the end
    if unmatched_findings:
        out_doc.add_page_break()
        fn_heading = out_doc.add_heading("Nicht zugeordnete Befunde", level=2)
        fn_heading.runs[0].font.color.rgb = RGBColor(0x1A, 0x3A, 0x5C)

        for finding in unmatched_findings:
            severity = finding.get("severity", "green")
            color = SEVERITY_COLORS.get(severity, "C6EFCE")
            label = SEVERITY_LABELS.get(severity, severity)

            fn_para = out_doc.add_paragraph()
            fn_para.add_run(f"[{label}] ").bold = True
            fn_para.add_run(f'Textstelle: „{finding.get("quote", "")[:200]}" — ')
            fn_para.add_run(finding.get("comment", ""))

            pPr = fn_para._p.get_or_add_pPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), color)
            pPr.append(shd)

    # Append findings appendix table
    add_findings_appendix(out_doc, findings)

    if unmatched_findings:
        logger.warning("%d Findings konnten nicht zugeordnet werden", len(unmatched_findings))

    # Save output
    output_path = os.path.join(output_dir, f"review_{uuid.uuid4()}.docx")
    out_doc.save(output_path)
    logger.info("Output-DOCX gespeichert: %s", output_path)
    return output_path
