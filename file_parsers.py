"""
File Parsers Module for Technical Manuals (PDF & DOCX)
Supports Page Range Filtering, Layout Detection, and Structured Extraction
حقوق الملكية الفكرية وتطوير النظام: إسلام عبد البديع
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Tuple, Optional, Dict, Any


def parse_page_ranges(range_str: str, max_pages: int) -> List[int]:
    """
    Parses a page range string such as '1-15, 22-45', '1, 3, 5-8', or 'All'.
    Returns a sorted, unique list of 1-based page numbers.
    Never clips pages if the user explicitly specifies a higher page range than max_pages.
    """
    if not range_str or not range_str.strip():
        return list(range(1, max(1, max_pages) + 1))
    
    clean_str = range_str.strip().lower()
    if clean_str in ["all", "الكل", "جميع الصفحات", "كامل المرجع", "*"]:
        return list(range(1, max(1, max_pages) + 1))
    
    chunks = clean_str.split(",")
    
    # First pass: determine if user requested pages exceeding max_pages
    effective_max = max(1, max_pages)
    for chunk in chunks:
        chunk = chunk.strip()
        m = re.match(r"^(\d+)\s*[-:]\s*(\d+)$", chunk)
        if m:
            effective_max = max(effective_max, int(m.group(1)), int(m.group(2)))
        elif chunk.isdigit():
            effective_max = max(effective_max, int(chunk))
        elif re.match(r"^-\s*(\d+)$", chunk):
            effective_max = max(effective_max, int(re.match(r"^-\s*(\d+)$", chunk).group(1)))
            
    selected_pages = set()
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        
        # Match range like "1-15" or "1:15"
        range_match = re.match(r"^(\d+)\s*[-:]\s*(\d+)$", chunk)
        if range_match:
            start = int(range_match.group(1))
            end = int(range_match.group(2))
            if start > end:
                start, end = end, start
            for p in range(start, end + 1):
                if 1 <= p <= effective_max:
                    selected_pages.add(p)
            continue
        
        # Match open start like "-10" (pages 1 to 10)
        open_start_match = re.match(r"^-\s*(\d+)$", chunk)
        if open_start_match:
            end = min(int(open_start_match.group(1)), effective_max)
            for p in range(1, end + 1):
                selected_pages.add(p)
            continue
        
        # Match open end like "10-" (pages 10 to effective_max)
        open_end_match = re.match(r"^(\d+)\s*-$", chunk)
        if open_end_match:
            start = max(1, int(open_end_match.group(1)))
            for p in range(start, effective_max + 1):
                selected_pages.add(p)
            continue
            
        # Match single page number like "5"
        if chunk.isdigit():
            p = int(chunk)
            if 1 <= p <= effective_max:
                selected_pages.add(p)
    
    # If no valid pages were matched, default to all pages
    if not selected_pages:
        return list(range(1, effective_max + 1))
        
    return sorted(list(selected_pages))


def get_pdf_page_count(file_bytes: bytes) -> int:
    """Returns the total number of pages in a PDF file."""
    # 1. PyMuPDF (fitz) - high precision
    try:
        import pymupdf as fitz
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        count = len(doc)
        doc.close()
        if count > 0:
            return count
    except Exception:
        pass
    
    # 2. Fallback to pypdf
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        count = len(reader.pages)
        if count > 0:
            return count
    except Exception:
        pass
        
    return 1


def extract_text_from_pdf(
    file_bytes: bytes,
    page_range_str: str = "All"
) -> Tuple[str, Dict[str, Any]]:
    """
    Extracts text from PDF bytes filtered by page ranges.
    Returns (extracted_text, metadata_dict).
    """
    total_pages = get_pdf_page_count(file_bytes)
    target_pages = parse_page_ranges(page_range_str, total_pages)
    
    extracted_text_chunks = []
    page_details = []
    
    # Primary: PyMuPDF (fitz)
    used_fitz = False
    try:
        import pymupdf as fitz
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for p_num in target_pages:
            page_idx = p_num - 1  # 0-indexed
            if 0 <= page_idx < len(doc):
                page = doc[page_idx]
                text = page.get_text("text")
                if text and text.strip():
                    cleaned = text.strip()
                    extracted_text_chunks.append(f"--- [صفحة {p_num}] ---\n{cleaned}")
                    page_details.append({"page": p_num, "char_count": len(cleaned)})
        doc.close()
        used_fitz = True
    except Exception:
        used_fitz = False
        
    # Fallback: pypdf if fitz failed or was unavailable
    if not used_fitz:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for p_num in target_pages:
            page_idx = p_num - 1
            if 0 <= page_idx < len(reader.pages):
                page = reader.pages[page_idx]
                text = page.extract_text() or ""
                if text.strip():
                    cleaned = text.strip()
                    extracted_text_chunks.append(f"--- [صفحة {p_num}] ---\n{cleaned}")
                    page_details.append({"page": p_num, "char_count": len(cleaned)})
                    
    combined_text = "\n\n".join(extracted_text_chunks)
    meta = {
        "file_type": "PDF",
        "total_pages": total_pages,
        "selected_pages_count": len(target_pages),
        "selected_pages_list": target_pages,
        "extracted_chars": len(combined_text),
        "page_details": page_details,
    }
    return combined_text, meta


def get_docx_approx_page_count(file_bytes: bytes) -> int:
    """
    Accurately determines page count of a DOCX file:
    Combines:
    1. Word app.xml metadata (<Pages>)
    2. Rendered layout breaks (<w:lastRenderedPageBreak/>)
    3. Manual breaks (<w:br type="page"/>)
    4. Text & table density calculation (~150 words per Arabic technical manual page)
    Always returns the maximum realistic count so stale metadata never under-reports.
    """
    import zipfile
    import xml.etree.ElementTree as ET
    import docx
    
    meta_pages = 0
    xml_breaks = 0
    manual_breaks = 1
    total_words = 0
    total_table_rows = 0

    # 1. Inspect docx zip archive for Word metadata and layout breaks
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            if "docProps/app.xml" in z.namelist():
                try:
                    app_xml = z.read("docProps/app.xml")
                    root = ET.fromstring(app_xml)
                    for elem in root.iter():
                        if elem.tag.endswith("Pages") and elem.text and elem.text.strip().isdigit():
                            val = int(elem.text.strip())
                            if val > 0:
                                meta_pages = val
                                break
                except Exception:
                    pass
            
            if "word/document.xml" in z.namelist():
                try:
                    doc_xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
                    last_rendered = doc_xml.count("<w:lastRenderedPageBreak")
                    explicit_br = doc_xml.count('type="page"')
                    if last_rendered > 0:
                        xml_breaks = last_rendered + 1
                    elif explicit_br > 0:
                        xml_breaks = explicit_br + 1
                except Exception:
                    pass
    except Exception:
        pass

    # 2. Extract words and tables using python-docx
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        for p in doc.paragraphs:
            words = len(p.text.split())
            total_words += words
            for run in p.runs:
                if "w:br" in run._r.xml and 'type="page"' in run._r.xml:
                    manual_breaks += 1
                    
        for table in doc.tables:
            total_table_rows += len(table.rows)
            for row in table.rows:
                for cell in row.cells:
                    total_words += len(cell.text.split())
    except Exception:
        pass

    # Density calculation:
    # 150 words per Arabic technical manual page
    density_pages = max(1, round(total_words / 150) + (total_table_rows // 5))

    # Crucial: Take the maximum realistic count.
    # Stale/default metadata often under-reports (e.g. 15 pages for a 60-page manual).
    return max(1, meta_pages, xml_breaks, manual_breaks, density_pages)


def extract_text_from_docx(
    file_bytes: bytes,
    page_range_str: str = "All"
) -> Tuple[str, Dict[str, Any]]:
    """
    Extracts text from DOCX bytes with structured page/section parsing.
    Guarantees 100% full content extraction with zero truncation when 'All' is selected.
    Returns (extracted_text, metadata_dict).
    """
    import docx
    doc = docx.Document(io.BytesIO(file_bytes))
    
    approx_pages = get_docx_approx_page_count(file_bytes)
    clean_range = (page_range_str or "").strip().lower()
    is_full_doc = clean_range in ["all", "الكل", "كامل المرجع", "جميع الصفحات", "*", ""]
    
    total_words = sum(len(p.text.split()) for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total_words += len(cell.text.split())
                
    words_per_slice = max(100, round(total_words / max(1, approx_pages)))
    
    # Collect all content items in sequential order
    extracted_chunks = []
    page_details = []
    
    current_page_idx = 1
    current_page_lines = []
    current_words = 0
    
    target_pages = parse_page_ranges(page_range_str, approx_pages)
    target_pages_set = set(target_pages)
    
    # Process paragraphs
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
            
        has_break = False
        for run in p.runs:
            if ("w:br" in run._r.xml and 'type="page"' in run._r.xml) or "w:lastRenderedPageBreak" in run._r.xml:
                has_break = True
                break
                
        current_page_lines.append(txt)
        current_words += len(txt.split())
        
        if has_break or current_words >= words_per_slice:
            chunk_str = "\n".join(current_page_lines).strip()
            if chunk_str:
                if is_full_doc or (current_page_idx in target_pages_set):
                    extracted_chunks.append(f"--- [صفحة/قسم {current_page_idx}] ---\n{chunk_str}")
                    page_details.append({"page": current_page_idx, "char_count": len(chunk_str)})
            current_page_idx += 1
            current_page_lines = []
            current_words = 0
            
    # Process tables
    for table in doc.tables:
        table_rows = []
        for row in table.rows:
            row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
            if row_text:
                table_rows.append(row_text)
        if table_rows:
            tbl_txt = "[جدول فني]:\n" + "\n".join(table_rows)
            current_page_lines.append(tbl_txt)
            current_words += len(tbl_txt.split())
            if current_words >= words_per_slice:
                chunk_str = "\n".join(current_page_lines).strip()
                if chunk_str:
                    if is_full_doc or (current_page_idx in target_pages_set):
                        extracted_chunks.append(f"--- [صفحة/قسم {current_page_idx}] ---\n{chunk_str}")
                        page_details.append({"page": current_page_idx, "char_count": len(chunk_str)})
                current_page_idx += 1
                current_page_lines = []
                current_words = 0
                
    # Append any remaining content in the buffer
    if current_page_lines:
        chunk_str = "\n".join(current_page_lines).strip()
        if chunk_str:
            if is_full_doc or (current_page_idx in target_pages_set) or (max(target_pages) >= approx_pages):
                extracted_chunks.append(f"--- [صفحة/قسم {current_page_idx}] ---\n{chunk_str}")
                page_details.append({"page": current_page_idx, "char_count": len(chunk_str)})
                
    combined_text = "\n\n".join(extracted_chunks)
    actual_detected_pages = max(approx_pages, current_page_idx, len(page_details), 1)
    
    meta = {
        "file_type": "DOCX",
        "total_pages": actual_detected_pages,
        "selected_pages_count": len(page_details),
        "selected_pages_list": [d["page"] for d in page_details],
        "extracted_chars": len(combined_text),
        "page_details": page_details,
    }
    return combined_text, meta


def extract_source_document(
    file_name: str,
    file_bytes: bytes,
    page_range_str: str = "All"
) -> Tuple[str, Dict[str, Any]]:
    """
    Unified dispatcher for extracting text from PDF, DOCX, or TXT files.
    """
    lower_name = file_name.lower()
    if lower_name.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes, page_range_str)
    elif lower_name.endswith((".docx", ".doc")):
        return extract_text_from_docx(file_bytes, page_range_str)
    elif lower_name.endswith(".txt"):
        text = file_bytes.decode("utf-8", errors="ignore")
        return text, {
            "file_type": "TXT",
            "total_pages": 1,
            "selected_pages_count": 1,
            "selected_pages_list": [1],
            "extracted_chars": len(text),
            "page_details": [{"page": 1, "char_count": len(text)}]
        }
    else:
        raise ValueError(f"نوع الملف غير مدعوم: {file_name}. الصيغ المدعومة هي PDF و DOCX و TXT.")
