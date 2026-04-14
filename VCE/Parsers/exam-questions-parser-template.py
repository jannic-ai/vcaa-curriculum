#!/usr/bin/env python3
"""
VCE Exam Questions Parser - CONFIG-DRIVEN (PDF extraction)
============================================================

Extracts exam questions from converted (text-layer) PDF exam papers.
Generates a CSV of exam questions linked to exam report and expected
qualities codes.

Supports multiple exam structures:
- Literature: text-based (Section A + B per literary text)
- Flat: no sections (e.g. Accounting - 8 questions)
- Sectioned Written: Section A + B with written questions (BM, Legal Studies)
- MC + Written: Section A multiple-choice + Section B written (Economics)

Usage:
    python exam-questions-parser-template.py <config.json>

Requires:
- Converted PDFs with extractable text (not scanned images)
- PyMuPDF (fitz) for PDF text extraction
- Config must have documents.exam_pdfs dict keyed by year

CSV Output:
- vcaa_vce_sd_{slug}_exam_questions.csv
"""

import os
import re
import csv
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Dict, Optional, Tuple

# Force UTF-8 stdout on Windows to avoid cp1252 encoding errors
# when printing exam text containing Unicode characters (e.g. minus sign U+2212)
if sys.platform == 'win32' and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: PyMuPDF (fitz) is required. Install with: pip install PyMuPDF")
    sys.exit(1)


# ============================================================================
# WINDOWS LONG PATH HELPER
# ============================================================================

def _long_path(p: str) -> str:
    """Prefix with \\\\?\\ on Windows to bypass MAX_PATH (260 char) limit."""
    if sys.platform == "win32" and not p.startswith("\\\\?\\"):
        return "\\\\?\\" + str(Path(p).resolve())
    return p


# ============================================================================
# CONFIG LOADER
# ============================================================================

def load_config(config_path: str) -> Dict:
    """Load subject config from JSON file and resolve paths."""
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # Resolve VCE root directory -- prefer VCE_ROOT env var (set by pipeline worker),
    # fall back to walking up from config file path (works from VCE/_etl/configs/)
    vce_root_env = os.environ.get('VCE_ROOT', '')
    if vce_root_env:
        vce_dir = Path(vce_root_env)
    else:
        config_p = Path(config_path).resolve()
        etl_dir = config_p.parent.parent   # configs/ -> _etl-staging/
        vce_dir = etl_dir.parent           # _etl-staging/ -> tpa-monorepo/ (wrong!)
        # Actually need the VCE root; fallback for standalone use
        # When running from _etl-staging/templates/, walk up differently
        # Best effort: check if parent_folder resolves under vce_dir
        candidate = vce_dir / config['parent_folder']
        if not candidate.is_dir():
            # Try the default VCE_ROOT path
            vce_dir = Path(r"C:\Users\parky\OneDrive\Jannic PA\Education PA\Curriculum\Australian"
                           r"\Victoria\Victoria - VCAA - Senior Years\VCE")

    subject_dir = vce_dir / config['parent_folder']
    docs_dir = subject_dir / "Documentation"

    # Tier 0: Pre-set exam_schedule bypasses all path resolution (used by parser_runner)
    if 'exam_schedule' in config:
        config.setdefault('output_dir', str(subject_dir / "graphRAG_csv" / "Assessment" / "Final Exam" / "Exam"))
        config.setdefault('name', config.get('subject', 'UNKNOWN'))
        config.setdefault('subject_slug', config.get('subject', '').lower().replace(' ', '_'))
        return config

    # Resolve exam PDF paths -- Tier 1: explicit config, Tier 3: auto-detect
    # exam_pdfs supports two formats:
    #   Single-stream: { "2024": "filename.pdf" }
    #   Multi-stream:  { "2024": { "ANC": "ancient.pdf", "AUS": "australian.pdf" } }
    exam_pdfs = config.get('documents', {}).get('exam_pdfs', {})
    exam_schedule = []  # list of (year, stream_code_or_None, pdf_path)

    def _filename_from(value: str) -> str:
        """Accept either a bare filename or a GitHub URL; return just the basename.

        The pipeline worker downloads URL-valued entries into docs_dir before
        invoking us, so we only ever need the filename relative to that dir.
        """
        from urllib.parse import unquote
        if value.startswith(("http://", "https://")):
            return unquote(value.rsplit("/", 1)[-1])
        return value

    if exam_pdfs:
        for year in sorted(exam_pdfs.keys()):
            value = exam_pdfs[year]
            if isinstance(value, str):
                # Single-stream: value is a filename or URL
                exam_schedule.append((year, None, str(docs_dir / _filename_from(value))))
            elif isinstance(value, dict):
                # Multi-stream: value is { stream_code: filename|url }
                for stream_code in sorted(value.keys()):
                    exam_schedule.append((year, stream_code, str(docs_dir / _filename_from(value[stream_code]))))
    else:
        # Tier 3: Auto-detect *-exam.pdf in documentation folder
        streams_config = config.get('streams', {})
        if docs_dir.is_dir():
            for match in sorted(docs_dir.glob("*-exam.pdf")):
                if match.name.startswith("~$"):
                    continue
                year_match = re.search(r'(\d{4})', match.name)
                if not year_match:
                    continue
                year = year_match.group(1)
                # Try to match stream from filename against streams config
                stream_code = None
                if streams_config:
                    slug = config.get('subject_slug', '')
                    # Extract the part between subject slug and '-exam.pdf'
                    # e.g. "2024-vce-history-ancient-exam.pdf" -> "ancient"
                    name_part = match.stem  # e.g. "2024-vce-history-ancient-exam"
                    name_part = re.sub(r'^\d{4}-vce-', '', name_part)
                    name_part = re.sub(r'-exam$', '', name_part)
                    name_part = re.sub(rf'^{re.escape(slug)}-?', '', name_part)
                    # Match against stream names (case-insensitive slug comparison)
                    for code, name_or_dict in streams_config.items():
                        stream_name = name_or_dict if isinstance(name_or_dict, str) else name_or_dict.get('name', '')
                        # Extract last word(s) of stream name for slug matching
                        stream_slug = stream_name.lower().split(':')[-1].strip().split()[-1]
                        if name_part.lower() == stream_slug or code.lower() in name_part.lower():
                            stream_code = code
                            break
                exam_schedule.append((year, stream_code, str(match)))
            if exam_schedule:
                print(f"  Auto-detected exam PDFs: {[(y, sc) for y, sc, _ in exam_schedule]}")

    config['exam_schedule'] = exam_schedule
    config['output_dir'] = str(subject_dir / "graphRAG_csv" / "Assessment" / "Final Exam" / "Exam")

    config.setdefault('name', config.get('subject', 'UNKNOWN'))
    config.setdefault('subject_slug', config.get('subject', '').lower().replace(' ', '_'))

    return config


# ============================================================================
# UNICODE NORMALIZATION
# ============================================================================

def fix_mojibake(text: str) -> str:
    """Fix double-encoded UTF-8 text (mojibake).

    Detects patterns like â€™ (should be ') where UTF-8 bytes were
    decoded as Windows-1252. Re-encodes to CP1252 bytes then decodes as UTF-8.
    """
    if not text or '\u00e2\u20ac' not in text:
        return text
    try:
        return text.encode('cp1252').decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        return text


def normalize_unicode(text: str) -> str:
    """Normalize smart punctuation from DOCX/PDF to ASCII equivalents.

    IMPORTANT: All parsers must apply this to text extracted from documents.
    PDF files may contain smart quotes, em/en dashes, and other Unicode
    punctuation that displays as mojibake when tools read CSV files with
    wrong encoding. Normalizing to ASCII prevents this.
    """
    if not text:
        return ""
    text = fix_mojibake(text)
    replacements = {
        '\u2018': "'",   # left single quotation mark
        '\u2019': "'",   # right single quotation mark (apostrophe)
        '\u201C': '"',   # left double quotation mark
        '\u201D': '"',   # right double quotation mark
        '\u2013': '-',   # en dash
        '\u2014': '-',   # em dash
        '\u2026': '...',  # horizontal ellipsis
        '\u00A0': ' ',   # non-breaking space
        '\u2009': ' ',   # thin space
        '\u2003': ' ',   # em space
        '\u2002': ' ',   # en space
        '\u2008': ' ',   # punctuation space
        '\u00AD': '',    # soft hyphen
        '\u200B': '',    # zero-width space
        '\uFB01': 'fi',  # latin small ligature fi
        '\uFB02': 'fl',  # latin small ligature fl
        '\uFB00': 'ff',  # latin small ligature ff
        '\uFB03': 'ffi', # latin small ligature ffi
        '\uFB04': 'ffl', # latin small ligature ffl
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    return text


def sanitize_rows(rows):
    """Apply normalize_unicode to all string values in rows before CSV write."""
    for row in rows:
        if isinstance(row, dict):
            for key, val in row.items():
                if isinstance(val, str):
                    row[key] = normalize_unicode(val)
        elif isinstance(row, list):
            for i, val in enumerate(row):
                if isinstance(val, str):
                    row[i] = normalize_unicode(val)
    return rows


# ============================================================================
# TEXT CLEANING HELPERS
# ============================================================================

# Regex patterns for OCR noise filtering
PAGE_HEADER_RE = re.compile(
    r'^(Page\s+\d+\s+of\s+\d+|'
    r'\d{4}\s+VCE\s+\w[\w\s]*|'
    r'Question\s+(and\s+Answer\s+)?Book\s+\d{4}|'
    r'Section\s+[A-Z]\s+\d{4}|'
    r'\d{4}\s+VCE\s+[\w\s]+Section\s+[A-Z]|'
    r'Section\s+[A-Z]\s+Page\s+\d+)',
    re.IGNORECASE
)
NOISE_RE = re.compile(
    r'^(Do\s+not\s+write\s+in\s+this\s+area|'
    r'SUPERVISOR\s+TO\s+ATTACH|'
    r'Write\s+your\s+student\s+number|'
    r'This\s+page\s+is\s+blank|'
    r'Victorian\s+Curri|'   # relaxed prefix to handle OCR garbling
    r'VCAA\s+\d{4}|'
    r'End\s+of\s+(Answer|Question)\s+Book|'
    r'End\s+of\s+examination)',
    re.IGNORECASE
)

# OCR form-field garbage: sequences of single characters separated by spaces
# that represent empty form boxes (e.g. "o T [ T [ [ T T o[l =]")
FORM_GARBAGE_RE = re.compile(r'^[\s\[\](){}|oOTIBREN=\-\+\*\.,;:!?\d<>@&#~`^%$/\\]{3,}$')


def _is_garbled_line(line: str) -> bool:
    """Detect garbled lines from PDF figure/image extraction.

    PDF pages with maps, charts, and images produce garbage text when the
    text extractor tries to read image content. These lines have:
    - Very low ratio of alphabetic characters to total non-space chars
    - Mostly single-character 'words' (fragmented OCR noise)
    - Dense sequences of symbols and punctuation

    Returns True if the line is likely garbled and should be removed.
    """
    if not line or len(line) < 4:
        return False
    # Don't filter sub-question labels like "a.", "b."
    if re.match(r'^[a-z]\.\s', line):
        return False
    # Don't filter Question headers
    if re.match(r'^Question\s+\d', line, re.IGNORECASE):
        return False
    # Don't filter Figure/Table captions
    if re.match(r'^(Figure|Table|Source|Note)\s+', line, re.IGNORECASE):
        return False

    non_space = line.replace(' ', '')
    if len(non_space) < 4:
        return False

    alpha_count = sum(1 for c in non_space if c.isalpha())
    alpha_ratio = alpha_count / len(non_space)

    # Lines with < 40% alphabetic characters are likely garbled
    if alpha_ratio < 0.40:
        return True

    # Check for fragmented text: many single-char "words"
    words = line.split()
    if len(words) >= 4:
        single_chars = sum(1 for w in words if len(w) == 1)
        if single_chars / len(words) > 0.5:
            return True

    return False


def clean_page_text(text: str) -> str:
    """Remove page headers, footers, and OCR noise from extracted PDF text.

    PDF answer-book pages often contain form-field artifacts in the left margin:
    single characters, symbols, non-ASCII glyphs (Â£, 8Â», etc.) that appear as
    leading garbage on content lines. These are stripped to expose the real text.
    """
    lines = text.split('\n')
    cleaned = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if PAGE_HEADER_RE.match(stripped):
            continue
        if NOISE_RE.match(stripped):
            continue
        if FORM_GARBAGE_RE.match(stripped):
            continue
        if _is_garbled_line(stripped):
            continue
        # Skip lines that are just 1-2 characters (form artifacts, margin marks)
        # ONLY preserve sub-question labels like "a.", "b.", etc.
        if len(stripped) <= 2:
            if not re.match(r'^[a-z]\.$', stripped):
                continue
        # Strip leading OCR garbage tokens (1-3 chars) before Question headers
        # e.g. "s Question 10" → "Question 10", "ES Question 15" → "Question 15"
        stripped = re.sub(r'^\S{1,3}\s+(?=Question\s+\d)', '', stripped)
        # Strip leading OCR garbage tokens before sub-question labels
        # e.g. "_g a. Complete..." → "a. Complete...", "H5 b. Prepare..." → "b. Prepare..."
        # e.g. "Â£ a. Discuss..." → "a. Discuss..."
        stripped = re.sub(r'^\S{1,3}\s+(?=[a-z]\.\s)', '', stripped)
        # Strip leading tokens containing non-ASCII characters (form-field artifacts)
        # e.g. "8Â» Credit Sales..." → "Credit Sales...", "Â£ Allowance..." → "Allowance..."
        # Only strip if what follows looks like real content (starts with uppercase or digit)
        stripped = re.sub(r'^\S*[^\x00-\x7F]\S*\s+(?=[A-Z0-9])', '', stripped)
        # Strip single ASCII form-field markers before content lines
        # e.g. "& 31 October..." → "31 October...", "8 On 6 June..." → "On 6 June..."
        stripped = re.sub(r'^[&+*#~^\\|]\s+', '', stripped)
        # Strip underscore-prefixed form artifacts (e.g. "_g changing..." → "changing...")
        stripped = re.sub(r'^_\S{0,2}\s+', '', stripped)
        # Strip single-letter form artifacts before continuation text
        # e.g. "g Narrations..." → "Narrations...", "E The following..." → "The following..."
        # Only if followed by an uppercase letter (real sentence continuation)
        stripped = re.sub(r'^[a-zA-Z]\s+(?=[A-Z])', '', stripped)
        # Strip single digit followed by space before content (form line numbers)
        # e.g. "8 On 6 June..." → "On 6 June...", "3 The..." → "The..."
        stripped = re.sub(r'^\d\s+(?=[A-Z][a-z]+)', '', stripped)
        # Strip trailing form-field garbage: single chars/symbols at end of line
        # PDF answer-book pages have right-margin form artifacts too:
        # "period. g", "3 marks &", "equipment. P", "provided. =3"
        # Strategy: strip trailing whitespace + 1-3 char garbage tokens after
        # content that ends with punctuation, "marks", or a word.
        stripped = re.sub(r'\s+\S?[^\x00-\x7F]\S*$', '', stripped)  # trailing non-ASCII
        stripped = re.sub(r'\s+[=%;()\[\]{}><&#|\\/*+~^]+\S{0,2}$', '', stripped)  # trailing symbols (incl. =3)
        # Trailing single letter/digit after sentence-ending punctuation
        # e.g. "period. g" → "period.", "100%. 3" → "100%.", "marks o" → "marks"
        stripped = re.sub(r'(?<=[.!?:,])\s+[a-zA-Z\d]{1,2}$', '', stripped)
        # Trailing single letter after "marks" (e.g. "4 marks o" → "4 marks")
        stripped = re.sub(r'(?<=marks)\s+[a-zA-Z]{1,2}$', '', stripped)
        # Trailing single letter after content (form-field margin marker)
        # e.g. "equipment. P" → "equipment.", "Receivable o" → "Receivable",
        #      "$20000 g" → "$20000"
        stripped = re.sub(r'(?<=[a-z0-9])\s+[A-Za-z]$', '', stripped)
        # Trailing single digit after text (form artifact, not part of content)
        # e.g. "made by the 3" → "made by the" (the 3 is a margin number)
        # Only strip if preceded by a letter (not another digit — "30000" is real)
        # IMPORTANT: Protect "Question N" headers — single-digit question numbers
        # must NOT be stripped. Use negative lookbehind for "Question".
        stripped = re.sub(r'(?<!Question)(?<=[a-zA-Z])\s+\d$', '', stripped)
        cleaned.append(stripped)
    return '\n'.join(cleaned)


def clean_question_text(text: str) -> str:
    """Clean extracted question text: collapse whitespace, normalize unicode."""
    text = normalize_unicode(text)
    # Collapse multiple spaces
    text = re.sub(r'  +', ' ', text)
    # Remove "Question N continues on the next page." artifacts
    text = re.sub(r'Question\s+\d+\s+continues\s+on\s+the\s+next\s+page\.?', '', text, flags=re.IGNORECASE)
    return text.strip()


def _split_context_segments(text: str) -> List[Tuple[str, str]]:
    """Split question context into typed segments for multi-row output.

    Returns [(content_type, text), ...] where content_type is one of:
    - 'Question' for paragraph/instruction text
    - 'QuestionOption' for numbered or checkbox option lists
    - 'QuestionHeader' for standalone short header lines (e.g. "Planning")

    Groups consecutive lines of the same type together.
    """
    if not text or not text.strip():
        return []

    lines = text.split('\n')
    # Classify each line
    CHECKBOX_RE = re.compile(r'^[\u25a1\u25a0\u2610\u2611\u2612\u2713\u2717\u00a8\u2022]\s+')  # □ ■ ☐ ☑ ☒ ✓ ✗ ¨ •
    NUMBERED_OPT_RE = re.compile(r'^\d+\.\s+\S')  # "1. actor", "2. director"

    classified: List[Tuple[str, str]] = []  # (type, line)
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if CHECKBOX_RE.match(stripped) or NUMBERED_OPT_RE.match(stripped):
            classified.append(('QuestionOption', stripped))
        elif len(stripped.split()) <= 3 and stripped[0].isupper() and not stripped.endswith('.'):
            # Short capitalised line without full stop → likely a header
            classified.append(('QuestionHeader', stripped))
        else:
            classified.append(('Question', stripped))

    # Group consecutive lines of the same type
    if not classified:
        return []
    segments: List[Tuple[str, str]] = []
    current_type = classified[0][0]
    current_lines = [classified[0][1]]
    for ctype, cline in classified[1:]:
        if ctype == current_type:
            current_lines.append(cline)
        else:
            segments.append((current_type, '\n'.join(current_lines)))
            current_type = ctype
            current_lines = [cline]
    segments.append((current_type, '\n'.join(current_lines)))
    return segments


def _is_real_mid_context(text: str) -> bool:
    """Check if trailing text is real mid-question context (not form/answer noise).

    Real context is prose providing additional information for subsequent sub-questions
    (e.g. "The following information is also provided..." in Accounting, RBA policy
    quotes in Economics). Noise includes form-field garbage, answer templates, blank
    manuscript areas, and page footer text from combined Q&A book PDFs.
    """
    if not text or len(text) < 30:
        return False
    # Must have enough alphabetic characters (> 60% of non-whitespace)
    non_ws = re.sub(r'\s', '', text)
    if not non_ws:
        return False
    alpha_ratio = sum(1 for c in non_ws if c.isalpha()) / len(non_ws)
    if alpha_ratio < 0.6:
        return False
    # Must have enough real words (8+ words of 2+ alpha chars = prose, not template)
    words = [w for w in text.split() if len(w) >= 2 and w[0].isalpha()]
    if len(words) < 8:
        return False
    # Filter known noise patterns
    if re.search(
        r'Blank\s+manuscript|End\s+of\s+(Question|Section|examination)|'
        r'Victorian\s+Curri|This\s+page\s+is\s+blank|[©□]|'
        r'Examination\s+continues\s+on|silent\s+working\s+time|'
        r'Select\s+one\s*\(',
        text, re.IGNORECASE):
        return False
    return True


# ============================================================================
# DATA STRUCTURES
# ============================================================================

@dataclass
class SectionInfo:
    key: str              # "A", "B", "ALL"
    name: str
    section_type: str     # "written", "multiple_choice", "case_study", "flat"
    start_page: int       # 0-indexed inclusive
    end_page: int         # 0-indexed exclusive
    unit_as_code: str = ""
    er_code_base: str = ""
    expected_marks: int = 0   # from cover page (0 = unknown)


@dataclass
class ExamStructure:
    sections: List[SectionInfo]
    qb_end_page: int      # page where Question Book ends (-1 = parse all)
    total_pages: int
    expected_total_marks: int = 0  # from cover page (0 = unknown)
    section_choice: Dict[str, Tuple[int, int]] = field(default_factory=dict)  # "N of M" questions


@dataclass
class SubQuestion:
    letter: str           # "a", "b", "c"
    text: str
    marks: int
    trailing_context: str = ""  # mid-question context after this sub-q (before next)


@dataclass
class Question:
    number: int
    total_marks: int
    context: str          # scenario/preamble text before sub-questions
    sub_questions: List[SubQuestion] = field(default_factory=list)
    content_type: str = "Question"  # "Question" or "MultipleChoice"


# ============================================================================
# PASS 1: DETECT EXAM STRUCTURE
# ============================================================================

# Section header: must be at start of line, uppercase letter A-E only.
# Optionally followed by dash + name (e.g. "Section B - Case study") or standalone ("Section A").
# NOT case-insensitive — avoids matching "section" in question text.
SECTION_RE = re.compile(r'^Section\s+([A-E])(?:\s*[-\u2013\u2014]\s*(.*))?\s*$', re.MULTILINE)
QUESTION_HEADER_RE = re.compile(r'Question\s+(\d+)\s*\((\d+)\s*marks?\)', re.IGNORECASE)
MC_OPTION_RE = re.compile(r'^[A-D]\.\s+', re.MULTILINE)
ANSWER_BOOK_RE = re.compile(r'Answer\s+Book', re.IGNORECASE)
END_SECTION_RE = re.compile(r'End\s+of\s+Section\s+([A-Z])', re.IGNORECASE)


def detect_exam_structure(pdf_path: str, config: Dict) -> ExamStructure:
    """Pass 1: Scan PDF to determine section layout and question book boundary."""
    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    # Detect Question Book / Answer Book boundary
    qb_end_page = total_pages  # default: parse all pages
    for i in range(total_pages):
        text = doc[i].get_text()
        # Look for "Answer Book" as a header (not in "Question and Answer Book")
        # Accounting has separate QB + AB; others have combined "Question and Answer Book"
        first_200 = text[:200]
        if 'Answer Book' in first_200 and 'Question' not in first_200:
            qb_end_page = i
            print(f"    Answer Book detected at page {i} - parsing pages 0-{i-1}")
            break

    # Also check page 0 for "Question Book of N pages" pattern
    page0_text = doc[0].get_text()
    qb_pages_match = re.search(r'Question\s+Book\s+of\s+(\d+)\s+pages?', page0_text, re.IGNORECASE)
    if qb_pages_match:
        qb_page_count = int(qb_pages_match.group(1))
        # QB starts at page 0 (cover), so actual content pages are 0..(qb_page_count-1)
        # But page numbers in the PDF are 0-indexed, and the cover is page 0
        # "Question Book of 12 pages" means pages 0-11 in the PDF
        if qb_page_count < total_pages:
            qb_end_page = qb_page_count
            print(f"    Question Book: {qb_page_count} pages (parsing pages 0-{qb_page_count-1})")

    # Detect section headers and their page positions
    section_pages: Dict[str, Tuple[int, str]] = {}  # key -> (page, name)
    section_ends: Dict[str, int] = {}  # key -> page of "End of Section X"

    for i in range(min(qb_end_page, total_pages)):
        text = doc[i].get_text()

        # Check for section headers
        for m in SECTION_RE.finditer(text):
            section_key = m.group(1).upper()
            section_name = (m.group(2) or '').strip()
            # Only take the first occurrence of each section
            if section_key not in section_pages:
                # Skip if this is in the contents/TOC on page 0
                if i == 0:
                    continue
                # Skip if "Section X" appears deep in the page text — likely a
                # reference within question content (e.g. musical structure maps
                # in Music exams: "Section A ... Section B"). Real exam section
                # headers appear near the top of their page.
                if m.start() > 300:
                    continue
                section_pages[section_key] = (i, section_name)

        # Check for "End of Section X"
        for m in END_SECTION_RE.finditer(text):
            section_key = m.group(1).upper()
            section_ends[section_key] = i

    # Extract expected marks from cover page (page 0)
    # Patterns: "Section A (N questions, M marks)", "N questions (M marks)"
    # OCR often garbles: "1%5marks" = 15, "4Omarks" = 40, "toOmarks" = 100
    cover_section_marks: Dict[str, int] = {}
    cover_section_choice: Dict[str, Tuple[int, int]] = {}  # key -> (answer_n, of_m)
    cover_total_marks = 0

    def _ocr_fix_marks(raw: str) -> int:
        """Fix common OCR garbling in marks numbers."""
        raw = raw.replace('%', '').replace('O', '0').replace('o', '0')
        raw = raw.replace('l', '1').replace('I', '1').replace('t', '1')
        raw = re.sub(r'[^0-9]', '', raw)
        return int(raw) if raw else 0

    for line in page0_text.split('\n'):
        s = line.strip()
        # Per-section: "Section A (N questions, Mmarks)" or "Section A (N of M questions, Mmarks)"
        sec_m = re.search(r'Section\s+([A-E])\s*\([^)]*?,\s*(\S+?)\s*marks\s*\)', s, re.IGNORECASE)
        if sec_m:
            cover_section_marks[sec_m.group(1).upper()] = _ocr_fix_marks(sec_m.group(2))
            # Detect "N of M questions" (student chooses N of M available)
            choice_m = re.search(r'(\d+)\s+of\s+(\d+)\s+questions?', s, re.IGNORECASE)
            if choice_m:
                cover_section_choice[sec_m.group(1).upper()] = (
                    int(choice_m.group(1)), int(choice_m.group(2)))
        # Flat exam total: "N questions (Mmarks)" without "Section"
        flat_m = re.search(r'questions?\s*\((\S+?)\s*marks\s*\)', s, re.IGNORECASE)
        if flat_m and 'Section' not in s:
            cover_total_marks = _ocr_fix_marks(flat_m.group(1))

    if cover_section_marks:
        cover_total_marks = sum(cover_section_marks.values())

    doc.close()

    # Build section info from config + detected pages
    config_sections = config.get('exam_sections', {})

    if not section_pages:
        # Flat exam (no sections detected) - e.g. Accounting
        section_key = 'ALL' if 'ALL' in config_sections else 'ALL'
        sec_config = config_sections.get('ALL', config_sections.get(section_key, {}))
        sections = [SectionInfo(
            key='ALL',
            name=sec_config.get('name', 'All questions'),
            section_type='flat',
            start_page=0,
            end_page=qb_end_page,
            unit_as_code=sec_config.get('unit_as_code', ''),
            er_code_base=sec_config.get('exam_report_code', ''),
            expected_marks=cover_total_marks,
        )]
        print(f"    Structure: flat (no sections), {qb_end_page} QB pages")
    else:
        # Sectioned exam
        sorted_keys = sorted(section_pages.keys())
        sections = []
        for idx, key in enumerate(sorted_keys):
            page, detected_name = section_pages[key]
            sec_config = config_sections.get(key, {})

            # Determine section name (prefer config, fall back to detected)
            name = sec_config.get('name', detected_name or f'Section {key}')

            # Determine section type
            section_type = sec_config.get('section_type', 'written')
            if section_type == 'written' and 'case study' in name.lower():
                section_type = 'case_study'

            # Determine end page: next section start or QB end
            if key in section_ends:
                end_page = section_ends[key] + 1  # include the "End of" page
            elif idx + 1 < len(sorted_keys):
                end_page = section_pages[sorted_keys[idx + 1]][0]
            else:
                end_page = qb_end_page

            sections.append(SectionInfo(
                key=key,
                name=name,
                section_type=section_type,
                start_page=page,
                end_page=end_page,
                unit_as_code=sec_config.get('unit_as_code', ''),
                er_code_base=sec_config.get('exam_report_code', ''),
                expected_marks=cover_section_marks.get(key, 0),
            ))
            print(f"    Section {key}: '{name}' (type={section_type}, pages {page}-{end_page-1})")

    return ExamStructure(sections=sections, qb_end_page=qb_end_page, total_pages=total_pages,
                         expected_total_marks=cover_total_marks,
                         section_choice=cover_section_choice)


# ============================================================================
# PASS 2: GENERIC EXAM EXTRACTION
# ============================================================================

# Sub-question pattern: letter followed by period at start of line
# e.g. "a. Describe one way..." or "b. Explain two restraining forces..."
SUBQ_START_RE = re.compile(r'^([a-z])\.\s+', re.MULTILINE)

# Marks at end of a line: "3 marks" or "1 mark"
MARKS_RE = re.compile(r'(\d+)\s+marks?\s*$', re.IGNORECASE | re.MULTILINE)

# MC question header (no marks in parentheses)
MC_QUESTION_RE = re.compile(r'^Question\s+(\d+)\s*$', re.MULTILINE)


def extract_section_text(pdf_path: str, start_page: int, end_page: int) -> str:
    """Extract and clean text from a range of PDF pages."""
    doc = fitz.open(pdf_path)
    all_text = []
    for i in range(start_page, min(end_page, len(doc))):
        page_text = doc[i].get_text()
        cleaned = clean_page_text(page_text)
        if cleaned.strip():
            all_text.append(cleaned)
    doc.close()
    return '\n'.join(all_text)


def extract_written_questions(text: str) -> List[Question]:
    """Extract written questions with sub-parts from section text.

    Handles:
    - Question N (M marks) headers
    - Sub-questions: a. text N marks, b. text N marks
    - Questions without sub-parts (single block)
    - Multi-page questions ("continues on the next page")
    """
    questions = []

    # Preprocess: join sub-question labels on their own line with the next line
    # e.g. "a.\nPrepare the..." → "a. Prepare the..."
    text = re.sub(r'^([a-z])\.\s*\n', r'\1. ', text, flags=re.MULTILINE)

    # Find all question headers with marks
    q_matches = list(QUESTION_HEADER_RE.finditer(text))
    if not q_matches:
        return questions

    for qi, q_match in enumerate(q_matches):
        q_num = int(q_match.group(1))
        q_total_marks = int(q_match.group(2))

        # Get text block for this question (until next question or end)
        start = q_match.end()
        end = q_matches[qi + 1].start() if qi + 1 < len(q_matches) else len(text)
        q_text = text[start:end].strip()

        # Clean the question text
        q_text = clean_question_text(q_text)

        # Find sub-questions
        sub_matches = list(SUBQ_START_RE.finditer(q_text))

        if sub_matches:
            # Context is everything before the first sub-question
            context = q_text[:sub_matches[0].start()].strip()

            sub_questions = []
            for si, s_match in enumerate(sub_matches):
                letter = s_match.group(1)
                s_start = s_match.end()
                s_end = sub_matches[si + 1].start() if si + 1 < len(sub_matches) else len(q_text)
                s_text = q_text[s_start:s_end].strip()

                # Extract marks from sub-question text.
                # Some sub-questions have roman numeral sub-parts (i, ii, iii)
                # each with their own "N mark(s)" — sum all matches in that case.
                all_marks = list(MARKS_RE.finditer(s_text))
                trailing_context = ""
                if len(all_marks) > 1:
                    # Multiple marks entries (e.g. i. 1 mark, ii. 1 mark, iii. 1 mark)
                    marks = sum(int(m.group(1)) for m in all_marks)
                    trailing_raw = s_text[all_marks[-1].end():].strip()
                    s_text = s_text[:all_marks[0].start()].strip()
                elif all_marks:
                    marks = int(all_marks[0].group(1))
                    trailing_raw = s_text[all_marks[0].end():].strip()
                    s_text = s_text[:all_marks[0].start()].strip()
                else:
                    marks = 0
                    trailing_raw = ""

                # Detect mid-question context: substantial text after marks line
                # and before the next sub-question (e.g. "The following information
                # is also provided..." in Accounting, RBA extracts in Economics)
                if trailing_raw and len(trailing_raw) > 30:
                    candidate = clean_question_text(trailing_raw)
                    if _is_real_mid_context(candidate):
                        trailing_context = candidate

                sub_questions.append(SubQuestion(
                    letter=letter, text=s_text, marks=marks,
                    trailing_context=trailing_context))

            questions.append(Question(
                number=q_num,
                total_marks=q_total_marks,
                context=context,
                sub_questions=sub_questions,
            ))
        else:
            # No sub-questions - single question block
            # Remove marks from the end if present
            q_clean = q_text
            marks_match = MARKS_RE.search(q_clean)
            if marks_match:
                q_clean = q_clean[:marks_match.start()].strip()

            questions.append(Question(
                number=q_num,
                total_marks=q_total_marks,
                context=q_clean,
                sub_questions=[],
            ))

    return questions


def extract_mc_questions(text: str) -> List[Question]:
    """Extract multiple-choice questions from section text.

    Each MC question has:
    - Question N header (no marks - all 1 mark each)
    - Question text
    - Options A. B. C. D.
    """
    questions = []

    # Find MC question headers: "Question N" at start of line (may have trailing noise from OCR)
    q_pattern = re.compile(r'^Question\s+(\d+)\b', re.MULTILINE)
    q_matches = list(q_pattern.finditer(text))

    for qi, q_match in enumerate(q_matches):
        q_num = int(q_match.group(1))
        q_marks = 1  # MC questions are 1 mark each

        start = q_match.end()
        end = q_matches[qi + 1].start() if qi + 1 < len(q_matches) else len(text)
        q_text = text[start:end].strip()

        # Clean
        q_text = clean_question_text(q_text)

        # Split into question stem and options
        # Find first option "A." at start of line
        option_start = re.search(r'^A\.\s+', q_text, re.MULTILINE)
        if option_start:
            stem = q_text[:option_start.start()].strip()
            options_text = q_text[option_start.start():].strip()

            # Parse individual options
            option_matches = list(re.finditer(r'^([A-D])\.\s+', options_text, re.MULTILINE))
            options = []
            for oi, o_match in enumerate(option_matches):
                o_start = o_match.end()
                o_end = option_matches[oi + 1].start() if oi + 1 < len(option_matches) else len(options_text)
                o_text = options_text[o_start:o_end].strip()
                options.append(f"{o_match.group(1)}. {o_text}")

            content = stem + '\n' + ' | '.join(options) if options else q_text
        else:
            content = q_text

        questions.append(Question(
            number=q_num,
            total_marks=q_marks,
            context=content,
            sub_questions=[],
            content_type='MultipleChoice',
        ))

    return questions


def extract_case_study_preamble(text: str) -> str:
    """Extract case study preamble text before the first question.

    For BM Section B, the case study scenario appears between "Case study"
    header and Question 1.
    """
    # Find first "Question N" header
    q_match = QUESTION_HEADER_RE.search(text)
    if not q_match:
        q_match = re.search(r'^Question\s+\d+', text, re.MULTILINE)
    if not q_match:
        return ""

    preamble = text[:q_match.start()].strip()

    # Look for "Case study" header as start marker
    cs_match = re.search(r'^Case\s+study\s*$', preamble, re.MULTILINE | re.IGNORECASE)
    if cs_match:
        preamble = preamble[cs_match.end():].strip()
    else:
        # Fallback: strip section header, instructions, and bullet lines
        lines = preamble.split('\n')
        filtered = []
        past_instructions = False
        for line in lines:
            stripped = line.strip()
            if re.match(r'^(Section\s+[A-Z]|Instructions)', stripped, re.IGNORECASE):
                past_instructions = False
                continue
            if re.match(r'^[+\-*\u2022]\s+', stripped):
                continue  # Skip bullet-prefixed instruction lines
            if stripped:
                filtered.append(stripped)
        preamble = ' '.join(filtered).strip()

    return clean_question_text(preamble)


def extract_generic_exam(pdf_path: str, year: str, config: Dict, structure: ExamStructure,
                         stream_code: Optional[str] = None, stream_name: str = '') -> List[List]:
    """Pass 2: Extract questions from all sections and build CSV rows."""
    subject_code = config['subject_code']
    yy = year[-2:]
    rows = []
    # For multi-stream subjects, insert stream code into all codes
    stream_prefix = stream_code or ''

    meta_base = [
        config.get('subject_area', ''),
        config['name'],
        config.get('_csv_stream_code', stream_code or ''),  # SubjectStreamCode (config fallback)
        config.get('_csv_stream_name', stream_name),        # SubjectStreamName (config fallback)
        'Year 12',
        'Final Exam',
        'Exam Questions',
        year,
    ]

    sections_questions: Dict[str, List[Question]] = {}  # for sanity check

    for section in structure.sections:
        print(f"    Extracting section {section.key}: {section.name} (type={section.section_type})")

        # Extract text for this section's page range
        text = extract_section_text(pdf_path, section.start_page, section.end_page)
        if not text.strip():
            print(f"      WARNING: No text extracted for section {section.key}")
            continue

        # Build codes
        section_code = section.key  # "A", "B", or "ALL" for flat exams
        unit_as = section.unit_as_code

        # ExamReportCode: year-encode the base code
        # For multi-stream, stream code goes between subject code and year
        if section.er_code_base:
            # Insert stream code into base before year-encoding
            if stream_prefix:
                base = section.er_code_base
                # Insert stream code after subject code prefix (VCE + subject_code)
                prefix_len = 3 + len(subject_code)  # "VCE" + subject_code
                base = base[:prefix_len] + stream_prefix + base[prefix_len:]
                er_code = _year_encode_er_code(base, subject_code, yy)
            else:
                er_code = _year_encode_er_code(section.er_code_base, subject_code, yy)
        elif section_code:
            er_code = f'VCE{subject_code}{stream_prefix}{yy}ER{section_code}'
        else:
            er_code = f'VCE{subject_code}{stream_prefix}{yy}ER'

        seq = 1

        # Section header row — include total marks so teacher assistant can retrieve it
        header_label = f'Section {section.key}: {section.name}' if section.key != 'ALL' else section.name
        header_marks = str(section.expected_marks) if section.expected_marks > 0 else ''
        rows.append(meta_base + [
            unit_as, er_code, '', section_code, section.name,
            '', '', header_marks, 'Header', header_label, '', str(seq)
        ])
        seq += 1

        # Extract instructions from section start
        instructions = _extract_instructions(text)
        if instructions:
            rows.append(meta_base + [
                unit_as, er_code, '', section_code, section.name,
                '', '', '', 'SubHeader', 'Instructions', '', str(seq)
            ])
            seq += 1
            for instr in instructions:
                rows.append(meta_base + [
                    unit_as, er_code, '', section_code, section.name,
                    '', '', '', 'Instruction', instr, '', str(seq)
                ])
                seq += 1

        # Case study preamble (for BM Section B type)
        if section.section_type == 'case_study':
            preamble = extract_case_study_preamble(text)
            if preamble:
                rows.append(meta_base + [
                    unit_as, er_code, '', section_code, section.name,
                    '', '', '', 'CaseStudy', preamble, '', str(seq)
                ])
                seq += 1

        # Extract questions based on section type
        if section.section_type == 'multiple_choice':
            questions = extract_mc_questions(text)
        else:
            questions = extract_written_questions(text)

        print(f"      Questions found: {len(questions)}")
        sections_questions[section.key] = questions

        # Validate sub-question sequences (detect gaps from OCR extraction issues)
        _validate_subquestions(questions, section.key, year, stream_code)

        # Convert questions to CSV rows
        for q in questions:
            eq_code = f'VCE{subject_code}{stream_prefix}{yy}EQ{section_code}{q.number}'

            if q.content_type == 'MultipleChoice':
                # MC: single row per question
                rows.append(meta_base + [
                    unit_as, er_code, eq_code, section_code, section.name,
                    str(q.number), f'Q{q.number}', str(q.total_marks),
                    'MultipleChoice', q.context, '', str(seq)
                ])
                seq += 1
            elif q.sub_questions:
                # Written question with sub-parts
                # Context/scenario rows — split into typed segments
                if q.context:
                    segments = _split_context_segments(q.context)
                    for seg_i, (seg_type, seg_text) in enumerate(segments):
                        marks_val = str(q.total_marks) if seg_i == 0 else ''
                        rows.append(meta_base + [
                            unit_as, er_code, eq_code, section_code, section.name,
                            str(q.number), f'Q{q.number}', marks_val,
                            seg_type, seg_text, '', str(seq)
                        ])
                        seq += 1

                # Sub-question rows
                for sq in q.sub_questions:
                    q_label = f'Q{q.number}{sq.letter}'
                    rows.append(meta_base + [
                        unit_as, er_code, eq_code, section_code, section.name,
                        f'{q.number}{sq.letter}', q_label, str(sq.marks),
                        'SubQuestion', sq.text, '', str(seq)
                    ])
                    seq += 1

                    # Mid-question context: additional info/stimulus between sub-questions
                    # e.g. "The following information is also provided..." in Accounting
                    if sq.trailing_context:
                        rows.append(meta_base + [
                            unit_as, er_code, eq_code, section_code, section.name,
                            str(q.number), f'Q{q.number}', '',
                            'MidQuestionContext', sq.trailing_context, '', str(seq)
                        ])
                        seq += 1
            else:
                # Written question without sub-parts (essay/single question)
                segments = _split_context_segments(q.context)
                if segments:
                    for seg_i, (seg_type, seg_text) in enumerate(segments):
                        marks_val = str(q.total_marks) if seg_i == 0 else ''
                        rows.append(meta_base + [
                            unit_as, er_code, eq_code, section_code, section.name,
                            str(q.number), f'Q{q.number}', marks_val,
                            seg_type, seg_text, '', str(seq)
                        ])
                        seq += 1
                else:
                    rows.append(meta_base + [
                        unit_as, er_code, eq_code, section_code, section.name,
                        str(q.number), f'Q{q.number}', str(q.total_marks),
                        'Question', q.context, '', str(seq)
                    ])
                    seq += 1

    # Sanity check: marks consistency across 3 levels
    print(f"    --- Marks sanity check ---")
    _sanity_check_marks(sections_questions, structure, year, stream_code)

    return rows


def _year_encode_er_code(base_code: str, subject_code: str, yy: str) -> str:
    """Year-encode an ExamReportCode by inserting yy before 'ER'.

    Examples:
        VCEACER -> VCEAC25ER
        VCEBMERA -> VCEBM25ERA
        VCEECER -> VCEEC25ER
    """
    # Insert year after VCE + subject_code (before the 'ER' section marker)
    # Using subject_code length avoids false 'ER' matches (e.g. VCE + RS = VC[ER]S...)
    insert_pos = 3 + len(subject_code)
    if insert_pos <= len(base_code):
        return base_code[:insert_pos] + yy + base_code[insert_pos:]
    # Fallback: append year
    return base_code + yy


def _extract_instructions(text: str) -> List[str]:
    """Extract instruction lines from section text (bullets after 'Instructions')."""
    instructions = []
    lines = text.split('\n')
    in_instructions = False
    for line in lines:
        stripped = line.strip()
        if re.match(r'^Instructions\s*$', stripped, re.IGNORECASE):
            in_instructions = True
            continue
        if in_instructions:
            # Instruction bullets start with +, -, *, or bullet chars
            if re.match(r'^[+\-*\u2022]\s+', stripped):
                instr = re.sub(r'^[+\-*\u2022]\s+', '', stripped).strip()
                instructions.append(clean_question_text(instr))
            elif stripped.startswith('Answer all') or stripped.startswith('Write your') or \
                 stripped.startswith('Use the') or stripped.startswith('Follow the') or \
                 stripped.startswith('Choose the') or stripped.startswith('A correct') or \
                 stripped.startswith('Marks will') or stripped.startswith('No marks') or \
                 stripped.startswith('At the end'):
                instructions.append(clean_question_text(stripped))
            elif re.match(r'^Question\s+\d+', stripped, re.IGNORECASE):
                break  # Hit first question
            elif stripped and not instructions:
                continue  # Skip non-instruction text before bullets
            elif stripped:
                break  # Non-bullet line after instructions = done
    return instructions


def _validate_subquestions(questions: List[Question], section_key: str, year: str,
                           stream_code: Optional[str] = None):
    """Post-extraction validation: detect gaps in sub-question letter sequences.

    If a question has sub-questions starting at 'b' or 'c' (skipping 'a'), or has
    gaps like a,c (missing b), warn the user — this usually means OCR garbage
    prevented extraction.
    """
    stream_label = f" [{stream_code}]" if stream_code else ""
    for q in questions:
        if not q.sub_questions:
            continue
        letters = [sq.letter for sq in q.sub_questions]
        # Check if starts at 'a'
        if letters and letters[0] != 'a':
            print(f"      WARNING: Q{q.number} ({year}{stream_label} Section {section_key}) "
                  f"sub-questions start at '{letters[0]}' — missing earlier parts: "
                  f"got {letters}")
        # Check for gaps in sequence
        for i in range(1, len(letters)):
            expected = chr(ord(letters[i - 1]) + 1)
            if letters[i] != expected:
                print(f"      WARNING: Q{q.number} ({year}{stream_label} Section {section_key}) "
                      f"sub-question gap: '{letters[i - 1]}' -> '{letters[i]}' "
                      f"(expected '{expected}')")
        # Check for 0-mark sub-questions (usually means marks regex failed due to OCR noise)
        for sq in q.sub_questions:
            if sq.marks == 0:
                print(f"      WARNING: Q{q.number}{sq.letter} ({year}{stream_label} Section {section_key}) "
                      f"has 0 marks — check for OCR artifacts in text: "
                      f"'{sq.text[-60:]}'")
    # Also check total-mark questions without sub-questions that have 0 marks
    for q in questions:
        if not q.sub_questions and q.total_marks == 0:
            print(f"      WARNING: Q{q.number} ({year}{stream_label} Section {section_key}) "
                  f"has 0 total marks — check extraction")


def _sanity_check_marks(sections_questions: Dict[str, List[Question]],
                        structure: ExamStructure,
                        year: str, stream_code: Optional[str] = None):
    """Three-level marks sanity check before CSV output.

    1. Sub-question marks sum == question total_marks
    2. Question marks sum == section expected marks (if known from cover page)
    3. Section marks sum == exam total marks (if known from cover page)

    Reports FAIL for check 1 (data integrity), WARN for checks 2-3 (OCR-derived).
    """
    stream_label = f" [{stream_code}]" if stream_code else ""
    label = f"{year}{stream_label}"
    all_ok = True
    section_computed: Dict[str, int] = {}

    for section in structure.sections:
        questions = sections_questions.get(section.key, [])
        section_total = 0

        for q in questions:
            # Check 1: sub-question marks sum vs question total_marks
            if q.sub_questions:
                sub_sum = sum(sq.marks for sq in q.sub_questions)
                if sub_sum != q.total_marks:
                    sub_detail = '+'.join(f'{sq.letter}({sq.marks})' for sq in q.sub_questions)
                    print(f"      FAIL: Q{q.number} ({label} Section {section.key}) "
                          f"sub-question marks {sub_detail} = {sub_sum}, "
                          f"but question total = {q.total_marks}")
                    all_ok = False
                section_total += q.total_marks
            else:
                section_total += q.total_marks

        section_computed[section.key] = section_total

        # Check 2: section marks sum vs expected (from cover page)
        if section.expected_marks > 0:
            expected = section.expected_marks
            choice = structure.section_choice.get(section.key)
            if choice:
                # "N of M questions" — cover shows student total (N questions),
                # but we extract all M questions. Scale expected up.
                answer_n, of_m = choice
                if answer_n > 0 and of_m > answer_n:
                    # All questions should have equal marks for clean scaling
                    # e.g. 2 of 3 questions, 50 marks → each Q = 25 → all 3 = 75
                    scaled = expected * of_m // answer_n
                    print(f"      Section {section.key} ({label}): "
                          f"{section_total} marks (answer {answer_n} of {of_m} questions, "
                          f"cover says {expected} marks for {answer_n})")
                    # Don't fail on this — just report for manual review
                    continue
            if section_total != expected:
                print(f"      FAIL: Section {section.key} ({label}) "
                      f"question marks sum = {section_total}, "
                      f"but expected = {expected} (from cover page)")
                all_ok = False
            else:
                print(f"      Section {section.key} ({label}): "
                      f"{section_total} marks - PASS")
        else:
            print(f"      Section {section.key} ({label}): "
                  f"{section_total} marks (no expected total to check against)")

    # Check 3: overall exam total
    exam_total = sum(section_computed.values())
    if structure.expected_total_marks > 0:
        has_choice = bool(structure.section_choice)
        if has_choice:
            # Exam has "N of M" sections — cover total is student total, not all questions
            print(f"      Exam total ({label}): {exam_total} marks "
                  f"(cover says {structure.expected_total_marks}, "
                  f"but exam has choice questions)")
        elif exam_total != structure.expected_total_marks:
            print(f"      FAIL: Exam total ({label}) "
                  f"= {exam_total}, but expected = {structure.expected_total_marks} "
                  f"(from cover page)")
            all_ok = False
        else:
            print(f"      Exam total ({label}): "
                  f"{exam_total} marks - PASS")
    else:
        print(f"      Exam total ({label}): {exam_total} marks "
              f"(no expected total to check against)")

    return all_ok


def _resolve_stream_name(config: Dict, stream_code: Optional[str]) -> str:
    """Look up stream display name from config.

    streams config supports two formats:
        "ANC": "Ancient History"              (string = name)
        "ANC": { "name": "Ancient History" }  (dict with name key)
    Returns empty string if stream_code is None.
    """
    if not stream_code:
        return ''
    streams = config.get('streams', {})
    value = streams.get(stream_code, '')
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        return value.get('name', '')
    return ''


# ============================================================================
# LITERATURE EXAM EXTRACTOR (UNCHANGED)
# ============================================================================

def extract_literature_exam(pdf_path: str, year: str, config: Dict,
                            stream_code: Optional[str] = None, stream_name: str = '') -> List[Dict]:
    """Extract exam questions from a Literature exam PDF.

    Literature exams have a standard structure:
    - Section A: Developing interpretations (Q1 = 6 marks, Q2 = 14 marks per text)
    - Section B: Close analysis (1 question = 20 marks per text)
    - 30 texts across 5 categories (Novels, Plays, Short stories, Other literature, Poetry)
    """
    doc = fitz.open(pdf_path)
    print(f"    Pages: {len(doc)}")

    # Find Section A text pages (look for "Text no." pattern)
    # Use dict keyed by text_num to deduplicate TOC vs actual text pages
    texts_dict = {}
    section_b_start = None

    for i in range(len(doc)):
        page = doc[i]
        text = page.get_text()
        if not text.strip():
            continue

        # Detect Section B start (cover page has "Section B - Task Book" in first few lines)
        # Check first 80 chars to handle "Literature\nSection B - Task Book" variant
        # but not the materials list ("Section B - Task Book of 64 pages") on page 0
        first_80 = text.strip()[:80]
        if 'Section B' in first_80 and 'Task Book' in first_80 and 'pages' not in first_80 and section_b_start is None:
            section_b_start = i
            continue

        # Extract text entries from Section A (before Section B)
        if section_b_start is None:
            num_match = re.search(r'Text\s*no\.\s*(\d+)\s+(.+)', text)
            if num_match:
                text_num = int(num_match.group(1))
                title_line = num_match.group(2).strip()

                # Extract concept from Q2 (only present on actual text pages, not TOC)
                concept_match = re.search(
                    r'concept\s+of\s+(.+?)\s+is\s+endorsed', text, re.IGNORECASE)
                concept = concept_match.group(1).strip() if concept_match else ''

                # Determine category
                cat_match = re.search(
                    r'^(Novels|Plays|Short stories|Other literature|Poetry)\s*$',
                    text, re.MULTILINE)
                category = cat_match.group(1) if cat_match else ''

                # Clean up title (remove OCR artifacts)
                title_line = re.sub(r'\s+', ' ', title_line).strip()

                # Later pages (actual text) overwrite earlier TOC entries
                texts_dict[text_num] = {
                    'num': text_num,
                    'title': title_line,
                    'category': category,
                    'concept': concept,
                }

    texts = [texts_dict[k] for k in sorted(texts_dict.keys())]
    print(f"    Texts found: {len(texts)}")
    if not texts:
        return []

    # Build CSV rows
    subject_code = config['subject_code']
    rows = []

    # Get section config
    section_a = config.get('exam_sections', {}).get('A', {})
    section_b = config.get('exam_sections', {}).get('B', {})

    section_a_name = section_a.get('name', 'Developing interpretations')
    section_b_name = section_b.get('name', 'Close analysis')
    section_a_unit_as = section_a.get('unit_as_code', '')
    section_b_unit_as = section_b.get('unit_as_code', '')
    # Year-encoded codes: derive from exam year (e.g. "2025" -> "25")
    yy = year[-2:]
    section_a_er_code = f'VCE{subject_code}{yy}ERA'
    section_b_er_code = f'VCE{subject_code}{yy}ERB'
    eq_code_a1 = f'VCE{subject_code}{yy}EQA1'
    eq_code_a2 = f'VCE{subject_code}{yy}EQA2'
    eq_code_b = f'VCE{subject_code}{yy}EQB'

    meta_base = [
        config['subject_area'],
        config['name'],
        config.get('_csv_stream_code', stream_code or ''),  # SubjectStreamCode (config fallback)
        config.get('_csv_stream_name', stream_name),        # SubjectStreamName (config fallback)
        'Year 12',
        'Final Exam',
        'Exam Questions',
        year,
    ]

    seq = 1

    # --- Section A header ---
    rows.append(meta_base + [
        section_a_unit_as, section_a_er_code, '', 'A', section_a_name,
        '', '', '', 'Header', f'Section A: {section_a_name}', '', str(seq)
    ])
    seq += 1

    # Section A instructions
    instructions_a = [
        ('There are two questions for each text in Section A. You must answer both questions for your chosen text.', 'Bullet'),
        ('One passage has been set for each text.', 'Bullet'),
        ('You must use the set passage as the basis of your responses to both questions.', 'Bullet'),
        ('Your selected text for Section A must be from a different category than your selected text for Section B.', 'Bullet'),
    ]
    rows.append(meta_base + [
        section_a_unit_as, section_a_er_code, '', 'A', section_a_name,
        '', '', '', 'SubHeader', 'Instructions', '', str(seq)
    ])
    seq += 1
    for instr_text, content_type in instructions_a:
        rows.append(meta_base + [
            section_a_unit_as, section_a_er_code, '', 'A', section_a_name,
            '', '', '', content_type, instr_text, '', str(seq)
        ])
        seq += 1

    # Section A questions per text
    for t in texts:
        text_label = f"Text {t['num']}"
        title = t['title']
        category = t['category']

        # Text header
        rows.append(meta_base + [
            section_a_unit_as, section_a_er_code, eq_code_a1, 'A', section_a_name,
            '', text_label, '', 'SubHeader',
            f"{category} - {title}", '', str(seq)
        ])
        seq += 1

        # Q1 (6 marks) - standard question
        rows.append(meta_base + [
            section_a_unit_as, section_a_er_code, eq_code_a1, 'A', section_a_name,
            '1', f'{text_label} Q1', '6', 'Paragraph',
            'Explore the significance of the passage below in the text.',
            '', str(seq)
        ])
        seq += 1

        # Q2 (14 marks) - concept-specific
        q2_text = (
            f"Using the passage as a focus, discuss the ways in which the concept of "
            f"{t['concept']} is endorsed, challenged and/or marginalised by the text."
        )
        rows.append(meta_base + [
            section_a_unit_as, section_a_er_code, eq_code_a2, 'A', section_a_name,
            '2', f'{text_label} Q2', '14', 'Paragraph',
            q2_text, '', str(seq)
        ])
        seq += 1

    # --- Section B header ---
    seq_b = 1
    rows.append(meta_base + [
        section_b_unit_as, section_b_er_code, '', 'B', section_b_name,
        '', '', '', 'Header', f'Section B: {section_b_name}', '', str(seq_b)
    ])
    seq_b += 1

    # Section B instructions
    instructions_b = [
        ('You are required to complete one task based on one text.', 'Bullet'),
        ('Three passages have been set for each text.', 'Bullet'),
        ('You must use two or more of the set passages as the basis for a discussion about the selected text.', 'Bullet'),
        ('Your selected text for Section B must be from a different category than your selected text for Section A.', 'Bullet'),
    ]
    rows.append(meta_base + [
        section_b_unit_as, section_b_er_code, '', 'B', section_b_name,
        '', '', '', 'SubHeader', 'Instructions', '', str(seq_b)
    ])
    seq_b += 1
    for instr_text, content_type in instructions_b:
        rows.append(meta_base + [
            section_b_unit_as, section_b_er_code, '', 'B', section_b_name,
            '', '', '', content_type, instr_text, '', str(seq_b)
        ])
        seq_b += 1

    # Section B questions per text
    for t in texts:
        text_label = f"Text {t['num']}"
        title = t['title']
        category = t['category']

        rows.append(meta_base + [
            section_b_unit_as, section_b_er_code, eq_code_b, 'B', section_b_name,
            '', text_label, '', 'SubHeader',
            f"{category} - {title}", '', str(seq_b)
        ])
        seq_b += 1

        rows.append(meta_base + [
            section_b_unit_as, section_b_er_code, eq_code_b, 'B', section_b_name,
            '1', f'{text_label} Q1', '20', 'Paragraph',
            f"Use two or more of the set passages as the basis for a discussion of {title}.",
            '', str(seq_b)
        ])
        seq_b += 1

    return rows


# ============================================================================
# CSV WRITER
# ============================================================================

FIELDNAMES = [
    'SubjectArea', 'Subject', 'SubjectStreamCode', 'SubjectStreamName', 'Band',
    'AssessmentType', 'AssessmentInformationDetails', 'AssessmentYears', 'UnitASCode',
    'ExamReportCode', 'EQCode', 'SectionCode', 'SectionName', 'QuestionNumber',
    'QuestionLabel', 'Marks', 'ContentType', 'Content', 'StimulusSource', 'Sequence'
]


def write_csv(all_rows: List[List], output_path: str):
    """Write exam questions CSV."""
    sanitize_rows(all_rows)
    lp = _long_path(output_path)
    with open(lp, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(FIELDNAMES)
        writer.writerows(all_rows)
    print(f"  Wrote {output_path} ({len(all_rows)} data rows)")


# ============================================================================
# MAIN
# ============================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python exam-questions-parser-template.py <config.json>")
        sys.exit(1)

    config = load_config(sys.argv[1])

    print("=" * 70)
    print(f"VCE {config['name']} Exam Questions Parser")
    print("=" * 70)

    if not config.get('exam_sections'):
        print("  No exam_sections configured - skipping")
        sys.exit(0)

    exam_schedule = config.get('exam_schedule', [])
    if not exam_schedule:
        print("  No exam PDFs configured (documents.exam_pdfs) - skipping")
        sys.exit(0)

    output_dir = Path(_long_path(config['output_dir']))
    output_dir.mkdir(parents=True, exist_ok=True)
    slug = config.get('subject_slug', 'unknown')

    all_rows = []

    for year, stream_code, pdf_path in exam_schedule:
        if not Path(pdf_path).exists():
            stream_label = f" [{stream_code}]" if stream_code else ""
            print(f"\n  WARNING: {year}{stream_label} exam PDF not found: {pdf_path}")
            continue

        stream_name = _resolve_stream_name(config, stream_code)
        # Config-level fallback for CSV columns: when no Music-style streams,
        # use config subject_stream fields for SubjectStreamCode/SubjectStreamName
        config['_csv_stream_code'] = stream_code or config.get('subject_stream_code', '')
        config['_csv_stream_name'] = stream_name or config.get('subject_stream_name', '')
        stream_label = f" [{stream_code}: {stream_name}]" if stream_code else ""
        print(f"\n  Processing {year}{stream_label} exam: {Path(pdf_path).name}")

        # Dispatch: Literature uses its own extractor, others use generic
        exam_format = config.get('exam_format', '')
        subject_slug = config.get('subject_slug', '')
        if exam_format == 'literature' or subject_slug == 'literature':
            rows = extract_literature_exam(pdf_path, year, config, stream_code, stream_name)
        else:
            structure = detect_exam_structure(pdf_path, config)
            rows = extract_generic_exam(pdf_path, year, config, structure, stream_code, stream_name)

        all_rows.extend(rows)
        print(f"    Total rows for {year}{stream_label}: {len(rows)}")

    if all_rows:
        output_path = str(Path(config['output_dir']) / f'vcaa_vce_sd_{slug}_exam_questions.csv')
        write_csv(all_rows, output_path)
    else:
        print("\n  No exam questions extracted")

    print("\n" + "=" * 70)
    print("COMPLETE")
    print("=" * 70)


if __name__ == '__main__':
    main()
