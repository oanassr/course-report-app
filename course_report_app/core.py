from __future__ import annotations

import json
import gc
import hashlib
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from collections import defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from statistics import mean
from typing import Any

import pdfplumber
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
JOBS_DIR = DATA_DIR / "jobs"
_PLAN_CACHE: dict[str, tuple[list["PlanCourse"], dict[str, str]]] = {}
_OCR_ENGINE: Any | None = None
_REPORT_CACHE: dict[str, list["ReportOccurrence"]] = {}
ITEMS = [
    "تم تزويدي في بداية دراستي للمقرر ببيانات شاملة عنه: مخرجات التعلم، استراتيجيات التعليم والتعلم، طرق التقييم",
    "يراعي المنهج الدراسي التطورات العلمية والتقنية والمهنية في مجال التخصص",
    "استراتيجيات التعليم والتعلم (المحاضرة، المناقشة، العصف الذهني، التدريب العملي، حل المشكلات) متنوعة ومناسبة لطبيعة ومستوى المقرر",
    "طرق التقييم (الواجبات، التكليفات، الاختبارات) متنوعة ومناسبة لطبيعة ومستوى المقرر",
    "تم إعلان درجات الأنشطة والتكليفات والاختبارات الفصلية، مع تقديم التغذية الراجعة",
    "أستاذ المقرر تدريسه فعال والتزم بمواعيد المحاضرات والإعداد الجيد لها وتفعيل الساعات المكتبية",
    "أستاذ المقرر لديه خبرة ومهارة عالية بمحتوى المقرر",
    "طبق أستاذ المقرر آليات فعالة لمتابعة انتظام حضور الطلاب وتقدمهم الدراسي والمشاركة الفعالة في أنشطة المقرر",
    "التزم أستاذ المقرر بطرق التقييم التي اشتمل عليها توصيف المقرر",
    "مصادر التعلم (كتب، مراجع، قواعد المعلومات، البلاك بورد... إلخ) كافية ومناسبة لمتطلبات المقرر",
    "تتوفر خدمات تهيئة ودعم فني مناسبة ساهمت في استخدامي الفعال لمصادر التعلم (كتب، مراجع، قواعد المعلومات، البلاك بورد... إلخ)",
]

# Some university plan PDFs use a private-use font for the semester heading.
# These signatures are a fallback for that export format; normal Arabic text
# continues through the regular parser below.
PLAN_TERM_GLYPHS = {
    "": "الثاني",
    "أل": "الأول",
    "": "الخامس",
    "": "الثالث",
    "": "السابع",
    "": "الرابع",
    "": "السادس",
    "": "الثامن",
}

# Hijri year-like numbers that appear in dates inside PDF text — these codes
# should not be mistaken for course numbers (e.g., 1446, 1447, 1448, 1449…).
_HIJRI_YEAR_PATTERN = re.compile(r"1[34]\d{2}$")


@dataclass
class PlanCourse:
    term: str
    code: str
    code_key: str
    name: str
    hours: str


@dataclass
class ReportOccurrence:
    page: int
    code: str
    code_key: str
    name: str
    activity: str
    kind: str
    items: dict[str, float]
    percents: dict[str, float]


def ensure_dirs() -> None:
    JOBS_DIR.mkdir(parents=True, exist_ok=True)


def normalize_text(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "").replace("\n", " ").strip()


def fix_pdf_arabic(text: str) -> str:
    """Fix Arabic text that PDF readers extract in visual (RTL display) order.

    Only applies the token-reversal heuristic to lines that appear visually
    reversed — lines where the Arabic words are in the wrong order but each
    word's letters are individually correct.  Plain Arabic text (already in
    logical order) passes through unchanged.
    """
    lines = [unicodedata.normalize("NFKC", line).strip() for line in (text or "").splitlines()]
    fixed_lines: list[str] = []
    for line in lines:
        tokens = [token for token in re.split(r"\s+", line) if token]
        # Heuristic: if the line contains Arabic but looks visually reversed
        # (digits appear before the text that should follow them in RTL),
        # reverse the token order and flip each Arabic token.  Otherwise keep
        # the line as-is so correctly-ordered Arabic is not corrupted.
        has_arabic = any(re.search(r"[؀-ۿ]", t) for t in tokens)
        if has_arabic and _looks_visually_reversed(tokens):
            fixed: list[str] = []
            for token in reversed(tokens):
                if re.search(r"[؀-ۿ]", token):
                    fixed.append(token[::-1])
                else:
                    fixed.append(token)
            fixed_lines.append(" ".join(fixed))
        else:
            fixed_lines.append(line)
    return clean_arabic_text("\n".join(fixed_lines))


def _looks_visually_reversed(tokens: list[str]) -> bool:
    """Return True if the token sequence appears to be in visual RTL order.

    In visual RTL order the last Arabic word of a sentence appears first.
    A reliable signal: digits embedded in course-code–like patterns often
    appear at the START of the token list when the line is reversed.
    """
    if not tokens:
        return False
    first = tokens[0]
    # A line starting with a bare digit followed by a separator strongly
    # suggests visual ordering (e.g. "3131ريتسجام-3" instead of "3-ريتسجام3131").
    if re.match(r"^\d{3,4}[؀-ۿ]", first):
        return True
    # Arabic presentation-form characters (U+FB50–U+FDFF, U+FE70–U+FEFF)
    # are written in visual order in some old PDFs.
    if any("ﭐ" <= ch <= "﷿" or "ﹰ" <= ch <= "﻿" for ch in first):
        return True
    return False


def clean_arabic_text(text: str) -> str:
    replacements = {
        "األول": "الأول",
        "االول": "الأول",
        "ألاول": "الأول",
        "الاول": "الأول",
        "ألول": "الأول",
        "الثانى": "الثاني",
        "الثانى": "الثاني",
        "ألعمال": "الأعمال",
        "الأعمال": "الأعمال",
        "اإلحصائية": "الإحصائية",
        "الإحصائية": "الإحصائية",
        "اإلدارة": "الإدارة",
        "الإدارة": "الإدارة",
        "اإلدارى": "الإداري",
        "اإلدارية": "الإدارية",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def term_key(text: str) -> str:
    text = clean_arabic_text(text)
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    return re.sub(r"\s+", "", text)


def term_sort_key(text: str) -> tuple[int, str]:
    order = {
        "الاول": 1,
        "الثاني": 2,
        "الثالث": 3,
        "الرابع": 4,
        "الخامس": 5,
        "السادس": 6,
        "السابع": 7,
        "الثامن": 8,
    }
    key = term_key(text)
    return (order.get(key, 999), key)


def code_key(raw_code: str) -> str | None:
    text = normalize_text(raw_code)
    match = re.search(r"(\d)\D+(\d{3,4})", text)
    if not match:
        return None
    # A course number can legitimately start with 13xx (for example 1313).
    # Normalize only padding zeros so exports such as 0102 and 102 match.
    return f"{match.group(1)}-{int(match.group(2))}"


def display_code(key: str) -> str:
    credit, number = key.split("-", 1)
    return f"{credit}-نما{number}"


def course_code(raw_code: str, key: str | None = None) -> str:
    """Return the code text from the source instead of inventing a prefix."""
    fixed = fix_pdf_arabic(raw_code)
    found = re.findall(r"\d-[^\s:]+?\d{3,4}", fixed)
    if found:
        return found[-1]
    # The report form puts the number and prefix in visual order, e.g.
    # 3131حسب-3, which represents 3-حسب1313.
    reversed_form = re.search(r"(\d{3,4})([^\s:]+)-(\d)", fixed)
    if reversed_form:
        prefix = reversed_form.group(2)
        return f"{reversed_form.group(3)}-{prefix}{reversed_form.group(1)[::-1]}"
    return display_code(key) if key else normalize_text(raw_code)


def plan_term(raw_term: str) -> str:
    # Check private-use glyph signatures first (special plan font).
    for signature, term in PLAN_TERM_GLYPHS.items():
        if signature in raw_term or signature[::-1] in raw_term:
            return term
    fixed = fix_pdf_arabic(raw_term)
    for signature, term in PLAN_TERM_GLYPHS.items():
        if signature in fixed or signature[::-1] in fixed:
            return term
    # Look for known Arabic term names in the text.
    known = {
        "الأول": "الأول", "الاول": "الأول", "اول": "الأول",
        "الثاني": "الثاني", "الثانى": "الثاني",
        "الثالث": "الثالث",
        "الرابع": "الرابع",
        "الخامس": "الخامس",
        "السادس": "السادس",
        "السابع": "السابع",
        "الثامن": "الثامن",
    }
    text_norm = re.sub(r"\s+", "", fixed).replace("أ", "ا").replace("إ", "ا")
    for key, canonical in known.items():
        key_norm = re.sub(r"\s+", "", key).replace("أ", "ا").replace("إ", "ا")
        if key_norm in text_norm:
            return canonical
    return fixed


def _ocr_pdf_blocks(pdf_path: Path) -> list[tuple[int, int, int, int, int, str]]:
    """Render an image-only PDF and read Arabic text when OCR is installed."""
    try:
        import fitz
        import numpy as np
        from rapidocr import EngineType, LangDet, LangRec, ModelType, OCRVersion, RapidOCR
    except ImportError as exc:
        raise ValueError(
            "الملف صورة ممسوحة ولا يحتوي نصًا. لم يتم تثبيت مكونات OCR اللازمة لقراءة العربية."
        ) from exc

    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        _OCR_ENGINE = RapidOCR(
            params={
                "Det.engine_type": EngineType.ONNXRUNTIME,
                "Det.lang_type": LangDet.CH,
                "Det.model_type": ModelType.MOBILE,
                "Det.ocr_version": OCRVersion.PPOCRV5,
                "Rec.engine_type": EngineType.ONNXRUNTIME,
                "Rec.lang_type": LangRec.ARABIC,
                "Rec.model_type": ModelType.MOBILE,
                "Rec.ocr_version": OCRVersion.PPOCRV5,
            }
        )
    engine = _OCR_ENGINE
    blocks: list[tuple[int, int, int, int, int, str]] = []
    with fitz.open(str(pdf_path)) as document:
        for page_number, page in enumerate(document, start=1):
            # Use scale 1.5 to balance readability vs memory usage.
            pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
                pixmap.height, pixmap.width, pixmap.n
            )
            result = engine(image)
            boxes = getattr(result, "boxes", None)
            texts = getattr(result, "txts", None)
            for box, text in zip(
                boxes if boxes is not None else (),
                texts if texts is not None else (),
            ):
                x0 = int(min(point[0] for point in box))
                y0 = int(min(point[1] for point in box))
                x1 = int(max(point[0] for point in box))
                y1 = int(max(point[1] for point in box))
                if text:
                    blocks.append((page_number, x0, y0, x1, y1, str(text)))
            del result, image, pixmap
            if page_number % 3 == 0:
                gc.collect()
    return blocks


def _ocr_pdf_text(pdf_path: Path) -> str:
    return "\n".join(block[-1] for block in _ocr_pdf_blocks(pdf_path))


def _ocr_plan_courses(plan_pdf: Path) -> list[PlanCourse]:
    blocks = _ocr_pdf_blocks(plan_pdf)
    courses: list[PlanCourse] = []
    by_page: dict[int, list[tuple[int, int, int, int, int, str]]] = defaultdict(list)
    for block in blocks:
        by_page[block[0]].append(block)
    for page_blocks in by_page.values():
        page_width = max((block[3] for block in page_blocks), default=1)
        page_height = max((block[4] for block in page_blocks), default=1)
        for block in page_blocks:
            _, x0, y0, x1, y1, raw_line = block
            line = normalize_text(raw_line).replace("ـ", "")
            match = re.search(r"(?<!\d)(\d)\s*-\s*([؀-ۿA-Za-z]+)\s*(\d{3,4})(?!\d)", line)
            if not match:
                match = re.search(r"(?<!\d)(\d{3,4})\s*([؀-ۿ]+)\s*-\s*(\d)(?!\d)", line)
                if match:
                    credit, prefix, number = match.group(3), match.group(2), match.group(1)
                else:
                    continue
            else:
                credit, prefix, number = match.group(1), match.group(2), match.group(3)
            if _HIJRI_YEAR_PATTERN.match(number):
                continue
            key = f"{credit}-{number}"
            code = f"{credit}-{prefix}{number}"
            center_x = (x0 + x1) / 2
            center_y = (y0 + y1) / 2
            row_band = (
                0 if center_y < page_height * 0.43
                else 1 if center_y < page_height * 0.57
                else 2 if center_y < page_height * 0.72
                else 3
            )
            right_table = center_x > page_width / 2
            all_terms = (
                ("الأول", "الثاني"),
                ("الثالث", "الرابع"),
                ("الخامس", "السادس"),
                ("السابع", "الثامن"),
            )
            term = all_terms[row_band][0 if right_table else 1]
            candidates = []
            for other in page_blocks:
                _, ox0, oy0, ox1, oy1, other_text = other
                other_center = (ox0 + ox1) / 2
                if abs(((oy0 + oy1) / 2) - center_y) > 45 or other is block:
                    continue
                if other_center >= center_x or re.search(r"\d", other_text):
                    continue
                cleaned = normalize_text(other_text).strip(" -")
                if len(cleaned) > 2 and re.search(r"[؀-ۿ]", cleaned):
                    candidates.append(
                        (-len(cleaned), abs(((oy0 + oy1) / 2) - center_y), -other_center, cleaned)
                    )
            name = min(candidates)[3] if candidates else "مقرر مستخرج عبر OCR"
            courses.append(PlanCourse(term=term, code=code, code_key=key, name=name, hours=""))
    unique: dict[tuple[str, str], PlanCourse] = {}
    for course in courses:
        unique.setdefault((term_key(course.term), course.code_key), course)
    return list(unique.values())


def is_course_code(value: str) -> bool:
    return code_key(value) is not None


def has_private_glyphs(value: str) -> bool:
    # Unicode Private Use Areas: Basic (E000–F8FF) and Supplementary (F0000–FFFFF).
    return any(
        "" <= ch <= ""
        or "\U000F0000" <= ch <= "\U000FFFFF"
        for ch in value or ""
    )


def rating(avg: float) -> str:
    pct = avg * 20
    if pct >= 90:
        return "ممتاز"
    if pct >= 80:
        return "جيد جداً"
    if pct >= 70:
        return "جيد"
    if pct >= 60:
        return "مرضي"
    return "بحاجة إلى تحسين"


def extract_plan_courses(plan_pdf: Path) -> tuple[list[PlanCourse], dict[str, str]]:
    cache_key = hashlib.sha256(plan_pdf.read_bytes()).hexdigest()
    cached = _PLAN_CACHE.get(cache_key)
    if cached:
        return deepcopy(cached[0]), dict(cached[1])
    courses: list[PlanCourse] = []
    meta: dict[str, str] = {"program": "", "college": "", "department": "", "plan": ""}
    with pdfplumber.open(str(plan_pdf)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        fixed_text = normalize_text(text)
        # Try multiple label variants for program name.
        for pattern in [
            r"(?:التخصص|البرنامج|اسم\s+البرنامج|اسم\s+التخصص)\s*[:]\s*([^\n\r]{3,60}?)(?:\s+\d{3,}|\s+الإصدار|$)",
        ]:
            program_match = re.search(pattern, fixed_text)
            if program_match:
                name = program_match.group(1).strip(" :-")
                if name and not name.isdigit():
                    meta["program"] = name
                    break
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table or len(table) < 2:
                    continue
                # The term heading is typically in the first cell of the first row.
                first_cell = (table[0][0] or "") if table[0] else ""
                term = plan_term(first_cell)
                if not term or "المجموع" in term or len(term) > 30:
                    continue
                # Skip header row (index 1) — data starts at index 2.
                data_start = 2 if len(table) >= 3 else 1
                for row in table[data_start:]:
                    cells = [cell or "" for cell in row]
                    code_cell = next((cell for cell in reversed(cells) if is_course_code(cell)), "")
                    key = code_key(code_cell)
                    if not key:
                        continue
                    code_index = cells.index(code_cell)
                    name_cell = cells[code_index - 1] if code_index > 0 else ""
                    hours_cell = cells[code_index - 2] if code_index > 1 else ""
                    courses.append(
                        PlanCourse(
                            term=term,
                            code=course_code(code_cell, key),
                            code_key=key,
                            name=normalize_text(name_cell),
                            hours=normalize_text(hours_cell),
                        )
                    )
    # Private-use glyphs can hide the display text while the embedded PDF
    # text still exposes reliable course keys. Do not OCR the whole text plan
    # in that case; report names/codes fill confirmed matches later.
    needs_ocr = not courses
    ocr_courses: list[PlanCourse] = []
    if needs_ocr:
        try:
            ocr_courses = _ocr_plan_courses(plan_pdf)
        except Exception as exc:
            if not courses:
                raise ValueError(f"تعذر تشغيل OCR العربي لقراءة الخطة المصورة: {exc}") from exc
        if ocr_courses:
            meta["ocr_used"] = "true"
        if not courses:
            courses = ocr_courses
    ocr_by_key = {course.code_key: course for course in ocr_courses}
    for course in courses:
        replacement = ocr_by_key.get(course.code_key)
        if replacement:
            course.code = replacement.code
            course.name = replacement.name
        else:
            if has_private_glyphs(course.code):
                course.code = f"رمز غير مكتمل ({course.code_key})"
            if has_private_glyphs(course.name):
                course.name = "مقرر غير مقروء (يحتاج تأكيد)"
    if not courses:
        raise ValueError(
            "تعذر استخراج مقررات الخطة حتى بعد تشغيل OCR العربي. "
            "تأكد من وضوح الصفحات وجودة المسح."
        )
    _PLAN_CACHE[cache_key] = (deepcopy(courses), dict(meta))
    if len(_PLAN_CACHE) > 4:
        _PLAN_CACHE.pop(next(iter(_PLAN_CACHE)))
    return courses, meta


def _extract_page_code(text: str) -> str | None:
    # First pass: look for a code on the same line as "رمز المقرر".
    for line in text.splitlines():
        line_norm = fix_pdf_arabic(line)
        if "رمز" in line_norm and "مقرر" in line_norm:
            key = code_key(line)
            if not key:
                key = code_key(line_norm)
            if key:
                return key
        # Presentation form of "رمز المقرر" found in some PDFs.
        if "ﺭﺮــﻘـﻤﻟﺍ ﺰـﻣﺭ" in line:
            key = code_key(line)
            if key:
                return key
    # Second pass: pick the first course-code–like pattern in the page text,
    # skipping Hijri year numbers.
    fixed_text = fix_pdf_arabic(text)
    matches = re.findall(r"\d-[^\s:]{1,12}\d{3,4}", fixed_text)
    for match in matches:
        key = code_key(match)  # code_key already filters Hijri years
        if key:
            return key
    return None


def _fix_legacy_pdf_text(text: str) -> str:
    """Normalize report metadata lines exported in visual Arabic order."""
    fixed_lines: list[str] = []
    for raw_line in (text or "").splitlines():
        if any("\ufb50" <= char <= "\ufdff" or "\ufe70" <= char <= "\ufeff" for char in raw_line):
            normalized = unicodedata.normalize("NFKC", raw_line)
            tokens = [token for token in re.split(r"\s+", normalized) if token]
            reversed_tokens = [
                token[::-1] if re.search(r"[؀-ۿ]", token) else token
                for token in reversed(tokens)
            ]
            fixed_lines.append(" ".join(reversed_tokens))
        else:
            fixed_lines.append(raw_line)
    return "\n".join(fixed_lines)


def _extract_page_name(text: str) -> str:
    normalized = _fix_legacy_pdf_text(text).replace("ـ", "")
    # Pattern 1: "اسم المقرر : <name>"
    match = re.search(
        r"اسم\s+المقرر\s*[:]\s*([^\n\r]{3,80}?)(?:\s+(?:طبيعة|رمز|نوع|عدد)|\s*$)",
        normalized,
    )
    if match:
        return match.group(1).strip(" :-")
    # Pattern 2: multiline — name may be on the next line after the label
    for line in normalized.splitlines():
        if "اسم" in line and "المقرر" in line:
            parts = re.split(r"[:]", line, maxsplit=1)
            if len(parts) == 2:
                name = parts[1].strip(" :-")
                if len(name) > 2:
                    return name
    return ""


def _extract_page_display_code(text: str, key: str) -> str:
    normalized = _fix_legacy_pdf_text(text).replace("ـ", "")
    for line in normalized.splitlines():
        if "رمز" in line and "مقرر" in line:
            value = line.split("رمز المقرر", 1)[-1] if "رمز المقرر" in line else line
            value = value.split("رقم المقرر", 1)[0]
            reversed_form = re.search(r"(\d{3,4})([^\s:]+)-(\d)", value)
            if reversed_form:
                return f"{reversed_form.group(3)}-{reversed_form.group(2)}{reversed_form.group(1)[::-1]}"
            direct = re.search(r"\d-[^\s:]{1,12}\d{3,4}", value)
            if direct:
                return direct.group(0)
            return display_code(key)
    return display_code(key)


def _extract_activity(text: str) -> str:
    """Detect whether the activity type is practical (عملي) or theoretical (نظري).

    Checks both Arabic presentation forms (present in some legacy PDFs) and
    standard Unicode Arabic characters (after NFKC normalization).
    """
    normalized = normalize_text(text)
    # Standard Arabic Unicode check (most PDFs after normalization).
    if "عملي" in normalized or "عملى" in normalized:
        return "عملي"
    if "نظري" in normalized or "نظرى" in normalized:
        return "نظري"
    # Arabic Presentation Form A fallback (legacy/scanned PDFs).
    if "ﻲﻠﻤﻋ" in text or "ﻋﻤﻠﻲ" in text:
        return "عملي"
    if "ﻱﺮﻈﻧ" in text or "ﻧﻈﺮﻱ" in text:
        return "نظري"
    return ""


def _extract_item_rows(page: Any) -> tuple[dict[str, float], dict[str, float]]:
    """Extract per-item averages and percentages from the measurement table.

    The standard report table has:
      column 0  → percentage (النسبة المئوية)
      column 1  → average   (المتوسط الحسابي)
      columns 2–N → intermediate data
      last column → item number (1–11)

    This function is tolerant of tables with fewer columns than expected and
    tries alternative column positions when the standard ones fail.
    """
    items: dict[str, float] = {}
    percents: dict[str, float] = {}
    for table in page.extract_tables():
        for row in table:
            if not row or len(row) < 3:
                continue
            # Search for item number (1–11) in the last few cells.
            item_num: int | None = None
            for cell in reversed(row[-3:]):
                val = normalize_text(cell or "")
                if val.isdigit() and 1 <= int(val) <= 11:
                    item_num = int(val)
                    break
            if item_num is None:
                continue
            key = str(item_num)
            # Standard positions: col 0 = percent, col 1 = average.
            try:
                avg = float(normalize_text(row[1] or ""))
                pct = float(normalize_text(row[0] or ""))
                items[key] = avg
                percents[key] = pct
                continue
            except (ValueError, IndexError):
                pass
            # Fallback: scan for a float in [1.0, 5.0] which is a Likert average.
            for cell in row[1: len(row) - 1]:
                try:
                    val = float(normalize_text(cell or ""))
                    if 1.0 <= val <= 5.0:
                        items[key] = val
                        break
                except ValueError:
                    pass
    return items, percents


def extract_report_occurrences(report_pdf: Path) -> list[ReportOccurrence]:
    occurrences: list[ReportOccurrence] = []
    with pdfplumber.open(str(report_pdf)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            key = _extract_page_code(text)
            if not key:
                continue
            items, percents = _extract_item_rows(page)
            occurrences.append(
                ReportOccurrence(
                    page=page_number,
                    code=_extract_page_display_code(text, key),
                    code_key=key,
                    name=_extract_page_name(text),
                    activity=_extract_activity(text),
                    kind="تفصيلي" if items else "ملخص",
                    items=items,
                    percents=percents,
                )
            )
    return occurrences


def create_analysis(
    job_dir: Path,
    selected_terms: list[str],
    semester_label: str = "",
) -> dict[str, Any]:
    plan_courses, meta = extract_plan_courses(job_dir / "plan.pdf")
    if semester_label:
        meta["semester_label"] = semester_label.strip()
    report_occurrences = extract_report_occurrences(job_dir / "report.pdf")
    if not any(item.kind == "تفصيلي" for item in report_occurrences):
        raise ValueError(
            "تعذر قراءة التقرير التفصيلي: لم يتم العثور على صفحات قياس تحتوي على بنود التقييم. "
            "تأكد من رفع التقرير التفصيلي الأصلي وليس صورة ممسوحة أو صفحة ملخص فقط."
        )
    by_key: dict[str, list[ReportOccurrence]] = defaultdict(list)
    for occurrence in report_occurrences:
        by_key[occurrence.code_key].append(occurrence)

    selected_term_keys = {term_key(term) for term in selected_terms}
    selected = [
        course for course in plan_courses
        if not selected_term_keys or term_key(course.term) in selected_term_keys
    ]
    matches = []
    for course in selected:
        occurrences = by_key.get(course.code_key, [])
        detail = [item for item in occurrences if item.kind == "تفصيلي"]
        summaries = [item for item in occurrences if item.kind == "ملخص"]
        source_name = course.name
        source_code = course.code
        report_source = next((item for item in detail if item.name), None)
        if report_source:
            if not source_name or has_private_glyphs(source_name) or source_name.startswith("مقرر غير مقروء"):
                source_name = report_source.name
            if not source_code or has_private_glyphs(source_code) or source_code.startswith("رمز غير مكتمل"):
                source_code = report_source.code
        status = "مطابق" if detail else "غير موجود"
        if len(detail) > 1:
            status = "مطابق مع تكرار"
        matches.append(
            {
                "selected": bool(detail),
                "status": status,
                "term": course.term,
                "plan_code": source_code,
                "code_key": course.code_key,
                "name": source_name,
                "hours": course.hours,
                "detail_pages": [item.page for item in detail],
                "summary_pages": [item.page for item in summaries],
                "activities": sorted({item.activity for item in detail if item.activity}),
                "source_keys": [course.code_key],
            }
        )
    all_terms = {course.term for course in plan_courses}
    if meta.get("ocr_used") == "true":
        all_terms.update(
            {"الأول", "الثاني", "الثالث", "الرابع", "الخامس", "السادس", "السابع", "الثامن"}
        )
    analysis = {
        "meta": meta,
        "selected_terms": selected_terms,
        "matches": matches,
        "all_terms": sorted(all_terms, key=term_sort_key),
        "report_occurrences": [asdict(item) for item in report_occurrences],
    }
    (job_dir / "analysis.json").write_text(
        json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return analysis


def load_analysis(job_dir: Path) -> dict[str, Any]:
    return json.loads((job_dir / "analysis.json").read_text(encoding="utf-8"))


def compute_course_stats(analysis: dict[str, Any], match: dict[str, Any]) -> dict[str, Any]:
    source_keys = set(match.get("source_keys") or [match["code_key"]])
    occurrences = [
        ReportOccurrence(**item)
        for item in analysis["report_occurrences"]
        if item["code_key"] in source_keys and item["kind"] == "تفصيلي"
    ]
    item_values: dict[str, float | None] = {}
    raw_item_values: list[float] = []
    for item_number in range(1, 12):
        values = [
            occ.items[str(item_number)]
            for occ in occurrences
            if str(item_number) in occ.items
        ]
        raw_value = mean(values) if values else None
        if raw_value is not None:
            raw_item_values.append(raw_value)
        item_values[str(item_number)] = round(raw_value, 2) if raw_value is not None else None
    raw_avg = mean(raw_item_values) if raw_item_values else None
    avg = round(raw_avg, 2) if raw_avg is not None else None
    return {
        "item_values": item_values,
        "average": avg,
        "percent": round(raw_avg * 20, 2) if raw_avg is not None else None,
        "rating": rating(avg) if avg is not None else "غير محسوب",
        "detail_pages": sorted({occ.page for occ in occurrences}),
    }


def set_rtl(paragraph: Any, align: Any | None = WD_ALIGN_PARAGRAPH.RIGHT) -> None:
    if align is not None:
        paragraph.alignment = align
    ppr = paragraph._p.get_or_add_pPr()
    bidi = ppr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        ppr.append(bidi)
    bidi.set(qn("w:val"), "1")


def set_run_rtl(run: Any) -> None:
    rpr = run._element.get_or_add_rPr()
    rtl = rpr.find(qn("w:rtl"))
    if rtl is None:
        rtl = OxmlElement("w:rtl")
        rpr.append(rtl)
    rtl.set(qn("w:val"), "1")


def set_table_rtl(table: Any) -> None:
    tbl_pr = table._tbl.tblPr
    bidi = tbl_pr.find(qn("w:bidiVisual"))
    if bidi is None:
        bidi = OxmlElement("w:bidiVisual")
        tbl_pr.append(bidi)
    bidi.set(qn("w:val"), "1")


def set_cell_text(
    cell: Any,
    text: str,
    bold: bool = False,
    size: int = 8,
    shade: str | None = None,
    align: Any = WD_ALIGN_PARAGRAPH.CENTER,
) -> None:
    cell.text = ""
    if shade:
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = tc_pr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tc_pr.append(shd)
        shd.set(qn("w:fill"), shade)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    set_rtl(p, align)
    run = p.add_run(text)
    set_run_rtl(run)
    run.bold = bold
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:cs"), "Arial")
    run.font.size = Pt(size)
    if shade == "1F4E79":
        run.font.color.rgb = RGBColor(255, 255, 255)


def set_table_borders(table: Any) -> None:
    set_table_rtl(table)
    borders = table._tbl.tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table._tbl.tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "D9D9D9")


def clear_borders(table: Any) -> None:
    borders = table._tbl.tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table._tbl.tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "nil")


def repeat_header_row(row: Any) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = tr_pr.find(qn("w:tblHeader"))
    if header is None:
        header = OxmlElement("w:tblHeader")
        tr_pr.append(header)
    header.set(qn("w:val"), "true")


def add_header(doc: Document, college_ar: str = "كلية الأعمال", college_en: str = "College of Business") -> None:
    section = doc.sections[0]
    header = section.header
    for child in list(header._element):
        header._element.remove(child)
    table = header.add_table(rows=1, cols=3, width=Cm(25.5))
    clear_borders(table)
    left = table.rows[0].cells[0].paragraphs[0]
    left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for line in ["Kingdom of Saudi Arabia", "Ministry of Education", "King Khalid University", college_en]:
        run = left.add_run(line)
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 80, 30)
        left.add_run("\n")
    center = table.rows[0].cells[1].paragraphs[0]
    center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    center.add_run("King Khalid University")
    right = table.rows[0].cells[2].paragraphs[0]
    set_rtl(right)
    for line in ["المملكة العربية السعودية", "وزارة التعليم", "جامعة الملك خالد", college_ar]:
        run = right.add_run(line)
        run.bold = True
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:cs"), "Arial")
        set_run_rtl(run)
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 80, 30)
        right.add_run("\n")
    for child in list(section.footer._element):
        section.footer._element.remove(child)


def add_heading(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    set_rtl(p)
    run = p.add_run(text)
    set_run_rtl(run)
    run.bold = True
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:cs"), "Arial")
    run.font.size = Pt(13)


def add_paragraph(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    set_rtl(p)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_run_rtl(run)
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:cs"), "Arial")
    run.font.size = Pt(11)


def build_docx(job_dir: Path, confirmed_matches: list[dict[str, Any]]) -> Path:
    analysis = load_analysis(job_dir)
    output_path = job_dir / "تقرير_تحليل_استبانة_المقررات.docx"
    selected = [match for match in confirmed_matches if match.get("selected")]
    if not selected:
        raise ValueError("لا توجد مقررات محددة لتوليد التقرير.")
    for match in selected:
        stats = compute_course_stats(analysis, match)
        match.update(stats)

    meta = analysis.get("meta", {})
    program = meta.get("program") or "تحليل بيانات الأعمال"
    semester_label = meta.get("semester_label", "").strip()
    terms_text = " و".join(analysis.get("selected_terms") or [])
    # Use the explicit semester label if provided; otherwise fall back to
    # the Arabic term names (e.g. "الأول و الثاني").
    semester_display = semester_label if semester_label else terms_text

    doc = Document()
    body = doc._body._element
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Cm(1.3)
    section.bottom_margin = Cm(1.3)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.header_distance = Cm(0.4)
    section.footer_distance = Cm(0.4)

    college_ar = meta.get("college_ar", "كلية الأعمال")
    college_en = meta.get("college_en", "College of Business")
    add_header(doc, college_ar=college_ar, college_en=college_en)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(title)
    title_text = (
        f"تقرير تحليل استبانة تقييم جودة المقررات الدراسية"
        f" لبرنامج {program}"
        + (f" للفصل {semester_display}" if semester_display else "")
    )
    title_run = title.add_run(title_text)
    set_run_rtl(title_run)
    title_run.bold = True
    title_run.font.name = "Arial"
    title_run._element.rPr.rFonts.set(qn("w:cs"), "Arial")
    title_run.font.size = Pt(16)

    add_heading(doc, "تحليل وتقييم جودة المقررات الدراسية")
    add_paragraph(
        doc,
        f"يحلل هذا التقرير نتائج استبانة تقييم المقررات المطروحة وفق الخطة الدراسية لبرنامج {program}. "
        f"تم اعتماد {len(selected)} مقررات"
        + (f" من {terms_text}" if terms_text else "")
        + "، مع تثبيت رموز وأسماء المقررات كما وردت في الخطة.",
    )
    add_paragraph(
        doc,
        "احتسب المتوسط الحسابي لكل بند من صفحات القياس التفصيلية فقط. وعند تكرار المقرر في أكثر من صفحة، "
        "تم أخذ متوسط كل بند عبر جميع صفحات القياس الخاصة بالمقرر، واستبعاد صفحات الملخص من الحساب حتى لا تتكرر القيمة.",
    )

    add_heading(doc, "نتائج التقييم لكل مقرر")
    table = doc.add_table(rows=1 + len(ITEMS), cols=1 + len(selected))
    table.autofit = True
    set_table_borders(table)
    for col, header in enumerate(
        ["البند"] + [f"{m['name']}\n{m['plan_code']}" for m in selected]
    ):
        set_cell_text(table.rows[0].cells[col], header, bold=True, size=7, shade="1F4E79")
    repeat_header_row(table.rows[0])
    for row_idx, item_text in enumerate(ITEMS, start=1):
        set_cell_text(
            table.rows[row_idx].cells[0],
            item_text,
            size=7,
            shade="EAF2F8",
            align=WD_ALIGN_PARAGRAPH.RIGHT,
        )
        for col_idx, match in enumerate(selected, start=1):
            value = match["item_values"].get(str(row_idx))
            set_cell_text(
                table.rows[row_idx].cells[col_idx],
                "-" if value is None else f"{value:.2f}",
                size=8,
            )

    add_heading(doc, "المتوسطات الحسابية لتقييم المقررات")
    summary = doc.add_table(rows=1 + len(selected), cols=6)
    set_table_borders(summary)
    for col, header in enumerate(
        ["المقرر", "الفصل في الخطة", "المتوسط الحسابي", "النسبة المئوية", "التقدير", "صفحات القياس"]
    ):
        set_cell_text(summary.rows[0].cells[col], header, bold=True, size=8, shade="1F4E79")
    repeat_header_row(summary.rows[0])
    for row_idx, match in enumerate(selected, start=1):
        avg_str = f"{match['average']:.2f}" if match.get("average") is not None else "-"
        pct_str = f"{match['percent']:.2f}" if match.get("percent") is not None else "-"
        values = [
            f"{match['name']}\n{match['plan_code']}",
            match["term"],
            avg_str,
            pct_str,
            match.get("rating", "-"),
            ", ".join(map(str, match.get("detail_pages", []))),
        ]
        for col, value in enumerate(values):
            set_cell_text(
                summary.rows[row_idx].cells[col],
                value,
                size=8,
                align=WD_ALIGN_PARAGRAPH.RIGHT if col == 0 else WD_ALIGN_PARAGRAPH.CENTER,
            )

    # Sort only courses that have a computed average.
    scored = [(m["average"], m) for m in selected if m.get("average") is not None]
    averages = sorted(scored, key=lambda item: item[0])

    add_heading(doc, "تحليل النتائج")
    add_heading(doc, "نقاط القوة")
    top_n = averages[-3:] if len(averages) >= 3 else averages
    top_text = "، ".join(
        f"{m['name']} ({m['plan_code']}) بمتوسط {avg:.2f}"
        for avg, m in reversed(top_n)
    ) if top_n else "لا توجد بيانات كافية"
    add_paragraph(
        doc,
        f"أعلى المقررات في النتائج هي: {top_text}. "
        "وتعكس هذه النتائج مستوى مرتفعاً في وضوح معلومات المقرر، وتنوع استراتيجيات التعليم والتقييم، وفاعلية التدريس.",
    )

    add_heading(doc, "مجالات التحسين")
    bottom_n = averages[:3] if len(averages) >= 3 else averages
    bottom_text = "، ".join(
        f"{m['name']} ({m['plan_code']}) بمتوسط {avg:.2f}"
        for avg, m in bottom_n
    ) if bottom_n else "لا توجد بيانات كافية"
    add_paragraph(
        doc,
        f"تتركز فرص التحسين في المقررات الأقل متوسطاً: {bottom_text}. "
        "ويوصى بتحليل البنود التفصيلية لهذه المقررات، خاصة البنود المرتبطة بتنوع استراتيجيات التعليم والتقييم، والتغذية الراجعة، ومصادر التعلم.",
    )

    add_heading(doc, "مصادر القياس ومعالجة التكرار")
    sources = doc.add_table(rows=1 + len(selected), cols=4)
    set_table_borders(sources)
    for col, header in enumerate(
        ["رمز الخطة", "اسم المقرر في الخطة", "صفحات القياس المستخدمة", "ملاحظة المطابقة"]
    ):
        set_cell_text(sources.rows[0].cells[col], header, bold=True, size=8, shade="1F4E79")
    repeat_header_row(sources.rows[0])
    for row_idx, match in enumerate(selected, start=1):
        pages = match.get("detail_pages", [])
        note = (
            "ظهر المقرر في أكثر من موضع، واحتسب المتوسط من جميع صفحات القياس التفصيلية دون صفحات الملخص."
            if len(pages) > 1
            else "احتسب المتوسط من جميع صفحات القياس التفصيلية دون صفحات الملخص."
        )
        values = [
            match["plan_code"],
            match["name"],
            ", ".join(map(str, pages)),
            note,
        ]
        for col, value in enumerate(values):
            set_cell_text(sources.rows[row_idx].cells[col], value, size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)

    add_heading(doc, "خطة التحسين")
    improvement = doc.add_table(rows=5, cols=4)
    set_table_borders(improvement)
    for col, header in enumerate(["مجال التحسين", "الإجراء المقترح", "المسؤول", "الإطار الزمني"]):
        set_cell_text(improvement.rows[0].cells[col], header, bold=True, size=8, shade="1F4E79")
    repeat_header_row(improvement.rows[0])
    improvement_rows = [
        [
            "تحسين المقررات الأقل متوسطاً",
            f"مراجعة البنود الأقل في كل مقرر، وتطوير أنشطة تطبيقية وحالات عملية مرتبطة ببرنامج {program}.",
            "منسق البرنامج وأعضاء هيئة التدريس",
            "الفصل التالي",
        ],
        [
            "توحيد جودة التغذية الراجعة",
            "تحديد آلية واضحة لإعلان درجات الأنشطة والاختبارات وتقديم تغذية راجعة منتظمة للطلاب.",
            "أعضاء هيئة التدريس",
            "طوال الفصل",
        ],
        [
            "تعزيز مصادر التعلم",
            "تحديث مصادر التعلم الرقمية، وتفعيل استخدام البلاك بورد وقواعد المعلومات لدعم تعلم الطلاب.",
            "منسق المقرر والدعم الفني",
            "قبل بداية الفصل",
        ],
        [
            "المتابعة الدورية",
            "إعادة قياس أثر التحسين في نهاية الفصل ومقارنة المتوسطات بنتائج هذا التقرير.",
            "لجنة الجودة ومنسق البرنامج",
            "نهاية الفصل",
        ],
    ]
    for row_idx, row in enumerate(improvement_rows, start=1):
        for col, value in enumerate(row):
            set_cell_text(improvement.rows[row_idx].cells[col], value, size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)

    doc.save(str(output_path))
    return output_path
