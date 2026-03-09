# -*- coding: utf-8 -*-
"""
Consistency checks between:
- config/programs.py (PROGRAM_*_PAGES)
- src/columns_config.py (SUBJECT_COLUMNS)

Purpose:
Fail fast before generation if parsing map and diploma layout drift apart.
"""

import re
from typing import Dict, Set

from config.programs import PROGRAM_IT_PAGES, PROGRAM_ACCOUNTING_PAGES
from src.columns_config import SUBJECT_COLUMNS
from core.utils import normalize_key


_IGNORED_EXTRA_KEYWORDS = (
    "орташабаллы",
    "average",
    "среднийбалл",
    "қорытындыаттестациялау",
    "итоговаяаттестация",
)


def _subject_identifier(name: str) -> str:
    """
    Build stable subject identifier.

    Priority:
    1) Module/outcome code (КМ/ПМ/БМ/ОН/РО + number)
    2) Fallback normalized full name
    """
    n = str(name or "").strip()
    if not n:
        return ""

    m = re.match(r"^(БМ|КМ|ПМ|ОН|РО)\s*\.?\s*0*(\d+(?:\.\d+)?)", n, re.IGNORECASE)
    if m:
        prefix = m.group(1).lower()
        num = m.group(2)
        if prefix in ("км", "пм"):
            return f"m{num}"
        if prefix in ("он", "ро"):
            return f"o{num}"
        return f"bm{num}"

    return normalize_key(n)


def _program_pages(program_code: str):
    if program_code == "3F":
        return PROGRAM_IT_PAGES
    if program_code == "3D":
        return PROGRAM_ACCOUNTING_PAGES
    raise ValueError(f"Unsupported program code: {program_code}")


def _is_subject_trackable(hours, credits) -> bool:
    """
    Trackable means subject is expected to have a source column in SUBJECT_COLUMNS.
    Electives/attestation rows often have empty hours/credits and are skipped.
    """
    h = "" if hours is None else str(hours).strip()
    c = "" if credits is None else str(credits).strip()
    return bool(h) or bool(c)


def validate_program_subject_mapping(program_code: str) -> Dict[str, object]:
    """
    Compare subject universe from PROGRAM_PAGES vs SUBJECT_COLUMNS.

    Returns report dict with mismatches and sample names.
    Does not raise.
    """
    pages = _program_pages(program_code)
    col_map = SUBJECT_COLUMNS.get(program_code, {})

    expected_keys: Set[str] = set()
    expected_name_by_key: Dict[str, str] = {}

    for _, subjects in pages.items():
        for subject in subjects:
            if getattr(subject, "is_module_header", False):
                continue

            if not _is_subject_trackable(subject.hours, subject.credits):
                continue

            nkz = _subject_identifier(subject.name_kz)
            nru = _subject_identifier(subject.name_ru)

            if nkz:
                expected_keys.add(nkz)
                expected_name_by_key[nkz] = subject.name_kz
            if nru:
                expected_keys.add(nru)
                expected_name_by_key[nru] = subject.name_ru

    actual_keys: Set[str] = set()
    actual_name_by_key: Dict[str, str] = {}

    for _, info in col_map.items():
        kz_name = str(info.get("kz", "") or "").strip()
        ru_name = str(info.get("ru", "") or "").strip() or kz_name

        nkz = _subject_identifier(kz_name)
        nru = _subject_identifier(ru_name)

        if nkz:
            actual_keys.add(nkz)
            actual_name_by_key[nkz] = kz_name
        if nru:
            actual_keys.add(nru)
            actual_name_by_key[nru] = ru_name

    missing_in_columns = sorted(k for k in (expected_keys - actual_keys) if k)

    extras_raw = expected_keys.intersection(set())  # no-op for readability
    del extras_raw

    extras = sorted(
        key for key in (actual_keys - expected_keys)
        if key
        and len(key) <= 80
        and not any(word in key for word in _IGNORED_EXTRA_KEYWORDS)
    )

    return {
        "program_code": program_code,
        "ok": not missing_in_columns and not extras,
        "missing_count": len(missing_in_columns),
        "extra_count": len(extras),
        "missing_examples": [expected_name_by_key[k] for k in missing_in_columns[:10]],
        "extra_examples": [actual_name_by_key[k] for k in extras[:10]],
    }


def assert_program_subject_mapping(program_code: str) -> None:
    """
    Validate PROGRAM_PAGES vs SUBJECT_COLUMNS.
    Raises only when missing count is critically high (>3).
    Prints warnings for small mismatches (e.g. subject in model without source column).
    """
    report = validate_program_subject_mapping(program_code)
    if report["missing_count"] == 0:
        return

    # Small number of missing → non-blocking warning (e.g. ОН 4.4 not in source)
    if report["missing_count"] <= 3:
        print(
            f"[WARN] {report['missing_count']} subject(s) in model but not in "
            f"SUBJECT_COLUMNS for {program_code} (will use hardcoded hours/credits): "
            f"{report['missing_examples']}"
        )
        return

    msg = (
        f"Subject mapping mismatch for {program_code}: "
        f"missing={report['missing_count']}, extra={report['extra_count']}\n"
        f"Missing examples: {report['missing_examples']}\n"
        f"Extra examples: {report['extra_examples']}"
    )
    raise ValueError(msg)
