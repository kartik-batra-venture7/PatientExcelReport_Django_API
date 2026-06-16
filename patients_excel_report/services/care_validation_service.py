"""
Port of CareValidationService.cs

Validates Personal Care codes and Mobile Text Notes per the
ExcelScript matching logic used in the original .NET service.

Returns a 3-tuple: (pc_status, mobile_status, validation_result)
"""

import re
from typing import Optional
from config.settings import _DENIED_CLS_PATTERNS, _DENIED_RESPITE_PATTERNS, _VALID_CODES, _VALID_RESPITE_PC_CODES


# ---------------------------------------------------------------------------
# Denied-word lists shifted to config/settings.py
# ---------------------------------------------------------------------------


class CareValidationService:
    """Validates Personal Care codes and Mobile Text Note content."""
    # ------------------------------------------------------------------
    # Single-value validate  (used by legacy ExcelProcessor / PatientController)
    # ------------------------------------------------------------------
    def validate(
        self,
        pc_value: str,
        mobile_value: str,
        duties_not_performed: bool = False,
        is_personal_care_column_present: bool = True,
        check_respite_for_mobile_text_value: Optional[list] = None,
    ) -> tuple[str, str, str]:
        """
        Validate a single personal care value + mobile note.

        Returns (pc_status, mobile_note_value, validation_result).
        """
        if check_respite_for_mobile_text_value is None:
            check_respite_for_mobile_text_value = []

        personal_care_status = ""
        mobile_note_value = ""
        validation_result = ""

        # ---- Mobile note processing ----
        if not mobile_value or not mobile_value.strip():
            mobile_note_value = "Text Absent"
        else:
            text = mobile_value.strip()
            words = text.split()

            if len(words) < 5:
                mobile_note_value = "Text less than 5 words"
            else:
                check_respite = False
                if check_respite_for_mobile_text_value:
                    first = check_respite_for_mobile_text_value[0]
                    # first is (row_idx, col_idx, description, value)
                    desc = first[2] if len(first) > 2 else ""
                    val = first[3] if len(first) > 3 else ""
                    check_respite = (
                        bool(desc)
                        and "Respite T1005" in desc
                        and bool(val)
                        and val in _VALID_RESPITE_PC_CODES
                    )

                detected_words: set[str] = set()
                if not check_respite:
                    for pattern in _DENIED_CLS_PATTERNS:
                        for match in pattern.finditer(text):
                            if match.group():
                                detected_words.add(match.group())
                    if not detected_words:
                        mobile_note_value = "OK"
                    else:
                        mobile_note_value = "CLS words detected - " + ", ".join(detected_words)
                else:
                    for pattern in _DENIED_RESPITE_PATTERNS:
                        for match in pattern.finditer(text):
                            if match.group():
                                detected_words.add(match.group())
                    if not detected_words:
                        mobile_note_value = "OK"
                    else:
                        mobile_note_value = "Respite words detected - " + ", ".join(detected_words)

        # ---- Personal care processing ----
        has_value = bool(pc_value and pc_value.strip())

        if is_personal_care_column_present:
            has_error = has_value and pc_value not in _VALID_CODES
            has_valid = pc_value in _VALID_CODES if has_value else False

            if has_error:
                personal_care_status = "Wrong values inserted"
            elif has_valid:
                personal_care_status = "OK"
            else:
                personal_care_status = "No value inserted"
        else:
            personal_care_status = "No duties performed"

        # ---- Final validation result ----
        is_mobile_ok = mobile_note_value == "OK"
        is_pc_ok = personal_care_status in ("OK", "")

        if not has_value and not (mobile_value and mobile_value.strip()):
            validation_result = ""
        elif is_mobile_ok and is_pc_ok:
            validation_result = "TRUE"
        else:
            validation_result = "FALSE"

        # Normalise wording (mirrors .NET Replace calls)
        mobile_note_value = mobile_note_value.replace(
            "CLS word detected", "CLS words detected"
        ).replace("Denied", "CLS words detected")

        return personal_care_status, mobile_note_value, validation_result

    # ------------------------------------------------------------------
    # Multi-value validate  (used by ExcelReaderService)
    # ------------------------------------------------------------------
    def validate_multiple(
        self,
        pc_values: list[str],
        mobile_value: str,
        duties_not_performed: bool,
        is_personal_care_column_present: bool,
        check_respite_for_mobile_text_value: Optional[list] = None,
    ) -> tuple[str, str, str]:
        """
        Validate multiple personal care values (from multiple PC rows)
        together with a mobile note.

        Returns (pc_status, mobile_status, validation_result).
        """
        if check_respite_for_mobile_text_value is None:
            check_respite_for_mobile_text_value = []
        if pc_values is None:
            pc_values = []

        has_any_value = any(v and v.strip() for v in pc_values)
        has_invalid   = any(v and v.strip() and v.strip() not in _VALID_CODES for v in pc_values)
        has_missing   = any(not v or not v.strip() for v in pc_values)
        all_valid     = all(v and v.strip() in _VALID_CODES for v in pc_values) if pc_values else False

        if not is_personal_care_column_present:
            pc_status = "No duties performed"
        elif not has_any_value:
            pc_status = "No value inserted"
        elif has_invalid:
            pc_status = "Wrong value inserted"
        elif has_missing:
            pc_status = "Missing value"
        elif all_valid:
            pc_status = "OK"
        else:
            pc_status = "Invalid combination of values"

        # Mobile validation reuses single-value logic
        _, mobile_status, _ = self.validate(
            "",
            mobile_value,
            duties_not_performed,
            is_personal_care_column_present,
            check_respite_for_mobile_text_value,
        )

        # Final validation result
        is_mobile_ok = mobile_status == "OK"
        is_pc_ok = pc_status in ("OK", "")

        if not has_any_value and not (mobile_value and mobile_value.strip()):
            validation_result = ""
        elif is_mobile_ok and is_pc_ok:
            validation_result = "TRUE"
        else:
            validation_result = "FALSE"

        return pc_status, mobile_status, validation_result
