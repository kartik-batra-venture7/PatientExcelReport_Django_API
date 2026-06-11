"""
Port of TextSimilarityService.cs

Computes word-level Jaccard-style similarity between two strings and
returns whether it exceeds a configurable threshold (default 0.75).
Used to detect copy-pasted Mobile Text Notes across days.
"""

import re


class TextSimilarityService:
    """Word-overlap similarity used to flag copy-pasted mobile notes."""

    _SPLIT_PATTERN = re.compile(r"[ ,\.;:\-_/\\()\[\]{}\t\n]+")

    @staticmethod
    def _tokenize(text: str) -> set[str]:
        if not text or not text.strip():
            return set()
        tokens = TextSimilarityService._SPLIT_PATTERN.split(text.lower())
        # ignore single-character tokens (mirrors the C# `w.Length > 1` filter)
        return {t for t in tokens if len(t) > 1}

    @staticmethod
    def calculate_word_similarity(text1: str, text2: str) -> float:
        words1 = TextSimilarityService._tokenize(text1)
        words2 = TextSimilarityService._tokenize(text2)

        if not words1 and not words2:
            return 1.0
        if not words1 or not words2:
            return 0.0

        intersection = words1 & words2
        min_count = min(len(words1), len(words2))
        return len(intersection) / min_count

    def is_above_threshold(self, text1: str, text2: str, threshold: float = 0.75) -> bool:
        return self.calculate_word_similarity(text1, text2) >= threshold
