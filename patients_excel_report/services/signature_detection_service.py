"""
Port of SignatureDetectionService.cs

Analyses an image (bytes) to determine whether a signature is present.
Any pixel that is non-transparent AND non-white counts as ink.
Images larger than 300×300 are downscaled before analysis.
"""

import io
from PIL import Image


MAX_ANALYSIS_DIMENSION = 300


class SignatureDetectionService:
    """Detects whether an embedded Excel image contains a signature."""

    def detect(self, img_bytes: bytes) -> str:
        """
        Quick detection (used by the legacy PatientController).
        Samples every N-th pixel; returns 'signed' if more than 5
        non-white opaque pixels are found.
        """
        try:
            image = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
            width, height = image.size
            step = max(1, width // 50)
            non_white = 0

            pixels = image.load()
            for y in range(0, height, step):
                for x in range(0, width, step):
                    r, g, b, a = pixels[x, y]
                    if a > 200 and (r + g + b) < 740:
                        non_white += 1
                        if non_white > 5:
                            return "signed"
            return "unsigned"
        except Exception:
            return "unsigned"

    def analyze_image_non_empty_pixel_percent(self, image_bytes: bytes) -> str:
        """
        Full pixel analysis (used by ExcelReaderService).
        Downscales to 300×300 max, then counts non-white non-transparent
        pixels. Any non-empty pixel → 'signed'.
        """
        try:
            image = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
            w, h = image.size

            if w > MAX_ANALYSIS_DIMENSION or h > MAX_ANALYSIS_DIMENSION:
                scale = min(MAX_ANALYSIS_DIMENSION / w, MAX_ANALYSIS_DIMENSION / h)
                new_w = max(1, round(w * scale))
                new_h = max(1, round(h * scale))
                image = image.resize((new_w, new_h), Image.LANCZOS)
                w, h = image.size

            total = w * h
            if total == 0:
                return "unsigned"

            pixels = image.load()
            non_empty = 0

            for y in range(h):
                for x in range(w):
                    r, g, b, a = pixels[x, y]
                    alpha = a / 255.0

                    # nearly transparent → skip
                    if alpha < 0.05:
                        continue

                    # luminance (ITU-R BT.709)
                    lum = 0.2126 * (r / 255.0) + 0.7152 * (g / 255.0) + 0.0722 * (b / 255.0)

                    # pure white + fully opaque → background, skip
                    if lum > 0.98 and alpha > 0.99:
                        continue

                    non_empty += 1

            return "signed" if non_empty > 0 else "unsigned"
        except Exception:
            return "unsigned"
