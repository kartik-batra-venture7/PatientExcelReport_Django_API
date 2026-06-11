"""
Port of ImageExtractor.cs

Extracts images anchored to a specific worksheet from an .xlsx file,
returning a dict mapping cell address (e.g. "F25") → image bytes.

.xlsx files are standard ZIP archives.  The drawing XML at
xl/drawings/drawingN.xml describes TwoCellAnchor elements whose
<xdr:from> child gives the (0-based) column and row of the top-left
corner of the image.  The actual image bytes live in
xl/drawings/media/ (referenced via relationship IDs).
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import Optional


# XML namespaces used in drawing XML
_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Relationship type for images
_IMAGE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


def _column_letter_from_index(zero_based_index: int) -> str:
    """Convert a 0-based column index to an Excel column letter (A, B, … AA, …)."""
    index = zero_based_index + 1
    col = ""
    while index > 0:
        rem = (index - 1) % 26
        col = chr(ord("A") + rem) + col
        index = (index - 1) // 26
    return col


def _get_sheet_drawing_path(
    zf: zipfile.ZipFile, sheet_name: str
) -> Optional[tuple[str, str]]:
    """
    Return (drawing_xml_path, drawing_rels_path) for the given sheet name,
    or None if the sheet has no drawing part.
    """
    # 1. Find the sheet's rId from workbook.xml
    try:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    except Exception:
        return None

    wb_ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    sheet_rid: Optional[str] = None
    for sheet_el in wb_xml.findall(".//ns:sheet", wb_ns):
        if sheet_el.get("name", "").lower() == sheet_name.lower():
            sheet_rid = sheet_el.get(f"{{{r_ns}}}id")
            break

    if not sheet_rid:
        return None

    # 2. Resolve the sheet path via workbook relationships
    try:
        wb_rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except Exception:
        return None

    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    sheet_path: Optional[str] = None
    for rel in wb_rels_xml.findall(f"{{{rels_ns}}}Relationship"):
        if rel.get("Id") == sheet_rid:
            target = rel.get("Target", "")
            sheet_path = "xl/" + target if not target.startswith("xl/") else target
            break

    if not sheet_path:
        return None

    # 3. Find the sheet's _rels file and look for a drawing relationship
    sheet_dir = sheet_path.rsplit("/", 1)[0]
    sheet_file = sheet_path.rsplit("/", 1)[1]
    sheet_rels_path = f"{sheet_dir}/_rels/{sheet_file}.rels"

    try:
        sheet_rels_xml = ET.fromstring(zf.read(sheet_rels_path))
    except Exception:
        return None

    drawing_rid: Optional[str] = None
    drawing_target: Optional[str] = None
    for rel in sheet_rels_xml.findall(f"{{{rels_ns}}}Relationship"):
        if "drawing" in rel.get("Type", "").lower():
            drawing_rid = rel.get("Id")
            drawing_target = rel.get("Target", "")
            break

    if not drawing_target:
        return None

    # drawing_target is relative to the sheet's directory
    if drawing_target.startswith("../"):
        drawing_path = "xl/" + drawing_target[3:]
    elif drawing_target.startswith("/"):
        drawing_path = drawing_target.lstrip("/")
    else:
        drawing_path = f"{sheet_dir}/{drawing_target}"

    drawing_dir = drawing_path.rsplit("/", 1)[0]
    drawing_file = drawing_path.rsplit("/", 1)[1]
    drawing_rels_path = f"{drawing_dir}/_rels/{drawing_file}.rels"

    return drawing_path, drawing_rels_path


def extract_images_anchored_to_worksheet(
    excel_path: str, sheet_name: str
) -> dict[str, bytes]:
    """
    Extract images anchored to *sheet_name* from *excel_path*.

    Returns a dict  { "F25": <bytes>, "H25": <bytes>, … }
    where the key is the cell address of the image's top-left corner.
    """
    result: dict[str, bytes] = {}

    try:
        with zipfile.ZipFile(excel_path, "r") as zf:
            paths = _get_sheet_drawing_path(zf, sheet_name)
            if not paths:
                return result

            drawing_path, drawing_rels_path = paths

            # Load drawing XML
            try:
                drawing_xml = ET.fromstring(zf.read(drawing_path))
            except Exception:
                return result

            # Load drawing relationships (image rId → file path)
            image_rels: dict[str, str] = {}
            try:
                rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
                rels_xml = ET.fromstring(zf.read(drawing_rels_path))
                for rel in rels_xml.findall(f"{{{rels_ns}}}Relationship"):
                    if "image" in rel.get("Type", "").lower():
                        image_rels[rel.get("Id", "")] = rel.get("Target", "")
            except Exception:
                return result

            drawing_dir = drawing_path.rsplit("/", 1)[0]

            # Parse TwoCellAnchor elements
            for anchor in drawing_xml.findall(".//xdr:twoCellAnchor", _NS):
                try:
                    from_el = anchor.find("xdr:from", _NS)
                    if from_el is None:
                        continue

                    col_el = from_el.find("xdr:col", _NS)
                    row_el = from_el.find("xdr:row", _NS)
                    if col_el is None or row_el is None:
                        continue

                    col_id = int(col_el.text)
                    row_id = int(row_el.text)

                    col_letter = _column_letter_from_index(col_id)
                    excel_row = row_id + 1
                    cell_address = f"{col_letter}{excel_row}"

                    # Find the blip embed relationship ID
                    blip = anchor.find(".//a:blip", _NS)
                    if blip is None:
                        continue

                    embed = blip.get(
                        f"{{{_NS['r']}}}embed"
                    )
                    if not embed:
                        continue

                    img_rel_target = image_rels.get(embed)
                    if not img_rel_target:
                        continue

                    # Resolve image path inside ZIP
                    if img_rel_target.startswith("../"):
                        img_path = "xl/" + img_rel_target[3:]
                    elif img_rel_target.startswith("/"):
                        img_path = img_rel_target.lstrip("/")
                    else:
                        img_path = f"{drawing_dir}/{img_rel_target}"

                    img_bytes = zf.read(img_path)

                    if cell_address not in result:
                        result[cell_address] = img_bytes

                except Exception:
                    continue  # skip broken anchor

    except Exception:
        pass

    return result
