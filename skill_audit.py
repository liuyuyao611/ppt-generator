# Adapted from the user-provided local paper-to-editable-slides audit helper.
# Byte input support keeps uploaded documents and checks in session memory.
"""Check SVG/native PPTX text-object invariants, not visual fidelity."""
import argparse
import io
import json
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

S = "http://www.w3.org/2000/svg"
N = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
     "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def normalized(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def audit(svg_path: Path | bytes, pptx_path: Path | bytes | None, ids: list[str]) -> dict:
    root = ET.fromstring(svg_path) if isinstance(svg_path, bytes) else ET.parse(svg_path).getroot()
    errors, warnings = [], []
    if root.tag != f"{{{S}}}svg":
        errors.append("Not a namespaced SVG document")
    texts = list(root.iter(f"{{{S}}}text"))
    tspans = list(root.iter(f"{{{S}}}tspan"))
    if not texts:
        errors.append("SVG has no editable text objects")
    if tspans:
        errors.append("SVG violates the no-tspan contract")
    for tag in ("foreignObject", "textPath", "script", "style"):
        if list(root.iter(f"{{{S}}}{tag}")):
            errors.append("Unsupported SVG content: " + tag)
    if root.get("transform"):
        errors.append("Root transform changes the design-coordinate scale")
    transformed = sum(1 for node in root.iter() if node.get("transform"))
    if transformed:
        warnings.append("Transforms present; review nested/group scaling")
    if list(root.iter(f"{{{S}}}path")):
        warnings.append("Paths may be valid editable geometry; manually verify no text was outlined")
    all_ids = [node.get("id") for node in root.iter() if node.get("id")]
    if len(all_ids) != len(set(all_ids)):
        errors.append("SVG has duplicate object IDs")
    svg_paragraphs = {}
    for identifier in ids:
        found = [node for node in texts if node.get("id") == identifier]
        if len(found) != 1:
            errors.append(f"Expected one SVG text for {identifier}; found {len(found)}")
            continue
        if len(found[0]):
            errors.append(f"Paragraph {identifier} has nested elements")
        svg_paragraphs[identifier] = normalized("".join(found[0].itertext()))
        if "\n" in (found[0].text or ""):
            warnings.append(f"Raw newline in {identifier} is not proof of SVG multiline support")
    images = list(root.iter(f"{{{S}}}image"))
    external = [node.get("href", node.get("{http://www.w3.org/1999/xlink}href", ""))
                for node in images]
    if any(not href.startswith("data:image/") for href in external):
        warnings.append("SVG has non-embedded images; distribute assets or embed them")
    summary = {"svg_text_objects": len(texts), "svg_tspan_objects": len(tspans),
               "svg_image_objects": len(images), "requested_paragraphs": ids}
    if pptx_path:
        shape_records, picture_count, norm_autofit = [], 0, 0
        with zipfile.ZipFile(io.BytesIO(pptx_path) if isinstance(pptx_path, bytes) else pptx_path) as package:
            slides = sorted(name for name in package.namelist()
                            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name))
            for slide in slides:
                tree = ET.fromstring(package.read(slide))
                picture_count += len(tree.findall(".//p:pic", N))
                norm_autofit += len(tree.findall(".//a:normAutofit", N))
                for shape in tree.findall(".//p:sp", N):
                    body = shape.find("p:txBody", N)
                    if body is None:
                        continue
                    name_node = shape.find("p:nvSpPr/p:cNvPr", N)
                    shape_name = name_node.get("name", "") if name_node is not None else ""
                    paragraphs = []
                    for paragraph in body.findall("a:p", N):
                        # Soft line breaks and native paragraphs stay inside one shape.
                        parts = []
                        for child in paragraph:
                            if child.tag == f"{{{N['a']}}}br":
                                parts.append("\n")
                            else:
                                parts.extend(node.text or "" for node in child.iter(f"{{{N['a']}}}t"))
                        paragraphs.append("".join(parts))
                    shape_records.append({"slide": slide, "name": shape_name,
                                          "text": normalized("\n".join(paragraphs)),
                                          "paragraph_count": len(paragraphs)})
        if not shape_records:
            errors.append("PPTX has no native text boxes; raster-only output is not editable")
        if norm_autofit:
            warnings.append("PPTX uses normAutofit; inspect fontScale and unintended font shrinking")
        for identifier in ids:
            found = [record for record in shape_records if record["name"] == identifier]
            if len(found) != 1:
                errors.append(f"Expected one native text box named {identifier}; found {len(found)}")
            elif identifier in svg_paragraphs and found[0]["text"] != svg_paragraphs[identifier]:
                errors.append("SVG/PPTX paragraph text differs: " + identifier)
        summary.update(pptx_slides=len(slides), pptx_text_boxes=len(shape_records),
                       pptx_pictures=picture_count, pptx_norm_autofit=norm_autofit)
    return {"passed": not errors, "summary": summary, "errors": errors, "warnings": warnings,
            "boundary": "Structure only; no native PowerPoint conversion, rendering or facts verified"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--svg", type=Path, required=True)
    parser.add_argument("--pptx", type=Path)
    parser.add_argument("--paragraph-id", action="append", default=[])
    args = parser.parse_args()
    try:
        result = audit(args.svg, args.pptx, args.paragraph_id)
        print(json.dumps(result))
        return 0 if result["passed"] else 2
    except (ValueError, OSError, ET.ParseError, zipfile.BadZipFile, KeyError) as error:
        print(json.dumps({"passed": False, "errors": [str(error)]}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
