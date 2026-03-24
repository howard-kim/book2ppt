"""
Structured data → PowerPoint generator

Clones the 5 template slides for each unit and updates the content text boxes.
Template slide structure (by index):
  0 – Passage  : content in "TextBox 2"
  1 – Q1       : content in "TextBox 2"
  2 – Q2       : content in "TextBox 2"
  3 – Mission  : content in "TextBox 2"
  4 – Syntax   : content in "TextBox 4"
"""

import copy
from pathlib import Path
from typing import Dict, List

from pptx import Presentation
from pptx.oxml.ns import qn

_FIELDS = ["passage", "q1", "q2", "mission", "syntax"]
_CONTENT_BOX = ["TextBox 2", "TextBox 2", "TextBox 2", "TextBox 2", "TextBox 4"]


def generate_ppt(
    data: List[Dict[str, str]],
    template_path: str,
    output_path: str,
) -> None:
    tp = Path(template_path)
    if not tp.exists():
        raise FileNotFoundError(f"Template not found: {tp}")

    prs = Presentation(str(tp))

    if len(prs.slides) < 5:
        raise ValueError(
            f"template.pptx must have at least 5 slides, got {len(prs.slides)}"
        )

    # Keep references to the 5 template slides before we add new ones
    template_slides = [prs.slides[i] for i in range(5)]

    for unit_no, unit in enumerate(data, start=1):
        print(f"[INFO] Generating slides for unit {unit_no} / {len(data)} …")
        for slide_idx, (field, box_name) in enumerate(zip(_FIELDS, _CONTENT_BOX)):
            new_slide = _clone_slide(prs, template_slides[slide_idx])
            _set_text(new_slide, box_name, unit.get(field, ""))

    # Remove the original 5 template slides (now at index 0-4)
    for _ in range(5):
        _delete_slide(prs, 0)

    prs.save(str(output_path))
    print(f"[INFO] Saved → {output_path}  ({len(data) * 5} slides total)")


def _clone_slide(prs: Presentation, source) -> object:
    """Append a deep copy of source slide to the presentation."""
    new_slide = prs.slides.add_slide(source.slide_layout)
    sp_tree = new_slide.shapes._spTree
    for elem in list(sp_tree):
        sp_tree.remove(elem)
    for elem in source.shapes._spTree:
        sp_tree.append(copy.deepcopy(elem))
    return new_slide


def _set_text(slide, box_name: str, text: str) -> None:
    """Replace all text in a named text box."""
    for shape in slide.shapes:
        if shape.name == box_name and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            lines = text.split("\n") if text else [""]
            for line_no, line in enumerate(lines):
                para = tf.paragraphs[0] if line_no == 0 else tf.add_paragraph()
                para.add_run().text = line
            return
    print(f"[WARN] Text box '{box_name}' not found on slide")


def _delete_slide(prs: Presentation, slide_index: int) -> None:
    """Remove a slide by index from the presentation."""
    xml_slides = prs.slides._sldIdLst
    sldId = xml_slides[slide_index]
    rId = sldId.get(qn("r:id"))
    xml_slides.remove(sldId)
    prs.part.drop_rel(rId)
