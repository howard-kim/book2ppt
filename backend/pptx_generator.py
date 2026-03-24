"""
Structured data → PowerPoint generator
"""

from pathlib import Path
from typing import Dict, List

from pptx import Presentation
from pptx.util import Pt

LAYOUT_PASSAGE = 0
LAYOUT_Q1      = 1
LAYOUT_Q2      = 2
LAYOUT_MISSION = 3
LAYOUT_SYNTAX  = 4

PH_BODY      = 1
SYNTAX_FONT  = "Consolas"


def generate_ppt(
    data: List[Dict[str, str]],
    template_path: str,
    output_path: str,
) -> None:
    """
    Generate a .pptx file from structured unit data.
    Produces exactly 5 slides per unit: Passage → Q1 → Q2 → Mission → Syntax.
    """
    tp = Path(template_path)
    if not tp.exists():
        raise FileNotFoundError(f"Template not found: {tp}")

    prs = Presentation(str(tp))

    for unit_no, unit in enumerate(data, start=1):
        print(f"[INFO] Generating slides for unit {unit_no} / {len(data)} …")
        _add_slide(prs, LAYOUT_PASSAGE, unit["passage"])
        _add_slide(prs, LAYOUT_Q1,      unit["q1"])
        _add_slide(prs, LAYOUT_Q2,      unit["q2"])
        _add_slide(prs, LAYOUT_MISSION, unit["mission"])
        _add_slide(prs, LAYOUT_SYNTAX,  unit["syntax"], monospace=True)

    prs.save(str(output_path))
    print(f"[INFO] Saved → {output_path}  ({len(data) * 5} slides total)")


def _add_slide(
    prs: Presentation,
    layout_idx: int,
    text: str,
    monospace: bool = False,
) -> None:
    layout = prs.slide_layouts[layout_idx]
    slide  = prs.slides.add_slide(layout)

    body_ph = next(
        (ph for ph in slide.placeholders if ph.placeholder_format.idx == PH_BODY),
        None,
    )
    if body_ph is None:
        raise Exception(
            f"Layout {layout_idx} has no placeholder idx={PH_BODY}. "
            "Check template.pptx."
        )

    tf = body_ph.text_frame
    tf.word_wrap = True

    for line_no, line in enumerate(text.split("\n")):
        para = tf.paragraphs[0] if line_no == 0 else tf.add_paragraph()
        run  = para.add_run()
        run.text = line
        if monospace:
            run.font.name = SYNTAX_FONT
            run.font.size = Pt(11)
