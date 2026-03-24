"""
IDML → Structured data parser
================================
Parses an Adobe InDesign IDML file (ZIP + XML) into a list of units.

Reading order strategy:
    1. Read spreads in document order from designmap.xml
    2. Within each spread, sort TextFrames by (y, x) position — top→bottom, left→right
    3. Extract paragraph text from each story in that order
    4. Apply rule-based unit parsing on the flat paragraph list
"""

import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET

_IDPKG_NS = "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging"
_SPREAD_SEP = "\x00SPREAD\x00"   # sentinel inserted between spreads

_KOREAN_CHAPTER_NAMES = {
    "첫 문장과 링킹",
    "배경과 반전",
    "부분 링킹",
    "정보 분류",
    "체화",
}

_STRUCTURAL_LABELS = {
    "Logic Note", "CONTENTS", "CONTENTSc",
    "A", "B", "C", "D", "E",
    "①", "②", "③", "④", "⑤",
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_idml(idml_path: str) -> List[Dict[str, str]]:
    path = Path(idml_path)
    if not path.exists():
        raise FileNotFoundError(f"IDML file not found: {path}")

    paragraphs = _extract_all_paragraphs(path)
    print(f"[DEBUG] Total paragraphs extracted: {len(paragraphs)}")
    return _build_units(paragraphs)


# ---------------------------------------------------------------------------
# Step 1: IDML ZIP → flat list of paragraphs (in reading order)
# ---------------------------------------------------------------------------

def _extract_all_paragraphs(idml_path: Path) -> List[str]:
    with zipfile.ZipFile(idml_path, "r") as zf:
        all_names = set(zf.namelist())

        # Spreads in document order from designmap.xml
        spread_files = _get_spread_order(zf)
        print(f"[DEBUG] Spreads in order: {spread_files}")

        # Collect (spread_idx, y, x, story_id)
        frame_order: List[Tuple[int, float, float, str]] = []
        for spread_idx, spread_file in enumerate(spread_files):
            if spread_file not in all_names:
                continue
            frames = _get_frames_from_spread(zf, spread_file)
            for y, x, story_id in frames:
                frame_order.append((spread_idx, y, x, story_id))

        # Sort: spread order → y (top→bottom) → x (left→right)
        frame_order.sort(key=lambda t: (t[0], t[1], t[2]))

        # Deduplicate: same story can appear multiple times (threaded frames)
        seen: set = set()
        paragraphs: List[str] = []
        current_spread = None
        for spread_idx, _, _, story_id in frame_order:
            # Insert a sentinel between spreads so the parser never crosses boundaries
            if spread_idx != current_spread:
                if current_spread is not None:
                    paragraphs.append(_SPREAD_SEP)
                current_spread = spread_idx
            if story_id in seen:
                continue
            seen.add(story_id)
            story_file = f"Stories/Story_{story_id}.xml"
            if story_file not in all_names:
                continue
            story_paras = _parse_story_xml(zf, story_file)
            paragraphs.extend(story_paras)

    return paragraphs


def _get_spread_order(zf: zipfile.ZipFile) -> List[str]:
    """Return spread file paths in document order from designmap.xml."""
    with zf.open("designmap.xml") as f:
        tree = ET.parse(f)
    root = tree.getroot()
    tag = f"{{{_IDPKG_NS}}}Spread"
    return [e.get("src") for e in root.iter(tag) if e.get("src")]


def _get_frames_from_spread(
    zf: zipfile.ZipFile,
    spread_file: str,
) -> List[Tuple[float, float, str]]:
    """
    Return (y, x, story_id) for every TextFrame in this spread.
    Position is taken from the first PathPointType anchor.
    """
    with zf.open(spread_file) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    frames: List[Tuple[float, float, str]] = []
    for tf in root.iter("TextFrame"):
        story_id = tf.get("ParentStory")
        if not story_id:
            continue
        pt = tf.find(".//PathPointType")
        if pt is not None:
            try:
                x, y = map(float, pt.get("Anchor", "0 0").split())
            except ValueError:
                x, y = 0.0, 0.0
        else:
            x, y = 0.0, 0.0
        frames.append((y, x, story_id))

    return frames


def _parse_story_xml(zf: zipfile.ZipFile, story_file: str) -> List[str]:
    """
    Parse one Story XML → list of paragraph strings.

    Structure:
        <ParagraphStyleRange>          ← one paragraph
            <CharacterStyleRange>
                <Content>text</Content>
            </CharacterStyleRange>
            <Br/>                      ← soft line break within paragraph
        </ParagraphStyleRange>
    """
    with zf.open(story_file) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    paragraphs: List[str] = []
    for para in root.iter("ParagraphStyleRange"):
        parts: List[str] = []
        for child in para:
            if child.tag == "CharacterStyleRange":
                for sub in child:
                    if sub.tag == "Content" and sub.text:
                        parts.append(sub.text)
            elif child.tag == "Br":
                parts.append("\n")
        text = "".join(parts).strip()
        if text:
            paragraphs.append(text)

    return paragraphs


# ---------------------------------------------------------------------------
# Step 2: flat paragraphs → structured units (rule-based state machine)
# ---------------------------------------------------------------------------

def _is_section_marker(text: str) -> bool:
    return text in ("Mission", "Syntax")


def _build_units(paragraphs: List[str]) -> List[Dict[str, str]]:
    units: List[Dict[str, str]] = []
    i = 0
    n = len(paragraphs)

    while i < n:
        t = paragraphs[i]
        # Skip separators, structural markers, and skip lines
        if t == _SPREAD_SEP or _is_skip_line(t) or _is_section_marker(t):
            i += 1
            continue

        print(f"[DEBUG] index={i}  text={t[:60]!r}")
        result = _parse_one_unit(paragraphs, i)
        if result["data"]["passage"] or result["data"]["q1"]:
            units.append(result["data"])
            print(f"[DEBUG] Unit {len(units)} parsed. Next index → {result['next_index']}\n")
        i = result["next_index"]

    return units


def _is_skip_line(text: str) -> bool:
    """
    Return True for lines that are structural labels / decorations,
    not part of content units.
    Skips: page numbers, chapter headers, single characters/numbers,
           'Logic Note', 'A'/'B' answer labels, circled numbers (①~⑤),
           Korean chapter names, etc.
    """
    t = text.strip()
    if not t:
        return True
    # Single character tokens
    if len(t) == 1:
        return True
    # Known structural labels and Korean chapter titles
    if t in _STRUCTURAL_LABELS or t in _KOREAN_CHAPTER_NAMES:
        return True
    # Chapter headers like "CHAPTER 1 ..."
    if t.lower().startswith("chapter"):
        return True
    # Codes like "01c", "02c"
    if len(t) <= 4 and t[:-1].isdigit() and t.endswith("c"):
        return True
    # Pure page numbers
    if t.isdigit() and len(t) <= 3:
        return True
    return False


def _parse_one_unit(
    paragraphs: List[str],
    start: int,
    passage_override: "str | None" = None,
) -> Dict:
    """
    Parse exactly one 5-field unit starting at `start`.
    Returns {"data": {...}, "next_index": int}.
    Raises Exception if any required section is absent.
    """
    n = len(paragraphs)
    i = start

    # ── PASSAGE ────────────────────────────────────────────────────────────
    if passage_override:
        passage = passage_override
        print(f"[DEBUG]   → passage injected from previous unit's syntax")
    else:
        passage_lines: List[str] = []
        while i < n:
            t = paragraphs[i]
            if _is_skip_line(t):
                i += 1
                continue
            if t.startswith("1.") or t in ("Mission", "Syntax"):
                break
            passage_lines.append(t)
            i += 1

        passage = "\n".join(passage_lines).strip()
    print(f"[DEBUG]   → passage detected")

    # ── QUESTION 1 (optional) ──────────────────────────────────────────────
    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    q1_lines: List[str] = []
    if i < n and paragraphs[i].startswith("1."):
        while i < n:
            t = paragraphs[i]
            if _is_skip_line(t):
                i += 1
                continue
            if t.startswith("2.") or t in ("Mission", "Syntax"):
                break
            q1_lines.append(t)
            i += 1

    q1 = "\n".join(q1_lines).strip()
    print(f"[DEBUG]   → q1 detected")

    # ── QUESTION 2 (optional) ───────────────────────────────────────────────
    # Stop Q2 after answer choices: once we've seen ①②③④⑤ lines,
    # the next plain-prose paragraph is Syntax content (not part of Q2).
    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    _CHOICE_CHARS = set("①②③④⑤")

    q2_lines: List[str] = []
    if i < n and paragraphs[i].startswith("2."):
        has_choices = False
        while i < n:
            t = paragraphs[i]
            if _is_skip_line(t):
                i += 1
                continue
            if t in ("Mission", "Syntax"):
                break
            if t.startswith("1.") and q2_lines:
                break
            # After seeing answer choices, stop at the next prose paragraph
            if has_choices and (not t or t[0] not in _CHOICE_CHARS):
                break
            if t and t[0] in _CHOICE_CHARS:
                has_choices = True
            q2_lines.append(t)
            i += 1

    q2 = "\n".join(q2_lines).strip()
    print(f"[DEBUG]   → q2 detected")

    # ── SYNTAX + MISSION content ────────────────────────────────────────────
    # After Q2 answer choices, the next paragraphs (before Mission/Syntax labels)
    # are: [0] = Syntax content, [1] = Mission content.
    pre_label: List[str] = []
    while i < n:
        t = paragraphs[i]
        if _is_skip_line(t):
            i += 1
            continue
        if t in ("Mission", "Syntax"):
            break
        if t.startswith("1."):
            break
        pre_label.append(t)
        i += 1

    syntax  = pre_label[0] if len(pre_label) > 0 else ""
    mission = "\n".join(pre_label[1:]) if len(pre_label) > 1 else ""
    print(f"[DEBUG]   → mission: {mission[:60]!r}")
    print(f"[DEBUG]   → syntax:  {syntax[:60]!r}")

    # ── Skip Mission / Syntax labels and trailing decorations ───────────────
    while i < n:
        t = paragraphs[i]
        if _is_skip_line(t) or t in ("Mission", "Syntax"):
            i += 1
            continue
        break  # hit next unit's passage or Q1

    return {
        "data": {
            "passage": passage,
            "q1":      q1,
            "q2":      q2,
            "mission": mission,
            "syntax":  syntax,
        },
        "next_index": i,
    }
