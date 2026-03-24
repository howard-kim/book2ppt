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
        for _, _, _, story_id in frame_order:
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
                y, x = map(float, pt.get("Anchor", "0 0").split())
            except ValueError:
                y, x = 0.0, 0.0
        else:
            y, x = 0.0, 0.0
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

def _build_units(paragraphs: List[str]) -> List[Dict[str, str]]:
    units: List[Dict[str, str]] = []
    i = 0
    n = len(paragraphs)

    while i < n:
        # Skip non-content paragraphs until we find a passage start
        # (i.e. skip anything that looks like a header/label before content)
        if _is_skip_line(paragraphs[i]):
            i += 1
            continue

        print(f"[DEBUG] index={i}  text={paragraphs[i][:60]!r}")
        result = _parse_one_unit(paragraphs, i)
        units.append(result["data"])
        i = result["next_index"]
        print(f"[DEBUG] Unit {len(units)} parsed. Next index → {i}\n")

    return units


def _is_skip_line(text: str) -> bool:
    """
    Return True for lines that are structural labels / decorations,
    not part of content units.
    Skips: page numbers, chapter headers, single characters/numbers,
           'Logic Note', 'A'/'B' answer labels, circled numbers (①~⑤), etc.
    """
    t = text.strip()
    if not t:
        return True
    # Single character (answer labels like A, B, circled numbers)
    if len(t) == 1:
        return True
    # Circled numbers ①②③④⑤ or similar
    if t in ("①", "②", "③", "④", "⑤"):
        return True
    # Common structural labels
    if t in ("Logic Note", "CONTENTS", "CONTENTSc"):
        return True
    # Chapter headers like "CHAPTER 1 ...", "01c", "02c" etc.
    if t.lower().startswith("chapter"):
        return True
    if len(t) <= 4 and t[:-1].isdigit() and t.endswith("c"):
        return True
    # Pure page numbers
    if t.isdigit() and len(t) <= 3:
        return True
    return False


def _parse_one_unit(paragraphs: List[str], start: int) -> Dict:
    """
    Parse exactly one 5-field unit starting at `start`.
    Returns {"data": {...}, "next_index": int}.
    Raises Exception if any required section is absent.
    """
    n = len(paragraphs)
    i = start

    # ── PASSAGE ────────────────────────────────────────────────────────────
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
    if not passage:
        raise Exception(f"[index={start}] Missing PASSAGE before Question 1.")
    print(f"[DEBUG]   → passage detected")

    # ── QUESTION 1 ─────────────────────────────────────────────────────────
    # skip structural noise between passage and Q1
    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    if i >= n or not paragraphs[i].startswith("1."):
        raise Exception(
            f"[index={i}] Expected Question 1 ('1.'), got: {paragraphs[i]!r}"
        )

    q1_lines: List[str] = []
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
    if not q1:
        raise Exception(f"[index={i}] Missing content for Question 1.")
    print(f"[DEBUG]   → q1 detected")

    # ── QUESTION 2 ─────────────────────────────────────────────────────────
    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    if i >= n or not paragraphs[i].startswith("2."):
        raise Exception(
            f"[index={i}] Expected Question 2 ('2.'), got: {paragraphs[i]!r}"
        )

    q2_lines: List[str] = []
    while i < n:
        t = paragraphs[i]
        if _is_skip_line(t):
            i += 1
            continue
        if t in ("Mission", "Syntax"):
            break
        if t.startswith("1.") and q2_lines:
            break
        q2_lines.append(t)
        i += 1

    q2 = "\n".join(q2_lines).strip()
    if not q2:
        raise Exception(f"[index={i}] Missing content for Question 2.")
    print(f"[DEBUG]   → q2 detected")

    # ── MISSION ────────────────────────────────────────────────────────────
    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    if i >= n or paragraphs[i] != "Mission":
        raise Exception(
            f"[index={i}] Expected 'Mission', got: {paragraphs[i]!r}"
        )
    i += 1

    mission_lines: List[str] = []
    while i < n:
        t = paragraphs[i]
        if _is_skip_line(t):
            i += 1
            continue
        if t == "Syntax":
            break
        mission_lines.append(t)
        i += 1

    mission = "\n".join(mission_lines).strip()
    if not mission:
        raise Exception(f"[index={i}] Missing content for Mission.")
    print(f"[DEBUG]   → mission detected")

    # ── SYNTAX ─────────────────────────────────────────────────────────────
    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    if i >= n or paragraphs[i] != "Syntax":
        raise Exception(
            f"[index={i}] Expected 'Syntax', got: {paragraphs[i]!r}"
        )
    i += 1

    while i < n and _is_skip_line(paragraphs[i]):
        i += 1

    if i >= n:
        raise Exception(f"[index={i}] Missing content after 'Syntax'.")

    syntax = paragraphs[i].strip()
    i += 1

    if not syntax:
        raise Exception(f"[index={i-1}] Syntax content is empty.")
    print(f"[DEBUG]   → syntax detected")

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
