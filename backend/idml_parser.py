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
from typing import Dict, List, Optional, Tuple
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

        # Collect (spread_idx, page_bucket, y, x, story_id)
        frame_order: List[Tuple[int, int, float, float, str]] = []
        for spread_idx, spread_file in enumerate(spread_files):
            if spread_file not in all_names:
                continue
            frames = _get_frames_from_spread(zf, spread_file)
            for y, x, story_id in frames:
                # In a two-page spread, all left-page items should be read before
                # right-page items even if some right-page captions sit higher.
                page_bucket = 0 if x < 0 else 1
                frame_order.append((spread_idx, page_bucket, y, x, story_id))

        # Sort: spread order → page side (left→right) → y (top→bottom) → x
        frame_order.sort(key=lambda t: (t[0], t[1], t[2], t[3]))

        # Deduplicate: same story can appear multiple times (threaded frames)
        seen: set = set()
        paragraphs: List[str] = []
        current_spread = None
        for spread_idx, _, _, _, story_id in frame_order:
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

    We need the frame position in spread coordinates, not local frame coordinates.
    In many two-page spreads, left and right page frames reuse similar anchor values,
    and only `ItemTransform` distinguishes which page they belong to.
    """
    with zf.open(spread_file) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    frames: List[Tuple[float, float, str]] = []
    for tf in root.iter("TextFrame"):
        story_id = tf.get("ParentStory")
        if not story_id:
            continue

        tx, ty = _get_item_translation(tf.get("ItemTransform"))
        pt = tf.find(".//PathPointType")
        if pt is not None:
            try:
                anchor_x, anchor_y = map(float, pt.get("Anchor", "0 0").split())
            except ValueError:
                anchor_x, anchor_y = 0.0, 0.0
        else:
            anchor_x, anchor_y = 0.0, 0.0

        x = tx + anchor_x
        y = ty + anchor_y
        frames.append((y, x, story_id))

    return frames


def _get_item_translation(item_transform: Optional[str]) -> Tuple[float, float]:
    """Extract translation (tx, ty) from an InDesign ItemTransform matrix."""
    if not item_transform:
        return 0.0, 0.0

    parts = item_transform.split()
    if len(parts) != 6:
        return 0.0, 0.0

    try:
        tx = float(parts[4])
        ty = float(parts[5])
    except ValueError:
        return 0.0, 0.0

    return tx, ty


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
    page_chunks = _split_into_page_chunks(paragraphs)
    units: List[Dict[str, str]] = []

    for page_no, chapter, chunk in page_chunks:
        result = _parse_page_chunk(page_no, chunk)
        if any(result.values()):
            result["chapter"] = chapter
            units.append(result)
            print(f"[DEBUG] Unit {len(units)} parsed from page {page_no}\n")

    return units


def _split_into_page_chunks(paragraphs: List[str]) -> List[Tuple[str, str, List[str]]]:
    """Split flat paragraph stream into page-sized chunks using page numbers."""
    chunks: List[Tuple[str, str, List[str]]] = []
    current_page: Optional[str] = None
    current_chapter = "default"
    pending_chapter: Optional[str] = None
    current_chunk: List[str] = []

    for text in paragraphs:
        if text == _SPREAD_SEP:
            continue

        if text in _KOREAN_CHAPTER_NAMES:
            pending_chapter = text
            continue

        if text.lower().startswith("chapter"):
            continue

        if _is_page_marker(text):
            if current_page is not None:
                chunks.append((current_page, current_chapter, current_chunk))
            if pending_chapter is not None:
                current_chapter = pending_chapter
                pending_chapter = None
            current_page = text
            current_chunk = []
            continue

        if current_page is None:
            continue

        current_chunk.append(text)

    if current_page is not None:
        chunks.append((current_page, current_chapter, current_chunk))

    return chunks


def _parse_page_chunk(page_no: str, chunk: List[str]) -> Dict[str, str]:
    """Parse one page chunk into passage/q1/q2/mission/syntax fields."""
    print(f"[DEBUG] page={page_no} size={len(chunk)}")

    passage_lines: List[str] = []
    q1_lines: List[str] = []
    q2_lines: List[str] = []
    q3_lines: List[str] = []
    mission_lines: List[str] = []
    syntax_lines: List[str] = []

    current_section = "passage"

    cleaned = _normalize_choice_lines([text for text in chunk if text != _SPREAD_SEP])

    for idx, text in enumerate(cleaned):
        next_text = _next_meaningful_text(cleaned, idx + 1)
        future_texts = _next_meaningful_texts(cleaned, idx + 1, limit=3)

        if text == _SPREAD_SEP or text == "Logic Note":
            continue
        if text in ("A", "B"):
            continue
        if _is_page_marker(text):
            break

        if text == "Mission":
            current_section = "mission"
            continue
        if text == "Syntax":
            current_section = "syntax"
            continue

        if text.startswith("1."):
            current_section = "q1"
        elif text.startswith("2."):
            current_section = "q2"
        elif text.startswith("3."):
            current_section = "q3"
        elif _looks_like_question_start(text, next_text, future_texts):
            if current_section == "passage":
                current_section = "q1"
            elif current_section == "q1":
                current_section = "q2"
            elif current_section == "q2":
                current_section = "q3"

        if current_section == "passage":
            passage_lines.append(text)
        elif current_section == "q1":
            q1_lines.append(text)
        elif current_section == "q2":
            q2_lines.append(text)
        elif current_section == "q3":
            q3_lines.append(text)
        elif current_section == "mission":
            mission_lines.append(text)
        elif current_section == "syntax":
            syntax_lines.append(text)

    result = {
        "passage": "\n".join(passage_lines).strip(),
        "q1": "\n".join(q1_lines).strip(),
        "q2": "\n".join(q2_lines).strip(),
        "q3": "\n".join(q3_lines).strip(),
        "mission": "\n".join(mission_lines).strip(),
        "syntax": "\n".join(syntax_lines).strip(),
    }

    print(f"[DEBUG]   → passage: {result['passage'][:60]!r}")
    print(f"[DEBUG]   → q1: {result['q1'][:60]!r}")
    print(f"[DEBUG]   → q2: {result['q2'][:60]!r}")
    print(f"[DEBUG]   → q3: {result['q3'][:60]!r}")
    print(f"[DEBUG]   → mission: {result['mission'][:60]!r}")
    print(f"[DEBUG]   → syntax: {result['syntax'][:60]!r}")
    return result


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


def _is_page_marker(text: str) -> bool:
    t = text.strip()
    return t.isascii() and t.isdigit() and len(t) <= 3


def _looks_like_choice_marker(text: str) -> bool:
    return text.startswith(("①", "②", "③", "④", "⑤"))


def _looks_like_question_start(
    text: str,
    next_text: Optional[str],
    future_texts: Optional[List[str]] = None,
) -> bool:
    if text.startswith(("1.", "2.")):
        return True
    if text.startswith("3."):
        return True
    if _looks_like_choice_marker(text):
        return False
    future_texts = future_texts or []

    question_prompt_keywords = (
        "choose the best place",
        "where does the following sentence fit",
        "where should the following sentence go",
        "choose the statement",
        "choose the best title",
        "which of the following",
        "what is the main idea",
        "what is the passage mainly about",
        "밑줄 친",
        "문맥상",
        "내용과 일치하지 않는",
    )
    lowered = text.lower()
    if any(keyword in lowered for keyword in question_prompt_keywords):
        if any(_looks_like_choice_marker(item) for item in future_texts):
            return True

    if next_text is None or not any(
        _looks_like_choice_marker(item) for item in [next_text, *future_texts]
    ):
        return False

    # Long prose paragraphs can contain question marks inside the passage
    # ("Who owns the Moon?", "But has modern love really set us free?") while
    # the actual question prompt appears later. Treat generic keyword matching
    # as a question start only for prompt-sized lines, not passage-sized blocks.
    if _looks_like_long_prose(text):
        return False

    question_keywords = (
        "?",
        "것은",
        "고르시오",
        "알맞은",
        "적절",
        "옳은",
        "틀린",
        "main idea",
        "main topic",
        "best title",
        "true according to the passage",
        "NOT true",
        "best fits",
        "most appropriate",
        "can be inferred",
        "mainly about",
        "summary",
        "supported by",
        "identical with",
        "refers to",
    )
    return any(keyword.lower() in lowered for keyword in question_keywords)


def _next_meaningful_text(chunk: List[str], start: int) -> Optional[str]:
    for i in range(start, len(chunk)):
        text = chunk[i]
        if text == _SPREAD_SEP or text == "Logic Note":
            continue
        if text in ("A", "B", "Mission", "Syntax"):
            continue
        return text
    return None


def _next_meaningful_texts(chunk: List[str], start: int, limit: int) -> List[str]:
    items: List[str] = []
    for i in range(start, len(chunk)):
        text = chunk[i]
        if text == _SPREAD_SEP or text == "Logic Note":
            continue
        if text in ("A", "B", "Mission", "Syntax"):
            continue
        items.append(text)
        if len(items) >= limit:
            break
    return items


def _looks_like_long_prose(text: str) -> bool:
    stripped = text.strip()
    if len(stripped) >= 220:
        return True

    sentence_markers = stripped.count(". ") + stripped.count("? ") + stripped.count("! ")
    if sentence_markers >= 3:
        return True

    return stripped.count(";") >= 2


def _normalize_choice_lines(lines: List[str]) -> List[str]:
    normalized: List[str] = []
    i = 0

    while i < len(lines):
        text = lines[i].strip()

        if text in {"①", "②", "③", "④", "⑤"}:
            if i + 1 < len(lines):
                next_text = lines[i + 1].strip()
                if next_text and not _looks_like_choice_marker(next_text):
                    normalized.append(f"{text} {next_text}")
                    i += 2
                    continue

        normalized.append(lines[i])
        i += 1

    return normalized


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
