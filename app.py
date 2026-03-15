"""
文档对比与编辑助手
Document Diff & Edit Assistant

完全本地运行，不调用任何外部 API。
Runs fully locally — no external API calls.
"""

import io
import re
import html as html_module
from typing import Dict, List, Optional, Tuple
import difflib

import streamlit as st

# ── Optional dependency: python-docx ──────────────────────────────────────────
try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
# FILE READING & SECTION EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

# ── Pre-compiled regex patterns ───────────────────────────────────────────────

# ATX heading: allows 0-3 leading ASCII spaces (CommonMark), then 1-6 `#`,
# then either ASCII/tab/ideographic (U+3000) whitespace OR no space at all
# (common in Chinese writing), then the heading text, then optional closing #s.
_ATX_RE = re.compile(
    r"^[ \t]{0,3}"                       # up to 3 leading spaces
    r"(#{1,6})"                           # 1–6 hash signs
    r"(?:[ \t\u3000]+|(?=[^\s#]))"        # space(s) / full-width space, or no space
    r"(.+?)$"                             # heading text (non-greedy, to EOL)
)
_SETEXT_H1_RE = re.compile(r"^[ \t]{0,3}=+[ \t]*$")   # ===
_SETEXT_H2_RE = re.compile(r"^[ \t]{0,3}-{3,}[ \t]*$") # ---

# Heading style name prefixes recognised across Word localisations
_HEADING_PREFIXES: Tuple[str, ...] = (
    "Heading",   # English
    "标题",       # Simplified Chinese (zh-CN)
    "標題",       # Traditional Chinese (zh-TW)
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _flush_section(
    sections: Dict[str, str],
    header: Optional[str],
    lines: List[str],
) -> None:
    """Commit accumulated lines into *sections* under *header*."""
    content = "\n".join(lines).strip()
    key = header if header is not None else "(序言)"
    if content or header is not None:
        sections[key] = content


def _heading_level(style_name: str) -> Optional[int]:
    """
    Return the heading level (1–6) when *style_name* matches a known Word
    heading style, or ``None`` otherwise.

    Handles English ("Heading 1"), Simplified Chinese ("标题 1"), and
    Traditional Chinese ("標題 1") Word localisations.  The trailing number
    may be separated from the prefix by a space or not.
    """
    if not style_name:
        return None
    for prefix in _HEADING_PREFIXES:
        if style_name.startswith(prefix):
            suffix = style_name[len(prefix):].strip()
            if suffix.isdigit():
                return max(1, min(6, int(suffix)))
            return 1  # bare "Heading" / "标题" with no number → H1
    return None


# ── Section extractor ─────────────────────────────────────────────────────────

def extract_sections(text: str) -> Dict[str, str]:
    """
    Split *text* into sections keyed by their heading label.

    Recognises:
    - ATX headings ``# H1`` … ``###### H6`` — 0–3 leading spaces tolerated;
      full-width ideographic space (U+3000) accepted as separator; no-space
      variant (``#标题``) also matched.
    - Setext headings: text line followed by ``===`` (H1) or ``---`` (H2).

    Falls back to ``{"(全文)": text}`` when no headings are detected.
    """
    if not text or not text.strip():
        return {"(全文)": text}

    lines = text.splitlines()
    sections: Dict[str, str] = {}
    current_header: Optional[str] = None
    current_lines:  List[str]    = []

    i = 0
    while i < len(lines):
        line = lines[i]

        # ATX-style heading
        atx = _ATX_RE.match(line)
        if atx:
            _flush_section(sections, current_header, current_lines)
            hashes = atx.group(1)
            title  = atx.group(2).strip()
            # Strip optional trailing closing hashes (e.g. "# Title ##")
            title  = re.sub(r"\s+#+\s*$", "", title).strip()
            current_header = f"{hashes} {title}"
            current_lines  = []
            i += 1
            continue

        # Setext-style heading (underline on the next line)
        if i + 1 < len(lines):
            nxt     = lines[i + 1]
            stripped = line.strip()
            if stripped and _SETEXT_H1_RE.match(nxt):
                _flush_section(sections, current_header, current_lines)
                current_header = f"# {stripped}"
                current_lines  = []
                i += 2
                continue
            if stripped and _SETEXT_H2_RE.match(nxt):
                _flush_section(sections, current_header, current_lines)
                current_header = f"## {stripped}"
                current_lines  = []
                i += 2
                continue

        current_lines.append(line)
        i += 1

    _flush_section(sections, current_header, current_lines)

    # No usable structure found → single section covering the whole document
    if not sections or list(sections.keys()) == ["(序言)"]:
        return {"(全文)": text}
    return sections


# ── Keyword-based section extractor ──────────────────────────────────────────

def _parse_keywords(raw_input: str) -> List[str]:
    """
    Parse a comma/newline-delimited string into a clean, deduplicated keyword list.
    Strips surrounding whitespace; drops empty entries; preserves insertion order.
    """
    parts = re.split(r"[,\n\r]+", raw_input)
    seen:   set       = set()
    result: List[str] = []
    for part in parts:
        kw = part.strip()
        if kw and kw not in seen:
            seen.add(kw)
            result.append(kw)
    return result


def extract_sections_by_keywords(
    text: str,
    keywords: List[str],
) -> Dict[str, str]:
    """
    Split *text* into sections by matching lines against *keywords*.

    A line becomes a section heading if it contains at least one keyword
    (case-insensitive substring match).  Special regex characters in each
    keyword are escaped so they are treated as literal strings.

    Falls back to ``{"(全文)": text}`` when no lines match any keyword.
    If *keywords* is empty, delegates to the standard ``extract_sections()``.
    """
    if not text or not text.strip():
        return {"(全文)": text}
    if not keywords:
        return extract_sections(text)

    # Build a single combined pattern; escape every keyword literally
    pattern = re.compile(
        "|".join(re.escape(kw) for kw in keywords),
        re.IGNORECASE,
    )

    lines           = text.splitlines()
    sections:       Dict[str, str] = {}
    current_header: Optional[str] = None
    current_lines:  List[str]     = []

    for line in lines:
        if pattern.search(line.strip()):
            _flush_section(sections, current_header, current_lines)
            current_header = line.strip()
            current_lines  = []
        else:
            current_lines.append(line)

    _flush_section(sections, current_header, current_lines)

    if not sections or list(sections.keys()) == ["(序言)"]:
        return {"(全文)": text}
    return sections


# ── File reader ───────────────────────────────────────────────────────────────

def read_file_structured(uploaded_file) -> Tuple[str, Dict[str, str]]:
    """
    Read an uploaded file and return ``(full_plain_text, sections_dict)``.

    *sections_dict* maps heading labels to the body text beneath each heading.

    Safety notes:
    - Rewinds the file stream to position 0 before reading (guards against a
      caller that has already partially consumed the buffer).
    - Reads the raw bytes **once** and reuses the in-memory copy for all
      subsequent parsing, eliminating any stream-pointer drift.
    - Handles empty files and common corruption errors with user-friendly
      messages rather than bare tracebacks.
    """
    # Guard: rewind in case the stream was previously consumed
    try:
        uploaded_file.seek(0)
    except (AttributeError, OSError):
        pass

    fname = uploaded_file.name
    name  = fname.lower()

    # ── .docx ──────────────────────────────────────────────────────────────
    if name.endswith(".docx"):
        if not DOCX_AVAILABLE:
            st.error(
                "需要 python-docx 才能读取 .docx 文件。\n"
                "请运行: `pip install python-docx`"
            )
            return "", {"(全文)": ""}

        # Read the entire file once into memory
        try:
            file_bytes = uploaded_file.read()
        except Exception as exc:
            st.error(f"无法读取 {fname}：{exc}")
            return "", {"(全文)": ""}

        if not file_bytes:
            st.warning(f"⚠️ {fname} 是空文件，已跳过。")
            return "", {"(全文)": ""}

        # Parse the .docx archive (raises zipfile.BadZipFile if corrupt)
        try:
            doc = DocxDocument(io.BytesIO(file_bytes))
        except Exception as exc:
            st.error(
                f"无法解析 {fname}（文件可能已损坏或不是有效的 .docx 格式）：{exc}"
            )
            return "", {"(全文)": ""}

        all_paras:      List[str] = []
        sections:       Dict[str, str] = {}
        current_header: Optional[str] = None
        current_paras:  List[str] = []

        for para in doc.paragraphs:
            ptext  = para.text
            pstyle = para.style.name if para.style else ""
            level  = _heading_level(pstyle)

            if level is not None and ptext.strip():
                _flush_section(sections, current_header, current_paras)
                current_header = f"{'#' * level} {ptext.strip()}"
                current_paras  = []
            else:
                if ptext.strip():
                    current_paras.append(ptext)
            all_paras.append(ptext)

        _flush_section(sections, current_header, current_paras)
        full_text = "\n".join(all_paras)

        if not sections or list(sections.keys()) == ["(序言)"]:
            return full_text, {"(全文)": full_text}
        return full_text, sections

    # ── .md / .txt ─────────────────────────────────────────────────────────
    try:
        raw = uploaded_file.read()
    except Exception as exc:
        st.error(f"无法读取 {fname}：{exc}")
        return "", {"(全文)": ""}

    if not raw:
        st.warning(f"⚠️ {fname} 是空文件，已跳过。")
        return "", {"(全文)": ""}

    # Decode: prefer UTF-8, fall back to Latin-1 (never raises)
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("latin-1", errors="replace")

    # Defensive check for binary / null-byte content
    if "\x00" in text:
        st.warning(f"⚠️ {fname} 包含二进制内容，部分字符可能显示异常。")
        text = text.replace("\x00", "")

    return text, extract_sections(text)


# ══════════════════════════════════════════════════════════════════════════════
# PREPROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def preprocess(
    text: str,
    ignore_case: bool,
    ignore_whitespace: bool,
    ignore_comments: bool,
) -> str:
    """Apply normalisation options before diffing."""
    if ignore_comments:
        # HTML / Markdown block comments  <!-- … -->
        text = re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)
        # LaTeX line comments  % to end-of-line
        text = re.sub(r"(?m)^\s*%.*$", "", text)
        # Tidy up blank lines introduced by removal
        text = re.sub(r"\n{3,}", "\n\n", text)

    if ignore_whitespace:
        lines = [ln.strip() for ln in text.splitlines()]
        lines = [ln for ln in lines if ln]          # drop empty lines
        text = "\n".join(lines)

    if ignore_case:
        text = text.lower()

    return text


# ══════════════════════════════════════════════════════════════════════════════
# DIFF ENGINE
# ══════════════════════════════════════════════════════════════════════════════

# Type aliases
SBSRow  = Tuple[str, Optional[int], str]   # (tag, line_num, html_content)
ILRow   = Tuple[str, Optional[int], Optional[int], str]  # (tag, ln1, ln2, html)


def _word_diff_html(line1: str, line2: str) -> Tuple[str, str]:
    """
    Highlight word-level changes within two replaced lines.
    Returns a pair of HTML strings (for the left panel, for the right panel).
    """
    # Split on whitespace boundaries, keeping delimiters
    words1 = re.split(r"(\s+)", line1)
    words2 = re.split(r"(\s+)", line2)
    matcher = difflib.SequenceMatcher(None, words1, words2, autojunk=False)

    left_parts:  List[str] = []
    right_parts: List[str] = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        c1 = html_module.escape("".join(words1[i1:i2]))
        c2 = html_module.escape("".join(words2[j1:j2]))
        if tag == "equal":
            left_parts.append(c1)
            right_parts.append(c2)
        elif tag == "replace":
            left_parts.append(f'<mark class="wd-del">{c1}</mark>')
            right_parts.append(f'<mark class="wd-ins">{c2}</mark>')
        elif tag == "delete":
            left_parts.append(f'<mark class="wd-del">{c1}</mark>')
        elif tag == "insert":
            right_parts.append(f'<mark class="wd-ins">{c2}</mark>')

    return "".join(left_parts), "".join(right_parts)


def build_side_by_side(text1: str, text2: str) -> Tuple[List[SBSRow], List[SBSRow]]:
    """
    Produce two aligned lists of rows for the side-by-side view.
    Tags: 'equal' | 'delete' | 'insert' | 'replace' | 'empty'
    """
    lines1 = text1.splitlines()
    lines2 = text2.splitlines()
    matcher = difflib.SequenceMatcher(None, lines1, lines2, autojunk=False)

    left:  List[SBSRow] = []
    right: List[SBSRow] = []
    ln1 = ln2 = 1

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for a, b in zip(lines1[i1:i2], lines2[j1:j2]):
                left.append(("equal", ln1, html_module.escape(a)))
                right.append(("equal", ln2, html_module.escape(b)))
                ln1 += 1; ln2 += 1

        elif tag == "replace":
            chunk1, chunk2 = lines1[i1:i2], lines2[j1:j2]
            n = max(len(chunk1), len(chunk2))
            for k in range(n):
                has1, has2 = k < len(chunk1), k < len(chunk2)
                if has1 and has2:
                    h1, h2 = _word_diff_html(chunk1[k], chunk2[k])
                    left.append(("replace", ln1, h1));  ln1 += 1
                    right.append(("replace", ln2, h2)); ln2 += 1
                elif has1:
                    left.append(("delete", ln1, html_module.escape(chunk1[k]))); ln1 += 1
                    right.append(("empty", None, ""))
                else:
                    left.append(("empty", None, ""))
                    right.append(("insert", ln2, html_module.escape(chunk2[k]))); ln2 += 1

        elif tag == "delete":
            for line in lines1[i1:i2]:
                left.append(("delete", ln1, html_module.escape(line))); ln1 += 1
                right.append(("empty", None, ""))

        elif tag == "insert":
            for line in lines2[j1:j2]:
                left.append(("empty", None, ""))
                right.append(("insert", ln2, html_module.escape(line))); ln2 += 1

    return left, right


def build_inline(text1: str, text2: str) -> List[ILRow]:
    """Produce a flat list of rows for the inline (unified) view."""
    lines1 = text1.splitlines()
    lines2 = text2.splitlines()
    matcher = difflib.SequenceMatcher(None, lines1, lines2, autojunk=False)

    result: List[ILRow] = []
    ln1 = ln2 = 1

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for line in lines1[i1:i2]:
                result.append(("equal", ln1, ln2, html_module.escape(line)))
                ln1 += 1; ln2 += 1
        elif tag == "replace":
            for line in lines1[i1:i2]:
                result.append(("delete", ln1, None, html_module.escape(line))); ln1 += 1
            for line in lines2[j1:j2]:
                result.append(("insert", None, ln2, html_module.escape(line))); ln2 += 1
        elif tag == "delete":
            for line in lines1[i1:i2]:
                result.append(("delete", ln1, None, html_module.escape(line))); ln1 += 1
        elif tag == "insert":
            for line in lines2[j1:j2]:
                result.append(("insert", None, ln2, html_module.escape(line))); ln2 += 1

    return result


def compute_stats(text1: str, text2: str) -> Dict:
    """Return summary statistics for the diff."""
    lines1 = text1.splitlines()
    lines2 = text2.splitlines()
    matcher = difflib.SequenceMatcher(None, lines1, lines2, autojunk=False)

    added = deleted = unchanged = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            unchanged += i2 - i1
        elif tag == "replace":
            deleted += i2 - i1
            added   += j2 - j1
        elif tag == "delete":
            deleted += i2 - i1
        elif tag == "insert":
            added += j2 - j1

    return {
        "added":      added,
        "deleted":    deleted,
        "unchanged":  unchanged,
        "similarity": round(matcher.ratio() * 100, 1),
        "lines1":     len(lines1),
        "lines2":     len(lines2),
    }


# ══════════════════════════════════════════════════════════════════════════════
# HTML RENDERING
# ══════════════════════════════════════════════════════════════════════════════

_CSS = """
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
       background: #ffffff; font-size: 13px; }

/* ── Word-level highlights ─────────────────────────────────────────────── */
mark.wd-del { background: #ffb3ba; color: #000; border-radius: 2px;
              padding: 1px 2px; font-style: normal; }
mark.wd-ins { background: #b3ffbc; color: #000; border-radius: 2px;
              padding: 1px 2px; font-style: normal; }

/* ── Common ─────────────────────────────────────────────────────────────── */
.mono { font-family: "SFMono-Regular", "Consolas", "Liberation Mono",
        "Menlo", monospace; }

/* ── Side-by-side ───────────────────────────────────────────────────────── */
.sbs-wrap     { display: flex; gap: 6px; }
.sbs-panel    { flex: 1 1 0; min-width: 0; border: 1px solid #d0d7de;
               border-radius: 6px; overflow: hidden; }
.sbs-header   { background: #f6f8fa; padding: 7px 12px; font-weight: 600;
               font-size: 12.5px; border-bottom: 1px solid #d0d7de;
               white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sbs-body     { overflow-x: auto; overflow-y: visible; padding-bottom: 14px; }

/* ── Custom scrollbars (WebKit) ─────────────────────────────────────────── */
.sbs-body::-webkit-scrollbar,
.il-body::-webkit-scrollbar   { height: 8px; width: 8px; }
.sbs-body::-webkit-scrollbar-track,
.il-body::-webkit-scrollbar-track { background: #f0f2f4; border-radius: 4px; }
.sbs-body::-webkit-scrollbar-thumb,
.il-body::-webkit-scrollbar-thumb { background: #bbbfc4; border-radius: 4px; }
.sbs-body::-webkit-scrollbar-thumb:hover,
.il-body::-webkit-scrollbar-thumb:hover { background: #9198a1; }

.diff-row            { display: flex; align-items: stretch; min-height: 21px; }
.diff-row:hover      { filter: brightness(0.965); }
.diff-ln             { font-size: 11.5px; color: #848d97; min-width: 44px;
                      text-align: right; padding: 2px 10px 2px 6px;
                      user-select: none; border-right: 1px solid #d0d7de;
                      flex-shrink: 0; line-height: 17px; }
.diff-code           { padding: 2px 10px 2px 14px; flex: 1;
                      white-space: pre-wrap; word-break: break-word;
                      line-height: 17px; }

/* row-level colours */
.row-equal            { background: #ffffff; }
.row-delete           { background: #ffeef0; }
.row-delete .diff-ln  { background: #ffd7d9; }
.row-insert           { background: #e6ffed; }
.row-insert .diff-ln  { background: #ccffd8; }
.row-replace-l        { background: #fff5b1; }
.row-replace-l .diff-ln { background: #fce77c; }
.row-replace-r        { background: #dafbdf; }
.row-replace-r .diff-ln { background: #aff5bc; }
.row-empty            { background: #f6f8fa; }
.row-empty .diff-ln   { background: #f0f2f4; }

/* diff symbols */
.sym-del { color: #cf222e; font-weight: bold; margin-right: 5px;
           font-family: monospace; }
.sym-ins { color: #1a7f37; font-weight: bold; margin-right: 5px;
           font-family: monospace; }
.sym-chg { color: #9a6700; font-weight: bold; margin-right: 5px;
           font-family: monospace; }

/* ── Inline ─────────────────────────────────────────────────────────────── */
.il-wrap            { border: 1px solid #d0d7de; border-radius: 6px;
                     overflow: hidden; }
.il-header          { background: #f6f8fa; padding: 7px 12px; font-weight: 600;
                     font-size: 12.5px; border-bottom: 1px solid #d0d7de; }
.il-body            { overflow-x: auto; overflow-y: visible; padding-bottom: 14px; }

.il-row             { display: flex; align-items: stretch; min-height: 21px; }
.il-row:hover       { filter: brightness(0.965); }
.il-lns             { display: flex; border-right: 1px solid #d0d7de;
                     flex-shrink: 0; }
.il-ln              { font-size: 11.5px; color: #848d97; min-width: 44px;
                     text-align: right; padding: 2px 8px;
                     user-select: none; line-height: 17px; }
.il-sep             { width: 1px; background: #d0d7de; }
.il-sym             { min-width: 22px; text-align: center; padding: 2px 3px;
                     flex-shrink: 0; line-height: 17px;
                     font-family: monospace; font-weight: bold; }
.il-code            { padding: 2px 10px 2px 6px; flex: 1;
                     white-space: pre-wrap; word-break: break-word;
                     line-height: 17px; }

.ilrow-equal            { background: #ffffff; }
.ilrow-delete           { background: #ffeef0; }
.ilrow-delete .il-lns   { background: #ffd7d9; }
.ilrow-delete .il-sym   { color: #cf222e; }
.ilrow-insert           { background: #e6ffed; }
.ilrow-insert .il-lns   { background: #ccffd8; }
.ilrow-insert .il-sym   { color: #1a7f37; }
</style>
"""


def _sbs_row(tag: str, ln: Optional[int], content: str, side: str) -> str:
    """Render one row in a side-by-side panel."""
    ln_str = str(ln) if ln is not None else ""
    body   = content if content else "&nbsp;"

    if tag == "equal":
        row_cls = "row-equal"
        sym = ""
    elif tag == "delete":
        row_cls = "row-delete"
        sym = '<span class="sym-del">−</span>'
    elif tag == "insert":
        row_cls = "row-insert"
        sym = '<span class="sym-ins">+</span>'
    elif tag == "replace":
        row_cls = f"row-replace-{side}"
        sym = '<span class="sym-chg">~</span>'
    else:  # empty
        row_cls = "row-empty"
        sym = ""

    return (
        f'<div class="diff-row {row_cls}">'
        f'<span class="diff-ln mono">{ln_str}</span>'
        f'<span class="diff-code mono">{sym}{body}</span>'
        f'</div>\n'
    )


def render_side_by_side(
    left: List[SBSRow],
    right: List[SBSRow],
    name1: str,
    name2: str,
) -> str:
    n1 = html_module.escape(name1)
    n2 = html_module.escape(name2)

    left_rows  = "".join(_sbs_row(t, ln, c, "l") for t, ln, c in left)
    right_rows = "".join(_sbs_row(t, ln, c, "r") for t, ln, c in right)

    return f"""{_CSS}
<div class="sbs-wrap">
  <div class="sbs-panel">
    <div class="sbs-header">&#128196; {n1}</div>
    <div class="sbs-body">{left_rows}</div>
  </div>
  <div class="sbs-panel">
    <div class="sbs-header">&#128196; {n2}</div>
    <div class="sbs-body">{right_rows}</div>
  </div>
</div>
"""


def render_inline(rows: List[ILRow], name1: str, name2: str) -> str:
    n1 = html_module.escape(name1)
    n2 = html_module.escape(name2)

    parts: List[str] = []
    for tag, ln1, ln2, content in rows:
        ln1_s = str(ln1) if ln1 is not None else ""
        ln2_s = str(ln2) if ln2 is not None else ""
        body  = content if content else "&nbsp;"

        if tag == "equal":
            row_cls = "ilrow-equal";  sym = "&nbsp;"
        elif tag == "delete":
            row_cls = "ilrow-delete"; sym = "−"
        else:
            row_cls = "ilrow-insert"; sym = "+"

        parts.append(
            f'<div class="il-row {row_cls}">'
            f'<div class="il-lns">'
            f'<span class="il-ln mono">{ln1_s}</span>'
            f'<span class="il-sep"></span>'
            f'<span class="il-ln mono">{ln2_s}</span>'
            f'</div>'
            f'<span class="il-sym">{sym}</span>'
            f'<span class="il-code mono">{body}</span>'
            f'</div>\n'
        )

    return f"""{_CSS}
<div class="il-wrap">
  <div class="il-header">&#128196; {n1} &nbsp;&#8594;&nbsp; {n2}&ensp;(Inline Diff)</div>
  <div class="il-body">{"".join(parts)}</div>
</div>
"""


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT APP
# ══════════════════════════════════════════════════════════════════════════════

def _estimate_height(n_diff_rows: int) -> int:
    """
    Compute a sensible iframe height for the diff view.
    Row height is ~21 px; header adds ~40 px; bottom scrollbar ~14 px.
    Clamp between 240 px (small screens) and 680 px (large screens) so the
    frame never dominates the viewport on either extreme.
    """
    raw = n_diff_rows * 21 + 56
    # For very small diffs keep a compact minimum; for large diffs cap earlier
    # on small screens by using a tighter ceiling when there are few rows.
    lo  = 240
    hi  = 680 if n_diff_rows > 30 else min(raw + 40, 480)
    return min(max(raw, lo), hi)


def main() -> None:
    st.set_page_config(
        page_title="文档对比助手",
        page_icon="📝",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # Light global style tweaks
    st.markdown(
        """
        <style>
        [data-testid="stMetric"] {
            border: 1px solid #e6e9ef;
            border-radius: 8px;
            padding: 10px 16px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("📝 文档对比与编辑助手")
    st.caption("完全本地运行 · 隐私安全 · 支持 `.docx` / `.md` / `.txt`")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ 预处理选项")
        st.markdown("对比前对文本进行规范化")

        ignore_case = st.toggle(
            "🔡 忽略大小写",
            value=False,
            help="将所有字符转为小写后再对比",
        )
        ignore_whitespace = st.toggle(
            "⬜ 忽略空白 / 空行",
            value=False,
            help="去除每行首尾空格，并删除空行",
        )
        ignore_comments = st.toggle(
            "💬 忽略注释",
            value=False,
            help="移除 HTML/Markdown 注释 `<!-- -->` 和 LaTeX 行注释 `%`",
        )

        st.divider()

        st.header("🔍 自定义标题识别")
        st.caption("覆盖 Markdown `#` 与 Word 样式的自动识别")
        custom_kw_raw = st.text_area(
            "标题关键词（留空则自动识别）",
            value="",
            height=110,
            placeholder="一行一个，或用逗号分隔\n\n示例：\n摘要\n引言, 背景\n方法论\n结论",
            help=(
                "包含这些关键词的行将被识别为章节标题（大小写不敏感，子串匹配）。\n\n"
                "特殊字符（如括号、点号）会被自动转义，无需手动处理。\n\n"
                "留空则使用 Markdown `#` 标题 / Word 标题样式自动识别。"
            ),
            key="custom_kw_input",
            label_visibility="collapsed",
        )
        custom_keywords = _parse_keywords(custom_kw_raw)
        if custom_keywords:
            st.caption(
                f"已启用自定义识别：{len(custom_keywords)} 个关键词  \n"
                + " · ".join(f"`{kw}`" for kw in custom_keywords[:6])
                + (" …" if len(custom_keywords) > 6 else "")
            )

        st.divider()

        st.header("🖥️ 视图模式")
        view_mode = st.radio(
            "对比视图",
            ["📐 并排对比 (Side-by-Side)", "📋 行内对比 (Inline)"],
            help=(
                "**并排**：左右面板分别显示两个文件，逐行对齐。\n\n"
                "**行内**：在同一列显示增删行（类似 git diff）。"
            ),
        )

        st.divider()
        st.markdown(
            "🔒 **隐私保障**\n\n"
            "所有处理均在本地内存完成，\n"
            "不调用任何外部 API 或网络服务。"
        )

        if not DOCX_AVAILABLE:
            st.warning(
                "⚠️ 未安装 python-docx，\n"
                "`.docx` 文件暂不可用。\n\n"
                "安装命令：\n```\npip install python-docx\n```"
            )

    # ── File Upload ───────────────────────────────────────────────────────────
    st.subheader("📂 上传文件（2 ~ 3 个）")

    allowed_types = ["md", "txt"] + (["docx"] if DOCX_AVAILABLE else [])
    cols = st.columns(3)
    uploaded: List = []
    for i, col in enumerate(cols):
        with col:
            f = st.file_uploader(
                f"文件 {i + 1}",
                type=allowed_types,
                key=f"upload_{i}",
            )
            if f is not None:
                uploaded.append(f)

    if len(uploaded) < 2:
        st.info("👆 请至少上传 **2 个文件** 以开始对比。")
        st.stop()

    # ── Read & Extract Sections ───────────────────────────────────────────────
    file_data: Dict[str, Dict] = {}
    for f in uploaded:
        raw, sections = read_file_structured(f)
        # Custom keywords override auto-detection for all file types
        if custom_keywords:
            sections = extract_sections_by_keywords(raw, custom_keywords)
        file_data[f.name] = {"raw": raw, "sections": sections}

    names = list(file_data.keys())

    # ── Select Pair ───────────────────────────────────────────────────────────
    if len(names) == 3:
        st.subheader("🔀 选择对比文件对")
        pair_labels = [
            f"{names[0]}  ←→  {names[1]}",
            f"{names[0]}  ←→  {names[2]}",
            f"{names[1]}  ←→  {names[2]}",
        ]
        pair_map = [(0, 1), (0, 2), (1, 2)]
        sel = st.selectbox("选择两个文件", pair_labels, label_visibility="collapsed")
        i1, i2 = pair_map[pair_labels.index(sel)]
    else:
        i1, i2 = 0, 1

    name1, name2 = names[i1], names[i2]
    headers1 = list(file_data[name1]["sections"].keys())
    headers2 = list(file_data[name2]["sections"].keys())

    # ── Section Filter ────────────────────────────────────────────────────────
    has_structure = len(headers1) > 1 or len(headers2) > 1

    sel1: List[str] = headers1
    sel2: List[str] = headers2

    if has_structure:
        with st.expander("📑 区域筛选（可选）", expanded=False):
            st.caption(
                "选择参与对比的章节。默认选中全部；取消勾选某节后，"
                "对比时将跳过该节内容。"
            )
            fc1, fc2 = st.columns(2)
            with fc1:
                sel1 = st.multiselect(
                    f"📄 {name1}",
                    options=headers1,
                    default=headers1,
                    key=f"sec_{name1}_{name2}_1",
                    placeholder="（选择章节…）",
                )
            with fc2:
                sel2 = st.multiselect(
                    f"📄 {name2}",
                    options=headers2,
                    default=headers2,
                    key=f"sec_{name1}_{name2}_2",
                    placeholder="（选择章节…）",
                )

            # Show a quick preview of which sections are active
            if sel1 != headers1 or sel2 != headers2:
                active1 = len(sel1) or len(headers1)
                active2 = len(sel2) or len(headers2)
                st.info(
                    f"当前筛选：**{name1}** 选中 {active1}/{len(headers1)} 节 · "
                    f"**{name2}** 选中 {active2}/{len(headers2)} 节",
                    icon="📑",
                )

            if not sel1:
                sel1 = headers1          # guard: empty → use all
            if not sel2:
                sel2 = headers2

    # Assemble the text slices to compare, then preprocess
    raw1 = "\n\n".join(
        file_data[name1]["sections"][h] for h in sel1
        if h in file_data[name1]["sections"]
    )
    raw2 = "\n\n".join(
        file_data[name2]["sections"][h] for h in sel2
        if h in file_data[name2]["sections"]
    )
    text1 = preprocess(raw1, ignore_case, ignore_whitespace, ignore_comments)
    text2 = preprocess(raw2, ignore_case, ignore_whitespace, ignore_comments)

    # ── Statistics ────────────────────────────────────────────────────────────
    stats = compute_stats(text1, text2)

    st.subheader("📊 差异统计")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("相似度", f"{stats['similarity']}%")
    m2.metric(f"{name1} 行数", stats["lines1"])
    m3.metric(f"{name2} 行数", stats["lines2"])
    m4.metric("新增行 ＋", stats["added"])
    m5.metric("删除行 −", stats["deleted"])

    # ── Diff View ─────────────────────────────────────────────────────────────
    st.subheader("🔍 差异视图")

    if "并排" in view_mode:
        left_rows, right_rows = build_side_by_side(text1, text2)
        diff_html = render_side_by_side(left_rows, right_rows, name1, name2)
        n_rows    = len(left_rows)
    else:
        inline_rows = build_inline(text1, text2)
        diff_html   = render_inline(inline_rows, name1, name2)
        n_rows      = len(inline_rows)

    if stats["lines1"] == 0 and stats["lines2"] == 0:
        st.warning("两个文件在当前预处理设置下均为空白内容。")
    elif stats["similarity"] == 100.0:
        st.success("✅ 两个文件完全相同（当前预处理设置下）。")
    else:
        st.info(
            "💡 请使用下方对比框内部的滚动条进行查看。"
            "并排视图中，两侧面板可独立横向滚动。",
            icon="ℹ️",
        )
        st.components.v1.html(diff_html, height=_estimate_height(n_rows), scrolling=True)

    # ── Editor ────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("✏️ 在线编辑器")

    hdr_l, hdr_r = st.columns([4, 2])
    with hdr_l:
        st.caption("选择基础文件进行修改，支持 Markdown 格式，修改后可直接导出。")
    with hdr_r:
        edit_base = st.selectbox(
            "基础文件",
            names,
            key="edit_base_select",
            label_visibility="collapsed",
        )

    # Per-file editor state — preserves edits when switching files
    ss_init_key    = f"editor_init_{edit_base}"
    ss_content_key = f"editor_content_{edit_base}"

    if ss_init_key not in st.session_state:
        st.session_state[ss_content_key] = file_data[edit_base]["raw"]
        st.session_state[ss_init_key]    = True

    edited = st.text_area(
        "编辑区",
        key=ss_content_key,
        height=420,
        label_visibility="collapsed",
        placeholder="在此编辑文档内容…",
    )

    # Reset button
    if st.button("↩️ 重置为原始内容", help="丢弃所有编辑，恢复上传时的原始文件内容"):
        st.session_state[ss_content_key] = file_data[edit_base]["raw"]
        st.rerun()

    # Word / character count
    wc = len(edited.split())
    cc = len(edited)
    st.caption(f"字数：**{wc}**&ensp;·&ensp;字符数：**{cc}**")

    # ── Export ────────────────────────────────────────────────────────────────
    st.subheader("💾 导出")

    base_stem = edit_base.rsplit(".", 1)[0]
    ex1, ex2, _ = st.columns([2, 2, 3])

    with ex1:
        st.download_button(
            label="⬇️ 导出为 Markdown (.md)",
            data=edited.encode("utf-8"),
            file_name=f"{base_stem}_edited.md",
            mime="text/markdown",
            use_container_width=True,
        )
    with ex2:
        st.download_button(
            label="⬇️ 导出为纯文本 (.txt)",
            data=edited.encode("utf-8"),
            file_name=f"{base_stem}_edited.txt",
            mime="text/plain",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
