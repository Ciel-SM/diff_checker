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
# FILE READING
# ══════════════════════════════════════════════════════════════════════════════

def read_file(uploaded_file) -> str:
    """Read an uploaded file and return its plain-text content."""
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".docx"):
            if not DOCX_AVAILABLE:
                st.error(
                    "需要 python-docx 才能读取 .docx 文件。\n"
                    "请运行: `pip install python-docx`"
                )
                return ""
            doc = DocxDocument(io.BytesIO(uploaded_file.read()))
            return "\n".join(p.text for p in doc.paragraphs)
        else:
            raw = uploaded_file.read()
            try:
                return raw.decode("utf-8")
            except UnicodeDecodeError:
                return raw.decode("latin-1", errors="replace")
    except Exception as exc:
        st.error(f"读取 {uploaded_file.name} 时出错：{exc}")
        return ""


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

    # ── Read & Preprocess ─────────────────────────────────────────────────────
    file_data: Dict[str, Dict[str, str]] = {}
    for f in uploaded:
        raw  = read_file(f)
        proc = preprocess(raw, ignore_case, ignore_whitespace, ignore_comments)
        file_data[f.name] = {"raw": raw, "processed": proc}

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
    text1 = file_data[name1]["processed"]
    text2 = file_data[name2]["processed"]

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
