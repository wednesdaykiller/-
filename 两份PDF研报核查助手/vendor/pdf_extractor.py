# -*- coding: utf-8 -*-
"""
PDF 解析模块 —— 金融人工智能竞赛 · 研报核查方向

作者：sjj
职责：从财报 / 研报 PDF 中提取：
  1. 全文文本（绑定页码）              -> PyMuPDF (pymupdf)
  2. 财务表格（表头 + 行数据 + 页码 + 坐标）-> pdfplumber
  3. 数字标准化（"1,234.56万元" -> 数值 + 单位）
  4. 输出统一 JSON，供后续 RAG 检索 / 大模型核查模块直接使用

用法（命令行）：
    python pdf_extractor.py 测试财报.pdf
    python pdf_extractor.py 测试财报.pdf -o output.json
    python pdf_extractor.py 测试财报.pdf --preview        # 只打印预览不存文件
    python pdf_extractor.py 测试财报.pdf --docx           # 提取结果自动保存为桌面上的 Word 报告
    python pdf_extractor.py 测试财报.pdf --docx D:\a.docx  # 指定 Word 报告保存路径

依赖安装（PyCharm Terminal 或命令行）：
    pip install pymupdf pdfplumber pandas python-docx
"""

import math
import unicodedata
import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime
from typing import Any, Dict, List, Optional
from collections import Counter

import pymupdf as fitz          # 文本提取（PyMuPDF 新版推荐直接 import pymupdf）
import pdfplumber               # 表格提取

# Word 导出（python-docx，依赖：pip install python-docx）
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 日志：控制台 + 文件，满足竞赛「过程可追溯」
# ============================================================
LOG_FILE = "pdf_extractor.log"

logger = logging.getLogger("pdf_extractor")
logger.setLevel(logging.INFO)

_fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S")
_console = logging.StreamHandler(sys.stdout)
_console.setFormatter(_fmt)
logger.addHandler(_console)

_file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(_fmt)
logger.addHandler(_file_handler)


# ============================================================
# 1. 数字标准化
# ============================================================
# 完整匹配单值，防止从期间、代码、范围或说明文字中截出第一个数字。
_UNITS = sorted(['亿元', '万元', '百万元', '千元', '亿美元', '万港元', '亿港元',
                 '万美元', '人民币元', '港元', '美元', '元/股', '元/公斤',
                 '元/吨', '元', '亿', '万', '千', '%', '‰', '股', '万股', '亿股',
                 '倍', '吨', '万吨', '公斤', '人', '家', '个百分点', '基点'], key=len, reverse=True)
_UNIT_PATTERN = '|'.join(re.escape(u) for u in _UNITS)
_NUM_RE = re.compile(r'(?P<num>[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|(?=\.\d))(?:\.\d+)?)(?:\s*(?P<unit>' + _UNIT_PATTERN + r'))?')


def normalize_number(value: Any) -> Dict[str, Any]:
    """保留 raw，完整单值才生成 number；百分数不除以 100，不做单位换算。"""
    raw = '' if value is None else str(value).strip()
    result = {'raw': raw, 'number': None, 'unit': None}
    text = unicodedata.normalize('NFKC', raw).replace('−', '-').replace('﹣', '-')
    text = text.strip()
    negative = False
    # 会计括号负数：允许 (123万元) 和 (123)万元；拒绝带负号的双重符号。
    if text.startswith('('):
        wrapped = re.fullmatch(r'\(([^()]+)\)\s*(' + _UNIT_PATTERN + r')?', text)
        if not wrapped:
            return result
        text = wrapped[1].strip() + (wrapped[2] or '')
        if text.startswith(('+', '-')):
            return result
        negative = True
    match = _NUM_RE.fullmatch(text)
    if not match:
        return result
    digits = match['num'].lstrip('+-').split('.')[0]
    # 带前导零的多位整数常为代码。0、0.25 等仍可解析。
    if len(digits) > 1 and digits.startswith('0') and ',' not in digits:
        return result
    number = float(match['num'].replace(',', ''))
    if not math.isfinite(number):
        return result
    result.update(number=-number if negative else number, unit=match['unit'])
    return result


def _label_unit(label):
    """只读取明确括号单位或“单位：…”，不从说明文字猜单位。"""
    text = unicodedata.normalize('NFKC', str(label))
    candidates = re.findall(r'\(([^()]*)\)', text)
    marked = re.search(r'单位\s*[:：]\s*(.+)$', text)
    if marked:
        candidates.append(marked[1])
    units = set()
    for value in candidates:
        value = value.strip()
        if value in _UNITS:
            units.add(value)
        elif re.fullmatch(r'%\s*(?:YoY|YOY|yoy)', value):
            units.add('%')
    return next(iter(units)) if len(units) == 1 else None


def _protect_table_units(table):
    """仅补确定的单位；原文单元格单位优先，raw 与表格结构不变。"""
    headers = table['headers']
    if not headers:
        return
    table_unit = _label_unit(table['table_name']) or _label_unit(headers[0])
    unit_columns = [h for h in headers if h in ('单位', '计量单位')]
    section_unit = None
    for row in table['rows']:
        label = row[headers[0]]['raw']
        # 字段列、代码列、单位列不是测量值，保留原文而不生成金融数值。
        for h in headers:
            if h == headers[0] or re.search(r'代码|编号', h) or h in unit_columns:
                row[h]['number'], row[h]['unit'] = None, None
        numbers = [row[h] for h in headers[1:] if row[h]['number'] is not None]
        if label and not any(row[h]['raw'] for h in headers[1:]):
            section_unit = _label_unit(label)
        explicit = {unicodedata.normalize('NFKC', row[h]['raw']).strip() for h in unit_columns
                    if row[h]['raw']}
        row_unit = next(iter(explicit)) if len(explicit) == 1 and next(iter(explicit)) in _UNITS else None
        for h in headers[1:]:
            cell = row[h]
            if cell['number'] is None or cell['unit']:
                continue
            unit = row_unit or _label_unit(label) or _label_unit(h) or section_unit
            # 表名金额单位不可传播到利润率、每股指标、估值倍数等不同量纲。
            if not unit and not re.search(r'率|同比|环比|每股|EPS|P/E|P/B|PE\b|PB\b|RO[A-Z]|倍|比例', label, re.I):
                unit = table_unit
            cell['unit'] = unit


# ============================================================
# 2. 文本提取（PyMuPDF，带页码，按版面坐标重组）
# ============================================================
# 旧启发式仅保留给分栏判断；不再用于删除正文行。
# 研报常见页眉噪声（逐行过滤，不影响正文）
_HEADER_NOISE_PREFIXES = (
    "请务必阅读正文之后的免责声明",
    "证券研究报告",
)


def _is_noise_line(text: str) -> bool:
    """判断是否为页眉/页脚噪声行：固定免责声明、纯页码等。"""
    t = text.strip()
    if not t:
        return True
    if t.isdigit() and len(t) <= 4:          # 纯页码
        return True
    return any(t.startswith(p) for p in _HEADER_NOISE_PREFIXES)



_MARGIN_RATIO = 0.08
_MARGIN_REPEAT_RATIO = 0.60
_MARGIN_MIN_PAGES = 3


def _margin_signature(text: str, page_num: int) -> str:
    """只归一化空白及明确页码；不替换日期、年份、金额等普通数字。"""
    text = re.sub(r"\s+", "", text)
    # 独立页码必须与物理页码一致，避免把页边的财务数字当页码。
    if text == str(page_num):
        return "<page>"
    text = re.sub(r"第\d+页", "第<page>页", text)
    return re.sub(r"(?<!\d)[-–—]\d+[-–—](?!\d)", "-<page>-", text)


def _margin_candidates(words: List[tuple], bounds, page_num: int) -> List[tuple]:
    """返回 (区域, 签名, 相对x, 相对y, word索引列表)。

    仅收集完全处于上下 8% 的词。按 y 聚行，再按较大横向间隙拆片，
    避免左侧页眉与右侧独立正文绑在一起。坐标只用于检测，不重排正文。
    """
    candidates = []
    for side in ("top", "bottom"):
        selected = [
            (i, w) for i, w in enumerate(words)
            if (w[1] >= bounds.y0 and w[3] <= bounds.y0 + bounds.height * _MARGIN_RATIO
                if side == "top" else
                w[1] >= bounds.y1 - bounds.height * _MARGIN_RATIO and w[3] <= bounds.y1)
        ]
        rows = []
        for item in sorted(selected, key=lambda item: (item[1][1], item[1][0])):
            if rows and abs(item[1][1] - rows[-1][0][1][1]) <= 2:
                rows[-1].append(item)
            else:
                rows.append([item])
        for row in rows:
            chunks = []
            for item in sorted(row, key=lambda item: item[1][0]):
                if chunks and item[1][0] - chunks[-1][-1][1][2] <= 24:
                    chunks[-1].append(item)
                else:
                    chunks.append([item])
            for chunk in chunks:
                signature = _margin_signature(" ".join(w[4] for _, w in chunk), page_num)
                if signature:
                    candidates.append((
                        side, signature,
                        (chunk[0][1][0] - bounds.x0) / bounds.width,
                        (min(w[1] for _, w in chunk) - bounds.y0) / bounds.height,
                        [i for i, _ in chunk],
                    ))
    return candidates


def _detect_repeated_margins(doc) -> Dict[int, set]:
    """跨页检测重复页边片段，返回每页需要排除的 word 索引。

    至少 3 页且覆盖全文件 60% 的不同页面，位置相近才确认。
    短文档/不确定片段保留；统计与日志不进入 JSON。
    """
    import math

    groups = {}
    for page_num, page in enumerate(doc, start=1):
        for side, signature, x, y, indices in _margin_candidates(
            page.get_text("words"), page.rect, page_num
        ):
            clusters = groups.setdefault((side, signature), [])
            cluster = next((c for c in clusters
                            if abs(c[0] - x) <= 0.03 and abs(c[1] - y) <= 0.008), None)
            if cluster is None:
                cluster = [x, y, []]
                clusters.append(cluster)
            cluster[2].append((page_num, indices))
    threshold = max(_MARGIN_MIN_PAGES, math.ceil(doc.page_count * _MARGIN_REPEAT_RATIO))
    excluded = {}
    patterns = 0
    for clusters in groups.values():
        for _, _, occurrences in clusters:
            if len({page_num for page_num, _ in occurrences}) < threshold:
                continue
            patterns += 1
            for page_num, indices in occurrences:
                excluded.setdefault(page_num, set()).update(indices)
    logger.info("页眉/页脚检测：%d 个重复模式，覆盖 %d 页，排除 %d 个词",
                patterns, len(excluded), sum(map(len, excluded.values())))
    return excluded

def _has_body_column(rows: List[tuple]) -> bool:
    """该栏是否为真正的正文栏：存在 y 连续（≤16pt）且平均行宽 ≥200pt 的段落，
    且段落行数 ≥ 8。

    区分「正文栏」（连续长段落，需要独立成栏）与「并排图表/无边框表格」
    （y 对齐的窄行，应该按行拼接而不是分栏拆开）。
    """
    if not rows:
        return False
    sr = sorted(rows, key=lambda r: r[0])
    cur_len, cur_widths = 1, [sr[0][2] - sr[0][1]]
    best = 0

    def flush() -> None:
        nonlocal cur_len, cur_widths, best
        if cur_len >= 8 and sum(cur_widths) / len(cur_widths) >= 200:
            best = max(best, cur_len)
        cur_len, cur_widths = 1, []

    for i in range(1, len(sr)):
        if sr[i][0] - sr[i - 1][0] <= 18:
            cur_len += 1
            cur_widths.append(sr[i][2] - sr[i][1])
        else:
            flush()
            cur_len, cur_widths = 1, [sr[i][2] - sr[i][1]]
    flush()
    return best >= 8


def _split_columns(rows: List[tuple], page_width: float) -> List[List[tuple]]:
    """检测页面是否分栏：x0 存在明显空隙（>6% 页宽）且左栏是真正的正文栏
    （连续宽段落 ≥8 行、平均行宽 ≥200pt）时，按左右两栏输出；否则整体一栏。

    研报典型布局是「左栏正文 + 右栏信息」：左栏正文是稳定特征（连续长段落），
    右栏信息行数/宽度变化大，不能用右栏做判据。仅按 x 空隙分栏会误伤图表页
    （两图并排、无边框表格的行 y 对齐，拆开反而打断表格），
    所以必须同时满足「左栏是连续宽正文」才分栏。
    """
    if len(rows) < 6:
        return [rows]
    xs = sorted(r[1] for r in rows)
    max_gap, boundary = 0.0, None
    for i in range(len(xs) - 1):
        gap = xs[i + 1] - xs[i]
        if gap > max_gap:
            max_gap, boundary = gap, (xs[i] + xs[i + 1]) / 2
    if max_gap <= page_width * 0.06:
        return [rows]
    left = [r for r in rows if r[1] <= boundary]
    right = [r for r in rows if r[1] > boundary]
    if len(left) >= 3 and len(right) >= 3 and _has_body_column(left):
        return [left, right]
    return [rows]


def _is_digit_token(text: str) -> bool:
    """是否为纯数字/数值词（含 %、括号、小数点、货币符号、E/A 年份后缀）。"""
    t = text.replace("%", "").replace("(", "").replace(")", "")
    t = t.replace(",", "").replace(".", "").replace("-", "").replace(" ", "")
    t = t.replace("/", "").rstrip("EA")
    if not t:
        return True
    return t.isdigit()


def _words_to_lines(words: List[tuple]) -> List[tuple]:
    """词 → 视觉行列表 [(y0, x0, x1, text)]：先按 y 聚类（差 ≤2pt 为同行），
    行内按 x 从左到右拼接。必须在分栏之后对单栏调用，避免左右栏混拼。

    阈值取 2pt：研报两栏布局中左右栏同一水平线的行 y0 可差 2–3pt，
    取 3pt 会把跨栏的标题行误拼成一行；同行正文/表格词基线一致，差 <1pt。
    """
    if not words:
        return []
    ws = sorted(words, key=lambda w: (w[1], w[0]))
    line_rows: List[tuple] = []
    cur_y: Optional[float] = None
    cur: List[tuple] = []
    for w in ws:
        if cur_y is not None and abs(w[1] - cur_y) <= 2.0:
            cur.append(w)
        else:
            if cur:
                cur.sort(key=lambda a: a[0])
                text = " ".join(a[4] for a in cur).strip()
                line_rows.append((round(cur_y, 1), round(cur[0][0], 1),
                                  round(cur[-1][2], 1), text))
            cur_y, cur = w[1], [w]
    if cur:
        cur.sort(key=lambda a: a[0])
        text = " ".join(a[4] for a in cur).strip()
        line_rows.append((round(cur_y, 1), round(cur[0][0], 1),
                          round(cur[-1][2], 1), text))
    return line_rows


def _native_word_lines(words: List[tuple]) -> List[List[tuple]]:
    """使用 PDF 原生 block/line 标识，避免先按 y 聚类而混合相邻栏。"""
    groups = {}
    for w in words:
        if len(w) < 8:
            return []  # 无原生行证据时不猜测栏界
        groups.setdefault((w[5], w[6]), []).append(w)
    return [sorted(row, key=lambda w: w[0]) for row in groups.values()]


def _split_word_columns(words: List[tuple], page_width: float) -> List[List[tuple]]:
    """以连续正文及原生行之间的空白识别栏界，支持多个并排正文栏。

    候选正文至少 8 行；整行不可跨栏界。跨栏标题/页末说明作为独立区域，
    内部存在跨栏段落时保守回退单栏，留待后续阅读顺序模块处理。
    表格数字列、短标签及孤立图题不能独自构成正文栏证据。
    """
    rows = _native_word_lines(words)
    if not rows or page_width <= 0:
        return [words]

    def box(row):
        return (min(w[0] for w in row), min(w[1] for w in row),
                max(w[2] for w in row), max(w[3] for w in row))

    boxes = [box(row) for row in rows]
    prose = []
    for row, bounds in zip(rows, boxes):
        text = ''.join(w[4] for w in row)
        letters = sum(c.isalpha() for c in text)
        if (bounds[2] - bounds[0] >= max(80, page_width * 0.16)
                and len(text) >= 14 and letters >= len(text) * 0.45):
            prose.append(bounds)
    # 按行首聚簇，再寻找纵向连续正文；容许段首缩进。
    clusters = []
    for bounds in sorted(prose):
        cluster = next((c for c in clusters
                        if abs(c[0][0] - bounds[0]) <= page_width * 0.045), None)
        if cluster is None:
            clusters.append([bounds])
        else:
            cluster.append(bounds)
    anchors = []
    for cluster in clusters:
        runs = []
        for bounds in sorted(cluster, key=lambda b: b[1]):
            if (runs and bounds[1] - runs[-1][-1][1]
                    <= max(24, (bounds[3] - bounds[1]) * 2.8)):
                runs[-1].append(bounds)
            else:
                runs.append([bounds])
        for run in runs:
            if len(run) >= 8:
                anchors.append((min(b[0] for b in run), min(b[1] for b in run),
                                max(b[2] for b in run), max(b[3] for b in run)))
    cuts = []
    for anchor in anchors:
        nearby = [b for b in boxes if b[1] < anchor[3] and b[3] > anchor[1]]
        right = [b[0] for b in nearby if b[0] >= anchor[2] + 6]
        left = [b[2] for b in nearby if b[2] <= anchor[0] - 6]
        candidates = []
        if right:
            candidates.append((anchor[2] + min(right)) / 2)
        if left:
            candidates.append((max(left) + anchor[0]) / 2)
        for cut in candidates:
            if any(b[0] < cut < b[2] for b in nearby):
                continue
            sides = [[row for row, b in zip(rows, boxes)
                      if b[1] < anchor[3] and b[3] > anchor[1]
                      and (b[2] <= cut if side == 0 else b[0] >= cut)]
                     for side in (0, 1)]
            if any(len(side) < 5 or sum(any(c.isalpha() for c in ''.join(w[4] for w in row))
                                       for row in side) < 3 for side in sides):
                continue
            if not any(abs(cut - c) < 6 for c in cuts):
                cuts.append(cut)
    if not cuts:
        return [words]
    cuts.sort()
    # 同一个原生文本块可能把横跨页面的页首信息拆成多行（如左侧副标题、
    # 右侧报告类型）。在连续正文开始之前，保持该页首区域整体输出。
    blocks = {}
    for row in rows:
        for w in row:
            blocks.setdefault(w[5], []).append(w)
    anchor_top = min(a[1] for a in anchors)
    header_bottom = max((box(block)[3] for block in blocks.values()
                         if box(block)[3] <= anchor_top
                         and any(box(block)[0] < cut < box(block)[2] for cut in cuts)),
                        default=float('-inf'))
    header = [i for i, b in enumerate(boxes) if b[3] <= header_bottom]
    crossing = [i for i, b in enumerate(boxes) if any(b[0] < c < b[2] for c in cuts)]
    content = [i for i in range(len(rows)) if i not in crossing and i not in header]
    if not content:
        return [words]
    top = min(boxes[i][1] for i in content)
    bottom = max(boxes[i][3] for i in content)
    # 中途跨栏内容需要分区阅读顺序，本阶段不强行拆分。
    if any(boxes[i][3] > top and boxes[i][1] < bottom for i in crossing):
        return [words]
    prefix = [w for i in range(len(rows))
              if i in header or (i in crossing and boxes[i][3] <= top) for w in rows[i]]
    suffix = [w for i in crossing if boxes[i][1] >= bottom for w in rows[i]]
    columns = [[] for _ in range(len(cuts) + 1)]
    for i in content:
        column = sum(boxes[i][0] >= cut for cut in cuts)
        columns[column].extend(rows[i])
    return [region for region in [prefix, *columns, suffix] if region]


# P0-4：只改变区块顺序，不删除图表文字，不修改表格输出。
_FIGURE_CAPTION_RE = re.compile(r"^(?:图\s*\d+\s*[:：.]|Fig(?:ure)?\.?\s*\d+\s*[:：.])", re.I)
_SOURCE_LINE_RE = re.compile(r"^(?:资料来源|数据来源|来源|Sources?)\s*[:：]", re.I)


def _word_bounds(words: List[tuple]) -> tuple:
    return (min(w[0] for w in words), min(w[1] for w in words),
            max(w[2] for w in words), max(w[3] for w in words))


def _order_text_sections(words: List[tuple], page_width: float) -> List[List[tuple]]:
    """先保留已确认的整页分栏；否则尝试按水平空白分区，再在区内分栏。

    解决“上部双栏—中间通栏—下部双栏”等布局。若所有小区仍为单栏，
    保持原视觉行输出，不额外拆碎普通段落、表格或图表。
    """
    if not words:
        return []
    columns = _split_word_columns(words, page_width)
    if len(columns) > 1:
        return columns
    rows = _native_word_lines(words)
    if not rows:
        return [words]
    bounds = [_word_bounds(row) for row in rows]
    heights = sorted(b[3] - b[1] for b in bounds)
    min_gap = max(8.0, heights[len(heights) // 2] * 0.8)
    bands = []
    bottom = float('-inf')
    for row in sorted(rows, key=lambda r: (_word_bounds(r)[1], _word_bounds(r)[0])):
        box = _word_bounds(row)
        if not bands or box[1] - bottom >= min_gap:
            bands.append([])
        bands[-1].extend(row)
        bottom = max(bottom, box[3])
    regions = [_split_word_columns(band, page_width) for band in bands]
    if not any(len(region) > 1 for region in regions):
        return [words]
    return [column for region in regions for column in region]


def _paired_figure_bands(words: List[tuple], page_width: float) -> List[tuple]:
    """检测两个并排且各自有来源行的图文区，返回 (上界, 下界, 栏界)。

    标题、来源与贯穿区块的空白须同时支持；目录、无来源、跨栏正文、
    边界不明确均回退。编号仅用于识别通用图题格式，不依赖公司或页码。
    """
    rows = _native_word_lines(words)
    if not rows:
        return []
    boxes = [_word_bounds(row) for row in rows]
    texts = [' '.join(w[4] for w in row).strip() for row in rows]
    captions = [i for i, text in enumerate(texts) if _FIGURE_CAPTION_RE.match(text)
                and not re.search(r"\.{3,}|…|·{3,}", text)]
    sources = [i for i, text in enumerate(texts) if _SOURCE_LINE_RE.match(text)]
    candidates = []
    for pos, first in enumerate(captions):
        for second in captions[pos + 1:]:
            left, right = sorted((first, second), key=lambda i: boxes[i][0])
            lb, rb = boxes[left], boxes[right]
            if (abs(lb[1] - rb[1]) > 3 or rb[0] - lb[0] < page_width * 0.2
                    or rb[0] - lb[2] < 6):
                continue
            top = min(lb[1], rb[1])
            if sum(abs(boxes[i][1] - top) <= 3 for i in captions) != 2:
                continue  # 三图及更多并排布局不套用双图区块规则
            next_caption = min((boxes[i][1] for i in captions if boxes[i][1] > top + 3),
                               default=float('inf'))
            ends = []
            for caption in (left, right):
                cb = boxes[caption]
                matches = [i for i in sources
                           if boxes[i][1] >= max(lb[3], rb[3]) + 20
                           and boxes[i][1] < next_caption
                           and abs(boxes[i][0] - cb[0]) <= max(4, page_width * 0.015)]
                if not matches:
                    break
                ends.append(min(matches, key=lambda i: boxes[i][1]))
            if len(ends) != 2 or abs(boxes[ends[0]][1] - boxes[ends[1]][1]) > 12:
                continue
            bottom = max(boxes[i][3] for i in ends)
            # 允许来源行在栏宽末端折行。只接收邻近、同缩进、同字号高度的短尾行。
            for source in ends:
                sb = boxes[source]
                if len(texts[source]) < 24:
                    continue
                edge = rb[0] if source == ends[0] else max(b[2] for b in boxes)
                if edge - sb[2] > page_width * 0.06:
                    continue
                tails = [i for i, b in enumerate(boxes)
                         if 0 <= b[1] - sb[3] <= (sb[3] - sb[1]) * 0.9
                         and abs(b[0] - sb[0]) <= 3
                         and b[3] - b[1] <= (sb[3] - sb[1]) * 1.2
                         and b[2] - b[0] <= (sb[2] - sb[0]) * 0.5
                         and len(texts[i]) <= 20
                         and not _FIGURE_CAPTION_RE.match(texts[i])
                         and not _SOURCE_LINE_RE.match(texts[i])]
                if tails:
                    bottom = max(bottom, boxes[min(tails, key=lambda i: boxes[i][1])][3])
            # 区块上下边界不得切穿文字。
            if any((w[1] < top - 0.1 and w[3] > top + 0.1)
                   or (w[1] < bottom - 0.1 and w[3] > bottom + 0.1) for w in words):
                continue
            inside = [w for w in words if w[1] >= top - 0.1 and w[3] <= bottom + 0.1]
            # 在两图题之间寻找贯穿完整区块的空白，而非直接按中线切词。
            lo, hi = lb[2], rb[0]
            occupied = sorted((max(lo, w[0]), min(hi, w[2])) for w in inside
                              if w[2] > lo and w[0] < hi)
            gaps = []
            cursor = lo
            for start, end in occupied:
                if start > cursor:
                    gaps.append((cursor, start))
                cursor = max(cursor, end)
            if hi > cursor:
                gaps.append((cursor, hi))
            if not gaps:
                continue
            start, end = max(gaps, key=lambda gap: gap[1] - gap[0])
            if end - start < 4:
                continue
            cut = (start + end) / 2
            # 原生行也不能跨越栏界，防止将区内通栏正文或表格拆开。
            if any(b[1] >= top - 0.1 and b[3] <= bottom + 0.1
                   and b[0] < cut < b[2] for b in boxes):
                continue
            candidates.append((top, bottom, cut))
    bands = []
    for band in sorted(candidates):
        if not bands or band[0] > bands[-1][1]:
            bands.append(band)
    return bands


def _reading_order_regions(words: List[tuple], page_width: float) -> List[List[tuple]]:
    """区域从上到下，已确认的并排区块从左到右，区域内沿用视觉行排序。"""
    if not words:
        return []
    bands = _paired_figure_bands(words, page_width)
    if not bands:
        return _order_text_sections(words, page_width)
    remaining = list(words)
    regions = []
    for top, bottom, cut in bands:
        before = [w for w in remaining if w[3] < top]
        inside = [w for w in remaining if w[1] >= top - 0.1 and w[3] <= bottom + 0.1]
        # 按词实例做不相交划分，保留重复文字；浮点边界含糊时整页回退。
        before_ids = {id(w) for w in before}
        inside_ids = {id(w) for w in inside}
        if before_ids & inside_ids:
            return _order_text_sections(words, page_width)
        after = [w for w in remaining if id(w) not in before_ids | inside_ids]
        if any(w[1] < bottom - 0.1 for w in after):
            return _order_text_sections(words, page_width)
        regions.extend(_order_text_sections(before, page_width))
        regions.extend([side for side in ([w for w in inside if w[2] <= cut],
                                          [w for w in inside if w[0] >= cut]) if side])
        remaining = after
    regions.extend(_order_text_sections(remaining, page_width))
    return regions


# P0-5：仅修复有版面证据的中文正文软换行，不修改文字或跨页拼接。
_CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
_PARAGRAPH_LABEL_RE = re.compile(
    r"^(?:[图表]\s*\d+\s*[:：.]|(?:Fig(?:ure)?\.?|Table)\s*\d+\s*[:：.]"
    r"|(?:资料来源|数据来源|来源|Sources?)\s*[:：])", re.I)
_NEW_ITEM_RE = re.compile(r"^(?:[•●▪\uf06c]|[（(]?\d+(?:\.\d+)*[、）)]|[一二三四五六七八九十]+[、）])")


def _page_line_styles(page) -> Dict[tuple, tuple]:
    """原生行的主文字字体/字号；dict 仅提供样式，正文仍全部来自 words。"""
    # 与 words 使用相同 flags，排除图像块引起的 block 编号偏移。
    styles = {}
    for index, block in enumerate(page.get_text('dict', flags=fitz.TEXTFLAGS_WORDS)['blocks']):
        if block.get('type') != 0:
            continue
        for line_index, line in enumerate(block.get('lines', [])):
            spans = [s for s in line.get('spans', []) if s.get('text', '').strip()]
            if not spans:
                continue
            main = max(spans, key=lambda s: (len(_CJK_RE.findall(s['text'])), len(s['text'])))
            styles[(block.get('number', index), line_index)] = (
                main['font'], round(main['size'], 2), bool(main.get('flags', 0) & 16))
    return styles


def _join_soft_break(left: str, right: str) -> Optional[str]:
    """返回续行拼接符；None 表示数字/词边界有歧义，保留换行。"""
    a, b = left[-1], right[0]
    # 不把两行数值、负号、小数/千分位碎片猜成一个新数字。
    if a.isascii() and b.isascii() and a.isdigit() and b.isalnum():
        return None
    if (a.isascii() and b.isascii()
            and (a.isalnum() or a in '.,+-/%(')
            and (b.isdigit() or b in '.,+-/%')):
        return None
    if a.isascii() and b.isascii() and a.isalnum() and b.isalnum():
        return ' '
    return ''


def _merge_paragraph_lines(words: List[tuple], page_width: float,
                           styles: Dict[tuple, tuple]) -> List[str]:
    """在一个已排序区块内，保守地消除中文正文软换行。

    多个原生行拼成的视觉行视为表格/复杂区域，不参与合并。字体、行距、
    缩进、段落右边界、标点及项目起点须同时支持；缺少样式则原样返回。
    """
    visual = _words_to_lines(words)
    if not visual:
        return []
    groups = []
    for w in sorted(words, key=lambda w: (w[1], w[0])):
        if groups and abs(w[1] - groups[-1][0][1]) <= 2.0:
            groups[-1].append(w)
        else:
            groups.append([w])
    records = []
    block_right = {}
    for row, group in zip(visual, groups):
        ids = {(w[5], w[6]) for w in group if len(w) >= 8}
        native = next(iter(ids)) if len(ids) == 1 and all(len(w) >= 8 for w in group) else None
        bounds = _word_bounds(group)
        record = {'text': row[3], 'box': bounds, 'native': native,
                  'style': styles.get(native)}
        records.append(record)
        if native is not None:
            block_right[native[0]] = max(block_right.get(native[0], bounds[2]), bounds[2])

    def can_join(prev, current):
        a, b = prev['text'], current['text']
        pa, pb = prev['box'], current['box']
        sa, sb = prev['style'], current['style']
        if not sa or not sb or not prev['native'] or not current['native']:
            return False
        numeric_tail = bool(re.search(r'(?:为|至|达|约)$', a)
                            and re.fullmatch(r'[\d\s.,，/％%+\-（）()]+[。.]?', b)
                            and prev['native'][0] == current['native'][0])
        if ((sa[0] != sb[0] and not numeric_tail)
                or abs(sa[1] - sb[1]) > 0.5 or sa[2] != sb[2]):
            return False
        em = max(sa[1], sb[1])
        if not (em * 0.9 <= pb[1] - pa[1] <= em * 1.8
                and 0 <= pb[1] - pa[3] <= em * 0.8):
            return False
        if not (-em * 2.5 <= pb[0] - pa[0] <= em * 0.6):
            return False  # 下一行新增缩进通常代表新段落
        if _PARAGRAPH_LABEL_RE.match(a) or _PARAGRAPH_LABEL_RE.match(b) or _NEW_ITEM_RE.match(b):
            return False
        if re.match(r'^[•●▪\uf06c]', a) and not re.search(r'[:：]', a):
            return False
        if re.search(r'[。！？!?；;：:]\s*[”’」』）)]?$', a):
            return False
        if re.search(r'\.{3,}|…|·{3,}', a + b):
            return False
        count = len(_CJK_RE.findall(a))
        if count < 10 or count < len(a.replace(' ', '')) * 0.2:
            return False
        if pa[2] - pa[0] < max(120, page_width * 0.28):
            return False
        target_right = max(pa[2], pb[2], block_right[prev['native'][0]])
        if target_right - pa[2] > em * 1.5:
            return False  # 短行通常是段尾，不强接下一段
        cjk_b = len(_CJK_RE.findall(b))
        if cjk_b < 4 or (cjk_b < 10 and len(b.strip()) < 16):
            short_tail = bool(_CJK_RE.search(b) and re.search(r'[。！？!?；;）)]$', b))
            if not numeric_tail and not short_tail:
                return False
        # 跨原生 block 需要更严格的起点/右端连续性，禁止把新缩进段落接上。
        if prev['native'][0] != current['native'][0]:
            if abs(pb[0] - pa[0]) > em * 0.3 or abs(pb[2] - pa[2]) > em * 1.5:
                return False
        return _join_soft_break(a, b) is not None

    result = []
    for index, record in enumerate(records):
        if index and can_join(records[index - 1], record):
            separator = _join_soft_break(records[index - 1]['text'], record['text'])
            result[-1] += separator + record['text']
        else:
            result.append(record['text'])
    return result


# P0-6：图题/来源 + 图形证据 + 数字标签共同确认；不按关键词全局删数字。
_CHART_NUMBER_RE = re.compile(r'^[+\-−]?(?:\d[\d,]*)(?:\.\d+)?(?:[%％]|年|月|H[12])?$')
_CHART_NOTE_RE = re.compile(r'^(?:注|备注|说明|Notes?)\s*[:：]', re.I)


def _chart_candidates(words: List[tuple], page_rect) -> List[dict]:
    """带来源的编号图候选区域；并排图沿用 P0-4 栏界，单图保守匹配来源。"""
    rows = _native_word_lines(words)
    boxes = [_word_bounds(row) for row in rows]
    texts = [' '.join(w[4] for w in row).strip() for row in rows]
    captions = [i for i, t in enumerate(texts) if _FIGURE_CAPTION_RE.match(t)
                and not re.search(r'\.{3,}|…|·{3,}', t)]
    sources = [i for i, t in enumerate(texts) if _SOURCE_LINE_RE.match(t)]
    pairs = _paired_figure_bands(words, page_rect.width)
    candidates = []
    for caption in captions:
        cb = boxes[caption]
        aligned = [i for i in captions if abs(boxes[i][1] - cb[1]) <= 3]
        pair = next((b for b in pairs if abs(b[0] - cb[1]) <= 3), None)
        if len(aligned) != 1 and pair is None:
            continue
        if pair:
            left = cb[0] < pair[2]
            x0 = cb[0] - 3 if left else pair[2]
            x1 = pair[2] if left else page_rect.x1
        else:
            x0, x1 = cb[0] - 3, page_rect.x1
        next_caption = min((boxes[i][1] for i in captions if boxes[i][1] > cb[1] + 3),
                           default=page_rect.y1)
        matches = [i for i in sources
                   if cb[3] + 30 <= boxes[i][1] < next_caption
                   and abs(boxes[i][0] - cb[0]) <= max(4, page_rect.width * 0.015)
                   and boxes[i][2] <= x1 + 1]
        if not matches:
            continue
        source = min(matches, key=lambda i: boxes[i][1])
        # 候选上下界避开图题、来源，且不跨到下一幅图。
        bbox = (x0, cb[3] + 0.5, x1, boxes[source][1] - 0.5)
        protected_blocks = {rows[caption][0][5], rows[source][0][5]}
        for i, t in enumerate(texts):
            if (_CHART_NOTE_RE.match(t) and boxes[i][1] >= bbox[1]
                    and boxes[i][3] <= bbox[3] and x0 <= boxes[i][0] < x1):
                protected_blocks.add(rows[i][0][5])
        candidates.append({'caption': texts[caption], 'bbox': bbox,
                           'protected_blocks': protected_blocks})
    return candidates


def _chart_graphics_evidence(bbox: tuple, drawings: list, images: list) -> bool:
    """必须有柱/曲线/图像元素，单独矩形背景或表格边框不算图形证据。"""
    x0, y0, x1, y1 = bbox

    def inside(rect):
        return (rect[0] >= x0 - 1 and rect[1] >= y0 - 1
                and rect[2] <= x1 + 1 and rect[3] <= y1 + 1)

    substantial_images = [img['bbox'] for img in images if inside(img['bbox'])
                          and img['bbox'][2] - img['bbox'][0] >= 4
                          and img['bbox'][3] - img['bbox'][1] >= 4]
    if len(substantial_images) >= 3:
        return True
    if any((r[2] - r[0]) >= (x1 - x0) * 0.4
           and (r[3] - r[1]) >= (y1 - y0) * 0.4 for r in substantial_images):
        return True
    bars = 0
    diagonal = 0
    for path in drawings:
        r = path['rect']
        if not inside(r):
            continue
        fill = path.get('fill')
        if (fill is not None and any(c < 0.9 for c in fill)
                and 2 <= r.width <= (x1 - x0) * 0.18
                and r.height >= 5 and path.get('fill_opacity', 1) > 0):
            bars += 1
        if path.get('color') is not None and path.get('stroke_opacity', 1) > 0:
            for item in path['items']:
                if item[0] == 'c':
                    diagonal += 1
                elif item[0] == 'l' and abs(item[1].x - item[2].x) > 1 and abs(item[1].y - item[2].y) > 1:
                    diagonal += 1
    return bars >= 3 or diagonal >= 2


def _detect_chart_text(page, words: List[tuple]) -> List[dict]:
    """返回已确认区域及待排除 words，保留所有候选以外的文字。

    无绘图/图像证据、数字不足、含长正文或表格标题时保留。图内纯图片本身
    无可提取文字，本阶段不 OCR。返回值仅用于过滤/日志，不进入原 JSON。
    """
    candidates = _chart_candidates(words, page.rect)
    if not candidates:
        return []
    try:
        drawings, images = page.get_drawings(), page.get_image_info()
    except Exception as exc:
        logger.warning('图形证据读取失败，保留本页图表文字：%s', exc)
        return []
    accepted = []
    for candidate in candidates:
        x0, y0, x1, y1 = candidate['bbox']
        inside = [w for w in words if w[0] >= x0 and w[2] <= x1
                  and w[1] >= y0 and w[3] <= y1
                  and w[5] not in candidate['protected_blocks']]
        if not inside:
            continue
        # 拒绝切穿边界文字，也不截取横跨两区的原生行。
        rows = _native_word_lines(words)
        if any(b[0] < x1 and b[2] > x0 and b[1] < y1 and b[3] > y0
               and not (b[0] >= x0 and b[2] <= x1 and b[1] >= y0 and b[3] <= y1)
               for b in map(_word_bounds, rows)):
            continue
        numeric_labels = [w[4].strip() for w in inside
                          if _CHART_NUMBER_RE.fullmatch(w[4].strip())]
        numbers = len(numeric_labels)
        sparse_percent_chart = numbers >= 2 and all(t.endswith(('%', '％')) for t in numeric_labels)
        if (numbers < 6 and not sparse_percent_chart) or numbers < len(inside) * 0.35:
            continue
        unsafe = False
        for row in _native_word_lines(inside):
            text = ''.join(w[4] for w in row)
            if (re.match(r'^(?:表\s*\d+|Table\s*\d+)', text, re.I)
                    or (len(_CJK_RE.findall(text)) >= 14 and re.search(r'[，。；]', text))
                    or len(re.findall(r'[A-Za-z]{2,}', text)) >= 12):
                unsafe = True
                break
        if unsafe or not _chart_graphics_evidence(candidate['bbox'], drawings, images):
            continue
        accepted.append({'caption': candidate['caption'], 'bbox': candidate['bbox'], 'words': inside})
    return accepted


def _extract_page_ordered(
    page, skip_rects: Optional[List[tuple]] = None,
    excluded_word_indices: Optional[set] = None,
    filter_chart_text: bool = True,
) -> str:
    """按版面坐标重组文本：区块阅读顺序 → 区内分栏/拼行 → 正文软换行修复。

    研报 PDF 常为「左栏正文 + 右栏信息」的两栏布局，或带浮动图表。
    - get_text("text") 按内部内容流输出，会把同一行的数据拆散成列；
    - dict 模式在某些 PDF（如含多区块无边框大表）下会漏行；
    - 若先按 y 聚类再分栏，左右栏同一行的词会混拼。
    因此用 words 模式（原子级文本对象）：先排除表格区域 → 词级分栏 →
    栏内按 y 聚类成行、行内按 x 拼接，栏与栏之间空一行。
    """
    rects = list(skip_rects or [])
    words = [
        w for w in page.get_text("words")
        if not (
            rects and any(
                r[0] - 2 <= (w[0] + w[2]) / 2 <= r[2] + 2
                and r[1] - 2 <= (w[1] + w[3]) / 2 <= r[3] + 2
                for r in rects
            )
        )
    ]
    if not words:
        return ""
    excluded_words = {tuple(w) for i, w in enumerate(page.get_text("words"))
                      if i in (excluded_word_indices or set())}
    parts: List[str] = []
    body_words = [w for w in words if tuple(w) not in excluded_words]
    styles = _page_line_styles(page)
    # 先按原完整内容确定阅读顺序，再隔离图区，避免删词后改变分栏/区块判断。
    regions = _reading_order_regions(body_words, page.rect.width)
    charts = _detect_chart_text(page, body_words) if filter_chart_text else []
    chart_words = {tuple(w) for chart in charts for w in chart['words']}
    if charts:
        logger.info('第 %d 页图表正文过滤：%d 个图区，%d 个词',
                    page.number + 1, len(charts), sum(tuple(w) in chart_words for w in body_words))
    for col_words in regions:
        retained = [w for w in col_words if tuple(w) not in chart_words]
        lines = _merge_paragraph_lines(retained, page.rect.width, styles)
        parts.append("\n".join(lines))
    return "\n\n".join(parts).strip()


def extract_text_with_page(
    pdf_path: str,
    skip_bboxes: Optional[Dict[int, List[tuple]]] = None,
    filter_chart_text: bool = True,
) -> List[Dict[str, Any]]:
    """逐页提取文本（重复页边检测、版面排序、续行拼接及可选图表文字过滤）。

    filter_chart_text: 默认过滤确认的图区文字；False 可恢复 P0-5 正文。
    skip_bboxes: {页码: [表格bbox,...]}，与表格提取同源，避免表格内容
    在正文中重复出现，也避免漏掉表格未覆盖的区域（如无边框表区块）。
    """
    pages: List[Dict[str, Any]] = []
    doc = fitz.open(pdf_path)
    try:
        total = doc.page_count
        excluded = _detect_repeated_margins(doc)
        for i, page in enumerate(doc, start=1):
            rects = (skip_bboxes or {}).get(i, [])
            raw_lines = _extract_page_ordered(page, rects, excluded.get(i), filter_chart_text).split("\n")
            clean_lines = [ln for ln in raw_lines if ln.strip()]
            text = "\n".join(clean_lines).strip()
            pages.append({
                "page_num": i,
                "text": text,
                "char_count": len(text),
            })
            if i % 10 == 0 or i == total:
                logger.info("文本提取进度：%d/%d 页", i, total)
    finally:
        doc.close()
    return pages


# ============================================================
# 3. 表格提取（pdfplumber，带页码 / 坐标 / 表名推断）
# ============================================================
def _infer_table_name(page, table_bbox: List[float], max_dist: float = 60.0) -> str:
    """
    简单推断表格标题：取表格 bbox 上方最近的一行文本作为候选标题。
    只用于可读性提示，不影响数据正确性。
    """
    try:
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        tbl_top = table_bbox[1]
        lines: Dict[float, List[tuple]] = {}
        for w in words:
            # 只考虑表格上方的词，且距离不超过 max_dist
            if w["top"] < tbl_top and tbl_top - w["bottom"] < max_dist:
                lines.setdefault(round(w["top"], 1), []).append((w["x0"], w["text"]))
        if not lines:
            return ""
        # 取最靠近表格的那一行
        nearest_top = min(lines.keys(), key=lambda t: tbl_top - t)
        words_in_line = sorted(lines[nearest_top], key=lambda x: x[0])
        title = "".join(w for _, w in words_in_line).strip()
        if title and len(title) <= 40:
            return title
    except Exception as e:  # noqa: BLE001 —— 表名推断失败不影响主流程
        logger.debug("表名推断失败：%s", e)
    return ""


def _clean_table_geometry(page):
    """去掉同色单元格底色内的嵌套底色；保留外框、细线和描边对象。"""
    rects = [r for r in page.rects if r['width'] > 2 and r['height'] > 2
             and r.get('fill') and not r.get('stroke')]
    removed = set()
    for r in rects:
        if any(s is not r and s.get('non_stroking_color') == r.get('non_stroking_color')
               and s['width'] * s['height'] > r['width'] * r['height'] * 1.05
               and s['x0'] <= r['x0'] + .3 and s['x1'] >= r['x1'] - .3
               and s['top'] <= r['top'] + .3 and s['bottom'] >= r['bottom'] - .3
               for s in rects):
            removed.add(id(r))
    return page.filter(lambda obj: id(obj) not in removed)


def _table_text(value):
    """只规整表头；数据 raw 保留原有换行。"""
    return re.sub(r'\s+', '', str(value or ''))


def _year_header(value):
    return bool(re.fullmatch(r'(?:19|20)\d{2}(?:[AE]|H[12]|Q[1-4])?', _table_text(value)))


def _recover_horizontal_table(page, table, stop_y=None):
    """仅补全带年份表头、连续横线的缺竖线表，未确认的无框区域不猜测。"""
    stop_y = stop_y if stop_y is not None else page.height * .94
    raw = table.extract()
    if len(raw) != 1 or sum(_year_header(v) for v in raw[0]) < 2:
        return table
    cells = table.rows[0].cells
    if any(c is None for c in cells):
        return table
    xs = [c[0] for c in cells] + [cells[-1][2]]
    # 同一水平线可以是多个单元格的短线，按 y 聚合后验证覆盖率。
    groups = []
    for e in sorted(page.edges, key=lambda e: e['top']):
        if e['orientation'] != 'h' or e['top'] < table.bbox[3] - 2:
            continue
        a, b = max(xs[0], e['x0']), min(xs[-1], e['x1'])
        if b <= a:
            continue
        if not groups or e['top'] - groups[-1][0] > 1:
            groups.append([e['top'], []])
        groups[-1][1].append((a, b))
    ys = [table.bbox[1], table.bbox[3]]
    max_gap = max(16, min(28, (table.bbox[3] - table.bbox[1]) * 1.3))
    for y, spans in groups:
        if y >= stop_y - 1:
            break
        if y <= ys[-1] + 2:
            continue
        if y - ys[-1] > max_gap:
            break
        end, covered = xs[0], 0
        for a, b in sorted(spans):
            covered += max(0, b - max(end, a))
            end = max(end, b)
        if covered < (xs[-1] - xs[0]) * .9:
            break
        ys.append(y)
    if len(ys) < 4:
        return table
    # 末行可能没有下边框。只在相同列内确认字段和至少两个数字后补末边。
    lower = min(stop_y - 1, ys[-1] + max_gap)
    if lower > ys[-1] + 3:
        tail = page.crop((xs[0], ys[-1] + 1, xs[-1], lower)).extract_words()
        if tail:
            top = min(w['top'] for w in tail)
            line = [w for w in tail if abs(w['top'] - top) < 3]
            values = [''.join(w['text'] for w in line if a <= (w['x0'] + w['x1']) / 2 < b)
                      for a, b in zip(xs, xs[1:])]
            if values[0] and sum(bool(re.fullmatch(r'[-+]?\d[\d,.]*[%％]?', v))
                                 for v in values[1:]) >= 2:
                ys.append(min(lower, max(w['bottom'] for w in line) + 2))
    found = page.crop((xs[0], ys[0], xs[-1], ys[-1])).find_tables({
        'vertical_strategy': 'explicit', 'explicit_vertical_lines': xs,
        'horizontal_strategy': 'explicit', 'explicit_horizontal_lines': ys,
    })
    if not found:
        return table
    candidate = found[0]
    rows = candidate.extract()[1:]
    # 至少两行同时有字段及两项数字，避免把附近正文扩进表格。
    numeric = lambda s: bool(re.fullmatch(r'[-+（(]?\d[\d,.]*[%％）)]?', _table_text(s)))
    if sum(bool(r[0]) and sum(numeric(v) for v in r[1:]) >= 2 for r in rows) < 2:
        return table
    return candidate


def _unique_table_headers(labels):
    result, used = [], set()
    for i, label in enumerate(labels):
        base = label or f'col{i + 1}'
        name, suffix = base, 1
        while name in used:
            name = f'{base}_{suffix}'
            suffix += 1
        used.add(name)
        result.append(name)
    return result


def _table_matrix(table):
    """仅对有实际合并几何的单元格传递文字；普通空格绝不填数。"""
    raw = table.extract()
    xs = sorted({c[0] for c in table.cells})
    ys = sorted({c[1] for c in table.cells})
    matrix = [list(row) for row in raw]
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            if cell is None or not raw[i][j]:
                continue
            for ii, y in enumerate(ys):
                for jj, x in enumerate(xs):
                    if cell[0] - .1 <= x < cell[2] - .5 and cell[1] - .1 <= y < cell[3] - .5:
                        if matrix[ii][jj] is None:
                            matrix[ii][jj] = raw[i][j]
    return raw, matrix, xs


def _table_header_depth(table, raw, max_rows):
    """第二层需由跨列表头和下一行短标签共同支持，不把前三行一律当表头。"""
    depth = 1
    for i in range(min(max(1, max_rows) - 1, len(raw) - 2)):
        xs = sorted({c[0] for c in table.cells})
        spans = [c for c in table.rows[i].cells if c and
                 sum(c[0] - .1 <= x < c[2] - .5 for x in xs) > 1]
        next_labels = [v for v in raw[i + 1] if v]
        if not spans or not next_labels or not all(
                _year_header(v) or (len(_table_text(v)) <= 12 and
                                   not re.search(r'\d', _table_text(v))) for v in next_labels):
            break
        depth += 1
    return depth


def extract_tables_with_pdfplumber(
    pdf_path: str,
    header_rows: int = 3,
    infer_table_names: bool = True,
) -> List[Dict[str, Any]]:
    """P0-7/P1：修复边框、隔离图形、保留列关系，并展开有依据的多层表头。

    header_rows 是自动识别表头层数的上限；JSON 键及 cell 结构不变。
    相邻页仅复用严格匹配的表头，不合并跨页行或跨页表格。
    """
    tables = []
    previous = []
    with pdfplumber.open(pdf_path) as pdf, fitz.open(pdf_path) as doc:
        for page_idx, source in enumerate(pdf.pages, 1):
            current = []
            try:
                page = _clean_table_geometry(source)
                found = page.find_tables()
                fp = doc[page_idx - 1]
                charts = _detect_chart_text(fp, fp.get_text('words'))
            except Exception as exc:
                logger.warning('第 %d 页表格检测异常：%s', page_idx, exc)
                previous = []
                continue
            for table in found:
                try:
                    stop_y = min((t.bbox[1] for t in found if t.bbox[1] > table.bbox[3] + 2
                                  and min(t.bbox[2], table.bbox[2]) > max(t.bbox[0], table.bbox[0])),
                                 default=page.height * .94)
                    table = _recover_horizontal_table(page, table, stop_y)
                    raw, matrix, xs = _table_matrix(table)
                    # 原始完全空列只删除，不再把两侧内容拆成独立小表。
                    active = [j for j in range(len(xs)) if any(r[j] for r in raw)]
                    inherited = None
                    if table.bbox[1] < page.height * .15 and not re.match(
                            r'^表\s*\d+', _infer_table_name(source, table.bbox)):
                        for prev in previous:
                            indexes = [min(range(len(xs)), key=lambda j: abs(xs[j] - x))
                                       for x in prev['xs']]
                            if (len(set(indexes)) == len(indexes)
                                    and all(abs(xs[j] - x) < 2 for j, x in zip(indexes, prev['xs']))
                                    and set(active).issubset(indexes)
                                    and abs(table.bbox[2] - prev['right']) < 2):
                                # 有本页年份表头或重复原表头时，由本页自行解析。
                                if not any(_year_header(v) for v in raw[0]) and not all(
                                        _table_text(raw[0][j]) == h for j, h in zip(indexes, prev['headers'])):
                                    inherited, active = prev, indexes
                                    break
                    if len(active) < 2 or len(raw) < 2:
                        logger.debug('第 %d 页过滤单列/单行对象 %s', page_idx, table.bbox)
                        continue
                    # 图内数值网格不是表格；只过滤完全落在已确认图区内的对象。
                    x0, y0, x1, y1 = table.bbox
                    if any(c['bbox'][0] <= x0 and c['bbox'][1] <= y0
                           and c['bbox'][2] >= x1 and c['bbox'][3] >= y1 for c in charts):
                        logger.info('第 %d 页过滤图内伪表格 %s', page_idx, table.bbox)
                        continue
                    if inherited:
                        headers, depth = inherited['headers'], 0
                    else:
                        depth = _table_header_depth(table, raw, header_rows)
                        labels = []
                        for j in active:
                            parts = list(dict.fromkeys(_table_text(matrix[i][j]) for i in range(depth)
                                                       if matrix[i][j]))
                            labels.append(' / '.join(parts))
                        headers = _unique_table_headers(labels)
                    rows = [{h: normalize_number(str(matrix[i][j] or '').strip())
                             for h, j in zip(headers, active)}
                            for i in range(depth, len(raw)) if any(raw[i][j] for j in active)]
                    if not rows:
                        continue
                    name = _infer_table_name(source, table.bbox) if infer_table_names else ''
                    if infer_table_names and inherited:
                        name = inherited['name']
                    # 表内财务标题自带单位，优先于图注/附近正文。
                    if infer_table_names and re.search(r'[（(].*(?:元|%|％|倍).*[）)]', headers[0]):
                        name = headers[0]
                    output = dict(page_num=page_idx, table_id=f'page{page_idx}_table{len(current)}',
                                  bbox=table.bbox, table_name=name, headers=headers,
                                  rows=rows, row_count=len(rows))
                    _protect_table_units(output)
                    tables.append(output)
                    current.append({'headers': headers, 'xs': [xs[j] for j in active],
                                    'right': x1, 'bottom': y1, 'name': name})
                except Exception as exc:
                    logger.warning('第 %d 页候选表格处理异常：%s', page_idx, exc)
            previous = [t for t in current if t['bottom'] > page.height * .85]
            logger.info('第 %d 页：%d 个边框候选 → %d 个结构表格', page_idx, len(found), len(current))
    return tables


# ============================================================
# 4. 主流程：组装统一 JSON
# ============================================================
def _ocr_financial_tables(records, page_num, page_width):
    """基于连续年份锚点的财务表初版。原 OCR 文字和不确定性留存，不修猜错字。"""
    rows = _ocr_rows(records)
    starts = []
    for row in rows:
        group = []
        for item in row['items']:
            value = re.sub(r'\s+', '', item['text'])
            if _year_header(value):
                if group and int(value[:4]) != int(re.sub(r'\s+', '', group[-1]['text'])[:4]) + 1:
                    if len(group) >= 3:
                        starts.append((row, group))
                    group = []
                group.append(item)
            else:
                if len(group) >= 3:
                    starts.append((row, group))
                group = []
        if len(group) >= 3:
            starts.append((row, group))
    output, audit = [], []
    for header_row, years in starts:
        anchors = [(w['bbox'][0] + w['bbox'][2]) / 2 for w in years]
        spacing = sorted(b-a for a,b in zip(anchors,anchors[1:]))[len(anchors)//2-1]
        left_labels = [w for w in header_row['items'] if w['bbox'][2] < years[0]['bbox'][0]
                       and not _year_header(w['text'])]
        label = max(left_labels,key=lambda w:w['bbox'][2]) if left_labels else None
        if not label:
            nearby = [w for w in records if w['bbox'][2] < years[0]['bbox'][0]
                      and abs((w['bbox'][1]+w['bbox'][3])/2-header_row['cy']) < 11
                      and re.search(r'会计年度|指标|项目|科目', w['text'])]
            label = min(nearby,key=lambda w:abs((w['bbox'][1]+w['bbox'][3])/2-header_row['cy'])) if nearby else None
        if not label:
            continue
        x0 = label['bbox'][0] - 3
        x1 = min(page_width, anchors[-1] + spacing * .5)
        boundary = years[0]['bbox'][0] - 2
        y0 = min(w['bbox'][1] for w in [label]+years)
        stop = min((row['cy'] - 3 for row, other in starts if row['cy'] > header_row['cy'] + 5
                    and x0 < other[0]['bbox'][0] < x1), default=float('inf'))
        column_records = [w for w in records if w is not label and x0 <= (w['bbox'][0]+w['bbox'][2])/2 <= x1
                          and w['bbox'][1] > header_row['cy'] and w['bbox'][1] < stop]
        number_rows = []
        value_records = [w for w in column_records if (w['bbox'][0]+w['bbox'][2])/2 >= boundary]
        for row in _ocr_rows(value_records):
            slots = [[] for _ in anchors]
            for w in row['items']:
                cx = (w['bbox'][0]+w['bbox'][2])/2
                if cx < boundary:
                    continue
                j = min(range(len(anchors)),key=lambda j:abs(anchors[j]-cx))
                if abs(anchors[j]-cx) < spacing*.45 and re.search(r'\d|^[OoIl-]+$',w['text']):
                    slots[j].append(w)
            if sum(bool(s) for s in slots) >= max(2,len(anchors)-2):
                if number_rows and row['cy'] - number_rows[-1]['cy'] > 32:
                    break
                number_rows.append({'cy':row['cy'],'slots':slots})
            elif number_rows and any(re.match(r'资料来源|数据来源|注[:：]|风险',w['text']) for w in row['items']):
                break
        if len(number_rows) < 2:
            continue
        headers = _unique_table_headers([label['text']] + [re.sub(r'\s+','',w['text']) for w in years])
        label_groups = [[] for _ in number_rows]
        for w in column_records:
            if w['bbox'][2] > boundary:
                continue
            cy = (w['bbox'][1]+w['bbox'][3])/2
            if cy > number_rows[-1]['cy'] + 10:
                continue
            j = min(range(len(number_rows)), key=lambda j:abs(number_rows[j]['cy']-cy))
            # 单独换行的括号单位属于上一数据行，不移到下一行同比。
            if re.fullmatch(r'[（(][^()（）]{1,10}[）)]',w['text'].strip()) and _label_unit(w['text']):
                preceding = [i for i,r in enumerate(number_rows) if r['cy'] < cy]
                if preceding:
                    j = preceding[-1]
            label_groups[j].append(w)
        built_rows, cell_audit = [], []
        for i,nrow in enumerate(number_rows):
            labels = label_groups[i]
            label_text = '\n'.join(w['text'] for w in sorted(labels,key=lambda w:w['bbox'][1]))
            row_dict = {headers[0]: {'raw':label_text,'number':None,'unit':None}}
            for j,slot in enumerate(nrow['slots']):
                raw = ' '.join(w['text'] for w in slot)
                cell = normalize_number(raw)
                confidence = min([w['confidence'] for w in slot+labels+[years[j]]],default=0)
                flags = _ocr_risk_flags(raw,confidence)
                if not slot:
                    flags.append('unrecognized_or_blank_cell')
                if raw and cell['number'] is None:
                    flags.append('not_a_strict_single_number')
                if len(slot)>1:
                    flags.append('multiple_boxes_in_cell')
                if not labels:
                    flags.append('missing_row_label')
                if flags:
                    cell['number'],cell['unit']=None,None
                row_dict[headers[j+1]] = cell
                cell_audit.append({'row_index':i,'header':headers[j+1], 'raw':raw,
                                   'confidence':confidence,'flags':flags,
                                   'source_bboxes':[w['bbox'] for w in slot]})
            built_rows.append(row_dict)
        above = [w for w in records if x0 <= w['bbox'][0] < x1 and y0-22 < w['bbox'][1]
                 and w['bbox'][3] <= y0+1]
        name = ' '.join(w['text'] for w in sorted(above,key=lambda w:w['bbox'][0])
                        if len(w['text']) < 30 and not re.match(r'风险|资料来源|数据来源',w['text']))
        table_id = f'page{page_num}_ocrtable{len(output)}'
        y1 = max(w['bbox'][3] for nr in number_rows for slot in nr['slots'] for w in slot)
        table = dict(page_num=page_num,table_id=table_id,bbox=[x0,y0,x1,y1],table_name=name,
                     headers=headers,rows=built_rows,row_count=len(built_rows))
        _protect_table_units(table)
        output.append(table)
        audit.append({'table_id':table_id,'source':'ocr','review_required':True,
                      'method':'year_anchor_alignment','cells':cell_audit})
    return output,audit


def _ocr_risk_flags(text, confidence, threshold=.90):
    flags = []
    if confidence < threshold:
        flags.append('low_confidence')
    # 不自动把字母改成数字；金融混合串和比较范围留待复核。
    if re.fullmatch(r'[+-]?\d{1,3}\.\d{3}', text.strip()):
        flags.append('ambiguous_decimal_or_thousands_separator')
    if re.search(r'\d[OoIl]|[OoIl]\d', text):
        flags.append('ambiguous_alphanumeric')
    if re.search(r'\d\s*[-~～至]\s*\d|[<>≤≥]', text):
        flags.append('range_or_comparison')
    return flags


def _ocr_rows(records):
    """按文字框垂直中心聚行；不依赖 OCR 服务返回顺序。"""
    rows = []
    for item in sorted(records, key=lambda r: ((r['bbox'][1] + r['bbox'][3]) / 2, r['bbox'][0])):
        cy = (item['bbox'][1] + item['bbox'][3]) / 2
        height = item['bbox'][3] - item['bbox'][1]
        if rows and abs(cy - rows[-1]['cy']) <= max(2, height * .45):
            rows[-1]['items'].append(item)
        else:
            rows.append({'cy': cy, 'items': [item]})
    for row in rows:
        row['items'].sort(key=lambda r: r['bbox'][0])
    return rows


def _run_ocr_page(page, engine, dpi=240):
    import numpy as np
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False)
    rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    found, _ = engine(rgb[:, :, ::-1].copy())
    records = []
    for box, text, confidence in found or []:
        xs, ys = zip(*box)
        bbox = [min(xs) * page.rect.width / pix.width, min(ys) * page.rect.height / pix.height,
                max(xs) * page.rect.width / pix.width, max(ys) * page.rect.height / pix.height]
        records.append({'text': str(text), 'confidence': float(confidence), 'bbox': bbox,
                        'flags': _ocr_risk_flags(str(text), float(confidence))})
    return records


def _apply_ocr(pdf_path, pages, tables, audit, enabled=True, dpi=240):
    engine = None
    with fitz.open(pdf_path) as doc:
        for page, output in zip(doc, pages):
            layer_chars = len(re.sub(r'\s+', '', page.get_text()))
            if layer_chars >= 10 or not page.get_images():
                continue
            page_num = page.number + 1
            if not enabled:
                audit.append({'page_num': page_num, 'status': 'disabled', 'records': []})
                logger.warning('第 %d 页疑似扫描页，OCR 已关闭', page_num)
                continue
            if engine is None:
                try:
                    local_deps = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ocr_deps")
                    if os.path.isdir(local_deps) and local_deps not in sys.path:
                        sys.path.insert(0, local_deps)
                    import onnxruntime
                    onnxruntime.disable_telemetry_events()
                    from rapidocr_onnxruntime import RapidOCR
                    engine = RapidOCR(intra_op_num_threads=2, inter_op_num_threads=1)
                except Exception as exc:
                    raise RuntimeError('扫描页需要本地 OCR。请用当前 Python 安装 requirements_ocr.txt 后重试；'
                                       '可使用 --no-ocr 仅做扫描检测。') from exc
            logger.info('第 %d/%d 页正在本地 OCR', page_num, len(doc))
            records = _run_ocr_page(page, engine, dpi)
            if not records:
                raise RuntimeError(f'第 {page_num} 页 OCR 没有返回文字，请检查扫描清晰度。')
            # 保留 OCR 原始文字；版面只做坐标排序，不对数字进行猜测替换。
            words = [tuple(r['bbox']) + (r['text'], i, 0, 0) for i, r in enumerate(records)]
            regions = _reading_order_regions(words, page.rect.width)
            text = '\n'.join(' '.join(w['text'] for w in row['items'])
                             for region in regions
                             for row in _ocr_rows([{'text': w[4], 'bbox': list(w[:4])} for w in region])).strip()
            output.update(text=text, char_count=len(text))
            detected, table_audit = _ocr_financial_tables(records, page_num, page.rect.width)
            # 仅纯扫描页补 OCR 表；原生表格不重复输出。
            if not any(t['page_num'] == page_num for t in tables):
                tables.extend(detected)
            audit.append({'page_num': page_num, 'status': 'completed', 'engine': 'rapidocr_onnxruntime',
                          'dpi': dpi, 'records': records, 'tables': table_audit,
                          'review_record_count': sum(bool(r['flags']) for r in records)})
            logger.info('第 %d 页 OCR 完成：%d 个文字框，%d 个财务表候选', page_num, len(records), len(detected))
    tables.sort(key=lambda t: (t['page_num'], t['bbox'][1] if t['bbox'] else 0))


def _link_continued_tables(result):
    """只关联证据充分的相邻页表；主 JSON 保留分页表，链内每行保留来源。"""
    tables = result['tables']
    links, dimensions = [], {}
    with fitz.open(result['file_path']) as doc:
        dimensions = {i + 1: (p.rect.width, p.rect.height) for i, p in enumerate(doc)}
    candidates = []
    for a in tables:
        if not a.get('bbox') or a['page_num'] not in dimensions:
            continue
        wa, ha = dimensions[a['page_num']]
        if a['bbox'][3] < .85 * ha:
            continue
        for b in tables:
            if b['page_num'] != a['page_num'] + 1 or not b.get('bbox'):
                continue
            wb, hb = dimensions[b['page_num']]
            if b['bbox'][1] > .15 * hb or a['headers'] != b['headers']:
                continue
            if any(abs(a['bbox'][k] / wa - b['bbox'][k] / wb) > .008 for k in (0,2)):
                continue
            if a.get('table_name') and b.get('table_name') and a['table_name'] != b['table_name']:
                continue
            candidates.append((a['table_id'], b['table_id']))
    for a,b in candidates:
        if sum(x == a for x,y in candidates) == 1 and sum(y == b for x,y in candidates) == 1:
            links.append((a,b))
    by_id = {t['table_id']: t for t in tables}
    next_id = dict(links)
    destinations = {b for a,b in links}
    chains = []
    for start in next_id:
        if start in destinations:
            continue
        members, cur = [], start
        while cur not in members:
            members.append(cur)
            if cur not in next_id:
                break
            cur = next_id[cur]
        chains.append({'table_ids': members, 'headers': by_id[start]['headers'],
                       'method': 'matching_headers_adjacent_pages_aligned_edges',
                       'rows': [{'source_page': by_id[t]['page_num'], 'source_table_id': t,
                                 'source_row_index': i, 'cells': row}
                                for t in members for i,row in enumerate(by_id[t]['rows'])],
                       'note': '关联视图，不自动拼接跨页断行；主 JSON 分页表保持不变。'})
    return chains


def extract_pdf(
    pdf_path: str,
    out_json: Optional[str] = None,
    header_rows: int = 3,
    infer_table_names: bool = True,
    filter_chart_text: bool = True,
    use_ocr: bool = True,
    ocr_audit: Optional[list] = None,
) -> Dict[str, Any]:
    """解析一份 PDF，返回统一结构，可选写入 JSON 文件。"""
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"找不到文件：{pdf_path}")

    logger.info("开始解析：%s", pdf_path)
    start = datetime.now()

    tables = extract_tables_with_pdfplumber(pdf_path, header_rows, infer_table_names)

    # 注意：不跳过表格区域提取文本。研报大表（如资产负债+利润+现金流+关键指标
    # 同页并排）的网格线范围与内容范围常不一致（内容会溢出网格线），按坐标跳过
    # 会误删内容。因此正文保留表格文本，表格另行结构化输出；P0-6 仅隔离确认的图文区。
    text_pages = extract_text_with_page(pdf_path, filter_chart_text=filter_chart_text)

    audit = ocr_audit if ocr_audit is not None else []
    _apply_ocr(pdf_path, text_pages, tables, audit, enabled=use_ocr)
    total_chars = sum(p["char_count"] for p in text_pages)

    result = {
        "file_name": os.path.basename(pdf_path),
        "file_path": os.path.abspath(pdf_path),
        "extract_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_pages": len(text_pages),
        "total_tables": len(tables),
        "total_chars": total_chars,
        "pages": text_pages,
        "tables": tables,
    }

    elapsed = (datetime.now() - start).total_seconds()
    logger.info(
        "解析完成：%d 页 / %d 个表格 / %d 字符，耗时 %.1f 秒",
        len(text_pages), len(tables), total_chars, elapsed,
    )

    if out_json:
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        logger.info("已写入：%s", out_json)
    return result


# ============================================================
# 4.5 导出 Word：把提取结果整理成可直接阅读的 Word 报告
# ============================================================
def _set_run_font(run, cn_font: str, size: float, bold: bool = False) -> None:
    """统一设置 run 的中西文字体、字号与加粗（中文字体走 eastAsia）。"""
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)


def _add_heading_cn(doc, text: str, level: int) -> None:
    """真实 Heading 样式 + 黑体 + 黑色（避免默认主题蓝）。"""
    h = doc.add_heading("", level=level)
    run = h.add_run(text)
    size, before, after = {1: (16, 14, 6), 2: (14, 11, 5)}.get(level, (12, 9, 4))
    _set_run_font(run, "黑体", size, bold=True)
    h.paragraph_format.space_before = Pt(before)
    h.paragraph_format.space_after = Pt(after)


def _set_cell_bg(cell, color: str) -> None:
    """给单元格加浅色底纹（表头用）。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tc_pr.append(shd)


def _fill_cell(cell, value: str, align_right: bool, bold: bool = False) -> None:
    """填单元格：垂直居中、无缩进、按需右对齐。"""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align_right else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.first_line_indent = Pt(0)
    _set_run_font(p.add_run(str(value)), "宋体", 10.5, bold=bold)


def _looks_numeric(text: str) -> bool:
    return bool(re.match(r"^\s*-?\d", str(text)))


def export_to_docx(result: Dict[str, Any], out_docx: str) -> None:
    """把提取结果写入 Word 报告：A4 / 2.5cm 边距 / 黑色标题 / 浅灰表头。"""
    from docx import Document
    from docx.shared import Cm

    doc = Document()

    # 页面：A4 + 上下左右 2.5cm
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.5)

    # Normal：宋体 12pt、1.5 倍行距、首行缩进 2 字符
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(12)
    _rpr = normal.element.get_or_add_rPr()
    _rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    npf = normal.paragraph_format
    npf.line_spacing = 1.5
    npf.first_line_indent = Pt(24)

    # 主标题
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run_font(title.add_run(f"{result['file_name']} —— PDF 提取报告"), "黑体", 18, bold=True)
    title.paragraph_format.space_after = Pt(18)

    # 副标题：提取时间与统计
    subtitle = doc.add_paragraph(style="Subtitle")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.first_line_indent = Pt(0)
    _set_run_font(
        subtitle.add_run(
            f"提取时间：{result['extract_time']}　|　共 {result['total_pages']} 页 / "
            f"{result['total_tables']} 个表格 / {result['total_chars']} 字符"
        ),
        "楷体", 12,
    )
    subtitle.paragraph_format.space_after = Pt(12)

    # 一、正文文本
    _add_heading_cn(doc, "一、正文文本", 1)
    for p in result["pages"]:
        _add_heading_cn(doc, f"第 {p['page_num']} 页（{p['char_count']} 字）", 2)
        body = doc.add_paragraph()
        _set_run_font(body.add_run(p["text"] or "（本页无文本）"), "宋体", 12)

    # 二、表格数据
    _add_heading_cn(doc, "二、表格数据", 1)
    if not result["tables"]:
        body = doc.add_paragraph()
        _set_run_font(body.add_run("（未检测到表格）"), "宋体", 12)
    for i, t in enumerate(result["tables"], start=1):
        # 表题（Caption，置于表上方，与表同页）
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.first_line_indent = Pt(0)
        cap.paragraph_format.space_after = Pt(4)
        cap.paragraph_format.keep_with_next = True
        name = t.get("table_name") or "（未识别表名）"
        _set_run_font(cap.add_run(f"表{i}：{name}（第 {t['page_num']} 页）"), "宋体", 10.5, bold=True)

        headers = t["headers"]
        rows = t["rows"]
        # 逐列判断是否数字列（多数单元格以数字开头则整列右对齐）
        numeric_cols = []
        for h in headers:
            hits = 0
            for row in rows:
                cell_val = row.get(h)
                raw = cell_val.get("raw", "") if isinstance(cell_val, dict) else str(cell_val or "")
                if _looks_numeric(raw):
                    hits += 1
            numeric_cols.append(hits > len(rows) / 2 if rows else False)

        table = doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = "Table Grid"
        # 表头：浅灰底纹 + 加粗居中
        for j, h in enumerate(headers):
            _fill_cell(table.cell(0, j), h, align_right=False, bold=True)
            _set_cell_bg(table.cell(0, j), "D9D9D9")
        # 数据行
        for r_i, row in enumerate(rows, start=1):
            for j, h in enumerate(headers):
                cell_val = row.get(h)
                raw = cell_val.get("raw", "") if isinstance(cell_val, dict) else str(cell_val or "")
                _fill_cell(table.cell(r_i, j), raw, align_right=numeric_cols[j])

    doc.save(out_docx)


# ============================================================
# 5. 命令行入口
# ============================================================
def _preview(result: Dict[str, Any], table_show: int = 3, text_show: int = 2) -> None:
    print("\n" + "=" * 60)
    print(f"文件：{result['file_name']}   页数：{result['total_pages']}   表格：{result['total_tables']}")
    print("=" * 60)
    for p in result["pages"][:text_show]:
        print(f"\n【第 {p['page_num']} 页 · 文本预览】")
        print(p["text"][:200].replace("\n", " "))
    for t in result["tables"][:table_show]:
        print(f"\n【表格 {t['table_id']} · 第 {t['page_num']} 页】标题：{t['table_name'] or '(未识别)'}")
        print("表头：", " | ".join(t["headers"]))
        for row in t["rows"][:5]:
            vals = []
            for k, v in row.items():
                if isinstance(v, dict):
                    vals.append(f"{k}={v['raw']}({v['number']}{v['unit'] or ''})")
                else:
                    vals.append(f"{k}={v}")
            print("   ", " | ".join(vals)[:200])


def build_quality_report(result, ocr_audit=None):
    """独立质量报告和内容索引；原提取 JSON 不增加字段。告警不等于提取错误。"""
    issues, pages, content = [], [], []
    tables = result.get('tables', [])
    ocr_audit = ocr_audit or []
    ocr_pages = {p['page_num']: p for p in ocr_audit}
    ids = [t.get('table_id') for t in tables]
    if len(ids) != len(set(ids)):
        issues.append({'code': 'duplicate_table_id', 'severity': 'error'})
    with fitz.open(result['file_path']) as doc:
        for i, page in enumerate(doc, 1):
            text_page = next((p for p in result['pages'] if p['page_num'] == i), None)
            pts = [t for t in tables if t['page_num'] == i]
            layer_chars = len(re.sub(r'\s+', '', page.get_text()))
            image_count = len(page.get_images())
            flags = []
            if i in ocr_pages and ocr_pages[i]['status'] == 'completed':
                flags.append('ocr_review_required')
                issues.append({'code': flags[-1], 'severity': 'warning', 'page_num': i})
            elif layer_chars < 10:
                flags.append('suspected_scan' if image_count else 'low_text')
                issues.append({'code': flags[-1], 'severity': 'warning', 'page_num': i})
            if text_page is None or text_page['char_count'] != len(text_page['text']):
                issues.append({'code': 'page_text_count_mismatch', 'severity': 'error', 'page_num': i})
            values = [cell for t in pts for row in t['rows'] for cell in row.values()]
            unresolved = sum(c.get('number') is None and bool(re.search(r'\d', c.get('raw', ''))) for c in values)
            unit_missing = sum(c.get('number') is not None and not c.get('unit') for c in values)
            pages.append({'page_num': i, 'text_layer_chars': layer_chars,
                          'output_chars': text_page['char_count'] if text_page else 0,
                          'table_count': len(pts), 'image_count': image_count,
                          'unparsed_digit_cells': unresolved, 'numeric_cells_without_unit': unit_missing,
                          'extraction_method': ('ocr' if i in ocr_pages and ocr_pages[i]['status'] == 'completed'
                                                else 'no_text' if layer_chars < 10 else 'native'),
                          'flags': flags})
            content.append({'type': 'page_text', 'page_num': i, 'ref': f'pages[{i - 1}].text',
                            'may_include_tables_and_chart_text': True})
            charts = _detect_chart_text(page, page.get_text('words'))
            for chart in charts:
                content.append({'type': 'chart', 'page_num': i, 'bbox': list(chart['bbox']),
                                'caption': chart['caption'], 'detection': 'heuristic'})
    for i, t in enumerate(tables):
        content.append({'type': 'table', 'page_num': t['page_num'], 'table_id': t['table_id'],
                        'bbox': t['bbox'], 'ref': f'tables[{i}]'})
        if t['row_count'] != len(t['rows']) or len(t['headers']) != len(set(t['headers'])) or any(
                set(row) != set(t['headers']) for row in t['rows']):
            issues.append({'code': 'table_structure_mismatch', 'severity': 'error', 'table_id': t['table_id']})
    if (result['total_pages'] != len(pages) or result['total_tables'] != len(tables)
            or result['total_chars'] != sum(p['char_count'] for p in result['pages'])):
        issues.append({'code': 'document_count_mismatch', 'severity': 'error'})
    logger.info('质量检查：%d 页，%d 条告警/错误；内容索引 %d 项', len(pages), len(issues), len(content))
    chains = _link_continued_tables(result)
    return {'report_version': 2, 'ocr': ocr_audit, 'continued_tables': chains, 'file_name': result['file_name'],
            'summary': {'pages': len(pages), 'tables': len(tables),
                        'errors': sum(i['severity'] == 'error' for i in issues),
                        'warnings': sum(i['severity'] == 'warning' for i in issues)},
            'pages': pages, 'content_index': content, 'issues': issues,
            'notes': ['OCR 状态以 ocr 记录为准；置信度不是准确率，扫描金融数据需要复核。',
                      '含数字但未解析的单元格可能是代码、期间或复杂文字，不等于丢失。',
                      '图表索引仅覆盖启发式确认区域，page_text 仍可能包含表格和图文。']}


def main() -> None:
    parser = argparse.ArgumentParser(description="PDF 解析模块（文本 + 表格 + 页码留痕）")
    parser.add_argument("pdf", help="要解析的 PDF 文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出 JSON 文件路径（默认不存）")
    parser.add_argument("--docx", nargs="?", const="__desktop__", default=None,
                        help="输出 Word 报告路径；不带值则自动保存到桌面")
    parser.add_argument("--preview", action="store_true", help="解析后打印预览")
    parser.add_argument("--header-rows", type=int, default=3, help="自动识别表头层数上限（默认 3）")
    parser.add_argument("--no-table-name", action="store_true", help="关闭表格标题推断")
    parser.add_argument('--keep-chart-text', action='store_true',
                        help='保留正文中的图表文字（关闭 P0-6 过滤，便于与 P0-5 对照）')
    parser.add_argument("--quality-report", help="独立质量报告路径；指定 -o 时默认写同名 _quality.json")
    parser.add_argument("--no-ocr", action="store_true", help="关闭扫描页本地 OCR，仅输出已有文字层")
    args = parser.parse_args()
    report_path = args.quality_report or (os.path.splitext(args.output)[0] + "_quality.json" if args.output else None)
    targets = [os.path.realpath(p) for p in (args.pdf, args.output, report_path) if p]
    if len(targets) != len(set(targets)):
        parser.error("输入 PDF、提取 JSON 和质量报告必须使用不同路径")

    ocr_audit = []
    result = extract_pdf(
        args.pdf,
        out_json=args.output,
        header_rows=args.header_rows,
        infer_table_names=not args.no_table_name,
        filter_chart_text=not args.keep_chart_text,
        use_ocr=not args.no_ocr,
        ocr_audit=ocr_audit,
    )
    report_path = args.quality_report or (os.path.splitext(args.output)[0] + "_quality.json" if args.output else None)
    if report_path:
        report = build_quality_report(result, ocr_audit)
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        logger.info("质量报告已写入：%s", report_path)
    if args.preview:
        _preview(result)
    if args.docx:
        if args.docx == "__desktop__":
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.isdir(desktop):
                desktop = os.path.dirname(os.path.abspath(args.pdf))
            out_docx = os.path.join(desktop, os.path.splitext(os.path.basename(args.pdf))[0] + "_提取结果.docx")
        else:
            out_docx = args.docx
        export_to_docx(result, out_docx)
        logger.info("已生成 Word 报告：%s", out_docx)


if __name__ == "__main__":
    main()
