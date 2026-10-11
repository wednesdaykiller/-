# -*- coding: utf-8 -*-
"""
规则加载与归一化模块 —— 金融人工智能竞赛 · 研报核查方向

职责：
  1. 加载《金融规则制定3.0.xlsx》四张表：核查字段 / 换算表 / 常见错误 / 业务关键词库
  2. 建立「别名 → 标准字段名」映射，长名优先（避免"净利润"被"归母净利"抢匹配）
  3. 单位换算：千元/万元/百万元/元 → 统一亿元；%、倍 保留
  4. 时间归一化：26H1 / 2026H1 / 1H26 / 2025年 / 2025A / 2025E / 25Q1 ...
     → (year, period)，period ∈ {"FY","H1","H2","Q1","Q2","Q3","Q4"}
  5. 集中存放常见错误表的判定阈值（规则1~37 的可程序化部分）

不负责：字段抽取、大模型调用、报告输出。
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

logger = logging.getLogger("rules")


# ============================================================
# 数据结构
# ============================================================
@dataclass
class FieldRule:
    """核查字段表中的一行。"""
    canonical: str                       # 标准字段名（如 "营业收入"）
    report_name: str                     # 财报标准叫法
    aliases: List[str] = field(default_factory=list)   # 研报可能叫法
    report_unit: str = ""                # 财报单位
    broker_unit: str = ""                # 研报单位
    standard_unit: str = ""              # 统一单位
    note: str = ""                       # 备注（口径说明等）


@dataclass
class NormalizedValue:
    """一个被抽出来的数值。"""
    field: str               # 标准字段名（未知时为原始别名）
    period: Tuple[int, str]  # (year, period)，如 (2024, "FY")
    raw_text: str            # 原文（留痕）
    number: Optional[float]  # 数值（已按单位换算成 standard_unit 的量纲）
    unit: str                # 原始单位（"亿元"/"万元"/"%"...）
    raw_number: Optional[float] = None  # 原始未换算数值（以 unit 为单位读出来的数）
    std_unit: str = "亿元"   # 归一化后的单位
    source: str = ""         # 来源：page{n}/table_id/正文句
    page_num: int = 0
    sub_label: str = ""      # 分产品/分渠道/分品牌时的子类名（如"书写工具"）
    company: str = ""        # 子公司/主体名（如有）
    is_estimate: bool = False  # 是否预测值（E）


# ============================================================
# 1. 规则加载
# ============================================================
# 金额单位 → 到"亿元"的换算倍率
_UNIT_TO_YIYI = {
    "元": 1e-8,
    "千元": 1e-5,
    "万元": 1e-4,
    "百万元": 1e-2,
    "亿元": 1.0,
    "亿": 1.0,
    "万": 1e-4,
    "千": 1e-5,
    "百万": 1e-2,
}

# 期间正则：2024 / 2024A / 2024年报 / 2024年 / 2024H1 / 25H1 / 1Q25 / 2025E / 2025Q1...
_YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})")
_HQ_RE = re.compile(r"(?<!\d)(\d{2})\s*([HQ])\s*([1-4])")
_HQ_FULL_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})\s*([HQ])\s*([1-4])")
_SUFFIX_RE = re.compile(r"(A|E|Reported|Estimate|年报|半年报|一季报|三季报|中报|预测|预告)")


class RuleBook:
    def __init__(self, rule_xlsx: str):
        self.path = rule_xlsx
        self.fields: Dict[str, FieldRule] = {}      # canonical -> FieldRule
        self.alias_index: List[Tuple[str, str]] = []  # (alias_lower, canonical)，按 alias 长度降序
        self._load(rule_xlsx)

    # ---------- 内部：读四张表 ----------
    def _load(self, path: str) -> None:
        wb = openpyxl.load_workbook(path, data_only=True)

        # 1) 核查字段
        ws = wb["核查字段"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            canonical = str(row[0]).strip()
            report_name = str(row[1] or "").strip()
            alias_cell = str(row[2] or "").strip()
            aliases = [a.strip() for a in re.split(r"[，,、/]", alias_cell) if a.strip()]
            # canonical 本身也加入别名
            aliases = [canonical] + aliases
            self.fields[canonical] = FieldRule(
                canonical=canonical,
                report_name=report_name,
                aliases=aliases,
                report_unit=str(row[3] or "").strip(),
                broker_unit=str(row[4] or "").strip(),
                standard_unit=str(row[5] or "").strip(),
                note=str(row[7] or "").strip() if len(row) > 7 else "",
            )
        # 别名索引：长 alias 优先，避免"净利润"抢了"归母净利润"
        for canon, fr in self.fields.items():
            for a in fr.aliases:
                self.alias_index.append((a.lower(), canon))
        self.alias_index.sort(key=lambda x: -len(x[0]))

        logger.info("已加载核查字段 %d 个，别名索引 %d 条",
                    len(self.fields), len(self.alias_index))
        wb.close()

    # ---------- 字段名匹配 ----------
    def match_field(self, text: str) -> Optional[str]:
        """从一段科目文本里匹配标准字段名；匹配不到返回 None。
        长名优先：'归母净利润（万元）' 命中 '归母净利' 而不是 '净利润'。"""
        if not text:
            return None
        t = text.lower().replace(" ", "")
        for alias, canon in self.alias_index:
            a = alias.replace(" ", "")
            if a and a in t:
                return canon
        return None

    # ---------- 单位换算 ----------
    @staticmethod
    def to_yiyyi(number: float, unit: Optional[str]) -> Optional[float]:
        """把 number 按 unit 换算成「亿元」。比率（%）、倍、家、元/股 不换算。"""
        if number is None or unit is None:
            return number
        u = unit.strip()
        # 比率 / 倍数 / 家数 / 每股 不换算
        if u in ("%", "％", "倍", "家", "个", "元/股", "元", ""):
            # "元" 出现在金额语境时要换算，但如果是每股指标已被元/股捕获；
            # 单独的"元"按金额处理
            if u == "元":
                return number * 1e-8
            return number
        factor = _UNIT_TO_YIYI.get(u)
        if factor is None:
            # 尝试去掉末尾的"元"再查（如"人民币亿元"）
            for k, v in _UNIT_TO_YIYI.items():
                if u.endswith(k) or k in u:
                    return number * v
            return number
        return number * factor

    # ---------- 时间归一化 ----------
    @staticmethod
    def normalize_period(text: str) -> Optional[Tuple[int, str]]:
        """把表头/正文里的时间表述归一为 (year, period)。
        period ∈ {"FY","H1","H2","Q1","Q2","Q3","Q4"}。
        识别失败返回 None。"""
        if not text:
            return None
        t = str(text).strip()

        # 1) 显式 H/Q：2024H1、25H1、1Q25、2025Q3
        m = _HQ_FULL_RE.search(t)
        if m:
            year = int(m.group(1)); kind = m.group(2); q = m.group(3)
            if kind == "H":
                return (year, "H1" if q == "1" else "H2")
            return (year, f"Q{q}")
        m = _HQ_RE.search(t)
        if m:
            yy = int(m.group(1)); year = 2000 + yy
            kind = m.group(2); q = m.group(3)
            if kind == "H":
                return (year, "H1" if q == "1" else "H2")
            return (year, f"Q{q}")

        # 2) 1Q26 / 2Q25 这种"季度在前"
        m = re.search(r"(?<!\d)([1-4])\s*Q\s*(\d{2})(?!\d)", t, re.IGNORECASE)
        if m:
            q = m.group(1); year = 2000 + int(m.group(2))
            return (year, f"Q{q}")

        # 3) 普通年份：2024 / 2024年 / 2024A / 2024年报 / 2025E / 2025年预测
        m = _YEAR_RE.search(t)
        if m:
            year = int(m.group(1))
            tail = t[m.end():]
            # 后缀判断：半年报/中报 → H1；三季报 → Q3累计；一季报 → Q1
            if any(k in tail for k in ("半年报", "中报")):
                return (year, "H1")
            if "三季报" in tail or "前三季" in tail:
                return (year, "Q3")
            if "一季报" in tail:
                return (year, "Q1")
            if "年报" in tail or re.search(r"[年A\s]|$", tail):
                # 全年口径（含 A 年报、E 预测年）
                return (year, "FY")
            return (year, "FY")
        return None

    @staticmethod
    def period_key(p: Tuple[int, str]) -> str:
        return f"{p[0]}{p[1]}"


# ============================================================
# 2. 判定阈值（对应常见错误表 1~37 的可程序化部分）
# ============================================================
# 金额类（相对差 = |研报-财报|/财报，单位统一到亿元后比较）
AMT_ERR_REL = 0.05      # 规则1：>5% 判错
AMT_WARN_REL = 0.02     # 规则2：2%~5% 疑似；规则3：≤2% 不判错
AMT_CONVERT_ERR = 0.01  # 规则9：换算后仍差>1% 判错

# 比率类（绝对差，单位：百分点）
RATIO_ERR_PP = 1.0      # 规则4：>1pp 判错
RATIO_WARN_PP = 0.5     # 规则5：0.5~1pp 疑似；规则6：<0.5pp 不判错

# 判定结果枚举
JUDGE_OK = "不判错"
JUDGE_ERR = "判错"
JUDGE_WARN = "标记异常，待人工确认"
JUDGE_MISSING = "资料缺失，无法核验"
JUDGE_NEW_FIELD = "待补充字段"
JUDGE_SAME_SOURCE = "同源"
