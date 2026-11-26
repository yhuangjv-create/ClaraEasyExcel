"""Keyword → formula intent detection.

A small hand-tuned keyword map for MVP v1.0. No NLP / LLM yet.
Each formula has a list of keyword groups; user input is tokenized and
matched against each group. Matched groups contribute to a score, returned
in descending order.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class IntentHit:
    formula: str
    score: float
    matched_keywords: tuple[str, ...]


# keyword groups — any token in a group matches; multiple groups hit → score += len(group)
INTENT_MAP: dict[str, list[tuple[str, ...]]] = {
    "VLOOKUP": [
        ("匹配", "查找", "搜索", "对应", "vlookup", "lookup", "找一下", "对应值"),
        ("匹配表", "查找表", "数据表", "字典表"),
    ],
    "SUMIF": [
        ("求和", "汇总", "加总", "sumif", "条件求和", "合计"),
        ("条件", "满足", "大于", "小于", "等于"),
    ],
    "COUNTIF": [
        ("计数", "个数", "次数", "countif", "统计数量", "出现次数"),
        ("条件", "满足", "大于", "小于", "等于", "包含"),
    ],
    "IFERROR": [
        ("错误", "报错", "容错", "iferror", "空白", "显示无", "避免报错", "#n/a", "#div/0"),
    ],
    "CONCATENATE": [
        ("拼接", "合并文本", "连起来", "concatenate", "concat", "&", "组合起来", "加在一起"),
        ("文本", "文字", "字符串", "单元格内容"),
    ],
}


def detect(query: str) -> list[IntentHit]:
    """Tokenize `query` (split by whitespace + punctuation) and score each formula.

    Returns list of hits sorted by score desc, score > 0 only.
    """
    if not query or not query.strip():
        return []

    # normalize: lowercase, split on CJK punctuation + ASCII whitespace
    tokens = []
    buf = []
    for ch in query.lower():
        if ch.isspace() or ch in "，。！？、；：,.;:!?，":
            if buf:
                tokens.append("".join(buf))
                buf = []
        else:
            buf.append(ch)
    if buf:
        tokens.append("".join(buf))
    # also split compound tokens on ASCII whitespace
    flat: list[str] = []
    for t in tokens:
        flat.extend([x for x in t.split(" ") if x])
    tokens = flat

    hits: list[IntentHit] = []
    for formula, groups in INTENT_MAP.items():
        matched_kw: set[str] = set()
        score = 0.0
        for group in groups:
            for kw in group:
                for token in tokens:
                    if kw.lower() in token or token in kw.lower():
                        matched_kw.add(kw)
                        score += 1.0
                        break
        if score > 0:
            hits.append(IntentHit(formula=formula, score=score, matched_keywords=tuple(sorted(matched_kw))))

    hits.sort(key=lambda h: h.score, reverse=True)
    return hits
