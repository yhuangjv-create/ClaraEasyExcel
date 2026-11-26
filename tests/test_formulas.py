"""Smoke tests for intent detection + formula build + base diagnose."""
from __future__ import annotations

from mkr.formulas.base import all_formula_names, get_formula
from mkr.intent import detect


def test_registry_loaded() -> None:
    names = all_formula_names()
    assert "VLOOKUP" in names
    assert "SUMIF" in names
    assert "COUNTIF" in names
    assert "IFERROR" in names
    assert "CONCAT" in names
    assert len(names) == 5  # no duplicates — aliases don't create new cards

    # CONCATENATE is an ALIAS (not canonical) but get_formula still works
    assert get_formula("CONCATENATE") is get_formula("CONCAT")  # same instance


def test_intent_vlookup() -> None:
    hits = detect("我想根据订单号匹配价格")
    assert any(h.formula == "VLOOKUP" for h in hits)
    hits2 = detect("vlookup 找不到")
    assert any(h.formula == "VLOOKUP" for h in hits2)


def test_intent_sumif() -> None:
    hits = detect("条件求和 大于100")
    assert any(h.formula == "SUMIF" for h in hits)


def test_intent_countif() -> None:
    hits = detect("统计已完成的个数")
    assert any(h.formula == "COUNTIF" for h in hits)


def test_intent_iferror() -> None:
    hits = detect("避免 #N/A 报错")
    assert any(h.formula == "IFERROR" for h in hits)


def test_intent_concat() -> None:
    hits = detect("把姓名和工号拼接起来")
    assert any(h.formula in ("CONCAT", "CONCATENATE") for h in hits)


def test_intent_empty() -> None:
    assert detect("") == []
    assert detect("   ") == []


def test_vlookup_build() -> None:
    f = get_formula("VLOOKUP")
    assert f is not None
    body = f.build(lookup="A2", table="$D:$F", col_index="3", match_mode="精确匹配 (FALSE)")
    assert body.startswith("VLOOKUP(")
    assert "FALSE" in body
    # $ auto-absolute applied
    assert "$D:$F" in body


def test_vlookup_diagnose_na() -> None:
    f = get_formula("VLOOKUP")
    sugg = f.diagnose("=VLOOKUP(A2,$D:$F,3,FALSE)", "#N/A")
    assert len(sugg) >= 2


def test_sumif_build() -> None:
    f = get_formula("SUMIF")
    body = f.build(range="B:B", criteria='">100"', sum_range="C:C")
    assert body.startswith("SUMIF(")
    assert ">100" in body


def test_sumif_build_no_sumrange() -> None:
    f = get_formula("SUMIF")
    body = f.build(range="B:B", criteria='"甲"', sum_range="")
    assert body.count(",") == 1


def test_countif_build() -> None:
    f = get_formula("COUNTIF")
    body = f.build(range="A:A", criteria='"已完成"')
    assert body.startswith("COUNTIF(")


def test_iferror_build() -> None:
    f = get_formula("IFERROR")
    body = f.build(inner="VLOOKUP(A2,$D:$F,3,FALSE)", fallback='""')
    assert body.startswith("IFERROR(VLOOKUP(")


def test_concat_build() -> None:
    f = get_formula("CONCAT")
    body = f.build(items='A2, "-", B2')
    assert body == 'CONCAT(A2, "-", B2)'
