from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple
import pandas as pd
import numpy as np

REQUIRED_COLUMNS = {
    "매출상세": ["회계연도", "회계월", "고객코드", "고객명", "제품코드", "제품명", "수량", "판매단가", "매출액"],
    "제조원가상세": ["회계연도", "회계월", "제품코드", "제품명", "원가구분", "금액"],
    "판관비상세": ["회계연도", "회계월", "부서코드", "부서명", "계정코드", "계정과목", "거래처코드", "거래처명", "금액"],
    "예산": ["회계연도", "회계월", "손익항목", "예산금액"],
}

NUMERIC_COLUMNS = {
    "매출상세": ["회계연도", "회계월", "수량", "판매단가", "매출액"],
    "제조원가상세": ["회계연도", "회계월", "금액"],
    "판관비상세": ["회계연도", "회계월", "금액"],
    "예산": ["회계연도", "회계월", "예산금액"],
}

@dataclass(frozen=True)
class Period:
    year: int
    month: int

    @property
    def label(self) -> str:
        return f"{self.year}년 {self.month:02d}월"


def period_key(df: pd.DataFrame) -> pd.Series:
    return pd.to_numeric(df["회계연도"], errors="coerce").astype("Int64") * 100 + pd.to_numeric(df["회계월"], errors="coerce").astype("Int64")


def normalize_data(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    for name, df in data.items():
        x = df.copy()
        x.columns = [str(c).strip() for c in x.columns]
        for col in NUMERIC_COLUMNS.get(name, []):
            if col in x.columns:
                x[col] = pd.to_numeric(x[col], errors="coerce")
        out[name] = x
    return out


def validate_data(data: Dict[str, pd.DataFrame]) -> List[str]:
    issues: List[str] = []
    for sheet, cols in REQUIRED_COLUMNS.items():
        if sheet not in data:
            issues.append(f"필수 자료 '{sheet}'가 없습니다.")
            continue
        missing = [c for c in cols if c not in data[sheet].columns]
        if missing:
            issues.append(f"{sheet}: 필수 열 누락 - {', '.join(missing)}")
    return issues


def available_periods(data: Dict[str, pd.DataFrame]) -> List[Period]:
    periods = set()
    for name in ["매출상세", "제조원가상세", "판관비상세"]:
        df = data.get(name)
        if df is None or df.empty or "회계연도" not in df or "회계월" not in df:
            continue
        for y, m in df[["회계연도", "회계월"]].dropna().drop_duplicates().itertuples(index=False, name=None):
            periods.add((int(y), int(m)))
    return [Period(y, m) for y, m in sorted(periods)]


def _filter_period(df: pd.DataFrame, p: Period) -> pd.DataFrame:
    return df[(pd.to_numeric(df["회계연도"], errors="coerce") == p.year) & (pd.to_numeric(df["회계월"], errors="coerce") == p.month)].copy()


def pnl(data: Dict[str, pd.DataFrame], p: Period) -> Dict[str, float]:
    sales = _filter_period(data["매출상세"], p)
    cogs = _filter_period(data["제조원가상세"], p)
    sga = _filter_period(data["판관비상세"], p)
    revenue = float(pd.to_numeric(sales["매출액"], errors="coerce").fillna(0).sum())
    cogs_amt = float(pd.to_numeric(cogs["금액"], errors="coerce").fillna(0).sum())
    sga_amt = float(pd.to_numeric(sga["금액"], errors="coerce").fillna(0).sum())
    gross = revenue - cogs_amt
    op = gross - sga_amt
    return {
        "매출액": revenue,
        "매출원가": cogs_amt,
        "매출총이익": gross,
        "판매비와관리비": sga_amt,
        "영업이익": op,
        "매출총이익률": gross / revenue if revenue else np.nan,
        "영업이익률": op / revenue if revenue else np.nan,
    }


def pnl_comparison(data: Dict[str, pd.DataFrame], prior: Period, current: Period) -> pd.DataFrame:
    a = pnl(data, prior)
    b = pnl(data, current)
    rows = []
    for item in ["매출액", "매출원가", "매출총이익", "판매비와관리비", "영업이익"]:
        prev = a[item]
        cur = b[item]
        rows.append({
            "손익항목": item,
            prior.label: prev,
            current.label: cur,
            "증감금액": cur - prev,
            "증감률": (cur - prev) / abs(prev) if prev else np.nan,
        })
    return pd.DataFrame(rows)


def operating_profit_bridge(data: Dict[str, pd.DataFrame], prior: Period, current: Period) -> pd.DataFrame:
    a = pnl(data, prior)
    b = pnl(data, current)
    revenue_effect = b["매출액"] - a["매출액"]
    cogs_effect = -(b["매출원가"] - a["매출원가"])
    sga_effect = -(b["판매비와관리비"] - a["판매비와관리비"])
    return pd.DataFrame([
        {"구분": "기준월 영업이익", "영향금액": a["영업이익"], "유형": "기준"},
        {"구분": "매출액 증감 영향", "영향금액": revenue_effect, "유형": "변동"},
        {"구분": "매출원가 증감 영향", "영향금액": cogs_effect, "유형": "변동"},
        {"구분": "판매비와관리비 증감 영향", "영향금액": sga_effect, "유형": "변동"},
        {"구분": "분석월 영업이익", "영향금액": b["영업이익"], "유형": "결과"},
    ])


def sales_price_volume_analysis(data: Dict[str, pd.DataFrame], prior: Period, current: Period, dims: List[str] | None = None) -> pd.DataFrame:
    dims = dims or ["제품코드", "제품명", "고객코드", "고객명"]
    df = data["매출상세"].copy()
    prev = _filter_period(df, prior)
    cur = _filter_period(df, current)

    def agg(x: pd.DataFrame, suffix: str) -> pd.DataFrame:
        g = x.groupby(dims, dropna=False).agg(수량=("수량", "sum"), 매출액=("매출액", "sum")).reset_index()
        g[f"평균판매단가_{suffix}"] = np.where(g["수량"] != 0, g["매출액"] / g["수량"], 0.0)
        return g.rename(columns={"수량": f"수량_{suffix}", "매출액": f"매출액_{suffix}"})

    p = agg(prev, "기준")
    c = agg(cur, "분석")
    m = p.merge(c, on=dims, how="outer").fillna(0)

    q0 = m["수량_기준"].astype(float)
    q1 = m["수량_분석"].astype(float)
    p0 = m["평균판매단가_기준"].astype(float)
    p1 = m["평균판매단가_분석"].astype(float)

    existing = (q0 > 0) & (q1 > 0)
    new = (q0 == 0) & (q1 > 0)
    discontinued = (q0 > 0) & (q1 == 0)

    m["판매량효과"] = np.where(existing, (q1 - q0) * p0, 0.0)
    m["판매단가효과"] = np.where(existing, q1 * (p1 - p0), 0.0)
    m["신규매출효과"] = np.where(new, m["매출액_분석"], 0.0)
    m["판매중단효과"] = np.where(discontinued, -m["매출액_기준"], 0.0)
    m["매출증감"] = m["매출액_분석"] - m["매출액_기준"]
    m["효과합계"] = m[["판매량효과", "판매단가효과", "신규매출효과", "판매중단효과"]].sum(axis=1)
    m["검산차이"] = m["매출증감"] - m["효과합계"]
    return m.sort_values("매출증감", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)


def change_by_dimension(df: pd.DataFrame, prior: Period, current: Period, dims: List[str], value_col: str = "금액", impact_sign: float = -1.0) -> pd.DataFrame:
    p = _filter_period(df, prior).groupby(dims, dropna=False)[value_col].sum().rename("기준월금액")
    c = _filter_period(df, current).groupby(dims, dropna=False)[value_col].sum().rename("분석월금액")
    m = pd.concat([p, c], axis=1).fillna(0).reset_index()
    m["증감금액"] = m["분석월금액"] - m["기준월금액"]
    m["영업이익영향"] = m["증감금액"] * impact_sign
    return m.sort_values("영업이익영향", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)


def cogs_change(data: Dict[str, pd.DataFrame], prior: Period, current: Period, dims: List[str] | None = None) -> pd.DataFrame:
    return change_by_dimension(data["제조원가상세"], prior, current, dims or ["원가구분"], "금액", -1.0)


def sga_change(data: Dict[str, pd.DataFrame], prior: Period, current: Period, dims: List[str] | None = None) -> pd.DataFrame:
    return change_by_dimension(data["판관비상세"], prior, current, dims or ["계정과목"], "금액", -1.0)


def budget_comparison(data: Dict[str, pd.DataFrame], current: Period) -> pd.DataFrame:
    actual = pnl(data, current)
    b = _filter_period(data["예산"], current).groupby("손익항목")["예산금액"].sum().to_dict()
    rows = []
    for item in ["매출액", "매출원가", "판매비와관리비", "영업이익"]:
        actual_amt = actual[item]
        budget_amt = float(b.get(item, np.nan))
        variance = actual_amt - budget_amt if not np.isnan(budget_amt) else np.nan
        rows.append({
            "손익항목": item,
            "예산금액": budget_amt,
            "실적금액": actual_amt,
            "차이금액": variance,
            "달성률": actual_amt / budget_amt if budget_amt not in (0, np.nan) and not np.isnan(budget_amt) else np.nan,
        })
    return pd.DataFrame(rows)


def combined_root_causes(data: Dict[str, pd.DataFrame], prior: Period, current: Period) -> pd.DataFrame:
    rows = []
    sales = sales_price_volume_analysis(data, prior, current)
    for effect in ["판매량효과", "판매단가효과", "신규매출효과", "판매중단효과"]:
        val = float(sales[effect].sum())
        if abs(val) > 1e-9:
            rows.append({"영역": "매출", "원인": effect, "영업이익영향": val})
    c = cogs_change(data, prior, current, ["원가구분"])
    for _, r in c.iterrows():
        if abs(r["영업이익영향"]) > 1e-9:
            rows.append({"영역": "매출원가", "원인": str(r["원가구분"]), "영업이익영향": float(r["영업이익영향"])})
    s = sga_change(data, prior, current, ["계정과목"])
    for _, r in s.iterrows():
        if abs(r["영업이익영향"]) > 1e-9:
            rows.append({"영역": "판매비와관리비", "원인": str(r["계정과목"]), "영업이익영향": float(r["영업이익영향"])})
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values("영업이익영향", key=lambda x: x.abs(), ascending=False).reset_index(drop=True)


def management_report(data: Dict[str, pd.DataFrame], prior: Period, current: Period, materiality: float = 1_000_000) -> str:
    prev = pnl(data, prior)
    cur = pnl(data, current)
    delta = cur["영업이익"] - prev["영업이익"]
    causes = combined_root_causes(data, prior, current)
    causes = causes[causes["영업이익영향"].abs() >= materiality] if not causes.empty else causes
    neg = causes[causes["영업이익영향"] < 0].nsmallest(3, "영업이익영향") if not causes.empty else causes
    pos = causes[causes["영업이익영향"] > 0].nlargest(3, "영업이익영향") if not causes.empty else causes
    budget = budget_comparison(data, current)
    op_budget = budget[budget["손익항목"] == "영업이익"]

    def won(v: float) -> str:
        sign = "-" if v < 0 else ""
        return f"{sign}{abs(v):,.0f}원"

    lines = [
        f"[{current.label} 월마감 손익 요약]",
        f"- 영업이익: {won(cur['영업이익'])} (기준월 {won(prev['영업이익'])}, 증감 {won(delta)})",
        f"- 영업이익률: {cur['영업이익률']:.1%} (기준월 {prev['영업이익률']:.1%})",
    ]
    if not op_budget.empty and not pd.isna(op_budget.iloc[0]["차이금액"]):
        lines.append(f"- 예산 대비 영업이익 차이: {won(float(op_budget.iloc[0]['차이금액']))}")
    if len(neg):
        lines.append("- 주요 감소 요인(데이터 계산 기준):")
        for _, r in neg.iterrows():
            lines.append(f"  · {r['영역']} / {r['원인']}: {won(float(r['영업이익영향']))}")
    if len(pos):
        lines.append("- 주요 개선 요인(데이터 계산 기준):")
        for _, r in pos.iterrows():
            lines.append(f"  · {r['영역']} / {r['원인']}: +{abs(float(r['영업이익영향'])):,.0f}원")
    lines.append("- 주의: 위 내용은 입력 데이터와 산식으로 확인된 금액 영향이며, 실제 업무 원인은 관련 부서·전표·거래내역 확인 후 확정해야 합니다.")
    return "\n".join(lines)
