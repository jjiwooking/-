from __future__ import annotations

import io
import json
import os
from copy import deepcopy
from typing import Dict

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from analysis_engine import (
    Period,
    REQUIRED_COLUMNS,
    available_periods,
    budget_comparison,
    cogs_change,
    combined_root_causes,
    management_report,
    normalize_data,
    operating_profit_bridge,
    pnl,
    pnl_comparison,
    sales_price_volume_analysis,
    sga_change,
    validate_data,
)

ROOT = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(ROOT, "sample_data", "demo", "csv")
SETTINGS_FILE = os.path.join(ROOT, "settings", "기본_화면설정.json")

st.set_page_config(page_title="월마감 손익 변동 원인분석", layout="wide")


def load_default_settings():
    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def init_state():
    if "settings" not in st.session_state:
        st.session_state.settings = load_default_settings()
    if "data" not in st.session_state:
        st.session_state.data = None
    if "menu_config" not in st.session_state:
        s = st.session_state.settings
        st.session_state.menu_config = pd.DataFrame([
            [1, True, s["메뉴1"], "자료"],
            [2, True, s["메뉴2"], "손익"],
            [3, True, s["메뉴3"], "원인"],
            [4, True, s["메뉴4"], "추적"],
            [5, True, s["메뉴5"], "보고"],
            [6, True, s["메뉴6"], "설정"],
        ], columns=["순서", "표시", "메뉴명", "메뉴키"])
    if "tool_config" not in st.session_state:
        st.session_state.tool_config = pd.DataFrame([
            [1, True, "영업이익 브리지", "bridge"],
            [2, True, "매출 판매량·판매단가 분석", "sales"],
            [3, True, "매출원가 변동 분석", "cogs"],
            [4, True, "판매비와관리비 변동 분석", "sga"],
            [5, True, "예산 대비 실적", "budget"],
        ], columns=["순서", "표시", "분석도구명", "도구키"])
    if "column_label_config" not in st.session_state:
        names = [
            "회계연도","회계월","매출일자","고객코드","고객명","제품코드","제품명","수량","판매단가","매출액",
            "원가구분","금액","부서코드","부서명","계정코드","계정과목","거래처코드","거래처명","손익항목","예산금액",
            "매출원가","매출총이익","판매비와관리비","영업이익","기준월금액","분석월금액","증감금액","증감률","영업이익영향",
            "실적금액","차이금액","달성률","구분","영향금액","유형","원인","영역","판매량효과","판매단가효과","신규매출효과","판매중단효과"
        ]
        st.session_state.column_label_config = pd.DataFrame([[x,x] for x in names], columns=["내부열이름","화면표시명"])


def col_label(name):
    cfg = st.session_state.get("column_label_config")
    if cfg is None or cfg.empty:
        return str(name)
    hit = cfg[cfg["내부열이름"].astype(str) == str(name)]
    return str(hit.iloc[0]["화면표시명"]) if len(hit) else str(name)


def generic_column_config(df):
    cfg = {}
    for col in df.columns:
        label = col_label(col)
        name = str(col)
        if col in ["증감률", "달성률"]:
            cfg[col] = st.column_config.NumberColumn(label, format="percent")
        elif "수량" in name:
            cfg[col] = st.column_config.NumberColumn(label, format="%,d", step=1)
        elif any(k in name for k in ["금액", "단가", "영업이익영향", "영향금액", "매출증감", "효과", "검산차이"]):
            cfg[col] = money_config(label)
        else:
            cfg[col] = st.column_config.Column(label)
    return cfg


def money_config(label=None):
    suffix = str(st.session_state.settings.get("금액단위", "원"))
    return st.column_config.NumberColumn(label, format=f"%,d{suffix}", step=1)


def percent_config(label=None):
    return st.column_config.NumberColumn(label, format="percent")


def won(v):
    suffix = str(st.session_state.settings.get("금액단위", "원"))
    if pd.isna(v):
        return "-"
    sign = "-" if float(v) < 0 else ""
    return f"{sign}{abs(float(v)):,.0f}{suffix}"


def load_sample():
    data = {
        "매출상세": pd.read_csv(os.path.join(SAMPLE_DIR, "매출상세.csv")),
        "제조원가상세": pd.read_csv(os.path.join(SAMPLE_DIR, "제조원가상세.csv")),
        "판관비상세": pd.read_csv(os.path.join(SAMPLE_DIR, "판관비상세.csv")),
        "예산": pd.read_csv(os.path.join(SAMPLE_DIR, "예산.csv")),
    }
    return normalize_data(data)


def to_excel_bytes(data: Dict[str, pd.DataFrame]):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        workbook = writer.book
        money_fmt = workbook.add_format({"num_format": '#,##0"원";[Red](#,##0"원");-'})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#17365D", "font_color": "white", "align": "center"})
        for name, df in data.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            for c, col in enumerate(df.columns):
                ws.write(0, c, col, header_fmt)
                width = max(10, min(24, max(len(str(col)) + 2, 12)))
                if any(k in str(col) for k in ["금액", "단가"]):
                    ws.set_column(c, c, 16, money_fmt)
                else:
                    ws.set_column(c, c, width)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns)-1, 0))
    return out.getvalue()


def standardize_from_mapping(raw_sheets: Dict[str, pd.DataFrame], sheet_map: Dict[str, str], col_maps: Dict[str, Dict[str, str]]):
    out = {}
    for logical, source_sheet in sheet_map.items():
        src = raw_sheets[source_sheet].copy()
        cmap = col_maps[logical]
        selected = {}
        for target in REQUIRED_COLUMNS[logical]:
            source = cmap.get(target)
            if source and source != "(없음)" and source in src.columns:
                selected[target] = src[source]
            else:
                selected[target] = pd.Series([None] * len(src))
        out[logical] = pd.DataFrame(selected)
        # Keep optional date fields when exact names exist.
        for optional in ["매출일자"]:
            if optional in src.columns and optional not in out[logical].columns:
                out[logical][optional] = src[optional]
    return normalize_data(out)


def load_uploaded_excel(uploaded):
    raw = pd.read_excel(uploaded, sheet_name=None, engine="openpyxl")
    return {str(k): v for k, v in raw.items()}


def data_editor_config(name, df):
    cfg = {}
    for col in df.columns:
        if any(k in str(col) for k in ["금액", "단가"]):
            cfg[col] = money_config(col_label(col))
        elif col == "수량":
            cfg[col] = st.column_config.NumberColumn(col_label(col), format="%,d", step=1)
        elif col in ["회계연도", "회계월"]:
            cfg[col] = st.column_config.NumberColumn(col_label(col), format="%d", step=1)
        else:
            cfg[col] = st.column_config.Column(col_label(col))
    return cfg


def choose_periods(data):
    periods = available_periods(data)
    if len(periods) < 2:
        st.error("비교 가능한 월이 2개 이상 필요합니다.")
        st.stop()
    labels = [p.label for p in periods]
    c1, c2 = st.columns(2)
    with c1:
        prior_label = st.selectbox(st.session_state.settings.get("기준월명", "기준월"), labels, index=max(len(labels)-2, 0))
    with c2:
        current_label = st.selectbox(st.session_state.settings.get("분석월명", "분석월"), labels, index=len(labels)-1)
    prior = periods[labels.index(prior_label)]
    current = periods[labels.index(current_label)]
    if (current.year, current.month) <= (prior.year, prior.month):
        st.warning("분석월은 기준월보다 뒤의 기간을 선택하는 것을 권장합니다.")
    return prior, current


def require_data():
    if st.session_state.data is None:
        st.info("먼저 '자료 넣기'에서 샘플 또는 회사 데이터를 불러오세요.")
        st.stop()
    issues = validate_data(st.session_state.data)
    if issues:
        st.error("자료 구조를 먼저 확인해야 합니다.\n\n" + "\n".join(f"- {x}" for x in issues))
        st.stop()
    return normalize_data(st.session_state.data)


def page_data():
    st.subheader("재무 자료 불러오기 및 직접 수정")
    st.caption("ERP/Excel에서 추출한 자료를 넣고, 화면 안에서 금액·계정·제품·고객·부서명을 직접 수정할 수 있습니다.")
    c1, c2 = st.columns([1, 2])
    with c1:
        if st.button("샘플 데이터 불러오기", use_container_width=True, type="primary"):
            st.session_state.data = load_sample()
            st.success("한빛정밀 샘플 데이터를 불러왔습니다.")
    with c2:
        uploaded = st.file_uploader("회사 Excel 업로드", type=["xlsx", "xlsm"])

    if uploaded is not None:
        try:
            raw = load_uploaded_excel(uploaded)
            st.markdown("#### 1) 시트 연결")
            sheet_names = list(raw.keys())
            sheet_map = {}
            for logical in REQUIRED_COLUMNS:
                default = sheet_names.index(logical) if logical in sheet_names else 0
                sheet_map[logical] = st.selectbox(f"{logical} ← Excel 시트", sheet_names, index=default, key=f"sheetmap_{logical}")
            st.markdown("#### 2) 열 이름 연결")
            col_maps = {}
            for logical, source_sheet in sheet_map.items():
                with st.expander(f"{logical} 열 매핑", expanded=False):
                    cols = ["(없음)"] + [str(c) for c in raw[source_sheet].columns]
                    col_maps[logical] = {}
                    for target in REQUIRED_COLUMNS[logical]:
                        idx = cols.index(target) if target in cols else 0
                        col_maps[logical][target] = st.selectbox(f"{target} ← 원본 열", cols, index=idx, key=f"colmap_{logical}_{target}")
            if st.button("업로드 자료 적용", type="primary"):
                st.session_state.data = standardize_from_mapping(raw, sheet_map, col_maps)
                st.success("업로드 자료를 프로그램 표준 구조로 연결했습니다.")
        except Exception as e:
            st.error(f"Excel 읽기 오류: {e}")

    if st.session_state.data is None:
        return

    st.divider()
    st.markdown("### 대시보드 안에서 원본자료 수정")
    edited = {}
    for name, df in st.session_state.data.items():
        with st.expander(f"{name} 직접 수정", expanded=(name == "매출상세")):
            edited[name] = st.data_editor(
                df,
                num_rows="dynamic",
                hide_index=True,
                width="stretch",
                column_config=data_editor_config(name, df),
                key=f"editor_{name}",
            )
    st.session_state.data = normalize_data(edited)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "수정한 전체 자료 Excel 다운로드",
            data=to_excel_bytes(st.session_state.data),
            file_name="수정_재무분석자료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        if st.button("샘플 원본으로 되돌리기", use_container_width=True):
            st.session_state.data = load_sample()
            st.rerun()


def page_pnl():
    data = require_data()
    prior, current = choose_periods(data)
    a = pnl(data, prior); b = pnl(data, current)
    st.markdown("### 핵심 손익")
    items = [
        (st.session_state.settings.get("매출명", "매출액"), "매출액"),
        (st.session_state.settings.get("원가명", "매출원가"), "매출원가"),
        (st.session_state.settings.get("판관비명", "판매비와관리비"), "판매비와관리비"),
        (st.session_state.settings.get("영업이익명", "영업이익"), "영업이익"),
    ]
    cols = st.columns(4)
    for col, (label, key) in zip(cols, items):
        delta = b[key] - a[key]
        with col:
            st.metric(label, won(b[key]), won(delta))
    st.caption(f"영업이익률: {b['영업이익률']:.1%} / 기준월 {a['영업이익률']:.1%}")

    comp = pnl_comparison(data, prior, current)
    st.markdown("### 월별 손익 비교")
    comp_cfg = generic_column_config(comp)
    comp_cfg[prior.label] = money_config(prior.label)
    comp_cfg[current.label] = money_config(current.label)
    st.dataframe(comp, hide_index=True, width="stretch", column_config=comp_cfg)

    st.markdown("### 예산 대비 실적")
    bud = budget_comparison(data, current)
    st.dataframe(bud, hide_index=True, width="stretch", column_config=generic_column_config(bud))
    st.info("계산 결과를 바꾸려면 '자료 넣기'에서 원본 금액을 수정하세요. 결과는 즉시 다시 계산됩니다.")


def page_causes():
    data = require_data()
    prior, current = choose_periods(data)
    st.markdown("### 통합 주요 영향 요인")
    causes_all = combined_root_causes(data, prior, current)
    min_amt = float(st.session_state.settings.get("최소중요금액", 1_000_000))
    causes_view = causes_all[causes_all["영업이익영향"].abs() >= min_amt].copy() if not causes_all.empty else causes_all
    if not causes_view.empty:
        top = causes_view.head(12).sort_values("영업이익영향")
        fig0 = go.Figure(go.Bar(x=top["영업이익영향"], y=top["영역"] + " / " + top["원인"], orientation="h"))
        fig0.update_layout(xaxis_tickformat=",.0f", margin=dict(t=10,b=20,l=20,r=20), height=420)
        st.plotly_chart(fig0, width="stretch")
        st.dataframe(causes_view, hide_index=True, width="stretch", column_config=generic_column_config(causes_view))
    tools = st.session_state.tool_config.copy()
    tools = tools[tools["표시"] == True].sort_values("순서")
    for _, tool in tools.iterrows():
        key = tool["도구키"]
        st.markdown(f"### {tool['분석도구명']}")
        if key == "bridge":
            bridge = operating_profit_bridge(data, prior, current)
            a = pnl(data, prior); b = pnl(data, current)
            rev_effect = float(bridge.loc[bridge['구분']=='매출액 증감 영향','영향금액'].iloc[0])
            cogs_effect = float(bridge.loc[bridge['구분']=='매출원가 증감 영향','영향금액'].iloc[0])
            sga_effect = float(bridge.loc[bridge['구분']=='판매비와관리비 증감 영향','영향금액'].iloc[0])
            fig = go.Figure(go.Waterfall(
                orientation="v", measure=["absolute","relative","relative","relative","total"],
                x=[prior.label, "매출액", "매출원가", "판매비와관리비", current.label],
                y=[a['영업이익'], rev_effect, cogs_effect, sga_effect, b['영업이익']],
                text=[won(a['영업이익']), won(rev_effect), won(cogs_effect), won(sga_effect), won(b['영업이익'])],
                textposition="outside", connector={"line":{"color":"gray"}}
            ))
            fig.update_layout(showlegend=False, yaxis_tickformat=",.0f", margin=dict(t=20,b=20,l=20,r=20), height=420)
            st.plotly_chart(fig, width="stretch")
            st.dataframe(bridge, hide_index=True, width="stretch", column_config=generic_column_config(bridge))
        elif key == "sales":
            sales = sales_price_volume_analysis(data, prior, current)
            summary = pd.DataFrame({
                "원인": ["판매량효과","판매단가효과","신규매출효과","판매중단효과"],
                "영업이익영향": [sales[x].sum() for x in ["판매량효과","판매단가효과","신규매출효과","판매중단효과"]]
            })
            st.dataframe(summary, hide_index=True, width="stretch", column_config=generic_column_config(summary))
            with st.expander("제품·고객별 매출 변동 상세", expanded=True):
                show_cols = ["제품명","고객명","수량_기준","수량_분석","평균판매단가_기준","평균판매단가_분석","판매량효과","판매단가효과","매출증감","검산차이"]
                st.dataframe(sales[show_cols], hide_index=True, width="stretch", column_config=generic_column_config(sales[show_cols]))
            st.caption("판매량효과와 판매단가효과는 동일 제품·고객 조합의 기준월/분석월 수량과 평균판매단가로 계산합니다. 신규/판매중단 거래는 별도로 분리합니다.")
        elif key == "cogs":
            dims = st.multiselect("원가 분석 기준", ["원가구분","제품명"], default=["원가구분"], key="cogs_dims")
            x = cogs_change(data, prior, current, dims or ["원가구분"])
            st.dataframe(x, hide_index=True, width="stretch", column_config=generic_column_config(x))
        elif key == "sga":
            dims = st.multiselect("판관비 분석 기준", ["계정과목","부서명","거래처명"], default=["계정과목"], key="sga_dims")
            x = sga_change(data, prior, current, dims or ["계정과목"])
            st.dataframe(x, hide_index=True, width="stretch", column_config=generic_column_config(x))
        elif key == "budget":
            x = budget_comparison(data, current)
            st.dataframe(x, hide_index=True, width="stretch", column_config=generic_column_config(x))
        st.divider()


def page_trace():
    data = require_data()
    prior, current = choose_periods(data)
    st.markdown("### 숫자 → 세부 항목 → 원본 행 추적")
    causes = combined_root_causes(data, prior, current)
    if not causes.empty:
        st.dataframe(causes.head(12), hide_index=True, width="stretch", column_config=generic_column_config(causes.head(12)))
    area = st.radio("확인할 영역", ["매출", "매출원가", "판매비와관리비"], horizontal=True)
    if area == "매출":
        x = sales_price_volume_analysis(data, prior, current)
        product = st.selectbox("제품", sorted(data["매출상세"]["제품명"].dropna().astype(str).unique()))
        customer_opts = ["전체"] + sorted(data["매출상세"]["고객명"].dropna().astype(str).unique())
        customer = st.selectbox("고객", customer_opts)
        detail = data["매출상세"].copy()
        detail = detail[detail["제품명"].astype(str)==product]
        if customer != "전체": detail = detail[detail["고객명"].astype(str)==customer]
        detail = detail[((detail["회계연도"]==prior.year)&(detail["회계월"]==prior.month)) | ((detail["회계연도"]==current.year)&(detail["회계월"]==current.month))]
        st.dataframe(detail, hide_index=True, width="stretch", column_config=generic_column_config(detail))
    elif area == "매출원가":
        product = st.selectbox("제품", sorted(data["제조원가상세"]["제품명"].dropna().astype(str).unique()))
        cost_type = st.selectbox("원가구분", ["전체"] + sorted(data["제조원가상세"]["원가구분"].dropna().astype(str).unique()))
        detail = data["제조원가상세"].copy()
        detail = detail[detail["제품명"].astype(str)==product]
        if cost_type != "전체": detail = detail[detail["원가구분"].astype(str)==cost_type]
        detail = detail[((detail["회계연도"]==prior.year)&(detail["회계월"]==prior.month)) | ((detail["회계연도"]==current.year)&(detail["회계월"]==current.month))]
        st.dataframe(detail, hide_index=True, width="stretch", column_config=generic_column_config(detail))
    else:
        account = st.selectbox("계정과목", sorted(data["판관비상세"]["계정과목"].dropna().astype(str).unique()))
        department = st.selectbox("부서", ["전체"] + sorted(data["판관비상세"]["부서명"].dropna().astype(str).unique()))
        detail = data["판관비상세"].copy()
        detail = detail[detail["계정과목"].astype(str)==account]
        if department != "전체": detail = detail[detail["부서명"].astype(str)==department]
        detail = detail[((detail["회계연도"]==prior.year)&(detail["회계월"]==prior.month)) | ((detail["회계연도"]==current.year)&(detail["회계월"]==current.month))]
        st.dataframe(detail, hide_index=True, width="stretch", column_config=generic_column_config(detail))
    st.caption("현재 버전은 업로드된 상세자료의 원본 행까지 추적합니다. ERP 전표번호가 있는 파일을 연결하면 다음 버전에서 전표번호까지 바로 내려갈 수 있습니다.")


def page_report():
    data = require_data()
    prior, current = choose_periods(data)
    min_amt = float(st.session_state.settings.get("최소중요금액", 1_000_000))
    base = management_report(data, prior, current, min_amt)
    key = f"report_{prior.year}{prior.month}_{current.year}{current.month}"
    if key not in st.session_state:
        st.session_state[key] = base
    st.markdown("### 경영보고 초안")
    st.caption("문구는 직접 수정할 수 있습니다. 자동 작성 내용은 계산된 숫자만 사용하고, 확인되지 않은 실제 원인을 단정하지 않습니다.")
    st.session_state[key] = st.text_area("보고서 직접 수정", st.session_state[key], height=360)
    st.download_button("경영보고 TXT 다운로드", st.session_state[key].encode("utf-8-sig"), file_name=f"{current.year}_{current.month:02d}_월마감_손익보고.txt", mime="text/plain")

    causes = combined_root_causes(data, prior, current)
    if not causes.empty:
        st.markdown("### 보고서 근거표")
        st.dataframe(causes, hide_index=True, width="stretch", column_config=generic_column_config(causes))


def page_settings():
    st.markdown("### 화면/용어를 코딩 없이 수정")
    st.caption("제목, 회사명, 메뉴명, 재무 용어, 금액 단위, 분석도구명·표시여부·순서를 수정할 수 있습니다.")

    s = st.session_state.settings
    rows = [[k, str(v)] for k, v in s.items()]
    edited = st.data_editor(pd.DataFrame(rows, columns=["설정항목","값"]), hide_index=True, num_rows="fixed", width="stretch", key="settings_editor")
    if st.button("화면 설정 적용", type="primary"):
        new = {}
        for _, r in edited.iterrows():
            k, v = str(r["설정항목"]), str(r["값"])
            if k == "최소중요금액":
                try: v = float(v.replace(",",""))
                except: v = 1_000_000
            elif k == "기본증감률경보기준":
                try: v = float(v)
                except: v = 0.2
            new[k] = v
        st.session_state.settings = new
        st.success("화면 설정을 적용했습니다.")

    st.markdown("### 메뉴 설정")
    st.session_state.menu_config = st.data_editor(st.session_state.menu_config, hide_index=True, num_rows="fixed", width="stretch", key="menu_editor")
    st.markdown("### 분석 툴 설정")
    st.session_state.tool_config = st.data_editor(st.session_state.tool_config, hide_index=True, num_rows="fixed", width="stretch", key="tool_editor")
    st.markdown("### 표의 열 제목 설정")
    st.caption("내부 계산용 열 이름은 유지하고, 화면에 보이는 제목만 자유롭게 바꿉니다.")
    st.session_state.column_label_config = st.data_editor(st.session_state.column_label_config, hide_index=True, num_rows="dynamic", width="stretch", key="column_label_editor")

    payload = {
        "settings": st.session_state.settings,
        "menu_config": st.session_state.menu_config.to_dict(orient="records"),
        "tool_config": st.session_state.tool_config.to_dict(orient="records"),
        "column_label_config": st.session_state.column_label_config.to_dict(orient="records"),
    }
    st.download_button("현재 전체 설정 JSON 다운로드", json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"), file_name="월마감_손익분석_화면설정.json", mime="application/json")
    uploaded = st.file_uploader("이전에 저장한 설정 JSON 불러오기", type=["json"], key="setting_upload")
    if uploaded is not None and st.button("설정 JSON 적용"):
        try:
            obj = json.load(uploaded)
            st.session_state.settings = obj["settings"]
            st.session_state.menu_config = pd.DataFrame(obj["menu_config"])
            st.session_state.tool_config = pd.DataFrame(obj["tool_config"])
            if "column_label_config" in obj:
                st.session_state.column_label_config = pd.DataFrame(obj["column_label_config"])
            st.success("설정을 불러왔습니다.")
            st.rerun()
        except Exception as e:
            st.error(f"설정 파일 오류: {e}")


init_state()
s = st.session_state.settings
st.title(str(s.get("앱제목", "월마감 손익 변동 원인분석 시스템")))
st.caption(f"{s.get('회사명','')} · {s.get('앱설명','')}")

menu = st.session_state.menu_config.copy()
menu = menu[menu["표시"] == True].sort_values("순서")
if menu.empty:
    st.error("표시할 메뉴가 없습니다. 화면 설정 파일을 초기화하세요.")
    st.stop()
options = menu["메뉴명"].tolist()
selected = st.sidebar.radio("메뉴", options)
menu_key = menu.loc[menu["메뉴명"] == selected, "메뉴키"].iloc[0]

if menu_key == "자료": page_data()
elif menu_key == "손익": page_pnl()
elif menu_key == "원인": page_causes()
elif menu_key == "추적": page_trace()
elif menu_key == "보고": page_report()
elif menu_key == "설정": page_settings()
