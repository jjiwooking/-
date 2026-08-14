import os, sys, math
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(__file__))
sys.path.insert(0, ROOT)
from analysis_engine import *


def load():
    base = os.path.join(ROOT, "sample_data", "demo", "csv")
    return normalize_data({
        "매출상세": pd.read_csv(os.path.join(base, "매출상세.csv")),
        "제조원가상세": pd.read_csv(os.path.join(base, "제조원가상세.csv")),
        "판관비상세": pd.read_csv(os.path.join(base, "판관비상세.csv")),
        "예산": pd.read_csv(os.path.join(base, "예산.csv")),
    })


def close(a,b,tol=0.001):
    assert abs(a-b) <= tol, (a,b)


def run():
    d=load(); p0=Period(2026,6); p1=Period(2026,7)
    assert validate_data(d)==[]
    a=pnl(d,p0); b=pnl(d,p1)
    close(a['매출액'],5_000_000_000); close(b['매출액'],4_850_000_000)
    close(a['매출원가'],3_200_000_000); close(b['매출원가'],3_330_000_000)
    close(a['판매비와관리비'],1_280_000_000); close(b['판매비와관리비'],1_120_000_000)
    close(a['영업이익'],520_000_000); close(b['영업이익'],400_000_000)
    close(b['영업이익']-a['영업이익'],-120_000_000)
    bridge=operating_profit_bridge(d,p0,p1)
    close(bridge.iloc[1:4]['영향금액'].sum(),-120_000_000)
    sales=sales_price_volume_analysis(d,p0,p1)
    close(sales['매출증감'].sum(),-150_000_000)
    close(sales['검산차이'].abs().sum(),0)
    c=cogs_change(d,p0,p1,['원가구분'])
    close(c['증감금액'].sum(),130_000_000)
    raw=c[c['원가구분']=='원재료비'].iloc[0]
    close(raw['증감금액'],130_000_000); close(raw['영업이익영향'],-130_000_000)
    s=sga_change(d,p0,p1,['계정과목'])
    close(s['증감금액'].sum(),-160_000_000)
    ship=s[s['계정과목']=='운반비'].iloc[0]
    close(ship['영업이익영향'],-60_000_000)
    adv=s[s['계정과목']=='광고선전비'].iloc[0]
    close(adv['영업이익영향'],150_000_000)
    bud=budget_comparison(d,p1)
    op=bud[bud['손익항목']=='영업이익'].iloc[0]
    close(op['차이금액'],-300_000_000)
    causes=combined_root_causes(d,p0,p1)
    close(causes['영업이익영향'].sum(),-120_000_000)
    report=management_report(d,p0,p1)
    assert '400,000,000원' in report and '-120,000,000원' in report
    assert '실제 업무 원인' in report
    print('PASS: 20 core checks')

if __name__=='__main__':
    run()
