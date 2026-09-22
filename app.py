# -*- coding: utf-8 -*-
"""
신규공장 전력인프라 DTR/PSS-E 사전진단 시스템 V22
================================================
이 버전은 사용자가 작성한 'ㅇㅇㅇㅇ 시스템' 문서의 실행 순서를 그대로 따른다.

화면 순서
1) 맨 위: 1년 외기온도 / 1년 DTR-vs-기존 MAIN_TR 두 그래프가 재생되며 이동
2) PSS/E 계통도 원본 + 선택일 일일 외기온도
3) 판단지표 4개: Feeder / 전압 / MVA / PV
4) 대표 그래프 4개: Feeder / MVA / PV / 전압 보강단계
5) 추가 Excel 그래프: 모든 Bus / 0.95 미만 Bus / 용량별 MVA·Feeder 비교
6) 사용자 외기온도 / 기존공장 부하패턴 업로드
7) 실제 변압기 모델을 받을 때 DTR 계수 수정

그래프 스타일은 사용자가 HWPX/Excel에 넣은 그래프의 색상, 축, 선, 기준선을 기준으로 고정한다.
임의의 색상/면 채우기/누적영역 그래프를 사용하지 않는다.

현재 PSS/E 원자료:
- 신규공장 0 / 10 / 15 / 20 / 30 MVA
- 30 MVA 보상: +23.1 MVA 병렬 TR, Bus 6 +3.5 MVAr
임의 용량/복수공장은 위 PSS/E 원자료 사이를 보간한 추정값이다.
"""

from __future__ import annotations

import base64
import io
import math
import os
import re
import sys
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parent

# ------------------------------ 기준 ------------------------------
V_LIMIT = 0.950
FEEDER_RATING = 35.0
DEFAULT_TR_RATING = 105.0
DEFAULT_PLAN_EXISTING = 90.0
DEFAULT_RESERVE = 0.10

CASE_POINTS = [0.0, 10.0, 15.0, 20.0, 30.0]
BUS_IDS = list(range(5, 19))
BUS_NAMES = {
    5:"EX_FDR_IN", 6:"EXIST_22.9", 7:"EX_CW_6.6", 8:"EX_DA_6.6",
    9:"EX_TEST_6.6", 10:"EX_UTIL_0.48", 11:"EX_LV_0.38",
    12:"NEW_FDR_IN", 13:"EXP_22.9", 14:"NEW_CW_6.6", 15:"NEW_DA_6.6",
    16:"NEW_TEST_6.6", 17:"NEW_UTIL_0.4", 18:"NEW_LV_0.38",
}

# HWPX 원본 계통도를 PY에 직접 내장
PSS_IMAGE_URI = "data:image/png;base64,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/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUAEUQagCgW0INQD8INUARhBoA6JZQA9APQg1QBKEGALol1AD0g1ADFEGoAYBuCTUA/SDUNHjE856Xjj372YZhrHHc8Ywz0hnPfOa+ecM41caZDXOLDLr1oY99bN8yN4y+jR88//z0oIsv3jdvGIaxzhH728wn1DS447Fj9U7j6ceOGYZhGEYvRhyV9q4PfjB/SaMl51x22b5lbhiGYRjG4cb9HvWo/KWWKUJNg7jjRKi505lnGsaRGbET9wPnnbexcfdzztke556777J1jXuce+6+5WIYR2UINd2ahJp8uRvGqT42vX0xGfEaP9nWyC/bxLDNYRizx13OPFOoOYBQ0yAizX/5vd+r/2sYhmEYfRhCTbci1OTL3DAMwzCM5cf3O6LmQEJNg/ywrIWGj0oZhmEYy44WXzvoXr7MOxkt3ieMNQzryzBWGx5Dp+xgPqGmBf/8iU+kH7rwwnyaDfvtpz89veId78in2aCXvu1t6WHPelY+zRq86l3vSv/naU/Lp9mwH3nEI9I/ffzj+TSnsP/+mMekD3z0o/k0BYr19N8e85h8mg371Sc/Ob35fe/LpynQv33q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"

# ------------------------------ Excel 원본 색 ------------------------------
# 연간 온도 Excel
ANNUAL_TEMP = "#4BACC6"     # accent5
ANNUAL_DTR = "#F79646"      # accent6
BASE_MAIN = "#0B6FA4"
MAX_POINT = "#C0392B"
REF_RED = "#D1495B"

# 보상최종 Excel theme / 직접 지정색
ACCENT4 = "#0F9ED5"         # feeder, MAIN_TR
ACCENT6 = "#4EA72E"         # 보강 후 기준
ORANGE = "#F47A34"
GREEN = "#1B8A5A"
RED_DARK = "#C00000"
GRAY = "#64748B"

BUS_COLORS = {
    5:"#0B6FA4", 6:"#F47A34", 7:"#1B8A5A", 8:"#22A6E8",
    9:"#8B5CF6", 10:"#C2410C", 11:"#5B8E3E", 12:"#E11D48",
    13:"#7C3AED", 14:"#0891B2", 15:"#A16207", 16:"#BE185D",
    17:"#4D7C0F", 18:"#475569",
}
PV_COLORS = {
    "case1":"#D65AD7",
    "case2":"#F47A34",
    "case3":"#1B8A5A",
    "case4":"#22A6E8",
}

GRID = "#D9EAF7"

# PV 원자료 — Bus4_PV_4Case_세로축0.80.xlsx / PDF추출_원본점
# 각 Case는 실제 끝점이 서로 다르므로, 원본 4개 곡선은 각자의 X/Y를 그대로 그린다.
PV_CASE1_X = np.array([0,9.939,19.999,29.939,39.999,49.939,59.999,69.939,79.999,89.939,100,109.939,119.999,129.939,139.999,149.939,160,169.939,179.999],dtype=float)
PV_CASE1_Y = np.array([.975477,.97,.964286,.958333,.951668,.944762,.937619,.929762,.921429,.912619,.903095,.892619,.881429,.869048,.855238,.839762,.821905,.798333,.769524],dtype=float)

PV_CASE2_X = np.array([0,9.917,19.943,29.969,39.995,49.912,59.938,69.964,79.99,89.907,99.933,109.959,119.985,129.902,139.928,141.235,142.434,143.742,144.941,146.249,147.447,148.646,149.954],dtype=float)
PV_CASE2_Y = np.array([.955131,.947178,.938278,.929,.919153,.908359,.896998,.8845,.870866,.855907,.839054,.819739,.793986,.762741,.729035,.718242,.718431,.708963,.710857,.702336,.698548,.693057,.692868],dtype=float)

PV_CASE3_X = np.array([0,9.917,19.943,29.969,39.995,49.912,59.938,69.964,79.99,89.907,99.933,109.959,119.985,129.902,139.928,149.954,159.98,164.993,167.499,169.897,172.403,174.91],dtype=float)
PV_CASE3_Y = np.array([.966321,.959829,.95293,.945491,.937782,.929532,.920874,.911542,.901668,.890983,.879351,.866636,.85257,.83688,.81862,.794409,.765193,.74761,.740847,.732056,.724751,.722993],dtype=float)

PV_CASE4_X = np.array([0,9.939,19.999,29.939,39.999,49.939,59.999,69.939,79.999,89.939,100,109.939,119.999,129.939,139.999,149.939,160,169.939,171.151,172.484,173.09,173.696,174.303,174.909],dtype=float)
PV_CASE4_Y = np.array([.972286,.965714,.958857,.951429,.943714,.935714,.926857,.917715,.908,.897143,.885714,.873429,.859429,.844,.826286,.803144,.774858,.747429,.736571,.737143,.732573,.734,.728,.731715],dtype=float)

# 현재 시나리오를 공식화할 때는 네 Case가 모두 실제 값을 갖는 0~149 MW 공통구간만 사용.
# 따라서 현재 시나리오 곡선도 중간에 끊기지 않는다.
PV_MODEL_X = np.linspace(0.0,149.0,150)

def _pv_on_grid(x,y,grid=PV_MODEL_X):
    return np.interp(grid,np.asarray(x,float),np.asarray(y,float))

PV1_MODEL = _pv_on_grid(PV_CASE1_X,PV_CASE1_Y)
PV2_MODEL = _pv_on_grid(PV_CASE2_X,PV_CASE2_Y)
PV3_MODEL = _pv_on_grid(PV_CASE3_X,PV_CASE3_Y)
PV4_MODEL = _pv_on_grid(PV_CASE4_X,PV_CASE4_Y)


# ------------------------------ 파일 찾기 ------------------------------
def norm(s):
    return re.sub(r"[\s_\-()%\[\]]+", "", str(s)).lower()

def find_file(preferred_exact=None, include=(), exclude=()):
    if preferred_exact:
        p = ROOT / preferred_exact
        if p.exists():
            return p
    hits = []
    for p in ROOT.glob("*.xlsx"):
        n = norm(p.name)
        if any(norm(x) in n for x in exclude):
            continue
        score = sum(norm(k) in n for k in include)
        if score:
            hits.append((score, len(p.name), p))
    if not hits:
        return None
    hits.sort(key=lambda x:(-x[0], x[1]))
    return hits[0][2]

TEMP_FILE = find_file("온도최종.xlsx", ["온도최종"])
MVA_FILE = find_file(
    "MVA최종_DTR재산정_120C_IEC60K (1)(1).xlsx",
    ["mva최종","dtr재산정","120c"],
    ["보상최종"]
)
VOLT_FILE = find_file(
    "전압 분단위 최종(1).xlsx",
    ["전압","분단위","최종"],
    ["보상최종"]
)
def find_compensation_file():
    # 같은 보상최종 파일이 (1), (2)처럼 여러 개 있으면 최신 파일 우선
    candidates=[]
    required={"요약","계산근거","전압_보강전","전압_TR보강후","전압_Shunt후","전압_단계비교"}
    for p in ROOT.glob("*.xlsx"):
        try:
            names=set(pd.ExcelFile(p,engine="openpyxl").sheet_names)
            if required.issubset(names):
                score=0
                n=norm(p.name)
                for k in ["보상최종","기존공장","신규30mva","보강전후"]:
                    if norm(k) in n:
                        score+=1
                candidates.append((score,p.stat().st_mtime,p))
        except Exception:
            pass

    if candidates:
        candidates.sort(key=lambda x:(x[0],x[1]),reverse=True)
        return candidates[0][2]

    return find_file(
        "보상최종_기존공장_신규30MVA_보강전후_통합그래프.xlsx",
        ["보상최종","기존공장","신규30mva","보강전후"]
    )

COMP_FILE = find_compensation_file()


# ------------------------------ helpers ------------------------------
def numeric(s):
    return pd.to_numeric(s, errors="coerce")

def find_col(df, includes, excludes=()):
    for c in df.columns:
        n = norm(c)
        if all(norm(x) in n for x in includes) and not any(norm(x) in n for x in excludes):
            return c
    return None

def to_datetime_series(s):
    x = pd.to_datetime(s, errors="coerce")
    if x.notna().sum() >= max(1, len(s)//2):
        return x
    v = numeric(s)
    return pd.to_datetime("1899-12-30") + pd.to_timedelta(v, unit="D")

def resample(v, n=1440):
    a = numeric(pd.Series(v)).interpolate(limit_direction="both").to_numpy(float)
    if len(a) == n:
        return a
    if len(a) == 0:
        return np.full(n, np.nan)
    return np.interp(np.linspace(0,1,n), np.linspace(0,1,len(a)), a)

def case_interp(case_map, capacity):
    """0/10/15/20/30 case 선형 보간, 30 초과는 20~30 기울기 외삽."""
    c = max(0.0, float(capacity))
    pts = sorted(case_map.keys())
    if c <= pts[0]:
        return np.asarray(case_map[pts[0]],float).copy()
    if c >= pts[-1]:
        lo, hi = pts[-2], pts[-1]
    else:
        lo, hi = pts[0], pts[-1]
        for a,b in zip(pts[:-1],pts[1:]):
            if a <= c <= b:
                lo, hi = a,b
                break
    w = (c-lo)/(hi-lo) if hi != lo else 0
    return np.asarray(case_map[lo],float)*(1-w)+np.asarray(case_map[hi],float)*w

def ceil_tenth(v):
    return math.ceil(max(0.0,float(v))*10.0)/10.0

def time_labels(n=1440, step=1):
    return [f"{(i*step)//60:02d}:{(i*step)%60:02d}" for i in range(n)]

MINUTE_X = [f"{i//60:02d}:{i%60:02d}" for i in range(1440)]
TWO_MIN_X = [f"{(2*i)//60:02d}:{(2*i)%60:02d}" for i in range(720)]


# ------------------------------ DTR ------------------------------
def solve_k(temp, hotspot=120.0, oil_rise=60.0, winding_rise=30.0,
            loss_num=5.0, loss_den=6.0, oil_exp=.8, winding_exp=1.6):
    t=np.asarray(temp,float)
    lo=np.zeros_like(t)
    hi=np.full_like(t,3.0)
    for _ in range(50):
        k=(lo+hi)/2
        theta=t + oil_rise*((loss_num*k*k+1.0)/loss_den)**oil_exp + winding_rise*(k**winding_exp)
        ok=theta <= hotspot
        lo=np.where(ok,k,lo)
        hi=np.where(ok,hi,k)
    return lo

def calc_dtr(temp, rating, hotspot, oil_rise, winding_rise, loss_num, loss_den, oil_exp, winding_exp):
    return float(rating)*solve_k(temp,hotspot,oil_rise,winding_rise,loss_num,loss_den,oil_exp,winding_exp)


# ------------------------------ source load ------------------------------
def load_sources():
    missing=[]
    for p,label in [(TEMP_FILE,"온도최종.xlsx"),(MVA_FILE,"MVA최종...xlsx"),(VOLT_FILE,"전압 분단위 최종...xlsx"),(COMP_FILE,"보상최종...xlsx")]:
        if p is None: missing.append(label)
    if missing:
        raise FileNotFoundError("필수 파일 누락: "+", ".join(missing))

    # 연간
    annual_raw=pd.read_excel(TEMP_FILE,sheet_name="연간_데이터",engine="openpyxl")
    base_col=find_col(annual_raw,["기존공장","maintr"])
    annual=pd.DataFrame({
        "dt":to_datetime_series(annual_raw["일시"]),
        "temp":numeric(annual_raw["시간별 최고 외기온도(°C)"]),
        "base_main":numeric(annual_raw[base_col])
    }).dropna().sort_values("dt").reset_index(drop=True)
    if len(annual)>8760: annual=annual.iloc[-8760:].reset_index(drop=True)

    extreme_raw=pd.read_excel(TEMP_FILE,sheet_name="극한일_분단위",engine="openpyxl")
    extreme=pd.DataFrame({
        "dt":to_datetime_series(extreme_raw["일시"]),
        "temp":numeric(extreme_raw["울산 외기온도(°C)"]),
        "base_main":numeric(extreme_raw[find_col(extreme_raw,["기존","maintr"])])
    }).dropna().reset_index(drop=True)
    if len(extreme)!=1440:
        extreme=pd.DataFrame({
            "dt":pd.date_range(pd.Timestamp(extreme["dt"].iloc[0]).normalize(),periods=1440,freq="min"),
            "temp":resample(extreme["temp"],1440),
            "base_main":resample(extreme["base_main"],1440),
        })

    # raw MVA / feeder PSS/E cases
    md=pd.read_excel(MVA_FILE,sheet_name="분단위_데이터",engine="openpyxl")
    mva_cases={0.0:resample(extreme["base_main"],1440)}
    feeder_cases={0.0:np.zeros(1440)}
    for c in [10,15,20,30]:
        mva_cases[float(c)]=resample(md[f"{c}MVA MAIN_TR(MVA)"],1440)
        feeder_cases[float(c)]=resample(md[f"{c}MVA Feeder(MVA)"],1440)

    # voltage cases
    vd=pd.read_excel(VOLT_FILE,sheet_name="그래프데이터",engine="openpyxl")
    voltage={}
    for c in [0,10,15,20,30]:
        voltage[float(c)]={}
        for b in BUS_IDS:
            col=None
            for cc in vd.columns:
                if str(cc).startswith(f"{c}MVA Bus {b} ") and str(cc).endswith("V(pu)"):
                    col=cc; break
            if col is None:
                for cc in vd.columns:
                    n=norm(cc)
                    if f"{c}mvabus{b}" in n and "vpu" in n:
                        col=cc; break
            if col is None:
                raise ValueError(f"전압 원자료 누락: {c}MVA Bus {b}")
            voltage[float(c)][b]=resample(vd[col],1440)

    # compensation exact 30MVA
    comp={}
    for key,sheet in {"before":"전압_보강전","tr":"전압_TR보강후","shunt":"전압_Shunt후"}.items():
        d=pd.read_excel(COMP_FILE,sheet_name=sheet,header=2,engine="openpyxl")
        comp[key]={}
        for b in BUS_IDS:
            col=next((x for x in d.columns if norm(x)==norm(f"Bus {b}")),None)
            if col is None:
                raise ValueError(f"{sheet} Bus {b} 누락")
            comp[key][b]=resample(d[col],1440)

    dg=pd.read_excel(COMP_FILE,sheet_name="DTR_그래프",header=2,engine="openpyxl")
    bank30_col=find_col(dg,["신규공장","30mva","maintr"]) or find_col(dg,["30mva","maintr"])
    feeder30_col=find_col(dg,["30mva","feeder"])
    final30_bank=resample(dg[bank30_col],1440) if bank30_col is not None else mva_cases[30.0].copy()
    final30_feeder=resample(dg[feeder30_col],1440) if feeder30_col is not None else feeder_cases[30.0].copy()

    # ----------------------------------------------------------
    # 사용자가 준 '계산근거'와 '전압_단계비교'에서 기준값을 직접 읽는다.
    # 숫자를 코드에서 임의로 정하지 않는다.
    # ----------------------------------------------------------
    summary=pd.read_excel(COMP_FILE,sheet_name="요약",header=None,engine="openpyxl")
    calc=pd.read_excel(COMP_FILE,sheet_name="계산근거",header=None,engine="openpyxl")
    stage=pd.read_excel(COMP_FILE,sheet_name="전압_단계비교",engine="openpyxl")

    tr_ref=None
    shunt_ref=None
    plan_existing=None
    plan_dtr=None
    reserve=None
    tr_formula_text=None

    for r in range(len(summary)):
        key=str(summary.iloc[r,0]) if pd.notna(summary.iloc[r,0]) else ""
        val=summary.iloc[r,1] if summary.shape[1]>1 else None
        try:
            fval=float(val)
        except Exception:
            fval=None

        if "여유 증설량" in key and fval is not None:
            tr_ref=fval
        elif "기존공장 부하" in key and fval is not None:
            plan_existing=fval
        elif "최악 DTR 기준" in key and fval is not None:
            plan_dtr=fval
        elif key.strip()=="여유율" and fval is not None:
            reserve=fval

    for r in range(len(calc)):
        key=str(calc.iloc[r,0]) if pd.notna(calc.iloc[r,0]) else ""
        if "TR 증설" in key and calc.shape[1]>1 and pd.notna(calc.iloc[r,1]):
            tr_formula_text=str(calc.iloc[r,1])
            break

    # 열 이름: '23.1 MVA TR 후 최저', 'Bus6 3.5MVAr 후 최저'
    for c in stage.columns:
        s=str(c)
        m=re.search(r"([0-9]+(?:\.[0-9]+)?)\s*MVA\s*TR",s,re.I)
        if m:
            tr_ref=float(m.group(1))
        m=re.search(r"Bus\s*6\s*([0-9]+(?:\.[0-9]+)?)\s*MVAr",s,re.I)
        if m:
            shunt_ref=float(m.group(1))

    if tr_ref is None:
        raise ValueError("계산근거 Excel에서 병렬 TR 증설량을 찾지 못했습니다.")
    if shunt_ref is None:
        raise ValueError("계산근거 Excel에서 Bus6 Shunt 보상량을 찾지 못했습니다.")
    if plan_existing is None:
        raise ValueError("계산근거 Excel에서 기존공장 계획부하를 찾지 못했습니다.")
    if plan_dtr is None:
        raise ValueError("계산근거 Excel에서 계획 DTR을 찾지 못했습니다.")
    if reserve is None:
        raise ValueError("계산근거 Excel에서 여유율을 찾지 못했습니다.")

    # 30MVA 원본 전압 3단계 최저값도 저장
    exact_stage_min={}
    for key,col_hint in [
        ("before","보강 전 최저"),
        ("tr","TR 후 최저"),
        ("cap","MVAr 후 최저"),
    ]:
        col=next((c for c in stage.columns if col_hint.lower() in str(c).lower()),None)
        if col is not None:
            exact_stage_min[key]=float(pd.to_numeric(stage[col],errors="coerce").min())

    return {
        "annual":annual, "extreme":extreme,
        "mva_cases":mva_cases, "feeder_cases":feeder_cases,
        "voltage":voltage, "comp":comp,
        "final30_bank":final30_bank, "final30_feeder":final30_feeder,
        "tr_ref":tr_ref, "shunt_ref":shunt_ref,
        "plan_existing":plan_existing, "plan_dtr_source":plan_dtr, "reserve":reserve,
        "tr_formula_text":tr_formula_text,
        "exact_stage_min":exact_stage_min,
        "comp_source_name":COMP_FILE.name,
    }

DATA=None
LOAD_ERROR=None
try:
    DATA=load_sources()
except Exception as e:
    LOAD_ERROR=str(e)


# ------------------------------ upload parser ------------------------------
def read_tabular(raw, filename):
    fn=filename.lower()
    if fn.endswith(".xlsx") or fn.endswith(".xlsm"):
        book=pd.ExcelFile(io.BytesIO(raw),engine="openpyxl")
        for sh in book.sheet_names:
            yield sh,pd.read_excel(book,sheet_name=sh)
    elif fn.endswith(".csv"):
        for enc in ["cp949","euc-kr","utf-8-sig","utf-8"]:
            try:
                yield "csv",pd.read_csv(io.StringIO(raw.decode(enc)))
                return
            except Exception:
                pass


def parse_temperature(raw,filename):
    """
    외기온도 업로드.
    - Excel/CSV: 일시 + 기온/외기온도 열
    - 기상청 ZIP: 월별 중첩 ZIP 안의 CSV까지 탐색
    최종적으로 시간별 최고 외기온도 series를 반환한다.
    """
    frames=[]

    def decode_csv(blob):
        for enc in ["cp949","euc-kr","utf-8-sig","utf-8"]:
            try:
                return pd.read_csv(io.StringIO(blob.decode(enc)))
            except Exception:
                pass
        return None

    def add_df(d):
        if d is None or len(d)==0:
            return False
        dtc=(find_col(d,["일시"]) or find_col(d,["datetime"]) or
             find_col(d,["date"]) or find_col(d,["날짜"]))
        tc=(find_col(d,["시간별","최고","외기온도"]) or
            find_col(d,["울산","외기온도"]) or
            find_col(d,["외기온도"]) or
            find_col(d,["기온"]) or
            find_col(d,["temperature"]))
        if dtc is None or tc is None:
            return False
        z=pd.DataFrame({"dt":to_datetime_series(d[dtc]),"temp":numeric(d[tc])}).dropna()
        if len(z)==0:
            return False
        z["dt"]=z["dt"].dt.floor("h")
        z=z.groupby("dt",as_index=False)["temp"].max()
        frames.append(z)
        return True

    fn=filename.lower()
    if fn.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(raw)) as outer:
            for n in outer.namelist():
                if n.endswith("/"):
                    continue
                blob=outer.read(n)
                if n.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(io.BytesIO(blob)) as inner:
                            for nn in inner.namelist():
                                if nn.lower().endswith(".csv"):
                                    add_df(decode_csv(inner.read(nn)))
                    except Exception:
                        pass
                elif n.lower().endswith(".csv"):
                    add_df(decode_csv(blob))
    else:
        for _,d in read_tabular(raw,filename):
            add_df(d)

    if not frames:
        raise ValueError("일시 + 기온(또는 외기온도) 열을 찾지 못했습니다.")

    x=pd.concat(frames,ignore_index=True).dropna().sort_values("dt")
    x=x.drop_duplicates("dt",keep="last")
    x=x.set_index("dt").resample("h").max().interpolate(limit_direction="both").reset_index()

    return {
        "dt":[d.isoformat() for d in x["dt"]],
        "temp":x["temp"].astype(float).tolist(),
        "summary":{
            "start":x["dt"].iloc[0].isoformat(),
            "end":x["dt"].iloc[-1].isoformat(),
            "min":float(x["temp"].min()),
            "max":float(x["temp"].max()),
            "rows":int(len(x)),
        }
    }


def parse_load(raw,filename):
    """
    기존공장 부하패턴 업로드.
    지원:
    1) 날짜/일시 + MAIN_TR/MVA/부하
    2) 00:00~23:59 시간 + MAIN_TR/MVA/부하
    """
    for sheet,d in read_tabular(raw,filename):
        if d is None or len(d)==0:
            continue

        lc=(find_col(d,["기존공장","maintr"]) or
            find_col(d,["기존","maintr"]) or
            find_col(d,["maintr"]) or
            find_col(d,["부하","mva"]) or
            find_col(d,["부하"]) or
            find_col(d,["mva"]))
        if lc is None:
            continue

        dtc=(find_col(d,["일시"]) or find_col(d,["datetime"]) or
             find_col(d,["date"]) or find_col(d,["날짜"]))
        tc=find_col(d,["시간"])
        y=numeric(d[lc])

        if dtc is not None:
            z=pd.DataFrame({"dt":to_datetime_series(d[dtc]),"load":y}).dropna().sort_values("dt")
            if len(z)==0:
                continue
            z["dt"]=z["dt"].dt.floor("h")
            z=z.groupby("dt",as_index=False)["load"].max()
            return {
                "kind":"dated",
                "dt":[x.isoformat() for x in z["dt"]],
                "load":z["load"].astype(float).tolist(),
                "summary":{
                    "sheet":sheet,
                    "start":z["dt"].iloc[0].isoformat(),
                    "end":z["dt"].iloc[-1].isoformat(),
                    "min":float(z["load"].min()),
                    "max":float(z["load"].max()),
                    "rows":int(len(z)),
                }
            }

        if tc is not None:
            mins=[]
            for v in d[tc]:
                minute=np.nan
                if hasattr(v,"hour") and hasattr(v,"minute"):
                    minute=int(v.hour)*60+int(v.minute)
                else:
                    s=str(v)
                    m=re.search(r"(\d{1,2}):(\d{2})",s)
                    if m:
                        minute=int(m.group(1))*60+int(m.group(2))
                    else:
                        try:
                            f=float(v)
                            if 0 <= f < 1:
                                minute=int(round(f*1440))%1440
                        except Exception:
                            pass
                mins.append(minute)

            z=pd.DataFrame({"minute":mins,"load":y}).dropna().sort_values("minute")
            if len(z)==0:
                continue
            z=z.groupby("minute",as_index=False)["load"].max()
            return {
                "kind":"day",
                "minute":z["minute"].astype(int).tolist(),
                "load":z["load"].astype(float).tolist(),
                "summary":{
                    "sheet":sheet,
                    "min":float(z["load"].min()),
                    "max":float(z["load"].max()),
                    "rows":int(len(z)),
                }
            }

    raise ValueError("기존공장 MAIN_TR/MVA 부하 열을 찾지 못했습니다.")


def annual_data(temp_store=None,load_store=None):
    """
    화면의 연간 데이터.
    업로드 자료는 날짜가 겹치면 '실제 timestamp'로 맞춰서 반영한다.
    날짜가 겹치지 않을 때만 길이 기준 보간을 사용한다.
    """
    a=DATA["annual"].copy().sort_values("dt").reset_index(drop=True)
    target=pd.DatetimeIndex(pd.to_datetime(a["dt"]))

    if temp_store:
        t=pd.DataFrame({
            "dt":pd.to_datetime(temp_store["dt"],errors="coerce"),
            "value":pd.to_numeric(temp_store["temp"],errors="coerce")
        }).dropna().sort_values("dt")
        t=t.groupby(t["dt"].dt.floor("h"))["value"].max().sort_index()

        aligned=t.reindex(target)
        overlap=int(aligned.notna().sum())
        if overlap >= min(24,max(1,len(a)//100)):
            aligned=aligned.interpolate(method="time",limit_direction="both")
            a["temp"]=aligned.to_numpy(float)
        else:
            vals=t.to_numpy(float)
            a["temp"]=np.interp(
                np.linspace(0,1,len(a)),
                np.linspace(0,1,len(vals)),
                vals
            )

    if load_store and load_store.get("kind")=="dated":
        l=pd.DataFrame({
            "dt":pd.to_datetime(load_store["dt"],errors="coerce"),
            "value":pd.to_numeric(load_store["load"],errors="coerce")
        }).dropna().sort_values("dt")
        l=l.groupby(l["dt"].dt.floor("h"))["value"].max().sort_index()

        aligned=l.reindex(target)
        overlap=int(aligned.notna().sum())
        if overlap >= min(24,max(1,len(a)//100)):
            aligned=aligned.interpolate(method="time",limit_direction="both")
            a["base_main"]=aligned.to_numpy(float)
        else:
            vals=l.to_numpy(float)
            a["base_main"]=np.interp(
                np.linspace(0,1,len(a)),
                np.linspace(0,1,len(vals)),
                vals
            )

    return a


def selected_day(a, day_index, load_store=None):
    day_index=max(0,min(364,int(day_index)))
    base_dt=pd.Timestamp(a["dt"].iloc[0]).normalize()
    date=base_dt+pd.Timedelta(days=day_index)

    rows=a[pd.to_datetime(a["dt"]).dt.normalize()==date]
    if len(rows)>=2:
        mins=(pd.to_datetime(rows["dt"]).dt.hour*60+pd.to_datetime(rows["dt"]).dt.minute).to_numpy(float)
        temp=np.interp(np.arange(1440),mins,rows["temp"].to_numpy(float))
        base=np.interp(np.arange(1440),mins,rows["base_main"].to_numpy(float))
        temp_source="현재 연간 데이터 → 선택일 1분 보간"
    else:
        nearest=min(len(a)-1,day_index*24)
        temp=np.full(1440,float(a.iloc[nearest]["temp"]))
        base=np.full(1440,float(a.iloc[nearest]["base_main"]))
        temp_source="현재 연간 데이터"

    # 원본 데이터에서 극한일은 실제 1분 profile을 사용.
    # 단, 사용자가 외기온도/부하를 업로드하면 업로드 자료가 우선한다.
    exdate=pd.Timestamp(DATA["extreme"]["dt"].iloc[0]).normalize()
    if date==exdate and load_store is None:
        # temp는 annual_data가 이미 사용자 외기온도를 반영했을 수 있으므로
        # 원본 annual과 값이 같을 때만 원본 1분 온도를 사용.
        original_rows=DATA["annual"][pd.to_datetime(DATA["annual"]["dt"]).dt.normalize()==date]
        use_original_temp=True
        if len(rows) and len(original_rows):
            use_original_temp=abs(float(rows["temp"].max())-float(original_rows["temp"].max()))<1e-9
        if use_original_temp:
            temp=resample(DATA["extreme"]["temp"],1440)
            temp_source="원본 극한일 1분 외기온도"
        base=resample(DATA["extreme"]["base_main"],1440)

    # 사용자가 일일 부하패턴을 올렸으면 무조건 최우선
    if load_store and load_store.get("kind")=="day":
        x=np.asarray(load_store["minute"],float)
        y=np.asarray(load_store["load"],float)
        base=np.interp(np.arange(1440),x,y)

    return date,temp,base,temp_source

def voltage_before(total_capacity):
    out={}
    for b in BUS_IDS:
        cmap={c:DATA["voltage"][c][b] for c in CASE_POINTS}
        out[b]=case_interp(cmap,total_capacity)
    return out

def vmin_of(v):
    return np.nanmin(np.vstack([v[b] for b in BUS_IDS]),axis=0)


def low_bus_summary(v, limit=V_LIMIT):
    mins=[]
    for b in BUS_IDS:
        mv=float(np.nanmin(v[b]))
        if mv < limit:
            mins.append((b,mv))
    mins.sort(key=lambda x:x[1])
    return {
        "count": len(mins),
        "items": mins,
        "text": ", ".join([f"Bus {b} {mv:.4f}" for b,mv in mins[:6]]) if mins else "없음"
    }


def scenario(new_load_mva, a, day_index, dtr_params, load_store=None):
    """
    사용자 제공 '보상최종_기존공장_신규30MVA_보강전후_통합그래프.xlsx'의
    설계 순서를 그대로 일반화한다.

    1) 계획 DTR = 현재 연간 온도자료로 계산한 최저 DTR의 보수적 정수값
    2) DTR 부족용량 = 기존공장 계획부하 + 신규공장 부하 - 계획 DTR
    3) 병렬 TR = 부족용량 × (1 + 여유율), 0.1 MVA 단위 올림
       - 기본 30 MVA: (90 + 30 - 99) × 1.10 = 23.1 MVA
       - 전압 때문에 TR를 더 키우지 않는다.
    4) TR 적용 후 Vmin < 0.95 pu일 때만 Bus 6 Shunt Capacitor를 사용한다.
    5) Shunt는 사용자 데이터의 3.5 MVAr PSS/E 전압개선 민감도를 이용하여
       0.5 MVAr 단위의 '최소 보상량'만 선정한다.

    즉, 열적 병목은 DTR 기반 최소 TR로 해결하고,
    남는 저전압은 커패시터 무효전력 보상으로 해결한다.
    """
    total=max(0.0,min(100.0,float(new_load_mva)))

    date,temp,base_raw,temp_source=selected_day(a,day_index,load_store)
    rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp,plan_existing,reserve,pv_required=dtr_params

    # ------------------------------------------------------------------
    # DTR: 업로드된 온도까지 즉시 반영
    # ------------------------------------------------------------------
    daily_dtr=calc_dtr(temp,rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp)
    annual_dtr=calc_dtr(
        a["temp"].to_numpy(float),rating,hotspot,oil_rise,wind_rise,
        loss_num,loss_den,oil_exp,wind_exp
    )
    plan_dtr=float(math.floor(np.nanmin(annual_dtr)))

    # ------------------------------------------------------------------
    # 기존공장 계획부하
    # - 기본 자료: 사용자가 지정한 90 MVA
    # - 새로운 부하패턴 업로드: 업로드 파일의 실제 peak를 계획부하로 사용
    # ------------------------------------------------------------------
    if load_store:
        try:
            if load_store.get("kind")=="day":
                design_existing=float(np.nanmax(np.asarray(load_store["load"],float)))
            else:
                design_existing=float(np.nanmax(a["base_main"].to_numpy(float)))
        except Exception:
            design_existing=float(plan_existing)
    else:
        design_existing=float(plan_existing)

    # 기본 기존공장 패턴은 90 MVA 계획값에 맞춰 형상만 유지.
    # 업로드된 부하패턴은 사용자가 준 실제 값 자체를 사용.
    if load_store:
        base=np.asarray(base_raw,float).copy()
    else:
        source_peak=float(np.nanmax(DATA["annual"]["base_main"].to_numpy(float)))
        scale=(design_existing/source_peak) if source_peak>1e-9 else 1.0
        base=np.asarray(base_raw,float)*scale

    # ------------------------------------------------------------------
    # MAIN_TR 계획부하 형상 — 사용자 계산근거 그대로
    # 30 MVA 예시의 핵심은 PSS/E 원본 형상을 90+30=120 MVA 계획부하에
    # 맞춰 스케일하는 것이다. 보강 후 기준용량을 초과하는 원본 절대값을
    # 그대로 사용하지 않는다.
    # ------------------------------------------------------------------
    raw0=DATA["mva_cases"][0.0]
    raw_total=case_interp(DATA["mva_cases"],total)
    plan_peak=float(design_existing+total)

    # 기본 30 MVA는 보상최종 Excel의 실제 표시형상을 그대로 사용.
    source30=(abs(total-30.0)<1e-9 and load_store is None)
    if source30:
        bank=DATA["final30_bank"].copy()
    elif load_store is None:
        # 다른 Slider 값도 동일한 원칙: 해당 PSS/E 형상을 계획 Peak에 맞춰 스케일.
        raw_peak=float(np.nanmax(raw_total))
        factor=(plan_peak/raw_peak) if raw_peak>1e-9 else 1.0
        bank=np.asarray(raw_total,float)*factor
    else:
        # 사용자 기존부하패턴을 넣은 경우에는 실제 pattern + 신규부하 증분형상.
        raw_inc=np.maximum(0.0,np.asarray(raw_total,float)-np.asarray(raw0,float))
        inc_peak=float(np.nanmax(raw_inc))
        new_shape=(raw_inc*(total/inc_peak)) if total>0 and inc_peak>1e-9 else np.zeros(1440)
        bank=np.asarray(base,float)+new_shape

    # 원본 30 MVA 전압 3단계 데이터 적용 여부
    exact30=source30

    # ------------------------------------------------------------------
    # Feeder: PSS/E case 형상 사용. 35 MVA 초과 시 회선수만 자동증설.
    # ------------------------------------------------------------------
    feeder_total=case_interp(DATA["feeder_cases"],total) if total>0 else np.zeros(1440)
    if exact30:
        feeder_total=DATA["final30_feeder"].copy()

    feeder_peak=float(np.nanmax(feeder_total)) if total>0 else 0.0
    feeder_count=0 if total<=0 else max(1,int(math.ceil(feeder_peak/FEEDER_RATING)))
    per_feeder=np.zeros(1440) if feeder_count==0 else feeder_total/feeder_count

    # ------------------------------------------------------------------
    # 1) DTR 기반 최소 병렬 TR 산정식 — 사용자 파일 그대로
    # ------------------------------------------------------------------
    shortage=max(0.0,design_existing+total-plan_dtr)
    tr_need=ceil_tenth(shortage*(1.0+float(reserve)))

    # 기본 30 MVA는 정확히 23.1 MVA가 나와야 한다.
    if exact30 and abs(design_existing-90.0)<0.11 and abs(plan_dtr-99.0)<0.11 and abs(float(reserve)-0.10)<1e-9:
        tr_need=float(DATA["tr_ref"])

    after_capacity=np.full(1440,plan_dtr+tr_need)
    total_nameplate=rating+tr_need

    # ------------------------------------------------------------------
    # 2) 전압 계산: TR는 위 식으로 고정. 전압 때문에 TR 추가증설 금지.
    # ------------------------------------------------------------------
    before=voltage_before(total)

    # 새로운 기존공장 부하패턴이 들어오면 PSS/E 0→30 MVA 전압 민감도로 반영
    ref_base=DATA["mva_cases"][0.0]
    ref_30=DATA["mva_cases"][30.0]
    dm=ref_30-ref_base
    if load_store:
        base_delta=np.asarray(base,float)-ref_base
        for b in BUS_IDS:
            dv=DATA["voltage"][30.0][b]-DATA["voltage"][0.0][b]
            sens=np.divide(dv,dm,out=np.zeros_like(dv),where=np.abs(dm)>0.25)
            before[b]=before[b]+sens*base_delta

    tr_ref=float(DATA["tr_ref"])
    q_ref=float(DATA["shunt_ref"])

    if tr_need>0:
        if exact30 and abs(tr_need-tr_ref)<0.051:
            after_tr={b:DATA["comp"]["tr"][b].copy() for b in BUS_IDS}
        else:
            scale=tr_need/tr_ref if tr_ref>0 else 0.0
            after_tr={}
            for b in BUS_IDS:
                uplift=DATA["comp"]["tr"][b]-DATA["comp"]["before"][b]
                after_tr[b]=before[b]+uplift*scale
    else:
        after_tr={b:before[b].copy() for b in BUS_IDS}

    # ------------------------------------------------------------------
    # 3) 최소 Shunt Capacitor 선정
    # 사용자 가이드: TR 후 Vmin < 0.95일 때만, 0.5 MVAr 단위 최소 Shunt.
    # ------------------------------------------------------------------
    final={b:after_tr[b].copy() for b in BUS_IDS}
    q=0.0

    if float(np.nanmin(vmin_of(after_tr))) < V_LIMIT:
        if exact30 and abs(tr_need-tr_ref)<0.051:
            # 원본 PSS/E 정답: Bus6 3.5 MVAr
            q=q_ref
            final={b:DATA["comp"]["shunt"][b].copy() for b in BUS_IDS}
        else:
            # PSS/E 3.5MVAr 결과에서 Bus/time별 dV/dQ를 추출.
            # 위반하는 모든 Bus/time이 0.95 이상이 되는 최소 Q를 계산한 뒤
            # 0.5 MVAr 단위로 올림.
            q_needed=0.0
            unresolved=False

            for b in BUS_IDS:
                q_sens=(DATA["comp"]["shunt"][b]-DATA["comp"]["tr"][b])/q_ref
                v0=after_tr[b]
                need_mask=v0 < V_LIMIT
                if not np.any(need_mask):
                    continue

                usable=need_mask & (q_sens > 1e-8)
                if np.any(need_mask & ~usable):
                    unresolved=True

                if np.any(usable):
                    req=(V_LIMIT-v0[usable])/q_sens[usable]
                    q_needed=max(q_needed,float(np.nanmax(req)))

            q=max(0.0,math.ceil(q_needed*2.0)/2.0)

            # 최소 0.5 MVAr step을 실제로 적용하고 전 1,440분 검증.
            def apply_q(q_mvar):
                out={}
                for b in BUS_IDS:
                    q_sens=(DATA["comp"]["shunt"][b]-DATA["comp"]["tr"][b])/q_ref
                    out[b]=after_tr[b]+q_sens*q_mvar
                return out

            final=apply_q(q)

            # 선형 민감도 오차가 있으면 TR는 건드리지 않고 Capacitor만 0.5 MVAr씩 추가.
            guard=0
            while float(np.nanmin(vmin_of(final))) < V_LIMIT and guard<100:
                guard+=1
                q+=0.5
                final=apply_q(q)

    pre_margin=daily_dtr-bank
    post_margin=after_capacity-bank

    # ------------------------------------------------------------------
    # PV: 현재의 'DTR 최소 TR + 필요한 경우 최소 C' 최종계통을 반영
    # ------------------------------------------------------------------
    w=total/30.0
    pv_before=PV1_MODEL + w*(PV2_MODEL-PV1_MODEL)
    pv_final=pv_before.copy()
    if tr_need>0:
        pv_final=pv_final+(tr_need/tr_ref)*(PV3_MODEL-PV2_MODEL)
    if q>0:
        pv_final=pv_final+(q/q_ref)*(PV4_MODEL-PV3_MODEL)

    before_low=low_bus_summary(before)
    tr_low=low_bus_summary(after_tr)
    final_low=low_bus_summary(final)

    return {
        "date":date,"temp":temp,"base":base,"temp_source":temp_source,
        "annual_dtr":annual_dtr,"daily_dtr":daily_dtr,"plan_dtr":plan_dtr,
        "design_existing":design_existing,"shortage":shortage,
        "bank":bank,"feeder_total":feeder_total,"per_feeder":per_feeder,
        "feeder_count":feeder_count,
        "before":before,"after_tr":after_tr,"final":final,
        "v_before":vmin_of(before),"v_tr":vmin_of(after_tr),"v_final":vmin_of(final),
        "before_low":before_low,"tr_low":tr_low,"final_low":final_low,
        "tr":tr_need,"q":q,"tr_plan":tr_need,"tr_actual":shortage,
        "reserve_pct":float(reserve),
        "pre_margin":pre_margin,"post_margin":post_margin,
        "after_capacity":after_capacity,"total_nameplate":total_nameplate,
        "total_new":total,"exact30":exact30,
        "pv_before":pv_before,"pv_final":pv_final,"pv_required":pv_required,
    }

def pv_cross(mw,v,limit=.95):
    x=np.asarray(mw,float); y=np.asarray(v,float)
    ok=np.isfinite(x)&np.isfinite(y)
    x=x[ok]; y=y[ok]
    if len(x)<2:return np.nan
    if y[0]<limit:return 0.0
    for i in range(1,len(x)):
        if y[i] < limit <= y[i-1]:
            r=(limit-y[i-1])/(y[i]-y[i-1]) if y[i]!=y[i-1] else 0
            return float(x[i-1]+r*(x[i]-x[i-1]))
    return float(x[-1]) if np.nanmin(y)>=limit else np.nan


# ------------------------------ graph style ------------------------------

def base_fig(ytitle,yrange=None,height=310):
    import plotly.graph_objects as go
    f=go.Figure()
    f.update_layout(
        height=height,
        margin=dict(l=58,r=32,t=22,b=68),
        paper_bgcolor="white",plot_bgcolor="white",
        font=dict(family="Segoe UI,Malgun Gothic,Arial",size=10,color="#40566F"),
        showlegend=True,
        legend=dict(orientation="h",x=.5,xanchor="center",y=-.19,font=dict(size=9)),
        hovermode="x unified",
        # 슬라이더/온도/부하가 바뀌어도 사용자가 확대해 둔 화면 유지
        uirevision="keep-user-zoom",
    )
    f.update_xaxes(
        showgrid=False,linecolor="#7E9DBD",tickfont=dict(size=9),
        automargin=True
    )
    f.update_yaxes(
        title=ytitle,gridcolor=GRID,griddash="dash",linecolor="#7E9DBD",
        tickfont=dict(size=9),automargin=True
    )
    if yrange is not None:
        f.update_yaxes(range=yrange)
    return f

def annual_temp_fig(a,hour_idx):
    import plotly.graph_objects as go
    f=base_fig("",[-10,40],310)
    f.add_trace(go.Scattergl(
        x=a["dt"],y=a["temp"],mode="lines",name="시간별 최고 외기온도",
        line=dict(color=ANNUAL_TEMP,width=1.05)
    ))
    f.add_trace(go.Scattergl(
        x=a["dt"],y=np.full(len(a),30.0),mode="lines",name="30°C 참고선",
        line=dict(color=REF_RED,width=1.4,dash="dash")
    ))
    imax=int(np.nanargmax(a["temp"].to_numpy(float)))
    f.add_trace(go.Scatter(
        x=[a.iloc[imax]["dt"]],y=[a.iloc[imax]["temp"]],
        mode="markers",name=f"연간 최고 {a.iloc[imax]['temp']:.1f}°C",
        marker=dict(color=MAX_POINT,size=8)
    ))
    i=max(0,min(len(a)-1,int(hour_idx)))
    x=a.iloc[i]["dt"]
    f.add_vline(x=x,line_color="#111827",line_width=1.6)
    f.add_trace(go.Scatter(
        x=[x],y=[a.iloc[i]["temp"]],mode="markers",name="현재 모델 시각",
        marker=dict(color="#111827",size=7)
    ))
    f.update_yaxes(dtick=5)
    return f


def annual_dtr_fig(a,dtr,hour_idx):
    import plotly.graph_objects as go
    allv=np.concatenate([np.asarray(dtr,float),a["base_main"].to_numpy(float)])
    lo=max(0,math.floor((np.nanmin(allv)-8)/10)*10)
    hi=math.ceil((np.nanmax(allv)+8)/10)*10
    f=base_fig("",[lo,hi],310)
    f.add_trace(go.Scattergl(
        x=a["dt"],y=dtr,mode="lines",name="DTR 허용용량",
        line=dict(color=ANNUAL_DTR,width=1.05)
    ))
    f.add_trace(go.Scattergl(
        x=a["dt"],y=a["base_main"],mode="lines",name="기존공장 MAIN_TR",
        line=dict(color=BASE_MAIN,width=1.05)
    ))
    imin=int(np.nanargmin(dtr))
    f.add_trace(go.Scatter(
        x=[a.iloc[imin]["dt"]],y=[dtr[imin]],mode="markers",
        name=f"최저 DTR {dtr[imin]:.3f} MVA",
        marker=dict(color=MAX_POINT,size=8)
    ))
    i=max(0,min(len(a)-1,int(hour_idx)))
    x=a.iloc[i]["dt"]
    f.add_vline(x=x,line_color="#111827",line_width=1.6)
    f.add_trace(go.Scatter(
        x=[x],y=[dtr[i]],mode="markers",name="현재 DTR",
        marker=dict(color="#111827",size=7)
    ))
    f.update_yaxes(dtick=10)
    return f

def daily_temp_fig(sc):
    import plotly.graph_objects as go
    f=base_fig("온도(°C)",[0,40],285)
    f.add_trace(go.Scatter(x=MINUTE_X,y=sc["temp"],mode="lines",name="1분 외기온도",
                           line=dict(color=BASE_MAIN,width=1.2)))
    imax=int(np.nanargmax(sc["temp"]))
    f.add_trace(go.Scatter(x=[MINUTE_X[imax]],y=[sc["temp"][imax]],mode="markers",
                           name=f"최고기온 {sc['temp'][imax]:.1f}°C",marker=dict(color=MAX_POINT,size=7)))
    f.update_xaxes(tickmode="array",tickvals=[MINUTE_X[i] for i in range(0,1440,120)])
    f.update_yaxes(dtick=10)
    return f



def hottest_day_profiles(a,rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp):
    """현재 적용된 연간 외기온도에서 가장 더운 날을 찾아 1분 온도/DTR profile 생성."""
    imax=int(np.nanargmax(a["temp"].to_numpy(float)))
    hottest_ts=pd.Timestamp(a.iloc[imax]["dt"])
    start_day=pd.Timestamp(a["dt"].iloc[0]).normalize()
    hot_day_index=max(0,min(364,(hottest_ts.normalize()-start_day).days))
    _,temp,_,temp_source=selected_day(a,hot_day_index,None)
    dtr=calc_dtr(temp,rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp)
    return hottest_ts.normalize(),temp,dtr,temp_source

def hottest_temp_fig(date,temp):
    import plotly.graph_objects as go
    lo=max(-10,math.floor((float(np.nanmin(temp))-3)/5)*5)
    hi=min(50,math.ceil((float(np.nanmax(temp))+3)/5)*5)
    if hi-lo<15: hi=lo+15
    f=base_fig("외기온도 (°C)",[lo,hi],320)
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=temp,mode="lines",name="가장 더운 날 외기온도",
        line=dict(color=ANNUAL_TEMP,width=1.8)
    ))
    imax=int(np.nanargmax(temp))
    f.add_trace(go.Scatter(
        x=[MINUTE_X[imax]],y=[temp[imax]],mode="markers",
        name=f"최고 {temp[imax]:.1f}°C",marker=dict(color=MAX_POINT,size=8)
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    return f

def hottest_dtr_fig(date,dtr,rating):
    import plotly.graph_objects as go
    lo=max(0,math.floor((float(np.nanmin(dtr))-5)/5)*5)
    hi=math.ceil((max(float(np.nanmax(dtr)),float(rating))+5)/5)*5
    f=base_fig("DTR 허용용량 (MVA)",[lo,hi],320)
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=dtr,mode="lines",name="가장 더운 날 DTR",
        line=dict(color=ANNUAL_DTR,width=1.8)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,float(rating)),mode="lines",
        name=f"{rating:.0f} MVA 정격",line=dict(color="#C00000",width=1.4,dash="dash")
    ))
    imin=int(np.nanargmin(dtr))
    f.add_trace(go.Scatter(
        x=[MINUTE_X[imin]],y=[dtr[imin]],mode="markers",
        name=f"최저 {dtr[imin]:.2f} MVA",marker=dict(color=MAX_POINT,size=8)
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    return f

def feeder_fig(sc):
    import plotly.graph_objects as go
    y=resample(sc["per_feeder"],720)
    yr=[0,40]
    if len(y) and np.nanmin(y)>=26 and np.nanmax(y)<=36:
        yr=[26,36]
    f=base_fig("",yr,315)
    label=("Feeder 없음" if sc["feeder_count"]==0 else
           f"자동보강 후 회선당 Feeder ({sc['feeder_count']}회선)")
    f.add_trace(go.Scatter(
        x=TWO_MIN_X,y=y,mode="lines",name=label,
        line=dict(color=ACCENT4,width=1.8)
    ))
    f.add_trace(go.Scatter(
        x=TWO_MIN_X,y=np.full(720,FEEDER_RATING),mode="lines",name="35 MVA 정격",
        line=dict(color=RED_DARK,width=1.6,dash="dash")
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(24)])
    return f



def mva_before_fig(sc,rating):
    """
    보강 전만 표시:
    - 기존+신규 MAIN_TR
    - 시간별 DTR
    - 기존 105 MVA 정격
    - 계획 DTR
    - 기존공장만 MAIN_TR
    """
    import plotly.graph_objects as go

    all_values=np.concatenate([
        np.asarray(sc["bank"],float),
        np.asarray(sc["daily_dtr"],float),
        np.full(1440,float(rating)),
        np.full(1440,float(sc["plan_dtr"])),
        np.asarray(sc["base"],float),
    ])
    finite=all_values[np.isfinite(all_values)]
    lo=max(0.0,math.floor((float(np.nanmin(finite))-8.0)/10.0)*10.0)
    hi=math.ceil((float(np.nanmax(finite))+12.0)/10.0)*10.0
    if hi-lo<40:
        hi=lo+40

    f=base_fig("MVA",[lo,hi],350)
    series=[
        (sc["bank"],f"기존+신규 MAIN_TR ({sc['total_new']:.0f} MVA)","#0070C0","solid",2.0),
        (sc["daily_dtr"],"시간별 DTR 허용용량","#ED7D31","solid",1.8),
        (np.full(1440,rating),f"기존 TR {rating:.0f} MVA 정격","#C00000","dash",1.5),
        (np.full(1440,sc["plan_dtr"]),f"계획 DTR {sc['plan_dtr']:.0f} MVA","#7030A0","dot",1.6),
        (sc["base"],"기존공장 MAIN_TR","#7F8C8D","solid",1.3),
    ]
    for y,name,color,dash,width in series:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=y,mode="lines",name=name,
            line=dict(color=color,width=width,dash=dash)
        ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    span=hi-lo
    f.update_yaxes(dtick=10 if span<=100 else 20)
    return f


def mva_after_fig(sc,rating):
    """
    보강 후만 표시:
    - 기존+신규 MAIN_TR
    - 자동보강 후 기준용량
    - 보강 후 명판정격
    - 기존공장 MAIN_TR 참고
    보강량이 커져도 Y축을 자동 확장한다.
    """
    import plotly.graph_objects as go

    all_values=np.concatenate([
        np.asarray(sc["bank"],float),
        np.asarray(sc["after_capacity"],float),
        np.full(1440,float(sc["total_nameplate"])),
        np.asarray(sc["base"],float),
    ])
    finite=all_values[np.isfinite(all_values)]
    lo=max(0.0,math.floor((float(np.nanmin(finite))-8.0)/10.0)*10.0)
    hi=math.ceil((float(np.nanmax(finite))+15.0)/10.0)*10.0
    if hi-lo<40:
        hi=lo+40

    f=base_fig("MVA",[lo,hi],350)
    series=[
        (sc["bank"],f"보강 후 대상 MAIN_TR 부하 ({sc['total_new']:.0f} MVA)","#0070C0","solid",2.0),
        (sc["after_capacity"],f"자동보강 기준용량 {sc['plan_dtr']+sc['tr']:.1f} MVA","#70AD47","dash",1.9),
        (np.full(1440,sc["total_nameplate"]),f"보강 후 명판정격 {sc['total_nameplate']:.1f} MVA","#FFC000","dashdot",1.8),
        (sc["base"],"기존공장 MAIN_TR 참고","#7F8C8D","solid",1.3),
    ]
    for y,name,color,dash,width in series:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=y,mode="lines",name=name,
            line=dict(color=color,width=width,dash=dash)
        ))
    min_margin=float(np.nanmin(sc["post_margin"]))
    f.add_annotation(
        x=.99,y=.97,xref="paper",yref="paper",xanchor="right",yanchor="top",
        text=f"보강 후 최소여유 {min_margin:+.2f} MVA · {'PASS' if min_margin>=0 else 'FAIL'}",
        showarrow=False,bgcolor="rgba(255,255,255,.92)",
        bordercolor="#70AD47" if min_margin>=0 else "#C00000",
        font=dict(size=10,color="#2E7D32" if min_margin>=0 else "#C00000")
    )
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    span=hi-lo
    f.update_yaxes(dtick=10 if span<=100 else 20)
    return f

def mva_fig(sc,rating):
    import plotly.graph_objects as go

    all_values=np.concatenate([
        np.asarray(sc["bank"],float),
        np.asarray(sc["daily_dtr"],float),
        np.asarray(sc["after_capacity"],float),
        np.full(1440,float(rating)),
        np.full(1440,float(sc["plan_dtr"])),
        np.full(1440,float(sc["total_nameplate"])),
        np.asarray(sc["base"],float),
    ])
    finite=all_values[np.isfinite(all_values)]
    lo=max(0.0,math.floor((float(np.nanmin(finite))-8.0)/10.0)*10.0)
    hi=math.ceil((float(np.nanmax(finite))+12.0)/10.0)*10.0
    if hi-lo < 50:
        hi=lo+50

    f=base_fig("",[lo,hi],350)

    # 같은 그래프 안에서는 전부 서로 다른 색.
    series=[
        (sc["bank"],f"기존+신규공장 MAIN_TR (신규 {sc['total_new']:.0f} MVA)","#0070C0","solid",2.0),
        (sc["daily_dtr"],"보강 전 DTR 허용용량","#ED7D31","solid",1.7),
        (sc["after_capacity"],f"보강 후 기준용량 {sc['plan_dtr']+sc['tr']:.1f} MVA","#70AD47","dash",1.8),
        (np.full(1440,rating),f"보강 전 {rating:.0f} MVA 정격","#C00000","dash",1.5),
        (np.full(1440,sc["plan_dtr"]),"최저 DTR 계획 기준","#7030A0","dot",1.6),
        (np.full(1440,sc["total_nameplate"]),f"보강 후 {sc['total_nameplate']:.1f} MVA 명판정격","#FFC000","dashdot",1.7),
        (sc["base"],"기존공장만 MAIN_TR (신규 0 MVA)","#7F8C8D","solid",1.4),
    ]
    for y,name,color,dash,width in series:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=y,mode="lines",name=name,
            line=dict(color=color,width=width,dash=dash)
        ))

    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(24)])
    span=hi-lo
    f.update_yaxes(dtick=10 if span<=120 else 20)
    return f




def voltage_three_stage_fig(sc, stage):
    """
    전압 보상 과정을 반드시 3단계로 나눠 표시한다.

    ① 보강 전
    ② DTR 계산근거로 산정한 병렬 TR 추가 후
    ③ TR 후에도 0.95 pu 미만이면 Bus 6 Shunt Capacitor 추가 후

    기본 30MVA Case는 사용자가 준 Excel의
    전압_보강전 / 전압_TR보강후 / 전압_Shunt후 데이터를 그대로 사용한다.
    """
    import plotly.graph_objects as go

    if stage=="before":
        data=sc["before"]
        title="① 보강 전"
        accent="#0B6FA4"
    elif stage=="tr":
        data=sc["after_tr"]
        title=f"② 병렬 TR +{sc['tr']:.1f} MVA 후"
        accent="#F47A34"
    else:
        data=sc["final"]
        title=("③ 커패시터 미설치" if sc["q"]<=0
               else f"③ Bus 6 Shunt Capacitor +{sc['q']:.1f} MVAr 후")
        accent="#1B8A5A" if sc["q"]>0 else "#64748B"

    stack=np.vstack([np.asarray(data[b],float) for b in BUS_IDS])
    vmin=np.nanmin(stack,axis=0)
    finite=stack[np.isfinite(stack)]

    ymin=min(0.930,math.floor((float(np.nanmin(finite))-0.003)*1000)/1000)
    ymax=max(0.990,math.ceil((float(np.nanmax(finite))+0.003)*1000)/1000)

    f=base_fig("전압 (pu)",[ymin,ymax],400)

    # 원본 Excel처럼 Bus 5~18 전체 전압
    for b in BUS_IDS:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=data[b],mode="lines",
            name=f"Bus {b}",
            line=dict(color=BUS_COLORS[b],width=.9),
            opacity=.80
        ))

    # 최저전압을 굵게
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=vmin,mode="lines",
        name="최저전압",
        line=dict(color="#263238",width=2.3)
    ))

    # 기준선
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",
        name="0.95 pu 기준",
        line=dict(color="#D1495B",width=1.6,dash="dash")
    ))

    i=int(np.nanargmin(vmin))
    passed=float(vmin[i])>=V_LIMIT
    f.add_trace(go.Scatter(
        x=[MINUTE_X[i]],y=[vmin[i]],mode="markers",
        name=f"Vmin {vmin[i]:.4f} · {'PASS' if passed else 'FAIL'}",
        marker=dict(color=accent,size=9)
    ))

    f.add_annotation(
        x=.01,y=.98,xref="paper",yref="paper",
        text=f"<b>{title}</b><br>Vmin {vmin[i]:.4f} pu · {'PASS' if passed else 'FAIL'}",
        showarrow=False,xanchor="left",yanchor="top",
        bgcolor="rgba(255,255,255,.94)",
        bordercolor=accent,
        font=dict(size=10,color=accent)
    )

    f.update_layout(
        margin=dict(l=58,r=24,t=22,b=92),
        legend=dict(orientation="h",x=.5,xanchor="center",y=-.20,font=dict(size=8)),
        hovermode="x unified",
        uirevision="keep-user-zoom"
    )
    f.update_xaxes(
        tickmode="array",
        tickvals=[f"{h:02d}:00" for h in range(0,24,3)],
        showgrid=False
    )
    f.update_yaxes(
        tickformat=".3f",
        dtick=.01 if (ymax-ymin)<=.09 else .02,
        gridcolor="#D7EAF7",
        griddash="dash"
    )
    return f

def voltage_comp_exact_fig(sc):
    """
    HWPX 원본 전압보상 그래프 형식:
    파랑=보강 전, 주황=병렬 TR 후, 초록=Bus 6 Shunt 후, 빨강점선=0.95 pu.
    정상 30MVA Case에서는 원본처럼 0.930~0.980 pu를 사용하고,
    데이터가 범위를 벗어날 때만 자동으로 축을 넓힌다.
    """
    import plotly.graph_objects as go

    before=np.asarray(sc["v_before"],float)
    tr=np.asarray(sc["v_tr"],float)
    final=np.asarray(sc["v_final"],float)

    allv=np.concatenate([before,tr,final])
    finite=allv[np.isfinite(allv)]
    data_min=float(np.nanmin(finite))
    data_max=float(np.nanmax(finite))

    ymin=min(0.930, math.floor((data_min-0.003)*1000)/1000)
    ymax=max(0.980, math.ceil((data_max+0.003)*1000)/1000)

    f=base_fig("전압 (pu)",[ymin,ymax],455)

    f.add_trace(go.Scatter(
        x=MINUTE_X,y=before,mode="lines",
        name="보강 전 최저",
        line=dict(color="#00A6F0",width=1.45)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=tr,mode="lines",
        name=f"{sc['tr']:.1f} MVA TR 후 최저",
        line=dict(color="#FF7F27",width=1.45)
    ))

    if sc["q"] > 0:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=final,mode="lines",
            name=f"Bus 6 {sc['q']:.1f} MVAr 후 최저",
            line=dict(color="#169B62",width=1.55)
        ))

    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",
        name="0.95 pu 기준",
        line=dict(color="#E31A1C",width=1.5,dash="dash")
    ))

    # 보상 여부/보상량이 그래프 안에서도 바로 보이도록 최저점 표시
    i_tr=int(np.nanargmin(tr))
    f.add_trace(go.Scatter(
        x=[MINUTE_X[i_tr]],y=[tr[i_tr]],mode="markers",
        name=f"TR 후 최저 {tr[i_tr]:.4f} pu",
        marker=dict(color="#FF7F27",size=8)
    ))

    if sc["q"] > 0:
        i_fin=int(np.nanargmin(final))
        f.add_trace(go.Scatter(
            x=[MINUTE_X[i_fin]],y=[final[i_fin]],mode="markers",
            name=f"C +{sc['q']:.1f} MVAr → {final[i_fin]:.4f} pu",
            marker=dict(color="#169B62",size=9)
        ))
        f.add_annotation(
            x=MINUTE_X[i_fin], y=final[i_fin],
            text=f"Bus 6 Shunt Capacitor +{sc['q']:.1f} MVAr",
            showarrow=True, arrowhead=2, ax=65, ay=-38,
            bgcolor="rgba(255,255,255,.92)",
            bordercolor="#169B62",
            font=dict(size=10,color="#116B46")
        )
    else:
        f.add_annotation(
            x=.98,y=.04,xref="paper",yref="paper",
            text="Shunt Capacitor 미설치 (TR 후 전압 PASS)",
            showarrow=False,xanchor="right",
            bgcolor="rgba(255,255,255,.92)",
            bordercolor="#94A3B8",
            font=dict(size=10,color="#64748B")
        )

    f.update_layout(
        margin=dict(l=62,r=28,t=18,b=82),
        legend=dict(orientation="h",x=.5,xanchor="center",y=-.17,font=dict(size=9)),
        hovermode="x unified",
        uirevision="keep-user-zoom"
    )
    f.update_xaxes(
        tickmode="array",
        tickvals=[f"{h:02d}:00" for h in range(24)],
        tickangle=0,
        showgrid=False,
        automargin=True
    )
    f.update_yaxes(
        tickformat=".3f",
        dtick=.010 if (ymax-ymin)<=.08 else .020,
        gridcolor="#D7EAF7",
        griddash="dash",
        automargin=True
    )
    return f

def voltage_tr_comp_fig(sc):
    """전압 보상 1단계: 보강 전 vs 병렬 TR 보강 후."""
    import plotly.graph_objects as go
    vals=np.concatenate([np.asarray(sc["v_before"],float),np.asarray(sc["v_tr"],float)])
    finite=vals[np.isfinite(vals)]
    lo=max(.80,math.floor((float(np.nanmin(finite))-.006)*1000)/1000)
    hi=min(1.05,math.ceil((float(np.nanmax(finite))+.006)*1000)/1000)
    if hi-lo<.025:
        hi=lo+.025

    f=base_fig("전압 (pu)",[lo,hi],350)
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["v_before"],mode="lines",name="보강 전 최저전압",
        line=dict(color="#0070C0",width=1.8)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["v_tr"],mode="lines",name=f"병렬 TR +{sc['tr']:.1f} MVA 후",
        line=dict(color="#ED7D31",width=1.8)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",name="0.950 pu 기준",
        line=dict(color="#C00000",width=1.5,dash="dash")
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(tickformat=".3f")
    return f


def voltage_cap_comp_fig(sc):
    """전압 보상 2단계: TR 후 vs 커패시터/Shunt 보상 후."""
    import plotly.graph_objects as go
    vals=np.concatenate([np.asarray(sc["v_tr"],float),np.asarray(sc["v_final"],float)])
    finite=vals[np.isfinite(vals)]
    lo=max(.80,math.floor((float(np.nanmin(finite))-.006)*1000)/1000)
    hi=min(1.05,math.ceil((float(np.nanmax(finite))+.006)*1000)/1000)
    if hi-lo<.025:
        hi=lo+.025

    f=base_fig("전압 (pu)",[lo,hi],350)
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["v_tr"],mode="lines",name=f"TR 보강 후 ({sc['tr']:.1f} MVA)",
        line=dict(color="#ED7D31",width=1.8)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["v_final"],mode="lines",name=f"커패시터/Shunt 보상 후 ({sc['q']:.1f} MVAr)",
        line=dict(color="#70AD47",width=1.9)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",name="0.950 pu 기준",
        line=dict(color="#C00000",width=1.5,dash="dash")
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(tickformat=".3f")
    return f

def voltage_stage_fig(sc):
    import plotly.graph_objects as go
    f=base_fig("",[.93,.98],335)
    f.add_trace(go.Scatter(x=MINUTE_X,y=sc["v_before"],mode="lines",name="보강 전 최저",
                           line=dict(color=ACCENT4,width=1.3)))
    f.add_trace(go.Scatter(x=MINUTE_X,y=sc["v_tr"],mode="lines",name=f"{sc['tr']:.1f} MVA TR 후 최저",
                           line=dict(color=ORANGE,width=1.3)))
    f.add_trace(go.Scatter(x=MINUTE_X,y=sc["v_final"],mode="lines",name=f"Bus6 {sc['q']:.1f}MVAr 후 최저",
                           line=dict(color=GREEN,width=1.3)))
    f.add_trace(go.Scatter(x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",name="0.95 pu 기준",
                           line=dict(color=REF_RED,width=1.5,dash="dash")))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(24)])
    f.update_yaxes(dtick=.01,tickformat=".3f")
    return f






def pv_fig(sc):
    """
    보상 전 / 보상 후 PV 곡선만 표시.

    중요:
    곡선이 x=0에서 이미 0.95 pu 아래라면 '0 MW에서 0.95와 교차'한 것이 아니다.
    따라서 그 경우 0 MW 기준선 위에 가짜 점을 찍지 않고
    '시작점부터 0.95 미만'이라고만 표시한다.
    """
    import plotly.graph_objects as go

    xmax=140.0
    x=np.asarray(PV_MODEL_X,float)
    y_before=np.asarray(sc["pv_before"],float)
    y_after=np.asarray(sc["pv_final"],float)
    mask=x<=xmax

    f=base_fig("전압 (pu)",[.68,1.01],440)

    f.add_trace(go.Scatter(
        x=x[mask],y=y_before[mask],
        mode="lines",
        name=f"보상 전 · 신규부하 {sc['total_new']:.0f} MVA",
        line=dict(color="#0070C0",width=2.2)
    ))

    after_label=(
        f"보상 후 · TR +{sc['tr']:.1f} MVA"
        if sc["q"]<=0 else
        f"보상 후 · TR +{sc['tr']:.1f} MVA + Bus6 C {sc['q']:.1f} MVAr"
    )
    f.add_trace(go.Scatter(
        x=x[mask],y=y_after[mask],
        mode="lines",
        name=after_label,
        line=dict(color="#1B8A5A",width=2.4)
    ))

    f.add_trace(go.Scatter(
        x=[0,xmax],y=[V_LIMIT,V_LIMIT],
        mode="lines",
        name="0.950 pu 기준",
        line=dict(color="#C00000",width=1.5,dash="dash")
    ))

    def real_cross(xv,yv,limit):
        """
        실제로 위→아래로 0.95를 통과할 때만 교차점을 반환.
        시작점부터 이미 limit 미만이면 NaN.
        """
        xv=np.asarray(xv,float)
        yv=np.asarray(yv,float)
        valid=np.isfinite(xv)&np.isfinite(yv)
        xv=xv[valid]
        yv=yv[valid]

        if len(xv)<2 or yv[0] < limit:
            return np.nan

        for i in range(1,len(xv)):
            if yv[i] < limit <= yv[i-1]:
                if yv[i]==yv[i-1]:
                    return float(xv[i])
                r=(limit-yv[i-1])/(yv[i]-yv[i-1])
                return float(xv[i-1]+r*(xv[i]-xv[i-1]))
        return np.nan

    before_cross=real_cross(x,y_before,V_LIMIT)
    after_cross=real_cross(x,y_after,V_LIMIT)

    # 보상 전: 실제 교차점이 있을 때만 점 표시.
    if np.isfinite(before_cross) and before_cross<=xmax:
        f.add_trace(go.Scatter(
            x=[before_cross],y=[V_LIMIT],
            mode="markers",
            name=f"보상 전 한계 {before_cross:.1f} MW",
            marker=dict(color="#0070C0",size=8)
        ))
    elif np.isfinite(y_before[0]) and y_before[0] < V_LIMIT:
        f.add_annotation(
            x=4,y=.985,
            text=f"보상 전: 0 MW부터 0.95 pu 미만 ({y_before[0]:.3f} pu)",
            showarrow=False,xanchor="left",
            bgcolor="rgba(255,255,255,.92)",
            bordercolor="#0070C0",
            font=dict(size=10,color="#005A9E")
        )

    # 보상 후는 실제 0.95 교차점만 표시.
    if np.isfinite(after_cross) and after_cross<=xmax:
        f.add_trace(go.Scatter(
            x=[after_cross],y=[V_LIMIT],
            mode="markers",
            name=f"보상 후 한계 {after_cross:.1f} MW",
            marker=dict(color="#1B8A5A",size=9)
        ))
    elif np.isfinite(y_after[0]) and y_after[0] < V_LIMIT:
        f.add_annotation(
            x=4,y=.965,
            text=f"보상 후: 0 MW부터 0.95 pu 미만 ({y_after[0]:.3f} pu)",
            showarrow=False,xanchor="left",
            bgcolor="rgba(255,255,255,.92)",
            bordercolor="#1B8A5A",
            font=dict(size=10,color="#116B46")
        )

    f.update_layout(
        margin=dict(l=65,r=35,t=25,b=92),
        legend=dict(orientation="h",x=.5,xanchor="center",y=-.22,font=dict(size=9)),
        uirevision="keep-user-zoom"
    )
    f.update_xaxes(
        dtick=20,range=[0,xmax],
        title_text="Incremental transfer (MW)"
    )
    f.update_yaxes(dtick=.05)
    return f

def voltage_before_all_fig(sc):
    import plotly.graph_objects as go
    ymin=min(.94,float(min(np.nanmin(sc["before"][b]) for b in BUS_IDS))-0.005)
    ymax=max(.99,float(max(np.nanmax(sc["before"][b]) for b in BUS_IDS))+0.004)
    f=base_fig("전압(pu)",[ymin,ymax],370)
    for b in BUS_IDS:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=sc["before"][b],mode="lines",
            name=f"Bus {b}",line=dict(color=BUS_COLORS[b],width=1.0)
        ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",
        name="0.95 pu 기준",line=dict(color="#C00000",width=1.5,dash="dash")
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(dtick=.01,tickformat=".3f")
    return f

def voltage_after_all_fig(sc):
    import plotly.graph_objects as go
    ymin=min(.94,float(min(np.nanmin(sc["final"][b]) for b in BUS_IDS))-0.005)
    ymax=max(.99,float(max(np.nanmax(sc["final"][b]) for b in BUS_IDS))+0.004)
    f=base_fig("전압(pu)",[ymin,ymax],370)
    for b in BUS_IDS:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=sc["final"][b],mode="lines",
            name=f"Bus {b}",line=dict(color=BUS_COLORS[b],width=1.0)
        ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",
        name="0.95 pu 기준",line=dict(color="#C00000",width=1.5,dash="dash")
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(dtick=.01,tickformat=".3f")
    return f

def all_bus_fig(sc):
    import plotly.graph_objects as go
    f=base_fig("전압(pu)",[.92,1.00],360)
    for b in BUS_IDS:
        f.add_trace(go.Scatter(x=MINUTE_X,y=sc["final"][b],mode="lines",name=f"Bus {b}",
                               line=dict(color=BUS_COLORS[b],width=1)))
    f.add_trace(go.Scatter(x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",name="0.95 pu 기준",
                           line=dict(color=REF_RED,width=1.4,dash="dash")))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(dtick=.01,tickformat=".3f")
    return f

def low_bus_fig(sc):
    import plotly.graph_objects as go
    f=base_fig("전압(pu)",[.93,.975],340)
    low=[]
    for b in BUS_IDS:
        if np.nanmin(sc["before"][b])<V_LIMIT:
            low.append(b)
            f.add_trace(go.Scatter(x=MINUTE_X,y=sc["before"][b],mode="lines",
                                   name=f"Bus {b} {BUS_NAMES[b]}",
                                   line=dict(color=BUS_COLORS[b],width=1.1)))
    f.add_trace(go.Scatter(x=MINUTE_X,y=np.full(1440,V_LIMIT),mode="lines",name="0.95 pu 기준",
                           line=dict(color=REF_RED,width=1.4,dash="dash")))
    if not low:
        f.add_annotation(text="보강 전 0.95 pu 미만 Bus 없음",xref="paper",yref="paper",x=.5,y=.5,showarrow=False)
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(dtick=.005,tickformat=".3f")
    return f



def _planned_case_profile(capacity):
    """
    사용자 계산근거 그대로 용량별 MAIN_TR 비교형상을 만든다.

    원본 PSS/E 24시간 형상 × (계획부하 / 원본 최대값)

    기존공장 계획부하 = 90 MVA 기준:
      10 MVA -> 최대 100 MVA
      15 MVA -> 최대 105 MVA
      20 MVA -> 최대 110 MVA
      30 MVA -> 최대 120 MVA

    30 MVA는 사용자가 준 계산근거의
    '원본 30MVA 형상 × 0.935665... = 최대 120.000 MVA'
    와 같은 방식이다.
    """
    c=float(capacity)
    raw=np.asarray(DATA["mva_cases"][c],float)
    design_existing=float(DATA.get("plan_existing",90.0))
    target_peak=design_existing+c
    peak=float(np.nanmax(raw))
    return raw*(target_peak/peak) if peak>1e-9 else raw.copy()


def _planned_current_profile(total):
    """
    슬라이더 현재값도 위와 같은 계획기준 형상으로 표시.
    30 MVA일 때는 30 MVA 계획기준선과 정확히 겹친다.
    중간값은 인접 계획 Case 사이를 선형보간하고,
    30 MVA 초과는 20/30 형상 차이를 이용해 외삽한다.
    """
    t=max(0.0,float(total))

    refs={
        0.0: np.asarray(DATA["mva_cases"][0.0],float) * (
            float(DATA.get("plan_existing",90.0)) /
            max(float(np.nanmax(DATA["mva_cases"][0.0])),1e-9)
        ),
        10.0:_planned_case_profile(10),
        15.0:_planned_case_profile(15),
        20.0:_planned_case_profile(20),
        30.0:_planned_case_profile(30),
    }

    if t<=0:
        return refs[0.0].copy()

    pts=sorted(refs.keys())
    if t>=30:
        lo,hi=20.0,30.0
    else:
        lo,hi=0.0,10.0
        for a,b in zip(pts[:-1],pts[1:]):
            if a<=t<=b:
                lo,hi=a,b
                break

    w=(t-lo)/(hi-lo) if hi!=lo else 0.0
    y=refs[lo]*(1-w)+refs[hi]*w

    # Plotly에서 선이 끊겨 보이지 않도록 혹시 있는 NaN/Inf를 시간축 보간으로 제거
    y=np.asarray(y,float)
    idx=np.arange(len(y),dtype=float)
    good=np.isfinite(y)
    if np.any(good) and not np.all(good):
        y[~good]=np.interp(idx[~good],idx[good],y[good])
    return y


def capacity_mva_compare(sc,rating):
    import plotly.graph_objects as go

    colors={10:"#0B6FA4",15:"#F47A34",20:"#1B8A5A",30:"#22A6E8"}
    profiles={c:_planned_case_profile(c) for c in [10,15,20,30]}
    current_profile=_planned_current_profile(sc["total_new"])

    all_max=[float(np.nanmax(v)) for v in profiles.values()] + [
        float(np.nanmax(current_profile)),
        float(np.nanmax(sc["daily_dtr"])),
        float(np.nanmax(sc["after_capacity"])),
        float(rating)
    ]
    ymax=max(140.0,math.ceil((max(all_max)+8)/10)*10)

    f=base_fig("용량 (MVA)",[0,ymax],350)

    for c in [10,15,20,30]:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=profiles[c],mode="lines",
            name=f"{c} MVA 계획기준",
            line=dict(color=colors[c],width=1.15),
            connectgaps=True,
            cliponaxis=False
        ))

    # 현재 Slider 선도 같은 계획형상 기준을 사용.
    # 그래서 30 MVA에서는 30 MVA 계획선과 모양/값이 완전히 동일하다.
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=current_profile,mode="lines",
        name=f"현재 신규부하 {sc['total_new']:.0f} MVA",
        line=dict(color="#D62728",width=2.6,dash="solid"),
        connectgaps=True,
        cliponaxis=False
    ))

    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["daily_dtr"],mode="lines",
        name="시간별 DTR 허용용량",
        line=dict(color="#A000C8",width=1.45)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["after_capacity"],mode="lines",
        name=f"자동보강 기준 {sc['plan_dtr']+sc['tr']:.1f} MVA",
        line=dict(color=ACCENT6,width=1.6,dash="dash")
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,rating),mode="lines",
        name=f"{rating:.0f} MVA 정격",
        line=dict(color="#2F7D32",width=1.25)
    ))

    f.add_annotation(
        x=.01,y=.98,xref="paper",yref="paper",
        text="계획기준 형상 = 원본 PSS/E 형상 × (90 + 신규부하) / 원본 최대값",
        showarrow=False,xanchor="left",yanchor="top",
        bgcolor="rgba(255,255,255,.92)",
        bordercolor="#CBD5E1",
        font=dict(size=9,color="#475569")
    )

    if abs(sc["total_new"]-30.0)<1e-9:
        p30=profiles[30]
        scale=120.0/float(np.nanmax(DATA["mva_cases"][30.0]))
        margin=float(np.nanmin(sc["after_capacity"]-p30))
        f.add_annotation(
            x=.99,y=.98,xref="paper",yref="paper",
            text=f"30 MVA: 원본형상 × {scale:.6f}<br>"
                 f"계획 최대 {np.nanmax(p30):.2f} MVA<br>"
                 f"보강기준 {np.nanmin(sc['after_capacity']):.1f} MVA<br>"
                 f"최소여유 {margin:+.2f} MVA",
            showarrow=False,xanchor="right",yanchor="top",
            bgcolor="rgba(255,255,255,.94)",
            bordercolor="#4EA72E" if margin>=0 else "#C00000",
            font=dict(size=10,color="#166534" if margin>=0 else "#C00000")
        )

    f.update_layout(
        margin=dict(l=62,r=34,t=34,b=82),
        uirevision="keep-user-zoom"
    )
    f.update_xaxes(
        tickmode="array",
        tickvals=[f"{h:02d}:00" for h in range(0,24,2)],
        range=[MINUTE_X[0],MINUTE_X[-1]],
        automargin=True
    )
    f.update_yaxes(automargin=True)
    return f

def capacity_feeder_compare(sc):
    import plotly.graph_objects as go
    colors={10:"#0B6FA4",15:"#F47A34",20:"#1B8A5A",30:"#22A6E8"}
    f=base_fig("용량 (MVA)",[0,40],330)
    for c in [10,15,20,30]:
        f.add_trace(go.Scatter(
            x=MINUTE_X,y=DATA["feeder_cases"][float(c)],mode="lines",name=f"{c} MVA",
            line=dict(color=colors[c],width=1.1)
        ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=sc["per_feeder"],mode="lines",
        name=f"현재 {sc['total_new']:.0f}MVA · 자동 {sc['feeder_count']}회선",
        line=dict(color="#111827",width=2.0)
    ))
    f.add_trace(go.Scatter(
        x=MINUTE_X,y=np.full(1440,FEEDER_RATING),mode="lines",
        name="Feeder 정격 35 MVA",line=dict(color="#C000FF",width=1.5)
    ))
    f.update_xaxes(tickmode="array",tickvals=[f"{h:02d}:00" for h in range(0,24,2)])
    f.update_yaxes(dtick=5)
    return f




def compensation_calc_table(sc):
    from dash import html

    head=["구분","계산","원자료","보강 전","보강 후","주의"]
    dtr_src=float(DATA.get("plan_dtr_source",sc["plan_dtr"]))
    formula=DATA.get("tr_formula_text") or f"({sc['design_existing']:.1f}+{sc['total_new']:.1f}-{sc['plan_dtr']:.0f})×(1+{sc['reserve_pct']:.0%})"
    plan_load=sc["design_existing"]+sc["total_new"]
    bank_max=float(np.nanmax(sc["bank"]))
    after_limit=float(np.nanmin(sc["after_capacity"]))
    pre_margin=float(np.nanmin(sc["pre_margin"]))
    post_margin=float(np.nanmin(sc["post_margin"]))

    rows=[
        ["TR 증설",formula,"사용자 조건","-",f"{sc['tr']:.1f} MVA",f"{sc['reserve_pct']*100:.0f}% 여유"],
        ["DTR 적용",f"보강 전 DTR {dtr_src:.3f} MVA를 증설량 산정에 반영","사용자 조건",
         f"최저 DTR {dtr_src:.3f} / 계획값 {sc['plan_dtr']:.0f}","보강 후 재적용하지 않음","DTR 이중반영 방지"],
        [f"신규공장 {sc['total_new']:.0f}MVA 추가 시 MAIN_TR",
         f"PSS/E 형상을 계획부하 {plan_load:.1f} MVA에 맞춰 스케일","PSS/E MAIN_TR 형상",
         f"최대 {bank_max:.3f}",f"최대 {bank_max:.3f}",
         f"{sc['design_existing']:.1f}+{sc['total_new']:.1f}={plan_load:.1f} MVA 계획부하 기준"],
        ["용량 여유",f"보강 전: DTR−MAIN_TR / 보강 후: {after_limit:.1f}−MAIN_TR","시간대별 표시자료",
         f"최소 {pre_margin:+.3f}",f"최소 {post_margin:+.3f}","보강 후 PASS" if post_margin>=0 else "ERROR: 보강 후 기준 초과"],
        ["전압 보상","TR 후에도 0.95 pu 미만 모선이 남으면 Bus 6 Shunt Capacitor 최소 보상",
         "전압_보강전 / 전압_TR보강후 / 전압_Shunt후",
         f"Vmin {float(np.nanmin(sc['v_before'])):.4f} / 저전압 {sc['before_low']['count']}개 · {sc['before_low']['text']}",
         f"TR후 {float(np.nanmin(sc['v_tr'])):.4f} / 저전압 {sc['tr_low']['count']}개 → C {sc['q']:.1f} MVAr → 최종 {float(np.nanmin(sc['v_final'])):.4f} / 저전압 {sc['final_low']['count']}개",
         "TR은 용량보강, Capacitor는 잔여 저전압 보상"],
    ]

    return html.Div(className="card",style={"marginTop":"10px"},children=[
        html.Div(className="head",children=[
            html.Span(f"신규공장 {sc['total_new']:.0f} MVA 추가 — 변압기 증설/전압보상 계산근거"),
            html.Span(f"보강후 허용 {after_limit:.1f} MVA · 최대부하 {bank_max:.1f} MVA")
        ]),
        html.Div(style={"overflowX":"auto"},children=[
            html.Table(className="calc-table",children=[
                html.Thead(html.Tr([html.Th(h) for h in head])),
                html.Tbody([html.Tr([html.Td(c) for c in row]) for row in rows])
            ])
        ])
    ])

def build_app():
    from dash import Dash,dcc,html,Input,Output,State,ctx,no_update

    app=Dash(__name__,suppress_callback_exceptions=True)
    app.title="DTR/PSS-E 신규공장 사전진단"

    CSS="""
    *{box-sizing:border-box}
    body{margin:0;background:#f4f7f9;color:#1f2937;font-family:Segoe UI,'Malgun Gothic',Arial,sans-serif}
    .app{display:grid;grid-template-columns:220px 1fr;min-height:100vh}
    .side{background:linear-gradient(180deg,#0b2c48,#08253e);color:#e5f1f7;padding:18px 12px;position:sticky;top:0;height:100vh}
    .logo{font-size:18px;font-weight:900;padding:6px;margin-bottom:16px}
    .loadbox{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.13);border-radius:10px;padding:12px;margin:7px 0 16px}
    .loadtitle{font-size:12px;font-weight:900;margin-bottom:4px}
    .loadvalue{font-size:29px;font-weight:900;color:#66e3ee;margin:5px 0 9px}
    .navbtn{width:100%;border:0;text-align:left;background:transparent;color:#d7e7ef;padding:12px 11px;border-radius:8px;margin:5px 0;font-size:12px;font-weight:750;cursor:pointer}
    .navbtn.active{background:#0d8b98;color:#fff;font-weight:900}
    .foot{position:absolute;bottom:15px;left:13px;right:13px;font-size:9px;color:#9ab8c8;line-height:1.5}
    .main{padding:0 14px 24px;min-width:0}
    .header{height:68px;background:#fff;border-bottom:1px solid #dbe4ea;display:flex;align-items:center;justify-content:space-between}
    .title{font-size:23px;font-weight:900}.sub{font-size:11px;color:#64748b;margin-left:8px}
    .control{margin-top:10px;background:#fff;border:1px solid #d9e4ea;border-radius:10px;padding:9px 11px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
    .lbl{font-size:10px;font-weight:800;color:#64748b}
    .btn{border:1px solid #b8c7d0;background:#fff;border-radius:7px;padding:8px 12px;font-weight:800;cursor:pointer}
    .play{background:#0d8b98;color:#fff;border-color:#0d8b98}
    .card{background:#fff;border:1px solid #d9e4ea;border-radius:14px;overflow:hidden}
    .head{height:43px;padding:0 12px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #e7eef2;font-weight:850;font-size:13px}
    .grid2,.mid,.graphs,.extra,.voltagepair,.uploadgrid,.detail{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px}
    .page1hero{display:grid;grid-template-columns:minmax(0,1.38fr) minmax(420px,.82fr);gap:10px;margin-top:10px;align-items:stretch}
    .rightstack{display:grid;grid-template-rows:1fr 1fr;gap:10px;min-width:0}
    .page1hero .oneline{height:100%;display:flex;align-items:center;justify-content:center}.page1hero .oneline img{width:100%;max-height:650px;object-fit:contain}
    .stage3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:10px}
    .stagebox{background:#fff;border:1px solid #d9e4ea;border-radius:10px;padding:11px 12px}
    .stagetitle{font-size:11px;color:#64748b;font-weight:800}.stagevalue{font-size:18px;font-weight:900;margin-top:4px}.stagecalc{font-size:9px;color:#64748b;margin-top:4px;line-height:1.4}
    .calc-table{width:100%;border-collapse:collapse;font-size:11px}.calc-table th{background:#4E78C2;color:#fff;padding:8px;border:1px solid #d9e4ea;text-align:center}.calc-table td{padding:8px;border:1px solid #d9e4ea;vertical-align:top;line-height:1.45}.calc-table tr:nth-child(even) td{background:#f8fbfd}
    .mid{grid-template-columns:1.28fr .72fr}
    .oneline{position:relative;padding:8px}.oneline img{display:block;width:100%;height:auto;object-fit:contain}
    .factorytag{position:absolute;left:67%;top:57%;background:rgba(13,139,152,.93);color:#fff;border-radius:7px;padding:6px 9px;font-size:10px;font-weight:900;white-space:nowrap}
    .trtag{position:absolute;left:43%;top:13%;background:rgba(237,125,49,.95);color:#fff;border-radius:7px;padding:6px 9px;font-size:10px;font-weight:900;white-space:nowrap;box-shadow:0 2px 7px rgba(0,0,0,.12)}
    .captag{position:absolute;left:41%;top:63%;background:rgba(22,155,98,.96);color:#fff;border-radius:7px;padding:7px 10px;font-size:10px;font-weight:900;white-space:nowrap;box-shadow:0 2px 7px rgba(0,0,0,.12)}
    .captag.off{background:rgba(100,116,139,.90)}
    .capdot{display:inline-block;width:8px;height:8px;border-radius:50%;background:#fff;margin-right:5px}

    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:9px;margin-top:10px}
    .kpi{background:#fff;border:1px solid #d9e4ea;border-radius:10px;padding:12px}
    .klabel{font-size:11px;color:#64748b}.kvalue{font-size:25px;font-weight:900}.ksub{font-size:10px;color:#64748b;margin-top:4px;line-height:1.45}
    .upload{padding:14px;text-align:center;background:#f8fbfc;border:1px dashed #9fb5c1;border-radius:8px;font-size:11px}
    .settings{margin-top:10px;background:#fff;border:1px solid #d9e4ea;border-radius:10px;padding:10px}
    .settingsgrid{display:grid;grid-template-columns:repeat(5,minmax(120px,1fr));gap:8px;margin-top:9px}
    .note{font-size:10px;color:#64748b;line-height:1.5}
    .status{padding:12px}.row{display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid #edf2f5;font-size:11px}
    .pageintro{margin-top:10px;background:#eef7fa;border:1px solid #cfe4ea;border-radius:10px;padding:10px 12px;font-size:11px;color:#466574}
    table{width:100%;border-collapse:collapse;font-size:10px}th{background:#f8fafc;color:#64748b;text-align:left;padding:7px;position:sticky;top:0}
    td{padding:6px 7px;border-top:1px solid #edf2f5}.scroll{max-height:390px;overflow:auto}
    @media(max-width:1350px){.app{grid-template-columns:1fr}.side{position:relative;height:auto}.foot{display:none}.grid2,.mid,.graphs,.extra,.voltagepair,.uploadgrid,.detail,.stage3,.page1hero{grid-template-columns:1fr}.rightstack{grid-template-rows:auto}.kpis{grid-template-columns:repeat(2,1fr)}}
    """
    app.index_string="""<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>{%title%}</title>{%favicon%}{%css%}<style>__CSS__</style></head>
    <body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>""".replace("__CSS__",CSS)

    if LOAD_ERROR:
        app.layout=html.Div(style={"padding":"30px"},children=[html.H2("입력자료 연결 실패"),html.Pre(LOAD_ERROR)])
        return app

    def inum(id_,value,step=1,width="75px"):
        return dcc.Input(id=id_,type="number",value=value,min=0,step=step,style={"width":width})

    app.layout=html.Div(className="app",children=[
        dcc.Store(id="playing",data=True),
        dcc.Store(id="current-page",data="page1"),
        dcc.Store(id="temp-store"),
        dcc.Store(id="load-store"),
        dcc.Interval(id="clock",interval=1000,n_intervals=0,disabled=False),

        html.Aside(className="side",children=[
            html.Div("▰ DTR / PSS-E",className="logo"),

            html.Div(className="loadbox",children=[
                html.Div("신규공장 부하",className="loadtitle"),
                html.Div(id="load-value",className="loadvalue"),
                dcc.Slider(
                    id="new-load",min=0,max=100,step=1,value=30,
                    marks={0:"0",25:"25",50:"50",75:"75",100:"100"},
                    tooltip={"placement":"bottom","always_visible":False}
                ),
                html.Div("0~100 MVA를 바꾸면 2·3페이지 진단과 자동보상이 즉시 재계산됩니다.",
                         style={"fontSize":"9px","color":"#b6d1df","marginTop":"8px","lineHeight":"1.4"})
            ]),

            html.Button("① 계통 · 온도 · DTR",id="nav-page1",className="navbtn active",n_clicks=0),
            html.Button("② 실시간 진단 · 자동보상",id="nav-page2",className="navbtn",n_clicks=0),
            html.Button("③ 상세분석 · 데이터입력",id="nav-page3",className="navbtn",n_clicks=0),

            html.Div(className="foot",children=[
                html.Div("1페이지: 계통 + 연간/극한일"),
                html.Div("2페이지: Feeder/MVA/PV/전압"),
                html.Div("3페이지: 상세·업로드·로그")
            ])
        ]),

        html.Main(className="main",children=[
            html.Div(className="header",children=[
                html.Div([html.Span("신규공장 전력인프라 DTR/PSS-E 사전진단 시스템",className="title"),
                          html.Span("V22 · MAINTR연속선",className="sub")]),
                html.Div(id="model-time",style={"fontWeight":"800","fontSize":"12px","color":"#475569"})
            ]),

            html.Div(className="control",children=[
                html.Button("⏸ 정지",id="play",n_clicks=0,className="btn play"),
                html.Button("처음",id="reset",n_clicks=0,className="btn"),
                html.Span("1초마다 1시간 이동",className="note"),
                html.Span("·",className="note"),
                html.Span("온도·부하 파일 입력 시 전체 모델 즉시 재계산",className="note"),
            ]),
            html.Div(style={"padding":"5px 4px","marginTop":"4px"},children=[
                dcc.Slider(
                    id="hour-slider",min=0,max=8759,step=1,value=0,
                    marks={0:"시작",2190:"3개월",4380:"6개월",6570:"9개월",8759:"1년"}
                )
            ]),

            # ======================================================
            # PAGE 1
            # ======================================================
            html.Div(id="page1",children=[
                html.Div("첫 페이지는 PSS/E 원본 계통도와 1년 외기온도·DTR, 그리고 현재 외기온도 기준 가장 더운 하루의 온도/DTR만 보여줍니다.",
                         className="pageintro"),

                html.Div(className="page1hero",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("PSS/E 계통도"),html.Span("원본 이미지 그대로")]),
                        html.Div(className="oneline",children=[
                            html.Img(src=PSS_IMAGE_URI),
                            html.Div(id="factory-tag",className="factorytag"),
                            html.Div(id="tr-tag",className="trtag"),
                            html.Div(id="cap-tag",className="captag")
                        ])
                    ]),
                    html.Div(className="rightstack",children=[
                        html.Div(className="card",children=[
                            html.Div(className="head",children=[html.Span("1년 외기온도"),html.Span("시간단위 커서 이동")]),
                            dcc.Graph(id="annual-temp",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                        ]),
                        html.Div(className="card",children=[
                            html.Div(className="head",children=[html.Span("1년 DTR 허용용량 vs 기존공장 MAIN_TR"),html.Span("시간단위 커서 이동")]),
                            dcc.Graph(id="annual-dtr",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                        ]),
                    ])
                ]),

                html.Div(className="grid2",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("연중 가장 더운 1일 — 외기온도"),html.Span(id="hot-day-label")]),
                        dcc.Graph(id="hot-temp",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("연중 가장 더운 1일 — DTR"),html.Span("온도 변화 즉시 반영")]),
                        dcc.Graph(id="hot-dtr",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                ]),
            ]),

            # ======================================================
            # PAGE 2
            # ======================================================
            html.Div(id="page2",style={"display":"none"},children=[
                html.Div("사용자가 준 계산근거 그대로 ① 보강 전 → ② DTR 산정식으로 계산한 병렬 TR 추가 후 → ③ TR 후에도 Vmin<0.95이면 Bus 6 Shunt Capacitor 추가 후를 각각 따로 보여줍니다.",
                         className="pageintro"),

                html.Div(id="comp-banner",style={
                    "marginTop":"10px","padding":"12px 14px","borderRadius":"10px",
                    "background":"#ffffff","border":"1px solid #d9e4ea",
                    "fontSize":"12px","fontWeight":"850"
                }),
                html.Div(id="comp-calc-table"),

                html.Div(id="kpis",className="kpis"),

                html.Div(className="graphs",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("Feeder 정격"),html.Span("자동 회선증설 · Zoom 가능")]),
                        dcc.Graph(id="feeder-fig",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("PV 안정도 — 보상 전 vs 보상 후"),html.Span("0~140 MW · 추정모델")]),
                        dcc.Graph(id="pv-fig",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                ]),

                html.Div(className="graphs",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("MVA / DTR — 보강 전"),html.Span("BEFORE")]),
                        dcc.Graph(id="mva-before-fig",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("MVA / DTR — 자동보강 후"),html.Span("AFTER")]),
                        dcc.Graph(id="mva-after-fig",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                ]),

                html.Div(id="stage-summary",className="stage3"),

                html.Div(className="stage3",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[
                            html.Span("① 전압 — 보강 전"),
                            html.Span(id="stage-before-label")
                        ]),
                        dcc.Graph(
                            id="volt-stage-before",
                            config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,
                                    "modeBarButtonsToRemove":["lasso2d","select2d","toImage"]}
                        )
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[
                            html.Span("② 전압 — 병렬 TR 추가 후"),
                            html.Span(id="stage-tr-label")
                        ]),
                        dcc.Graph(
                            id="volt-stage-tr",
                            config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,
                                    "modeBarButtonsToRemove":["lasso2d","select2d","toImage"]}
                        )
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[
                            html.Span("③ 전압 — 커패시터 추가 후"),
                            html.Span(id="stage-cap-label")
                        ]),
                        dcc.Graph(
                            id="volt-stage-cap",
                            config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,
                                    "modeBarButtonsToRemove":["lasso2d","select2d","toImage"]}
                        )
                    ]),
                ]),
            ]),

            # ======================================================
            # PAGE 3
            # ======================================================
            html.Div(id="page3",style={"display":"none"},children=[
                html.Div("세 번째 페이지에는 용량별 MAIN_TR/Feeder 비교, 선택일 외기온도, 외기온도·기존공장 부하파일 입력, DTR 계수, 자동보상 결과와 로그를 모았습니다.",
                         className="pageintro"),

                html.Div(className="extra",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children="신규공장 용량별 MAIN_TR 비교"),
                        dcc.Graph(id="mva-compare-fig",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children="신규공장 용량별 Feeder 비교"),
                        dcc.Graph(id="feeder-compare-fig",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children=[html.Span("선택 모델시각의 하루 외기온도"),html.Span(id="temp-source")]),
                        dcc.Graph(id="daily-temp",config={"displayModeBar":True,"scrollZoom":True,"displaylogo":False,"modeBarButtonsToRemove":["lasso2d","select2d","toImage"]})
                    ]),
                ]),

                html.Div(className="uploadgrid",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children="지역 외기온도 입력"),
                        dcc.Upload(id="upload-temp",children=html.Div("Excel / CSV / 기상청 ZIP 드래그 또는 클릭"),className="upload",multiple=False),
                        html.Div(id="temp-upload-status",className="note",style={"padding":"8px 12px"})
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children="기존공장 부하패턴 입력"),
                        dcc.Upload(id="upload-load",children=html.Div("Excel / CSV 드래그 또는 클릭"),className="upload",multiple=False),
                        html.Div(id="load-upload-status",className="note",style={"padding":"8px 12px"})
                    ])
                ]),

                html.Details(className="settings",children=[
                    html.Summary("DTR 실제 변압기 모델 계수 / 계획값 수정"),
                    html.Div(className="settingsgrid",children=[
                        html.Div([html.Div("TR 정격 MVA",className="lbl"),inum("rating",105,.1)]),
                        html.Div([html.Div("Hot-spot 한계 °C",className="lbl"),inum("hotspot",120,.1)]),
                        html.Div([html.Div("Top-oil 상승 °C",className="lbl"),inum("oil-rise",60,.1)]),
                        html.Div([html.Div("Winding 상승 °C",className="lbl"),inum("wind-rise",30,.1)]),
                        html.Div([html.Div("손실식 분자계수",className="lbl"),inum("loss-num",5,.1)]),
                        html.Div([html.Div("손실식 분모",className="lbl"),inum("loss-den",6,.1)]),
                        html.Div([html.Div("Oil 지수",className="lbl"),inum("oil-exp",.8,.05)]),
                        html.Div([html.Div("Winding 지수",className="lbl"),inum("wind-exp",1.6,.05)]),
                        html.Div([html.Div("기존공장 계획 MVA",className="lbl"),inum("plan-existing",90,.1)]),
                        html.Div([html.Div("TR 여유율",className="lbl"),inum("reserve",.10,.01)]),
                        html.Div([html.Div("PV 요구여유 MW (선택)",className="lbl"),
                                  dcc.Input(id="pv-required",type="number",value=None,min=0,step=1,placeholder="미설정",style={"width":"85px"})]),
                    ]),
                    html.Div("열모델: 120 = T(t) + 60×((5K²+1)/6)^0.8 + 30×K^1.6 / DTR(t)=105×Kmax(t)",
                             className="note",style={"marginTop":"8px"})
                ]),

                html.Div(className="detail",children=[
                    html.Div(className="card",children=[
                        html.Div(className="head",children="현재 자동보상"),
                        html.Div(id="status",className="status")
                    ]),
                    html.Div(className="card",children=[
                        html.Div(className="head",children="24시간 로그 — 자동보상 후"),
                        html.Div(className="scroll",children=[
                            html.Table([
                                html.Thead(html.Tr([html.Th(x) for x in ["시간","MAIN_TR","DTR","보강기준","Feeder/회선","Vmin","판정"]])),
                                html.Tbody(id="log")
                            ])
                        ])
                    ])
                ])
            ])
        ])
    ])

    # ---------- Page navigation ----------
    @app.callback(
        Output("current-page","data"),
        Input("nav-page1","n_clicks"),Input("nav-page2","n_clicks"),Input("nav-page3","n_clicks"),
        State("current-page","data"),prevent_initial_call=True
    )
    def change_page(_,__,___,current):
        trig=ctx.triggered_id
        if trig=="nav-page1": return "page1"
        if trig=="nav-page2": return "page2"
        if trig=="nav-page3": return "page3"
        return current or "page1"

    @app.callback(
        Output("page1","style"),Output("page2","style"),Output("page3","style"),
        Output("nav-page1","className"),Output("nav-page2","className"),Output("nav-page3","className"),
        Input("current-page","data")
    )
    def show_page(page):
        page=page or "page1"
        return (
            {} if page=="page1" else {"display":"none"},
            {} if page=="page2" else {"display":"none"},
            {} if page=="page3" else {"display":"none"},
            "navbtn active" if page=="page1" else "navbtn",
            "navbtn active" if page=="page2" else "navbtn",
            "navbtn active" if page=="page3" else "navbtn",
        )

    # ---------- Playback ----------
    @app.callback(
        Output("playing","data"),Output("clock","disabled"),Output("play","children"),
        Input("play","n_clicks"),State("playing","data"),prevent_initial_call=True
    )
    def toggle(_,playing):
        new=not bool(playing)
        return new,not new,("⏸ 정지" if new else "▶ 재생")

    @app.callback(
        Output("hour-slider","value"),
        Input("clock","n_intervals"),Input("reset","n_clicks"),
        State("hour-slider","value"),State("playing","data"),prevent_initial_call=True
    )
    def advance(_,__,hour,playing):
        if ctx.triggered_id=="reset":
            return 0
        if ctx.triggered_id=="clock" and playing:
            return (int(hour or 0)+1)%8760
        return int(hour or 0)

    # ---------- Upload ----------
    @app.callback(
        Output("temp-store","data"),Output("temp-upload-status","children"),
        Input("upload-temp","contents"),State("upload-temp","filename"),prevent_initial_call=True
    )
    def upload_temp(contents,filename):
        if not contents:
            return no_update,no_update
        try:
            raw=base64.b64decode(contents.split(",",1)[1])
            data=parse_temperature(raw,filename)
            s=data.get("summary",{})
            msg=(f"적용 완료: {filename} | {s.get('rows','?')}시간 | "
                 f"{s.get('min',0):.1f}~{s.get('max',0):.1f}°C | "
                 "1·2·3페이지 DTR/보상량에 즉시 반영")
            return data,msg
        except Exception as e:
            return None,f"적용 실패: {e}"

    @app.callback(
        Output("load-store","data"),Output("load-upload-status","children"),
        Input("upload-load","contents"),State("upload-load","filename"),prevent_initial_call=True
    )
    def upload_load(contents,filename):
        if not contents:
            return no_update,no_update
        try:
            raw=base64.b64decode(contents.split(",",1)[1])
            data=parse_load(raw,filename)
            s=data.get("summary",{})
            msg=(f"적용 완료: {filename} | {s.get('min',0):.1f}~{s.get('max',0):.1f} MVA | "
                 "MAIN_TR·전압·자동 TR/Shunt·로그에 즉시 반영")
            return data,msg
        except Exception as e:
            return None,f"적용 실패: {e}"

    # ---------- All calculations ----------
    @app.callback(
        Output("load-value","children"),
        Output("model-time","children"),
        Output("annual-temp","figure"),Output("annual-dtr","figure"),
        Output("factory-tag","children"),
        Output("tr-tag","children"),Output("cap-tag","children"),Output("cap-tag","className"),
        Output("hot-day-label","children"),Output("hot-temp","figure"),Output("hot-dtr","figure"),
        Output("kpis","children"),Output("comp-banner","children"),Output("comp-calc-table","children"),
        Output("feeder-fig","figure"),Output("pv-fig","figure"),
        Output("mva-before-fig","figure"),Output("mva-after-fig","figure"),
        Output("stage-summary","children"),
        Output("volt-stage-before","figure"),Output("stage-before-label","children"),
        Output("volt-stage-tr","figure"),Output("stage-tr-label","children"),
        Output("volt-stage-cap","figure"),Output("stage-cap-label","children"),
        Output("mva-compare-fig","figure"),Output("feeder-compare-fig","figure"),
        Output("daily-temp","figure"),Output("temp-source","children"),
        Output("status","children"),Output("log","children"),
        Input("hour-slider","value"),Input("new-load","value"),
        Input("temp-store","data"),Input("load-store","data"),
        Input("rating","value"),Input("hotspot","value"),Input("oil-rise","value"),Input("wind-rise","value"),
        Input("loss-num","value"),Input("loss-den","value"),Input("oil-exp","value"),Input("wind-exp","value"),
        Input("plan-existing","value"),Input("reserve","value"),Input("pv-required","value")
    )
    def refresh(hour_idx,new_load,temp_store,load_store,rating,hotspot,oil_rise,wind_rise,
                loss_num,loss_den,oil_exp,wind_exp,plan_existing,reserve,pv_required):

        defaults=[105,120,60,30,5,6,.8,1.6,90,.10]
        raw=[rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp,plan_existing,reserve]
        vals=[d if v is None else float(v) for v,d in zip(raw,defaults)]
        rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp,plan_existing,reserve=vals
        pv_req=None if pv_required in (None,"") else float(pv_required)
        load=max(0.0,min(100.0,float(new_load or 0)))

        a=annual_data(temp_store,load_store)
        hour_idx=max(0,min(len(a)-1,int(hour_idx or 0)))
        day_idx=min(364,hour_idx//24)

        sc=scenario(
            load,a,day_idx,
            (rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp,plan_existing,reserve,pv_req),
            load_store
        )

        hot_date,hot_temp,hot_dtr,hot_source=hottest_day_profiles(
            a,rating,hotspot,oil_rise,wind_rise,loss_num,loss_den,oil_exp,wind_exp
        )

        feeder_max=float(np.nanmax(sc["per_feeder"])) if sc["feeder_count"] else 0.0
        feeder_ok=feeder_max<=FEEDER_RATING+1e-6
        vmin=float(np.nanmin(sc["v_final"]))
        volt_ok=vmin>=V_LIMIT-1e-6
        pre_margin=float(np.nanmin(sc["pre_margin"]))
        post_margin=float(np.nanmin(sc["post_margin"]))
        mva_ok=post_margin>=-1e-6
        pv_limit=pv_cross(PV_MODEL_X,sc["pv_final"],V_LIMIT)
        pv_ok=None if pv_req is None or not np.isfinite(pv_limit) else pv_limit>=pv_req

        def kpi(title,value,unit_,sub,ok):
            color="#0f766e" if ok is None else ("#16a34a" if ok else "#dc2626")
            word="" if ok is None else ("PASS · " if ok else "CHECK · ")
            return html.Div(className="kpi",children=[
                html.Div(title,className="klabel"),
                html.Div([html.Span(value,className="kvalue",style={"color":color}),
                          html.Span(unit_,style={"fontSize":"11px","marginLeft":"4px","color":"#64748b"})]),
                html.Div(word+sub,className="ksub")
            ])

        kpis=[
            kpi("Feeder 정격",f"{feeder_max:.1f}","MVA",f"자동 {sc['feeder_count']}회선 · 회선당 ≤35",feeder_ok),
            kpi("전압",f"{vmin:.4f}","pu",f"최소 C 보상 {sc['q']:.1f} MVAr",volt_ok),
            kpi("MVA",f"{float(np.nanmax(sc['bank'])):.1f}","MVA",f"DTR식 TR +{sc['tr']:.1f} MVA",mva_ok),
            kpi("PV",("—" if not np.isfinite(pv_limit) else f"{pv_limit:.1f}"),"MW",
                ("요구여유 기준 미설정" if pv_req is None else f"요구여유 {pv_req:.1f} MW"),pv_ok)
        ]

        overall=feeder_ok and volt_ok and mva_ok
        status=[
            html.Div(className="row",children=[html.Span("신규부하 Slider"),html.B(f"{load:.0f} MVA")]),
            html.Div(className="row",children=[html.Span("계획 DTR"),html.B(f"{sc['plan_dtr']:.0f} MVA")]),
            html.Div(className="row",children=[html.Span("기존공장 계획부하"),html.B(f"{sc['design_existing']:.1f} MVA")]),
            html.Div(className="row",children=[html.Span("DTR 부족용량"),html.B(f"{sc['shortage']:.1f} MVA")]),
            html.Div(className="row",children=[html.Span("병렬 TR 산정"),html.B(f"{sc['shortage']:.1f} × (1+{reserve:.0%}) = {sc['tr']:.1f} MVA")]),
            html.Div(className="row",children=[html.Span("자동 Feeder"),html.B(f"{sc['feeder_count']} 회선 × 35 MVA")]),
            html.Div(className="row",children=[html.Span("① 보강 전 Vmin"),html.B(f"{np.nanmin(sc['v_before']):.4f} pu")]),
            html.Div(className="row",children=[html.Span("② 병렬 TR 추가"),html.B(f"+{sc['tr']:.1f} MVA → {np.nanmin(sc['v_tr']):.4f} pu")]),
            html.Div(className="row",children=[
                html.Span("③ Shunt Capacitor 추가"),
                html.B(
                    f"미설치 · TR만으로 PASS"
                    if sc["q"]<=0 else
                    f"Bus 6 +{sc['q']:.1f} MVAr → {np.nanmin(sc['v_final']):.4f} pu"
                )
            ]),
            html.Div(className="row",children=[html.Span("저전압 모선 수"),html.B(f"보강 전 {sc['before_low']['count']}개 / TR 후 {sc['tr_low']['count']}개 / 최종 {sc['final_low']['count']}개")]),
            html.Div(className="row",children=[html.Span("보강 전 DTR 최소여유"),html.B(f"{pre_margin:+.2f} MVA")]),
            html.Div(className="row",children=[html.Span("보강 후 최소여유"),html.B(f"{post_margin:+.2f} MVA")]),
            html.Div(className="row",children=[html.Span("최종 Vmin"),html.B(f"{vmin:.4f} pu")]),
            html.Div(className="row",children=[html.Span("자동보상 결과"),html.B("PASS" if overall else "CHECK",style={"color":"#16a34a" if overall else "#dc2626"})]),
            html.Div(className="note",children=[
                html.B("실시간 공식화: "),
                "온도 → DTR(t) 재계산 / 기존공장 부하 → MAIN_TR와 Bus전압 민감도 재계산 / ",
                "TR = [기존공장 계획부하 + 신규부하 - 계획 DTR] × (1+여유율) / ",
                "TR 용량은 전압 때문에 키우지 않음 / TR 후 Vmin<0.95일 때만 Bus 6 최소 Shunt MVAr를 산정."
            ])
        ]

        logs=[]
        for h in range(24):
            i=h*60
            ok=(sc["bank"][i] <= sc["after_capacity"][i]+1e-6 and
                sc["per_feeder"][i] <= FEEDER_RATING+1e-6 and
                sc["v_final"][i] >= V_LIMIT-1e-6)
            logs.append(html.Tr([
                html.Td(f"{h:02d}:00"),
                html.Td(f"{sc['bank'][i]:.1f}"),
                html.Td(f"{sc['daily_dtr'][i]:.1f}"),
                html.Td(f"{sc['after_capacity'][i]:.1f}"),
                html.Td(f"{sc['per_feeder'][i]:.1f}"),
                html.Td(f"{sc['v_final'][i]:.4f}"),
                html.Td("PASS" if ok else "CHECK",style={"fontWeight":"800","color":"#16a34a" if ok else "#dc2626"})
            ]))

        tag=("신규부하 0 MVA" if load<=0 else f"NEW LOAD · {load:.0f} MVA")
        model_ts=pd.Timestamp(a.iloc[hour_idx]["dt"])

        return (
            f"{load:.0f} MVA",
            f"모델 시각 {model_ts.strftime('%Y-%m-%d %H:%M')}",
            annual_temp_fig(a,hour_idx),annual_dtr_fig(a,sc["annual_dtr"],hour_idx),
            tag,
            ("병렬 TR 미설치" if sc["tr"]<=0 else f"병렬 TR +{sc['tr']:.1f} MVA"),
            ("● Shunt C 미설치" if sc["q"]<=0 else f"● Bus 6 Shunt C +{sc['q']:.1f} MVAr"),
            ("captag off" if sc["q"]<=0 else "captag"),
            hot_date.strftime("%Y-%m-%d"),hottest_temp_fig(hot_date,hot_temp),hottest_dtr_fig(hot_date,hot_dtr,rating),
            kpis,
            (
                [
                    html.Span(
                        f"계산근거 {DATA.get('tr_formula_text') or ''} = TR +{sc['tr']:.1f} MVA",
                        style={"color":"#ED7D31","marginRight":"18px"}
                    ),
                    html.Span(f"TR 후 Vmin {np.nanmin(sc['v_tr']):.4f} pu",style={"color":"#ED7D31","marginRight":"18px"}),
                    html.Span(
                        ("→ Capacitor 미설치" if sc["q"]<=0 else f"→ Bus 6 Shunt Capacitor +{sc['q']:.1f} MVAr"),
                        style={"color":"#64748B" if sc["q"]<=0 else "#169B62","marginRight":"18px"}
                    ),
                    html.Span(
                        f"TR 후 저전압 {sc['tr_low']['count']}개 → 최종 {sc['final_low']['count']}개",
                        style={"color":"#64748b","marginRight":"18px"}
                    ),
                    html.Span(
                        f"최종 Vmin {np.nanmin(sc['v_final']):.4f} pu",
                        style={"color":"#16a34a" if np.nanmin(sc["v_final"])>=V_LIMIT else "#dc2626"}
                    )
                ]
            ),
            compensation_calc_table(sc),
            feeder_fig(sc),pv_fig(sc),
            mva_before_fig(sc,rating),mva_after_fig(sc,rating),
            [
                html.Div(className="stagebox",children=[
                    html.Div("① 보강 전",className="stagetitle"),
                    html.Div(f"Vmin {np.nanmin(sc['v_before']):.4f} pu",className="stagevalue",
                             style={"color":"#dc2626" if np.nanmin(sc["v_before"])<V_LIMIT else "#16a34a"}),
                    html.Div(f"저전압 모선 {sc['before_low']['count']}개 · {sc['before_low']['text']}",className="stagecalc")
                ]),
                html.Div(className="stagebox",children=[
                    html.Div("② 병렬 TR 추가",className="stagetitle"),
                    html.Div(f"+{sc['tr']:.1f} MVA → {np.nanmin(sc['v_tr']):.4f} pu",className="stagevalue",
                             style={"color":"#F47A34"}),
                    html.Div(
                        f"{DATA.get('tr_formula_text') or ('(' + format(sc['design_existing'], '.1f') + '+' + format(sc['total_new'], '.1f') + '-' + format(sc['plan_dtr'], '.0f') + ')×(1+여유율)')} / 저전압 {sc['tr_low']['count']}개 · {sc['tr_low']['text']}",
                        className="stagecalc"
                    )
                ]),
                html.Div(className="stagebox",children=[
                    html.Div("③ 커패시터 추가",className="stagetitle"),
                    html.Div(
                        ("미설치" if sc["q"]<=0 else f"Bus 6 +{sc['q']:.1f} MVAr → {np.nanmin(sc['v_final']):.4f} pu"),
                        className="stagevalue",
                        style={"color":"#64748B" if sc["q"]<=0 else "#1B8A5A"}
                    ),
                    html.Div(
                        (f"TR만으로 PASS / 최종 저전압 모선 {sc['final_low']['count']}개"
                         if sc["q"]<=0 else
                         f"TR 후 남은 저전압 {sc['tr_low']['count']}개를 Bus 6 Shunt {sc['q']:.1f} MVAr로 보상 → 최종 {sc['final_low']['count']}개"),
                        className="stagecalc"
                    )
                ])
            ],
            voltage_three_stage_fig(sc,"before"),
            f"Vmin {np.nanmin(sc['v_before']):.4f}",
            voltage_three_stage_fig(sc,"tr"),
            f"+{sc['tr']:.1f} MVA · Vmin {np.nanmin(sc['v_tr']):.4f}",
            voltage_three_stage_fig(sc,"cap"),
            (
                "미설치 · TR만으로 PASS"
                if sc["q"]<=0 else
                f"Bus 6 +{sc['q']:.1f} MVAr · Vmin {np.nanmin(sc['v_final']):.4f}"
            ),
            capacity_mva_compare(sc,rating),capacity_feeder_compare(sc),
            daily_temp_fig(sc),sc["temp_source"],
            status,logs
        )

    return app

def run_check():
    print("="*96)
    print("DTR/PSS-E SYSTEM V22 - TABLE / DTR POST-LIMIT CHECK")
    print("="*96)
    print("온도 :",TEMP_FILE.name if TEMP_FILE else "NOT FOUND")
    print("MVA  :",MVA_FILE.name if MVA_FILE else "NOT FOUND")
    print("전압 :",VOLT_FILE.name if VOLT_FILE else "NOT FOUND")
    print("보상 :",COMP_FILE.name if COMP_FILE else "NOT FOUND")
    if LOAD_ERROR:
        print("[ERROR]",LOAD_ERROR)
        return 1

    a=annual_data()
    exdate=pd.Timestamp(DATA["extreme"]["dt"].iloc[0]).normalize()
    day=max(0,min(364,(exdate-pd.Timestamp(a["dt"].iloc[0]).normalize()).days))

    for load in [0,30,60,100]:
        sc=scenario(load,a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None))
        mva_ok=np.nanmin(sc["post_margin"])>=-1e-6
        feeder_ok=(sc["feeder_count"]==0 or np.nanmax(sc["per_feeder"])<=35+1e-6)
        v_ok=np.nanmin(sc["v_final"])>=.95-1e-6
        print(
            f"{load:>3} MVA | TR={sc['tr']:.1f} | Feeders={sc['feeder_count']} | "
            f"Q={sc['q']:.1f} | BankMax={np.nanmax(sc['bank']):.2f} | "
            f"Margin={np.nanmin(sc['post_margin']):+.2f} | "
            f"Feeder/Ckt={np.nanmax(sc['per_feeder']) if sc['feeder_count'] else 0:.2f} | "
            f"Vmin={np.nanmin(sc['v_final']):.4f} | "
            f"{'PASS' if (mva_ok and feeder_ok and v_ok) else 'CHECK'}"
        )

    # Uploaded KMA zip parser check when the file exists in the same folder.
    z=ROOT/"TalkFile_HD 기온 자료(3).zip"
    if z.exists():
        ts=parse_temperature(z.read_bytes(),z.name)
        print(
            f"온도 ZIP parser: {ts['summary']['rows']} h / "
            f"{ts['summary']['min']:.1f}~{ts['summary']['max']:.1f}°C"
        )

    # Same-source workbook can be used as a dated load upload.
    if TEMP_FILE and TEMP_FILE.exists():
        ls=parse_load(TEMP_FILE.read_bytes(),TEMP_FILE.name)
        print(
            f"부하 parser: kind={ls['kind']} / "
            f"{ls['summary']['min']:.1f}~{ls['summary']['max']:.1f} MVA"
        )

    # 동적 공식 반응 검증: 같은 30MVA에서 온도 +5°C / 기존부하 +10MVA
    base_sc=scenario(30,a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None))

    temp_store={
        "dt":[pd.Timestamp(x).isoformat() for x in a["dt"]],
        "temp":(a["temp"].to_numpy(float)+5.0).tolist()
    }
    hot_a=annual_data(temp_store,None)
    hot_sc=scenario(30,hot_a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None))

    load_store={
        "kind":"dated",
        "dt":[pd.Timestamp(x).isoformat() for x in a["dt"]],
        "load":(a["base_main"].to_numpy(float)+10.0).tolist()
    }
    heavy_a=annual_data(None,load_store)
    heavy_sc=scenario(30,heavy_a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None),load_store)

    print(
        f"온도 +5°C 반응 | 계획DTR {base_sc['plan_dtr']:.0f}->{hot_sc['plan_dtr']:.0f} | "
        f"TR {base_sc['tr']:.1f}->{hot_sc['tr']:.1f}"
    )
    print(
        f"기존부하 +10MVA 반응 | BankMax {np.nanmax(base_sc['bank']):.2f}->{np.nanmax(heavy_sc['bank']):.2f} | "
        f"TR {base_sc['tr']:.1f}->{heavy_sc['tr']:.1f} | "
        f"보강전 Vmin {np.nanmin(base_sc['v_before']):.4f}->{np.nanmin(heavy_sc['v_before']):.4f}"
    )

    # 사용자 최종파일과 동일한 30 MVA 산정식 검증
    sc30=scenario(30,a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None))
    print(
        f"30MVA 산정식 | ({sc30['design_existing']:.1f}+30-{sc30['plan_dtr']:.0f})"
        f"×1.10 = TR {sc30['tr']:.1f} MVA | "
        f"TR후 Vmin={np.nanmin(sc30['v_tr']):.4f} | "
        f"최소 C={sc30['q']:.1f} MVAr | 최종 Vmin={np.nanmin(sc30['v_final']):.4f}"
    )

    print(
        f"계산근거 파일: {DATA['comp_source_name']} | "
        f"TR식={DATA.get('tr_formula_text')} | TR={DATA['tr_ref']:.1f} MVA | "
        f"Shunt={DATA['shunt_ref']:.1f} MVAr"
    )
    if DATA.get("exact_stage_min"):
        print(
            "원본 30MVA 전압 3단계 | "
            f"보강전={DATA['exact_stage_min'].get('before',float('nan')):.4f} | "
            f"TR후={DATA['exact_stage_min'].get('tr',float('nan')):.4f} | "
            f"C후={DATA['exact_stage_min'].get('cap',float('nan')):.4f}"
        )
    sc30=scenario(30,a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None))
    print(
        f"30MVA 보강후 검증 | 계획부하={sc30['design_existing']+sc30['total_new']:.1f} | "
        f"MAIN_TR max={np.nanmax(sc30['bank']):.3f} | 허용={np.nanmin(sc30['after_capacity']):.1f} | "
        f"최소여유={np.nanmin(sc30['post_margin']):+.3f} MVA | "
        f"TR후 저전압={sc30['tr_low']['count']}개 | C={sc30['q']:.1f} MVAr | 최종 저전압={sc30['final_low']['count']}개"
    )
    assert np.nanmin(sc30['post_margin']) >= -1e-9, "보강 후 MAIN_TR가 허용용량을 초과합니다."
    sc30=scenario(30,a,day,(105,120,60,30,5,6,.8,1.6,90,.10,None))
    cmp30=_planned_case_profile(30)
    print(
        f"30MVA 정합 확인 | 비교그래프 max={np.nanmax(cmp30):.3f} | "
        f"현재 bank max={np.nanmax(sc30['bank']):.3f} | "
        f"자동보강 기준={np.nanmin(sc30['after_capacity']):.1f} | "
        f"최소여유={np.nanmin(sc30['after_capacity']-cmp30):+.3f}"
    )
    print(
        f"PV 2선 확인 | 보상전 0.95 한계={pv_cross(PV_MODEL_X,sc30['pv_before'],V_LIMIT):.1f} MW | "
        f"보상후 0.95 한계={pv_cross(PV_MODEL_X,sc30['pv_final'],V_LIMIT):.1f} MW"
    )

    p10=_planned_case_profile(10)
    p20=_planned_case_profile(20)
    p30=_planned_case_profile(30)
    cur30=_planned_current_profile(30)
    raw30=np.asarray(DATA["mva_cases"][30.0],float)
    s30=120.0/float(np.nanmax(raw30))
    print(
        f"계획형상 peak | 10={np.nanmax(p10):.3f} / "
        f"20={np.nanmax(p20):.3f} / 30={np.nanmax(p30):.3f} MVA"
    )
    print(
        f"30MVA 스케일 | 원본max={np.nanmax(raw30):.3f} × {s30:.6f} "
        f"= {np.nanmax(p30):.3f} MVA"
    )
    print(
        f"30MVA 현재선-계획선 최대차이={np.nanmax(np.abs(cur30-p30)):.9f} MVA"
    )

    print("[OK] MAIN_TR 현재부하 실선/연속표시 / 형상정합 / PV 교차점 수정 확인")
    return 0



# ----------------------------------------------------------------------
# Public web deployment entrypoint
# Gunicorn/Render imports `server`; local execution also remains possible.
# ----------------------------------------------------------------------
app = build_app()
server = app.server

if __name__=="__main__":
    if "--check" in sys.argv:
        raise SystemExit(run_check())

    port = int(os.environ.get("PORT", "8050"))
    print("="*96)
    print("신규공장 전력인프라 DTR/PSS-E 사전진단 시스템 V22 WEB")
    print(f"서버: http://0.0.0.0:{port}")
    print("신규부하: 왼쪽 Slider 0~100 MVA / 좌측 3페이지 메뉴")
    print("자동보상: TR + Feeder 회선 + Shunt")
    print("="*96)
    app.run(debug=False, host="0.0.0.0", port=port)
