#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
전체 4,954개 단어를 필터형 XLSX로 만든다.
- 강화 완료분(enriched_cumulative.json): 전 컬럼(코어/어원/의미파생/유반의어/예문) 채움
- 미강화분(cards_raw.json): 표제어·품사·CEFR·뜻(원본)·예문·출처 기본 정보만 채움
강화가 진행될수록 자동으로 더 풍부해진다.

사용: python3 make_full_xlsx.py [out.xlsx]
"""
import json, os, sys, re
import build_xlsx

HERE = os.path.dirname(os.path.abspath(__file__))

def clean_pos(raw_pos, hw, cefr):
    """원본 '품사/원본' 필드에서 표제어·CEFR을 떼고 품사만 남긴다."""
    if not raw_pos:
        return ""
    s = raw_pos
    # 끝의 CEFR 등급 제거
    s = re.sub(r'\b[ABC][12]\b', '', s)
    # 표제어(콤마 분리 변형 포함) 앞부분 제거
    head = hw.replace(' / ', ', ')
    for token in re.split(r'[,/]', head):
        token = token.strip()
        if token:
            s = re.sub(r'^\s*' + re.escape(token) + r'\b', '', s)
    return re.sub(r'^[\s,]+', '', s).strip()

def build_full(out_path="out_full.xlsx"):
    raw = json.load(open(os.path.join(HERE, "cards_raw.json"), encoding="utf-8"))
    enr = {c["num"]: c for c in json.load(open(os.path.join(HERE, "enriched_cumulative.json"), encoding="utf-8"))}
    full = []
    for r in raw:
        num = r["num"]
        if num in enr:
            full.append(enr[num])
            continue
        # 미강화: 원본 기본 정보로 최소 카드 구성
        mean = (r.get("mean") or "").strip()
        full.append({
            "num": num,
            "hw": r.get("headword", "").strip(),
            "pos": clean_pos(r.get("pos", ""), r.get("headword", ""), r.get("cefr", "")),
            "cefr": r.get("cefr", ""),
            "core": "",
            "senses": [mean] if mean else [],
            "etym": "",
            "deriv": "",
            "syn": [], "anti": [], "idiom": [], "colloc": [],
            "ex_orig": (r.get("example", "") or "").replace("\n", " ").strip(),
            "ex_src": (r.get("source", "") or "").split(" | ")[0].strip(),
            "ex_add": "",
        })
    # 임시 파일로 저장 후 build_xlsx 재사용
    tmp = os.path.join(HERE, "_full_tmp.json")
    json.dump(full, open(tmp, "w"), ensure_ascii=False)
    build_xlsx.build("_full_tmp.json", out_path)
    os.remove(tmp)
    enriched_n = len(enr)
    print("FULL: %d words total, %d enriched (%.0f%%)" %
          (len(full), enriched_n, 100*enriched_n/len(full)))

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "out_full.xlsx"
    build_full(out)
