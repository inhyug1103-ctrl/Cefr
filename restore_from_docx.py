#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX(완료분 단어장) -> enriched_cumulative.json 복원기.

새 세션에서 enriched_cumulative.json을 잃었을 때, 가장 최근 DOCX에서
마스터 강화본을 되살린다. (인수인계 문서 60번 복구 절차)

사용: python3 restore_from_docx.py <완료분.docx> [out.json]
"""
import sys, re, html, json, zipfile

LBL = ['핵심 코어', '뜻갈래', '어원적 기원', '의미파생원리', '유의어', '반의어',
       '숙어·구동사', '관용·연어', '실제 예문', '회독·메모']
# 항목 구분자는 양쪽에 공백이 있는 가운뎃점("   ·   ")만. 텍스트 내부의
# 공백 없는 "·"(예: 삼촌·외삼촌)는 분리하지 않는다.
SEP_DOT = r'\s{2,}·\s{2,}'
SEP_SENSE = r'\s{3,}'

def splitlist(s, sep):
    s = (s or '').strip()
    if s in ('', '-'):
        return []
    return [p.strip() for p in re.split(sep, s) if p.strip()]

def restore(docx):
    x = zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')
    runs = re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', x, re.S)
    runs = [html.unescape(t) for t in runs]
    starts = [i for i, t in enumerate(runs) if re.match(r'^#\d{4}\s*$', t)]
    cards = []
    for si, s in enumerate(starts):
        end = starts[si + 1] if si + 1 < len(starts) else len(runs)
        seg = runs[s:end]
        num = re.match(r'^#(\d{4})', seg[0].strip()).group(1)
        hw = seg[1].strip()
        posc = seg[2].strip()
        m = re.match(r'^(.*)·\s*([A-C][12])\s*$', posc)
        pos = m.group(1).strip() if m else posc
        cefr = m.group(2) if m else ''
        pos_idx = {seg[j].strip(): j for j in range(3, len(seg)) if seg[j].strip() in LBL}

        def block(label):
            if label not in pos_idx:
                return []
            a = pos_idx[label] + 1
            nxt = min([pos_idx[l] for l in LBL if l in pos_idx and pos_idx[l] > pos_idx[label]]
                      + [len(seg)])
            return seg[a:nxt]

        def joined(label):
            return ' '.join(block(label)).strip()

        exblk = block('실제 예문')
        ex_orig = ex_src = ex_add = ''
        k = 0
        while k < len(exblk):
            t = exblk[k].strip()
            if t == '[발췌]':
                ex_orig = exblk[k + 1].strip() if k + 1 < len(exblk) else ''
                if k + 2 < len(exblk):
                    sm = re.match(r'^\s*\((.*)\)\s*$', exblk[k + 2])
                    if sm:
                        ex_src = sm.group(1).strip(); k += 1
                k += 2; continue
            if t == '[추가]':
                ex_add = exblk[k + 1].strip() if k + 1 < len(exblk) else ''
                k += 2; continue
            k += 1
        cards.append(dict(
            num=num, hw=hw, pos=pos, cefr=cefr,
            core=joined('핵심 코어'),
            senses=splitlist(joined('뜻갈래'), SEP_SENSE),
            etym=joined('어원적 기원'),
            deriv=joined('의미파생원리'),
            syn=splitlist(joined('유의어'), SEP_DOT),
            anti=splitlist(joined('반의어'), SEP_DOT),
            idiom=splitlist(joined('숙어·구동사'), SEP_DOT),
            colloc=splitlist(joined('관용·연어'), SEP_DOT),
            ex_orig=ex_orig, ex_src=ex_src, ex_add=ex_add))
    return cards

if __name__ == '__main__':
    docx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else 'enriched_cumulative.json'
    cards = restore(docx)
    json.dump(cards, open(out, 'w'), ensure_ascii=False, indent=1)
    print('restored %d cards -> %s' % (len(cards), out))
