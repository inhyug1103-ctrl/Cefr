#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
InDesign 반입용 파일 생성
1) Tagged Text (.txt, UTF-16LE, <UNICODE-WIN>) : 단락/문자 스타일 포함 → InDesign '파일>가져오기'
2) Data Merge CSV (UTF-8 BOM)                  : InDesign 데이터 병합(Data Merge)용
사용: python3 make_indesign.py
"""
import json, os, csv
import build_textbook as bt

HERE = os.path.dirname(os.path.abspath(__file__))

def joinlist(xs, sep=' · '):
    xs = [x for x in (xs or []) if x and x != '-']
    return sep.join(xs) if xs else ''

def load_days():
    cum = json.load(open(os.path.join(HERE, 'enriched_cumulative.json'), encoding='utf-8'))
    cum.sort(key=lambda x: x['num'])
    return bt.chunk_days(cum)

# ---------- 1) InDesign Tagged Text ----------
def tt_esc(s):
    s = '' if s is None else str(s)
    return s.replace('\\', '\\\\').replace('<', '\\<').replace('>', '\\>')

def make_tagged_text(days, path):
    L = ['<UNICODE-WIN>', '<vsn:14.0><FeatureSet:InDesign-Roman>']
    for dayno, lv, cards in days:
        L.append('<ParaStyle:DayHead>DAY %02d \\u00b7 %s (#%s\\u2013#%s)'
                 % (dayno, lv, cards[0]['num'], cards[-1]['num']))
        for c in cards:
            L.append('<ParaStyle:Headword>%s\t<CharStyle:POS>%s \\u00b7 %s<CharStyle:>'
                     % (tt_esc(c['hw']), tt_esc(c.get('pos', '')), tt_esc(c.get('cefr', ''))))
            for name, val in [
                ('Core', c.get('core', '')),
                ('Senses', joinlist(c.get('senses'), '  ')),
                ('Etym', c.get('etym', '')),
                ('Deriv', c.get('deriv', '')),
                ('SynAnt', '유의어 ' + (joinlist(c.get('syn')) or '-') + '   |   반의어 ' + (joinlist(c.get('anti')) or '-')),
                ('Phrase', '숙어 ' + (joinlist(c.get('idiom')) or '-') + '   |   연어 ' + (joinlist(c.get('colloc')) or '-')),
                ('Example', c.get('ex_add', '')),
            ]:
                if val:
                    L.append('<ParaStyle:%s>%s' % (name, tt_esc(val)))
    text = '\r\n'.join(L) + '\r\n'
    with open(os.path.join(HERE, path), 'wb') as f:
        f.write(b'\xff\xfe')                 # UTF-16LE BOM
        f.write(text.encode('utf-16-le'))
    print('built', path)

# ---------- 2) Data Merge CSV ----------
def make_csv(days, path):
    cols = ['day', 'num', 'hw', 'pos', 'cefr', 'core', 'senses',
            'etym', 'deriv', 'syn', 'anti', 'idiom', 'colloc', 'ex_add']
    n = 0
    with open(os.path.join(HERE, path), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(cols)
        for dayno, lv, cards in days:
            for c in cards:
                n += 1
                w.writerow([
                    'DAY %02d' % dayno, c['num'], c.get('hw', ''), c.get('pos', ''), c.get('cefr', ''),
                    c.get('core', ''), joinlist(c.get('senses'), '  '),
                    c.get('etym', ''), c.get('deriv', ''),
                    joinlist(c.get('syn')), joinlist(c.get('anti')),
                    joinlist(c.get('idiom')), joinlist(c.get('colloc')), c.get('ex_add', ''),
                ])
    print('built', path, '(%d rows)' % n)

if __name__ == '__main__':
    days = load_days()
    make_tagged_text(days, 'CEFR_indesign_taggedtext.txt')
    make_csv(days, 'CEFR_indesign_datamerge.csv')
    print('days:', len(days))
