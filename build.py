#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
수능영단어 심화강화판 빌더 (build2.js 재구성판)

입력 : enriched_cumulative.json  (마스터 강화본)
       cards_raw.json            (원본 예문/출처 병합용)
       template.docx             (포장 템플릿 - word/document.xml만 교체)
       template_intro.xml        (표지/읽는 법 인트로)
출력 : out.docx
포맷 : 카드 1장 = 표(table) 1개. 색 체계
       핵심 코어=노랑(FFF2C2) · 어원=하늘색(E2EEF8) · 의미파생=주황(FBE3D4)
       예문=옅은 하늘(F4F8FC) · 라벨=연회색(F0F3F7) · 헤더=네이비(1F4E79)
"""
import json, re, zipfile, shutil, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = 'Malgun Gothic'
RF = ('<w:rFonts w:ascii="Malgun Gothic" w:cs="Malgun Gothic" '
      'w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/>')
BORD = ('<w:tcBorders><w:top w:val="single" w:color="CCCCCC" w:sz="3"/>'
        '<w:left w:val="single" w:color="CCCCCC" w:sz="3"/>'
        '<w:bottom w:val="single" w:color="CCCCCC" w:sz="3"/>'
        '<w:right w:val="single" w:color="CCCCCC" w:sz="3"/></w:tcBorders>')

def esc(s):
    s = '' if s is None else str(s)
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    s = s.replace('"', '&quot;').replace("'", '&apos;')
    return s

def rpr(sz, color=None, b=False, i=False):
    out = '<w:rPr>' + RF
    if b: out += '<w:b/><w:bCs/>'
    if i: out += '<w:i/><w:iCs/>'
    if color: out += '<w:color w:val="%s"/>' % color
    out += '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz, sz)
    return out

def run(text, sz, color=None, b=False, i=False):
    return ('<w:r>' + rpr(sz, color, b, i) +
            '<w:t xml:space="preserve">%s</w:t></w:r>' % esc(text))

def para(runs_xml, after=0):
    return ('<w:p><w:pPr><w:spacing w:after="%d" w:line="230" '
            'w:lineRule="auto"/></w:pPr>%s</w:p>' % (after, runs_xml))

def cell(w, body, fill=None, gridspan=1, mar_l=100, mar_r=100, valign=True):
    tc = '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/>' % w
    if gridspan > 1:
        tc += '<w:gridSpan w:val="%d"/>' % gridspan
    elif gridspan == 1 and w in (3330,):
        tc += '<w:gridSpan w:val="1"/>'
    tc += BORD
    if fill:
        tc += '<w:shd w:fill="%s" w:val="clear"/>' % fill
    tc += ('<w:tcMar><w:top w:type="dxa" w:w="28"/>'
           '<w:left w:type="dxa" w:w="%d"/><w:bottom w:type="dxa" w:w="28"/>'
           '<w:right w:type="dxa" w:w="%d"/></w:tcMar>' % (mar_l, mar_r))
    if valign:
        tc += '<w:vAlign w:val="center"/>'
    tc += '</w:tcPr>' + body + '</w:tc>'
    return tc

def label_cell(text):
    body = para(run(text, 16, '1F4E79', b=True))
    return cell(1350, body, fill='F0F3F7', mar_l=90, mar_r=70)

def row(cells):
    return '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + ''.join(cells) + '</w:tr>'

def full_row(label, value_runs, fill=None):
    val = cell(8010, para(value_runs), fill=fill, gridspan=3)
    return row([label_cell(label), val])

def two_col_row(l1, v1, l2, v2):
    return row([
        label_cell(l1), cell(3330, para(run(v1, 17)), gridspan=1),
        label_cell(l2), cell(3330, para(run(v2, 17)), gridspan=1),
    ])

def joinlist(xs, sep='   ·   '):
    xs = [x for x in (xs or []) if x and x != '-']
    return sep.join(xs) if xs else '-'

def card_table(c):
    rows = []
    # header
    hdr = ('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' +
           run('#%s  ' % c['num'], 16, 'AEC6E0', b=True) +
           run(c['hw'], 26, 'FFFFFF', b=True) +
           run('   %s · %s' % (c.get('pos', ''), c.get('cefr', '')), 16, 'AEC6E0') +
           '</w:p>')
    hcell = ('<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9360"/>'
             '<w:gridSpan w:val="4"/>' + BORD +
             '<w:shd w:fill="1F4E79" w:val="clear"/>'
             '<w:tcMar><w:top w:type="dxa" w:w="46"/>'
             '<w:left w:type="dxa" w:w="120"/><w:bottom w:type="dxa" w:w="46"/>'
             '<w:right w:type="dxa" w:w="120"/></w:tcMar></w:tcPr>' + hdr + '</w:tc>')
    rows.append(row([hcell]))
    # 핵심 코어 (yellow, bold sz18)
    rows.append(full_row('핵심 코어', run(c.get('core', ''), 18, b=True), fill='FFF2C2'))
    # 뜻갈래 (no fill, sz17)
    rows.append(full_row('뜻갈래', run(joinlist(c.get('senses'), '   '), 17)))
    # 어원적 기원 (sky, sz16)
    rows.append(full_row('어원적 기원', run(c.get('etym', ''), 16), fill='E2EEF8'))
    # 의미파생원리 (orange, sz16)
    rows.append(full_row('의미파생원리', run(c.get('deriv', ''), 16), fill='FBE3D4'))
    # 유의어 / 반의어
    rows.append(two_col_row('유의어', joinlist(c.get('syn')),
                            '반의어', joinlist(c.get('anti'))))
    # 숙어·구동사 / 관용·연어
    rows.append(two_col_row('숙어·구동사', joinlist(c.get('idiom')),
                            '관용·연어', joinlist(c.get('colloc'))))
    # 실제 예문
    ex_paras = ''
    if c.get('ex_orig'):
        r = run('[발췌] ', 15, '1F4E79', b=True) + run(c['ex_orig'], 16, i=True)
        if c.get('ex_src'):
            r += run('  (%s)' % c['ex_src'], 13, '999999')
        ex_paras += para(r)
    if c.get('ex_add'):
        ex_paras += para(run('[추가] ', 15, 'C55A11', b=True) + run(c['ex_add'], 16))
    if not ex_paras:
        ex_paras = para(run('-', 16))
    rows.append(row([label_cell('실제 예문'),
                     cell(8010, ex_paras, fill='F4F8FC', gridspan=3)]))
    # 회독·메모
    rows.append(row([label_cell('회독·메모'),
                     cell(8010, para(run('1·2·3회독 ☐ ☐ ☐    메모:', 14, '888888')),
                          gridspan=3)]))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="9360"/>'
           '<w:tblBorders><w:top w:val="single" w:color="auto" w:sz="4"/>'
           '<w:left w:val="single" w:color="auto" w:sz="4"/>'
           '<w:bottom w:val="single" w:color="auto" w:sz="4"/>'
           '<w:right w:val="single" w:color="auto" w:sz="4"/>'
           '<w:insideH w:val="single" w:color="auto" w:sz="4"/>'
           '<w:insideV w:val="single" w:color="auto" w:sz="4"/></w:tblBorders>'
           '</w:tblPr><w:tblGrid><w:gridCol w:w="1350"/><w:gridCol w:w="3330"/>'
           '<w:gridCol w:w="1350"/><w:gridCol w:w="3330"/></w:tblGrid>' +
           ''.join(rows) + '</w:tbl>')
    return tbl

SPACER = ('<w:p><w:pPr><w:spacing w:after="90"/></w:pPr><w:r>' + rpr(2) +
          '<w:t xml:space="preserve"></w:t></w:r></w:p>')
SECTPR = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
          '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
          'w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType/>'
          '<w:docGrid w:linePitch="360"/></w:sectPr>')

def build(cum_path='enriched_cumulative.json', out_path='out.docx'):
    cum = json.load(open(os.path.join(HERE, cum_path), encoding='utf-8'))
    cum.sort(key=lambda x: x['num'])
    intro = open(os.path.join(HERE, 'template_intro.xml'), encoding='utf-8').read()
    # update preview note dynamically
    nums = [c['num'] for c in cum]
    note = '※ 수록: #%s–#%s (총 %d개). 동일 양식으로 전 단계 확장.' % (
        nums[0], nums[-1], len(cum))
    intro = re.sub(r'(?<=<w:t xml:space="preserve">)※ 프리뷰[^<]*(?=</w:t>)',
                   esc(note), intro)
    parts = [intro]
    for c in cum:
        parts.append(card_table(c))
        parts.append(SPACER)
    body = ''.join(parts) + SECTPR
    head = open(os.path.join(HERE, 'template_head.xml'), encoding='utf-8').read()
    document = head + body + '</w:body></w:document>'
    # repackage: copy template, replace word/document.xml
    src = os.path.join(HERE, 'template.docx')
    out = os.path.join(HERE, out_path)
    if os.path.exists(out):
        os.remove(out)
    zin = zipfile.ZipFile(src, 'r')
    zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.namelist():
        if item == 'word/document.xml':
            zout.writestr(item, document)
        else:
            zout.writestr(item, zin.read(item))
    zin.close(); zout.close()
    print('built %s  (%d cards, #%s-#%s)' % (out_path, len(cum), nums[0], nums[-1]))

if __name__ == '__main__':
    cum = sys.argv[1] if len(sys.argv) > 1 else 'enriched_cumulative.json'
    out = sys.argv[2] if len(sys.argv) > 2 else 'out.docx'
    build(cum, out)
