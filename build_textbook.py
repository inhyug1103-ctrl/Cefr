#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CEFR 기준 영단어 -> 'DAY 단위 어휘 교재' 빌더
- 표지 / 머리말·구성과 특징·학습법 / 목차(DAY)
- 레벨별(A1~C2, 심화) -> DAY(40단어) 본문: DAY 배너 + 카드(build.card_table 재사용) + DAY 셀프체크
- 알파벳 INDEX(찾아보기)
사용: python3 build_textbook.py [enriched_cumulative.json] [out.docx]
"""
import json, os, sys, zipfile
import build  # 카드 렌더링·헬퍼 재사용

HERE = os.path.dirname(os.path.abspath(__file__))
esc, run, para, rpr = build.esc, build.run, build.para, build.rpr
BORD = build.BORD
WIDTH = 9360
PER_DAY = 40
LEVEL_ORDER = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2', '심화']
LEVEL_DESC = {
    'A1': '기초 (초등~중1)', 'A2': '초급 (중1~중2)', 'B1': '중급 (중2~중3)',
    'B2': '중상급 (고1~수능 기본)', 'C1': '상급 (수능 심화·상위권)',
    'C2': '최상급 (특목·SAT/GRE급)', '심화': '심화 부록 (AP·GRE·문학·전문)',
}

def pagebreak():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def center_p(runs_xml, after=120, before=0):
    return ('<w:p><w:pPr><w:spacing w:after="%d" w:before="%d"/>'
            '<w:jc w:val="center"/></w:pPr>%s</w:p>' % (after, before, runs_xml))

def heading(text, sz=30, color='1F4E79', after=160, before=240):
    return ('<w:p><w:pPr><w:spacing w:after="%d" w:before="%d"/></w:pPr>%s</w:p>'
            % (after, before, run(text, sz, color, b=True)))

def bullet(text, sz=18):
    return ('<w:p><w:pPr><w:spacing w:after="60"/><w:ind w:left="360" w:hanging="240"/>'
            '</w:pPr>%s%s</w:p>' % (run('•  ', sz, '1F4E79', b=True), run(text, sz)))

def one_cell_table(body_xml, fill='1F4E79'):
    tc = ('<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/>' % WIDTH + BORD +
          '<w:shd w:fill="%s" w:val="clear"/>' % fill +
          '<w:tcMar><w:top w:type="dxa" w:w="80"/><w:left w:type="dxa" w:w="160"/>'
          '<w:bottom w:type="dxa" w:w="80"/><w:right w:type="dxa" w:w="160"/></w:tcMar>'
          '</w:tcPr>' + body_xml + '</w:tc>')
    return ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/>' % WIDTH +
            '<w:tblBorders><w:top w:val="single" w:color="auto" w:sz="4"/>'
            '<w:left w:val="single" w:color="auto" w:sz="4"/>'
            '<w:bottom w:val="single" w:color="auto" w:sz="4"/>'
            '<w:right w:val="single" w:color="auto" w:sz="4"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>' % WIDTH +
            '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + tc + '</w:tr></w:tbl>')

def day_banner(dayno, level, cards):
    n0, n1 = cards[0]['num'], cards[-1]['num']
    body = ('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' +
            run('DAY %02d' % dayno, 30, 'FFFFFF', b=True) +
            run('    %s · %s' % (level, LEVEL_DESC.get(level, '')), 16, 'AEC6E0') +
            run('      #%s–#%s · %d단어' % (n0, n1, len(cards)), 14, 'AEC6E0') +
            '</w:p>')
    return one_cell_table(body, fill='1F4E79')

def selfcheck(dayno, cards):
    """DAY 셀프체크: 영단어를 보고 뜻을 적는 4열 표."""
    hdr = one_cell_table(
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' +
        run('✍  DAY %02d REVIEW' % dayno, 18, 'FFFFFF', b=True) +
        run('   — 뜻을 적어 보며 셀프 체크하세요', 13, 'FCE4D6') + '</w:p>',
        fill='C55A11')
    # 4열(영어|빈칸|영어|빈칸) 표
    rows = []
    half = (len(cards) + 1) // 2
    left, right = cards[:half], cards[half:]
    for i in range(half):
        def engcell(c):
            if not c:
                return build.cell(2040, para(run('', 14)))
            return build.cell(2040, para(run(c['hw'], 15, '1F4E79', b=True)))
        def blankcell(c):
            return build.cell(2640, para(run('', 14)))
        lc = left[i] if i < len(left) else None
        rc = right[i] if i < len(right) else None
        rows.append(build.row([engcell(lc), blankcell(lc), engcell(rc), blankcell(rc)]))
    grid = ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/>' % WIDTH +
            '<w:tblBorders><w:top w:val="single" w:color="CCCCCC" w:sz="3"/>'
            '<w:left w:val="single" w:color="CCCCCC" w:sz="3"/>'
            '<w:bottom w:val="single" w:color="CCCCCC" w:sz="3"/>'
            '<w:right w:val="single" w:color="CCCCCC" w:sz="3"/>'
            '<w:insideH w:val="single" w:color="CCCCCC" w:sz="3"/>'
            '<w:insideV w:val="single" w:color="CCCCCC" w:sz="3"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2040"/><w:gridCol w:w="2640"/>'
            '<w:gridCol w:w="2040"/><w:gridCol w:w="2640"/></w:tblGrid>' +
            ''.join(rows) + '</w:tbl>')
    return hdr + grid

def chunk_days(cum):
    """레벨 순서대로 40단어씩 DAY로 분할. [(dayno, level, cards), ...]"""
    by_level = {lv: [] for lv in LEVEL_ORDER}
    for c in cum:
        by_level.setdefault(c.get('cefr', ''), []).append(c)
    days, dayno = [], 0
    for lv in LEVEL_ORDER:
        cards = by_level.get(lv, [])
        for i in range(0, len(cards), PER_DAY):
            dayno += 1
            days.append((dayno, lv, cards[i:i + PER_DAY]))
    return days

def cover():
    p = []
    p.append('<w:p><w:pPr><w:spacing w:before="1400" w:after="80"/><w:jc w:val="center"/></w:pPr>'
             + run('CEFR 기준 영단어', 56, '1F4E79', b=True) + '</w:p>')
    p.append(center_p(run('어휘 교재 · DAY 학습판', 26, '2E75B6'), after=60))
    p.append(center_p(run('기초(A1)부터 수능, 그 너머(특목·SAT/GRE)까지', 18, '555555'), after=400))
    p.append(center_p(run('A1 · A2 · B1 · B2 · C1 · C2  +  심화', 18, '1F4E79', b=True), after=40))
    p.append(center_p(run('총 6,000단어 · 어원 중심 심층 학습', 16, '888888'), after=600))
    p.append(center_p(run('— 한 권으로 끝내는 CEFR 등급별 어휘 —', 15, '999999'), after=0))
    return ''.join(p) + pagebreak()

def preface():
    p = [heading('머리말', 30, before=200)]
    p.append(para(run('이 책은 단어를 "외우는" 데서 그치지 않고 "이해하는" 어휘 교재입니다. '
                      '모든 표제어를 실제로 검증된 어원에 뿌리내려 풀이하고, 같은 어근끼리 묶어 '
                      '한 단어를 알면 열 단어가 따라오도록 설계했습니다. CEFR 등급(A1~C2)에 '
                      '심화 부록을 더해, 기초 회화부터 수능·특목·유학 시험까지 한 권으로 잇습니다.', 17), after=160))
    p.append(heading('이 책의 구성과 특징', 24))
    for t in [
        '6요소 심층 카드: ① 핵심 코어 ② 뜻갈래 ③ 실제 어원 ④ 의미파생원리 ⑤ 유의어·반의어 ⑥ 숙어·연어·예문',
        '어원은 "실제로 검증된 것만" 수록 — 외우기용 가짜 어원 없음(불확실하면 그렇다고 명시).',
        '어근 클러스터: tenere(붙잡다)·specere(보다)·ducere(이끌다) 등 한 뿌리로 무리 지어 암기.',
        'CEFR 등급별 배열 + 심화 부록: 자기 수준의 DAY만 골라 학습 가능.',
        '한국 학습자 혼동 포인트 표시: elude/allude, defuse/diffuse, discreet/discrete 등.',
        'DAY마다 셀프 체크(REVIEW)로 그날 배운 단어를 즉시 점검.',
    ]:
        p.append(bullet(t))
    p.append(heading('학습법 (중1 → 고3 수능)', 24))
    for t in [
        '하루 1 DAY(40단어)가 기본. 무리하면 20단어씩 반일치 학습.',
        '1회독: 코어+뜻만 빠르게 / 2회독: 어원·파생까지 / 3회독: 예문·연어까지 — 카드의 ☐☐☐에 체크.',
        '중1~중2: A1~A2  ·  중3~고1: B1~B2  ·  고2~고3: C1(= 수능 어휘 완성선)  ·  최상위/유학: C2·심화.',
        'DAY 끝 REVIEW에서 막힌 단어는 다음 날 먼저 복습(누적 복습).',
        '단어가 끝이 아닙니다 — 문법·구문독해·듣기·기출 풀이를 반드시 병행하세요.',
    ]:
        p.append(bullet(t))
    return ''.join(p) + pagebreak()

def toc(days):
    p = [heading('목차  (DAY 구성)', 30, before=120)]
    cur = None
    for dayno, lv, cards in days:
        if lv != cur:
            cur = lv
            p.append('<w:p><w:pPr><w:spacing w:after="40" w:before="160"/></w:pPr>'
                     + run('【 %s 】 %s' % (lv, LEVEL_DESC.get(lv, '')), 18, 'C55A11', b=True)
                     + '</w:p>')
        line = ('<w:p><w:pPr><w:spacing w:after="20"/><w:ind w:left="360"/></w:pPr>'
                + run('DAY %02d' % dayno, 15, '1F4E79', b=True)
                + run('   #%s–#%s   (%d단어)' % (cards[0]['num'], cards[-1]['num'], len(cards)), 14, '555555')
                + '</w:p>')
        p.append(line)
    return ''.join(p) + pagebreak()

def index_section(cum):
    p = [heading('찾아보기 (INDEX) — 알파벳순', 28, before=120)]
    items = sorted(cum, key=lambda c: c['hw'].lower())
    # 4열 표로 압축
    cols = 4
    cellw = WIDTH // cols
    rows, buf = [], []
    def mkcell(c):
        body = para(run(c['hw'], 13, b=True) + run('  %s' % c['num'], 11, '999999'))
        return build.cell(cellw, body, mar_l=60, mar_r=40)
    per_row = []
    for c in items:
        per_row.append(mkcell(c))
        if len(per_row) == cols:
            rows.append(build.row(per_row)); per_row = []
    if per_row:
        while len(per_row) < cols:
            per_row.append(build.cell(cellw, para(run('', 11))))
        rows.append(build.row(per_row))
    grid = ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/></w:tblPr>' % WIDTH +
            '<w:tblGrid>' + ('<w:gridCol w:w="%d"/>' % cellw) * cols + '</w:tblGrid>' +
            ''.join(rows) + '</w:tbl>')
    return pagebreak() + ''.join(p) + grid

def build_textbook(cum_path='enriched_cumulative.json', out_path='out_textbook.docx'):
    cum = json.load(open(os.path.join(HERE, cum_path), encoding='utf-8'))
    cum.sort(key=lambda x: x['num'])
    days = chunk_days(cum)
    parts = [cover(), preface(), toc(days)]
    cur_level = None
    for dayno, lv, cards in days:
        parts.append(pagebreak())
        if lv != cur_level:
            cur_level = lv
            parts.append(center_p(run('· · ·  %s  ·  %s  · · ·' % (lv, LEVEL_DESC.get(lv, '')),
                                      20, 'C55A11', b=True), after=160, before=200))
        parts.append(day_banner(dayno, lv, cards))
        parts.append(build.SPACER)
        for c in cards:
            parts.append(build.card_table(c))
            parts.append(build.SPACER)
        parts.append(build.SPACER)
        parts.append(selfcheck(dayno, cards))
    parts.append(index_section(cum))
    body = ''.join(parts) + build.SECTPR
    head = open(os.path.join(HERE, 'template_head.xml'), encoding='utf-8').read()
    document = head + body + '</w:body></w:document>'
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
    print('built %s  (%d cards, %d days)' % (out_path, len(cum), len(days)))

if __name__ == '__main__':
    cp = sys.argv[1] if len(sys.argv) > 1 else 'enriched_cumulative.json'
    op = sys.argv[2] if len(sys.argv) > 2 else 'out_textbook.docx'
    build_textbook(cp, op)
