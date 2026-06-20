#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CEFR 기준 영단어 -> 'DAY 단위 어휘 교재' 빌더 (디자인 고급화판)
- 표지(컬러 밴드) / 머리말·구성·학습법 / 목차
- 레벨별 컬러 디바이더 -> DAY(40단어): 레벨색 배너 + 카드 + 셀프체크
- 알파벳 INDEX / 꼬리말 페이지 번호
사용: python3 build_textbook.py [enriched_cumulative.json] [out.docx]
"""
import json, os, sys, zipfile
import build

HERE = os.path.dirname(os.path.abspath(__file__))
esc, run, para, rpr = build.esc, build.run, build.para, build.rpr
BORD = build.BORD
WIDTH = 9360
PER_DAY = 40
LEVEL_ORDER = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2', '심화']
LEVEL_DESC = {
    'A1': '기초 · 초등~중1', 'A2': '초급 · 중1~중2', 'B1': '중급 · 중2~중3',
    'B2': '중상급 · 고1~수능 기본', 'C1': '상급 · 수능 심화·상위권',
    'C2': '최상급 · 특목·SAT/GRE급', '심화': '심화 부록 · AP·GRE·문학·전문',
}
LEVEL_COLOR = {
    'A1': '2E75B6', 'A2': '1E8449', 'B1': 'CA6F1E', 'B2': '7D3C98',
    'C1': 'B03A2E', 'C2': '0E6655', '심화': '2C3E50',
}
LEVEL_GOAL = {
    'A1': '가장 기본적인 일상·학교 어휘. 영어 학습의 토대를 세우는 단계.',
    'A2': '간단한 글과 대화를 이해하는 초급 어휘. 중등 내신의 기본기.',
    'B1': '문단 독해가 가능해지는 중급 어휘. 본격 독해의 출발점.',
    'B2': '수능 지문의 기본 골격을 이루는 중상급 어휘.',
    'C1': '수능 고난도·상위권 변별 어휘. 여기까지가 수능 어휘 완성선.',
    'C2': '특목고·SAT·GRE급 최상위 어휘. 변별과 유학을 위한 구간.',
    '심화': 'AP·GRE·문학·전문 분야의 초고난도 어휘. 어휘력의 정점.',
}

def pagebreak():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def center_p(runs_xml, after=120, before=0):
    return ('<w:p><w:pPr><w:spacing w:after="%d" w:before="%d"/><w:jc w:val="center"/>'
            '</w:pPr>%s</w:p>' % (after, before, runs_xml))

def heading(text, sz=30, color='1F4E79', after=160, before=240, bar=True):
    """좌측 컬러 바를 가진 섹션 제목."""
    pre = run('▍ ', sz, color, b=True) if bar else ''
    return ('<w:p><w:pPr><w:spacing w:after="%d" w:before="%d"/></w:pPr>%s%s</w:p>'
            % (after, before, pre, run(text, sz, color, b=True)))

def bullet(text, sz=18):
    return ('<w:p><w:pPr><w:spacing w:after="70"/><w:ind w:left="380" w:hanging="260"/>'
            '</w:pPr>%s%s</w:p>' % (run('•  ', sz, 'C55A11', b=True), run(text, sz)))

def band(body_xml, fill, w=WIDTH, top=80, bottom=80):
    """단색 컬러 밴드(1셀 표)."""
    tc = ('<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/>' % w +
          '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>'
          '<w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'
          '<w:shd w:fill="%s" w:val="clear"/>' % fill +
          '<w:tcMar><w:top w:type="dxa" w:w="%d"/><w:left w:type="dxa" w:w="200"/>'
          '<w:bottom w:type="dxa" w:w="%d"/><w:right w:type="dxa" w:w="200"/></w:tcMar>'
          '</w:tcPr>' % (top, bottom) + body_xml + '</w:tc>')
    return ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/>' % w +
            '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
            '<w:right w:val="nil"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>' % w +
            '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + tc + '</w:tr></w:tbl>')

def cover():
    p = ['<w:p><w:pPr><w:spacing w:after="0" w:before="600"/></w:pPr></w:p>']
    top = (center_p(run('CEFR 기준', 30, 'FFFFFF', b=True), after=40, before=120) +
           center_p(run('영  단  어', 60, 'FFFFFF', b=True), after=120) +
           center_p(run('VOCABULARY  BY  CEFR', 16, 'BFD7EE'), after=120))
    p.append(band(top, '1F4E79', top=400, bottom=400))
    p.append('<w:p><w:pPr><w:spacing w:after="0" w:before="200"/></w:pPr></w:p>')
    sub = (center_p(run('어휘 교재 · DAY 학습판', 24, 'FFFFFF', b=True), after=60, before=80) +
           center_p(run('기초(A1)부터 수능, 그 너머(특목·SAT/GRE)까지', 15, 'FCE4D6'), after=80))
    p.append(band(sub, 'C55A11', top=160, bottom=160))
    p.append('<w:p><w:pPr><w:spacing w:after="0" w:before="500"/></w:pPr></w:p>')
    p.append(center_p(run('A1 · A2 · B1 · B2 · C1 · C2  +  심화', 18, '1F4E79', b=True), after=40))
    p.append(center_p(run('총 6,000단어 · 154 DAY · 어원 중심 심층 학습', 15, '888888'), after=0))
    return ''.join(p) + pagebreak()

def preface():
    p = [heading('머리말', 30, before=120)]
    p.append(para(run('이 책은 단어를 "외우는" 데서 그치지 않고 "이해하는" 어휘 교재입니다. '
                      '모든 표제어를 실제로 검증된 어원에 뿌리내려 풀이하고, 같은 어근끼리 묶어 '
                      '한 단어를 알면 열 단어가 따라오도록 설계했습니다. CEFR 등급(A1~C2)에 '
                      '심화 부록을 더해, 기초 회화부터 수능·특목·유학 시험까지 한 권으로 잇습니다.', 17), after=180))
    p.append(heading('이 책의 구성과 특징', 24))
    for t in [
        '6요소 심층 카드: ① 핵심 코어 ② 뜻갈래 ③ 실제 어원 ④ 의미파생원리 ⑤ 유의어·반의어 ⑥ 숙어·연어·예문',
        '어원은 "실제로 검증된 것만" 수록 — 외우기용 가짜 어원 없음(불확실하면 그렇다고 명시).',
        '어근 클러스터: tenere(붙잡다)·specere(보다)·ducere(이끌다) 등 한 뿌리로 무리 지어 암기.',
        'CEFR 등급별 컬러 구분 + 심화 부록: 자기 수준의 DAY만 골라 학습.',
        '한국 학습자 혼동 포인트 표시: elude/allude, defuse/diffuse, discreet/discrete 등.',
        'DAY마다 셀프체크(REVIEW)로 그날 배운 단어를 즉시 점검.',
    ]:
        p.append(bullet(t))
    p.append(heading('학습법 (중1 → 고3 수능)', 24))
    for t in [
        '하루 1 DAY(40단어)가 기본. 무리하면 20단어씩 반일치 학습.',
        '1회독: 코어+뜻 / 2회독: 어원·파생 / 3회독: 예문·연어 — 카드의 ☐☐☐에 체크.',
        '중1~2: A1·A2  ·  중3~고1: B1·B2  ·  고2~3: C1(수능 어휘 완성선)  ·  최상위/유학: C2·심화.',
        'DAY 끝 REVIEW에서 막힌 단어는 다음 날 먼저 복습(누적 복습).',
        '단어가 끝이 아닙니다 — 문법·구문독해·듣기·기출 풀이를 반드시 병행하세요.',
    ]:
        p.append(bullet(t))
    return ''.join(p) + pagebreak()

def toc(days):
    p = [heading('목차  ·  DAY 구성', 30, before=120)]
    cur = None
    for dayno, lv, cards in days:
        if lv != cur:
            cur = lv
            p.append(band('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' +
                          run('%s' % lv, 18, 'FFFFFF', b=True) +
                          run('   %s' % LEVEL_DESC.get(lv, ''), 13, 'FFFFFF') + '</w:p>',
                          LEVEL_COLOR.get(lv, '1F4E79'), top=50, bottom=50))
        p.append('<w:p><w:pPr><w:spacing w:after="20" w:before="20"/><w:ind w:left="360"/></w:pPr>'
                 + run('DAY %02d' % dayno, 15, '1F4E79', b=True)
                 + run('    #%s–#%s    (%d단어)' % (cards[0]['num'], cards[-1]['num'], len(cards)), 14, '666666')
                 + '</w:p>')
    return ''.join(p) + pagebreak()

def level_divider(lv):
    color = LEVEL_COLOR.get(lv, '1F4E79')
    body = (center_p(run('LEVEL', 16, 'FFFFFF'), after=20, before=200) +
            center_p(run(lv, 72, 'FFFFFF', b=True), after=40) +
            center_p(run(LEVEL_DESC.get(lv, ''), 20, 'FFFFFF', b=True), after=200) +
            center_p(run(LEVEL_GOAL.get(lv, ''), 15, 'F2F2F2'), after=200, before=0))
    return pagebreak() + band(body, color, top=700, bottom=700) + pagebreak()

def day_banner(dayno, level, cards):
    color = LEVEL_COLOR.get(level, '1F4E79')
    n0, n1 = cards[0]['num'], cards[-1]['num']
    body = ('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' +
            run('DAY %02d' % dayno, 32, 'FFFFFF', b=True) +
            run('    %s' % level, 18, 'FFFFFF', b=True) +
            run(' · %s' % LEVEL_DESC.get(level, ''), 13, 'F2F2F2') +
            run('      #%s–#%s · %d단어' % (n0, n1, len(cards)), 13, 'F2F2F2') +
            '</w:p>')
    return band(body, color, top=110, bottom=110)

def selfcheck(dayno, cards):
    hdr = band('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' +
               run('✍  DAY %02d  REVIEW' % dayno, 18, 'FFFFFF', b=True) +
               run('   — 영단어를 보고 뜻을 적어 보세요', 13, 'FCE4D6') + '</w:p>',
               'C55A11', top=70, bottom=70)
    half = (len(cards) + 1) // 2
    left, right = cards[:half], cards[half:]
    rows = []
    for i in range(half):
        lc = left[i] if i < len(left) else None
        rc = right[i] if i < len(right) else None
        def eng(c): return build.cell(2040, para(run(c['hw'], 15, '1F4E79', b=True)) if c else para(run('', 14)))
        def blk(c): return build.cell(2640, para(run('', 14)))
        rows.append(build.row([eng(lc), blk(lc), eng(rc), blk(rc)]))
    grid = ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/>' % WIDTH +
            '<w:tblBorders><w:top w:val="single" w:color="D9D9D9" w:sz="3"/>'
            '<w:left w:val="single" w:color="D9D9D9" w:sz="3"/>'
            '<w:bottom w:val="single" w:color="D9D9D9" w:sz="3"/>'
            '<w:right w:val="single" w:color="D9D9D9" w:sz="3"/>'
            '<w:insideH w:val="single" w:color="D9D9D9" w:sz="3"/>'
            '<w:insideV w:val="single" w:color="D9D9D9" w:sz="3"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2040"/><w:gridCol w:w="2640"/>'
            '<w:gridCol w:w="2040"/><w:gridCol w:w="2640"/></w:tblGrid>' +
            ''.join(rows) + '</w:tbl>')
    return hdr + grid

def chunk_days(cum):
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

def index_section(cum):
    p = [pagebreak(), heading('찾아보기 · INDEX (알파벳순)', 28, before=120)]
    items = sorted(cum, key=lambda c: c['hw'].lower())
    cols, cellw = 4, WIDTH // 4
    rows, buf = [], []
    for c in items:
        buf.append(build.cell(cellw, para(run(c['hw'], 12, b=True) + run('  %s' % c['num'], 10, '999999')),
                              mar_l=60, mar_r=30))
        if len(buf) == cols:
            rows.append(build.row(buf)); buf = []
    if buf:
        while len(buf) < cols:
            buf.append(build.cell(cellw, para(run('', 10))))
        rows.append(build.row(buf))
    grid = ('<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/></w:tblPr>' % WIDTH +
            '<w:tblGrid>' + ('<w:gridCol w:w="%d"/>' % cellw) * cols + '</w:tblGrid>' +
            ''.join(rows) + '</w:tbl>')
    return ''.join(p) + grid

# 꼬리말(페이지 번호) 주입용
FOOTER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>'
    + '<w:r>' + rpr(14, '9AA5B1') + '<w:t xml:space="preserve">CEFR 기준 영단어  ·  </w:t></w:r>'
    + '<w:fldSimple w:instr=" PAGE "><w:r>' + rpr(14, '9AA5B1') + '<w:t>1</w:t></w:r></w:fldSimple>'
    + '</w:p></w:ftr>'
)
SECTPR_F = ('<w:sectPr><w:footerReference w:type="default" r:id="rId100"/>'
            '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
            '<w:pgMar w:top="1000" w:right="1080" w:bottom="1000" w:left="1080" '
            'w:header="600" w:footer="500" w:gutter="0"/><w:pgNumType/>'
            '<w:docGrid w:linePitch="360"/></w:sectPr>')

def build_textbook(cum_path='enriched_cumulative.json', out_path='out_textbook.docx'):
    cum = json.load(open(os.path.join(HERE, cum_path), encoding='utf-8'))
    cum.sort(key=lambda x: x['num'])
    days = chunk_days(cum)
    parts = [cover(), preface(), toc(days)]
    cur_level = None
    for dayno, lv, cards in days:
        if lv != cur_level:
            cur_level = lv
            parts.append(level_divider(lv))
        else:
            parts.append(pagebreak())
        parts.append(day_banner(dayno, lv, cards))
        parts.append(build.SPACER)
        for c in cards:
            parts.append(build.card_table(c))
            parts.append(build.SPACER)
        parts.append(build.SPACER)
        parts.append(selfcheck(dayno, cards))
    parts.append(index_section(cum))
    body = ''.join(parts) + SECTPR_F
    head = open(os.path.join(HERE, 'template_head.xml'), encoding='utf-8').read()
    document = head + body + '</w:body></w:document>'
    src = os.path.join(HERE, 'template.docx')
    out = os.path.join(HERE, out_path)
    if os.path.exists(out):
        os.remove(out)
    zin = zipfile.ZipFile(src, 'r')
    zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.namelist():
        data = zin.read(item)
        if item == 'word/document.xml':
            zout.writestr(item, document); continue
        if item == 'word/_rels/document.xml.rels':
            s = data.decode('utf-8').replace(
                '</Relationships>',
                '<Relationship Id="rId100" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
                'Target="footer1.xml"/></Relationships>')
            zout.writestr(item, s); continue
        if item == '[Content_Types].xml':
            s = data.decode('utf-8').replace(
                '</Types>',
                '<Override ContentType="application/vnd.openxmlformats-officedocument.'
                'wordprocessingml.footer+xml" PartName="/word/footer1.xml"/></Types>')
            zout.writestr(item, s); continue
        zout.writestr(item, data)
    zout.writestr('word/footer1.xml', FOOTER_XML)
    zin.close(); zout.close()
    print('built %s  (%d cards, %d days)' % (out_path, len(cum), len(days)))

if __name__ == '__main__':
    cp = sys.argv[1] if len(sys.argv) > 1 else 'enriched_cumulative.json'
    op = sys.argv[2] if len(sys.argv) > 2 else 'out_textbook.docx'
    build_textbook(cp, op)
