#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
배치(batch_XXXX_YYYY.json)를 마스터(enriched_cumulative.json)에 누적하고
원본 예문/출처(cards_raw.json)를 병합한다.

사용: python3 merge_batch.py batch_0321_0370.json
출력 효과: enriched_cumulative.json 갱신 (정렬·중복 제거·예문 병합)
이후 `python3 build.py`로 렌더링.
"""
import sys, json, os

HERE = os.path.dirname(os.path.abspath(__file__))

def load(p):
    return json.load(open(os.path.join(HERE, p), encoding='utf-8'))

def merge(batch_path):
    cum = load('enriched_cumulative.json')
    have = {c['num'] for c in cum}
    added = 0
    for c in load(batch_path):
        if c['num'] not in have:
            cum.append(c); have.add(c['num']); added += 1
    cum.sort(key=lambda x: x['num'])
    # 원본 예문/출처 병합 (출처는 ' | ' 앞 짧은 코드만 표시용으로 저장)
    raw = {r['num']: r for r in load('cards_raw.json')}
    for c in cum:
        o = raw.get(c['num'], {})
        c['ex_orig'] = (o.get('example', '') or '').replace('\n', ' ').strip()
        c['ex_src'] = (o.get('source', '') or '').split(' | ')[0].strip()
    json.dump(cum, open(os.path.join(HERE, 'enriched_cumulative.json'), 'w'),
              ensure_ascii=False, indent=1)
    print('merged %s: +%d -> %d cards (#%s-#%s)' %
          (batch_path, added, len(cum), cum[0]['num'], cum[-1]['num']))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('usage: python3 merge_batch.py <batch.json>'); sys.exit(1)
    merge(sys.argv[1])
