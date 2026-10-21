"""윤곽선 글자 읽기용 참조 글자표(js/glyph-ref.js)를 만듭니다.

실행: python3 scripts/make-glyph-ref.py <LiberationSans-Regular.ttf 경로> [DejaVuSans.ttf 경로]
      (fontTools 필요. Liberation 은 https://github.com/liberationfonts/liberation-fonts 의 2.1.x,
       DejaVu Sans 는 matplotlib 에 들어 있는 것 — 빼면 matplotlib 에서 찾습니다.)

- CAD 가 PDF 로 내보낼 때 글자를 「윤곽선(칠한 도형)」 으로 바꾸면 PDF 에 글자 층이 없습니다.
  도구는 그 도형을 이 참조 글자와 모양으로 맞춰 봐서 글자로 되돌립니다(js/logic.js 4-1).
- 참조 글꼴은 Liberation Sans(SIL Open Font License 1.1)입니다. 산세리프(Arial · Helvetica 계열)라
  CAD 도면 글꼴과 모양이 가깝습니다. 글꼴 파일 자체가 아니라 글자 윤곽을 꺾은선으로 바꾼 좌표만 넣습니다.
- 「1」 은 Liberation 에는 아래 받침이 있지만 Helvetica 계열에는 없어서, 받침 없는 「1」 을 직접 그린 다각형으로 하나 더 넣었습니다(이름 1#h).
- CAD 글꼴마다 「a」 「,」 숫자 모양이 조금씩 달라, 두 번째 산세리프 DejaVu Sans(Bitstream Vera 라이선스)의 숫자 · 기호 · 자주 쓰는 글자를
  「글자#dv」 이름으로 더 넣었습니다(가장 닮은 것을 고름).
- 좌표 단위: 대문자 높이 = 100, 기준선 y = 0, 위로 +.
"""
import json
import sys
from pathlib import Path

from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import RecordingPen

ROOT = Path(__file__).resolve().parent.parent
CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,+-±()/°ØΦφ×%#*:=<>[]'\""
EXTRA = {
    # 헬베티카형 「1」(받침 없음)
    '1#h': [[[33, 0], [46, 0], [46, 100], [36, 100], [31, 91], [22, 84], [9, 79], [9, 69], [22, 73], [33, 79]]],
}


DV_CHARS = "0123456789.,+-±()°Øx/aAzRMCDEFHLNPST#"


def outlines(path, chars=CHARS):
    f = TTFont(path)
    gs = f.getGlyphSet()
    cmap = f.getBestCmap()
    cap = getattr(f['OS/2'], 'sCapHeight', 0)
    if not cap:  # 옛 OS/2 표에는 대문자 높이가 없어 「H」 의 높이로
        from fontTools.pens.boundsPen import BoundsPen
        bp = BoundsPen(gs)
        gs[cmap[ord('H')]].draw(bp)
        cap = bp.bounds[3]
    out = {}
    for ch in chars:
        gn = cmap.get(ord(ch))
        if not gn:
            print('없음:', ch, file=sys.stderr)
            continue
        pen = RecordingPen()
        gs[gn].draw(pen)
        polys, cur, last = [], None, None
        for op, args in pen.value:
            if op == 'moveTo':
                cur, last = [args[0]], args[0]
            elif op == 'lineTo':
                cur.append(args[0]); last = args[0]
            elif op == 'qCurveTo':  # TrueType 2차 곡선(중간 점은 암시)
                offs, end, on = list(args[:-1]), args[-1], last
                for i, o in enumerate(offs):
                    nxt = ((o[0] + offs[i + 1][0]) / 2, (o[1] + offs[i + 1][1]) / 2) if i < len(offs) - 1 else end
                    for k in range(1, 5):
                        t = k / 4
                        cur.append(((1 - t) ** 2 * on[0] + 2 * (1 - t) * t * o[0] + t * t * nxt[0],
                                    (1 - t) ** 2 * on[1] + 2 * (1 - t) * t * o[1] + t * t * nxt[1]))
                    on = nxt
                last = end
            elif op == 'curveTo':
                c1, c2, p3 = args
                p0 = last
                for k in range(1, 5):
                    t = k / 4
                    a, b, c, d = (1 - t) ** 3, 3 * t * (1 - t) ** 2, 3 * t * t * (1 - t), t ** 3
                    cur.append((a * p0[0] + b * c1[0] + c * c2[0] + d * p3[0], a * p0[1] + b * c1[1] + c * c2[1] + d * p3[1]))
                last = p3
            elif op in ('closePath', 'endPath'):
                if cur and len(cur) > 2:
                    polys.append(cur)
                cur = None
        out[ch] = [[[round(x * 100 / cap), round(y * 100 / cap)] for x, y in p] for p in polys]
    return out


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    g = outlines(sys.argv[1])
    g.update(EXTRA)
    dv = sys.argv[2] if len(sys.argv) > 2 else None
    if not dv:
        try:
            import matplotlib
            dv = str(Path(matplotlib.__file__).parent / 'mpl-data' / 'fonts' / 'ttf' / 'DejaVuSans.ttf')
        except ImportError:
            dv = None
    if dv and Path(dv).exists():
        for k, v in outlines(dv, DV_CHARS).items():
            g[k + '#dv'] = v
    flat = {k: [[c for pt in p for c in pt] for p in v] for k, v in g.items()}
    body = json.dumps(flat, ensure_ascii=False, separators=(',', ':'))
    js = ('/* 윤곽선 글자 읽기용 참조 글자표 — scripts/make-glyph-ref.py 가 만든 파일(직접 고치지 마세요)\n'
          '   출처: Liberation Sans Regular 2.1.5 (SIL Open Font License 1.1, https://github.com/liberationfonts/liberation-fonts)\n'
          '         + DejaVu Sans 일부(「글자#dv」, Bitstream Vera 라이선스 · https://dejavu-fonts.github.io)\n'
          '   글꼴 파일이 아니라 글자 윤곽을 꺾은선으로 바꾼 좌표입니다. 단위: 대문자 높이 100, 기준선 0, 위로 +.\n'
          '   글자마다 [바깥·구멍 윤곽 …], 윤곽 하나 = [x1, y1, x2, y2, …]. 「1#h」 = 받침 없는 1(직접 그림). */\n'
          '(function (root, data) {\n'
          "  if (typeof module === 'object' && module.exports) module.exports = data;\n"
          '  else root.DRGlyphRef = data;\n'
          "})(typeof self !== 'undefined' ? self : this, " + body + ');\n')
    (ROOT / 'js' / 'glyph-ref.js').write_text(js, encoding='utf-8')
    print('js/glyph-ref.js', len(g), '글자', len(js) // 1024, 'KB')


if __name__ == '__main__':
    main()
