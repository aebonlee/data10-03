"""가상 예시 도면 3종(풍선 없음 / 풍선 있음)과 정답 엑셀을 만듭니다.

실행: python3 scripts/make-samples.py      (reportlab · openpyxl 필요, 글꼴은 matplotlib 에 든 DejaVuSans)

- 실제 회사 도면이 아닙니다. 치수·품번은 모두 지어낸 값입니다.
- 글자는 벡터 PDF 의 글자 층으로 들어갑니다(스캔 그림이 아님). 1단계 도구가 읽는 형태입니다.
- 일부러 넣은 함정: 테두리 구역 번호(1~8, A~F), 표제란·개정란의 숫자, 일반공차 ±0.5, 주기 문장 속 숫자,
  참고 치수 (100), 화살촉(가는 삼각형), 따로 떨어진 공차 글자(「45」 「±0.1」), 위·아래로 쌓인 공차,
  세로로 돌아간 치수, 기하공차 틀의 칸 글자, 용접 기호(삼각형) 옆 숫자, 글자 없이 형상만 있는 용접 기호,
  그리고 도면 C 에는 글자가 아니라 선으로 그려진 치수 「8」 하나(1단계로는 못 읽는 것이 정상 — 2단계 OCR 대상).
- 정답 엑셀은 사람이 풍선을 단 순서(보기별로 묶음)로 번호를 매겼습니다. 도구의 자동 번호 순서와 일부러 다릅니다.
- 도면 D(2차, 2026-10-21)는 실제 도면에서 드러난 형태를 흉내 낸 가상 도면입니다: 글자 층이 없고 글자가 모두 칠한 윤곽선,
  지름 Ø · × · ° · 기하공차 기호(⌖ ▱ ∥ ⊥)는 선으로 그림, 같은 크기로 위아래 쌓인 공차, 따로 떨어진 「±1」 · 「6x」+「M12」,
  비스듬한 글, ISO 용접 기호(기준선 + 점선 · 선으로 그린 삼각형 · a 크기 · 개수x길이 · 기준선 아래 부기 · 양면 · 크기 없는 개선 기호),
  도면에 원래 그려진 품번 풍선(윤곽선 숫자 — 검사 풍선 아님), 검사 풍선은 색 원 + 글자 층 번호 + 빨간 지시선(가운데서 두 토막).

- 도면 E(3차, 2026-10-21)는 새 실제 도면에서 드러난 형태를 지어낸 값으로 흉내 낸 가상 도면입니다: 글자 층은 있으나 CAD 가 자간을 넓혀
  「2 6 2」 「2 7 , 5」 「± 1」 처럼 글자 사이에 빈칸을 넣음, 주기의 일반공차(ISO 13920-BF), 선으로 그은 Φ(꺾은선 원) · ⏤(진직도),
  한 줄 선 글꼴(파랑 · 회색, 도구의 참조와 다른 꼴로 직접 그림)로 쓴 치수 「125」 · 용접 크기 「a4」 · 부기 「DO NOT GRIND」 · 글자 층 「16」 앞의 선 글자 「M」,
  「2x」 개수, 양쪽 각장 「z6 z8」, J 형 용접 기호(세로선 + 호, 크기 없음), 함정(꺾은선 원 안 품번 「5」, 숨은선이 붙은 긴 외형선 + 작은 네모).
  정답은 수강생 기준서 양식(No. · Inspection Item · Standard · Tolerance MAX/MIN · Inspection Equipment, 치수 · 각장 시트가 따로, 번호도 따로)이고,
  풍선 있는 PDF 도 치수용(_풍선있음) · 용접용(_풍선있음_용접) 두 개입니다.

출력: samples/*.pdf · samples/*_정답.xlsx · js/sample-pdfs.js(브라우저에서 바로 여는 base64 사본)
"""
import base64
import json
import math
import os
import sys
from pathlib import Path

from reportlab.lib.colors import black, Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / 'samples'


def find_font():
    try:
        import matplotlib
        p = Path(matplotlib.__file__).parent / 'mpl-data' / 'fonts' / 'ttf' / 'DejaVuSans.ttf'
        if p.exists():
            return str(p)
    except ImportError:
        pass
    for p in ['/Library/Fonts/DejaVuSans.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']:
        if os.path.exists(p):
            return p
    sys.exit('DejaVuSans.ttf 를 찾지 못했습니다(pip install matplotlib 로 함께 설치됩니다).')


pdfmetrics.registerFont(TTFont('DV', find_font()))
MM = 72 / 25.4
BLUE = Color(0.1, 0.2, 0.55)
RED = Color(0.78, 0.1, 0.1)


class Sheet:
    """mm · 왼쪽 위 원점 좌표로 그리는 얇은 도우미"""

    def __init__(self, path, w_mm, h_mm):
        self.w, self.h = w_mm, h_mm
        self.c = canvas.Canvas(str(path), pagesize=(w_mm * MM, h_mm * MM))
        self.c.setTitle('예시 도면(가상)')
        self.c.setAuthor('data10-03 sample')
        self.c.setLineWidth(0.35 * MM)

    def P(self, x, y):
        return x * MM, (self.h - y) * MM

    def line(self, x1, y1, x2, y2, w=0.25):
        self.c.setLineWidth(w * MM)
        self.c.line(*self.P(x1, y1), *self.P(x2, y2))

    def rect(self, x, y, w, h, lw=0.35):
        self.c.setLineWidth(lw * MM)
        X, Y = self.P(x, y + h)
        self.c.rect(X, Y, w * MM, h * MM, stroke=1, fill=0)

    def circle(self, x, y, r, lw=0.35):
        self.c.setLineWidth(lw * MM)
        self.c.circle(*self.P(x, y), r * MM, stroke=1, fill=0)

    def text(self, x, y, s, size=12, rot=0, anchor='start', color=black):
        """x, y = 글자 기준선 시작점(mm). anchor=middle 이면 가운데."""
        c = self.c
        c.saveState()
        c.setFillColor(color)
        c.setFont('DV', size)
        X, Y = self.P(x, y)
        c.translate(X, Y)
        if rot:
            c.rotate(rot)
        wpt = pdfmetrics.stringWidth(s, 'DV', size)
        dx = -wpt / 2 if anchor == 'middle' else (-wpt if anchor == 'end' else 0)
        c.drawString(dx, 0, s)
        c.restoreState()
        return wpt / MM  # 글자 폭(mm)

    def width(self, s, size):
        return pdfmetrics.stringWidth(s, 'DV', size) / MM

    def poly(self, pts, close=True, fill=False, lw=0.25):
        c = self.c
        c.setLineWidth(lw * MM)
        p = c.beginPath()
        p.moveTo(*self.P(*pts[0]))
        for q in pts[1:]:
            p.lineTo(*self.P(*q))
        if close:
            p.close()
        c.drawPath(p, stroke=1, fill=1 if fill else 0)

    def arrow(self, x, y, dx, dy, L=3.0, W=0.9):
        """(x,y) 가 촉 끝, (dx,dy) 방향으로 향함 — 가늘고 긴 삼각형(용접 기호와 구분돼야 함)"""
        n = math.hypot(dx, dy)
        ux, uy = dx / n, dy / n
        bx, by = x - ux * L, y - uy * L
        self.c.setFillColor(black)
        self.poly([(x, y), (bx - uy * W / 2, by + ux * W / 2), (bx + uy * W / 2, by - ux * W / 2)], fill=True, lw=0.1)

    # 치수선(가로): x1~x2, y 높이, 글자는 위
    def dim_h(self, x1, x2, y, y_from=None):
        if y_from is not None:
            self.line(x1, y_from, x1, y + 1.5, 0.18)
            self.line(x2, y_from, x2, y + 1.5, 0.18)
        self.line(x1, y, x2, y, 0.18)
        self.arrow(x1, y, -1, 0)
        self.arrow(x2, y, 1, 0)

    def dim_v(self, x, y1, y2, x_from=None):
        if x_from is not None:
            self.line(x_from, y1, x - 1.5, y1, 0.18)
            self.line(x_from, y2, x - 1.5, y2, 0.18)
        self.line(x, y1, x, y2, 0.18)
        self.arrow(x, y1, 0, -1)
        self.arrow(x, y2, 0, 1)

    def gdt(self, x, y, cells, size=11):
        """기하공차 틀: 칸마다 따로 쓴 글자(실제 CAD 출력처럼). 반환: 틀 오른쪽 끝 x"""
        hgt = 7.0
        cx = x
        for s in cells:
            w = max(7.0, self.width(s, size) + 3.0)
            self.rect(cx, y - hgt + 1.8, w, hgt, 0.25)
            self.text(cx + w / 2, y, s, size, anchor='middle')
            cx += w
        return cx

    def weld(self, jx, jy, rx, ry, side_text_left=None, side_text_right=None, size=11, ref_len=26, ref_dir=1):
        """용접 기호: 접합부(jx,jy) → 화살표 → 기준선 시작(rx,ry), 기준선 아래(화살표 쪽)에 필릿 삼각형.
        ref_dir=-1 이면 기준선이 왼쪽으로 뻗습니다."""
        self.line(jx, jy, rx, ry, 0.25)
        self.arrow(jx, jy, jx - rx, jy - ry)
        ex = rx + ref_len * ref_dir
        self.line(rx, ry, ex, ry, 0.25)
        tx = (rx + 9) if ref_dir > 0 else (ex + 12)
        # 직각 이등변 삼각형(다리 4mm) — 기준선 아래
        self.poly([(tx, ry), (tx + 4.0, ry), (tx, ry + 4.0)], close=True, fill=False, lw=0.3)
        if side_text_left:
            self.text(tx - 1.2, ry + 3.6, side_text_left, size, anchor='end')
        if side_text_right:
            self.text(tx + 5.2, ry + 3.6, side_text_right, size)

    def frame(self, title, dwg, scale, mat, notes, rev=True, split_notes=False):
        w, h = self.w, self.h
        self.rect(10, 10, w - 20, h - 20, 0.5)
        # 테두리 구역 번호 — 숫자만 따로 있는 글자(풍선 번호·치수로 오인하기 쉬움)
        ncol = 8 if w > 350 else 6
        for i in range(ncol):
            cx = 10 + (w - 20) / ncol * (i + 0.5)
            self.text(cx, 7.2, str(i + 1), 9, anchor='middle')
            self.text(cx, h - 4.2, str(i + 1), 9, anchor='middle')
            if i:
                self.line(10 + (w - 20) / ncol * i, 10, 10 + (w - 20) / ncol * i, 13, 0.2)
        rows = 'ABCDEF' if h > 250 else 'ABCD'
        for j, ch in enumerate(rows):
            cy = 10 + (h - 20) / len(rows) * (j + 0.5)
            self.text(5.0, cy + 1.5, ch, 9, anchor='middle')
            self.text(w - 5.0, cy + 1.5, ch, 9, anchor='middle')
        # 표제란(오른쪽 아래)
        tw, th = (180, 40) if w > 350 else (130, 34)
        x0, y0 = w - 10 - tw, h - 10 - th
        self.rect(x0, y0, tw, th, 0.4)
        rowsz = th / 4
        for k in range(1, 4):
            self.line(x0, y0 + rowsz * k, x0 + tw, y0 + rowsz * k, 0.2)
        self.line(x0 + tw / 2, y0, x0 + tw / 2, y0 + th, 0.2)
        fs = 9 if w > 350 else 7.5
        cells = [('TITLE', title), ('DWG NO.', dwg), ('SCALE', scale), ('MATERIAL', mat),
                 ('DATE', '2026-10-20'), ('UNIT', 'mm'), ('SHEET', '1/1'), ('REV', 'A')]
        for k, (lab, val) in enumerate(cells):
            cx = x0 + (tw / 2) * (k % 2) + 2
            cy = y0 + rowsz * (k // 2) + rowsz * 0.7
            self.text(cx, cy, lab, fs * 0.8)
            self.text(cx + (24 if w > 350 else 17), cy, val, fs)
        self.text(x0 + 2, y0 - 2.5, 'GENERAL TOLERANCE ±0.5', fs * 0.85)
        self.text(x0 + tw - 2, y0 - 2.5, 'FICTITIOUS SAMPLE DRAWING', fs * 0.8, anchor='end')
        # 개정란(오른쪽 위)
        if rev:
            rx0 = w - 10 - 110
            self.rect(rx0, 13, 110, 14, 0.3)
            self.line(rx0, 20, rx0 + 110, 20, 0.2)
            for lab, dx in (('REV', 3), ('DESCRIPTION', 18), ('DATE', 80)):
                self.text(rx0 + dx, 18.2, lab, 8)
            for val, dx in (('A', 3), ('FIRST ISSUE', 18), ('2026-10-01', 80)):
                self.text(rx0 + dx, 25.2, val, 8)
        # 주기(NOTES) — 문장 속 숫자는 검사 대상이 아님
        nx, ny = 18, h - 10 - th + 2
        self.text(nx, ny, 'NOTES', 10)
        for k, line in enumerate(notes):
            yy = ny + 6 + k * 5.5
            if split_notes:
                # CAD 출력처럼 낱말마다 따로 — 낱말 사이가 글자 높이보다 넓어 한 덩어리로 붙지 않습니다
                cx = nx
                for word in line.split(' '):
                    cx += self.text(cx, yy, word, 9) + 3.6
            else:
                self.text(nx, yy, line, 9)

    def balloon(self, x, y, no, target=None, via=None, r=4.0):
        """풍선: 원 + 번호. target 이 있으면 원 둘레에서 지시선을 긋습니다(via = 꺾이는 점)."""
        c = self.c
        c.setStrokeColor(RED)
        self.circle(x, y, r, 0.3)
        self.text(x, y + 1.6, str(no), 10, anchor='middle', color=RED)
        if target:
            first = via or target
            dx, dy = first[0] - x, first[1] - y
            n = math.hypot(dx, dy)
            sx, sy = x + dx / n * r, y + dy / n * r
            self.line(sx, sy, *first, 0.25)
            if via:
                self.line(*via, *target, 0.25)
        c.setStrokeColor(black)

    def save(self):
        self.c.showPage()
        self.c.save()


class OutlineSheet(Sheet):
    """글자를 칠한 윤곽선으로 그리는 도면(글자 층 없음) — CAD 가 「글자를 윤곽선으로」 내보낸 PDF 를 흉내냅니다.
    윤곽은 DejaVu Sans(도구의 참조 글꼴 Liberation Sans 와 다른 글꼴)에서 꺼냅니다. 글 한 줄 = 칠하기 한 번."""

    def __init__(self, path, w_mm, h_mm):
        super().__init__(path, w_mm, h_mm)
        from fontTools.ttLib import TTFont as FTFont
        self.ft = FTFont(find_font())
        self.gs = self.ft.getGlyphSet()
        self.cmap = self.ft.getBestCmap()
        self.upm = self.ft['head'].unitsPerEm
        self.hmtx = self.ft['hmtx']

    def _glyph_path(self, path, ch, ox, k):
        from fontTools.pens.basePen import BasePen
        gn = self.cmap.get(ord(ch))
        if not gn:
            return 0
        class P(BasePen):
            def _moveTo(s2, pt): path.moveTo(ox + pt[0] * k, pt[1] * k)
            def _lineTo(s2, pt): path.lineTo(ox + pt[0] * k, pt[1] * k)
            def _curveToOne(s2, a, b, c): path.curveTo(ox + a[0] * k, a[1] * k, ox + b[0] * k, b[1] * k, ox + c[0] * k, c[1] * k)
            def _closePath(s2): path.close()
        self.gs[gn].draw(P(self.gs))
        return self.hmtx[gn][0] * k

    def width(self, s, size):
        k = size / self.upm
        return sum(self.hmtx[self.cmap[ord(ch)]][0] * k for ch in s if ord(ch) in self.cmap) / MM

    def text(self, x, y, s, size=12, rot=0, anchor='start', color=black, tracking=0.0):
        c = self.c
        c.saveState()
        c.setFillColor(color)
        X, Y = self.P(x, y)
        c.translate(X, Y)
        if rot:
            c.rotate(rot)
        k = size / self.upm
        wpt = self.width(s, size) * MM + tracking * MM * max(0, len(s) - 1)
        ox = -wpt / 2 if anchor == 'middle' else (-wpt if anchor == 'end' else 0)
        p = c.beginPath()
        for ch in s:
            if ch != ' ':
                self._glyph_path(p, ch, ox, k)
            ox += self.hmtx[self.cmap[ord(ch)]][0] * k + tracking * MM
        c.drawPath(p, stroke=0, fill=1)
        c.restoreState()
        return wpt / MM

    def real_text(self, x, y, s, size, color):
        """글자 층 글자(풍선 번호용 — 풍선 도구가 단 번호처럼)"""
        Sheet.text(self, x, y, s, size, anchor='middle', color=color)

    # 선으로 그린 기호들(글자 높이 hh mm 기준)
    def sym_dia(self, x, y, hh):
        """지름 Ø: 원 + 가운데를 지나는 비스듬한 선. (x, y) = 기준선 왼쪽"""
        r = hh * 0.36
        cx, cy = x + r, y - hh / 2
        self.circle(cx, cy, r, 0.18)
        self.line(cx - r * 0.25, cy + r * 1.35, cx + r * 0.25, cy - r * 1.35, 0.18)
        return 2 * r + hh * 0.25

    def sym_times(self, cx, cy, hh):
        d = hh * 0.28
        self.line(cx - d, cy - d, cx + d, cy + d, 0.18)
        self.line(cx - d, cy + d, cx + d, cy - d, 0.18)

    def sym_deg(self, cx, cy, hh):
        r = hh * 0.14
        self.poly([(cx + r * math.cos(t * math.pi / 4), cy + r * math.sin(t * math.pi / 4)) for t in range(9)], close=False, lw=0.15)

    def gdt_frame(self, x, y, sym, cells, hh=3.0):
        """기하공차 틀: 첫 칸은 선으로 그린 기호, 나머지 칸은 윤곽선 글자. (x, y) = 틀 왼쪽 아래"""
        H = hh * 2.0
        self.rect(x, y - H, H, H, 0.18)
        cx, cy, d = x + H / 2, y - H / 2, hh * 0.45
        if sym == 'pos':
            self.circle(cx, cy, d * 0.75, 0.18); self.line(cx - d * 1.4, cy, cx + d * 1.4, cy, 0.18); self.line(cx, cy - d * 1.4, cx, cy + d * 1.4, 0.18)
        elif sym == 'flat':
            self.poly([(cx - d * 1.1, cy + d * 0.7), (cx + d * 0.5, cy + d * 0.7), (cx + d * 1.1, cy - d * 0.7), (cx - d * 0.5, cy - d * 0.7), (cx - d * 1.1, cy + d * 0.7)], close=False, lw=0.18)
        elif sym == 'par':
            for off in (-0.35, 0.35):
                self.line(cx + d * off - d * 0.35, cy + d * 1.1, cx + d * off + d * 0.35, cy - d * 1.1, 0.18)
        elif sym == 'perp':
            self.line(cx, cy - d * 1.2, cx, cy + d * 1.1, 0.18); self.line(cx - d * 1.1, cy + d * 1.1, cx + d * 1.1, cy + d * 1.1, 0.18)
        xx = x + H
        for cell in cells:
            dia = cell.startswith('Ø')
            body = cell[1:] if dia else cell
            fs = hh / 0.729 * MM
            w = max(H, self.width(body, fs) + (hh if dia else 0) + 2.5)
            self.rect(xx, y - H, w, H, 0.18)
            tx = xx + 1.2
            if dia:
                tx += self.sym_dia(tx, y - H / 2 + hh / 2, hh)
            self.text(tx, y - H / 2 + hh / 2, body, fs)
            xx += w
        return xx

    def weld_ref(self, jx, jy, rx, ry, L, syms, size=11):
        """ISO 2553 용접 기호: 화살표(jx,jy) → 기준선 시작(rx,ry), 기준선(실선) + 0.8mm 아래 점선, 끝에 꼬리.
        syms = [(side, kind, u)] side = -1 위(실선 쪽) / +1 아래, kind = fillet | bevel, u = 기준선 시작에서 기호까지(mm)"""
        self.line(jx, jy, rx, ry, 0.2)
        self.arrow(jx, jy, jx - rx, jy - ry)
        self.line(rx, ry, rx + L, ry, 0.2)
        x = rx + 0.6
        while x < rx + L - 1.2:
            self.line(x, ry + 0.8, x + 1.0, ry + 0.8, 0.13)
            x += 1.8
        self.line(rx + L, ry, rx + L + 3, ry - 3, 0.2)
        self.line(rx + L, ry, rx + L + 3, ry + 3, 0.2)
        for side, kind, u in syms:
            t = rx + u
            base = ry + (0.8 if side > 0 else 0)
            if kind == 'fillet':
                self.poly([(t, base), (t, base + 4.0 * side), (t + 4.0, base)], close=False, lw=0.2)
            else:
                self.line(t, base, t, base + 4.0 * side, 0.2)
                self.line(t, base, t + 3.6, base + 4.0 * side, 0.2)

    def ix_balloon(self, x, y, no, target, r=5.0):
        """풍선 도구(IX 류)가 단 검사 풍선: 분홍 원 + 파란 테두리 + 글자 층 번호 + 빨간 지시선(가운데서 두 토막)"""
        c = self.c
        c.saveState()
        dx, dy = target[0] - x, target[1] - y
        n = math.hypot(dx, dy)
        sx, sy = x + dx / n * r, y + dy / n * r
        mx, my = (sx + target[0]) / 2, (sy + target[1]) / 2
        c.setStrokeColor(Color(1, 0, 0))
        self.line(mx, my, sx, sy, 0.3)
        self.line(mx, my, *target, 0.3)
        c.setStrokeColor(Color(31 / 255, 73 / 255, 125 / 255))
        c.setFillColor(Color(242 / 255, 220 / 255, 219 / 255))
        c.setLineWidth(0.35 * MM)
        c.circle(*self.P(x, y), r * MM, stroke=1, fill=1)
        c.restoreState()
        self.real_text(x, y + 1.8, str(no), 13, black)


# ── 도면 정의 ─────────────────────────────────────────────────
# 검사 항목 = (번호, 구분, 종류, 원문, 공칭, 상한공차, 하한공차, 단위, 풍선(x,y), 지시선 끝(x,y)|None, 꺾임점|None)
# 정답 번호는 「사람이 단 순서」(보기별로 묶음) — 도구의 자동 순서와 일부러 다릅니다.

def drawing_a(s, items_only=False):
    """A3 브래킷 BRK-100 — 치수·공차·쌓인 공차·끼워맞춤·기하공차·나사·각도·거칠기"""
    # 정면도: 판 120 x 80 (60..180, 60..140)
    s.c.roundRect(*s.P(60, 140), 120 * MM, 80 * MM, 5 * MM, stroke=1, fill=0)
    s.circle(120, 100, 6)
    s.circle(97.5, 100, 4.25)
    s.circle(142.5, 100, 4.25)
    for cx in (97.5, 120, 142.5):
        s.line(cx, 92, cx, 108, 0.13)
    s.line(52, 100, 188, 100, 0.13)
    # 측면도(SECTION A-A): 두께 15 (230..245), 위쪽 30° 경사
    s.poly([(230, 60), (241, 60), (245, 64), (245, 140), (230, 140)], close=True, lw=0.35)
    s.line(230, 96, 245, 96, 0.18)
    s.line(230, 104, 245, 104, 0.18)
    s.line(226, 100, 249, 100, 0.13)
    s.text(237.5, 50, 'SECTION A-A', 11, anchor='middle')
    # 데이텀 A(네모 안 글자) — 검사 항목 아님
    s.rect(234.5, 148, 6, 6, 0.25)
    s.text(237.5, 152.6, 'A', 10, anchor='middle')
    s.line(237.5, 148, 237.5, 140, 0.2)

    # 1 가로 120
    s.dim_h(60, 180, 156, y_from=141)
    s.text(120, 154, '120', 12, anchor='middle')
    # 2 세로 80 (90° 돌린 글자)
    s.dim_v(46, 60, 140, x_from=59)
    s.text(44, 100, '80', 12, rot=90, anchor='middle')
    # 3 구멍 간격 45 ±0.1 — 두 글자 조각으로 따로
    s.dim_h(97.5, 142.5, 50, y_from=95)
    w45 = s.width('45', 12)
    s.text(113.5, 48, '45', 12)
    s.text(113.5 + w45 + 0.9, 48, '±0.1', 12)
    # 4 Ø12 H7 지시선
    s.line(124.2, 95.4, 136, 76, 0.2)
    s.arrow(124.2, 95.4, -0.53, 0.85)
    s.line(136, 76, 156, 76, 0.2)
    s.text(137, 74.5, 'Ø12 H7', 12)
    # 5 2-Ø8.5
    s.line(145.5, 103.2, 160, 120, 0.2)
    s.arrow(145.5, 103.2, -0.65, -0.76)
    s.line(160, 120, 182, 120, 0.2)
    s.text(161, 118.5, '2-Ø8.5', 12)
    # 6 R5
    s.line(178.5, 138.5, 186, 146, 0.2)
    s.text(186.5, 148, 'R5', 12)
    # 7 C1 (측면도 모서리)
    s.line(243, 62, 252, 55, 0.2)
    s.text(252.5, 54.5, 'C1', 12)
    # 8 M8x1.25
    s.line(245, 100, 262, 112, 0.2)
    s.line(262, 112, 290, 112, 0.2)
    s.text(263, 110.5, 'M8x1.25', 12)
    # 9 30°±0.5°
    s.text(253, 76, '30°±0.5°', 12)
    s.line(245, 64, 252, 74, 0.15)
    # 10 25 +0.05/-0.02 (위·아래로 쌓인 작은 공차)
    s.dim_h(195, 220, 135, y_from=None)
    s.text(200, 133, '25', 12)
    wd = s.width('25', 12)
    s.text(200 + wd + 0.8, 129.6, '+0.05', 7.5)
    s.text(200 + wd + 0.8, 133.4, '-0.02', 7.5)
    # 11 40 +0.1/0 (한 줄 공차)
    s.dim_h(80, 160, 170, y_from=157)
    s.text(108, 168, '40 +0.1/0', 12)
    # 12 평면도 ⏥ 0.05
    s.gdt(66, 186, ['⏥', '0.05'])
    s.line(72, 179.2, 72, 141, 0.2)
    # 13 직각도 ⊥ 0.1 A
    s.gdt(222, 172, ['⊥', '0.1', 'A'])
    s.line(228, 165.2, 232, 141, 0.2)
    # 14 Ra 1.6
    s.text(84, 70, 'Ra 1.6', 11)
    # 함정: 참고 치수 (100)
    s.text(120, 195, '(100)', 12, anchor='middle')

    s.frame('BRACKET', 'BRK-100', '1:1', 'SS400',
            ['1. UNSPECIFIED CHAMFER C0.5', '2. REMOVE ALL BURRS AND SHARP EDGES', '3. SURFACE TREATMENT: ZINC PLATING 8um'])
    return [
        (1, '치수', '길이', '120', 120, None, None, 'mm', (128, 147), None, None),
        (2, '치수', '길이', '80', 80, None, None, 'mm', (36, 88), None, None),
        (3, '공차', '길이', '45±0.1', 45, 0.1, -0.1, 'mm', (106, 40), None, None),
        (4, '공차', '지름', 'Ø12 H7', 12, 0.018, 0, 'mm', (166, 70), None, None),
        (5, '치수', '지름', '2-Ø8.5', 8.5, None, None, 'mm', (190, 112), None, None),
        (6, '치수', '반지름', 'R5', 5, None, None, 'mm', (199, 145), None, None),
        (7, '공차', '길이', '40 +0.1/0', 40, 0.1, 0, 'mm', (140, 163), None, None),
        (8, '기하공차', '평면도', '⏥ 0.05', None, 0.05, 0, 'mm', (52, 194), (66.5, 186.5), None),
        (9, '기타', '표면 거칠기', 'Ra 1.6', 1.6, None, None, 'μm', (80, 80), None, None),
        (10, '치수', '모따기', 'C1', 1, None, None, 'mm', (264, 46), None, None),
        (11, '공차', '각도', '30°±0.5°', 30, 0.5, -0.5, '°', (282, 68), None, None),
        (12, '치수', '나사', 'M8x1.25', 8, None, None, 'mm', (291, 104), None, None),
        (13, '공차', '길이', '25 +0.05/-0.02', 25, 0.05, -0.02, 'mm', (208, 118), (205, 128), None),
        (14, '기하공차', '직각도', '⊥ 0.1 A', None, 0.1, 0, 'mm', (212, 186), (222.5, 171), (216, 180)),
    ]


def drawing_b(s):
    """A3 용접 구조물 WLD-200 — 용접 기호(크기 글자 · 형상만), 쌓인 공차, 평행도"""
    # 정면도: 바닥판 300 x 20 (50..350, 150..170), 세움판 20 x 130 (190..210, 20..150) — 1:2 축척이라 그림은 절반
    s.rect(50, 130, 150, 12, 0.35)       # 바닥판
    s.rect(117, 60, 16, 70, 0.35)        # 세움판
    s.circle(85, 136, 5)                 # Ø20 (1:2 → 반지름 5mm)
    for cx in (60, 75, 175, 190):
        s.circle(cx, 136, 2.75, 0.25)    # 4-Ø11
    s.c.roundRect(*s.P(240, 142), 60 * MM, 40 * MM, 5 * MM, stroke=1, fill=0)  # 평면도 일부(R10)
    s.line(240, 122, 300, 122, 0.13)

    # 1 300
    s.dim_h(50, 200, 160, y_from=143)
    s.text(125, 158, '300', 12, anchor='middle')
    # 2 150 (세로, 돌린 글자)
    s.dim_v(105, 60, 142, x_from=116)
    s.text(103, 101, '150', 12, rot=90, anchor='middle')
    # 3 100±0.5
    s.dim_h(133, 183, 52, y_from=59)
    s.text(146, 50, '100±0.5', 12)
    # 4 50 +0.2/0 쌓인 공차
    s.dim_h(85, 110, 176, y_from=143)
    s.text(90, 174, '50', 12)
    wd = s.width('50', 12)
    s.text(90 + wd + 0.8, 170.6, '+0.2', 7.5)
    s.text(90 + wd + 0.8, 174.4, '0', 7.5)
    # 5 Ø20
    s.line(88.5, 132.5, 98, 112, 0.2)
    s.arrow(88.5, 132.5, -0.42, 0.9)
    s.line(98, 112, 112, 112, 0.2)
    s.text(98.5, 110.5, 'Ø20', 12)
    # 6 4-Ø11
    s.line(58, 133.6, 52, 118, 0.2)
    s.line(52, 118, 30, 118, 0.2)
    s.text(31, 116.5, '4-Ø11', 12)
    # 7 R10
    s.line(298, 140, 306, 150, 0.2)
    s.text(306.5, 152, 'R10', 12)
    # 8 C3
    s.line(133, 61, 142, 46, 0.2)
    s.text(142.5, 45, 'C3', 12)
    # 9 평행도 ∥ 0.2 B
    s.gdt(232, 184, ['∥', '0.2', 'B'])
    s.line(240, 177.2, 245, 143, 0.2)
    s.rect(312, 118, 6, 6, 0.25)
    s.text(315, 122.6, 'B', 10, anchor='middle')
    # 10 용접: 6 ◺ C.W
    s.weld(117, 125, 150, 92, '6', 'C.W')
    # 11 용접: a4
    s.weld(133, 128, 166, 112, 'a4', None)
    # 12 용접: 5 ◺ 4-50(100) 단속
    s.weld(117, 72, 98, 64, '5', '4-50(100)', ref_len=40, ref_dir=-1)
    # 13 용접: 형상만(글자 없음)
    s.weld(133, 70, 160, 64, None, None)
    # 함정: 참고 치수
    s.text(270, 108, '(250)', 12, anchor='middle')

    s.frame('WELDED BASE', 'WLD-200', '1:2', 'SS400',
            ['1. WELD SIZE 6 UNLESS OTHERWISE SPECIFIED', '2. ALL WELDS TO BE GROUND SMOOTH 2 MM', '3. PAINT AFTER WELDING'], split_notes=True)
    return [
        (1, '치수', '길이', '300', 300, None, None, 'mm', (138, 153), None, None),
        (2, '치수', '길이', '150', 150, None, None, 'mm', (96, 88), None, None),
        (3, '공차', '길이', '100±0.5', 100, 0.5, -0.5, 'mm', (178, 44), None, None),
        (4, '공차', '길이', '50 +0.2/0', 50, 0.2, 0, 'mm', (110, 184), None, None),
        (5, '치수', '지름', 'Ø20', 20, None, None, 'mm', (120, 104), None, None),
        (6, '치수', '지름', '4-Ø11', 11, None, None, 'mm', (26, 108), None, None),
        (7, '치수', '반지름', 'R10', 10, None, None, 'mm', (321, 147), None, None),
        (8, '치수', '모따기', 'C3', 3, None, None, 'mm', (156, 38), None, None),
        (9, '기하공차', '평행도', '∥ 0.2 B', None, 0.2, 0, 'mm', (222, 196), (232.5, 183), None),
        (10, '용접', '필릿', '6 C.W', 6, None, None, 'mm', (190, 84), (176, 92), None),
        (11, '용접', '목두께', 'a4', 4, None, None, 'mm', (200, 106), (192, 112), None),
        (12, '용접', '단속 필릿', '5 4-50(100)', 5, None, None, 'mm', (50, 82), (62, 70), None),
        (13, '용접', '용접 형상', '(용접 형상)', None, None, None, '', (184, 58), (172, 66), None),
    ]


def seg_digit_8(s, x, y, hgt=4.2):
    """글자가 아니라 선으로 그린 「8」 — CAD 가 글자를 선으로 바꿔 내보낸 경우를 흉내냅니다"""
    w = hgt * 0.55
    s.rect(x, y - hgt, w, hgt / 2, 0.3)
    s.rect(x, y - hgt / 2, w, hgt / 2, 0.3)


def drawing_c(s):
    """A4 가로 축 SFT-300 — 끼워맞춤 공차, 동심도·흔들림, 나사 등급, 선으로 그린 치수"""
    cy = 80
    # 축 외형(1:1): Ø25(30..50) Ø30(50..85) Ø40(85..120) Ø30(120..135) M12(135..150)
    for x0, x1, d in ((30, 50, 25), (50, 85, 30), (85, 120, 40), (120, 135, 30), (135, 150, 12)):
        s.rect(x0, cy - d / 2, x1 - x0, d, 0.35)
    s.line(24, cy, 156, cy, 0.13)

    # 1 120 전체 길이
    s.dim_h(30, 150, 125, y_from=106)
    s.text(90, 123, '120', 12, anchor='middle')
    # 2 Ø40
    s.dim_v(102, 60, 100)
    s.text(100, 72, 'Ø40', 12, rot=90, anchor='middle')
    # 3 Ø30 g6
    s.line(67, 65, 60, 46, 0.2)
    s.arrow(67, 65, 0.35, 0.94)
    s.line(60, 46, 40, 46, 0.2)
    s.text(41, 44.5, 'Ø30 g6', 12)
    # 4 Ø25 k6
    s.line(40, 92.5, 34, 104, 0.2)
    s.line(34, 104, 14, 104, 0.2)
    s.text(15, 102.5, 'Ø25 k6', 12)
    # 5 35±0.05
    s.dim_h(50, 85, 112, y_from=96)
    s.text(57, 110, '35±0.05', 12)
    # 6 20 0/-0.1 쌓인 공차
    s.dim_h(30, 50, 30, y_from=67)
    s.text(33, 28, '20', 12)
    wd = s.width('20', 12)
    s.text(33 + wd + 0.8, 24.6, '0', 7.5)
    s.text(33 + wd + 0.8, 28.4, '-0.1', 7.5)
    # 7 M12x1.5-6g
    s.line(145, 74, 156, 58, 0.2)
    s.line(156, 58, 186, 58, 0.2)
    s.text(157, 56.5, 'M12x1.5-6g', 12)
    # 8 C0.5
    s.line(150, 86, 160, 96, 0.2)
    s.text(160.5, 98, 'C0.5', 12)
    # 9 동심도 ◎ Ø0.02 A
    s.gdt(185, 80, ['◎', 'Ø0.02', 'A'])
    s.line(185, 77, 120.5, 66, 0.2)
    # 10 원주 흔들림 ↗ 0.03 A
    s.gdt(185, 116, ['↗', '0.03', 'A'])
    s.line(185, 113, 121, 99, 0.2)
    # 11 15° (모따기 각)
    s.text(156, 72, '15°', 12)
    # 12 선으로 그린 「8」 — 홈 폭
    s.dim_h(120, 128, 54, y_from=64)
    seg_digit_8(s, 122.6, 52.2)
    # 데이텀 A
    s.rect(24, 64, 6, 6, 0.25)
    s.text(27, 68.6, 'A', 10, anchor='middle')

    s.frame('SHAFT', 'SFT-300', '1:1', 'SM45C', ['1. HARDEN 40-45 HRC', '2. UNSPECIFIED CHAMFER C0.3'], rev=False)
    return [
        (1, '치수', '길이', '120', 120, None, None, 'mm', (100, 132), None, None),
        (2, '치수', '지름', 'Ø40', 40, None, None, 'mm', (108, 92), None, None),
        (3, '공차', '지름', 'Ø30 g6', 30, -0.007, -0.020, 'mm', (67, 41), None, None),
        (4, '공차', '지름', 'Ø25 k6', 25, 0.015, 0.002, 'mm', (12, 94), None, None),
        (5, '공차', '길이', '35±0.05', 35, 0.05, -0.05, 'mm', (82, 117), None, None),
        (6, '공차', '길이', '20 0/-0.1', 20, 0, -0.1, 'mm', (22, 20), None, None),
        (7, '공차', '나사', 'M12x1.5-6g', 12, None, None, 'mm', (194, 50), None, None),
        (8, '치수', '모따기', 'C0.5', 0.5, None, None, 'mm', (178, 98), None, None),
        (9, '기하공차', '동심도', '◎ Ø0.02 A', None, 0.02, 0, 'mm', (224, 70), (215, 76), None),
        (10, '기하공차', '원주 흔들림', '↗ 0.03 A', None, 0.03, 0, 'mm', (222, 124), (211, 115), None),
        (11, '치수', '각도', '15°', 15, None, None, '°', (168, 66), None, None),
        (12, '치수', '길이', '8', 8, None, None, 'mm', (132, 45), None, None),
    ]


def drawing_d(s):
    """A3 용접 프레임 FRM-400 — 글자 층 없는 윤곽선 도면(2차: 실제 도면에서 틀린 유형을 흉내)"""
    HH = 12 * 0.729 / MM  # 12pt 글자의 대문자 높이(mm)
    # 형상: 바닥판 · 세움판 · 구멍
    s.rect(60, 150, 220, 20, 0.35)
    s.rect(150, 60, 20, 90, 0.35)
    s.circle(100, 160, 6, 0.3)
    s.line(30, 175, 320, 175, 0.13)
    # 1 640 ±1 (한 덩어리)
    s.dim_h(60, 280, 186, y_from=171)
    s.text(150, 184, '640 ±1', 12)
    # 2 58 + 쌓인 공차 「+1」 「0」(같은 크기 · 90° 돌림)
    s.dim_v(46, 150, 170, x_from=59)
    s.text(44, 166, '58', 12, rot=90)
    s.text(44 - HH - 1.2, 166 - s.width('58', 12) - 1.2, '+1', 12, rot=90)
    s.text(44, 166 - s.width('58', 12) - 1.2, '0', 12, rot=90)
    # 3 Ø45 +0,02 (같은 줄) 위에 +0,08 (같은 크기) — Ø 는 선으로
    s.line(94, 156, 80, 136, 0.2)
    s.arrow(94, 156, 0.57, 0.82)
    s.line(80, 136, 50, 136, 0.2)
    x = 52
    x += s.sym_dia(x, 134.5, HH)
    w60 = s.text(x, 134.5, '45', 12)
    s.text(x + w60 + 2.0, 134.5, '+0,02', 12)
    s.text(x + w60 + 2.0, 134.5 - HH - 2.0, '+0,08', 12)
    # 4 2 × 45° — × · ° 는 선으로, 숫자는 따로
    s.line(170, 62, 182, 50, 0.2)
    s.line(182, 50, 205, 50, 0.2)
    w2 = s.text(183, 48.5, '2', 12)
    s.sym_times(183 + w2 + 2.2, 48.5 - HH / 2, HH)
    w45 = s.text(183 + w2 + 4.4, 48.5, '45', 12)
    s.sym_deg(183 + w2 + 4.4 + w45 + 0.9, 48.5 - HH + 0.5, HH)
    # 5 352 와 ±1 이 따로(넓은 간격)
    s.dim_h(170, 280, 120, y_from=148)
    w352 = s.text(205, 118, '352', 12)
    s.text(205 + w352 + 3.0, 118, '±1', 12)
    # 6 6x M12 (따로)
    s.line(105, 157, 118, 140, 0.2)
    s.line(118, 140, 145, 140, 0.2)
    w7 = s.text(119, 138.5, '6x', 12)
    s.text(119 + w7 + 2.5, 138.5, 'M12', 12)
    # 7 위치도 ⌖ Ø0.3 · 8 평면도 ▱ 1.5 CZ · 9 평행도 ∥ 0.2 B · 10 직각도 ⊥ Ø0.2 B
    s.gdt_frame(225, 75, 'pos', ['Ø0.3'])
    s.gdt_frame(225, 88, 'flat', ['1.5 CZ'])
    s.gdt_frame(225, 101, 'par', ['0.2', 'B'])
    s.gdt_frame(225, 114, 'perp', ['Ø0.2', 'B'])
    for yy in (69, 82, 95, 108):
        s.line(225, yy, 171, yy + 3, 0.15)
    # 11 1240 (세로) · 12 27,5° (비스듬) · 13 1141 ±1
    s.dim_v(300, 60, 170, x_from=281)
    s.text(298, 125, '1240', 12, rot=90)
    s.text(330, 140, '27,5°', 12, rot=25)
    s.dim_h(60, 170, 40, y_from=59)
    s.text(100, 38, '1141 ±1', 12)
    # 용접 A: a3.5 ◺ 4x40 / 아래 부기 2줄
    s.weld_ref(150, 100, 110, 92, 46, [(-1, 'fillet', 18)])
    s.text(112, 90.5, 'a3.5', 12)
    s.text(133, 90.5, '4x40', 12)
    s.text(111, 98, 'DO NOT WELD', 10)
    s.text(111, 102.5, '#12 AREA', 10)
    s.line(110, 104, 138, 104, 0.15)
    # 용접 B: 양면 a6.3 / a6.3
    s.weld_ref(170, 128, 190, 140, 40, [(-1, 'fillet', 18), (1, 'fillet', 18)])
    s.text(193, 138.5, 'a6.3', 12)
    s.text(193, 146.5, 'a6.3', 12)
    # 용접 C: 크기 없는 개선 기호(형상만)
    s.weld_ref(170, 70, 190, 60, 34, [(-1, 'bevel', 16)])
    # 함정: 도면에 원래 그려진 품번 풍선(윤곽선 숫자 12) · 표제란은 frame 이 그림
    s.circle(120, 120, 4.0, 0.3)
    s.text(120 - s.width('12', 12) / 2, 120 + HH / 2, '12', 12)
    s.line(123, 117, 150, 110, 0.2)
    # 함정: 빗금(짧은 평행선 여러 개 — 기하공차 ∥ 로 잡히면 안 됨)
    for k in range(8):
        s.line(62 + k * 3, 168, 66 + k * 3, 152, 0.13)
    # 함정: 네모 속 「//」 옆에 칸 글자가 아닌 낱말(TYP) — 기하공차 항목 아님
    s.rect(330, 60, 7, 7, 0.18)
    for off in (-0.6, 0.6):
        s.line(333.5 + off - 0.5, 61.5, 333.5 + off + 0.5, 65.5, 0.18)
    s.text(338.5, 65.0, 'TYP', 12)
    # 함정: 용접 기준선에 한 끝만 닿는 도면 선(기호가 아님) — 용접 A 기준선 위 · 아래
    s.line(146, 92, 146, 99.5, 0.18)

    s.frame('WELDED FRAME', 'FRM-400', '1:5', 'SS400', ['1. ALL WELDS a3 UNLESS SPECIFIED', '2. PAINT AFTER WELDING'])
    return [
        (1, '공차', '길이', '640 ±1', 640, 1, -1, 'mm', (140, 196), (152, 185), None),
        (2, '공차', '길이', '58 +1/0', 58, 1, 0, 'mm', (30, 150), (40, 158), None),
        (3, '공차', '지름', 'Ø45 +0.08/+0.02', 45, 0.08, 0.02, 'mm', (40, 122), (54, 131), None),
        (4, '치수', '모따기', '2 x 45°', 2, None, None, 'mm', (196, 38), (190, 46), None),
        (5, '공차', '길이', '352 ±1', 352, 1, -1, 'mm', (222, 128), (210, 120), None),
        (6, '치수', '나사', '6x M12', 12, None, None, 'mm', (132, 128), (126, 136), None),
        (7, '기하공차', '위치도', '⌖ Ø0.3', None, 0.3, 0, 'mm', (262, 66), (250, 70), None),
        (8, '기하공차', '평면도', '▱ 1.5 CZ', None, 1.5, 0, 'mm', (270, 82), (258, 83), None),
        (9, '기하공차', '평행도', '∥ 0.2 B', None, 0.2, 0, 'mm', (270, 96), (257, 96), None),
        (10, '기하공차', '직각도', '⊥ Ø0.2 B', None, 0.2, 0, 'mm', (276, 110), (257, 109), None),
        (11, '치수', '길이', '1240', 1240, None, None, 'mm', (312, 112), (300, 118), None),
        (12, '치수', '각도', '27.5°', 27.5, None, None, '°', (350, 126), (336, 135), None),
        (13, '공차', '길이', '1141 ±1', 1141, 1, -1, 'mm', (88, 28), (104, 35), None),
        (14, '용접', '목두께', 'a3.5', 3.5, None, None, 'mm', (100, 80), (114, 87), None),
        (15, '용접', '단속', '4x40', 40, None, None, 'mm', (150, 80), (138, 87), None),
        (16, '주기', '용접 부기', 'DO NOT WELD #12 AREA', None, None, None, '', (96, 112), (112, 103), None),
        (17, '용접', '목두께', 'a6.3', 6.3, None, None, 'mm', (208, 128), (198, 135), None),
        (18, '용접', '목두께', 'a6.3', 6.3, None, None, 'mm', (182, 154), (194, 147), None),
        (19, '용접 형상', '용접 형상', '(용접 형상)', None, None, None, '', (216, 50), (208, 56), None),
    ]


# 한 줄 선 글꼴(가상 CAD 선 글꼴) — 도구의 참조(js/logic.js STROKE_FONT)와 일부러 다른 꼴로 직접 그림(팔각형 0 · 각진 2 · 5 …).
# 단위: 대문자 높이 10, 기준선 0, 위로 +. 글자 사이 간격 = 폭 + 2.4
SFONT = {
    '0': ([[(1, 0), (5, 0), (6, 2), (6, 8), (5, 10), (1, 10), (0, 8), (0, 2), (1, 0)]], 6),
    '1': ([[(1.5, 8.5), (3, 10), (3, 0)]], 4),
    '2': ([[(0, 8), (1, 10), (5, 10), (6, 8), (6, 6.5), (0, 0), (6, 0)]], 6),
    '4': ([[(5, 0), (5, 10), (0, 3.2), (6.2, 3.2)]], 6.2),
    '5': ([[(6, 10), (1, 10), (0.5, 5.5), (4.5, 6), (6, 4.5), (6, 1.5), (4.5, 0), (1.5, 0), (0, 1.2)]], 6),
    '6': ([[(5.5, 9), (4.5, 10), (1.5, 10), (0, 8), (0, 2), (1.5, 0), (4.5, 0), (6, 1.5), (6, 4), (4.5, 5.5), (1.5, 5.5), (0, 4)]], 6),
    'a': ([[(0.5, 6), (2, 7), (4, 7), (5, 6), (5, 0)], [(5, 3.5), (1.5, 3.5), (0, 2.2), (0.3, 0.6), (1.8, 0), (3.5, 0), (5, 1.2)]], 5),
    'D': ([[(0, 0), (0, 10), (3.5, 10), (5.5, 8.5), (6, 5), (5.5, 1.5), (3.5, 0), (0, 0)]], 6),
    'O': ([[(1.2, 0), (5.4, 0), (6.6, 2), (6.6, 8), (5.4, 10), (1.2, 10), (0, 8), (0, 2), (1.2, 0)]], 6.6),
    'N': ([[(0, 0), (0, 10), (6, 0), (6, 10)]], 6),
    'T': ([[(0, 10), (6.4, 10)], [(3.2, 10), (3.2, 0)]], 6.4),
    'G': ([[(6, 8.5), (5, 10), (1.5, 10), (0, 8), (0, 2), (1.5, 0), (5, 0), (6, 1.5), (6, 4.5), (3.5, 4.5)]], 6),
    'R': ([[(0, 0), (0, 10), (4.5, 10), (6, 8.5), (6, 6.8), (4.5, 5.2), (0, 5.2)], [(3.5, 5.2), (6, 0)]], 6),
    'I': ([[(0, 0), (0, 10)]], 0),
    'M': ([[(0, 0), (0, 10), (3.3, 3.5), (6.6, 10), (6.6, 0)]], 6.6),
}
GRAY = Color(0.45, 0.45, 0.45)


class StrokeSheet(Sheet):
    """글자 층(DejaVu) 도면 + 한 줄 선 글꼴 · 꺾은선 기호 · 풍선 도구 풍선(도면 E)"""

    def stroke_text(self, x, y, s, hh=3.0, color=GRAY, lw=0.18):
        """선 글꼴 글: (x, y) = 기준선 시작(mm), hh = 대문자 높이(mm). 반환 = 폭(mm)"""
        k = hh / 10
        c = self.c
        c.saveState()
        c.setStrokeColor(color)
        ox = x
        for ch in s:
            if ch == ' ':
                ox += 5.0 * k
                continue
            strokes, w = SFONT[ch]
            for pl in strokes:
                self.poly([(ox + px * k, y - py * k) for px, py in pl], close=False, lw=lw)
            ox += (w + 2.4) * k
        c.restoreState()
        return ox - x - 2.4 * k

    def poly_circle(self, cx, cy, r, n=24, lw=0.18):
        """곡선이 아니라 꺾은선(점 n 개)으로 그린 원 — 선 글꼴 Φ · 품번 풍선"""
        self.poly([(cx + r * math.cos(2 * math.pi * t / n), cy + r * math.sin(2 * math.pi * t / n)) for t in range(n + 1)], close=False, lw=lw)

    def sym_phi(self, x, y, hh):
        """Φ: 꺾은선 원 + 조금 기운 세로선. (x, y) = 기준선 왼쪽. 반환 = 폭"""
        r = hh * 0.36
        cx, cy = x + r, y - hh / 2
        self.poly_circle(cx, cy, r)
        self.line(cx - r * 0.2, y, cx + r * 0.2, y - hh, 0.18)
        return 2 * r + hh * 0.3

    def real_text(self, x, y, s, size, color):
        Sheet.text(self, x, y, s, size, anchor='middle', color=color)

    weld_ref = OutlineSheet.weld_ref
    ix_balloon = OutlineSheet.ix_balloon


def drawing_e(s):
    """A3 용접 브래킷 WB-500 — 3차: 자간 넓힌 글자 층 · 일반공차 ISO 13920-BF · 선 글꼴 · Φ · ⏤ · 2x · z6 z8 · J 기호.
    반환 = [{'std': 'dim'|'weld', 'no', 'sheet': (Inspection Item, Standard, MAX, MIN, Equipment), 'b': 풍선, 't': 지시선 끝}] — 값은 모두 지어낸 것"""
    HH = 12 * 0.729 / MM
    BLUE2 = Color(0.1, 0.25, 0.75)
    # 형상
    s.rect(60, 70, 120, 40, 0.35)
    s.rect(200, 70, 60, 40, 0.35)
    # D1 262 — 자간 넓힌 글(「2 6 2」), 공차 없음 → 일반공차 ISO 13920-BF(120~400: ±2)
    s.dim_h(60, 180, 58, y_from=69)
    s.text(112, 56, '2 6 2', 12)
    # D2 27,5 ±1 — 자간 넓힌 「2 7 , 5」 와 「± 1」 이 따로
    s.dim_h(200, 212.5, 58, y_from=69)
    w = s.text(199, 56, '2 7 , 5', 12)
    s.text(199 + w + 1.5, 56, '± 1', 12)
    # D3 Φ71 +0,046/0 — Φ 는 꺾은선 원 + 세로선, 공차는 같은 크기로 위아래
    s.line(120, 110, 100, 128, 0.2)
    s.arrow(120, 110, 0.74, -0.67)
    s.line(100, 128, 60, 128, 0.2)
    x = 62
    x += s.sym_phi(x, 126.5, HH)
    w = s.text(x, 126.5, '71', 12)
    s.text(x + w + 2.0, 126.5, '0', 12)
    s.text(x + w + 2.0, 126.5 - HH - 2.0, '+0,046', 12)
    # D4 2x M20 — 「M」 만 선 글꼴(검정 · 글자 층 「20」 바로 앞)
    s.line(150, 110, 140, 150, 0.2)
    s.line(140, 150, 110, 150, 0.2)
    w2 = s.text(112, 148.5, '2x', 12)
    wm = s.stroke_text(112 + w2 + 1.8, 148.5, 'M', HH, color=black)
    s.text(112 + w2 + 1.8 + wm + 0.9, 148.5, '20', 12)
    # D5 진직도 ⏤ 0,3 — 틀 첫 칸에 가로선 하나
    s.rect(272, 120, 7, 7, 0.18)
    s.line(273.8, 123.5, 277.2, 123.5, 0.18)
    s.rect(279, 120, 13, 7, 0.18)
    s.text(280.5, 125.0, '0,3', 12)
    s.line(272, 123.5, 262, 105, 0.15)
    # D6 146 — 파란 한 줄 선 글꼴(글자 층 없음), 공차 없음 → ±2
    s.dim_h(200, 325, 168, y_from=111)
    s.stroke_text(255, 166, '146', HH, color=BLUE2)
    # D7 2x Ø11 H7 — 글자 층 한 줄(개수 2)
    s.circle(230, 90, 6.5, 0.3)
    s.line(234.6, 85.4, 245, 78, 0.2)
    s.line(245, 78, 270, 78, 0.2)
    s.text(246, 76.5, '2x Ø11 H7', 12)
    # 용접 W1 a4 ◺ — 회색 선 글꼴 「a4」, 아래 부기 「DO NOT GRIND」(선 글꼴)
    s.weld_ref(75, 186, 55, 196, 44, [(-1, 'fillet', 26)])
    s.stroke_text(58, 194.2, 'a4', HH)
    s.stroke_text(57, 202.8, 'DO NOT GRIND', 2.6)
    # 용접 W2 z6 z8 ◺(글자 층)
    s.weld_ref(165, 186, 140, 196, 50, [(-1, 'fillet', 36)])
    s.text(142, 194.5, 'z6 z8', 12)
    # 용접 W3 J 형 기호(세로선 + 호) — 크기 글자 없음 → 용접 형상
    s.weld_ref(245, 186, 225, 196, 34, [])
    s.line(244, 196, 244, 191, 0.2)
    s.poly([(244, 193.6), (245.2, 193.5), (246.3, 193.0), (247.1, 192.2), (247.6, 191.1), (247.7, 190.0)], close=False, lw=0.2)
    # 함정: 꺾은선 원 안 품번 「5」(검사 항목 아님) — 용접 W2 기준선 위 오른쪽
    s.poly_circle(181, 186.5, 3.6, lw=0.25)
    s.text(181 - s.width('5', 12) / 2, 186.5 + HH / 2, '5', 12)
    s.line(183.5, 189.2, 186, 192, 0.2)
    # 함정: 숨은선(점선)이 1mm 아래 붙은 긴 외형선 + 그 선에 붙은 작은 네모(화살 · 꼬리 없음 → 용접 기준선 아님)
    s.line(60, 226, 180, 226, 0.35)
    xx = 61
    while xx < 178:
        s.line(xx, 227, xx + 1.2, 227, 0.13)
        xx += 2.0
    s.rect(150, 220, 3, 6, 0.25)

    s.frame('WELD BRACKET', 'WB-500', '1:2', 'S355', ['1. WELDED CONSTRUCTION TOLERANCES ISO 13920-BF', '2. REMOVE ALL BURRS'])
    return [
        {'std': 'dim', 'no': 1, 'sheet': ('Dimension', '262', 2, -2, 'CMM'), 'b': (126, 44), 't': (120, 52)},
        {'std': 'dim', 'no': 2, 'sheet': ('Dimension', '27.5', 1, -1, 'CMM'), 'b': (204, 42), 't': (204, 51)},
        {'std': 'dim', 'no': 3, 'sheet': ('Hole', '71', 0.046, 0, 'CMM'), 'b': (76, 140), 't': (72, 130)},
        {'std': 'dim', 'no': 4, 'sheet': ('TAP', 'M20x2.5', '-', '-', 'Plug Gage'), 'b': (118, 160), 't': (121, 151)},
        {'std': 'dim', 'no': 5, 'sheet': ('Straightness', '0.3', 0, -0.3, 'CMM'), 'b': (300, 123.5), 't': (292, 123.5)},
        {'std': 'dim', 'no': 6, 'sheet': ('Dimension', '146', 2, -2, 'CMM'), 'b': (262, 180), 't': (259, 169)},
        {'std': 'dim', 'no': 7, 'sheet': ('Hole', '11', 0.018, 0, 'CMM'), 'b': (262, 64), 't': (258, 73)},
        {'std': 'weld', 'no': 1, 'sheet': ('Weld', 'a4\n(06x06)', '-', '-', 'Leg Length\nGage'), 'b': (50, 180), 't': (60, 191)},
        {'std': 'weld', 'no': 2, 'sheet': ('Weld', 'z6 z8', '-', '-', 'Steel Ruler'), 'b': (130, 178), 't': (146, 191)},
        {'std': 'weld', 'no': 3, 'sheet': ('DO NOT GRIND', None, '-', '-', 'Visual'), 'b': (44, 210), 't': (56, 201)},
        {'std': 'weld', 'no': 4, 'sheet': ('Weld', None, '-', '-', 'Visual'), 'b': (258, 178), 't': (247, 189)},
    ]


def write_sheet_answer(path, items):
    """수강생 기준서 양식(Standard Inspection Report)과 같은 칸: 머리행 2줄(Tolerance 아래 MAX · MIN), 치수 · 각장 시트가 따로"""
    wb = Workbook()
    wb.remove(wb.active)
    for std, title in (('dim', '치수'), ('weld', '각장')):
        ws = wb.create_sheet(title)
        ws.append(['Standard Inspection Report For Dimension (예시 — 가상 도면, 지어낸 값)'])
        ws.append([])
        ws.append(['No.', 'Inspection Item', 'Standard', 'Tolerance', '', 'Inspection Equipment', "Q'ty of Sample", 'Checking Sample'])
        ws.append(['', '', '', 'MAX', 'MIN', '', '', 1, 2])
        for it in items:
            if it['std'] != std:
                continue
            item, sd, mx, mn, eq = it['sheet']
            ws.append([it['no'], item, sd, mx, mn, eq, '○'])
        for col, w in zip('ABCDEFGHI', (6, 18, 14, 8, 8, 18, 10, 8, 8)):
            ws.column_dimensions[col].width = w
    wb.save(str(path))


DRAWINGS = [
    ('A', '도면A_브래킷', 420, 297, drawing_a),
    ('B', '도면B_용접구조물', 420, 297, drawing_b),
    ('C', '도면C_축', 297, 210, drawing_c),
    ('D', '도면D_윤곽선글자', 420, 297, drawing_d),
    ('E', '도면E_선글꼴_기준서양식', 420, 297, drawing_e),
]


def write_answer(path, title, items):
    wb = Workbook()
    ws = wb.active
    ws.title = '검사항목'
    ws.append(['예시 정답 — ' + title + ' (가상 도면, 사람이 풍선을 단 결과라고 가정)'])
    ws['A1'].font = Font(bold=True)
    head = ['번호', '구분', '종류', '원문', '공칭값', '상한공차', '하한공차', '단위', '풍선 X(mm)', '풍선 Y(mm)']
    ws.append(head)
    for c in ws[2]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F5FA8')
        c.alignment = Alignment(horizontal='center')
    for it in items:
        no, cat, kind, raw, nom, up, lo, unit, (bx, by) = it[:9]
        ws.append([no, cat, kind, raw, nom, up, lo, unit, bx, by])
    for col, w in zip('ABCDEFGHIJ', (6, 10, 12, 18, 9, 9, 9, 6, 11, 11)):
        ws.column_dimensions[col].width = w
    wb.save(str(path))


def main():
    SAMPLES.mkdir(exist_ok=True)
    js = {}
    truth = {}
    sheets = {}
    for key, name, w, h, fn in DRAWINGS:
        plain = SAMPLES / (name + '_풍선없음.pdf')
        ball = SAMPLES / (name + '_풍선있음.pdf')
        K = OutlineSheet if key == 'D' else StrokeSheet if key == 'E' else Sheet
        s = K(plain, w, h)
        items = fn(s)
        s.save()
        if key == 'E':
            # 기준서 양식: 치수 · 용접 풍선 PDF 가 따로(번호도 기준서마다 1번부터)
            ballw = SAMPLES / (name + '_풍선있음_용접.pdf')
            for path_, std in ((ball, 'dim'), (ballw, 'weld')):
                s2 = K(path_, w, h)
                fn(s2)
                for it in items:
                    if it['std'] == std:
                        s2.ix_balloon(it['b'][0], it['b'][1], it['no'], it['t'])
                s2.save()
            write_sheet_answer(SAMPLES / (name + '_정답.xlsx'), items)
            js[key + '_plain'] = base64.b64encode(plain.read_bytes()).decode()
            js[key + '_ball'] = base64.b64encode(ball.read_bytes()).decode()
            js[key + '_ballw'] = base64.b64encode(ballw.read_bytes()).decode()
            sheets[key] = {}
            for std, title in (('dim', '치수'), ('weld', '각장')):
                aoa = [['No.', 'Inspection Item', 'Standard', 'Tolerance', '', 'Inspection Equipment'], ['', '', '', 'MAX', 'MIN', '']]
                aoa += [[it['no']] + list(it['sheet']) for it in items if it['std'] == std]
                sheets[key][title] = aoa
            print(name, '검사 항목', len(items), '개(치수 · 용접) ·', plain.stat().st_size, '/', ball.stat().st_size, '/', ballw.stat().st_size, 'bytes')
            continue
        s2 = K(ball, w, h)
        fn(s2)
        for it in items:
            no, bxy, target, via = it[0], it[8], it[9], it[10]
            if key == 'D':
                s2.ix_balloon(bxy[0], bxy[1], no, target)
            else:
                s2.balloon(bxy[0], bxy[1], no, target, via)
        s2.save()
        write_answer(SAMPLES / (name + '_정답.xlsx'), name, items)
        js[key + '_plain'] = base64.b64encode(plain.read_bytes()).decode()
        js[key + '_ball'] = base64.b64encode(ball.read_bytes()).decode()
        truth[key] = [{'no': it[0], 'cat': it[1], 'kind': it[2], 'raw': it[3], 'nominal': it[4], 'upper': it[5], 'lower': it[6], 'unit': it[7]} for it in items]
        print(name, '검사 항목', len(items), '개 ·', plain.stat().st_size, '/', ball.stat().st_size, 'bytes')
    meta = [{'key': k, 'name': n, 'w': w, 'h': h} for k, n, w, h, _ in DRAWINGS]
    out = ROOT / 'js' / 'sample-pdfs.js'
    out.write_text(
        '/* 예시 도면(가상) — scripts/make-samples.py 로 만든 PDF 의 base64 사본과 정답입니다.\n'
        '   truthSheets = 수강생 기준서 양식 정답(도면 E, 시트별 2차원 배열 — DRLogic.parseSheetTruth 로 읽음).\n'
        '   로컬 파일(file://)로 열어도 「예시 도면 열기」가 되도록 담았습니다. 직접 고치지 말고 스크립트를 다시 돌리세요. */\n'
        '(function (root) {\n  root.DRSample = {\n'
        '    drawings: ' + json.dumps(meta, ensure_ascii=False) + ',\n'
        '    truth: ' + json.dumps(truth, ensure_ascii=False) + ',\n'
        '    truthSheets: ' + json.dumps(sheets, ensure_ascii=False) + ',\n'
        '    pdf: ' + json.dumps(js) + '\n  };\n'
        '  if (typeof module === \'object\' && module.exports) module.exports = root.DRSample;\n'
        '})(typeof self !== \'undefined\' ? self : this);\n', encoding='utf-8')
    print('js/sample-pdfs.js', out.stat().st_size, 'bytes')


if __name__ == '__main__':
    main()
