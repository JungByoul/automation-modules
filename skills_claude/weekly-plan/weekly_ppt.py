# -*- coding: utf-8 -*-
"""
weekly-plan 스킬 전용 pptx 생성기 (표지 없음 · 2페이지)
================================================================
사용법:
  1) 아래 PLAN 딕셔너리를 그 주 내용으로 채운다 (레이아웃 코드는 건드리지 말 것).
  2) python weekly_ppt.py   실행 → OUT_PATH 에 pptx 생성.

레이아웃은 고정, 내용(PLAN)만 매주 교체한다.
  - 1페이지: 전제 박스 + 최우선 카드 2개
  - 2페이지: 월~금 5열 카드(연차일은 회색) + 하단 ActionItem 바
  - ★ 로 시작하는 항목은 붉은 볼드(최우선)로 자동 강조된다.
의존성: python-pptx  (없으면 pip install python-pptx)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR

# ================================================================
# ↓↓↓ 매주 여기만 수정 ↓↓↓
# ================================================================
OUT_PATH = r"C:\Users\정별\2_Personal\2.10_TODO\주간_개인업무_계획\주간계획_N월M주차(YYMMDD~YYMMDD)_v0.1_YYYYMMDD.pptx"

PLAN = {
    "title": "이번주 개인업무 계획 — N월 M주차",
    "subtitle": "YYMMDD(월) ~ YYMMDD(금)  ·  작성 YYYY-MM-DD  ·  작성자 별",
    # 전제: 각 항목 한 줄
    "premise": [
        "수요일(10/28) 연차  →  실질 가용 4일 (월·화·목·금)",
        "MDM 오픈 일정 : 11/2 기존 시스템 block  →  11/4 오픈",
    ],
    # 최우선 카드 2개. 각 카드: {"head": 제목, "lines": [{"t":.., "hi":강조여부}]}
    "top": [
        {"head": "★ 1.  export 데이터 분석", "lines": [
            {"t": "전체데이터 SharePoint 추출 자체는 금방 끝남", "hi": False},
            {"t": "중요한 건 꺼낸 export 데이터를 분석하는 것", "hi": True},
            {"t": "보고 형태로 정리 → 공유 / 논의 가능하게", "hi": False},
        ]},
        {"head": "★ 2.  교육자료 (클로드 코드)", "lines": [
            {"t": "기획·구성 : 오늘(월) 마감  /  제작 : 화요일 마감", "hi": True},
            {"t": "1주차 클코 설치+명령어+mode별 차이", "hi": False},
            {"t": "2주차 skill+자동화툴1(메일 접수대장)", "hi": False},
            {"t": "3주차 하네스 환경+자동화툴2(위키 자동화)", "hi": False},
            {"t": "4주차 본인 업무 구축 워크플로우 공유", "hi": False},
        ]},
    ],
    # 일자별 5열. is_off=True 면 회색(연차 등). ★ 로 시작하는 item은 붉은 볼드.
    "days": [
        {"name": "10/26 월", "off": False, "items": [
            "주간보고 회의 겸 MDM 운영 보고 (오전)",
            "★ 교육자료 기획·구성 끝내기 (오늘 마감)",
            "★ 자동화 패턴(LLM 활용) 정리 끝내기 (오늘 마감)",
            "청라에너지 데이터 수령 → 확인",
            "배포 체크 / 재배포 큐 적체 확인(퇴근 전)",
            "MDM 접수건 처리(삭제건·중복설비 확인)",
        ]},
        {"name": "10/27 화", "off": False, "items": [
            "★ 교육자료 제작 → 마감",
            "설비 추출 : SharePoint로 한번에 (금방 끝남)",
            "★ export 데이터 분석 착수",
            "배포 체크 / 큐 소진 확인",
            "소득세 감면 신청 확인(지윤 선임)",
        ]},
        {"name": "10/28 수", "off": True, "items": ["── 연차 ──"]},
        {"name": "10/29 목", "off": False, "items": [
            "★ MDM 월간보고 제작 착수",
            "★ export 데이터 분석 이어서",
            "포스코인터 : 업무 현황 공유받기 (플젝 중단·영어 없음)",
            "워크샵 : 이슈 체크만 (불참)",
            "RTDB : 이슈 체크만 (참여 아님)",
        ]},
        {"name": "10/30 금", "off": False, "items": [
            "★ MDM 월간보고 제작 이어서",
            "★ export 데이터 분석 → 보고 형태로 정리",
            "배포 최종 체크",
        ]},
    ],
    # 하단 ActionItem 바 (한 줄, ' · '로 구분)
    "actionitems": "MDM 월간보고 제작(29일 착수)  ·  교육자료(기획 월/제작 화)  ·  설비 추출→SharePoint  ·  청라에너지 데이터 확인(견적12월·수행3~4월)  ·  RTDB 이슈 체크",
}
# ================================================================
# ↑↑↑ 매주 여기까지 ↑↑↑  (아래 레이아웃 코드는 수정 불필요)
# ================================================================

NAVY   = RGBColor(0x1F, 0x3A, 0x5F)
BLUE   = RGBColor(0x2E, 0x6D, 0xB4)
ACCENT = RGBColor(0xC2, 0x3B, 0x22)
GRAY   = RGBColor(0x55, 0x55, 0x55)
LIGHT  = RGBColor(0xED, 0xF1, 0xF7)
DARK   = RGBColor(0x22, 0x22, 0x22)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = "맑은 고딕"


def build(plan, out_path):
    prs = Presentation()
    prs.slide_width  = Inches(13.333)
    prs.slide_height = Inches(7.5)
    SW, SH = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]

    def add_box(slide, l, t, w, h, fill=None, line=None):
        sp = slide.shapes.add_shape(1, l, t, w, h)
        sp.shadow.inherit = False
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid(); sp.fill.fore_color.rgb = fill
        if line is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb = line; sp.line.width = Pt(0.75)
        return sp

    def txt(slide, l, t, w, h, lines, anchor=MSO_ANCHOR.TOP):
        tb = slide.shapes.add_textbox(l, t, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = Pt(4); tf.margin_right = Pt(4)
        tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
        first = True
        for ln in lines:
            text = ln.get("t", "")
            size = ln.get("sz", 14); color = ln.get("c", DARK)
            bold = ln.get("b", False); indent = ln.get("ind", 0)
            space = ln.get("sp", 4)
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.space_after = Pt(space); p.space_before = Pt(0); p.level = indent
            r = p.add_run(); r.text = text
            r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = color; r.font.name = FONT
        return tb

    def header(slide, title, sub):
        add_box(slide, 0, 0, SW, Inches(0.95), fill=NAVY)
        add_box(slide, 0, Inches(0.95), SW, Pt(4), fill=ACCENT)
        t = slide.shapes.add_textbox(Inches(0.5), Inches(0.12), Inches(12), Inches(0.75))
        tf = t.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run(); r.text = title
        r.font.size = Pt(24); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT
        p2 = tf.add_paragraph()
        r2 = p2.add_run(); r2.text = sub
        r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xC9, 0xD6, 0xE8); r2.font.name = FONT

    # ---------- 슬라이드 1 ----------
    s1 = prs.slides.add_slide(blank)
    header(s1, plan["title"], plan["subtitle"])

    pre_lines = [{"t": "이번주 전제", "sz": 14, "c": NAVY, "b": True, "sp": 5}]
    for i, p in enumerate(plan["premise"]):
        pre_lines.append({"t": "· " + p, "sz": 13, "c": DARK,
                          "sp": 2 if i == len(plan["premise"]) - 1 else 3})
    prem_h = Inches(0.75 + 0.30 * len(plan["premise"]))
    add_box(s1, Inches(0.5), Inches(1.35), Inches(12.33), prem_h, fill=LIGHT)
    txt(s1, Inches(0.7), Inches(1.45), Inches(12.0), prem_h, pre_lines)

    top_y = Emu(Inches(1.35) + prem_h + Inches(0.25))
    txt(s1, Inches(0.5), top_y, Inches(6), Inches(0.4),
        [{"t": "이번주 최우선", "sz": 16, "c": ACCENT, "b": True, "sp": 0}])

    card_y = Emu(top_y + Inches(0.5))
    card_h = Emu(SH - card_y - Inches(0.4))
    card_w = Inches(6.0)
    xs = [Inches(0.5), Inches(6.83)]
    for card, cx in zip(plan["top"], xs):
        add_box(s1, cx, card_y, card_w, card_h, fill=WHITE, line=BLUE)
        add_box(s1, cx, card_y, card_w, Inches(0.6), fill=BLUE)
        txt(s1, Emu(cx + Inches(0.15)), card_y, Inches(5.7), Inches(0.6),
            [{"t": card["head"], "sz": 15, "c": WHITE, "b": True, "sp": 0}],
            anchor=MSO_ANCHOR.MIDDLE)
        body = []
        for j, ln in enumerate(card["lines"]):
            body.append({"t": "· " + ln["t"], "sz": 13,
                         "c": ACCENT if ln.get("hi") else DARK,
                         "b": bool(ln.get("hi")), "sp": 6})
        txt(s1, Emu(cx + Inches(0.2)), Emu(card_y + Inches(0.75)),
            Inches(5.6), Emu(card_h - Inches(0.9)), body)

    # ---------- 슬라이드 2 ----------
    s2 = prs.slides.add_slide(blank)
    header(s2, "일자별 계획 & ActionItem", "★ = 최우선")

    col_gap = Inches(0.15)
    left0 = Inches(0.4)
    total_w = SW - Inches(0.8)
    n = len(plan["days"])
    col_w = Emu(int((total_w - col_gap * (n - 1)) / n))
    top = Inches(1.35)
    col_h = Inches(4.75)

    x = left0
    for d in plan["days"]:
        hdr_color = GRAY if d.get("off") else NAVY
        add_box(s2, x, top, col_w, Inches(0.5), fill=hdr_color)
        txt(s2, x, top, col_w, Inches(0.5),
            [{"t": d["name"], "sz": 13, "c": WHITE, "b": True, "sp": 0}],
            anchor=MSO_ANCHOR.MIDDLE)
        add_box(s2, x, Emu(top + Inches(0.5)), col_w, Emu(col_h - Inches(0.5)),
                fill=WHITE, line=RGBColor(0xC9, 0xD6, 0xE8))
        lines = []
        for it in d["items"]:
            star = it.startswith("★")
            lines.append({"t": it if star else "• " + it, "sz": 10.5,
                          "c": ACCENT if star else DARK, "b": star, "sp": 7})
        txt(s2, Emu(x + Inches(0.08)), Emu(top + Inches(0.6)),
            Emu(col_w - Inches(0.16)), Emu(col_h - Inches(0.7)), lines)
        x = Emu(x + col_w + col_gap)

    ai_top = Inches(6.35)
    add_box(s2, Inches(0.4), ai_top, total_w, Inches(0.85), fill=LIGHT)
    txt(s2, Inches(0.6), ai_top, Emu(total_w - Inches(0.4)), Inches(0.85), [
        {"t": "회의 연계 ActionItem", "sz": 12, "c": NAVY, "b": True, "sp": 3},
        {"t": plan["actionitems"], "sz": 11, "c": DARK, "sp": 0},
    ], anchor=MSO_ANCHOR.MIDDLE)

    prs.save(out_path)
    print("saved:", out_path)


if __name__ == "__main__":
    build(PLAN, OUT_PATH)
