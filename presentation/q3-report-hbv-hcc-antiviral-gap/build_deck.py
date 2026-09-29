# -*- coding: utf-8 -*-
"""Build the Q3 progress-report deck from the MOHW template (python-pptx)."""
import copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_MARKER_STYLE, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn
from pptx.text.text import Font
from lxml import etree

TEMPLATE = sys.argv[1]
OUT = sys.argv[2]

FONT = '微軟正黑體'
NAVY = RGBColor(0x1F, 0x3B, 0x57)
BLUE = RGBColor(0x2F, 0x6D, 0xB5)
RED = RGBColor(0xC6, 0x3D, 0x2F)
AMBER = RGBColor(0xD9, 0x8E, 0x1B)
GREEN = RGBColor(0x2E, 0x8B, 0x57)
GREY = RGBColor(0x6B, 0x72, 0x80)
LGREY = RGBColor(0xDA, 0xDD, 0xE2)
LIGHT = RGBColor(0xF3, 0xF5, 0xF8)
LBLUE = RGBColor(0xE6, 0xEF, 0xF9)
LRED = RGBColor(0xFB, 0xEB, 0xE8)
TEXT = RGBColor(0x2B, 0x2B, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ICE = RGBColor(0xBF, 0xD6, 0xEE)

# ---------------------------------------------------------------- helpers
def _ensure_ea(rPr, name=FONT):
    latin = rPr.find(qn('a:latin'))
    if latin is None:
        Font(rPr).name = name
        latin = rPr.find(qn('a:latin'))
    latin.set('typeface', name)
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = etree.SubElement(rPr, qn('a:ea'))
        latin.addnext(ea)
    ea.set('typeface', name)


def set_run_font(run, size=None, bold=None, color=None, italic=None, name=FONT):
    f = run.font
    f.name = name
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color is not None:
        f.color.rgb = color
    _ensure_ea(run._r.get_or_add_rPr(), name)


def _fill_paragraphs(tf, paras, size, bold, color, align, line_spacing=None, space_after=None):
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        if space_after is not None:
            p.space_after = Pt(space_after)
        runs = para if isinstance(para, list) else [(para, {})]
        for rt, opts in runs:
            r = p.add_run()
            r.text = rt
            set_run_font(r, size=opts.get('size', size), bold=opts.get('bold', bold),
                         color=opts.get('color', color), italic=opts.get('italic'))


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=TEXT, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, margin=0.0, line_spacing=None, space_after=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = anchor
    paras = text if isinstance(text, list) else [text]
    _fill_paragraphs(tf, paras, size, bold, color, align, line_spacing, space_after)
    return tb


def add_bullets(slide, x, y, w, h, items, size=14, color=TEXT, bullet_color=NAVY,
                space_after=6, line_spacing=1.08, indent=0.24):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    for i, item in enumerate(items):
        level = 0
        runs = item
        if isinstance(item, tuple):
            runs, level = item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        pPr = p._p.get_or_add_pPr()
        marL = int(Inches(indent + 0.28 * level))
        pPr.set('marL', str(marL))
        pPr.set('indent', str(-int(Inches(indent))))
        buClr = etree.SubElement(pPr, qn('a:buClr'))
        srgb = etree.SubElement(buClr, qn('a:srgbClr'))
        srgb.set('val', str(bullet_color))
        buFont = etree.SubElement(pPr, qn('a:buFont'))
        buFont.set('typeface', 'Arial')
        buChar = etree.SubElement(pPr, qn('a:buChar'))
        buChar.set('char', '•' if level == 0 else '–')
        runs = runs if isinstance(runs, list) else [(runs, {})]
        for rt, opts in runs:
            r = p.add_run()
            r.text = rt
            set_run_font(r, size=opts.get('size', size), bold=opts.get('bold', False),
                         color=opts.get('color', color))
    return tb


def add_rect(slide, x, y, w, h, fill=LIGHT, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None,
             text=None, size=14, bold=False, color=TEXT, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, margin=0.08, line_spacing=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    tf = s.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = anchor
    if text is not None:
        paras = text if isinstance(text, list) else [text]
        _fill_paragraphs(tf, paras, size, bold, color, align, line_spacing)
    return s


def add_circle_num(slide, x, y, d, num, fill=NAVY, color=WHITE, size=16):
    return add_rect(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL, text=str(num),
                    size=size, bold=True, color=color, margin=0)


def add_pill(slide, x, y, w, h, text, fill=GREEN, color=WHITE, size=11):
    return add_rect(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5,
                    text=text, size=size, bold=True, color=color, margin=0)


def add_table(slide, x, y, w, data, col_widths, row_height=0.42, font_size=12, header_fill=NAVY,
              header_color=WHITE, zebra=True, align_cols=None, first_col_bold=False):
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_height * rows))
    tbl = gf.table
    tbl.first_row = True
    tbl.horz_banding = False
    tbl.vert_banding = False
    for j, cw in enumerate(col_widths):
        tbl.columns[j].width = Inches(cw)
    for i in range(rows):
        tbl.rows[i].height = Inches(row_height)
        for j in range(cols):
            cell = tbl.cell(i, j)
            cell.margin_left = cell.margin_right = Inches(0.07)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            val = data[i][j]
            if isinstance(val, list) and val and isinstance(val[0], list):
                paras = val
            else:
                paras = [val if isinstance(val, list) else [(str(val), {})]]
            for k, runs in enumerate(paras):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                for rt, opts in runs:
                    r = p.add_run()
                    r.text = rt
                    set_run_font(r, size=opts.get('size', font_size),
                                 bold=opts.get('bold', (i == 0) or (first_col_bold and j == 0)),
                                 color=opts.get('color', header_color if i == 0 else TEXT))
                if align_cols:
                    p.alignment = align_cols[j]
                else:
                    p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.fore_color.rgb = LIGHT if (zebra and i % 2 == 0) else WHITE
    return gf


def fix_chart_fonts(chart):
    chart.font.name = FONT
    for el in chart._chartSpace.iter():
        if el.tag in (qn('a:defRPr'), qn('a:rPr'), qn('a:endParaRPr')):
            _ensure_ea(el)


def style_value_axis(ch, vmin=0, vmax=None, major=None, fmt=None, size=10, hide_labels=False, title=None):
    va = ch.value_axis
    va.minimum_scale = vmin
    if vmax is not None:
        va.maximum_scale = vmax
    if major is not None:
        va.major_unit = major
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = LGREY
    va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(size)
    va.tick_labels.font.color.rgb = GREY
    if fmt:
        va.tick_labels.number_format = fmt
        va.tick_labels.number_format_is_linked = False
    if hide_labels:
        va.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    if title:
        va.has_title = True
        va.axis_title.text_frame.text = title
        r = va.axis_title.text_frame.paragraphs[0].runs[0]
        set_run_font(r, size=size, bold=False, color=GREY)
    return va


def style_cat_axis(ch, size=10.5, color=TEXT):
    ca = ch.category_axis
    ca.has_major_gridlines = False
    ca.format.line.color.rgb = RGBColor(0xBF, 0xC4, 0xCC)
    ca.tick_labels.font.size = Pt(size)
    ca.tick_labels.font.color.rgb = color
    return ca


def set_chart_title(ch, text, size=13, color=NAVY):
    ch.has_title = True
    ch.chart_title.text_frame.text = text
    p = ch.chart_title.text_frame.paragraphs[0]
    set_run_font(p.runs[0], size=size, bold=True, color=color)
    ch.chart_title.include_in_layout = False


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ---------------------------------------------------------------- template handling
prs = Presentation(TEMPLATE)
slides = list(prs.slides)
S_TITLE, S_OVERVIEW, S_BG, S_DIFF, S_QA = slides

# slide-number placeholder to clone onto new slides (taken from template slide 2)
SLDNUM = None
for shp in S_OVERVIEW.shapes:
    if shp.is_placeholder and shp.placeholder_format.type is not None and 'SLIDE_NUMBER' in str(shp.placeholder_format.type):
        SLDNUM = copy.deepcopy(shp._element)
assert SLDNUM is not None

TITLE_GEOM = (Emu(477726), Emu(262947), Emu(10515600), Emu(810029))


def set_title(shape, text, size=40):
    tf = shape.text_frame
    p = tf.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    r = p.add_run()
    r.text = text
    set_run_font(r, size=size, bold=True)


def strip_template_extras(slide):
    """Remove the red '如不敷使用' note and the empty content placeholder."""
    for shp in list(slide.shapes):
        if shp.has_text_frame and '如不敷使用' in shp.text_frame.text:
            shp._element.getparent().remove(shp._element)
        elif shp.is_placeholder and shp.placeholder_format.idx == 1:
            shp._element.getparent().remove(shp._element)


def new_slide(title):
    layout = [l for l in prs.slide_layouts if l.name == '標題及內容'][0]
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 1:
            ph._element.getparent().remove(ph._element)
    t = s.shapes.title
    t.left, t.top, t.width, t.height = TITLE_GEOM
    set_title(t, title)
    el = copy.deepcopy(SLDNUM)
    el.find('.//' + qn('p:cNvPr')).set('id', str(900 + len(prs.slides)))
    s.shapes._spTree.append(el)
    return s


# ================================================================ Slide 1: title
for shp in S_TITLE.shapes:
    if not shp.has_text_frame:
        continue
    txt = shp.text_frame.text
    if txt.startswith('子計畫'):
        p = shp.text_frame.paragraphs[0]
        runs = p.runs
        runs[0].text = '子計畫：慢性 B 型肝炎患者肝癌診斷前'
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        p.add_line_break()
        r2 = p.add_run()
        r2.text = '抗病毒治療缺口之全國性分析'
        r2._r.insert(0, copy.deepcopy(runs[0]._r.get_or_add_rPr()))
        shp.top = Inches(3.05)
        shp.height = Inches(1.6)
    elif txt.startswith('報告者'):
        p = shp.text_frame.paragraphs[0]
        runs = p.runs
        runs[0].text = '報告者：許耀峻 教授'
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
set_notes(S_TITLE, '（約 20 秒）各位委員好，我是義大醫院許耀峻，報告本子計畫第三季成果：慢性 B 型肝炎患者在肝癌診斷前抗病毒治療缺口的全國性分析。')

# ================================================================ Slide 2: 第三季工作執行進度說明
strip_template_extras(S_OVERVIEW)
set_title(S_OVERVIEW.shapes.title, '第三季工作執行進度說明')
add_text(S_OVERVIEW, 0.6, 1.42, 6.6, 0.4, '本季執行進度', size=20, bold=True, color=NAVY)
progress = [
    ('全國性研究世代建置', '串聯健保申報資料 (2009–2023) 與檢驗結果資料庫 (2015–2023)，納入 34,550 位 HBV 相關肝癌成人'),
    ('肝癌前抗病毒治療率與趨勢分析', '整體治療率、2015–2023 年度趨勢、年齡分層，以及觀察期標準化之敏感度分析'),
    ('照護路徑斷點與監測型態分析', '未治療者分類、HBV DNA 等檢測缺口、治療者與未治療者檢測頻率比較、未治療相關因素'),
    ('存活分析與論文撰寫', '肝癌後存活比較 (Kaplan–Meier)；完成論文撰寫並投稿 Hepatology'),
]
for i, (head, desc) in enumerate(progress):
    y = 2.0 + i * 1.15
    add_circle_num(S_OVERVIEW, 0.6, y + 0.08, 0.48, i + 1)
    add_text(S_OVERVIEW, 1.28, y, 4.55, 1.05,
             [[(head, {'size': 15, 'bold': True, 'color': NAVY})], [(desc, {'size': 12.5, 'color': GREY})]],
             space_after=3)
    add_pill(S_OVERVIEW, 5.95, y + 0.1, 0.9, 0.32, '已完成')
card = add_rect(S_OVERVIEW, 7.5, 1.42, 5.23, 5.25, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_text(S_OVERVIEW, 7.8, 1.62, 4.7, 0.3, '研究主題', size=12, bold=True, color=ICE)
add_text(S_OVERVIEW, 7.8, 1.95, 4.7, 1.15,
         'A Nationwide Analysis of Gaps in Antiviral Treatment Before Hepatocellular Carcinoma in Patients with Chronic Hepatitis B',
         size=14.5, bold=True, color=WHITE, line_spacing=1.05)
add_text(S_OVERVIEW, 7.8, 3.12, 4.7, 0.65, '慢性 B 型肝炎患者肝癌診斷前抗病毒治療缺口之全國性分析',
         size=13, color=ICE)
keynums = [('34,550', '納入之 HBV 相關肝癌患者 (2015–2023 診斷)'),
           ('68.7%', '肝癌診斷前未曾接受抗病毒治療'),
           ('83.8%', '已在肝病照護 ≥1 年之未治療者，無 HBV DNA 檢測紀錄')]
for i, (num, lab) in enumerate(keynums):
    y = 3.95 + i * 0.85
    add_text(S_OVERVIEW, 7.8, y, 1.75, 0.6, num, size=26, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(S_OVERVIEW, 9.6, y, 2.9, 0.6, lab, size=11.5, color=ICE, anchor=MSO_ANCHOR.MIDDLE)
set_notes(S_OVERVIEW, '（約 40 秒）本季完成四項工作：建置全國性世代、治療率趨勢分析、照護路徑斷點與監測型態分析、以及存活分析與論文投稿。右側是三個關鍵數字：34,550 位病人、68.7% 在肝癌前未曾治療、已在照護中的未治療者有 83.8% 從未做過 HBV DNA。')

# ================================================================ Slide 3: 研究背景與目的
strip_template_extras(S_BG)
set_title(S_BG.shapes.title, '研究背景與目的')
add_text(S_BG, 0.6, 1.42, 6.9, 0.4, '研究背景', size=20, bold=True, color=NAVY)
add_bullets(S_BG, 0.6, 1.92, 6.9, 3.3, [
    '慢性 B 型肝炎 (CHB) 是全球肝硬化與肝癌 (HCC) 的主要原因：2022 年約 2.54 億人帶原、約 110 萬人死亡，亞太地區負擔最重',
    '抗病毒治療可顯著降低 CHB 患者的 HCC 風險；台灣自 2003 年起健保給付抗病毒藥物，末期肝病與 HCC 負擔已明顯下降',
    '然而本團隊先前於南台灣單一醫療體系 (2004–2022) 發現，僅 25% 的 HBV 相關 HCC 患者在診斷前曾接受抗病毒治療',
    '族群層級的照護路徑資料仍有限，需超越已進入專科照護者，找出診斷、監測、轉介與治療的上游缺口',
], size=14.5, space_after=8)
add_rect(S_BG, 0.6, 5.35, 6.9, 1.2, fill=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06,
         text='台灣單一保險人全民健保涵蓋 >99% 人口，並可串聯檢驗結果，是在未經篩選的 CHB 族群中檢視整段照護路徑的獨特場域',
         size=13.5, bold=True, color=NAVY, align=PP_ALIGN.LEFT, margin=0.18)
add_rect(S_BG, 7.9, 1.42, 4.83, 5.13, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_text(S_BG, 8.15, 1.62, 4.3, 0.4, '研究目的', size=20, bold=True, color=WHITE)
aims = ['量化全國 HBV 相關 HCC 患者於診斷前的抗病毒治療率，及 2015–2023 年度趨勢',
        '描述未治療者於 HCC 前的肝病照護歷程與監測型態 (ALT、HBV DNA、腹部超音波)',
        '找出與未治療相關的因素，並依預先定義的照護路徑斷點分類未治療者']
for i, a in enumerate(aims):
    y = 2.35 + i * 1.3
    add_circle_num(S_BG, 8.15, y + 0.05, 0.45, i + 1, fill=WHITE, color=NAVY, size=15)
    add_text(S_BG, 8.78, y, 3.7, 1.15, a, size=14, color=WHITE, line_spacing=1.08)
set_notes(S_BG, '（約 45 秒）背景：抗病毒治療能降低肝癌風險，但我們先前單一體系的資料顯示只有四分之一的 HBV 肝癌患者在診斷前治療過。本研究利用全民健保串聯檢驗資料，在未經篩選的全國族群量化這個缺口、描述監測型態，並找出未治療的相關因素與斷點。')

# ================================================================ Slide 4: 研究設計與方法
s4 = new_slide('研究設計與方法')
def flowbox(x, y, w, h, fill, line1, line2, s1=12.5, s2=16):
    add_rect(s4, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1,
             text=[[(line1, {'size': s1, 'color': WHITE})], [(line2, {'size': s2, 'bold': True, 'color': WHITE})]],
             margin=0.06, line_spacing=1.0)
def darrow(x, y):
    add_rect(s4, x, y, 0.3, 0.3, fill=RGBColor(0xA0, 0xA6, 0xB0), shape=MSO_SHAPE.DOWN_ARROW)
flowbox(0.6, 1.42, 4.2, 0.8, BLUE, '健保資料庫中 HBsAg 陽性之慢性 HBV 感染者', 'n = 622,655')
darrow(2.55, 2.27)
flowbox(0.6, 2.62, 4.2, 0.8, BLUE, '首次肝癌診斷（重大傷病證明）', 'n = 47,257')
add_rect(s4, 5.0, 2.42, 1.9, 1.2, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1,
         text=[[('排除', {'size': 11, 'bold': True, 'color': RED})],
               [('2016 年前取得證明', {'size': 10, 'color': TEXT})],
               [('n = 12,700', {'size': 10, 'bold': True, 'color': TEXT})],
               [('未滿 18 歲  n = 7', {'size': 10, 'color': TEXT})]], margin=0.08, align=PP_ALIGN.LEFT, line_spacing=1.0)
darrow(2.55, 3.47)
flowbox(0.6, 3.82, 4.2, 0.8, NAVY, '最終分析世代（指標日 = 證明日 − 90 天）', 'n = 34,550')
darrow(1.45, 4.67)
darrow(3.65, 4.67)
flowbox(0.6, 5.02, 2.0, 0.95, BLUE, '診斷前曾接受抗病毒治療', '10,799 (31.3%)', s1=11.5, s2=15)
flowbox(2.8, 5.02, 2.0, 0.95, RED, '診斷前未接受抗病毒治療', '23,751 (68.7%)', s1=11.5, s2=15)
add_text(s4, 0.6, 6.12, 6.3, 0.55,
         '資料期間：健保申報 2009–2023、檢驗結果 2015–2023；HCC 證明年 2016–2023（指標日 2015/10–2023/9）。IRB：義大醫療 EMRP-115-035',
         size=10.5, color=GREY, line_spacing=1.05)
add_rect(s4, 7.1, 1.42, 5.63, 5.25, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_text(s4, 7.35, 1.58, 5.2, 0.4, '定義與分析方法', size=18, bold=True, color=NAVY)
add_bullets(s4, 7.35, 2.05, 5.15, 4.5, [
    [('慢性 HBV：', {'bold': True, 'color': NAVY}), ('HBsAg 陽性；', {}), ('HCC：', {'bold': True, 'color': NAVY}), ('重大傷病登記新證明 (2016–2023)', {})],
    [('抗病毒治療：', {'bold': True, 'color': NAVY}), ('指標日前任何核苷(酸)類似物或干擾素處方', {})],
    [('進入肝病照護：', {'bold': True, 'color': NAVY}), ('最早帶有肝病診斷之申報紀錄；缺漏或晚於指標日者觀察期設為 0', {})],
    [('未治療者分類：', {'bold': True, 'color': NAVY}), ('(i) 肝病紀錄 <1 年（含完全無紀錄）；(ii) ≥1 年但無 ALT、HBV DNA、超音波紀錄；(iii) 有監測仍未治療', {})],
    [('檢測頻率：', {'bold': True, 'color': NAVY}), ('負二項迴歸 (offset = 記錄病程年數)，校正年齡、性別、肝硬化、診斷年', {})],
    [('未治療因素：', {'bold': True, 'color': NAVY}), ('多變項邏輯斯迴歸；', {}), ('存活：', {'bold': True, 'color': NAVY}), ('Kaplan–Meier 與 log-rank（探索性）', {})],
    [('趨勢：', {'bold': True, 'color': NAVY}), ('Cochran–Armitage 檢定；以觀察期直接標準化作敏感度分析', {})],
], size=12.5, space_after=5)
set_notes(s4, '（約 60 秒）資料來源是健保申報加檢驗結果資料庫。從 62 萬多位 HBsAg 陽性者中，找出 2016 年後新取得重大傷病證明的肝癌成人共 34,550 位；指標日定為證明日前 90 天。依指標日前是否曾有核苷酸類似物或干擾素處方分成治療與未治療兩組。未治療者再依照護路徑分成三類，並比較檢測頻率、分析相關因素與存活。')

# ================================================================ Slide 5: 研究族群特性
s5 = new_slide('研究族群特性（Table 1）')
hdr = [[('特性', {})], [[('全體', {})], [('N = 34,550', {'size': 10.5, 'bold': False})]],
       [[('未治療', {})], [('n = 23,751', {'size': 10.5, 'bold': False})]], [[('已治療', {})], [('n = 10,799', {'size': 10.5, 'bold': False})]]]
rows = [
    ['診斷時年齡，中位數（歲）', '63.4', '63.0', '64.4'],
    ['男性', '78.8%', '79.8%', '76.6%'],
    ['肝硬化', '30.9%', [('15.9%', {'bold': True, 'color': RED})], [('63.9%', {'bold': True, 'color': BLUE})]],
    ['糖尿病', '40.4%', '36.1%', '49.7%'],
    ['ALT，中位數（U/L）', '–', '43', '30'],
    ['AST，中位數（U/L）', '–', '56', '34'],
    ['血小板，中位數（×10³/µL）', '–', '188', '148'],
    ['甲型胎兒蛋白，中位數（ng/mL）', '–', '42.0', '7.4'],
    ['HBV DNA，中位數（log₁₀ IU/mL）', '–', [('4.51', {'bold': True, 'color': RED})], [('0（未檢出）', {'bold': True, 'color': BLUE})]],
]
add_table(s5, 0.6, 1.42, 7.9, [hdr] + rows, [3.1, 1.6, 1.6, 1.6], row_height=0.46, font_size=12.5)
add_text(s5, 0.6, 6.12, 7.9, 0.55,
         '檢驗值取指標日前後 90 天內最接近之結果；所有比較 P<0.001（類別變項：卡方檢定；連續變項：Mann–Whitney U）。HBV DNA 可得比例 47.7%。',
         size=10.5, color=GREY, line_spacing=1.05)
add_text(s5, 8.9, 1.42, 3.83, 0.4, '重點觀察', size=18, bold=True, color=NAVY)
obs = [('15.9% vs 63.9%', RED, '未治療者較少被記錄有肝硬化——反映未進入評估，而非真的沒有肝硬化'),
       ('4.51 vs 0 log₁₀ IU/mL', RED, '未治療者診斷時 HBV DNA 顯著較高；ALT、AST 亦較高 (43 vs 30、56 vs 34 U/L)'),
       ('42.0 vs 7.4 ng/mL', RED, '未治療者 AFP 較高（第三四分位 1,582 vs 36），血小板較高 (188 vs 148)')]
for i, (big, col, lab) in enumerate(obs):
    y = 1.95 + i * 1.6
    add_rect(s5, 8.9, y, 3.83, 1.45, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_text(s5, 9.1, y + 0.12, 3.45, 0.45, big, size=19, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s5, 9.1, y + 0.6, 3.45, 0.8, lab, size=11.5, color=TEXT, line_spacing=1.05)
set_notes(s5, '（約 45 秒）中位年齡 63 歲、近八成男性。值得注意的是：未治療者只有 15.9% 被記錄有肝硬化，而治療者為 63.9%；未治療者確診時 HBV DNA、ALT、AST 和 AFP 都顯著較高，顯示病毒活性未受控制，且病程較晚。')

# ================================================================ Slide 6: 結果一 治療率趨勢
s6 = new_slide('結果一：肝癌診斷前的抗病毒治療率與趨勢')
cd = CategoryChartData()
cd.categories = ['2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023']
cd.add_series('治療比例 (%)', [21.9, 24.3, 26.2, 29.3, 31.7, 31.9, 34.5, 37.1, 39.4])
gf = s6.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.6), Inches(1.38), Inches(7.9), Inches(4.75), cd)
ch = gf.chart
ch.has_legend = False
set_chart_title(ch, '肝癌診斷前曾接受抗病毒治療之比例 (%)，依 HCC 診斷年', size=13)
plot = ch.plots[0]
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format = '0.0'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.ABOVE
dl.font.size = Pt(11)
dl.font.bold = True
dl.font.color.rgb = NAVY
ser = plot.series[0]
ser.smooth = False
ser.format.line.color.rgb = BLUE
ser.format.line.width = Pt(2.75)
ser.marker.style = XL_MARKER_STYLE.CIRCLE
ser.marker.size = 8
ser.marker.format.fill.solid()
ser.marker.format.fill.fore_color.rgb = BLUE
ser.marker.format.line.color.rgb = WHITE
style_value_axis(ch, 0, 50, 10, fmt='0"%"')
style_cat_axis(ch)
fix_chart_fonts(ch)
add_text(s6, 0.6, 6.15, 7.9, 0.55,
         '2015 與 2023 為部分年度（約 3 個月與 9 個月）。各年 Wilson 95% CI 重疊，趨勢依 Cochran–Armitage 檢定 (P<0.0001) 與邏輯斯斜率 (OR 1.10/年，95% CI 1.09–1.11) 判定。',
         size=10.5, color=GREY, line_spacing=1.05)
add_rect(s6, 8.9, 1.42, 3.83, 1.7, fill=LRED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
add_text(s6, 9.1, 1.5, 3.45, 0.8, '68.7%', size=40, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
add_text(s6, 9.1, 2.32, 3.45, 0.7, 'HCC 診斷前未曾接受抗病毒治療（23,751 / 34,550）', size=12, color=TEXT, line_spacing=1.05)
add_rect(s6, 8.9, 3.27, 3.83, 1.7, fill=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
add_text(s6, 9.1, 3.35, 3.45, 0.8, '21.9% → 39.4%', size=28, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
add_text(s6, 9.1, 4.17, 3.45, 0.7, '2015 → 2023 治療比例逐年上升（趨勢 P<0.0001）', size=12, color=TEXT, line_spacing=1.05)
add_bullets(s6, 8.9, 5.15, 3.83, 1.55, [
    '各年齡層 (<50、50–59、60–69、≥70 歲) 皆一致上升 (Figure 2)',
    '以觀察期分布標準化後，上升趨勢不變 (Supp. Figure 2)',
    '即使到 2023 年，仍有六成患者在肝癌前未曾治療',
], size=11.5, space_after=3)
set_notes(s6, '（約 50 秒）核心數字：68.7% 的病人在肝癌診斷前從未接受抗病毒治療。好消息是治療比例從 2015 年的 21.9% 上升到 2023 年的 39.4%，趨勢顯著，各年齡層一致，且以觀察期標準化後趨勢不變；但即使到最近一年，仍有六成病人未治療。')

# ================================================================ Slide 7: 結果二 照護路徑斷點
s7 = new_slide('結果二：未治療者的照護路徑斷點')
add_text(s7, 0.6, 1.38, 12.13, 0.4, '23,751 位未治療者：依 HCC 前的肝病照護歷程分類', size=16, bold=True, color=NAVY)
bx, bw, by, bh = 0.6, 12.13, 1.95, 0.55
segs = [(37.7, RED, '37.7%'), (3.6, AMBER, ''), (58.7, NAVY, '58.7%')]
cx = bx
for pct, col, lab in segs:
    w = bw * pct / 100
    add_rect(s7, cx, by, w, bh, fill=col, text=lab, size=15, bold=True, color=WHITE, margin=0)
    cx += w
add_text(s7, bx + bw * 0.377 - 0.3, by - 0.33, 1.2, 0.3, '3.6%', size=12, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
cards = [
    ('37.7%', 'n = 8,950', RED, '進入肝病照護不到 1 年',
     '其中 7,847 人（占全部未治療者 33.0%）在 HCC 前完全沒有任何帶肝病診斷的就醫紀錄；門檻延長至 2 年僅多歸入 794 人'),
    ('3.6%', 'n = 849', AMBER, '在照護 ≥1 年，但從未檢測',
     '整個肝病照護期間沒有任何 ALT、HBV DNA 或腹部超音波的醫囑紀錄'),
    ('58.7%', 'n = 13,952', NAVY, '已在照護且有監測，仍未治療',
     '在肝病照護 ≥1 年、至少有一項檢測紀錄，卻始終未接受抗病毒治療——是最大宗的族群'),
]
cw, gap = 3.9, 0.215
for i, (big, n, col, head, desc) in enumerate(cards):
    x = 0.6 + i * (cw + gap)
    add_rect(s7, x, 2.85, cw, 2.55, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    add_text(s7, x + 0.25, 3.0, 2.0, 0.6, big, size=30, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s7, x + 2.2, 3.0, 1.5, 0.6, n, size=12.5, color=GREY, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    add_text(s7, x + 0.25, 3.68, cw - 0.5, 0.4, head, size=15, bold=True, color=TEXT)
    add_text(s7, x + 0.25, 4.15, cw - 0.5, 1.2, desc, size=12, color=TEXT, line_spacing=1.08)
add_rect(s7, 0.6, 5.7, 12.13, 0.72, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15,
         text='兩大斷點 →  ① 未進入肝病照護（約 1/3 的未治療者）   ② 已在照護中，卻未完成病毒學評估（見下頁）',
         size=14.5, bold=True, color=WHITE)
set_notes(s7, '（約 50 秒）把 23,751 位未治療者依照護歷程分類：37.7% 在肝癌前進入肝病照護不到一年，其中絕大多數（7,847 人、占全部未治療者三分之一）根本沒有任何肝病相關就醫紀錄；3.6% 在照護中但從未做任何檢測；最大宗 58.7% 已在照護、有做檢測，卻始終沒有走向治療。這指出兩個斷點：沒進入照護、以及在照護中沒完成病毒學評估。')

# ================================================================ Slide 8: 結果三 HBV DNA 檢測缺口
s8 = new_slide('結果三：照護中的 HBV DNA 檢測缺口')
add_rect(s8, 0.6, 1.45, 0.18, 0.18, fill=RED)
add_text(s8, 0.85, 1.4, 1.6, 0.28, 'HCC 前未治療', size=11, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s8, 2.45, 1.45, 0.18, 0.18, fill=BLUE)
add_text(s8, 2.7, 1.4, 1.6, 0.28, 'HCC 前已治療', size=11, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
add_text(s8, 4.4, 1.4, 4.1, 0.28, '無該項檢測紀錄之比例 (%)，依記錄病程長短', size=11, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
panels = [('ALT', [20.9, 6.5], [1.9, 0.1]), ('HBV DNA', [88.7, 83.5], [25.6, 21.1]), ('腹部超音波', [49.2, 23.0], [7.2, 1.2])]
for i, (name, untreated, treated) in enumerate(panels):
    cd = CategoryChartData()
    cd.categories = ['1–<2 年', '≥2 年']
    cd.add_series('HCC 前未治療', untreated)
    cd.add_series('HCC 前已治療', treated)
    x = 0.6 + i * 2.66
    gf = s8.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(1.75), Inches(2.6), Inches(4.3), cd)
    ch = gf.chart
    ch.has_legend = False
    set_chart_title(ch, name, size=12.5, color=TEXT)
    plot = ch.plots[0]
    plot.gap_width = 55
    plot.overlap = -5
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0.0'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(9.5)
    dl.font.color.rgb = TEXT
    for ser, col in zip(plot.series, (RED, BLUE)):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = col
        ser.invert_if_negative = False
    style_value_axis(ch, 0, 100, 20, size=9, hide_labels=(i > 0))
    style_cat_axis(ch, size=10)
    fix_chart_fonts(ch)
add_text(s8, 0.6, 6.12, 7.9, 0.55,
         '對象：肝病照護紀錄 ≥1 年者 25,151 人（未治療 14,801、已治療 10,350）；「無檢測紀錄」指觀察期間內無任何該項醫囑紀錄（Figure 3）。',
         size=10.5, color=GREY, line_spacing=1.05)
add_rect(s8, 8.9, 1.42, 3.83, 1.9, fill=LRED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
add_text(s8, 9.1, 1.5, 3.45, 0.8, '83.8%', size=40, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
add_text(s8, 9.1, 2.3, 3.45, 0.95, '已在肝病照護 ≥1 年的 14,801 位未治療者中，從未有 HBV DNA 定量紀錄（n = 12,402）', size=12, color=TEXT, line_spacing=1.05)
add_bullets(s8, 8.9, 3.5, 3.83, 3.2, [
    '同一群人：無超音波 24.4%、無 ALT 僅 7.3%',
    '58.6%（8,667 人）有 ALT 與超音波，卻沒有 HBV DNA',
    '觀察期由 1–2 年延長至 ≥2 年：ALT 缺測 20.9%→6.5%、超音波 49.2%→23.0%，HBV DNA 僅 88.7%→83.5%',
    [('缺口不在「沒就醫」，而在「沒做病毒學評估」', {'bold': True, 'color': RED})],
], size=12, space_after=5)
set_notes(s8, '（約 50 秒）聚焦已在肝病照護至少一年的 14,801 位未治療者：83.8% 從來沒有 HBV DNA 定量紀錄，相對地沒做 ALT 的只有 7.3%。超過一半的人有 ALT 也有超音波，唯獨沒有 HBV DNA。觀察期拉長後 ALT 與超音波的缺測大幅下降，HBV DNA 幾乎不變。缺口不是沒就醫，而是沒做病毒學評估。')

# ================================================================ Slide 9: 結果四 檢測頻率
s9 = new_slide('結果四：治療者與未治療者的檢測頻率差異')
cd = CategoryChartData()
cd.categories = ['ALT', '腹部超音波', 'HBeAg', 'HBV DNA']
cd.add_series('校正 IRR', [2.16, 2.45, 3.71, 9.73])
gf = s9.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(1.38), Inches(7.5), Inches(4.7), cd)
ch = gf.chart
ch.has_legend = False
set_chart_title(ch, '每年檢測次數：治療者相對未治療者之校正發生率比 (adjusted IRR)', size=13)
plot = ch.plots[0]
plot.gap_width = 45
plot.vary_by_categories = False
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format = '0.00"×"'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(12)
dl.font.bold = True
dl.font.color.rgb = TEXT
ser = plot.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = BLUE
ser.invert_if_negative = False
pt = ser.points[3]
pt.format.fill.solid()
pt.format.fill.fore_color.rgb = RED
style_value_axis(ch, 0, 12, 2, size=10)
style_cat_axis(ch, size=12)
fix_chart_fonts(ch)
add_text(s9, 0.6, 6.12, 7.5, 0.55,
         '負二項迴歸，offset = 記錄病程年數；校正年齡、性別、肝硬化與 HCC 診斷年。對象：肝病照護 >1 年者 25,151 人（治療 10,350、未治療 14,801）。',
         size=10.5, color=GREY, line_spacing=1.05)
tbl = [[[('檢測', {})], [('校正 IRR (95% CI)', {})], [('時間變動模型', {})]],
       ['ALT', '2.16 (2.11–2.20)', '2.26'],
       ['腹部超音波', '2.45 (2.39–2.50)', '2.69'],
       ['HBeAg', '3.71 (3.58–3.85)', '3.94'],
       [[('HBV DNA', {'bold': True, 'color': RED})], [('9.73 (9.31–10.18)', {'bold': True, 'color': RED})], [('11.38', {'bold': True, 'color': RED})]]]
add_table(s9, 8.4, 1.42, 4.33, tbl, [1.33, 1.9, 1.1], row_height=0.38, font_size=11)
add_bullets(s9, 8.4, 3.55, 4.33, 3.15, [
    '治療者每年 HBV DNA 檢測次數為未治療者的 8.8–11.7 倍，各診斷年度差距穩定 (Figure 4C)',
    'ALT、超音波差距僅 2–3 倍，HBeAg 3.6–4.3 倍 → 差距集中於病毒學評估',
    '各年檢測率的表面下降，源於記錄病程（分母）逐年拉長 (Figure 4B)',
    '此差異同時反映「走向治療的路徑」與「治療後的例行監測」，非單純照護不足',
], size=11.5, space_after=4)
set_notes(s9, '（約 40 秒）比較檢測頻率：校正後，治療者的 ALT 與超音波檢測約為未治療者的 2 倍多、HBeAg 近 4 倍，而 HBV DNA 將近 10 倍，且各年度穩定。差距高度集中在病毒學評估。')

# ================================================================ Slide 10: 結果五 未治療相關因素
s10 = new_slide('結果五：與未治療相關的因素（Table 3）')
hdr = [[('共變項', {})], [[('A. 全體世代', {})], [('N = 34,550', {'size': 10.5, 'bold': False})]],
       [[('B. 肝病照護 ≥1 年', {})], [('N = 25,151', {'size': 10.5, 'bold': False})]]]
rows = [
    ['年齡（首次肝病診斷時，每 +10 歲）', '1.24 (1.21–1.27)', '1.09 (1.07–1.12)'],
    ['男性（vs 女性）', '1.30 (1.22–1.37)', '1.12 (1.06–1.20)'],
    ['無糖尿病', '1.85 (1.75–1.92)', '1.25 (1.18–1.32)'],
    ['投保薪資 Q1（最低）vs Q4（最高）', '1.67 (1.56–1.78)', '1.27 (1.18–1.37)'],
    ['投保薪資 Q2 vs Q4', '1.43 (1.34–1.53)', '1.03 (0.96–1.11)'],
    ['投保薪資 Q3 vs Q4', '1.24 (1.16–1.33)', '1.31 (1.22–1.40)'],
    ['中度都市化 vs 高度都市化', '0.96 (0.91–1.01)', '0.98 (0.92–1.04)'],
    ['鄉村 vs 高度都市化', '1.00 (0.93–1.07)', '1.10 (1.02–1.18)'],
]
add_table(s10, 0.6, 1.42, 8.0, [hdr] + rows, [3.6, 2.2, 2.2], row_height=0.47, font_size=12)
add_text(s10, 0.6, 5.75, 8.0, 0.9,
         '數值為校正勝算比 aOR (95% CI)，>1 表示較可能未治療。肝硬化屬給付適應症、位於治療決定的路徑上，故未納入主模型（納入時 aOR 0.112，95% CI 0.106–0.118）。',
         size=10.5, color=GREY, line_spacing=1.05)
add_rect(s10, 8.9, 1.42, 3.83, 3.75, fill=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
add_text(s10, 9.1, 1.58, 3.45, 0.4, '解讀', size=18, bold=True, color=NAVY)
add_bullets(s10, 9.1, 2.05, 3.45, 4.5, [
    '全體世代：年長、男性、無糖尿病、低投保薪資者較可能未治療（薪資呈梯度）；都市化程度無關',
    '限縮於已在照護 ≥1 年者後，所有關聯明顯減弱 → 這些因素主要作用於「是否進入照護」',
    '限縮後鄉村居住與未治療有輕微關聯 (aOR 1.10)，薪資梯度不再單調',
    '世代以 HCC 為條件，估計值描述未治療者的組成，不可外推為一般 CHB 族群的預測因子',
], size=12, space_after=6)
add_rect(s10, 8.9, 5.35, 3.83, 1.32, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06,
         text='關鍵：社會人口因素主要影響「是否進入照護」，而非進入照護後的治療決定',
         size=12.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT, margin=0.18)
set_notes(s10, '（約 45 秒）多變項分析：年長、男性、沒有糖尿病、投保薪資較低的人較可能未治療。但限縮在已進入照護至少一年的人之後，所有關聯都明顯減弱，表示這些因素主要影響「有沒有進入照護」，而不是進入照護後有沒有被治療。')

# ================================================================ Slide 11: 結果六 存活
s11 = new_slide('結果六：肝癌診斷後的存活（探索性分析）')
add_rect(s11, 0.6, 1.42, 3.85, 1.95, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
add_text(s11, 0.8, 1.52, 3.5, 0.35, 'HCC 後中位存活（年）', size=12.5, bold=True, color=GREY)
add_text(s11, 0.8, 1.9, 3.5, 0.7, [[('5.62', {'color': BLUE}), (' vs ', {'size': 18, 'color': GREY}), ('1.71', {'color': RED})]],
         size=30, bold=True, anchor=MSO_ANCHOR.MIDDLE)
add_text(s11, 0.8, 2.62, 3.5, 0.7, '治療者 (95% CI 5.31–5.88) vs 未治療者 (1.65–1.77)；log-rank P<0.0001', size=11, color=TEXT, line_spacing=1.05)
add_rect(s11, 4.65, 1.42, 3.85, 1.95, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
add_text(s11, 4.85, 1.52, 3.5, 0.35, '5 年累積存活率', size=12.5, bold=True, color=GREY)
add_text(s11, 4.85, 1.9, 3.5, 0.7, [[('52.9%', {'color': BLUE}), (' vs ', {'size': 18, 'color': GREY}), ('31.9%', {'color': RED})]],
         size=30, bold=True, anchor=MSO_ANCHOR.MIDDLE)
add_text(s11, 4.85, 2.62, 3.5, 0.7, '10,797 位治療者 vs 23,750 位未治療者（存活狀態可確認者，Figure 5）', size=11, color=TEXT, line_spacing=1.05)
hdr = [[('分層（Supp. Table S3）', {})], [('治療者中位存活（年）', {})], [('未治療者中位存活（年）', {})]]
rows = [
    ['年齡 <50 歲', '7.39', '1.65'], ['年齡 50–64 歲', '6.63', '1.82'], ['年齡 ≥65 歲', '4.61', '1.64'],
    ['男性 / 女性', '5.54 / 5.81', '1.61 / 2.11'],
    ['HCC 前照護紀錄 ≥1 年', '5.79', '2.65'], ['HCC 前照護紀錄 <1 年', '3.35', '0.92'],
    ['無 / 有肝硬化紀錄', '未達 / 4.79', '1.56 / 2.52'],
]
add_table(s11, 0.6, 3.55, 7.9, [hdr] + rows, [3.3, 2.3, 2.3], row_height=0.385, font_size=11.5)
add_rect(s11, 8.9, 1.42, 3.83, 5.25, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
add_text(s11, 9.1, 1.58, 3.45, 0.4, '解讀與限制', size=18, bold=True, color=NAVY)
add_bullets(s11, 9.1, 2.05, 3.45, 4.5, [
    '各年齡、性別、診斷年代、觀察期長短與肝硬化分層，方向皆一致',
    '限 2019 年後診斷者，結果方向不變',
    '屬探索性、未校正之描述分析：受適應症干擾、lead-time 與監測偏差影響，且無腫瘤分期資料',
    [('不應解讀為抗病毒治療的因果效果', {'bold': True, 'color': RED}), ('，但顯示未治療者確診時病情較晚、預後較差', {})],
], size=12, space_after=6)
set_notes(s11, '（約 45 秒）探索性存活分析：治療者中位存活 5.62 年、未治療者 1.71 年；五年存活率 52.9% 對 31.9%，各分層方向一致。必須強調這是未校正、沒有腫瘤分期的描述性比較，不能解讀為治療的因果效果。')

# ================================================================ Slide 12: 討論
s12 = new_slide('討論：兩大斷點與政策意涵')
add_rect(s12, 0.6, 1.38, 5.95, 4.25, fill=LRED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_text(s12, 0.85, 1.52, 5.45, 0.4, '斷點一：未進入肝病照護', size=17, bold=True, color=RED)
add_bullets(s12, 0.85, 2.02, 5.45, 2.45, [
    '1/3 未治療者在 HCC 前無任何肝病相關申報；放寬照護定義後比例幾乎不變',
    '台灣 HBV 多為周產期或幼年感染，這些患者多已帶原數十年卻從未進入照護',
    '2015–16 社區調查 (n=50,909)：34.6% HBsAg 陽性者不知感染；知情者中 43.4% 未就醫，多因無症狀',
    '縮短轉介延遲或加強追蹤的介入對此群體無效——他們根本不在路徑上',
], size=12, space_after=5)
add_rect(s12, 0.85, 4.55, 5.45, 0.9, fill=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12,
         text='對策：HBsAg 陽性結果應直接觸發轉介與結構化長期追蹤，而非交由個人自行處理',
         size=12.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT, margin=0.15)
add_rect(s12, 6.78, 1.38, 5.95, 4.25, fill=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_text(s12, 7.03, 1.52, 5.45, 0.4, '斷點二：在照護中，卻未做病毒學評估', size=17, bold=True, color=NAVY)
add_bullets(s12, 7.03, 2.02, 5.45, 2.45, [
    '5/6 已在照護的未治療者無 HBV DNA；ALT 屬常規生化，HBV DNA 需刻意開單',
    '健保給付條件：HBV DNA 無法單獨建立治療資格，須先有 ALT 異常或肝硬化；非肝硬化者曾需 ALT >2×ULN 兩次',
    '肝炎論質計酬 (2010 起) 以 ALT、AFP、超音波為指標，HBV DNA 從未納入——給付與品質誘因同時忽略病毒學評估',
    '未治療者僅 15.9% 記錄有肝硬化（治療者 63.9%）：肝硬化這條路徑未被開啟',
], size=12, space_after=5)
add_rect(s12, 7.03, 4.55, 5.45, 0.9, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12,
         text='對策：反射式 (reflex) 或流程驅動的風險評估，並於符合條件時及時治療',
         size=12.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT, margin=0.15)
add_rect(s12, 0.6, 5.88, 12.13, 0.62, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3,
         text='國際比較：韓國 2018 年治療率 46.8%、美國保險族群 37%、國際多中心世代 17% —— 多數 CHB 患者未經治療即發生 HCC，是全球共通現象',
         size=11.5, color=NAVY, margin=0.1)
set_notes(s12, '（約 60 秒）兩個斷點需要不同對策。第一，三分之一的未治療者從未進入肝病照護，社區調查也顯示很多帶原者不知情或知情但不就醫，因此 HBsAg 陽性應直接觸發轉介與結構化追蹤。第二，已在照護中的人五分之六沒做 HBV DNA，這和給付條件以 ALT 為門檻、論質計酬未納入 HBV DNA 有關，需要反射式、流程驅動的病毒學評估與及時治療。')

# ================================================================ Slide 13: 限制與後續
s13 = new_slide('研究限制與後續規劃')
add_text(s13, 0.6, 1.42, 6.9, 0.4, '研究限制', size=20, bold=True, color=NAVY)
add_bullets(s13, 0.6, 1.92, 6.9, 4.7, [
    '世代以 HCC 為條件：關聯不可外推至一般 CHB 族群，也無法估計治療的預防效果',
    '存活比較為描述性；重大傷病資料庫無腫瘤分期',
    '無法得知開單醫師科別，無法直接判定照護場域',
    '申報資料自 2009 年起，2009 年前的肝病照護不可觀察；自費檢驗與用藥未納入',
    '臨床診斷至證明日的間隔無法逐一測得，90 天位移為近似',
    '無法確定個別患者未治療的原因，亦無法判定其是否符合給付條件——正因病毒學評估極少執行',
    '結果受台灣給付框架形塑，未必適用於其他制度',
], size=13, space_after=6)
add_rect(s13, 7.9, 1.42, 4.83, 5.25, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_text(s13, 8.15, 1.58, 4.4, 0.4, '進度總結與後續規劃', size=18, bold=True, color=WHITE)
plan = [('已完成', GREEN, '全國性世代建置、治療率趨勢、照護路徑斷點、檢測頻率與未治療因素分析、存活分析'),
        ('已完成', GREEN, '論文撰寫並投稿 Hepatology（Original Article）'),
        ('下一季', AMBER, '依審查意見補充分析與修訂稿件'),
        ('下一季', AMBER, '將研究發現整理為政策建議：HBsAg 陽性轉介機制、反射式 HBV DNA 檢測'),
        ('規劃中', AMBER, '評估未治療者中符合給付條件之比例（需補足病毒學資料）')]
for i, (tag, col, txt) in enumerate(plan):
    y = 2.12 + i * 0.9
    add_pill(s13, 8.15, y + 0.05, 0.85, 0.3, tag, fill=col, size=10.5)
    add_text(s13, 9.12, y, 3.45, 0.85, txt, size=12, color=WHITE, line_spacing=1.05)
set_notes(s13, '（約 40 秒）限制：世代以肝癌為條件、無腫瘤分期、無法得知科別與自費項目、90 天位移為近似。進度：本季完成所有主要分析並投稿 Hepatology；下一季依審查意見修訂，並整理政策建議。')

# ================================================================ Slide 14: 面臨的困難或需協助之處
strip_template_extras(S_DIFF)
set_title(S_DIFF.shapes.title, '面臨的困難或需協助之處')
diffs = [
    ('資料近用與分析效率',
     '分析須於指定資料中心進行，僅能輸出通過揭露審查之彙整結果，圖表需以匯出的彙整值於中心外重建，迭代分析與審查週期較長。',
     '需協助：協調資料中心之分析時段與輸出審查時程。'),
    ('資料庫欄位限制',
     '健保資料無開單醫師科別、腫瘤分期，亦無自費檢驗與用藥紀錄，限制照護場域與預後的判讀。',
     '需協助：期望與計畫內其他子計畫（癌症登記、醫院電子病歷）建立資料串聯，補足分期與科別資訊。'),
    ('研究發現之政策轉譯',
     '本研究指出給付條件與論質計酬設計均未涵蓋 HBV DNA 評估，且 HBsAg 陽性者缺乏轉介機制。',
     '需協助：計畫辦公室協助安排與衛福部／健保署之政策溝通管道。'),
]
for i, (head, desc, need) in enumerate(diffs):
    y = 1.5 + i * 1.65
    add_rect(S_DIFF, 0.6, y, 12.13, 1.4, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_circle_num(S_DIFF, 0.85, y + 0.45, 0.5, i + 1)
    add_text(S_DIFF, 1.6, y + 0.13, 10.9, 0.4, head, size=15.5, bold=True, color=NAVY)
    add_text(S_DIFF, 1.6, y + 0.52, 10.9, 0.95,
             [[(desc, {})], [(need, {'bold': True, 'color': RED})]], size=12.5, color=TEXT, line_spacing=1.05, space_after=2)
set_notes(S_DIFF, '（約 30 秒）三項困難：資料中心的近用與輸出審查流程影響分析效率；健保資料缺少科別、分期與自費項目，希望能與其他子計畫的資料串聯；以及研究發現的政策轉譯，需要計畫辦公室協助建立與衛福部、健保署的溝通管道。')

# ================================================================ Slide 15: Q & A (template as-is)
set_notes(S_QA, '謝謝各位，敬請指教。')

# ---------------------------------------------------------------- reorder: 困難 + Q&A go last
sldIdLst = prs.slides._sldIdLst
items = list(sldIdLst)
order = items[:3] + items[5:] + [items[3], items[4]]
for el in items:
    sldIdLst.remove(el)
for el in order:
    sldIdLst.append(el)

prs.save(OUT)
print('saved', OUT, 'slides:', len(prs.slides))
