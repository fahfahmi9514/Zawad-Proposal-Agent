"""
builder.py — Professional PPTX builder for Zawad proposals.
Design: dark navy header, dual logos, teal badge, icon boxes, RTL Arabic.
28-35 slides per proposal matching professional benchmark.
"""

import os
import re
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from lxml import etree

# ── Colors ────────────────────────────────────────────────────────
NAVY   = RGBColor(0x0A, 0x26, 0x47)
MINT   = RGBColor(0x2E, 0xC4, 0xB6)
TEAL   = RGBColor(0x0E, 0x7C, 0x86)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY  = RGBColor(0xF5, 0xF7, 0xFA)
SLATE  = RGBColor(0x1B, 0x2A, 0x3A)
MGRAY  = RGBColor(0xDE, 0xE3, 0xEB)
MGREEN = RGBColor(0xC9, 0xEF, 0xEA)

W, H = 33.87, 19.05

ASSETS = Path(__file__).parent / 'assets'
ZAWAD_LOGO    = str(ASSETS / 'zawad_logo.png')
RIYADAH_CARD  = str(ASSETS / 'riyadah_card.png')
PARTNER_LOGOS = str(ASSETS / 'partner_logos.png')

# ── Language labels ────────────────────────────────────────────────
_LABELS = {
    'ar': {
        'team_title':     'فريق العمل والسير الذاتية',
        'team_sub':       'الكفاءات العُمانية المتخصصة',
        'team_page':      'الصفحة {n}',
    },
    'en': {
        'team_title':     'Team & CVs',
        'team_sub':       'Specialized Omani Competencies',
        'team_page':      'Page {n}',
    },
}

def _lbl(lang, key, **kwargs):
    text = _LABELS.get(lang, _LABELS['ar']).get(key, _LABELS['ar'].get(key, key))
    if kwargs:
        text = text.format(**kwargs)
    return text

def _align_for(lang):
    return PP_ALIGN.LEFT if lang == 'en' else PP_ALIGN.RIGHT

def _is_rtl(lang):
    return lang != 'en'


PAST_PROJECTS = [
    ("البرنامج التنفيذي في الإدارة المحلية", "الأكاديمية السلطانية للإدارة", "83,000", "25", "داخل", "تدريب وكوتشن"),
    ("تطوير القيادات — الدفعة الخامسة", "الأكاديمية السلطانية للإدارة", "119,000", "25", "داخل", "تدريب، كوتشينج"),
    ("البرنامج التنفيذي في التخطيط الاستراتيجي", "وزارة التنمية الاجتماعية", "12,000", "30", "داخل", "تدريب"),
    ("تعزيز القدرات التخطيطية للمحافظات", "الأكاديمية السلطانية للإدارة", "52,000", "50", "داخل", "تدريب، استشارات، كوتشينج"),
    ("تطوير القيادات — الدفعة الرابعة", "الأكاديمية السلطانية للإدارة", "78,000", "25", "داخل", "تدريب، كوتشينج"),
    ("تقييم المشاركين في برنامج اعتماد", "الأكاديمية السلطانية للإدارة", "109,000", "200", "داخل", "تقييم واختيار"),
    ("تمكين للقيادة", "الشركة العُمانية لدرفلة الألمنيوم", "43,000", "32", "داخل", "تدريب، تقييم، كوتشينج"),
    ("التخطيط الاستراتيجي وإدارة التغيير", "وزارة التنمية الاجتماعية", "11,000", "30", "داخل", "تدريب"),
    ("برنامج إمكان 3", "وزارة الخدمة المدنية", "81,900", "120", "داخل", "تدريب، كوتشينج"),
    ("بناء القدرات التخطيطية والاقتصادية", "الأكاديمية السلطانية للإدارة", "52,000", "50", "داخل", "تدريب، استشارات"),
    ("تقارير استشارية", "شركة أوكيو", "10,347", "—", "داخل", "استشارات"),
    ("إعداد المدربين", "الصندوق الاجتماعي للتنمية", "5,810", "25", "داخل", "تدريب"),
]


# ═══════════════════════════════════════════════════════════════════
# LOW-LEVEL HELPERS
# ═══════════════════════════════════════════════════════════════════

def _blank_slide(prs):
    blank = prs.slide_layouts[6]
    return prs.slides.add_slide(blank)


def _rect(slide, x, y, w, h, color, rounded=False, radius=0.08, border=None):
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE as MSO
    stype = MSO.ROUNDED_RECTANGLE if rounded else MSO.RECTANGLE
    shape = slide.shapes.add_shape(stype, Cm(x), Cm(y), Cm(w), Cm(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if border:
        shape.line.color.rgb = border
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    if rounded:
        shape.adjustments[0] = radius
    return shape


def _oval(slide, x, y, w, h, color):
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE as MSO
    shape = slide.shapes.add_shape(MSO.OVAL, Cm(x), Cm(y), Cm(w), Cm(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _txbox(slide, x, y, w, h):
    # STRICT RULE: box auto-sizes to content — never too big, never overflow
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tb.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    tb.text_frame.word_wrap = True
    return tb


def _para(tf, text, size, bold=False, color=SLATE, align=PP_ALIGN.RIGHT,
          rtl=True, italic=False, space_before=0, font_name='Times New Roman'):
    if len(tf.paragraphs) == 1 and not tf.paragraphs[0].text:
        p = tf.paragraphs[0]
        p.clear()
    else:
        p = tf.add_paragraph()
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    _set_rtl(p, rtl)
    run = p.add_run()
    run.text = _fix_spacing(str(text)) if text else ''
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return p


def _set_rtl(paragraph, rtl=True):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('rtl', '1' if rtl else '0')


def _img(slide, path, x, y, w, h):
    if path and os.path.exists(path):
        slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))


def _tf_clear(tf):
    for _ in range(len(tf.paragraphs) - 1):
        p = tf.paragraphs[-1]._p
        p.getparent().remove(p)
    tf.paragraphs[0].clear()
    bodyPr = tf._txBody.bodyPr
    for tag in (qn('a:spAutoFit'), qn('a:normAutofit'), qn('a:lstAutofit')):
        el = bodyPr.find(tag)
        if el is not None:
            bodyPr.remove(el)
    if bodyPr.find(qn('a:noAutofit')) is None:
        etree.SubElement(bodyPr, qn('a:noAutofit'))


def _ar(n):
    """Return number as Western digits (no Arabic-Indic conversion)."""
    return str(n)


_AR = '[\u0600-\u065f\u066a-\u06ef\u06fa-\u06ff\u0750-\u077f\ufb50-\ufdff\ufe70-\ufeff]'

# Characters invalid in XML 1.0 (causes PowerPoint "problem with content" errors)
_INVALID_XML = re.compile(
    '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F￾￿]'
)

def _strip_invalid_xml(text: str) -> str:
    """Remove characters that are illegal in XML 1.0 and corrupt PPTX files."""
    return _INVALID_XML.sub('', text)

_LAT = '[A-Za-z]'
_DIG = '[0-9]'
_AI  = '[٠-٩۰-۹]'  # Arabic-Indic digits ٠١٢٣ / ۰۱۲۳

def _fix_spacing(text):
    """Normalize spacing at Arabic↔Latin, Arabic↔digit, and bracket boundaries."""
    if not text:
        return text

    # Arabic ↔ Latin letters (both directions)
    text = re.sub('(' + _AR + ')(' + _LAT + ')', r'\1 \2', text)
    text = re.sub('(' + _LAT + ')(' + _AR + ')', r'\1 \2', text)

    # Arabic ↔ Western digits (both directions)
    text = re.sub('(' + _AR + ')(' + _DIG + ')', r'\1 \2', text)
    text = re.sub('(' + _DIG + ')(' + _AR + ')', r'\1 \2', text)

    # Arabic ↔ Arabic-Indic digits (both directions)
    text = re.sub('(' + _AR + ')(' + _AI + ')', r'\1 \2', text)
    text = re.sub('(' + _AI + ')(' + _AR + ')', r'\1 \2', text)

    # Brackets ()[]{}  — inner spaces first, then outer
    # Inner: space after opening bracket if not already spaced
    text = re.sub(r'\((\S)', r'( \1', text)
    text = re.sub(r'\[(\S)', r'[ \1', text)
    text = re.sub(r'\{(\S)', r'{ \1', text)
    # Inner: space before closing bracket if not already spaced
    text = re.sub(r'(\S)\)', r'\1 )', text)
    text = re.sub(r'(\S)\]', r'\1 ]', text)
    text = re.sub(r'(\S)\}', r'\1 }', text)
    # Outer: space before opening bracket
    text = re.sub(r'(\S)\(', r'\1 (', text)
    text = re.sub(r'(\S)\[', r'\1 [', text)
    text = re.sub(r'(\S)\{', r'\1 {', text)
    # Outer: space after closing bracket
    text = re.sub(r'\)(\S)', r') \1', text)
    text = re.sub(r'\](\S)', r'] \1', text)
    text = re.sub(r'\}(\S)', r'} \1', text)

    # Collapse any double spaces introduced above
    text = re.sub(r'  +', ' ', text)
    return text


# ═══════════════════════════════════════════════════════════════════
# COMPOSITE HELPERS
# ═══════════════════════════════════════════════════════════════════

def _header(slide, title, section_num, client_logo=None):
    """Dark navy header bar with logo box, section title and badge."""
    _rect(slide, 0, 0, W, 2.7, NAVY)
    _rect(slide, 0.25, 0.18, 5.3, 2.35, WHITE, rounded=True, radius=0.05)
    _img(slide, ZAWAD_LOGO, 2.8, 0.25, 2.5, 2.1)

    if client_logo and os.path.exists(client_logo):
        _img(slide, client_logo, 0.4, 0.3, 2.2, 2.0)
        _rect(slide, 2.72, 0.45, 0.05, 1.7, MGRAY)

    tx = _txbox(slide, 5.8, 0.2, 24.5, 1.55)
    tf = tx.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf)
    _para(tf, title, 32, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    _oval(slide, W - 2.85, 0.2, 2.3, 2.3, MINT)
    tx2 = _txbox(slide, W - 2.85, 0.2, 2.3, 2.3)
    tf2 = tx2.text_frame
    tf2.word_wrap = False
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf2)
    p = tf2.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = str(section_num).zfill(2)
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = WHITE
    run.font.name = 'Times New Roman'

    tx3 = _txbox(slide, 5.8, 1.5, 24.5, 1.15)
    tf3 = tx3.text_frame
    tf3.word_wrap = True
    tf3.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf3)
    return tf3


def _divider(slide):
    _rect(slide, 0, 2.7, W, 0.12, MINT)


def _icon_circle(slide, cx, cy, r, icon, icon_size=18, bg=NAVY):
    d = r * 2
    _oval(slide, cx - r, cy - r, d, d, bg)
    tx = _txbox(slide, cx - r, cy - r, d, d)
    tf = tx.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = icon
    run.font.size = Pt(icon_size)
    run.font.color.rgb = WHITE
    run.font.name = 'Segoe UI Emoji'


def _info_box(slide, x, y, w, h, icon, title, value,
              bg=WHITE, title_color=TEAL, value_color=SLATE):
    _rect(slide, x, y, w, h, bg, rounded=True, radius=0.08, border=MGRAY)
    ic_x = x + w - 0.9
    ic_y = y + h / 2
    _icon_circle(slide, ic_x, ic_y, 0.55, icon, 14, MINT)
    tx = _txbox(slide, x + 0.2, y + 0.15, w - 1.6, h - 0.3)
    tf = tx.text_frame
    tf.word_wrap = True
    _tf_clear(tf)
    _para(tf, title, 13, bold=True, color=title_color, align=PP_ALIGN.RIGHT)
    _para(tf, value, 14, bold=False, color=value_color, align=PP_ALIGN.RIGHT)


def _stat_box(slide, x, y, w, h, number, label):
    _rect(slide, x, y, w, h, NAVY, rounded=True, radius=0.1)
    tx = _txbox(slide, x, y, w, h)
    tf = tx.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf)
    _para(tf, _ar(number), 32, bold=True, color=MINT, align=PP_ALIGN.RIGHT, rtl=True)
    _para(tf, label, 14, bold=False, color=WHITE, align=PP_ALIGN.RIGHT, rtl=True, space_before=2)


def _orb(slide, x, y, d, color):
    _oval(slide, x, y, d, d, color)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 1: COVER
# ═══════════════════════════════════════════════════════════════════

def _slide_cover(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, NAVY)
    _orb(slide, -3.5, -3.5, 10, MINT)
    _orb(slide, W - 7, -4, 10, TEAL)
    _orb(slide, -2, H - 6, 8, TEAL)
    _orb(slide, W - 5, H - 6, 9, MINT)

    box_w, box_h = 12, 3.8
    box_x, box_y = (W - box_w) / 2, 1.5
    _rect(slide, box_x, box_y, box_w, box_h, WHITE, rounded=True, radius=0.08)
    _img(slide, ZAWAD_LOGO, box_x + box_w / 2 + 0.3, box_y + 0.25, 5.2, 3.2)

    if client_logo and os.path.exists(client_logo):
        _img(slide, client_logo, box_x + 0.4, box_y + 0.3, 5.0, 3.1)
        _rect(slide, box_x + box_w / 2 - 0.03, box_y + 0.5, 0.06, 2.7, MGRAY)

    tx = _txbox(slide, 2, 6.2, W - 4, 1.5)
    tf = tx.text_frame
    tf.word_wrap = False
    _tf_clear(tf)
    _para(tf, data.get('program_category', 'البرنامج التدريبي'), 18,
          color=MGREEN, align=PP_ALIGN.CENTER, rtl=True)

    tx2 = _txbox(slide, 1.5, 7.3, W - 3, 4.0)
    tf2 = tx2.text_frame
    tf2.word_wrap = True
    _tf_clear(tf2)
    _para(tf2, data.get('program_title', ''), 36, bold=True,
          color=WHITE, align=PP_ALIGN.CENTER, rtl=True)

    _rect(slide, 6, 11.6, W - 12, 0.12, MINT)
    _rect(slide, 3, 12.2, W - 6, 5.0, RGBColor(0x0D, 0x2D, 0x55),
          rounded=True, radius=0.1, border=MINT)

    info_lines = [
        ('مقدَّم إلى: ' + data.get('client_name', ''), 16, True, WHITE),
        (data.get('submitted_to', ''), 15, False, MGREEN),
        ('رقم المناقصة: ' + data.get('tender_number', '—') +
         '      أكتوبر 2026', 14, False, LGRAY),
    ]
    ty = 12.6
    for line, sz, bold, col in info_lines:
        tx3 = _txbox(slide, 4, ty, W - 8, 0.9)
        tf3 = tx3.text_frame
        tf3.word_wrap = False
        _tf_clear(tf3)
        _para(tf3, line, sz, bold=bold, color=col, align=PP_ALIGN.RIGHT, rtl=True)
        ty += 1.4


# ═══════════════════════════════════════════════════════════════════
# SLIDE 2: COVER LETTER
# ═══════════════════════════════════════════════════════════════════

def _slide_cover_letter(prs, data, client_logo=None):
    """Formal Arabic cover letter slide matching 98087 slide 3 style."""
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)

    client = data.get('client_name', 'الجهة المستفيدة')
    program = data.get('program_title', 'البرنامج التدريبي')
    days_val = str(data.get('stats', {}).get('days', '—'))
    participants_val = str(data.get('stats', {}).get('participants', '—'))
    location_val = data.get('stats', {}).get('location', 'مقر الجهة')

    tf_hdr = _header(slide, 'خطاب التقديم', '01', client_logo)
    _divider(slide)
    _para(tf_hdr, f'العرض الفني والمالي لمناقصة "{program}"', 14, color=MINT, align=PP_ALIGN.RIGHT)

    # ── WHITE LETTER SECTION ──────────────────────────────────────────
    letter_y, letter_h = 3.0, 8.5
    _rect(slide, 0.5, letter_y, W - 1.0, letter_h, WHITE)

    cx = 1.1           # inner left margin
    cw = W - 2.2       # inner width

    # Addressee line
    tx = _txbox(slide, cx, 3.35, cw, 0.9)
    tf = tx.text_frame
    tf.word_wrap = True
    _tf_clear(tf)
    _para(tf, f'الأفاضل/ {client} المحترمين', 16, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)

    # Salutation
    tx2 = _txbox(slide, cx, 4.35, cw, 0.7)
    tf2 = tx2.text_frame
    _tf_clear(tf2)
    _para(tf2, 'تحية طيبة وبعد،', 13, color=SLATE, align=PP_ALIGN.RIGHT)

    # Body paragraph 1: company intro + program scope
    body1 = (
        f'تتقدم شركة زواد للتطوير والاستشارات، الشركة العُمانية المتخصصة في التدريب والاستشارات '
        f'وتطوير القيادات ومقرها العذيبة – مسقط، بهذا العرض الفني والمالي لتنفيذ برنامج "{program}" '
        f'لعدد ({_ar(participants_val)}) مشاركاً من المستويات القيادية المستهدفة لدى {client}.'
    )
    tx3 = _txbox(slide, cx, 5.15, cw, 2.5)
    tf3 = tx3.text_frame
    tf3.word_wrap = True
    _tf_clear(tf3)
    _para(tf3, body1, 13, color=SLATE, align=PP_ALIGN.RIGHT)

    # Body paragraph 2: commitment to execution
    body2 = (
        f'ونلتزم بتنفيذ البرنامج حضورياً في {location_val} على مدى ({_ar(days_val)}) أيام تدريبية '
        f'باللغة العربية، وفق التصميم والمحتوى التعليمي المحدد، بما يحقق الأهداف التنموية '
        f'المرجوة من هذا البرنامج الذي صُمِّم خصيصاً لتلبية احتياجات {client}.'
    )
    tx4 = _txbox(slide, cx, 7.75, cw, 2.0)
    tf4 = tx4.text_frame
    tf4.word_wrap = True
    _tf_clear(tf4)
    _para(tf4, body2, 13, color=SLATE, align=PP_ALIGN.RIGHT)

    # Closing line
    tx5 = _txbox(slide, cx, 10.0, cw, 0.8)
    tf5 = tx5.text_frame
    _tf_clear(tf5)
    _para(tf5, 'وتفضلوا بقبول فائق الاحترام والتقدير،', 13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)

    # ── NAVY COMMITMENTS SECTION ──────────────────────────────────────
    navy_y = letter_y + letter_h  # 11.5
    navy_h = H - navy_y - 0.45   # ~7.1 cm
    _rect(slide, 0.5, navy_y, W - 1.0, navy_h, NAVY)

    # Section heading
    tx6 = _txbox(slide, cx, navy_y + 0.3, cw, 0.9)
    tf6 = tx6.text_frame
    _tf_clear(tf6)
    _para(tf6, 'يرتكز عرضنا على ثلاثة التزامات جوهرية', 15, bold=True, color=MINT, align=PP_ALIGN.RIGHT)

    # Three side-by-side commitment boxes
    commitments = [
        (
            'التطبيق قبل التلقين',
            'تجربة تعلم تنفيذية يغلب عليها التطبيق الميداني، يخرج منها كل مشارك '
            'بخطة عمل تطبيقية جاهزة للتنفيذ الفوري في وحدته الحكومية.',
        ),
        (
            'محاكاة من واقع البيئة العُمانية',
            'حالات وسيناريوهات حقيقية من بيئة العمل الحكومي العُماني تضمن الانتقال '
            'الفوري من التعلم النظري إلى التطبيق الميداني المباشر.',
        ),
        (
            'قياس موثّق للتعلم',
            f'تقييم قبلي وبعدي، وبطاقة ملاحظة أداء، وتقرير فردي لكل مشارك '
            f'من ({_ar(participants_val)}) مشارك، ومتابعة للأثر بعد ٦٠–٩٠ يوماً.',
        ),
    ]

    gap = 0.4
    box_w = (cw - gap * 2) / 3
    box_h = navy_h - 1.35
    box_y = navy_y + 1.2

    for i, (title_c, desc_c) in enumerate(commitments):
        bx = cx + i * (box_w + gap)
        _rect(slide, bx, box_y, box_w, box_h, WHITE, rounded=True, radius=0.15)
        # Mint accent strip at top of each box
        _rect(slide, bx, box_y, box_w, 0.35, MINT)
        # Title
        tx_t = _txbox(slide, bx + 0.2, box_y + 0.45, box_w - 0.4, 1.1)
        tf_t = tx_t.text_frame
        tf_t.word_wrap = True
        _tf_clear(tf_t)
        _para(tf_t, title_c, 13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        # Description
        tx_d = _txbox(slide, bx + 0.2, box_y + 1.65, box_w - 0.4, box_h - 2.1)
        tf_d = tx_d.text_frame
        tf_d.word_wrap = True
        _tf_clear(tf_d)
        _para(tf_d, desc_c, 11, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 3: TABLE OF CONTENTS
# ═══════════════════════════════════════════════════════════════════

def _slide_toc(prs, data):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    _rect(slide, 0, 0, W, 2.7, NAVY)
    _rect(slide, 0, 0, 5.5, 2.7, WHITE)
    _img(slide, ZAWAD_LOGO, 0.15, 0.1, 5.2, 2.5)
    _rect(slide, 0, 2.7, W, 0.12, MINT)

    tx = _txbox(slide, 5.8, 0.2, W - 6.3, 1.55)
    tf = tx.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf)
    _para(tf, 'المحتوى', 30, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    tx2 = _txbox(slide, 5.8, 1.6, W - 6.3, 1.0)
    tf2 = tx2.text_frame
    tf2.word_wrap = False
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf2)
    _para(tf2, 'فهرس أقسام العرض الفني والمالي', 16,
          color=MINT, align=PP_ALIGN.RIGHT)

    sections = [
        ('01', 'خطاب التقديم والملخص التنفيذي'),
        ('02', 'نبذة عن زواد للتطوير والاستشارات'),
        ('03', 'فهمنا لنطاق العمل وأهداف البرنامج'),
        ('04', 'المنهجية المقترحة والتصميم التفصيلي'),
        ('05', 'خطة العمل: المراحل والأنشطة والمخرجات'),
        ('06', 'الإطار الزمني التفصيلي'),
        ('07', 'فريق العمل والسير الذاتية'),
        ('08', 'أدوات وأساليب قياس التعلم'),
        ('09', 'الحوكمة وضمان الجودة'),
        ('10', 'الابتكار والقيمة المضافة'),
        ('11', 'خبرة الشركة ومشاريعنا المكتملة'),
        ('12', 'العرض المالي'),
        ('13', 'الإقرارات والالتزامات والمرفقات'),
    ]

    cols = [sections[:7], sections[7:]]
    col_w = W / 2 - 1.0

    table_y = 3.7
    avail_h = H - table_y - 0.25
    max_rows = 7
    step = avail_h / max_rows
    bh = step - 0.2

    for ci, col in enumerate(cols):
        cx = W - 0.5 - ci * (col_w + 0.3)
        for ri, (num, label) in enumerate(col):
            bx = cx - col_w
            by = table_y + ri * step
            _rect(slide, bx, by, col_w, bh, WHITE, rounded=True, radius=0.12, border=MGRAY)
            _rect(slide, bx, by, 1.6, bh, MINT, rounded=True, radius=0.12)
            tx3 = _txbox(slide, bx, by, 1.6, bh)
            tf3 = tx3.text_frame
            tf3.word_wrap = False
            tf3.vertical_anchor = MSO_ANCHOR.MIDDLE
            _tf_clear(tf3)
            p = tf3.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = num
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = WHITE
            run.font.name = 'Times New Roman'
            tx4 = _txbox(slide, bx + 1.7, by + 0.15, col_w - 2.0, bh - 0.3)
            tf4 = tx4.text_frame
            tf4.word_wrap = True
            tf4.vertical_anchor = MSO_ANCHOR.MIDDLE
            _tf_clear(tf4)
            _para(tf4, label, 13, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 3: EXECUTIVE SUMMARY — STATS
# ═══════════════════════════════════════════════════════════════════

def _slide_stats(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'الملخص التنفيذي', '01', client_logo)
    _divider(slide)
    _para(tf3, 'البرنامج في أرقام وأبرز التزاماتنا', 15, color=MINT, align=PP_ALIGN.RIGHT)

    stats = data.get('key_stats', [
        {'number': str(data.get('stats', {}).get('participants', '—')), 'label': 'مشاركاً'},
        {'number': str(data.get('stats', {}).get('days', '—')), 'label': 'أيام تدريبية حضورية'},
        {'number': '3', 'label': 'وحدات تعليمية'},
        {'number': '2', 'label': 'جلستان حواريتان'},
        {'number': '9', 'label': 'أسابيع حتى التسليم'},
        {'number': '90', 'label': 'يوماً صلاحية العرض'},
    ])

    # Stat boxes — taller so numbers are prominent
    stats_y, stats_h = 3.0, 4.5
    n = len(stats)
    bw = (W - 1.0) / n
    for i, s in enumerate(stats):
        bx = 0.5 + i * bw
        _stat_box(slide, bx, stats_y, bw - 0.2, stats_h,
                  s.get('number', '—'), s.get('label', ''))

    # Two-column layout below stat tiles (matches 98087 slide 4)
    cols_y = stats_y + stats_h + 0.2
    cols_h = H - cols_y - 0.3
    col_w = (W - 1.0) / 2   # each column half the slide width

    participants = str(data.get('stats', {}).get('participants', '—'))
    days = str(data.get('stats', {}).get('days', '—'))
    location = data.get('stats', {}).get('location', 'مقر الجهة')
    client = data.get('client_name', 'الجهة المستفيدة')

    # ── RIGHT COLUMN (NAVY): ملخص الالتزامات الرئيسية ───────────────
    rc_x = 0.5 + col_w   # right column left edge
    _rect(slide, rc_x, cols_y, col_w, cols_h, NAVY)

    # Right column heading
    tx = _txbox(slide, rc_x + 0.2, cols_y + 0.2, col_w - 0.4, 0.9)
    tf = tx.text_frame; _tf_clear(tf)
    _para(tf, 'ملخص الالتزامات الرئيسية', 14, bold=True, color=MINT, align=PP_ALIGN.RIGHT)

    right_rows = data.get('summary_right', [
        ('الفئة المستهدفة', f'{_ar(participants)} مشاركاً من المستويات القيادية'),
        ('لغة ومكان التنفيذ', f'العربية – {location}'),
        ('مدة البرنامج', f'{_ar(days)} أيام تدريبية حضورية'),
        ('صلاحية العرض', f'{_ar(90)} يوماً من تاريخ الإصدار'),
    ])
    r_row_h = (cols_h - 1.2) / len(right_rows)
    for ri, (lbl, val) in enumerate(right_rows):
        ry = cols_y + 1.2 + ri * r_row_h
        bg = RGBColor(0x0E, 0x2F, 0x5F) if ri % 2 == 0 else TEAL  # darker navy / teal
        _rect(slide, rc_x, ry, col_w, r_row_h, bg, border=RGBColor(0x1A, 0x4A, 0x80))
        # label
        tx = _txbox(slide, rc_x + 0.2, ry, col_w - 0.4, r_row_h * 0.45)
        tf = tx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
        _para(tf, lbl, 11, bold=True, color=MINT, align=PP_ALIGN.RIGHT)
        # value
        tx2 = _txbox(slide, rc_x + 0.2, ry + r_row_h * 0.45, col_w - 0.4, r_row_h * 0.55)
        tf2 = tx2.text_frame; tf2.word_wrap = True; tf2.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf2)
        _para(tf2, val, 12, color=WHITE, align=PP_ALIGN.RIGHT)

    # ── LEFT COLUMN (WHITE): رسالة العرض + data rows ─────────────────
    lc_x = 0.5
    _rect(slide, lc_x, cols_y, col_w, cols_h, WHITE)
    _rect(slide, lc_x, cols_y, col_w, 0.08, MINT)  # mint top accent

    # "رسالة العرض:" heading
    tx = _txbox(slide, lc_x + 0.2, cols_y + 0.2, col_w - 0.4, 0.75)
    tf = tx.text_frame; _tf_clear(tf)
    _para(tf, 'رسالة العرض:', 14, bold=True, color=TEAL, align=PP_ALIGN.RIGHT)

    # Key message — large prominent text (the "big para" the user wants)
    msg = data.get('sections', {}).get('key_message',
        'تمكين القيادات التنفيذية بأحدث ممارسات القيادة الاستراتيجية لتحقيق مستهدفات '
        'رؤية عُمان 2040 بكفاءة واقتدار، وترجمة الأهداف المؤسسية إلى نتائج ملموسة وقابلة للقياس.')
    msg_h = 4.2
    tx2 = _txbox(slide, lc_x + 0.2, cols_y + 1.1, col_w - 0.4, msg_h)
    tf2 = tx2.text_frame; tf2.word_wrap = True; tf2.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf2)
    _para(tf2, msg, 18, bold=False, color=NAVY, align=PP_ALIGN.RIGHT)

    # Mint divider inside left column
    mid_y = cols_y + 1.1 + msg_h + 0.15
    _rect(slide, lc_x + 0.2, mid_y, col_w - 0.4, 0.08, MINT)

    left_rows = data.get('summary_left', [
        ('مكونات البرنامج', f'وحدات تعليمية + جلسات حوارية + تقييمات'),
        ('فترة الإعداد', f'{_ar(3)} أسابيع من تاريخ الإسناد'),
        ('تسليم المخرجات', f'خلال {_ar(15)} يوم عمل من انتهاء البرنامج'),
        ('عملة العرض', 'الريال العُماني، شاملاً الضرائب والرسوم'),
    ])
    l_row_h = (H - mid_y - 0.25) / len(left_rows)
    for ri, (lbl, val) in enumerate(left_rows):
        ry = mid_y + 0.1 + ri * l_row_h
        bg = LGRAY if ri % 2 == 0 else WHITE
        _rect(slide, lc_x, ry, col_w, l_row_h, bg, border=MGRAY)
        tx = _txbox(slide, lc_x + 0.3, ry, col_w - 0.5, l_row_h * 0.45)
        tf = tx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
        _para(tf, lbl, 11, bold=True, color=TEAL, align=PP_ALIGN.RIGHT)
        tx2 = _txbox(slide, lc_x + 0.3, ry + l_row_h * 0.45, col_w - 0.5, l_row_h * 0.55)
        tf2 = tx2.text_frame; tf2.word_wrap = True; tf2.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf2)
        _para(tf2, val, 12, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 4: COMMITMENTS GRID
# ═══════════════════════════════════════════════════════════════════

def _slide_commitments_grid(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'الملخص التنفيذي', '01', client_logo)
    _divider(slide)
    _para(tf3, 'ملخص الالتزامات الرئيسية', 15, color=MINT, align=PP_ALIGN.RIGHT)

    grid = data.get('commitments_grid', [
        {'icon': '\U0001f465', 'title': 'الفئة المستهدفة',
         'value': str(data.get('stats', {}).get('participants', '30')) + ' مشاركاً'},
        {'icon': '\U0001f4cd', 'title': 'لغة ومكان التنفيذ',
         'value': 'العربية — ' + data.get('stats', {}).get('location', 'مقر الجهة')},
        {'icon': '⌛', 'title': 'فترة الإعداد', 'value': '3 أسابيع من تاريخ الإسناد'},
        {'icon': '\U0001f4cb', 'title': 'تسليم المخرجات النهائية', 'value': 'خلال 15 يوم عمل من انتهاء البرنامج'},
        {'icon': '\U0001f4b0', 'title': 'عملة العرض', 'value': 'الريال العُماني، شاملاً جميع الضرائب'},
        {'icon': '\U0001f6e1', 'title': 'صلاحية العرض', 'value': '90 يوماً من تاريخ تقديم العرض'},
    ])

    bw = (W - 1.2) / 2
    bh = 3.0
    gap_x = 0.4
    gap_y = 0.35

    for i, item in enumerate(grid[:6]):
        row = i // 2
        col = 1 - (i % 2)
        bx = 0.5 + col * (bw + gap_x)
        by = 3.0 + row * (bh + gap_y)
        _info_box(slide, bx, by, bw, bh,
                  item.get('icon', '•'),
                  item.get('title', ''),
                  item.get('value', ''))


# ═══════════════════════════════════════════════════════════════════
# SLIDE 5: ABOUT ZAWAD — WITH RIYADAH CARD
# ═══════════════════════════════════════════════════════════════════

def _slide_company(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'عن الشركة', '02', client_logo)
    _divider(slide)
    _para(tf3, 'الشريك القائد  ·  سلطنة عُمان', 15, color=MINT, align=PP_ALIGN.RIGHT)

    _img(slide, RIYADAH_CARD, 0.5, 3.0, 5.5, 6.3)

    overview = data.get('sections', {}).get('company_overview',
        'شركة زواد للتطوير والاستشارات هي شركة رائدة في مجال تقديم حلول التدريب والاستشارات المبتكرة في سلطنة عُمان.')

    tx = _txbox(slide, 6.5, 3.0, W - 7.0, 7.2)
    tf = tx.text_frame
    tf.word_wrap = True
    _tf_clear(tf)
    paras = overview.split('\n') if '\n' in overview else [overview]
    for para in paras:
        if para.strip():
            _para(tf, para.strip(), 14, color=SLATE, align=PP_ALIGN.RIGHT)

    _rect(slide, 0.5, 9.6, W - 1.0, 1.5, MINT, rounded=True, radius=0.1)
    tx2 = _txbox(slide, 0.5, 9.6, W - 1.0, 1.5)
    tf2 = tx2.text_frame
    tf2.word_wrap = False
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf2)
    _para(tf2, '✔  حاصلة على بطاقة تسجيل ريادة — هيئة تنمية المؤسسات الصغيرة والمتوسطة  ✔',
          15, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    pillars = [
        ('\U0001f4a1', 'الابتكار', 'حلول تدريبية مبتكرة ومتطورة'),
        ('★', 'التميز', 'أعلى معايير الجودة في التنفيذ'),
        ('\U0001f3af', 'التخصص', 'حلول تدريب متخصصة لكل قطاع'),
    ]
    pw = (W - 1.5) / 3
    for i, (icon, name, desc) in enumerate(pillars):
        px = 0.5 + (2 - i) * (pw + 0.25)
        _rect(slide, px, 11.4, pw, 4.0, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _icon_circle(slide, px + pw / 2, 12.6, 0.65, icon, 18, MINT)
        tx3 = _txbox(slide, px + 0.2, 13.6, pw - 0.4, 1.0)
        tf3 = tx3.text_frame
        tf3.word_wrap = True
        _tf_clear(tf3)
        _para(tf3, name, 15, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        tx4 = _txbox(slide, px + 0.2, 14.7, pw - 0.4, 0.6)
        tf4 = tx4.text_frame
        tf4.word_wrap = True
        _tf_clear(tf4)
        _para(tf4, desc, 12, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 6: PARTNER LOGOS
# ═══════════════════════════════════════════════════════════════════

def _slide_clients(prs, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'عملاؤنا', '02', client_logo)
    _divider(slide)
    _para(tf3, 'نفخر بثقة مجموعة واسعة من الجهات الحكومية والخاصة', 15,
          color=MINT, align=PP_ALIGN.RIGHT)

    _rect(slide, 1.0, 3.0, W - 2.0, 14.8, WHITE, rounded=True, radius=0.1, border=MGRAY)
    _img(slide, PARTNER_LOGOS, 1.5, 3.5, W - 3.0, 14.0)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 7: UNDERSTANDING REQUIREMENTS
# ═══════════════════════════════════════════════════════════════════

def _slide_requirements(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'فهمنا لنطاق العمل والسياق', '03', client_logo)
    _divider(slide)
    _para(tf3, 'تحليلنا للمتطلبات والفئة المستهدفة', 15, color=MINT, align=PP_ALIGN.RIGHT)

    req = data.get('sections', {}).get('requirements_understanding', '')
    parts = req.split('\n\n') if '\n\n' in req else [req]

    _rect(slide, 0.5, 3.0, W * 0.58 - 0.7, 15.7, WHITE, rounded=True, radius=0.08, border=MGRAY)
    tx = _txbox(slide, 0.8, 3.3, W * 0.58 - 1.3, 15.1)
    tf = tx.text_frame
    tf.word_wrap = True
    _tf_clear(tf)
    _para(tf, 'السياق والمتطلبات', 16, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    for part in parts:
        if part.strip():
            _para(tf, part.strip(), 14, color=SLATE, align=PP_ALIGN.RIGHT, space_before=6)

    ctx_boxes = data.get('context_boxes', [
        {'icon': '⏱', 'title': 'مدخلات قصيرة', 'desc': 'مدخلات مفاهيمية لا تتجاوز 20 دقيقة'},
        {'icon': '\U0001f465', 'title': 'حالات من الواقع العُماني', 'desc': 'حالات من العمل الحكومي يطرحها المشاركون'},
        {'icon': '\U0001f5e3', 'title': 'بيئة نقاش آمنة', 'desc': 'وفق قاعدة هاوس: الأفكار لا تُنسب لأصحابها'},
        {'icon': '\U0001f4cb', 'title': 'مخرجات للاستخدام الفوري', 'desc': 'خطة تواصل مجتمعي لمدة 90 يوماً'},
    ])

    bx = W * 0.58 + 0.3
    bw = W - bx - 0.5
    bh = 3.5
    for i, box in enumerate(ctx_boxes[:4]):
        by = 3.0 + i * (bh + 0.3)
        _rect(slide, bx, by, bw, bh, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _icon_circle(slide, bx + bw - 0.9, by + bh / 2, 0.6, box.get('icon', '•'), 16, TEAL)
        tx2 = _txbox(slide, bx + 0.3, by + 0.3, bw - 1.8, bh - 0.6)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        _tf_clear(tf2)
        _para(tf2, box.get('title', ''), 14, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        _para(tf2, box.get('desc', ''), 13, color=SLATE, align=PP_ALIGN.RIGHT, space_before=4)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 8: EXPECTED OUTPUTS GRID
# ═══════════════════════════════════════════════════════════════════

def _slide_outputs(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'المخرجات المتوقعة', '03', client_logo)
    _divider(slide)
    _para(tf3, 'ما سيتمكن المشاركون من تحقيقه', 15, color=MINT, align=PP_ALIGN.RIGHT)

    outputs = data.get('sections', {}).get('expected_outputs', [
        {'number': '01', 'title': 'قيادات حكومية مؤهلة', 'description': 'وجاهزة لقيادة التواصل المجتمعي بكفاءة وفاعلية'},
        {'number': '02', 'title': 'قدرات في إدارة السمعة', 'description': 'المؤسسية والتعامل مع القضايا العامة والأزمات'},
        {'number': '03', 'title': 'تطبيق المعارف والمهارات', 'description': 'عبر المحاكاة التفاعلية والمواقف العملية'},
        {'number': '04', 'title': 'تغذية راجعة فردية', 'description': 'لكل مشارك تحدد نقاط القوة وفرص التحسين'},
        {'number': '05', 'title': 'قائمة المشاركين المستوفين', 'description': 'لمتطلبات الإتمام وفق معايير الأكاديمية'},
        {'number': '06', 'title': 'تقارير دورية عن سير التنفيذ', 'description': 'ونتائج قياس التعلم والمشاركين'},
        {'number': '07', 'title': 'تقرير ختامي بالنتائج', 'description': 'والتحديات والدروس المستفادة والتوصيات'},
        {'number': '08', 'title': 'خطة تواصل مجتمعي', 'description': 'لمدة 90 يوماً لكل مشارك في 30 وحدة حكومية'},
    ])

    cols = 4
    rows = (len(outputs) + cols - 1) // cols
    bw = (W - 1.0) / cols - 0.3
    bh = (H - 3.8) / rows - 0.3

    for i, out in enumerate(outputs):
        row = i // cols
        col = cols - 1 - (i % cols)
        bx = 0.5 + col * (bw + 0.3)
        by = 3.2 + row * (bh + 0.3)

        is_last = (i == len(outputs) - 1)
        bg = MINT if is_last else WHITE
        tc = WHITE if is_last else NAVY
        dc = WHITE if is_last else SLATE

        _rect(slide, bx, by, bw, bh, bg, rounded=True, radius=0.12,
              border=None if is_last else MGRAY)

        _oval(slide, bx + bw - 1.3, by + 0.2, 1.0, 0.9,
              WHITE if is_last else NAVY)
        tx_num = _txbox(slide, bx + bw - 1.3, by + 0.2, 1.0, 0.9)
        tf_num = tx_num.text_frame
        tf_num.word_wrap = False
        tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf_num)
        p = tf_num.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = out.get('number', str(i + 1))
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = MINT if is_last else WHITE
        r.font.name = 'Times New Roman'

        tx2 = _txbox(slide, bx + 0.2, by + 0.2, bw - 1.7, bh - 0.4)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        _tf_clear(tf2)
        _para(tf2, out.get('title', ''), 14, bold=True, color=tc, align=PP_ALIGN.RIGHT)
        _para(tf2, out.get('description', ''), 13, color=dc,
              align=PP_ALIGN.RIGHT, space_before=3)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 9: METHODOLOGY
# ═══════════════════════════════════════════════════════════════════

def _slide_methodology(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'المنهجية المقترحة', '04', client_logo)
    _divider(slide)
    _para(tf3, 'المرجعية الدولية والتصميم التفصيلي', 15, color=MINT, align=PP_ALIGN.RIGHT)

    methodology = data.get('sections', {}).get('methodology', '')
    parts = methodology.split('\n\n') if '\n\n' in methodology else [methodology]

    _rect(slide, 0.5, 3.0, W - 1.0, 4.5, WHITE, rounded=True, radius=0.08, border=MGRAY)
    tx = _txbox(slide, 0.9, 3.3, W - 1.8, 4.0)
    tf = tx.text_frame
    tf.word_wrap = True
    _tf_clear(tf)
    for part in parts[:2]:
        if part.strip():
            _para(tf, part.strip(), 14, color=SLATE, align=PP_ALIGN.RIGHT)

    meth_items = data.get('methodology_items', [
        {'icon': '\U0001f4da', 'title': 'تصميم التعلم', 'desc': 'دورة كولب + مبادئ تعلم البالغين'},
        {'icon': '\U0001f517', 'title': 'نقل التعلم إلى العمل', 'desc': 'نموذج 70-20-10 للتطوير القيادي'},
        {'icon': '\U0001f3af', 'title': 'منظور القيادة', 'desc': 'القيادة التكيفية — كلية كينيدي هارفارد'},
        {'icon': '\U0001f4ca', 'title': 'تقييم البرنامج', 'desc': 'نموذج كيركباتريك وأسلوب الحالات الناجحة'},
    ])

    bw = (W - 1.5) / 4 - 0.2
    for i, item in enumerate(meth_items):
        bx = 0.5 + (3 - i) * (bw + 0.27)
        _rect(slide, bx, 7.8, bw, 4.5, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _icon_circle(slide, bx + bw / 2, 9.2, 0.7, item.get('icon', '•'), 20, TEAL)
        tx2 = _txbox(slide, bx + 0.2, 10.2, bw - 0.4, 1.1)
        tf2 = tx2.text_frame
        _tf_clear(tf2)
        _para(tf2, item.get('title', ''), 14, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        tx3 = _txbox(slide, bx + 0.2, 11.2, bw - 0.4, 1.0)
        tf3 = tx3.text_frame
        tf3.word_wrap = True
        _tf_clear(tf3)
        _para(tf3, item.get('desc', ''), 12, color=SLATE, align=PP_ALIGN.RIGHT)

    _rect(slide, 0.5, 12.6, W - 1.0, 6.2, NAVY, rounded=True, radius=0.1)
    tx4 = _txbox(slide, 1.0, 12.9, W - 2.0, 5.6)
    tf4 = tx4.text_frame
    tf4.word_wrap = True
    _tf_clear(tf4)
    _para(tf4, 'التزامنا بالجودة', 16, bold=True, color=MINT, align=PP_ALIGN.RIGHT)
    qual_text = data.get('sections', {}).get('quality_commitment',
        'نلتزم بتصميم برنامج تدريبي يستجيب للسياق العُماني الحكومي ويُترجم المعارف إلى ممارسات قابلة للتطبيق الفوري.')
    _para(tf4, qual_text, 14, color=LGRAY, align=PP_ALIGN.RIGHT, space_before=4)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 10+: WORK PLAN DAYS
# ═══════════════════════════════════════════════════════════════════

def _slide_workplan_day(prs, day_data, day_num, total_days, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'خطة العمل والأنشطة', '05', client_logo)
    _divider(slide)
    _para(tf3, f'اليوم {_ar(day_num)} من {_ar(total_days)}', 15, color=MINT, align=PP_ALIGN.RIGHT)

    _rect(slide, 0.5, 3.0, W - 1.0, 2.2, NAVY, rounded=True, radius=0.1)
    tx = _txbox(slide, 0.8, 3.1, W - 1.6, 2.0)
    tf = tx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_clear(tf)
    _para(tf, day_data.get('title', f'اليوم {_ar(day_num)}'), 18, bold=True,
          color=WHITE, align=PP_ALIGN.RIGHT)

    activities = day_data.get('activities', [])
    n = len(activities)
    if n == 0:
        return

    available_h = H - 6.0
    bh = min(4.0, (available_h - 0.2 * (n - 1)) / n)
    bw = W - 1.0

    for i, act in enumerate(activities):
        by = 5.5 + i * (bh + 0.25)
        _rect(slide, 0.5, by, bw, bh, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _rect(slide, 0.5, by, 0.35, bh, MINT, rounded=False)

        _oval(slide, bw - 0.7, by + bh / 2 - 0.45, 0.9, 0.9, NAVY)
        tx_num = _txbox(slide, bw - 0.7, by + bh / 2 - 0.45, 0.9, 0.9)
        tf_num = tx_num.text_frame
        tf_num.word_wrap = False
        tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf_num)
        p = tf_num.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(i + 1)
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = WHITE
        r.font.name = 'Times New Roman'

        tx2 = _txbox(slide, 1.1, by + 0.2, bw - 2.2, bh - 0.4)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf2)
        fs = 15 if len(act) < 120 else 13
        _para(tf2, act, fs, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE: TEAM
# ═══════════════════════════════════════════════════════════════════

def _slide_team_page(prs, members_page, page_num, total_pages, slide_num, client_logo, lang='ar'):
    """Render one team slide for up to 3 members."""
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)

    subtitle_suffix = ''
    if total_pages > 1:
        subtitle_suffix = '  —  ' + _lbl(lang, 'team_page', n=page_num)

    tf3 = _header(slide, _lbl(lang, 'team_title'), str(slide_num).zfill(2), client_logo)
    _divider(slide)
    _para(tf3, _lbl(lang, 'team_sub') + subtitle_suffix, 15,
          color=MINT, align=_align_for(lang), rtl=_is_rtl(lang))

    n = len(members_page)
    cols = min(n, 3)
    bw = (W - 1.0) / cols - 0.3
    bh = H - 4.0

    for i, member in enumerate(members_page):
        col = (cols - 1 - i) if lang != 'en' else i
        bx = 0.5 + col * (bw + 0.3)

        _rect(slide, bx, 3.0, bw, bh, WHITE, rounded=True, radius=0.12, border=MGRAY)
        _rect(slide, bx, 3.0, bw, 1.5, NAVY, rounded=True, radius=0.12)
        _rect(slide, bx, 3.8, bw, 0.7, NAVY)

        # Avatar circle with initial
        _oval(slide, bx + bw / 2 - 1.2, 3.2, 2.4, 2.4, MINT)
        tx_init = _txbox(slide, bx + bw / 2 - 1.2, 3.2, 2.4, 2.4)
        tf_init = tx_init.text_frame
        tf_init.word_wrap = False
        tf_init.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf_init)
        p = tf_init.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        name_key = 'name_en' if lang == 'en' else 'name'
        name = member.get(name_key) or member.get('name', '')
        r.text = name[0] if name else '★'
        r.font.size = Pt(28)
        r.font.bold = True
        r.font.color.rgb = WHITE
        r.font.name = 'Times New Roman'

        # Name
        tx2 = _txbox(slide, bx + 0.3, 6.0, bw - 0.6, 1.4)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf2)
        _para(tf2, name, 15, bold=True, color=NAVY,
              align=_align_for(lang), rtl=_is_rtl(lang))

        # Years badge (small, on avatar top-right)
        years = str(member.get('years', '')).strip()
        if years:
            _rect(slide, bx + bw - 1.5, 3.05, 1.2, 0.65, TEAL, rounded=True, radius=0.08)
            tx_yr = _txbox(slide, bx + bw - 1.5, 3.05, 1.2, 0.65)
            tf_yr = tx_yr.text_frame
            tf_yr.word_wrap = False
            tf_yr.vertical_anchor = MSO_ANCHOR.MIDDLE
            _tf_clear(tf_yr)
            p_yr = tf_yr.paragraphs[0]
            p_yr.alignment = PP_ALIGN.CENTER
            r_yr = p_yr.add_run()
            r_yr.text = years
            r_yr.font.size = Pt(9)
            r_yr.font.bold = True
            r_yr.font.color.rgb = WHITE
            r_yr.font.name = 'Times New Roman'

        # Role badge
        role_key = 'role_en' if lang == 'en' else 'role'
        role = member.get(role_key) or member.get('role', '')
        _rect(slide, bx + 0.5, 7.4, bw - 1.0, 1.1, MINT, rounded=True, radius=0.08)
        tx3 = _txbox(slide, bx + 0.5, 7.4, bw - 1.0, 1.1)
        tf3_role = tx3.text_frame
        tf3_role.word_wrap = True
        tf3_role.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf3_role)
        _para(tf3_role, role, 12, bold=True, color=WHITE,
              align=_align_for(lang), rtl=_is_rtl(lang))

        # Experience text
        exp_key = 'experience_en' if lang == 'en' else 'experience'
        exp = member.get(exp_key) or member.get('experience', '')
        tx4 = _txbox(slide, bx + 0.3, 8.7, bw - 0.6, bh - 7.1)
        tf4 = tx4.text_frame
        tf4.word_wrap = True
        _tf_clear(tf4)
        _para(tf4, exp, 11, color=SLATE,
              align=_align_for(lang), rtl=_is_rtl(lang))

        # Qualifications bullets (thin MINT divider + bullet list)
        quals_key = 'qualifications_en' if lang == 'en' else 'qualifications'
        quals = member.get(quals_key) or member.get('qualifications', [])
        if quals and isinstance(quals, list):
            divider_y = 3.0 + bh - 2.6
            _rect(slide, bx + 0.4, divider_y, bw - 0.8, 0.06, MINT)
            qt = _txbox(slide, bx + 0.3, divider_y + 0.12, bw - 0.6, 2.4)
            qtf = qt.text_frame
            qtf.word_wrap = True
            _tf_clear(qtf)
            for q in quals[:4]:
                bullet = ('• ' if lang == 'en' else '◈ ') + str(q)
                _para(qtf, bullet, 10, color=TEAL,
                      align=_align_for(lang), rtl=_is_rtl(lang))


def _slide_team_pages(prs, data, client_logo, lang='ar'):
    """Render team across multiple slides (3 members per slide)."""
    team = data.get('sections', {}).get('team', [])
    if not team:
        return
    page_size = 3
    pages = [team[i:i + page_size] for i in range(0, len(team), page_size)]
    total_pages = len(pages)
    for page_num, members_page in enumerate(pages, 1):
        _slide_team_page(prs, members_page, page_num, total_pages,
                         slide_num=7 + (page_num - 1),
                         client_logo=client_logo, lang=lang)


def _slide_team(prs, data, client_logo):
    """Legacy alias — calls multi-page version with Arabic."""
    _slide_team_pages(prs, data, client_logo, lang='ar')


# ═══════════════════════════════════════════════════════════════════
# SLIDE: TRACK RECORD — PAST PROJECTS TABLE
# ═══════════════════════════════════════════════════════════════════

def _slide_track_record(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'المشاريع المكتملة — زواد', '11', client_logo)
    _divider(slide)
    _para(tf3, 'سجل إنجازاتنا الموثق', 15, color=MINT, align=PP_ALIGN.RIGHT)

    headers = ['المستفيدون', 'نطاق العمل', 'الجهة المستفيدة', 'اسم المشروع', 'م']
    col_ws_raw = [3.0, 6.5, 7.0, 11.5, 1.5]
    total_w = sum(col_ws_raw)
    scale = (W - 1.0) / total_w
    col_ws = [c * scale for c in col_ws_raw]

    header_h = 1.2
    row_h = 1.4
    table_x = 0.5
    table_y = 3.1

    cx = table_x
    for hdr, cw in zip(headers, col_ws):
        _rect(slide, cx, table_y, cw, header_h, NAVY)
        tx = _txbox(slide, cx + 0.05, table_y, cw - 0.1, header_h)
        tf = tx.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf)
        _para(tf, hdr, 10, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, rtl=True)
        cx += cw

    projects = PAST_PROJECTS[:10]
    for ri, row_data in enumerate(projects):
        row_color = WHITE if ri % 2 == 0 else LGRAY
        row_y = table_y + header_h + ri * row_h
        cx = table_x
        cells = [row_data[3], row_data[5], row_data[1], row_data[0], _ar(ri + 1)]
        for j, (cell, cw) in enumerate(zip(cells, col_ws)):
            is_name = (j == 3)
            _rect(slide, cx, row_y, cw, row_h, row_color, border=MGRAY)
            tx = _txbox(slide, cx + 0.05, row_y + 0.1, cw - 0.1, row_h - 0.2)
            tf = tx.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            _tf_clear(tf)
            col_color = TEAL if is_name else SLATE
            _para(tf, str(cell), 10, bold=is_name, color=col_color, align=PP_ALIGN.RIGHT, rtl=True)
            cx += cw


# ═══════════════════════════════════════════════════════════════════
# SLIDE: INNOVATION / ADDED VALUE
# ═══════════════════════════════════════════════════════════════════

def _slide_innovation(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'الابتكار والقيمة المضافة', '10', client_logo)
    _divider(slide)
    _para(tf3, 'ما الذي يميزنا عن غيرنا', 15, color=MINT, align=PP_ALIGN.RIGHT)

    inno_text = data.get('sections', {}).get('innovation', '')
    points = data.get('sections', {}).get('innovation_points', [])

    if not points:
        pts = inno_text.split('\n') if inno_text else []
        points = [{'title': p.strip(), 'description': ''} for p in pts if p.strip()][:6]

    default_icons = ['\U0001f310', '\U0001f4a1', '\U0001f3c6', '\U0001f91d', '\U0001f4ca', '⭐']

    cols = 3
    rows = (len(points) + cols - 1) // cols
    bw = (W - 1.2) / cols - 0.25
    bh = (H - 4.0) / rows - 0.3

    for i, pt in enumerate(points[:6]):
        row = i // cols
        col = cols - 1 - (i % cols)
        bx = 0.5 + col * (bw + 0.25)
        by = 3.2 + row * (bh + 0.3)

        _rect(slide, bx, by, bw, bh, WHITE, rounded=True, radius=0.12, border=MGRAY)

        icon = pt.get('icon', default_icons[i % len(default_icons)])
        _icon_circle(slide, bx + bw / 2, by + 1.2, 0.7, icon, 20, TEAL)

        tx = _txbox(slide, bx + 0.3, by + 2.3, bw - 0.6, 1.2)
        tf = tx.text_frame
        tf.word_wrap = True
        _tf_clear(tf)
        _para(tf, pt.get('title', ''), 14, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)

        desc = pt.get('description', '')
        if desc:
            tx2 = _txbox(slide, bx + 0.3, by + 3.5, bw - 0.6, bh - 3.7)
            tf2 = tx2.text_frame
            tf2.word_wrap = True
            _tf_clear(tf2)
            _para(tf2, desc, 13, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE: FINANCIAL SUMMARY
# ═══════════════════════════════════════════════════════════════════

def _slide_financial_summary(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'العرض المالي', '12', client_logo)
    _divider(slide)
    _para(tf3, 'ملخص التكاليف والشروط المالية', 15, color=MINT, align=PP_ALIGN.RIGHT)

    pricing = data.get('sections', {}).get('pricing', {})
    subtotal = pricing.get('subtotal_omr', 0)
    vat = pricing.get('vat_omr', 0)
    total = pricing.get('grand_total_omr', 0)

    vat_rate = float(pricing.get('vat_rate', 0.05))
    vat_pct  = int(round(vat_rate * 100))
    validity = data.get('stats', {}).get('validity', '90 يوماً')
    bond_rate = data.get('bond_rate', None)
    client_name = data.get('client_name', '')
    _GOV = ('وزارة', 'سلطنة', 'حكومة', 'أمانة', 'بلدية', 'مجلس', 'هيئة', 'ديوان', 'دائرة', 'أكاديمية', 'gov')
    is_gov = any(k in client_name for k in _GOV)

    boxes = [
        ('الإجمالي الكلي شامل الضريبة', f'{total:,.0f}', 'ريال عماني', MINT),
        (f'ضريبة القيمة المضافة ({vat_pct}%)', f'{vat:,.0f}', 'ريال عماني', TEAL),
        ('المجموع قبل الضريبة', f'{subtotal:,.0f}', 'ريال عماني', NAVY),
    ]
    bw = (W - 1.5) / 3 - 0.3
    for i, (label, val, unit, col) in enumerate(boxes):
        bx = 0.5 + i * (bw + 0.45)
        _rect(slide, bx, 3.2, bw, 4.5, col, rounded=True, radius=0.12)
        tx = _txbox(slide, bx + 0.2, 3.4, bw - 0.4, 4.1)
        tf = tx.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf)
        _para(tf, val, 30, bold=True, color=WHITE, align=PP_ALIGN.CENTER, rtl=False)
        _para(tf, unit, 13, color=MGREEN if col == MINT else WHITE,
              align=PP_ALIGN.RIGHT, rtl=True, space_before=3)
        _para(tf, label, 14, color=WHITE, align=PP_ALIGN.RIGHT, rtl=True, space_before=8)

    compliance = [('\U0001f4b0', f'ضريبة القيمة المضافة {vat_pct}%', 'مضمنة في السعر الكلي')]
    if is_gov:
        compliance.append(('\U0001f4c4', 'مساهمة المنافسة 1.2%', 'وفق التعميم رقم 13/2024'))
    if bond_rate:
        bond_pct = int(round(float(bond_rate) * 100))
        compliance.append(('\U0001f6e1', f'ضمان حسن التنفيذ {bond_pct}%', 'من قيمة العقد'))
    compliance.append(('\U0001f4c5', 'صلاحية العرض', validity))
    n_comp = len(compliance)
    cw = (W - 1.5) / n_comp - 0.25
    for i, (icon, title, val) in enumerate(compliance):
        cx = 0.5 + (n_comp - 1 - i) * (cw + 0.25)
        _rect(slide, cx, 8.3, cw, 3.5, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _icon_circle(slide, cx + cw / 2, 9.2, 0.6, icon, 16, NAVY)
        tx = _txbox(slide, cx + 0.2, 10.2, cw - 0.4, 1.4)
        tf = tx.text_frame
        tf.word_wrap = True
        _tf_clear(tf)
        _para(tf, title, 13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        _para(tf, val, 12, color=TEAL, align=PP_ALIGN.RIGHT, space_before=3)

    _rect(slide, 0.5, 12.1, W - 1.0, 2.2, MGREEN, rounded=True, radius=0.1)
    tx2 = _txbox(slide, 1.0, 12.3, W - 2.0, 1.8)
    tf2 = tx2.text_frame
    tf2.word_wrap = True
    _tf_clear(tf2)
    _para(tf2, 'يتوافق هذا العرض مع أهداف رؤية عُمان 2040 في بناء الكفاءات الوطنية وتطوير القيادات الحكومية.',
          14, color=TEAL, align=PP_ALIGN.RIGHT, bold=True)


# ═══════════════════════════════════════════════════════════════════
# SLIDE: FINANCIAL TABLE (LINE ITEMS)
# ═══════════════════════════════════════════════════════════════════

def _slide_financial_table(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'العرض المالي التفصيلي', '12', client_logo)
    _divider(slide)
    _para(tf3, 'تفصيل بنود التكاليف', 15, color=MINT, align=PP_ALIGN.RIGHT)

    pricing = data.get('sections', {}).get('pricing', {})
    items = pricing.get('line_items', [])

    if not items:
        return

    headers = ['الإجمالي (ع.ر)', 'السعر (ع.ر)', 'الوحدة', 'الكمية', 'وصف البند', 'م']
    col_ws_raw = [4.5, 4.0, 3.0, 2.5, 16.0, 1.5]
    total_w = sum(col_ws_raw)
    scale = (W - 1.0) / total_w
    col_ws = [c * scale for c in col_ws_raw]

    header_h = 1.3
    row_h = 1.4
    table_x = 0.5
    table_y = 3.2

    cx = table_x
    for hdr, cw in zip(headers, col_ws):
        _rect(slide, cx, table_y, cw, header_h, NAVY)
        tx = _txbox(slide, cx + 0.05, table_y + 0.1, cw - 0.1, header_h - 0.2)
        tf = tx.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf)
        _para(tf, hdr, 12, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, rtl=True)
        cx += cw

    for ri, item in enumerate(items):
        row_color = WHITE if ri % 2 == 0 else LGRAY
        row_y = table_y + header_h + ri * row_h
        def _fmt(v):
            try:
                return f'{int(round(float(v))):,}'
            except (ValueError, TypeError):
                return str(v)
        cells = [
            _fmt(item.get('total_omr', '')),
            _fmt(item.get('rate_omr', '')),
            str(item.get('unit', '')),
            str(item.get('quantity', '')),
            str(item.get('description', '')),
            str(ri + 1),
        ]
        cx = table_x
        for cell, cw in zip(cells, col_ws):
            _rect(slide, cx, row_y, cw, row_h, row_color, border=MGRAY)
            tx = _txbox(slide, cx + 0.05, row_y + 0.1, cw - 0.1, row_h - 0.2)
            tf = tx.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            _tf_clear(tf)
            _para(tf, cell, 12, color=SLATE, align=PP_ALIGN.RIGHT, rtl=True)
            cx += cw

    subtotal = pricing.get('subtotal_omr', 0)
    vat = pricing.get('vat_omr', 0)
    total = pricing.get('grand_total_omr', 0)
    last_y = table_y + header_h + len(items) * row_h

    vat_rate2 = float(pricing.get('vat_rate', 0.05))
    vat_pct2  = int(round(vat_rate2 * 100))
    for label, val, bg in [
        ('المجموع قبل الضريبة', f'{int(round(subtotal)):,} ريال', LGRAY),
        (f'ضريبة القيمة المضافة {vat_pct2}%', f'{int(round(vat)):,} ريال', MGREEN),
        ('الإجمالي الكلي', f'{int(round(total)):,} ريال', NAVY),
    ]:
        _rect(slide, table_x, last_y, W - 1.0, 1.1, bg, border=MGRAY)
        tx = _txbox(slide, table_x + 0.2, last_y, W - 3.5, 1.1)
        tf = tx.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf)
        col = WHITE if bg == NAVY else SLATE
        _para(tf, label, 13, bold=True, color=col, align=PP_ALIGN.RIGHT, rtl=True)
        tx2 = _txbox(slide, W - 3.0, last_y, 2.5, 1.1)
        tf2 = tx2.text_frame
        tf2.word_wrap = False
        tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf2)
        _para(tf2, val, 13, bold=True, color=MINT if bg == NAVY else NAVY,
              align=PP_ALIGN.CENTER, rtl=False)
        last_y += 1.1


# ═══════════════════════════════════════════════════════════════════
# SLIDE: GOVERNANCE & QUALITY ASSURANCE
# ═══════════════════════════════════════════════════════════════════

def _slide_governance(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'ضمان الجودة والحوكمة', '11', client_logo)
    _divider(slide)
    _para(tf3, 'آليات الإشراف والمراجعة والاعتماد', 15, color=MINT, align=PP_ALIGN.RIGHT)

    stages = [
        ('\U0001f4cb', 'مراجعة داخلية', 'يراجع فريق زواد جميع المخرجات قبل التسليم للتأكد من الجودة والدقة'),
        ('\U0001f91d', 'تقييم المشاركين', 'استبيانات رضا يومية وتقييم بعد انتهاء البرنامج (Kirkpatrick Level 1 & 2)'),
        ('\U0001f4ca', 'تقرير الأداء', 'تقرير ختامي شامل بالنتائج والتوصيات وقياس الأثر خلال 15 يوم عمل'),
        ('\U0001f3c6', 'اعتماد الجهة', 'موافقة الجهة المستفيدة على كل مرحلة قبل الانتقال للمرحلة التالية'),
    ]

    bw = (W - 1.5) / 2 - 0.3
    bh = (H - 4.0) / 2 - 0.4

    for i, (icon, title, desc) in enumerate(stages):
        row = i // 2
        col = 1 - (i % 2)
        bx = 0.5 + col * (bw + 0.6)
        by = 3.3 + row * (bh + 0.5)

        _rect(slide, bx, by, bw, bh, WHITE, rounded=True, radius=0.12, border=MGRAY)
        _rect(slide, bx, by, bw, 0.28, TEAL, rounded=False)
        _icon_circle(slide, bx + bw - 1.2, by + bh / 2, 0.8, icon, 22, MINT)

        tx = _txbox(slide, bx + 0.4, by + 0.45, bw - 2.0, 1.1)
        tf = tx.text_frame
        tf.word_wrap = True
        _tf_clear(tf)
        _para(tf, title, 16, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)

        tx2 = _txbox(slide, bx + 0.4, by + 1.6, bw - 2.0, bh - 2.0)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        _tf_clear(tf2)
        _para(tf2, desc, 13, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE: RISK MANAGEMENT MATRIX
# ═══════════════════════════════════════════════════════════════════

def _slide_risk_matrix(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'إدارة المخاطر', '09', client_logo)
    _divider(slide)
    _para(tf3, 'تحديد المخاطر المحتملة وخطط المعالجة', 15, color=MINT, align=PP_ALIGN.RIGHT)

    risks = data.get('sections', {}).get('risks', [])
    if not risks:
        risks = [
            {'risk': 'تأخر في بدء البرنامج',
             'probability': 'متوسطة', 'impact': 'عالي',
             'preventive': 'التنسيق المبكر مع الجهة قبل شهر',
             'contingency': 'تفعيل خطة الجدول البديل'},
            {'risk': 'انخفاض نسبة الحضور',
             'probability': 'منخفضة', 'impact': 'متوسط',
             'preventive': 'قوائم احتياطية وتذكيرات متعددة',
             'contingency': 'إعادة جدولة الجلسات المتأثرة'},
            {'risk': 'ضعف التفاعل والمشاركة',
             'probability': 'منخفضة', 'impact': 'متوسط',
             'preventive': 'أنشطة تفاعلية ومحتوى مخصص للجمهور',
             'contingency': 'تعديل الأسلوب التدريبي بشكل آني'},
            {'risk': 'صعوبة تطبيق المخرجات',
             'probability': 'متوسطة', 'impact': 'عالي',
             'preventive': 'ربط المحتوى بالسياق المؤسسي الفعلي',
             'contingency': 'جلسات متابعة وكوتشينج لما بعد البرنامج'},
            {'risk': 'تغيير في الجهة المشرفة',
             'probability': 'منخفضة', 'impact': 'منخفض',
             'preventive': 'توثيق كامل لكل مراحل التنفيذ',
             'contingency': 'اجتماع تعريفي فوري مع الجهة الجديدة'},
        ]

    headers = ['خطة الطوارئ', 'الإجراء الوقائي', 'الأثر', 'الاحتمالية', 'المخاطرة', 'م']
    # Wider columns for text-heavy contingency/preventive; narrow for numbers/status
    col_ws_raw = [7.5, 7.5, 2.5, 2.8, 10.5, 1.5]
    total_w = sum(col_ws_raw)
    scale = (W - 1.0) / total_w
    col_ws = [c * scale for c in col_ws_raw]

    header_h = 1.3
    row_h = 2.2          # was 1.5 — taller rows so text never overflows
    table_x = 0.5
    table_y = 3.2

    cx = table_x
    for hdr, cw in zip(headers, col_ws):
        _rect(slide, cx, table_y, cw, header_h, NAVY)
        tx = _txbox(slide, cx + 0.05, table_y, cw - 0.1, header_h)
        tf = tx.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf)
        _para(tf, hdr, 11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, rtl=True)
        cx += cw

    impact_colors = {'عالي': RGBColor(0xFF, 0x5C, 0x5C), 'متوسط': RGBColor(0xFF, 0xB6, 0x00),
                     'منخفض': RGBColor(0x4C, 0xAF, 0x50), 'منخفضة': RGBColor(0x4C, 0xAF, 0x50),
                     'متوسطة': RGBColor(0xFF, 0xB6, 0x00), 'عالية': RGBColor(0xFF, 0x5C, 0x5C)}

    for ri, risk in enumerate(risks[:5]):
        row_color = WHITE if ri % 2 == 0 else LGRAY
        row_y = table_y + header_h + ri * row_h
        cells = [
            risk.get('contingency', ''), risk.get('preventive', ''),
            risk.get('impact', ''), risk.get('probability', ''),
            risk.get('risk', ''), _ar(ri + 1),
        ]
        cx = table_x
        for j, (cell, cw) in enumerate(zip(cells, col_ws)):
            cell_color = row_color
            if j in (2, 3):
                cell_color = impact_colors.get(cell, row_color)
            _rect(slide, cx, row_y, cw, row_h, cell_color, border=MGRAY)
            tx = _txbox(slide, cx + 0.08, row_y + 0.1, cw - 0.16, row_h - 0.2)
            tf = tx.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            _tf_clear(tf)
            txt_color = WHITE if j in (2, 3) and cell_color != row_color else SLATE
            font_sz = 10 if j in (0, 1, 5) else 11
            is_center = j in (2, 3)
            _para(tf, cell, font_sz, bold=(j in (2, 3)), color=txt_color,
                  align=PP_ALIGN.RIGHT, rtl=True)
            cx += cw


# ═══════════════════════════════════════════════════════════════════
# SLIDE: DECLARATIONS & REQUIRED ATTACHMENTS
# ═══════════════════════════════════════════════════════════════════

def _slide_declarations(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'الإقرارات والمستندات المطلوبة', '14', client_logo)
    _divider(slide)
    _para(tf3, 'المتطلبات الرسمية للعرض', 15, color=MINT, align=PP_ALIGN.RIGHT)

    bond_rate_d = data.get('bond_rate', None)
    if bond_rate_d:
        bond_pct_d = int(round(float(bond_rate_d) * 100))
        bond_decl = ('✅', 'ضمان حسن التنفيذ', f'على استعداد لتقديم ضمان حسن التنفيذ {bond_pct_d}% عند الإسناد')
    else:
        bond_decl = ('✅', 'الالتزام بالتنفيذ', 'نلتزم بتنفيذ البرنامج وفق الشروط المتفق عليها بالكامل')
    declarations = [
        ('✅', 'إقرار الاعتراف بشروط المناقصة', 'نُقر بالاطلاع على جميع شروط وأحكام المناقصة وقبولها'),
        ('✅', 'السجل التجاري', 'مرفق سجل تجاري ساري المفعول صادر من وزارة التجارة والصناعة وترويج الاستثمار'),
        ('✅', 'شهادة التسجيل الضريبي', 'مرفق شهادة التسجيل في الضريبة الانتقائية — رقم VAT مفعّل'),
        ('✅', 'السيرة الذاتية لفريق العمل', 'مرفقة السير الذاتية لجميع أعضاء الفريق المعينين للتنفيذ'),
        bond_decl,
        ('✅', 'إقرار عدم التعارض', 'نُقر بعدم وجود أي تعارض في المصالح مع الجهة المُناقِصة'),
    ]

    bh = 1.55
    bw_left = W * 0.55
    bw_right = W * 0.43

    left_items = declarations[:3]
    right_items = declarations[3:]

    for i, (icon, title, desc) in enumerate(left_items):
        by = 3.2 + i * (bh + 0.22)
        bx = W - bw_left - 0.4
        _rect(slide, bx, by, bw_left, bh, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _rect(slide, bx, by, 0.28, bh, MINT)
        _icon_circle(slide, bx + bw_left - 1.0, by + bh / 2, 0.5, icon, 14, TEAL)
        tx = _txbox(slide, bx + 0.5, by + 0.1, bw_left - 1.8, 0.75)
        tf = tx.text_frame
        _tf_clear(tf)
        _para(tf, title, 13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        tx2 = _txbox(slide, bx + 0.5, by + 0.8, bw_left - 1.8, 0.65)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        _tf_clear(tf2)
        _para(tf2, desc, 11, color=SLATE, align=PP_ALIGN.RIGHT)

    for i, (icon, title, desc) in enumerate(right_items):
        by = 3.2 + i * (bh + 0.22)
        bx = 0.4
        _rect(slide, bx, by, bw_right, bh, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _rect(slide, bx, by, 0.28, bh, TEAL)
        _icon_circle(slide, bx + bw_right - 1.0, by + bh / 2, 0.5, icon, 14, MINT)
        tx = _txbox(slide, bx + 0.5, by + 0.1, bw_right - 1.8, 0.75)
        tf = tx.text_frame
        _tf_clear(tf)
        _para(tf, title, 13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        tx2 = _txbox(slide, bx + 0.5, by + 0.8, bw_right - 1.8, 0.65)
        tf2 = tx2.text_frame
        tf2.word_wrap = True
        _tf_clear(tf2)
        _para(tf2, desc, 11, color=SLATE, align=PP_ALIGN.RIGHT)

    sig_y = H - 2.8
    _rect(slide, 0.5, sig_y, W - 1.0, 2.3, WHITE, rounded=True, radius=0.08, border=MGRAY)
    tx_s = _txbox(slide, 1.0, sig_y + 0.2, W - 3.5, 1.0)
    tf_s = tx_s.text_frame
    _tf_clear(tf_s)
    _para(tf_s, 'المفوَّض بالتوقيع: .....................................    التاريخ: ......../........./' + '2026',
          13, color=SLATE, align=PP_ALIGN.RIGHT)
    tx_s2 = _txbox(slide, 1.0, sig_y + 1.1, W - 3.5, 0.9)
    tf_s2 = tx_s2.text_frame
    _tf_clear(tf_s2)
    _para(tf_s2, 'الختم الرسمي: .....................................    رقم التسجيل: ................................',
          13, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE: COMMITMENTS
# ═══════════════════════════════════════════════════════════════════

def _slide_commitments_page(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'الإقرارات والالتزامات', '13', client_logo)
    _divider(slide)
    _para(tf3, 'التزاماتنا تجاهكم', 15, color=MINT, align=PP_ALIGN.RIGHT)

    commit_text = data.get('sections', {}).get('commitments', '')
    lines = [l.strip() for l in commit_text.split('\n') if l.strip()]

    defaults = [
        'نلتزم بتقديم برنامج تدريبي مصمم خصيصاً لاحتياجات الجهة، مع ضمان جودة المحتوى وملاءمته للسياق المؤسسي العُماني.',
        'نضمن حضور المدربين المعتمدين في جميع الجلسات، مع الالتزام بالجدول الزمني المتفق عليه دون أي تأخير.',
        'نوفر تقارير تقييمية شاملة بعد كل جلسة تدريبية، وتقرير ختامي مفصّل يشمل مقياس الأثر وفق نموذج كيركباتريك.',
        'نلتزم بمبدأ السرية التامة لجميع معلومات الجهة والمشاركين، ولا يُشارك أي محتوى خارج نطاق البرنامج.',
        'نوفر دعماً فنياً وتنسيقياً طوال فترة تنفيذ البرنامج، مع نقطة اتصال واحدة مخصصة للجهة.',
        'نلتزم بتقديم ضمان حسن التنفيذ وفق اشتراطات المناقصة، وإعادة أي جلسة لم تحقق معيار الجودة المتفق عليه.',
    ]
    if len(lines) < 3:
        lines = lines + defaults[len(lines):]

    icons = ['✔', '\U0001f91d', '\U0001f4c5', '⏰', '\U0001f512', '\U0001f3c6']
    bh = (H - 4.2) / max(len(lines), 1) - 0.3
    bh = min(bh, 2.8)

    for i, line in enumerate(lines[:6]):
        by = 3.2 + i * (bh + 0.3)
        _rect(slide, 0.5, by, W - 1.0, bh, WHITE, rounded=True, radius=0.1, border=MGRAY)
        _icon_circle(slide, W - 1.7, by + bh / 2, 0.6, icons[i % len(icons)], 16, MINT)
        _rect(slide, 0.5, by, 0.3, bh, TEAL)
        tx = _txbox(slide, 1.0, by + 0.15, W - 3.5, bh - 0.3)
        tf = tx.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf)
        _para(tf, line, 15, color=SLATE, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════
# SLIDE: GANTT CHART
# ═══════════════════════════════════════════════════════════════════

def _slide_gantt(prs, data, client_logo=None):
    """Gantt chart slide — project phases over weeks."""
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, LGRAY)
    tf3 = _header(slide, 'الجدول الزمني التنفيذي', '06', client_logo)
    _divider(slide)
    _para(tf3, 'مراحل التنفيذ والإطار الزمني التفصيلي', 15, color=MINT, align=PP_ALIGN.RIGHT)

    # Layout constants
    tbl_x = 0.5
    tbl_y = 3.1
    tbl_w = W - 1.0
    tbl_h = H - tbl_y - 0.3

    # Default phases/tasks — overridable from data
    gantt_data = data.get('gantt_phases', [
        {
            'phase': 'مرحلة الإعداد',
            'color': RGBColor(0x0A, 0x26, 0x47),  # NAVY
            'tasks': [
                {'task': 'دراسة وثائق المناقصة وتحليل الاحتياج', 'start': 1, 'end': 1},
                {'task': 'تصميم المحتوى ومواءمة المواد', 'start': 1, 'end': 2},
                {'task': 'التنسيق مع الجهة واعتماد الخطة', 'start': 2, 'end': 3},
            ]
        },
        {
            'phase': 'مرحلة التنفيذ',
            'color': RGBColor(0x2E, 0xC4, 0xB6),  # MINT
            'tasks': [
                {'task': 'تنفيذ أيام التدريب الحضورية', 'start': 4, 'end': 6},
                {'task': 'جلسات التوجيه والكوتشينج الفردي', 'start': 4, 'end': 7},
                {'task': 'التقييمات القبلية والبعدية', 'start': 4, 'end': 6},
            ]
        },
        {
            'phase': 'مرحلة الإغلاق',
            'color': RGBColor(0x0E, 0x7C, 0x86),  # TEAL
            'tasks': [
                {'task': 'إعداد التقارير الختامية', 'start': 7, 'end': 8},
                {'task': 'تسليم المخرجات النهائية', 'start': 8, 'end': 9},
                {'task': 'قياس الأثر والمتابعة البعدية', 'start': 9, 'end': 11},
            ]
        },
    ])

    # Calculate total weeks
    all_ends = [t['end'] for ph in gantt_data for t in ph['tasks']]
    total_weeks = max(all_ends) if all_ends else 9
    total_tasks = sum(len(ph['tasks']) for ph in gantt_data)

    # RTL layout: phase on RIGHT, task next, weeks fill LEFT (week 1 rightmost)
    phase_col_w = 3.2
    task_col_w = 7.5
    tbl_right = tbl_x + tbl_w              # right edge of table area

    phase_left = tbl_right - phase_col_w   # phase column left edge
    task_left = phase_left - task_col_w    # task column left edge
    week_area_w = task_left - tbl_x        # weeks fill from tbl_x to task_left
    week_col_w = week_area_w / total_weeks

    row_h = tbl_h / (total_tasks + 1)
    header_h = min(row_h, 0.8)
    row_h = (tbl_h - header_h) / total_tasks

    # Header row — RTL: phase RIGHT, task middle-right, weeks left→right but week 1 is rightmost
    # Phase header (rightmost)
    _rect(slide, phase_left, tbl_y, phase_col_w, header_h, NAVY)
    tx = _txbox(slide, phase_left, tbl_y, phase_col_w, header_h)
    tf = tx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
    _para(tf, 'المرحلة', 10, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    # Task header
    _rect(slide, task_left, tbl_y, task_col_w, header_h, NAVY)
    tx = _txbox(slide, task_left, tbl_y, task_col_w, header_h)
    tf = tx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
    _para(tf, 'المهمة / النشاط', 10, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    # Week column headers: week 1 is rightmost (adjacent to task column), week N leftmost
    for w in range(total_weeks):
        wx = task_left - (w + 1) * week_col_w
        _rect(slide, wx, tbl_y, week_col_w, header_h, TEAL)
        tx = _txbox(slide, wx, tbl_y, week_col_w, header_h)
        tf = tx.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
        _para(tf, _ar(w + 1), 9, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, rtl=True)

    # Data rows
    row_idx = 0
    for ph in gantt_data:
        phase_color = ph.get('color', NAVY)
        tasks = ph['tasks']

        for t_idx, task in enumerate(tasks):
            ry = tbl_y + header_h + row_idx * row_h
            bg = WHITE if row_idx % 2 == 0 else LGRAY

            # Phase cell (merged across its tasks) — on the RIGHT
            if t_idx == 0:
                _rect(slide, phase_left, ry, phase_col_w, row_h * len(tasks), phase_color)
                tx = _txbox(slide, phase_left, ry, phase_col_w, row_h * len(tasks))
                tf = tx.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
                _para(tf, ph['phase'], 10, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

            # Task cell — adjacent to phase on the left
            _rect(slide, task_left, ry, task_col_w, row_h, bg, border=MGRAY)
            tx = _txbox(slide, task_left + 0.1, ry, task_col_w - 0.2, row_h)
            tf = tx.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; _tf_clear(tf)
            _para(tf, task['task'], 10, color=SLATE, align=PP_ALIGN.RIGHT)

            # Week background cells: week 1 rightmost, week N leftmost
            for w in range(total_weeks):
                wx = task_left - (w + 1) * week_col_w
                _rect(slide, wx, ry, week_col_w, row_h, bg, border=MGRAY)

            # Colored bar: spans start..end weeks in RTL coordinates
            s, e = task['start'], task['end']
            bar_x = task_left - e * week_col_w + 0.05
            bar_w = (e - s + 1) * week_col_w - 0.1
            bar_y = ry + row_h * 0.2
            bar_h_val = row_h * 0.6
            _rect(slide, bar_x, bar_y, bar_w, bar_h_val, phase_color, rounded=True, radius=0.08)

            row_idx += 1


# ═══════════════════════════════════════════════════════════════════
# SLIDE: CLOSING
# ═══════════════════════════════════════════════════════════════════

def _slide_closing(prs, data, client_logo):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, W, H, NAVY)

    _orb(slide, -3, -3, 9, MINT)
    _orb(slide, W - 6, -3, 9, TEAL)
    _orb(slide, -2, H - 5, 7, TEAL)
    _orb(slide, W - 5, H - 5, 8, MINT)

    box_w, box_h = 12, 3.8
    box_x, box_y = (W - box_w) / 2, 2.0
    _rect(slide, box_x, box_y, box_w, box_h, WHITE, rounded=True, radius=0.08)
    _img(slide, ZAWAD_LOGO, box_x + box_w / 2 + 0.3, box_y + 0.25, 5.2, 3.2)
    if client_logo and os.path.exists(client_logo):
        _img(slide, client_logo, box_x + 0.4, box_y + 0.3, 5.0, 3.1)
        _rect(slide, box_x + box_w / 2 - 0.03, box_y + 0.5, 0.06, 2.7, MGRAY)

    tx = _txbox(slide, 2, 6.5, W - 4, 2.0)
    tf = tx.text_frame
    tf.word_wrap = False
    _tf_clear(tf)
    _para(tf, 'شكراً لثقتكم', 36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    closing_msg = ('نتطلع إلى شراكة مع ' + data.get('client_name', 'الجهة المحترمة') +
                   ' لتنفيذ هذا البرنامج النوعي، دعماً لمستهدفات رؤية عُمان 2040.')
    tx2 = _txbox(slide, 2.5, 8.8, W - 5, 1.8)
    tf2 = tx2.text_frame
    tf2.word_wrap = True
    _tf_clear(tf2)
    _para(tf2, closing_msg, 14, color=MGREEN, align=PP_ALIGN.RIGHT)

    contacts = [
        ('\U0001f310', 'الموقع الإلكتروني', 'www.zawad.om'),
        ('✉', 'البريد الإلكتروني', 'info@zawad.om'),
        ('\U0001f4de', 'الهاتف', '+968 99232775'),
        ('\U0001f4cd', 'العنوان', 'العذيبة، مسقط'),
    ]
    box_top = 10.5
    box_h = 5.0
    cw = (W - 2.0) / 4 - 0.3
    for i, (icon, label, val) in enumerate(contacts):
        cx = 1.0 + (3 - i) * (cw + 0.3)
        _rect(slide, cx, box_top, cw, box_h, RGBColor(0x0D, 0x2D, 0x55),
              rounded=True, radius=0.12, border=MINT)
        # Icon centred in upper half of box
        _icon_circle(slide, cx + cw / 2, box_top + 1.3, 0.65, icon, 18, MINT)
        # Label
        tx3 = _txbox(slide, cx + 0.15, box_top + 2.6, cw - 0.3, 0.85)
        tf3 = tx3.text_frame
        tf3.word_wrap = True
        tf3.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf3)
        _para(tf3, label, 11, color=MINT, align=PP_ALIGN.RIGHT)
        # Value — taller box so text never clips
        tx4 = _txbox(slide, cx + 0.15, box_top + 3.45, cw - 0.3, 1.3)
        tf4 = tx4.text_frame
        tf4.word_wrap = True
        tf4.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_clear(tf4)
        _para(tf4, val, 12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, rtl=False)

    tx5 = _txbox(slide, 2, 16.2, W - 4, 0.9)
    tf5 = tx5.text_frame
    _tf_clear(tf5)
    _para(tf5, 'والله ولي التوفيق', 13, color=MINT, align=PP_ALIGN.RIGHT, italic=True)


# ═══════════════════════════════════════════════════════════════════
# LOGO FETCHER
# ═══════════════════════════════════════════════════════════════════

def _fetch_client_logo(data: dict) -> Optional[str]:
    """Try to download client logo. Returns local path if successful, else None."""
    import urllib.request
    import tempfile

    client_name = data.get('client_name', '')
    client_website = data.get('client_website', '')
    if not client_name:
        return None

    save_dir = Path(tempfile.gettempdir()) / 'zawad_logos'
    save_dir.mkdir(exist_ok=True)
    safe_name = ''.join(c for c in client_name if c.isalnum() or c in ' _-')[:40]
    logo_path = save_dir / f'{safe_name}.png'

    if logo_path.exists():
        return str(logo_path)

    domain = ''
    if client_website:
        domain = client_website.replace('https://', '').replace('http://', '').split('/')[0]
    elif 'أكاديمية' in client_name or 'academy' in client_name.lower():
        domain = 'royalacademy.gov.om'
    elif 'وزارة الخدمة المدنية' in client_name:
        domain = 'moc.gov.om'
    elif 'وزارة التنمية الاجتماعية' in client_name:
        domain = 'mosd.gov.om'
    elif 'أوكيو' in client_name or 'oq' in client_name.lower():
        domain = 'oqomholdinggroup.com'
    elif 'ناما' in client_name or 'nama' in client_name.lower():
        domain = 'namaholding.om'

    if domain:
        try:
            url = f'https://logo.clearbit.com/{domain}?size=200'
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as resp:
                data_bytes = resp.read()
                # Must be a valid PNG (magic bytes) or JPEG — reject HTML/WebP error pages
                is_png  = data_bytes[:8] == b'\x89PNG\r\n\x1a\n'
                is_jpeg = data_bytes[:2] == b'\xff\xd8'
                if (is_png or is_jpeg) and len(data_bytes) > 1000:
                    with open(logo_path, 'wb') as f:
                        f.write(data_bytes)
                    return str(logo_path)
        except Exception:
            pass

    return None


# ═══════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════

def _sanitize_data(obj):
    """
    STRICT RULE — applied to every string in proposal data before building:
      1. Arabic/number spacing: always space between digits and Arabic on both sides
      2. Bracket spacing: always space after ( and before ) when adjacent to Arabic
      3. Pricing: integer format with commas, no decimals
    This runs recursively on the full data dict so no field can bypass it.
    """
    if isinstance(obj, str):
        return _fix_spacing(_strip_invalid_xml(obj))
    if isinstance(obj, dict):
        return {k: _sanitize_data(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_data(i) for i in obj]
    return obj


def build_pptx(data: dict, output_path: str):
    # STRICT RULE: sanitize ALL strings before any slide is drawn
    data = _sanitize_data(data)
    lang = data.get('language', 'ar')

    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)

    client_logo = _fetch_client_logo(data)

    _slide_cover(prs, data, client_logo)
    _slide_cover_letter(prs, data, client_logo)
    _slide_toc(prs, data)
    _slide_stats(prs, data, client_logo)
    _slide_commitments_grid(prs, data, client_logo)
    _slide_company(prs, data, client_logo)
    _slide_clients(prs, client_logo)
    _slide_requirements(prs, data, client_logo)
    _slide_outputs(prs, data, client_logo)
    _slide_methodology(prs, data, client_logo)

    workplan = data.get('sections', {}).get('workplan', [])
    for i, day in enumerate(workplan):
        _slide_workplan_day(prs, day, i + 1, len(workplan), client_logo)

    _slide_gantt(prs, data, client_logo)
    _slide_team_pages(prs, data, client_logo, lang=lang)
    _slide_track_record(prs, data, client_logo)
    _slide_innovation(prs, data, client_logo)
    _slide_governance(prs, data, client_logo)
    _slide_risk_matrix(prs, data, client_logo)
    _slide_financial_summary(prs, data, client_logo)
    _slide_financial_table(prs, data, client_logo)
    _slide_commitments_page(prs, data, client_logo)
    _slide_declarations(prs, data, client_logo)
    _slide_closing(prs, data, client_logo)

    # Post-processing: remove any normAutofit that coexists with noAutofit (OOXML violation)
    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for slide in prs.slides:
        for el in slide._element.iter('{%s}bodyPr' % ns_a):
            norm = el.find('{%s}normAutofit' % ns_a)
            no   = el.find('{%s}noAutofit'   % ns_a)
            if norm is not None and no is not None:
                el.remove(norm)

    prs.save(output_path)
