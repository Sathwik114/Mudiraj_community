import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PUBLIC_DIR = os.path.join(BASE_DIR, "public")
DOWNLOADS_DIR = os.path.join(PUBLIC_DIR, "downloads")
LOGO_PATH = os.path.join(PUBLIC_DIR, "mudiraj_logo.png")
APPLY_SCREENSHOT_PATH = os.path.join(PUBLIC_DIR, "apply_form_screenshot.png")

os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Executive Color Palette
NAVY_DARK = RGBColor(15, 35, 71)        # #0f2347 (Header, Brand Apex)
NAVY_PRIMARY = RGBColor(30, 58, 138)    # #1e3a8a (Primary Theme)
NAVY_DEEP = RGBColor(23, 37, 84)        # #172554
AMBER_ACCENT = RGBColor(217, 119, 6)    # #d97706 (Gold Accent / Buttons)
GOLD_LIGHT = RGBColor(251, 191, 36)     # #fbbf24 (Highlights, Badges)
GOLD_BG = RGBColor(254, 243, 199)       # #fef3c7 (Warning / Pending bg)
SLATE_DARK = RGBColor(30, 41, 59)       # #1e293b (Body Text)
SLATE_MUTED = RGBColor(100, 116, 139)   # #64748b (Secondary Text)
SLATE_LIGHT = RGBColor(248, 250, 252)   # #f8fafc (Card Canvas bg)
SLATE_BORDER = RGBColor(203, 213, 225)  # #cbd5e1 (Form Field Border)
WHITE = RGBColor(255, 255, 255)
GREEN_SUCCESS = RGBColor(16, 185, 129)  # #10b981 (Approved / Success)
GREEN_BG = RGBColor(220, 252, 231)      # #dcfce7
RED_DANGER = RGBColor(220, 38, 38)      # #dc2626 (Required / Rejected)
RED_BG = RGBColor(254, 226, 226)        # #fee2e2
BLUE_INFO = RGBColor(37, 99, 235)       # #2563eb (Card ready / Info)
BLUE_BG = RGBColor(219, 234, 254)       # #dbeafe
PURPLE_ACCENT = RGBColor(126, 34, 206)  # #7e22ce
PURPLE_BG = RGBColor(243, 232, 255)     # #f3e8ff

def init_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

# ------------------------------------------------------------------------------
# BASE SLIDE DECORATIONS (Header, Footer, Titles)
# ------------------------------------------------------------------------------
def add_header(slide, title_text, category_badge="ANDHRA PRADESH MUDIRAJ MAHASABHA", subtitle_text=""):
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = NAVY_DARK
    header_box.line.color.rgb = NAVY_DARK

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.06))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = AMBER_ACCENT
    stripe.line.color.rgb = AMBER_ACCENT

    left_pos = Inches(0.6)
    if os.path.exists(LOGO_PATH):
        try:
            slide.shapes.add_picture(LOGO_PATH, Inches(0.4), Inches(0.12), width=Inches(0.9), height=Inches(0.9))
            left_pos = Inches(1.5)
        except Exception:
            left_pos = Inches(0.6)

    tb = slide.shapes.add_textbox(left_pos, Inches(0.1), Inches(11.2), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_badge = tf.paragraphs[0]
    p_badge.text = category_badge.upper()
    p_badge.font.size = Pt(9.5)
    p_badge.font.bold = True
    p_badge.font.color.rgb = GOLD_LIGHT

    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.size = Pt(19)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE

    if subtitle_text:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.size = Pt(10.5)
        p_sub.font.color.rgb = RGBColor(203, 213, 225)

def add_footer(slide, current_slide, total_slides, doc_title):
    f_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(7.05), Inches(12.133), Inches(0.02))
    f_line.fill.solid()
    f_line.fill.fore_color.rgb = SLATE_BORDER
    f_line.line.color.rgb = SLATE_BORDER

    tb_left = slide.shapes.add_textbox(Inches(0.6), Inches(7.1), Inches(8.5), Inches(0.35))
    tf_left = tb_left.text_frame
    tf_left.margin_left = tf_left.margin_top = tf_left.margin_right = tf_left.margin_bottom = 0
    p = tf_left.paragraphs[0]
    p.text = f"Andhra Pradesh Mudiraj Community Portal | {doc_title} | Official User Manual"
    p.font.size = Pt(9)
    p.font.color.rgb = SLATE_MUTED

    tb_right = slide.shapes.add_textbox(Inches(9.5), Inches(7.1), Inches(3.233), Inches(0.35))
    tf_right = tb_right.text_frame
    tf_right.margin_left = tf_right.margin_top = tf_right.margin_right = tf_right.margin_bottom = 0
    p_r = tf_right.paragraphs[0]
    p_r.alignment = PP_ALIGN.RIGHT
    p_r.text = f"Slide {current_slide} of {total_slides} • Downloadable from Navbar"
    p_r.font.size = Pt(9)
    p_r.font.bold = True
    p_r.font.color.rgb = NAVY_PRIMARY

def add_title_slide(prs, title, subtitle, badge, extra_notes, doc_label):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY_DARK
    bg.line.color.rgb = NAVY_DARK

    stripe_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18))
    stripe_top.fill.solid()
    stripe_top.fill.fore_color.rgb = AMBER_ACCENT
    stripe_top.line.color.rgb = AMBER_ACCENT

    stripe_bot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.32), Inches(13.333), Inches(0.18))
    stripe_bot.fill.solid()
    stripe_bot.fill.fore_color.rgb = AMBER_ACCENT
    stripe_bot.line.color.rgb = AMBER_ACCENT

    if os.path.exists(LOGO_PATH):
        try:
            slide.shapes.add_picture(LOGO_PATH, Inches(5.916), Inches(0.7), width=Inches(1.5), height=Inches(1.5))
        except Exception:
            pass

    top_pos = Inches(2.35) if os.path.exists(LOGO_PATH) else Inches(1.8)
    tb = slide.shapes.add_textbox(Inches(1.0), top_pos, Inches(11.333), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True

    p_badge = tf.paragraphs[0]
    p_badge.alignment = PP_ALIGN.CENTER
    p_badge.text = badge.upper()
    p_badge.font.size = Pt(13)
    p_badge.font.bold = True
    p_badge.font.color.rgb = GOLD_LIGHT

    p_space = tf.add_paragraph()
    p_space.text = ""
    p_space.font.size = Pt(6)

    p_title = tf.add_paragraph()
    p_title.alignment = PP_ALIGN.CENTER
    p_title.text = title
    p_title.font.size = Pt(30)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE

    p_sub = tf.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.text = subtitle
    p_sub.font.size = Pt(17)
    p_sub.font.color.rgb = RGBColor(226, 232, 240)

    p_space2 = tf.add_paragraph()
    p_space2.text = ""
    p_space2.font.size = Pt(8)

    for note in extra_notes:
        p_note = tf.add_paragraph()
        p_note.alignment = PP_ALIGN.CENTER
        p_note.text = note
        p_note.font.size = Pt(11.5)
        p_note.font.color.rgb = RGBColor(148, 163, 184)

    return slide

# ------------------------------------------------------------------------------
# RICH GRAPHICAL UI BUILDERS
# ------------------------------------------------------------------------------
def draw_mockup_window(slide, left, top, width, height, window_title="Online Form", url_badge="https://mudiraj.ap.gov.in/apply"):
    outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    outer.fill.solid()
    outer.fill.fore_color.rgb = SLATE_LIGHT
    outer.line.color.rgb = SLATE_BORDER
    outer.line.width = Pt(1.5)

    head_h = Inches(0.42)
    win_head = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, head_h)
    win_head.fill.solid()
    win_head.fill.fore_color.rgb = NAVY_DEEP
    win_head.line.color.rgb = NAVY_DEEP

    dot_colors = [RGBColor(239, 68, 68), RGBColor(245, 158, 11), RGBColor(16, 185, 129)]
    for i, col in enumerate(dot_colors):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.12 + i * 0.16), top + Inches(0.14), Inches(0.12), Inches(0.12))
        dot.fill.solid()
        dot.fill.fore_color.rgb = col
        dot.line.fill.background()

    url_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.7), top + Inches(0.08), width - Inches(0.85), Inches(0.26))
    url_box.fill.solid()
    url_box.fill.fore_color.rgb = RGBColor(15, 23, 42)
    url_box.line.color.rgb = RGBColor(51, 65, 85)
    url_box.line.width = Pt(1)

    tb = slide.shapes.add_textbox(left + Inches(0.8), top + Inches(0.09), width - Inches(1.0), Inches(0.24))
    tf = tb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"🔒 {url_badge}  |  {window_title}"
    p.font.size = Pt(8.5)
    p.font.color.rgb = RGBColor(203, 213, 225)

    inner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.08), top + head_h + Inches(0.02), width - Inches(0.16), height - head_h - Inches(0.1))
    inner.fill.solid()
    inner.fill.fore_color.rgb = WHITE
    inner.line.color.rgb = RGBColor(241, 245, 249)
    inner.line.width = Pt(1)

    return inner

def draw_input_field(slide, left, top, width, label, value_or_placeholder, is_required=True, is_select=False, badge=""):
    tb_lbl = slide.shapes.add_textbox(left, top, width, Inches(0.24))
    tf_lbl = tb_lbl.text_frame
    tf_lbl.margin_left = tf_lbl.margin_top = tf_lbl.margin_right = tf_lbl.margin_bottom = 0
    p_lbl = tf_lbl.paragraphs[0]
    run_l = p_lbl.add_run()
    run_l.text = label
    run_l.font.size = Pt(9)
    run_l.font.bold = True
    run_l.font.color.rgb = NAVY_PRIMARY

    if is_required:
        run_req = p_lbl.add_run()
        run_req.text = " *"
        run_req.font.bold = True
        run_req.font.color.rgb = RED_DANGER
        run_req.font.size = Pt(9.5)

    if badge:
        run_b = p_lbl.add_run()
        run_b.text = f"  [{badge}]"
        run_b.font.bold = True
        run_b.font.size = Pt(7.5)
        run_b.font.color.rgb = AMBER_ACCENT

    box_top = top + Inches(0.24)
    box_h = Inches(0.36)
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, box_top, width, box_h)
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = SLATE_BORDER
    box.line.width = Pt(1)

    tb_txt = slide.shapes.add_textbox(left + Inches(0.1), box_top + Inches(0.06), width - Inches(0.35), Inches(0.24))
    tf_txt = tb_txt.text_frame
    tf_txt.margin_left = tf_txt.margin_top = tf_txt.margin_right = tf_txt.margin_bottom = 0
    p_txt = tf_txt.paragraphs[0]
    p_txt.text = value_or_placeholder
    p_txt.font.size = Pt(8.5)
    p_txt.font.color.rgb = SLATE_DARK if not value_or_placeholder.startswith("Select") and not value_or_placeholder.startswith("Enter") else SLATE_MUTED

    if is_select:
        tb_arr = slide.shapes.add_textbox(left + width - Inches(0.3), box_top + Inches(0.05), Inches(0.25), Inches(0.25))
        tf_arr = tb_arr.text_frame
        tf_arr.margin_left = tf_arr.margin_top = tf_arr.margin_right = tf_arr.margin_bottom = 0
        p_arr = tf_arr.paragraphs[0]
        p_arr.alignment = PP_ALIGN.CENTER
        p_arr.text = "▼"
        p_arr.font.size = Pt(8)
        p_arr.font.color.rgb = NAVY_PRIMARY

def draw_button(slide, left, top, width, height, text, bg_color=AMBER_ACCENT, text_color=WHITE):
    btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    btn.fill.solid()
    btn.fill.fore_color.rgb = bg_color
    btn.line.color.rgb = bg_color

    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = text
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = text_color
    return btn

def draw_info_table(slide, left, top, width, height, headers, rows):
    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    tbl = table_shape.table

    for c_idx, h_text in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.LEFT

    for r_idx, row_data in enumerate(rows):
        bg_col = WHITE if r_idx % 2 == 0 else SLATE_LIGHT
        for c_idx, val in enumerate(row_data):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col
            p = cell.text_frame.paragraphs[0]
            p.text = str(val)
            p.font.size = Pt(8.5)
            p.font.color.rgb = SLATE_DARK
            if c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY_PRIMARY

add_info_table = draw_info_table

def draw_id_card(slide, left, top, is_front=True):
    w = Inches(3.3)
    h = Inches(4.9)

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = SLATE_BORDER
    card.line.width = Pt(1.5)

    if is_front:
        c_head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, Inches(1.05))
        c_head.fill.solid()
        c_head.fill.fore_color.rgb = NAVY_DARK
        c_head.line.color.rgb = NAVY_DARK

        c_stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + Inches(1.05), w, Inches(0.05))
        c_stripe.fill.solid()
        c_stripe.fill.fore_color.rgb = AMBER_ACCENT
        c_stripe.line.color.rgb = AMBER_ACCENT

        tb_h = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.12), w - Inches(0.2), Inches(0.85))
        tf_h = tb_h.text_frame
        tf_h.word_wrap = True
        p1 = tf_h.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = "ANDHRA PRADESH MUDIRAJ MAHASABHA"
        p1.font.bold = True
        p1.font.size = Pt(8.5)
        p1.font.color.rgb = GOLD_LIGHT

        p2 = tf_h.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = "OFFICIAL DIGITAL MEMBERSHIP CARD"
        p2.font.size = Pt(7.5)
        p2.font.color.rgb = WHITE

        ph = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(1.25), Inches(1.1), Inches(1.3))
        ph.fill.solid()
        ph.fill.fore_color.rgb = RGBColor(226, 232, 240)
        ph.line.color.rgb = NAVY_PRIMARY
        ph.line.width = Pt(1)

        tb_ph = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(1.7), Inches(1.1), Inches(0.4))
        p_ph = tb_ph.text_frame.paragraphs[0]
        p_ph.alignment = PP_ALIGN.CENTER
        p_ph.text = "MEMBER\nPHOTO"
        p_ph.font.size = Pt(8)
        p_ph.font.bold = True
        p_ph.font.color.rgb = SLATE_MUTED

        tb_det = slide.shapes.add_textbox(left + Inches(1.4), top + Inches(1.25), w - Inches(1.5), Inches(1.3))
        tf_det = tb_det.text_frame
        tf_det.word_wrap = True

        p_name = tf_det.paragraphs[0]
        p_name.text = "RAMESH BABU M."
        p_name.font.size = Pt(11)
        p_name.font.bold = True
        p_name.font.color.rgb = NAVY_PRIMARY

        p_id = tf_det.add_paragraph()
        p_id.text = "ID: MUD-000842"
        p_id.font.size = Pt(9)
        p_id.font.bold = True
        p_id.font.color.rgb = AMBER_ACCENT

        p_wing = tf_det.add_paragraph()
        p_wing.text = "Wing: Youth Team"
        p_wing.font.size = Pt(8.5)
        p_wing.font.bold = True
        p_wing.font.color.rgb = SLATE_DARK

        p_dist = tf_det.add_paragraph()
        p_dist.text = "Dist: Anantapur"
        p_dist.font.size = Pt(8)
        p_dist.font.color.rgb = SLATE_MUTED

        mid = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.15), top + Inches(2.7), w - Inches(0.3), Inches(1.6))
        mid.fill.solid()
        mid.fill.fore_color.rgb = SLATE_LIGHT
        mid.line.color.rgb = SLATE_BORDER
        mid.line.width = Pt(1)

        tb_mid = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(2.75), w - Inches(0.5), Inches(1.5))
        tf_mid = tb_mid.text_frame
        tf_mid.word_wrap = True

        items = [
            ("Constituency", "Kalyandurg"),
            ("Mandal", "Beluguppa"),
            ("Designation", "Active Community Member"),
            ("Valid Thru", "Lifelong (Registered 2026)"),
            ("Verification", "State Portal Verified ✓")
        ]
        for idx, (k, v) in enumerate(items):
            p = tf_mid.paragraphs[0] if idx == 0 else tf_mid.add_paragraph()
            p.font.size = Pt(8)
            run_k = p.add_run()
            run_k.text = k + ": "
            run_k.font.bold = True
            run_k.font.color.rgb = NAVY_DEEP
            run_v = p.add_run()
            run_v.text = v
            run_v.font.color.rgb = SLATE_DARK

        c_foot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + h - Inches(0.4), w, Inches(0.4))
        c_foot.fill.solid()
        c_foot.fill.fore_color.rgb = NAVY_PRIMARY
        c_foot.line.color.rgb = NAVY_PRIMARY

        tb_cf = slide.shapes.add_textbox(left, top + h - Inches(0.36), w, Inches(0.3))
        p_cf = tb_cf.text_frame.paragraphs[0]
        p_cf.alignment = PP_ALIGN.CENTER
        p_cf.text = "FRONT SIDE • OFFICIAL PVC SMARTCARD"
        p_cf.font.size = Pt(7.5)
        p_cf.font.bold = True
        p_cf.font.color.rgb = GOLD_LIGHT

    else:
        c_head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, Inches(0.7))
        c_head.fill.solid()
        c_head.fill.fore_color.rgb = NAVY_DARK
        c_head.line.color.rgb = NAVY_DARK

        tb_h = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.12), w - Inches(0.2), Inches(0.5))
        p1 = tb_h.text_frame.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = "RESIDENTIAL & VERIFICATION DETAILS"
        p1.font.bold = True
        p1.font.size = Pt(8.5)
        p1.font.color.rgb = GOLD_LIGHT

        qr_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.2), top + Inches(0.85), Inches(1.2), Inches(1.2))
        qr_box.fill.solid()
        qr_box.fill.fore_color.rgb = WHITE
        qr_box.line.color.rgb = NAVY_PRIMARY
        qr_box.line.width = Pt(1.5)

        tb_qr = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(1.2), Inches(1.2), Inches(0.4))
        p_qr = tb_qr.text_frame.paragraphs[0]
        p_qr.alignment = PP_ALIGN.CENTER
        p_qr.text = "[ QR CODE ]\nScan to Verify"
        p_qr.font.size = Pt(7.5)
        p_qr.font.bold = True
        p_qr.font.color.rgb = NAVY_PRIMARY

        tb_det = slide.shapes.add_textbox(left + Inches(1.5), top + Inches(0.85), w - Inches(1.6), Inches(1.3))
        tf_det = tb_det.text_frame
        tf_det.word_wrap = True

        p_bg = tf_det.paragraphs[0]
        p_bg.text = "Blood Group: O+ (Positive)"
        p_bg.font.size = Pt(8)
        p_bg.font.bold = True
        p_bg.font.color.rgb = RED_DANGER

        p_mob = tf_det.add_paragraph()
        p_mob.text = "Emergency: 98765 43210"
        p_mob.font.size = Pt(8)
        p_mob.font.bold = True
        p_mob.font.color.rgb = NAVY_PRIMARY

        p_addr = tf_det.add_paragraph()
        p_addr.text = "Address: D.No 4-82, Mudiraj Street, Beluguppa Village, Anantapur - 515761"
        p_addr.font.size = Pt(7.5)
        p_addr.font.color.rgb = SLATE_DARK

        inst_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.15), top + Inches(2.2), w - Inches(0.3), Inches(1.5))
        inst_box.fill.solid()
        inst_box.fill.fore_color.rgb = SLATE_LIGHT
        inst_box.line.color.rgb = SLATE_BORDER

        tb_inst = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(2.25), w - Inches(0.5), Inches(1.4))
        tf_inst = tb_inst.text_frame
        tf_inst.word_wrap = True
        p_i1 = tf_inst.paragraphs[0]
        p_i1.text = "• Official identity document issued by AP Mudiraj Mahasabha."
        p_i1.font.size = Pt(7.5)
        p_i1.font.color.rgb = SLATE_DARK

        p_i2 = tf_inst.add_paragraph()
        p_i2.text = "• In case of loss or inquiry, call State Helpline: 1800-425-MUDI."
        p_i2.font.size = Pt(7.5)
        p_i2.font.color.rgb = SLATE_DARK

        p_i3 = tf_inst.add_paragraph()
        p_i3.text = "• Online verification at https://mudirajcommunity.ap.gov.in/id-card"
        p_i3.font.size = Pt(7.5)
        p_i3.font.color.rgb = NAVY_PRIMARY

        tb_sig = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(3.9), w - Inches(0.4), Inches(0.4))
        p_sig = tb_sig.text_frame.paragraphs[0]
        p_sig.alignment = PP_ALIGN.RIGHT
        p_sig.text = "Authorized Signatory\n[ State President Seal ]"
        p_sig.font.size = Pt(7.5)
        p_sig.font.bold = True
        p_sig.font.color.rgb = NAVY_DEEP

        c_foot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + h - Inches(0.4), w, Inches(0.4))
        c_foot.fill.solid()
        c_foot.fill.fore_color.rgb = NAVY_PRIMARY
        c_foot.line.color.rgb = NAVY_PRIMARY

        tb_cf = slide.shapes.add_textbox(left, top + h - Inches(0.36), w, Inches(0.3))
        p_cf = tb_cf.text_frame.paragraphs[0]
        p_cf.alignment = PP_ALIGN.CENTER
        p_cf.text = "BACK SIDE • RESIDENTIAL & QR CODE"
        p_cf.font.size = Pt(7.5)
        p_cf.font.bold = True
        p_cf.font.color.rgb = GOLD_LIGHT

def draw_status_card(slide, left, top, width, height, status_title, badge_color, badge_bg, desc_items, icon_char="●"):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = SLATE_BORDER
    card.line.width = Pt(1.2)

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.08))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = badge_color
    stripe.line.color.rgb = badge_color

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.18), top + Inches(0.18), width - Inches(0.36), Inches(0.38))
    badge.fill.solid()
    badge.fill.fore_color.rgb = badge_bg
    badge.line.color.rgb = badge_color
    badge.line.width = Pt(1)

    tb_b = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.22), width - Inches(0.36), Inches(0.3))
    p_b = tb_b.text_frame.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = f"{icon_char}  {status_title}"
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = badge_color

    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.68), width - Inches(0.4), height - Inches(0.78))
    tf = tb.text_frame
    tf.word_wrap = True

    for idx, item in enumerate(desc_items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.font.size = Pt(8.8)
        p.font.color.rgb = SLATE_DARK
        p.space_after = Pt(2)
        if isinstance(item, tuple):
            rk = p.add_run()
            rk.text = "• " + item[0] + ": "
            rk.font.bold = True
            rk.font.color.rgb = badge_color
            rv = p.add_run()
            rv.text = item[1]
        else:
            p.text = "• " + item

def add_bottom_callout(slide, text, highlight_text=""):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(6.32), Inches(12.133), Inches(0.6))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(238, 242, 255)
    box.line.color.rgb = RGBColor(199, 210, 254)
    box.line.width = Pt(1)

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(6.38), Inches(11.7), Inches(0.48))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.font.size = Pt(9.5)
    if highlight_text:
        run_h = p.add_run()
        run_h.text = highlight_text + " "
        run_h.font.bold = True
        run_h.font.color.rgb = NAVY_PRIMARY
    run_t = p.add_run()
    run_t.text = text
    run_t.font.color.rgb = SLATE_DARK


# ==============================================================================
# BUILD PRESENTATION 1: BEAUTIFUL WEBSITE USER GUIDE & ALL FORMS MANUAL
# ==============================================================================
def build_beautiful_user_guide():
    prs = init_presentation()
    total_slides = 13
    doc_title = "Website User Guide & Form Manual"

    # Slide 1: Title Slide
    add_title_slide(
        prs,
        title="Andhra Pradesh Mudiraj Community Portal",
        subtitle="Complete Website & Interactive Forms User Manual",
        badge="Official Portal User Guide & Form Walkthrough",
        extra_notes=[
            "Step-by-Step Manual for Every Page, Form, Validation Rule & Digital Service",
            "State Jurisdiction: Andhra Pradesh | Serving 20 Lakh+ Community Members",
            "Includes Real Form UI Walkthrough: Membership Registration, Status Tracker & Smartcard"
        ],
        doc_label=doc_title
    )

    # Slide 2: Portal Navigation & Interactive Forms Map
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Website Pages & Interactive Forms Map", "PORTAL SITEMAP", "Visual layout of all public informational pages, interactive forms, and digital downloads")
    
    col_w = Inches(2.8)
    col_gap = Inches(0.24)
    top_pos = Inches(1.35)
    card_h = Inches(4.8)

    # Col 1: Public Pages
    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), top_pos, col_w, card_h)
    c1.fill.solid()
    c1.fill.fore_color.rgb = WHITE
    c1.line.color.rgb = SLATE_BORDER
    s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), top_pos, col_w, Inches(0.08))
    s1.fill.solid()
    s1.fill.fore_color.rgb = NAVY_PRIMARY
    tb1 = slide.shapes.add_textbox(Inches(0.75), top_pos + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "INFORMATIONAL PAGES"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p_t = tf1.add_paragraph()
    p_t.text = "Public Portals"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    pages = [
        ("Home ( / )", "Central state portal with statistics snapshot, mission & shortcuts."),
        ("About Us ( /about )", "Heritage, 20L+ members mandate, and 30-member committee tables."),
        ("Structure ( /organization )", "Interactive 5-tier tree explorer with unit counters & live search."),
        ("Leadership ( /leadership )", "Directory of office bearers across Main, Youth & Mahila wings."),
        ("Contact ( /contact )", "Vijayawada HQ & 6 district support helpdesks.")
    ]
    for k, v in pages:
        p_item = tf1.add_paragraph()
        p_item.font.size = Pt(8.5)
        p_item.space_before = Pt(4)
        rk = p_item.add_run()
        rk.text = "• " + k + "\n  "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p_item.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    # Col 2: Interactive Forms
    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + col_w + col_gap), top_pos, col_w, card_h)
    c2.fill.solid()
    c2.fill.fore_color.rgb = WHITE
    c2.line.color.rgb = SLATE_BORDER
    s2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + col_w + col_gap), top_pos, col_w, Inches(0.08))
    s2.fill.solid()
    s2.fill.fore_color.rgb = AMBER_ACCENT
    tb2 = slide.shapes.add_textbox(Inches(0.75 + col_w + col_gap), top_pos + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "INTERACTIVE FORMS"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    p_t = tf2.add_paragraph()
    p_t.text = "Forms to Fill"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    forms = [
        ("Form 1: Membership (/apply)", "Public registration: basic info, Aadhaar, photo & smart PIN lookup."),
        ("Form 2: Status Query (/membership)", "Lookup tool using Application No (APP-YYYY-XXXX) or Mobile."),
        ("Form 3: Digital ID Card (/id-card)", "Verification & rendering of official dual-sided PVC smartcards."),
        ("Form 4: Helpdesk Form (/contact)", "Grievance, inquiry & leadership nomination contact form.")
    ]
    for k, v in forms:
        p_item = tf2.add_paragraph()
        p_item.font.size = Pt(8.5)
        p_item.space_before = Pt(6)
        rk = p_item.add_run()
        rk.text = "• " + k + "\n  "
        rk.font.bold = True
        rk.font.color.rgb = AMBER_ACCENT
        rv = p_item.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    # Col 3: Smart Form Features
    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + (col_w + col_gap) * 2), top_pos, col_w, card_h)
    c3.fill.solid()
    c3.fill.fore_color.rgb = WHITE
    c3.line.color.rgb = SLATE_BORDER
    s3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + (col_w + col_gap) * 2), top_pos, col_w, Inches(0.08))
    s3.fill.solid()
    s3.fill.fore_color.rgb = GREEN_SUCCESS
    tb3 = slide.shapes.add_textbox(Inches(0.75 + (col_w + col_gap) * 2), top_pos + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = "FORM INTELLIGENCE"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = GREEN_SUCCESS

    p_t = tf3.add_paragraph()
    p_t.text = "Smart Validation"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    features = [
        ("Zero Account Hassle", "No login or password needed for citizens to submit applications."),
        ("Pincode Auto-Fill", "Typing a 6-digit Indian PIN queries postal data to auto-fill Mandal & District."),
        ("10-Digit Mobile Regex", "Client-side verification ensures valid Indian mobile format (^[6-9]\\d{9}$)."),
        ("12-Digit Aadhaar Check", "Aadhaar validation ensures genuine citizen representation."),
        ("Instant SMS Alert", "Immediate confirmation SMS sent with tracking ID upon submission.")
    ]
    for k, v in features:
        p_item = tf3.add_paragraph()
        p_item.font.size = Pt(8.5)
        p_item.space_before = Pt(4)
        rk = p_item.add_run()
        rk.text = "• " + k + "\n  "
        rk.font.bold = True
        rk.font.color.rgb = GREEN_SUCCESS
        rv = p_item.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    # Col 4: Guide Downloads
    c4 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + (col_w + col_gap) * 3), top_pos, col_w, card_h)
    c4.fill.solid()
    c4.fill.fore_color.rgb = WHITE
    c4.line.color.rgb = SLATE_BORDER
    s4 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + (col_w + col_gap) * 3), top_pos, col_w, Inches(0.08))
    s4.fill.solid()
    s4.fill.fore_color.rgb = BLUE_INFO
    tb4 = slide.shapes.add_textbox(Inches(0.75 + (col_w + col_gap) * 3), top_pos + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf4 = tb4.text_frame
    tf4.word_wrap = True
    p = tf4.paragraphs[0]
    p.text = "DOWNLOADABLE GUIDES"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = BLUE_INFO

    p_t = tf4.add_paragraph()
    p_t.text = "Navbar Downloads"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    dls = [
        ("Website Guide (.pptx)", "Complete visual manual of all pages & forms, downloadable directly in navbar."),
        ("Admin Guide (.pptx)", "Dedicated manual for administrators showing review, approval & leadership setup."),
        ("Top Banner Button", "1-click direct download button prominently featured in the top header banner."),
        ("Navbar Menu Dropdown", "Expanded dropdown showing quick links to all 8 pages with guide download.")
    ]
    for k, v in dls:
        p_item = tf4.add_paragraph()
        p_item.font.size = Pt(8.5)
        p_item.space_before = Pt(6)
        rk = p_item.add_run()
        rk.text = "• " + k + "\n  "
        rk.font.bold = True
        rk.font.color.rgb = BLUE_INFO
        rv = p_item.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    add_bottom_callout(slide, "Both the Public Website User Guide and Admin Guide PPTs are available for download at any time from the top navigation bar.", "Quick Access:");
    add_footer(slide, 2, total_slides, doc_title)

    # Slide 3: How to Navigate & Download PPTs in the Navbar
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "How to Navigate the Website & Download Guides in the Navbar", "NAVIGATION & DOWNLOADS", "Visual guide to the public header, top banner, menu links, and guide download dropdown")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(6.0), Inches(4.8), "Website Navigation Header", "https://mudiraj.ap.gov.in/")
    
    tb_b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.85), Inches(5.8), Inches(0.35))
    tb_b.fill.solid()
    tb_b.fill.fore_color.rgb = NAVY_DARK
    tb_b.line.fill.background()
    tb_bt = slide.shapes.add_textbox(Inches(0.75), Inches(1.88), Inches(5.7), Inches(0.3))
    p = tb_bt.text_frame.paragraphs[0]
    p.text = "Andhra Pradesh Mudiraj Mahasabha  |  📥 [ Download Website Guide (PPT) ]  |  Admin Login"
    p.font.size = Pt(7.5)
    p.font.color.rgb = GOLD_LIGHT

    mh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(2.25), Inches(5.8), Inches(0.6))
    mh.fill.solid()
    mh.fill.fore_color.rgb = WHITE
    mh.line.color.rgb = SLATE_BORDER
    mh.line.width = Pt(1)

    mht = slide.shapes.add_textbox(Inches(0.8), Inches(2.32), Inches(5.6), Inches(0.45))
    p = mht.text_frame.paragraphs[0]
    p.text = "🦁 MUDIRAJ PORTAL    Home   About   Structure   Leaders   Status   [ Website Guide (PPT) ▼ ]"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    dd = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.8), Inches(2.9), Inches(3.6), Inches(2.9))
    dd.fill.solid()
    dd.fill.fore_color.rgb = WHITE
    dd.line.color.rgb = AMBER_ACCENT
    dd.line.width = Pt(1.5)

    dd_t = slide.shapes.add_textbox(Inches(2.9), Inches(2.95), Inches(3.4), Inches(2.8))
    tf_dd = dd_t.text_frame
    tf_dd.word_wrap = True

    p = tf_dd.paragraphs[0]
    p.text = "📥 Download Website Guide (.pptx)\nAll Pages Manual • 13 Slides Presentation (572 KB)"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    dd_items = [
        "• Home Page (State Central Portal)",
        "• About Organization & 30-Member Teams",
        "• 5-Tier Structure Explorer (/organization)",
        "• Leadership Directory 3 Wings (/leadership)",
        "• Status Verification & Lookup (/membership)",
        "• Apply for Membership Form (/apply)",
        "• Digital Dual-Sided ID Card (/id-card)",
        "• Helpdesk & District Centers (/contact)"
    ]
    for item in dd_items:
        p = tf_dd.add_paragraph()
        p.text = item
        p.font.size = Pt(7.5)
        p.font.color.rgb = NAVY_PRIMARY
        p.space_before = Pt(1.5)

    add_info_table(
        slide, Inches(6.8), Inches(1.35), Inches(5.9), Inches(4.8),
        ["Navbar Element", "Action & How to Use", "Resulting Destination"],
        [
            ["Top Banner Download Button", "Click 'Download Website Guide (PPT)' in the top gold banner.", "Direct 1-click download of the complete PPTX presentation."],
            ["Website Guide Dropdown", "Click the 'Website Guide (PPT) ▼' dropdown button in the main menu.", "Opens interactive menu showing all 8 pages + direct download link."],
            ["Page Shortcuts in Dropdown", "Click any page listed inside the dropdown (e.g. Apply Form, ID Card).", "Instantly navigates directly to that specific page or form."],
            ["Admin Login Link", "Click 'Admin Login' on the top-right corner of the top banner.", "Opens the secure administrator authentication screen (/login)."],
            ["Mobile Friendly Navigation", "On smartphones, tap the hamburger icon (☰) to expand.", "Opens mobile drawer with all navigation links and download buttons."]
        ]
    )

    add_bottom_callout(slide, "Visitors can access and download this presentation from any page across the entire website via the persistent top navigation bar.", "Navbar Tip:");
    add_footer(slide, 3, total_slides, doc_title)

    # ==============================================================================
    # Slide 4: Form 1 - Membership Application Form ( /apply )
    # MATCHES EXACT USER SCREENSHOT
    # ==============================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        "Form 1: Membership Application ( /apply ) — Complete Form Walkthrough",
        "CITIZEN REGISTRATION • NO LOGIN REQUIRED",
        "Visual walkthrough of the real form: Basic Details, Aadhaar, Photo upload, and Smart Pincode auto-fetch"
    )

    if os.path.exists(APPLY_SCREENSHOT_PATH):
        frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.32), Inches(6.75), Inches(4.88))
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = NAVY_PRIMARY
        frame.line.width = Pt(1.5)

        ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.32), Inches(6.75), Inches(0.35))
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = NAVY_DARK
        ribbon.line.color.rgb = NAVY_DARK
        tb_r = slide.shapes.add_textbox(Inches(0.65), Inches(1.35), Inches(6.55), Inches(0.28))
        p_r = tb_r.text_frame.paragraphs[0]
        p_r.text = "🌐 REAL FORM SCREENSHOT • https://mudirajcommunity.ap.gov.in/apply"
        p_r.font.size = Pt(8.5)
        p_r.font.bold = True
        p_r.font.color.rgb = GOLD_LIGHT

        # Embed Actual Screenshot
        slide.shapes.add_picture(APPLY_SCREENSHOT_PATH, Inches(0.65), Inches(1.72), width=Inches(6.55), height=Inches(3.1))

        # Bottom Callout Card under image
        info_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(4.88), Inches(6.55), Inches(1.22))
        info_card.fill.solid()
        info_card.fill.fore_color.rgb = SLATE_LIGHT
        info_card.line.color.rgb = SLATE_BORDER
        info_card.line.width = Pt(1)

        tb_ic = slide.shapes.add_textbox(Inches(0.75), Inches(4.92), Inches(6.35), Inches(1.12))
        tf_ic = tb_ic.text_frame
        tf_ic.word_wrap = True

        p1 = tf_ic.paragraphs[0]
        p1.text = "FORM ARCHITECTURE & TWO-SECTION WORKFLOW:"
        p1.font.size = Pt(8.5)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY

        p2 = tf_ic.add_paragraph()
        p2.font.size = Pt(7.8)
        p2.space_before = Pt(2)
        r1 = p2.add_run()
        r1.text = "• Section 1 (Basic Details): "
        r1.font.bold = True
        r1.font.color.rgb = AMBER_ACCENT
        r2 = p2.add_run()
        r2.text = "Name, Gender, Blood Group, Occupation, +91 Mobile, Email, Photo & 12-digit Aadhaar.\n"
        r2.font.color.rgb = SLATE_DARK

        r3 = p2.add_run()
        r3.text = "• Section 2 (Address & Location): "
        r3.font.bold = True
        r3.font.color.rgb = GREEN_SUCCESS
        r4 = p2.add_run()
        r4.text = "6-digit Pincode (queries Postal API to suggest address) + Street / Colony name.\n"
        r4.font.color.rgb = SLATE_DARK

        r5 = p2.add_run()
        r5.text = "• Action Button: "
        r5.font.bold = True
        r5.font.color.rgb = NAVY_PRIMARY
        r6 = p2.add_run()
        r6.text = "Click '[ Review & Submit Application ]' to validate and send application for verification."
        r6.font.color.rgb = SLATE_DARK

    # Right: Comprehensive Bilingual Field Guide Table
    add_info_table(
        slide, Inches(7.45), Inches(1.32), Inches(5.3), Inches(4.88),
        ["Field (తెలుగు)", "Rules / Type", "Purpose & Instructions"],
        [
            ["Full Name (పూర్తి పేరు) *", "Text (min 3 chars)", "Legal name matching Aadhaar card. Printed on ID."],
            ["Gender (లింగం) *", "Male / Female / Other", "Females automatically qualify for Mahila Team wing."],
            ["Blood Group (రక్త వర్గం) *", "A+, B+, O+, AB+ (Rh +/-)", "Printed on PVC Smartcard for medical emergencies."],
            ["Occupation (వృత్తి) *", "Text box", "Profession (Agriculture, Business, Student, Employee)."],
            ["Mobile (+91) (మొబైల్) *", "10-Digit (^[6-9]\\d{9}$)", "Receives SMS with Tracking ID, Member ID & Password."],
            ["Email Address (ఇమెయిల్) *", "Valid email format", "For digital certificates, receipts & official notices."],
            ["Passport Photo *", "Upload button (< 2MB)", "Frontal face portrait (JPG/PNG). Printed on smartcard."],
            ["Aadhaar Number *", "12-Digit (Not 0 or 1)", "UIDAI format. Unique verified identifier in database."],
            ["Pincode (పిన్‌కోడ్) *", "6-Digit Postal PIN", "Auto-fetches Post Office, Mandal & District via API."],
            ["Street / Colony (వీధి) *", "Text (min 2 chars)", "Door / house number, street name, and landmark."],
            ["Review & Submit", "Primary Button", "Validates client inputs and queues application in Pending."]
        ]
    )

    add_bottom_callout(
        slide,
        "All fields with a red asterisk (*) are mandatory. Submitting immediately generates your Application Number (APP-2026-XXXXX) and dispatches an SMS confirmation.",
        "User Manual Tip:"
    )
    add_footer(slide, 4, total_slides, doc_title)

    # ==============================================================================
    # Slide 5: Form 1 (Part 2) - Application Submission Success & What Happens Next
    # ==============================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(
        slide,
        "Form 1 (Part 2): Submission Success Screen, Tracking ID & SMS Delivery",
        "POST-SUBMISSION CONFIRMATION",
        "What happens immediately after submitting the form: Instant confirmation, tracking reference & SMS"
    )

    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Application Submitted Successfully", "https://mudiraj.ap.gov.in/apply")

    sc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.85), Inches(5.0), Inches(4.15))
    sc.fill.solid()
    sc.fill.fore_color.rgb = WHITE
    sc.line.color.rgb = GREEN_SUCCESS
    sc.line.width = Pt(1.5)

    cir = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.05), Inches(1.95), Inches(0.65), Inches(0.65))
    cir.fill.solid()
    cir.fill.fore_color.rgb = GREEN_BG
    cir.line.color.rgb = GREEN_SUCCESS
    cir.line.width = Pt(1.5)
    tb_c = slide.shapes.add_textbox(Inches(3.05), Inches(2.05), Inches(0.65), Inches(0.45))
    p = tb_c.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "✓"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = GREEN_SUCCESS

    tb_st = slide.shapes.add_textbox(Inches(0.95), Inches(2.65), Inches(4.8), Inches(0.6))
    tf_st = tb_st.text_frame
    tf_st.word_wrap = True
    p = tf_st.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Membership Application Submitted Successfully!"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = GREEN_SUCCESS

    p2 = tf_st.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = "Your application is now in Pending status for Administrator verification."
    p2.font.size = Pt(8)
    p2.font.color.rgb = SLATE_MUTED

    det_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(3.3), Inches(4.6), Inches(1.4))
    det_box.fill.solid()
    det_box.fill.fore_color.rgb = SLATE_LIGHT
    det_box.line.color.rgb = SLATE_BORDER
    tb_det = slide.shapes.add_textbox(Inches(1.15), Inches(3.35), Inches(4.4), Inches(1.3))
    tf_det = tb_det.text_frame
    tf_det.word_wrap = True

    items = [
        ("Application Number", "APP-2026-89421  [SAVE THIS]"),
        ("Application Status", "Pending (Under Review)"),
        ("Applicant Name", "Ramesh Babu Mudiraj"),
        ("Occupation", "Agriculture & Business"),
        ("Registered Email", "name@gmail.com")
    ]
    for idx, (k, v) in enumerate(items):
        p = tf_det.paragraphs[0] if idx == 0 else tf_det.add_paragraph()
        p.font.size = Pt(8)
        rk = p.add_run()
        rk.text = k + ": "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_DARK

    sms_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(4.8), Inches(4.6), Inches(0.55))
    sms_box.fill.solid()
    sms_box.fill.fore_color.rgb = GOLD_BG
    sms_box.line.color.rgb = AMBER_ACCENT

    tb_sms = slide.shapes.add_textbox(Inches(1.15), Inches(4.85), Inches(4.4), Inches(0.45))
    tf_sms = tb_sms.text_frame
    tf_sms.word_wrap = True
    p = tf_sms.paragraphs[0]
    p.text = "📱 Password Generation Formula upon Approval:\nFirst 3 letters of last name + First 3 digits of mobile (sent via SMS)"
    p.font.size = Pt(7.5)
    p.font.color.rgb = NAVY_DEEP

    draw_button(slide, Inches(1.05), Inches(5.45), Inches(4.6), Inches(0.4), "TRACK APPLICATION STATUS NOW (/membership) →", NAVY_PRIMARY)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Stage Number", "Stage Name", "Responsible Authority", "Expected Timeframe & Action"],
        [
            ["Stage 1", "Form Submission & Tracking ID", "Portal Engine (Automated)", "Instant. Unique reference APP-YYYY-XXXX generated and SMS dispatched."],
            ["Stage 2", "Pending District Verification", "District Review Coordinator", "1 to 2 Business Days. Checks address authenticity and Aadhaar record."],
            ["Stage 3", "Membership Approval & ID Issuance", "State Admin Panel", "Generates permanent Member ID (MUD-00000001) and auto-generated password."],
            ["Stage 4", "Digital ID Card Generation", "Citizen Portal (/id-card)", "Immediate upon approval. High-resolution dual-sided card becomes printable."]
        ]
    )

    add_bottom_callout(
        slide,
        "Save your Application Number (APP-2026-XXXXX). You can check status at any time on /membership using either this number or your mobile number.",
        "Next Steps:"
    )
    add_footer(slide, 5, total_slides, doc_title)

    # Slide 6: Form 2 - Membership Status Lookup & Verification Portal
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Form 2: Membership Status Lookup & Verification Portal ( /membership )", "APPLICATION TRACKING TOOL", "How to check application progress and understand the 4 status badges")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(12.133), Inches(1.4), "Membership Status Verification Form", "https://mudiraj.ap.gov.in/membership")
    
    draw_input_field(slide, Inches(0.9), Inches(1.85), Inches(7.5), "Enter Application Tracking ID OR 10-Digit Mobile Number", "e.g. APP-2026-89421  or  9876543210", is_required=True)
    draw_button(slide, Inches(8.6), Inches(2.05), Inches(3.8), Inches(0.42), "CHECK APPLICATION STATUS →", AMBER_ACCENT)

    card_w = Inches(2.8)
    gap = Inches(0.24)
    top_pos = Inches(2.95)
    h_pos = Inches(3.25)

    draw_status_card(
        slide, Inches(0.6), top_pos, card_w, h_pos,
        "STATUS 1: PENDING", AMBER_ACCENT, GOLD_BG,
        [
            ("Badge Color", "Amber Gold (Clock Icon)"),
            ("Meaning", "Application submitted and queued for verification."),
            ("Next Step", "District coordinator reviews demographic data."),
            ("Timeline", "Typically takes 24 to 48 hours."),
            ("User Action", "No action required. Wait for SMS update.")
        ],
        "⏳"
    )

    draw_status_card(
        slide, Inches(0.6 + card_w + gap), top_pos, card_w, h_pos,
        "STATUS 2: APPROVED", GREEN_SUCCESS, GREEN_BG,
        [
            ("Badge Color", "Emerald Green (Check Icon)"),
            ("Meaning", "Membership confirmed! Permanent Member ID issued."),
            ("Member ID", "Assigned format: MUD-00000001."),
            ("Rights", "Official community member & voting eligibility."),
            ("User Action", "Click 'Download ID Card' to view smartcard.")
        ],
        "✓"
    )

    draw_status_card(
        slide, Inches(0.6 + (card_w + gap) * 2), top_pos, card_w, h_pos,
        "STATUS 3: REJECTED", RED_DANGER, RED_BG,
        [
            ("Badge Color", "Rose Red (Alert Icon)"),
            ("Meaning", "Application could not be verified by reviewer."),
            ("Reasons", "Unclear photo, invalid address, or outside AP."),
            ("Correction", "Reviewer notes explain exact reason for correction."),
            ("User Action", "Click 'Re-Apply' with corrected details.")
        ],
        "✕"
    )

    draw_status_card(
        slide, Inches(0.6 + (card_w + gap) * 3), top_pos, card_w, h_pos,
        "STATUS 4: CARD READY", BLUE_INFO, BLUE_BG,
        [
            ("Badge Color", "Royal Blue (Smartcard Icon)"),
            ("Meaning", "Digital dual-sided PVC card has been generated."),
            ("Features", "Front photo & ID + Back QR code & address."),
            ("Formats", "High-res PDF print format or PNG image."),
            ("User Action", "Print on standard CR-80 PVC plastic card.")
        ],
        "🪪"
    )

    add_bottom_callout(slide, "You can look up status using either your Application ID (APP-2026-XXXXX) or the 10-digit mobile number entered during registration.", "Lookup Tip:");
    add_footer(slide, 6, total_slides, doc_title)

    # Slide 7: Form 3 - Digital Dual-Sided ID Card Generator (/id-card)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Form 3: Digital Dual-Sided ID Card Portal ( /id-card )", "PVC SMARTCARD GENERATOR", "Visual breakdown of the front and back of the official Mudiraj Membership Smartcard")
    
    draw_id_card(slide, Inches(0.6), Inches(1.35), is_front=True)
    draw_id_card(slide, Inches(4.2), Inches(1.35), is_front=False)

    add_info_table(
        slide, Inches(7.8), Inches(1.35), Inches(4.9), Inches(4.8),
        ["Smartcard Feature", "Specification & Details", "Verification Method"],
        [
            ["Standard Card Size", "CR-80 PVC (85.6 mm x 54 mm), 30 mil thickness.", "Standard ATM / Driving License card size."],
            ["Front Side Elements", "Logo, Member Photo, Member ID, Full Name, Wing, District.", "Visual inspection with official AP state emblem."],
            ["Back Side Elements", "Residential Address, Blood Group, Emergency Phone, Signature.", "Provides vital medical & emergency contact data."],
            ["QR Verification Code", "Unique encoded URL pointing to https://mudiraj.ap.gov.in/verify/...", "Any smartphone camera scans QR code to verify live validity."],
            ["High-Res PDF Download", "Vector-sharp 300 DPI PDF with front and back layout.", "Ready for direct sublimation PVC thermal card printers."],
            ["Digital Wallet Ready", "Save as PNG image directly to smartphone gallery.", "Convenient paperless identification for community events."]
        ]
    )

    add_bottom_callout(slide, "Members can download their high-resolution PVC Smartcard PDF at any time from /id-card using their Member ID or mobile number.", "Card Portal:");
    add_footer(slide, 7, total_slides, doc_title)

    # Slide 8: Form 4 - Public Contact & Helpdesk Form (/contact)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Form 4: Helpdesk & Grievance Contact Form ( /contact )", "PUBLIC INQUIRY & GRIEVANCE FORM", "Connecting community members with State Headquarters and District Support Centers")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Citizen Helpdesk & Inquiry Form", "https://mudiraj.ap.gov.in/contact")
    
    draw_input_field(slide, Inches(0.85), Inches(1.9), Inches(4.95), "Inquiry Category *", "Membership Registration Assistance", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(2.55), Inches(4.95), "Your Full Name *", "Ramesh Babu Mudiraj", is_required=True)
    
    draw_input_field(slide, Inches(0.85), Inches(3.2), Inches(2.35), "Mobile Number *", "9876543210", is_required=True)
    draw_input_field(slide, Inches(3.45), Inches(3.2), Inches(2.35), "Your District *", "Anantapur", is_required=True, is_select=True)

    tb_msg_lbl = slide.shapes.add_textbox(Inches(0.85), Inches(3.85), Inches(4.95), Inches(0.24))
    p = tb_msg_lbl.text_frame.paragraphs[0]
    p.text = "Your Detailed Message / Grievance *"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    msg_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(4.1), Inches(4.95), Inches(0.95))
    msg_box.fill.solid()
    msg_box.fill.fore_color.rgb = WHITE
    msg_box.line.color.rgb = SLATE_BORDER
    msg_box.line.width = Pt(1)

    tb_mtxt = slide.shapes.add_textbox(Inches(0.95), Inches(4.15), Inches(4.75), Inches(0.85))
    tf_mtxt = tb_mtxt.text_frame
    tf_mtxt.word_wrap = True
    p = tf_mtxt.paragraphs[0]
    p.text = "Need help updating my address details on my submitted application APP-2026-89421. Please assist."
    p.font.size = Pt(8.5)
    p.font.color.rgb = SLATE_DARK

    draw_button(slide, Inches(0.85), Inches(5.2), Inches(4.95), Inches(0.42), "SEND MESSAGE TO DISTRICT HELPDESK →", AMBER_ACCENT)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Inquiry Category", "Routing Destination", "Expected Response Time", "Escalation Level"],
        [
            ["Membership Assistance", "District Coordinator Queue", "Within 24 Hours", "State Membership Registrar"],
            ["Leadership Nominations", "State Executive Committee", "Within 3 Business Days", "State President Office"],
            ["Grievance / Dispute", "Legal & Community Welfare Cell", "Within 48 Hours", "Apex Grievance Tribunal"],
            ["Technical / Website Bug", "IT Systems & Webmaster", "Within 12 Hours", "Chief Technology Officer"],
            ["State Headquarters", "Vijayawada Central Office", "Direct Phone: 1800-425-MUDI", "Mon - Sat (9:00 AM - 6:00 PM)"]
        ]
    )

    add_bottom_callout(slide, "Urgent queries can also be answered via the State Toll-Free Community Helpline: 1800-425-MUDI (Toll Free).", "Helpline:");
    add_footer(slide, 8, total_slides, doc_title)

    # Slide 9: Form 5 - Admin Login Gateway Form (/login)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Form 5: Admin Login Gateway Form ( /login )", "ADMINISTRATOR AUTHENTICATION", "Secure authentication screen for district coordinators, verifiers, and state executives")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Admin Security Gateway", "https://mudiraj.ap.gov.in/login")
    
    ac = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(1.9), Inches(4.5), Inches(4.1))
    ac.fill.solid()
    ac.fill.fore_color.rgb = WHITE
    ac.line.color.rgb = NAVY_PRIMARY
    ac.line.width = Pt(1.5)

    tb_ab = slide.shapes.add_textbox(Inches(1.2), Inches(2.05), Inches(4.3), Inches(0.6))
    tf_ab = tb_ab.text_frame
    p = tf_ab.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "🔒 ADMIN CONTROL PANEL"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p2 = tf_ab.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = "Authorized Personnel Only • IP Monitored"
    p2.font.size = Pt(7.5)
    p2.font.color.rgb = SLATE_MUTED

    draw_input_field(slide, Inches(1.35), Inches(2.75), Inches(4.0), "Username / Official Email *", "admin@mudiraj.org", is_required=True)
    draw_input_field(slide, Inches(1.35), Inches(3.5), Inches(4.0), "Password *", "••••••••••••", is_required=True, badge="ENCRYPTED")
    
    tb_rem = slide.shapes.add_textbox(Inches(1.35), Inches(4.25), Inches(4.0), Inches(0.3))
    p = tb_rem.text_frame.paragraphs[0]
    p.text = "☑ Keep session active on this workstation"
    p.font.size = Pt(8)
    p.font.color.rgb = SLATE_DARK

    draw_button(slide, Inches(1.35), Inches(4.7), Inches(4.0), Inches(0.44), "SECURE SIGN IN TO ADMIN DASHBOARD 🔒", NAVY_PRIMARY)

    tb_dl = slide.shapes.add_textbox(Inches(1.35), Inches(5.25), Inches(4.0), Inches(0.4))
    p = tb_dl.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "📥 [ Download Admin Features Guide (PPT) ]"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Administrative Role", "Assigned Jurisdiction", "Accessible Admin Features", "Security Standard"],
        [
            ["State Super Admin", "All 26 AP Districts Statewide", "Full database access, hierarchy setup, leadership appointments, reports.", "Two-Factor Auth + Audit Logging"],
            ["District Coordinator", "Assigned District (e.g. Guntur)", "Review & approve pending applications, verify local identity.", "District-scoped role permissions"],
            ["Wing Reviewer", "Main / Youth / Mahila Wing", "Manage wing-specific portfolios, events, and committees.", "Wing-specific view & edit rights"],
            ["Helpdesk Operator", "Grievance Management Cell", "Respond to citizen inquiries and status support tickets.", "Read-only member record access"]
        ]
    )

    add_bottom_callout(slide, "Admins can also download the dedicated Admin Features Guide (PPT) directly from the admin dashboard and topbar.", "Admin Access:");
    add_footer(slide, 9, total_slides, doc_title)

    # Slide 10: Page Guide - Organization Tree Explorer (/organization)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page Guide: 5-Tier Organization Tree Explorer ( /organization )", "HIERARCHY EXPLORER", "Navigating the interactive structural breakdown from State to Village Gramam levels")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Organization Hierarchy Explorer", "https://mudiraj.ap.gov.in/organization")
    
    draw_input_field(slide, Inches(0.85), Inches(1.85), Inches(4.95), "Search Any District, Constituency or Mandal", "e.g. Anantapur, Kalyandurg, Beluguppa...", is_required=False, badge="LIVE FILTER")

    node_top = Inches(2.65)
    nodes = [
        ("🏛️ TIER 1: ANDHRA PRADESH STATE CENTRAL (APEX)", NAVY_DARK, GOLD_LIGHT, Inches(0)),
        ("  ↳ 🏢 TIER 2: Anantapur District Committee (26 Total)", NAVY_PRIMARY, WHITE, Inches(0.2)),
        ("    ↳ 🗳️ TIER 3: Kalyandurg Assembly Constituency (175 Total)", AMBER_ACCENT, WHITE, Inches(0.4)),
        ("      ↳ 📍 TIER 4: Beluguppa Mandal Unit (679 Total)", GREEN_SUCCESS, WHITE, Inches(0.6)),
        ("        ↳ 🏡 TIER 5: Beluguppa Gramam / Village (Grassroots)", SLATE_DARK, WHITE, Inches(0.8))
    ]
    for idx, (n_text, bg_col, txt_col, indent) in enumerate(nodes):
        nb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85 + indent), node_top + Inches(idx * 0.52), Inches(4.95 - indent), Inches(0.42))
        nb.fill.solid()
        nb.fill.fore_color.rgb = bg_col
        nb.line.color.rgb = bg_col
        tb = slide.shapes.add_textbox(Inches(0.95 + indent), node_top + Inches(idx * 0.52 + 0.08), Inches(4.8 - indent), Inches(0.3))
        p = tb.text_frame.paragraphs[0]
        p.text = n_text
        p.font.size = Pt(8)
        p.font.bold = True
        p.font.color.rgb = txt_col

    draw_button(slide, Inches(0.85), Inches(5.35), Inches(4.95), Inches(0.38), "VIEW COMMITTEE LEADERS FOR THIS UNIT →", NAVY_PRIMARY)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Hierarchy Tier", "Count in AP", "Governance Role", "Leadership Wings Present"],
        [
            ["Tier 1: State", "1 State Apex", "Statewide policy, government lobbying, apex welfare decisions.", "Main Team (30), Youth Team (30), Mahila Team (30)"],
            ["Tier 2: District", "26 Districts", "District-level administration, coordination with district collectorate.", "Main Team (30), Youth Team (30), Mahila Team (30)"],
            ["Tier 3: Assembly", "175 Assemblies", "Legislative constituency representation and civic advocacy.", "Main Team (30), Youth Team (30), Mahila Team (30)"],
            ["Tier 4: Mandal", "679 Mandals", "Town/block welfare, local dispute resolution, youth job drives.", "Main Team (30), Youth Team (30), Mahila Team (30)"],
            ["Tier 5: Gramam", "Village Units", "Doorstep citizen engagement, local festivals, member enrollment.", "Grassroots Village Advisory Panels"]
        ]
    )

    add_bottom_callout(slide, "Click on any unit node in the tree to immediately open its designated Main, Youth, and Mahila leadership committees.", "Tree Feature:");
    add_footer(slide, 10, total_slides, doc_title)

    # Slide 11: Page Guide - Leadership Directory (/leadership)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page Guide: Appointed Leadership Directory ( /leadership )", "OFFICE BEARERS DIRECTORY", "Connecting with verified office bearers across Main, Youth, and Mahila Wings")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Leadership Directory", "https://mudiraj.ap.gov.in/leadership")
    
    draw_input_field(slide, Inches(0.85), Inches(1.85), Inches(2.35), "Filter by Level", "State Level", is_required=False, is_select=True)
    draw_input_field(slide, Inches(3.45), Inches(1.85), Inches(2.35), "Filter by Wing", "Youth Wing", is_required=False, is_select=True)

    lc1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(2.6), Inches(4.95), Inches(1.2))
    lc1.fill.solid()
    lc1.fill.fore_color.rgb = WHITE
    lc1.line.color.rgb = SLATE_BORDER

    ph1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(2.7), Inches(1.0), Inches(1.0))
    ph1.fill.solid()
    ph1.fill.fore_color.rgb = RGBColor(241, 245, 249)
    ph1.line.color.rgb = NAVY_PRIMARY
    tb_ph1 = slide.shapes.add_textbox(Inches(0.95), Inches(3.05), Inches(1.0), Inches(0.3))
    p = tb_ph1.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "PHOTO"
    p.font.size = Pt(7.5)
    p.font.bold = True
    p.font.color.rgb = SLATE_MUTED

    tb_ld1 = slide.shapes.add_textbox(Inches(2.1), Inches(2.68), Inches(3.6), Inches(1.05))
    tf_ld1 = tb_ld1.text_frame
    p = tf_ld1.paragraphs[0]
    p.text = "SRI K. CHINNA MUDIRAJ"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p_d = tf_ld1.add_paragraph()
    p_d.text = "President — Youth Team (State Level)"
    p_d.font.size = Pt(8.5)
    p_d.font.bold = True
    p_d.font.color.rgb = AMBER_ACCENT

    p_sub = tf_ld1.add_paragraph()
    p_sub.text = "Verified ID: MUD-000001  |  Location: AP Central"
    p_sub.font.size = Pt(7.5)
    p_sub.font.color.rgb = SLATE_MUTED

    lc2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(3.95), Inches(4.95), Inches(1.2))
    lc2.fill.solid()
    lc2.fill.fore_color.rgb = WHITE
    lc2.line.color.rgb = SLATE_BORDER

    ph2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(4.05), Inches(1.0), Inches(1.0))
    ph2.fill.solid()
    ph2.fill.fore_color.rgb = RGBColor(241, 245, 249)
    ph2.line.color.rgb = NAVY_PRIMARY
    tb_ph2 = slide.shapes.add_textbox(Inches(0.95), Inches(4.4), Inches(1.0), Inches(0.3))
    p = tb_ph2.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "PHOTO"
    p.font.size = Pt(7.5)
    p.font.bold = True
    p.font.color.rgb = SLATE_MUTED

    tb_ld2 = slide.shapes.add_textbox(Inches(2.1), Inches(4.03), Inches(3.6), Inches(1.05))
    tf_ld2 = tb_ld2.text_frame
    p = tf_ld2.paragraphs[0]
    p.text = "SMT. T. LAKSHMI DEVI MUDIRAJ"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p_d = tf_ld2.add_paragraph()
    p_d.text = "General Secretary — Mahila Team (State Level)"
    p_d.font.size = Pt(8.5)
    p_d.font.bold = True
    p_d.font.color.rgb = GREEN_SUCCESS

    p_sub = tf_ld2.add_paragraph()
    p_sub.text = "Verified ID: MUD-000004  |  Location: AP Central"
    p_sub.font.size = Pt(7.5)
    p_sub.font.color.rgb = SLATE_MUTED

    draw_button(slide, Inches(0.85), Inches(5.3), Inches(4.95), Inches(0.4), "VIEW FULL 30-MEMBER EXECUTIVE COMMITTEE →", NAVY_PRIMARY)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Wing / Committee", "Core Focus & Mandate", "Key Designations Included", "Number of Leaders"],
        [
            ["Main Team Wing", "Core socio-political representation, legal advocacy, government schemes.", "President, 6 Vice Presidents, 2 Gen Secs, 6 Joint Secs, Treasurer.", "~30 Office Bearers per Level"],
            ["Youth Team Wing", "Education assistance, competitive exam guidance, IT skill drives, sports.", "Youth President, Youth Working President, Zonal Coordinators.", "~30 Youth Leaders per Level"],
            ["Mahila Team Wing", "Women empowerment, micro-enterprise support, girl child education.", "Mahila President, Working President, District Coordinators.", "~30 Women Leaders per Level"],
            ["Total Appointed Scale", "3 Wings x 30 Positions = 90 Leaders per administrative unit.", "State (90) + 26 Districts (2,340) + 175 Assemblies (15,750).", "Over 20,000+ Appointed Leaders"]
        ]
    )

    add_bottom_callout(slide, "Every leader displayed in the directory has an official verified Member ID confirming authentic representation.", "Leadership Integrity:");
    add_footer(slide, 11, total_slides, doc_title)

    # Slide 12: Applicant Troubleshooting & Form FAQs
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Applicant Troubleshooting & Frequently Asked Questions (FAQs)", "SUPPORT & COMMON ISSUES", "Quick solutions to the most common questions and validation errors when filling forms")
    
    card_w = Inches(5.8)
    card_h = Inches(2.25)
    gap_x = Inches(0.533)
    gap_y = Inches(0.25)

    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), card_w, card_h)
    c1.fill.solid()
    c1.fill.fore_color.rgb = WHITE
    c1.line.color.rgb = SLATE_BORDER
    tb1 = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), card_w - Inches(0.4), card_h - Inches(0.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "❓ ISSUE 1: 'Mobile Number Already Registered'"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT
    items = [
        ("Cause", "An application has already been submitted using this 10-digit mobile number."),
        ("Solution", "Go to /membership to check the existing application status instead of re-applying."),
        ("Correction", "If you need to update incorrect details, contact your District Helpdesk.")
    ]
    for k, v in items:
        p = tf1.add_paragraph()
        p.font.size = Pt(8.5)
        rk = p.add_run()
        rk.text = "• " + k + ": "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_DARK

    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + card_w + gap_x), Inches(1.4), card_w, card_h)
    c2.fill.solid()
    c2.fill.fore_color.rgb = WHITE
    c2.line.color.rgb = SLATE_BORDER
    tb2 = slide.shapes.add_textbox(Inches(0.8 + card_w + gap_x), Inches(1.5), card_w - Inches(0.4), card_h - Inches(0.2))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "❓ ISSUE 2: 'Aadhaar Number Validation Error'"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT
    items = [
        ("Cause", "Aadhaar must be exactly 12 digits and cannot begin with 0 or 1."),
        ("Solution", "Double-check your 12-digit UIDAI card number and re-enter without spaces."),
        ("Security", "Aadhaar numbers are encrypted and never shown publicly on ID cards.")
    ]
    for k, v in items:
        p = tf2.add_paragraph()
        p.font.size = Pt(8.5)
        rk = p.add_run()
        rk.text = "• " + k + ": "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_DARK

    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4 + card_h + gap_y), card_w, card_h)
    c3.fill.solid()
    c3.fill.fore_color.rgb = WHITE
    c3.line.color.rgb = SLATE_BORDER
    tb3 = slide.shapes.add_textbox(Inches(0.8), Inches(1.5 + card_h + gap_y), card_w - Inches(0.4), card_h - Inches(0.2))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = "❓ ISSUE 3: 'Photo Upload Fails or Takes Too Long'"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT
    items = [
        ("Cause", "The image file exceeds the 2 MB file size limit or is in an unsupported format."),
        ("Solution", "Compress the photo using standard mobile gallery tools before uploading."),
        ("Format", "Use JPG, JPEG, or PNG format with a clean frontal portrait view.")
    ]
    for k, v in items:
        p = tf3.add_paragraph()
        p.font.size = Pt(8.5)
        rk = p.add_run()
        rk.text = "• " + k + ": "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_DARK

    c4 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + card_w + gap_x), Inches(1.4 + card_h + gap_y), card_w, card_h)
    c4.fill.solid()
    c4.fill.fore_color.rgb = WHITE
    c4.line.color.rgb = SLATE_BORDER
    tb4 = slide.shapes.add_textbox(Inches(0.8 + card_w + gap_x), Inches(1.5 + card_h + gap_y), card_w - Inches(0.4), card_h - Inches(0.2))
    tf4 = tb4.text_frame
    tf4.word_wrap = True
    p = tf4.paragraphs[0]
    p.text = "❓ ISSUE 4: 'Lost My Application Tracking ID'"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT
    items = [
        ("Recovery", "You DO NOT need your Application ID to track your membership."),
        ("Solution", "Simply type your registered 10-digit mobile number on /membership."),
        ("SMS History", "Search your mobile SMS messages for sender 'MDSABH' or 'MUDIRAJ'.")
    ]
    for k, v in items:
        p = tf4.add_paragraph()
        p.font.size = Pt(8.5)
        rk = p.add_run()
        rk.text = "• " + k + ": "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_DARK

    add_bottom_callout(slide, "For any unresolved issues, contact the State Helpline at 1800-425-MUDI or email support@mudiraj.org.", "Assistance:");
    add_footer(slide, 12, total_slides, doc_title)

    # Slide 13: Executive Summary & Support Directory
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Executive Summary & State Support Directory", "DIRECTORY & HELPLINE", "Official contact centers, state headquarters, and downloadable resources")
    
    col3_w = Inches(3.8)
    col3_gap = Inches(0.36)

    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), col3_w, Inches(4.7))
    c1.fill.solid()
    c1.fill.fore_color.rgb = WHITE
    c1.line.color.rgb = SLATE_BORDER
    s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.4), col3_w, Inches(0.08))
    s1.fill.solid()
    s1.fill.fore_color.rgb = NAVY_PRIMARY

    tb1 = slide.shapes.add_textbox(Inches(0.8), Inches(1.55), col3_w - Inches(0.4), Inches(4.4))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "STATE HEADQUARTERS"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p_t = tf1.add_paragraph()
    p_t.text = "Vijayawada Central Apex"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    items = [
        ("Address", "Mudiraj Bhavan, MG Road, Governorpet, Vijayawada, AP - 520002"),
        ("Toll-Free Helpline", "1800-425-MUDI (Toll Free Across AP)"),
        ("Direct Phone", "+91 866 2489000 / 2489001"),
        ("Official Email", "portal@mudiraj.org / support@mudiraj.org"),
        ("Working Hours", "Monday to Saturday: 9:00 AM – 6:00 PM IST"),
        ("State Jurisdiction", "Apex authority over all 26 District Mahasabhas.")
    ]
    for k, v in items:
        p = tf1.add_paragraph()
        p.font.size = Pt(8.5)
        p.space_before = Pt(4)
        rk = p.add_run()
        rk.text = "• " + k + ":\n  "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + col3_w + col3_gap), Inches(1.4), col3_w, Inches(4.7))
    c2.fill.solid()
    c2.fill.fore_color.rgb = WHITE
    c2.line.color.rgb = SLATE_BORDER
    s2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + col3_w + col3_gap), Inches(1.4), col3_w, Inches(0.08))
    s2.fill.solid()
    s2.fill.fore_color.rgb = AMBER_ACCENT

    tb2 = slide.shapes.add_textbox(Inches(0.8 + col3_w + col3_gap), Inches(1.55), col3_w - Inches(0.4), Inches(4.4))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "REGIONAL SUPPORT HUBS"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    p_t = tf2.add_paragraph()
    p_t.text = "Zonal Verification Centers"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    zones = [
        ("Rayalaseema Zone", "Anantapur, Kurnool, Kadapa, Chittoor, Nandyal, Tirupati centers."),
        ("Coastal Andhra Zone", "Krishna, Guntur, Bapatla, Palnadu, Prakasam, Nellore centers."),
        ("North Coastal Zone", "Visakhapatnam, Vizianagaram, Srikakulam, Anakapalli centers."),
        ("Godavari Zone", "East Godavari, West Godavari, Kakinada, Konaseema centers."),
        ("Verification Role", "Physical address audit & member dispute arbitration.")
    ]
    for k, v in zones:
        p = tf2.add_paragraph()
        p.font.size = Pt(8.5)
        p.space_before = Pt(4)
        rk = p.add_run()
        rk.text = "• " + k + ":\n  "
        rk.font.bold = True
        rk.font.color.rgb = AMBER_ACCENT
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + (col3_w + col3_gap) * 2), Inches(1.4), col3_w, Inches(4.7))
    c3.fill.solid()
    c3.fill.fore_color.rgb = WHITE
    c3.line.color.rgb = SLATE_BORDER
    s3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + (col3_w + col3_gap) * 2), Inches(1.4), col3_w, Inches(0.08))
    s3.fill.solid()
    s3.fill.fore_color.rgb = GREEN_SUCCESS

    tb3 = slide.shapes.add_textbox(Inches(0.8 + (col3_w + col3_gap) * 2), Inches(1.55), col3_w - Inches(0.4), Inches(4.4))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = "DIGITAL RESOURCES"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = GREEN_SUCCESS

    p_t = tf3.add_paragraph()
    p_t.text = "Quick Access Links"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    links = [
        ("Online Application Form", "https://mudiraj.ap.gov.in/apply"),
        ("Application Status Tracker", "https://mudiraj.ap.gov.in/membership"),
        ("Digital ID Card Generator", "https://mudiraj.ap.gov.in/id-card"),
        ("Leadership Directory", "https://mudiraj.ap.gov.in/leadership"),
        ("Structure Explorer", "https://mudiraj.ap.gov.in/organization"),
        ("Download Website Guide", "Top Banner & Navbar Dropdown (.pptx)"),
        ("Download Admin Guide", "Admin Topbar & Dashboard (.pptx)")
    ]
    for k, v in links:
        p = tf3.add_paragraph()
        p.font.size = Pt(8.5)
        p.space_before = Pt(3.5)
        rk = p.add_run()
        rk.text = "• " + k + ":\n  "
        rk.font.bold = True
        rk.font.color.rgb = GREEN_SUCCESS
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    add_bottom_callout(slide, "Thank you for using the Andhra Pradesh Mudiraj Community Portal — Empowering 20 Lakh+ community members through digital governance.", "Community Mission:");
    add_footer(slide, 13, total_slides, doc_title)

    out_path = os.path.join(DOWNLOADS_DIR, "Mudiraj_Website_User_Guide.pptx")
    prs.save(out_path)
    root_path = os.path.join(PUBLIC_DIR, "Mudiraj_Website_User_Guide.pptx")
    prs.save(root_path)
    print(f"Generated Beautiful Website User Guide: {out_path} ({total_slides} slides)")


# ==============================================================================
# BUILD PRESENTATION 2: BEAUTIFUL ADMIN FEATURES GUIDE & OPERATIONS MANUAL
# ==============================================================================
def build_beautiful_admin_guide():
    prs = init_presentation()
    total_slides = 12
    doc_title = "Admin Panel Features & Management Manual"

    # Slide 1: Title Slide
    add_title_slide(
        prs,
        title="Admin Control Panel & Operations Manual",
        subtitle="Complete Administrator Guide to Member Management & Hierarchy Setup",
        badge="Authorized Personnel & Executive Guide",
        extra_notes=[
            "Administrative Workflows: Member Registration, Application Review & Leadership Setup",
            "Covers: Direct Member Add Modal, Status Approval Engine, 30-Member Committees & Reports",
            "State Jurisdiction: Andhra Pradesh Central Administration | Serving 26 Districts"
        ],
        doc_label=doc_title
    )

    # Slide 2: Admin Architecture & Modules Overview
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Control Panel Architecture & Core Modules", "SYSTEM OVERVIEW", "Master layout of administrative functions, role scopes, and database management modules")
    
    card_w = Inches(2.8)
    gap = Inches(0.24)
    top_pos = Inches(1.4)
    card_h = Inches(4.75)

    draw_status_card(
        slide, Inches(0.6), top_pos, card_w, card_h,
        "Module 1: Member Hub", NAVY_PRIMARY, BLUE_BG,
        [
            ("Page URL", "/admin/members"),
            ("Direct Add Modal", "Register pre-verified citizens directly into active database."),
            ("Member Table", "Search, filter by district/mandal, sort, and edit member profiles."),
            ("ID Override", "Assign or regenerate unique lifelong Member IDs (MUD-XXXXX)."),
            ("Status Control", "Toggle status: Active, Suspended, or Honorary Member.")
        ],
        "👥"
    )

    draw_status_card(
        slide, Inches(0.6 + card_w + gap), top_pos, card_w, card_h,
        "Module 2: Review Engine", AMBER_ACCENT, GOLD_BG,
        [
            ("Page URL", "/admin/applications"),
            ("Pending Queue", "Review submitted public applications (APP-YYYY-XXXX)."),
            ("Demographic Audit", "Inspect name, father's name, DOB, and address against proof."),
            ("One-Click Approval", "Generates Member ID and dispatches approval SMS."),
            ("Rejection with Notes", "Sends structured rejection reason back to applicant.")
        ],
        "📋"
    )

    draw_status_card(
        slide, Inches(0.6 + (card_w + gap) * 2), top_pos, card_w, card_h,
        "Module 3: Leadership Hub", GREEN_SUCCESS, GREEN_BG,
        [
            ("Page URL", "/admin/leaders & /admin/teams"),
            ("3 Dedicated Wings", "Main Team (Core), Youth Team, and Mahila Team."),
            ("30 Positions", "President, 6 VPs, 2 Gen Secs, 6 Joint Secs, Treasurer, Committee."),
            ("5 Administrative Tiers", "Appoint leaders at State, District, Constituency, and Mandal."),
            ("Leader Directory", "Publishes verified leaders to public directory immediately.")
        ],
        "👑"
    )

    draw_status_card(
        slide, Inches(0.6 + (card_w + gap) * 3), top_pos, card_w, card_h,
        "Module 4: Hierarchy & Reports", PURPLE_ACCENT, PURPLE_BG,
        [
            ("Page URL", "/admin/organizations & /admin/reports"),
            ("5-Tier Hierarchy", "Configure 26 AP Districts, 175 Assemblies, and Mandals."),
            ("NRI Diaspora Hub", "Manage international chapters across USA, Gulf, UK, Australia."),
            ("Excel Data Export", "Export filtered member records with complete contact fields."),
            ("Topbar Manual DL", "Direct button in topbar to download this Admin Guide PPT.")
        ],
        "📊"
    )

    add_bottom_callout(slide, "All admin modules feature role-based permissions, automated audit trails, and 1-click Excel reporting.", "System Security:");
    add_footer(slide, 2, total_slides, doc_title)

    # Slide 3: Admin Form 1 - Authentication & Login Gateway
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Form 1: Security Gateway & Authentication ( /login )", "LOGIN & ACCESS CONTROL", "Session security, multi-role credential verification, and brute-force protection")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Admin Security Gateway", "https://mudiraj.ap.gov.in/login")
    
    ac = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(1.85), Inches(4.5), Inches(4.1))
    ac.fill.solid()
    ac.fill.fore_color.rgb = WHITE
    ac.line.color.rgb = NAVY_PRIMARY
    ac.line.width = Pt(1.5)

    tb_ab = slide.shapes.add_textbox(Inches(1.2), Inches(1.95), Inches(4.3), Inches(0.6))
    tf_ab = tb_ab.text_frame
    p = tf_ab.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "🔒 STATE ADMIN CONTROL GATEWAY"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p2 = tf_ab.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = "Session Encryption • Role Permission Verification"
    p2.font.size = Pt(7.5)
    p2.font.color.rgb = SLATE_MUTED

    draw_input_field(slide, Inches(1.35), Inches(2.65), Inches(4.0), "Official Administrator Email *", "admin@mudiraj.org", is_required=True)
    draw_input_field(slide, Inches(1.35), Inches(3.35), Inches(4.0), "Security Password *", "••••••••••••", is_required=True, badge="ENCRYPTED")
    
    tb_rem = slide.shapes.add_textbox(Inches(1.35), Inches(4.1), Inches(4.0), Inches(0.3))
    p = tb_rem.text_frame.paragraphs[0]
    p.text = "☑ Maintain secure session on this workstation"
    p.font.size = Pt(8)
    p.font.color.rgb = SLATE_DARK

    draw_button(slide, Inches(1.35), Inches(4.55), Inches(4.0), Inches(0.44), "SECURE ADMIN SIGN IN 🔒", NAVY_PRIMARY)

    tb_dl = slide.shapes.add_textbox(Inches(1.35), Inches(5.1), Inches(4.0), Inches(0.4))
    p = tb_dl.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "📥 [ Download Admin Features Guide (PPT) ]"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Administrative Role", "Assigned Scope", "Access Privileges", "Security Protocols"],
        [
            ["Super Administrator", "Entire AP State (All 26 Districts)", "Unrestricted read/write across all modules, user role creation, database backup.", "Hardware 2FA + Static IP Binding"],
            ["District Review Coordinator", "Specific Assigned District", "Review & approve pending membership applications for assigned district.", "District-scoped role permissions"],
            ["Wing State Secretary", "Main, Youth, or Mahila Wing", "Manage committee appointments, leader designations, and wing events.", "Wing-specific edit privileges"],
            ["Grievance Officer", "Citizen Contact Desk", "Review inquiries submitted via /contact, route tickets, dispatch replies.", "Read-only member records access"]
        ]
    )

    add_bottom_callout(slide, "Admin sessions automatically terminate after 30 minutes of inactivity to safeguard community census data.", "Security Standard:");
    add_footer(slide, 3, total_slides, doc_title)

    # Slide 4: Admin Form 2 - Direct Member Registration Modal (/admin/members)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Form 2: Direct Member Registration Modal ( /admin/members )", "DIRECT MEMBER ENROLLMENT", "Registering verified citizens directly into the active database without queue delays")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Direct Member Enrollment Modal", "https://mudiraj.ap.gov.in/admin/members")
    
    mh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(1.85), Inches(4.95), Inches(0.35))
    mh.fill.solid()
    mh.fill.fore_color.rgb = NAVY_PRIMARY
    mht = slide.shapes.add_textbox(Inches(0.95), Inches(1.88), Inches(4.75), Inches(0.3))
    p = mht.text_frame.paragraphs[0]
    p.text = "👤 DIRECT MEMBER ENROLLMENT FORM  [X Close]"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = WHITE

    draw_input_field(slide, Inches(0.85), Inches(2.25), Inches(4.95), "Full Name (Aadhaar Verified) *", "e.g. Chinna Mudiraj", is_required=True)
    
    draw_input_field(slide, Inches(0.85), Inches(2.9), Inches(2.35), "Mobile Number *", "9876543210", is_required=True)
    draw_input_field(slide, Inches(3.45), Inches(2.9), Inches(2.35), "Wing Category *", "Youth Team", is_required=True, is_select=True)

    draw_input_field(slide, Inches(0.85), Inches(3.55), Inches(2.35), "District *", "Anantapur", is_required=True, is_select=True)
    draw_input_field(slide, Inches(3.45), Inches(3.55), Inches(2.35), "Mandal *", "Beluguppa", is_required=True, is_select=True)

    draw_input_field(slide, Inches(0.85), Inches(4.2), Inches(2.35), "Member ID (Auto/Custom)", "MUD-000843", is_required=True, badge="AUTO")
    draw_input_field(slide, Inches(3.45), Inches(4.2), Inches(2.35), "Membership Status", "Active Member", is_required=True, is_select=True)

    draw_button(slide, Inches(0.85), Inches(4.9), Inches(4.95), Inches(0.42), "SAVE MEMBER & GENERATE DIGITAL ID CARD →", AMBER_ACCENT)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Form Field", "Input Control", "Validation Rule", "Administrative Purpose"],
        [
            ["Full Name", "Text Box", "Mandatory (*). Letters & spaces.", "Records verified name into MongoDB master member database."],
            ["Mobile Number", "10-Digit Numeric", "Unique constraint in DB.", "Serves as primary identifier and SMS receipt recipient."],
            ["Wing Category", "Select Dropdown", "Main / Youth / Mahila.", "Assigns member to their leadership pool and committee quota."],
            ["District & Mandal", "Cascading Select", "Valid AP administrative unit.", "Locks member into local voter and leadership jurisdiction."],
            ["Member ID", "Auto-Generated", "Prefix MUD- with 6-digit sequence.", "Immediate generation bypassing 'Pending' queue."],
            ["Status Override", "Dropdown", "Active / Suspended / Honorary.", "Allows direct granting of active voting rights."]
        ]
    )

    add_bottom_callout(slide, "The Direct Registration Modal is ideal for enrolling local VIPs, senior patrons, and office bearers at regional conferences.", "Admin Utility:");
    add_footer(slide, 4, total_slides, doc_title)

    # Slide 5: Admin Form 3 - Application Review & Decision Modal (/admin/applications)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Form 3: Application Review & Decision Modal ( /admin/applications )", "APPLICATION REVIEW ENGINE", "Evaluating pending submissions, inspecting proofs, and executing Approval or Rejection")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Application Review Decision Modal", "https://mudiraj.ap.gov.in/admin/applications")
    
    rb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(1.85), Inches(4.95), Inches(0.4))
    rb.fill.solid()
    rb.fill.fore_color.rgb = NAVY_PRIMARY
    rbt = slide.shapes.add_textbox(Inches(0.95), Inches(1.9), Inches(4.75), Inches(0.3))
    p = rbt.text_frame.paragraphs[0]
    p.text = "📋 REVIEW APPLICATION: APP-2026-89421"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = GOLD_LIGHT

    det = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(2.35), Inches(4.95), Inches(1.5))
    det.fill.solid()
    det.fill.fore_color.rgb = SLATE_LIGHT
    det.line.color.rgb = SLATE_BORDER
    tb_d = slide.shapes.add_textbox(Inches(0.95), Inches(2.4), Inches(4.75), Inches(1.4))
    tf_d = tb_d.text_frame
    items = [
        ("Applicant Name", "Ramesh Babu Mudiraj (Age 32, Male)"),
        ("Father's Name", "Venkataswamy Mudiraj"),
        ("Contact Mobile", "+91 98765 43210  (SMS Ready ✓)"),
        ("Location", "Beluguppa Mandal, Kalyandurg, Anantapur - 515761"),
        ("Requested Wing", "Youth Team (Eligible based on age 32)")
    ]
    for idx, (k, v) in enumerate(items):
        p = tf_d.paragraphs[0] if idx == 0 else tf_d.add_paragraph()
        p.font.size = Pt(7.8)
        rk = p.add_run()
        rk.text = k + ": "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_DARK

    draw_input_field(slide, Inches(0.85), Inches(3.95), Inches(4.95), "Assign Official Permanent Member ID *", "MUD-000842 (Auto-Generated Next in Sequence)", is_required=True)

    draw_button(slide, Inches(0.85), Inches(4.75), Inches(2.35), Inches(0.46), "✓ APPROVE APPLICATION", GREEN_SUCCESS)
    draw_button(slide, Inches(3.45), Inches(4.75), Inches(2.35), Inches(0.46), "✕ REJECT (WITH REASON)", RED_DANGER)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Verification Step", "Coordinator Checklist", "Approval Trigger", "Rejection Reason Options"],
        [
            ["Demographic Check", "Verify spelling of applicant and father's name matches proof.", "Names clear and verified.", "Unclear identity proof or mismatched name."],
            ["Jurisdiction Check", "Confirm applicant resides within the selected AP district & mandal.", "Mandal verified inside district boundary.", "Address outside Andhra Pradesh state."],
            ["Wing Verification", "Validate age criteria (< 35 for Youth, Female for Mahila).", "Age criteria satisfied.", "Over-age for Youth Team (transferred to Main)."],
            ["Duplicate Detection", "Database checks for existing mobile number or Aadhaar hash.", "Zero existing duplicate found.", "Duplicate application already approved."],
            ["Post-Action Dispatch", "Generates permanent Member ID and dispatches confirmation SMS.", "Instant SMS with Member ID.", "SMS explains reason with link to re-apply."]
        ]
    )

    add_bottom_callout(slide, "Approving an application automatically generates their permanent Member ID and makes their digital PVC card available instantly.", "Review Workflow:");
    add_footer(slide, 5, total_slides, doc_title)

    # Slide 6: Admin Form 4 - 30-Member Leadership & Committee Assignment
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Form 4: 30-Member Committee Leadership Setup ( /admin/leaders )", "COMMITTEE APPOINTMENT ENGINE", "Appointing and structuring the 30-member executive committees across all wings and tiers")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Leadership Appointment Form", "https://mudiraj.ap.gov.in/admin/leaders")
    
    draw_input_field(slide, Inches(0.85), Inches(1.85), Inches(2.35), "1. Committee Tier *", "District Level", is_required=True, is_select=True)
    draw_input_field(slide, Inches(3.45), Inches(1.85), Inches(2.35), "2. Wing Committee *", "Youth Team (30 Positions)", is_required=True, is_select=True)

    draw_input_field(slide, Inches(0.85), Inches(2.55), Inches(4.95), "3. Select District Jurisdiction *", "Anantapur District (AP)", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(3.25), Inches(4.95), "4. Leadership Designation Portfolio *", "President — Youth Team (1 Sanctioned Position)", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(3.95), Inches(4.95), "5. Select Verified Member from Database *", "Search by Member ID or Name (MUD-000001)...", is_required=True, badge="AUTO-LINK")

    draw_button(slide, Inches(0.85), Inches(4.75), Inches(4.95), Inches(0.46), "APPOINT LEADER & PUBLISH TO DIRECTORY →", AMBER_ACCENT)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Executive Designation", "Quota", "Core Portfolio Responsibility", "Published Scope"],
        [
            ["President / Chairman", "1 Post", "Apex executive head of the committee unit.", "State / District / Mandal Directory"],
            ["Working President", "1 Post", "Direct executive deputy assisting president.", "Public Leadership Directory"],
            ["Vice Presidents", "6 Posts", "Zonal and regional coordination across subdivisions.", "Regional Public Cards"],
            ["General Secretaries", "2 Posts", "Administrative records, correspondence, and meetings.", "State & District Hubs"],
            ["Joint Secretaries", "6 Posts", "Program execution and grassroots mobilization.", "Subdivision Directory"],
            ["Treasurer", "1 Post", "Financial accounting, audit records, welfare fund.", "State / District Accounts"],
            ["Executive Members", "13 Posts", "Advisory council and grassroots representation.", "Full Committee Roster"]
        ]
    )

    add_bottom_callout(slide, "Every administrative unit supports 3 wings x 30 positions = 90 total leadership positions per administrative unit.", "Committee Scale:");
    add_footer(slide, 6, total_slides, doc_title)

    # Slide 7: Admin Form 5 - 5-Tier Organization Hierarchy Structure Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Form 5: 5-Tier Organization Hierarchy Management ( /admin/organizations )", "HIERARCHY CONFIGURATOR", "Adding, organizing, and maintaining Districts, Assembly Constitutions, Mandals, and Gramams")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Organization Hierarchy Form", "https://mudiraj.ap.gov.in/admin/organizations")
    
    draw_input_field(slide, Inches(0.85), Inches(1.85), Inches(4.95), "Select Hierarchy Tier to Add / Edit *", "Tier 4: Mandal / Revenue Block", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(2.55), Inches(4.95), "Parent District (26 AP Districts) *", "Anantapur District", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(3.25), Inches(4.95), "Parent Assembly Constituency *", "Kalyandurg Assembly (175 Total)", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(3.95), Inches(4.95), "New Mandal Name *", "e.g. Beluguppa Mandal", is_required=True)

    draw_button(slide, Inches(0.85), Inches(4.75), Inches(4.95), Inches(0.46), "SAVE HIERARCHY UNIT & UPDATE TREE →", NAVY_PRIMARY)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Administrative Tier", "Current Count", "Administrative Scope", "Cascading Impact"],
        [
            ["Tier 1: State", "1 Apex State", "Andhra Pradesh State Central Executive.", "Governs all statewide committees and census."],
            ["Tier 2: District", "26 Districts", "26 official restructured districts of Andhra Pradesh.", "Routes member applications to district reviewers."],
            ["Tier 3: Assembly", "175 Assemblies", "175 legislative assembly constituencies.", "Enables legislative constituency level representation."],
            ["Tier 4: Mandal", "679 Mandals", "Revenue mandals and urban municipal corporations.", "Forms grassroots towns & block operational units."],
            ["Tier 5: Gramam", "Village Units", "Panchayat villages and municipal wards.", "Enables village doorstep membership identification."]
        ]
    )

    add_bottom_callout(slide, "Adding or reorganizing a unit instantly updates the interactive tree on /organization and the public registration dropdowns on /apply.", "Real-Time Sync:");
    add_footer(slide, 7, total_slides, doc_title)

    # Slide 8: Admin Form 6 - NRI Global Chapter Management Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Form 6: Global NRI Chapter Management ( /admin/nri-teams )", "DIASPORA CHAPTER REGISTRATION", "Connecting and organizing overseas Mudiraj community chapters across the globe")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Global Chapter Registration", "https://mudiraj.ap.gov.in/admin/nri-teams")
    
    draw_input_field(slide, Inches(0.85), Inches(1.85), Inches(4.95), "Overseas Country *", "United States of America (USA)", is_required=True, is_select=True)
    draw_input_field(slide, Inches(0.85), Inches(2.55), Inches(4.95), "Chapter Name / Region *", "e.g. Mudiraj Association of North America (MANA)", is_required=True)
    
    draw_input_field(slide, Inches(0.85), Inches(3.25), Inches(2.35), "Coordinator Name *", "Dr. S. K. Mudiraj", is_required=True)
    draw_input_field(slide, Inches(3.45), Inches(3.25), Inches(2.35), "Overseas Phone *", "+1 (408) 555-0192", is_required=True)

    draw_input_field(slide, Inches(0.85), Inches(3.95), Inches(4.95), "Official Chapter Email *", "mana@mudiraj.org", is_required=True)

    draw_button(slide, Inches(0.85), Inches(4.75), Inches(4.95), Inches(0.46), "REGISTER GLOBAL CHAPTER & APPOINT LIAISON →", AMBER_ACCENT)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Global Region", "Active Diaspora Hubs", "Key Mandate & Activities", "Liaison Coordination"],
        [
            ["North America", "USA (Dallas, Bay Area, New Jersey), Canada", "Higher education scholarships, career mentorship, diaspora conventions.", "State NRI Affairs Secretary"],
            ["Gulf Countries", "UAE (Dubai), Saudi Arabia, Kuwait, Qatar", "Expatriate worker welfare, emergency legal aid, repatriation support.", "Overseas Labor Welfare Cell"],
            ["United Kingdom & Europe", "UK (London), Germany, Netherlands", "Student exchange programs, tech community networks, cultural meets.", "European Chapter Convener"],
            ["Australia & Asia-Pacific", "Australia (Sydney, Melbourne), Singapore", "Diaspora business investments, community entrepreneurship funds.", "Trade & Investment Liaison"]
        ]
    )

    add_bottom_callout(slide, "NRI chapters enable overseas members to maintain permanent cultural ties and support grassroots education initiatives in Andhra Pradesh.", "Global Reach:");
    add_footer(slide, 8, total_slides, doc_title)

    # Slide 9: Admin Reports & Multi-Format Data Export
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Module 7: Multi-Format Reports & Excel Data Export ( /admin/reports )", "REPORTS & EXCEL EXPORT", "Filtering, auditing, and exporting complete member records for electoral and census use")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Census Reports Generator", "https://mudiraj.ap.gov.in/admin/reports")
    
    draw_input_field(slide, Inches(0.85), Inches(1.85), Inches(2.35), "Select District Filter", "All 26 AP Districts", is_required=False, is_select=True)
    draw_input_field(slide, Inches(3.45), Inches(1.85), Inches(2.35), "Select Wing Filter", "Youth Team Wing", is_required=False, is_select=True)

    draw_input_field(slide, Inches(0.85), Inches(2.55), Inches(2.35), "Membership Status", "Approved Active Members", is_required=False, is_select=True)
    draw_input_field(slide, Inches(3.45), Inches(2.55), Inches(2.35), "Date Range", "Last 30 Days (Oct 2026)", is_required=False, is_select=True)

    rc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(3.25), Inches(4.95), Inches(1.2))
    rc.fill.solid()
    rc.fill.fore_color.rgb = SLATE_LIGHT
    rc.line.color.rgb = SLATE_BORDER
    tb_rc = slide.shapes.add_textbox(Inches(0.95), Inches(3.35), Inches(4.75), Inches(1.0))
    tf_rc = tb_rc.text_frame
    p = tf_rc.paragraphs[0]
    p.text = "📊 MATCHING RECORDS FOUND: 14,820 MEMBERS"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p2 = tf_rc.add_paragraph()
    p2.text = "Includes: Member ID, Full Name, Father's Name, Phone, District, Mandal, Wing."
    p2.font.size = Pt(8)
    p2.font.color.rgb = SLATE_MUTED

    draw_button(slide, Inches(0.85), Inches(4.65), Inches(4.95), Inches(0.48), "📥 EXPORT COMPLETE DATASET TO EXCEL (.XLSX) →", GREEN_SUCCESS)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Export Format", "Included Columns", "Recommended Use", "Security & Auditing"],
        [
            ["Excel (.xlsx) Multi-Tab", "Member ID, Name, Father's Name, Mobile, Blood Group, Address, Wing, Status.", "Electoral list preparation, district census, meeting attendance.", "Export action logged with admin username and timestamp."],
            ["CSV Format (Comma Delimited)", "Raw tab-delimited records for bulk SMS dispatch.", "SMS blast campaigns for state conventions and announcements.", "Requires Super Admin clearance for full phone exports."],
            ["PDF Membership Roll", "Printable formatted voter rolls sorted alphabetically by mandal.", "Physical voter verification at district leadership elections.", "Includes official digital watermark and QR verification."],
            ["Statistical Summary", "Bar charts and breakdown tables by district and age bracket.", "Executive committee presentations and strategic welfare planning.", "Available in real-time on dashboard."]
        ]
    )

    add_bottom_callout(slide, "Reports are generated directly using xlsx-js-style with professional table headers, frozen top rows, and alternating row colors.", "Excel Engine:");
    add_footer(slide, 9, total_slides, doc_title)

    # Slide 10: Admin Security, Role-Based Access & Audit Logging
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Security, Role-Based Access & Audit Logging", "CYBERSECURITY & AUDITING", "Protecting community census data with multi-tier access control and immutable audit logs")
    
    card_w = Inches(2.8)
    gap = Inches(0.24)
    top_pos = Inches(1.4)
    card_h = Inches(4.75)

    draw_status_card(
        slide, Inches(0.6), top_pos, card_w, card_h,
        "Pillar 1: Data Encryption", NAVY_PRIMARY, BLUE_BG,
        [
            ("HTTPS / TLS 1.3", "All browser-to-server traffic encrypted with SSL certificates."),
            ("Password Hashing", "Bcrypt / Argon2 cryptographic password salting and hashing."),
            ("PII Protection", "Aadhaar and personal identity hashes stored securely."),
            ("No Plaintext Passwords", "Zero plain-text credential storage across database.")
        ],
        "🔒"
    )

    draw_status_card(
        slide, Inches(0.6 + card_w + gap), top_pos, card_w, card_h,
        "Pillar 2: Role Access (RBAC)", AMBER_ACCENT, GOLD_BG,
        [
            ("Granular Roles", "Super Admin, District Coordinator, Wing Reviewer, Helpdesk."),
            ("District Scope", "District coordinators can only view/edit their assigned district."),
            ("Principle of Least Privilege", "Sensitive export buttons hidden from basic viewers."),
            ("Session Tokens", "Secure HTTP-only cookies with 30-minute timeout.")
        ],
        "🛡️"
    )

    draw_status_card(
        slide, Inches(0.6 + (card_w + gap) * 2), top_pos, card_w, card_h,
        "Pillar 3: Immutable Audit Log", GREEN_SUCCESS, GREEN_BG,
        [
            ("Audit Trail", "Every record modification, approval, or deletion is recorded."),
            ("Recorded Fields", "Admin User ID, Timestamp, Action Type, Previous & New Value."),
            ("Export Logging", "Bulk Excel downloads generate security alert to Super Admin."),
            ("Tamper Proof", "Audit logs cannot be edited or erased by normal admins.")
        ],
        "📝"
    )

    draw_status_card(
        slide, Inches(0.6 + (card_w + gap) * 3), top_pos, card_w, card_h,
        "Pillar 4: Disaster Recovery", PURPLE_ACCENT, PURPLE_BG,
        [
            ("Automated Backups", "Daily automated encrypted snapshots of MongoDB database."),
            ("Redundant Storage", "Geographically separated secondary backup nodes."),
            ("Quick Rollback", "One-click restoration to previous hourly checkpoint."),
            ("99.9% Uptime SLA", "Engineered for high availability on cloud infrastructure.")
        ],
        "💾"
    )

    add_bottom_callout(slide, "Strict compliance with data privacy standards ensures complete confidentiality of member contact information and family records.", "Compliance:");
    add_footer(slide, 10, total_slides, doc_title)

    # Slide 11: Topbar Quick Actions & PPT Guide Download
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Topbar Quick Actions & PPT Guide Download in Navbar", "ADMIN NAVBAR & SHORTCUTS", "Direct access to downloadable administrator guides and operational topbar utilities")
    
    draw_mockup_window(slide, Inches(0.6), Inches(1.35), Inches(5.5), Inches(4.8), "Admin Shell Navigation Topbar", "https://mudiraj.ap.gov.in/admin")
    
    tb_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(1.85), Inches(4.95), Inches(0.65))
    tb_h.fill.solid()
    tb_h.fill.fore_color.rgb = WHITE
    tb_h.line.color.rgb = SLATE_BORDER
    tb_h.line.width = Pt(1)

    tb_ht = slide.shapes.add_textbox(Inches(0.95), Inches(1.92), Inches(4.75), Inches(0.5))
    p = tb_ht.text_frame.paragraphs[0]
    p.text = "🛡️ ADMIN CONSOLE   [ 📥 Download Admin Guide (PPT) ]   👤 Admin"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    sb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(2.6), Inches(1.5), Inches(3.2))
    sb.fill.solid()
    sb.fill.fore_color.rgb = NAVY_DARK
    sb.line.color.rgb = NAVY_DARK

    tb_sb = slide.shapes.add_textbox(Inches(0.9), Inches(2.65), Inches(1.4), Inches(3.1))
    tf_sb = tb_sb.text_frame
    sb_links = [
        "📊 Dashboard",
        "👥 Members",
        "📋 Applications",
        "👑 Leaders",
        "🏛️ Organization",
        "🌐 NRI Teams",
        "📈 Reports",
        "📥 Admin Guide"
    ]
    for idx, l in enumerate(sb_links):
        p = tf_sb.paragraphs[0] if idx == 0 else tf_sb.add_paragraph()
        p.text = l
        p.font.size = Pt(7.5)
        p.font.color.rgb = GOLD_LIGHT if idx == 7 else WHITE
        p.font.bold = (idx == 7)
        p.space_before = Pt(3)

    cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.45), Inches(2.6), Inches(3.35), Inches(3.2))
    cb.fill.solid()
    cb.fill.fore_color.rgb = WHITE
    cb.line.color.rgb = SLATE_BORDER

    tb_cb = slide.shapes.add_textbox(Inches(2.55), Inches(2.7), Inches(3.15), Inches(3.0))
    tf_cb = tb_cb.text_frame
    tf_cb.word_wrap = True
    p = tf_cb.paragraphs[0]
    p.text = "📥 Download Admin Features Guide"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    p2 = tf_cb.add_paragraph()
    p2.text = "Click to download the official 12-slide PowerPoint presentation covering all admin management features."
    p2.font.size = Pt(7.8)
    p2.font.color.rgb = SLATE_DARK
    p2.space_before = Pt(3)

    draw_button(slide, Inches(2.55), Inches(4.2), Inches(3.15), Inches(0.4), "DOWNLOAD ADMIN PPT (544 KB) →", AMBER_ACCENT)

    add_info_table(
        slide, Inches(6.4), Inches(1.35), Inches(6.3), Inches(4.8),
        ["Topbar Element", "Location & Trigger", "Administrative Action", "Target Users"],
        [
            ["Admin Guide PPT Button", "Top Navigation Bar (Right Side)", "1-click download of the complete Admin Guide PowerPoint (.pptx).", "All authorized coordinators"],
            ["Sidebar Admin Guide Link", "Sidebar Navigation Menu", "Direct download shortcut without navigating to other pages.", "District reviewers"],
            ["Dashboard Action Button", "Admin Overview Dashboard", "Prominently featured download button alongside registration actions.", "New administrators"],
            ["Public Website Guide Link", "Switch to Citizen Portal (/)", "Inspect the citizen-facing website guide from the member view.", "Support desk staff"],
            ["Session Logout", "User Profile Menu (Top Right)", "Terminates encrypted session token and returns to /login.", "All active admins"]
        ]
    )

    add_bottom_callout(slide, "The Admin Guide presentation is always accessible in the topbar of every admin screen for quick reference during operations.", "Navbar Guide:");
    add_footer(slide, 11, total_slides, doc_title)

    # Slide 12: Admin Operational Checklist & Escalation Protocol
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Operational Checklist & Escalation Protocol", "OPERATIONS & SLA", "Daily operating standards, verification procedures, and emergency contact channels")
    
    col3_w = Inches(3.8)
    col3_gap = Inches(0.36)

    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.4), col3_w, Inches(4.7))
    c1.fill.solid()
    c1.fill.fore_color.rgb = WHITE
    c1.line.color.rgb = SLATE_BORDER
    s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.4), col3_w, Inches(0.08))
    s1.fill.solid()
    s1.fill.fore_color.rgb = NAVY_PRIMARY

    tb1 = slide.shapes.add_textbox(Inches(0.8), Inches(1.55), col3_w - Inches(0.4), Inches(4.4))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "DAILY OPERATIONAL CHECKLIST"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    p_t = tf1.add_paragraph()
    p_t.text = "District Review Routine"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    tasks = [
        ("Step 1: Check Pending Queue", "Log in daily at 9:00 AM to inspect newly submitted applications."),
        ("Step 2: Scrutinize Photos", "Ensure applicant photo is upright, clear, and has a clean background."),
        ("Step 3: Verify Jurisdiction", "Confirm selected Mandal matches the applicant's residential address."),
        ("Step 4: Execute Approvals", "Approve valid records to generate lifelong Member IDs."),
        ("Step 5: Provide Rejection Notes", "If rejecting, write clear feedback explaining what to correct.")
    ]
    for k, v in tasks:
        p = tf1.add_paragraph()
        p.font.size = Pt(8.5)
        p.space_before = Pt(4)
        rk = p.add_run()
        rk.text = "• " + k + ":\n  "
        rk.font.bold = True
        rk.font.color.rgb = NAVY_PRIMARY
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + col3_w + col3_gap), Inches(1.4), col3_w, Inches(4.7))
    c2.fill.solid()
    c2.fill.fore_color.rgb = WHITE
    c2.line.color.rgb = SLATE_BORDER
    s2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + col3_w + col3_gap), Inches(1.4), col3_w, Inches(0.08))
    s2.fill.solid()
    s2.fill.fore_color.rgb = AMBER_ACCENT

    tb2 = slide.shapes.add_textbox(Inches(0.8 + col3_w + col3_gap), Inches(1.55), col3_w - Inches(0.4), Inches(4.4))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "SERVICE LEVEL AGREEMENTS (SLA)"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = AMBER_ACCENT

    p_t = tf2.add_paragraph()
    p_t.text = "Turnaround Targets"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    slas = [
        ("Application Review", "Maximum 48 business hours from public submission."),
        ("ID Card Generation", "Instantaneous upon approval (0-second delay)."),
        ("Citizen Helpdesk Queries", "Response dispatched within 24 hours via phone or SMS."),
        ("Committee Appointments", "Leadership directory updated immediately upon save."),
        ("Data Export Audits", "Excel audit reports archived for 12 months.")
    ]
    for k, v in slas:
        p = tf2.add_paragraph()
        p.font.size = Pt(8.5)
        p.space_before = Pt(4)
        rk = p.add_run()
        rk.text = "• " + k + ":\n  "
        rk.font.bold = True
        rk.font.color.rgb = AMBER_ACCENT
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6 + (col3_w + col3_gap) * 2), Inches(1.4), col3_w, Inches(4.7))
    c3.fill.solid()
    c3.fill.fore_color.rgb = WHITE
    c3.line.color.rgb = SLATE_BORDER
    s3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6 + (col3_w + col3_gap) * 2), Inches(1.4), col3_w, Inches(0.08))
    s3.fill.solid()
    s3.fill.fore_color.rgb = GREEN_SUCCESS

    tb3 = slide.shapes.add_textbox(Inches(0.8 + (col3_w + col3_gap) * 2), Inches(1.55), col3_w - Inches(0.4), Inches(4.4))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = "ADMINISTRATIVE ESCALATION"
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = GREEN_SUCCESS

    p_t = tf3.add_paragraph()
    p_t.text = "State Support Desk"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    contacts = [
        ("IT Operations Desk", "it@mudiraj.org | Ext. 104"),
        ("Database Admin (DBA)", "dba@mudiraj.org | Direct: +91 866 2489004"),
        ("State Registrar", "registrar@mudiraj.org | Direct: +91 866 2489002"),
        ("Dispute Tribunal", "tribunal@mudiraj.org | Vijayawada Central HQ"),
        ("Working Hours", "Mon - Sat (9:00 AM - 6:00 PM IST)"),
        ("Emergency Server Hotline", "+91 94400 00000 (24x7 System Monitoring)")
    ]
    for k, v in contacts:
        p = tf3.add_paragraph()
        p.font.size = Pt(8.5)
        p.space_before = Pt(3.5)
        rk = p.add_run()
        rk.text = "• " + k + ":\n  "
        rk.font.bold = True
        rk.font.color.rgb = GREEN_SUCCESS
        rv = p.add_run()
        rv.text = v
        rv.font.color.rgb = SLATE_MUTED

    add_bottom_callout(slide, "Admin Control Panel operations are overseen by the Andhra Pradesh State Mudiraj Mahasabha Central Committee.", "Governance:");
    add_footer(slide, 12, total_slides, doc_title)

    out_path = os.path.join(DOWNLOADS_DIR, "Mudiraj_Admin_Features_Guide.pptx")
    prs.save(out_path)
    root_path = os.path.join(PUBLIC_DIR, "Mudiraj_Admin_Features_Guide.pptx")
    prs.save(root_path)
    print(f"Generated Beautiful Admin Features Guide: {out_path} ({total_slides} slides)")

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================
if __name__ == "__main__":
    print("=" * 60)
    print("GENERATING BEAUTIFUL PROFESSIONAL POWERPOINT GUIDES WITH REAL UI MOCKUPS")
    print("=" * 60)
    build_beautiful_user_guide()
    build_beautiful_admin_guide()
    print("=" * 60)
    print("ALL POWERPOINT PRESENTATIONS SUCCESSFULLY GENERATED!")
    print("=" * 60)
