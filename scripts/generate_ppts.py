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

os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Color Palette
NAVY_DARK = RGBColor(15, 35, 71)       # #0f2347
NAVY_PRIMARY = RGBColor(30, 58, 138)   # #1e3a8a
AMBER_ACCENT = RGBColor(217, 119, 6)   # #d97706
GOLD_LIGHT = RGBColor(251, 191, 36)    # #fbbf24
SLATE_DARK = RGBColor(30, 41, 59)      # #1e293b
SLATE_MUTED = RGBColor(100, 116, 139)  # #64748b
SLATE_LIGHT = RGBColor(241, 245, 249)  # #f1f5f9
WHITE = RGBColor(255, 255, 255)
GREEN_SUCCESS = RGBColor(16, 185, 129) # #10b981
SKY_BLUE = RGBColor(14, 116, 144)      # #0e7490
BORDER_COLOR = RGBColor(226, 232, 240) # #e2e8f0

def init_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def add_header(slide, title_text, category_badge="ANDHRA PRADESH MUDIRAJ MAHASABHA", subtitle_text=""):
    # Header Background Banner
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = NAVY_DARK
    header_box.line.color.rgb = NAVY_DARK

    # Accent Stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.06))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = AMBER_ACCENT
    stripe.line.color.rgb = AMBER_ACCENT

    # Logo in header if exists
    if os.path.exists(LOGO_PATH):
        try:
            slide.shapes.add_picture(LOGO_PATH, Inches(0.4), Inches(0.12), width=Inches(0.9), height=Inches(0.9))
        except Exception:
            pass

    # Header Text
    left_pos = Inches(1.5) if os.path.exists(LOGO_PATH) else Inches(0.6)
    tb = slide.shapes.add_textbox(left_pos, Inches(0.12), Inches(11.0), Inches(0.9))
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
    # Footer line
    f_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(7.05), Inches(12.133), Inches(0.02))
    f_line.fill.solid()
    f_line.fill.fore_color.rgb = BORDER_COLOR
    f_line.line.color.rgb = BORDER_COLOR

    # Footer Left
    tb_left = slide.shapes.add_textbox(Inches(0.6), Inches(7.1), Inches(8.5), Inches(0.35))
    tf_left = tb_left.text_frame
    tf_left.margin_left = tf_left.margin_top = tf_left.margin_right = tf_left.margin_bottom = 0
    p = tf_left.paragraphs[0]
    p.text = f"Andhra Pradesh Mudiraj Community Portal | {doc_title}"
    p.font.size = Pt(9)
    p.font.color.rgb = SLATE_MUTED

    # Footer Right
    tb_right = slide.shapes.add_textbox(Inches(10.0), Inches(7.1), Inches(2.733), Inches(0.35))
    tf_right = tb_right.text_frame
    tf_right.margin_left = tf_right.margin_top = tf_right.margin_right = tf_right.margin_bottom = 0
    p_r = tf_right.paragraphs[0]
    p_r.alignment = PP_ALIGN.RIGHT
    p_r.text = f"Slide {current_slide} of {total_slides}"
    p_r.font.size = Pt(9)
    p_r.font.bold = True
    p_r.font.color.rgb = NAVY_PRIMARY

def add_title_slide(prs, title, subtitle, badge, extra_notes, doc_label):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Full Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY_DARK
    bg.line.color.rgb = NAVY_DARK

    # Accent Gold Top Stripe
    stripe_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    stripe_top.fill.solid()
    stripe_top.fill.fore_color.rgb = AMBER_ACCENT
    stripe_top.line.color.rgb = AMBER_ACCENT

    # Accent Gold Bottom Stripe
    stripe_bot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.35), Inches(13.333), Inches(0.15))
    stripe_bot.fill.solid()
    stripe_bot.fill.fore_color.rgb = AMBER_ACCENT
    stripe_bot.line.color.rgb = AMBER_ACCENT

    # Logo Center/Top
    if os.path.exists(LOGO_PATH):
        try:
            slide.shapes.add_picture(LOGO_PATH, Inches(5.916), Inches(0.7), width=Inches(1.5), height=Inches(1.5))
        except Exception:
            pass

    # Title Box
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

def add_card(slide, left, top, width, height, title, items, badge_text="", highlight_color=NAVY_PRIMARY):
    # Card Background
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = BORDER_COLOR
    card.line.width = Pt(1)

    # Top border strip inside card
    card_stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.08))
    card_stripe.fill.solid()
    card_stripe.fill.fore_color.rgb = highlight_color
    card_stripe.line.color.rgb = highlight_color

    # Text box
    tb = slide.shapes.add_textbox(left + Inches(0.22), top + Inches(0.16), width - Inches(0.44), height - Inches(0.26))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Badge if any
    first = True
    if badge_text:
        p_b = tf.paragraphs[0]
        p_b.text = badge_text.upper()
        p_b.font.size = Pt(8.5)
        p_b.font.bold = True
        p_b.font.color.rgb = highlight_color
        first = False

    # Title
    p_t = tf.paragraphs[0] if first else tf.add_paragraph()
    p_t.text = title
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE_DARK

    p_s = tf.add_paragraph()
    p_s.text = ""
    p_s.font.size = Pt(3)

    for item in items:
        p = tf.add_paragraph()
        p.font.size = Pt(10)
        p.font.color.rgb = SLATE_DARK
        p.space_after = Pt(3)
        if isinstance(item, tuple):
            run_b = p.add_run()
            run_b.text = "• " + item[0] + ": "
            run_b.font.bold = True
            run_b.font.color.rgb = highlight_color
            run_t = p.add_run()
            run_t.text = item[1]
        else:
            p.text = "• " + item

def add_bottom_callout(slide, text, highlight_text=""):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(6.25), Inches(12.133), Inches(0.65))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(238, 242, 255)
    box.line.color.rgb = RGBColor(199, 210, 254)
    box.line.width = Pt(1)

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(6.32), Inches(11.7), Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.font.size = Pt(10.5)
    if highlight_text:
        run_h = p.add_run()
        run_h.text = highlight_text + " "
        run_h.font.bold = True
        run_h.font.color.rgb = NAVY_PRIMARY
    run_t = p.add_run()
    run_t.text = text
    run_t.font.color.rgb = SLATE_DARK


# ==============================================================================
# BUILD PRESENTATION 1: PUBLIC WEBSITE USER GUIDE (ALL PAGES & FORMS)
# ==============================================================================
def build_website_user_guide():
    prs = init_presentation()
    total_slides = 14
    doc_title = "Website User Guide & Form Manual (All Pages)"

    # Slide 1: Title
    add_title_slide(
        prs,
        title="Andhra Pradesh Mudiraj Community Portal",
        subtitle="Complete Website & Interactive Forms User Guide",
        badge="Official Portal & Form Completion Manual",
        extra_notes=[
            "Detailed Guide for Every Webpage, Input Form, Validation Rule & Digital Service",
            "State Jurisdiction: Andhra Pradesh | Serving 20 Lakh+ Community Members",
            "Covers: Public Registration Form, Status Query, ID Card Generator & Navigation"
        ],
        doc_label=doc_title
    )

    # Slide 2: Portal Navigation & Interactive Forms Map
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Website Pages & Interactive Forms Map", "PORTAL SITEMAP", "Master map of every webpage and user-fillable form across the website")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(3.8), Inches(4.6),
        "Informational Pages",
        [
            ("Home Page (/)", "State central portal with active community statistics snapshot, 5-tier structure, and leadership highlights."),
            ("About Us (/about)", "Community history, 20L+ members vision, youth education, mahila empowerment, and 30-member committee tables."),
            ("Structure (/organization)", "Interactive 5-tier organizational tree from State to Village Gramam with real-time active unit counters."),
            ("Leadership (/leadership)", "Directory of appointed office bearers across Main Team, Youth Team, and Mahila Team wings."),
            ("Contact Us (/contact)", "State Headquarters (Vijayawada), 6 district offices, and toll-free helpline desk.")
        ],
        "Information Hub",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(4.75), Inches(1.4), Inches(3.8), Inches(4.6),
        "Interactive Forms to Fill",
        [
            ("1. Membership Form (/apply)", "Full public registration form: personal info, contact, location hierarchy, education, occupation, and photo upload."),
            ("2. Status Query Form (/membership)", "Lookup form to track application progress or verify membership using ID, App No, or Mobile."),
            ("3. ID Card Form (/id-card)", "Lookup form to render and download/print official dual-sided PVC smartcards with QR code."),
            ("4. Admin Login Form (/login)", "Secure authentication gateway for authorized operators and reviewers.")
        ],
        "Forms to Fill",
        AMBER_ACCENT
    )
    add_card(
        slide, Inches(8.9), Inches(1.4), Inches(3.8), Inches(4.6),
        "Form Experience Highlights",
        [
            ("Zero Account Hassle", "Public members do not need to create accounts, remember passwords, or sign up to fill forms."),
            ("Smart Pincode Auto-Fill", "Typing a 6-digit Indian PIN code automatically queries the Postal API to suggest address details."),
            ("Instant Application ID", "Submitting the form immediately generates an official APP-YYYY-XXXX tracking reference."),
            ("SMS Confirmation Alerts", "Automated SMS is sent to the registered mobile number on submission and upon approval.")
        ],
        "Convenience Features",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "All forms feature real-time client-side validation, error tooltips, and responsive layout for mobile phones.", "User Experience:");
    add_footer(slide, 2, total_slides, doc_title)

    # Slide 3: Page 1 - Home Page
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page 1: Home Page ( / ) — Central State Portal", "STATE OVERVIEW & SHORTCUTS", "The gateway to community unity, governance, and online member services")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Key Sections on Home Page",
        [
            ("Hero Banner & State Mission", "Highlights unity, welfare, and official AP state jurisdiction with prominent action buttons."),
            ("Organizational Snapshot", "Live statistics displaying active Districts, Assembly Constitutions, Mandals, and Community Teams."),
            ("5-Tier Governance Overview", "Clear visual representation of administrative tiers from State to grassroots Gramams."),
            ("3 Leadership Wings", "Introduction to Main Team, Youth Team, and Mahila Team with dedicated mandates."),
            ("Membership Highlights", "3-step visual guide on registration, verification, and permanent ID issuance."),
            ("Leadership Spotlight", "Cards showcasing active State office bearers with their portfolios and locations.")
        ],
        "State Hub",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "How Visitors Can Use This Page",
        [
            ("1. Check Community Strength", "View the live snapshot count of registered mandals and sanctioned teams across Andhra Pradesh."),
            ("2. Direct Shortcut to Apply Form", "Click 'Apply for Membership' (Gold button) to directly open the public registration form."),
            ("3. Track Application Progress", "Click 'Check Application Status' (Outline button) to jump directly to the status verification tool."),
            ("4. Download Website Guide (PPT)", "Download this complete user guide directly from the navbar dropdown or top banner."),
            ("5. Admin Login Link", "Admins can access the login gateway from the top banner link ('Admin Login').")
        ],
        "User Actions",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "The Home Page provides immediate one-click shortcuts to both membership registration and application status lookup.", "Quick Access:")
    add_footer(slide, 3, total_slides, doc_title)

    # Slide 4: Page 2 - About Us Page
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page 2: About Us ( /about ) — Heritage & Structure", "COMMUNITY MANDATE", "History, mission, and the standard 30-member leadership structure")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Mission, Heritage & 3 Pillars",
        [
            ("Community Scale", "Connecting 20 Lakh+ Mudiraj members with historic roots in fisheries, agriculture, defense, and public governance."),
            ("Community Unity & Identity", "Lifelong permanent Membership ID (MUD-00000001 format) ensuring genuine representation."),
            ("Youth Education & Skill Drives", "Competitive exam guidance, scholarship assistance, and career mentorship through Youth Teams."),
            ("Mahila Empowerment", "Dedicated women leadership, micro-business guidance, and family welfare support through Mahila Teams."),
            ("State Jurisdiction", "Unified digital governance across all 26 restructured districts of Andhra Pradesh.")
        ],
        "Community Mandate",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "30-Member Committee Structure Table",
        [
            ("President / Chairman (1 Position)", "Heads the executive committee at State, District, Constitution, or Mandal level."),
            ("Vice Presidents (6 Positions)", "Zonal and regional coordination across subdivisions."),
            ("General Secretaries (2 Positions)", "Overall organizational operations, record keeping, and meetings."),
            ("Joint Secretaries (6 Positions)", "Assisting operations across specific zonal portfolios."),
            ("Treasurer (1 Position)", "Financial stewardship, accounts, and welfare fund audits."),
            ("Executive Committee (14 Positions)", "Grassroots representatives and advisory committee members.")
        ],
        "Sanctioned Hierarchy",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "Every administrative level (State, District, Constitution, Mandal) maintains 3 distinct wings of 30 leaders each (total 90 leaders per unit).", "Governance Scale:")
    add_footer(slide, 4, total_slides, doc_title)

    # Slide 5: Page 3 - Organization Tree Page
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page 3: Organization Tree ( /organization )", "HIERARCHY EXPLORER & SEARCH", "Interactive structural breakdown from State to Village Gramam levels")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "5-Tier Structure Explained",
        [
            ("Tier 1: State Level", "Apex body: Andhra Pradesh State Mudiraj Mahasabha governing all statewide committees."),
            ("Tier 2: District Level", "26 Official AP District Committees coordinating regional initiatives."),
            ("Tier 3: Assembly Constitution Level", "175 Assembly constituencies providing democratic political and civic representation."),
            ("Tier 4: Mandal Level", "Local administrative mandal units directly serving grassroots towns and clusters."),
            ("Tier 5: Gramam Level (Future)", "Village-level community bodies ensuring representation at the grassroots doorstep.")
        ],
        "Governance Tiers",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Interactive Tree Features & Quick Search",
        [
            ("Active Unit Counter Badges", "Real-time summary counters showing exact number of operational units at each tier."),
            ("Search Form Input", "Type any District, Constituency, or Mandal name (e.g. 'Vijayawada', 'Guntur') to highlight matching units instantly."),
            ("Expandable Node Hierarchy", "Click to expand District nodes to reveal child Constitutions, and further expand to see Mandals."),
            ("Direct Leadership Link", "Click 'View Leaders' on any unit node to jump straight to the Leadership Directory for that unit."),
            ("Unit Status Badges", "Clear visual indicators showing active units vs pre-configured bodies.")
        ],
        "Explorer Features",
        SKY_BLUE
    )
    add_bottom_callout(slide, "Users can navigate through hundreds of Mandals and Assembly Constitutions without page reloads using the real-time search box.", "User Tip:")
    add_footer(slide, 5, total_slides, doc_title)

    # Slide 6: Page 4 - Leadership Directory Page
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page 4: Leadership Directory ( /leadership )", "OFFICE BEARERS DIRECTORY & FILTERS", "Filter and connect with appointed committee leaders across all wings")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "The 3 Dedicated Wings",
        [
            ("Main Team (Core Executive)", "Heads community policy, socio-political representation, community disputes, and government relations."),
            ("Youth Team (Youth Leadership)", "Mobilizes student community, career mentorship, job drives, sports meets, and IT awareness."),
            ("Mahila Team (Women Leadership)", "Drives women safety, healthcare, education of daughters, and self-help entrepreneurship."),
            ("Strict Merit & Verification", "Every leader displayed must be a verified registered member with an issued Membership ID.")
        ],
        "Leadership Wings",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "How to Search & Filter Leaders",
        [
            ("Filter by Level Dropdown", "Select 'State', 'District', 'Constitution', or 'Mandal' from the level dropdown."),
            ("Filter by Wing Tabs", "Toggle between 'Main Team', 'Youth Team', and 'Mahila Team'."),
            ("Filter by Location Dropdown", "Select specific District or Mandal to view local office bearers."),
            ("Leader Profile Cards", "Displays Leader's Name, Official Designation (e.g., President, Gen Sec), Verified Membership ID, Photo, and Geographic Scope."),
            ("Direct Contact Scope", "Enables local community members to identify authorized leaders for assistance.")
        ],
        "Filter Controls",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "Only verified members approved by the State Admin with assigned official portfolios appear in this public directory.", "Integrity Assurance:")
    add_footer(slide, 6, total_slides, doc_title)

    # Slide 7: FORM FOCUS - Membership Registration Form (Part 1: Personal Profile)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "FORM FOCUS: Membership Application ( /apply ) — Part 1", "REGISTRATION FORM GUIDE", "Detailed breakdown of Personal Profile, Socio-Economic details & Photo upload")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Section 1: Personal Information Fields",
        [
            ("Full Name * (Mandatory)", "Enter applicant's complete legal name as stated on Aadhaar Card or Voter ID."),
            ("Father's / Husband's Name * (Mandatory)", "Parent or spouse name for authentic identity and genealogical records."),
            ("Mother's Name (Optional)", "Mother's name for complete family registry."),
            ("Date of Birth * (Mandatory)", "Date picker (YYYY-MM-DD); system automatically validates age requirements."),
            ("Gender Selection *", "Radio / dropdown selection: 'Male', 'Female', or 'Other'."),
            ("Blood Group Dropdown", "Options: A+, A-, B+, B-, O+, O-, AB+, AB-, or Unknown (crucial for community blood donation drives).")
        ],
        "Personal Profile",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Section 2: Socio-Economic Profile & Photo",
        [
            ("Educational Qualification", "Select highest degree achieved: 10th, Intermediate, Diploma, Graduate, Post Graduate, or Doctorate."),
            ("Profession / Occupation", "Specify primary vocation: Agriculture, Inland Fisheries, Business, Govt Employee, Private Sector, Student, or Self-Employed."),
            ("Passport Photo Upload", "Upload a clear, front-facing passport-sized portrait (PNG, JPG, WebP format)."),
            ("Live Image Preview", "The form immediately renders an instant thumbnail preview of the uploaded photo."),
            ("Digital ID Card Usage", "This uploaded photo will be printed directly onto your official dual-sided Digital Membership ID Card.")
        ],
        "Vocation & Photo",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "Fields marked with a red asterisk (*) are strictly required. Missing mandatory fields will highlight with inline error prompts.", "Form Requirement:")
    add_footer(slide, 7, total_slides, doc_title)

    # Slide 8: FORM FOCUS - Membership Registration Form (Part 2: Contact & Location)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "FORM FOCUS: Membership Application ( /apply ) — Part 2", "CONTACT & LOCATION HIERARCHY", "Contact validation, cascading location selectors, smart pincode auto-fill & submission")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Section 3: Contact & Government ID",
        [
            ("Primary Mobile Number * (Mandatory)", "10-digit Indian mobile number. Validates that the number starts with 6, 7, 8, or 9 without +91 or 0 prefix."),
            ("Alternate / WhatsApp Number", "Optional secondary contact number for WhatsApp circulars and committee announcements."),
            ("Email Address", "Optional email to receive digital membership notifications and meeting notices."),
            ("Government ID Type", "Select official document: 'Aadhaar Card', 'Voter ID', 'Ration Card', 'PAN Card', or 'Driving License'."),
            ("Government ID Number", "Enter corresponding document number for identity verification by the district reviewer.")
        ],
        "Contact & ID Proof",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Section 4: Cascading Location & Submission",
        [
            ("District Selector *", "Select your AP District from the list of 26 districts (e.g. NTR, Visakhapatnam, Guntur, Kurnool)."),
            ("Assembly Constitution *", "Dropdown automatically updates to display only the constituencies within the selected district."),
            ("Mandal Selector *", "Dropdown dynamically updates to show only mandals belonging to the chosen constitution."),
            ("6-Digit Postal PIN Code", "Entering your PIN code automatically triggers a lookup to fill local post office and street area."),
            ("Door No & Street Address", "House / Door number, Street name, Colony, or Village landmark."),
            ("Submission Outcome", "Click 'Submit Application' &rarr; Generates Application ID (APP-YYYY-XXXX) and dispatches automated SMS confirmation!")
        ],
        "Location & Submit",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "The cascading dropdowns ensure you cannot accidentally pick a mandal that does not belong to your district.", "Smart Validation:")
    add_footer(slide, 8, total_slides, doc_title)

    # Slide 9: FORM FOCUS - Membership Status & Verification Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "FORM FOCUS: Status Lookup Form ( /membership )", "VERIFY MEMBERSHIP & TRACKING", "How to query application status using ID, Application Number, or Mobile")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "The Status Lookup Form Input",
        [
            ("Universal Input Box", "A single, intuitive search box that accepts any of three identifier formats:"),
            ("Option 1: Membership ID", "Enter permanent ID: e.g., MUD-00000001 or MUD-00000042."),
            ("Option 2: Application Number", "Enter the tracking number received upon submitting: e.g., APP-2026-0009."),
            ("Option 3: Registered Mobile Number", "Enter the 10-digit phone number submitted in the application: e.g., 9848011223."),
            ("Click 'Check Status'", "Button triggers instant API lookup to /api/public/apply without page reload.")
        ],
        "Query Form",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Understanding Status Form Output",
        [
            ("Status: 'Active' (Approved)", "Application has been verified by the Admin! Displays permanent MUD-ID and a green 'Download Digital ID Card' action button."),
            ("Status: 'Pending'", "Application is in the review queue awaiting document and residence verification by District Committee."),
            ("Status: 'Rejected'", "Application was not approved (e.g., duplicate submission or unverified details)."),
            ("Dossier Summary Display", "Displays applicant's Full Name, Father's Name, Gender, Blood Group, District, Mandal, and Submission Date for verification.")
        ],
        "Result Display",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "If you lose your Application Number, simply enter your registered 10-digit mobile number to retrieve your status immediately.", "Convenience Tip:")
    add_footer(slide, 9, total_slides, doc_title)

    # Slide 10: FORM FOCUS - Digital ID Card Generation & Print Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "FORM FOCUS: Digital ID Card Form ( /id-card )", "OFFICIAL DIGITAL MEMBER CARD", "Generate, preview, and print official dual-sided smart ID cards")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "The ID Card Search Form",
        [
            ("Search Form Field", "Enter your approved Permanent Membership ID (e.g., MUD-00000001) or Application Number."),
            ("Click 'Search / Generate'", "Queries /api/public/status to verify active membership."),
            ("Active Status Requirement", "Only members with status 'Active' can generate ID cards. Pending applicants receive an informative notice."),
            ("Instant Visual Rendering", "Card renders immediately on screen in official royal blue and saffron gold themes.")
        ],
        "Card Search Form",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Dual-Sided Card Output & Print Controls",
        [
            ("Front Face of Card", "Official Mudiraj Emblem, Member Portrait, Name, Permanent ID, QR Code, and AP State Jurisdiction."),
            ("Back Face of Card", "Full Residential Address, Mobile Number, Blood Group, Issue Date, and State Authority Sign & Seal."),
            ("Print / PDF Action Button", "Click 'Print Card / Save as PDF' to open the optimized print dialog."),
            ("Standard Wallet Size", "Dimensions calibrated to standard CR-80 plastic smartcard (credit-card size) for PVC card printing."),
            ("A4 Sheet Compatibility", "Prints with high resolution on standard home and office A4 paper printers.")
        ],
        "Card Output",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "The embedded QR code on the front of the card allows instant verification of member authenticity by scanning with any smartphone camera.", "Anti-Forgery:")
    add_footer(slide, 10, total_slides, doc_title)

    # Slide 11: FORM FOCUS - Admin Login Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "FORM FOCUS: Admin Login Form ( /login )", "SECURE AUTHENTICATION GATEWAY", "Restricted access point for state administrators and district reviewers")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Admin Login Form Fields",
        [
            ("Username Field *", "Input authorized administrator username (e.g., admin or district coordinator username)."),
            ("Password Field *", "Input secure administrator password with hidden character masking."),
            ("Sign In Button", "Submits POST request to /api/auth/login with encrypted JSON payload."),
            ("Error Feedback Display", "Clear error banner appears if credentials are invalid or account is locked."),
            ("Public Separation", "Community members do not need to access this form; it is strictly for authorized reviewers.")
        ],
        "Login Form Fields",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Session Security & Access Handling",
        [
            ("HttpOnly Session Cookie", "On successful login, the server sets a tamper-proof, secure HTTP-only session cookie."),
            ("Automatic Redirection", "Authorized admins are immediately routed to the Admin Dashboard (/admin)."),
            ("Route Protection Guards", "Unauthorized attempts to access /admin routes directly are automatically redirected back to /login."),
            ("Logout Button", "Admins can safely terminate their session at any time from the bottom of the admin sidebar.")
        ],
        "Security Flow",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "Always ensure you log out of the admin panel on shared or public computers to protect community records.", "Security Rule:")
    add_footer(slide, 11, total_slides, doc_title)

    # Slide 12: Page 8 - Contact Us & Helpdesk
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Page 8: Contact Us & Helpdesk ( /contact )", "STATE & DISTRICT CENTERS", "Connect with State Headquarters, regional offices, and support helplines")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "State Headquarters & District Centers",
        [
            ("State Headquarters (Vijayawada)", "Andhra Pradesh Mudiraj Bhavan, M.G. Road, Benz Circle, Vijayawada, NTR District - 520010."),
            ("Visakhapatnam District Office", "MVP Colony, Sector 5, Visakhapatnam - 530017 (Phone: +91 891-2554411)."),
            ("Guntur District Office", "Arundelpet 4th Line, Guntur - 522002 (Phone: +91 863-2233445)."),
            ("Rayalaseema Centers", "Regional offices in Kurnool, Tirupati, and Anantapuramu with local coordinators.")
        ],
        "Office Network",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Helplines & Communication Channels",
        [
            ("Toll-Free Helpline", "1800-425-MUDIRAJ — Operational Mon to Sat, 9:30 AM to 6:00 PM."),
            ("Direct State Desk", "+91 866-2478900 for executive inquiries and committee nominations."),
            ("Official Email Support", "support@mudirajcommunityap.org for corrections, name updates, or grievances."),
            ("Form Help Support", "Contact helpdesk if you need assistance filling the online application form.")
        ],
        "Support Channels",
        SKY_BLUE
    )
    add_bottom_callout(slide, "State HQ and District Offices assist community members with offline applications, document verification, and ID collection.", "Assistance:")
    add_footer(slide, 12, total_slides, doc_title)

    # Slide 13: Complete Form Lifecycle Walkthrough
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Complete Form Lifecycle & Submission Walkthrough", "STEP-BY-STEP WORKFLOW", "From filling the form online to holding your official community ID card")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(3.8), Inches(4.6),
        "Step 1: Fill the Form",
        [
            ("Open /apply", "Navigate to the public application page from the navbar or homepage."),
            ("Fill Personal Info", "Input Full Name, Parent Name, DOB, Gender, and Blood Group."),
            ("Select AP Location", "Choose District &rarr; Constitution &rarr; Mandal &rarr; Village."),
            ("Input Contact & Photo", "Enter 10-digit mobile number, Govt ID details, and upload portrait."),
            ("Submit Form", "Click 'Submit Application'.")
        ],
        "Stage 1: Application",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(4.75), Inches(1.4), Inches(3.8), Inches(4.6),
        "Step 2: Processing & SMS",
        [
            ("Instant Reference", "Screen displays your unique Application Number (APP-YYYY-XXXX)."),
            ("SMS Confirmation", "Automated SMS notification arrives on your phone with the reference."),
            ("District Review", "Application enters District Admin Queue for residential verification."),
            ("Track Status", "Check /membership at any time using your mobile number or APP ID.")
        ],
        "Stage 2: Verification",
        AMBER_ACCENT
    )
    add_card(
        slide, Inches(8.9), Inches(1.4), Inches(3.8), Inches(4.6),
        "Step 3: Approval & ID Card",
        [
            ("Admin Approval", "Admin verifies details and issues permanent MUD-XXXXX Membership ID."),
            ("Approval SMS", "SMS confirmation sent with your new lifelong Membership ID."),
            ("Generate ID Card", "Visit /id-card, enter your ID, and view your official dual-sided card."),
            ("Print & Laminate", "Print card on paper or plastic PVC for official community identification.")
        ],
        "Stage 3: Completion",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "The entire online application process takes less than 3 minutes to complete from any smartphone or computer.", "Quick & Simple:")
    add_footer(slide, 13, total_slides, doc_title)

    # Slide 14: Form Troubleshooting & Common Questions
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Form Troubleshooting, Validation Rules & FAQs", "HELP & SUPPORT", "Answers to common questions and resolving form validation warnings")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Form Validation Rules Explained",
        [
            ("Mobile Number Error", "Ensure you enter exactly 10 digits starting with 6, 7, 8, or 9. Do not type '+91' or '0' at the start."),
            ("Mandal Dropdown Empty", "You must first select your District and Assembly Constitution; mandals populate automatically based on the selected constitution."),
            ("Pincode Not Populating", "Pincode lookup requires a valid 6-digit Indian postal code. You can also type your village or street name manually."),
            ("Photo Upload Size", "Use clear images under 5MB. Standard portrait orientation works best for ID cards."),
            ("Duplicate Mobile Error", "Each member should have a unique mobile number. If you already applied, check your status on /membership.")
        ],
        "Validation Tips",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Frequently Asked Questions (FAQs)",
        [
            ("How long does verification take?", "District committees usually review and approve applications within 24 to 48 hours."),
            ("Can I update my phone number later?", "Yes, contact your District Helpdesk or State Admin to update your contact or address details."),
            ("Is there any registration fee?", "Online application and digital ID card generation are 100% free for community members."),
            ("What if I forgot my Application Number?", "Simply enter your registered 10-digit mobile number on /membership to retrieve your record."),
            ("Need offline assistance?", "Visit any of our 6 regional district offices in AP for in-person form submission assistance.")
        ],
        "Member FAQs",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "For further help with forms, call the State Toll-Free Support line at 1800-425-MUDIRAJ (Mon-Sat, 9:30 AM to 6:00 PM).", "Helpline:")
    add_footer(slide, 14, total_slides, doc_title)

    out_path = os.path.join(DOWNLOADS_DIR, "Mudiraj_Website_User_Guide.pptx")
    prs.save(out_path)
    # Also save to public root for direct fallback access
    root_path = os.path.join(PUBLIC_DIR, "Mudiraj_Website_User_Guide.pptx")
    prs.save(root_path)
    print(f"[OK] Website User Guide saved to: {out_path} and {root_path}")
    return out_path


# ==============================================================================
# BUILD PRESENTATION 2: ADMIN PANEL FEATURES & FORMS GUIDE
# ==============================================================================
def build_admin_features_guide():
    prs = init_presentation()
    total_slides = 12
    doc_title = "Admin Console & Features Manual"

    # Slide 1: Title
    add_title_slide(
        prs,
        title="Andhra Pradesh Mudiraj Community",
        subtitle="Administrator Panel & Features Comprehensive Manual",
        badge="Official Admin Console Operations Guide",
        extra_notes=[
            "Administrative Hub for State Executives, District Reviewers & System Operators",
            "Managing 26 Districts | 175 Assembly Constitutions | 679+ Mandals | 3 Wings",
            "Complete Guide to Admin Forms: Direct Entry, Approval Queue, Masters, Teams & Settings"
        ],
        doc_label=doc_title
    )

    # Slide 2: Admin Architecture & Navigation Overview
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Console Architecture & Navigation", "SYSTEM OVERVIEW", "Modular administrative suite for state and district community operations")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(3.8), Inches(4.6),
        "Core Management Modules",
        [
            ("Dashboard (/admin)", "Real-time state metrics, member counts, application queues, and active units."),
            ("Members (/admin/members)", "Comprehensive member registry, profile editor, status toggle, Excel export, and manual entry."),
            ("Applications (/admin/applications)", "Approval queue for public applicants with one-click MUD-ID generation and SMS alerts.")
        ],
        "Registry & Queue",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(4.75), Inches(1.4), Inches(3.8), Inches(4.6),
        "Organizational Structure",
        [
            ("Masters (/admin/organizations)", "5-tier geographic hierarchy management: State, Districts, Constitutions, Mandals, Gramams."),
            ("Teams (/admin/teams)", "Governing 3 leadership wings (Main, Youth, Mahila) across all units."),
            ("NRI Teams (/admin/nri-teams)", "Global diaspora chapters (USA, UK, Gulf, Australia) and overseas committees.")
        ],
        "Structure & Teams",
        AMBER_ACCENT
    )
    add_card(
        slide, Inches(8.9), Inches(1.4), Inches(3.8), Inches(4.6),
        "Governance, Reports & Config",
        [
            ("Leaders (/admin/leaders)", "Appointment of verified members to ~30 sanctioned leadership portfolios."),
            ("Reports (/admin/reports)", "District-by-district breakdown, penetration analytics, and CSV report export."),
            ("Settings (/admin/settings)", "Admin password changes, India (+91) SMS gateway configuration, test SMS dispatcher, audit logs.")
        ],
        "System Control",
        SKY_BLUE
    )
    add_bottom_callout(slide, "The Admin Console includes responsive sidebar navigation, mobile drawer support, secure HttpOnly session cookies, and live data synchronization.", "Console Architecture:")
    add_footer(slide, 2, total_slides, doc_title)

    # Slide 3: Module 1 - Admin Dashboard
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 1: Admin Dashboard ( /admin )", "REAL-TIME STATE ANALYTICS", "High-level overview of membership metrics, applications, and geographic distribution")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Key Metrics & Live Statistics",
        [
            ("Total Members (Issued ID)", "Count of approved members with permanent MUD-XXXXX IDs in good standing."),
            ("Total Database Records", "Aggregate of all submissions including pending, approved, and inactive records."),
            ("Pending Applications Queue", "Immediate counter of applications awaiting admin verification."),
            ("Approved vs Inactive / Rejected", "Live breakdown of active community strength vs rejected or deactivated files."),
            ("Geographic Unit Counts", "Total operational Districts (26), Assembly Constitutions (175), and Mandals (679+)."),
            ("Leadership Capacity", "Total active teams (Main/Youth/Mahila) and appointed office bearers.")
        ],
        "Analytical Cards",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Dashboard Operations & Action Links",
        [
            ("Immediate Queue Access", "Click on 'Pending Applications' card to jump straight to the verification queue."),
            ("Direct Member Search", "Click 'Total Members' to launch the comprehensive member search table."),
            ("Download Admin Guide (PPT)", "Direct download button in the dashboard header and topbar navbar."),
            ("Executive Summary Briefing", "Provides at-a-glance state metrics ready for executive review meetings."),
            ("Auto-refreshing Data", "Metrics compute dynamically from the database on every dashboard view.")
        ],
        "Operational Value",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "The Dashboard gives administrators instant situational awareness over the entire 20 Lakh+ community membership base.", "Key Advantage:")
    add_footer(slide, 3, total_slides, doc_title)

    # Slide 4: Module 2 - Member Database Management & Forms
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 2: Member Database & Forms ( /admin/members )", "MEMBER REGISTRY & ADMIN FORMS", "Direct member registration form, edit profile form, search filters & Excel export")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Search & Multi-Filter Form Controls",
        [
            ("Keyword Search Form", "Search instantly across Name, Membership ID (MUD-), Application No (APP-), or Mobile Number."),
            ("Cascading Location Filters", "Filter dynamically by District, Assembly Constitution, or Mandal."),
            ("Status Filter Dropdown", "Filter by Active, Pending, Inactive, or All statuses."),
            ("Pagination & Sorting", "Sort by Name, Date, or ID with customizable rows-per-page (10, 25, 50, 100)."),
            ("Excel Export (.xlsx)", "One-click export of filtered or full member datasets with professional styling.")
        ],
        "Filter & Export Forms",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Admin Forms: Direct Registration & Edit Modal",
        [
            ("Direct Registration Form", "Click 'Register New Member Directly' &rarr; Opens member form in admin mode. Bypasses pending queue and immediately issues active MUD-ID."),
            ("Edit Member Profile Form", "Click 'Edit' on any member row &rarr; Opens modal form to update full name, father name, phone, address, blood group, or occupation."),
            ("Status Toggle Form Action", "Click 'Activate' or 'Deactivate' to toggle member standing without data deletion."),
            ("Digital ID Card Generator", "Launch the high-resolution printable ID Card directly from any member record."),
            ("Audit Trail", "All updates made via admin forms are automatically timestamped.")
        ],
        "Admin Member Forms",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "Admins can directly register offline or VIP applicants with instant active ID generation using the Direct Registration Form.", "Direct Enrollment:")
    add_footer(slide, 4, total_slides, doc_title)

    # Slide 5: Module 3 - Membership Applications Review Queue
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 3: Applications Approval Queue ( /admin/applications )", "VERIFICATION PIPELINE & FORMS", "Inspect applicant submissions, approve with auto-ID generation, or reject with reason")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Application Review Workflow",
        [
            ("1. Dedicated Status Tabs", "Organized into 'Pending Review', 'Approved & Active', 'Rejected', and 'All Applications'."),
            ("2. Dossier Inspection", "Admin reviews applicant's full name, father/husband name, date of birth, photo, and address proofs."),
            ("3. Duplicate Check", "Search by mobile number ensures applicant does not already have an existing Membership ID."),
            ("4. Geographic Validation", "Confirms the applicant resides within the selected AP District, Constitution, and Mandal.")
        ],
        "Review Protocol",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Approval & Rejection Execution",
        [
            ("One-Click Approval", "Click 'Approve' to instantly transition application to 'Active' status."),
            ("Automatic MUD-ID Assignment", "System automatically queries the highest sequence number and issues the next permanent ID (e.g. MUD-00000042)."),
            ("Automated SMS Alert", "Backend triggers Fast2SMS / Twilio gateway to dispatch approval SMS with the new Membership ID."),
            ("Rejection with Feedback Form", "Click 'Reject' &rarr; Logs reason for rejection (e.g., incorrect address or invalid document)."),
            ("Immediate Status Reflection", "Applicant can immediately verify their active status and download ID card on public site.")
        ],
        "Action Handlers",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "Approving an application automatically activates the member, issues their lifelong ID, and dispatches an automated SMS alert.", "Automated Pipeline:")
    add_footer(slide, 5, total_slides, doc_title)

    # Slide 6: Module 4 - Organization Masters Management & Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 4: Organization Masters ( /admin/organizations )", "5-TIER TERRITORIAL MASTERS FORM", "Create and maintain State, District, Constitution, Mandal, and Gramam units")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "5-Tier Hierarchy Governance",
        [
            ("Tier 1: State", "Root territorial jurisdiction: Andhra Pradesh State."),
            ("Tier 2: District", "26 Official Districts (e.g. NTR, Guntur, Krishna, Visakhapatnam, Kurnool)."),
            ("Tier 3: Assembly Constitution", "175 Legislative constituencies linked directly to their parent District."),
            ("Tier 4: Mandal", "679+ Mandals linked to their respective Assembly Constitution and District."),
            ("Tier 5: Gramam (Village)", "Grassroots village units linked to their parent Mandal.")
        ],
        "Territorial Tiers",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Master Unit Creation Form & Controls",
        [
            ("Select Level", "Choose organizational level: District, Constitution, Mandal, or Gramam."),
            ("Parent Unit Dropdown", "Select parent territorial body (e.g. Select District when adding a Constitution)."),
            ("Unit Name Input", "Enter official display name of the administrative unit (e.g. 'Mangalagiri')."),
            ("Unit Code Input", "Enter unique administrative abbreviation or identifier (e.g. 'MGL')."),
            ("Status Toggle", "Activate or Deactivate units with zero risk to existing member records."),
            ("Search & Paginated Grid", "Quickly find any unit among hundreds using instant keyword search.")
        ],
        "Unit Creation Form",
        SKY_BLUE
    )
    add_bottom_callout(slide, "Masters maintain the single source of truth for all location dropdowns across public registration and leadership assignments.", "Data Core:")
    add_footer(slide, 6, total_slides, doc_title)

    # Slide 7: Module 5 - Teams Management
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 5: Teams Management ( /admin/teams )", "3-WING GOVERNANCE", "Manage Main, Youth, and Mahila committee teams across every administrative unit")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "3 Leadership Wings Architecture",
        [
            ("Main Team", "Heads general administration, political representations, legal affairs, and welfare initiatives."),
            ("Youth Team", "Empowers students and youth with education scholarships, competitive coaching, and sports."),
            ("Mahila Team", "Coordinates women leadership, self-help groups, healthcare camps, and family welfare."),
            ("Standard Capacity", "Each team across any unit is sanctioned for ~30 distinct leadership portfolios.")
        ],
        "Wings Overview",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Team Exploration & Operations",
        [
            ("Filter by Tier Dropdowns", "Filter teams by State, District, Assembly Constitution, or Mandal."),
            ("Filter by Wing Tabs", "Toggle between Main Team, Youth Team, or Mahila Team views."),
            ("Team Roster Cards", "Displays team name, jurisdiction, sanctioned vs filled leadership seats, and leader count."),
            ("Inspect Team Members", "Expand any team card to review the list of appointed leaders and vacant positions."),
            ("Direct Assignment Link", "Direct shortcut to assign verified members into open leadership positions.")
        ],
        "Team Operations",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "Every single Mandal and District in AP is pre-structured with 3 teams, ensuring balanced representation for women and youth.", "Balanced Governance:")
    add_footer(slide, 7, total_slides, doc_title)

    # Slide 8: Module 6 - NRI Teams & Global Chapters Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 6: NRI Teams Management ( /admin/nri-teams )", "GLOBAL DIASPORA & FORMS", "Forms to create international chapters and appoint overseas committee leaders")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Create New NRI Chapter Form",
        [
            ("Chapter Name Input", "Enter country / chapter designation: e.g., 'United States of America (USA Chapter)'."),
            ("Country Code Input", "Enter standard 2-letter ISO country code: e.g., 'US', 'UK', 'AE', 'AU'."),
            ("Chapter Description", "Brief summary of community objectives and regional coordination scope."),
            ("Submit Action", "Initializes international chapter with Main, Youth, and Mahila wings ready for leader appointments.")
        ],
        "Chapter Creation Form",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Direct NRI Leader Appointment Form",
        [
            ("Overseas Leader Form Fields", "Full Name, Country, City of Residence, International Mobile Number, and Email."),
            ("Position Selection", "Designate portfolio: President, Vice President, General Secretary, or Coordinator."),
            ("Team Wing Selection", "Assign to NRI Main Team, NRI Youth Team, or NRI Mahila Team."),
            ("Link Existing Member Base", "Alternatively search and link registered diaspora members by their MUD-ID."),
            ("Global Directory Publication", "Appointed NRI leaders appear immediately in the public international directory.")
        ],
        "NRI Leader Form",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "The NRI module enables seamless coordination with overseas community executives across North America, Europe, the Middle East, and Asia-Pacific.", "Global Reach:")
    add_footer(slide, 8, total_slides, doc_title)

    # Slide 9: Module 7 - Leadership Directory Appointment Form
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 7: Leadership Management ( /admin/leaders )", "OFFICE BEARERS APPOINTMENT FORM", "Form to appoint verified members to ~30 official leadership designations")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Official Leadership Portfolios (~30 Roles)",
        [
            ("President / Chairman (1)", "Apex leader heading the committee at the designated level."),
            ("Vice Presidents (6)", "Senior leaders overseeing regional and departmental portfolios."),
            ("General Secretaries (2)", "Executive leadership handling official correspondence and meetings."),
            ("Joint Secretaries (6)", "Supporting executive operations across community zones."),
            ("Treasurer (1)", "Overseeing accounts, budget allocations, and welfare audits."),
            ("Executive Committee Members (14)", "Advisory and voting members guiding committee actions.")
        ],
        "Sanctioned Designations",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Leader Appointment Form Flow",
        [
            ("1. Search Member Input", "Type member's name or MUD-ID in the assignment panel search box."),
            ("2. Select Member Dossier", "Click matching member to load verified name, phone, and district."),
            ("3. Select Jurisdiction Level", "Choose State, District, Constitution, or Mandal unit."),
            ("4. Select Wing & Position", "Pick Main, Youth, or Mahila team and select the designated portfolio code."),
            ("5. Confirm Appointment", "System saves appointment, updates team roster, and publishes to public Leadership Directory."),
            ("Revocation / Term End Form", "Click 'Remove Leader' to revoke appointment and reopen vacant seat.")
        ],
        "Appointment Form",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "Only active registered members with verified MUD-IDs can be appointed to official leadership positions.", "Strict Qualification:")
    add_footer(slide, 9, total_slides, doc_title)

    # Slide 10: Module 8 - Reports & Analytical Intelligence
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 8: Reports & Analytics ( /admin/reports )", "STATEWIDE DATA INTELLIGENCE", "District-by-district breakdown, penetration analytics, and executive CSV export")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "District-by-District Breakdown Matrix",
        [
            ("District Name & Code", "Lists all 26 restructured districts of Andhra Pradesh with official codes."),
            ("Constitutions & Mandals Count", "Number of operational assembly constituencies and mandals per district."),
            ("Total Applications Submitted", "Cumulative applicant volume registered from each district."),
            ("Active Approved Members", "Verified community strength with issued MUD-IDs per district."),
            ("Pending Application Backlog", "Identifies districts requiring administrative review attention."),
            ("Assigned Leaders Count", "Tracks leadership appointment progress across district committees.")
        ],
        "Analytics Matrix",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "Executive Export & Decision Support",
        [
            ("One-Click CSV Export", "Click 'Export District Report (.csv)' to download a spreadsheet with complete district breakdowns."),
            ("State Executive Briefings", "Ready-made data for State Committee reviews, district president meetings, and audits."),
            ("Penetration Tracking", "Identifies underserved mandals and districts needing membership registration drives."),
            ("Resource Allocation", "Guides distribution of welfare funds, youth scholarships, and women self-help loans."),
            ("Dynamic Calculations", "All metrics recalculate in real time directly from the live SQL database.")
        ],
        "Executive Utilities",
        SKY_BLUE
    )
    add_bottom_callout(slide, "The Reports module provides state executives with data-driven insights to monitor community growth and leadership saturation.", "Strategic Value:")
    add_footer(slide, 10, total_slides, doc_title)

    # Slide 11: Module 9 - System Settings & Security Forms
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Module 9: Settings & Security Forms ( /admin/settings )", "PASSWORD, SMS GATEWAY & TEST FORMS", "Change admin password, configure SMS providers, dispatch test SMS & review audit logs")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(5.8), Inches(4.6),
        "Admin Password Change Form",
        [
            ("Current Password Input *", "Enter existing administrator password for verification."),
            ("New Password Input *", "Input secure new password (letters, digits, special characters)."),
            ("Confirm New Password *", "Re-type new password to avoid typographical errors."),
            ("Submit Action", "Updates encrypted password hash and invalidates old sessions."),
            ("Audit Trail", "Password change is immediately logged in the system audit trail.")
        ],
        "Password Form",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(6.8), Inches(1.4), Inches(5.9), Inches(4.6),
        "SMS Gateway & Test Dispatcher Forms",
        [
            ("SMS Gateway Config Form", "Configure SMS Provider (Fast2SMS or Twilio), API Key, Account SID, and Sender ID."),
            ("Save Credentials", "Encrypted storage ensures automated OTP and approval SMS dispatches work reliably."),
            ("Test SMS Dispatcher Form", "Enter an Indian 10-digit mobile number in the test box and click 'Send Test SMS'."),
            ("Live Status Feedback", "Screen displays immediate gateway response (Success, Message ID, Delivery status)."),
            ("SMS Delivery Log Table", "Inspect full history of all outgoing registration and approval SMS messages.")
        ],
        "SMS & Test Forms",
        AMBER_ACCENT
    )
    add_bottom_callout(slide, "The SMS test form allows administrators to verify live SMS delivery to any Indian mobile number with one click.", "Test Engine:")
    add_footer(slide, 11, total_slides, doc_title)

    # Slide 12: Admin Daily Operations Checklist & Best Practices
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, "Admin Daily Operations Checklist", "OPERATIONS & BEST PRACTICES", "Recommended daily, weekly, and monthly operational guidelines for administrators")
    
    add_card(
        slide, Inches(0.6), Inches(1.4), Inches(3.8), Inches(4.6),
        "Daily Operations Checklist",
        [
            ("1. Inspect Pending Queue", "Navigate to /admin/applications and check new submissions in the 'Pending' tab."),
            ("2. Verify Personal Details", "Verify applicant name, father's name, photo, and residential district/mandal."),
            ("3. Approve or Reject", "Approve genuine applications to auto-generate permanent MUD-IDs."),
            ("4. Review SMS Logs", "Check /admin/settings to confirm approval SMS alerts were successfully delivered.")
        ],
        "Daily Routine",
        NAVY_PRIMARY
    )
    add_card(
        slide, Inches(4.75), Inches(1.4), Inches(3.8), Inches(4.6),
        "Weekly Maintenance Tasks",
        [
            ("1. Review District Analytics", "Check /admin/reports to track district growth and pending application bottlenecks."),
            ("2. Fill Vacant Leadership Seats", "Coordinate with District Presidents to appoint leaders to vacant team positions."),
            ("3. Member Profile Updates", "Process any phone number or address updates requested by members."),
            ("4. NRI Chapter Coordination", "Review overseas diaspora applications and update NRI committee rosters.")
        ],
        "Weekly Routine",
        AMBER_ACCENT
    )
    add_card(
        slide, Inches(8.9), Inches(1.4), Inches(3.8), Inches(4.6),
        "Monthly Audits & Backups",
        [
            ("1. Full Database Export", "Export full member database to Excel (.xlsx) from /admin/members for offline archive."),
            ("2. Export District CSV", "Export monthly district statistical report from /admin/reports."),
            ("3. Security Audit Review", "Review audit logs on /admin/settings for any unusual administrative activities."),
            ("4. Password Rotation", "Regularly update admin password to maintain high console security.")
        ],
        "Monthly Routine",
        GREEN_SUCCESS
    )
    add_bottom_callout(slide, "Always log out of the Admin Console when finishing work by clicking 'Logout' at the bottom of the sidebar.", "Security Protocol:")
    add_footer(slide, 12, total_slides, doc_title)

    out_path = os.path.join(DOWNLOADS_DIR, "Mudiraj_Admin_Features_Guide.pptx")
    prs.save(out_path)
    # Also save to public root for direct fallback access
    root_path = os.path.join(PUBLIC_DIR, "Mudiraj_Admin_Features_Guide.pptx")
    prs.save(root_path)
    print(f"[OK] Admin Features Guide saved to: {out_path} and {root_path}")
    return out_path


if __name__ == "__main__":
    print("Regenerating PowerPoint Presentations with Detailed Forms Breakdown...")
    p1 = build_website_user_guide()
    p2 = build_admin_features_guide()
    print("Done! Both presentations successfully updated with comprehensive form guides.")
