"""
PEPS BOM Tool - User Guide PDF Generator
Run: C:\Python314\python.exe generate_user_guide.py
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
import os, datetime

# ── Brand colours ──────────────────────────────────────────────────────────
NAVY      = HexColor('#0f1923')
ACCENT    = HexColor('#1a56db')
ACCENT2   = HexColor('#0e3fa8')
INK2      = HexColor('#3a4a5a')
INK3      = HexColor('#7a8a9a')
LINE      = HexColor('#dde3ea')
BG_LIGHT  = HexColor('#f4f6f9')
GREEN     = HexColor('#065f46')
GREEN_BG  = HexColor('#d1fae5')
ORANGE    = HexColor('#92400e')
ORANGE_BG = HexColor('#fffbeb')
RED       = HexColor('#991b1b')
RED_BG    = HexColor('#fef2f2')
WHITE     = colors.white

W, H = A4  # 595 x 842 pts

OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        'PEPS_BOM_Tool_User_Guide.pdf')

# ── Page template with header/footer ───────────────────────────────────────
class PageTemplate:
    def __init__(self, title='PEPS BOM Tool — User Guide'):
        self.title = title

    def __call__(self, canv, doc):
        canv.saveState()
        # Header bar
        canv.setFillColor(NAVY)
        canv.rect(0, H - 36, W, 36, fill=1, stroke=0)
        canv.setFillColor(WHITE)
        canv.setFont('Helvetica-Bold', 9)
        canv.drawString(20, H - 22, 'PEPS BOM Tool')
        canv.setFillColor(HexColor('#5b9bf8'))
        canv.drawString(86, H - 22, '—')
        canv.setFillColor(HexColor('#aac8f8'))
        canv.setFont('Helvetica', 9)
        canv.drawString(96, H - 22, self.title)
        # Page number (right)
        canv.setFillColor(HexColor('#aac8f8'))
        canv.setFont('Helvetica', 8)
        canv.drawRightString(W - 20, H - 22, f'Page {doc.page}')
        # Footer
        canv.setFillColor(LINE)
        canv.rect(20, 20, W - 40, 0.5, fill=1, stroke=0)
        canv.setFillColor(INK3)
        canv.setFont('Helvetica', 7.5)
        canv.drawString(20, 10, 'PEPS BOM Automation Tool v2.2  ·  Confidential — Internal Use Only')
        canv.drawRightString(W - 20, 10, f'Generated {datetime.date.today().strftime("%d %b %Y")}')
        canv.restoreState()

# ── Style helpers ───────────────────────────────────────────────────────────
def styles():
    base = getSampleStyleSheet()

    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    return {
        'cover_title': ps('ct', fontSize=32, fontName='Helvetica-Bold',
                          textColor=WHITE, alignment=TA_CENTER, leading=38),
        'cover_sub':   ps('cs', fontSize=14, fontName='Helvetica',
                          textColor=HexColor('#aac8f8'), alignment=TA_CENTER, leading=20),
        'cover_ver':   ps('cv', fontSize=10, fontName='Helvetica',
                          textColor=HexColor('#7a8a9a'), alignment=TA_CENTER),

        'h1':  ps('h1', fontSize=18, fontName='Helvetica-Bold', textColor=ACCENT,
                  spaceBefore=16, spaceAfter=6, leading=22),
        'h2':  ps('h2', fontSize=13, fontName='Helvetica-Bold', textColor=NAVY,
                  spaceBefore=14, spaceAfter=4, leading=17),
        'h3':  ps('h3', fontSize=11, fontName='Helvetica-Bold', textColor=INK2,
                  spaceBefore=10, spaceAfter=3, leading=15),
        'body': ps('bd', fontSize=10, fontName='Helvetica', textColor=INK2,
                   leading=15, spaceAfter=5, alignment=TA_JUSTIFY),
        'step': ps('st', fontSize=10, fontName='Helvetica', textColor=INK2,
                   leading=15, spaceAfter=4, leftIndent=14),
        'note': ps('nt', fontSize=9.5, fontName='Helvetica', textColor=ORANGE,
                   leading=14, spaceAfter=4, leftIndent=8, rightIndent=8),
        'tip':  ps('tp', fontSize=9.5, fontName='Helvetica', textColor=GREEN,
                   leading=14, spaceAfter=4, leftIndent=8, rightIndent=8),
        'warn': ps('wn', fontSize=9.5, fontName='Helvetica', textColor=RED,
                   leading=14, spaceAfter=4, leftIndent=8, rightIndent=8),
        'mono': ps('mn', fontSize=9, fontName='Helvetica-Oblique', textColor=INK2,
                   leading=13, spaceAfter=3),
        'toc_section': ps('tc', fontSize=11, fontName='Helvetica-Bold',
                          textColor=NAVY, spaceAfter=2, spaceBefore=8),
        'toc_item':    ps('ti', fontSize=10, fontName='Helvetica',
                          textColor=INK2, spaceAfter=2, leftIndent=16),
        'label': ps('lb', fontSize=8.5, fontName='Helvetica-Bold',
                    textColor=INK3, leading=12, spaceAfter=1),
    }

S = styles()

# ── Reusable building blocks ────────────────────────────────────────────────
def section_banner(text, color=ACCENT):
    """Full-width coloured band with white text — used as major section header."""
    tbl = Table([[Paragraph(f'<font color="white"><b>{text}</b></font>',
                            ParagraphStyle('sb', fontSize=12, fontName='Helvetica-Bold',
                                           textColor=WHITE, leading=16))]],
                colWidths=[W - 80])
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), color),
        ('TOPPADDING',    (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING',   (0,0), (-1,-1), 12),
        ('RIGHTPADDING',  (0,0), (-1,-1), 12),
        ('ROUNDEDCORNERS', [4]),
    ]))
    return tbl

def badge(label, bg=ACCENT, fg=WHITE):
    """Small coloured badge/pill."""
    p = ParagraphStyle('bg', fontSize=8, fontName='Helvetica-Bold',
                       textColor=fg, leading=10)
    t = Table([[Paragraph(f'  {label}  ', p)]], colWidths=None)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    return t

def callout(text, kind='tip'):
    """Coloured callout box (tip / note / warning)."""
    cfg = {
        'tip':  (GREEN_BG,  GREEN,  '✔  '),
        'note': (ORANGE_BG, ORANGE, 'ℹ  '),
        'warn': (RED_BG,    RED,    '⚠  '),
    }[kind]
    bg, fg, icon = cfg
    style = ParagraphStyle('co', fontSize=9.5, fontName='Helvetica',
                           textColor=fg, leading=14, leftIndent=4)
    t = Table([[Paragraph(f'{icon}{text}', style)]],
              colWidths=[W - 80])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('BOX', (0,0), (-1,-1), 0.8, fg),
        ('TOPPADDING', (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('ROUNDEDCORNERS', [3]),
    ]))
    return t

def steps_table(rows):
    """Numbered step list rendered as a two-column table."""
    data = []
    for i, text in enumerate(rows, 1):
        num = Paragraph(f'<b>{i}</b>',
                        ParagraphStyle('sn', fontSize=10, textColor=WHITE,
                                       alignment=TA_CENTER, leading=14))
        txt = Paragraph(text,
                        ParagraphStyle('st2', fontSize=10, textColor=INK2,
                                       leading=14, spaceAfter=0))
        data.append([num, txt])

    t = Table(data, colWidths=[22, W - 80 - 30])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (0,-1), ACCENT),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',    (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING',   (0,0), (0,-1), 4),
        ('RIGHTPADDING',  (0,0), (0,-1), 4),
        ('LEFTPADDING',   (1,0), (1,-1), 10),
        ('RIGHTPADDING',  (1,0), (1,-1), 4),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [BG_LIGHT, WHITE]),
        ('BOX',   (0,0), (-1,-1), 0.5, LINE),
        ('GRID',  (0,0), (-1,-1), 0.3, LINE),
    ]))
    return t

def field_table(rows):
    """Two-column Field / Description reference table."""
    hdr = [Paragraph('<b>Field</b>',
                     ParagraphStyle('fh', fontSize=9, textColor=WHITE)),
           Paragraph('<b>Description</b>',
                     ParagraphStyle('fd', fontSize=9, textColor=WHITE))]
    data = [hdr]
    for f, d in rows:
        data.append([
            Paragraph(f'<b>{f}</b>', ParagraphStyle('ff', fontSize=9,
                      fontName='Helvetica-Bold', textColor=NAVY, leading=13)),
            Paragraph(d, ParagraphStyle('fd2', fontSize=9, textColor=INK2, leading=13)),
        ])
    t = Table(data, colWidths=[120, W - 80 - 128])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), ACCENT),
        ('TEXTCOLOR',     (0,0), (-1,0), WHITE),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [WHITE, BG_LIGHT]),
        ('GRID',  (0,0), (-1,-1), 0.4, LINE),
        ('TOPPADDING',    (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING',   (0,0), (-1,-1), 8),
        ('RIGHTPADDING',  (0,0), (-1,-1), 8),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
    ]))
    return t

def hr():
    return HRFlowable(width='100%', thickness=0.5, color=LINE,
                      spaceAfter=8, spaceBefore=4)

def sp(h=6):
    return Spacer(1, h)

def P(text, style='body'):
    return Paragraph(text, S[style])

# ═══════════════════════════════════════════════════════════════════════════
#  CONTENT
# ═══════════════════════════════════════════════════════════════════════════
def build_cover(story):
    """Full navy cover page."""
    story.append(Spacer(1, 120))
    # Logo text
    story.append(Paragraph(
        '<font color="#ffffff">PEPS </font><font color="#5b9bf8">BOM</font>'
        '<font color="#ffffff"> Tool</font>',
        ParagraphStyle('logo', fontSize=40, fontName='Helvetica-Bold',
                       alignment=TA_CENTER, leading=48)))
    story.append(sp(10))
    story.append(Paragraph('User Guide', S['cover_sub']))
    story.append(sp(6))
    story.append(Paragraph('BOM Automation Platform  ·  v2.2', S['cover_ver']))
    story.append(sp(50))
    # Divider line
    story.append(HRFlowable(width='60%', thickness=1, color=HexColor('#2d4a7a'),
                             hAlign='CENTER', spaceAfter=20, spaceBefore=20))
    story.append(Paragraph(
        'A complete reference for all users — from login to downloading<br/>'
        'your final product structure file.',
        ParagraphStyle('cs2', fontSize=12, fontName='Helvetica',
                       textColor=HexColor('#7a9cc8'), alignment=TA_CENTER,
                       leading=18)))
    story.append(sp(60))
    story.append(Paragraph(
        f'PepsiCo India Holdings Pvt. Ltd.  ·  {datetime.date.today().strftime("%B %Y")}',
        S['cover_ver']))
    story.append(PageBreak())


def build_toc(story):
    story.append(sp(20))
    story.append(section_banner('TABLE OF CONTENTS'))
    story.append(sp(12))

    sections = [
        ('1', 'Introduction & System Requirements', ''),
        ('2', 'Logging In', ''),
        ('3', 'Navigation Overview', ''),
        ('4', 'BOM Viewer', ''),
        ('',  '4.1  Searching & Filtering Products', ''),
        ('',  '4.2  Viewing BOM Components', ''),
        ('',  '4.3  Downloading Product Structure', ''),
        ('5', 'Global Replace', ''),
        ('',  '5.1  Replacing a Raw Material', ''),
        ('',  '5.2  Impact Preview & Rollback', ''),
        ('6', 'Product Modifications (BOM Wizard)', ''),
        ('',  '6.1  Step 1 — Approval Gate', ''),
        ('',  '6.2  Step 2 — Entry Point', ''),
        ('',  '6.3  Step 3 — Product Identity', ''),
        ('',  '6.4  Step 4 — Size & Colour Matrix', ''),
        ('',  '6.5  Step 5 — Component Table', ''),
        ('',  '6.6  Step 6 — Formula Coverage Check', ''),
        ('',  '6.7  Step 7 — Generate & Download', ''),
        ('7', 'Nearest BOM', ''),
        ('',  '7.1  Single Code Lookup', ''),
        ('',  '7.2  Bulk Upload', ''),
        ('',  '7.3  Downloading Results (Excel & MDCF)', ''),
        ('8', 'Run History', ''),
        ('9', 'Ramco MDCF', ''),
        ('10','Settings', ''),
        ('11','CMS — Control Mapping System (Admin)', ''),
        ('12','Troubleshooting & Tips', ''),
    ]

    for num, title, pg in sections:
        is_main = bool(num)
        style = S['toc_section'] if is_main else S['toc_item']
        prefix = f'<b>{num}. </b>' if num else '• '
        story.append(Paragraph(f'{prefix}{title}', style))

    story.append(PageBreak())


def build_intro(story):
    story.append(sp(10))
    story.append(section_banner('1.  INTRODUCTION & SYSTEM REQUIREMENTS'))
    story.append(sp(10))
    story.append(P(
        'The <b>PEPS BOM Automation Tool</b> is a web-based platform built specifically '
        'for PepsiCo India\'s mattress manufacturing BOM management. It connects to the '
        'existing macro-driven Excel BOM library (233 registered products, 21 000+ SKUs) '
        'and provides a fast, controlled interface for viewing, editing, creating, and '
        'downloading product-structure files ready for upload into Ramco ERP.'))
    story.append(sp(6))
    story.append(P('<b>What you can do with this tool:</b>'))
    for item in [
        'Browse and search every registered BOM by family, category, size, and colour.',
        'Replace a raw-material item code across <i>all</i> 232 product files in one click (with preview and rollback).',
        'Create new product BOMs through a guided 7-step wizard — with or without a base template.',
        'Map odd-size mattress orders to the nearest standard BOM and download the proxy structure.',
        'Bulk-upload a list of mattress codes with order quantities and get a combined BOM in Excel or MDCF format.',
        'Track every operation in Run History for audit purposes.',
        'Push CreateProductStructure / EditProductStructure files directly to Ramco via the MDCF connector.',
    ]:
        story.append(P(f'• {item}'))

    story.append(sp(10))
    story.append(P('<b>System Requirements</b>', 'h3'))
    story.append(field_table([
        ('Browser',        'Google Chrome or Microsoft Edge (latest version recommended). Firefox supported.'),
        ('Network',        'Must be connected to the same LAN as the server (192.168.0.133).'),
        ('URL',            'https://192.168.0.133:5005'),
        ('Login',          'Username and password provided by your system administrator.'),
        ('SSL Warning',    'On first visit, click Advanced → Proceed to site (unsafe). This is a self-signed internal certificate — it is safe on your company network.'),
        ('Supported Files','Upload files must be .xlsx (MDCF / template) or .xlsm (Macro BOM).'),
    ]))
    story.append(PageBreak())


def build_login(story):
    story.append(sp(10))
    story.append(section_banner('2.  LOGGING IN'))
    story.append(sp(10))
    story.append(P(
        'The tool requires a personal login. Your administrator will provide your '
        'username and password. Tabs visible to you depend on the role assigned to '
        'your account.'))
    story.append(sp(8))
    story.append(P('<b>Login Steps</b>', 'h2'))
    story.append(steps_table([
        'Open your browser and go to <b>https://192.168.0.133:5005</b>.',
        'If you see a security warning, click <b>Advanced</b> and then <b>Proceed to 192.168.0.133 (unsafe)</b>. This only appears on the very first visit.',
        'Enter your <b>Username</b> and <b>Password</b> in the login card.',
        'Click <b>Login</b>. You will be taken directly to the BOM Viewer home screen.',
        'If your credentials are incorrect, an error message appears in red below the form. Contact your administrator to reset your password.',
    ]))
    story.append(sp(10))
    story.append(callout(
        'Your session stays active until you close the browser tab or click Logout '
        '(top-right avatar menu). You do not need to log in again as long as the '
        'tab remains open.', 'tip'))
    story.append(sp(8))
    story.append(callout(
        'Passwords must be at least 6 characters. You can change your own password '
        'at any time from the Settings tab.', 'note'))
    story.append(PageBreak())


def build_navigation(story):
    story.append(sp(10))
    story.append(section_banner('3.  NAVIGATION OVERVIEW'))
    story.append(sp(10))
    story.append(P(
        'After login, the top navigation bar is always visible. It shows only the '
        'tabs your role permits. Click any tab to switch pages instantly — no page '
        'reload is needed.'))
    story.append(sp(8))
    story.append(field_table([
        ('BOM Viewer',         'Search and view the full component list of any registered product. Download the product-structure file.'),
        ('Global Replace',     'Replace a raw-material item code across all 232 product BOM files simultaneously.'),
        ('Product Modifications','7-step BOM wizard — create a new product or update an existing one by uploading a spec file or building manually.'),
        ('Nearest BOM',        'Map an odd-size mattress code to the nearest standard BOM. Supports single-code lookup and bulk Excel upload.'),
        ('Run History',        'Full audit log of every BOM download, replacement, and creation run.'),
        ('Ramco MDCF',         'Configure and push CreateProductStructure / EditProductStructure files to the Ramco ERP server.'),
        ('Settings',           'Change your password, switch light/dark theme, adjust font size and style.'),
        ('CMS',                'Admin-only. Create users, assign roles, manage tab-level access permissions.'),
    ]))
    story.append(sp(8))
    story.append(P(
        'The <b>top-right corner</b> shows the current user name and role badge. '
        'Click the avatar/name to reveal the <b>Logout</b> option. The header also '
        'shows the live product and SKU count (e.g. "v2.2 · 233 products · 21 086 SKUs").'))
    story.append(PageBreak())


def build_bom_viewer(story):
    story.append(sp(10))
    story.append(section_banner('4.  BOM VIEWER'))
    story.append(sp(10))
    story.append(P(
        'The BOM Viewer is the home screen. It lets you drill down into any of the '
        '233 registered products and see the full Bill of Materials — every component, '
        'quantity, UOM, and warehouse code — for any size and colour combination.'))

    story.append(sp(8))
    story.append(P('4.1  Searching & Filtering Products', 'h2'))
    story.append(P(
        'Use the four drop-down filters across the top bar to narrow down products. '
        'Each filter unlocks the next one.'))
    story.append(steps_table([
        '<b>Brand / Family</b> — Select a product family (e.g. "Spg spr", "Org Macro", "Peps Pk"). The list shows all families registered in the BOM library.',
        '<b>Category</b> — Refines by product category within the selected family.',
        '<b>Height</b> — Filter by mattress height in inches (e.g. 06, 08, 10).',
        '<b>Colour</b> — Filter by fabric/colour code (e.g. NL, BW, GR).',
        'The product list in the left panel updates instantly as you change filters.',
        'Click any product card in the list to load its BOM details on the right.',
    ]))

    story.append(sp(8))
    story.append(P('4.2  Viewing BOM Components', 'h2'))
    story.append(P(
        'After clicking a product, the right panel shows the Product Details section '
        'with a full component table. Use the <b>L (Length)</b>, <b>W (Width)</b>, '
        '<b>H (Height)</b>, and <b>Colour</b> selectors at the top of the detail '
        'panel to switch between size variants.'))
    story.append(field_table([
        ('Seq',         'Component sequence number in the BOM.'),
        ('Item Code',   'Raw material or sub-assembly item code as registered in Ramco.'),
        ('Description', 'Item description from the master list.'),
        ('Qty',         'Quantity per mattress unit for the selected size.'),
        ('UOM',         'Unit of measurement (NOS, KG, MTR, etc.).'),
        ('WH',          'Source warehouse code (CBERMFG, CBEFG, etc.).'),
    ]))

    story.append(sp(8))
    story.append(P('4.3  Downloading Product Structure', 'h2'))
    story.append(P(
        'Once you have selected a product and size, scroll to the bottom of the detail '
        'panel and click <b>Download CreateProductStructure</b>. This produces an Excel '
        '(.xlsx) file ready for direct upload into Ramco ERP.'))
    story.append(callout(
        'The downloaded file is a Ramco-compatible CreateProductStructure format. '
        'Do not modify the column headers or sheet names before uploading to Ramco.', 'warn'))
    story.append(PageBreak())


def build_global_replace(story):
    story.append(sp(10))
    story.append(section_banner('5.  GLOBAL REPLACE'))
    story.append(sp(10))
    story.append(P(
        'Global Replace lets you change a raw-material item code across <b>all 232 '
        'product BOM files</b> in one operation. Use this when a component is '
        'superseded, retired, or re-coded in the item master.'))
    story.append(callout(
        '⚠  This operation modifies every BOM file that contains the specified item code. '
        'Always preview the impact before executing. A full snapshot is created automatically '
        'so you can roll back if needed.', 'warn'))

    story.append(sp(8))
    story.append(P('5.1  Replacing a Raw Material', 'h2'))
    story.append(steps_table([
        'Go to the <b>Global Replace</b> tab.',
        'Enter the <b>Old Item Code</b> — the code you want to remove or replace. Press Tab; the tool looks it up in the item master and shows its description.',
        'Enter the <b>New Item Code</b> if you are replacing it. Leave blank if you only want to update the description or quantity without changing the code.',
        'Optionally enter a <b>New Description</b> and <b>New Quantity</b> override.',
        'Click <b>Scan Impact</b>. The right panel shows every product file, BOM row, and current value that will be affected — <i>before</i> any changes are made.',
        'Review the impact table carefully. When satisfied, click <b>Execute Replace</b>.',
        'A success toast confirms how many files were updated. The run is logged in Run History.',
    ]))

    story.append(sp(8))
    story.append(P('5.2  Impact Preview & Rollback', 'h2'))
    story.append(P(
        'The <b>Impact Preview</b> panel on the right shows a colour-coded table: '
        'every product that contains the old item code, with the row that will change '
        'highlighted in orange. The <b>Rollback</b> card at the bottom lists all '
        'previous snapshots — click <b>Restore</b> next to any snapshot to undo '
        'that entire replace operation.'))
    story.append(callout(
        'Snapshots are kept indefinitely. You can roll back to any previous state at '
        'any time, even weeks after the original replace.', 'tip'))
    story.append(PageBreak())


def build_product_wizard(story):
    story.append(sp(10))
    story.append(section_banner('6.  PRODUCT MODIFICATIONS — BOM WIZARD'))
    story.append(sp(10))
    story.append(P(
        'The Product Modifications wizard guides you through creating a new product '
        'BOM in 7 structured steps. Every new BOM requires a management sign-off '
        'reference before it can be generated.'))

    story.append(sp(6))
    story.append(P('Step 1 — Approval Gate', 'h2'))
    story.append(P('A sign-off reference is mandatory before any BOM can be created. This is logged permanently against the run record.'))
    story.append(field_table([
        ('Sign-off Reference No.', 'Internal approval reference (e.g. SCR-2026-041).'),
        ('Approval Date',          'Date the approval was granted.'),
        ('Approving Authority',    'Name and designation of the approver.'),
    ]))
    story.append(P('Click <b>Next →</b> to proceed.'))

    story.append(sp(6))
    story.append(P('Step 2 — Entry Point', 'h2'))
    story.append(P('Choose how to build the BOM:'))
    story.append(field_table([
        ('Upload Spec File',       'Upload a Macro (.xlsm) or MDCF (.xlsx) spec file. If the product already exists in the library it will be updated automatically.'),
        ('Manual / Clone Existing','Select any registered product as a template. All component rules are pre-filled from the base product — edit as needed.'),
    ]))

    story.append(sp(6))
    story.append(P('Step 3 — Product Identity', 'h2'))
    story.append(field_table([
        ('Product Name',      'Full product name (e.g. Peps Springkoil Bonnell Normal).'),
        ('Item Code Prefix',  'The prefix used in all mattress codes for this product (e.g. PEPSSPKBNLNM).'),
        ('Warehouse Code',    'Parent warehouse: CBEFG (Finished Goods), CBESFG (Semi-Finished), CBEWIP (WIP).'),
        ('Scenario',         'Intended use scenario for the BOM.'),
    ]))

    story.append(sp(6))
    story.append(P('Step 4 — Size & Colour Matrix', 'h2'))
    story.append(P(
        'Add every size and colour combination the product is manufactured in. '
        'The system uses this matrix to generate individual BOM rows for each SKU. '
        'Click <b>Add Row</b> to add a new L × W × H × Colour combination.'))

    story.append(sp(6))
    story.append(P('Step 5 — Component Table', 'h2'))
    story.append(P(
        'Define the raw materials and sub-assemblies. For each component, enter the '
        'item code, description, UOM, warehouse, and the quantity formula. Formulas '
        'can reference the size dimensions (L, W, H) to compute per-unit quantities '
        'automatically across all size variants.'))
    story.append(callout(
        'Click the Formula Guide (expandable panel on the right side) for a full '
        'reference of supported formula syntax and examples.', 'tip'))

    story.append(sp(6))
    story.append(P('Step 6 — Formula Coverage Check', 'h2'))
    story.append(P(
        'The tool validates that every formula produces a valid, positive result for '
        'every size in your matrix. Rows with issues are highlighted in red. Fix any '
        'errors before proceeding — the generator will not run with formula errors.'))

    story.append(sp(6))
    story.append(P('Step 7 — Generate & Download', 'h2'))
    story.append(steps_table([
        'Click <b>Generate BOM</b>. The server builds the full product-structure file for all size/colour combinations.',
        'A preview summary shows the total number of BOM rows generated.',
        'Click <b>Download CreateProductStructure</b> to save the Excel file ready for Ramco upload.',
        'The run is logged in Run History with the approval reference, operator name, and timestamp.',
    ]))
    story.append(PageBreak())


def build_nearest_bom(story):
    story.append(sp(10))
    story.append(section_banner('7.  NEAREST BOM'))
    story.append(sp(10))
    story.append(P(
        'The Nearest BOM module handles <b>odd-size mattress orders</b> — sizes that '
        'do not have their own BOM in the library. The tool finds the next standard '
        'size equal to or greater than the requested dimensions (independently for '
        'Length and Width), then uses that standard BOM as a proxy.'))
    story.append(callout(
        'The tool rounds L and W up independently to the next available standard size. '
        'If the requested size is larger than the maximum defined size in the library '
        'for that product, the tool automatically uses the maximum available BOM as the '
        'nearest reference and clearly flags this with an "above maximum size" note.', 'note'))

    story.append(sp(8))
    story.append(P('7.1  Single Code Lookup', 'h2'))
    story.append(steps_table([
        'Go to the <b>Nearest BOM</b> tab.',
        'Under <b>Find Nearest Standard BOM</b>, select the <b>Product Family</b> from the drop-down.',
        'Select the <b>Colour</b> for the product.',
        'Enter the <b>Length (in)</b>, <b>Width (in)</b>, and select the <b>Height (in)</b>.',
        'Enter the <b>Odd-Size Item Code</b> (the code that will appear in the output file — this is your customer\'s actual item code).',
        'Click <b>Find Nearest BOM</b>.',
        'The Mapping Result panel on the right shows the original code → nearest standard code arrow, the note explaining the rounding, and the full BOM component table.',
        'Click <b>Download CreateProductStructure</b> to download the proxy BOM for this code.',
    ]))

    story.append(sp(8))
    story.append(P('7.2  Bulk Upload', 'h2'))
    story.append(P(
        'Use bulk upload when you have multiple odd-size orders at once — for example '
        'an entire purchase order with different mattress codes and quantities.'))
    story.append(steps_table([
        'Click <b>Download upload template</b> (link in the Bulk Upload card or the Download Template button in the results area) to get the Excel input file.',
        'Fill in the template: Column A = S.No, Column B = Mattress Code, Column C = Qty. Each row is one order line.',
        'Save the file and click <b>Choose File</b> in the Bulk Upload card.',
        'Click <b>Process Upload</b>. The server processes each code — auto-detecting the product family and colour from the code prefix.',
        'The Bulk Upload Preview panel shows each matched code with its order quantity, the nearest standard BOM used, and the full component table.',
        '<b>Unmatched codes</b> (if any) are listed separately below the results with the specific reason (e.g. "prefix not recognised", "above maximum size — kindly check and update").',
    ]))

    story.append(sp(8))
    story.append(P('7.3  Downloading Results (Excel & MDCF)', 'h2'))
    story.append(P(
        'After a successful bulk upload, the download strip becomes active with '
        'three options:'))
    story.append(field_table([
        ('Download Upload Template', 'Re-download the blank input template at any time.'),
        ('Download Excel',           'Excel file with all matched BOMs combined — one sheet per matched code, with order quantities shown separately.'),
        ('Download MDCF (EditProductStructure)', 'Ramco-compatible EditProductStructure file for all matched codes, ready for direct upload into Ramco ERP.'),
    ]))
    story.append(callout(
        'The same codes uploaded together are de-duplicated — if the same mattress code '
        'appears more than once in your upload list, its BOM is shown only once but the '
        'quantities are summed.', 'tip'))
    story.append(PageBreak())


def build_run_history(story):
    story.append(sp(10))
    story.append(section_banner('8.  RUN HISTORY'))
    story.append(sp(10))
    story.append(P(
        'Every operation performed in the tool is automatically logged in Run History. '
        'This provides a complete audit trail for compliance and troubleshooting.'))
    story.append(sp(8))
    story.append(P('What is logged:', 'h3'))
    story.append(field_table([
        ('BOM Download',    'Every product-structure file downloaded from the BOM Viewer, including the product name, size, colour, and the operator\'s username.'),
        ('Global Replace',  'Every raw-material replacement — old code, new code, number of files affected, and the snapshot reference for rollback.'),
        ('New Product',     'Every BOM created through the Product Modifications wizard, with the sign-off reference and approver details.'),
        ('Nearest BOM',     'Single-code lookups and bulk-upload runs, showing the input codes and nearest matches used.'),
    ]))
    story.append(sp(8))
    story.append(P('How to use Run History:', 'h2'))
    story.append(steps_table([
        'Click the <b>Run History</b> tab.',
        'Use the filter chips at the top to show only a specific type of run (BOM Download, Global Replace, New Product, etc.).',
        'Click any run row to expand it and see the full details — components generated, approval references, operator, and timestamp.',
        'For Global Replace runs, the snapshot ID is shown — use this in the Global Replace tab\'s Rollback section if you need to undo that operation.',
    ]))
    story.append(PageBreak())


def build_ramco_mdcf(story):
    story.append(sp(10))
    story.append(section_banner('9.  RAMCO MDCF'))
    story.append(sp(10))
    story.append(P(
        'The Ramco MDCF tab provides the connection settings to push product-structure '
        'files directly to your Ramco ERP application server. Once configured, you can '
        'upload files from within the tool without saving them to your desktop first.'))
    story.append(sp(8))
    story.append(P('Connection Attributes', 'h2'))
    story.append(field_table([
        ('App Server URL',  'The URL or IP address of your Ramco application server.'),
        ('Port',            'The port Ramco listens on (e.g. 8080).'),
        ('Username',        'Your Ramco ERP login username.'),
        ('Password',        'Your Ramco ERP login password (stored encrypted on the server).'),
        ('Company Code',    'The Ramco company/entity code (e.g. CBEF, CBESFG).'),
        ('Module',          'Target Ramco module — typically "Manufacturing".'),
    ]))
    story.append(sp(8))
    story.append(callout(
        'Connection settings are saved per-user. Only users with the Ramco MDCF tab '
        'enabled in their role can access this section.', 'note'))
    story.append(callout(
        'If you receive a connection error, verify that your PC can reach the Ramco '
        'server IP address and that the Ramco service is running.', 'warn'))
    story.append(PageBreak())


def build_settings(story):
    story.append(sp(10))
    story.append(section_banner('10.  SETTINGS'))
    story.append(sp(10))
    story.append(P('The Settings tab lets you personalise your account and the tool\'s appearance.'))
    story.append(sp(8))

    story.append(P('Change Password', 'h2'))
    story.append(steps_table([
        'Go to <b>Settings</b> tab.',
        'Under <b>Change My Password</b>, enter your current password.',
        'Enter the new password (minimum 6 characters) and confirm it.',
        'Click <b>Update Password</b>. You will need to use the new password at your next login.',
    ]))

    story.append(sp(8))
    story.append(P('Theme', 'h2'))
    story.append(P(
        'Toggle between <b>Light</b> and <b>Dark</b> theme. Your preference is saved '
        'in your browser and restores automatically on next visit.'))

    story.append(sp(8))
    story.append(P('Font Size & Style', 'h2'))
    story.append(P(
        'Adjust the interface font size (Small / Normal / Large) and font style '
        '(DM Sans, Inter, Roboto, System UI) to suit your display setup. '
        'Changes take effect immediately.'))
    story.append(PageBreak())


def build_cms(story):
    story.append(sp(10))
    story.append(section_banner('11.  CMS — CONTROL MAPPING SYSTEM  (Admin Only)'))
    story.append(sp(10))
    story.append(P(
        'The CMS tab is visible only to users with the Admin or Developer role. '
        'It manages user accounts, passwords, roles, and tab-level access permissions.'))

    story.append(sp(8))
    story.append(P('Creating a New User', 'h2'))
    story.append(steps_table([
        'Go to the <b>CMS</b> tab.',
        'In the <b>Add New User</b> card, enter a <b>Username</b> (must be unique).',
        'Enter a <b>Password</b> (minimum 6 characters).',
        'Select a <b>Role</b> from the drop-down (User, Admin, Developer, or any custom role).',
        'Under <b>Allowed Tabs</b>, tick the tabs this user should have access to: BOM Viewer, Global Replace, Product Modifications, Nearest BOM, Run History, Ramco MDCF, Settings, CMS.',
        'Click <b>Create User</b>. The user can log in immediately with the credentials you set.',
    ]))

    story.append(sp(8))
    story.append(P('Managing Existing Users', 'h2'))
    story.append(P(
        'The <b>Existing Users</b> section lists every account with their role badge, '
        'last login date, and hashed password. Available actions:'))
    story.append(field_table([
        ('Edit User',      'Change the user\'s username, password, role, or allowed tabs.'),
        ('Deactivate',     'Disable the account without deleting it. The user cannot log in while deactivated.'),
        ('Copy Password Hash', 'Copy the bcrypt hash to clipboard (useful for manual migration).'),
    ]))

    story.append(sp(8))
    story.append(P('Roles & Hierarchy', 'h2'))
    story.append(field_table([
        ('Developer', 'Full system access + developer tools. All tabs enabled.'),
        ('Admin',     'All tabs enabled. Can manage users and roles.'),
        ('User',      'Standard user. Tabs enabled according to individual Allowed Tabs settings.'),
        ('Custom',    'Create custom roles with specific tab sets — assign to multiple users at once.'),
    ]))
    story.append(PageBreak())


def build_troubleshooting(story):
    story.append(sp(10))
    story.append(section_banner('12.  TROUBLESHOOTING & TIPS'))
    story.append(sp(10))

    issues = [
        ('Browser shows "Not Secure" warning',
         'This is expected — the tool uses a self-signed SSL certificate on your internal network. '
         'Click Advanced → Proceed to 192.168.0.133 (unsafe). You only need to do this once per browser. '
         'The connection is still encrypted within your LAN.'),
        ('Login page says "Invalid credentials"',
         'Double-check your username and password (case-sensitive). '
         'Contact your administrator to reset your password from the CMS tab.'),
        ('Bulk upload shows "No match found" for a code',
         'Check that the mattress code prefix is registered in the BOM library (e.g. PEPSPKORGBW, SKSGSR). '
         'Codes with fictional or test prefixes will not match. If the size is above the maximum defined size, '
         'the tool will show the maximum available BOM as a reference instead.'),
        ('"Above the defined maximum size" message',
         'The requested mattress size (L × W) is larger than the largest standard BOM available for that '
         'product family. The tool automatically falls back to the maximum-size BOM as the nearest reference. '
         'The download will still work — review the note in the output file and confirm with your product team '
         'if a new standard size needs to be added to the library.'),
        ('Download button does nothing',
         'Ensure your browser allows pop-ups and file downloads from 192.168.0.133. '
         'Check the browser\'s download manager — the file may have been blocked. '
         'Try a different browser (Chrome recommended).'),
        ('Page loads slowly or times out on first use',
         'The server pre-loads all BOM files on startup (~30 seconds). If you open the tool '
         'immediately after a server restart, wait 30–60 seconds and refresh the page.'),
        ('Global Replace affected fewer files than expected',
         'Use the Impact Preview (click Scan Impact) before executing to see the exact list '
         'of products that contain the old item code. The item code must match exactly — '
         'check for trailing spaces or capitalisation differences.'),
        ('CMS tab not visible',
         'Only Admin and Developer roles can see the CMS tab. Contact your administrator '
         'to enable this for your account if needed.'),
    ]

    for title, desc in issues:
        story.append(KeepTogether([
            P(f'<b>{title}</b>', 'h3'),
            P(desc),
            sp(4),
        ]))

    story.append(sp(10))
    story.append(section_banner('QUICK REFERENCE', color=INK2))
    story.append(sp(8))
    story.append(field_table([
        ('Tool URL',               'https://192.168.0.133:5005'),
        ('Login help',             'Contact your administrator — CMS tab'),
        ('Password change',        'Settings tab → Change My Password'),
        ('Add / edit users',       'CMS tab (Admin access required)'),
        ('Bulk upload template',   'Nearest BOM tab → Download upload template link, or the Download Template button in the results area'),
        ('Roll back a replacement','Global Replace tab → Rollback section → Restore snapshot'),
        ('Audit trail',            'Run History tab — all operations logged with username and timestamp'),
        ('Ramco upload',           'Ramco MDCF tab — configure connection, then push from BOM Viewer or Nearest BOM download'),
    ]))
    story.append(sp(20))
    story.append(Paragraph(
        '<font color="#7a8a9a">PEPS BOM Automation Tool v2.2  ·  '
        'PepsiCo India Holdings Pvt. Ltd.  ·  Confidential — Internal Use Only</font>',
        ParagraphStyle('footer', fontSize=8, alignment=TA_CENTER, leading=12)))


# ═══════════════════════════════════════════════════════════════════════════
#  COVER PAGE CANVAS (full navy background)
# ═══════════════════════════════════════════════════════════════════════════
class CoverCanvas:
    def __call__(self, canv, doc):
        if doc.page == 1:
            canv.saveState()
            canv.setFillColor(NAVY)
            canv.rect(0, 0, W, H, fill=1, stroke=0)
            # Accent stripe on left
            canv.setFillColor(ACCENT)
            canv.rect(0, 0, 8, H, fill=1, stroke=0)
            # Bottom bar
            canv.setFillColor(HexColor('#0a1220'))
            canv.rect(0, 0, W, 50, fill=1, stroke=0)
            canv.setFillColor(HexColor('#5b9bf8'))
            canv.setFont('Helvetica-Bold', 9)
            canv.drawString(20, 18,
                f'PepsiCo India Holdings Pvt. Ltd.  ·  Confidential — Internal Use Only  ·  {datetime.date.today().strftime("%B %Y")}')
            canv.restoreState()
        else:
            PageTemplate()(canv, doc)


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════
def main():
    doc = SimpleDocTemplate(
        OUT_PATH,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=55,
        bottomMargin=35,
        title='PEPS BOM Tool — User Guide',
        author='PEPS BOM Automation Tool v2.2',
        subject='User Guide',
    )

    story = []

    # Cover (no header/footer on page 1)
    build_cover(story)
    build_toc(story)
    build_intro(story)
    build_login(story)
    build_navigation(story)
    build_bom_viewer(story)
    build_global_replace(story)
    build_product_wizard(story)
    build_nearest_bom(story)
    build_run_history(story)
    build_ramco_mdcf(story)
    build_settings(story)
    build_cms(story)
    build_troubleshooting(story)

    doc.build(story, onFirstPage=CoverCanvas(), onLaterPages=PageTemplate())
    print('PDF saved -> ' + OUT_PATH)


if __name__ == '__main__':
    main()
