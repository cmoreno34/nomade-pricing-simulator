"""
Build a Word document for the Nomade Vans case, adapted to the interactive
pricing simulator. Run: `python build_case_docx.py` (requires python-docx).
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SIMULATOR_URL = "https://cmoreno34.github.io/nomade-pricing-simulator/"

ORANGE = RGBColor(0xF9, 0x73, 0x16)
BLUE   = RGBColor(0x25, 0x63, 0xEB)
GREY   = RGBColor(0x64, 0x74, 0x8B)
BLACK  = RGBColor(0x0F, 0x17, 0x2A)


def set_cell_bg(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_fill)
    tc_pr.append(shd)


def h(doc, text, level=1, color=BLACK):
    p = doc.add_heading(level=level)
    run = p.add_run(text)
    run.font.color.rgb = color
    return p


def para(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(text)
    return p


def numbered(doc, text):
    p = doc.add_paragraph(style='List Number')
    p.add_run(text)
    return p


def answer_box(doc, lines=6):
    """Draws a light-grey box for students to write their answer in."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = True
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, 'F8FAFC')
    cell.paragraphs[0].add_run('\n' * (lines - 1)).italic = True
    run = cell.paragraphs[0].runs[0]
    run.font.color.rgb = GREY
    doc.add_paragraph()


def build():
    doc = Document()

    # default style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ---------- Cover ----------
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run('CASE STUDY')
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = GREY

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('Nomade Vans — Pricing under Competition and Psychological Factors')
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = BLUE

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run('Interactive simulator edition — student workbook')
    r.italic = True
    r.font.size = Pt(12)
    r.font.color.rgb = GREY

    doc.add_paragraph()
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run('© César Moreno Pascual PhD — v1.0\n').italic = True
    meta.add_run('Simulator: ').italic = True
    link = meta.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE
    link.underline = True

    doc.add_page_break()

    # ---------- 1. Case background ----------
    h(doc, '1. Case background', level=1)
    para(doc,
         "Nomade Vans is a young Spanish start-up that rents converted camper vans "
         "to leisure travellers. After running a small fleet successfully for two "
         "seasons, the founders now want to formalise pricing for the coming year. "
         "Three direct competitors already operate in the same geographical market "
         "with published prices: Further VAN, People Camper and Ocean Vans.")
    para(doc,
         "The commercial team has three decisions to make for each rental period "
         "(day / 2-night weekend / 6-night week) and for each vehicle trim "
         "(Standard and Premium):")
    bullet(doc, "How much should we charge?")
    bullet(doc, "Where should we sit relative to the competition?")
    bullet(doc, "How should the number be presented (charm pricing, good-better-best…)?")
    para(doc,
         "Last spring we ran a quantitative survey with 90 respondents to measure "
         "willingness-to-pay (WTP). The survey data, together with a cost model, a "
         "competitor price sheet and the psychological-factors framework from the "
         "technical notes, are integrated into the interactive simulator you will "
         "use throughout this case.")

    # ---------- 2. Company ----------
    h(doc, '2. The company — Nomade Vans', level=1)
    para(doc,
         "Nomade Vans offers two trims: Standard (base conversion, proven kitchen + "
         "bed package) and Premium (solar roof, outdoor shower, expanded bed, "
         "premium upholstery). Each trim is rented at three usage tiers: a single "
         "day, a 2-night weekend and a 6-night week.")

    # ---------- 3. Cost structure ----------
    h(doc, '3. Cost structure', level=1)
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = 'Light Grid Accent 1'
    headers = ['Item', 'Standard', 'Premium']
    for i, txt in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    rows = [
        ('Variable cost per rental (€)', '30',     '30'),
        ('Fixed costs per year (€)',     '69,750', '77,750'),
        ('Annual demand potential (D)',  '5,000',  '5,000'),
    ]
    for i, (a, b, c) in enumerate(rows, 1):
        tbl.rows[i].cells[0].text = a
        tbl.rows[i].cells[1].text = b
        tbl.rows[i].cells[2].text = c
    doc.add_paragraph()
    para(doc,
         "Annual demand potential is scaled by period: ×1.0 for daily rentals, "
         "×0.365 for weekends (≈ 182 weekend rentals per year), ×0.0743 for "
         "weekly rentals (≈ 52 per year).", italic=True, size=10)

    # ---------- 4. Competitors ----------
    h(doc, '4. Competitors — reference prices', level=1)
    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = 'Light Grid Accent 1'
    headers = ['Competitor', 'Daily Std / Prem (€)', 'Weekend Std / Prem (€)', 'Weekly Std / Prem (€)']
    for i, txt in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    comp_rows = [
        ('Further VAN',   '85 / 95',   '80 / 90',   '60 / 70'),
        ('People Camper', '105 / 115', '100 / 110', '75 / 85'),
        ('Ocean Vans',    '98 / 129',  '93 / 123',  '68 / 95'),
    ]
    for i, row in enumerate(comp_rows, 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()
    para(doc,
         "These prices are the defaults preloaded in the simulator; you can change "
         "them on the Position & prices tab if you have fresher intelligence.", italic=True, size=10)

    # ---------- 5. Survey methodology ----------
    h(doc, '5. The WTP survey', level=1)
    para(doc,
         "We asked 90 Spanish residents (aged 20–60, representative of the target "
         "segment of leisure travellers) about their demographic profile and about "
         "the maximum price they would be willing to pay for a Nomade Vans camper, "
         "in six different usage scenarios. The simulator uses the aggregated "
         "price-frequency table as the basis of its demand curve.")

    # ---------- 6. 13 Questions ----------
    h(doc, '6. The 13 survey questions — please answer them yourself', level=1)
    para(doc,
         "Before you analyse the aggregated data, answer the 13 questions as if "
         "you were one of the 90 respondents. This exercise builds empathy with "
         "the customer and helps you detect possible biases in the aggregated "
         "data.", italic=True)

    questions = [
        ('Q1',  'Gender (Male / Female / Other):'),
        ('Q2',  'Age range (<20, 20–30, 30–40, 40–50, >50):'),
        ('Q3',  'Civil status (Single / Married / Divorced / Widowed):'),
        ('Q4',  'Education level (Primary / Secondary / University / Master or PhD):'),
        ('Q5',  'Employment status (Employed / Self-employed / Unemployed / Retired / Student):'),
        ('Q6',  'Leisure activities you enjoy (mark all that apply: Travel, Read, Sports, Adventure, Family time, Other):'),
        ('Q7',  'Days of vacation per year (<3, 3–5, 5–7, >7):'),
        ('Q8',  'Maximum price you would pay for a Standard Nomade Vans / day (€):'),
        ('Q9',  'Maximum price you would pay for a Standard Nomade Vans / weekend — 2 nights (€):'),
        ('Q10', 'Maximum price you would pay for a Standard Nomade Vans / week — 6 nights (€):'),
        ('Q11', 'Maximum price you would pay for a Premium Nomade Vans / day (€):'),
        ('Q12', 'Maximum price you would pay for a Premium Nomade Vans / weekend — 2 nights (€):'),
        ('Q13', 'Maximum price you would pay for a Premium Nomade Vans / week — 6 nights (€):'),
    ]
    for qid, text in questions:
        p = doc.add_paragraph()
        run = p.add_run(f'{qid}. ')
        run.bold = True
        p.add_run(text)
        answer_box(doc, lines=2)

    # ---------- 7. Using the simulator ----------
    doc.add_page_break()
    h(doc, '7. Using the interactive simulator', level=1)
    para(doc, 'Open: ', bold=False)
    p = doc.paragraphs[-1]
    link = p.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE
    link.underline = True
    para(doc, 'Five-step walkthrough:', bold=True)
    steps = [
        'Pick a rental period (Daily / Weekend / Week) and a version (Standard or Premium) at the top.',
        'Go to Competition-based. On the Profit curve, drag the green "Our price" line — KPIs for profit, demand, revenue and lost-vs-peak update live. The three coloured dashed lines are Further VAN, People Camper and Ocean Vans, so you can read at each competitor\'s x what our profit would be at their price.',
        'Go to Psychological factors. Try a three-tier menu (anchoring), a charm-pricing test (€80 vs €79), the Prospect-Theory value function, and a reference-price split.',
        'Open Position & prices. Drag on the chart or type the numeric price. Add a free-text analyst note.',
        'Open Download Excel. One click produces a .xlsx with the 13 questions, aggregated answers, your position and all the simulator charts as images.',
    ]
    for s in steps:
        numbered(doc, s)

    # ---------- 8. Formal questions ----------
    doc.add_page_break()
    h(doc, '8. Formal questions — your deliverables', level=1)
    para(doc,
         'Use the simulator to answer. For every question, take a screenshot or '
         'download the Excel and attach it as evidence.',
         italic=True)

    fq = [
        ('Q8.1 — WTP & optimum',
         'For each combination (period × version), identify the peak-profit price. '
         'Is it closer to the top or the bottom of the WTP distribution? Why?'),
        ('Q8.2 — Competitive positioning',
         'Where does Nomade sit vs Further VAN, People Camper and Ocean Vans on each '
         'of the three periods? Drag the green line to a competitor\'s price using the '
         'Match-competitor button. What would our profit be? What would the acceptance '
         'rate be?'),
        ('Q8.3 — Move up or move down?',
         'On the Profit curve, is there a competitor whose price sits clearly above '
         'our current peak (i.e., moving up toward them would reduce profit)? And a '
         'competitor below our peak? Justify whether to move up, match, or stay.'),
        ('Q8.4 — Anchoring',
         'Design a three-tier menu (basic / middle / premium) for the Weekend period. '
         'Where do you place the middle to maximise acceptance of the target price?'),
        ('Q8.5 — Charm pricing',
         'Test €80 vs €79 (Daily Standard). How many additional rentals does the +7% '
         'perception uplift buy? Does profit increase? When is charm pricing counter-'
         'productive?'),
        ('Q8.6 — Prospect Theory',
         'If you are forced to raise prices next season by €10, how would you frame it '
         'to customers so that it feels less like a loss? Use the value function as '
         'argument.'),
        ('Q8.7 — Reference price',
         'Set the reference price at the cheapest competitor. What share of respondents '
         'are above? What does that share mean for Nomade\'s positioning strategy?'),
        ('Q8.8 — Final recommendation',
         'Propose a price for each period × version. Write the analyst note inside the '
         'simulator, download the Excel, and attach a one-paragraph justification per '
         'decision.'),
    ]
    for qtitle, qtext in fq:
        p = doc.add_paragraph()
        r = p.add_run(qtitle)
        r.bold = True
        r.font.color.rgb = ORANGE
        doc.add_paragraph(qtext)
        answer_box(doc, lines=8)

    # ---------- 9. Evaluation ----------
    h(doc, '9. Evaluation criteria', level=1)
    bullet(doc, "Rigour — use of the simulator (drag evidence, screenshots).")
    bullet(doc, "Integration — competition + WTP + psychological factors combined into a coherent recommendation.")
    bullet(doc, "Clarity — clean writing, numbers justified, trade-offs explicit.")
    bullet(doc, "Attached Excel export with the final position included.")

    # ---------- Deliverables footer ----------
    doc.add_page_break()
    h(doc, 'Deliverables', level=1)
    numbered(doc, 'This workbook (Word), with every answer box filled in.')
    numbered(doc, 'The Excel export (.xlsx) downloaded from the simulator with your final position.')
    numbered(doc, 'A one-page executive summary of your price recommendation.')

    out = 'Nomade_Vans_Case_Study_Simulator_Edition.docx'
    doc.save(out)
    print('Saved:', out)


if __name__ == '__main__':
    build()
