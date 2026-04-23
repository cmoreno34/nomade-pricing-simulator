"""
Build the Word workbook for the Nomade Vans case — aligned to Nomade v3.0
(4 formal questions, strategic 4-scenario framework, competitor attributes,
daily-only competitor data disclaimer). The interactive simulator covers the
experimentation; this workbook is the paper reference.
"""
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SIMULATOR_URL = "https://cmoreno34.github.io/nomade-pricing-simulator/"

ORANGE = RGBColor(0xF9, 0x73, 0x16)
BLUE   = RGBColor(0x25, 0x63, 0xEB)
GREY   = RGBColor(0x64, 0x74, 0x8B)
BLACK  = RGBColor(0x0F, 0x17, 0x2A)
AMBER  = RGBColor(0x92, 0x40, 0x0E)


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


def para(doc, text, bold=False, italic=False, size=11, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(text)
    return p


def numbered(doc, text):
    p = doc.add_paragraph(style='List Number')
    p.add_run(text)
    return p


def note_box(doc, text, fill='FEF3C7', title='⚠ Data note', title_color=AMBER):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, fill)
    p = cell.paragraphs[0]
    r = p.add_run(title + '  ')
    r.bold = True; r.font.color.rgb = title_color
    p.add_run(text)
    doc.add_paragraph()


def answer_box(doc, lines=10):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, 'F8FAFC')
    p = cell.paragraphs[0]
    p.add_run('\n' * (lines - 1))
    doc.add_paragraph()


def make_header_row(tbl, headers):
    for i, txt in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        cell.text = txt
        for r in cell.paragraphs[0].runs:
            r.bold = True


def build():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ---------- Cover ----------
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('INTERACTIVE PRICING STRATEGY — CASE STUDY')
    r.bold = True; r.font.size = Pt(13); r.font.color.rgb = GREY

    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('Nomade Vans — Pricing under Competition and Psychological Factors')
    r.bold = True; r.font.size = Pt(22); r.font.color.rgb = BLUE

    sub = doc.add_paragraph(); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run('Simulator-driven edition · aligned to Nomade v3.0 — student workbook')
    r.italic = True; r.font.size = Pt(12); r.font.color.rgb = GREY

    doc.add_paragraph()
    meta = doc.add_paragraph(); meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run('© César Moreno Pascual PhD — based on Nomade v2.0 / v3.0\n').italic = True
    meta.add_run('Interactive simulator: ').italic = True
    link = meta.add_run(SIMULATOR_URL); link.font.color.rgb = BLUE; link.underline = True

    doc.add_page_break()

    # ---------- 1. Company summary ----------
    h(doc, '1. Company summary')
    para(doc,
         "Nomade Vans is a NEW entrant in the campervan rental market, specialised "
         "in sustainable and highly customisable van designs with fast delivery "
         "(within 1 month). Despite its superior features it faces the classic "
         "challenge of low brand awareness while competing against well-established "
         "incumbents.")
    tbl = doc.add_table(rows=6, cols=3); tbl.style = 'Light Grid Accent 1'
    make_header_row(tbl, ['Attribute', 'Nomade', 'Market average'])
    for i, row in enumerate([
        ('Market position',     'NEW ENTRANT',  'Established (8–9/10)'),
        ('Sustainability',      '10/10',        '4–6/10'),
        ('Customisation',       '10/10',        '3–6/10'),
        ('Brand recognition',   '3/10',         '8–9/10'),
        ('Service quality',     '7/10',         '7–9/10'),
    ], 1):
        for j, v in enumerate(row):
            tbl.rows[i].cells[j].text = v
    doc.add_paragraph()

    # ---------- 2. Data note ----------
    note_box(doc,
             "Competitor prices are only published as DAILY tariffs. Weekend "
             "and weekly rates in the simulator are ESTIMATES aggregated from "
             "the daily tariff (industry convention: weekend ≈ 5–10 % per-day "
             "discount, week ≈ 20–30 % total discount). Treat them as reference "
             "scenarios. Nomade WTP data itself is real (95 respondents).",
             title='⚠ Data note — incomplete dataset (v3.0)')

    # ---------- 3. Data available ----------
    h(doc, '2. Data available')

    # 2.1 Standard survey
    h(doc, '2.1 Survey — Standard camper (WTP)', level=2)
    for lbl, prices, resps in [
        ('Daily rentals',            [50,60,70,80,90,100,110,120,130],  [17,14,10,16,12,19,2,4,1]),
        ('Weekend (2-night) rentals',[80,100,120,140,160,180,200],      [12,7,18,13,11,15,11]),
        ('Weekly (7-night) rentals', [240,300,360,420,480,540,600,660,720],[14,10,10,15,15,8,10,3,6]),
    ]:
        para(doc, lbl, bold=True)
        tbl = doc.add_table(rows=2, cols=len(prices)+1); tbl.style = 'Light Grid Accent 1'
        tbl.rows[0].cells[0].text = 'Price (€)'
        tbl.rows[1].cells[0].text = 'Responses'
        for i, p in enumerate(prices, 1):
            tbl.rows[0].cells[i].text = str(p)
            tbl.rows[1].cells[i].text = str(resps[i-1])
        doc.add_paragraph()

    # 2.2 Premium survey
    h(doc, '2.2 Survey — Premium camper (WTP)', level=2)
    for lbl, prices, resps in [
        ('Daily rentals',            [60,70,80,90,100,110,120,130,140,150,160,180,200], [10,5,12,8,14,12,19,5,3,4,1,1,1]),
        ('Weekend (2-night) rentals',[120,140,160,180,200,220,240,260,280],              [10,6,13,8,22,2,18,4,2]),
        ('Weekly (7-night) rentals', [360,420,480,540,600,660,720,780,840],              [14,7,4,13,17,8,13,7,3]),
    ]:
        para(doc, lbl, bold=True)
        tbl = doc.add_table(rows=2, cols=len(prices)+1); tbl.style = 'Light Grid Accent 1'
        tbl.rows[0].cells[0].text = 'Price (€)'
        tbl.rows[1].cells[0].text = 'Responses'
        for i, p in enumerate(prices, 1):
            tbl.rows[0].cells[i].text = str(p)
            tbl.rows[1].cells[i].text = str(resps[i-1])
        doc.add_paragraph()

    # 2.3 Cost structure
    h(doc, '2.3 Cost structure', level=2)
    tbl = doc.add_table(rows=4, cols=4); tbl.style = 'Light Grid Accent 1'
    make_header_row(tbl, ['Item', 'Daily', 'Weekend', 'Weekly'])
    for i, row in enumerate([
        ('Variable cost (€/day)',       '30', '30', '30'),
        ('Fixed cost (€/year)',         '69,750 (Std) / 77,750 (Prem)', '—', '—'),
        ('Average demand (rentals/yr)', '5,000', '1,825', '521.43'),
    ], 1):
        for j, v in enumerate(row):
            tbl.rows[i].cells[j].text = v
    doc.add_paragraph()

    # 2.4 Competitors (with attributes)
    h(doc, '2.4 Competitor data (daily tariffs only)', level=2)
    tbl = doc.add_table(rows=5, cols=7); tbl.style = 'Light Grid Accent 1'
    make_header_row(tbl, ['Competitor','Std (€/d)','Prem (€/d)','Gap (€)','Establishment 0–10','Brand 0–10','Sustainability 0–10'])
    for i, row in enumerate([
        ('Further VAN',   '85', '95', '10', '8', '8', '5'),
        ('People Camper', '105','115','10', '9', '9', '4'),
        ('Ocean Vans',    '98', '129','31', '9', '9', '6'),
        ('Nomade',        '80', '100','20', '2', '3', '10'),
    ], 1):
        for j, v in enumerate(row):
            tbl.rows[i].cells[j].text = v
    doc.add_paragraph()
    para(doc, 'Competitor profiles:', bold=True)
    bullet(doc, 'Further VAN — well-established, affordable. Less differentiated; likely operational-efficiency focus.')
    bullet(doc, 'People Camper — premium brand, high-end positioning. Likely higher costs or a significant brand premium.')
    bullet(doc, 'Ocean Vans — sophisticated premium offering. Significant price gap suggests strong premium differentiation.')
    bullet(doc, 'Nomade — NEW ENTRANT. Highly differentiated (sustainable design, customisation, 1-month delivery) but low brand recognition.')

    # ---------- 4. Tool features / simulator walkthrough ----------
    doc.add_page_break()
    h(doc, '3. Using the interactive simulator')
    para(doc, 'URL: ')
    p = doc.paragraphs[-1]; link = p.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE; link.underline = True

    para(doc,
         'The simulator is a CLOSED LOOP. The Nomade price lives in one place and '
         'every tab reads or writes the same number. You do not have to re-type the '
         'same price in multiple tabs — pick one, iterate, and the Excel download '
         'will reflect whatever you left on "Position & prices".',
         italic=True, color=GREY)

    for s in [
        'Pick period (Day / Weekend / Week) and version (Standard / Premium) at the top.',
        'Competition-based → Profit curve: DRAG the green "Our price" line. Coloured dashed lines are competitors; KPIs (profit, demand, revenue, lost-vs-peak) update live. Use Match-competitor / Snap-to-peak quick-jump buttons.',
        'Psychological factors: all inputs pre-loaded from your current position. Experiment with Anchoring, Charm pricing, Prospect Theory (before/after), Reference price. Use Apply-to-Standard / Apply-to-Premium to push tested numbers into the final position.',
        'Position & prices: same interactive chart + numeric inputs for every period. Add the analyst note.',
        'Case answers: the four formal questions (see section 4) — write your answer in each box.',
        'Download Excel: single .xlsx with 8 sheets including Case answers and all 9 charts as images — this file IS the deliverable.',
    ]:
        numbered(doc, s)

    # ---------- 5. Strategic framework ----------
    doc.add_page_break()
    h(doc, '4. Strategic framework — the 4 scenarios')
    para(doc,
         'The pricing decision depends on comparing the WTP optimal price (Popt) '
         'with competitor prices (C) and on understanding whether the gap comes '
         'from cost structure or from differentiation.')
    tbl = doc.add_table(rows=9, cols=4); tbl.style = 'Light Grid Accent 1'
    make_header_row(tbl, ['Scenario', 'Status', 'Reason', 'Strategy'])
    for i, row in enumerate([
        ('a.1','Popt > C','More differentiated',     'NO — illogical'),
        ('a.2','Popt > C','Worse costs',             'YES (−) Penetration'),
        ('a.3','Popt < C','Less differentiated',     'YES (−−) Aggressive penetration'),
        ('a.4','Popt < C','Better costs',            'NO — great position'),
        ('b.1','Popt > C','More differentiated',     'RISKY — only for Apple-like brands'),
        ('b.2','Popt > C','Worse costs',             'NO — escape / reposition'),
        ('b.3','Popt < C','Less differentiated',     'YES (+) Skimming'),
        ('b.4','Popt < C','Better costs',            'NO — already right'),
    ], 1):
        for j, v in enumerate(row):
            tbl.rows[i].cells[j].text = v
    doc.add_paragraph()
    para(doc, 'Nomade\'s strategic position (v3 baseline):', bold=True)
    bullet(doc, 'Current status: Popt < C (€80 / €100 vs €85–€105 Std and €95–€129 Prem).')
    bullet(doc, 'Cost structure: similar or slightly better (sustainable-design efficiency).')
    bullet(doc, 'Differentiation: OBJECTIVELY HIGHER (sustainability 10/10, customisation 10/10).')
    bullet(doc, 'BUT low brand recognition (NEW entrant — 3/10 vs 8–9/10).')
    bullet(doc, 'This is SCENARIO b.3: Popt < C because customers perceive Nomade as less differentiated (the brand is unknown, not the product).')
    bullet(doc, 'Recommendation: SKIMMING (+) — moderate price increase to build brand perception and leverage the anchor effect.')

    # ---------- 6. FORMAL QUESTIONS (4 v3) ----------
    doc.add_page_break()
    h(doc, '5. Case questions')
    para(doc,
         'Four formal questions from Nomade v3.0. Write your answer either (a) '
         'inside the simulator\'s Case answers tab — which is exported into the '
         'Excel — or (b) directly in the answer boxes below.',
         italic=True)

    # Q1
    p = doc.add_paragraph()
    r = p.add_run('P1 — Competitive analysis and positioning')
    r.bold = True; r.font.color.rgb = ORANGE
    para(doc, 'With the data provided and the tool:', bold=True)
    bullet(doc, 'Give assumptions about the competitors\' cost structures and their brand-value differentiation.')
    bullet(doc, 'Remember NOMADE is NEW while the others are well-established.')
    bullet(doc, 'For each competitor, evaluate whether higher prices are due to: (a) higher operating costs, or (b) higher differentiation (brand recognition, quality, features).')
    bullet(doc, 'Build a positioning map of Nomade vs. competitors on the key attributes.')
    para(doc, 'Simulator tasks:', bold=True, color=GREY)
    bullet(doc, 'Read profit at each competitor\'s dashed line on the Profit curve.')
    bullet(doc, 'Fill the table of brand / cost assumptions per competitor (see section 2.4).')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=12)

    # Q2
    p = doc.add_paragraph()
    r = p.add_run('P2 — Apply the 4-scenario strategic framework')
    r.bold = True; r.font.color.rgb = ORANGE
    para(doc, 'Using the 4-scenario framework (section 4):', bold=True)
    bullet(doc, 'Identify which scenario applies to Nomade for each version (Standard / Premium).')
    bullet(doc, 'Is Popt > or < competitors\' prices? Is the reason cost structure or differentiation?')
    bullet(doc, 'Given the scenario, should Nomade pursue: (a) penetration (−) / aggressive penetration (−−), (b) skimming (+) / fast-aggressive (++), or (c) nothing?')
    bullet(doc, 'Justify with competition theory and the WTP profit analysis.')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=14)

    # Q3
    p = doc.add_paragraph()
    r = p.add_run('P3 — Weekend / week pricing strategy')
    r.bold = True; r.font.color.rgb = ORANGE
    para(doc, 'Competitor data is DAILY only:', bold=True)
    bullet(doc, 'Start from your daily pricing decisions (WTP optimum + competitive positioning).')
    bullet(doc, 'Assume industry conventions: weekend ≈ 5–10 % per-day discount, week ≈ 20–30 % total discount.')
    bullet(doc, 'Propose Nomade\'s weekend / week prices.')
    bullet(doc, 'Use anchoring (daily × 7 vs weekly) and the "sandwich" between Standard and Premium.')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=14)

    # Q4
    p = doc.add_paragraph()
    r = p.add_run('P4 — Complete pricing proposal')
    r.bold = True; r.font.color.rgb = ORANGE
    para(doc, 'Deliver the full pricing strategy with justification:', bold=True)
    bullet(doc, 'Complete price table (Standard / Premium × Day / Weekend / Week).')
    tbl = doc.add_table(rows=3, cols=4); tbl.style = 'Light Grid Accent 1'
    make_header_row(tbl, ['', 'Day', 'Weekend (per day)', 'Week (per day)'])
    tbl.rows[1].cells[0].text = 'Standard'
    tbl.rows[2].cells[0].text = 'Premium'
    doc.add_paragraph()
    bullet(doc, 'Comparison vs competitors (with your weekend/week assumptions).')
    bullet(doc, 'Justify with: (a) WTP analysis (profit curves), (b) competitive positioning (differentiation vs costs), (c) psychological tactics (anchoring, sandwich, charm).')
    bullet(doc, 'Expected business outcomes (profit, market positioning).')
    bullet(doc, 'Consider alternative structures (e.g., 3-tier with "enhanced" version).')
    bullet(doc, 'Risk mitigation (competitor reactions).')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=16)

    # ---------- 7. Evaluation ----------
    doc.add_page_break()
    h(doc, '6. Evaluation criteria')
    bullet(doc, 'Rigour — evidence inside the simulator (drag results, screenshots, numbers).')
    bullet(doc, 'Integration of competition + WTP + psychological factors into a single recommendation.')
    bullet(doc, 'Correct use of the 4-scenario strategic framework.')
    bullet(doc, 'Clarity and brevity — numbers justified, trade-offs explicit.')

    h(doc, 'Deliverable', level=2)
    para(doc,
         'The Excel (.xlsx) downloaded from the simulator is enough. Submit that file alone — '
         'it already contains every piece of evidence required:')
    bullet(doc, 'Guide — case background, data note and active view.')
    bullet(doc, 'Questions — the 13 survey items (reference only).')
    bullet(doc, 'Answers — Day / Weekend / Week — aggregated WTP distributions.')
    bullet(doc, 'Position — your final prices, competitor prices, peak / gap / acceptance KPIs and the analyst note.')
    bullet(doc, 'Case answers — your written answers to the four formal questions.')
    bullet(doc, 'Charts — the nine simulator charts embedded as images.')
    para(doc,
         'Before downloading, make sure you have filled the four Case answers and set the '
         'final prices in Position & prices. No additional Word or PowerPoint file is required.',
         italic=True, color=GREY)

    # ---------- APPENDIX ----------
    doc.add_page_break()
    h(doc, 'Appendix — The 13 survey questions (reference only)')
    para(doc,
         'These are the 13 items asked to the 95 survey respondents. They appear here only as '
         'reference so that you can interpret the aggregated WTP distribution displayed by the '
         'simulator. You are NOT asked to answer them yourself. Original Spanish wording in italics.',
         italic=True, color=GREY)

    items = [
        ('Q1',  'Please indicate your gender.',
                'Por favor, indique su género.'),
        ('Q2',  'What age range are you in?  <20 / 20-30 / 30-40 / 40-50 / 50-60 / >60',
                '¿En qué rango de edad se encuentra?'),
        ('Q3',  'What is your civil status?  Single / Married / Divorced / Widowed',
                '¿Cuál es su estado civil?'),
        ('Q4',  'What is your level of education?  Primary / Secondary / University / Master or PhD',
                '¿Cuál es su nivel de educación?'),
        ('Q5',  'Select your employment status.  Employed / Self-employed / Unemployed / Retired / Student',
                'Seleccione su situación laboral.'),
        ('Q6',  'What do you like to do in your free time? (multiple choice)  Travel · Reading · Sports · Adventure · Family time · Other',
                '¿Qué le gusta hacer en su tiempo libre?'),
        ('Q7',  'How many days do you usually go on vacation?  <3 / 3-5 / 5-7 / >7',
                '¿Cuántos días sueles irte de vacaciones?'),
        ('Q8',  'Maximum price you are willing to pay — STANDARD camper, per day.',
                'Precio máximo por el alquiler de una camper ESTÁNDAR al día.'),
        ('Q9',  'Maximum price you are willing to pay — STANDARD camper, weekend (2 nights).',
                'Precio máximo — camper ESTÁNDAR fin de semana (2 noches).'),
        ('Q10', 'Maximum price you are willing to pay — STANDARD camper, week (6 nights).',
                'Precio máximo — camper ESTÁNDAR semana (6 noches).'),
        ('Q11', 'Maximum price you are willing to pay — PREMIUM camper, per day.',
                'Precio máximo — camper PREMIUM al día.'),
        ('Q12', 'Maximum price you are willing to pay — PREMIUM camper, weekend (2 nights).',
                'Precio máximo — camper PREMIUM fin de semana (2 noches).'),
        ('Q13', 'Maximum price you are willing to pay — PREMIUM camper, week (6 nights).',
                'Precio máximo — camper PREMIUM semana (6 noches).'),
    ]
    for qid, english, spanish in items:
        p = doc.add_paragraph()
        run = p.add_run(f'{qid}. '); run.bold = True
        p.add_run(english)
        p2 = doc.add_paragraph()
        r2 = p2.add_run(spanish); r2.italic = True; r2.font.color.rgb = GREY; r2.font.size = Pt(10)

    # ---------- Save ----------
    stamp = datetime.now().strftime('%Y-%m-%d')
    out = f'Nomade_Vans_Case_v3.0_Simulator_Edition_{stamp}.docx'
    doc.save(out)
    print('Saved:', out)


if __name__ == '__main__':
    build()
