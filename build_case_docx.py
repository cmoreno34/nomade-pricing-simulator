"""
Build the Word workbook for the Nomade Vans case, adapted to the interactive
pricing simulator. Deliverables = the three formal questions from Nomade v2.0.
The 13 survey items stay as a reference appendix (what respondents were asked),
NOT as student tasks.
"""
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


def answer_box(doc, lines=8):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, 'F8FAFC')
    p = cell.paragraphs[0]
    p.add_run('\n' * (lines - 1))
    doc.add_paragraph()


def build():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ---------- Cover ----------
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('CASO PRÁCTICO · CASE STUDY')
    r.bold = True; r.font.size = Pt(13); r.font.color.rgb = GREY

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('Nomade Vans — Pricing under Competition and Psychological Factors')
    r.bold = True; r.font.size = Pt(22); r.font.color.rgb = BLUE

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run('Simulator-driven edition · v3.1 — student workbook')
    r.italic = True; r.font.size = Pt(12); r.font.color.rgb = GREY

    doc.add_paragraph()
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run('© César Moreno Pascual PhD — based on Nomade v2.0 by Ariane Atucha, Ángela Pesquera, Irache Gallego, Paula García\n').italic = True
    meta.add_run('Interactive simulator: ').italic = True
    link = meta.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE; link.underline = True

    doc.add_page_break()

    # ---------- 1. The story ----------
    h(doc, '1. The story · El caso')
    para(doc,
         "Nomade Vans es una empresa española que se centra en el diseño y la «camperización» "
         "de furgonetas sostenibles. Permite a los consumidores personalizar el producto a través "
         "de su web y promete cumplir con un plazo de entrega de un mes. A diferencia de otras "
         "empresas del sector, emplea un diseño único y coherente con los valores de "
         "sostenibilidad y el entorno natural.")
    para(doc,
         "Tras diseñar y camperizar furgonetas para la venta directa, se plantea la posibilidad "
         "de iniciar una nueva línea de negocio basada en el alquiler de camper vans. Para ello "
         "se establecen dos niveles de servicio — Standard y Premium — y tres métricas temporales "
         "— día, fin de semana (2 noches) y semana (6 noches). La combinación de tiers y periodos "
         "define una estructura de precios 2 × 3.")
    para(doc,
         "Nomade Vans is a Spanish start-up designing sustainable camperised vans. Two service "
         "tiers (Standard and Premium) and three rental periods (Day / Weekend / Week) define "
         "the price structure. To calibrate willingness-to-pay the team surveyed 95 people, and "
         "marketing identified three competitors — Further VAN experience, People Camper and "
         "Ocean Vans — currently operating only with daily tariffs.",
         italic=True, size=10, color=GREY)

    # ---------- 2. Data note ----------
    note_box(doc,
             "The three competitors publish DAILY tariffs only. Weekend and weekly prices "
             "shown throughout the simulator are ESTIMATES aggregated from the daily tariff "
             "(weekend ≈ −5 % per day, week ≈ −30 % per day). Treat them as reference "
             "scenarios, not observed data. Any number can be overridden on the "
             "Position & prices tab. The Nomade WTP data itself is real (95 respondents).",
             title='⚠ Data note — dataset is incomplete')

    # ---------- 3. Price structure (v2.0 optima) ----------
    h(doc, '2. Estructura de precios óptimos sin segmentar (Nomade v2.0 baseline)')
    tbl = doc.add_table(rows=3, cols=5)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['', 'Day', 'Weekend (total)', 'Week (total)', 'Per-day price']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    rows = [
        ('Standard', '80 €',  '120 €', '420 €', '60 € weekend / 60 € week'),
        ('Premium',  '100 €', '160 €', '540 €', '80 € weekend / 77 € week'),
    ]
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()

    # ---------- 4. Cost structure ----------
    h(doc, '3. Estructura de costes · Cost structure')
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['Item', 'Standard', 'Premium']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    for i, row in enumerate([
        ('Variable cost per rental (VC)', '30 €',     '30 €'),
        ('Fixed cost per year (CF)',     '69.750 €', '77.750 €'),
        ('Demand potential (D)',         '5.000',    '5.000'),
    ], 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()
    para(doc,
         "Demand scaling: day × 1.0, weekend × 0.365 (≈ 182 rentals/year), "
         "week × 0.0743 (≈ 52 rentals/year).", italic=True, size=10, color=GREY)

    # ---------- 5. Competitors ----------
    h(doc, '4. Competidores · Competitors (daily tariff — the only one published)')
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['Competitor', 'Standard (€)', 'Premium (€)']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    for i, row in enumerate([
        ('Further VAN experience', '85',  '95'),
        ('People Camper',          '105', '115'),
        ('Ocean Vans',             '98',  '129'),
    ], 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()

    # ---------- 6. Simulator walkthrough ----------
    h(doc, '5. Uso del simulador · How to use the simulator')
    para(doc, 'URL: ')
    p = doc.paragraphs[-1]; link = p.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE; link.underline = True

    for s in [
        'Pick a rental period (Day / Weekend / Week) and a version (Standard / Premium) at the top.',
        'Go to Competition-based → Profit curve. DRAG the green "Our price" line. Coloured dashed lines are the three competitors; KPIs (profit, demand, revenue, lost-vs-peak) update live. Use Match-competitor / Snap-to-peak.',
        'Go to Psychological factors — all inputs are pre-loaded from your current position. Experiment (Anchoring, Charm, Prospect Theory before/after price, Reference). Use Apply-to-Standard / Apply-to-Premium buttons to push a tested number back into Position & prices.',
        'Open Position & prices to type numbers directly and add the analyst note.',
        'Open Case answers — the three formal questions are there; write your answer in each box.',
        'Open Download Excel. One click produces a .xlsx with 8 sheets including Case answers and all 9 charts embedded as images.',
    ]:
        numbered(doc, s)

    doc.add_page_break()

    # ---------- 7. FORMAL QUESTIONS ----------
    h(doc, '6. Preguntas del caso · Formal deliverable questions')
    para(doc,
         'These three questions reproduce the original Nomade v2.0 deliverables. '
         'Write your answer either (a) inside the simulator\'s Case answers tab — '
         'everything will be exported to the Excel download — or (b) directly in '
         'the boxes below.', italic=True)

    # Q1
    p = doc.add_paragraph()
    r = p.add_run('P1 — '); r.bold = True; r.font.color.rgb = ORANGE
    r2 = p.add_run('With the indicated data, give assumptions of competitors\' cost '
                    'structures and brand-value differentiation compared to NOMADE. '
                    'Consider that NOMADE is new, and the others are well-established.')
    r2.bold = True
    para(doc, 'Simulator tasks:', bold=True)
    bullet(doc, 'Read profit at each competitor\'s price on the Profit curve (dashed lines). Fill the table.')
    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['Competitor', 'Their Std (€)', 'Our profit at their price (€)', 'Likely FC assumption']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    for i, name in enumerate(['Further VAN experience', 'People Camper', 'Ocean Vans'], 1):
        tbl.rows[i].cells[0].text = name
    doc.add_paragraph()
    bullet(doc, 'Justify NOMADE\'s cost disadvantage (new, smaller scale) and brand-value upside (design, sustainability, customisation).')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=10)

    # Q2
    p = doc.add_paragraph()
    r = p.add_run('P2 — '); r.bold = True; r.font.color.rgb = ORANGE
    r2 = p.add_run('Suggest a competitive strategy for Nomade, indicating whether '
                    'Nomade can improve its market position with a strategic view. '
                    'Also consider possible psychological factors.')
    r2.bold = True
    para(doc, 'Simulator tasks:', bold=True)
    bullet(doc, 'Drag the green line and use Match-competitor / Snap-to-peak. Take a screenshot.')
    bullet(doc, 'On Psychological factors, experiment in this order:')
    numbered(doc, 'Anchoring: design a Good-Better-Best sandwich. Apply the middle to Our Standard.')
    numbered(doc, 'Charm pricing: test round vs charm. Does the +7 % perception uplift beat the margin cost of −€1? Apply if it does.')
    numbered(doc, 'Prospect Theory: if you raise price by €10 next year, quantify the pain and draft one sentence to reframe it.')
    numbered(doc, 'Reference price: cheapest competitor as reference. What share of respondents sit above?')
    bullet(doc, 'Close with a <100-word positioning statement.')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=14)

    # Q3
    p = doc.add_paragraph()
    r = p.add_run('P3 — '); r.bold = True; r.font.color.rgb = ORANGE
    r2 = p.add_run('Suggest a pricing implementation strategy with the final list '
                    'prices and elaborate on possible promotional actions.')
    r2.bold = True
    para(doc, 'Simulator tasks:', bold=True)
    bullet(doc, 'Set final prices in Position & prices for each combination:')
    tbl = doc.add_table(rows=3, cols=4)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['', 'Day', 'Weekend (per day)', 'Week (per day)']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    tbl.rows[1].cells[0].text = 'Standard'
    tbl.rows[2].cells[0].text = 'Premium'
    doc.add_paragraph()
    bullet(doc, 'Justify each number (peak, WTP %, nearest competitor, psychological lever).')
    bullet(doc, 'Draft 2–3 promotions — target segment, price lever, expected pain/gain (Prospect Theory).')
    bullet(doc, 'Add the analyst note in the simulator, click Download Excel, attach the file.')
    para(doc, 'Your answer:', bold=True)
    answer_box(doc, lines=14)

    # ---------- 8. Evaluation ----------
    doc.add_page_break()
    h(doc, '7. Criterios de evaluación · Evaluation criteria')
    bullet(doc, 'Rigour (simulator evidence: screenshots, Excel, numbers).')
    bullet(doc, 'Integration of competition + WTP + psychological factors.')
    bullet(doc, 'Clarity & brevity — numbers justified, trade-offs explicit.')
    bullet(doc, 'Attached Excel export with final prices, position, Case answers and chart images.')

    h(doc, 'Entregables · Deliverables', level=2)
    numbered(doc, 'This workbook (Word), every answer box filled in — OR case answers written inside the simulator.')
    numbered(doc, 'The Excel (.xlsx) downloaded from the simulator — Guide, Questions, Answers (Day/Weekend/Week), Position, Case answers, Charts.')
    numbered(doc, 'A one-page executive summary.')

    # ---------- APPENDIX ----------
    doc.add_page_break()
    h(doc, 'Anexo · Appendix — The 13 survey questions (reference)')
    para(doc,
         'These are the 13 items asked to the 95 survey respondents. They appear here for '
         'reference so that you can interpret the aggregated WTP distribution displayed by '
         'the simulator. You are NOT asked to answer them yourself.', italic=True, color=GREY)

    items = [
        ('Q1',  'Por favor, indique su género.'),
        ('Q2',  '¿En qué rango de edad se encuentra?  <20 / 20-30 / 30-40 / 40-50 / 50-60 / >60'),
        ('Q3',  '¿Cuál es su estado civil?  Soltero / Casado / Divorciado / Viudo'),
        ('Q4',  '¿Cuál es su nivel de educación?  Primaria / Secundaria / Grado / Máster o Posgrado'),
        ('Q5',  'Seleccione su situación laboral.  Empleado / Autónomo / Desempleado / Jubilado / Estudiante'),
        ('Q6',  '¿Qué le gusta hacer en su tiempo libre? (Multi-select)  Viajar · Leer · Deporte · Aventura · Familia · Otras'),
        ('Q7',  '¿Cuántos días sueles irte de vacaciones?  <3 / 3-5 / 5-7 / >7'),
        ('Q8',  'Precio máximo dispuesto a pagar — camper ESTÁNDAR al día.'),
        ('Q9',  'Precio máximo dispuesto a pagar — camper ESTÁNDAR fin de semana (2 noches).'),
        ('Q10', 'Precio máximo dispuesto a pagar — camper ESTÁNDAR semana (6 noches).'),
        ('Q11', 'Precio máximo dispuesto a pagar — camper PREMIUM al día.'),
        ('Q12', 'Precio máximo dispuesto a pagar — camper PREMIUM fin de semana (2 noches).'),
        ('Q13', 'Precio máximo dispuesto a pagar — camper PREMIUM semana (6 noches).'),
    ]
    for qid, text in items:
        p = doc.add_paragraph()
        run = p.add_run(f'{qid}. '); run.bold = True
        p.add_run(text)

    from datetime import datetime
    stamp = datetime.now().strftime('%Y-%m-%d')
    out = f'Nomade_Vans_Case_v3.1_Simulator_Edition_{stamp}.docx'
    doc.save(out)
    print('Saved:', out)


if __name__ == '__main__':
    build()
