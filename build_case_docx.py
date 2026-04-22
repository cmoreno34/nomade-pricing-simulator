"""
Build the Word workbook for the Nomade Vans case, adapted to the interactive
pricing simulator. Questions are the real Spanish survey items and the 3 formal
deliverable questions from Nomade v2.0 (expanded with simulator tasks).
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
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = True
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, 'F8FAFC')
    p = cell.paragraphs[0]
    p.add_run('\n' * (lines - 1))
    doc.add_paragraph()


def q_header(doc, qid, title_es, title_en):
    p = doc.add_paragraph()
    r = p.add_run(f'{qid}. ')
    r.bold = True
    r.font.color.rgb = ORANGE
    p.add_run(title_es).bold = True
    p2 = doc.add_paragraph()
    p2.add_run(title_en).italic = True
    p2.runs[0].font.color.rgb = GREY


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
    r = sub.add_run('Simulator-driven edition · v3.0 — student workbook')
    r.italic = True; r.font.size = Pt(12); r.font.color.rgb = GREY

    doc.add_paragraph()
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run('© César Moreno Pascual PhD — based on Nomade v2.0 by Ariane Atucha, Ángela Pesquera, Irache Gallego, Paula García\n').italic = True
    meta.add_run('Interactive simulator: ').italic = True
    link = meta.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE; link.underline = True

    doc.add_page_break()

    # ---------- 1. Case background ----------
    h(doc, '1. Caso · Case background')
    para(doc,
         "Nomade Vans es una empresa española que se centra en el diseño y la "
         "«camperización» de furgonetas sostenibles. Permite a los consumidores "
         "personalizar el producto a través de su web y se ha planteado una nueva "
         "línea de negocio: el alquiler de camper vans.")
    para(doc,
         "Two service tiers are proposed: Standard Camper Rental and Premium Camper "
         "Rental (hybrid model, extra bed, indoor shower, air conditioning, solar "
         "power, extra kitchenware). Rentals are offered by day, weekend (2 nights) "
         "and week (6 nights). The combination gives a 2×3 price structure.",
         italic=True, size=10)
    para(doc,
         "A WTP survey was carried out with 95 respondents. Three competitors have "
         "been identified in the relevant market: Further VAN experience, People "
         "Camper and Ocean Vans. The simulator reproduces the WTP distribution, "
         "profit curves, competitor positioning and the main psychological factors "
         "from the technical notes.")

    # ---------- 2. Price structure ----------
    h(doc, '2. Estructura de precios óptimos sin segmentar (Nomade v2.0 ·\u202fbaseline)')
    tbl = doc.add_table(rows=3, cols=5)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['', 'Day', 'Weekend', 'Week', 'Weekly (per day)']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    rows = [
        ('Standard', '80 €', '120 € / 60 €/day', '420 €', '60 €'),
        ('Premium',  '100 €', '160 € / 80 €/day', '540 €', '77 €'),
    ]
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()

    # ---------- 3. Cost structure ----------
    h(doc, '3. Estructura de costes · Cost structure')
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['Item', 'Standard', 'Premium']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    rows = [
        ('Variable cost per rental (VC)', '30 €',     '30 €'),
        ('Fixed cost per year (CF)',     '69.750 €', '77.750 €'),
        ('Demand potential (D)',         '5.000',    '5.000'),
    ]
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()
    para(doc,
         "Demand scaling: day × 1.0, weekend × 0.365 (≈ 182 rentals/year), "
         "week × 0.0743 (≈ 52 rentals/year).", italic=True, size=10)

    # ---------- 4. Competitors ----------
    h(doc, '4. Competidores (precios de referencia diario)')
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['Competitor', 'Standard (€)', 'Premium (€)']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    comp_rows = [
        ('Further VAN experience', '85',  '95'),
        ('People Camper',          '105', '115'),
        ('Ocean Vans',             '98',  '129'),
    ]
    for i, row in enumerate(comp_rows, 1):
        for j, val in enumerate(row):
            tbl.rows[i].cells[j].text = val
    doc.add_paragraph()
    para(doc,
         "All three competitors operate only by day. Weekend and weekly tariffs in "
         "the simulator are estimated (−5 % and −30 % respectively), as reference "
         "scenarios.", italic=True, size=10)

    # ---------- 5. Simulator walkthrough ----------
    h(doc, '5. Uso del simulador · How to use the simulator')
    para(doc, 'URL: ')
    p = doc.paragraphs[-1]; link = p.add_run(SIMULATOR_URL)
    link.font.color.rgb = BLUE; link.underline = True

    steps = [
        'Select period (Daily / Weekend / Week) and version (Standard / Premium) at the top. Everything updates.',
        'On Competition-based → Profit curve, DRAG the green "Our price" line. Coloured dashed lines are competitors. KPIs (profit, demand, revenue, lost-vs-peak) update live. Use Match-competitor / Snap-to-peak buttons for quick moves.',
        'On Psychological factors, all inputs are pre-loaded from your current position. Experiment with Anchoring (good-better-best), Charm pricing (€80 vs €79), Prospect Theory (before/after price, perceived pain/gain) and Reference price (cheapest competitor by default). Use Apply-to-position buttons to push a tested number back into Position & prices.',
        'On Position & prices you see the same interactive chart and full numeric inputs for every period. Add the analyst note.',
        'On Download Excel, one click produces a .xlsx with 7 sheets: Guide, Questions, Answers (Day / Weekend / Week), Position and Charts (all 9 charts embedded as images).',
    ]
    for s in steps:
        numbered(doc, s)

    doc.add_page_break()

    # ---------- 6. 13 real survey questions ----------
    h(doc, '6. La encuesta · The 13 survey questions')
    para(doc,
         'Before interpreting the aggregated data, answer the 13 questions as if '
         'you were one of the 95 respondents. Items 1–7 build the buyer profile; '
         'items 8–13 are the WTP measurements. Phrasing is the original Spanish '
         'from the questionnaire, with English translation below.', italic=True)

    items = [
        ('Q1', 'Por favor, indique su género.',
                '(Please indicate your gender.)'),
        ('Q2', '¿En qué rango de edad se encuentra?  <20 / 20-30 / 30-40 / 40-50 / 50-60 / >60',
                '(What age range are you in?)'),
        ('Q3', '¿Cuál es su estado civil?  Soltero / Casado / Divorciado / Viudo',
                '(What is your civil status?)'),
        ('Q4', '¿Cuál es su nivel de educación?  Primaria / Secundaria / Grado Universitario / Máster o Posgrado',
                '(What is your level of education?)'),
        ('Q5', 'Seleccione su situación laboral.  Empleado / Autónomo / Desempleado / Jubilado / Estudiante',
                '(Select your employment status.)'),
        ('Q6', '¿Qué le gusta hacer en su tiempo libre? (Puede marcar más de una respuesta)  '
               'Viajar · Leer · Practicar algún deporte · Realizar actividades de aventura · '
               'Pasar tiempo en familia · Otras',
                '(What do you like to do in your free time? Multiple answers allowed.)'),
        ('Q7', '¿Cuántos días sueles irte de vacaciones?  '
               'Menos de 3 días / Entre 3 y 5 días / Entre 5 y 7 días / Más de una semana',
                '(How many days do you usually go on vacation?)'),
        ('Q8', '¿Cuál es el precio máximo que estás dispuesto a pagar por el alquiler '
               'de una camper ESTÁNDAR al día?  (50 / 60 / 70 / 80 / 90 / 100 / 110 / 120 / 130 €)',
                '(Max price per day — Standard.)'),
        ('Q9', '¿Cuál es el precio máximo que estás dispuesto a pagar por el alquiler '
               'de una camper ESTÁNDAR durante un fin de semana (2 noches)?',
                '(Max price per weekend — 2 nights — Standard.)'),
        ('Q10', '¿Cuál es el precio máximo que estás dispuesto a pagar por el alquiler '
                'de una camper ESTÁNDAR durante una semana (6 noches)?',
                '(Max price per week — 6 nights — Standard.)'),
        ('Q11', '¿Cuál es el precio máximo que estás dispuesto a pagar por el alquiler '
                'de la camper PREMIUM al día?',
                '(Max price per day — Premium.)'),
        ('Q12', '¿Cuál es el precio máximo que estás dispuesto a pagar por el alquiler '
                'de una camper PREMIUM durante un fin de semana (2 noches)?',
                '(Max price per weekend — 2 nights — Premium.)'),
        ('Q13', '¿Cuál es el precio máximo que estás dispuesto a pagar por el alquiler '
                'de una camper PREMIUM durante una semana (6 noches)?',
                '(Max price per week — 6 nights — Premium.)'),
    ]
    for qid, es, en in items:
        p = doc.add_paragraph()
        run = p.add_run(f'{qid}. '); run.bold = True
        p.add_run(es)
        p2 = doc.add_paragraph(); r2 = p2.add_run(en); r2.italic = True; r2.font.color.rgb = GREY
        answer_box(doc, lines=2)

    doc.add_page_break()

    # ---------- 7. Formal deliverable questions (v2.0 expanded) ----------
    h(doc, '7. Preguntas del caso · Formal deliverable questions')
    para(doc,
         'The three questions below reproduce the original Nomade v2.0 deliverables. '
         'Each one is expanded with concrete simulator tasks so you can ground your '
         'answer in evidence from the interactive tool.', italic=True)

    # Q1
    q_header(doc, 'P1',
             'Indicar las suposiciones sobre las estructuras de costes de los '
             'competidores y la diferenciación del valor de marca en comparación '
             'con NOMADE. Tenga en cuenta que NOMADE es nuevo y que los demás '
             'están bien establecidos.',
             '(State your assumptions about the competitors\' cost structures and '
             'the differentiation of Nomade\'s brand value versus the established '
             'rivals.)')
    para(doc, 'Tareas con el simulador · Tasks with the simulator:', bold=True)
    bullet(doc,
           'In Competition-based → Profit curve, read the profit we would earn '
           'at each competitor\'s price (dashed lines). Fill the table below.')
    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['Competitor', 'Their Std (€)', 'Our profit at their price (€)', 'Likely their FC (assumption)']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    for i, name in enumerate(['Further VAN experience', 'People Camper', 'Ocean Vans'], 1):
        tbl.rows[i].cells[0].text = name
    doc.add_paragraph()
    bullet(doc,
           'Justify why NOMADE\'s FC could be higher (new, smaller scale, fewer '
           'vehicles to amortise fixed costs) and what brand-value elements (design, '
           'sustainability, customisation) could compensate.')
    bullet(doc,
           'Conclude with one-line assumptions per competitor (cost structure + '
           'brand-value gap).')
    answer_box(doc, lines=10)

    # Q2
    q_header(doc, 'P2',
             'Sugerir una estrategia competitiva para Nomade, indicando si es '
             'posible mejorar su posición en el mercado con una visión estratégica, '
             'y considerar también algunos posibles factores psicológicos.',
             '(Recommend a competitive strategy — strategic positioning + '
             'psychological factors.)')
    para(doc, 'Tareas con el simulador · Tasks with the simulator:', bold=True)
    bullet(doc,
           'Use the Profit curve drag-handle to propose a Standard price that beats '
           'our current peak while remaining below People Camper. Screenshot it.')
    bullet(doc,
           'Open Psychological factors. Reset all controls from position. '
           'Experiment in this order:')
    numbered(doc,
             'Anchoring: design a Good-Better-Best sandwich. What middle price do '
             'you recommend? Apply it to Our Standard.')
    numbered(doc,
             'Charm pricing: test round vs charm. Does the +7 % perception uplift '
             'beat the margin cost of −€1? Apply if it does.')
    numbered(doc,
             'Prospect Theory: if you raise price by €10 next year, quantify the '
             'perceived pain. Draft a one-sentence framing to reduce it.')
    numbered(doc,
             'Reference price: set the cheapest competitor as reference. What share '
             'of respondents sit above? Does this support or weaken your Standard price?')
    bullet(doc,
           'Finish with a <100-word positioning statement: where does Nomade sit, '
           'why, and what is the psychological hook.')
    answer_box(doc, lines=14)

    # Q3
    q_header(doc, 'P3',
             'Sugerir una estrategia de implementación de precios con los precios '
             'de lista finales y elaborar algunas posibles acciones promocionales.',
             '(Final list prices and promotional actions.)')
    para(doc, 'Tareas con el simulador · Tasks with the simulator:', bold=True)
    bullet(doc,
           'For each combination (period × version) set the final price in '
           'Position & prices. Fill the table:')
    tbl = doc.add_table(rows=3, cols=4)
    tbl.style = 'Light Grid Accent 1'
    for i, txt in enumerate(['', 'Day', 'Weekend (per day)', 'Week (per day)']):
        cell = tbl.rows[0].cells[i]; cell.text = txt
        for r in cell.paragraphs[0].runs: r.bold = True
    tbl.rows[1].cells[0].text = 'Standard'
    tbl.rows[2].cells[0].text = 'Premium'
    doc.add_paragraph()
    bullet(doc,
           'Justify each number in one bullet, referring explicitly to: the peak, '
           'the WTP acceptance %, the nearest competitor and a psychological lever.')
    bullet(doc,
           'Draft 2–3 promotional actions compatible with your strategy (e.g., '
           'early-bird weekend, loyalty week-at-day-price, charm endings, '
           'good-better-best bundle). For each action, specify: target segment, '
           'price lever, expected pain / gain according to Prospect Theory.')
    bullet(doc,
           'Add the analyst note inside the simulator, click Download Excel, and '
           'attach the file to your submission.')
    answer_box(doc, lines=14)

    # ---------- 8. Evaluation ----------
    doc.add_page_break()
    h(doc, '8. Criterios de evaluación · Evaluation criteria')
    bullet(doc, 'Rigor (simulator evidence: screenshots, Excel, numbers).')
    bullet(doc, 'Integration of competition + WTP + psychological factors into a coherent story.')
    bullet(doc, 'Clarity & brevity — numbers justified, trade-offs explicit.')
    bullet(doc, 'Attached Excel export with final prices, position snapshot and chart images.')

    h(doc, 'Entregables · Deliverables', level=2)
    numbered(doc, 'This workbook (Word), every answer box filled in.')
    numbered(doc, 'The Excel file (.xlsx) downloaded from the simulator with your final position.')
    numbered(doc, 'A one-page executive summary of your pricing recommendation.')

    out = 'Nomade_Vans_Case_Study_Simulator_Edition.docx'
    doc.save(out)
    print('Saved:', out)


if __name__ == '__main__':
    build()
