# Nomade Vans — Pricing Simulator

An interactive, self-contained pricing simulator for the Nomade Vans case.
It combines **competition-based pricing** with **psychological factors**, lets you pick the graphs you want dynamically, and exports the questionnaire + the positioning you chose.

> Single-file app. No build step. Drop `index.html` anywhere and it runs.

---

## Features

### 1. Competition-based module
- WTP distribution with cumulative acceptance curve
- Profit curve with Std / Prem optimal markers and your own price marker
- Demand + Revenue combo chart
- Competitive positioning bar chart
- Gap vs each competitor
- Editable competitor prices (Further VAN, People Camper, Ocean Vans) per period

### 2. Psychological factors module
- **Anchoring — Good-Better-Best sandwich** (editable basic / middle / premium tiers)
- **Charm pricing — digit-9 effect** (round vs charm price, perception uplift)
- **Prospect Theory** value function (losses weigh more than gains)
- **Reference price** — how many respondents fall below / at / above

### 3. Dynamic graph picker
Every module has a chip-style picker — tick the graphs you want on the board. Recharts renders them side-by-side, responsive to period + version.

### 4. "How to use" walk-through
Built-in tab with a 5-step guide, per-tab explanations, per-chart "what this shows" callouts, and a click-to-expand glossary (WTP, peak price, anchoring, charm pricing, Prospect Theory, reference price, …).

### 5. Interactive Profit curve
- **Drag the green "Our price" line** left/right — KPIs for profit, demand, revenue and "lost vs peak" update live.
- Competitors (Further VAN, People Camper, Ocean Vans) appear as **dashed coloured lines on the Profit curve**, so you can read at each competitor's x what our profit would be at their price — and immediately see if moving up/down makes sense.
- Quick-jump buttons: *Snap to peak*, *Match Further VAN*, *Match People Camper*, *Match Ocean Vans*.
- The same chart lives on the Position & prices tab, so dragging it syncs the numeric inputs.

### 6. Student workbook (Word)
`Nomade_Vans_Case_Study_Simulator_Edition.docx` — a companion Word document built from `build_case_docx.py`. It contains the case background, cost structure, competitor prices, the 13 survey questions for the student to answer, a 5-step simulator walkthrough and 8 formal deliverable questions with answer boxes.

### 7. One-click Excel export
A single `.xlsx` download with **7 sheets**:
- **Guide** — short explanation of the file
- **Questions** — the 13 survey questions (text, type, options)
- **Answers — Day / Weekend / Week** — aggregated WTP tables (price, responses, %, cumulative acceptance)
- **Position** — your Nomade prices, competitor prices, KPIs (peak profit, gap vs average, acceptance %), psychological parameters and free-text analyst note
- **Charts** — all 9 simulator charts, embedded as PNG images

Opens in Excel, Numbers, Google Sheets or LibreOffice without any plug-in.

---

## Run locally

```bash
# any static server works
python -m http.server 8080
# then open http://localhost:8080
```

Or double-click `index.html` — it loads React, Recharts and Babel from a CDN.

---

## Publish to GitHub Pages

```bash
# inside this folder
git init
git add index.html README.md
git commit -m "Nomade Vans pricing simulator"
git branch -M main
git remote add origin https://github.com/<your-user>/nomade-pricing-simulator.git
git push -u origin main
```

Then on GitHub:

1. Repo → **Settings** → **Pages**
2. Source: **Deploy from a branch** → **main** → **/ (root)**
3. Save. The simulator will be live at `https://<your-user>.github.io/nomade-pricing-simulator/`.

No build, no Node, no workflow needed.

---

## Data sources

- Questionnaire of 90 respondents (see `Copy of Respuestas Cuestionario 90 respuestas.xlsx`)
- Competition-based pricing — `TECHNICAL NOTE_ competition based pricing v3.0.pdf`
- Psychological factors — `TECHNICAL NOTE_ psicologocal factors_ v 1.0.pdf`

## Credits

© César Moreno Pascual PhD.
