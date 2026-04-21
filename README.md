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

### 4. Download Q&A + Position
- 13 survey questions (demographics + 6 WTP questions)
- Aggregated WTP distribution for day / weekend / week, standard + premium
- **Your position in the simulator**: Nomade prices, competitor prices, peak-profit prices, gap vs competitor average, WTP acceptance rate, psychological parameters used and free-text analyst note
- Two formats: `JSON` (full structure) and `CSV` (flat, opens in Excel)

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
