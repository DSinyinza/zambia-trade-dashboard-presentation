# Zambia Agriculture Trade Dashboard Presentation
### *Storytelling with Data — 8-Slide, 5-Minute Deck*

> **"Zambia's widening K33bn trade deficit is not just a trade problem — it is a currency crisis amplifier: every Kwacha depreciation makes imports more expensive while masking the true stagnation of export volumes, and without structural diversification, the deficit will compound."**

---

## Overview

This repository contains a Python script that automatically generates a professional PowerPoint presentation (.pptx) analysing Zambia's agriculture import-export trade data from 2013 to 2022.

The presentation applies **Cole Nussbaumer Knaflic's "Storytelling with Data"** principles:
- One big idea per slide (action headlines)
- Sparse design — no bullet soup
- Big numbers front and centre
- Strategic, purposeful colour use
- Visual hierarchy preserved
- Annotations guide the audience's attention

---

## Quick Start

```bash
# 1. Install dependency
pip install python-pptx

# 2. Generate the presentation
python build_presentation.py

# 3. Open the output file
#    zambia_trade_presentation.pptx
```

---

## Files

| File | Description |
|------|-------------|
| `build_presentation.py` | Main Python script — generates the .pptx |
| `zambia_trade_presentation.pptx` | Generated presentation (run script first) |
| `data/trade_data.csv` | Year-by-year import/export data (2013–2022) |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |

---

## Presentation Structure

| Slide | Title | Purpose |
|-------|-------|---------|
| 1 | **Zambia's Trade Story: K33 Billion Warning** | Title & key numbers |
| 2 | **Zambia Flipped from Surplus to Structural Deficit in Under a Decade** | The situation — plot the conflict |
| 3 | **Imports Grew 428%. Exports Grew 138%. The Gap Tells the Story.** | Data — diverging bar chart |
| 4 | **The Numbers Lie: Currency Depreciation is Inflating Both Sides** | Twist — ZMW/USD context |
| 5 | **Urea Dependency: Zambia Pays More Each Year to Grow the Same Crops** | Import drivers — the fertiliser trap |
| 6 | **Raw Commodity Exports: High Volume, Low Value-Add, Capped Upside** | Export drivers — structural ceiling |
| 7 | **The Deficit Loop: How Trade Imbalance Deepens Currency Weakness** | Compounding risk cycle |
| 8 | **Three Priorities to Break the Deficit Loop** | Call to action |

---

## Key Insights

### 1. From Surplus to Structural Deficit
Zambia ran a **+K1.23bn trade surplus in 2013**. By 2022 the annual deficit was **-K9.27bn**, contributing to a cumulative decade-long shortfall of **K33.4bn**.

### 2. Import Explosion Driven by Currency Depreciation
Import amounts rose from **K4.2bn (2013) to K22.2bn (2022) — a 428% increase**. The Zambian Kwacha (ZMW) depreciated from approximately **5 ZMW/USD (2013) to 17–21 ZMW/USD (2020–22)**, meaning largely the same USD volume of goods costs 4× more in local currency terms.

### 3. Export Growth is Partially Illusory
Export values rose from K5.4bn to K12.9bn (+138%), but Zambia's main exports — tobacco, maize, cotton, oil cake — are raw commodities whose USD prices are globally determined. The ZMW rise in export revenues masks flat-to-modest volume growth.

### 4. The Fertiliser Trap
**Urea is Zambia's #1 import**. As ZMW weakens, fertiliser costs rise for farmers, threatening yields of the very export crops (tobacco, maize) that earn the forex needed to stabilise the Kwacha — a vicious cycle.

### 5. 2022 Warning Sign
Despite continuing ZMW weakness (which should inflate ZMW export figures), exports **fell** from K14.17bn (2021) to K12.95bn (2022) — the first export contraction, suggesting real volume or commodity price deterioration that currency effects could no longer mask.

### 6. Structural Vulnerability
Zambia is a **price-taker** in global commodity markets: it exports raw agricultural goods (low value-add) and imports processed inputs (high USD cost). Currency weakness disproportionately hurts import costs while export revenue gains are capped.

---

## Three Priorities to Break the Deficit Loop

1. **Reduce Urea Dependency** — invest in domestic fertiliser production or regional procurement in ZMW/barter arrangements to reduce USD outflow on the single largest import.

2. **Value-Add Exports** — process tobacco, cotton, and maize locally before export to capture more USD per unit and reduce dependence on global commodity price cycles.

3. **Import Substitution** — develop local capacity in top import categories (petroleum, mineral products) to reduce structural USD outflow and ease pressure on the Kwacha.

---

## Data Source

Zambia Agriculture Import-Export Dashboard 2013–2022  
FSIO Cohort Analysis — May 2026

| Year | Exports (ZMW bn) | Imports (ZMW bn) | Balance (ZMW bn) |
|------|-----------------|-----------------|-----------------|
| 2013 | 5.44 | 4.21 | +1.23 |
| 2014 | 4.73 | 4.62 | +0.11 |
| 2015 | 5.91 | 7.03 | -1.11 |
| 2016 | 6.23 | 7.29 | -1.06 |
| 2017 | 5.33 | 9.43 | -4.10 |
| 2018 | 5.59 | 9.24 | -3.65 |
| 2019 | 6.59 | 10.40 | -3.82 |
| 2020 | 9.70 | 15.46 | -5.76 |
| 2021 | 14.17 | 20.13 | -5.97 |
| 2022 | 12.95 | 22.22 | -9.27 |

---

## Design Principles

- **Dark navy theme** (`#0d1b2a`) for professional, high-contrast slides
- **Orange** (`#e85d04`) for imports and warnings
- **Teal-blue** (`#1d6fa4`) for exports and positive indicators
- **Red** (`#c1121f`) for deficits and danger signals
- **Gold** (`#f4a261`) for callouts and key quotes
- Visual "chart simulations" built with coloured rectangles (no external chart libraries required)
- All charts and graphics generated programmatically using `python-pptx`

---

*Generated by FSIO Cohort | Storytelling with Data Principles Applied*
