---
name: appinio-decks
description: Build on-brand Appinio presentations (.pptx) from the official Appinio Slide Deck Template 2026 and the Appinio Design System. Use whenever someone asks for a deck, slides, a presentation, a research results report, a sales or pitch deck, a company intro, a proposal or case study, or wants content turned into Appinio-branded slides. Produces a PowerPoint file that also opens in Google Slides and Keynote.
compatibility: Needs Python 3 with python-pptx and Pillow (code execution or a terminal). Inter font installed on the viewing machine.
---

# Appinio decks

You write the storyline and the words as a JSON outline; `scripts/build_deck.py` builds every slide on the **real layouts of the Appinio Slide Deck Template 2026** (`assets/template.pptx`) and draws charts, cards, tables, timelines and other components exactly as the Appinio Design System specifies. **Never build slides any other way** (no hand-written python-pptx, no other library).

## Step 1 — decide the deck style (always first)

| | `"style": "results"` | `"style": "master"` |
|---|---|---|
| For | Research results: study findings, question charts, methods (KDA, TURF, pricing, conjoint), brand health, tracking | Everything else: sales and pitch decks, company intros, proposals, capabilities, case studies, events, team, internal and thought-leadership decks |
| Heading | Dark grey uppercase **`eyebrow`** above the title | Grey **`subtitle`** under the title |
| Chapter | Large number, title and `description` at mid-height | "01" and title on one line at the top |
| Footer | Hairline divider, `note` under it, logo below | No divider; `note` beside the logo |
| Boxes | Grey 10 cards; takeaway panel beside charts | Blue 10 panels with white cards |
| Charts | Full data-viz rules (palettes, n-step scales, profiles) | Same palettes, keep charts simple: one message |

A deck never mixes the two styles. If unsure, ask: "Is this a results deck or a sales/company deck?"

## Step 2 — workflow

1. Plan the story (Content rules below). One message per slide; titles state the finding.
2. Write `outline.json`: `{"title": "...", "style": "results" | "master", "slides": [ ... ]}`. Check asset names with `python scripts/build_deck.py --list-assets`.
3. Build: `pip install python-pptx Pillow` if needed, then `python <skill-dir>/scripts/build_deck.py outline.json -o <name>.pptx` (`<skill-dir>` = the folder holding this file; if unsure: `find / -name build_deck.py -path "*appinio-decks*" 2>/dev/null`).
4. Fix every WARNING (usually text too long: shorten words, never shrink fonts) and rebuild until clean.
5. Deliver the .pptx; mention that recipients need the Inter font (`assets/fonts/`, free on Google Fonts).

**Structure:** `cover` → `agenda` (3+ chapters) → per chapter `chapter` (+ optional `section_intro`) → content → optional `contact` / `team` / `thank_you` → **always** `back_cover`.

## Layout catalog

Every slide may add `notes` (speaker notes). Highlight one phrase per slide with `[[double brackets]]` (Appinio Blue). In results decks use `eyebrow` (1–4 words); in master decks use `subtitle`. Data slides take `note` (base and source: "Base: n = 1,204").

**Shared (both styles; the style changes the look automatically)**

| type | fields | limits |
|---|---|---|
| `cover` | `title`, `subtitle?`, `photo?` (photo cover, master), `partner_logo?` (white logo file → co-brand) | title ≤ 2 short lines (`\n`) |
| `agenda` | `chapters`: [{`title`, `page?`}] | results: ≤ 6 with page numbers; master: ≤ 6 big rows, 7–20 two columns |
| `chapter` | `title`, `description?` (results only), `number?` | ≤ 24 chars |
| `section_intro` | `title`, `eyebrow`/`subtitle`, `intro` | 1–3 sentences |
| `key_findings` | `findings`: [str] | 3 (max 4) |
| `stat` | `title`, `stats`: [{`label`, `value`, `meaning`}], `highlight?` (index), `note?` | 1–3; value ≤ 5 chars |
| `chart` | `title`, `kind` (`bar` · `column` · `line` · `stacked_bar` · `stacked_column`), `categories`, `series`: [{`name`, `values`, `sig?`}], `number_format?` (`0%` for 0–1, `0"%"` for whole numbers), `takeaways?` (≤ 3), `colors?`, `note` | ≤ 12 categories. Stacked = 100 % bars for scales: one series per scale point, `colors` = the n-step row |
| `table` | `title`, `columns`, `rows`, `highlight_column?` (index, e.g. the recommended package), `first_column_width?` (px), `note?` | ≤ 8 rows, ≤ 6 cols |
| `funnel` | `title`, `stages`: [{`value` (0.61 or "61%"), `label`, `conversion?`}], `highlight?` (index), `note?` | 3–6 stages, brand or sales funnels |
| `fact_cards` | `title`, `cards`: [{`icon`, `label`, `value?`, `detail?`}] | 2–6 |
| `icon_points` | `title`, `points`: [{`icon`, `headline`, `text`}] | 2–4; icons for all or none |
| `steps` | `title`, `steps`: [{`title`, `text`}], `growing?` (master: dots grow, for growth stories) | 3–5 |
| `content` | `title`, `body`, `bullets` (≤ 5) or `body_right` | ~150 words |
| `image` | `title`, `body`, `photo` | ~60 words |
| `quote` | `quote`, `name`, `role`, `photo?` (portrait half) or `avatar?` + `client_logo?` (testimonial) | ≤ 25 words; real quotes only |
| `statement` | `lines`: [2 strings, one `[[phrase]]`], `gradient?` (true or a gradient name), `photo?` (culture/mission over a photo, once per deck); omit `lines` for the tagline | ≤ 40 chars/line |
| `contact` | `people`: [{`name`, `role`, `email`, `phone?`, `photo`}] | 1–2, real photos |
| `back_cover` | `variant?` ("lilac"), `people?` | always last |

**Results decks only**

| type | fields | notes |
|---|---|---|
| `profile` | `title`, `pairs`: [[negative pole, positive pole]], `series`: [{`name`, `values` (means 1–5)}], `takeaways?`, `note` | semantic differential / brand profile; lead brand first |

**Results extras on `chart` and `table`**
- Significance letters: on a chart series `"sig": ["", "B C", "A"]` (one entry per value); on a table `"sig"` is a matrix shaped like `rows`. Cuts are lettered A, B, C in series / column order (added to the legend or header automatically); letters print Bold in their group's colour. Always name the cuts and the test in `note`.
- Heat-filled tables: `"heatmap": true` or `{"bands": 3–5, "columns": [1, 2, 3]}` fills cells with the sequential n-step row (white text on dark steps).
- Deltas: `"delta_columns": [4]` prints values starting with + / − in Bold green / red.

**Master decks (sales, company, pitch)**

| type | fields | use for |
|---|---|---|
| `columns` | `title`, `subtitle?`, `columns`: [{`headline`, `text`}], `grid?` ("2x2") | 2–4 parallel points |
| `ideas` | `title`, `ideas`: [{`headline`, `text`}] | up to 6 ideas |
| `panel` | `title` (two-line claim/question), `bullets` (≤ 4) or `body` | a claim with supporting points on Blue 10 |
| `donut_split` | `value` (73 or 0.73), `caption`, `bullets` (≤ 5) | one striking share |
| `timeline` | `title`, `subtitle?`, `events`: [{`title`, `date?`, `text?`}] | 4–7 events, history or milestones |
| `from_to` | `title`, `rows`: [[from, to]], `from_label?`, `to_label?` | 3–6 transformation pairs |
| `comparison` | `title` (question), `body`, `old`: {`items`}, `new`: {`items`} | the usual way vs Appinio (≤ 4 items each) |
| `stat_mosaic` | `statement`, `cards`: [{`icon` or `value`, `text`}] | company intro statement + 3 proof cards |
| `highlights` | `title`, `figures`: [{`value`, `label`}], `gradient?` | up to 6 headline figures on a gradient |
| `gradient_cards` | `title`, `subtitle?`, `cards`: [str], `gradients?` [names] | 3–5 values, pillars, series |
| `growth` | `title`, `subtitle?`, `steps`: [{`title` (2–3 words), `text` (≤ 12 words)}], `gradient?` | 3–6 rising bars: growth plans, maturity, roadmaps |
| `text_images` | `title`, `subtitle?`, `points`: [{`headline`, `text`}], `photos` (3) | capability slide with photo collage |
| `photos` | `title`, `items`: [{`photo`, `headline`, `text`}] | three photos with captions |
| `team` | `title?`, `people`: [{`name`, `role`, `email?`, `phone?`, `photo`}] | 3 people |
| `logo_wall` | `title?`, `logos`: [file paths] | official client logo files only, ≤ 15 |
| `thank_you` | `title` ("Thank you!" or "Q&A"), `photo?` | closing |
| `map` | `title`, `subtitle?`, `markers`: [{`region` (North America, Europe, Asia, South America, Oceania) or `x`/`y`, `label?`, `text?`, `value?` (share → ring instead of pill)}] | markets, offices, fieldwork coverage |
| `markets` | `title`, `subtitle?`, `markets`: [{`country`, `label?`, `value?`, `text?`}] | 2–5 markets with round flags (USA, UK, Spain, France, Germany only) |
| `case_study` | `headline`, `client_logo`, `facts` {Industry, Geography, Method}, `challenge`, `approach`, `impact_value`, `impact_text`, `quote?`, `quote_by?` | approved client results only |
| `disclaimer` | — | roadmap/product decks shared externally; Legal’s fixed text |

Examples: `examples/results-outline.json` and `examples/master-outline.json`.

## Content rules (Appinio voice)

- Confident, plain-spoken, outcome-driven. **Titles state the finding**, not the topic ("Brand A leads, but the field below it is flat", not "Brand usage"). Put the topic in `eyebrow` / `subtitle`.
- "You / your team" for the client, "we" for Appinio, never "I". Sentence case; never type all caps.
- Digits in stats ("2B+", "n = 1,204"); spell out one to nine in prose. En dash for ranges, curly quotes. No Oxford comma. **No emoji.** No bold or italics inside sentences.
- Tagline is verbatim, two lines: "Rigorous, rapid research." / "Refreshingly human." — the template places it.
- Results decks: every chart gets a base in `note`; cuts are lettered A, B, C in series order and named in the note.
- Never invent data, quotes, people, photos, client logos or results. If something is missing, ask or leave it out.

## Colour and charts

Single series → Data Blue; multi-series → categorical palette in order (automatic). For scales pass `colors` with the exact n-step row from `references/dataviz-scales.json` (count the classes first): 5-point Likert → `["8152E3","A383E5","DFE0E2","FF867E","F45A51"]`; NPS groups → `["FF6E5D","FFCD4C","88D298"]`; ordered bins → the sequential row for n. Blue marks one phrase or one number per slide, never more.

## Assets

List names with `python scripts/build_deck.py --list-assets photos` (or `icons`, `gradients`, `flags`); pass names without extension, or any approved file path.

- **Photos** (`assets/photos/`): the Appinio image bank, 353 photos named by topic, e.g. `person-working-from-home-04`, `group-cooking-eating-02`, `fmcg-07`, `banking-03`, `pets-05` (+ 10 template photos). `references/photo-catalog.json` lists topic and orientation for each; pick by the slide's subject, portrait photos for half-slide layouts. Never put text over photos except the photo cover and photo statement.
- **Icons** (`assets/icons/`): the full Appinio Blue set (178). Only Blue; if none fits, use none.
- **Gradients** (`assets/gradients/`): `cover`, `contact`, `lilac` (template defaults) plus the 15 brand gradients `beta-blaze`, `bloom-index`, `clarity-wave`, `clickstream`, `focus-glow`, `fresh-take`, `hype-haze`, `insight-storm`, `midnight-metric`, `mindshare`, `pulse-check`, `sample-spark`, `segmint`, `spectral-phase`, `trendspotter`. Dark blue ones (`bloom-index`, `clickstream`, `midnight-metric`, `mindshare`, `sample-spark`, `segmint`, `spectral-phase`) carry white text best. One gradient family per deck.
- **Flags** (`assets/flags/`): USA, UK, Spain, France, Germany; add more as round PNGs named after the country. **Client logos**: only official files the user supplies.

## References (read only when needed)

- `references/brand-book.md` — the Appinio Design System brand book (both deck styles, data-viz rules, imagery prompt).
- `references/layouts/*.md` — the design system's spec for each slide component (`Slide*` = results & shared, `Master*` = master template).
- `references/tokens.json`, `references/dataviz-scales.json` — exact colour values.
- `references/photo-catalog.json` — topic, orientation and source file of every bank photo.
