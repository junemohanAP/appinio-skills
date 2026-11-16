Appinio is the AI-powered market research platform that helps teams make better decisions with actionable consumer insights. This system covers decks, documents and on-brand visuals. **Read the Presentations section before building any slide.**

Tagline (used verbatim, all three full stops): *Rigorous, rapid research. Refreshingly human.*

## Content fundamentals

- **Voice.** Confident, plain-spoken, outcome-driven. We sell decisions, not data: headlines state the finding, not the topic ("Brand A leads, but the field below it is flat", not "Brand usage").
- **Tone principles.** Stay human and authentic · simplify complex ideas · be brief and get to the point · adapt to the audience · be positive and approachable.
- **Person.** "You" and "your team"; "we" for Appinio. Never "I".
- **Case.** Sentence case everywhere: titles, labels, buttons, chart legends. **Never typed in all caps**; on slides, eyebrows and card labels are displayed uppercase via `text-transform`. Proper nouns keep their case (Appinio Boards, Appinio Conversations).
- **Numbers.** Digits in stats, charts and UI ("2B+ opinions", "n = 1,204"). Spell out one to nine in running prose. Tabular figures in tables.
- **Punctuation.** Curly apostrophes, en dash for ranges (12–19 March), em dash for asides. No Oxford comma.
- **Emoji.** Never.
- **Real lines:** "Teams make better decisions with Appinio" · "Insights for everyone" · CTAs: *Get free access*, *Book a demo*.

## Colour

**Proportion is the rule: 60 / 25 / 15.** Most of any slide, page or deck is neutral, so the brand colours can shine and mean something.

| Share | Colours | How it shows up |
| --- | --- | --- |
| 60% neutrals | `white` (most of it), `grey-900`, `blue-10`, `grey-10`, `beige` | Grounds, text, boxes. White is the default ground. |
| 25% Appinio Blue | `blue-100` | The highlighted phrase, the key number, the logo on light, primary marks. |
| 15% secondary | `green-100`, `orange-100` (roughly equal) | Small accents, shapes, illustration. Never text, never a full ground. |

- Use `blue-100` **for emphasis only**. Never colour eyebrows, every bullet, every border or every bar blue. If more than a quarter of a slide is blue, cut blue.
- Never use `blue-50`, `blue-25` or `blue-10` as decoration to "add brand". `blue-10` is for one callout panel per slide at most.
- `green-100` (2.5:1) and `orange-100` (3.0:1) fail as text on white. Use them only as fills behind nothing, or as 24px+ bold numerals in orange.
- `pink-*` belongs to **Appinio AI** and `purple-*` to **Appinio Labs**. Use them only when talking about those products.
- Text is `ink` (`grey-900`) on light grounds and `ink-inverse` (white) on dark ones. Use `ink-muted` (`grey-500`) only for footnotes, sources and base sizes. `grey-300` is never text.
- Status colours (`success`, `warning`, `danger`) are for product UI states only, never for brand expression.

## Typography

- **Decks use Inter**, exactly as the approved *Appinio Slide Deck Template 2026*: it works in Google Slides and PowerPoint for everyone. Load it in a deck from Google Fonts (`family=Inter:wght@400;500;600;700;800`).
- **Aeonik** is the brand face for designed marketing material (web, print, social). Inter is its fallback everywhere Aeonik is unavailable. Never use any other face (no Manrope, no system fonts as a design choice).
- **Hierarchy comes from size and a few heavy weights.** Slide titles are Inter Bold (700) with −0.5px tracking (display sizes Bold or ExtraBold); body copy is always Regular (400) in `grey-700`, which reads softer than pure ink; intros and point headings are SemiBold (600); eyebrows and labels are Bold uppercase in dark grey (`grey-700`). Never bold words inside a sentence.
- **Emphasis is colour.** Highlight the one phrase that matters in `highlight` (#0038FF) on white, `highlight-on-dark` (#BFCDFF) on dark grounds. No italics, underlines or size changes inside a sentence.
- Flush left by default. The template centres only big-number columns and the two-line statement slide.
- Nothing on a slide is smaller than 24px (the template's 19–20px footer text is raised to 24px).

## Presentations

The **Appinio Slide Deck Template 2026 (Light Mode)** — the master template — is the source of truth. Every deck is one of two kinds, and the kind decides the style. **Decide it first.**

| | **Results deck** | **Master-style deck** (everything else) |
| --- | --- | --- |
| What | Research results: study findings, question-type charts, methods (KDA, TURF, pricing, conjoint …), mental availability, tracking | Sales and pitch decks, company intros, proposals, capabilities, case studies, events, team, internal and thought-leadership decks |
| Components | Group **Slides · results & shared** | Group **Slides · master template**, plus the shared cover, contents, statement, quote, contact and back cover |
| Small heading | Dark grey (`grey-700`) uppercase eyebrow *above* the title (question type, cut) | Grey (`grey-300`) 45px ExtraBold subtitle *under* the title — no eyebrow |
| Chapter divider | `SlideChapter`: large number, title and description at mid-height | `MasterChapter`: "01" in #CCCCCC and the title on one line at the top |
| Footer | 1px divider at y 900, note under it, logo below | No divider: logo bottom-left at y 983, note at x 685 on the same row |
| Boxes | Grey 10 cards; takeaway panel beside each chart | Blue 10 (#E5EBFF) panels with white cards on them |
| Charts | Full data-visualisation rules below (palettes, n-step scales, significance letters) | Same palettes; keep charts simple — one message, no significance letters |

A deck never mixes the two styles. A results deck may borrow a master layout (e.g. `MasterTeam` for contacts, `MasterFlags` for markets), but restyles it to results conventions (eyebrow, divider footer).

Each component's README carries the exact markup in the Slides format (inline styles only, no classes). Build decks by copying that markup and changing words, data and pictures, nothing else.

**Master-style components:** `MasterCoverPhoto` (photo cover), `MasterChapter`, `MasterAgendaLong` (7–20 chapters), `MasterTextImages` (points + photo collage), `MasterComparison` (old way vs Appinio on a dark panel), `MasterFromTo`, `MasterStatMosaic` (statement + proof cards on Blue 10), `MasterHighlights` (six figures on a gradient), `MasterTimeline`, `MasterProcess` (3–5 steps), `MasterTable` (pricing, packages), `MasterDonutSplit`, `MasterGradientCards` (values, pillars), `MasterTeam`, `MasterTestimonial`, `MasterLogoWall`, `MasterCaseStudy`, `MasterMap`, `MasterStatementPhoto`, `MasterThankYou` (also Q&A), `MasterDisclaimer`, `MasterFlags`. Shared with results decks: `SlideCover`, `SlideCoBrandCover`, `SlideAgenda`, `SlideContent`, `SlideKeyFindings`, `SlideStat`, `SlideStatement`, `SlideImage`, `SlideQuote`, `SlideIconPoints`, `SlideFactCards`, `SlideContact`, `SlideBackCover` — in a master-style deck use them with the master title, footer and Blue 10 panels.

**Grid.** Margin 67px on all four sides (`slide-margin`). Two columns: left x 67–927, right x 993–1853 (`slide-gutter` 66px). Three columns: 551px each, 66px apart. Stay inside the grid.

**Type on slides** (Inter, all values at 1920 × 1080):

| Role | Style | Size / weight | Colour |
| --- | --- | --- | --- |
| Cover title | `slide-cover-title` | 171px Bold, lh 0.96 | white |
| Chapter title, statement | `slide-display` | 107px Bold | white on gradient / ink |
| Big number, "Key Findings" | `slide-big-number` | 133px Bold (ExtraBold for Key Findings) | ink |
| Quote | `slide-quote` | 72px Bold | ink |
| Message title, "Agenda" | `slide-title-lg` | 61px ExtraBold | ink |
| Eyebrow above a title, card and takeaway labels | `slide-eyebrow` | 24px Bold, uppercase, 2px tracking | `grey-700` #39404C |
| Slide title | `slide-title` | 45px Bold, −0.5px tracking | ink |
| Agenda line | `slide-agenda` | 43px SemiBold | ink, number #CCCCCC |
| Intro, point heading, number meaning | `slide-lead` | 32px SemiBold | ink |
| Body and bullets | `slide-body` | 24px Regular, lh 1.4 | `grey-700` #39404C |
| Label above a number or fact | `slide-label` | 24px Bold, uppercase, 2px tracking | `grey-700` |
| Footer note, base, source | `slide-footnote` | 24px Regular | `grey-500` #5F6675 |

**Slide-format limits.** Section padding can be at most 256px, so never use padding to push text into the right column: wrap the text in a block with `width:860px`, or pin it (`position:absolute; left:993px`). Text is never under 24px.

**Every content slide in a results deck** (master-style decks: see the table above and each `Master*` README):
- A dark grey eyebrow (question type, cut or section) above the title at x 67, y 67; title below it. No grey subtitle.
- Footer: a 1px `grey-100` (#D2D6DB) divider at x 67, y 900, 1786 wide (860 on half-photo slides); the footer note (base n =, source or legend, 24px `grey-500`) at x 67, y 914; the dark logo at x 67, y 990, 144 × 38. Every content slide has the divider.
- Section padding `67px 67px 210px`, body wrapped in a `flex:1` column so it **fills the space between the title and the divider**. Card rows, chart panels and tables stretch (`flex:1`, grid rows `1fr`). No empty white band under the content: if content is short, make the cards taller or the type larger, never leave a gap.
- Group points in **Box cards** (`grey-10` #F5F6F7 fill, radius 16, padding 40, 32px gaps), with a blue icon at the top where one fits (see `SlideFactCards`, `SlideIconPoints`, `SlideStat`). Takeaways sit in a Box panel beside the chart.
- Titles in sentence case. Eyebrows and card / takeaway labels are short (1–4 words), written in sentence case in the source and shown uppercase with `text-transform:uppercase`.

**Approved layouts.**
- **Chapter** (`SlideChapter`): left side, mid-height. Large number (200px, white at 55%), the chapter title (107px Bold) below it, a one-line description (40px) below that. Never number and title side by side at the top.
- **Contents** (`SlideAgenda`): "Contents" title, then two columns of rows: blue number, title, page number right-aligned in grey, a 1px divider under each row. A variant with a photo per chapter is fine for short decks.

**Tagline.** "Rigorous, rapid research." and "Refreshingly human." are always written as two whole lines, each `white-space:nowrap`. Never break inside "Rigorous, rapid research."

**Every deck:**
1. `SlideCover` (or `SlideCoBrandCover` with a partner): a gradient from the Gradients group, white logo top-left 280 × 73, 171px title.
2. `SlideAgenda` for three or more chapters.
3. Per chapter: `SlideChapter` on a gradient (the same gradient for all chapters), then optionally `SlideSectionIntro`.
4. Content: `SlideContent`, `SlideFactCards`, `SlideSplitChart`, `SlideProfileChart`, `SlideKeyFindings`, `SlideStat`, `SlideChart`, `SlideImage`, `SlideQuote`, `SlideIconPoints`, `SlideStatement`.
5. Optionally `SlideContact`, then always `SlideBackCover` (gradient, tagline, office addresses).

**Colour proportion (60 / 25 / 15).** Content slides are white. Gradients are for cover, chapter and back-cover slides and the contact panel only, never behind charts or body text. Blue is the highlighted phrase, the chart's lead series and the icons; green and orange appear only in charts and gradients.

**Gradients.** Use only files from the Gradients group, full-bleed, as the first element of the slide so text paints over them; text on them is white. Never recolour them, crop them to a sliver or put photos over them. Never draw a CSS gradient instead.

**Pictures.** Use photos from the Imagery group (or the full Appinio image bank). Text never sits on a photo, except the two master layouts built for it (`MasterCoverPhoto`, `MasterStatementPhoto`), which add an ink scrim. Half-slide photos bleed off the edge with square corners (x 993–1920 or 0–927, full height); inset photos and portraits use `radius-image` (32px). Never put text over a photo. If you must generate an image, use the brand prompt under Imagery.

**Flags and maps.** Countries are shown only with the round flags in the **Flags** group (USA, UK, Spain, France, Germany — add others to the group before using them); maps use the Blue 10 world map in the **Maps** group. Never draw or download flags.

**Icons.** Use ONLY the Icons Blue group on white slides (Icons Grey for monochrome needs). Each SVG already includes its round tinted background; show it at 75–96px. Never draw, generate or borrow icons. If no icon fits, use none.

**Charts — data-visualisation palettes** (from the Appinio *Data visualization* sheets; tokens `chart-*`, `seq-*`, `div-*`, `tl-*`, `gender-*`):

| Palette | Use for | Colours |
| --- | --- | --- |
| Categorical `chart-1`…`chart-20` | Unrelated categories: brands, cuts, codes. **Apply strictly in order** 01→20. | 01 Data Blue #779DFF · 02 Apricot #FFAF75 · 03 Purple #8152E3 · 04 Green #88D298 · 05 Red #CD4E5F · 06 Yellow #FFCD4C · 07 Dusty Pink #ECA1B8 · 08 Teal #5E9A8F · 09 Lilac #B997EF · 10 Orange #FF7D34 · 11 Ash #6E7AA5 · 12 Neon Red #FF6C63 · 13 Ski Blue #6285E1 · 14 Turquoise #8BD9D0 · 15 Periwinkle #AEBBFF · 16 Fuchsia #D1519E · 17 Mint Green #AACBB7 · 18 Peach #FF9882 · 19 Ochre #F4B13B · 20 Grass Green #67AC65 |
| Sequential `seq-1`…`seq-11` | Ordered / numeric values light → dark: heatmap fills, bins, intensity. Darkest = largest. | #D4E0FF #B5C9FF #96B3FF #779DFF #5987FF #3B71FF #3463E0 #2D56C2 #2648A3 #1E3B84 #172D66 |
| Diverging (neutral) `div-1`…`div-11` | Two extremes with a midpoint: Likert, agree ↔ disagree. Purple end = positive, red end = negative, `div-6` neutral. | #8152E3 #926BE4 #A383E5 #B49CE6 #C5B5E7 #DFE0E2 #FFB9B5 #FF9F9A #FF867E #FF6C63 #F45A51 |
| Traffic light `tl-1`…`tl-11` | Good vs bad: NPS, pass / watch / fail, positive vs negative. | #6B032E #A6293E #D0434A #FF6E5D #FF9052 #FFCD4C #ACD182 #88D298 #67A176 #4C795B #2C4B3C |
| Gender `gender-*` | Any gender split. | Total #779DFF · Female #8152E3 · Male #FF7D34 · Non-binary #FFCD4C |

**Colour logic — count the classes, then use that row.** Every chart picks its colours in three steps:
1. **Data type → palette.** Unrelated categories → categorical. Ordered values with no midpoint → sequential. Two poles with a midpoint (agree ↔ disagree) → diverging. Good ↔ bad → traffic light. Gender → gender palette.
2. **Count the classes n.** A 5-point scale is n = 5, NPS groups n = 3, four heatmap bands n = 4, six spend bins n = 6. "Other / Don't know / None" stays `chart-other` and does not count.
3. **Use the n-step row exactly** (below; also in `dataviz/scales.json`). Never take the first n colours of a sequential, diverging or traffic-light palette — the rows skip steps so neighbouring classes stay distinguishable. Categorical uses the first n colours, strictly in order.

Text on a fill: white on the dark steps (`seq` from #3463E0, `div` #8152E3, `tl` #6B032E–#D0434A and #4C795B–#2C4B3C), ink on the rest. Significance letters on a filled table cell take the cell's text colour (Bold); on white they take their group's series colour.

Typical question types: 5-point Likert → diverging 5 · 7-point bipolar → diverging 7 · NPS groups → traffic light 3 · 0–10 NPS distribution → traffic light 11 · heatmap tables → sequential 3–5 · frequency or ordered numeric bins → sequential n · brands / cuts / codes → categorical n.

*Sequential (light → dark):*

| n | Colours |
| --- | --- |
| 3 | `#779DFF` `#3B71FF` `#2D56C2` |
| 4 | `#96B3FF` `#5987FF` `#3463E0` `#2648A3` |
| 5 | `#B5C9FF` `#779DFF` `#3B71FF` `#2D56C2` `#1E3B84` |
| 6 | `#D4E0FF` `#96B3FF` `#5987FF` `#3463E0` `#2648A3` `#172D66` |
| 7 | `#D4E0FF` `#96B3FF` `#5987FF` `#3B71FF` `#3463E0` `#2648A3` `#172D66` |
| 8 | `#D4E0FF` `#B5C9FF` `#779DFF` `#5987FF` `#3463E0` `#2D56C2` `#1E3B84` `#172D66` |
| 9 | `#D4E0FF` `#B5C9FF` `#779DFF` `#5987FF` `#3B71FF` `#3463E0` `#2D56C2` `#1E3B84` `#172D66` |
| 10 | `#D4E0FF` `#B5C9FF` `#96B3FF` `#779DFF` `#5987FF` `#3463E0` `#2D56C2` `#2648A3` `#1E3B84` `#172D66` |
| 11 | `#D4E0FF` `#B5C9FF` `#96B3FF` `#779DFF` `#5987FF` `#3B71FF` `#3463E0` `#2D56C2` `#2648A3` `#1E3B84` `#172D66` |

*Diverging, neutral centre (positive pole first):*

| n | Colours |
| --- | --- |
| 2 | `#8152E3` `#F45A51` |
| 3 | `#8152E3` `#DFE0E2` `#F45A51` |
| 4 | `#8152E3` `#A383E5` `#FF867E` `#F45A51` |
| 5 | `#8152E3` `#A383E5` `#DFE0E2` `#FF867E` `#F45A51` |
| 6 | `#8152E3` `#A383E5` `#C5B5E7` `#FFB9B5` `#FF867E` `#F45A51` |
| 7 | `#8152E3` `#A383E5` `#C5B5E7` `#DFE0E2` `#FFB9B5` `#FF867E` `#F45A51` |
| 8 | `#8152E3` `#A383E5` `#B49CE6` `#C5B5E7` `#FFB9B5` `#FF9F9A` `#FF867E` `#F45A51` |
| 9 | `#8152E3` `#A383E5` `#B49CE6` `#C5B5E7` `#DFE0E2` `#FFB9B5` `#FF9F9A` `#FF867E` `#F45A51` |
| 10 | `#8152E3` `#926BE4` `#A383E5` `#B49CE6` `#C5B5E7` `#FFB9B5` `#FF9F9A` `#FF867E` `#FF6C63` `#F45A51` |
| 11 | `#8152E3` `#926BE4` `#A383E5` `#B49CE6` `#C5B5E7` `#DFE0E2` `#FFB9B5` `#FF9F9A` `#FF867E` `#FF6C63` `#F45A51` |

*Traffic light (bad → good):*

| n | Colours |
| --- | --- |
| 2 | `#FF6E5D` `#88D298` |
| 3 | `#FF6E5D` `#FFCD4C` `#88D298` |
| 4 | `#A6293E` `#FF6E5D` `#88D298` `#4C795B` |
| 5 | `#A6293E` `#FF6E5D` `#FFCD4C` `#88D298` `#4C795B` |
| 6 | `#A6293E` `#D0434A` `#FF6E5D` `#88D298` `#67A176` `#4C795B` |
| 7 | `#A6293E` `#D0434A` `#FF6E5D` `#FFCD4C` `#88D298` `#67A176` `#4C795B` |
| 8 | `#A6293E` `#D0434A` `#FF6E5D` `#FF9052` `#ACD182` `#88D298` `#67A176` `#4C795B` |
| 9 | `#A6293E` `#D0434A` `#FF6E5D` `#FF9052` `#FFCD4C` `#ACD182` `#88D298` `#67A176` `#4C795B` |
| 10 | `#6B032E` `#A6293E` `#D0434A` `#FF6E5D` `#FF9052` `#ACD182` `#88D298` `#67A176` `#4C795B` `#2C4B3C` |
| 11 | `#6B032E` `#A6293E` `#D0434A` `#FF6E5D` `#FF9052` `#FFCD4C` `#ACD182` `#88D298` `#67A176` `#4C795B` `#2C4B3C` | Value labels 24px in ink. No 3D, shadows or gradients in charts.

**Marking every chart** (see `SlideSplitChart`, `SlideProfileChart` and `SlideScaleLogic`):
- The legend sits directly above the chart, aligned with the plot's left edge, in series order; its swatch matches the mark (square for bars, line-with-dot for lines).
- Cuts are lettered A, B, C … in series order and named in the footer. Significance letters follow the value in Bold, each letter in its group's series colour ("64% A C").
- Split panels: each panel's header is the cut name in its series colour, with a 2px rule of that colour under it. Quote attributions on split slides take the same colour.
- Profiles, semantic differentials and trends are **connected lines** (an `<svg>` polyline with dots) on a light grey panel with gridlines at the scale's inner values; pole or axis labels are HTML text beside the SVG, never `<text>` inside it.
- NPS and other headline scores sit in a blue pill (white Bold text); gaps and deltas are Bold `tl-10` #4C795B when positive, `tl-3` #D0434A when negative.

**Using this system's pictures in a deck.** Gradients, photos, icons and logos are stored as assets here; each group's README lists every file with its asset id. Copy the ones you use into the deck with one server-side copy (Artifact `publish` with `asset: true`, `from_url` = this system, `asset_ids` = up to 10 ids per call) and use the new `/_blob/<id>` urls it returns.

## Imagery

Start from the Imagery group (one approved photo per category of the Appinio image bank). Photography is warm and authentic: real people and real moments, diverse in gender, age and ethnicity, never staged stock. AI images are allowed but must look natural (check faces, hands and lighting). Use Midjourney-style generation only for concepts that are hard to photograph. Images support the story; pair each with text.

**Every generated image uses this prompt, filling only the bracket:**

> Documentary photography, eye-level, of [subject — who, doing what, where], real everyday moment, diverse and natural-looking people, warm and authentic mood, soft natural lighting with gentle shadows, bright and sharp, slightly desaturated with a cool-leaning grade, no grain, no text, no logos. 28mm, f/11, Kodak portrait 400.

Never ask for studio backdrops, dramatic lighting, futuristic or "tech" visuals, illustrations, or stylised 3D. Show inset images with `radius-image` corners; half-slide images bleed with square edges.

## Logo

- Use the files in the Logos asset group. Never redraw, recolour, stretch or add effects.
- `appinio-logo-dark.svg` on light grounds and `appinio-logo-white.svg` on `grey-900`, `blue-100`, gradients and photos. Use the monogram only where the wordmark would be too small (footers under 120px wide, favicons).
- Place the logo left: top-left on covers (280 × 73), bottom-left on every other slide (144 × 38 at x 67, y 990 under the footer divider). Keep clear space equal to the height of the "o" on all sides.
- **Co-branding:** Appinio logo first, then a 2px divider, then the partner logo at the same height on the same centre line (see `SlideCoBrandCover`). Neither logo is bigger.

## Visual foundations

- **Spacing** is a 4pt grid (`space-*`). Product UI sections use 24–32px padding; marketing sections 80–100px.
- **Radii.** `radius-md` for buttons and inputs, `radius-lg` for cards, `radius-xl` for marketing feature cards, `radius-image` for all imagery, `radius-pill` for status capsules.
- **Borders** are 1px `line`. Use `focus` (blue-100) for focus and selected states.
- **Shadows** (`shadow-*`) are on Blue 200 ink and for product UI only. Slides use no shadows.
- **Motion** is quick and confident: 200ms, ease-out `cubic-bezier(.22,1,.36,1)`. No bounce or overshoot.

## Iconography

The official Appinio icon library has 178 outline icons on round tinted discs, in six colour collections. This system holds **Icons Blue** (Appinio Blue on Blue 10, the default) and **Icons Grey** (Grey 700 on Grey 10, for monochrome use). The Green, Orange, Pink (Appinio AI) and Purple (Appinio Labs) collections and the animated icons exist in the shared library and can be added the same way. Emoji and Unicode symbols are never icons.
