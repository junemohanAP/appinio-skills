# SlideSplitChart

Grouped bars by brand or cut: legend directly above, significance letters in their group colour, takeaway card.

Use for any result with two or more series: brands, age cuts, buyer groups.

**Provide:** categories, the values per series, significance letters, 2–3 takeaways, the base.

**Marking rules (every multi-series chart):**
- The legend sits **directly above the chart**, left-aligned with the bars, in series order. Its swatch matches the mark: a 20px rounded square for bars, a short line with a dot for line charts (`SlideProfileChart`).
- Series colours: the categorical palette strictly in order — `chart-1` Data Blue #779DFF, `chart-2` Apricot #FFAF75, `chart-3` Purple #8152E3, `chart-4` Green #88D298 … Cuts are lettered A, B, C in the same order, and the footer names them (e.g. "A 18–29, B 30–44, C 45–65").
- Significance letters follow the value ("64% A C") in Bold, **each letter in the colour of the group it refers to** (A = chart-1, B = chart-2, C = chart-3).
- Value labels 24px SemiBold at the end of each bar; category labels 24px Regular, left.
- Where a chart is split into panels (one per cut), each panel header is the cut's name in its series colour with a 2px rule of the same colour underneath.
- Quote attributions on split slides are coloured like their cut.
- NPS and other single scores sit in a blue (`blue-100`) pill with white Bold text, with the column header "NPS" as an eyebrow.
- Gaps and deltas in tables: positive in `tl-10` #4C795B, negative in `tl-3` #D0434A, Bold.
- Scales follow the colour logic (`SlideScaleLogic`): count the classes, use that row. NPS = traffic light 3 (#FF6E5D, #FFCD4C, #88D298); 5-point Likert = diverging 5 (#8152E3, #A383E5, #DFE0E2, #FF867E, #F45A51); heatmap tables = the sequential row for the number of bands, white text on the dark steps.
- Takeaway card on the right: Grey 10, radius 16, padding 40; each heading an eyebrow (dark grey), items split by 1px hairlines. The chart and card fill the area down to the footer divider.


