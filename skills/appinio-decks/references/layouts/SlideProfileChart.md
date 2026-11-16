# SlideProfileChart

Semantic differential / profile: connected lines per brand on a Grey panel, poles left and right, scale below.

Use for semantic differentials, brand profiles, and any trend or profile that should read as a line.

**Provide:** the word pairs (negative pole left, positive right), each brand's mean per pair (1–5), 2–3 takeaways, the base.

- One `<svg>` (700 × 560) on a light grey panel (#F5F6F8, radius 12): vertical gridlines at 2, 3 and 4 in `grey-50`; per brand a `<polyline>` (6px, round joins) through the means, with a 14px-radius dot at each point. Draw the lead brand last so it sits on top.
- Pole labels are HTML `<p>`s (28px SemiBold, ink) pinned at each row's y — never `<text>` inside the SVG. The 1–5 scale sits under the panel, each number centred on its x.
- Series colours: categorical order — lead brand `chart-1` Data Blue, then `chart-2` Apricot, `chart-3` Purple. The legend sits directly above the plot, aligned with its left edge, using the line-with-dot swatch.
- To change values: x = (mean − 1) / 4 × 700; y of row i = 56 + i × 112 (five rows). Keep the same y values for the labels.


