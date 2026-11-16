# SlideScaleLogic

Colour logic for scales: count the classes, then use that row of the sequential, diverging or traffic-light palette.

Reference card for the colour logic every chart follows (the Appinio Data visualization sheets). Also usable as a build-reference slide in a deck.

**The rule:**
1. Data type → palette: categorical (unrelated categories), sequential (ordered values, no midpoint), diverging (two poles and a midpoint — Likert), traffic light (good ↔ bad — NPS, thresholds), gender.
2. Count the classes n (a 5-point scale is 5; NPS groups 3; heatmap bands as many as you define; "Other / Don't know" does not count).
3. Use exactly the n-step row (filled cells in each row of this card; hex values in the README and in `dataviz/scales.json`). Never the first n colours of the 11. Categorical: first n, in order.

- Text on fills: white on the dark steps, ink on the light ones.
- Examples: 5-point Likert → diverging 5 = #8152E3 #A383E5 #DFE0E2 #FF867E #F45A51. NPS → traffic light 3 = #FF6E5D #FFCD4C #88D298. Four heatmap bands → sequential 4 = #96B3FF #5987FF #3463E0 #2648A3. Six ordered bins → sequential 6 = #D4E0FF #96B3FF #5987FF #3463E0 #2648A3 #172D66.


