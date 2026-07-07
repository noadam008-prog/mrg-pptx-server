# Design enhancement layer (PptxGenJS + raw DrawingML)

The render server builds decks with **PptxGenJS**, then optionally runs a **lossless
post-processor** (`enhance-pptx.mjs`) that injects raw OOXML design effects PptxGenJS
can't produce: **multi-stop gradients, pattern fills, glow, soft edge, drop shadow,
reflection, and gradient-filled text.**

It edits only the target slide's XML inside the `.pptx` zip; every other part
(charts, embedded workbooks, images, media) is repackaged byte-for-byte. Verified: a
same-buffer round-trip changes only the targeted slide and nothing else. (This is why
it beats routing the finished deck through python-pptx, whose full re-save silently
drops chart data-workbooks and media parts.)

## How to use it

1. **Tag the shape** you want to style when it's built, with `objectName`:

   ```js
   slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: 'CCCCCC' }, objectName: 'enh:hero' });
   slide.addText('TITLE', { x, y, w, h, color: '008C64', objectName: 'enh:title' });
   ```

   Multiple shapes may share one `objectName` — the enhancement applies to all of them.

2. **Pass an `enhancements` array** in the `/generate-pptx` request body. No-op if omitted,
   so existing callers are unaffected.

   ```jsonc
   {
     // ...normal payload...
     "enhancements": [
       { "objectName": "enh:hero",
         "fill": { "type": "gradient",
                   "stops": [ {"pos":0,"color":"00E0A0"}, {"pos":0.5,"color":"008C64"}, {"pos":1,"color":"222A30"} ],
                   "angle": 45 },
         "shadow": { "color":"000000", "blur":6, "dist":3, "angle":45, "alpha":40 } },

       { "objectName": "enh:title",
         "textGradient": { "stops": [ {"pos":0,"color":"FFFFFF"}, {"pos":1,"color":"B9F6D8"} ], "angle":90 } }
     ]
   }
   ```

## Effect reference

| key | shape | fields |
|---|---|---|
| `fill.type:"gradient"` | any | `stops:[{pos 0..1, color}]`, `angle` (deg) |
| `fill.type:"pattern"`  | any | `preset` (e.g. `wave`, `pct50`, `dnDiag`), `fore`, `back` |
| `fill.type:"solid"` / `"none"` | any | `color` / — |
| `glow`       | any | `color`, `size` (pt), `alpha` (0–100) |
| `shadow`     | any | `color`, `blur` (pt), `dist` (pt), `angle` (deg), `alpha` |
| `softEdge`   | any | `size` (pt) |
| `reflection` | any | `true` |
| `textGradient` | text | `stops`, `angle` — gradient across the glyphs |

Colours are hex without `#`. Unknown `objectName`s are logged as warnings and skipped;
if the whole step throws, the server serves the un-enhanced deck rather than failing.

## Built-in targets on the financial deck

Already tagged in `pptx-api-server.mjs`, ready to enhance:

- `enh:coverTitle` — the "Monthly Financials Metrics" cover title (try `textGradient`)
- `enh:kpiTile` — the actual-vs-budget KPI tiles (try gradient `fill` + `shadow`)
- `enh:takeawayLabel` — the black "Key Takeaways" label (try gradient `fill` + `glow`)

## Limits

Solid/gradient/pattern fills, and the effect list above, cover most "make it look
designed" needs. Freeform custom geometry still isn't available (neither library nor
this layer builds it); add a new effect by emitting its DrawingML in `enhance-pptx.mjs`.
