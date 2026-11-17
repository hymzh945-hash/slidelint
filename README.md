# SlideLint

Find what breaks a deck before the room does.

SlideLint reads the `.pptx` XML and reports the problems that actually show up on screen
and in front of a client: text that overflows its box, shapes hanging off the canvas,
overlapping text boxes, six font sizes on one slide, placeholder copy nobody deleted,
titles missing, images with no alt text, and slides nobody can present because there are
no speaker notes.

It does **not** render the deck. There is no screenshotting, no headless PowerPoint, no
cloud upload. It reads the same XML PowerPoint writes, so it runs in a second and works
offline.

## Install

```bash
pip install slidelint
```

Requires Python 3.9+.

## Use

```bash
slidelint audit -i deck.pptx
```

```
SlideLint 0.1.0 · board_review.pptx
  score      26/100  Will embarrass you live
  findings   25 (3 high, 12 medium, 10 low)
  slides     6 · layouts 11
  - placeholder text          1  [high]
  - text overflow             1  [high]
  - off slide                 1  [high]
  - overlapping text          1  [medium]
  - font outlier              4  [medium]
```

Checks, in plain language:

| Check | What it catches | Severity |
|---|---|---|
| `placeholder_text` | "Click to add text", Lorem ipsum, TODO, TBD left in | high |
| `text_overflow` | more characters in a box than the box can hold at that font size | high |
| `off_slide` | shape left/top/width puts content outside the canvas | high |
| `overlapping_text` | two text boxes covering each other | medium |
| `font_outlier` | a font used exactly once in the whole deck | medium |
| `font_size_sprawl` | more than four font sizes on one slide | medium |
| `tiny_text` | text under 12pt | medium |
| `missing_title` | a slide with no title text | medium |
| `duplicate_title` | the same title on two slides | low |
| `reading_load` | more than 90 words on one slide | low |
| `missing_alt_text` | image alt text empty or just a file name | medium |
| `no_speaker_notes` | slide with no notes | low |
| `unterminated_sentence` | long sentence with no closing punctuation | low |

## Score

The score is `100 * exp(-penalty / 120)`, where each check contributes
`severity weight × min(count, cap)`. High is 20, medium 8, low 2. Caps stop one noisy
check from flattening the whole result. 85+ ships, 65–85 is an hour of fixes, below 65
will embarrass you live.

## Try it without your own deck

```bash
slidelint demo -o slidelint-demo
```

Builds `demo_deck.pptx` — a deck with the mistakes people really send — then audits it
and writes `slidelint-report.html`.

## Free tier and Pro

The audit itself, the score and the per-check summary are free and stay free.

Pro adds the HTML report, CSV export for a review tracker, and a signed licence verified
offline (Ed25519, no server call, no seat phone-home). It is sold at
<https://hymz.gumroad.com>.

```bash
slidelint audit -i deck.pptx --report deck.html --csv findings.csv
slidelint license
```

## Licence checks are offline

The Pro key is a JSON file signed with Ed25519. Verification happens locally against a
public key compiled into the package — so the tool never calls home, and works on a
plane or inside a client's network.

## Built on

`python-pptx` (MIT) for the OOXML layer. SlideLint adds the checks, the scoring, the
report and the licence layer. See `NOTICE`.

## Licence

Business source: free to use, including commercially, on your own decks. Redistribution
or resale of the source requires a commercial licence. `LICENSE` holds the full text.
