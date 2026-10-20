# Aptaworks presentation system

## Visual authority

The user-supplied [HTML reference](upstream.html) guides the visual direction. The generated [HTML](../examples/example.html) and [PPTX](../examples/example.pptx) are the current authority: the user requested clean backgrounds without the original paper grid. Preserve typography, palette, and composition while honoring that override.

Use bold Lato titles, editorial italic numerals, small monospaced labels, crisp rules, orange connectors, navy cover/dividers, and the consultation closing. Do not add a graph-paper grid to generated presentation backgrounds. Generated HTML and PPTX use Lato/Lato Black, Georgia italic, and Consolas. The original reference uses Newsreader/DM Mono in the display/metadata roles. Font files are not bundled or embedded; verify availability before delivery.

## Geometry and palette

| Role | Value |
|---|---|
| Canvas | 1280 × 720 CSS pixels; 13⅓ × 7½ inches; 16:9 |
| Content columns | x=84–1196 |
| Metadata / eyebrow | y=22 / y=54; compact labels above the title |
| Content title | x=84, y=82; up to two short lines |
| Component area | y=184–636 for one-line titles; y=244–636 for two-line titles |
| Page number / footer rule | y=656 / y=688 |
| Presentation background | Clean paper or navy gradient; no background grid |
| Navy / deep navy | `#17234E` / `#0B1533` |
| Orange | `#EF7636` |
| Paper / rule / muted | `#F6F3EA` / `#C9D2DF` / `#5E6B82` |
| Proof-cell fill | `#BEC1CD`, matching the upstream white-over-navy grid treatment |

PPTX geometry uses CSS pixels at 96 DPI. Font sizes are written as points (`px × .75`). Standard content and section titles are **48px / 36pt**; cover titles are **64px / 48pt**. Keep the content header compact and move components upward rather than reserving space for unused title lines. The header brand is simply **Aptaworks**. Do not add “Editorial” or an edition label to the cover, header, or footer. The consultation closing retains its reference-specific type scale and placement.

Body sizes follow each information type. The packaged glyph advance data in `assets/font-metrics.json` supports portable density checks without shipping fonts. It estimates wrapping; the slide renderer remains authoritative. HTML previews use the reference's bottom dot navigation: 7px dots, a 22px orange active pill, direct slide selection, keyboard, wheel, and swipe controls. Navigation is browser chrome and is omitted from PPTX and print output.

## Consultation closing: preserve the reference composition

Use this closing by default for Aptaworks decks unless the user asks for another ending. Respect a requested total slide count.

| Element | Placement at 1280×720 |
|---|---|
| Gradient | Vertical: `#182B63` at 0%, `#152B5B` at 48%, `#16A6C8` at 100% |
| Grid | Omitted by the user's current presentation-style override |
| Headline | Centered, x=84–1196, y=156, Lato Black 41.6px |
| Secondary headline | Centered, y=210, 31.6px, white at 78% |
| Consultation subtitle | Centered, y=328, 35.8px |
| QR card | x=569, y=447, 142×142px, white with a restrained rounded corner |
| QR image | Centered inside the card with 8px padding |
| Complete contact row | Original `assets/Footer.png`, x=80, y=610, 1120×56px |

Reference copy:

```text
Maximize your operational efficiency and business growth
with Aptaworks' cutting-edge data solutions.
Get a consultation with our experts at
```

The original footer artwork preserves the phone, Instagram, LinkedIn, website, and email row: `(+62) 811 862 2782`, `aptaworks`, `aptaworks`, `www.apta.id`, `info@apta.id`. Do not move the QR to a side column, replace the gradient with solid navy, reduce the contact row to two items, or add a cover logo/header to this closing. Headline text remains editable; the supplied QR and footer artwork remain images.

## Component library

The English example contains 13 slides; `process` is also accepted as an alias for `flow`.

| Layout | Appropriate use | Source capacity |
|---|---|---|
| cover | Strong title, brief context, large editorial watermark | About 3 title lines, 2 support lines |
| ledger | Agenda, options, maturity levels, comparison lists | 1–5 simple rows or 1–3 detailed rows with status |
| kpi | A compact, contextualized metric strip | 3–4 value/label/context cells |
| cards | Distinct workstreams or contained scope summaries | 2–3 square rule-framed panels |
| flow / process | A sequential, editable delivery map | 2–4 connected nodes |
| timeline | Dated phases, milestones, handoffs | 3–4 date/label/description items |
| section | A transition with large editorial numbering | Title and one support line |
| table | Workplan, scope, comparison, evidence | 2–4 columns, 2–6 rows including the header |
| columns | Context and response, two related narratives | Exactly two labeled explanations |
| proof | Capability or evidence grid | Exactly six value/label/context cells |
| image | Source artwork, screenshot, product visual with context | One local PNG and one caption |
| quote | A principle or attributed statement | One statement and one attribution |
| closing | Reference consultation and contact slide | Headline, secondary headline, consultation subtitle |

Use cards when content needs containment. Use ledger rows for simple lists; do not turn every list into repeated cards. Cards have square frames, numbered labels, and no repeated orange top strip. Flows and timelines use native editable nodes/lines and concise descriptions. Tables remain native PowerPoint tables. KPI figures must come from the brief or be clearly labeled illustrative; never present example numbers as client results.
