# Are the Yi-family cabbage numbers realistic?

Background check requested 2026-08-28. **Outcome: the figures were
switched, and then independently confirmed by Nico's own calibration
spreadsheet**, `Cabbage_Production_Costs_2026_4.xlsx`.

The cost function in the deck is no longer "the MW numbers" — it is the
output of a sourced model. The spreadsheet builds it from the UF/IFAS
North-East Florida cabbage budget (August 2025), the BLS Occupational
Outlook Handbook (May 2024) for the family's forgone earnings, and USDA
NASS 2025 for yields and rents; the deck's figures are the spreadsheet's
"values in use" after its presentation rounding.

## The cost function now in the deck

    TC = 60,000 + 135·Q + 0.2·Q²        P = $400 per ton, then $210

| | Adopted (was MW) | My original |
|---|---|---|
| Fixed cost (TFC) | **$60,000** | $30,000 |
| Market price | **$400 / ton**, then $210 | $230 / ton, then $160 |
| Q\* at the first price | **662.5 tons** | 475 tons |
| ATC at Q\* | **$358.07 / ton** | $198 / ton |
| AVC at the second Q\* | **$172.50 / ton** | $100 / ton |
| Minimum ATC | **$354.09** (at Q = 547.7) | $195 (at Q = 387) |
| Implied acreage at Q\* | ≈ 33 acres | ≈ 23 acres |

Acreage uses the USDA 2024 US average cabbage yield, 406.5 cwt per acre
= 20.3 tons per acre.

## Why these, and not mine

**Prices received by growers, fresh-market cabbage, United States**
(USDA NASS, *Vegetables 2024 Summary*, February 2025, p. 20):

| Year | $ / cwt | $ / ton |
|---|---|---|
| 2022 | 28.80 | 576 |
| 2023 | 28.50 | 570 |
| 2024 | **31.10** | **622** |

The 2020 season-average price was $20.80 / cwt = $416 / ton (AgMRC, citing
USDA NASS).

**Production cost** — the UC study cited in my own speaker notes
(Takele, Daugovish & Vue, *Costs and Profitability Analysis for Cabbage
Production in the Oxnard Plain, Ventura County, 2012-13*, UC Cooperative
Extension):

- yield 1,050 cartons per acre × 45 lb = 47,250 lb = **23.6 tons per acre**
- total cost **$8,212 per acre** → **$348 per ton**
- break-even price $7.74 / carton on cash costs, $7.82 on total costs
  → **$344 – 348 per ton**
- the study's assumed market price, $8 per 45-lb carton, is **$355 per ton**

So the adopted ATC of $358 per ton sits on top of the $348 in the study,
and $400 per ton is close to USDA's 2020 average. My original pair (ATC
$198, price $230) was roughly half the study's cost and a third of the
current market price.

Both are still below today's market: a fully current version would use a
price nearer $600 per ton and a cost nearer $475 (the UC total carried
forward with CPI, roughly ×1.37 from 2013).

## Consistency check on the deck

Every printed figure is computed from the five constants at the top of
`_build_Module4.py` (`TFC`, `B_LIN`, `B_QUAD`, `P_HIGH`, `P_LOW`) — none is
typed twice. The identities hold exactly:

| Check | Result |
|---|---|
| MC(Q\*) = P at $400 | 400.000000 ✓ |
| MC(Q\*) = P at $210 | 210.000000 ✓ |
| TR − TC at Q\* = 662.5 | 265,000 − 237,218.75 = **27,781.25** |
| (P − ATC)·Q\* | (400 − 358.066)·662.5 = **27,781.25** ✓ same |
| TR − TC at Q\* = 187.5 | 39,375 − 92,343.75 = **−52,968.75** |
| (P − ATC)·Q\* | (210 − 492.5)·187.5 = **−52,968.75** ✓ same |
| MC cuts ATC at ATC's minimum | both 354.089023 at Q = 547.72 ✓ |
| Loss if operating vs. TFC | 52,968.75 < 60,000 → operate ✓ |
| P vs. AVC at Q\* = 187.5 | 210 > 172.5 → operate ✓ (same answer) |
| Q·(P − AVC) = the gap between the two losses | 187.5 × 37.5 = 7,031.25 = 60,000 − 52,968.75 ✓ |

The last line is the one the shut-down slide turns on: the algebra on
slide 30 (*Difference = Q·(P − AVC)*) is confirmed by the arithmetic on
slide 34.

**Cost table on slide 14** — computed, and identical to the table in the
MW deck:

| Q | TC | TFC | TVC |
|---|---|---|---|
| 50 | 67,250 | 60,000 | 7,250 |
| 200 | 95,000 | 60,000 | 35,000 |
| 400 | 146,000 | 60,000 | 86,000 |
| 600 | 213,000 | 60,000 | 153,000 |
| 800 | 296,000 | 60,000 | 236,000 |
| 1,000 | 395,000 | 60,000 | 335,000 |

## Cross-check against `Cabbage_Production_Costs_2026_4.xlsx`

The spreadsheet's *Assumptions* tab calibrates the cost function and then
offers a presentation rounding, which is switched on (cell B61 = 1):

| Spreadsheet cell | | Value | In the deck |
|---|---|---|---|
| B14 | TFC, computed | $59,620 | — |
| B65 | **TFC in use** | **$60,000** | `TFC` ✓ |
| B66 | **a in use** (linear, $/ton) | **135** | `B_LIN` ✓ |
| B67 | **b in use** (curvature) | **0.2** | `B_QUAD` ✓ |
| B86 | output at minimum ATC | 547.7226 tons | 547.7226 ✓ |
| B87 | minimum ATC | $354.089 | $354.089023 ✓ |

The cost table agrees at **every one of the twelve quantities** the
spreadsheet tabulates (50 to 1,100 tons) — TC, TFC and TVC all identical
to the cent. The deck's six-row table is a subset of those rows.

The spreadsheet documents the rounding rather than hiding it: TFC is
overstated by $380 (+0.64 %) and variable cost at the benchmark understated
by $1,560 (−0.61 %), so total cost at the 840-ton benchmark is $1,180
light — 0.37 %. Its own note says to keep that inside about 1 %.

### The one thing that does NOT match: the price

| | Spreadsheet | Deck |
|---|---|---|
| F.O.B. price (B77) | $473 / ton | — |
| broker fee (B78) | 10 % | — |
| **net price received (B79)** | **$425.70 / ton** | **$400 / ton** |
| profit-maximizing output (B84) | 726.75 tons | 662.5 tons |

The spreadsheet is explicit that this cell is a free scenario input
("PROFIT CHECK — SCENARIO PRICE (vary this freely; the cost curve stays
put)"), so the two are not in conflict — but they are not the same number.

**$400 per ton is defensible on its own.** The spreadsheet's own reference
values give Florida's 2025 season average as $428 per ton and the US
average as $488; after the same 10 % broker fee those are $385 and $439
net, so $400 sits between them. It is also round, which is why the worked
example comes out at TR = $265,000 and Q\* = 662.5 rather than
TR = $309,377 and Q\* = 726.75.

Two ways to make the two artifacts agree exactly, if that is wanted:

- set the deck's `P_HIGH` to 425.70 — one constant, but every worked figure
  becomes untidy; or
- set the spreadsheet's F.O.B. price (B77) to **444.44**, which nets
  exactly $400 — one cell, and the deck does not change at all.

The **low-price scenario** ($210 per ton) has no counterpart in the
spreadsheet, which models only one price at a time. It is consistent with
the *Current Prices* tab: the lowest August 2026 quote there is New York
round green in sacks at $6 per 50-lb sack, i.e. $240 per ton F.O.B. and
$216 net, so $210 reads as a price just under the current floor. At that
price the spreadsheet's shut-down test still says operate (net price minus
minimum AVC = 210 − 135 = 75 > 0), which is what slide 34 concludes.

## Still to do outside the deck

**The two PollEverywhere activities are keyed to the OLD answers.**
Nico is re-keying them himself; until then the six poll slides in the deck
are marked placeholders (see the Session Notes). The activities to re-enter:

| Poll | Currently offers | Should offer |
|---|---|---|
| "What is the optimal quantity produced?" (slides 21 and 32) | Q = 160 / 300 / 475 / 500 / none | the new Q\* = 662.5 as the right answer, with distractors around it |
| — | — | and for the second run, Q\* = 187.5 |

The embed URLs stay valid. Once the activities are re-keyed, running
`python _splice_media.py` puts the live slides back in place of the
placeholders.

## Sources

- `Cabbage_Production_Costs_2026_4.xlsx` (Nico, August 2026) — the calibration model the deck's cost function comes from; its *Sources* tab carries the full audit trail
- [USDA NASS, *Vegetables 2024 Summary* (February 2025)](https://www.nass.usda.gov/Publications/Todays_Reports/reports/vegean25.pdf) – cabbage price, yield and acreage tables, pp. 19 – 21
- [AgMRC, Cabbage commodity page](https://www.agmrc.org/commodities-products/vegetables/cabbage) – 2020 season-average price and yields
- [Takele, Daugovish & Vue, *Costs and Profitability Analysis for Cabbage Production in the Oxnard Plain, Ventura County, 2012-13*, UC Cooperative Extension](https://ucanr.edu/sites/Farm_Management/files/179212.pdf) – cost per acre, yield, break-even prices
