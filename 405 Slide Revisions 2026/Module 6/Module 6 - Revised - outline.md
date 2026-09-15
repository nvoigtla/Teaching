# Module 6 – Revised: deck outline

**Sources**
- `Module 6 - NV Slides/Module 6.pptx` (76 slides, 4:3, Nov 2024) – the full
  lecture deck; the union of the video and in-class material.
- `Module 6 - NV Slides/Module 6 -- Slides On-Campus Applications with
  Solutions.pptx` (52 slides, Mar 2022) – the applications-class deck.
- `Module 6 - NV Slides/Module 6 - Video 1…8.pptx` (8 / 6 / 10 / 8 / 7 / 5 /
  4 / 2 slides) – the taped videos. **These pin the video block boundaries.**
- `Module 6 - PG Slides/Module 6 - PG - Part 1 / 2 / 3.pptx` (43 / 30 / 31
  slides, 16:9) – comparison decks, consulted per the adoption protocol.

**Targets**
1. `Module 6 - Revised.pptx` – 13.33 × 7.5", single master, deck-standard
   chrome, native charts / tables / OMML, fade builds. Built as a **taped**
   module: 8 video title cards, `Module 6 · Video k · <topic>` tags, gold
   coverage pills, plus a full In-Class Examples block.
2. `Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx` – a
   separate 10-slide deck carrying the **full PG BMW example**, for retaping
   the practice video.

Status: **outline v3, settled — building against it.** Every video block
boundary below rests on a slide-level match against the eight taped decks.

---

## 0. Decisions taken (your answers, 2026-09-09)

1. **Taped now** – built as a taped module in one pass.
2. **Full In-Class Examples block**, and it **keeps the applications deck's
   recap slides** so the block runs standalone the way your 2022 deck did.
3. **Video blocks are the retaping script** – the PG adoptions and
   enrichments go *inside* the relevant video block, so the deck is what you
   would tape next time. The 2022 videos are then out of step with the deck
   until you retape; section 5 lists exactly which slides differ.
4. **BMW: graphical comparison in the main deck, full algebra in the
   practice-video deck.**
5. **Costco appears twice**: "Versioning: Costco" inside Video 4 as taped,
   and a second slide in the in-class block revisiting it as the versioning +
   two-part-tariff combination (your 2024 title).
6. **All sections stay in the agenda; no introduction section.** Video 1 is
   already a teaching video ("Simple vs. Complex Pricing"), so the outline
   has no introduction item and the front matter sits outside every video
   block. → the 9-row outline in section 1 stands.
7. **The Mad Optimist is kept** (both slides). Only the verbatim second copy
   at your slides 69–70 is not repeated; say the word and I will put it back.
8. **Zoo example: the applications-deck version** (F = $32, P = $4/visit,
   MC = $4).
9. **Exam dates refreshed from the Fall 2026 calendar** – see section 6.
10. **Auctions are dropped from this project** – PG Part 3 is not adopted
    here or carried forward (2026-09-09).
11. **Duplicate source decks deleted** from the folder root; the
    `Module 6 - NV Slides/` copies are the build inputs. All ten were
    byte-identical (sha1 verified before deleting).

---

## 1. Module outline (9 rows) and the video mapping

Your four-degree hierarchy, with the three advanced-pricing strategies as
sub-items — one row per video, so every section agenda highlights one row.

| # | Outline item | Description | Pill |
|---|---|---|---|
| 1 | Simple vs. Complex Pricing | The pricing dilemma and why complex pricing pays | Video 1 |
| 2 | First Degree: Perfect Price Discrimination | Charging every customer her own willingness to pay | Video 2 |
| 3 | Third Degree: Segment Pricing | Different prices for groups with different price sensitivity | Video 3 |
| 4 | Second Degree: Versioning and Coupons | Letting customers sort themselves into versions | Video 4 |
| 5 | Advanced Pricing Strategies | Extracting surplus from one customer who buys many units | Videos 5–7 |
| 5a | *Flat Fee Pricing* | Unlimited access for one all-or-nothing fee | Video 5 |
| 5b | *Two-Part Tariff* | A flat fee plus a usage fee set at marginal cost | Video 6 |
| 5c | *Block Pricing* | A lower price as the same customer buys more | Video 7 |
| 6 | Summary of Pricing Strategies | Which strategy fits which market conditions | Video 8 |

Video names are verbatim from the Fall 2026 calendar, so pills, title cards
and calendar cannot drift. Nine rows need a **0.615" row pitch** — inside the
agenda builder's 0.60" floor, so it fits, with title rows slightly tighter
than in Modules 3 and 4.

**Tags.** Front matter keeps two-level tags (`Module 6 · Logistics`,
`Module 6 · Course Roadmap`). Content inside video *k* reads
`Module 6 · Video k · <topic>`; agenda slides `Module 6 · Video k · Agenda`;
the summary closer `Module 6 · Summary`; the in-class block
`Module 6 · In Class · Examples · <topic>`, with its two section agendas as
`Module 6 · In Class · Agenda`; backup slides `Module 6 · Backup`.

---

## 2. Video block boundaries — established from the tapes

Every slide of the eight video decks was matched to a main-deck slide by
text-overlap scoring (`_map_videos.py`). All 44 content slides matched at
**≥ 0.71**, most at 1.00; the only low scores are the title cards, which
match the deck title slide as expected.

| Video | Main-deck slides, in the tape's own order | Tape length |
|---|---|---|
| V1 · Simple vs. Complex Pricing | 4, 5 (**both** outline slides), 6, 7, 8, 9, 10 | 13 min |
| V2 · First-Degree Price Discrimination | 11, 12, 14, 15, **13** (dystopian comes last) | 11 min |
| V3 · Segment Pricing | 17, 18, 19, 20, 21, 22, 23, **29**, 30 | 14 min |
| V4 · Versioning and Coupons | 31, 32, 33, 34, **59 (Costco)**, 37, 42 | 13 min |
| V5 · Flat Fee Pricing | 43, 44, 45, 46, 47, 49 | 16 min |
| V6 · Two-Part Tariffs | 50, 51, 52, 53 | 8 min |
| V7 · Block Pricing | 60, 61, 62 | 8 min |
| V8 · Summary of Pricing Strategies | 64 — **no agenda slide** | 6 min |

**Main-deck slides in NO video** — these are the in-class material: 2, 3
(announcements, roadmap), 16 (Uber exercise), 24–25 (airline puzzle +
Lufthansa solution), 26–28 (Lodine poll trio), 35 (AMC), 36 (Barbie), 38
(Wendy's), 39–41 (airline poll trio), 48 (WSJ Netflix), 54–58 (the whole zoo
worked example), 63 (ice cream), 65–73 (innovative + behavioral), 74 (module
summary), 75–76 (backup).

Three things that fall out of this and confirm the design:

- **Slide 52 ("Two-part tariff Summary") is title-only in the tape too.**
  It is a blank canvas you talk over, not an unfinished slide.
- **The zoo solution (56) is title-only because it is worked live in class.**
  The applications deck carries the completed version.
- **The dystopian slide exists in two versions by design.** Video 2 *states*
  it ("By tracking consumer behavior, tech companies are increasingly able
  to use first-degree price discrimination… Example: Uber"); the main deck
  and applications deck *ask* it ("Why are companies increasingly able to…?
  Any examples?"). The video block gets the statement version, the in-class
  block the question version.

---

## 3. Slide-by-slide plan — `Module 6 - Revised.pptx` (123 slides)

Legend: **[NV n]** = main deck slide n · **[NVapp n]** = applications deck
slide n · **[Vk sn]** = video k slide n · **[NEW – PG1/2 n]** = adopted from
PG · **[CARD]** = video title card · **[POLL]** = PollEverywhere slide,
spliced verbatim with notes and tags · **[NATIVE]** = rebuilt from
PowerPoint shapes.

### Front matter — outside every video block

| # | Title | Source | Treatment |
|---|---|---|---|
| 1 | *Complex Pricing* · *Advanced Pricing Strategies* | [NV 1] | Deck-standard title slide; the UCLA Anderson wordmark is dropped per the branding rule |
| 2 | Some Logistics | [NV 2] | Exam facts refreshed from the Fall 2026 calendar — see section 6 |
| 3 | Agenda for the Class | [NV 3] | Four-box course roadmap, "we are here" on box 4 |

### 1 · Simple vs. Complex Pricing — Video 1

| # | Title | Source | Treatment |
|---|---|---|---|
| 4 | *Simple vs. Complex Pricing* | [CARD] | `Module 6 · Video 1` |
| 5 | Outline of Module 6 | [NV 4] · [V1 s2] | Descriptive overview — all nine descriptions, every pill in colour |
| 6 | Outline of Module 6 | [NV 5] · [V1 s3] | Section agenda, item 1 banded. Your two consecutive outline slides are exactly this pair, and the tape has both |
| 7 | The Dilemma of Simple Pricing | [NV 6] · [V1 s4] | Bullets, wording kept |
| 8 | What if Netflix Used Simple Pricing per Movie? | [NV 7] · [V1 s5] | [NATIVE] Demand (dark red) + MR (navy) + max-TR rectangle; `P = 2 − Q/8` as OMML |
| 9 | Under Simple Pricing, Netflix Would "Lose" Revenues | [NV 8] · [V1 s6] | [NATIVE] **[NEW – PG1 9]** adopts PG's in-plot labelling: Consumer Surplus / Revenues / Unexploited Markets named inside the plot with leader lines instead of three floating labels |
| 10 | Complex (Non-Uniform) Pricing | [NV 9] · [V1 s7] | **[NEW – PG1 10]** adds two sub-bullets making the market-power condition explicit: perfect competition → price taker, no pricing strategy; market power → price searcher has pricing power |
| 11 | Preventing Resale: The Dum-Dums Case | **[NEW – PG1 20]** | Bloomberg, "Lollipops Hustle on Amazon Costs Family Candy Business Millions" — resellers buying Dum-Dums in bulk at Sam's Club and drop-shipping them on Amazon. The concrete case for the no-resale condition on the slide before |
| 12 | Three Degrees of Price Discrimination | [NV 10] · [V1 s8] | The three degrees as cream concept callouts |

### 2 · First Degree: Perfect Price Discrimination — Video 2

| # | Title | Source | Treatment |
|---|---|---|---|
| 13 | *First-Degree Price Discrimination* | [CARD] | `Module 6 · Video 2` |
| 14 | Outline of Module 6 | [NV 11] · [V2 s2] | Section agenda, item 2 banded |
| 15 | First Degree Price Discrimination | [NV 12] · [V2 s3] | Conditions; "Note: this is (still) a hypothetical scenario" kept |
| 16 | First Degree Price Discrimination (Assume MC = 0) | [NV 14] · [V2 s4] | [NATIVE] Demand = MPV, shaded revenue triangle, your three MPV callouts |
| 17 | First Degree Price Discrimination (Now MC > 0) | [NV 15] · [V2 s5] | [NATIVE] Same figure plus MC; "do not sell to low-MPV customer" callout |
| 18 | The Dystopian Future of Price Discrimination | [V2 s6] | The tape's statement version, "Example: Uber". **[NEW – PG1 5]** adds the WSJ headline "Welcome to the Grocery Store Where Prices Change 100 Times a Day" (electronic shelf labels) as a second clipping |
| 19 | Personalized Pricing in the Wild | **[NEW – PG1 15]** | The Japanese vending machine with a face-recognition camera reading the customer before quoting a price — the best image either deck has for first-degree pricing actually happening |

### 3 · Third Degree: Segment Pricing — Video 3

| # | Title | Source | Treatment |
|---|---|---|---|
| 20 | *Segment Pricing* | [CARD] | `Module 6 · Video 3` |
| 21 | Outline of Module 6 | [NV 17] · [V3 s2] | Section agenda, item 3 banded |
| 22 | Third Degree Price Discrimination: Segment Pricing | [NV 18] · [V3 s3] | Conditions kept |
| 23 | Segment Pricing | [NV 19] · [V3 s4] | [NATIVE] Demand with the two price bands (regular movie-goers / seniors); P₁, P₂ as real subscripts |
| 24 | Segment Pricing: Student Discounts | [NV 20] · [V3 s5] | Photo, rounded + shadowed, caption grouped |
| 25 | Segment Pricing: Discounts for Locals | [NV 21] · [V3 s6] | Disneyland + Taj Mahal |
| 26 | Segment Pricing: Timing | [NV 22] · [V3 s7] | camelcamelcamel history; source line kept |
| 27 | Segment Pricing: Ways to Segment Consumers | [NV 23] · [V3 s8] | Bullets kept, incl. the doctors-and-veterinarians line |
| 28 | Impediments to Segment Pricing: Orbitz | **[NEW – PG1 41]** | The WSJ Mac-vs-PC Orbitz room-rate comparison, same dates and same search — a direct illustration of your own "pricing based on search history" bullet |
| 29 | Segment Pricing | [V3 s9] | The tape's lighter version: two markets, "should the price be the same? Not necessarily", set MRᵢ = MCᵢ, rule = lower price in more elastic markets, plus the Practice Video link box. **Not** the main deck's heavier slide 29 |
| 30 | Segment vs. Uniform Pricing: The Graphical Answer | **[NEW – PG1 40]** | [NATIVE] Three panels — US market, German market, joint market — showing segment profits (B + C) > uniform profits (A). Makes the practice video's *result* visible without repeating the algebra |
| 31 | Challenges with Segment Pricing for BMW | [NV 30] · [V3 s10] | Arbitrage / backlash bullets kept. **[NEW – PG1 43]** adds the legal impediment (Robinson-Patman Act, 1936) and the closing line that versioning is how firms get round these problems — which hands off to Video 4 |

### 4 · Second Degree: Versioning and Coupons — Video 4

| # | Title | Source | Treatment |
|---|---|---|---|
| 32 | *Versioning and Coupons* | [CARD] | `Module 6 · Video 4` |
| 33 | Outline of Module 6 | [NV 31] · [V4 s2] | Section agenda, item 4 banded |
| 34 | Second Degree Price Discrimination: Versioning | [NV 32] · [V4 s3] | Conditions kept |
| 35 | Versioning | [NV 33] · [V4 s4] | [NATIVE] The five-version staircase under the demand curve |
| 36 | Versioning | [NV 34] · [V4 s5] | "Offer a product line and let users sort themselves"; three product photos |
| 37 | Versioning: Costco | [NV 59] · [V4 s6] | **[NEW – PG2 7]** — the picture placeholder on this slide is **empty** in both your decks. PG's Costco membership-tier image ($65/year Gold Star, $130/year Executive) fills it exactly; your business-model bullets and the WSJ-video link are kept |
| 38 | More Examples for Versioning | [NV 37] · [V4 s7] | Bullets kept, incl. the incentive-compatibility note |
| 39 | Second-Degree Price Discrimination: Coupons | [NV 42] · [V4 s8] | Coupon images kept; "coupons ensure incentive-compatibility" as the gold takeaway bar |

### 5a · Flat Fee Pricing — Video 5

| # | Title | Source | Treatment |
|---|---|---|---|
| 40 | *Flat Fee Pricing* | [CARD] | `Module 6 · Video 5` |
| 41 | Outline of Module 6 | [NV 43] · [V5 s2] | Section agenda, items 5 + 5a banded |
| 42 | Context | [NV 44] · [V5 s3] | Bullets kept |
| 43 | Advanced Pricing | [NV 45] · [V5 s4] | The three variants named |
| 44 | Example: Netflix Flat Fee (Assume MC = 0) | [NV 46] · [V5 s5] | [NATIVE] Whole triangle shaded; area = (16 × 2)/2 = 16 as OMML |
| 45 | Netflix Pricing Strategy | [NV 47] · [V5 s6] | Plan tiers + price-history chart; the mix-of-flat-fee-and-versioning line as the takeaway bar |
| 46 | ClassPass: Credits as Volume Pricing | **[NEW – PG2 16]** | Plan tiers (6 credits $19 … 150 credits $299) plus the credits-per-class mechanism — popular classes and peak times cost more credits. Simultaneously volume pricing and segment pricing by time of day |
| 47 | What Can Go Wrong with Flat Rates… | [NV 49] · [V5 s7] | Clipping kept; "only 65 customers bought the flat-rate pass" moves from your notes onto the slide |

### 5b · Two-Part Tariff — Video 6

| # | Title | Source | Treatment |
|---|---|---|---|
| 48 | *Two-Part Tariffs* | [CARD] | `Module 6 · Video 6` |
| 49 | Outline of Module 6 | [NV 50] · [V6 s2] | Section agenda, items 5 + 5b banded |
| 50 | Two-Part Tariff (Now MC > 0) | [NV 51] · [V6 s3] | [NATIVE] F\* as the consumer-surplus triangle above P\* = MC |
| 51 | Two-Part Tariff: Summary | [NV 52] · [V6 s4] | **Decision A** — title-only in both your decks *and* in the tape. **[NEW – PG2 18]** would fill it: flat fee T + usage fee P; console/games, Zipcar, Amazon Prime as one-line examples; set P\* = MC to avoid excessive use; set T\* = the whole consumer surplus = (P_max − P\*)·Q/2. Say if you would rather keep it blank to draw on |
| 52 | Two-Part Tariff: ZipCar | [NV 53] · [V6 s5] | [NATIVE] F = $50, P = $0.50/mile; the "would a flat fee alone work?" question kept |

### 5c · Block Pricing — Video 7

| # | Title | Source | Treatment |
|---|---|---|---|
| 53 | *Block Pricing* | [CARD] | `Module 6 · Video 7` |
| 54 | Outline of Module 6 | [NV 60] · [V7 s2] | Section agenda, items 5 + 5c banded |
| 55 | Block Pricing | [NV 61] · [V7 s3] | The definition as a hero concept box |
| 56 | Block Pricing: Photo Cards at Walmart | [NV 62] · [V7 s4] | [NATIVE] The A / B / C step blocks, three purchase options, MC = $0.05 |

### 6 · Summary of Pricing Strategies — Video 8

| # | Title | Source | Treatment |
|---|---|---|---|
| 57 | *Summary of Pricing Strategies* | [CARD] | `Module 6 · Video 8` |
| 58 | Outline of Module 6 | new | **Decision B** — the tape has no agenda slide here (V8 is just card + flowchart). Added so Video 8 follows the same card → agenda → content pattern as the other seven. Drop it if you would rather match the tape exactly |
| 59 | How to Extract Consumer Surplus? A Summary of Pricing Strategies | [NV 64] · [V8 s2] | [NATIVE] Decision tree as navy/gold boxes and connectors with Yes/No labels; "Source: Goolsbee, Levitt, Syverson" kept |

### Summary closer — outside the video blocks

| # | Title | Source | Treatment |
|---|---|---|---|
| 60 | Module 6: Summary | [NV 74] | The outline slide with every row lit and every description shown, then your summary bullets. `Module 6 · Summary` |

### In-Class Examples block

The applications deck's own sequence (its slides 7–52), recaps included, with
the PG case additions and the second Costco slide inserted at the points they
belong. Tagged `Module 6 · In Class · Examples · <topic>`.

| # | Title | Source | Treatment |
|---|---|---|---|
| 61 | *Slides for the On-Campus Applications Class* | new | Divider, no top bar |
| 62 | Outline of Module 6 | [NVapp 7] | The "In the Videos / In class today" split kept as it is |
| 63 | Dilemma of Simple Pricing, Explained by the Podcast | [NVapp 8] | The NYC burger: MC $4.50, price $18, whom you lose and whom you subsidise |
| 64 | Dilemma of Simple Pricing | [NVapp 9] | [NATIVE] The recap graph |
| 65 | In the News: Hospital Mergers | **[NEW – PG1 3]** | WSJ "The True Cost of Megamergers in Healthcare: Higher Prices" + the question on surgery, ICU and ER prices after mergers |
| 66 | In the News: Restaurant Dynamic Pricing | **[NEW – PG1 4]** | WSJ "Surge Pricing Is Coming to More Menus Near You" + the questions on profitability, why some chains refuse, and the generational split |
| 67 | Complex (Non-Uniform) Pricing | [NVapp 10] | Recap |
| 68 | First Degree Price Discrimination | [NVapp 11] | Recap |
| 69 | The Dystopian Future of Price Discrimination | [NVapp 12] · [NV 13] | The **question** version, with both your clippings |
| 70 | Uber Starts Charging What You're Willing to Pay? | [NVapp 13] · [NV 16] | The live volunteer exercise — confirmed absent from Video 2, so it lives only here |
| 71 | Third Degree Price Discrimination: Segment Pricing | [NVapp 14] | Recap |
| 72 | Consumer Segments in Medication | [NVapp 15] · [NV 26] | Poll set-up; Poll Break parallelogram |
| 73 | Lodine poll | [NVapp 16] · [NV 27] | [POLL] Round POLL pill |
| 74 | Solution | [NVapp 17] · [NV 28] | "Higher price for humans!"; the 3× figure from your notes goes on the slide. Poll Break parallelogram |
| 75 | Important: Segment Pricing Analytically | [NVapp 18] · [NV 29] | The heavier analytics slide, with the practice-video link box |
| 76 | Airline Pricing: We Can Now Solve the "Puzzle" from Week 1 | [NV 24] | Lufthansa fare screenshots; callouts rebuilt as rounded cards |
| 77 | Lufthansa Pricing: Solution | [NV 25] | [NATIVE] Your two-panel D / MR / MC diagram, with Q\* at the true MR = MC intersection and P\* on the demand curve |
| 78 | Segment Pricing in Movies | [NVapp 19] | Higher prices for new releases |
| 79 | Outline of Module 6 | [NVapp 20] | In-class section agenda, `Module 6 · In Class · Agenda` |
| 80 | Second Degree Price Discrimination: Versioning | [NVapp 21] | Recap |
| 81 | Recall: Versioning | [NVapp 22] | [NATIVE] The recap staircase |
| 82 | Versioning: AMC | [NV 35] | Your two `.emf` price tables rebuilt as **native PowerPoint tables** |
| 83 | Versioning: Dolls | [NVapp 23] | **Decision C** — your three price-label boxes are empty on this slide (and on the main deck's Barbie version). What should they read? |
| 84 | Wendy's Failed Attempt at Dynamic Pricing | [NV 38] | Clipping kept; the Warren quote from your notes as a cream quote callout |
| 85 | Airline Response to Low-Cost Carriers | [NVapp 24] · [NV 39] | Poll set-up; Poll Break parallelogram |
| 86 | Airline-response poll | [NVapp 25] · [NV 40] | [POLL] Round POLL pill |
| 87 | Versioning in Response to Low-Price Competition | [NVapp 26] · [NV 41] | Poll Break parallelogram; the Kodak Gold / FunFilm story stays in the notes. Backup link pill → slide 123 |
| 88 | Podcast on Price Discrimination | [NVapp 27] | Haggling, incentive compatibility, and the $9 coupon on a $4.50 burger |
| 89 | Outline of Module 6 | [NVapp 28] | In-class section agenda |
| 90 | Context | [NVapp 29] | Recap |
| 91 | Advanced Pricing | [NVapp 30] | Recap |
| 92 | Example: Netflix Flat Fee (Assume MC = 0) | [NVapp 31] | [NATIVE] Recap |
| 93 | Netflix Pricing Strategy | [NVapp 32] | Recap |
| 94 | WSJ Article on Netflix | [NVapp 33] · [NV 48] | `.emf` clippings rebuilt as a native quote callout; "what's wrong with this statement?" kept |
| 95 | What Can Go Wrong with Flat Rates… | [NVapp 34] | **Hidden in your deck — rebuilt and kept hidden** |
| 96 | What Can Go Wrong with Flat Rates… | [NVapp 35] | The visible version |
| 97 | Two-Part Tariff (Now MC > 0) | [NVapp 36] | [NATIVE] Recap |
| 98 | Pricing at a Local Zoo | [NVapp 37] · [NV 54] | Poll set-up; Poll Break parallelogram |
| 99 | Zoo poll | [NVapp 38] · [NV 55] | [POLL] Round POLL pill |
| 100 | Zoo Pricing – Solution | [NVapp 39] | [NATIVE] Two-part tariff: F = $32, P = $4/visit, MC = $4, profit $32. Your main deck's slide 56 is title-only because this is worked live |
| 101 | Zoo Pricing – Comparison to Simple Pricing | [NVapp 40] · [NV 58] | [NATIVE] MR line, P = $10, Q = 2.5, profit $25; problem-set pointer as the standard ✎ box |
| 102 | Zoo Pricing – Comparison to Flat Fee | [NVapp 41] | [NATIVE] F = $50, variable cost $20, profit $30 |
| 103 | Costco: Combination of Versioning and Two-Part Pricing | [NV 59] | The **second** Costco slide you asked for — your 2024 title and framing, with the membership-tier image |
| 104 | Block Pricing | [NVapp 42] | Recap |
| 105 | Block Pricing: Ice Cream | [NVapp 43] · [NV 63] | [NATIVE] 1 scoop $4 / 2 for $7 / 3 for $8.50, the A / B / C blocks and the "price of the 2nd scoop is $7 − $4 = $3" callout |
| 106 | How to Extract Consumer Surplus? A Summary | [NVapp 44] | Recap of the decision tree |
| 107 | Disneyland: Explain the Pricing Strategy | **[NEW – PG2 25]** | Park banner + the open question, gold Discussion Break badge. Opens the integrative case |
| 108 | Disneyland: Third-Degree Price Discrimination | **[NEW – PG2 27]** | The ticket selector: Ages 10+ vs. Ages 3–9 |
| 109 | Disneyland: Third-Degree Price Discrimination | **[NEW – PG2 28]** | The pricing calendar, $104 to $224 by date — intertemporal segmentation |
| 110 | Disneyland: Second-Degree Price Discrimination | **[NEW – PG2 26 / 29]** | Lightning Lane Multi / Single / Premier — versioning inside the park |
| 111 | Disneyland: Two-Part Tariff | **[NEW – PG2 29]** | Park admission as the flat fee, Lightning Lane per attraction as the usage fee |
| 112 | Disneyland: Volume Pricing | **[NEW – PG2 30]** | The four Magic Key annual passes |
| 113 | Disneyland: What Strategy Is Where | new | One card mapping each Disneyland device onto the degree / strategy it illustrates, so the case closes on the module's own framework. Gold takeaway bar |
| 114 | *Some Examples for Innovative Pricing Strategies* | [NVapp 45] · [NV 65] | Divider; your "(only if we have time)" subtitle kept |
| 115 | The Mad Optimist | [NV 66] · [NVapp 46] | Kept |
| 116 | The Mad Optimist – "Choose Your Price" | [NV 67] | Kept |
| 117 | Rent the Runway | [NVapp 47] · [NV 68] | Block pricing + versioning |
| 118 | The Economist Pricing Example | [NVapp 48] · [NV 71] | **Decision D** — the picture placeholder is **empty**; the 16% / 84% labels are there but the subscription-options image is missing. Rebuild the three options as a native table? |
| 119 | Dan Ariely's Experiment | [NVapp 49] · [NV 72] | Bar chart rebuilt native; the two-option vs. three-option result (32% vs. 84%) from your notes onto the slide |
| 120 | Behavioral Insight: Price Framing | [NV 73] | Three panels, each with its own caption revealed on its own click |

### Backup

| # | Title | Source | Treatment |
|---|---|---|---|
| 121 | *Backup* | [NV 75] · [NVapp 51] | Divider, no top bar |
| 122 | More Seating Versions | [NV 76] · [NVapp 52] | The *Airplane Seating Chart* cartoon (TheCooperReview) — First Class down through "Economy 2: The Reckoning" and "Satan's Den Economy". Backs the versioning slide, which already says "More seating versions…", so slide 87 gets a backup link pill and this slide the navy "← Back" button. Full-bleed figure, so no top bar and no tag, per the backup exception |

**Net: 76 + 52 in → 122 out.** Growth is 8 video title cards, 1 new V8
agenda slide, 1 Disneyland summary card, 13 adopted PG slides, and the
applications deck's own applications and recaps. Two verbatim duplicates
(your 69–70) are not repeated, and the applications deck's front matter
(its slides 1–6: title, three exam slides, "Plan for Today", course agenda)
is left out because it is class-session material spanning Modules 6–8 —
**tell me if you want "Final Exam: Structure" ([NVapp 3]) as a second
logistics slide.**

---

## 4. Second deliverable — `Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx`

The full PG BMW block as its own deck, for retaping the practice video. Same
chrome, tagged `Module 6 · Practice Video`. It repeats the two illustrative
slides that also appear in the main deck so the video stands alone.

| # | Title | Source | Treatment |
|---|---|---|---|
| 1 | *Optimal Pricing in Two Markets* | [CARD] | `Module 6 · Practice Video` |
| 2 | Segment Pricing Analytically | [NV 29] + **[NEW – PG1 31]** | Set MRᵢ = MCᵢ in each market; with equal MC, MR_a = MR_b = MC; if MR_a ≠ MR_b uniform pricing is not optimal. Gold rule bar: charge the higher price in the less elastic market |
| 3 | BMW and Segmentation | **[NEW – PG1 32]** | Set-up: Q in thousands of cars; US inverse demand `P = 160 − 4Q`; German `P = 140 − 5Q`; `MC = 40` in both (in $1,000). Asks for the profit-maximizing price in each. Group-exercise badge |
| 4 | Profit Maximization in the U.S. Market | **[NEW – PG1 33]** | OMML: TR = 160Q − 4Q², MR = 160 − 8Q, MR = MC ⇒ Q\* = 15, **P\* = 100** (answer line dark red) |
| 5 | Profit Maximization in Germany | **[NEW – PG1 34]** | MR = 140 − 10Q, MR = MC ⇒ Q\* = 10, **P\* = 90** (dark red) |
| 6 | BMW and Segmentation | **[NEW – PG1 35]** | Recap of both answers, then: what is the demand elasticity at the profit-maximizing price, and the Lerner Index? |
| 7 | Demand Elasticity and Mark-up: USA | **[NEW – PG1 36]** | Q = 40 − 0.25P, dQ/dP = −0.25, E_D = −0.25 · (100/15) = **−1.67**; Lerner (P − MC)/P = 60/100 = **0.6** = 1/\|E_D\|. Ties back to Module 5 |
| 8 | Demand Elasticity and Mark-up: Germany | **[NEW – PG1 37]** | Q = 28 − 0.2P, E_D = −0.2 · (90/10) = **−1.80**; Lerner = 50/90 = **0.56** = 1/\|E_D\| |
| 9 | Segment vs. Uniform Pricing: The Graphical Answer | **[NEW – PG1 40]** | [NATIVE] The three-panel figure; same slide as main-deck 30 |
| 10 | Segment Pricing: Take-Away | **[NEW – PG1 38]** | Set MR = MC market by market; charge the higher price where demand is less elastic (here the US). Problem-set pointer ✎ |

I checked all of PG's arithmetic: US Q\* = 15, P\* = 100, E_D = −1.667,
Lerner 0.6; Germany Q\* = 10, P\* = 90, E_D = −1.8, Lerner 0.556. Every
number is right and Lerner = 1/\|E_D\| holds in both markets.

**Decision E** — the existing practice video is 19 minutes on a different
example. If you retape with BMW, the calendar entry keeps its name and only
the Panopto link changes. Confirm you want the old example retired rather
than kept alongside.

---

## 5. What now differs from the 2022 tapes

Because the video blocks are the retaping script, these are the slides a
student watching the current videos would **not** see. If you do not retape,
this list is what to warn them about — or move these slides to the in-class
block instead.

| Video | Difference |
|---|---|
| V1 | +1 slide: Dum-Dums resale case (main 11). Enriched: main 9 (in-plot labels), main 10 (price-taker sub-bullets) |
| V2 | +1 slide: personalized-pricing vending machine (main 19). Enriched: main 18 (shelf-labels headline) |
| V3 | +2 slides: Orbitz (main 28), three-panel segment-vs-uniform figure (main 30). Enriched: main 31 (Robinson-Patman) |
| V4 | No new slides. Enriched: main 37 (Costco image fills the empty placeholder) |
| V5 | +1 slide: ClassPass (main 46) |
| V6 | Enriched: main 51 (two-part-tariff summary filled) — Decision A |
| V7 | No change |
| V8 | +1 slide: section agenda (main 58) — Decision B |

---

## 6. Slide 2 — exam facts refreshed from the Fall 2026 calendar

Your slide reads: practice final online, similar to the actual final, **5
hours**, any time between **March 19 and March 20**, two windows with
guaranteed TA availability on Zoom, and "many practice problems on
**Achieve**". Every one of those is out of date. From
`405 Calendar and Website/Course Calendar/_calendar_content.py` (anchor
Friday 2026-09-25):

- **Final exam window: Friday, December 11 – Sunday, December 13, 2026.**
- **One 3.5-hour window, the same for everyone**; the exact date and time to
  be announced in class. (Not 5 hours, and not two TA windows.)
- Online, open book, open notes; calculator allowed; proctoring software.
- Covers all material, Modules 1 – 8.
- About 20 multiple-choice questions and 3 – 4 problem-solving questions.
- Practice Final Exam comes in the exam-prep period, solutions on BruinLearn.
- **"Achieve" is gone from the course** — the practice material is now the
  online quizzes on BruinLearn (the calendar has "Online quiz on Module 6").

Proposed wording, matching what you approved for Module 4's logistics slide:

> **Some Logistics**
> - Practice Final Exam on BruinLearn — similar to the actual final
> - Final exam period: December 11 – 13, 2026
>   - One 3.5-hour window — the same window for everyone
>   - The exact date and time will be announced in class
> - Online, open book, open notes; covers all material, Modules 1 – 8
> - About 20 multiple-choice questions and 3 – 4 problems

**Decision F** — the problem-set number. Your zoo slide points to "Problem
Set 5, Problem 2" and PG's segment-pricing slide points to "Problem Set 4,
#4". For Fall 2026 the calendar has PS3 due Tue Nov 10, PS4 due Tue Nov 24
and PS5 due Thu Dec 3, and Module 6's applications class is Sat Nov 21 —
so PS4 or PS5 are both plausible. Per the pointer convention the box names
the number only, never the exercise, so I need the number confirmed.

---

## 7. PG adoptions — the approval list

Approve or reject line by line.

| # | Adoption | From | Goes to | Status |
|---|---|---|---|---|
| A1 | In-plot labelling of CS / Revenues / Unexploited Markets | PG1 9 | Main 9 | adopted |
| A2 | Price-taker vs. price-searcher sub-bullets | PG1 10 | Main 10 | adopted |
| A3 | Dum-Dums resale case (Bloomberg) | PG1 20 | Main 11 | **approved** |
| A4 | Electronic shelf labels headline (WSJ) | PG1 5 | Main 18 | **approved** |
| A5 | Face-recognition vending machine | PG1 15 | — | **WITHDRAWN** — byte-identical to my own slide 13's images (sha1 6471d638 / 4f2ec7c5); nothing to adopt |
| A6 | Orbitz Mac-vs-PC room rates (WSJ) | PG1 41 | Main 28 | **approved** |
| A7 | Three-panel segment-vs-uniform figure | PG1 40 | Main 30 + Practice 9 | **approved** |
| A8 | Robinson-Patman Act + hand-off to versioning | PG1 43 | Main 31 | adopted |
| A9 | Two-part-tariff summary content | PG2 18 | — | **WITHDRAWN** — slide 52 was never blank; my own six-paragraph summary was hidden inside mc:AlternateContent |
| A10 | Costco membership-tier image | PG2 7 | Main 37 + In-class 103 | **approved** |
| A11 | ClassPass credit pricing | PG2 16 | Main 46 | **approved** |
| A12 | Disneyland integrative case (6 slides + 1 summary card) | PG2 25–30 | In-class 107–113 | **approved** |
| A13 | Hospital mergers headline (WSJ) | PG1 3 | In-class 65 | **approved** |
| A14 | Restaurant dynamic pricing headline (WSJ) | PG1 4 | In-class 66 | **approved** |
| A15 | Full BMW algebra + elasticity + Lerner block | PG1 31–38 | Practice deck | **approved** |

**Not adopted:** PG Part 3 in full (auctions → Module 8); PG's flat outline
with auctions as item 5 (your hierarchy is kept); PG's Netflix
simple-pricing slide (PG2 13, a duplicate); PG's own polls (they run on her
PollEverywhere account and are never spliced).

---

## 8. Global treatments and the open decisions

**Global.** 4:3 → 13.33 × 7.5"; multiple masters → one; Calibri throughout
(Cambria Math in equations); navy `0B2B4E` / gold `E09F3E` / gray `555B66`;
white backgrounds; navy top bar with the video-numbered tag; Title Case
action titles; thin gray rule + gold accent strip; live slide-number footer
fields. **All hand-built diagrams rebuilt native** (main 8, 9, 16, 17, 23,
30, 35, 44, 50, 52, 56, 59; in-class 64, 77, 81, 92, 97, 100, 101, 102, 105,
106, 119) with navy arrow-tipped axes, axis titles anchored to the arrow
tips in label-width boxes, Bézier freeforms, and every marked point at the
true intersection. **Demand curves dark red `C00000`**, supply / MC navy,
swept deck-wide. Formulas as OMML with real subscripts. Poll trios get the
alternating Poll Break → POLL → Poll Break marks, drawn last, never
animated. Fade builds, one beat per click; pictures grouped with captions,
boxes with their text. No trailing periods; no institutional branding mark.

**Decisions settled (2026-09-09, Nico: "let's not deal with section 8 for
now").** No longer open; each is carried in the build script with a dated
comment so it is easy to revisit.

- **A · Two-part-tariff summary (main 51)** — **filled** with PG2 18's
  content. The slide is titled "Summary" and is blank in every copy; a
  retaping script should say what the summary is.
- **B · Video 8 agenda slide (main 58)** — **added**, so all eight blocks
  follow card → agenda → content.
- **C · Barbie / dolls price labels (in-class 83)** — **resolved from the
  source images, nothing invented.** Your slide already carries three Mattel
  screenshots reading `$9.99` (Barbie Farmer Doll), `$150.00` (Barbie Yves
  Saint Laurent Doll) and `$13.99` (Barbie Robotics Engineer Doll). Those
  images go under their dolls and the three empty text boxes are dropped —
  they would only repeat what the screenshots show.
- **D · The Economist options (in-class 118)** — **rebuilt as a native
  table** with the three options the slide's own labels imply: web-only $59,
  print-only $125, print + web $125, against the 16% / 84% split.
- **E · Old practice-video example** — no action needed now. The new BMW deck
  is built either way; the existing video stays up until you retape.
- **F · Problem-set numbers** — **each slide keeps the number its own source
  carries**: "Problem Set 5" on the zoo comparison (your deck) and "Problem
  Set 4" on the BMW take-away (PG's deck). Numbers only, never exercise
  numbers, per the pointer convention. Both are one constant in the build
  script, so a Fall 2026 renumbering is a one-line change.
- **G · Repeated titles in the in-class block** — **your exact titles are
  kept**, no "Recap:" prefix. Changing a slide title needs your say-so, and
  the duplicates are the same content anyway.
- **H · Deck split** — **one deck**, the Module 3 precedent. Splitting the
  in-class block into its own file later is mechanical.

**Auctions.** PG's Part 3 is dropped from this project entirely
(2026-09-09, Nico). Its extracted inventory and images stay on disk unused;
they regenerate from the untouched source deck in one command if ever wanted.
