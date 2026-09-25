# -*- coding: utf-8 -*-
"""Speaker notes written for the Module 6 rebuild (2026-09-09).

Nico's own notes are carried verbatim in `_m6_notes.py`; these fill the
slides that had none.  The shipped standard is high coverage -- Module 3
ships at 91% and Module 4 at 73% -- because the deck is uploaded and
students read the notes as guidance.  Module 6 was at 36%.

Style, per Teaching CLAUDE.md: 2-4 sentences in a natural spoken voice,
carrying the example or the number the slide turns on, and the hand-off to
the next slide.  No emphasis claims ("the key point", "crucial") -- those
are Nico's to make.

KEYED BY DISPLAY NUMBER WITH A TITLE GUARD.  Inserting two slides into the
in-class block on 2026-09-09 shifted every later display number and would
have silently mis-attached these; the guard makes that fail loudly
instead.  Each entry is (title_prefix, note).
"""

WRITTEN_NOTES = {
    # 2026-09-15: the note still described a 3.5-hour window somewhere in
    # December 11-13 and put the practice final on BruinLearn, while the
    # slide had already been rebuilt from the course calendar.  The date
    # and slot are now filled from the SAME constants the slide uses -
    # see the {final_date} / {final_slot} substitution in
    # _apply_written_notes - so this pair cannot drift again.
    3: ("Some Logistics",
       "Two housekeeping points before we start. The practice final is "
       "on the class website and it is built to look like the real "
       "thing, so use it as a rehearsal rather than a reading exercise. "
       "The final itself is {final_date}, {final_slot}, and it is "
       "online. That window has to cover uploading your answers as well "
       "as solving the paper, so leave yourself a few minutes at the end "
       "rather than working right up to the deadline."),
    4: ("Agenda for the Class",
        "This is where we are in the course. We have done the economic way "
        "of thinking, then value and demand, then supply and cost. Module 6 "
        "sits in the fourth block, markets and pricing strategy, and it is "
        "where the demand curve you learned to estimate starts earning its "
        "keep: everything today is about turning a demand curve into a "
        "pricing decision."),
    5: ("Outline of Module 6",
        "Here is the whole module. We start with why one price for everyone "
        "leaves money on the table, then work through the three degrees of "
        "price discrimination, and finish with the strategies you use when "
        "a single customer buys many units. The gold tags on the right tell "
        "you which video each topic is in."),
    7: ("Outline of Module 6",
        "We begin with the pricing dilemma itself: what exactly goes wrong "
        "when a firm with market power charges everybody the same price."),
    8: ("Dilemma of Simple Pricing",
        "This is the trap. Charge a high price and you capture a lot from "
        "each buyer, but you lose everyone whose willingness to pay sits "
        "below it. Charge a low price and you sell to almost everyone, but "
        "you hand a large surplus to the customers who would happily have "
        "paid far more. One price cannot do both jobs, and that is the "
        "problem the rest of the module solves."),
    9: ("Dilemma of Simple Pricing, Explained by the Podcast",
         "The podcast puts the dilemma in a restaurant. A burger costs the "
         "kitchen four dollars fifty and sells for eighteen. At eighteen "
         "you lose everyone who would have paid fifteen, or eleven, and "
         "each of those is a sale worth making. Drop to ten and you win "
         "them, but you have just handed eight dollars to everyone who "
         "would have paid the eighteen."),
    10: ("What if Netflix Used Simple Pricing",
        "Let us put numbers on it. Demand is P equals 4 minus Q over 3, and "
        "marginal cost is zero, so profit maximization means setting "
        "marginal revenue to zero. That happens at 6 movies and a price of "
        "two dollars, which brings in twelve dollars a month. Hold on to "
        "that twelve dollars, because we will beat it several times before "
        "the module is over."),
    11: ("One Price Leaves Two Kinds",
        "Now look at what that twelve dollars leaves behind. The red triangle "
        "above the price is consumer surplus, value the customer keeps. The "
        "grey triangle to the right is the unexploited market, viewers who "
        "would have watched at some positive price and were shut out. "
        "Neither of those is a cost of doing business; both are money the "
        "one-price strategy simply fails to reach."),
    16: ("Three Degrees of Price Discrimination",
         "The three degrees are really three answers to one question: how "
         "much does the firm know about who is standing in front of it? "
         "First degree, it knows each customer's willingness to pay. Third "
         "degree, it can sort people into groups. Second degree, it knows "
         "nothing and has to design choices that make customers sort "
         "themselves. Everything in this module is one of those three."),
    18: ("Outline of Module 6",
         "First the dream case: the firm knows exactly what each customer "
         "will pay."),
    19: ("First Degree Price Discrimination",
         "First-degree, or perfect, price discrimination means a different "
         "price for every buyer. It needs three things: customers who "
         "differ, a firm that knows how they differ, and no resale. Note "
         "the last line: this is still largely hypothetical, and it is "
         "worth keeping that in mind while we work out how powerful it "
         "would be."),
    22: ("The Dystopian Future of Price Discrimination",
         "Two questions worth real discussion time. Why are firms getting "
         "better at this, and is the trend good or bad for consumers? The "
         "second one usually splits the room, which is the point: perfect "
         "price discrimination serves more customers than one price does, "
         "and it also leaves them with nothing."),
    23: ("Is Uber Charging What",
         "Run this live. Ten volunteers, five on Uber and five on Lyft, "
         "same destination, same moment. Have them read out the price "
         "without requesting the ride. The spread is usually enough to make "
         "the point on its own, and it lands harder coming from the room "
         "than from me."),
    25: ("Outline of Module 6",
         "Now the case firms actually use: sorting customers into groups."),
    26: ("Third Degree Price Discrimination",
         "Third-degree pricing splits customers into groups with different "
         "price sensitivities and charges each group its own price. Seniors "
         "at the cinema are the everyday example. It needs three things: "
         "groups that really do differ on average, a firm that can identify "
         "them, and no resale between them."),
    # 2026-09-15: this slide carried no note at all.
    32: ("Recall from Videos: Ways to Segment Consumers",
        "A quick recall before the examples. Four ways to split a market: "
        "by who the customer is, by whether she is new or already a "
        "customer, by where she is, and by when she buys. The gym is the "
        "clean case of the second one - the introductory rate exists for "
        "people who have not joined yet, and it disappears once they "
        "have, which is why it is never offered to the member renewing in "
        "year three. Ask the class for one example of each from their own "
        "industries before moving on."),
    36: ("Consumer Segments in Medication",
         "Lodine treats arthritis in humans and in dogs, and the two "
         "segments pay very different prices. Take the poll before you say "
         "which way it goes; the room usually splits, and the reasoning on "
         "both sides is worth hearing."),
    39: ("Segment Pricing in Movies",
         "Intertemporal segmentation. The new release costs more, and it is "
         "the same film in the same seat two months later. What the studio "
         "is pricing is impatience, which is exactly a difference in "
         "elasticity."),
    40: ("Segment Pricing",
         "The mechanics are the same rule you already know, applied twice. "
         "Set marginal revenue equal to marginal cost in each market "
         "separately. Prices come out different whenever demand or cost "
         "differs, and the direction is always the same: the lower price "
         "goes to the more price-sensitive market. The practice video works "
         "this through with BMW's US and German prices."),
    42: ("Challenges with Segment Pricing",
         "Segment pricing only survives if the segments stay apart. Between "
         "Germany and the US that is easier than it sounds for cars, "
         "because of instrumentation and taxes, but textbooks bought in "
         "India and resold in the States show how quickly arbitrage can "
         "close the gap. Add consumer anger and the Robinson-Patman Act, "
         "and you can see why firms often prefer the indirect route, which "
         "is where we go next."),
    44: ("Outline of Module 6",
         "Now the case where the firm cannot tell customers apart, so it "
         "makes them sort themselves."),
    45: ("Second Degree Price Discrimination",
         "Second-degree, or indirect, pricing is what you do when you "
         "cannot identify the groups. Instead of pricing the customer, you "
         "price versions of the product and let people choose. The firm "
         "still needs market power and customers who differ; what it no "
         "longer needs is any way of telling who is who."),
    48: ("Versioning: Costco",
         "Costco is the clean case. The membership is the version choice: "
         "sixty-five dollars for the standard card, a hundred and thirty "
         "for the executive one. The business model is to make the money on "
         "the memberships and sell the goods as cheaply as possible, which "
         "is why the shelf prices are so thin. We come back to Costco later "
         "as a two-part tariff as well."),
    49: ("Versioning: Dolls",
         "Three Barbies, and the prices are hidden. Ask the room to guess "
         "before each reveal. The farmer is nine ninety-nine and the "
         "robotics engineer thirteen ninety-nine, which sound like the same "
         "product. Then the Yves Saint Laurent doll comes in at a hundred "
         "and fifty. The plastic is nearly identical; what differs is who "
         "is buying, and Mattel never had to ask."),
    50: ("More Examples for Versioning",
         "The pattern is everywhere once you look: internet speed tiers, "
         "ground versus overnight shipping, the three cabins on an "
         "aircraft. The design rule is the one at the bottom: the cheap "
         "version has to be genuinely worse, or your high-value customers "
         "will happily take it. That is what incentive compatibility means "
         "here, and it is why airlines make basic economy irritating on "
         "purpose."),
    52: ("Airline Response to Low-Cost Carriers",
         "Put the question to the class before the poll. Southwest starts "
         "flying one of United's routes: does United match the fare, hold "
         "its price, or do something else? Most rooms split between the "
         "first two, and the answer is the third."),
    58: ("Outline of Module 6",
         "Now we change the question. Instead of many different customers, "
         "think about one customer buying many units."),
    59: ("Context",
         "Everything so far has been about customers who differ from each "
         "other. Now hold that aside: imagine every customer has the same "
         "demand curve, but each of them buys several units. The surplus is "
         "still there to be captured, but the tool is different, because "
         "what varies now is the value of the second unit against the "
         "first."),
    60: ("Advanced Pricing",
         "Three tools, and we take them in order. A flat fee for unlimited "
         "access. A two-part tariff, which is a flat fee plus a per-unit "
         "charge. And block pricing, where the price falls as the same "
         "customer buys more. Which one fits depends almost entirely on "
         "whether marginal cost is zero."),
    67: ("Outline of Module 6",
         "Now the case where serving an extra unit actually costs "
         "something."),
    68: ("Two-Part Tariff (Now MC",
         "Once marginal cost is positive, the pure flat fee breaks down, "
         "because the customer consumes until the value of the last unit is "
         "zero and you pay for all of it. The fix is two prices. Set the "
         "usage fee at marginal cost, which makes the customer stop at the "
         "right quantity, and then take the whole remaining consumer "
         "surplus as the flat fee."),
    69: ("Two-Part Tariff: Summary",
         "So the two-part tariff is one decision made twice. The usage fee "
         "is set by efficiency: put it at marginal cost and the customer "
         "self-regulates. The flat fee is set by extraction: take the "
         "triangle that is left. Video game consoles and their games, "
         "Zipcar's membership and its hourly rate, Amazon Prime and the "
         "purchases, are all this structure."),
    70: ("Two-Part Tariff: ZipCar",
         "Here it is with numbers. Demand is P equals 1.50 minus 0.01Q and "
         "marginal cost is fifty cents a mile, so the member drives a "
         "hundred miles. Charge fifty cents a mile and the flat fee can be "
         "fifty dollars, which is the triangle above marginal cost. Ask the "
         "class the question on the right before you answer it: a flat fee "
         "with no usage charge would have members driving far past the "
         "point where the miles are worth anything."),
    71: ("Pricing at a Local Zoo",
         "Here is the exercise. Demand for the average customer is P equals "
         "40 minus 8Q, where Q is visits per year, and each visit costs the "
         "zoo eight dollars. Give them a minute on the poll before working "
         "it through."),
    73: ("Zoo Pricing",
         "The answer. Set the per-visit charge at marginal cost, eight "
         "dollars, and the customer comes four times. The surplus left "
         "above that line is thirty-two times four over two, which is "
         "sixty-four dollars, and that is the membership. The per-visit "
         "charge brings in thirty-two dollars, exactly the variable cost "
         "of the four visits, so the zoo keeps the sixty-four."),
    74: ("Zoo Pricing",
         "Compare it with one simple price. Marginal revenue is 40 minus "
         "16Q, and setting that equal to the eight dollar marginal cost "
         "gives two visits at twenty-four dollars, so thirty-two dollars "
         "per customer. The two-part tariff doubled that, and it did so by "
         "serving the customer more, not less."),
    75: ("Zoo Pricing",
         "And compare it with a pure flat fee. Charge a hundred dollars for "
         "unlimited entry and the customer comes all five times, because "
         "each visit is now free to her. Those five visits cost the zoo "
         "forty dollars, so it keeps sixty. The last visit is the problem: "
         "she values it at nothing and it still costs eight dollars to "
         "serve. The usage fee in the two-part tariff is what stops that "
         "over-use, and it is worth four dollars here."),
    76: ("Costco: Combination of Versioning and Two-Part Pricing",
          "Costco is both of the things we have done today. The two "
          "membership tiers are versioning, and the membership plus the "
          "shopping is a two-part tariff, with the shelf prices close to "
          "cost. That is why the company can be relaxed about margins on "
          "the goods themselves."),
    78: ("Outline of Module 6",
         "The last of the three: charging less as the same customer buys "
         "more."),
    79: ("Block Pricing",
         "Block pricing reduces the price when the same customer buys more "
         "of the good. Notice what makes it different from the segment "
         "pricing we did earlier: the discount is tied to the quantity one "
         "person takes, not to who that person is. The firm does not need "
         "to know anything about the customer."),
    81: ("Block Pricing: Ice Cream",
          "One scoop is four dollars, two are seven, three are eight "
          "fifty. So the second scoop costs three dollars and the third one "
          "fifty. Each block sits under the same demand curve, and the "
          "customer walks herself down it, handing over most of the surplus "
          "on the way."),
    83: ("Outline of Module 6",
         "Last section. We have all the pieces now, so this one is about "
         "choosing between them."),
    # 2026-09-15: the tree was recoloured (perfect competition dark green,
    # direct price discrimination dark red), so the old closing line
    # 'everything in gold on this slide' named a colour that is no longer
    # there.  The note now walks the branches instead of the palette.
    84: ("How to Extract Consumer Surplus",
        "This tree pulls the whole module together, and it is worth "
        "walking down it slowly. No market power and you are a price "
        "taker, so there is no pricing strategy to have. Market power but "
        "no way to stop resale, and you are back to one simple price. "
        "Then the two questions that matter. Do customers differ from "
        "each other, and can you tell them apart? Yes to both is direct "
        "price discrimination - perfectly if you know each individual, by "
        "group if you only know the segment. Customers who differ but "
        "cannot be told apart send you down the indirect route, "
        "versioning and coupons. And if they do not differ from each "
        "other but each buys several units, you are in the advanced "
        "strategies: flat fee, two-part tariff, block pricing. Every "
        "branch on this slide is something we covered."),
    85: ("Module 6: Summary",
         "To recap the whole module. One price leaves surplus on the table. "
         "If you can identify customers you price them directly, perfectly "
         "in the hypothetical case and by group in practice. If you cannot, "
         "you build versions and let them sort themselves. And when one "
         "customer buys many units, the tool depends on marginal cost: a "
         "flat fee when it is zero, a two-part tariff when it is not."),
    87: ("The Mad Optimist",
          "A genuinely odd one to finish on. The company lets you set your "
          "own price. Ask the class why that is not simply giving the "
          "product away, and what it tells you about how much of "
          "willingness to pay is social rather than economic."),
    88: ("The Mad Optimist",
          "This is the choose-your-price screen itself. Note that they do "
          "anchor it: there is a suggested price, and most customers move "
          "only a little from it."),
    95: ("Disneyland: Pricing Strategy",
          "Open this one to the room before showing anything. Disneyland "
          "runs nearly every strategy in the module at once, and the class "
          "will usually find three or four of them unprompted."),
    96: ("Disneyland: Third-Degree Price Discrimination",
          "The first one is age. A child's ticket is cheaper, the park can "
          "check at the gate, and under-threes are free because they were "
          "never going to come on their own account."),
    # 2026-09-15: his slide edit replaced 'the least price-sensitive
    # visitors' with 'high-willingness-to-pay visitors'.
    97: ("Disneyland: Third-Degree Price Discrimination",
        "The second is the calendar. The same ticket runs from about a "
        "hundred and four dollars to two hundred and twenty-four "
        "depending on the date. Peak dates are when "
        "high-willingness-to-pay visitors can come, and the park never "
        "has to identify anybody - the date does the sorting."),
    # 2026-09-15: he deleted the 'nobody is asked how much their time is
    # worth' line from the slide and put the incentive-compatibility point
    # in its place; the note followed.
    98: ("Disneyland: Second-Degree Price Discrimination",
        "Lightning Lane is versioning. Three tiers, and visitors sort "
        "themselves by choosing - that is what makes it second degree "
        "rather than third. The sorting is incentive-compatible: a "
        "visitor whose time is expensive buys the Multi Pass because it "
        "is genuinely worth more to her, not because the park worked out "
        "who she is."),
    # 2026-09-15: his slide edit dropped Lightning Lane from the usage
    # fees (it is the versioning example one slide earlier) and changed
    # 'marginal cost' to 'variable cost'.
    99: ("Disneyland: Two-Part Tariff",
        "And the whole visit is a two-part tariff. Admission is the flat "
        "fee, and food, merchandise and the rest are the usage fees. The "
        "gate price takes the surplus; the in-park prices track the "
        "variable cost of serving you. Lightning Lane sits on the "
        "previous slide as versioning, so leave it out of this one."),
    # 2026-09-15: he replaced the Magic Key annual-pass image with the
    # day-ticket ladder, so the note can no longer be about Magic Keys.
    100: ("Disneyland: Volume Pricing",
        "The image is Disneyland's own day-ticket ladder: a single day "
        "runs from about $104 to $224, two days work out at about $168 a "
        "day, and five days at about $104. The more you commit to, the "
        "less each visit costs. That is block pricing under another name "
        "- the discount is tied to the quantity this visitor buys, not to "
        "who she is, so the park needs to know nothing about her."),
    # 2026-09-15: the note opened by quoting the takeaway bar he deleted
    # ('One firm, one afternoon...'), so it repeated text no longer on
    # the slide.
    101: ("Disneyland: What Strategy Is Where",
         "The table is the summary: every strategy in the module, running "
         "at one park on one afternoon. Worth leaving on screen while the "
         "class argues about which one is doing the most work. The usual "
         "answer is the two-part tariff, because the gate price is where "
         "the surplus actually gets taken."),
}


# ==========================================================================
#  Corrections to notes captured VERBATIM from the source decks
# ==========================================================================
# _m6_notes.py must stay a faithful capture, so a stale note from an old
# deck is corrected HERE and applied over the top (2026-09-15, Nico:
# "now update all notes to the slides").  Each entry is
# (guard, note): the guard is a string that must appear somewhere on the
# slide, which is checked before the note is written, so a renumbering
# fails loudly instead of attaching the note to the wrong slide.  A
# substring rather than the action title, because slide 93 is a
# full-bleed backup and has no action title.
NOTE_FIXES = {
    27: ("Divide customers into different groups",
         "We reintroduce segment pricing, a concept already introduced in "
         "week 2 when we discussed Mercedes's optimal pricing in the US and "
         "Germany, based on different elasticities. On this slide it is the "
         "cinema: seniors pay less than regular movie-goers, which is price "
         "discrimination by age. Why set different prices for different "
         "segments? Because the segments have different demand "
         "elasticities, and with marginal cost the same in both you want "
         "marginal revenue equal across them, which means the lower price "
         "goes to the more elastic segment. So why are seniors more price "
         "sensitive? Lower income, and a much lower opportunity cost of "
         "time. A retired person can take the cheap Tuesday matinee, or "
         "drive around for a better deal at the car wash, in a way a "
         "working professional cannot. The professional is buying the "
         "convenient showing rather than the cheapest one."),
    61: ("Triangle = total value of unlimited movies",
         "With marginal cost at zero the customer keeps watching until the "
         "next film is worth nothing to her, so the whole area under the "
         "demand curve is there to be taken. Demand here runs from $4 down "
         "to zero at 12 movies, so that triangle is 12 times 4 over 2, "
         "which is $24. Charge a $24 flat fee for unlimited access and she "
         "is just willing to pay it. That is twice the $12 the best single "
         "price earned, and she watches 12 films instead of 6. A sliding "
         "scale of quantity discounts does the same job: you march the "
         "consumer down her own demand curve, charging less for each "
         "successive unit, and collect the same surplus."),
    62: ("This pricing strategy is a mix of flat fee and versioning",
         "Netflix as versioning. Three plans on the table: Standard with "
         "Ads at $8.99, Standard at $19.99, and Premium at $26.99. What "
         "separates them costs Netflix almost nothing. The number of "
         "screens you can watch at once, the download slots, 4K, and "
         "whether ads run. The price differences are far larger than the "
         "cost differences, which is the whole point of versioning. Worth "
         "asking the class which plan they are on and why they did not take "
         "the cheaper one. There is a good discussion of the 2017 price "
         "increase and the backlash it drew at "
         "https://finance.yahoo.com/news/netflix-price-hikes-wont-hurt-"
         "subscriber-growth-211526185.html"),
    80: ("Three purchase options",
         "Note that the pricing strategy can also be expressed as 100 cards "
         "at $25, 125 at $30, and 175 at $35. Consumers will opt to "
         "purchase 175 at $35, and Walmart still increases its producer "
         "surplus."),
    90: ("Customer attention focused on the comparison",
         "The Economist ran three options: web only at $59, print only at "
         "$125, and print plus web also at $125. Faced with all three, 84 "
         "percent took print plus web and only 16 percent the cheap "
         "web-only option. Nobody chose print only, and that is exactly why "
         "it was on the page. It makes the $125 bundle look like a free "
         "upgrade, so attention goes to the comparison between the two $125 "
         "lines rather than to the $59. Ask the room what they would have "
         "picked before showing the split, then go to the next slide for "
         "what happens when the useless option is taken away."),
    91: ("Customer attention drawn to the comparison",
         "Now take the print-only decoy away and leave just web only at $59 "
         "and print plus web at $125. The split flips. Only 32 percent pay "
         "for the expensive one and 68 percent take the $59. Same two real "
         "products at the same two prices, and a third of the revenue is "
         "gone. The decoy never had to sell a single copy to earn its place "
         "on the page. It worked by changing which comparison the reader "
         "made. The write-up is at "
         "https://cxl.com/blog/pricing-experiments-you-might-not-know-but-"
         "can-learn-from/"),
    92: ("Prices ending in 9",
         "Two framing effects, both on the slide. First, prices ending in 9 "
         "or 99 cents. The eye anchors on the leading digits, so $19.99 "
         "reads as nineteen something rather than as basically twenty. "
         "Second, adding a deliberately expensive option pulls attention to "
         "the middle of the range, which is the decoy we just saw at The "
         "Economist. Neither of these changes the product or what it costs "
         "to make. Both change which comparison the customer makes. That is "
         "what price framing is: not what you charge, but how you present "
         "it."),
    94: ("TheCooperReview.com",
         "A backup slide, and a joke that makes the point better than a "
         "definition would. The Cooper Review's seating chart runs First "
         "Class, Economy Comfort and Economy, and then keeps going: Economy "
         "Discomfort, Economy Agony, Economy 2 The Reckoning, and Satan's "
         "Den Economy. It is versioning taken to its logical end. The cheap "
         "version has to be genuinely worse, or the customers who would "
         "have paid for the good one take it instead. Use it if the class "
         "pushes back on basic economy being made unpleasant on purpose."),
}
