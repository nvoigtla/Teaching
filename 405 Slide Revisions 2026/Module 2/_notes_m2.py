# -*- coding: utf-8 -*-
"""Speaker notes for "Module 2 - In Class Revised.pptx" (BUILD INPUT).

Keyed by DISPLAY slide number.  Applied at the end of build() to every
slide that does not already set its own notes, so the substantive notes
carried over verbatim from Nico's original deck (slides 15, 16, 21, 26,
28, 52, 73 and the rest) stay untouched.

Not listed here, on purpose:
  * the 16 PollEverywhere slides (4, 5, 11, 12, 32, 33, 37, 38, 42, 43,
    49, 50, 61, 62, 69, 70) — their notes ARE the poll mechanism; the
    add-in reads the poll URL from them and a rewritten notes part
    crashes the slideshow deck-wide;
  * slide 13 (the live Excel embed) — spliced from the original deck, so
    its notes are injected by _splice_media.py instead.

Style follows the Teaching CLAUDE.md default: 2-4 sentences, natural
spoken voice, the concrete example named, and the hand-off to the next
slide.  These double as student guidance when the deck is uploaded.
"""

NOTES = {

    1: (
        "Welcome to Module 2. Module 1 gave us the general toolkit: "
        "supply and demand, opportunity cost, sunk costs, and thinking "
        "at the margin. This module puts that toolkit to work on one "
        "side of the market, demand. The goal is to get from “customers "
        "buy less when the price rises” to a number you can actually "
        "use in a pricing decision."),

    2: (
        "Quick housekeeping before we start. Problem Set 1 is due at the "
        "date on the slide, and the TA will email the solutions. Take "
        "the short math test on BruinLearn: it is diagnostic rather than "
        "graded, and it tells you whether to watch the TA math review "
        "videos before we reach the elasticity computations later in "
        "this module. The individual midterm assignment has a flexible "
        "3.5-hour window; budget about 3 hours of work plus half an hour "
        "to scan and upload."),

    3: (
        "A one-minute recap of Module 1. We set up the supply-and-demand "
        "framework for looking at a market, and we agreed on three "
        "habits of thought: count opportunity costs as real costs, "
        "ignore sunk costs, and compare marginal benefit with marginal "
        "cost rather than averages. Everything in this module is an "
        "application of that last habit. We apply it first to a "
        "consumer's purchase decision, and later to a firm's pricing "
        "decision."),

    6: (
        "This is where Module 2 sits in the course. Module 1 covered the "
        "basic principles and the economic way of thinking. We are now "
        "in Part 2, Value and Demand. Supply and cost come next, and "
        "then we put the two sides together in markets, pricing, and "
        "strategy. Keep the map in mind: the demand curve we build "
        "today is the object we price against for the rest of the "
        "course."),

    7: (
        "Six topics in this module. We start with the law of demand and "
        "where the demand curve actually comes from. Then elasticities: "
        "how strongly quantity responds to price, to income, and to the "
        "prices of other goods. The next three topics connect demand to "
        "money, through total revenue, the cases where a price increase "
        "raises revenue and the cases where it backfires, and marginal "
        "revenue. We close with demand estimation, which is how you "
        "would get these numbers out of data. Today's class covers the "
        "first two topics; the rest is in the post-work videos."),

    8: (
        "We start with the law of demand. The aim is not the slogan, "
        "since everyone already knows that demand slopes down. The aim "
        "is the reason behind it, because the reason is what tells you "
        "how much quantity will actually move when you change your "
        "price."),

    9: (
        "Start with the assumption, because it is the part people "
        "forget. Ceteris paribus means we hold everything else fixed: "
        "income, weather, product attributes, how customers perceive "
        "quality. Only under that condition can we say that a lower "
        "price raises quantity demanded, which is what makes the curve "
        "slope downward. If your volume rose last quarter and you also "
        "cut the price, you cannot attribute all of it to the price cut "
        "unless nothing else moved. There are two separate reasons for "
        "the downward slope, and we take them one at a time."),

    10: (
        "First reason: at a lower price, more people are willing to buy "
        "at all. Let us make that concrete with the pizza. Ask yourself "
        "what you would pay for one slice of Gjelina pizza when you are "
        "hungry. Not what it costs, but what it is worth to you. That "
        "maximum willingness to pay is your marginal personal value, "
        "MPV, of the slice. We collect your answers in the poll and "
        "build a demand curve out of them."),

    14: (
        "Second reason for the downward slope. Even a single customer "
        "buys more when the price falls. Think about the pizza again: "
        "you would pay a lot for the first slice, less for the second, "
        "less again for the third, and essentially nothing for the "
        "millionth. That is diminishing marginal personal value, and it "
        "is the second engine behind the demand curve."),

    19: (
        "This is the consumer's optimization in one picture. The "
        "downward-sloping curve is MPV: the value of the next movie "
        "falls once you have already watched several this week. The "
        "rising red line is marginal cost including opportunity cost. "
        "Each additional movie costs you the ticket plus the next-best "
        "use of two hours, and you give up your best alternatives "
        "first, so the cost of one more rises as you watch more. Keep "
        "watching as long as MPV is above MC, and stop where the two "
        "cross, at Q*. Read the same picture the other way and the MPV "
        "curve IS the demand curve: for any price on the vertical axis, "
        "it tells you how many movies this customer buys."),

    20: (
        "Market demand is the horizontal sum of individual demands. At "
        "each price, add up the quantities, not the prices. Follow one "
        "row across: at a price of 12, consumer 1 buys 1 unit and "
        "consumer 2 buys 2, so the market buys 3. Do that at every "
        "price and you trace out the aggregate curve on the right. Two "
        "consequences: market demand is flatter than any individual's, "
        "and it shifts when the number of customers changes even if no "
        "single customer's demand has moved."),

    23: (
        "When a product is worth more to me because other people use "
        "it, three things follow. Switching becomes harder, because "
        "leaving means leaving your network behind. Markets become "
        "tippy: once one side pulls ahead the advantage compounds, so "
        "small early leads decide the outcome. And the end state is "
        "often winner-take-all. The screenshot is StudiVZ next to "
        "Facebook around 2007, two nearly identical products, where the "
        "one with the larger network took the market."),

    24: (
        "Three reminders before we move to elasticities. Demand "
        "represents willingness to pay across all actual and potential "
        "customers, not only the ones who buy today. We draw it as a "
        "straight line for convenience, and we let price move while "
        "everything else is held constant. One notational nuisance to "
        "keep straight: we plot P on the vertical axis, but the demand "
        "function is Q as a function of P. So Q = 10 − P is the demand "
        "function, and P = 10 − Q is the inverse demand function, which "
        "is technically what we are drawing. Economists are sloppy "
        "about this distinction, but it matters as soon as you "
        "differentiate."),

    25: (
        "That completes the law of demand. Now the part that turns "
        "demand into a number you can use in a decision: elasticities."),

    30: (
        "Own-price elasticity is the percentage change in quantity "
        "demanded divided by the percentage change in the price of the "
        "same product. Because it is a ratio of two percentages it is "
        "unit-free: it does not matter whether you measure in dollars "
        "or euros, gallons or litres, so the number is comparable "
        "across products and countries. By the law of demand the sign "
        "is always negative, since price up means quantity down. Most "
        "of the confusion in this topic comes from that minus sign, so "
        "keep an eye on it."),

    31: (
        "A real decision. A consulting firm estimated the price "
        "elasticity of demand for water in Los Angeles at −0.4. The Los "
        "Angeles Department of Water and Power wants households to use "
        "10% less water, and the only lever it has is price. How big "
        "does the price increase have to be? Work it out on the poll "
        "before we do it together."),

    34: (
        "Elasticity is the percent change in quantity over the percent "
        "change in price, so rearrange it: the percent change in price "
        "equals the percent change in quantity divided by the "
        "elasticity. That is −10% divided by −0.4, which is +25%. So a "
        "25% price increase is needed to cut water use by 10%. The "
        "lesson for anyone setting utility prices is that when demand "
        "is inelastic, it takes a large price move to get a modest "
        "quantity response."),

    36: (
        "Now the same formula run in the other direction. CorePower "
        "Yoga cuts its price by 17% and the number of booked lessons "
        "doubles, which is a 100% increase in quantity. What is the "
        "implied elasticity, and does that make demand elastic or "
        "inelastic? Try it on the poll."),

    39: (
        "The percent change in quantity is +100% and the percent change "
        "in price is −17%, so the elasticity is roughly −6. Notice that "
        "the answer is not “6%”. Elasticity is a unit-free number, not "
        "a percentage. Since it lies below −1, demand for yoga lessons "
        "is elastic: these customers are very price-sensitive, which "
        "makes sense given how many substitutes a yoga studio has. "
        "Compare that with water at −0.4."),

    40: (
        "There are two ways to compute an elasticity, and which one you "
        "use depends on what you know. Method 1 is for when you have "
        "two observed price and quantity points. Compute both "
        "percentage changes relative to the initial point, then divide. "
        "One caution: this only approximates the percentage changes, "
        "and the approximation gets worse the larger the price change, "
        "because the answer depends on which of the two points you call "
        "the initial one. The TA math review videos cover this, and you "
        "will practise it on Problem Set 2."),

    46: (
        "Method 2 is for when you know the whole demand curve rather "
        "than two points. Then you do not need an approximation at all, "
        "because you can use the slope of the demand curve directly. "
        "Formally the slope is the derivative of Q with respect to P, "
        "written dQ/dP, and the elasticity is that slope multiplied by "
        "P over Q. For a linear demand curve the slope is constant, but "
        "the elasticity is not, because P and Q both change as you move "
        "along the curve. We come back to that in a few slides."),

    47: (
        "Let us work one all the way through. The demand function is "
        "given in inverse form, P = 100 − 0.25Q, and we want the "
        "elasticity at a price of 25. Step 1, solve for Q, which gives "
        "Q = 400 − 4P. Step 2, the derivative dQ/dP is −4. Step 3, at "
        "P = 25 the quantity is 400 − 100 = 300. Step 4, put the pieces "
        "into the formula: −4 times 25 divided by 300, which is about "
        "−0.33. Step 1 was only necessary because the function was "
        "handed to us with P on the left-hand side. If you already have "
        "Q as a function of P, start at step 2."),

    48: (
        "Your turn. You are given a firm's demand function and the "
        "price it charges. Compute the elasticity of demand at that "
        "price, using the same four steps we just went through."),

    51: (
        "The answer is −0.25. Step 1 is unnecessary here, because we "
        "already have Q as a function of P. The derivative dQ/dP is −1. "
        "At P = 2 the quantity is 10 − 2 = 8. So the elasticity is −1 "
        "times 2 divided by 8, which is −0.25. That sits between −1 and "
        "0, so demand is inelastic at this price, and as the next slide "
        "shows, that is not a place a firm wants to be."),

    53: (
        "Here is a result worth carrying out of this module: a firm "
        "should not operate on the inelastic part of its demand curve. "
        "The intuition is short. If demand is inelastic, raising the "
        "price loses relatively little quantity, so revenue goes up. "
        "Selling less at the same time lowers variable costs. Revenue "
        "up and costs down means profit up, so a firm sitting in the "
        "inelastic region is leaving money on the table. Uber is the "
        "example on the slide: a 2016 study estimated its elasticity at "
        "about −0.4, Uber then raised prices by an average of 18%, and "
        "prices and revenues are far higher today. We develop the "
        "argument properly in the Module 2 videos."),

    54: (
        "This is the estimated demand curve behind that number. At the "
        "2016 base price the elasticity was about −0.4, solidly "
        "inelastic. I will go through how a curve like this is "
        "estimated in office hours; the estimation itself is not exam "
        "material. What matters here is that these are not textbook "
        "numbers. Firms really do estimate their own demand curves, and "
        "the estimate changes the pricing decision."),

    55: (
        "Elasticity normally changes as you move along a demand curve, "
        "with two exceptions worth knowing. Perfectly elastic demand is "
        "horizontal: at that price customers will take any quantity, "
        "but raise the price by a cent and demand goes to zero. That is "
        "roughly the position of a very small seller in a large "
        "commodity market. Perfectly inelastic demand is vertical: "
        "quantity does not respond to price at all, which is the "
        "short-run picture for something like a life-saving drug. Real "
        "products sit between the two, but these are the anchors."),

    56: (
        "What makes demand elastic? Mostly substitutes: how many there "
        "are, how close they are, and how willing customers are to "
        "switch. Firm size relative to the relevant market matters too. "
        "A smaller firm faces more elastic demand, because its "
        "customers can move to a competitor without leaving the market "
        "at all. That is the bridge to the next slide, which is about "
        "the gap between the elasticity your firm faces and the "
        "elasticity the whole industry faces."),

    58: (
        "We have done own-price elasticity. The same construction, one "
        "percent change divided by another percent change, gives two "
        "more measures that managers use. Income elasticity tells you "
        "how demand moves with customers' income, which is what you "
        "need for a recession scenario. Cross-price elasticity tells "
        "you how demand for your product moves with the price of "
        "another product, which is what you need to identify "
        "substitutes and complements."),

    59: (
        "Income elasticity is the percentage change in quantity "
        "demanded divided by the percentage change in income. It "
        "answers a question every planner asks: if the economy grows 2% "
        "next year, what happens to my volume? Unlike own-price "
        "elasticity, the sign is not fixed here. It can be positive or "
        "negative, and that sign is exactly what classifies the good."),

    60: (
        "A worked case. Average income in the United States rises by "
        "2%, and as a result demand for the Rivian R3 rises by 5%. What "
        "is the income elasticity? Take a moment on the poll."),

    63: (
        "The percent change in quantity is 5% and the percent change in "
        "income is 2%, so the income elasticity is 2.5. It is positive, "
        "which makes the R3 a normal good, and it is above 1, which "
        "makes it a luxury good: demand grows faster than income. That "
        "is a strong reaction, and it is what you would expect for a "
        "discretionary big-ticket purchase. It also means a recession "
        "hits this product harder than the average."),

    64: (
        "The sign and the size of the income elasticity classify the "
        "good. Above zero it is a normal good, so demand rises with "
        "income. Above one it is a luxury good, so demand rises faster "
        "than income, like the Rivian we just computed. Below zero it "
        "is an inferior good: demand actually falls as people get "
        "richer, because they trade up to something else. Store brands, "
        "instant noodles, and secondhand goods are the standard "
        "examples. Nothing pejorative is meant by “inferior”. It is a "
        "statement about the income elasticity, not about quality."),

    65: (
        "Here is why the classification pays off. In 2007 Target's "
        "stock traded well above Walmart's. Then the economy entered "
        "recession and the two moved in opposite directions: Walmart's "
        "stock rose while Target's fell. Income elasticity is the "
        "reason. Walmart's basket behaves more like an inferior good, "
        "so falling incomes pushed shoppers toward it, while Target's "
        "more discretionary mix suffered. If you know your own income "
        "elasticity, you know which way a downturn will move your "
        "volume. The two series here are digitized from the printed "
        "figure, so read the shape rather than the exact levels."),

    67: (
        "The third measure. Cross-price elasticity is the percent "
        "change in the quantity demanded of good X divided by the "
        "percent change in the price of another good Y. The sign is "
        "what carries the meaning. Positive means the two are "
        "substitutes: Y gets more expensive and customers switch to X. "
        "Negative means they are complements: Y gets more expensive and "
        "demand for X falls along with it. This is the number pricing "
        "teams and antitrust authorities use to decide which products "
        "actually compete with each other."),

    68: (
        "A concrete one. A movie theater raises its ticket price from "
        "$15 to $18, and the quantity of popcorn it sells falls by 8%. "
        "What is the implied cross-price elasticity, and what does its "
        "sign tell you about the two products? Take it on the poll."),

    71: (
        "The percent change in the quantity of popcorn is −8%. The "
        "percent change in the ticket price is 18 minus 15 over 15, or "
        "+20%. So the cross-price elasticity is −8 divided by 20, which "
        "is −0.4. Again, not “−0.4%”, because the measure is unit-free. "
        "The negative sign says that popcorn and movie tickets are "
        "complements. That is a large part of why theaters keep ticket "
        "prices lower than you might expect: the ticket is what sells "
        "the popcorn."),

    74: (
        "A one-page summary of own-price elasticity. The definition and "
        "the intuition are at the top. The three categories are in the "
        "middle: inelastic between −1 and 0, unit-elastic at exactly "
        "−1, and elastic below −1. The two computation routes are at "
        "the bottom, the approximation from two observed points and the "
        "exact point elasticity when the whole demand function is "
        "known, where ΔQ/ΔP is the slope of Q as a function of P. Keep "
        "this slide next to you when you do Problem Set 2."),

    75: (
        "That is where we stop in class. The remaining topics, demand "
        "and revenue, elasticity and revenue, and marginal revenue, are "
        "covered in the post-work. Watch Module 2 Videos 1 and 2, and "
        "use Practice Videos 1 and 2 to work through the computations. "
        "Everything is on BruinLearn under “Module 2 Post-Work”."),

    76: (
        "The last topic is demand estimation, which is where these "
        "numbers come from in practice: market experiments and "
        "regression. Module 2 Video 3 covers it and Problem Set 2 gives "
        "you the practice. You will not be asked to run a regression "
        "yourself, but you do need to understand what the estimates "
        "mean and how much weight to put on them."),
}


# Injected by _splice_media.py, because slide 13 is spliced in from the
# original deck and its notes part is replaced wholesale.
SPLICED_NOTES = {
    13: (
        "Here is your own demand curve, built live from the numbers you "
        "just gave me. Sort the willingness-to-pay values from highest "
        "to lowest and you get a downward-sloping step function: at a "
        "high price only the few people with a high MPV buy, and as the "
        "price falls each additional person joins in. Nobody in this "
        "room assumed a downward slope. It came out of your answers. "
        "That is the first reason the demand curve slopes down."),
}
