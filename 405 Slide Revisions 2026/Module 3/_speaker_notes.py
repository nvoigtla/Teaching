"""MBA-friendly speaker notes for each of the 78 slides in Module 3
(post-revision-v1: Tesla → Rivian, designer → AI researcher, Burn60 →
ChatGPT, iPhone 11 → iPhone 17, Airbus → Alphabet, + 2 new slides on
Meta Reality Labs and AI training costs).

Style: 2 – 4 sentences, natural spoken voice, MBA tone, tied to the
underlying economic concept and the visible example.
"""

NOTES = {
    1: "Welcome – this is Module 3, Production and Costs. Last time we wrapped up the demand side; tonight we tackle the supply side and how output depends on inputs. By the end you'll have all the pieces you need for Module 4, where we put demand and costs together to find profit-maximizing decisions.",

    2: "Quick housekeeping before we dive in: Zoom rules, midterm timing, and any last-minute logistics. Then we move on – this slide is here to clear the desk.",

    3: "A 60-second reminder of where Module 2 left us: demand curves, price elasticity, and marginal revenue. The revenue side is settled; tonight we crack open the cost side. Once both sides are in hand, profit-maximization in Module 4 falls out almost mechanically.",

    4: "Today's flow in one slide. We do Production first – how inputs become output and how to choose them – then Costs – which costs actually matter for decisions, and why scale changes the picture. The class ends with a preview of how all this sets up Module 4.",

    5: "The big-picture frame for the night. Every executive decision you've ever made – pricing, hiring, capacity, sourcing, outsourcing – is at some level a production-and-cost decision. The point of this module is to give you a precise lens for the kind of thinking you already do informally.",

    6: "The module outline, shown once. Two parts and four sections: Production (short run + long run), then Costs (cost concepts + scale and scope). We'll spend roughly equal time on each part.",

    7: "Entering Part 1 – Production. The core question for the next 40 minutes: how does output depend on inputs? We'll do short-run hiring decisions first, then long-run input choice.",

    8: "The production function in its most basic form: output Q is a function of capital K, labor L, and materials M. Capital is more than buildings – it includes machinery, software, IP, AI systems – anything you've already paid for that keeps producing. Everything else in the module is built on this expression.",

    9: "The single most important time-scale distinction in this course. Short run = your capacity (K) is fixed. You can only adjust labor and materials. Long run = everything is variable, including the plant itself. The factory walls are literally what defines 'short run.'",

    10: "Concrete example: Tesla's Gigafactory weekly output as a function of workers, given fixed K. Note the corrected table – last year's version had an increasing-marginal-returns error that I've fixed in place. Look at the pattern: more workers, more output, but the gains per worker shrink.",

    11: "Same data as the previous slide, now plotted. The shape – rising but flattening – is the visual signature of diminishing returns. Memorize this curve shape; you'll see it everywhere.",

    12: "MPL – marginal product of labor – is simply the extra output you get from one more worker. It's the slope of the total-output curve from the last slide. With K fixed, MPL falls as L rises: the slope flattens.",

    13: "The headline of this section. Diminishing MPL is a near-universal feature of short-run production: each additional worker has to share the same fixed capital, so the marginal contribution shrinks. This isn't a quirk of Tesla – it's nearly always true.",

    14: "A favorite historical example. The Black Death killed roughly 40% of Europe's labor force in the 14th century. Wages of survivors rose sharply – consistent with their (now higher) marginal product. Real-world proof of marginal-product reasoning, 600 years before economists named it.",

    15: "Setup for the next eight slides. We're going to derive Tesla's optimal hiring level. Given: cars sell at $90K, materials cost $40K per car, capital is fixed. Question: how many workers should Tesla hire?",

    16: "MRPL = Marginal Revenue Product of Labor. In plain terms: how much extra revenue does one more worker produce? It equals MPL times the price per unit. This is what a worker is 'worth' to the firm in dollar terms.",

    17: "Same definition, more carefully. Notice: when MPL falls, MRPL falls. So even if the price stays constant, each additional worker is worth less than the previous one. The economic value of a marginal hire shrinks as you scale up.",

    18: "A specific number to anchor the concept. Walk through it: at 6,000 workers and 100 robots, output per additional worker is X, price per car is $50K (after subtracting materials), so MRPL is X times $50K.",

    19: "Quick PollEv – compute the MRPL at 6,000 employees and submit. Give them 30 seconds. The point isn't to get the number perfectly, it's to make sure everyone is doing the calculation in their head.",

    20: "Reveal the answer. Walk through anyone's confusion – the most common slip is forgetting to net out materials cost.",

    21: "The hiring rule in one sentence. Hire as long as the next worker brings in more than they cost. Stop the moment the next worker just breaks even. That's it – the rest is just applying this in different settings.",

    22: "The same rule, in algebra: MRPL = w at the optimum. Burn this into your brain. Every short-run hiring problem you encounter as an executive is some version of this comparison.",

    23: "We've been assuming wages are constant. Reality check: for big enough employers, hiring more workers can push the wage up. A frontier AI lab can't just pay the market wage when it adds 100 senior researchers in one year.",

    24: "The technical term is monopsony, but you don't need the word. The intuition: as a big employer hires more, the local talent pool tightens and you pay more for everyone, not just the new hire. The 'true' marginal cost of labor includes this wage-bidding-up effect.",

    25: "Real wage data across firms of different sizes. The pattern – larger firms tend to pay more – is consistent with the wage-search story, though many other things matter too (productivity, location, benefits).",

    26: "A torn-from-the-headlines wage-search example. Anthropic wants to poach a star researcher from Google DeepMind. She's rare; the third hire requires bumping up the existing senior researchers too. So the third hire is way more expensive than her salary alone. The 2024-2026 AI talent wars are the textbook wage-searcher story.",

    27: "PollEv – what's the full marginal cost of the new researcher? Watch for the common trap of just reporting her $5M salary; the real answer includes the raises paid to researchers 1 and 2.",

    28: "Reveal: marginal cost of the third researcher is $8M, not her $5M salary. The lesson: when you're a big enough buyer of scarce talent (or anything), your hiring moves the market price. Factor that in. The same logic applied at Meta when they paid up to keep AI researchers from leaving in 2024.",

    29: "The empirical question that closes this section. Do real-world wages roughly equal MRPL? The UC wage-comparison tool lets you check for yourself. Spoiler: yes, broadly, but with persistent gaps that economists still argue about.",

    30: "Switching gears now from short run to long run. In the long run, capacity is no longer fixed – we get to choose K AND L from scratch. New decision: what's the right MIX of capital and labor?",

    31: "Concrete setup: Rivian building its new plant in Stanton Springs, Georgia (recently revived after the VW partnership in 2024). They get to pick everything – plant size, machinery, workforce, layout. How should they choose?",

    32: "The general framework: we need a decision rule for combining inputs when both are variable. The rule will look familiar – it's the same logic as the short-run hiring rule, generalized.",

    33: "The 'bang for the buck' rule. Spend each additional dollar on whichever input gives you the most extra output per dollar. At the optimum, MP per dollar is the same across all inputs: MP_K / p_K = MP_L / w.",

    34: "Step-by-step procedure for applying the rule on the exam, and in practice. Compute MP/$ for each input; if they're not equal, shift dollars toward the higher one.",

    35: "Real data from Rivian's Georgia plant project. We'll apply the bang-for-the-buck rule to actual numbers and see whether the current mix is optimal.",

    36: "The production function for Georgia in numbers. Setup question: is Rivian's current K/L mix optimal? Don't answer yet – the next slide does the math.",

    37: "Apply the rule. Compute MP/$ for K and MP/$ for L. If they're not equal, the mix isn't optimal and Rivian should shift toward the input with higher bang-for-the-buck.",

    38: "PollEv – vote on whether Rivian's current mix is optimal. Some will say yes, some no. Reveal in the next slide.",

    39: "Reveal: the mix isn't optimal. Discuss which input was underused, and what Rivian should do to fix it – hire more L or buy more robots.",

    40: "Comparative statics. When input prices change, the optimal mix shifts: a tax on robots pushes Rivian toward more labor; rising wages push them toward more automation. Real strategic implications for any firm facing input-price shocks – including AI labs deciding between GPU spend and engineer headcount.",

    41: "The bang-for-the-buck rule isn't just for factories. You apply it every week at the grocery store – balancing what you spend on each item against the extra utility you get. Same logic, different decision.",

    42: "Part 2 – Costs. We've covered what to PRODUCE; now we cover what it COSTS, and crucially, which costs actually matter for decisions.",

    43: "The single most important cost concept for executives, in five words: sunk costs are not costs. They've already been spent; they cannot be recovered. Any forward-looking decision should ignore them. Period.",

    44: "A group-work exercise. Two cars, same daily operating cost, but different sunk amounts. Which should you drive today? Whatever your gut says, the answer is: ignore the sunk cost – it's the same either way.",

    45: "Hollywood case. Kevin Costner's Waterworld – they knew it would flop, why release it anyway? Answer: even a flop adds revenue net of marketing/release costs. The hundreds of millions already spent on production are sunk. Decision should be forward-looking only.",

    46: "The decision tree behind the Waterworld decision. Walk through the numbers: release-anyway revenue exceeds the marginal cost of release. So they released. Sunk costs are sunk.",

    47: "The same logic, in a current strategic context. Meta's Reality Labs has lost roughly $50B from 2020-2024 on Metaverse and VR investments. Wall Street keeps asking when it pays off. Zuckerberg keeps investing – correctly – because the past losses are sunk. The right question is forward-looking: does the next $10B have positive expected value? Same lesson as Waterworld, dressed in 2025 clothes.",

    48: "The flip side of sunk costs is opportunity cost. Apple killed Project Titan in 2024 after roughly a decade and $10B spent. The sunk costs were sunk. What killed the project was opportunity cost: those engineers and that capital had a higher-MPL use in Apple Intelligence and AI.",

    49: "Quick reference: fixed, variable, sunk, marginal, average, opportunity. A cheat sheet you'll refer to for the rest of the module. Make sure you can give a one-sentence example of each.",

    50: "Cost concepts in the real world – a page from Ross Stores' annual report. Have students classify each line as fixed or variable. The point is to ground the abstract concepts in something they'll see in a 10-K.",

    51: "A pricing case students will recognize from their own subscriptions. ChatGPT Plus costs $20/user/month; ChatGPT Team costs $25/user/month but requires a 2-user minimum. If you have one user on Plus and want to add a second, what's the marginal cost of that second user? Hint: it's not just $25.",

    52: "PollEv – compute MC of adding the second user. Watch for them assuming MC = the Team rate of $25.",

    53: "Reveal: MC = $30/user-month for the second user (you go from $20 to $50 total). The lesson: tiered subscription pricing can hide a marginal cost that's HIGHER than the average rate. The opposite of the classic 'volume discount' story – and increasingly common in SaaS.",

    54: "Same concept applied to finance: when a bigger loan comes with a worse rate, the marginal cost of the extra dollars is higher than the average rate the loan was quoted at. This is a very common executive trap when comparing financing options.",

    55: "Now back to Rivian. The Georgia plant's weekly total cost function – how total cost varies with output Q.",

    56: "Decompose total cost into its components: fixed plus variable, broken into sub-categories. Visual.",

    57: "Same Rivian data, but expressed per vehicle: average cost, average variable cost, and marginal cost. Three curves on one chart; make sure you can read each.",

    58: "iPhone manufacturing teardowns – what do the parts and assembly actually cost for the current-gen iPhone 17? Discuss what's included (components, labor, assembly) and what's not (R&D, marketing, retail).",

    59: "PollEv – estimate AVC of an iPhone 17 given the teardown data.",

    60: "Reveal: AVC ≈ $580 for iPhone 17 (vs ~$1,199 retail). About half of retail. The rest – fixed-cost recovery, gross margin – funds Apple's R&D, retail, and ecosystem.",

    61: "iPhone naïve cost function – total cost as output rises. Linear: a fixed cost plus a constant marginal cost.",

    62: "Same data per unit. Note the constant MC and declining AC – the classic shape when fixed costs spread over more units.",

    63: "Real cost functions are more complex – capacity constraints, batch effects, increasing returns. We dig into one specific reason next: economies of scale.",

    64: "§2.2 now – how do costs change as you change SCALE? This is where long-run matters again. We'll cover economies of scale, diseconomies, and scope.",

    65: "Short run vs. long run, restated for costs. Short run: stuck with your current plant. Long run: pick the right plant size for your output level. Bigger plants make sense for bigger volumes.",

    66: "The long-run average cost (LRAC) curve is the envelope of short-run AC curves: at each output level, the lowest possible cost across all possible plant sizes. This is the cost frontier you're aiming for in long-run planning.",

    67: "The classic insight: in many industries, bigger means cheaper. Increasing returns to scale. The next two slides unpack why, with an AI example and a classic aviation example.",

    68: "Three reasons bigger is often cheaper: tubes vs. cubes (volume scales faster than surface area, so storage is cheaper at scale), worker specialization, and indivisible capital that's only economic above a certain volume.",

    69: "The most dramatic modern example: AI training costs. Training a frontier model like GPT-5 or Claude Opus 4.7 reportedly costs hundreds of millions in compute. Once trained, the marginal cost per query is fractions of a cent. Spread across 300M+ users, that fixed cost dilutes to near-zero per user. NVIDIA H100/B200 GPU clusters cost $40M+ each – pure lumpy capital that's only economic at massive scale. This is exactly why only a handful of labs compete in foundation models.",

    70: "Aviation as the textbook example. Cost per seat-mile is much lower on a Boeing 787 than an Embraer ERJ-145. Bigger plane, bigger volume, dramatically lower unit cost.",

    71: "But scale has limits. Above some size, AC starts rising again: coordination costs, principal-agent problems, bureaucracy. There's a sweet spot – the bottom of the U-shaped LRAC curve.",

    72: "A different idea – economies of SCOPE. Sharing capabilities across products. Alphabet is the textbook case: Search, YouTube, Cloud, Waymo, and DeepMind all share data, infrastructure, and AI know-how. Scope is about breadth, not size. The same engineer or training pipeline can support multiple products.",

    73: "Class discussion: Amazon. Scale or scope, or both? Both – they have scale in fulfillment (giant warehouses), scope across product categories (retail to AWS to media to AI). Recognize both when you see them.",

    74: "Final mini-case – a real Shark Tank pitch. We'll dig into the firm's cost structure and evaluate two competing deal offers. Costs sit at the center of every entrepreneurial decision.",

    75: "The numbers behind the pitch: volume last year, expected this year, and the two deal structures (royalty per can vs. equity stake).",

    76: "PollEv – does this firm have economies of scale? Vote, then discuss what the cost data tells us.",

    77: "PollEv – given the cost data, which deal would you take? This is the decision the entrepreneurs had to make in real time.",

    78: "Synthesis. We've covered Production (MRPL=w, bang-for-the-buck) and Costs (sunk vs. opportunity, MC vs. AC, scale). In Module 4 we'll combine demand from M2 with cost from M3 to find profit-maximizing price and quantity. See you next time.",
}


if __name__ == "__main__":
    print(f"Total slides covered: {len(NOTES)}")
    print(f"Total characters of note text: {sum(len(v) for v in NOTES.values())}")
