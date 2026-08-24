# -*- coding: utf-8 -*-
import os
import _diff_cross as X
HERE = os.path.dirname(os.path.abspath(__file__))
MAIN = os.path.join(HERE, "Module 1 - Revised.pptx")
V = lambda n: os.path.join(HERE, "Videos Final", n)
MAP = [
 ("Module 1 - Video 1 - Introduction.pptx", [(1,67),(2,2),(3,9),(4,10),(5,11),(6,13),(7,68),(8,1),(9,69),(10,95),(11,100)]),
 ("Module 1 - Video 2 - Markets.pptx", [(1,70),(2,71),(3,72),(4,73),(5,74),(6,75),(7,76)]),
 ("Module 1 - Video 3 - Demand and Supply.pptx", [(i,i+76) for i in range(1,11)]),
 ("Module 1 - Video 4 - Equilibrium.pptx", [(i,i+86) for i in range(1,8)]),
]
for deck, pairs in MAP:
    print("="*80); print(deck)
    for dv, dm in pairs:
        msgs,_ = X.report(V(deck), dv, MAIN, dm)
        if msgs:
            print("--- video %d  ->  main %d" % (dv, dm))
            for m in msgs: print("      "+m)
        else:
            print("--- video %d  ->  main %d : identical" % (dv, dm))
