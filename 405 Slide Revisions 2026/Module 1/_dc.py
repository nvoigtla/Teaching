# -*- coding: utf-8 -*-
"""Compact text dump of selected slides of a deck (default: the canonical
Module 1 deck). Usage: _dc.py [--deck NAME] <display> [<display> ...]"""
import sys
import zipfile

from lxml import etree as ET

import _diff_slides as D

A, P, R = D.A, D.P, D.R
args = sys.argv[1:]
deck = "Module 1 - Revised.pptx"
if args and args[0] == "--deck":
    deck = args[1]
    args = args[2:]

z = zipfile.ZipFile(deck)
pres = ET.fromstring(z.read("ppt/presentation.xml"))
pr = {r.get("Id"): r.get("Target") for r in
      ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
order = ["ppt/" + pr[s.get(D.q(R, "id"))].lstrip("/")
         for s in pres.find(D.q(P, "sldIdLst"))]

for disp in [int(x) for x in args]:
    part = order[disp - 1]
    t = ET.fromstring(z.read(part))
    npic = sum(1 for _ in t.iter(D.q(P, "pic")))
    ngrp = sum(1 for _ in t.iter(D.q(P, "grpSp")))
    anim = "ANIM" if t.find(D.q(P, "timing")) is not None else "----"
    parts = []
    for sp in t.iter(D.q(P, "sp")):
        s = D.norm("".join(x.text or "" for x in sp.iter(D.q(A, "t"))))
        if not s or s.startswith("Management 405"):
            continue
        if len(s) <= 3 and s.isdigit():
            continue
        parts.append(s)
    print("[%d] %-14s pics=%d grps=%d %s | %s"
          % (disp, part.split('/')[-1], npic, ngrp, anim,
             "  ||  ".join(parts)[:430]))
z.close()
