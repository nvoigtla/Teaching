# -*- coding: utf-8 -*-
"""Port Nico's polished-video-deck edits back into _build_Module1.py.

Source: the four decks in "Videos Final/", diffed slide-by-slide against
"Module 1 - Revised.pptx" (id-keyed geometry/text/run diff + click
structure + raw spPr/rels diff). 2026-08-24.

Every replacement asserts its match count so a silent miss is impossible.
"""
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "_build_Module1.py")

src = io.open(PATH, encoding="utf-8").read()
edits = []


def rep(old, new, count=1, label=""):
    global src
    n = src.count(old)
    assert n == count, "expected %d match(es), found %d for %s\n%r" % (
        count, n, label or old[:60], old[:200])
    src = src.replace(old, new)
    edits.append(label or old[:60])


# ---------------------------------------------------------------------------
# 1. Palette: the dim used for outline items that are not currently covered.
#    Nico's video decks write schemeClr bg1 lumMod 75% over a white
#    background = #BFBFBF.
# ---------------------------------------------------------------------------
rep(
    "CREAM = RGBColor(0xFD, 0xF6, 0xE6)\n",
    "CREAM = RGBColor(0xFD, 0xF6, 0xE6)\n"
    "DIM = RGBColor(0xBF, 0xBF, 0xBF)           # 2026-08-24 (Nico): outline\n"
    "                                           # items not currently covered\n"
    "                                           # are shaded (his video decks\n"
    "                                           # use schemeClr bg1 lumMod 75%\n"
    "                                           # over white = #BFBFBF)\n",
    label="palette: DIM")

# ---------------------------------------------------------------------------
# 2. Video 3 is called "Demand and Supply" (title slide + section tag).
# ---------------------------------------------------------------------------
rep('TAG_V3       = "Module 1 · Video 3 · Supply and Demand"',
    'TAG_V3       = "Module 1 · Video 3 · Demand and Supply"',
    label="TAG_V3 -> Demand and Supply")
rep('return make_video_title(prs, "Supply and Demand", 3)',
    'return make_video_title(prs, "Demand and Supply", 3)',
    label="display 77 title -> Demand and Supply")

# ---------------------------------------------------------------------------
# 3. Display 2 — new sub-bullet.
# ---------------------------------------------------------------------------
rep(
    '            ("PhD in Economics in Barcelona", 1),\n'
    '            # the backup jump button is wired in wire_backup_links()\n',
    '            ("PhD in Economics in Barcelona", 1),\n'
    '            # 2026-08-24 (Nico, from the polished Video 1 deck)\n'
    '            ("Visiting positions at Harvard, University College "\n'
    '             "London", 1),\n'
    '            # the backup jump button is wired in wire_backup_links()\n',
    label="display 2: visiting positions sub-bullet")

# display 2's backup action button follows the extra line down
rep("_add_jump_button(sl(2), sl(96), left=Inches(9.80), top=Inches(4.555),",
    "# top hand-tweaked from 4.555 on 2026-08-24 (the new visiting-\n"
    "    # positions sub-bullet pushes the linked line down)\n"
    "    _add_jump_button(sl(2), sl(96), left=Inches(9.80), top=Inches(4.765),",
    label="display 2: jump button 4.555 -> 4.765")

# ---------------------------------------------------------------------------
# 4. Display 9 — "students" -> "partners".
# ---------------------------------------------------------------------------
rep('            ("Make better decisions as managers (and as consumers, "\n'
    '             "students, etc.)", 0),',
    '            # "students" -> "partners" 2026-08-24 (Nico)\n'
    '            ("Make better decisions as managers (and as consumers, "\n'
    '             "partners, etc.)", 0),',
    label="display 9: consumers, partners")

# ---------------------------------------------------------------------------
# 5. Outline data — "the next-best alternative".
# ---------------------------------------------------------------------------
rep('     "In class: the value of the best alternative you gave up"),',
    '     "In class: the value of the next-best alternative you gave up"),',
    label="M1_OUTLINE: next-best alternative")

# ---------------------------------------------------------------------------
# 6. Display 72 — Market Definition wording, emphasis and spacing.
# ---------------------------------------------------------------------------
rep(
    '            ("Extent of market", 0,\n'
    '             {\'bold\': True, \'bullet_style\': \'none\'}),\n'
    '            ("Which products belong to a market?", 0),\n'
    '            ("Simple test to identify the range of products in your "\n'
    '             "market: If the price of another product changes, will demand "\n'
    '             "for your product change?", 1),\n'
    '            ("Relevant for antitrust litigation (in mergers & "\n'
    '             "acquisitions)", 1),\n'
    '            ("Geography boundaries", 0),\n'
    '            ("Coffee shop in Venice (CA) v. gasoline retail v. gold", 1),\n',
    '            # 2026-08-24 (Nico, polished Video 2 deck): 18 pt above\n'
    '            # "Extent of market" (was the slide default 10),\n'
    '            # "Simple test" set bold, "Geographic" boundaries, and\n'
    '            # "v. gold" -> "vs. App purchases"\n'
    '            ("Extent of market", 0,\n'
    '             {\'bold\': True, \'bullet_style\': \'none\',\n'
    '              \'space_before_pts\': 18}),\n'
    '            ("Which products belong to a market?", 0),\n'
    '            ([("Simple test ", {\'bold\': True}),\n'
    '              ("to identify the range of products in your market: If "\n'
    '               "the price of another product changes, will demand for "\n'
    '               "your product change?", {})], 1, {}),\n'
    '            ("Relevant for antitrust litigation (in mergers & "\n'
    '             "acquisitions)", 1),\n'
    '            ("Geographic boundaries", 0),\n'
    '            ("Coffee shop in Venice (CA) vs. gasoline retail vs. App "\n'
    '             "purchases", 1),\n',
    label="display 72: market-definition bullets")

# ---------------------------------------------------------------------------
# 7. Display 74 — provenance line loses " - not leaked".
# ---------------------------------------------------------------------------
rep('              "(figures from documents the companies had to hand over in "\n'
    '              "the merger review \\u2014 not leaked)", size=18, italic=True,',
    '              # "\\u2014 not leaked" dropped 2026-08-24 (Nico)\n'
    '              "(figures from documents the companies had to hand over in "\n'
    '              "the merger review)", size=18, italic=True,',
    label="display 74: drop 'not leaked'")

# ---------------------------------------------------------------------------
# 8. Display 75 — the Covid bullet goes.
# ---------------------------------------------------------------------------
rep('            ("All entertainment?", 0),\n'
    '            ("Did the market change with Covid-19?", 0),\n',
    '            ("All entertainment?", 0),\n'
    '            # "Did the market change with Covid-19?" deleted\n'
    '            # 2026-08-24 (Nico)\n',
    label="display 75: drop Covid bullet")

# ---------------------------------------------------------------------------
# 9. Display 83 / 86 — the red movement arrow gets its arrowhead.
# ---------------------------------------------------------------------------
rep("        _fig_line(slide, fig, (2.5, 6.5), (4.5, 4.5), color=RED,\n"
    "                  weight_pt=4.5)\n",
    "        # arrowhead added 2026-08-24 (Nico): the movement along D\n"
    "        # points UP the curve, i.e. at the path start (2.5, 6.5)\n"
    "        _set_line_ends(\n"
    "            _fig_line(slide, fig, (2.5, 6.5), (4.5, 4.5), color=RED,\n"
    "                      weight_pt=4.5),\n"
    "            head='triangle', tail='none')\n",
    label="display 83: red arrowhead")

rep("        _fig_line(slide, fig, (3, 4), (4.5, 5.5), color=RED,\n"
    "                  weight_pt=4.5).name = \"sdarrow:i\"\n",
    "        # arrowhead added 2026-08-24 (Nico): the movement along S\n"
    "        # points UP the curve; the line is flipV, so that is the tail\n"
    "        _mv_s = _fig_line(slide, fig, (3, 4), (4.5, 5.5), color=RED,\n"
    "                          weight_pt=4.5)\n"
    "        _mv_s.name = \"sdarrow:i\"\n"
    "        _set_line_ends(_mv_s, tail='triangle')\n",
    label="display 86: red arrowhead")

# ---------------------------------------------------------------------------
# 10. Display 93 — italic S in the Note line, and the y-axis title nudged
#     clear of the longer "SHIFT IN SUPPLY AND DEMAND" header.
# ---------------------------------------------------------------------------
rep('            ("Note: An even larger shift in S will lead to a lower price. "\n'
    '             "But quantity unambiguously increases.", 0,\n'
    '             {\'bullet_style\': \'none\'}),',
    '            # S italicised 2026-08-24 (Nico)\n'
    '            ([("Note: An even larger shift in ", {}),\n'
    '              ("S", {\'italic\': True}),\n'
    '              (" will lead to a lower price. But quantity "\n'
    '               "unambiguously increases.", {})], 0,\n'
    '             {\'bullet_style\': \'none\'}),',
    label="display 93: italic S")

rep("        _v4_header(slide, [(\"SHIFT IN SUPPLY \", {}),\n"
    "                           (\"AND\", {'underline': True}),\n"
    "                           (\" DEMAND\", {})])\n"
    "        _v4_shift_chart(slide, d_shift=True, s_shift=True)\n",
    "        _v4_header(slide, [(\"SHIFT IN SUPPLY \", {}),\n"
    "                           (\"AND\", {'underline': True}),\n"
    "                           (\" DEMAND\", {})])\n"
    "        _v4_shift_chart(slide, d_shift=True, s_shift=True,\n"
    "                        # y-axis title hand-moved from (7.350, 1.730)\n"
    "                        # on 2026-08-24 (Nico) to clear this slide's\n"
    "                        # two-line header\n"
    "                        ylab_pos=(7.300, 2.020))\n",
    label="display 93: Price ($) nudge")

rep("def _v4_shift_chart(slide, *, d_shift=False, s_shift=False):",
    "def _v4_shift_chart(slide, *, d_shift=False, s_shift=False,\n"
    "                    ylab_pos=None):")
rep("    fig = SimpleFig(8.1, 6.40, 4.3, 4.05, 12, 12)\n"
    "    _fig_axes(slide, fig, label_size=16)\n"
    "    if d_shift and s_shift:",
    "    fig = SimpleFig(8.1, 6.40, 4.3, 4.05, 12, 12)\n"
    "    _fig_axes(slide, fig, label_size=16)\n"
    "    if ylab_pos is not None:\n"
    "        _move_shape_by_text(slide, \"Price ($)\", ylab_pos)\n"
    "    if d_shift and s_shift:",
    label="_v4_shift_chart: ylab_pos hook")

# ---------------------------------------------------------------------------
# 11. Display 100 — the photo caption moves up under its picture.
# ---------------------------------------------------------------------------
rep('    _add_text(slide, Inches(0.90), Inches(6.00), Inches(4.7), Inches(0.3),\n'
    '              "Portland Street, Southampton, UK", size=12, italic=True,',
    '    # top hand-tweaked from 6.00 on 2026-08-24 (Nico): the caption sits\n'
    '    # right under its picture\n'
    '    _add_text(slide, Inches(0.90), Inches(5.06), Inches(4.7), Inches(0.3),\n'
    '              "Portland Street, Southampton, UK", size=12, italic=True,',
    label="display 100: caption 6.00 -> 5.06")

# ---------------------------------------------------------------------------
# 12. Two small helpers the edits above rely on.
# ---------------------------------------------------------------------------
rep("def _v4_shift_chart(slide, *, d_shift=False, s_shift=False,\n",
    "def _set_line_ends(shape, *, head=None, tail=None):\n"
    "    \"\"\"Put arrowheads on a connector's <a:ln>, the way PowerPoint\n"
    "    writes them (<a:headEnd type=\"triangle\"/> / <a:tailEnd\n"
    "    type=\"none\"/>). head = the path START, tail = the path END, so a\n"
    "    flipV line's visual top is its tail. Both elements sort last inside\n"
    "    <a:ln>, so appending keeps the schema order.\"\"\"\n"
    "    ln = shape.line._get_or_add_ln()\n"
    "    for tag, val in ((\"a:headEnd\", head), (\"a:tailEnd\", tail)):\n"
    "        if val is None:\n"
    "            continue\n"
    "        el = ln.find(qn(tag))\n"
    "        if el is None:\n"
    "            el = ET.SubElement(ln, qn(tag))\n"
    "        el.set(\"type\", val)\n"
    "    return shape\n"
    "\n"
    "\n"
    "def _move_shape_by_text(slide, text, pos):\n"
    "    \"\"\"Move the first shape whose text is exactly `text` to `pos`\n"
    "    (inches). Used to port hand-nudged chart labels.\"\"\"\n"
    "    for sh in slide.shapes:\n"
    "        if sh.has_text_frame and sh.text_frame.text.strip() == text:\n"
    "            sh.left = int(Inches(pos[0]))\n"
    "            sh.top = int(Inches(pos[1]))\n"
    "            return sh\n"
    "    raise KeyError(\"no shape with text %r on this slide\" % text)\n"
    "\n"
    "\n"
    "def _v4_shift_chart(slide, *, d_shift=False, s_shift=False,\n",
    label="helpers: _set_line_ends / _move_shape_by_text")

# ---------------------------------------------------------------------------
# 13. Outline agenda slides: shade the items that are not currently covered
#     (Nico, 2026-08-24 — his polished video decks dim both the circle
#     digit and the item title of every non-current item; the descriptive
#     overview, which has no current item, stays fully navy).
# ---------------------------------------------------------------------------
rep("        run.font.size = Pt(25)\n"
    "        run.font.bold = True\n"
    "        run.font.color.rgb = NAVY\n"
    "        run.font.name = \"Calibri\"\n"
    "        rows = [([(item[0].upper() + item[1:],\n"
    "                   {'bold': True, 'size': 25, 'color': NAVY})], 0,\n"
    "                 {'bullet_style': 'none', 'space_before_pts': 0})]\n",
    "        run.font.size = Pt(25)\n"
    "        run.font.bold = True\n"
    "        # 2026-08-24 (Nico): on a section agenda the items that are not\n"
    "        # currently covered are shaded; the descriptive overview (which\n"
    "        # highlights everything) keeps them all navy\n"
    "        lit = descriptions or i in hi\n"
    "        run.font.color.rgb = NAVY if lit else DIM\n"
    "        run.font.name = \"Calibri\"\n"
    "        rows = [([(item[0].upper() + item[1:],\n"
    "                   {'bold': True, 'size': 25,\n"
    "                    'color': NAVY if lit else DIM})], 0,\n"
    "                 {'bullet_style': 'none', 'space_before_pts': 0})]\n",
    label="make_m1_outline: shade non-current items")

io.open(PATH, "w", encoding="utf-8").write(src)
print("applied %d edit(s):" % len(edits))
for e in edits:
    print("   - " + e)
