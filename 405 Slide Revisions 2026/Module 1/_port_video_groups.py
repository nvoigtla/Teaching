# -*- coding: utf-8 -*-
"""Second half of the Videos-Final port: name the shapes Nico grouped by
hand on displays 75 / 91 / 92 / 93, and register those groups with
_group_pass.py. 2026-08-24."""
import io
import os

HERE = os.path.dirname(os.path.abspath(__file__))
edits = []


def patch(path, pairs):
    src = io.open(path, encoding="utf-8").read()
    for old, new, label in pairs:
        n = src.count(old)
        assert n == 1, "expected 1 match, found %d for %s\n%r" % (
            n, label, old[:200])
        src = src.replace(old, new)
        edits.append("%s: %s" % (os.path.basename(path), label))
    io.open(path, "w", encoding="utf-8").write(src)


# ---------------------------------------------------------------------------
# _build_Module1.py — names on the V4 shift-chart pieces and the badge
# ---------------------------------------------------------------------------
BUILD = os.path.join(HERE, "_build_Module1.py")
patch(BUILD, [
    # the new-equilibrium guides get names so rule 5 can group them
    ("    _fig_guide(slide, fig, (4.5, 4.5), color=GRAY)\n"
     "    _fig_guide(slide, fig, (q1, p1), color=GRAY)\n",
     "    _fig_guide(slide, fig, (4.5, 4.5), color=GRAY)\n"
     "    # 2026-08-24 (Nico): the new-equilibrium guides and their labels\n"
     "    # are one group, and the shifted curve + its label + the shift\n"
     "    # arrow are another (names drive _group_pass.py rule 5)\n"
     "    g1h, g1v = _fig_guide(slide, fig, (q1, p1), color=GRAY)\n"
     "    g1h.name = \"sdguide:h:1\"\n"
     "    g1v.name = \"sdguide:v:1\"\n",
     "_v4_shift_chart: name the P1/Q1 guides"),

    ("    if d_shift:\n"
     "        _fig_line(slide, fig, (3.4, 8.6), (10.6, 1.4), color=GREEN_DK,\n"
     "                  weight_pt=2.75, dash='dash')\n"
     "        _fig_curve_label(slide, fig, 10.7, 1.9, \"D’\", color=GREEN_DK)\n"
     "    if s_shift:\n"
     "        _fig_line(slide, fig, (3.4, 0.9), (10.9, 8.4), color=BLUE_PED,\n"
     "                  weight_pt=2.75, dash='dash')\n"
     "        _fig_curve_label(slide, fig, 10.5, 8.7, \"S’\", color=BLUE_PED)\n",
     "    if d_shift:\n"
     "        _fig_line(slide, fig, (3.4, 8.6), (10.6, 1.4), color=GREEN_DK,\n"
     "                  weight_pt=2.75,\n"
     "                  dash='dash').name = \"sdcurve:Dp\"\n"
     "        _fig_curve_label(slide, fig, 10.7, 1.9, \"D’\",\n"
     "                         color=GREEN_DK).name = \"sdlabel:Dp\"\n"
     "    if s_shift:\n"
     "        _fig_line(slide, fig, (3.4, 0.9), (10.9, 8.4), color=BLUE_PED,\n"
     "                  weight_pt=2.75,\n"
     "                  dash='dash').name = \"sdcurve:Sp\"\n"
     "        _fig_curve_label(slide, fig, 10.5, 8.7, \"S’\",\n"
     "                         color=BLUE_PED).name = \"sdlabel:Sp\"\n",
     "_v4_shift_chart: name D' / S' and their labels"),

    ("    _fig_ylab(slide, fig, p1_lab, \"P1\", size=16)\n"
     "    _fig_xlab(slide, fig, 4.5, \"Q0\", size=16)\n"
     "    _fig_xlab(slide, fig, q1, \"Q1\", size=16)\n",
     "    _fig_ylab(slide, fig, p1_lab, \"P1\", size=16).name = \"sdylab:P1\"\n"
     "    _fig_xlab(slide, fig, 4.5, \"Q0\", size=16)\n"
     "    _fig_xlab(slide, fig, q1, \"Q1\", size=16).name = \"sdxlab:Q1\"\n",
     "_v4_shift_chart: name the P1 / Q1 labels"),

    # display 91's shift arrow + label
    ("        _add_arrow(slide, (9208767, 4715447), (9667492, 4309110),\n"
     "                   color=GREEN_DK, weight_pt=2.0, head=True)\n"
     "        _add_text(slide, fig.x(7.2), fig.y(3.4), Inches(1.5), Inches(0.6),\n"
     "                  \"Shift in\\ndemand\", size=13, bold=True, color=GREEN_DK,\n"
     "                  font=\"Calibri\")\n",
     "        _add_arrow(slide, (9208767, 4715447), (9667492, 4309110),\n"
     "                   color=GREEN_DK, weight_pt=2.0,\n"
     "                   head=True).name = \"sdarrow:shift\"\n"
     "        _add_text(slide, fig.x(7.2), fig.y(3.4), Inches(1.5), Inches(0.6),\n"
     "                  \"Shift in\\ndemand\", size=13, bold=True, color=GREEN_DK,\n"
     "                  font=\"Calibri\").name = \"sdlabel:shift\"\n",
     "display 91: name the shift arrow + label"),

    # display 92's shift arrow + label
    ("        _add_arrow(slide, (9634727, 3729609), (10044302, 4099941),\n"
     "                   color=BLUE_PED, weight_pt=2.0, head=True)\n"
     "        _add_text(slide, 9861803, 3551301, Inches(1.4), Inches(0.6),\n"
     "                  \"Shift in\\nsupply\", size=13, bold=True, color=BLUE_PED,\n"
     "                  font=\"Calibri\")\n",
     "        _add_arrow(slide, (9634727, 3729609), (10044302, 4099941),\n"
     "                   color=BLUE_PED, weight_pt=2.0,\n"
     "                   head=True).name = \"sdarrow:shift\"\n"
     "        _add_text(slide, 9861803, 3551301, Inches(1.4), Inches(0.6),\n"
     "                  \"Shift in\\nsupply\", size=13, bold=True, color=BLUE_PED,\n"
     "                  font=\"Calibri\").name = \"sdlabel:shift\"\n",
     "display 92: name the shift arrow + label"),

    # the discussion badge: box and its overlaid text, so display 75's
    # badge can be grouped the way Nico grouped it
    ("    run.font.color.rgb = NAVY\n"
     "    return shp\n"
     "\n"
     "\n"
     "def _add_callout_box(",
     "    run.font.color.rgb = NAVY\n"
     "    shp.name = \"sdbadge:box\"\n"
     "    txt.name = \"sdbadge:txt\"\n"
     "    return shp\n"
     "\n"
     "\n"
     "def _add_callout_box(",
     "_add_discussion_break: name box + text"),
])

# ---------------------------------------------------------------------------
# _group_pass.py — accept the badge prefix, and register the new groups
# ---------------------------------------------------------------------------
GROUP = os.path.join(HERE, "_group_pass.py")
patch(GROUP, [
    ('            if nm.startswith(("sdcurve:", "sdlabel:", "sdguide:",\n'
     '                              "sdxlab:", "sdylab:", "sdarrow:",\n'
     '                              "sdpic:", "sdcap:")):',
     '            if nm.startswith(("sdcurve:", "sdlabel:", "sdguide:",\n'
     '                              "sdxlab:", "sdylab:", "sdarrow:",\n'
     '                              "sdpic:", "sdcap:", "sdbadge:")):',
     "rule 5: accept sdbadge: names"),

    ('        "Sp": ["sdcurve:Sp", "sdlabel:Sp", "sdarrow:ii", "sdlabel:ii",\n'
     '               "sdguide:h:Q3", "sdguide:v:Q3", "sdxlab:Q3"],\n'
     '    },\n'
     '}',
     '        "Sp": ["sdcurve:Sp", "sdlabel:Sp", "sdarrow:ii", "sdlabel:ii",\n'
     '               "sdguide:h:Q3", "sdguide:v:Q3", "sdxlab:Q3"],\n'
     '    },\n'
     '    # 2026-08-24 (Nico, polished video decks). Display 75: the gold\n'
     '    # discussion badge and the text on it are one object. Displays\n'
     '    # 91-93: the shifted curve(s) with their labels and the shift\n'
     '    # arrow are one beat, and the new-equilibrium guides with P1/Q1\n'
     '    # are another. (On 91 he nested the D’ label around an inner\n'
     '    # group; a flat group of the same four shapes behaves identically\n'
     '    # and matches how he grouped 92 and 93.)\n'
     '    75: {\n'
     '        "badge": ["sdbadge:box", "sdbadge:txt"],\n'
     '    },\n'
     '    91: {\n'
     '        "Dp": ["sdcurve:Dp", "sdlabel:Dp", "sdarrow:shift",\n'
     '               "sdlabel:shift"],\n'
     '        "Q1": ["sdguide:h:1", "sdguide:v:1", "sdylab:P1",\n'
     '               "sdxlab:Q1"],\n'
     '    },\n'
     '    92: {\n'
     '        "Sp": ["sdcurve:Sp", "sdlabel:Sp", "sdarrow:shift",\n'
     '               "sdlabel:shift"],\n'
     '        "Q1": ["sdguide:h:1", "sdguide:v:1", "sdylab:P1",\n'
     '               "sdxlab:Q1"],\n'
     '    },\n'
     '    93: {\n'
     '        "shifts": ["sdcurve:Dp", "sdlabel:Dp", "sdcurve:Sp",\n'
     '                   "sdlabel:Sp"],\n'
     '        "Q1": ["sdguide:h:1", "sdguide:v:1", "sdylab:P1",\n'
     '               "sdxlab:Q1"],\n'
     '    },\n'
     '}',
     "CHART_GROUPS: displays 75 / 91 / 92 / 93"),
])

print("applied %d edit(s):" % len(edits))
for e in edits:
    print("   - " + e)
