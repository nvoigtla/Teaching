import io
for src_name, out_old, out_new in (
        ("_build_Module6.py", '"Module 6 - Revised.pptx"', '"Module 6 - Revised_test.pptx"'),
        ("_build_M6_practice.py", None, None)):
    s = io.open(src_name, encoding="utf-8").read()
    if out_old:
        s = s.replace("OUT = Path(__file__).parent / " + out_old,
                      "OUT = Path(__file__).parent / " + out_new)
        s = s.replace("    n_notes, n_kept = _apply_written_notes(prs)",
                      "    n_notes, n_kept = (0, 0)")
        io.open("_build_side.py", "w", encoding="utf-8", newline="\n").write(s)
    else:
        s = s.replace('Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx',
                      '_pv_test.pptx')
        io.open("_build_pv_side.py", "w", encoding="utf-8", newline="\n").write(s)
