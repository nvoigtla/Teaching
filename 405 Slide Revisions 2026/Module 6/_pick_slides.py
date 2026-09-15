import sys, re
path, spec = sys.argv[1], sys.argv[2]
want = set()
for part in spec.split(","):
    if "-" in part:
        a,b = part.split("-"); want |= set(range(int(a), int(b)+1))
    else: want.add(int(part))
txt = open(path, encoding="utf-8").read()
blocks = re.split(r"^## Slide ", txt, flags=re.M)
CHROME = ("13.33x0.42", "12.78x0.02", "2.20x0.05", "0.00,7.15", "Management 405  \u00b7  Complex", "12.50,7.20", "\u00b7 Roadmap", "0.28,0.00 12.00x0.42")
for b in blocks[1:]:
    n = int(re.match(r"\d+", b).group())
    if n not in want: continue
    print("## Slide " + b.split("\n",1)[0])
    for line in b.split("\n")[1:]:
        if any(c in line for c in CHROME): continue
        print(line)
