import sys, io
import _dump_timing as T
CAN = "Module 2 - In Class Revised.pptx"
TEST = "Module 2 - In Class Revised_test.pptx"
def MAP(d): return d
def cap(fn, *a):
    buf = io.StringIO(); old = sys.stdout; sys.stdout = buf
    try: fn(*a)
    finally: sys.stdout = old
    return buf.getvalue()
def strip(s):
    # drop spid (differs across builds) and the header line
    out=[]
    for ln in s.splitlines()[1:]:
        i = ln.find("spid=")
        if i >= 0:
            ln = ln[:i] + ln[i+5:].split(None,1)[1] if len(ln[i+5:].split(None,1))>1 else ln[:i]
        out.append(ln.rstrip())
    return "\n".join(out)
rng = [int(x) for x in sys.argv[1:]] or list(range(1,21))
for d in rng:
    a = cap(T.dump, CAN, d, "C")
    b = cap(T.dump, TEST, MAP(d), "B")
    if strip(a) == strip(b):
        print("=== canon %2d (build %2d)  timing OK" % (d, MAP(d)))
    else:
        print("=== canon %2d (build %2d)  TIMING DIFFERS" % (d, MAP(d)))
        print(a); print(b)
