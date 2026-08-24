import sys
import _diff_slides as D
CAN = "Module 2 - In Class Revised.pptx"
TEST = "Module 2 - In Class Revised_test.pptx"
def show(deck, d, lbl):
    s, n = D.dump(deck, d)
    print("--- %s display %d : %d shapes" % (lbl, d, len(s)))
    for i, sh in enumerate(s):
        print("  %2d %s%-4s (%7.3f,%7.3f) %7.3fx%6.3f | %s | %s"
              % (i, "  "*sh[1], sh[0], sh[2], sh[3], sh[4], sh[5], sh[6], sh[7]))
    return n
a=int(sys.argv[1]); b=int(sys.argv[2])
n1=show(CAN,a,"CANON"); n2=show(TEST,b,"BUILD")
if D.norm(n1)!=D.norm(n2):
    print("=== CANON NOTES ==="); print(n1)
    print("=== BUILD NOTES ==="); print(n2)
