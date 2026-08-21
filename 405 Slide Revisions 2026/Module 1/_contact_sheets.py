# Build labeled contact sheets from _renders_src for layout study.
import os
from PIL import Image, ImageDraw

FOLDER = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(FOLDER, '_renders_src')

SHEETS = [
    ('_sheet_ic1.png', [('ic', i) for i in range(1, 15)]),
    ('_sheet_ic2.png', [('ic', i) for i in range(15, 29)]),
    ('_sheet_ic3.png', [('ic', i) for i in range(29, 43)]),
    ('_sheet_ic4.png', [('ic', i) for i in range(43, 54)]),
    ('_sheet_vids1.png', [('v1', i) for i in range(1, 4)]
                       + [('v2', i) for i in range(1, 6)]
                       + [('v3', i) for i in range(1, 8)]),
    ('_sheet_vids2.png', [('v3', i) for i in range(8, 11)]
                       + [('v4', i) for i in range(1, 8)]),
    ('_sheet_mw.png', [('mw', i) for i in (11, 12, 26, 52, 55, 64, 67, 68, 71)]),
]

COLS = 3
PAD = 8
LABEL_H = 18

for out_name, cells in SHEETS:
    imgs = []
    for tag, i in cells:
        p = os.path.join(SRC, tag, 's%02d.png' % i)
        imgs.append((f'{tag} s{i}', Image.open(p)))
    w = max(im.width for _, im in imgs)
    h = max(im.height for _, im in imgs)
    rows = (len(imgs) + COLS - 1) // COLS
    sheet = Image.new('RGB', (COLS * (w + PAD) + PAD,
                              rows * (h + LABEL_H + PAD) + PAD), 'white')
    d = ImageDraw.Draw(sheet)
    for k, (label, im) in enumerate(imgs):
        r, c = divmod(k, COLS)
        x = PAD + c * (w + PAD)
        y = PAD + r * (h + LABEL_H + PAD)
        d.text((x + 2, y + 2), label, fill='red')
        sheet.paste(im, (x, y + LABEL_H))
    dest = os.path.join(FOLDER, out_name)
    sheet.save(dest)
    print(out_name, sheet.size)
