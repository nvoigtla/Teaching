# Contact sheets for the rebuilt deck's renders.
import os, sys
from PIL import Image, ImageDraw

FOLDER = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(FOLDER, '_renders_new')
COLS = 3
PAD = 8
LABEL_H = 18

files = sorted(f for f in os.listdir(SRC) if f.endswith('.png'))
nums = [int(f[1:3]) for f in files]
per_sheet = 12
groups = [nums[i:i + per_sheet] for i in range(0, len(nums), per_sheet)]
for gi, group in enumerate(groups, 1):
    imgs = [(f's{n}', Image.open(os.path.join(SRC, 's%02d.png' % n)))
            for n in group]
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
    dest = os.path.join(FOLDER, f'_new_sheet_{gi}.png')
    sheet.save(dest)
    print(dest, sheet.size)
