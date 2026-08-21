# One-off: replace the M2 config region of _animate.py with the M1 config.
import io, os

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, '_animate.py')
src = io.open(path, encoding='utf-8').read()

start = src.index('SKIP_TITLE = ')
end = src.index('def q(ns, t):')
new_cfg = io.open(os.path.join(HERE, '_anim_config_m1.txt'),
                  encoding='utf-8').read()
src = src[:start] + new_cfg + '\n\n' + src[end:]

# slide-count bound for "all"
src = src.replace('todo = [d for d in range(1, 78) if d not in SKIP]',
                  'todo = [d for d in range(1, 85) if d not in SKIP]')

io.open(path, 'w', encoding='utf-8').write(src)
print('patched _animate.py (config region %d..%d chars)' % (start, end))
