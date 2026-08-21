# Fetch candidate photos from Wikimedia Commons (search API + Special:FilePath).
# Downloads the top 3 hits per query at width 640 for review; the chosen one
# per key is re-downloaded at width 1400 by passing --final key=title.
import json
import os
import sys
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(HERE, '_source_images')
UA = {'User-Agent': 'Mozilla/5.0 (teaching-slides image fetch; contact: '
                    'nico.voigtlaender@anderson.ucla.edu)'}

QUERIES = {
    'coach': 'Coach store front',
    'michaelkors': 'Michael Kors store',
    'kroger': 'Kroger store exterior',
    'albertsons': 'Albertsons store exterior',
    'costco': 'Costco Wholesale exterior',
    'united': 'United Airlines Boeing 787',
    'wb': 'Warner Bros Studios water tower',
}


def api(params):
    url = ('https://commons.wikimedia.org/w/api.php?'
           + urllib.parse.urlencode(params))
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode('utf-8'))


def download(title, dest, width):
    fp = ('https://commons.wikimedia.org/wiki/Special:FilePath/'
          + urllib.parse.quote(title.replace('File:', ''))
          + '?width=%d' % width)
    req = urllib.request.Request(fp, headers=UA)
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read()
    with open(dest, 'wb') as f:
        f.write(data)
    return len(data)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--final':
        # --final key "File:..." : fetch the chosen file at width 1400
        key, title = sys.argv[2], sys.argv[3]
        dest = os.path.join(IMG_DIR, 'web_%s.jpg' % key)
        n = download(title, dest, 1400)
        print('final %s <- %s (%d bytes)' % (key, title, n))
        return
    manifest = {}
    for key, query in QUERIES.items():
        res = api({'action': 'query', 'list': 'search',
                   'srsearch': query, 'srnamespace': 6,
                   'srlimit': 6, 'format': 'json'})
        titles = [h['title'] for h in res['query']['search']
                  if h['title'].lower().endswith(('.jpg', '.jpeg', '.png'))]
        manifest[key] = []
        for i, title in enumerate(titles[:3]):
            dest = os.path.join(IMG_DIR, 'webpick_%s_%d.jpg' % (key, i))
            try:
                n = download(title, dest, 640)
                manifest[key].append(title)
                print('%s[%d] %s (%d bytes)' % (key, i, title, n))
            except Exception as e:
                print('%s[%d] FAILED %s: %s' % (key, i, title, e))
    with open(os.path.join(HERE, '_webpick_manifest.json'), 'w',
              encoding='utf-8') as f:
        json.dump(manifest, f, indent=1)


if __name__ == '__main__':
    main()
