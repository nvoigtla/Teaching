# -*- coding: utf-8 -*-
"""Read the running time of every module podcast listed in
_calendar_content.py PODCASTS, without downloading the whole file.

Dropbox serves the .m4a over HTTP with range support, so we ask only for
the first 256 KB (and, if the moov atom sits at the end of the file, the
last 1 MB) and parse the duration out of the moov/mvhd box.

Usage:
    python _podcast_minutes.py

Prints one line per episode with the rounded minutes to paste back into
the PODCASTS table. Episodes with no link yet are reported as pending.
"""
import struct
import subprocess
import sys

from _calendar_content import PODCASTS

RANGE_HEAD = "0-262143"      # 256 KB
RANGE_TAIL = "-1048576"      # last 1 MB


def _fetch(url, rng):
    """Range-GET `rng` bytes of `url`; b'' on any failure."""
    cmd = ["curl", "-sL", "--max-time", "120", "-r", rng, url]
    try:
        return subprocess.run(cmd, capture_output=True).stdout
    except OSError as exc:
        print("  curl failed: %s" % exc)
        return b""


def _find_mvhd(buf):
    """Return (timescale, duration) from the first usable mvhd box."""
    i = buf.find(b"mvhd")
    while i != -1:
        p = i + 4
        version = buf[p] if p < len(buf) else None
        if version == 0 and p + 20 <= len(buf):
            ts, dur = struct.unpack(">II", buf[p + 12:p + 20])
            if ts and dur:
                return ts, dur
        elif version == 1 and p + 32 <= len(buf):
            ts, dur = struct.unpack(">IQ", buf[p + 20:p + 32])
            if ts and dur:
                return ts, dur
        i = buf.find(b"mvhd", i + 1)
    return None


def duration_seconds(url):
    """Length of the audio at `url` in seconds, or None if unreadable."""
    # dl=1 gives the file itself; dl=0 gives Dropbox's HTML preview page
    direct = url.replace("dl=0", "dl=1")
    got = _find_mvhd(_fetch(direct, RANGE_HEAD))
    if got is None:
        got = _find_mvhd(_fetch(direct, RANGE_TAIL))
    if got is None:
        return None
    timescale, dur = got
    return dur / float(timescale)


def main():
    pending = []
    for mod in sorted(PODCASTS):
        for key, label in (("intro", "Intro"), ("wrap", "Wrap-Up")):
            url, recorded = PODCASTS[mod].get(key, (None, None))
            name = "Module %d %s" % (mod, label)
            if not url:
                pending.append(name)
                continue
            secs = duration_seconds(url)
            if secs is None:
                print("%-20s  COULD NOT READ (link dead or not audio?)" % name)
                continue
            mins = int(round(secs / 60.0))
            flag = ""
            if recorded is None:
                flag = "   <- put %d in PODCASTS" % mins
            elif recorded != mins:
                flag = "   <- PODCASTS says %s, update to %d" % (recorded, mins)
            print("%-20s  %2d:%02d  ->  %2d min%s"
                  % (name, int(secs) // 60, int(secs) % 60, mins, flag))
            sys.stdout.flush()
    if pending:
        print()
        print("no link yet: " + ", ".join(pending))


if __name__ == "__main__":
    main()
