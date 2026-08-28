# -*- coding: utf-8 -*-
"""Read the running time of every Panopto video the calendar links to,
and compare it with the minutes stored in _calendar_content.py.

Panopto's viewer fetches its own metadata from DeliveryInfo.aspx, which
returns a JSON blob with a Duration field (seconds). That endpoint answers
anonymously ONLY while the session is shared publicly -- otherwise it
replies {"ErrorMessage": "Unauthorized access.", "LoginRedirect": true},
which is reported here as LOCKED rather than as a failure.

Usage:
    python _video_minutes.py            # every linked video
    python _video_minutes.py m1 m3      # only keys starting m1 / m3

Videos whose length is still (++) in the calendar are flagged with the
number to paste in.
"""
import json
import subprocess
import sys

from _calendar_content import LINKS, WEEKS

SITE = "https://ucla-anderson.hosted.panopto.com"
DELIVERY = SITE + "/Panopto/Pages/Viewer/DeliveryInfo.aspx"


def delivery_id(url):
    """The session GUID out of a Panopto viewer URL, or None."""
    if "panopto" not in url.lower() or "id=" not in url:
        return None
    return url.split("id=", 1)[1].split("&")[0]


def duration_seconds(did):
    """Length of a Panopto session in seconds; None if locked/unreadable."""
    try:
        out = subprocess.run(
            ["curl", "-s", "--max-time", "60", "-X", "POST", DELIVERY,
             "-d", "deliveryId=%s&responseType=json&isEmbed=true" % did],
            capture_output=True).stdout
        data = json.loads(out.decode("utf8"))
    except (OSError, ValueError):
        return None
    if data.get("ErrorCode"):
        return None
    return data.get("Delivery", {}).get("Duration")


def linked_videos():
    """(week, linkkey, label, stored_minutes) for every linked video item,
    in calendar order, without duplicates."""
    seen, out = set(), []
    for wk in WEEKS:
        groups = list(wk.get("prep_groups", []))
        if wk.get("weekend"):
            groups += wk["weekend"]["groups"]
        for g in groups:
            for it in g["items"]:
                if it[0] == "v" and it[1] and it[1] not in seen:
                    seen.add(it[1])
                    out.append((wk["num"], it[1], it[2], it[3]))
    return out


def main():
    want = [a.lower() for a in sys.argv[1:]]
    rows = [r for r in linked_videos()
            if not want or any(r[1].lower().startswith(w) for w in want)]
    if not rows:
        print("no linked videos match %s" % (want or "(all)"))
        return

    locked = []
    for wknum, key, label, stored in rows:
        did = delivery_id(LINKS[key])
        secs = duration_seconds(did) if did else None
        if secs is None:
            locked.append(key)
            print("wk%-2d %-7s %-52s LOCKED (not shared publicly)"
                  % (wknum, key, label[:52]))
            continue
        mins = int(round(secs / 60.0))
        note = ""
        if stored is None:
            note = "   <- calendar says (++), put %d" % mins
        elif stored != mins:
            note = "   <- calendar says %d, update to %d" % (stored, mins)
        print("wk%-2d %-7s %-52s %2d:%02d -> %2d min%s"
              % (wknum, key, label[:52], int(secs) // 60, int(secs) % 60,
                 mins, note))
        sys.stdout.flush()

    if locked:
        print()
        print("locked (share the session publicly to read it): "
              + ", ".join(locked))


if __name__ == "__main__":
    main()
