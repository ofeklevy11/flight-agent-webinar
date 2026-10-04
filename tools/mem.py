"""mem.py - the agent's memory writer. Append-only markdown files in memory/.

  python tools/mem.py problem "what went wrong / how it was solved" [--skill exact|flex|cheapest|build]
  python tools/mem.py search --skill exact --query "TLV-BUD 01/11-20/11" --pairs 13 --sec 12 --file output/x.xlsx [--status ok|partial|failed] [--note "..."]
  python tools/mem.py dest --iata BUD --name Budapest --dates "01/11-20/11/2026" --cheapest 434 --nonstop yes --skill exact [--note "..."]
  python tools/mem.py pref "short sentence about what the user likes / default he asked for"
  python tools/mem.py learn "technical fact worth remembering next time"
  python tools/mem.py show problems|searches|destinations|prefs|learnings [-n 15]

Every entry gets a date and time (Israel local). Files never get rewritten, only appended to.
"""
import argparse
import io
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEM = os.path.join(ROOT, "memory")
FILES = {
    "problems": ("PROBLEMS.md", "# בעיות שנתקלנו בהן\n\nנוסף אוטומטית תוך כדי עבודה. כל שורה: מתי, באיזה סקיל, מה קרה ומה עשינו.\n\n| מתי | סקיל | מה קרה / מה פתר |\n|---|---|---|\n"),
    "searches": ("SEARCHES.md", "# יומן חיפושים\n\nכל ריצה אמיתית של הסוכן.\n\n| מתי | סקיל | חיפוש | פעולות שנסרקו | שניות | סטטוס | קובץ | הערה |\n|---|---|---|---|---|---|---|---|\n"),
    "destinations": ("DESTINATIONS.md", "# יעדים שנסרקו\n\nמה היה המחיר הזול ביותר שמצאנו, כדי שאפשר יהיה להשוות בפעם הבאה.\n\n| מתי | יעד | קוד | תאריכים | הכי זול ₪ | ישירה? | סקיל | הערה |\n|---|---|---|---|---|---|---|---|\n"),
    "prefs": ("PREFERENCES.md", "# העדפות של אופק\n\nברירות מחדל קבועות: יציאה מתל אביב, נוסע אחד, מחלקת תיירים, שקלים.\n\nמה שנלמד תוך כדי:\n\n"),
    "learnings": ("LEARNINGS.md", "# מה למדנו (טכני)\n\nעובדות שחסכו או יחסכו זמן.\n\n"),
}


def now():
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def path(kind):
    name, head = FILES[kind]
    p = os.path.join(MEM, name)
    os.makedirs(MEM, exist_ok=True)
    if not os.path.exists(p):
        with io.open(p, "w", encoding="utf-8") as f:
            f.write(head)
    return p


def esc(s):
    return str(s or "").replace("|", "/").replace("\n", " ")


def add(kind, line):
    with io.open(path(kind), "a", encoding="utf-8") as f:
        f.write(line + "\n")
    print(f"נוסף ל-{FILES[kind][0]}")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("problem"); p.add_argument("text"); p.add_argument("--skill", default="-")
    p = sub.add_parser("search")
    for k in ("skill", "query", "file", "note"):
        p.add_argument("--" + k, default="")
    p.add_argument("--pairs", default="1"); p.add_argument("--sec", default=""); p.add_argument("--status", default="ok")
    p = sub.add_parser("dest")
    for k in ("iata", "name", "dates", "cheapest", "nonstop", "skill", "note"):
        p.add_argument("--" + k, default="")
    p = sub.add_parser("pref"); p.add_argument("text")
    p = sub.add_parser("learn"); p.add_argument("text")
    p = sub.add_parser("show"); p.add_argument("kind", choices=["problems", "searches", "destinations", "prefs", "learnings"]); p.add_argument("-n", type=int, default=15)
    a = ap.parse_args()
    if a.cmd == "problem":
        add("problems", f"| {now()} | {esc(a.skill)} | {esc(a.text)} |")
    elif a.cmd == "search":
        add("searches", f"| {now()} | {esc(a.skill)} | {esc(a.query)} | {esc(a.pairs)} | {esc(a.sec)} | {esc(a.status)} | {esc(a.file)} | {esc(a.note)} |")
    elif a.cmd == "dest":
        add("destinations", f"| {now()} | {esc(a.name)} | {esc(a.iata)} | {esc(a.dates)} | {esc(a.cheapest)} | {esc(a.nonstop)} | {esc(a.skill)} | {esc(a.note)} |")
    elif a.cmd == "pref":
        add("prefs", f"- {now()}: {a.text}")
    elif a.cmd == "learn":
        add("learnings", f"- {now()}: {a.text}")
    else:
        lines = io.open(path(a.kind), encoding="utf-8").read().splitlines()
        body = [l for l in lines if l.startswith("|") or l.startswith("- ")]
        data = [l for l in body if not l.startswith("|---") and not l.startswith("| מתי")]
        print("\n".join(data[-a.n:]) or "(ריק)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
