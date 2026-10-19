"""meetup-deck command line.

    meetup-deck --template deck.pptx --event event.json --stills out/stills [--header header.jpeg] [--canva dir] --out deck.pptx

Exit codes: 0 written, 2 a missing input.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .build import Inputs, build


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="meetup-deck", description="Fill in the group's slide deck for one meetup.")
    ap.add_argument("--template", type=Path, required=True, help="the group's deck, exported as .pptx")
    ap.add_argument("--event", type=Path, required=True, help="events/<id>/event.json")
    ap.add_argument("--stills", type=Path, required=True, help="the Remotion stills of that event (Agenda.png, ...)")
    ap.add_argument("--header", type=Path, help="the Meetup header image, downloaded")
    ap.add_argument("--canva", type=Path, help="Canva exports: header.png, talk-1.png, ... (used before the stills)")
    ap.add_argument("--keep-fonts", action="store_true", help="keep the template's fonts instead of Amazon Ember everywhere")
    ap.add_argument("--out", type=Path, required=True, help="where to write the filled-in deck")
    args = ap.parse_args(argv)

    for p in (args.template, args.event):
        if not p.is_file():
            print(f"{p} not found", file=sys.stderr)
            return 2
    try:
        report = build(Inputs(args.template, args.event, args.stills, args.header, args.canva, None if args.keep_fonts else "Amazon Ember"), args.out)
    except FileNotFoundError as err:
        print(err, file=sys.stderr)
        return 2
    print(f"wrote {args.out}, visible slides in running order:")
    for line in report.outline:
        print(f"  {line}")
    for line in report.todo:
        print(f"still to do: {line}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
