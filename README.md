# meetup-deck

Fills in the group's slide deck for one meetup from its `event.json`, so the
deck the organizers already present from gets the real agenda from nothing
more than the Meetup link.

## Part of the meetup tools

One of the repositories behind the AWS User Groups Austria meetup evenings, in
[aws-user-groups-austria](https://github.com/aws-user-groups-austria). Start with
[`meetup-workspace`](https://github.com/aws-user-groups-austria/meetup-workspace): it clones all of them (`make
clone`), installs them and runs the whole evening (`make start`); the contract
between them is its `SPEC.md`.

| Repository | What |
|---|---|
| [`meetup-workspace`](https://github.com/aws-user-groups-austria/meetup-workspace) | The Makefile, the spec and the guide |
| [`meetup-sync`](https://github.com/aws-user-groups-austria/meetup-sync) | A Meetup page to `event.json` |
| [`meetup-event-vienna`](https://github.com/aws-user-groups-austria/meetup-event-vienna) | One folder per meetup, shared photos and logos |
| [`meetup-stream-templates`](https://github.com/aws-user-groups-austria/meetup-stream-templates) | The room screen: Remotion scenes and the web player |
| **[`meetup-deck`](https://github.com/aws-user-groups-austria/meetup-deck)** | The slide deck, filled in |
| [`aws-community-austria-website`](https://github.com/aws-user-groups-austria/aws-community-austria-website) | awscommunity.at |


The template is that deck, exported from Google Slides as `.pptx`. Its
slides, design and hidden backup slides stay as they are. Slides are found by
what they say, not by their position, so the template can keep changing:

| Template slide (visible) | What happens |
|---|---|
| Welcome, with `Meetup #52` | The event's number, from the title (`AWS Meetup Vienna #53 - hosted by RBI`) |
| Header: `#X Meetup`, one picture over the whole slide | The Canva header (`--canva <dir>/header.png`), else the Meetup event's header image, same box |
| Every `Shoutout` | The sponsors' logos side by side; their names until a logo exists, plus a `#Todo` box. Then the shoutouts at the start and the end become `Thank you to our Location & Food Sponsor`, and a copy of the shoutout, the host's own slot, goes right after the first talk's slide (speaker notes say what happens) |
| `Who we are` | The organizers in the purple area, spread evenly (rows of three and two): photo from `meetup-event-vienna/people/<name>.png`, name on one line, email from `community.json`; "Check our new website", the arrow and the QR code centred in the dark panel |
| `What we do` | Each group's Meetup picture links to the group; one QR code per group, centred in the dark panel |
| `See you soon`, subtitle `at our ...` | `at our next Meetup` and the next-events link, since this page cannot know the next date |
| after the header slide | New: the agenda, one row per part, wrapped titles under the title column |
| after `Introducing Our User Group Members` | New: one slide per talk (its Canva card `<dir>/talk-1.png`, else the room screen's speaker handover still), the break with "back at" and who is next, the closing with the feedback QR code and how long people can stay, and the next AWS Community Day with the community's speakers, its website screenshot, link and QR code (from `upcoming` in `event.json`) |

The agenda, break, closing and community day slides are native slides in
the deck's own design (`src/meetup_deck/native.py`): the layout of the
template's Shoutout slide, so the master's user group logo sits top right,
its triangle, and text in PT Sans, titles bold #9E5EFB, text white, as the
template's slides have them. They stay editable in Google Slides. Only the
talks are pictures. The times and the handover are in the speaker notes. The template's
own `#Todo` boxes (beside the slide, outside the visible area) are removed
where this fills the slide in; anything left to do gets a new one, and the
command prints it.

The deck holds organizer emails and photos, so neither the template nor the
result is ever committed: both are gitignored in `meetup-event-vienna`.

## Run it

Normally through `meetup-event-vienna`, which renders the stills, downloads
the header image and calls this. In `meetup-event-vienna`:

```sh
make deck EVENT=2026-10-20-aws-meetup-vienna-53
```

Directly, from this folder:

```sh
uv run meetup-deck --template ../meetup-event-vienna/deck/aws-ug-vienna-austria.pptx --event ../meetup-event-vienna/events/2026-10-20-aws-meetup-vienna-53/event.json --stills ../meetup-stream-templates/out/stills --header ../meetup-event-vienna/events/2026-10-20-aws-meetup-vienna-53/deck/header.jpeg --out ../meetup-event-vienna/events/2026-10-20-aws-meetup-vienna-53/deck/AWS-Meetup-Vienna-53.pptx
```

Into Google Slides: upload the `.pptx` to Drive and open it with Google
Slides (or File, Import slides, in the existing deck).

## Test it

The tests build a small template with the same marker slides, so they need
neither the real deck nor the network:

```sh
uv run pytest
```

## Layout

| Path | What |
|---|---|
| `src/meetup_deck/build.py` | Finds the marker slides and fills them in |
| `src/meetup_deck/cli.py` | The `meetup-deck` command |
| `tests/` | pytest, against a generated template |

Python 3.12 or later, `python-pptx` 1.x, Pillow, managed with `uv`.
