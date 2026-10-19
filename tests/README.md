# tests

pytest for `meetup_deck`. `conftest.py` builds a small template with the same
marker slides as the group's deck (welcome number, header picture, shoutout,
members, a hidden backup, see you soon), so the tests need neither the real
deck, which is never committed, nor the network. Run with `uv run pytest`
from `..`.
