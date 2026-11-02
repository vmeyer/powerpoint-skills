# Style file

Optional. A deck without a master starts with `pptx init decks/<slug>` on a neutral 16:9 template, and
`pptx theme decks/<slug> --color accent1=#0B5CAD --font major=Georgia` sets its colors and fonts at any time.
A style file does the same at `init` and can also change the check thresholds, which is handy when several
decks share them: `pptx init decks/<slug> --style-file style.json`.

```json
{
  "colors": {"accent1": "#0B5CAD", "accent2": "#F28C28", "dk2": "#222222"},
  "fonts": {"major": "Georgia", "minor": "Verdana"},
  "rules": {"min_size_pt": 10, "max_lines": 4}
}
```

- `colors`: any of `dk1 lt1 dk2 lt2 accent1`..`accent6 hlink folHlink`, each `#RRGGBB`.
- `fonts`: `major` (headings) and `minor` (body) family names.
- `rules`: check thresholds as numbers above 0: `max_lines` and `max_title_lines` (used by every check),
  `min_size_pt`, `min_size_small_pt`, `max_bullets`, `near_align_cm`, `title_min_words` (used with
  `--strict`) and `snap_cm` (how far boxes snap to a master's guides). `rules` also applies to a deck with a
  master (`--from`); colors and fonts only shape the neutral template.

Every key is optional; `init` rejects an unknown or malformed key and names it before it writes anything.
