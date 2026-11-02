# Style file

A small JSON file for users without a PowerPoint master. `pptx init decks/<slug> --style-file style.json`
builds the working copy on a neutral 16:9 template (Arial, grey-blue palette) with the file's colors and
fonts. Every key is optional, and an empty `{}` gives the plain neutral template.

```json
{
  "colors": {"accent1": "#0B5CAD", "accent2": "#F28C28", "dk2": "#222222"},
  "fonts": {"major": "Georgia", "minor": "Verdana"},
  "rules": {"min_size_pt": 10, "max_bullets": 6}
}
```

- `colors`: any of `dk1 lt1 dk2 lt2 accent1`..`accent6 hlink folHlink`, each `#RRGGBB`.
- `fonts`: `major` (headings) and `minor` (body) family names.
- `rules`: any of the check thresholds listed under "Check rules" in `style/brief.md`
  (`min_size_pt`, `min_size_small_pt`, `max_lines`, `max_title_lines`, `max_bullets`, `near_align_cm`,
  `snap_cm`, `title_min_words`), as numbers above 0. `rules` also applies when the deck comes from a
  master (`--from`).

`init` rejects an unknown or malformed key and names it (for example `accent9`) before it writes anything.
Colors and fonts only shape the neutral template; a master keeps its own theme.
