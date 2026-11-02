# powerpoint-skill

## Skills

_No skills yet._

## Installation

```
/plugin marketplace add vmeyer/claude-plugins
/plugin install powerpoint-skills@vmeyer-plugins
```

For local development, load the repo directly: `claude --plugin-dir ./powerpoint-skills`

## Setup

No setup needed yet.

## Customization

_TBD: what users can adapt (palette, templates, …) and where it lives._

## Output Convention

_TBD: where the skills write their results, e.g. `<type>/<topic-slug>/…`._

## File Structure

```
powerpoint-skills/
  .claude-plugin/plugin.json   # Plugin manifest
  skills/<skill>/SKILL.md      # One folder per skill
  references/                  # Shared by several skills
```
