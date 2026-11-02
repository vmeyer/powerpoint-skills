#!/usr/bin/env bash
#
# Regenerate .codex-plugin/plugin.json from the canonical .claude-plugin/plugin.json.
#
# The repository root IS the plugin. Claude Code (marketplace or --plugin-dir)
# reads .claude-plugin/plugin.json directly. Codex CLI reads only
# .codex-plugin/plugin.json, which this script derives from it plus a pointer to
# skills/. Codex plugins consume skills only (this plugin has no agents).
#
# Never edit the generated file by hand: change plugin.json and re-run.
#
# Usage:  ./scripts/sync-mirrors.sh [--check]
#   --check  exit non-zero if the generated manifest is out of date, without writing.
#
set -euo pipefail
cd "$(dirname "$0")/.."

CHECK=0
[[ "${1:-}" == "--check" ]] && CHECK=1

codex_manifest() {
  python3 - <<'PY'
import json
with open(".claude-plugin/plugin.json") as f:
    p = json.load(f)
out = {"name": p["name"], "version": p["version"], "description": p["description"]}
if "author" in p:
    out["author"] = p["author"]
for k in ("homepage", "repository", "license"):
    if k in p:
        out[k] = p[k]
out["skills"] = "./skills/"
print(json.dumps(out, indent=2))
PY
}

if [[ "$CHECK" == "1" ]]; then
  diff <(codex_manifest) .codex-plugin/plugin.json >/dev/null 2>&1 || { echo "✗ out of date: .codex-plugin/plugin.json (run sync-mirrors.sh)"; exit 1; }
  echo "✓ generated artifacts are up to date"
  exit 0
fi

mkdir -p .codex-plugin
codex_manifest > .codex-plugin/plugin.json
echo "✓ regenerated .codex-plugin/plugin.json from .claude-plugin/plugin.json"
