# Übergabe: PPTX-Skill (`pptx-deck`)

Kontext für die Fortsetzung in Claude Code oder Codex. Zuerst diese Datei lesen, dann `REQUIREMENTS.md`.

## Worum es geht

Volker baut einen agentenübergreifenden Skill, mit dem Claude Code, Codex u. a. PowerPoint-Decks erstellen, bearbeiten und prüfen. Er soll dem bestehenden Skill-Repo `https://github.com/vmeyer/diagram-skills` (Excalidraw- und HTML-Diagramme mit PNG-Render und Prüfschleife) gleichwertig sein. Ziel ist auch, die Ergebnisse von Claude und Codex mit derselben Engine fair zu vergleichen.

## Getroffene Entscheidungen

1. **Clean Room.** Anthropics eingebauter pptx-Skill ist proprietär (LICENSE.txt verbietet Kopieren außerhalb der Anthropic-Dienste und abgeleitete Werke). Er wurde nur auf Ebene der Fähigkeiten verglichen. Kein Code, keine Texte, keine Listen daraus übernehmen. Umsetzung nur aus öffentlichen Quellen (ECMA-376, python-pptx-, PptxGenJS-, LibreOffice-Doku).
2. **Kein eigenes Theming-System.** Der Skill liest einen bestehenden Master (.pptx/.potx) als Kontext, erzeugt daraus einen Stil-Steckbrief und baut neue Folien als Kopie dieser Datei auf ihren Layouts (REQUIREMENTS §5).
3. **Screenshot-Nachbau.** Aus einem Screenshot entsteht eine native, editierbare Folie im Stil des Masters – Struktur aus dem Screenshot, Stil aus dem Master (§6).
4. **Spezifikation statt direktem Code.** Der Agent schreibt eine YAML-Spezifikation, eine gemeinsame Python-Engine baut daraus die .pptx. Korrekturen immer in der Spezifikation.
5. **Vierstufige Prüfung** (Struktur, Rendering, deterministische Regeln, visueller Review) mit JSON-Report; Stufen 1–3 laufen auch ohne Agent in CI (§7).
6. **Aufbau wie diagram-skills:** `.claude-plugin/plugin.json`, `skills/<name>/SKILL.md`, gemeinsame Skripte unter `references/`, `uv` + `pyproject.toml`, Ausgabe nach `decks/<thema>/` (§10).
7. **Name `pptx-deck`** (vorläufig), um nicht mit Anthropics eingebautem `pptx`-Skill zu kollidieren.

## Wichtigste Erkenntnisse aus dem Fähigkeiten-Abgleich

- Bearbeitung bestehender Decks braucht einen XML-Pfad mit Hilfsskripten (Folie duplizieren inkl. Relationen, verwaiste Teile entfernen); python-pptx allein reicht nicht.
- XSD-gültig ≠ öffnet in PowerPoint (v. a. Diagramm-XML) → eigene semantische Checks.
- Bei Vorlagen nur neu eingeführte Validierungsfehler zählen (Baseline).
- LibreOffice ersetzt fehlende Schriften; ohne metrisch kompatiblen Ersatz sind Overflow-Prüfungen unzuverlässig.
- Inhalts- und Design-Methodik gehört in die SKILL.md (eine Kernaussage pro Folie, Publikum, Anti-Muster) – analog zu „Diagrams should argue“ im Excalidraw-Skill.

## Hinweise zum diagram-skills-Repo

- README beschreibt nur die Installation für Claude Code, nicht für Codex.
- Der Clone-Befehl in der README zeigt noch auf `coleam00/excalidraw-diagram-skill`.
- Validierung dort: JSON-Strukturcheck in `render_excalidraw.py` + visuelle Schleife; der PPTX-Skill braucht deutlich mehr Strukturprüfung und LibreOffice statt Playwright.

## Nächste Schritte

1. Offene Fragen in REQUIREMENTS §11 klären (vor allem: gleiches Repo?, Engine, erster Test-Master).
2. Einen echten Master und 2–3 Screenshots als Testfälle in `tests/` ablegen.
3. Mit `inspect_master.py` und `render_pptx.py` beginnen – sie sind Grundlage für alles Weitere.
4. Danach `build_deck.py` + Spezifikationsschema, dann `check_pptx.py`.

## Startprompt

> Lies HANDOFF.md und REQUIREMENTS.md. Lass uns zuerst die offenen Fragen in §11 klären und danach mit `inspect_master.py` beginnen.

Online-Version des Anforderungsdokuments (claude.ai): https://claude.ai/artifact/BFeJMaBDzX2mz7kHvaecR7
