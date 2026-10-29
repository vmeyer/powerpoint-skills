# Anforderungen: PPTX-Skill für mehrere Agenten

Stand: 29.10.2026 · Autor: Volker · Arbeitsname des Skills: `pptx-deck`

## 1. Ziel und Scope

Der Skill lässt jeden Coding-Agenten (Claude Code, Codex, Gemini CLI u. a.) PowerPoint-Dateien erstellen, bearbeiten und lesen – im Stil eines vorhandenen Masters, reproduzierbar und mit automatischer Prüfung inkl. PNG-Vorschau. Er ist das PPTX-Gegenstück zum bestehenden Excalidraw-Skill (`vmeyer/diagram-skills`) und folgt demselben Muster: erzeugen → rendern → prüfen → korrigieren.

**Im Scope**

- Neue Präsentationen aus Briefing, Gliederung oder strukturierten Daten erzeugen
- Bestehende .pptx bearbeiten (Text, Folien, Reihenfolge, Diagramme, Bilder)
- Inhalte aus .pptx extrahieren (Text, Notizen, Struktur, Bilder)
- Bestehenden Master (.pptx/.potx) als Stil-Kontext lesen und neue Folien im selben Stil bauen, auch nach Screenshot-Vorlage
- Automatische Prüfung: Dateivalidität, Rendering, Layout- und Stilregeln, visuelle Selbstkontrolle

**Nicht im Scope (v1)**

- Animationen und Folienübergänge erzeugen (nur erhalten, wenn vorhanden)
- Eingebettete Videos/Audio erzeugen
- Makros (.pptm) ausführen oder schreiben
- Live-Kollaboration oder Cloud-APIs (Google Slides, PowerPoint Online)

**Clean-Room-Prinzip:** Spezifikation und Umsetzung entstehen aus öffentlichen Quellen (ECMA-376 / ISO/IEC 29500, Doku von python-pptx, PptxGenJS, LibreOffice) – nicht aus Code oder Text proprietärer Skills (siehe HANDOFF.md, Lizenz).

## 2. Zielplattformen und Agenten-Kompatibilität

Der Skill folgt dem offenen Agent-Skills-Standard (agentskills.io): ein Ordner mit `SKILL.md` (YAML-Frontmatter `name`, `description`) plus Skripten. Alle Logik steckt in CLI-Skripten, nicht in agentenspezifischen Tools.

| Agent | Ablageort | Besonderheit |
| --- | --- | --- |
| Claude Code | Plugin (`.claude-plugin/`) oder `.claude/skills/pptx-deck/` | Lädt SKILL.md bei Bedarf |
| OpenAI Codex | `.agents/skills/pptx-deck/` (Repo) oder User-Scope | Optional `agents/openai.yaml` für UI-Metadaten |
| Gemini CLI, Cursor u. a. | Jeweiliger Skills-Ordner oder Verweis aus `AGENTS.md` | Fallback: `AGENTS.md` verweist auf SKILL.md |

- K1: Keine agentenspezifischen Tool-Aufrufe in der SKILL.md; nur Shell-Befehle.
- K2: Jedes Skript hat `--help`, klare Exit-Codes (0 = ok, 1 = Prüffehler, 2 = Aufruffehler) und optional `--json`.
- K3: Bildprüfung mit jedem multimodalen Agenten: Skript erzeugt PNGs, der Agent liest sie mit seinem Bild-Tool.
- K4: Agenten ohne Bildverständnis bekommen einen rein textuellen Prüfreport.
- K5: Gleiche Struktur wie der Excalidraw-Skill, damit beide Skills gleich bedient werden.

## 3. Funktionale Anforderungen

Priorität: **M** = Muss (v1), **S** = Soll, **K** = Kann.

**Erstellen**

| ID | Anforderung | Prio |
| --- | --- | --- |
| F1 | Deck aus einer deklarativen Spezifikation (YAML/JSON) erzeugen: Folien mit Layout-Name und Platzhalter-Inhalten | M |
| F2 | Ausschließlich Layouts des Masters nutzen (übergebene Quelldatei; ohne Quelle: neutrales Template); Inhalte in Platzhalter statt frei positionierter Textboxen | M |
| F3 | Folientypen: Titel, Agenda, Kapiteltrenner, Text, Zwei-Spalten, Bild + Text, Diagramm, Tabelle, Zitat, KPI-Kacheln, Abschluss | M |
| F4 | Native Diagramme (Balken, Linie, Kreis/Donut, gestapelt) mit editierbaren Daten statt Bildern | M |
| F5 | Native Tabellen im Tabellenstil der Quelldatei | M |
| F6 | Bilder einfügen (PNG, JPG, SVG) mit Seitenverhältnis-Erhalt und Zuschnitt auf Platzhalter | M |
| F7 | Sprechernotizen pro Folie | M |
| F8 | Excalidraw-Diagramme aus dem bestehenden Skill als SVG/PNG einbinden | S |
| F9 | Einfache Formen und Pfeile (Prozess, Timeline) als native Shapes | S |
| F10 | Folien-Metadaten: Titel, Autor, Sprache, Foliennummern, Fußzeile | S |
| F11 | Icons aus einer hinterlegten Icon-Bibliothek | K |

**Bearbeiten**

| ID | Anforderung | Prio |
| --- | --- | --- |
| F20 | Text in bestehenden Folien ersetzen, ohne Formatierung zu verlieren | M |
| F21 | Folien hinzufügen (aus Layout oder als Kopie einer bestehenden Folie), löschen, umsortieren, duplizieren | M |
| F22 | Diagrammdaten aktualisieren | S |
| F23 | Bestehende Fremd-Decks auf einen anderen Master umziehen (Re-Branding) | S |
| F24 | Kommentare lesen und optional hinzufügen | K |
| F25 | Decks zusammenführen und aufteilen | K |

**Lesen und Extrahieren**

| ID | Anforderung | Prio |
| --- | --- | --- |
| F30 | Text, Notizen und Gliederung als Markdown/JSON ausgeben | M |
| F31 | Struktur-Inventar: Folien, verwendete Layouts, Platzhalter, Shapes mit Position und Größe | M |
| F32 | Bilder und Medien extrahieren | S |
| F33 | Master/Theme analysieren: Farben, Schriften, Layouts mit Platzhaltern (Basis für den Stil-Steckbrief, S1) | M |

**Workflow-Regel:** Der Agent plant zuerst die Gliederung (Spezifikation), erzeugt dann das Deck, prüft es und korrigiert iterativ – maximal 3 Runden, danach Report an den Nutzer.

## 4. Unterstützte Ein- und Ausgabeformate

| Format | Rolle | Lesen | Schreiben | Prio |
| --- | --- | --- | --- | --- |
| .pptx | Präsentation | ja | ja | M |
| .potx | Master/Vorlage als Stilquelle | ja | nein (nur als Quelle) | M |
| .pptm | Präsentation mit Makros | ja, Makros werden ignoriert | nein | K |
| .odp | OpenDocument-Präsentation | über LibreOffice-Konvertierung | nein | K |
| .ppt (binär, alt) | Altbestand | über LibreOffice-Konvertierung | nein | K |
| YAML/JSON | Deck-Spezifikation | ja | ja | M |
| Markdown | Briefing oder Gliederung → Spezifikation | ja | ja (Extraktion) | M |
| CSV/XLSX | Diagramm- und Tabellendaten | ja | nein | S |
| PNG/JPG | Bilder, Screenshots, Folien-Renderings | ja | ja (Renderings) | M |
| SVG | Vektorgrafik, Icons, Excalidraw-Export | ja | nein | S |
| .excalidraw | Diagramm aus dem bestehenden Skill (via dessen Export) | ja | nein | S |
| PDF | Export für Review und Versand | nein | ja | M |
| JSON-Report | Prüfergebnis | – | ja | M |

- SVG: PowerPoint unterstützt SVG nativ, LibreOffice-Rendering teils nicht – immer PNG-Fallback einbetten.
- Schriften des Masters müssen in der Render-Umgebung installiert sein, sonst stimmen Umbrüche im PNG nicht mit PowerPoint überein.

## 5. Stil aus bestehendem Master übernehmen

Kein eigenes Theming-System: Der Skill liest eine vorhandene .pptx/.potx als Kontext und leitet daraus alle Stilinformationen ab. Diese Datei ist die Quelle der Wahrheit; neue Folien entstehen immer in einer Kopie davon, auf ihren Layouts.

| Ebene | Information | Nutzung |
| --- | --- | --- |
| Theme | Farbschema (dk1, lt1, dk2, lt2, accent1–6, Links), Schriftschema (Überschrift, Text) | Farben und Schriften nur als Theme-Referenz setzen |
| Folienmaster | Hintergrund, Logo, Fußzeile, Foliennummer, Text-Stile je Ebene | erben, nie nachbauen |
| Layouts | Name, Platzhalter mit Typ, Position und Größe | passendes Layout je Folie wählen |
| Bestehende Folien | Wiederkehrende Kompositionen: Karten, KPI-Kacheln, Bildbehandlung, typische Abstände und Ränder | als Klon-Vorlage und Stilreferenz |
| Diagramme, Tabellen | Farbfolge, Schriftgrößen, Achsen, Gitternetz, Tabellenstil | gleicher Stil für neue Diagramme und Tabellen |

- S1 (M): `inspect_master` erzeugt aus einer .pptx/.potx einen Stil-Steckbrief (Markdown + JSON) – Theme, Layouts mit Platzhaltern, Text-Stile, gemessene Abstände, Diagramm- und Tabellenstil – plus einen beschrifteten Kontaktabzug aller Layouts und Folien. Der Agent liest beides vor jeder Erstellung.
- S2 (M): Neue Decks entstehen als Kopie der Quelldatei; vorhandene Folien werden entfernt oder als Klon-Vorlage genutzt (L1). Master und Layouts bleiben unverändert.
- S3 (M): Farben und Schriften als Theme-Referenzen, nicht als feste Werte.
- S4 (M): Frei gestaltete Elemente nutzen die Abstände, Ränder, Eckenradien und Farben, die der Steckbrief in bestehenden Folien gemessen hat.
- S5 (S): Ohne Quelldatei greift ein mitgeliefertes neutrales, strukturiertes Template (L4).
- S6 (S): Die Prüfregeln (Stufe 3) leiten das Erlaubte aus dem Steckbrief ab.
- S7 (K): Optionale Ergänzungsdatei neben der Quelle für Regeln, die nicht im Master stehen (z. B. max. 5 Bullets).

## 6. Folien nach Screenshot nachbauen

Ein Screenshot – eigene Folie, fremdes Deck, Webseite oder Skizze – liefert Aufbau und Inhalt; das Ergebnis ist eine native, editierbare Folie im Stil der Quelldatei, kein eingefügtes Bild.

1. Analyse: Raster, Elemente, Hierarchie und Inhalte beschreiben; Diagrammwerte ablesen oder erfragen.
2. Mapping: Layout wählen, Elemente auf Platzhalter und native Shapes abbilden, Farben auf Theme-Rollen abbilden statt sie zu übernehmen.
3. Bauen und rendern.
4. Vergleich: Render neben Screenshot – gleiche Struktur, Stil der Quelldatei.

- SC1 (M): Eingabe PNG/JPG, auch mehrere Screenshots für mehrere Folien.
- SC2 (M): Ergebnis vollständig nativ. Bildausschnitte aus dem Screenshot nur für echte Fotos/Grafiken und nur nach Freigabe.
- SC3 (M): Stil kommt aus der Quelldatei, nicht aus dem Screenshot; Elemente ohne Entsprechung im Theme werden gemeldet.
- SC4 (M): `compare_screenshot` stellt Screenshot und Render nebeneinander in ein PNG; der Agent prüft Vollständigkeit und Anordnung.
- SC5 (S): Unleserliche Texte oder Zahlen werden markiert und erfragt, nicht geraten.
- SC6 (S): Modus „nur Struktur“: Screenshot liefert das Layout, Inhalte kommen aus dem Briefing.
- SC7 (K): Strukturähnlichkeit (z. B. SSIM auf Kanten-Maske) als Kennzahl.

## 7. Prüfung und Qualitätssicherung

Vier Stufen, von billig/deterministisch zu teuer/visuell; alle schreiben in einen gemeinsamen JSON-Report mit Fundstelle (Folie, Shape) und Schweregrad. Korrigiert wird immer in der Spezifikation, nie direkt in der .pptx.

| Stufe | Was | Wie | Ergebnis |
| --- | --- | --- | --- |
| 1 Struktur | Gültiges OOXML, öffnet ohne Reparaturdialog | ZIP-Integrität, Content-Types, Relationships, XSD gegen ECMA-376, semantische PowerPoint-Checks (L2), Roundtrip mit python-pptx | error |
| 2 Rendering | Datei lässt sich rendern | LibreOffice headless → PDF → PNG je Folie (pdftoppm) + Kontaktabzug | PNGs, error bei Abbruch |
| 3 Regeln | Layout- und Stilregeln | Geometrie aus dem XML gegen den Stil-Steckbrief (S1, S6) | error / warning |
| 4 Visuell | Gesamteindruck | Agent betrachtet PNGs anhand fester Checkliste | Befunde im Report |

**Regelkatalog Stufe 3**

| Regel | Prüfung | Schwere |
| --- | --- | --- |
| Overflow | Text passt in seinen Platzhalter; Messung über Font-Metriken | error |
| Safe-Area | Kein Shape außerhalb Folie oder Rand | error |
| Überlappung | Keine ungewollten Überlappungen | warning |
| Vorlagen-Reste | Keine leeren Platzhalter, kein Blindtext (Lorem, XXX, TODO), keine überzähligen Slots | error |
| Farben | Nur Theme-Farben | warning |
| Schriften | Nur Theme-Schriften, Größen im Bereich der Master-Text-Stile | error |
| Kontrast | WCAG 2.x, ≥ 4,5:1 für Text | warning |
| Textmenge | Bullets/Wörter/Zeichen innerhalb der Limits | warning |
| Bilder | Keine Verzerrung, Auflösung ≥ Schwelle, Alt-Text | warning |
| Pflicht-Elemente | Logo, Foliennummer, Fußzeile wie im Master | error |
| Konsistenz | Titel-Position/-Größe je Folientyp gleich; Diagrammbeschriftung einheitlich | warning |
| Layout-Treue | Jede Folie nutzt ein Master-Layout | warning |
| Rechtschreibung | Optional, z. B. LanguageTool lokal | info |

**Checkliste Stufe 4:** abgeschnittener/überlaufender Text und ungünstige Umbrüche · Elemente zu dicht am Rand oder ungleich verteilt · Lesbarkeit von Diagrammen · eine erkennbare Kernaussage pro Folie · Stimmigkeit mit bestehenden Folien der Quelldatei bzw. dem Screenshot.

- Q1 (M): Ein Befehl `check` führt alle Stufen aus und liefert Report (JSON + Textzusammenfassung) und PNG-Pfade.
- Q2 (M): Stufen 1–3 laufen ohne Agent, auch in CI/CD.
- Q3 (M): Report nennt pro Befund Folie, Shape-Name, Regel und Korrekturvorschlag.
- Q4 (M): Render-Umgebung mit den Schriften des Masters; fehlende Schrift = eigener Befund (Font-Substitution erkennen).
- Q5 (S): Stufe 4 „blind“: separater Prüfagent ohne Erstellungskontext.
- Q6 (S): Abweichungen LibreOffice vs. PowerPoint dokumentieren; optional Abgleich per PowerPoint-Export.
- Q7 (K): Visuelle Regressionstests gegen Referenz-PNGs für den Skill selbst.

## 8. Erkenntnisse aus dem Abgleich mit dem Anthropic-pptx-Skill

Abgleich nur auf Ebene der Fähigkeiten. Der Kern deckt sich; Anthropic ist weiter bei Bearbeitung, Datei-Robustheit und Inhalts-/Design-Methodik, dieser Entwurf bei deterministischen Regeln, JSON-Report und CI-Tauglichkeit.

- L1 (M): Zwei Pfade – Neuerstellung über die Engine, Bearbeitung bestehender Decks/Vorlagen auf XML-Ebene mit Hilfsskripten zum Duplizieren einer Folie samt Relationen und zum Entfernen verwaister Teile (python-pptx kann nicht duplizieren und verliert beim einfachen Text-Setzen die Formatierung).
- L2 (M): Validierung über XSD hinaus – PowerPoint lehnt Dateien ab, die XSD, python-pptx und LibreOffice akzeptieren (v. a. Diagramm-XML, Element-Reihenfolge). Eigene semantische Checks mit Korrekturhinweis.
- L3 (M): Baseline-Modus für Vorlagen: nur neu eingeführte Fehler zählen.
- L4 (M): Immer strukturiert: Theme-Farben/-Schriften, benannte Layouts mit Platzhaltern, Abschnitte, Foliennummer/Fußzeile im Layout.
- L5 (M): Render-treue Schriften: metrisch kompatibler Ersatz in der Render-Umgebung (z. B. Liberation, Carlito, Caladea); sonst Puffer und Overflow nur als Warnung.
- L6 (M): Vorlagen-Reste erkennen; überzählige Slots als ganze Gruppe löschen.
- L7 (S): Folienübergreifende Konsistenz und Layout-Abwechslung.
- L8 (S): Kontaktabzug nur zur Orientierung, Prüfung an voll aufgelösten Einzelbildern; nur geänderte Folien neu rendern.
- L9 (S): Barrierefreiheit: Alt-Texte, Textfelder als Text ausgezeichnet, Lesereihenfolge.
- D1 (M): Inhalts-Methodik in der SKILL.md: eine Kernaussage pro Folie, Aufbau nach Zweck, Publikum (Live vs. Pre-Read), Umfang nach Redezeit – Gegenstück zu „Diagrams should argue“.
- D2 (M): Design-Methodik: Folientyp-Bibliothek, visuelles Element statt Textwüste, eigene Anti-Muster-Liste generischer KI-Folien.

## 9. Nichtfunktionale Anforderungen

| ID | Anforderung |
| --- | --- |
| N1 | Läuft lokal und offline |
| N2 | Nur permissive Open-Source-Abhängigkeiten: python-pptx (MIT), lxml, Pillow, defusedxml, LibreOffice (MPL), poppler; optional PptxGenJS (MIT) |
| N3 | Eine Sprache für alle Skripte (Vorschlag: Python) |
| N4 | Setup per `uv sync`, optional Dockerfile mit LibreOffice und den Schriften des Masters |
| N5 | Deterministisch: gleiche Spezifikation → gleiche .pptx |
| N6 | Erzeugen < 5 s, `check` für 20 Folien < 60 s |
| N7 | SKILL.md ≤ 300 Zeilen; Details in `references/` (Progressive Disclosure) |
| N8 | Linux, macOS, Windows (WSL) |
| N9 | Kein Datenabfluss: Master und Inhalte bleiben lokal |
| N10 | Eigene Tests: Golden-Specs mit erwarteten Reports |

## 10. Vorgeschlagener Skill-Aufbau

Angelehnt an `vmeyer/diagram-skills`: Plugin-Manifest, Skills unter `skills/`, gemeinsame Skripte unter `references/`, `uv`, feste Ausgabe-Konvention. Übernommene Muster: Methodik vor Technik, Design-Prozess in Schritten, Pattern-Bibliothek, Pflicht-Schleife „Render & Validate“ mit Abbruchregel, abschnittsweises Erzeugen großer Ergebnisse.

```
pptx-skills/
├── .claude-plugin/plugin.json
├── skills/
│   └── pptx-deck/
│       ├── SKILL.md              # Methodik, Workflow, Render-&-Validate-Loop
│       └── references/
│           ├── spec-format.md    # Schema der Deck-Spezifikation
│           ├── slide-types.md    # Folientyp-Bibliothek
│           ├── screenshot.md     # Ablauf Screenshot → native Folie
│           └── editing.md        # XML-Pfad für bestehende Decks (L1)
├── references/
│   ├── neutral-template.potx     # Fallback ohne Quelldatei (S5)
│   ├── inspect_master.py         # .pptx/.potx → Stil-Steckbrief + Kontaktabzug (S1)
│   ├── build_deck.py             # Spezifikation + Quelldatei → .pptx
│   ├── edit_pptx.py              # duplizieren, sortieren, löschen, aufräumen
│   ├── extract_pptx.py           # .pptx → Markdown/JSON, Inventar
│   ├── render_pptx.py            # .pptx → PDF → PNG + Kontaktabzug
│   ├── compare_screenshot.py     # Screenshot und Render nebeneinander (SC4)
│   ├── check_pptx.py             # Stufen 1–3, JSON-Report
│   ├── schemas/                  # ECMA-376-XSDs, spec.schema.json
│   ├── Dockerfile                # optional: LibreOffice + Schriften
│   └── pyproject.toml            # uv
└── tests/                        # Golden-Specs, Beispiel-Master, erwartete Reports
```

**Ausgabe-Konvention:** `decks/<thema>/<thema>.pptx`, `<thema>.pdf`, `png/slide-NN.png`, `report.json`.

**Kern-Workflow in der SKILL.md**

1. Quelldatei mit `inspect_master.py` lesen; Steckbrief und Kontaktabzug ansehen
2. Bei Screenshot-Vorlage: analysieren und auf Layouts und Theme-Rollen abbilden
3. Spezifikation (YAML) schreiben und gegen `spec.schema.json` validieren
4. `build_deck.py` → .pptx (als Kopie der Quelldatei)
5. `check_pptx.py` → Report + PNGs; PNGs ansehen, bei Screenshot zusätzlich `compare_screenshot.py`
6. Befunde in der Spezifikation korrigieren, zurück zu 4 (max. 3 Runden)
7. Ergebnis, PDF und Report an den Nutzer

Die Spezifikation statt direktem Code ist bewusst gewählt: agentenneutral, diff-bar und fair für den Vergleich Claude vs. Codex, weil beide dieselbe Engine nutzen.

## 11. Offene Fragen

**Umsetzung**

- [ ] Gleiches Repo wie diagram-skills (ein Plugin, gemeinsame `references/`) oder eigenes Repo?
- [ ] Skill-Name `pptx-deck` statt `pptx` – einverstanden?
- [ ] Engine: alles Python (python-pptx + XML-Pfad) oder PptxGenJS für die Neuerstellung?
- [ ] Neutrales Fallback-Template (S5): hell oder dunkel – oder immer eine Quelldatei voraussetzen?
- [ ] Format: nur 16:9 oder auch 4:3?
- [ ] Bearbeiten fremder Decks (F20–F25, L1) schon in v1?
- [ ] PowerPoint-Rechner für Render-Abgleich verfügbar, oder bleibt LibreOffice die einzige Referenz?
- [ ] Vergleich Claude vs. Codex: Kennzahlen (Prüffehler nach Runde 1, Iterationen, Laufzeit, subjektive Note)?
- [ ] Sprache der Folien: nur Deutsch oder mehrsprachig?

**Stil und Screenshots**

- [ ] Welche Master-Datei dient als erster Testfall, und liegen ihre Schriften lokal vor?
- [ ] Screenshot-Nachbau: eher eigene alte Folien, fremde Decks oder Webseiten/Skizzen?
- [ ] Dürfen Fotos aus dem Screenshot ausgeschnitten und übernommen werden (SC2), oder immer Platzhalter?
