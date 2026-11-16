# appinio-decks skill

Agent skill (agentskills.io SKILL.md format) that builds Appinio-branded PowerPoint decks
from the official *Appinio Slide Deck Template 2026 (Light Mode)*.

## Contents
- `SKILL.md` – instructions the agent loads
- `scripts/build_deck.py` – renderer (JSON outline → .pptx); `--list-assets` lists names
- `assets/template.pptx` – the template's light-mode master with its 59 layouts (slides removed, images compressed);
  the layouts the script uses are renamed "Appinio – …"
- `assets/` – all 178 Appinio Blue icons, 18 gradients, the image bank (353 photos, 1600 px, by topic) + template photos, logos, flags, world map, Inter fonts (OFL)
- `references/` – design-system brand book (two deck styles), 41 component specs, tokens, data-viz scales, photo catalogue
- `examples/results-outline.json`, `examples/master-outline.json` – one deck per style, every layout

Requires Python 3 + `pip install python-pptx Pillow`. Viewers need Inter installed.

## Install
**LibreChat:** Skills panel → + → Upload a skill → `appinio-decks.zip`. Org-wide: put the folder in `DEPLOYMENT_SKILLS_DIR` (default `./skill`, restart) or a repo synced with `skillSync.github`. Agents need code execution enabled.
**Hermes:** copy the folder to `~/.hermes/skills/design/appinio-decks/`, run `/reload-skills`; Python deps on the host.

## Updating the template
When the brand team changes the Google template: export it as .pptx, delete all slides, keep the first
master, and re-apply the "Appinio – …" layout names (see the `names` map in the maintainer notes below),
then replace `assets/template.pptx`. Layout positions are read from the template, so geometry updates flow through.

Maintainer notes – layout index → name (first master): 0 Cover, 1 Cover photo, 4 Agenda, 5 Agenda long,
8 Chapter, 9 Section intro, 11 Text + photo right, 17 Two columns, 20 Two by two, 22 Three columns,
25 Four columns, 30 Three photos, 33 Full photo, 37 Statement, 39 Statement gradient, 40 Six ideas,
42 Big numbers, 43 Two-liner + panel, 44 Quote, 45 Quote + photo, 50 Team, 51 Back cover lilac,
52 Back cover, 54 Title only, 57 Contact, 58 Blank, 16 Text + photo collage, 32 Title + wide photo, 35 Statement photo.

## Adding assets
Drop files into the matching `assets/` folder; the file name (lower-case, hyphens) becomes the asset name.
Photos: resize to 1600 px on the long edge (JPEG ~q75) and add an entry to `references/photo-catalog.json`.
The partnership photos (Red Bull, Coca-Cola) are deliberately left out of the bank: add them only for decks about those partnerships.
