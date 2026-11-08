# Trip Tour Guide template: how this repo works

This repo is a reusable **trip tour guide**. The same trip data drives two outputs:
1. `index.html`: one self-contained, mobile-first web app (served by GitHub Pages; works offline when saved)
2. `Desert-to-Divide-Tour-Guide.pptx`: an illustrated PowerPoint guide

## Files
| Path | What it is |
|---|---|
| `src/data.js` | **All trip content.** Title, dates, start/finish times (countdown), stats, and `days[]`. Each day has stops, food, sleep, a pastel `color`, an `ink` text color and an `art` key. Also `checklist` and `heads`. |
| `src/art.js` | Original pastel SVG scenes (`canyon`, `monument`, `arch`, `black`, `mountains`, `road`, `turkey`) plus the Estonian flag. Add new scenes here for new places. |
| `src/src.html` | The app template (UI, CSS, behavior). Placeholders: `__FLAG__`, `__HEROART__`, `__DOWNLOADS__`, `__AUDIO__`, `/*__DATA__*/`. |
| `src/build.js` | Inlines the data, art and audio into `index.html`. Run `node src/build.js`. |
| `src/deck.js` | Builds the PowerPoint with pptxgenjs from the same data. Needs `pptxgenjs react react-dom react-icons sharp`, plus `PPTX_APPLY_THEME` pointing at the pptx skill's `scripts/apply_theme.js`. Uses the PNGs in `src/img/`. |
| `src/audio/` | `gobble.mp3` (the embedded sound). `tts.py` and `turkey.py` regenerate it with espeak-ng (`pip install espeakng-loader`), numpy and scipy, plus ffmpeg. |

## Rework for a new trip
1. Copy this repo to a new name (e.g. `trip-<place>`). Enable Pages: Settings → Pages → Deploy from branch → `main` / root.
2. Rewrite `src/data.js` for the new trip:
   - `title`, `subtitle`, `dates`, `start`, `finish` (ISO with offset)
   - For each day: `day`, `date`, `title`, `route`, `drive`, `color`/`ink`/`art`, `stops[]`, `food[]`, `sleep`.
   - Stop fields: `t` (time), `kind` (hike, fun, food, drive or home), `name`, `desc`, optional `options[]` (hike A/B/C), `list[]`, `map` (use the `M("place")` Google Maps helper), `url`, `star` (highlight), `gobble` (shows the sound button).
3. Personalize it: change the splash greeting in `src/src.html` (`Tere, Mari!` / `Head reisi`) and the flag in `art.js` (`ESTONIA`) for the traveler, and swap the gobble audio for any fun sound.
4. Add or modify scenes in `src/art.js` for the new places. Re-render the PNGs for the deck by rasterizing each `ART[key]` with sharp to 1600×900, plus `-sq` (900×900) and `-tall` (700×900) cover crops.
5. `node src/build.js && node src/deck.js`, open `index.html` to check it, then commit and push.
6. Verify the research before sharing: park hours and closures, restaurant off-season hours, and lodging availability.

## Conventions
- Pastel palette: peach `F7B89A`, rose `F2A7B8`, lavender `C3B1E1`, mint `A8D8C9`, sky `B8D8F0`. Ink `3D3A4B`.
- Keep `index.html` single-file (no external requests) so it works offline and from a downloaded copy.
- Check-offs are stored in `localStorage` under the `d2d2-` prefix. Change the prefix for a new trip.
