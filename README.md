# Orai Fieldwork App (EDD-I Orai, DVVNL)

Field survey PWA for Orai town, built from Kuthond Survey Application v20.19.

## Orai-specific changes
- `app.js` → `SERVER_URL` is a placeholder: paste the Orai Apps Script Web App URL (ends in `/exec`).
- Separate offline caches (`orai-fieldwork-pwa-*`, `orai-fieldwork-master-bg`) so it never clashes with the Kuthond app.
- Name/title: "Orai Fieldwork App".
- Survey, login, master-data, upload and sync logic unchanged.

## Live link (GitHub Pages)
https://chauhannagendra005-bit.github.io/orai-field-survey/

## Release rule
Every time any file changes, bump `CACHE_VERSION` in `sw.js` so phones pick up the new version.
