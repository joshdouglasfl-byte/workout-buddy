# Workout Buddy

Home-gym workout app for Josh and Michelle. Follow the `josh-app-dev` skill for shared conventions (same setup as Food Log).

- App: static files on GitHub Pages. `josh/` and `michelle/` are per-person entry pages (they set `window.WB_PERSON`); shared code in `app.js` and `app.css`. The root `index.html` redirects to `josh/` so Josh's original home-screen icon keeps working.
- Device storage is per person (`wb.<Person>.<key>`); Josh's pre-split `wb.<key>` data is copied over once on first load.
- Data: Google Sheet "Workout Buddy" (id `1hmU_XKOhSxIM_eculF_hw0agBoI6TKY7EnNb10x5go0`). Per-person tabs have a Person column (Sets, Events, Progress, Plan, Feedback, Exercise Overrides); a blank Person means Josh (rows from before Nov 21, 2026). Settings: blank Person = everyone, a Person row overrides for that person. Exercises is the shared library; per-person tempo/pace/on-off/form tips go in Exercise Overrides.
- API: `apps-script/Code.gs`, deployed by Josh from the standalone project "Workout Buddy API" (https://script.google.com/d/1fkLA4Vdm6ADWtL5yXShb7KZQmm4DbPHyCj1pUARok6aL2vVrud99B4N8/edit), not from the sheet's Extensions menu. `state` returns only the caller's rows (`person` in the request; missing = Josh) and echoes `person`, which tells the app the people-aware script is live. Michelle's page refuses to save until it is. Generic `read` / `append` (with `tab`) / `update` / `delete` match Food Log; build new features on them so the script doesn't need redeploying.
- Feedback: "Tell trainer" / "Tell developer" save to the Feedback tab (Date, Time, Person, For = Trainer or Developer, Message, Status, EntryID, Context). Developer notes from both people come to Josh's developer Claude; review New ones at the start of dev work and mark them Done. Older notes are Events rows with Type note_developer / note_trainer.
- Michelle's setup link: removed 2026-11-21 once her iPad was set up. A new device can use `michelle/#setup=<base64url {url,key}>` (the code that reads it is still in app.js).
- Settings > About shows app version, script status, last load, waiting count and last error: ask Josh for it first when troubleshooting.
- Voice packs in `voice/<pack>/`; Michelle defaults to the `josh` pack.
- Every push: bump `version.json` so open apps reload to the new version.
- Test before pushing: Playwright with the API stubbed via page.route at 390x844 and 820x1180 (see the josh-app-dev skill), and `node --check` plus a SpreadsheetApp mock for Code.gs.

## Watch calories (2026-11-21)
- Settings key `Watch Calories`: `Ask` (default, keypad after "Watch is stopped") or `Shortcut` (name kept from the dropped iPhone-shortcut plan). With `Shortcut` the app skips the keypad and logs an Events `watch_calories` row with Reason `shortcut`; Health Auto Export on that person's iPhone posts the watch workout to the Food Log script, which adds it to Food_Log Exercise with Source `Apple Watch`, EntryID `AW-<start yyyyMMddHHmm>` (see Food Log CLAUDE.md). Josh's row is set; set Michelle's (Person row) once her first watch workout arrives. No-watch flow unchanged.
- 2026-11-21: sheet time zone switched from America/Los_Angeles to America/New_York (matches Code.gs's hardcoded TZ; reads use display values, so nothing shifted).

## Coach notes and What's new (2026-11-21)
- Coach Notes tab (Date, Time, Person, From, Message, Status New/Read, EntryID): the trainer's messages, shown as a "From your trainer" card on home (unread only, last 2) with Got it (generic update Status = Read; ids kept in `wb.<Person>.coachRead` and re-sent until the sheet shows Read) and Reply (opens the trainer note). Settings > Trainer notes lists them all. Loaded after each `state` load.
- Michelle's trainer and dietitian run in Josh's Claude for now (workout-trainer skill, michelle-dietitian skill); her replies go out as Coach Notes.
- `changelog.json` + Settings > About > What's new (dot on the gear until read; `newsSeen`/`latestNews`). Seeded from git log on 2026-11-21; entries before version.json existed (Nov 18-19) use the date as their version.

## Security (2026-11-24; needs one redeploy + new keys)
Same model as Food Log (see its CLAUDE.md Security section).
- Keys only in Script Properties `KEYS` {key: {p: person, t: made}}; the person comes from the key (request `person` ignored). Old `KEY` / `FALLBACK_KEY` accepted until `dropOldKeys()` sets `LEGACY_OFF` (then blank FALLBACK_KEY at the next redeploy). No key = refused.
- Generic actions limited to `GENERIC_TABS` (read/update Coach Notes, append Feedback, no delete); per-person rows only; Person stamped and never changeable. Formula guard as in Food Log.
- App: `<meta name="referrer" content="no-referrer">` on every page; `state` person different from the page's person = load refused with a message (never shows or saves the other person's data).

### Rotate keys (also the first-time setup)
1. script.google.com/d/1fkLA4Vdm6ADWtL5yXShb7KZQmm4DbPHyCj1pUARok6aL2vVrud99B4N8/edit > (first time: Project Settings > Script Properties > `WEB_APP_URL` = the Web app URL from Deploy > Manage deployments; `getUrl()` gives /dev from the editor) > select `newKeys` > Run; the Execution log shows each person's `#setup=` link.
2. On each iPad/iPhone: delete the old Home Screen icon, open the link in Safari, Share > Add to Home Screen, then open the new icon: it shows Connect (iOS drops the #setup part when saving the web app), paste the whole setup link there once (`parseSetup`). Get links to devices by Universal Clipboard or AirDrop.
3. When every device works: Run `dropOldKeys`.
